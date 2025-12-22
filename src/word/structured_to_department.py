"""
structured_to_department.py

text_to_structured.pyの出力（構造化テキスト）を部署ごとにファイル分割する。

【入出力】
- 入力:   data/05_before_department/*.txt
- 出力:   data/06_after_department/{元のファイル名}_depart_（部署名）_{yyyymmddhhmmss}.txt
- ログ:   logs/depart_{yyyymmddhhmmss}.log
- 複数ファイル一括処理（タイムスタンプは全ファイル共通）

【処理仕様】
- [PARENT]ごとにセクション分割
- 各セクションの「- 部署: XX部」から部署名を抽出
- 部署ごとにセクションを連結し、空行で区切る
- ファイル先頭に空行2つ
- エラー時は該当セクションのみスキップ、致命的エラー時は_ERROR.txt出力
- ログ出力あり
"""

import re
import traceback
import sys
from pathlib import Path
from datetime import datetime

# 終了コード（呼び出し元へ通知する契約）
EXIT_OK = 0  # 正常終了（全ファイル成功、警告相当なし）
EXIT_ERROR = 1  # 致命的エラー（環境不備などにより処理継続不可能）
EXIT_WARNING = 2  # 完走したが問題あり（警告、または _ERROR.txt 出力を伴うファイル単位失敗を含む）

# Word内のマーカー
MARKER_PARENT = "[PARENT]"

# ディレクトリ設定（プロジェクトルートからの相対パス）
# このスクリプトの位置: (project_root)/src/word/structured_to_department.py
# プロジェクトルート = このスクリプトの2階層上
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
INPUT_DIR = PROJECT_ROOT / "data" / "05_before_department"  # 入力ディレクトリ
DONE_DIR = INPUT_DIR / "done"
OUTPUT_DIR = PROJECT_ROOT / "data" / "06_after_department"  # 出力ディレクトリ
LOG_DIR = PROJECT_ROOT / "logs"
FILE_PATTERN = "*.txt"
LINE_BREAK = "\n"  # 改行文字（出力用）

# グローバルログファイルハンドラ
log_file = None

def log(message, also_print=False):
    """ログメッセージをファイルに書き込む
    
    Args:
        message (str): ログメッセージ
        also_print (bool): コンソールにも出力するか
    """
    global log_file
    if log_file:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file.write(f"[{timestamp}] {message}{LINE_BREAK}")
        log_file.flush()
    if also_print:
        print(message)

had_warning = False      # 要素レベルのスキップ等
had_file_error = False   # ファイル単位の失敗（_ERROR.txt になるもの等）

def notify_warning(file_path: str, message: str):
    """
    要素レベルのワーニング（段落/表/テキストボックス等の部分的エラー）を通知する。

    - 処理は継続可能だが、当該要素はスキップされ出力結果が一部欠落する可能性がある。
    - 上位プロセスが機械的に検知できるよう、stderr に `WARNING:` で出力する。

    Args:
        file_path: 対象Wordファイルのパス（不明な場合は None/空文字列でも可）
        message: ワーニング内容（簡潔な要約）
    """ 
    global had_warning
    had_warning = True
    name = Path(file_path).name if file_path else "-"
    print(f"WARNING: {name}: {message}", file=sys.stderr)

def notify_file_error(file_path: str, message: str):
    """
    ファイル単位の失敗（当該ファイルが処理できず _ERROR.txt を出力する等）を通知する。

    - 当該ファイルは失敗扱いだが、全体処理は継続する。
    - 上位プロセスが機械的に検知できるよう、stderr に `ERROR:` で出力する。

    Args:
        file_path: 対象Wordファイルのパス
        message: エラー内容（簡潔な要約）
    """
    global had_file_error
    had_file_error = True
    name = Path(file_path).name if file_path else "-"
    print(f"ERROR: {name}: {message}", file=sys.stderr)

def notify_fatal(message: str):
    """
    致命的エラー（環境不備などで処理継続不能）を通知する。

    - 上位プロセスが機械的に検知できるよう、stderr に `FATAL:` で出力する。
    - 本関数は終了処理を行わない。呼び出し側が終了コード（例: 1）で終了することを想定する。

    Args:
        message: 致命的エラー内容（簡潔な要約）
    """
    print(f"FATAL: {message}", file=sys.stderr)

def extract_department(section_lines):
    """セクションから部署名を抽出する
    Args:
        section_lines (list of str): セクションの行リスト
        Returns:
        str or None: 抽出した部署名、見つからなければNone
    """
    for line in section_lines:
        m = re.match(r"- 部署:(.*)", line)
        if m:
            return m.group(1).strip()
    return None

def split_sections(lines) -> list[list[str]]:
    """
    行リストを[PARENT]セクションごとに分割する
    
    Args:
        lines (list of str): ファイル全体の行リスト
    
    Returns: list of list of str: セクションごとの行リストのリスト
    """
    sections = []
    current = []
    for line in lines:
        if line.strip() == MARKER_PARENT:
            if current:
                # 空セクション（全て空行や空文字列のみ）は除外
                if any(l.strip() != '' for l in current):
                    sections.append(current)
            current = [line]
        else:
            current.append(line)
    if current:
        if any(l.strip() != '' for l in current):
            sections.append(current)
    return sections

def process_single_file(input_path, base_filename, timestamp) -> tuple[bool, int, str]:
    """
    単一ファイルを処理し、部署ごとに分割出力する

    Args:
        input_path (Path): 入力ファイルのパス
        base_filename (str): 元のファイル名（拡張子なし）
        timestamp (str): タイムスタンプ文字列（yyyymmddhhmmss形式）
    
    Returns:
        tuple: (成功したか, PARENTセクション数, エラーメッセージ)
    """
    try:
        # ファイル読み込み
        with open(input_path, encoding='utf-8') as f:
            lines = [line.rstrip(LINE_BREAK) for line in f]

        # セクション分割
        sections: list[list[str]] = split_sections(lines)
        # 各セクションの先頭・末尾の空行を検査し、あればすべて削除
        for i, sec in enumerate(sections):
            while sec and sec[0].strip() == '':
                sec.pop(0)
            while sec and sec[-1].strip() == '':
                sec.pop()
            sections[i] = sec

        # ファイル内から日付を抽出（最初に現れる「- 日付:」）
        file_date = None
        date_pattern = re.compile(r"- 日付:\s*(.+)")
        for line in lines:
            m = date_pattern.match(line)
            if m:
                file_date = m.group(1).strip()
                break
        if not file_date:
            file_date = timestamp  # 日付が見つからない場合はタイムスタンプを使う

        # 部署ごとに正常・エラーセクションを分類
        dept_dict: dict[str, list[list[str]]] = {}
        dept_error_dict: dict[str, list[tuple[list[str], str]]] = {}
        error_sections = 0
        for sec in sections:
            try:
                dept = extract_department(sec)
                if not dept:
                    dept = "（部署名なし）"
                if dept not in dept_dict:
                    dept_dict[dept] = []
                dept_dict[dept].append(sec)
            except Exception as e:
                # 部署名が抽出できない場合も含め、エラーセクションとして記録
                error_sections += 1
                # 例外発生時は部署名不明扱い
                dept = "（部署名なし）"
                if dept not in dept_error_dict:
                    dept_error_dict[dept] = []
                dept_error_dict[dept].append((sec, f"{e}{LINE_BREAK}{traceback.format_exc()}"))
                # エラー発生セクションの全行をログ出力
                log(f"  セクションエラー: {e}{LINE_BREAK}{traceback.format_exc()}")
                for i, line in enumerate(sec):
                    log(f"    エラー発生セクション行[{i+1}]: {line}")
                continue

        # 出力ディレクトリ（06_after_department/yyyymmddhhmmss）を作成
        output_dir = OUTPUT_DIR / timestamp
        output_dir.mkdir(parents=True, exist_ok=True)

        # ファイル名用に日付をYYYYMMDD形式でゼロ埋め
        date_for_filename = None
        m = re.match(r"(\d{4})年(\d{1,2})月(\d{1,2})日", file_date)
        if m:
            year = m.group(1)
            month = m.group(2).zfill(2)
            day = m.group(3).zfill(2)
            date_for_filename = f"{year}{month}{day}"
        else:
            # フォーマット外は数字のみ抽出（従来通り）
            import re as _re
            date_for_filename = _re.sub(r'[^0-9]', '', file_date)
            if not date_for_filename:
                date_for_filename = file_date  # 万一数字がなければそのまま

        # 正常セクション出力
        for dept, sec_list in dept_dict.items():
            # 部署名なしのセクションが1つもなければファイル出力しない
            if dept == "（部署名なし）" and not sec_list:
                continue
            if not sec_list:
                continue
            # 部署名なしのセクションが1つもなければファイル出力しない
            if dept == "（部署名なし）" and all(len(s) == 0 for s in sec_list):
                continue
            output_lines = ['', '']
            for idx, sec in enumerate(sec_list):
                output_lines.extend(sec)
                if idx < len(sec_list) - 1:
                    output_lines.append('')
            out_name = f"{date_for_filename}_{dept}.txt"
            out_path = output_dir / out_name
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write(LINE_BREAK.join(output_lines))

        # エラーセクション出力（部署単位）
        for dept, err_sec_list in dept_error_dict.items():
            if not err_sec_list:
                continue
            output_lines = ['', '']
            for idx, (sec, err_msg) in enumerate(err_sec_list):
                output_lines.append('【このセクションはエラーのため出力されました】')
                output_lines.extend(sec)
                output_lines.append('')
                output_lines.append('【エラー内容】')
                output_lines.append(err_msg)
                if idx < len(err_sec_list) - 1:
                    output_lines.append('')
            out_name = f"{date_for_filename}_{dept}_ERROR.txt"
            out_path = output_dir / out_name

            # 出力ファイル書き込み（部署ごと）
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write(LINE_BREAK.join(output_lines))

        return True, len(sections), None

    except Exception as e:
        return (False, 0, str(e))

def main():
    """
    複数のtextファイルを一括処理するメイン関数
    """
    global log_file

    # 処理開始時刻（全ファイル共通）
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")

    # ログディレクトリを作成
    log_dir = Path(LOG_DIR)
    log_dir.mkdir(parents=True, exist_ok=True)

    # ログファイルを開く
    log_path = log_dir / f"depart_{timestamp}.log"
    log_file = open(log_path, 'w', encoding='utf-8')

    try:
        log("="*70)
        log("部署ごとにファイル分割構処理開始")
        log(f"処理開始時刻: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
        log("="*70)
        
        # 入力・出力・処理済みディレクトリのパス
        input_dir = Path(INPUT_DIR)
        output_dir = Path(OUTPUT_DIR)
        done_dir = Path(DONE_DIR)
        done_dir.mkdir(parents=True, exist_ok=True)
        
        log(f"入力ディレクトリ: {input_dir}")
        log(f"出力ディレクトリ: {output_dir}")
        log(f"ログディレクトリ: {log_dir}")
        log("")

        # 入力ディレクトリの存在確認
        if not input_dir.exists():
            error_msg = f"入力ディレクトリが存在しません: {input_dir}"
            log(error_msg, also_print=True)
            log("ディレクトリを作成してテキストファイルを配置してください。", also_print=True)
            return
        
        # 出力ディレクトリを作成
        output_dir.mkdir(parents=True, exist_ok=True)

        # 処理対象ファイルを取得
        all_files = list(input_dir.glob(FILE_PATTERN))        

        log(f"処理対象ファイル数: {len(all_files)}件")

        if not all_files:
            error_msg = f"処理対象ファイルが見つかりません: {input_dir / FILE_PATTERN}"
            log(error_msg, also_print=True)
            return

        print("="*70)

       # 処理結果を集計
        success_count = 0
        error_count = 0
        total_parent_sections = 0

        log("")
        log("ファイル処理開始")
        log("="*70)

        # 各ファイルを処理
        # ※1始まりの意味：画面表示用の通し番号
        for idx, input_file in enumerate(all_files, 1):
            log("")
            log(f"[{idx}/{len(all_files)}] 処理中: {input_file.name}")
            print(f"[{idx}/{len(all_files)}] 処理中: {input_file.name}")

            # 仮の出力ファイル名を生成（エラー時に変更される可能性あり）
            output_filename = f"{input_file.stem}_{timestamp}.txt"
            output_path = output_dir / output_filename            

            log(f"  入力ファイル: {input_file}")
            log(f"  出力ファイル: {output_path}")

            # 各ファイル処理
            success, parent_count, err_msg = process_single_file(str(input_file), str(output_path), timestamp)

            if success:
                log(f"  結果: 成功")
                log(f"  [PARENT]セクション数: {parent_count}件")

                success_count += 1
                total_parent_sections += parent_count

                # 成功時は入力ファイルをdoneディレクトリへ移動
                moved_path = done_dir / input_file.name
                try:
                    input_file.rename(moved_path)
                    log(f"  入力ファイルを移動: {input_file} -> {moved_path}")
                except Exception as move_err:
                    log(f"  入力ファイル移動失敗: {move_err}")
            else:
                # エラー時はファイル名を変更
                error_filename = f"{input_file.stem}_{timestamp}_depart_ERROR.txt"
                error_path = output_dir / error_filename

                # 既に作成されているファイルがあればリネーム
                if output_path.exists():
                    output_path.rename(error_path)
                    log(f"  ファイルをリネーム: {output_filename} -> {error_filename}")

                log(f"  結果: エラー")
                log(f"  エラー内容: {error_msg}")
                log(f"  エラーファイル: {error_filename}")

                print(f"  ✗ エラー: {error_msg}")
                print(f"     エラーファイル: {error_filename}")
                error_count += 1

        # 最終結果サマリー
        log("")
        log("="*70)
        log("処理完了")
        log(f"処理終了時刻: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
        log(f"成功: {success_count}件")
        log(f"エラー: {error_count}件")
        log(f"総[PARENT]セクション数: {total_parent_sections}件")
        log(f"出力先: {output_dir.resolve()}")
        log(f"ログファイル: {log_path.resolve()}")
        log("="*70)
        
    finally:
        if log_file:
            log_file.close()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"予期しないエラーが発生しました: {e}")
        print(f"{LINE_BREAK}--- スタックトレース ---")
        traceback.print_exc()
        sys.exit(1)
