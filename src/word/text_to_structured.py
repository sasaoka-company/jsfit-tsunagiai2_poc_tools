"""
text_to_structured.py 

目的:
  word_to_text.py の出力（マーカー付きテキスト）を、指定フォーマットの構造化テキストに変換する。
  親セクション（[PARENT]）・子セクション（[CHILD]）で分割し、質問・回答境界（[Q][A]）を付与し、
  主要問答と更問に整理して出力する。

入出力:
  入力:  (プロジェクトルート)/data/03_before_structured/*.txt
        入力には [PARENT][CHILD][QA_SPLIT] が含まれる（[Q][A] は本スクリプトが挿入）
  出力:  (プロジェクトルート)/data/04_after_structured/
        成功: {元のファイル名}_structured_{yyyymmddhhmmss}.txt
        エラー: {元のファイル名}_structured_{yyyymmddhhmmss}_ERROR.txt
  ログ:  (プロジェクトルート)/logs/struct_{yyyymmddhhmmss}.log
  ※複数ファイル一括処理、タイムスタンプは全ファイル共通
"""

import re
from datetime import datetime
from pathlib import Path
import traceback
import sys
from typing import Optional

# 終了コード（呼び出し元へ通知する契約）
EXIT_OK = 0  # 正常終了（全ファイル成功、警告相当なし）
EXIT_ERROR = 1  # 致命的エラー（環境不備などにより処理継続不可能）
EXIT_WARNING = 2  # 完走したが問題あり（警告、または _ERROR.txt 出力を伴うファイル単位失敗を含む）

# Word内のマーカー
MARKER_PARENT = "[PARENT]"
MARKER_CHILD = "[CHILD]"
MARKER_QA_SPLIT = "[QA_SPLIT]"

TAG_Q = "[Q] "  # 末尾に半角スペース
TAG_A = "[A] "  # 末尾に半角スペース
LINE_BREAK = "\n"  # 改行文字（出力用）

# ディレクトリ設定（プロジェクトルートからの相対パス）
# このスクリプトの位置: (project_root)/src/word/text_to_structured.py
# プロジェクトルート = このスクリプトの3階層上
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
INPUT_DIR = PROJECT_ROOT / "data" / "03_before_structured"
DONE_DIR = INPUT_DIR / "done"
OUTPUT_DIR = PROJECT_ROOT / "data" / "04_after_structured"
LOG_DIR = PROJECT_ROOT / "logs"
FILE_PATTERN = "*.txt"

# グローバルログファイルハンドラ
log_file = None


def log(message: str, also_print: bool = False) -> None:
    """ログメッセージをファイルに書き込む（必要に応じてコンソールにも出力）"""
    global log_file
    if log_file:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file.write(f"[{timestamp}] {message}{LINE_BREAK}")
        log_file.flush()
    if also_print:
        print(message)

had_warning = False      # 要素レベルのスキップ等
had_file_error = False   # ファイル単位の失敗（_ERROR.txt になるもの等）

def notify_warning(file_path: Optional[str], message: str):
    """
    要素レベルのワーニングを通知する。

    - 処理は継続可能だが、当該要素はスキップされ出力結果が一部欠落する可能性がある。
    - 上位プロセスが機械的に検知できるよう、stderr に `WARNING:` で出力する。

    Args:
        file_path: 対象ファイルのパス（不明な場合は None/空文字列でも可）
        message: ワーニング内容（簡潔な要約）
    """ 
    global had_warning
    had_warning = True
    name = Path(file_path).name if file_path else "-"
    print(f"WARNING: {name}: {message}", file=sys.stderr)

def notify_file_error(file_path: Optional[str], message: str):
    """
    ファイル単位の失敗（当該ファイルが処理できず _ERROR.txt を出力する等）を通知する。

    - 当該ファイルは失敗扱いだが、全体処理は継続する。
    - 上位プロセスが機械的に検知できるよう、stderr に `ERROR:` で出力する。

    Args:
        file_path: 対象ファイルのパス
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

def is_non_empty_line(line: str) -> bool:
    """非空行判定（空文字・空白のみを除外）"""
    return line is not None and line.strip() != ""


def next_non_empty_index(lines: list[str], start_index: int) -> int | None:
    """start_index 以降で最初に現れる非空行の index を返す（なければ None）"""
    for i in range(start_index, len(lines)):
        if is_non_empty_line(lines[i]):
            return i
    return None

def is_blank_line(line: str) -> bool:
    return line.strip() == ""


def trim_trailing_blank(lines: list[str]) -> list[str]:
    """末尾の空行を全て削除（内部の空行は保持）"""
    while lines and is_blank_line(lines[-1]):
        lines.pop()
    return lines


def trim_leading_blank(lines: list[str]) -> list[str]:
    """先頭の空行を全て削除（内部の空行は保持）"""
    while lines and is_blank_line(lines[0]):
        lines.pop(0)
    return lines


def normalize_blank_lines(lines: list[str]) -> list[str]:
    """
    連続する空行を1行に正規化する。
    ただし、ファイル先頭の空行2行は設計どおり維持する。
    """
    if len(lines) <= 2:
        return lines

    head = lines[:2]
    body = lines[2:]

    normalized: list[str] = []
    prev_blank = False
    for line in body:
        blank = is_blank_line(line)
        if blank and prev_blank:
            continue
        normalized.append(line)
        prev_blank = blank

    return head + normalized

# ============================================================
# 1. ファイル名から会見日付を取得
# ============================================================
def date_from_filename(input_filename: str) -> tuple[str, str]:
    """
    入力ファイル名先頭の yymmdd を日付として扱い、以下を返す
      - 出力日付: YYYY-MM-DD
      - 問答ID用: YYYYMMDD
    仕様:
      - 先頭6文字が数字、7文字目が '_' を前提
      - yyyy = 2000 + yy 固定
    """
    m = re.match(r"^(\d{6})_", input_filename)
    if not m:
        raise ValueError(f"ファイル名が想定形式ではありません（先頭 yymmdd_ 必須）: {input_filename}")

    yymmdd = m.group(1)
    yy = int(yymmdd[0:2])
    mm = int(yymmdd[2:4])
    dd = int(yymmdd[4:6])

    yyyy = 2000 + yy
    # 日付妥当性チェック
    datetime(yyyy, mm, dd)

    out_date = f"{yyyy:04d}-{mm:02d}-{dd:02d}"
    id_date = f"{yyyy:04d}{mm:02d}{dd:02d}"
    return out_date, id_date


# ============================================================
# マーカー行単独の検証
# ============================================================
def validate_marker_lines_are_alone(lines: list[str]) -> list[str]:
    """
    [PARENT] / [CHILD] / [QA_SPLIT] が行単独であることを検証。
    条件違反は致命（ERROR出力対象）。
    """
    errors: list[str] = []
    markers = [MARKER_PARENT, MARKER_CHILD, MARKER_QA_SPLIT]
    for idx, line in enumerate(lines, start=1):
        for marker in markers:
            if marker in line and line.strip() != marker:
                errors.append(f"{idx}行目: マーカーが行単独ではありません: {line!r}")
    return errors


# ============================================================
# 2-3. 問答ブロック/子ブロックの分割
# ============================================================
def split_parent_blocks(lines: list[str]) -> list[tuple[int, int]]:
    """
    [PARENT] の出現でブロックを分割し、各ブロックの (start, end) を返す。
    start は [PARENT] 行の index、end はブロック終端（次PARENT直前 or 文書末尾の次）。
    """
    parent_indices = [i for i, line in enumerate(lines) if line.strip() == MARKER_PARENT]
    blocks: list[tuple[int, int]] = []
    for idx, p in enumerate(parent_indices):
        end = parent_indices[idx + 1] if idx + 1 < len(parent_indices) else len(lines)
        blocks.append((p, end))
    return blocks


def split_child_blocks(lines: list[str], parent_start: int, parent_end: int) -> list[tuple[int, int]]:
    """
    親ブロック内の [CHILD] の出現で子ブロックを分割し、各ブロックの (start, end) を返す。
    start は [CHILD] 行の index、end は次CHILD直前 or 親ブロック終端。
    """
    child_indices = [i for i in range(parent_start, parent_end) if lines[i].strip() == MARKER_CHILD]
    blocks: list[tuple[int, int]] = []
    for idx, c in enumerate(child_indices):
        end = child_indices[idx + 1] if idx + 1 < len(child_indices) else parent_end
        blocks.append((c, end))
    return blocks


# ============================================================
# 事前バリデーション（致命条件）
# ============================================================
def validate_file_structure(lines: list[str]) -> list[str]:
    """
    新設計「事前バリデーション（文書構造チェック）」の致命条件を検証し、
    失敗理由のリストを返す（空ならOK）。
    """
    errors: list[str] = []

    # 1) マーカー行単独
    marker_line_errors = validate_marker_lines_are_alone(lines)
    if marker_line_errors:
        for err in marker_line_errors:
            errors.append(err)
            # エラー発生行をログ出力
            try:
                idx = int(err.split('行目')[0])
                log(f"  バリデーションエラー発生行[{idx}]: {lines[idx-1]}")
            except Exception:
                pass

    # 2) [PARENT] が1つもない
    parent_blocks = split_parent_blocks(lines)
    if not parent_blocks:
        errors.append(f"{MARKER_PARENT} が1つもありません。")
        return errors  # 以降は前提崩れなので早期終了

    # 3) 各問答ブロックに [CHILD] が1つもない
    for p_idx, (p_start, p_end) in enumerate(parent_blocks, start=1):
        child_blocks = split_child_blocks(lines, p_start, p_end)
        if not child_blocks:
            errors.append(f"{p_idx}番目の問答ブロック: {MARKER_CHILD} が1つもありません。")
            # エラー発生ブロックの内容をログ出力
            log(f"  バリデーションエラー: {p_idx}番目の問答ブロック（{MARKER_CHILD}なし）")
            for idx2, line in enumerate(lines[p_start:p_end], start=1):
                log(f"    ブロック内行[{idx2}]: {line}")
            continue

        # 4) 各子ブロックに [QA_SPLIT] が必ず1回
        for c_idx, (c_start, c_end) in enumerate(child_blocks, start=1):
            qa_splits = [i for i in range(c_start, c_end) if lines[i].strip() == MARKER_QA_SPLIT]
            if len(qa_splits) == 0:
                errors.append(f"{p_idx}番目問答 / {c_idx}番目子ブロック: [QA_SPLIT] がありません。")
                continue
            if len(qa_splits) >= 2:
                errors.append(f"{p_idx}番目問答 / {c_idx}番目子ブロック: [QA_SPLIT] が複数存在します。")
                continue

            qa_idx = qa_splits[0]

            # 5) [CHILD] の次の非空行（質問部）が存在しない
            q_line_idx = next_non_empty_index(lines, c_start + 1)
            if q_line_idx is None or q_line_idx >= qa_idx:
                errors.append(
                    f"{p_idx}番目問答 / {c_idx}番目子ブロック: "
                    f"[CHILD] の次の非空行（質問起点）が存在しない、または [QA_SPLIT] より後です。"
                )
                continue

            # 6) [QA_SPLIT] の次の非空行（回答部）が存在しない
            a_line_idx = next_non_empty_index(lines, qa_idx + 1)
            if a_line_idx is None or a_line_idx >= c_end:
                errors.append(
                    f"{p_idx}番目問答 / {c_idx}番目子ブロック: "
                    f"[QA_SPLIT] の次の非空行（回答起点）が存在しません。"
                )
                continue

    return errors


# ============================================================
# 6. 部署名の抽出（警告で継続）
# ============================================================
def extract_department_from_major_question(question_lines: list[str]) -> str | None:
    for line in question_lines:
        m = re.search(r"【([^】]*)】", line)
        if m:
            return m.group(1)  # 空文字もあり得るが「見つかった」扱い
    return None

# ============================================================
# 7. 質問・回答の自動挿入（子ブロック）
# ============================================================
def build_child_output_lines(
    child_content_lines: list[str],
    is_major: bool,
    followup_number: int | None,
) -> list[str]:
    """
    子ブロックの内容（[CHILD] 行は含まない）から出力行を構築する。
      - [QA_SPLIT] を境界に質問/回答分割
      - [Q]/[A] を「次の非空行」に行頭付与
      - [QA_SPLIT] 行は出力に含めない
    """
    # [QA_SPLIT] の index（必ず1つであることは事前バリデーションで保証）
    try:
        qa_rel = next(i for i, line in enumerate(child_content_lines) if line.strip() == MARKER_QA_SPLIT)

        question_part = child_content_lines[:qa_rel]
        answer_part = child_content_lines[qa_rel + 1 :]

        # [QA_SPLIT] 前後の「余計な空行」を抑制（見た目安定化）
        question_part = trim_trailing_blank(question_part)
        answer_part = trim_leading_blank(answer_part)

        # [Q] を付与する行（質問部の次の非空行）
        q_rel_idx = next_non_empty_index(question_part, 0)
        if q_rel_idx is None:
            # 事前バリデーションで弾かれている想定だが、保険
            raise ValueError("質問部に非空行がありません。")

        if not question_part[q_rel_idx].startswith(TAG_Q):
            question_part[q_rel_idx] = TAG_Q + question_part[q_rel_idx]

        # [A] を付与する行（回答部の次の非空行）
        a_rel_idx = next_non_empty_index(answer_part, 0)
        if a_rel_idx is None:
            raise ValueError("回答部に非空行がありません。")

        if not answer_part[a_rel_idx].startswith(TAG_A):
            answer_part[a_rel_idx] = TAG_A + answer_part[a_rel_idx]

        # 見出し
        title_lines: list[str] = []
        if is_major:
            title_lines.append("## 主要問答")
        else:
            title_lines.append(f"## 更問{followup_number}")

        # 出力組み立て
        out: list[str] = []
        out.append(MARKER_CHILD)
        out.extend(title_lines)
        out.extend(question_part)
        # out.append("")  # [Q]と[A]の間に空行を1行挿入
        out.extend(answer_part)

        # 最終チェック（出力上の [Q]/[A] が1回ずつであること）
        q_count = sum(1 for line in out if line.startswith(TAG_Q))
        a_count = sum(1 for line in out if line.startswith(TAG_A))
        if q_count != 1 or a_count != 1:
            raise ValueError(f"子ブロックの [Q]/[A] 出現数が不正です（[Q]={q_count}, [A]={a_count}）。")

        # [Q] が [A] より先に出ること
        q_pos = next(i for i, line in enumerate(out) if line.startswith(TAG_Q))
        a_pos = next(i for i, line in enumerate(out) if line.startswith(TAG_A))
        if q_pos >= a_pos:
            raise ValueError("子ブロックの [Q] が [A] より後に出現しています。")

        # 子ブロック末尾の空行は削除（ブロック区切りは呼び出し側で付与）
        out = trim_trailing_blank(out)
        return out
    except Exception as e:
        # エラー発生時、ブロック内容をログ出力
        log(f"  build_child_output_linesエラー: {e}")
        for idx, line in enumerate(child_content_lines, start=1):
            log(f"    エラー発生CHILDブロック行[{idx}]: {line}")
        raise

# ============================================================
# 1ファイル処理
# ============================================================
def process_single_file(input_path: Path, output_ok_path: Path, output_error_path: Path) -> tuple[bool, dict, str | None]:
    """
    1ファイルを新設計に従って変換。
    Returns:
      (成功したか, 統計情報dict, エラーメッセージ)
    """
    stats = {
        "parent_count": 0,
        "major_count": 0,
        "followup_count": 0,
        "warning_count": 0,
    }

    try:
        with open(input_path, encoding="utf-8") as f:
            lines = [line.rstrip(LINE_BREAK) for line in f]

        # --- 事前バリデーション（致命条件） ---
        fatal_errors = validate_file_structure(lines)
        if fatal_errors:
            # ファイル単位失敗として上位へ通知
            notify_file_error(str(input_path), "入力ファイルの構造が不正です（バリデーションエラー）")

            # ERRORファイル出力（最小限の内容）
            log(f"致命バリデーションエラー: {input_path.name}", also_print=True)

            for msg in fatal_errors:
                log(f"  - {msg}")
            out_err_lines = ["", "", "ERROR: 入力ファイルの構造が不正です。"] + [f"- {e}" for e in fatal_errors]
            with open(output_error_path, "w", encoding="utf-8") as f:
                f.write(LINE_BREAK.join(out_err_lines).rstrip(LINE_BREAK) + LINE_BREAK)
            return False, stats, " / ".join(fatal_errors)

        # --- 日付の取得（ファイル名由来） ---
        out_date, id_date = date_from_filename(input_path.name)

        # --- 親ブロック処理 ---
        output_lines: list[str] = ["", ""]  # ファイル先頭に空行2つ
        parent_blocks = split_parent_blocks(lines)

        # 問答ID連番（ファイルごとにリセット）
        qa_seq = 1

        for p_idx, (p_start, p_end) in enumerate(parent_blocks, start=1):
            stats["parent_count"] += 1
            stats["major_count"] += 1

            # ブロック間は空行で区切る（先頭は既に空行2つがあるため、ここでは2ブロック目以降に空行）
            if p_idx > 1:
                output_lines.append("")

            # 子ブロック抽出
            child_blocks = split_child_blocks(lines, p_start, p_end)
            if not child_blocks:
                raise ValueError(f"{p_idx}番目の問答ブロック: {MARKER_CHILD} が1つもありません。")

            # 主要問答（最初の子ブロック）
            major_child_start, major_child_end = child_blocks[0]
            major_child_content = lines[major_child_start + 1 : major_child_end]  # [CHILD] 行を除外
            qa_rel = next(i for i, line in enumerate(major_child_content) if line.strip() == MARKER_QA_SPLIT)
            major_question_raw = major_child_content[:qa_rel]  # [CHILD] 直下〜 [QA_SPLIT] 直前（仕様どおり）

            # 部署抽出（主要問答の質問範囲から）
            dept = extract_department_from_major_question(major_question_raw)
            if dept is None:
                dept = ""  # 出力は空値
            # 親チャンク出力
            qa_id = f"{id_date}-{qa_seq:03d}"
            output_lines.append(MARKER_PARENT)
            output_lines.append(f"# 問答ID: {qa_id}")
            output_lines.append(f"- 日付: {out_date}")
            output_lines.append(f"- 部署: {dept}")

            # 親チャンクの主要質問（主要問答の質問部をコピーし、先頭の非空行に [Q] を付与）
            parent_q_lines = list(major_question_raw)  # コピー

            # 親チャンクの主要質問末尾の空行を削除（区切り空行と二重になりやすいため）
            parent_q_lines = trim_trailing_blank(parent_q_lines)

            q_rel_idx = next_non_empty_index(parent_q_lines, 0)
            if q_rel_idx is None:
                # 事前バリデーションで弾かれている想定
                raise ValueError("主要質問（親チャンク）に非空行がありません。")
            if not parent_q_lines[q_rel_idx].startswith(TAG_Q):
                parent_q_lines[q_rel_idx] = TAG_Q + parent_q_lines[q_rel_idx]
            output_lines.extend(parent_q_lines)

            # 親チャンクと子ブロックは空行で区切る
            output_lines.append("")

            # 子ブロック出力（主要問答＋更問）
            followup_no = 0
            for c_idx, (c_start, c_end) in enumerate(child_blocks, start=1):
                child_content = lines[c_start + 1 : c_end]  # [CHILD] 行は含めない

                is_major = (c_idx == 1)
                if not is_major:
                    followup_no += 1
                    stats["followup_count"] += 1

                child_out = build_child_output_lines(
                    child_content_lines=child_content,
                    is_major=is_major,
                    followup_number=(None if is_major else followup_no),
                )
                output_lines.extend(child_out)

                # 子ブロック間は空行
                # （最後の子の後も空行を入れてよいが、過剰にならないようここで統一的に1行）
                if c_idx < len(child_blocks):
                    output_lines.append("")

            qa_seq += 1

        # 出力（成功）
        output_lines = normalize_blank_lines(output_lines)
        with open(output_ok_path, "w", encoding="utf-8") as f:
            f.write(LINE_BREAK.join(output_lines).rstrip(LINE_BREAK) + LINE_BREAK)

        return True, stats, None

    except Exception as e:
        # 例外はファイル全体の致命として ERROR 出力
        err_msg = str(e)

        # ファイル単位失敗として上位へ通知
        notify_file_error(str(input_path), f"実行時例外: {err_msg}")

        log(f"致命エラー: {input_path.name} / {err_msg}", also_print=True)
        log(traceback.format_exc())

        out_err_lines = ["", "", "ERROR: 実行時例外が発生しました。", f"- {err_msg}"]
        with open(output_error_path, "w", encoding="utf-8") as f:
            f.write(LINE_BREAK.join(out_err_lines).rstrip(LINE_BREAK) + LINE_BREAK)

        return False, stats, err_msg


def main() -> int:
    """複数のtextファイルを一括処理するメイン関数（新設計準拠）"""
    global log_file

    # 処理開始時刻（全ファイル共通）
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")

    # ログディレクトリ作成
    LOG_DIR.mkdir(parents=True, exist_ok=True)

    # ログファイルを開く
    log_path = LOG_DIR / f"struct_{timestamp}.log"
    log_file = open(log_path, "w", encoding="utf-8")

    try:
        log("=" * 70)
        log("構造化文書への変換処理開始（新設計準拠）")
        log(f"処理開始時刻: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
        log("=" * 70)

        INPUT_DIR.mkdir(parents=True, exist_ok=True)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        DONE_DIR.mkdir(parents=True, exist_ok=True)

        log(f"入力ディレクトリ: {INPUT_DIR}")
        log(f"出力ディレクトリ: {OUTPUT_DIR}")
        log(f"ログディレクトリ: {LOG_DIR}")
        log("")

        all_files = list(INPUT_DIR.glob(FILE_PATTERN))
        log(f"処理対象ファイル数: {len(all_files)}件", also_print=True)

        if not all_files:
            msg = f"処理対象ファイルが見つかりません: {INPUT_DIR / FILE_PATTERN}"
            log(msg, also_print=True)
            # 致命として通知
            notify_fatal(msg)

            return EXIT_ERROR

        # 集計
        success_count = 0
        error_count = 0
        total_parent = 0
        total_major = 0
        total_followup = 0
        total_warning = 0

        log("")
        log("ファイル処理開始")
        log("=" * 70)

        for idx, input_file in enumerate(all_files, start=1):
            log("")
            log(f"[{idx}/{len(all_files)}] 処理中: {input_file.name}", also_print=True)

            # 出力ファイル名（新仕様）
            ok_filename = f"{input_file.stem}_structured_{timestamp}.txt"
            err_filename = f"{input_file.stem}_structured_{timestamp}_ERROR.txt"
            ok_path = OUTPUT_DIR / ok_filename
            err_path = OUTPUT_DIR / err_filename

            log(f"  入力ファイル: {input_file}")
            log(f"  成功出力: {ok_path}")
            log(f"  ERROR出力: {err_path}")

            success, stats, err_msg = process_single_file(
                input_path=input_file,
                output_ok_path=ok_path,
                output_error_path=err_path,
            )

            if success:
                success_count += 1
                total_parent += stats["parent_count"]
                total_major += stats["major_count"]
                total_followup += stats["followup_count"]
                total_warning += stats["warning_count"]

                log(f"  結果: 成功")
                log(f"  総問答数: {stats['parent_count']}")
                log(f"  主要問答数: {stats['major_count']}")
                log(f"  更問数: {stats['followup_count']}")
                if stats["warning_count"] > 0:
                    log(f"  警告数: {stats['warning_count']}")

                # 成功時は入力ファイルを done へ移動
                moved_path = DONE_DIR / input_file.name
                try:
                    input_file.rename(moved_path)
                    log(f"  入力ファイルを移動: {input_file} -> {moved_path}")
                except Exception as move_err:
                    # 要素レベル警告として上位へ通知
                    notify_warning(str(input_file), f"done への移動失敗: {move_err}")

                    log(f"  入力ファイル移動失敗: {move_err}")
            else:
                error_count += 1
                log(f"  結果: ERROR")
                log(f"  エラー内容: {err_msg}")
                log(f"  エラーファイル: {err_filename}", also_print=True)

        # サマリー
        log("")
        log("=" * 70)
        log("処理完了")
        log(f"処理終了時刻: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
        log(f"成功: {success_count}件")
        log(f"ERROR: {error_count}件")
        log(f"総問答数: {total_parent}件")
        log(f"主要問答数: {total_major}件")
        log(f"更問数: {total_followup}件")
        if total_warning > 0:
            log(f"警告数: {total_warning}件")
        log(f"出力先: {OUTPUT_DIR.resolve()}")
        log(f"ログファイル: {log_path.resolve()}")
        log("=" * 70)

        # 完走後の終了コード集約
        if had_warning or had_file_error or error_count > 0:
            return EXIT_WARNING
        return EXIT_OK
    
    finally:
        if log_file:
            log_file.close()


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        msg = f"致命的エラーが発生しました: {e}"
        notify_fatal(msg)
        print(f"{LINE_BREAK}--- スタックトレース ---", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)
        sys.exit(EXIT_ERROR)
