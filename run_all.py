"""
run_all.py

word_to_text.py
→ move_output_to_structured.py
→ text_to_structured.py
→ move_structured_to_department.py
→ structured_to_department.py
の順で自動実行するバッチスクリプト。

- 各スクリプトのパスは src/word/ 配下を想定
- 各ステップでエラーがあれば中断し、メッセージを表示
- Python仮想環境が有効な状態で実行することを推奨
"""

import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent / "src" / "word"

SCRIPTS = [
    "word_to_text.py",
    "move_output_to_structured.py",
    "text_to_structured.py",
    "move_structured_to_department.py",
    "structured_to_department.py"
]

# 終了コード
EXIT_OK = 0  # 正常終了（全ファイル成功、警告相当なし）
EXIT_ERROR = 1  # 致命的エラー（環境不備などにより処理継続不可能）
EXIT_WARNING = 2  # 完走したが問題あり（警告、または _ERROR.txt 出力を伴うファイル単位失敗を含む）

def run_script(script_name):
    script_path = SCRIPT_DIR / script_name
    print(f"\n=== {script_name} 実行開始 ===")

    result = subprocess.run([sys.executable, str(script_path)])
    code = result.returncode

    # 0: 正常終了
    if code == EXIT_OK:
        print(f"=== {script_name} 実行完了 ===\n")
        return EXIT_OK
    
    # 2: 警告ありで完走（ここでは継続）
    if code == EXIT_WARNING:
        print(f"[警告] {script_name} は警告ありで完走しました（exit={EXIT_WARNING}）。処理を継続します。")
        print(f"=== {script_name} 実行完了（警告あり） ===\n")
        return EXIT_WARNING

    # 1 もしくは想定外: 中断
    print(f"\n[エラー] {script_name} の実行に失敗しました（exit={code}）。処理を中断します。")
    sys.exit(code)    

def main():
    had_warning = False

    for script in SCRIPTS:
        code = run_script(script)
        if code == EXIT_WARNING:
            had_warning = True
    if had_warning:
        print("\n全処理が完了しました（警告あり）。\n")
        return EXIT_WARNING
            
    print("\n全処理が正常に完了しました。\n")

if __name__ == "__main__":
    sys.exit(main())
