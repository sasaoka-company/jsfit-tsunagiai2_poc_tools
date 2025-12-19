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

def run_script(script_name):
    script_path = SCRIPT_DIR / script_name
    print(f"\n=== {script_name} 実行開始 ===")
    result = subprocess.run([sys.executable, str(script_path)])
    if result.returncode != 0:
        print(f"\n[エラー] {script_name} の実行に失敗しました。処理を中断します。")
        sys.exit(result.returncode)
    print(f"=== {script_name} 実行完了 ===\n")

if __name__ == "__main__":
    for script in SCRIPTS:
        run_script(script)
    print("\n全処理が正常に完了しました。\n")
