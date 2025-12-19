"""
move_structured_to_department.py

data/04_after_structured フォルダ内の全ファイル（*.txt）を data/05_before_structured にコピーし、
コピー後は data/04_after_structured/done に移動するスクリプト。

- コピーは上書き（同名ファイルがあれば before_structured 側を上書き）
- 移動先ディレクトリがなければ自動作成
"""

import shutil
from pathlib import Path

# ディレクトリ設定
PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
INPUT_DIR = PROJECT_ROOT / "data" / "04_after_structured"
OUTPUT_DIR = PROJECT_ROOT / "data" / "05_before_department"
DONE_DIR = INPUT_DIR / "done"

# ディレクトリ作成
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
DONE_DIR.mkdir(parents=True, exist_ok=True)

# 対象ファイル取得
files = list(INPUT_DIR.glob("*.txt"))

for file in files:
    dest = OUTPUT_DIR / file.name
    shutil.copy2(file, dest)
    print(f"コピー: {file} -> {dest}")
    moved = DONE_DIR / file.name
    shutil.move(str(file), str(moved))
    print(f"移動: {file} -> {moved}")

print("完了")
