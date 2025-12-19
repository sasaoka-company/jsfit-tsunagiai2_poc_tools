# README_SHORT.md（第三者向け・機械的手順）

## 0. 事前準備（初回のみ）
### 0-1. uv をインストール
#### 方法A（推奨）
~~~powershell
winget install --id=astral-sh.uv -e
~~~

#### 方法B
~~~powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
~~~

### 0-2. uv が動くことを確認
（新しい PowerShell / cmd を開いてから実行）
~~~powershell
uv --version
~~~

---

## 1. 実行手順（毎回）
### 1-1. 入力を配置
- `data/01_input/` に `*.docx` を置く

### 1-2. 実行
~~~bat
scripts\run.cmd
~~~

---

## 2. 成功/失敗の確認（毎回）
### 2-1. 実行ログを確認
1) `logs/` を開く  
2) `run_all_*.log` があれば **最新ファイル**を開く（無ければ次へ）  
3) 次のログが「最新の実行分」として作成されているか確認する（同名が複数ある場合は更新日時が新しいもの）
- `process_*.log`
- `struct_*.log`
- `depart_*.log`

### 2-2. エラー判定（ファイルの有無だけで判定）
以下の場所を順に確認し、`*_ERROR.txt` が **1つでもあれば失敗扱い（要対応）** とします。

#### (1) 02_output（Word→Text）
- `data/02_output/` と `data/02_output/done/` を開く
- `*_ERROR.txt` が無いことを確認する

#### (2) 04_after_structured（Text→Structured）
- `data/04_after_structured/` と `data/04_after_structured/done/` を開く
- `*_ERROR.txt` が無いことを確認する

#### (3) 06_after_department（最終出力）
1) `data/06_after_department/` を開く  
2) **最新のタイムスタンプフォルダ**（例：`20251219113501`）を開く  
3) `*_ERROR.txt` が無いことを確認する  

- いずれかに `*_ERROR.txt` があれば **失敗扱い（要対応）**

---

## 3. エラー時の手順（毎回）
### 3-1. run.cmd が途中で止まる／エラー表示が出る
- `logs/run_all_*.log` の最新を開く  
- そのログを添付して連絡する

### 3-2. `*_ERROR.txt` がある
1) `logs/run_all_*.log` があれば **最新**を添付する（無ければ次へ）  
2) `*_ERROR.txt` が出ている場所に応じて、次のログとファイルを揃えて連絡する  

- 02にある（`data/02_output/` または `data/02_output/done/`）  
  - `logs/process_*.log`（最新）  
  - `data/02_output/ または data/02_output/done/ にある *_ERROR.txt（該当ファイル）`

- 04にある（`data/04_after_structured/` または `data/04_after_structured/done/`）  
  - `logs/struct_*.log`（最新）  
  - `data/04_after_structured/ または data/04_after_structured/done/ にある *_ERROR.txt（該当ファイル）`

- 06にある（`data/06_after_department/<timestamp>/`）  
  - `logs/depart_*.log`（最新）  
  - `data/06_after_department/<timestamp>/ にある *_ERROR.txt（該当ファイル）`

---

## 4. PATH（uv が見つからない場合）
### 4-1. 確認（PowerShell）
~~~powershell
Get-Command uv
uv --version
~~~

### 4-2. 確認（cmd）
~~~bat
where uv
uv --version
~~~

### 4-3. PATH 追加（ユーザー環境変数）
ユーザーの Path に次を追加して、ターミナルを開き直す：
~~~text
%USERPROFILE%\.local\bin
~~~
