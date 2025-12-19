# jsfit-tsunagiai2-poc-tools（Windows 実行手順）

本ツールは、Word（.docx）からマーカー範囲を抽出し、構造化 → 部署別分割までを一括実行します。

---

## 1. クイックスタート（最短）
1) `data/01_input` に入力 Word（*.docx）を置く
2) `scripts\run.cmd` を実行する  
3) 出力は `data/06_after_department/<yyyymmddhhmmss>/` に生成されます  
4) `data/02_output/`（および `data/02_output/done/`）→ `data/04_after_structured/`（および `data/04_after_structured/done/`）→ `data/06_after_department/` 配下の「最新タイムスタンプフォルダ」の順に開き、`*_ERROR.txt` が無いことを確認してください（下記「成功確認」参照）

---

## 2. 前提条件
- Windows 10/11
- uv がインストールされていること（次章参照）
- 初回実行時はネットワーク接続が必要な場合があります（依存関係・Python取得のため）

---

## 3. uv のインストール（Windows）
以下のいずれかで uv をインストールしてください。

### 方法A: WinGet（推奨）
~~~powershell
winget install --id=astral-sh.uv -e
~~~

### 方法B: 公式 PowerShell インストーラ
~~~powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
~~~

インストール後、ターミナルを開き直して次を確認してください：
~~~powershell
uv --version
~~~

### uv が見つからない場合（PATH 設定）

`run.cmd` 実行時に「uv が見つからない」と表示される場合は、次を確認してください。

#### 1) ターミナルを開き直す（最優先）
uv をインストールした直後は、既に開いている PowerShell / コマンドプロンプトには PATH が反映されません。  
**いったんすべてのターミナルを閉じて、新しく開き直してから**以下を実行してください。

~~~powershell
uv --version
~~~

---

#### 2) PATH が通っているか確認する
PowerShell の場合：

~~~powershell
Get-Command uv
uv --version
~~~

コマンドプロンプト（cmd）の場合：

~~~bat
where uv
uv --version
~~~

`uv --version` が通らない場合は、次の「PATH 追加」を行ってください。

---

#### 3) Windows（GUI）で PATH を追加する（ユーザー環境変数）
一般的なインストール先として、次のフォルダをユーザーの Path に追加します。

- `%USERPROFILE%\.local\bin`  
  例：`C:\Users\<ユーザー名>\.local\bin`

##### 追加手順（Windows 10/11）
1. **スタートメニュー**を開き、`環境変数` と入力  
2. 表示される **「システム環境変数の編集」** をクリック  
3. 表示されたウィンドウ（「システムのプロパティ」）で **「環境変数(N)...」** をクリック  
4. 上側の枠 **「ユーザー環境変数（<ユーザー名>）」** から **「Path」** を選択し、**「編集(E)...」** をクリック  
5. **「新規(N)」** をクリックし、次を追加  
~~~
%USERPROFILE%\.local\bin
~~~
6. すべてのウィンドウで **「OK」** をクリックして閉じる  
7. **PowerShell / cmd を開き直して**、次で確認

~~~powershell
uv --version
~~~

---

#### 4) それでも解決しない場合
uv は環境により別の場所にインストールされることがあります。  
（注）本ツールに同梱の `scripts/run.cmd` は、uv が PATH に無い場合でも代表的な場所を探索して実行を試みます。  
それでも見つからない場合は、次を添えて問い合わせてください。

- `logs\run_all_*.log`（作成されている場合）
- `where uv`（cmd）または `Get-Command uv`（PowerShell）の結果
- uv のインストール方法（WinGet / PowerShell インストーラ 等）

---

## 4. 入力データ
- 入力フォルダ：`data/01_input/`
- 対象：`*.docx`
- `~$` で始まる一時ファイルは処理対象外です

---

## 5. 実行方法
### 推奨（バッチ）
~~~bat
scripts\run.cmd
~~~

---

## 6. 出力（最終成果物）
最終成果物は次の配下に出力されます：

- `data/06_after_department/<yyyymmddhhmmss>/`

出力例：
- `data/06_after_department/20251219113501/20250701_総務部.txt`

※ `<yyyymmddhhmmss>` は実行時刻のタイムスタンプです。

---

## 7. 中間ファイルの扱い（自動）
本ツールは工程間の受け渡しを自動で行います。

- `data/02_output/*.txt` は `data/03_before_structured/` にコピーされ、コピー後 `data/02_output/done/` に移動されます  
  - `03_before_structured` 側に同名ファイルがある場合は上書きされます
- `data/04_after_structured/*.txt` は `data/05_before_department/` にコピーされ、コピー後 `data/04_after_structured/done/` に移動されます  
  - `05_before_department` 側に同名ファイルがある場合は上書きされます

---

## 8. 成功確認（重要）
本ツールのエラーは一律ではなく、以下の3種類があります。  
「実行が最後まで完了したか」と「成果物が完全か」は別物なので、必ず手順に沿って確認してください。

### 8.1 エラーの種類（重要）
#### (1) 要素レベルのエラー（継続）
- 対象：段落／表／テキストボックス／セクション解析などの一部要素
- 動作：**該当要素のみスキップして継続**（同じファイル内の次要素、または次ファイルへ進む）
- 記録：ログに詳細（エラー箇所、内容、スタックトレース）

#### (2) ファイル単位のエラー（そのファイルは中断して継続）
- 対象：入力ファイルの破損、読み込み失敗、マーカー構造の致命的不整合など
- 動作：**当該ファイルの処理を中断し、`*_ERROR.txt` を出力して次のファイルへ進む**
- 認識：*_ERROR.txt は「当該工程で正常に処理できなかった内容がある」印です（02/04では概ね“ファイル単位の中断”を意味します。06では“部署単位の分離出力”の場合があります）

#### (3) スクリプト単位の失敗（パイプライン中断）
- 対象：スクリプト自体が異常終了（return code ≠ 0）
- 動作：`run_all.py` が **そこで処理を中断**します（以降の工程は実行されません）
- 認識：コンソールに「[エラー] xxx.py の実行に失敗しました。処理を中断します。」が表示されます

---

### 8.2 実行（ジョブ）が完了したことの確認
- コンソールに「全処理が正常に完了しました。」が表示される  
  ※これは「全スクリプトが正常終了した」ことを意味します
- `logs/` 配下にログが生成されている（工程別ログ：process/struct/depart）

---

### 8.3 成果物が完全であることの確認（推奨）
最終成果物の完全性は、次の手順で確認します。

1) `data/06_after_department/<timestamp>/` に出力ファイルが生成されている  
2) 同フォルダ内に `*_ERROR.txt` が存在しない  
3) 上流工程の `_ERROR.txt` も確認する（02/04 で `_ERROR.txt` がある場合、当該ファイルは後続工程に進みません）  
   - `data/02_output/` **または** `data/02_output/done/`
   - `data/04_after_structured/` **または** `data/04_after_structured/done/`

補足：
- `structured_to_department.py` では、セクション解析等で問題が発生した場合に「部署単位の `_ERROR.txt`」が生成され得ます。  
  この場合、同じ部署について通常ファイルと `_ERROR.txt` が併存することがあります（=部分成功）。  
  併存時は、通常ファイルに含まれたセクションは正常抽出できていますが、`_ERROR.txt` 側に回ったセクションは欠落の可能性があります。

※ *_ERROR.txt が無い場合でも、要素レベルのエラーは発生し得ます。品質確認が必要な場合は 9.3（ログ確認）も実施してください。

---

## 9. エラー時の確認手順（整理版）

### 9.1 パイプラインが途中で止まった（スクリプト単位の失敗）
- 画面に「[エラー] xxx.py の実行に失敗しました。処理を中断します。」が出ます
- 直前に実行していたスクリプト名（例：`text_to_structured.py`）を控え、
  `logs/` 配下の該当ログを確認してください
- （ある場合）`logs/run_all_*.log`（実行全体ログ）を添付して問い合わせると切り分けが速いです

---

### 9.2 `*_ERROR.txt` が出た（工程別の意味に従って確認）
`*_ERROR.txt` は「正常に処理できなかった内容がある」印です。  
ただし **工程により意味が異なります**（ファイル単位の中断／部署単位の分離出力）。  
`*_ERROR.txt` が出た場所に応じて、次を確認してください。

- `data/02_output/*_ERROR.txt`（**または `data/02_output/done/*_ERROR.txt`**）→ `logs/process_*.log` を確認  
  （Word→Text工程：当該 docx はファイル単位で中断）
- `data/04_after_structured/*_ERROR.txt`（**または `data/04_after_structured/done/*_ERROR.txt`**）→ `logs/struct_*.log` を確認  
  （構造化工程：バリデーション不合格・読み込み失敗等で当該 txt はファイル単位で中断）
- `data/06_after_department/<timestamp>/*_ERROR.txt` → `logs/depart_*.log` を確認  
  （部署分割工程：**部署単位で「エラーになったセクションのみ」を分離出力している可能性があります（=部分成功）**。  
   併存（通常ファイル＋_ERROR.txt）の場合は、ログでエラー対象セクションを確認してください。）

問い合わせ時に添付するもの（可能な範囲で）：
- 該当ログ（`logs/*.log`）
- 該当 `*_ERROR.txt`
- 入力ファイル（docx または txt）

---

### 9.3 `_ERROR.txt` は無いが、要素レベルエラーの有無を確認したい（ログ確認を先に実施）
出力ファイルの内容確認は件数・分量的に現実的ではないため、**最初にログでエラー／警告の有無を判定**します。

#### 手順（推奨：ログファースト）
1) 該当実行のログファイルを特定する  
  - **（推奨）`logs/run_all_<timestamp>.log` がある場合は、まずこの最新ファイルを開き、そこに出ている process/struct/depart のログ名（timestamp）を辿ってください**  
  - `logs/process_<timestamp>.log`（Word→Text）  
  - `logs/struct_<timestamp>.log`（Text→Structured）  
  - `logs/depart_<timestamp>.log`（Structured→Department）  
  - （ある場合）`logs/run_all_<timestamp>.log`（実行全体ログ）

2) 各ログでエラー／警告を検索する（上から順に）
- 検索キーワード例：`ERROR` / `Exception` / `Traceback` / `WARNING` / `警告`

3) エラー／警告が見つかった場合の扱い
- **ERROR / Traceback / Exception がある**  
  - 要素レベルエラー（段落／表／テキストボックス／セクション解析等）の可能性があります  
  - ログ内の「対象ファイル名」「対象要素」「スタックトレース」を確認し、必要に応じて入力（docx/txt）やマーキングを修正して再実行してください
- **WARNING / 警告のみがある**  
  - 例：部署名 `【…】` の抽出不能など、処理は継続するが一部情報が空になる可能性があります  
  - 運用上の許容可否に応じて再実行／修正を判断してください
- **エラー／警告が見つからない**  
  - ログ上は正常に処理が進行した可能性が高いです（完全性を保証するものではありません）

#### ログ確認が難しい場合（補助）
- まず `logs/run_all_<timestamp>.log`（実行全体ログ）がある場合は、ここで一括検索すると効率的です  
- それでも判断が難しい場合は、問い合わせ時に以下を添付してください：  
  - `logs/run_all_<timestamp>.log`（あれば）  
  - `logs/process_<timestamp>.log` / `logs/struct_<timestamp>.log` / `logs/depart_<timestamp>.log`  
  - 該当する `_ERROR.txt`（存在する場合）

---

## 10. ログ出力先
- Word→Text：`logs/process_<timestamp>.log`
- Text→Structured：`logs/struct_<timestamp>.log`
- Structured→Department：`logs/depart_<timestamp>.log`

（推奨）実行全体ログ：`logs/run_all_<timestamp>.log`
