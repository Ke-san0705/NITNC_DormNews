# NITNC_DormNews

沼津高専寮生会新聞を作るための、**Microsoft Forms → 回答Excel → HTML / PDF / PNG** 自動組版システムです。A4縦・1ページに、3本の記事、16部署（うち三役は右欄）の報告、教員コメントを掲載します。

> **ファイル名はすべて英数字に統一しました。** 新聞の日本語本文や部署名は従来どおり利用できます。Windows上でZIPを展開しても、日本語名のファイルが文字化けしにくい構成です。

## 1. 最初に覚えるファイル名

| 用途 | ファイル名・場所 |
| --- | --- |
| 発行号・記事・写真・導入文・教員コメント | `issues/Vol06/edition.json` |
| Vol.6のMicrosoft Forms回答Excel | `issues/Vol06/answer.xlsx` |
| Vol.6の記事画像 | `issues/Vol06/assets/` |
| フォーム回答の集約と校正 | `generate.py` |
| 新号の準備 | `new_issue.py` |
| 新聞のHTML/CSSレイアウト | `templates/newspaper.html` |
| Windows向け簡易起動 | `start_windows.bat` |
| 回答の提出状況 | `output/Vol6/status.html` |

**旧名称との対応：** `回答.xlsx` → `answer.xlsx`、`classic` → `typeC`、`blue` → `typeB`。旧ファイル名・旧フォルダを使っていた場合は、後述の移行方法を確認してください。

## 2. 出力する新聞の名前

例えば**Vol.6**なら、`--design both`で以下のファイルが作成されます。

```text
NITNC_DormNews/
├── generate.py
├── new_issue.py
├── edition.json
├── requirements.txt
├── templates/
│   └── newspaper.html
├── issues/
│   └── Vol06/
│       ├── edition.json
│       ├── answer.xlsx                # Formsから取得したファイルを自分で配置
│       └── assets/
│           ├── Halloween.jpg
│           └── Halloween2.jpg
├── samples/
│   ├── demo_answer.xlsx
│   └── demo_answer.csv
└── output/
    └── Vol6/
        ├── typeB/
        │   ├── DormNewsVol6_typeB.pdf
        │   ├── DormNewsVol6_typeB.png
        │   └── DormNewsVol6_typeB.html
        ├── typeC/
        │   ├── DormNewsVol6_typeC.pdf
        │   ├── DormNewsVol6_typeC.png
        │   └── DormNewsVol6_typeC.html
        ├── status.html
        ├── status.json
        ├── merged_data.json
        └── error.txt                  # 失敗したときだけ
```

号数が7なら `DormNewsVol7_typeB.pdf` / `DormNewsVol7_typeC.pdf` のように自動変更されます。**typeBはカラフルな現行改良版、typeCは新聞風のデザイン**です。両方とも下部の発行情報はありません。typeCの上部にあった英語の題字も表示しません。


### 新レイアウトの余白とデザイン

- **typeB（現行改良版）**：タイトルは従来の青い塗りつぶし（`#1469b6`）に戻し、黄色い帯を背面に配置します。
- **typeC（新聞風）**：英語の題字は表示せず、落ち着いた新聞調のレイアウトを維持します。
- **共通**：記事①②は本文約4行・写真欄付き、記事③は左に文章・右に画像を配置。左下の発行者・号数表記を削除した分も本部部会へ使い、**各部署の報告欄は以前の高さ509px**に揃えています。
- 横長記事を追加したため、**各記事の領域内は従来よりコンパクト**です。長い記事本文・見出しは自動で削らず、はみ出し検査で停止します。

## 3. 初回セットアップ（Windows 11）

Python 3.11以降を用意して、`generate.py`のある**プロジェクトルート**でPowerShellを開きます。最初に次のコマンドを実行してください。

```powershell
python -m pip install -r requirements.txt
python -m playwright install chromium
```

インストール確認：

```powershell
python --version
python generate.py --demo --volume 5 --design both
```

`--demo`では個人情報を含まない `samples/demo_answer.xlsx` を使います。`--volume 5` とルートの `edition.json` の号数を一致させてください。デモ新聞も英数字のファイル名で出力されます。

## 4. Vol.6を作る（毎号の基本操作）

**① 記事を編集**　`issues/Vol06/edition.json` をVS Codeで開き、`volume=6`、発行年・月、導入文、記事1〜3、先生方のコメントを編集して保存します。

**② 画像を置く**　`issues/Vol06/assets/` に写真を保存し、各記事の `images` にその相対パスを書きます。

**③ FormsのExcelを置く**　その号のForms回答をExcel形式（`.xlsx`）で取得し、必ず**`issues/Vol06/answer.xlsx`** という名前で保存します。別の号のExcelを流用しないでください。

**④ 生成する**　プロジェクトルートで次を実行します。

```powershell
python generate.py --input "issues/Vol06/answer.xlsx" --edition "issues/Vol06/edition.json" --volume 6 --design both
```

**⑤ 確認する**　`output/Vol6/typeB/DormNewsVol6_typeB.pdf` と `output/Vol6/typeC/DormNewsVol6_typeC.pdf` を開き、両方の校正を行います。

## 5. Vol.7以降の新号を準備する

例：Vol.7を2026年12月に発行する場合。

```powershell
python new_issue.py --volume 7 --year 2026 --month 12
```

自動で `issues/Vol07/edition.json`、`issues/Vol07/assets/`、`issues/Vol07/PUT_answer_xlsx_HERE.txt` が作成されます。`edition.json` に**記事1・2・3の入力欄**が含まれます。同じコマンドを再実行しても既存の原稿と写真は上書きされません。

回答Excelを `issues/Vol07/answer.xlsx` に保存したら、次を実行します。

```powershell
python generate.py --input "issues/Vol07/answer.xlsx" --edition "issues/Vol07/edition.json" --volume 7 --design both
```

実行場所は `Vol07` 内ではなく **`NITNC_DormNews` の直下**です。

## 6. edition.jsonで編集できる内容

主なキーは次のとおりです。

| キー | 内容 |
| --- | --- |
| `volume` | 号数。例：`6` |
| `issue_year` / `issue_month` | 発行年・月。例：`2026` / `11` |
| `intro` | 題字の下の導入文 |
| `article1` | 左のメイン記事 |
| `article2` | 右のメイン記事 |
| `article3` | メイン2記事の下にある**横長記事** |
| `teacher1` | 寮務主事のコメント |
| `teacher2` | 寮監のコメント |

記事3の書き方（記事1・2も同じ`images`形式）：

```json
"article3": {
  "category": "CAMPUS NEWS / 03",
  "title": "次回イベントのお知らせ",
  "body": "記事の本文をここに記入します。",
  "date_label": "12 / 01",
  "images": ["assets/Event.jpg"]
}
```

記事3のレイアウトは**左上にタイトル、左下に本文、右側に画像**です。メイン記事2本の本文欄は約4行分を想定し、写真を大きく取っています。本文やタイトルが長すぎる場合は、文字が切れるのを防ぐため生成が停止することがあります。

### 画像0・1・2枚の指定方法

以下の配列だけを編集します。

```json
"images": []
```
画像0枚。`date_label`による日付カードを表示します。

```json
"images": ["assets/Halloween.jpg"]
```
画像1枚。写真枠に大きく表示されます。

```json
"images": ["assets/Halloween.jpg", "assets/Halloween2.jpg"]
```
画像2枚。写真枠内で左右に並びます。

指定パスは**その号の`edition.json`から見た相対パス**です。対応する画像形式はJPG / PNG / WebP（1枚8MB以下）。画像ファイル名も英数字・アンダースコアを推奨します。旧方式の `"image": "..."` も1枚指定なら使用可能ですが、**同じ記事で`image`と`images`の両方に画像を指定しないでください。**

`image_caption`は設定データとして保持できますが、現状は紙面への表示に対応していません。

## 7. Microsoft Formsの回答について

Formsの質問は**部署／①活動報告／②寮生への連絡事項／③その他**です。部署・①・②は必須、③は任意。部署は16種類あり、「秀峰寮東」「秀峰寮西」は別々に集計します。三役は右側の専用欄です。

各部署は何度でも回答可能で、**完了時刻が最新の1回分を採用**します。①②③は1セットとして上書きし、最新版の③が空欄なら旧回答の③は残しません。未提出の部署は未提出として扱います。

**掲載目安は各部署につき①活動報告30字・②連絡事項30字・③その他30字（合計90字）**です。文字数には①②③の番号を含みません。①②は必須、③は任意です。typeB・typeCとも報告欄の高さを以前の**509px**に復元し、本文の標準文字サイズも従来の**10px**に維持しました。16部署すべてが①②③に30字ずつ入力したサンプルで両デザインのA4・1ページ生成を確認しています。

ただし、**30字以内ならどんな文章でも必ず入るという保証はありません**。改行・英数字・記号の組み合わせで折り返しが変わり、必要に応じて文字サイズが縮小されます。紙面からはみ出す場合はPDFを出さずエラーにするため、発行前に実際のプレビューを確認してください。フォームには各設問の上限を30文字として設定することを推奨します。

実回答Excelには個人情報が含まれる場合があるため、**GitHubにアップロードしないでください**。

## 8. 出力方法・デザイン選択

| コマンド末尾 | 動作 |
| --- | --- |
| `--design both` | typeBとtypeCを両方生成 |
| `--design typeB` | typeBだけ生成 |
| `--design typeC` | typeCだけ生成 |
| `--watch` | Excel・edition.jsonのファイル更新を監視して再生成 |
| `--publish` | 未提出部署・文字数目安超過・仮コメントなどを厳しくチェック |
| `--demo` | `samples/demo_answer.xlsx` で生成テスト |

従来の`--design blue`・`--design classic`も互換性のため受け付けますが、**生成フォルダ名・ファイル名はtypeB / typeC**に統一されます。

OneDriveの同期Excelを使用する場合：

```powershell
python generate.py --input "issues/Vol06/answer.xlsx" --edition "issues/Vol06/edition.json" --volume 6 --design both --watch
```

`--watch`は、PCを起動している間だけ実行されます。停止するには`Ctrl + C`です。**Microsoft Formsにプログラムが直接ログインする仕組みではありません。**

Windowsでは、`start_windows.bat`をダブルクリックして号数を入力する方法も使えます。これは号別の`issues/VolXX/answer.xlsx`と`edition.json`を自動指定します。

## 9. 出力ファイルとエラーへの対処

出力は通常、`output/Vol6/`の下にまとまります。

| ファイル | 用途 |
| --- | --- |
| `typeB/DormNewsVol6_typeB.pdf` | typeBの完成PDF |
| `typeC/DormNewsVol6_typeC.pdf` | typeCの完成PDF |
| `typeB/DormNewsVol6_typeB.png` | typeBのプレビューPNG |
| `typeC/DormNewsVol6_typeC.png` | typeCのプレビューPNG |
| `status.html` | 16部署の提出状況（ブラウザで確認） |
| `status.json` | 処理状況や警告の詳細 |
| `merged_data.json` | 最新回答をまとめたデータ。寮内情報を含むので取扱注意 |
| `error.txt` | 生成失敗時のエラーメッセージ |

**よくある問題：**

- `answer.xlsx`が見つからない：`issues/VolXX/`内に存在するか確認してください。
- `edition.json`の号数と`--volume`が異なる：両方を同じ数字にしてください。
- Excelの日時が読めない：最新の`report_engine.py`を使用してください。
- PDFが生成されない：`error.txt`や`status.json`を読み、はみ出した記事・部署の文章を短くしてください。
- 写真が見つからない：画像を`issues/VolXX/assets/`に置き、`images`のスペルを確認してください。
- ZIP展開後の旧日本語ファイルが残る：次項の移行スクリプトを実行してください。

改修を検証する場合：

```powershell
python -m unittest test_system.py -v
```

## 10. 既存プロジェクトから移行する場合

**最初に、`issues`フォルダを含むプロジェクト全体をバックアップしてください。** 旧ファイルの名前を変えてから今回のプログラム一式を上書きする運用が安全です。

対応スクリプト`migrate_filenames.py`を追加しました。これは既存の`issues/VolXX/回答.xlsx`を`answer.xlsx`へ、古い生成ファイル名を英語に変更するためのものです。

```powershell
python migrate_filenames.py
```

すでに同じ名前のファイルがある場合は**上書きせずに警告**します。移行後は新しい`generate.py`で新聞を再生成することをおすすめします。旧ファイル名がGitに登録されている場合は、`git add -A`で削除と追加を同時に記録できます。

## 11. GitHubへの反映

`NITNC_DormNews`のルートで次を実行します。

```powershell
git status
git add -A
git diff --cached --name-only
git commit -m "Use English-only filenames for typeB and typeC outputs"
git push origin main
```

`.gitignore`では、実回答Excel・記事写真・生成物を原則除外し、架空の`demo_answer.xlsx`と`demo_answer.csv`だけをテスト用に追跡する設定です。実行前に`git diff --cached --name-only`で確認してください。

**注意：** このREADMEは現行仕様に対応したガイドです。発行前に、誤字脱字・掲載画像の許可・教員コメント・Forms回答の内容を必ず確認してください。
