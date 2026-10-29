# GitHubへの初回登録 — NITNC_DormNews

**保存先**：`Ke-san0705/NITNC_DormNews`

## 1. GitHubに空のリポジトリを作成

1. GitHubにサインインし、右上の **+ → New repository** を開きます。
2. **Repository name**：`NITNC_DormNews`
3. **Visibility**：まずは **Private（非公開）** を推奨します。寮生会関係の情報を扱うためです。
4. **Add a README file / .gitignore / license** はすべて付けず、空のリポジトリを作成します。
5. **Create repository** を押します。

## 2. このプロジェクトを登録（Windows PowerShell）

Gitがインストールされていない場合は [Git for Windows](https://git-scm.com/downloads/win) をインストールします。
このZIPを展開し、`generate.py`がある`NITNC_DormNews`フォルダを開き、アドレスバーに `powershell` と入力してEnterを押します。続けて：

```powershell
git init -b main
git add .
git status --short
git commit -m "Initial commit: Dorm News generator"
git remote add origin https://github.com/Ke-san0705/NITNC_DormNews.git
git push -u origin main
```

初めてGitを使う場合、`git commit`の前にユーザー情報を設定する必要があります（アドレスはご自身のGitHub設定に合わせてください）。

```powershell
git config --global user.name "Ke-san0705"
git config --global user.email "GitHubで設定したコミット用メールアドレス"
```

`git push`時にはGitHubへのサインインが求められる場合があります。**トークンやパスワードをこのチャットには送らないでください。**

## 3. 更新するとき

```powershell
git add .
git commit -m "Update newspaper generator"
git push
```

## 4. 登録前チェック

- `git status --short`で`output/`、実際の回答Excel、記事写真が含まれていないことを確認。
- `samples/`のファイルは架空の回答データです。実データに置き換えないでください。
- `.gitignore`は**追加前のファイルだけ**を対象にします。すでにGitに登録した機密ファイルは、`.gitignore`を追加しただけでは消えません。
- ファイルの形式はUTF-8を推奨します。

## 5. 自動生成の基本操作

初回のみ：

```powershell
python -m pip install -r requirements.txt
python -m playwright install chromium
```

デモ実行：

```powershell
python generate.py --demo --volume 5
```

新号の準備（原稿と`assets/`を同時作成）：

```powershell
python new_issue.py --volume 7 --year 2027 --month 1
```

実際のVol.7を生成する場合：

```powershell
python generate.py --input "issues/Vol07/回答.xlsx" --edition "issues/Vol07/edition.json" --volume 7 --design both
```

その他の詳しい操作方法は`README.md`を参照してください。
