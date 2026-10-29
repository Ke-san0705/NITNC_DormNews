# 記事の画像を0〜2枚にする方法

`issues/Vol06/edition.json`など、新聞の号ごとの`edition.json`を編集してください。

## 画像0枚（日付カード）

```json
"article1": {
  "category": "DORM LIFE / 01",
  "title": "イベントのお知らせ",
  "body": "記事本文",
  "images": [],
  "date_label": "11 / 01"
}
```

## 画像1枚（全面）

```json
"article1": {
  "category": "DORM LIFE / 01",
  "title": "イベント開催",
  "body": "記事本文",
  "images": ["assets/写真1.jpg"]
}
```

## 画像2枚（横並び）

```json
"article1": {
  "category": "DORM LIFE / 01",
  "title": "イベント開催",
  "body": "記事本文",
  "images": ["assets/写真1.jpg", "assets/写真2.jpg"]
}
```

`article2`でも同様です。画像は`edition.json`と同じ場所にある`assets/`フォルダに入れます。

```text
NITNC_DormNews/
└── issues/
    └── Vol06/
        ├── edition.json
        ├── 回答.xlsx
        └── assets/
            ├── 写真1.jpg
            └── 写真2.jpg
```

生成コマンドは従来どおり、`generate.py`のある`NITNC_DormNews`フォルダで実行します。

```powershell
python generate.py --input "issues/Vol06/回答.xlsx" --edition "issues/Vol06/edition.json" --volume 6 --design both
```

### 注意

- 画像は多くても**2枚**です。3枚以上はエラーになります。
- 対応形式は JPEG / PNG / WebP、ファイルサイズは**1枚8MBまで**です。
- 写真は枠に合わせて**中央でトリミング**されます。2枚の場合は左右均等幅です。
- 既存の `"image": "assets/写真1.jpg"` は従来どおり使えます。`images`と`image`を同時に指定しないでください。
- GitHub公開用`.gitignore`で実際の画像は除外されています。画像は手元で管理してください。
- 生成画像はデモでも配布前に内容・掲載許可を確認してください。
