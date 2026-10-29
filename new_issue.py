"""次号の原稿ファイルを作成。発行年と月は号数と独立して指定可能。"""
import argparse
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def parse_issue_date(args):
    if args.date:
        m = re.fullmatch(r"(\d{4})年(\d{1,2})月", args.date.strip())
        if not m:
            raise SystemExit('--dateには「2026年11月」の形式を指定してください。')
        yr, mo = int(m.group(1)), int(m.group(2))
        if args.year is not None and args.year != yr or args.month is not None and args.month != mo:
            raise SystemExit('--dateと--year/--monthが一致しません。')
    else:
        if args.year is None or args.month is None:
            raise SystemExit('--yearと--monthの両方を指定してください。例：--year 2026 --month 11')
        yr, mo = args.year, args.month
    if not (2000 <= yr <= 2100 and 1 <= mo <= 12):
        raise SystemExit('発行年・月の値を確認してください。')
    return yr, mo


def main():
    p = argparse.ArgumentParser(description='寮生会新聞の次号用フォルダを準備')
    p.add_argument('--volume', type=int, required=True, help='号数（例：6）')
    p.add_argument('--year', type=int, help='発行年（例：2026）')
    p.add_argument('--month', type=int, help='発行月（例：11）')
    p.add_argument('--date', help='旧方式との互換用。例：2026年11月')
    a = p.parse_args()
    if a.volume < 1:
        raise SystemExit('号数は1以上で指定してください。')
    year, month = parse_issue_date(a)
    dest = ROOT / 'issues' / f'Vol{a.volume:02}'
    edition_path = dest / 'edition.json'
    # 画像の保管場所を号ごとに用意する。繰り返し実行しても既存画像は消さない。
    dest.mkdir(parents=True, exist_ok=True)
    assets_dir = dest / 'assets'
    assets_dir.mkdir(parents=True, exist_ok=True)
    if edition_path.exists():
        # 原稿の誤上書きを防ぎながら、assetsフォルダだけは補完できる。
        print('[既存] 原稿は上書きしません:', edition_path)
        print('[確認] 画像フォルダ:', assets_dir)
        existing = json.loads(edition_path.read_text(encoding='utf-8'))
        if existing.get('issue_year') != year or existing.get('issue_month') != month:
            print('[注意] 指定した発行年・月と既存edition.jsonが異なります。既存の設定は変更していません。')
        print('生成コマンド:')
        print(f'python generate.py --input "issues/Vol{a.volume:02}/回答.xlsx" --edition "issues/Vol{a.volume:02}/edition.json" --volume {a.volume} --design both')
        return
    edition = json.loads((ROOT / 'edition.json').read_text(encoding='utf-8'))
    edition['volume'] = a.volume
    edition['issue_year'] = year
    edition['issue_month'] = month
    edition.pop('issue_date', None)
    edition['intro'] = '寮生会新聞の導入文を入力してください。'
    edition['notices'] = ['お知らせ①（入力してください）', 'お知らせ②（入力してください）', 'お知らせ③（入力してください）']
    for i in (1, 2):
        edition[f'article{i}']['title'] = f'記事{i}のタイトルを入力してください'
        edition[f'article{i}']['body'] = f'記事{i}の本文を入力してください'
        edition[f'article{i}']['image'] = ''
        edition[f'article{i}']['date_label'] = '写真を設定'
    edition['teacher1'] = '（寮務主事のコメントを入力）'
    edition['teacher2'] = '（寮監のコメントを入力）'
    edition_path.write_text(json.dumps(edition, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    (dest / 'ここにフォームの回答.xlsxを保存.txt').write_text(
        'Formsの新しい号の回答Excelを同じフォルダに置きます。前号の回答は使いません。\n', encoding='utf-8')
    print('[作成]', edition_path)
    print('[作成]', assets_dir)
    print('左下表示：寮生会庶務/vol.' + str(a.volume) + '/' + str(year) + '/' + str(month))
    print('Formsの複製と次号回答Excelの設定は、発行担当者が行ってください。')
    print('生成コマンド:')
    print(f'python generate.py --input "issues/Vol{a.volume:02}/回答.xlsx" --edition "issues/Vol{a.volume:02}/edition.json" --volume {a.volume} --design both')


if __name__ == '__main__':
    main()
