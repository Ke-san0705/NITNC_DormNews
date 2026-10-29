import csv
import json
import tempfile
import unittest
from pathlib import Path

from report_engine import FORM_DEPARTMENTS, PAPER_ORDER, InputProblem, load_and_merge, merge_rows
from generate import ROOT, compile_once, validate_edition

H = ['Id','開始時刻','完了時刻','メールアドレス','名前','部署','①活動報告','②寮生に対しての連絡事項','③その他']


class NewspaperTests(unittest.TestCase):
    def test_department_list(self):
        self.assertEqual(len(FORM_DEPARTMENTS), 16)
        self.assertEqual(FORM_DEPARTMENTS[0], '三役')
        self.assertEqual(FORM_DEPARTMENTS[-1], '庶務長')
        self.assertEqual(len(PAPER_ORDER), 15)
        self.assertIn('秀峰寮東', PAPER_ORDER)
        self.assertIn('秀峰寮西', PAPER_ORDER)
        self.assertNotIn('秀峰寮', FORM_DEPARTMENTS)

    def test_new_formal_names_and_legacy_aliases(self):
        self.assertIn('秀峰寮東', FORM_DEPARTMENTS)
        self.assertIn('秀峰寮西', FORM_DEPARTMENTS)
        self.assertNotIn('秀峰東寮', FORM_DEPARTMENTS)
        self.assertNotIn('秀峰西寮', FORM_DEPARTMENTS)
        data=[H,
            ['1','','2026/10/29 10:00','','','秀峰東寮','旧形式東①','旧形式東②',''],
            ['2','','2026/10/29 10:01','','','秀峰西寮','旧形式西①','旧形式西②',''],
            ['3','','2026/10/29 10:10','','','秀峰寮東','正式東①','正式東②','']]
        merged=merge_rows(data, 5)
        self.assertEqual(merged.latest['秀峰寮東']['activity'], '正式東①')
        self.assertEqual(merged.latest['秀峰寮西']['activity'], '旧形式西①')
        self.assertEqual(merged.audit['submitted'], 2)

    def test_latest_override_entire_packet(self):
        data=[H,
            ['1','','2026/10/29 10:00','','','翔峰寮','旧①','旧②','旧③'],
            ['2','','2026/10/29 11:00','','','翔峰寮','新①','新②','']]
        merged=merge_rows(data,5)
        self.assertEqual(merged.latest['翔峰寮']['activity'],'新①')
        self.assertNotIn('旧③',merged.latest['翔峰寮']['display'])
        self.assertEqual(merged.audit['duplicates'][0]['count'],2)

    def test_submit_time_over_row_order(self):
        data=[H,['3','','2026/10/29 11:00','','','三役','新','報告',''],
             ['1','','2026/10/28 11:00','','','三役','旧','旧報告','']]
        self.assertEqual(merge_rows(data,5).latest['三役']['activity'],'新')

    def test_sample_files_equal(self):
        root=ROOT/'samples'
        a=load_and_merge(root/'動作確認用_架空回答.csv',5)
        b=load_and_merge(root/'動作確認用_架空回答.xlsx',5)
        self.assertEqual(a.latest['翔峰寮']['display'], b.latest['翔峰寮']['display'])
        self.assertEqual(a.latest['秀峰寮東']['display'], b.latest['秀峰寮東']['display'])
        self.assertEqual(a.latest['秀峰寮西']['display'], b.latest['秀峰寮西']['display'])

    def test_shuhou_east_west_independent(self):
        data=[H,
            ['1','','2026/10/29 10:00','','','秀峰寮東','東①旧','東②旧','東③旧'],
            ['2','','2026/10/29 10:02','','','秀峰寮西','西①','西②','西③'],
            ['3','','2026/10/29 10:10','','','秀峰寮東','東①新','東②新','']]
        result=merge_rows(data,5)
        self.assertEqual(result.latest['秀峰寮東']['activity'],'東①新')
        self.assertNotIn('東③旧',result.latest['秀峰寮東']['display'])
        self.assertEqual(result.latest['秀峰寮西']['activity'],'西①')
        self.assertEqual(result.audit['submitted'],2)

    def test_old_ambiguous_shuhou_rejected(self):
        data=[H,['1','','2026/10/29 10:00','','','秀峰寮','作業','連絡','']]
        with self.assertRaisesRegex(InputProblem,'東西を区別'):
            merge_rows(data,5)

    def test_corrupt_csv_rejected(self):
        source=Path('/mnt/data/寮生会新聞 Vol.csv')
        if source.exists():
            with self.assertRaises(InputProblem):load_and_merge(source,5)

    def test_uknown_department_rejected(self):
        data=[H,['2','','2026/10/29 11:00','','','架空部署','a','b','']]
        with self.assertRaises(InputProblem):merge_rows(data,5)

    def test_both_pdfs_created_one_page(self):
        with tempfile.TemporaryDirectory() as tmp:
            out=Path(tmp)/'issue'
            r=compile_once(ROOT/'samples'/'動作確認用_架空回答.xlsx',ROOT/'edition.json',out,5,['classic','blue'],False,95)
            self.assertEqual(r['submitted'],5)
            for layout in ('classic','blue'):
                pdfs=list((out/layout).glob('*.pdf'))
                self.assertEqual(len(pdfs),1)
                self.assertGreater(pdfs[0].stat().st_size,8000)
            self.assertTrue((out/'提出状況.html').exists())
            for layout in ('classic','blue'):
                html_file=next((out/layout).glob('*.html'))
                html=html_file.read_text(encoding='utf-8')
                self.assertIn('data-department=\"秀峰寮東\"', html)
                self.assertIn('data-department=\"秀峰寮西\"', html)
                self.assertNotIn('data-department=\"予備\"', html)
                self.assertIn('秀峰寮東', html)
                self.assertIn('秀峰寮西', html)
                self.assertIn('<footer class="footer">寮生会庶務/vol.5/2026/10</footer>', html)
                self.assertIn('発行：2026年10月', html)
            status=(out/'提出状況.html').read_text(encoding='utf-8')
            self.assertIn('5/16部署', status)

    def test_footer_changes_with_edition_and_validates_month(self):
        with tempfile.TemporaryDirectory() as td:
            temp = Path(td)
            edition = json.loads((ROOT/'edition.json').read_text(encoding='utf-8'))
            edition.update(volume=6, issue_year=2027, issue_month=2)
            edition['article1']['image'] = ''  # 写真の参照は本テストでは不要
            target = temp/'edition.json'
            target.write_text(json.dumps(edition, ensure_ascii=False), encoding='utf-8')
            output = temp/'Vol6'
            compile_once(ROOT/'samples'/'動作確認用_架空回答.xlsx', target, output, 6, ['classic','blue'], False, 95)
            for mode in ('classic','blue'):
                html = next((output/mode).glob('*.html')).read_text(encoding='utf-8')
                self.assertIn('<footer class="footer">寮生会庶務/vol.6/2027/2</footer>', html)
                self.assertIn('発行：2027年2月', html)
            edition['issue_month'] = 13
            target.write_text(json.dumps(edition, ensure_ascii=False), encoding='utf-8')
            with self.assertRaisesRegex(InputProblem, 'issue_month'):
                validate_edition(target, 6)

    def test_too_long_body_blocks_pdf(self):
        with tempfile.TemporaryDirectory() as tmp:
            d=Path(tmp)
            src=d/'too_long.csv'
            with src.open('w',encoding='utf-8-sig',newline='') as f:
                w=csv.writer(f);w.writerow(H)
                w.writerow(['1','','2026/10/29 10:00','','','翔峰寮','長'*140,'連絡'*100,''])
            out=d/'out'
            with self.assertRaises(InputProblem):
                compile_once(src,ROOT/'edition.json',out,5,['classic'],False,95)
            self.assertFalse(list((out/'classic').glob('*.pdf')))
            state=json.loads((out/'提出状況.json').read_text(encoding='utf-8'))
            self.assertIn('OVERFLOW',state['status'])


if __name__ == '__main__':unittest.main(verbosity=2)
