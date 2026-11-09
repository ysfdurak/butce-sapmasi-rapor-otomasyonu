import sys,json,tempfile,copy,shutil,unittest
from pathlib import Path
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from model import Dataset
from notes import Notes
from note_review import PENDING,previous,report_text
from openpyxl import load_workbook
import jobs,butce

class NoteReviewTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data=Dataset(ROOT/'ornek/Ayarlar.xlsx',(2026,10));cls.prev=Dataset(ROOT/'ornek/Ayarlar.xlsx',(2026,9));cls.files=sorted((ROOT/'ornek/onceki_yorumlar').glob('Onceki*.xlsx'));cls.expected=json.loads((ROOT/'ornek/onceki_yorumlar/beklenen.json').read_text())
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.folder=Path(self.temp.name);self.notes=Notes(self.folder/'notes.sqlite3');self.months=list(range(1,11))
    def tearDown(self):self.temp.cleanup()
    def carry(self):return self.notes.review.carry(self.data,self.months,self.files)
    def rows(self):return self.notes.review.rows(self.data,self.months)
    def entry(self,r,action='approve',draft=None):return {'department':r['department'],'code':r['code'],'token':r['token'],'action':action,'draft':r['draft'] if draft is None else draft}
    def test_example_files_have_ten_pending_and_different_month_scope(self):
        self.assertEqual(self.carry()['created'],10);rows=self.rows();self.assertEqual(len(rows),10)
        self.assertTrue(all(r['status']=='pending' and not r['sameMonths'] for r in rows))
        self.assertEqual(len(self.notes.list(self.data,self.months)),10)
        self.assertTrue(all(r['status']=='pending' for r in self.notes.list(self.data,self.months).values()))
    def test_two_unchanged_eight_edited_in_two_bulk_actions(self):
        self.carry();expect={(r['department'],r['code']):r for r in self.expected}
        first=[self.entry(r) for r in self.rows() if expect[r['department'],r['code']]['action']=='Aynen onayla']
        self.assertEqual(len(first),2);self.notes.review.apply(self.data,self.months,first)
        edits=[self.entry(r,draft=expect[r['department'],r['code']]['new']) for r in self.rows() if r['status']=='pending']
        self.assertEqual(len(edits),8);self.notes.review.apply(self.data,self.months,edits)
        self.assertTrue(all(r['status']=='approved' for r in self.rows()))
        saved=self.notes.list(self.data,self.months)
        for key,record in expect.items():self.assertEqual(saved[key]['note'],record['new']);self.assertEqual(saved[key]['previousClosing'],'2026-09')
    def test_existing_current_note_is_not_overwritten(self):
        item=self.expected[0];self.notes.save(self.data,self.months,[{'department':item['department'],'code':item['code'],'note':'Bu ay girilmiş açıklama'}])
        result=self.carry();self.assertEqual(result['created'],9);self.assertEqual(result['skipped'],1)
        self.assertEqual(self.notes.list(self.data,self.months)[item['department'],item['code']]['note'],'Bu ay girilmiş açıklama')
        self.assertEqual(sum(r['protected'] for r in self.rows()),1)
    def test_repeat_import_preserves_drafts_rejections_and_approvals(self):
        self.carry();rows=self.rows();self.notes.review.apply(self.data,self.months,[self.entry(rows[0],'reject'),self.entry(rows[1],'draft','Düzenlenen taslak'),self.entry(rows[2])])
        before=self.rows();self.assertEqual(self.carry()['skipped'],10);self.assertEqual(before,self.rows())
    def test_rejected_note_is_not_exported_as_effective_comment(self):
        self.carry();r=self.rows()[0];self.notes.review.apply(self.data,self.months,[self.entry(r,'reject')])
        self.assertNotIn((r['department'],r['code']),self.notes.list(self.data,self.months))
    def test_draft_save_remains_unconfirmed(self):
        self.carry();r=self.rows()[0];self.notes.review.apply(self.data,self.months,[self.entry(r,'draft','Yeni taslak')])
        result=self.notes.list(self.data,self.months)[r['department'],r['code']];self.assertEqual(result['note'],'Yeni taslak');self.assertEqual(result['status'],'pending')
    def test_optimistic_lock_prevents_losing_other_window_edit(self):
        self.carry();r=self.rows()[0];self.notes.review.apply(self.data,self.months,[self.entry(r,'draft','Diğer pencere')])
        with self.assertRaisesRegex(ValueError,'değişti'):self.notes.review.apply(self.data,self.months,[self.entry(r)])
        self.assertEqual(self.rows()[0]['draft'],'Diğer pencere')
    def test_entire_bulk_operation_rolls_back_if_one_token_is_invalid(self):
        self.carry();rows=self.rows();a=self.entry(rows[0]);b=self.entry(rows[1]);b['token']='invalid'
        with self.assertRaises(ValueError):self.notes.review.apply(self.data,self.months,[a,b])
        self.assertTrue(all(r['status']=='pending' for r in self.rows()))
    def test_manual_edit_after_carry_is_protected(self):
        self.carry();r=self.rows()[0];self.notes.save(self.data,self.months,[{'department':r['department'],'code':r['code'],'note':'Yeni elle giriş'}])
        with self.assertRaisesRegex(ValueError,'korunuyor'):self.notes.review.apply(self.data,self.months,[self.entry(r)])
        self.assertEqual(self.notes.list(self.data,self.months)[r['department'],r['code']]['note'],'Yeni elle giriş')
    def test_wrong_closing_report_is_rejected_atomically(self):
        current=ROOT/'ornek/aciklama/Aciklama_Ornek_Mudurluk_01.xlsx'
        with self.assertRaisesRegex(ValueError,'önceki'):self.notes.review.carry(self.data,self.months,[self.files[0],current])
        self.assertEqual(self.rows(),[])
    def test_calendar_previous_closing_including_year_rollover(self):
        self.assertEqual(previous((2026,1)),(2025,12));self.assertEqual(previous((2026,10)),(2026,9))
    def test_approved_local_history_is_brought_without_files(self):
        item=self.expected[0];self.notes.save(self.prev,list(range(1,10)),[{'department':item['department'],'code':item['code'],'note':'Yerel geçmiş'}])
        self.assertEqual(self.notes.review.carry(self.data,self.months)['created'],1);self.assertEqual(self.rows()[0]['oldNote'],'Yerel geçmiş')
    def test_multiple_old_month_selections_are_not_guessed(self):
        item=self.expected[0]
        for m in ([1],[2]):self.notes.save(self.prev,m,[{'department':item['department'],'code':item['code'],'note':'Farklı kapsam'}])
        result=self.notes.review.carry(self.data,self.months);self.assertEqual(result['created'],0);self.assertTrue(any('birden fazla' in s for s in result['issues']))
    def test_department_scope_is_applied(self):
        name=self.expected[0]['department'];r=self.notes.review.carry(self.data,self.months,self.files,departments=[name]);self.assertEqual(r['created'],2);self.assertTrue(all(r['department']==name for r in self.rows()))
    def test_same_values_do_not_approve_the_comment(self):
        item=self.expected[0];self.notes.save(self.prev,[1],[{'department':item['department'],'code':item['code'],'note':'İnceleme bekleyen'}])
        self.notes.review.carry(self.data,[1]);rows=self.notes.review.rows(self.data,[1]);self.assertFalse(rows[0]['dataChanged']);self.assertEqual(rows[0]['status'],'pending')
    def test_changed_amount_invalidates_review_token(self):
        self.carry();r=self.rows()[0];d=copy.deepcopy(self.data);key=(r['department'],2026,1);rec=d.records[key]['items'][r['code']];d.records[key]['items'][r['code']]=(rec[0]+100,rec[1])
        with self.assertRaisesRegex(ValueError,'değişti'):self.notes.review.apply(d,self.months,[self.entry(r)])
    def test_pending_export_red_labeled_and_source_months_visible(self):
        manager=self.folder/'manager.xlsx';general=self.folder/'general.xlsx';shutil.copy2(ROOT/'ornek/aciklama/Aciklama_Ornek_Mudurluk_01.xlsx',manager);shutil.copy2(ROOT/'ornek/aciklama/Aciklama_Ornek_Direktorluk_Birlesik.xlsx',general)
        notes={('Giriş',1):{'note':'=Bu bir formül değil','status':'pending','previousClosing':'2026-09','previousMonths':'aylar:01,02,03,04,05,06,07,08,09'}}
        jobs.write_notes([manager,general],notes)
        b=load_workbook(manager);_,members,cols,items=butce.explanation_layout(b,'Müdürlük');r=next(v for k,v in items.items() if k[0]==1);cell=b['Giderler'].cell(r,cols['Giriş']+4)
        self.assertTrue(cell.value.startswith(PENDING));self.assertIn('2026-09',cell.value);self.assertTrue(cell.font.color.rgb.endswith('C63842'));self.assertEqual(cell.data_type,'s');self.assertEqual(b['Açıklama Durumu']['H2'].data_type,'s');b.close()
        b=load_workbook(general);self.assertIn('Giriş: '+PENDING,b['Alt Gider Açıklamaları']['D2'].value);self.assertTrue(b['Alt Gider Açıklamaları']['D2'].font.color.rgb.endswith('C63842'));b.close()
        with self.assertRaisesRegex(ValueError,'bekleyen'):self.notes.import_files(self.data,self.months,[manager])
    def test_approved_export_has_no_pending_label(self):
        self.assertEqual(report_text({'note':'Onaylandı','status':'approved','previousClosing':'2026-09'}),'Onaylandı')
    def test_duplicate_entries_rejected_without_partial_commit(self):
        self.carry();r=self.rows()[0]
        with self.assertRaisesRegex(ValueError,'yineleniyor'):self.notes.review.apply(self.data,self.months,[self.entry(r),self.entry(r)])
        self.assertEqual(self.rows()[0]['status'],'pending')
    def test_approved_comment_becomes_pending_again_after_amount_change(self):
        self.carry();r=self.rows()[0];self.notes.review.apply(self.data,self.months,[self.entry(r)])
        d=copy.deepcopy(self.data);rec=d.records[r['department'],2026,1]['items'][r['code']];d.records[r['department'],2026,1]['items'][r['code']]=(rec[0]+100,rec[1])
        value=self.notes.list(d,self.months)[r['department'],r['code']];self.assertEqual(value['status'],'pending');self.assertTrue(value['stale']);self.assertEqual(value['note'],r['draft'])
