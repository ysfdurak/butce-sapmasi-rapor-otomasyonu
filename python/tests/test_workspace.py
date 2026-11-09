import sys,unittest,tempfile,shutil,json,time,threading
from pathlib import Path
from decimal import Decimal
from unittest.mock import patch
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from app import Application
from model import Dataset
import workspace,configuration

class WorkspaceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):cls.base=Dataset(ROOT/'ornek/Ayarlar.xlsx',(2026,10))
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.path=Path(self.tmp.name)
    def tearDown(self):self.tmp.cleanup()
    def application(self):
        from notes import Notes
        from concurrent.futures import ThreadPoolExecutor
        a=Application.__new__(Application);a.state_dir=self.path;a.notes=Notes(self.path/'notes.sqlite3')
        with a.notes.connect() as c:c.execute('CREATE TABLE IF NOT EXISTS preferences (key TEXT PRIMARY KEY,value TEXT)')
        a.data=self.base;a.executor=ThreadPoolExecutor(max_workers=1);a.lock=threading.Lock();a.job=None;a.downloads={};a.workspace=workspace.Workspace(a)
        self.addCleanup(a.executor.shutdown,wait=True)
        return a
    def test_trace_merged_item_reconciles_and_excludes_counts(self):
        s={'filters':{'year':[2026],'month':[1,3,5],'department':['Giriş']}}
        t=workspace.trace(self.base,s,code=1)
        a=sum(r[-1] for r in t['rows'] if r[4]=='Gerçekleşen');b=sum(r[-1] for r in t['rows'] if r[4]=='Bütçe')
        self.assertEqual(a,t['actualTL']);self.assertEqual(b,t['budgetTL'])
        self.assertEqual(len(t['rows']),12) # two raw items merged, three months, two scenarios
        self.assertTrue(all(int(''.join(filter(str.isdigit,r[8])))>=65 for r in t['rows']))
        self.assertTrue(all(r[5][-2:] in ['01','03','05'] for r in t['rows']))
    def test_multi_year_trace_hides_budget(self):
        t=workspace.trace(self.base,{'filters':{'year':[2025,2026],'month':[2],'department':['Giriş']}},code=1)
        self.assertIsNone(t['budgetTL']);self.assertTrue(all(r[4]=='Gerçekleşen' for r in t['rows']))
    def test_quality_reports_scope_and_exclusions(self):
        q=workspace.quality(self.base,{'bmLocal':['Giriş'],'opexLocal':['Tartım']})
        self.assertTrue(any(r[0]=='Tartım' and r[1]=='Birim maliyet' for r in q['exclusions']))
        self.assertEqual(q['expectedDepartments'],12);self.assertEqual(q['receivedDepartments'],12)
    def test_snapshot_replay_is_identical_and_tamper_detected(self):
        a=self.application();rid,p=a.workspace.capture(self.base,'report',{'year':2026,'months':[1,3,5]})
        a.workspace.finish(rid,'done',{})
        path,manifest=a.workspace.replay_settings(rid);replay=Dataset(path,manifest['closing'])
        self.assertEqual(replay.query({})['general'],self.base.query({})['general'])
        self.assertEqual(len(manifest['sources']),13)
        path.write_bytes(b'corrupt')
        with self.assertRaisesRegex(ValueError,'ayar kopyası'):a.workspace.replay_settings(rid)
    def test_snapshot_rejects_added_source_file(self):
        a=self.application();rid,path=a.workspace.capture(self.base,'report',{'year':2026,'months':[1]})
        (path.parent/'maliyet'/'extra.xlsx').write_bytes(b'wrong')
        with self.assertRaisesRegex(ValueError,'listesi'):a.workspace.replay_settings(rid)
    def test_mapping_review_and_explicit_rename_keep_id_and_notes(self):
        a=self.application();a.workspace.accept({'fingerprint':configuration.digest(self.base.settings)})
        key='mapping:'+self.base.identity;baseline=a.preference(key);old='Giriş';new='Eski Giriş'
        baseline['mappings'][new]=baseline['mappings'].pop(old);baseline['ids'][new]=baseline['ids'].pop(old);a.save_preference(key,baseline)
        a.notes.save(self.base,[1],[{'department':old,'code':1,'note':'Geçmiş açıklama'}])
        with a.notes.connect() as c:c.execute('UPDATE notes SET department=? WHERE department=?',(new,old))
        self.assertIn(old,a.workspace.review()['added'])
        a.workspace.accept({'fingerprint':configuration.digest(self.base.settings),'renames':[{'old':new,'new':old}]})
        self.assertEqual(a.preference(key)['ids'][old],baseline['ids'][new])
        self.assertEqual(a.notes.list(self.base,[1])[old,1]['note'],'Geçmiş açıklama')
    def test_conflicting_rename_does_not_move_notes(self):
        a=self.application()
        with self.assertRaisesRegex(ValueError,'tekil'):a.workspace.accept({'fingerprint':configuration.digest(self.base.settings),'renames':[{'old':'Unknown','new':'Giriş'}]})
        self.assertIsNone(a.preference('mapping:'+self.base.identity))
    def test_cancel_load_preserves_previous_dataset(self):
        a=self.application();started=threading.Event();release=threading.Event()
        def factory(*args,**kw):started.set();release.wait(5);kw['progress']('Dosya adımı sonu');return self.base
        with patch('app.Dataset',side_effect=factory):
            a.start_job('load',{'settings':str(self.base.settings),'closing':[2026,10]});self.assertTrue(started.wait(3))
            a.job['cancelRequested']=True;release.set()
            for _ in range(100):
                if a.job['state']!='running':break
                time.sleep(.01)
        self.assertEqual(a.job['state'],'cancelled');self.assertIs(a.data,self.base)
    def test_changed_source_during_capture_is_rejected(self):
        a=self.application();original=workspace.configuration.digest
        with patch('workspace.configuration.digest',side_effect=lambda p:('changed' if 'inputs' in str(p) else original(p))):
            with self.assertRaisesRegex(ValueError,'kopyalanırken'):a.workspace.capture(self.base,'report',{'year':2026,'months':[1]})
        self.assertEqual(a.workspace.history()[0]['state'],'error')
    def test_layout_preview_uses_actual_cells(self):
        r=configuration.layout_preview(self.base.settings)
        self.assertEqual(r['sample'][0][1][0],'BF59');self.assertTrue(r['parameters'])
    def test_layout_failure_preserves_settings(self):
        shutil.copytree(ROOT/'ornek',self.path/'ornek');p=self.path/'ornek/Ayarlar.xlsx';before=p.read_bytes()
        with self.assertRaises(Exception):configuration.update_layout(p,configuration.digest(p),[['Maliyet','ilk_gider_satiri',1]],[2026,10],Dataset)
        self.assertEqual(p.read_bytes(),before)
    def test_replay_path_traversal_rejected(self):
        a=self.application()
        for rid in ('../other','a/b','a\\b',''):
            with self.assertRaises(ValueError):a.workspace.replay_settings(rid)
    def test_successful_layout_change_is_backed_up_and_atomic(self):
        shutil.copytree(ROOT/'ornek',self.path/'ornek');p=self.path/'ornek/Ayarlar.xlsx';before=p.read_bytes()
        data,backup=configuration.update_layout(p,configuration.digest(p),[['Kontrol','tolerans',0.02]],[2026,10],Dataset)
        self.assertEqual(Path(backup).read_bytes(),before);self.assertEqual(data.settings,p.resolve())
        self.assertEqual(Decimal(str(data.config['kontrol']['tolerans'])),Decimal('0.02'))
    def test_report_job_snapshot_and_replay_use_frozen_notes(self):
        a=self.application();a.notes.save(self.base,[1],[{'department':'Giriş','code':1,'note':'İlk açıklama'}])
        results=[]
        def output(data,year,months,path,notes,checkpoint=None):
            results.append((data.query({'filters':{'month':months}})['general']['cards']['actual'],notes['Giriş',1]['note']))
            return {'folder':str(path),'files':[]}
        def wait():
            end=time.monotonic()+40
            while a.job['state']=='running' and time.monotonic()<end:time.sleep(.02)
            self.assertEqual(a.job['state'],'done',a.job)
        with patch('app.jobs.report',side_effect=output):
            a.start_job('report',{'year':2026,'months':[1],'output':str(self.path/'out')});wait();rid=a.job['result']['runId']
            a.notes.save(self.base,[1],[{'department':'Giriş','code':1,'note':'Sonradan değişti'}])
            a.start_job('replay',{'id':rid,'output':str(self.path/'replay')});wait()
        self.assertEqual(results[0],results[1]);self.assertEqual(results[1][1],'İlk açıklama')
    def test_trace_rejects_changed_sources(self):
        with patch('workspace.inventory',return_value=[]):
            with self.assertRaisesRegex(ValueError,'değişmiş'):workspace.trace(self.base,{})
