"""Single-user, loopback-only dashboard. No network services or Excel automation."""
from http.server import BaseHTTPRequestHandler,ThreadingHTTPServer
from http.cookies import SimpleCookie
from urllib.parse import urlsplit,parse_qs,quote
from pathlib import Path
from decimal import Decimal
from concurrent.futures import ThreadPoolExecutor
import argparse,json,secrets,threading,sqlite3,traceback,subprocess,sys,os,webbrowser,time
from datetime import datetime
from model import Dataset
from notes import Notes
import jobs,butce
import configuration
import workspace

BASE=Path(__file__).resolve().parent
MAX_BODY=2*1024*1024
def serial(value):
    if isinstance(value,Decimal):return float(value)
    if isinstance(value,Path):return str(value)
    raise TypeError(type(value).__name__)

class Application:
    def __init__(self,settings=None,closing=None,state_dir=None):
        self.state_dir=Path(state_dir or BASE/'kullanici');self.state_dir.mkdir(parents=True,exist_ok=True)
        self.notes=Notes(self.state_dir/'kayitlar.sqlite3')
        with self.notes.connect() as c:c.execute('CREATE TABLE IF NOT EXISTS preferences (key TEXT PRIMARY KEY,value TEXT)')
        saved=self.preference('startup',{})
        settings=Path(settings or saved.get('settings') or BASE/'ornek/Ayarlar.xlsx')
        if not settings.is_file():settings=BASE/'ornek/Ayarlar.xlsx'
        if closing is None:
            if saved.get('settings')==str(settings.resolve()) and saved.get('closing'):closing=tuple(saved['closing'])
            else:
                cfg=butce.read_settings(settings);folder=settings.parent/cfg['maliyet_klasoru'];periods=[]
                for path in folder.glob('*.xlsx'):
                    if path.name.startswith('~$'):continue
                    try:periods.append(butce.filename_period(path))
                    except ValueError:pass
                if not periods:raise ValueError('Kaynak adlarında yıl/ay bulunamadı.')
                closing=max(periods)
        self.data=Dataset(settings,closing)
        self.executor=ThreadPoolExecutor(max_workers=1);self.lock=threading.Lock();self.job=None
        self.workspace=workspace.Workspace(self)
        self.downloads={};self.token=secrets.token_urlsafe(32)
    def preference(self,key,default=None):
        with self.notes.connect() as c:row=c.execute('SELECT value FROM preferences WHERE key=?',(key,)).fetchone()
        if row:
            try:return json.loads(row[0])
            except ValueError:pass
        return default
    def save_preference(self,key,value):
        with self.notes.connect() as c:c.execute('INSERT OR REPLACE INTO preferences VALUES (?,?)',(key,json.dumps(value,ensure_ascii=False)))
    def meta(self):
        meta=self.data.meta()
        meta['savedFilters']=configuration.reconcile_selection(self.preference('filters',None),self.preference('filterOptions',meta['options']),meta['options'])
        try:meta['configuration']=configuration.preview(self.data.settings)
        except Exception as error:
            cfg=self.data.config
            meta['configuration']={'settings':str(self.data.settings),'fingerprint':None,'locations':{key:str((self.data.root/cfg[code]).resolve()) for key,code in configuration.KEYS.items()},'mappings':meta['mappings'],'error':str(error)}
        meta['mappingReview']=self.workspace.review()
        meta['output']=self.preference('output',str(BASE/'ciktilar'))
        return meta
    def query(self,state):
        data=self.data;result=data.query(state)
        notes=self.notes.list(data,result['selection']['months']) if result['selection']['years']==[data.budget_year] else {}
        for row in result['priority']['rows']:row.update(notes.get((row['department'],row['code']),{'note':'','stale':False}))
        result['noteCount']=sum(bool(n['note']) for n in notes.values())
        return result
    def start_job(self,kind,args):
        with self.lock:
            if self.job and self.job['state']=='running':raise ValueError('Bir işlem devam ediyor; tamamlanmasını bekleyin.')
            job={'id':secrets.token_hex(8),'state':'running','kind':kind,'started':datetime.now().isoformat(timespec='seconds'),'stage':'Sırada','cancelRequested':False,'cancellable':True}
            self.job=job
        def checkpoint(message):
            with self.lock:
                job['stage']=message
                cancelled=job['cancelRequested']
            if cancelled:raise workspace.Cancelled('İşlem güvenli adım sınırında durduruldu; önceki raporlar korundu.')
        def work():
            run_id=None
            try:
                checkpoint('İşlem hazırlanıyor')
                if kind=='load':
                    closing=args['closing']
                    if not isinstance(closing,list) or len(closing)!=2 or any(type(v) is not int for v in closing):raise ValueError('Kapanış yılı ve ayı tam sayı olmalı.')
                    candidate=Dataset(args['settings'],closing,progress=checkpoint)
                    checkpoint('Kaynaklar doğrulandı')
                    self.data=candidate
                    self.save_preference('startup',{'settings':str(candidate.settings),'closing':list(candidate.closing)})
                    self.save_preference('filters',None)
                    result={'message':'Kaynaklar yenilendi.','meta':self.meta()}
                elif kind=='configure':
                    candidate,backup=configuration.update_locations(args['settings'],args['expectedSettings'],args['fingerprint'],args['locations'],args['closing'],Dataset)
                    self.data=candidate
                    self.save_preference('startup',{'settings':str(candidate.settings),'closing':list(candidate.closing)})
                    self.save_preference('filters',None)
                    result={'message':'Kaynak yolları ayar Excel’ine kaydedildi ve tüm eşleştirmeler yeniden okundu.','backup':backup,'meta':self.meta()}
                elif kind=='layout':
                    checkpoint('Tablo düzeni doğrulanıyor')
                    candidate,backup=configuration.update_layout(self.data.settings,args['fingerprint'],args['parameters'],list(self.data.closing),lambda p,c:Dataset(p,c,progress=checkpoint))
                    self.data=candidate
                    result={'message':'Tablo düzeni doğrulandı ve kaydedildi.','backup':backup,'meta':self.meta()}
                elif kind in ('report','replay'):
                    if kind=='replay':
                        snapshot,previous=self.workspace.replay_settings(args['id'])
                        original=previous['args'];months=original['months'];year=original['year']
                        candidate=Dataset(snapshot,previous['closing'],progress=checkpoint)
                        candidate.identity=previous['dataset']
                    else:
                        months=args['months'];year=args['year']
                        checkpoint('Güncel kaynaklar okunuyor')
                        candidate=Dataset(self.data.settings,self.data.closing,progress=checkpoint)
                    if type(year) is not int or not isinstance(months,list) or not months or any(type(m) is not int or not 1<=m<=12 for m in months):raise ValueError('Rapor yılı ve ayları geçersiz.')
                    # Read sources again so monthly revisions are reflected in the exported reports.
                    frozen_notes={(r['department'],r['code']):{k:v for k,v in r.items() if k not in ('department','code')} for r in previous.get('notes',[])} if kind=='replay' else self.notes.list(candidate,sorted(set(months)))
                    run_id,snapshot=self.workspace.capture(candidate,kind,{'year':year,'months':sorted(set(months)),'output':args['output']},checkpoint,notes=frozen_notes)
                    frozen=Dataset(snapshot,candidate.closing,progress=checkpoint);frozen.identity=candidate.identity
                    result=jobs.report(frozen,year,sorted(set(months)),args['output'],frozen_notes,checkpoint=checkpoint)
                    result['runId']=run_id
                    self.workspace.finish(run_id,'done',result)
                    if kind=='report':self.data=candidate
                    self.save_preference('output',args['output'])
                    result['meta']=self.meta()
                elif kind=='import':
                    months=self.data.selection(args['state'])[2]
                    if self.data.selection(args['state'])[1]!=[self.data.budget_year]:raise ValueError('Açıklamalar için bütçe yılını tek başına seçin.')
                    count=self.notes.import_files(self.data,months,args['paths'])
                    result={'message':f'{count} bölüm–gider açıklaması yüklendi.'}
                elif kind=='merge':result=jobs.merge(args['general'],args['managers'],args['output'])
                else:raise ValueError('Bilinmeyen işlem.')
                for path in result.get('files',[]):
                    key=secrets.token_urlsafe(16);self.downloads[key]=Path(path)
                    result.setdefault('downloads',[]).append({'name':Path(path).name,'url':'/download/'+key})
                with self.lock:job.update(state='done',result=result)
            except Exception as e:
                status='cancelled' if isinstance(e,workspace.Cancelled) else 'error'
                if run_id:self.workspace.finish(run_id,status,{'error':str(e)})
                with self.lock:job.update(state=status,error=str(e))
                with (self.state_dir/'hatalar.log').open('a',encoding='utf-8') as f:f.write(datetime.now().isoformat()+'\n'+traceback.format_exc()+'\n')
        self.executor.submit(work)
        return dict(job)

class Server(ThreadingHTTPServer):
    daemon_threads=True
    allow_reuse_address=False
    def __init__(self,address,application):
        self.application=application
        super().__init__(address,Handler)
        self.origin=f'http://127.0.0.1:{self.server_port}'

class Handler(BaseHTTPRequestHandler):
    server_version='ButceYerel/1.0'
    def log_message(self,*args):pass
    def send(self,code,body=b'',kind='application/json; charset=utf-8',headers=None):
        if isinstance(body,(dict,list)):body=json.dumps(body,ensure_ascii=False,default=serial,allow_nan=False).encode()
        if isinstance(body,str):body=body.encode('utf-8')
        self.send_response(code);self.send_header('Content-Type',kind);self.send_header('Content-Length',str(len(body)))
        self.send_header('Cache-Control','no-store');self.send_header('X-Content-Type-Options','nosniff')
        self.send_header('Referrer-Policy','no-referrer')
        self.send_header('Content-Security-Policy',"default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'")
        for k,v in (headers or {}).items():self.send_header(k,v)
        self.end_headers()
        try:self.wfile.write(body)
        except (BrokenPipeError,ConnectionResetError):pass
    def authorized(self):
        expected=urlsplit(self.server.origin).netloc
        if self.headers.get('Host')!=expected:return False
        cookie=SimpleCookie()
        try:cookie.load(self.headers.get('Cookie',''))
        except Exception:return False
        session=cookie.get('butce_session')
        return bool(session and secrets.compare_digest(session.value,self.server.application.token))
    def do_GET(self):
        path=urlsplit(self.path);app=self.server.application
        if path.path=='/' and self.headers.get('Host')==urlsplit(self.server.origin).netloc:
            token=parse_qs(path.query).get('token',[''])[0]
            if token and secrets.compare_digest(token,app.token):
                return self.send(302,b'',headers={'Location':'/','Set-Cookie':'butce_session='+app.token+'; HttpOnly; SameSite=Strict; Path=/'})
        if not self.authorized():return self.send(403,{'error':'Yerel uygulamayı Baslat.bat ile açın.'})
        try:
            if path.path in ('/','/app.js','/style.css'):
                name={'/':'index.html','/app.js':'app.js','/style.css':'style.css'}[path.path]
                kind={'/':'text/html; charset=utf-8','/app.js':'text/javascript; charset=utf-8','/style.css':'text/css; charset=utf-8'}[path.path]
                return self.send(200,(BASE/'web'/name).read_bytes(),kind)
            if path.path=='/api/meta':return self.send(200,app.meta())
            if path.path=='/favicon.ico':return self.send(204,b'','image/x-icon')
            if path.path=='/api/job':return self.send(200,app.job or {'state':'idle'})
            if path.path.startswith('/download/'):
                file=app.downloads.get(path.path[len('/download/'):])
                if not file or not file.is_file():return self.send(404,{'error':'Dosya bulunamadı.'})
                return self.send(200,file.read_bytes(),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',{'Content-Disposition':"attachment; filename*=UTF-8''"+quote(file.name)})
            return self.send(404,{'error':'Sayfa bulunamadı.'})
        except Exception as e:self.send(400,{'error':str(e)})
    def do_POST(self):
        if not self.authorized():return self.send(403,{'error':'Yetkisiz yerel istek.'})
        if self.headers.get('Origin') not in (None,self.server.origin):return self.send(403,{'error':'Kaynak adresi geçersiz.'})
        try:
            size=int(self.headers.get('Content-Length','0'))
            if not 0<size<=MAX_BODY:return self.send(413,{'error':'İstek boyutu geçersiz.'})
            if self.headers.get('Content-Type','').split(';')[0]!='application/json':return self.send(415,{'error':'JSON isteği gerekli.'})
            args=json.loads(self.rfile.read(size));path=urlsplit(self.path).path;app=self.server.application
            if not isinstance(args,dict):raise ValueError('İstek nesne olmalı.')
            if path=='/api/query':return self.send(200,app.query(args))
            if path=='/api/save-selection':
                app.data.selection(args);options=app.data.meta()['options']
                app.save_preference('filters',configuration.reconcile_selection(args,options,options));app.save_preference('filterOptions',options)
                return self.send(200,{'ok':True})
            if path=='/api/workspace':return self.send(200,{'quality':workspace.quality(app.data,args.get('state',{})),'review':app.workspace.review(),'history':app.workspace.history(),'views':app.preference('views:'+app.data.identity,{}),'currentJob':{k:app.job.get(k) for k in ('state','stage','cancelRequested')} if app.job else None})
            if path=='/api/trace':return self.send(200,workspace.trace(app.data,args.get('state',{}),args.get('department'),args.get('code')))
            if path=='/api/mapping-accept':return self.send(200,app.workspace.accept(args))
            if path=='/api/layout-preview':return self.send(200,configuration.layout_preview(app.data.settings,args.get('department')))
            if path=='/api/views':
                name=args.get('name','').strip()
                if not name or len(name)>80:raise ValueError('Görünüm adı 1–80 karakter olmalı.')
                views=app.preference('views:'+app.data.identity,{})
                if args.get('delete'):views.pop(name,None)
                else:
                    app.data.selection(args['state']);views[name]={'state':args['state'],'options':app.data.meta()['options']}
                app.save_preference('views:'+app.data.identity,views);return self.send(200,{'ok':True})
            if path=='/api/cancel':
                with app.lock:
                    if app.job and app.job['state']=='running':app.job['cancelRequested']=True
                return self.send(200,{'message':'Durdurma istendi; mevcut dosya işleminin bitmesi bekleniyor.'})
            if path=='/api/settings-preview':return self.send(200,configuration.preview(args['settings']))
            if path=='/api/job':return self.send(202,app.start_job(args['kind'],args.get('args',{})))
            if path in ('/api/comment-review','/api/comment-carry','/api/comment-apply'):
                if app.job and app.job['state']=='running':raise ValueError('Devam eden işlem tamamlandıktan sonra açıklamaları düzenleyin.')
                names,years,months=app.data.selection(args.get('state',{}))
                if years!=[app.data.budget_year] or not months:raise ValueError('Açıklama incelemesi için bütçe yılını tek başına ve en az bir ay seçin.')
                if path=='/api/comment-carry':
                    registry=app.preference('mapping:'+app.data.identity,{})
                    ids=registry.get('ids',{});active={ids[n]:n for n in app.data.departments if n in ids}
                    aliases={n:active[uid] for n,uid in ids.items() if uid in active}
                    result=app.notes.review.carry(app.data,months,args.get('paths'),aliases,departments=names)
                    return self.send(200,result)
                if path=='/api/comment-apply':
                    if any(r.get('department') not in names for r in args.get('entries',[])):raise ValueError('Seçim kapsamı dışındaki açıklama değiştirilemez.')
                    return self.send(200,app.notes.review.apply(app.data,months,args['entries']))
                return self.send(200,{'rows':[r for r in app.notes.review.rows(app.data,months) if r['department'] in names],'closing':list(app.data.closing),'months':months})
            if path=='/api/note':
                names,years,months=app.data.selection(args['state'])
                if years!=[app.data.budget_year] or args['department'] not in names:raise ValueError('Bütçe yılını tek başına ve ilgili bölümü seçin.')
                count=app.notes.save(app.data,months,[args]);return self.send(200,{'count':count})
            if path=='/api/pick':
                mode=args.get('mode','file')
                if mode not in ('file','files','folder'):raise ValueError('Geçersiz dosya seçimi.')
                result=subprocess.run([sys.executable,str(BASE/'dialog.py'),mode],capture_output=True,text=True,encoding='utf-8',timeout=600)
                if result.returncode:raise ValueError('Dosya penceresi açılamadı; yolu Kaynaklar bölümüne elle yazabilirsiniz.')
                return self.send(200,{'value':json.loads(result.stdout)})
            if path=='/api/shutdown':
                self.send(200,{'ok':True});threading.Thread(target=self.server.shutdown,daemon=True).start();return
            return self.send(404,{'error':'İşlem bulunamadı.'})
        except Exception as e:self.send(400,{'error':str(e)})

def main():
    parser=argparse.ArgumentParser(description='Yerel Python Bütçe Dashboard')
    parser.add_argument('--settings',type=Path);parser.add_argument('--closing',help='YYYY-AA')
    parser.add_argument('--no-browser',action='store_true');parser.add_argument('--port',type=int,default=0)
    parser.add_argument('--state-dir',type=Path)
    args=parser.parse_args()
    app=Application(args.settings,butce.period(args.closing) if args.closing else None,args.state_dir)
    server=Server(('127.0.0.1',args.port),app)
    url=server.origin+'/?token='+app.token
    print('Dashboard bu bilgisayarda açıldı. Kapatmak için uygulamadaki Kapat düğmesini kullanın.',flush=True)
    if args.no_browser:print(url,flush=True)
    else:webbrowser.open(url,new=1)
    try:server.serve_forever(poll_interval=.25)
    except KeyboardInterrupt:pass
    finally:server.server_close();app.executor.shutdown(wait=True)
if __name__=='__main__':main()
