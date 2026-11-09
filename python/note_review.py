"""Previous-closing comments as unconfirmed, bulk-reviewable local drafts."""
from pathlib import Path
from datetime import datetime
from decimal import Decimal
import json,hashlib
from openpyxl import load_workbook
import butce

PENDING='[Kontrol bekliyor — önceki dönemden]'
def previous(closing):
    y,m=closing
    return (y,m-1) if m>1 else (y-1,12)
def stamp():return datetime.now().isoformat(timespec='microseconds')
def checksum(value):return hashlib.sha256(json.dumps(value,ensure_ascii=False,sort_keys=True,default=str).encode()).hexdigest()
def literal(cell,value):cell.value=value;cell.number_format='@';cell.data_type='s'
def report_text(value):
    if value.get('status')=='pending':return f'{PENDING} {value["previousClosing"]} · {value["previousMonths"]}\n{value["note"]}'
    return value.get('note','')

class Review:
    def __init__(self,notes):
        self.notes=notes
        with notes.connect() as c:c.execute('''CREATE TABLE IF NOT EXISTS note_review (
        dataset TEXT, closing TEXT, months TEXT, department TEXT, code INTEGER,
        previous_closing TEXT, previous_months TEXT, old_note TEXT, draft TEXT,
        status TEXT, old_actual TEXT, old_budget TEXT, source TEXT, updated TEXT,
        PRIMARY KEY(dataset,closing,months,department,code))''')
    def _manual(self,c,data,months):
        return {(r[0],r[1]):r for r in c.execute('SELECT department,code,note,actual,budget,source,updated FROM notes WHERE dataset=? AND closing=? AND months=?',self.notes.key(data,months))}
    def rows(self,data,months,connection=None):
        if connection is None:
            with self.notes.connect() as c:return self.rows(data,months,c)
        manual=self._manual(connection,data,months);result=[];seen=set()
        for row in connection.execute('SELECT department,code,previous_closing,previous_months,old_note,draft,status,old_actual,old_budget,source,updated FROM note_review WHERE dataset=? AND closing=? AND months=? ORDER BY department,code',self.notes.key(data,months)):
            name,code,oldclose,oldmonths,oldnote,draft,status,olda,oldb,source,updated=row
            if name not in data.departments or code not in data.catalog:continue
            a=data.aggregate([name],[data.budget_year],months,code=code);b=data.aggregate([name],[data.budget_year],months,'budget',code=code)
            own=manual.get((name,code));protected=bool(own and own[5]!='Geçmiş yorum incelemesi: '+oldclose)
            stored_status=status
            stale=bool(status=='approved' and own and (a is None or b is None or Decimal(own[3])!=a or Decimal(own[4])!=b))
            if stale:status='pending'
            if protected:draft=own[2]
            same=oldmonths==butce.month_signature(months)
            tolerance=Decimal(str(data.config['kontrol']['tolerans']))
            changed=not same or olda is None or oldb is None or a is None or b is None or abs(Decimal(olda)-a)>tolerance or abs(Decimal(oldb)-b)>tolerance
            token=checksum([row,own,str(a),str(b),data.catalog[code]])
            result.append({'department':name,'manager':data.config['bolumler'][name]['mudurluk'],'code':code,'item':data.catalog[code][0],'oldNote':oldnote,'draft':draft,'status':status,'storedStatus':stored_status,'stale':stale,'previousClosing':oldclose,'previousMonths':oldmonths,'months':butce.month_signature(months),'sameMonths':same,'dataChanged':changed,'actual':a,'budget':b,'oldActual':Decimal(olda) if olda else None,'oldBudget':Decimal(oldb) if oldb else None,'source':source,'updated':updated,'token':token,'protected':protected,'currentNote':own[2] if own else ''})
            seen.add((name,code))
        for (name,code),own in manual.items():
            if (name,code) not in seen and name in data.departments and code in data.catalog:result.append({'department':name,'manager':data.config['bolumler'][name]['mudurluk'],'code':code,'item':data.catalog[code][0],'oldNote':'','draft':own[2],'status':'current','previousClosing':'','previousMonths':'','sameMonths':True,'dataChanged':False,'actual':data.aggregate([name],[data.budget_year],months,code=code),'budget':data.aggregate([name],[data.budget_year],months,'budget',code=code),'source':own[5],'protected':True,'currentNote':own[2],'token':checksum(own)})
        return result
    def carry(self,data,months,paths=None,aliases=None,departments=None):
        from workspace import inventory
        if inventory(data)!=data.fingerprint:raise ValueError('Kaynak dosyaları değişmiş; önce kaynakları yenileyin.')
        months=sorted(set(months))
        if not months or any(type(m) is not int or not 1<=m<=12 for m in months):raise ValueError('En az bir geçerli ay seçin.')
        oldclose=butce.label(previous(data.closing));candidates={};issues=[]
        def add(name,code,record):
            if departments is not None and name not in departments:return
            if name not in data.departments or code not in data.catalog:
                issues.append(f'Eşleşmeyen bölüm/gider: {name} / {code}');return
            key=name,code
            if key in candidates and candidates[key]!=record:raise ValueError('Aynı bölüm/gidere farklı geçmiş açıklamalar seçildi; hiçbir kayıt yüklenmedi.')
            candidates[key]=record
        if paths:
            identities=set()
            for path in paths:
                book=load_workbook(path,data_only=False)
                try:
                    if '_Rapor Kimliği' not in book:raise ValueError('Rapor kimliği yok: '+Path(path).name)
                    kind=book['_Rapor Kimliği']['B1'].value
                    if kind!='Müdürlük':raise ValueError('Geçmiş yorumlar için açıklamalı müdürlük raporlarını seçin.')
                    identity,members,cols,items=butce.explanation_layout(book,kind)
                    if identity[1]!=oldclose or str(identity[0])[:4]!=oldclose[:4]:raise ValueError('Yalnız bir önceki takvim kapanışının raporları kullanılabilir: '+oldclose)
                    identities.add(identity[2])
                    for (code,group,label),r in items.items():
                        if code not in data.catalog or data.catalog[code]!=(label,group):continue
                        for original in members:
                            name=(aliases or {}).get(original,original);cell=book['Giderler'].cell(r,cols[original]+4);value=cell.value
                            if value in (None,''):continue
                            if cell.data_type=='f' or not isinstance(value,str) or len(value)>32767:raise ValueError('Geçmiş açıklama sabit metin olmalı.')
                            if value.startswith('[Kontrol bekliyor') and '\n' in value:
                                issues.append(f'{name} / {code}: Önceki dosyadaki yorum da kontrol bekliyordu.')
                                value=value.split('\n',1)[1]
                            a=butce.number(book['Giderler'].cell(r,cols[original]).value);b=butce.number(book['Giderler'].cell(r,cols[original]+1).value)
                            add(name,code,{'text':value,'months':identity[2],'actual':str(a) if a is not None else None,'budget':str(b) if b is not None else None,'source':str(Path(path).resolve())})
                finally:book.close()
            if len(identities)>1:raise ValueError('Seçilen geçmiş raporların ay kümeleri farklı; tek ay kümesi seçin.')
        else:
            with self.notes.connect() as c:
                bykey={}
                for name,code,text,ms,a,b,source in c.execute('SELECT department,code,note,months,actual,budget,source FROM notes WHERE dataset=? AND closing=? AND note<>\'\'',(data.identity,oldclose)):
                    prior_review=c.execute('SELECT status FROM note_review WHERE dataset=? AND closing=? AND months=? AND department=? AND code=?',(data.identity,oldclose,ms,name,code)).fetchone()
                    if prior_review and prior_review[0]!='approved' and source=='Geçmiş yorum incelemesi: '+butce.label(previous(previous(data.closing))):
                        issues.append(f'{name} / {code}: Önceki kapanışın yorumu henüz onaylı değil; ilgili rapordan getirin.');continue
                    bykey.setdefault((name,code),[]).append({'text':text,'months':ms,'actual':a,'budget':b,'source':source})
                for (name,code),records in bykey.items():
                    exact=[r for r in records if r['months']==butce.month_signature(months)]
                    if exact:add(name,code,exact[0])
                    elif len(records)==1:add(name,code,records[0])
                    else:issues.append(f'{name} / {code}: Önceki kapanışta birden fazla ay seçimi var; ilgili rapor dosyasını seçin.')
        created=skipped=0
        with self.notes.connect() as c:
            c.execute('BEGIN IMMEDIATE');manual=self._manual(c,data,months)
            for (name,code),record in candidates.items():
                key=(*self.notes.key(data,months),name,code)
                if (name,code) in manual or c.execute('SELECT 1 FROM note_review WHERE dataset=? AND closing=? AND months=? AND department=? AND code=?',key).fetchone():skipped+=1;continue
                c.execute('INSERT INTO note_review VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)',(*key,oldclose,record['months'],record['text'],record['text'],'pending',record['actual'],record['budget'],record['source'],stamp()));created+=1
        if not candidates and not paths:issues.append('Önceki kapanışta kayıtlı açıklama yok; açıklamalı müdürlük dosyalarını seçebilirsiniz.')
        return {'created':created,'skipped':skipped,'issues':issues,'previousClosing':oldclose}
    def apply(self,data,months,entries):
        from workspace import inventory
        if inventory(data)!=data.fingerprint:raise ValueError('Kaynak dosyaları değişmiş; önce kaynakları yenileyin.')
        if not isinstance(entries,list) or len(entries)>5000:raise ValueError('Geçersiz toplu işlem.')
        with self.notes.connect() as c:
            c.execute('BEGIN IMMEDIATE');current={(r['department'],r['code']):r for r in self.rows(data,months,c)};seen=set();staged=[]
            for entry in entries:
                key=entry.get('department'),entry.get('code');row=current.get(key)
                if key in seen or not row or row['protected'] or row['status']=='current':raise ValueError('Kayıt yineleniyor veya bu dönemin açıklaması korunuyor; listeyi yenileyin.')
                seen.add(key)
                if entry.get('token')!=row['token']:raise ValueError('Açıklama veya tutarlar değişti; listeyi yenileyin. Hiçbir kayıt değiştirilmedi.')
                action=entry.get('action');text=entry.get('draft',row['draft'])
                if action not in ('draft','approve','reject') or not isinstance(text,str) or len(text)>32000:raise ValueError('Açıklama en çok 32.000 karakter olmalı; işlem geçersiz.')
                if action=='approve' and (not text.strip() or row['actual'] is None or row['budget'] is None):raise ValueError('Boş açıklama veya eksik tutar onaylanamaz.')
                if action=='draft' and row['status']=='approved':raise ValueError('Onaylı metni değiştirmek için düzenleyip tekrar onaylayın.')
                staged.append((entry,row,text))
            for entry,row,text in staged:
                action=entry['action'];key=(*self.notes.key(data,months),row['department'],row['code']);when=stamp()
                status={'draft':'pending','approve':'approved','reject':'rejected'}[action]
                c.execute('UPDATE note_review SET draft=?,status=?,updated=? WHERE dataset=? AND closing=? AND months=? AND department=? AND code=?',(text,status,when,*key))
                if action=='approve':c.execute('INSERT OR REPLACE INTO notes VALUES (?,?,?,?,?,?,?,?,?,?)',(*key,text,str(row['actual']),str(row['budget']),'Geçmiş yorum incelemesi: '+row['previousClosing'],when))
                elif action=='reject':c.execute('DELETE FROM notes WHERE dataset=? AND closing=? AND months=? AND department=? AND code=?',key)
        return {'count':len(staged)}
    def effective(self,data,months):
        result={}
        for row in self.rows(data,months):
            if row['status']=='pending' and not row['protected'] and row['draft'].strip():result[row['department'],row['code']]={'note':row['draft'],'status':'pending','previousClosing':row['previousClosing'],'previousMonths':row['previousMonths'],'source':row['source'],'stale':row.get('stale',False),'updated':row['updated']}
        return result
