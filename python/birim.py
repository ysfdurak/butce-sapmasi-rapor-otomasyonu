"""Endeks eşleşmeli birim maliyet: seçili ay oranları ve karşılaştırılabilir gider ağırlıkları."""
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as letter


def read_budget_counts(root, config, snapshots, closing):
    from butce import positioned_columns, number
    books={};result={}
    try:
        for name in config['bolumler']:
            snap=snapshots.get((closing,name))
            if not snap:continue
            path=Path(root)/config['maliyet_klasoru']/snap[3]
            if path not in books:books[path]=load_workbook(path,data_only=True)
            sheet=books[path][name];_,cols=positioned_columns(sheet,config['maliyet'],closing)
            result[name]={p:number(sheet.cell(config['maliyet']['islem_adedi_satiri'],c).value) for p,c in cols.items()}
    finally:
        for book in books.values():book.close()
    return result


def add_unit_cost(payload, config, indices):
    from butce import key, period_index
    known={k for k,p in indices}
    mapped={n:bool(m.get('endeks') and key(m['endeks']) in known) for n,m in config['bolumler'].items()}
    for row in payload['rows']:
        name,y,month=row[3],row[4],row[6];alias=config['bolumler'][name].get('endeks') or ''
        factor=period_index(indices,alias,[(y,month)]) if mapped[name] else None
        row.extend([float(factor) if factor is not None else None,int(factor is not None),int(mapped[name]),alias])
    payload['headers']+=['Bölüm aylık endeks çarpanı','Endeks çarpanı mevcut','Endeks eşleşti','Endeks bölüm adı']
    payload['unitDepartments']=[{'name':n,'manager':m['mudurluk']} for n,m in config['bolumler'].items()]
    payload['unitModel']=unit_model(payload)
    payload['unitLayout']=unit_layout(payload,[payload['budgetYear']])
    return payload


def unit_model(p):
    n=len(p['rows'])+1
    def s(c):return f"'Veri'!${c}$2:${c}${n}"
    vis=s('L');ys=s('E');dept=s('D')
    cells={'A1':'Bütçe yılı','B1':p['budgetYear'],'A2':'Yıl sayısı','B2':len(p['years']),'A3':'Bölüm sayısı','B3':len(p['unitDepartments']),
           'A4':'Durum','B4':'=IF(Hesap!B3=0,"Seçime uygun veri yok.","TL / işlem · Endeks eşleşmeli ve karşılaştırılabilir bölümler")',
           'P8':'Yıl','Q8':'Karşılaştırılabilir gerçekleşen gider TL','R8':'Ağırlıklı reel iyileşme','S8':'Seçili kayıt','T8':'Ağırlık toplamı',
           'A8':'Bölüm','B8':'Yıl','C8':'Gerçekleşen BM TL/işlem','D8':'Önceki yıl BM TL/işlem','E8':'Bütçe BM TL/işlem',
           'F8':'Dönem endeks çarpanı','G8':'Endekslenmiş önceki BM','H8':'Reel iyileşme %','I8':'Gerçekleşen gider TL',
           'J8':'Karşılaştırılabilir','K8':'Gider ağırlığı','L8':'Ağırlıklı reel iyileşme','M8':'Görünür satır','N8':'Enflasyon %'}
    last=10+len(p['unitDepartments'])*len(p['years'])-1
    for yi,y in enumerate(p['years']):
        rr=9+yi;cells[f'P{rr}']=y;cells[f'S{rr}']=f'=SUMIFS({vis},{ys},{y})'
        cells[f'Q{rr}']=f'=SUMIFS($I$10:$I${last},$B$10:$B${last},{y},$J$10:$J${last},1)'
        cells[f'R{rr}']=f'=IF(AND({y}=$B$1,Q{rr}>0),SUMIFS($L$10:$L${last},$B$10:$B${last},{y},$J$10:$J${last},1),"")'
        cells[f'T{rr}']=f'=IF(Q{rr}>0,SUMIFS($K$10:$K${last},$B$10:$B${last},{y},$J$10:$J${last},1),"")'
        for di,d in enumerate(p['unitDepartments']):
            r=10+di*len(p['years'])+yi;name=d['name'].replace('"','""')
            crit=f'{vis},1,{ys},{y},{dept},"{name}"'
            count=f'SUMIFS({vis},{ys},{y},{dept},"{name}")'
            cells[f'A{r}']=d['name'];cells[f'B{r}']=y
            for col,cost,cv,count_col,flag,gate in [('C','H','J','AI','AK',''),('D','AO','AP','AM','AN',''),('E','I','K','AJ','AL',f'{y}=$B$1,')]:
                ok=f'AND({gate}{count}>0,SUMIFS({s(cv)},{crit})={count},SUMIFS({s(flag)},{crit})={count},COUNTIFS({s(count_col)},"<=0",{crit})=0,SUMIFS({s(count_col)},{crit})>0)'
                cells[f'{col}{r}']=f'=IF({ok},SUMIFS({s(cost)},{crit})/SUMIFS({s(count_col)},{crit}),"")'
            cells[f'I{r}']=f'=IF(AND({count}>0,SUMIFS({s("J")},{crit})={count}),SUMIFS({s("H")},{crit}),"")'
            cells[f'F{r}']=f'=IF(AND({count}>0,SUMIFS({s("AR")},{crit})={count}),SUMIFS({s("AQ")},{crit})/{count},"")'
            cells[f'G{r}']=f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(F{r})),D{r}*F{r},"")'
            cells[f'H{r}']=f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(G{r}),G{r}>0),(G{r}-C{r})/G{r},"")'
            cells[f'N{r}']=f'=IF(ISNUMBER(F{r}),F{r}-1,"")'
            mapped=f'SUMIFS({s("AS")},{crit})={count}'
            cells[f'J{r}']=f'=IF(AND({y}=$B$1,{count}>0,{mapped},ISNUMBER(C{r}),C{r}>=0,ISNUMBER(D{r}),D{r}>0,ISNUMBER(H{r}),ISNUMBER(I{r}),I{r}>=0),1,0)'
            cells[f'K{r}']=f'=IF(AND(J{r}=1,$Q${rr}>0),I{r}/$Q${rr},"")'
            cells[f'L{r}']=f'=IF(AND(ISNUMBER(K{r}),ISNUMBER(H{r})),K{r}*H{r},"")'
            cells[f'M{r}']=f'=IF({y}=$B$1,J{r},IF(AND({count}>0,{mapped},ISNUMBER(C{r})),1,0))'
    return cells


def unit_layout(p, years):
    active=p['budgetYear'] in years;fields=[]
    for y in sorted(years):
        if active and y in [p['budgetYear']-1,p['budgetYear']]:continue
        fields.append((y,'C',f'{y} Gerçekleşen BM'))
    if active:
        y=p['budgetYear']
        fields += [(y,c,label) for c,label in [('D',f'{y-1} Gerçekleşen BM'),('C',f'{y} Gerçekleşen BM'),('E',f'{y} Bütçe BM'),('N',f'{y-1}–{y} Enflasyon %'),('G',f'Endekslenmiş {y-1} BM'),('H','Reel iyileşme %'),('K',f'{y} Gider ağırlığı %')]]
    eligible=[]
    for di,d in enumerate(p['unitDepartments']):
        current=[r for r in p['rows'] if r[3]==d['name'] and r[4]==p['budgetYear'] and r[6] in p['months']]
        history=[r for r in p['rows'] if r[3]==d['name'] and r[4] in years and r[6] in p['months']]
        if active:
            ok=bool(current) and all(r[9] and r[36] and r[34]>0 and r[41] and r[39] and r[38]>0 and r[43] and r[44] for r in current) and sum(r[40] for r in current)>0 and sum(r[7] for r in current)>=0
        else:ok=any((subset:=[r for r in history if r[4]==year]) and all(r[44] and r[9] and r[36] and r[34]>0 for r in subset) for year in years)
        if ok:eligible.append((di,d))
    headers=[];labels=[];cells=[]
    for i,(di,d) in enumerate(eligible):labels.append({'range':f'B{15+i}:F{15+i}','value':d['name'],'index':di})
    for i,(y,c,label) in enumerate(fields):
        col=7+3*i;headers.append({'range':f'{letter(col)}13:{letter(col+2)}14','label':label,'percent':c in ['N','H','K']})
        for j,(di,d) in enumerate(eligible):
            ref=f"'Birim Hesap'!{c}{10+di*len(p['years'])+p['years'].index(y)}"
            cells.append({'range':f'{letter(col)}{15+j}:{letter(col+2)}{15+j}','formula':f'=IF(ISNUMBER({ref}),{ref},"")','percent':c in ['N','H','K'],'colored':c=='H'})
    return {'headers':headers,'labels':labels,'cells':cells,'lastCol':letter(6+3*len(fields)),'lastRow':14+len(eligible),
            'cardRow':17+len(eligible),'cardRef':f"'Birim Hesap'!R{9+p['years'].index(p['budgetYear'])}",'active':active}
