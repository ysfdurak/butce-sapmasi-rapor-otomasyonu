"""Power Query rapor şablonunu üretir: powerquery/Butce_Rapor_Sablonu.xlsx

Şablon makro içermez ve Excel 2016 ile uyumludur (yalnız SUMIFS, COUNTIFS, INDEX, MATCH, SUMPRODUCT, RANK gibi
klasik işlevler). Power Query sorgusu (powerquery/Butce_Veri.pq) tek tabloyu "Veri" sayfasının A1 hücresine
yükler; tüm rapor sayfaları bu tablodan formüllerle hesaplanır. Kapsam, rapor yılı, ay seçimi ve eşikler
Kontrol Paneli'nden değiştirilir ve yenileme gerektirmeden anında uygulanır.

Kullanım:
    python tools/pq_sablon.py                    # boş şablon (Veri sayfası sorgu yüklemesini bekler)
    python tools/pq_sablon.py --ornek CIKTI.xlsx # örnek veriyle dolu kopya (test ve tanıtım için)
"""

import argparse
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))
from pq_veri import COLUMNS  # noqa: E402

MAXD, MAXL, MAXK, MAXI, MAXG, MAXR, MAXW, MAXM, MAXA = 60, 120, 2000, 80, 60, 1500, 1000, 500, 500
BLOCKS = 40  # Giderler sayfasındaki bölüm bloğu sayısı
MAXO, MAXF = 2000, 200  # Açıklama Onayı satırı, girdi kaydı satırı
# Açıklama Onayı sayfası: Butce_Aciklama sorgusunun sütunları (A..W) ve sağındaki formül sütunları (Y..AA)
ONAY_COLS = ["Anahtar", "Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Açıklama", "Yazan", "Yazma tarihi",
             "Kapanış", "Aylar", "Gerçekleşen (dosyada)", "Bütçe (dosyada)", "Kaynak dosya", "Dosya tarihi",
             "Okuma notu", "Kaynakta", "Durum", "Onay notu", "Son işlem", "İşlemi yapan", "İşlem zamanı",
             "İşlem geçmişi", "Kayıtlı durum"]
OC = {name: L(i) for i, name in enumerate(ONAY_COLS, 1)}
ONAY = "'Açıklama Onayı'"
DURUMLAR = ["Bekliyor", "Onaylandı", "Reddedildi", "Bu döneme aktar"]
# _Bolum sayfasında aylık sütun blokları (1. ayın sütun numarası - 1)
RATIO0, DIP0, DIPMISS0 = 54, 66, 78  # BC..BN endeks oranları, BO..BZ aylık dip toplam, CA..CL dip eksik bayrağı

# ----------------------------------------------------------------------------- Veri tablosu sütunları
V = {name: L(i) for i, name in enumerate(COLUMNS, 1)}
VKEY = "Veri!$A:$A"
VG = "Veri!$M:$X"  # G1..G12
VB = "Veri!$Y:$AJ"  # B1..B12


def vcol(name):
    return f"Veri!${V[name]}:${V[name]}"


def find(key_expr):
    return f'IFERROR(MATCH({key_expr},{VKEY},0),"")'


# ----------------------------------------------------------------------------- Ortak hesap hücreleri ('_H')
H = "'_H'"
KAPYIL, KAPAY, YIL, BUTCEYILI, KAPSAM, SELCOUNT = (f"{H}!$B${r}" for r in (2, 3, 4, 5, 6, 7))
ESIK, ORAN, OZETN, TOLERANS = (f"{H}!$B${r}" for r in (8, 9, 10, 11))
FLAGS, YTDF, RESTF = f"{H}!$C$14:$N$14", f"{H}!$C$15:$N$15", f"{H}!$C$16:$N$16"
SELLIST = f"{H}!$C$18:$N$18"
P = "'Kontrol Paneli'"


def sel(r, rng=VG, flags=FLAGS):
    """Seçili ayların toplamı; seçili aylardan biri boşsa boş (Python 'total' kuralı)."""
    return (
        f'IF({r}="","",IF(SUMPRODUCT({flags}*ISNUMBER(INDEX({rng},{r},0)))<SUM({flags}),"",'
        f"SUMPRODUCT({flags}*INDEX({rng},{r},0))))"
    )


def pos(r, rng=VG, flags=FLAGS):
    """İşlem adedi gibi pozitif olması gereken toplam: seçili aylardan biri boş/sıfır/negatifse boş."""
    return (
        f'IF({r}="","",IF(SUMPRODUCT({flags}*(INDEX({rng},{r},0)>0))<SUM({flags}),"",'
        f"SUMPRODUCT({flags}*INDEX({rng},{r},0))))"
    )


def num(*cells):
    return "AND(" + ",".join(f"ISNUMBER({c})" for c in cells) + ")"


# ----------------------------------------------------------------------------- Biçimler
MONEY = "#,##0.00;-#,##0.00;0.00"
COUNT = "#,##0"
PCT = "0.00%"
HEADER_FILL = PatternFill("solid", fgColor="E7EEF2")
INPUT_FILL = PatternFill("solid", fgColor="FFF4DA")
TITLE = Font(size=16, bold=True, color="233C57")
NOTE = Font(italic=True, color="5B6B75")
BOLD = Font(bold=True, color="263746")
RED_FILL, GREEN_FILL = PatternFill("solid", fgColor="FFC7CE"), PatternFill("solid", fgColor="C6EFCE")
MONTHS = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]


def header(ws, row, titles, col=1, height=34):
    for i, title in enumerate(titles):
        cell = ws.cell(row, col + i, title)
        cell.fill = HEADER_FILL
        cell.font = BOLD
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = height


def widths(ws, values):
    for col, width in values.items():
        ws.column_dimensions[col].width = width


def fmt(ws, cols, first, last, number_format):
    for c in cols:
        for r in range(first, last + 1):
            ws[f"{c}{r}"].number_format = number_format


def direction_colors(ws, ref):
    first = ref.split(":")[0]
    ws.conditional_formatting.add(ref, FormulaRule(formula=[f'{first}="Aleyhte"'], font=Font(color="BE303A", bold=True)))
    ws.conditional_formatting.add(ref, FormulaRule(formula=[f'{first}="Lehte"'], font=Font(color="1B754C", bold=True)))


def total_rows_style(ws, ref, level_cell):
    first_col = ref.split(":")[0]
    row = "".join(ch for ch in first_col if ch.isdigit())
    col = "".join(ch for ch in level_cell if ch.isalpha())
    ws.conditional_formatting.add(
        ref, FormulaRule(formula=[f'${col}{row}="Müdürlük toplamı"'], fill=PatternFill("solid", bgColor="EFF3F5"),
                         font=Font(bold=True))
    )
    ws.conditional_formatting.add(
        ref, FormulaRule(formula=[f'OR(${col}{row}="Direktörlük toplamı",${col}{row}="Genel toplam")'],
                         fill=PatternFill("solid", bgColor="D8E4E8"), font=Font(bold=True))
    )


# ============================================================================= Sayfalar
def build_panel(wb):
    ws = wb.active
    ws.title = "Kontrol Paneli"
    ws["A1"] = "Bütçe sapması raporu · Kontrol Paneli"
    ws["A1"].font = TITLE
    ws["A2"] = ("Kaynak dosyalar değiştiğinde: Veri › Tümünü Yenile. Aşağıdaki seçimler (sarı hücreler) değişince "
                "tüm sayfalar yenileme gerekmeden anında güncellenir.")
    ws["A2"].font = NOTE
    ws["B3"] = "Dosya"
    ws["B3"].font = BOLD
    ws["C3"] = (f'=IF({H}!$B$12="","Direktörlük raporu","Müdürlük dosyası: "&{H}!$B$12&" (yalnız bu müdürlüğün verisi)")')
    ws["B4"] = "Kaynak kapanış"
    ws["C4"] = (f'=IF({KAPYIL}="","Veri yok: önce sorguyu yükleyin (Veri › Tümünü Yenile)",'
                f'INDEX(\'_H\'!$C$19:$N$19,1,{KAPAY})&" "&{KAPYIL})')
    labels = {
        5: ("Rapor yılı", "Boş = kapanış yılı. Önceki bir yıl seçilirse bütçe karşılaştırması yapılmaz."),
        6: ("Kapsam", "Tümü = direktörlük raporu; bir müdürlük seçilirse o müdürlüğün raporu."),
        8: ("Ay seçimi", "Elle seçimde aşağıdaki satırda ayların altına 1 yazın."),
        13: ("Önemlilik tutarı (TL)", "Boş = Ayarlar.xlsx (Rapor › onemlilik_tutari) veya 1.000.000."),
        14: ("Önemlilik oranı", "Boş = Ayarlar.xlsx veya %10. Tutar VE oran birlikte aşılmalı."),
        15: ("Özet kalem sayısı", "Yönetici Özeti'nde gösterilecek en büyük aleyhte/lehte kalem sayısı (en çok 10)."),
    }
    for r, (label, note) in labels.items():
        ws.cell(r, 2, label).font = BOLD
        if r not in (8,):
            ws.cell(r, 3).fill = INPUT_FILL
        ws.cell(r, 16 if r == 8 else 4, note).font = NOTE
    ws["C6"] = "Tümü"
    ws["C8"] = "Yıl başından kapanışa"
    ws["C8"].fill = INPUT_FILL
    dv = DataValidation(type="list", formula1='"Yıl başından kapanışa,Yalnız kapanış ayı,Elle seçim"')
    ws.add_data_validation(dv)
    dv.add("C8")
    scope = DataValidation(type="list", formula1=f"'_Liste'!$A$1:$A${MAXD + 1}")
    ws.add_data_validation(scope)
    scope.add("C6")
    ws["B9"] = "Ay"
    ws["B10"] = "Elle seçim (1)"
    ws["B11"] = "Uygulanan"
    for m in range(12):
        c = 3 + m
        ws.cell(9, c, MONTHS[m][:3]).font = BOLD
        ws.cell(10, c).fill = INPUT_FILL
        ws.cell(11, c, f'=IF(INDEX({FLAGS},1,{m + 1})=1,"✓","")').alignment = Alignment(horizontal="center")
    manual = DataValidation(type="list", formula1='"1"', allow_blank=True)
    ws.add_data_validation(manual)
    manual.add("C10:N10")
    ws["B16"] = "Onaylayan (adınız)"
    ws["B16"].font = BOLD
    ws["C16"].fill = INPUT_FILL
    ws["D16"] = "Açıklama Onayı'nda yaptığınız işlemler bu adla kaydedilir."
    ws["D16"].font = NOTE
    ws["C13"].number_format = COUNT
    ws["C14"].number_format = "0%"
    ws["B17"] = "Durum"
    ws["B17"].font = BOLD
    ws["B18"] = "Veri satırı"
    ws["C18"] = '=MAX(0,COUNTA(Veri!$A:$A)-1)'
    ws["B19"] = "Uyarı"
    ws["C19"] = f"=COUNT('Uyarılar'!$A$2:$A${MAXW + MAXD + MAXG + 3})"
    ws["B20"] = "Önemli sapma"
    ws["C20"] = "=COUNT('_Kalem'!$S$2:$S$2001)"
    ws["B21"] = "Açıklaması eksik önemli sapma"
    ws["C21"] = "=COUNTIF('Önemli Sapmalar'!$K$2:$K$501,\"Açıklama gerekli\")"
    ws["B22"] = "Doğrulanamayan kontrol"
    ws["C22"] = ("=COUNTIF('Veri Kontrolü'!$H$4:$H$723,\"?*\")-COUNTIF('Veri Kontrolü'!$H$4:$H$723,\"Doğrulandı\")")
    ws["B23"] = "Seçim uyarısı"
    ws["C23"] = (f'=IF({KAPYIL}="","",IF({YIL}>{KAPYIL},"Rapor yılı kapanıştan sonra olamaz",'
                 f'IF({SELCOUNT}=0,"En az bir ay seçin (kapanış ayından sonraki aylar seçilemez)","Seçimler geçerli")))')
    ws["B24"] = "Son yenileme"
    ws["C24"] = f'=IFERROR(INDEX({vcol("Metin")},MATCH("Kimlik",{VKEY},0))&"","")'
    ws["B25"] = "Sayfalar"
    ws["B25"].font = BOLD
    guide = [
        ("Yönetici Özeti", "Toplamlar, yön, yıl sonu tahmini, kontrol durumu, en büyük önemli sapmalar"),
        ("Pano", "Müdürlük / bölüm / gider grubu açılır listeleri, tablolar ve grafikler"),
        ("Genel Özet", "Bölüm, müdürlük ve direktörlük toplamları; önceki yıl, bütçe, OPEX / ciro"),
        ("Giderler", "Gider grupları ve kalemleri × bölümler; açıklamalar Açıklamalar sayfasından"),
        ("Birim Maliyet", "İşlem başına maliyet, endekslenmiş önceki yıl, reel iyileşme ve ağırlıklar"),
        ("Önemli Sapmalar", "Eşiği aşan kalemler ve açıklama durumu"),
        ("Yıl Sonu Tahmini", "Gerçekleşen + kalan bütçe ve aylık ortalama gidiş"),
        ("Hacim ve Birim Maliyet", "Bütçe farkının hacim ve birim maliyet etkisine ayrıştırılması"),
        ("Veri Kontrolü", "Kontrol dosyasıyla dip toplam mutabakatı ve ciro"),
        ("Revizyonlar · Uyarılar", "Önceki kapanışa göre değişen tutarlar; okuma ve veri uyarıları"),
        ("Açıklamalar", "Açıklama girişi: müdürlük dosyasında müdürlük, ana raporda finans yazar"),
        ("Açıklama Onayı", "Müdürlüklerden dönen açıklamalar (Butce_Aciklama sorgusu), onay ve işlem geçmişi"),
        ("Dağıtım", "Müdürlük dosyalarını hazırlama listesi"),
        ("Girdi Kaydı", "Raporun okuduğu dosyalar, boyutları ve değiştirilme zamanları"),
    ]
    for i, (name, text) in enumerate(guide, 26):
        ws.cell(i, 2, name).font = Font(bold=True, color="233C57")
        ws.cell(i, 3, text)
    widths(ws, {"A": 2, "B": 30, "C": 22, "D": 14})
    for c in range(4, 16):
        ws.column_dimensions[L(c)].width = 6 if c < 15 else 6
    ws.column_dimensions["D"].width = 8
    ws.sheet_view.showGridLines = False
    for name, ref in (("Kullanici", "'Kontrol Paneli'!$C$16"), ("Rapor_Donemi", "'_H'!$B$21"),
                      ("Rapor_Aylari", "'_H'!$B$20")):
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
    return ws


def build_helpers(wb):
    h = wb.create_sheet("_H")
    rows = {
        2: ("Kapanış yılı", f'=IFERROR(INDEX({vcol("Yıl")},MATCH("Kapanis",{VKEY},0)),"")'),
        3: ("Kapanış ayı", f'=IFERROR(INDEX({vcol("Sayı")},MATCH("Kapanis",{VKEY},0)),"")'),
        4: ("Rapor yılı", f'=IF({P}!$C$5="",B2,{P}!$C$5)'),
        5: ("Bütçe yılı", "=IF(B4=B2,1,0)"),
        6: ("Kapsam", f'=IF(B12<>"",B12,IF({P}!$C$6="","Tümü",{P}!$C$6))'),
        7: ("Seçili ay sayısı", "=SUM(C14:N14)"),
        8: ("Önemlilik tutarı", f'=IF({P}!$C$13<>"",{P}!$C$13,IFERROR(INDEX({vcol("Sayı")},MATCH("Ayar|onemlilik_tutari",{VKEY},0)),1000000))'),
        9: ("Önemlilik oranı", f'=IF({P}!$C$14<>"",{P}!$C$14,IFERROR(INDEX({vcol("Sayı")},MATCH("Ayar|onemlilik_orani",{VKEY},0)),0.1))'),
        10: ("Özet kalem sayısı", f'=MIN(10,IF({P}!$C$15<>"",{P}!$C$15,IFERROR(INDEX({vcol("Sayı")},MATCH("Ayar|ozet_kalem_sayisi",{VKEY},0)),5)))'),
        11: ("Tolerans", f'=IFERROR(INDEX({vcol("Sayı")},MATCH("Ayar|tolerans",{VKEY},0)),0.01)'),
        12: ("Müdürlük dosyası", f'=IFERROR(INDEX({vcol("Müdürlük")},MATCH("Kimlik",{VKEY},0))&"","")'),
    }
    for r, (label, formula) in rows.items():
        h.cell(r, 1, label)
        h.cell(r, 2, formula)
    h["A13"], h["A14"], h["A15"], h["A16"], h["A17"], h["A18"], h["A19"] = (
        "Ay", "Seçili", "Ocak–kapanış", "Kapanış sonrası", "Birikimli", "i. seçili ay", "Ay adı")
    for m in range(1, 13):
        c = L(2 + m)
        h[f"{c}13"] = m
        h[f"{c}14"] = (f'=IF($B$2="",0,IF(AND($B$4=$B$2,{c}$13>$B$3),0,IF({P}!$C$8="Elle seçim",'
                       f'IF({P}!{L(2 + m)}$10=1,1,0),IF({P}!$C$8="Yalnız kapanış ayı",IF({c}$13=$B$3,1,0),'
                       f'IF({c}$13<=$B$3,1,0)))))')
        h[f"{c}15"] = f'=IF($B$3="",0,IF({c}$13<=$B$3,1,0))'
        h[f"{c}16"] = f"=1-{c}15"
        h[f"{c}17"] = f"=IF({c}14=1,SUM($C$14:{c}14),0)"
        h[f"{c}18"] = f'=IFERROR(MATCH({m},$C$17:$N$17,0),"")'
        h[f"{c}19"] = MONTHS[m - 1]
    # Açıklama anahtarı: kapanış ve seçili aylar ("2026:1,2,3"). Açıklama yalnız aynı seçim için geçerlidir.
    h["A20"], h["A21"] = "Aylar imzası", "Kapanış etiketi"
    h["B20"] = ('=IF($B$2="","",$B$4&":"&SUBSTITUTE(TRIM(' + "&".join(f'IF({L(2 + m)}14=1,"{m} ","")' for m in range(1, 13))
                + ')," ",","))')
    h["B21"] = '=IF($B$2="","",$B$2&"-"&RIGHT("0"&$B$3,2))'
    h.sheet_state = "hidden"

    lst = wb.create_sheet("_Liste")
    lst["A1"], lst["B1"], lst["C1"] = "Tümü", "Tümü", "Tümü"
    for i in range(1, MAXD + 1):
        lst[f"A{i + 1}"] = f'=IFERROR(INDEX({vcol("Müdürlük")},MATCH("Mudurluk|{i}",{VKEY},0)),"")'
        lst[f"B{i + 1}"] = f'=IFERROR(INDEX({vcol("Ad")},MATCH("GrupAd|{i}",{VKEY},0)),"")'
        lst[f"C{i + 1}"] = f"=IFERROR(INDEX('_Bolum'!$C$2:$C${MAXD + 1},MATCH({i},'_Bolum'!$J$2:$J${MAXD + 1},0)),\"\")"
    lst.sheet_state = "hidden"


BOLUM_COLS = {}


def build_bolum(wb):
    ws = wb.create_sheet("_Bolum")
    cols = [
        ("k", lambda r: f"=ROW()-1"),
        ("rB", lambda r: "=" + find(f'"Bolum|"&A{r}')),
        ("Bölüm", lambda r: f'=IF(B{r}="","",INDEX({vcol("Bölüm")},B{r}))'),
        ("Müdürlük", lambda r: f'=IF(B{r}="","",INDEX({vcol("Müdürlük")},B{r}))'),
        ("Direktörlük", lambda r: f'=IF(B{r}="","",INDEX({vcol("Direktörlük")},B{r}))'),
        ("Endeks", lambda r: f'=IF(B{r}="","",INDEX({vcol("Ad")},B{r})&"")'),
        ("KGrup", lambda r: f'=IF(B{r}="","",INDEX({vcol("Kod")},B{r}))'),
        ("Var", lambda r: f'=IF(B{r}="","",IF(INDEX({vcol("Kaynak")},B{r})&""="",0,1))'),
        ("Kapsamda", lambda r: f'=IF(C{r}="","",IF(OR({KAPSAM}="Tümü",D{r}={KAPSAM}),1,0))'),
        ("Sıra", lambda r: f'=IF(I{r}=1,COUNTIF($I$2:I{r},1),"")'),
        ("rTop", lambda r: f'=IF(C{r}="","",{find(chr(34) + "Toplam|" + chr(34) + f"&C{r}&" + chr(34) + "|" + chr(34) + f"&{YIL}")})'),
        ("A", lambda r: "=" + sel(f"K{r}")),
        ("B", lambda r: f'=IF({BUTCEYILI}=1,{sel(f"K{r}", VB)},"")'),
        ("rTopP", lambda r: f'=IF(C{r}="","",{find(chr(34) + "Toplam|" + chr(34) + f"&C{r}&" + chr(34) + "|" + chr(34) + f"&({YIL}-1)")})'),
        ("P", lambda r: "=" + sel(f"N{r}")),
        ("rAdet", lambda r: f'=IF(C{r}="","",{find(chr(34) + "Adet|" + chr(34) + f"&C{r}&" + chr(34) + "|" + chr(34) + f"&{YIL}")})'),
        ("X", lambda r: "=" + pos(f"P{r}")),
        ("rAdetP", lambda r: f'=IF(C{r}="","",{find(chr(34) + "Adet|" + chr(34) + f"&C{r}&" + chr(34) + "|" + chr(34) + f"&({YIL}-1)")})'),
        ("Y", lambda r: "=" + pos(f"R{r}")),
        ("Qb", lambda r: f'=IF({BUTCEYILI}=1,{pos(f"P{r}", VB)},"")'),
        ("rEnd", lambda r: f'=IF(F{r}="","",{find(chr(34) + "Endeks|" + chr(34) + f"&F{r}&" + chr(34) + "|" + chr(34) + f"&{YIL}")})'),
        ("rEndP", lambda r: f'=IF(F{r}="","",{find(chr(34) + "Endeks|" + chr(34) + f"&F{r}&" + chr(34) + "|" + chr(34) + f"&({YIL}-1)")})'),
        ("Endeks oranı", lambda r: f'=IF(OR(U{r}="",V{r}=""),"",IF(COUNT({L(RATIO0 + 1)}{r}:{L(RATIO0 + 12)}{r})<{SELCOUNT},"",SUM({L(RATIO0 + 1)}{r}:{L(RATIO0 + 12)}{r})/{SELCOUNT}))'),
        ("Endeks var", lambda r: f'=IF(F{r}="",0,IF(COUNTIF({VKEY},"Endeks|"&F{r}&"|*")>0,1,0))'),
        ("rTopK", lambda r: f'=IF(C{r}="","",{find(chr(34) + "Toplam|" + chr(34) + f"&C{r}&" + chr(34) + "|" + chr(34) + f"&{KAPYIL}")})'),
        ("YTD", lambda r: "=" + sel(f"Y{r}", VG, YTDF)),
        ("YTD bütçe", lambda r: "=" + sel(f"Y{r}", VB, YTDF)),
        ("Kalan bütçe", lambda r: f'=IF(Y{r}="","",IF({KAPAY}=12,0,{sel(f"Y{r}", VB, RESTF)}))'),
        ("BM uygun", lambda r: (
            f'=IF(C{r}="",0,IF(AND(H{r}=1,X{r}=1),IF({BUTCEYILI}=1,IF({num(f"L{r}", f"Q{r}", f"O{r}", f"S{r}", f"W{r}")},'
            f'IF(AND(L{r}>=0,Q{r}>0,O{r}>0,S{r}>0),1,0),0),IF({num(f"L{r}", f"Q{r}")},IF(Q{r}>0,1,0),0)),0))')),
        ("Endekslenmiş", lambda r: f'=IF(AND(AC{r}=1,{num(f"O{r}", f"W{r}")}),O{r}*W{r},"")'),
        ("BM", lambda r: f'=IF(AC{r}=1,L{r}/Q{r},"")'),
        ("Baz BM", lambda r: f'=IF(AND(AC{r}=1,ISNUMBER(AD{r})),AD{r}/S{r},"")'),
        ("İyileşme", lambda r: f'=IF({num(f"AE{r}", f"AF{r}")},IF(AF{r}=0,"",(AF{r}-AE{r})/AF{r}),"")'),
        ("Bütçe BM", lambda r: f'=IF(AND(AC{r}=1,{BUTCEYILI}=1,{num(f"M{r}", f"T{r}")}),IF(T{r}>0,M{r}/T{r},""),"")'),
        ("Müd. maliyet", lambda r: f'=IF(AC{r}=1,SUMIFS($L$2:$L${MAXD + 1},$D$2:$D${MAXD + 1},D{r},$AC$2:$AC${MAXD + 1},1),"")'),
        ("Dir. maliyet", lambda r: f'=IF(AC{r}=1,SUMIFS($L$2:$L${MAXD + 1},$E$2:$E${MAXD + 1},E{r},$AC$2:$AC${MAXD + 1},1),"")'),
        ("Müd. ağırlık", lambda r: f'=IF(AC{r}=1,IF(AI{r}=0,"",L{r}/AI{r}),"")'),
        ("Müd. ağ.×iyil.", lambda r: f'=IF({num(f"AK{r}", f"AG{r}")},AK{r}*AG{r},"")'),
        ("Dir. ağırlık", lambda r: f'=IF(AC{r}=1,IF(AJ{r}=0,"",L{r}/AJ{r}),"")'),
        ("Dir. ağ.×iyil.", lambda r: f'=IF({num(f"AM{r}", f"AG{r}")},AM{r}*AG{r},"")'),
        ("A×iyil.", lambda r: f'=IF({num(f"L{r}", f"AG{r}")},L{r}*AG{r},"")'),
        ("Bütçe BM (hacim)", lambda r: f'=IF(AND({BUTCEYILI}=1,{num(f"M{r}", f"T{r}")}),IF(T{r}=0,"",M{r}/T{r}),"")'),
        ("Hacim etkisi", lambda r: f'=IF({num(f"L{r}", f"AP{r}", f"Q{r}")},(Q{r}-T{r})*AP{r},"")'),
        ("BM etkisi", lambda r: f'=IF({num(f"L{r}", f"AP{r}", f"Q{r}")},L{r}-AP{r}*Q{r},"")'),
    ]
    # AS..AW: eksik bayrakları (1 = sayı değil) toplamların katılığı için
    miss = [("eksik A", "L"), ("eksik P", "O"), ("eksik B", "M"), ("eksik YTD", "Z"), ("eksik YTDb", "AA")]
    miss += [("eksik Kalan", "AB"), ("eksik Hacim", "AQ"), ("eksik BMetkisi", "AR")]
    for i, (title, _) in enumerate(cols, 1):
        ws.cell(1, i, title)
    first_miss = len(cols) + 1  # AT
    for j, (title, src) in enumerate(miss):
        ws.cell(1, first_miss + j, title)
    assert first_miss + len(miss) - 1 < RATIO0 + 1
    for m in range(1, 13):
        ws.cell(1, RATIO0 + m, f"Endeks oranı {m}")
        ws.cell(1, DIP0 + m, f"Dip {m}")
        ws.cell(1, DIPMISS0 + m, f"Dip eksik {m}")
    for r in range(2, MAXD + 2):
        for i, (_, formula) in enumerate(cols, 1):
            ws.cell(r, i, formula(r))
        for j, (_, src) in enumerate(miss):
            ws.cell(r, first_miss + j, f'=IF(C{r}="",0,IF(ISNUMBER({src}{r}),0,1))')
        for m in range(1, 13):
            cur, prev = f"INDEX({VG},$U{r},{m})", f"INDEX({VG},$V{r},{m})"
            ws.cell(r, RATIO0 + m, f'=IF(OR($U{r}="",$V{r}=""),"",IF(INDEX({FLAGS},1,{m})=1,IF(AND({cur}>0,{prev}>0),{cur}/{prev},""),""))')
            ws.cell(r, DIP0 + m, f'=IF($K{r}="","",IF(ISNUMBER(INDEX({VG},$K{r},{m})),INDEX({VG},$K{r},{m}),""))')
            ws.cell(r, DIPMISS0 + m, f'=IF($C{r}="",0,IF(ISNUMBER({L(DIP0 + m)}{r}),0,1))')
    BOLUM_COLS.update({title: L(first_miss + j) for j, (title, _) in enumerate(miss)})
    ws.sheet_state = "hidden"


B = "'_Bolum'"


def brng(col):
    return f"{B}!${col}$2:${col}${MAXD + 1}"


def build_duzen(wb):
    ws = wb.create_sheet("_Duzen")
    titles = ["k", "rD", "Başlık", "Tür", "Müdürlük", "Direktörlük", "Bölüm", "Kapsamda", "Sıra",
              "ölçüt bölüm", "ölçüt müdürlük", "ölçüt direktörlük", "Bölüm sayısı"]
    crit = lambda r: f"{brng('C')},$J{r},{brng('D')},$K{r},{brng('E')},$L{r}"  # noqa: E731

    def strict(r, col, miss_title):
        return (f'=IF(OR($C{r}="",$M{r}=0),"",IF(SUMIFS({brng(BOLUM_COLS[miss_title])},{crit(r)})>0,"",'
                f"SUMIFS({brng(col)},{crit(r)})))")

    def strict_eligible(r, col):
        # BM uygun bölümler üzerinde toplam (uygunluk bu değerlerin sayı olmasını gerektirir); uygun bölüm yoksa boş
        return (f'=IF(OR($C{r}="",$N{r}=0),"",SUMIFS({brng(col)},{crit(r)},{brng("AC")},1))')

    values = [
        ("Uygun bölüm", lambda r: f'=IF($C{r}="",0,COUNTIFS({crit(r)},{brng("AC")},1))'),  # N
        ("A", lambda r: strict(r, "L", "eksik A")),  # O
        ("P", lambda r: strict(r, "O", "eksik P")),  # P
        ("B", lambda r: f'=IF({BUTCEYILI}=1,{strict(r, "M", "eksik B")[1:]},"")'),  # Q
        ("YTD", lambda r: strict(r, "Z", "eksik YTD")),  # R
        ("YTD bütçe", lambda r: strict(r, "AA", "eksik YTDb")),  # S
        ("Kalan", lambda r: strict(r, "AB", "eksik Kalan")),  # T
        ("Yıllık bütçe", lambda r: f'=IF({num(f"S{r}", f"T{r}")},S{r}+T{r},"")'),  # U
        ("Tahmin", lambda r: f'=IF({num(f"R{r}", f"T{r}")},R{r}+T{r},"")'),  # V
        ("Ortalama gidiş", lambda r: f'=IF(ISNUMBER(R{r}),R{r}*12/{KAPAY},"")'),  # W
        ("Hacim etkisi", lambda r: strict(r, "AQ", "eksik Hacim")),  # X
        ("BM etkisi", lambda r: strict(r, "AR", "eksik BMetkisi")),  # Y
        ("aE", lambda r: strict_eligible(r, "L")),  # Z
        ("xE", lambda r: strict_eligible(r, "Q")),  # AA
        ("bE", lambda r: f'=IF({BUTCEYILI}=1,{strict_eligible(r, "O")[1:]},"")'),  # AB
        ("yE", lambda r: f'=IF({BUTCEYILI}=1,{strict_eligible(r, "S")[1:]},"")'),  # AC
        ("adjE", lambda r: (f'=IF(OR($C{r}="",$N{r}=0,{BUTCEYILI}=0),"",IF(COUNTIFS({crit(r)},{brng("AC")},1,{brng("AD")},">-1E+307")<$N{r},"",'
                            f'SUMIFS({brng("AD")},{crit(r)},{brng("AC")},1)))')),  # AD
        ("mw", lambda r: strict_eligible(r, "AK")),  # AE
        ("mwr", lambda r: (f'=IF(OR($C{r}="",$N{r}=0),"",IF(COUNTIFS({crit(r)},{brng("AC")},1,{brng("AL")},">-1E+307")<$N{r},"",'
                           f'SUMIFS({brng("AL")},{crit(r)},{brng("AC")},1)))')),  # AF
        ("dw", lambda r: strict_eligible(r, "AM")),  # AG
        ("dwr", lambda r: (f'=IF(OR($C{r}="",$N{r}=0),"",IF(COUNTIFS({crit(r)},{brng("AC")},1,{brng("AN")},">-1E+307")<$N{r},"",'
                           f'SUMIFS({brng("AN")},{crit(r)},{brng("AC")},1)))')),  # AH
        ("a×r", lambda r: (f'=IF(OR($C{r}="",$N{r}=0),"",IF(COUNTIFS({crit(r)},{brng("AC")},1,{brng("AO")},">-1E+307")<$N{r},"",'
                           f'SUMIFS({brng("AO")},{crit(r)},{brng("AC")},1)))')),  # AI
        ("BM", lambda r: f'=IF(AND({num(f"Z{r}", f"AA{r}")}),IF(AA{r}=0,"",Z{r}/AA{r}),"")'),  # AJ
        ("BM sıra", lambda r: f'=IF(AND(H{r}=1,ISNUMBER(AJ{r})),COUNTIFS($H$2:H{r},1,$AJ$2:AJ{r},">-1E+307"),"")'),  # AK
        ("Opex", lambda r: f'=O{r}'),  # AL
    ]
    for i, t in enumerate(titles + [t for t, _ in values], 1):
        ws.cell(1, i, t)
    for r in range(2, MAXL + 2):
        ws[f"A{r}"] = "=ROW()-1"
        ws[f"B{r}"] = "=" + find(f'"Duzen|"&A{r}')
        ws[f"C{r}"] = f'=IF(B{r}="","",INDEX({vcol("Ad")},B{r}))'
        ws[f"D{r}"] = f'=IF(B{r}="","",INDEX({vcol("Grup")},B{r}))'
        ws[f"E{r}"] = f'=IF(B{r}="","",INDEX({vcol("Müdürlük")},B{r})&"")'
        ws[f"F{r}"] = f'=IF(B{r}="","",INDEX({vcol("Direktörlük")},B{r})&"")'
        ws[f"G{r}"] = f'=IF(B{r}="","",INDEX({vcol("Bölüm")},B{r})&"")'
        ws[f"H{r}"] = (f'=IF(C{r}="","",IF({KAPSAM}="Tümü",1,IF(AND(OR(D{r}="Bölüm",D{r}="Müdürlük toplamı"),'
                       f'E{r}={KAPSAM}),1,0)))')
        ws[f"I{r}"] = f'=IF(H{r}=1,COUNTIF($H$2:H{r},1),"")'
        ws[f"J{r}"] = f'=IF(D{r}="Bölüm",G{r},"*")'
        ws[f"K{r}"] = f'=IF(D{r}="Müdürlük toplamı",E{r},"*")'
        ws[f"L{r}"] = f'=IF(D{r}="Direktörlük toplamı",F{r},"*")'
        ws[f"M{r}"] = f'=IF(C{r}="",0,COUNTIFS({crit(r)}))'
        for i, (_, formula) in enumerate(values, 14):
            ws.cell(r, i, formula(r))
    ws.sheet_state = "hidden"


def build_kalem(wb):
    ws = wb.create_sheet("_Kalem")
    titles = ["k", "rK", "Bölüm", "Müdürlük", "Kod", "Kalem", "Grup", "Anahtar", "Kapsamda", "rG", "A", "B", "rGP", "P",
              "Fark", "Oran", "Önemli", "|fark|", "Sıra", "|aleyhte|", "Aleyhte sıra", "|lehte|", "Lehte sıra", "Yön",
              "eksik A", "eksik B", "Açıklama", "Onay satırı", "Onay durumu", "Onay kontrolü", "Elle açıklama"]
    for i, t in enumerate(titles, 1):
        ws.cell(1, i, t)
    last = MAXK + 1
    for r in range(2, last + 1):
        ws[f"A{r}"] = "=ROW()-1"
        ws[f"B{r}"] = "=" + find(f'"Kalem|"&A{r}')
        ws[f"C{r}"] = f'=IF(B{r}="","",INDEX({vcol("Bölüm")},B{r}))'
        ws[f"D{r}"] = f'=IF(B{r}="","",INDEX({vcol("Müdürlük")},B{r}))'
        ws[f"E{r}"] = f'=IF(B{r}="","",INDEX({vcol("Kod")},B{r}))'
        ws[f"F{r}"] = f'=IF(B{r}="","",INDEX({vcol("Ad")},B{r}))'
        ws[f"G{r}"] = f'=IF(B{r}="","",INDEX({vcol("Grup")},B{r}))'
        ws[f"H{r}"] = f'=IF(B{r}="","",C{r}&"|"&E{r})'
        ws[f"I{r}"] = f'=IF(C{r}="","",IF(OR({KAPSAM}="Tümü",D{r}={KAPSAM}),1,0))'
        ws[f"J{r}"] = f'=IF(C{r}="","",{find(chr(34) + "Gider|" + chr(34) + f"&C{r}&" + chr(34) + "|" + chr(34) + f"&E{r}&" + chr(34) + "|" + chr(34) + f"&{YIL}")})'
        ws[f"K{r}"] = "=" + sel(f"J{r}")
        ws[f"L{r}"] = f'=IF({BUTCEYILI}=1,{sel(f"J{r}", VB)},"")'
        ws[f"M{r}"] = f'=IF(C{r}="","",{find(chr(34) + "Gider|" + chr(34) + f"&C{r}&" + chr(34) + "|" + chr(34) + f"&E{r}&" + chr(34) + "|" + chr(34) + f"&({YIL}-1)")})'
        ws[f"N{r}"] = "=" + sel(f"M{r}")
        ws[f"O{r}"] = f'=IF({num(f"K{r}", f"L{r}")},K{r}-L{r},"")'
        ws[f"P{r}"] = f'=IF({num(f"O{r}", f"L{r}")},IF(L{r}=0,"",O{r}/L{r}),"")'
        ws[f"Q{r}"] = (f'=IF(AND(I{r}=1,ISNUMBER(O{r})),IF(AND(ABS(N(O{r}))>={ESIK},OR(P{r}="",ABS(N(P{r}))>={ORAN})),1,0),0)')
        ws[f"R{r}"] = f'=IF(Q{r}=1,ABS(O{r}),"")'
        ws[f"S{r}"] = f'=IF(Q{r}=1,RANK(R{r},$R$2:$R${last})+COUNTIF($R$2:R{r},R{r})-1,"")'
        ws[f"T{r}"] = f'=IF(AND(Q{r}=1,N(O{r})>0),ABS(O{r}),"")'
        ws[f"U{r}"] = f'=IF(ISNUMBER(T{r}),RANK(T{r},$T$2:$T${last})+COUNTIF($T$2:T{r},T{r})-1,"")'
        ws[f"V{r}"] = f'=IF(AND(Q{r}=1,N(O{r})<0),ABS(O{r}),"")'
        ws[f"W{r}"] = f'=IF(ISNUMBER(V{r}),RANK(V{r},$V$2:$V${last})+COUNTIF($V$2:V{r},V{r})-1,"")'
        ws[f"X{r}"] = f'=IF(ISNUMBER(O{r}),IF(O{r}>0,"Aleyhte",IF(O{r}<0,"Lehte","Bütçede")),"")'
        ws[f"Y{r}"] = f'=IF(C{r}="",0,IF(ISNUMBER(K{r}),0,1))'
        ws[f"Z{r}"] = f'=IF(C{r}="",0,IF(ISNUMBER(L{r}),0,1))'
        # Açıklama: önce bu dosyanın Açıklamalar sayfası; yoksa müdürlükten dönen, onaylı ve tutarı güncel açıklama
        ws[f"AE{r}"] = f"=IF(H{r}=\"\",\"\",IFERROR(INDEX('Açıklamalar'!$D$2:$D${MAXA + 1},MATCH(H{r},'Açıklamalar'!$H$2:$H${MAXA + 1},0))&\"\",\"\"))"
        ws[f"AB{r}"] = f'=IF(H{r}="","",IFERROR(MATCH(H{r}&"|"&{H}!$B$21&"|"&{H}!$B$20,{ONAY}!$A:$A,0),""))'
        ws[f"AC{r}"] = f'=IF(AB{r}="","",INDEX({ONAY}!${OC["Durum"]}:${OC["Durum"]},AB{r})&"")'
        ws[f"AD{r}"] = f'=IF(AB{r}="","",INDEX({ONAY}!$AA:$AA,AB{r})&"")'
        ws[f"AA{r}"] = (f'=IF(AE{r}<>"",AE{r},IF(OR(AB{r}="",LEFT(AD{r},6)<>"Güncel"),"",'
                        f'IF(AC{r}="Onaylandı",INDEX({ONAY}!$F:$F,AB{r})&"",IF(AC{r}="Bekliyor","[Onay bekliyor] "&INDEX({ONAY}!$F:$F,AB{r}),""))))')
    ws.sheet_state = "hidden"


K_ = "'_Kalem'"


def krng(col):
    return f"{K_}!${col}$2:${col}${MAXK + 1}"


def build_satir(wb):
    ws = wb.create_sheet("_Satir")
    for i, t in enumerate(["i", "rS", "Tür", "Grup", "Kod", "Ad"], 1):
        ws.cell(1, i, t)
    for r in range(2, MAXI + 2):
        ws[f"A{r}"] = "=ROW()-1"
        ws[f"B{r}"] = "=" + find(f'"Satir|"&A{r}')
        ws[f"C{r}"] = f'=IF(B{r}="","",INDEX({vcol("Metin")},B{r}))'
        ws[f"D{r}"] = f'=IF(B{r}="","",INDEX({vcol("Grup")},B{r}))'
        ws[f"E{r}"] = f'=IF(B{r}="","",IF(C{r}="kalem",INDEX({vcol("Kod")},B{r}),""))'
        ws[f"F{r}"] = f'=IF(B{r}="","",INDEX({vcol("Ad")},B{r}))'
    # Blok başlıkları: j. kapsamdaki bölüm
    ws["H1"] = "Blok"
    ws["I1"] = "Bölüm"
    for j in range(1, BLOCKS + 1):
        ws[f"H{j + 1}"] = j
        ws[f"I{j + 1}"] = f"=IFERROR(INDEX({brng('C')},MATCH({j},{brng('J')},0)),\"\")"
    ws.sheet_state = "hidden"


def build_kgrup(wb):
    ws = wb.create_sheet("_KGrup")
    titles = ["k", "rK", "Kontrol bölümü", "Anahtar", "Dir. anahtarı", "Direktörlük", "Müdürlük", "Bölümler",
              "Üye", "Kapsamdaki üye", "Kapsamda", "Sıra", "Eşleşen satır", "rC"]
    for i, t in enumerate(titles, 1):
        ws.cell(1, i, t)
    for m in range(1, 13):
        ws.cell(1, 14 + m, f"Dip {m}")  # O..Z
        ws.cell(1, 26 + m, f"Kontrol {m}")  # AA..AL
    for r in range(2, MAXG + 2):
        ws[f"A{r}"] = "=ROW()-1"
        ws[f"B{r}"] = "=" + find(f'"KGrup|"&A{r}')
        ws[f"C{r}"] = f'=IF(B{r}="","",INDEX({vcol("Ad")},B{r}))'
        ws[f"D{r}"] = f'=IF(B{r}="","",INDEX({vcol("Metin")},B{r}))'
        ws[f"E{r}"] = f'=IF(B{r}="","",INDEX({vcol("Grup")},B{r}))'
        ws[f"F{r}"] = f'=IF(B{r}="","",INDEX({vcol("Direktörlük")},B{r}))'
        ws[f"G{r}"] = f'=IF(B{r}="","",INDEX({vcol("Müdürlük")},B{r}))'
        ws[f"H{r}"] = f'=IF(B{r}="","",INDEX({vcol("Bölüm")},B{r}))'
        ws[f"I{r}"] = f'=IF(B{r}="",0,COUNTIF({brng("G")},A{r}))'
        ws[f"J{r}"] = f'=IF(B{r}="",0,COUNTIFS({brng("G")},A{r},{brng("I")},1))'
        ws[f"K{r}"] = f'=IF(AND(I{r}>0,J{r}=I{r}),1,0)'
        ws[f"L{r}"] = f'=IF(K{r}=1,COUNTIF($K$2:K{r},1),"")'
        ws[f"M{r}"] = f'=IF(B{r}="",0,COUNTIF({VKEY},"Kontrol|"&E{r}&"|"&D{r}&"|*|"&{YIL}))'
        ws[f"N{r}"] = f'=IF(B{r}="","",{find(chr(34) + "Kontrol|" + chr(34) + f"&E{r}&" + chr(34) + "|" + chr(34) + f"&D{r}&" + chr(34) + "|1|" + chr(34) + f"&{YIL}")})'
        for m in range(1, 13):
            dip, miss = L(DIP0 + m), L(DIPMISS0 + m)
            ws.cell(r, 14 + m, f'=IF(I{r}=0,"",IF(SUMIFS({brng(miss)},{brng("G")},A{r})>0,"",SUMIFS({brng(dip)},{brng("G")},A{r})))')
            ws.cell(r, 26 + m, f'=IF(OR(M{r}<>1,N{r}=""),"",IF(ISNUMBER(INDEX({VG},N{r},{m})),INDEX({VG},N{r},{m}),""))')
    ws.sheet_state = "hidden"


def build_lists(wb):
    """Revizyon ve uyarı listeleri (kapsama göre süzülür)."""
    rev = wb.create_sheet("_Rev")
    for i, t in enumerate(["k", "rR", "Bölüm", "Müdürlük", "Kapsamda", "Sıra"], 1):
        rev.cell(1, i, t)
    for r in range(2, MAXR + 2):
        rev[f"A{r}"] = "=ROW()-1"
        rev[f"B{r}"] = "=" + find(f'"Revizyon|"&A{r}')
        rev[f"C{r}"] = f'=IF(B{r}="","",INDEX({vcol("Bölüm")},B{r}))'
        rev[f"D{r}"] = f'=IF(B{r}="","",INDEX({vcol("Müdürlük")},B{r}))'
        rev[f"E{r}"] = f'=IF(C{r}="","",IF(OR({KAPSAM}="Tümü",D{r}={KAPSAM}),1,0))'
        rev[f"F{r}"] = f'=IF(E{r}=1,COUNTIF($E$2:E{r},1),"")'
    rev.sheet_state = "hidden"

    uy = wb.create_sheet("_Uy")
    for i, t in enumerate(["k", "rU", "Dönem", "Bölüm", "Kaynak", "Uyarı", "Kapsamda", "Sıra"], 1):
        uy.cell(1, i, t)
    r = 2
    for k in range(1, MAXW + 1):  # sorgunun uyarıları
        uy[f"A{r}"] = k
        uy[f"B{r}"] = "=" + find(f'"Uyari|"&A{r}')
        uy[f"C{r}"] = f'=IF(B{r}="","",INDEX({vcol("Ad")},B{r})&"")'
        uy[f"D{r}"] = f'=IF(B{r}="","",INDEX({vcol("Bölüm")},B{r})&"")'
        uy[f"E{r}"] = f'=IF(B{r}="","",INDEX({vcol("Kaynak")},B{r})&"")'
        uy[f"F{r}"] = f'=IF(B{r}="","",INDEX({vcol("Metin")},B{r})&"")'
        uy[f"G{r}"] = (f'=IF(B{r}="",0,IF(OR(D{r}="",{KAPSAM}="Tümü"),1,IF(ISNA(MATCH(D{r},{brng("C")},0)),1,'
                       f'IF(INDEX({brng("D")},MATCH(D{r},{brng("C")},0))={KAPSAM},1,0))))')
        r += 1
    period = f'{KAPYIL}&"-"&RIGHT("0"&{KAPAY},2)'
    for d in range(1, MAXD + 1):  # kapanış dosyası olmayan bölümler ve eksik endeks
        br = d + 1
        uy[f"A{r}"] = MAXW + d
        uy[f"C{r}"] = f"={period}"
        uy[f"D{r}"] = f"={B}!$C${br}"
        uy[f"F{r}"] = (f'=IF({B}!$C${br}="","",IF({B}!$H${br}=0,"Kapanış dosyası yok veya kaynak okunamadı",'
                       f'IF(AND({B}!$F${br}<>"",{B}!$W${br}=""),"Seçilen dönem için endeks hesabı tamamlanamadı","")))')
        uy[f"G{r}"] = f'=IF(F{r}="",0,IF({B}!$I${br}=1,1,0))'
        r += 1
    extras = [
        (f'=IF(AND({KAPYIL}<>"",{BUTCEYILI}=0),{YIL}&" bütçesi bu kapanış dosyasında yok. Bütçe karşılaştırması boş bırakıldı.","")'),
        (f'=IF({KAPYIL}="","",IFERROR(IF(SUMPRODUCT({FLAGS}*ISNUMBER(INDEX({VG},MATCH("Ciro|"&{YIL},{VKEY},0),0)))<{SELCOUNT},"Seçilen dönem için şirket cirosu eksik",""),"Seçilen dönem için şirket cirosu eksik"))'),
    ]
    G_ = "'_KGrup'"
    for g in range(2, MAXG + 2):  # kapsam birleşik kontrol grubunu bölüyorsa (Python müdürlük raporundaki uyarı)
        uy[f"A{r}"] = MAXW + MAXD + 1
        uy[f"C{r}"] = f"={period}"
        uy[f"F{r}"] = f'=IF(AND({G_}!$J${g}>0,{G_}!$J${g}<{G_}!$I${g}),{G_}!$C${g}&": Birleşik bölüm kontrolü direktörlük raporunda yer alır","")'
        uy[f"G{r}"] = f'=IF(F{r}="",0,1)'
        r += 1
    for formula in extras:
        uy[f"A{r}"] = MAXW + MAXD + 1
        uy[f"C{r}"] = f"={period}"
        uy[f"F{r}"] = formula
        uy[f"G{r}"] = f'=IF(F{r}="",0,1)'
        r += 1
    for rr in range(2, r):
        uy[f"H{rr}"] = f'=IF(G{rr}=1,COUNTIF($G$2:G{rr},1),"")'
    uy.sheet_state = "hidden"
    return r - 1  # son satır


# ----------------------------------------------------------------------------- Görünür rapor sayfaları
D_ = "'_Duzen'"


def drng(col):
    return f"{D_}!${col}$2:${col}${MAXL + 1}"


def dval(col, k):
    return f'IF({k}="","",INDEX({drng(col)},{k}))'


def build_genel(wb):
    ws = wb.create_sheet("Genel Özet")
    header(ws, 1, ["Bölüm / müdürlük / direktörlük", "Satır türü", "Önceki yıl gerçekleşen (TL)", "Gerçekleşen (TL)",
                   "Bütçe (TL)", "Bütçe − gerçekleşen (TL)", "Bütçe aşım oranı (%)", "OPEX / şirket cirosu", "k"])
    ciro = f'IFERROR({sel(find(chr(34) + "Ciro|" + chr(34) + "&" + YIL))},"")'
    ws["K1"] = "Seçili ayların şirket cirosu"
    ws["K2"] = "=" + ciro
    for j in range(1, MAXL + 1):
        r = j + 1
        k = f"$I{r}"
        ws[f"I{r}"] = f"=IFERROR(MATCH({j},{drng('I')},0),\"\")"
        ws[f"A{r}"] = "=" + dval("C", k)
        ws[f"B{r}"] = "=" + dval("D", k)
        ws[f"C{r}"] = "=" + dval("P", k)
        ws[f"D{r}"] = "=" + dval("O", k)
        ws[f"E{r}"] = "=" + dval("Q", k)
        ws[f"F{r}"] = f'=IF({num(f"D{r}", f"E{r}")},E{r}-D{r},"")'
        ws[f"G{r}"] = f'=IF({num(f"D{r}", f"E{r}")},IF(E{r}=0,"",(D{r}-E{r})/E{r}),"")'
        ws[f"H{r}"] = f'=IF(AND(ISNUMBER(D{r}),ISNUMBER($K$2)),IF($K$2=0,"",D{r}/$K$2),"")'
    fmt(ws, "CDEF", 2, MAXL + 1, MONEY)
    fmt(ws, "GH", 2, MAXL + 1, PCT)
    ws["K2"].number_format = MONEY
    total_rows_style(ws, f"A2:H{MAXL + 1}", "B2")
    ws.conditional_formatting.add(f"F2:F{MAXL + 1}", FormulaRule(formula=["AND(ISNUMBER(F2),F2>0)"], font=Font(color="1B754C")))
    ws.conditional_formatting.add(f"F2:F{MAXL + 1}", FormulaRule(formula=["AND(ISNUMBER(F2),F2<0)"], font=Font(color="BE303A")))
    ws.conditional_formatting.add(f"G2:G{MAXL + 1}", FormulaRule(formula=["AND(ISNUMBER(G2),G2>0)"], font=Font(color="BE303A")))
    ws.conditional_formatting.add(f"G2:G{MAXL + 1}", FormulaRule(formula=["AND(ISNUMBER(G2),G2<0)"], font=Font(color="1B754C")))
    ws.column_dimensions["I"].hidden = True
    widths(ws, {"A": 34, "B": 20, "C": 20, "D": 20, "E": 20, "F": 20, "G": 14, "H": 14, "K": 22})
    ws.freeze_panes = "A2"


def build_giderler(wb):
    ws = wb.create_sheet("Giderler")
    S = "'_Satir'"
    ws["A1"] = "Giderler"
    ws["A1"].font = TITLE
    ws["A2"] = "Kırmızı: bütçe aşıldı · Yeşil: bütçe altında. Açıklamalar 'Açıklamalar' sayfasından gelir."
    ws["A2"].font = NOTE
    header(ws, 5, ["Gider grubu", "İndeks", "Gider kalemi"])
    titles = ["Gerçekleşen (TL)", "Bütçe (TL)", "Bütçe − gerçekleşen (TL)", "Bütçe aşım oranı (%)", "Açıklama"]
    edge = Side(style="thin", color="233C57")
    for j in range(BLOCKS):
        start = 4 + 5 * j
        ws.merge_cells(start_row=4, start_column=start, end_row=4, end_column=start + 4)
        cell = ws.cell(4, start, f"={S}!$I${j + 2}&\"\"")
        cell.fill = HEADER_FILL
        cell.font = BOLD
        cell.alignment = Alignment(horizontal="center")
        header(ws, 5, titles, start)
        ws.column_dimensions[L(start + 4)].width = 30
    for i in range(1, MAXI + 1):
        r = i + 5
        sr = i + 1
        tip, grp, kod = f"{S}!$C${sr}", f"{S}!$D${sr}", f"{S}!$E${sr}"
        ws[f"A{r}"] = f'=IF({tip}="","",IF({tip}="kalem","",{grp}))'
        ws[f"B{r}"] = f'=IF({tip}="kalem",{kod},"")'
        ws[f"C{r}"] = f"=IF({S}!$F${sr}=\"\",\"\",{S}!$F${sr})"
        for j in range(BLOCKS):
            start = 4 + 5 * j
            dept = f"{L(start)}$4"
            g, b, d, o, n = (L(start + x) for x in range(5))

            def agg(value_col, miss_col):
                by_item = f'IFERROR(INDEX({krng(value_col)},MATCH({dept}&"|"&{kod},{krng("H")},0)),"")'
                grp_filter = f'{krng("C")},{dept},{krng("G")},{grp}'
                all_filter = f'{krng("C")},{dept}'
                group_total = (f'IF(COUNTIFS({grp_filter})=0,"",IF(SUMIFS({krng(miss_col)},{grp_filter})>0,"",'
                               f"SUMIFS({krng(value_col)},{grp_filter})))")
                all_total = (f'IF(COUNTIFS({all_filter})=0,"",IF(SUMIFS({krng(miss_col)},{all_filter})>0,"",'
                             f"SUMIFS({krng(value_col)},{all_filter})))")
                return (f'=IF(OR({dept}="",{tip}=""),"",IF({tip}="kalem",{by_item},IF({tip}="grup",{group_total},'
                        f"{all_total})))")

            ws[f"{g}{r}"] = agg("K", "Y")
            ws[f"{b}{r}"] = agg("L", "Z")
            ws[f"{d}{r}"] = f'=IF({num(f"{g}{r}", f"{b}{r}")},{b}{r}-{g}{r},"")'
            ws[f"{o}{r}"] = f'=IF({num(f"{g}{r}", f"{b}{r}")},IF({b}{r}=0,"",({g}{r}-{b}{r})/{b}{r}),"")'
            ws[f"{n}{r}"] = (f'=IF(OR({dept}="",{tip}<>"kalem"),"",IFERROR(INDEX({krng("AA")},'
                             f'MATCH({dept}&"|"&{kod},{krng("H")},0))&"",""))')
    last = MAXI + 5
    for j in range(BLOCKS):
        start = 4 + 5 * j
        g, b, d, o, n = (L(start + x) for x in range(5))
        fmt(ws, [g, b, d], 6, last, MONEY)
        fmt(ws, [o], 6, last, PCT)
        ws.conditional_formatting.add(f"{d}6:{o}{last}", FormulaRule(formula=[f"AND(ISNUMBER(${d}6),${d}6<0)"], fill=RED_FILL))
        ws.conditional_formatting.add(f"{d}6:{o}{last}", FormulaRule(formula=[f"AND(ISNUMBER(${d}6),${d}6>0)"], fill=GREEN_FILL))
        for x in range(4):
            ws.column_dimensions[L(start + x)].width = 16
        # Önemli sapma açıklama hücreleri turuncu
        ws.conditional_formatting.add(
            f"{n}6:{n}{last}",
            FormulaRule(formula=[f'IFERROR(INDEX({krng("Q")},MATCH({g}$4&"|"&$B6,{krng("H")},0))=1,FALSE)'],
                        fill=PatternFill("solid", bgColor="F8CBAD")),
        )
    ws.conditional_formatting.add(f"A6:C{last}", FormulaRule(formula=['AND($A6<>"",$B6="")'], font=Font(bold=True)))
    widths(ws, {"A": 22, "B": 8, "C": 30})
    ws.freeze_panes = "D6"
    del edge


def bnum(col, brow):
    """Bölüm satırının _Bolum'daki sayısal değeri (yoksa boş)."""
    v = f"INDEX({brng(col)},{brow})"
    return f'IFERROR(IF(ISNUMBER({v}),{v},""),"")'


def build_birim(wb):
    ws = wb.create_sheet("Birim Maliyet")
    header(ws, 1, ["Bölüm / müdürlük / direktörlük", "Müdürlük / satır türü", "Toplam gider (TL)", "İşlem adedi",
                   "Birim maliyet (TL/işlem)", "Önceki yıl toplam gider (TL)", "Önceki yıl işlem adedi",
                   "Önceki yıl birim maliyet (TL/işlem)", "Endekslenmiş önceki yıl BM (TL/işlem)",
                   "Birim maliyet iyileşmesi (%)", "Dönem endeksi", "Müdürlük gider ağırlığı",
                   "Direktörlük gider ağırlığı", "Bütçe BM (TL/işlem)", "k"])
    for j in range(1, MAXL + 1):
        r = j + 1
        k = f"$O{r}"
        ws[f"O{r}"] = f"=IFERROR(MATCH({j},{drng('AK')},0),\"\")"
        lvl = f"INDEX({drng('D')},{k})"
        brow = f"MATCH(INDEX({drng('G')},{k}),{brng('C')},0)"
        only_budget = f'OR({k}="",{BUTCEYILI}=0)'
        ws[f"A{r}"] = "=" + dval("C", k)
        ws[f"B{r}"] = f'=IF({k}="","",IF({lvl}="Bölüm",INDEX({drng("E")},{k}),{lvl}))'
        ws[f"C{r}"] = "=" + dval("Z", k)
        ws[f"D{r}"] = "=" + dval("AA", k)
        ws[f"E{r}"] = "=" + dval("AJ", k)
        ws[f"F{r}"] = f'=IF({only_budget},"",{dval("AB", k)})'
        ws[f"G{r}"] = f'=IF({only_budget},"",{dval("AC", k)})'
        ws[f"H{r}"] = f'=IF({num(f"F{r}", f"G{r}")},IF(G{r}=0,"",F{r}/G{r}),"")'
        adj = f"INDEX({drng('AD')},{k})"
        ws[f"I{r}"] = f'=IF({only_budget},"",IF({num(adj, f"G{r}")},IF(G{r}=0,"",{adj}/G{r}),""))'
        ar = f"INDEX({drng('AI')},{k})"
        ws[f"J{r}"] = (f'=IF({only_budget},"",IF({lvl}="Bölüm",{bnum("AG", brow)},IF({lvl}="Müdürlük toplamı",'
                       f'{dval("AF", k)},IF({lvl}="Direktörlük toplamı",{dval("AH", k)},'
                       f'IF({num(ar, f"C{r}")},IF(C{r}=0,"",{ar}/C{r}),"")))))')
        ws[f"K{r}"] = f'=IF({only_budget},"",IF({lvl}="Bölüm",{bnum("W", brow)},""))'
        ws[f"L{r}"] = (f'=IF({only_budget},"",IF({lvl}="Bölüm",{bnum("AK", brow)},'
                       f'IF({lvl}="Müdürlük toplamı",{dval("AE", k)},"")))')
        # Müdürlük dosyasında diğer bölümlerin verisi olmadığından direktörlük ağırlığı hesaplanmaz
        ws[f"M{r}"] = (f'=IF(OR({only_budget},{H}!$B$12<>""),"",IF({lvl}="Bölüm",{bnum("AM", brow)},'
                       f'IF(OR({lvl}="Müdürlük toplamı",{lvl}="Direktörlük toplamı"),{dval("AG", k)},"")))')
        ws[f"N{r}"] = f'=IF({only_budget},"",IF({lvl}="Bölüm",{bnum("AH", brow)},""))'
    ws["Q1"] = f'=IF({H}!$B$12<>"","Direktörlük gider ağırlığı yalnız direktörlük raporunda hesaplanır.","")'
    ws["Q1"].font = NOTE
    last = MAXL + 1
    fmt(ws, "CEFHIN", 2, last, MONEY)
    fmt(ws, "DG", 2, last, COUNT)
    fmt(ws, "JLM", 2, last, PCT)
    fmt(ws, "K", 2, last, "0.000")
    total_rows_style(ws, f"A2:N{last}", "B2")
    ws.column_dimensions["O"].hidden = True
    widths(ws, {"A": 30, "B": 22, **{L(c): 17 for c in range(3, 15)}})
    ws.freeze_panes = "A2"


def build_onemli(wb):
    ws = wb.create_sheet("Önemli Sapmalar")
    header(ws, 1, ["Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Gider grubu", "Gerçekleşen (TL)", "Bütçe (TL)",
                   "Gerçekleşen − bütçe (TL)", "Sapma (%)", "Yön", "Açıklama durumu", "k"])
    for j in range(1, MAXM + 1):
        r = j + 1
        k = f"$L{r}"
        ws[f"L{r}"] = f"=IFERROR(MATCH({j},{krng('S')},0),\"\")"
        for col, src in zip("ABCDEFGHIJ", ["D", "C", "E", "F", "G", "K", "L", "O", "P", "X"]):
            ws[f"{col}{r}"] = f'=IF({k}="","",INDEX({krng(src)},{k}))'
        ws[f"K{r}"] = (f'=IF({k}="","",IF(INDEX({krng("AA")},{k})="","Açıklama gerekli",'
                       f'IF(LEFT(INDEX({krng("AA")},{k}),15)="[Onay bekliyor]","Onay bekliyor","Açıklandı")))')
    ws["N1"] = "Önemlilik eşiği"
    ws["N1"].font = BOLD
    ws["N2"] = f'="≥ "&FIXED({ESIK},0)&" TL ve ≥ %"&{ORAN}*100&" (tutar ve oran birlikte)"'
    ws["N3"] = f'=IF({BUTCEYILI}=0,"Seçili yıl bütçe yılı değil; önemlilik kontrolü uygulanmadı.",IF(COUNT($L$2:$L${MAXM + 1})=0,"Önemlilik eşiğini aşan sapma yok.",""))'
    last = MAXM + 1
    fmt(ws, "FGH", 2, last, MONEY)
    fmt(ws, "I", 2, last, PCT)
    direction_colors(ws, f"J2:J{last}")
    ws.conditional_formatting.add(f"K2:K{last}", FormulaRule(formula=['K2="Açıklama gerekli"'], fill=PatternFill("solid", bgColor="F8CBAD")))
    ws.conditional_formatting.add(f"K2:K{last}", FormulaRule(formula=['K2="Onay bekliyor"'], fill=PatternFill("solid", bgColor="FFEB9C")))
    for j in range(1, MAXM + 1):  # açıklama bekleyenlerin sırası (Açıklamalar sayfasındaki liste için)
        r = j + 1
        ws[f"M{r}"] = f'=IF(OR(K{r}="",K{r}="Açıklandı"),"",COUNTIF($K$2:K{r},"Açıklama gerekli")+COUNTIF($K$2:K{r},"Onay bekliyor"))'
    ws.column_dimensions["L"].hidden = True
    ws.column_dimensions["M"].hidden = True
    widths(ws, {"A": 20, "B": 18, "C": 9, "D": 28, "E": 20, "F": 18, "G": 18, "H": 20, "I": 11, "J": 11, "K": 18, "N": 50})
    ws.freeze_panes = "A2"


def build_tahmin(wb):
    ws = wb.create_sheet("Yıl Sonu Tahmini")
    header(ws, 1, ["Bölüm / müdürlük / direktörlük", "Satır türü", "Gerçekleşen Ocak–kapanış (TL)",
                   "Bütçe Ocak–kapanış (TL)", "Gerçekleşen − bütçe Ocak–kapanış (TL)", "Yıllık bütçe (TL)",
                   "Yıl sonu tahmini: gerçekleşen + kalan bütçe (TL)", "Tahmin − yıllık bütçe (TL)", "Yön",
                   "Yıl sonu tahmini: aylık ortalama gidiş (TL)", "Ortalama gidiş − yıllık bütçe (TL)", "k"])
    for j in range(1, MAXL + 1):
        r = j + 1
        k = f"$L{r}"
        ws[f"L{r}"] = f"=IFERROR(MATCH({j},{drng('I')},0),\"\")"
        ws[f"A{r}"] = "=" + dval("C", k)
        ws[f"B{r}"] = "=" + dval("D", k)
        ws[f"C{r}"] = "=" + dval("R", k)
        ws[f"D{r}"] = "=" + dval("S", k)
        ws[f"E{r}"] = f'=IF({num(f"C{r}", f"D{r}")},C{r}-D{r},"")'
        ws[f"F{r}"] = "=" + dval("U", k)
        ws[f"G{r}"] = "=" + dval("V", k)
        ws[f"H{r}"] = f'=IF({num(f"G{r}", f"F{r}")},G{r}-F{r},"")'
        ws[f"I{r}"] = f'=IF(ISNUMBER(H{r}),IF(H{r}>0,"Aleyhte",IF(H{r}<0,"Lehte","Bütçede")),"")'
        ws[f"J{r}"] = "=" + dval("W", k)
        ws[f"K{r}"] = f'=IF({num(f"J{r}", f"F{r}")},J{r}-F{r},"")'
    last = MAXL + 1
    fmt(ws, "CDEFGHJK", 2, last, MONEY)
    direction_colors(ws, f"I2:I{last}")
    total_rows_style(ws, f"A2:K{last}", "B2")
    ws["N1"] = f'="Kapanış yılı "&{KAPYIL}&" için; Ocak–"&INDEX(\'_H\'!$C$19:$N$19,1,MAX(1,N({KAPAY})))&". Seçili ay listesinden bağımsızdır."'
    ws["N1"].font = NOTE
    ws.column_dimensions["L"].hidden = True
    widths(ws, {"A": 30, "B": 20, **{L(c): 19 for c in range(3, 12)}})
    ws.freeze_panes = "A2"


def build_hacim(wb):
    ws = wb.create_sheet("Hacim ve Birim Maliyet")
    header(ws, 1, ["Bölüm / müdürlük / direktörlük", "Satır türü", "Gerçekleşen gider (TL)", "Bütçe gider (TL)",
                   "Gerçekleşen − bütçe (TL)", "Gerçekleşen işlem adedi", "Bütçe işlem adedi",
                   "Gerçekleşen birim maliyet (TL/işlem)", "Bütçe birim maliyet (TL/işlem)", "Hacim etkisi (TL)",
                   "Birim maliyet etkisi (TL)", "k"])
    for j in range(1, MAXL + 1):
        r = j + 1
        k = f"$L{r}"
        lvl = f"INDEX({drng('D')},{k})"
        brow = f"MATCH(INDEX({drng('G')},{k}),{brng('C')},0)"
        off = f'OR({k}="",{BUTCEYILI}=0)'
        ws[f"L{r}"] = f"=IFERROR(MATCH({j},{drng('I')},0),\"\")"
        ws[f"A{r}"] = "=" + dval("C", k)
        ws[f"B{r}"] = "=" + dval("D", k)
        ws[f"C{r}"] = f'=IF({off},"",{dval("O", k)})'
        ws[f"D{r}"] = f'=IF({off},"",{dval("Q", k)})'
        ws[f"E{r}"] = f'=IF({num(f"C{r}", f"D{r}")},C{r}-D{r},"")'
        ws[f"F{r}"] = f'=IF({off},"",IF({lvl}="Bölüm",{bnum("Q", brow)},""))'
        ws[f"G{r}"] = f'=IF({off},"",IF({lvl}="Bölüm",{bnum("T", brow)},""))'
        ws[f"H{r}"] = f'=IF({num(f"C{r}", f"F{r}")},IF(F{r}=0,"",C{r}/F{r}),"")'
        ws[f"I{r}"] = f'=IF({num(f"D{r}", f"G{r}")},IF(G{r}=0,"",D{r}/G{r}),"")'
        ws[f"J{r}"] = f'=IF({off},"",IF({lvl}="Bölüm",{bnum("AQ", brow)},{dval("X", k)}))'
        ws[f"K{r}"] = f'=IF({off},"",IF({lvl}="Bölüm",{bnum("AR", brow)},{dval("Y", k)}))'
    last = MAXL + 1
    fmt(ws, "CDEHIJK", 2, last, MONEY)
    fmt(ws, "FG", 2, last, COUNT)
    total_rows_style(ws, f"A2:K{last}", "B2")
    ws["N1"] = ("Hacim etkisi = (gerçekleşen adet − bütçe adet) × bütçe birim maliyeti. Birim maliyet etkisi = "
                "gerçekleşen gider − bütçe birim maliyeti × gerçekleşen adet. Toplamları gerçekleşen − bütçe farkını verir.")
    ws["N1"].font = NOTE
    ws.column_dimensions["L"].hidden = True
    widths(ws, {"A": 30, "B": 20, **{L(c): 18 for c in range(3, 12)}})
    ws.freeze_panes = "A2"


def build_kontrol(wb):
    ws = wb.create_sheet("Veri Kontrolü")
    G_ = "'_KGrup'"

    def g(col):
        return f"{G_}!${col}$2:${col}${MAXG + 1}"

    ws["A1"] = "Veri kontrolü: kaynak dip toplamı ↔ kontrol dosyası"
    ws["A1"].font = TITLE
    header(ws, 3, ["Kontrol bölümü", "Müdürlük", "Direktörlük", "Ay", "Alınan dip toplam TL", "Kontrol TL", "Fark TL",
                   "Durum", "Kaynak bölümler", "grup", "ay"])
    n = MAXG * 12
    for j in range(1, n + 1):
        r = j + 3
        ws[f"J{r}"] = f'=IF({SELCOUNT}=0,"",IFERROR(MATCH(INT(({j}-1)/{SELCOUNT})+1,{g("L")},0),""))'
        ws[f"K{r}"] = f'=IF(J{r}="","",INDEX({SELLIST},1,MOD({j}-1,{SELCOUNT})+1))'
        kk, mm = f"$J{r}", f"$K{r}"
        ws[f"A{r}"] = f'=IF({kk}="","",INDEX({g("C")},{kk}))'
        ws[f"B{r}"] = f'=IF({kk}="","",INDEX({g("G")},{kk}))'
        ws[f"C{r}"] = f'=IF({kk}="","",INDEX({g("F")},{kk}))'
        ws[f"D{r}"] = f'=IF({kk}="","",{YIL}&"-"&RIGHT("0"&{mm},2))'
        ws[f"E{r}"] = f"=IF({kk}=\"\",\"\",INDEX({G_}!$O$2:$Z${MAXG + 1},{kk},{mm}))"
        ws[f"F{r}"] = f"=IF({kk}=\"\",\"\",INDEX({G_}!$AA$2:$AL${MAXG + 1},{kk},{mm}))"
        ws[f"G{r}"] = f'=IF({num(f"E{r}", f"F{r}")},E{r}-F{r},"")'
        matches = f"INDEX({g('M')},{kk})"
        ws[f"H{r}"] = (f'=IF({kk}="","",IF({matches}=0,"Bölüm eşleşmedi",IF({matches}>1,"Birden fazla bölüm satırı eşleşti",'
                       f'IF(NOT(ISNUMBER(E{r})),"Kaynak dip toplamı eksik",IF(NOT(ISNUMBER(F{r})),"Kontrol tutarı veya ay bulunamadı",'
                       f'IF(ABS(G{r})>{TOLERANS},"Veri doğrulanamadı: tutar farklı","Doğrulandı"))))))')
        ws[f"I{r}"] = f'=IF({kk}="","",INDEX({g("H")},{kk}))'
    last = n + 3
    fmt(ws, "EFG", 4, last, MONEY)
    ws.conditional_formatting.add(f"H4:H{last}", FormulaRule(formula=['AND(H4<>"",H4<>"Doğrulandı")'], fill=PatternFill("solid", bgColor="FFC7CE")))
    ws.conditional_formatting.add(f"H4:H{last}", FormulaRule(formula=['H4="Doğrulandı"'], fill=PatternFill("solid", bgColor="C6EFCE")))
    ws["M1"] = "Ciro kontrolü"
    ws["M1"].font = Font(bold=True, color="233C57")
    header(ws, 3, ["Dönem", "Gerçekleşen ciro", "Bütçe ciro"], 13)
    crow = find(chr(34) + "Ciro|" + chr(34) + "&" + YIL)
    for i in range(1, 13):
        r = i + 3
        mi = f"INDEX({SELLIST},1,{i})"
        ws[f"M{r}"] = f'=IF({mi}="","",{YIL}&"-"&RIGHT("0"&{mi},2))'
        ws[f"N{r}"] = f'=IF(OR({mi}="",{crow}=""),"",IF(ISNUMBER(INDEX({VG},{crow},{mi})),INDEX({VG},{crow},{mi}),""))'
        ws[f"O{r}"] = f'=IF(OR({mi}="",{crow}=""),"",IF(ISNUMBER(INDEX({VB},{crow},{mi})),INDEX({VB},{crow},{mi}),""))'
    fmt(ws, "NO", 4, 15, MONEY)
    ws.column_dimensions["J"].hidden = True
    ws.column_dimensions["K"].hidden = True
    widths(ws, {"A": 26, "B": 22, "C": 24, "D": 10, "E": 18, "F": 18, "G": 14, "H": 30, "I": 30, "M": 12, "N": 18, "O": 18})
    ws.freeze_panes = "A4"


def build_revizyon(wb):
    ws = wb.create_sheet("Revizyonlar")
    R_ = "'_Rev'"
    header(ws, 1, ["Kapanış", "Etkilenen ay", "Müdürlük", "Bölüm", "İndeks", "Alt gider kalemi", "Tür", "Eski TL",
                   "Yeni TL", "Fark TL", "Önceki dosya → yeni dosya", "k"])
    for j in range(1, MAXR + 1):
        r = j + 1
        k = f"$L{r}"
        ws[f"L{r}"] = f"=IFERROR(INDEX({R_}!$B$2:$B${MAXR + 1},MATCH({j},{R_}!$F$2:$F${MAXR + 1},0)),\"\")"
        ws[f"A{r}"] = f'=IF({k}="","",INDEX({vcol("Metin")},{k}))'
        ws[f"B{r}"] = f'=IF({k}="","",INDEX({vcol("Yıl")},{k})&"-"&RIGHT("0"&INDEX({vcol("Sayı")},{k}),2))'
        ws[f"C{r}"] = f'=IF({k}="","",INDEX({vcol("Müdürlük")},{k}))'
        ws[f"D{r}"] = f'=IF({k}="","",INDEX({vcol("Bölüm")},{k}))'
        ws[f"E{r}"] = f'=IF({k}="","",INDEX({vcol("Kod")},{k}))'
        ws[f"F{r}"] = f'=IF({k}="","",INDEX({vcol("Ad")},{k}))'
        ws[f"G{r}"] = f'=IF({k}="","",INDEX({vcol("Grup")},{k}))'
        ws[f"H{r}"] = f'=IF({k}="","",IF(ISNUMBER(INDEX({vcol("G1")},{k})),INDEX({vcol("G1")},{k}),""))'
        ws[f"I{r}"] = f'=IF({k}="","",IF(ISNUMBER(INDEX({vcol("G2")},{k})),INDEX({vcol("G2")},{k}),""))'
        ws[f"J{r}"] = f'=IF({num(f"H{r}", f"I{r}")},I{r}-H{r},"")'
        ws[f"K{r}"] = f'=IF({k}="","",INDEX({vcol("Kaynak")},{k}))'
    fmt(ws, "HIJ", 2, MAXR + 1, MONEY)
    ws.column_dimensions["L"].hidden = True
    widths(ws, {"A": 10, "B": 12, "C": 18, "D": 16, "E": 8, "F": 26, "G": 12, "H": 16, "I": 16, "J": 14, "K": 50})
    ws.freeze_panes = "A2"


def build_uyarilar(wb, last_uy):
    ws = wb.create_sheet("Uyarılar")
    U_ = "'_Uy'"
    header(ws, 1, ["Sıra", "Dönem", "Bölüm", "Kaynak", "Uyarı", "k"])
    for j in range(1, MAXW + MAXD + MAXG + 3):
        r = j + 1
        k = f"$F{r}"
        ws[f"F{r}"] = f"=IFERROR(MATCH({j},{U_}!$H$2:$H${last_uy},0),\"\")"
        ws[f"A{r}"] = f'=IF({k}="","",{j})'
        for col, src in zip("BCDE", "CDEF"):
            ws[f"{col}{r}"] = f'=IF({k}="","",INDEX({U_}!${src}$2:${src}${last_uy},{k}))'
    ws.column_dimensions["F"].hidden = True
    widths(ws, {"A": 6, "B": 10, "C": 18, "D": 28, "E": 90})
    ws.freeze_panes = "A2"


def build_aciklama(wb):
    ws = wb.create_sheet("Açıklamalar")
    header(ws, 1, ["Bölüm", "Gider kodu", "Gider kalemi", "Açıklama", "Yazan", "Tarih", "Durum", "anahtar",
                   "Gerçekleşen (seçili aylar)", "Bütçe (seçili aylar)", "Aylar", "Kapanış", "Dosya müdürlüğü"])
    for r in range(2, MAXA + 2):
        # Açıklamanın yazıldığı andaki tutar ve seçim: ana rapor bunlarla "tutar değişti" kontrolü yapar.
        ws[f"I{r}"] = f'=IF(H{r}="","",IFERROR(INDEX({krng("K")},MATCH(H{r},{krng("H")},0)),""))'
        ws[f"J{r}"] = f'=IF(H{r}="","",IFERROR(INDEX({krng("L")},MATCH(H{r},{krng("H")},0)),""))'
        ws[f"K{r}"] = f'=IF(H{r}="","",{H}!$B$20)'
        ws[f"L{r}"] = f'=IF(H{r}="","",{H}!$B$21)'
        ws[f"M{r}"] = f'=IF(H{r}="","",{H}!$B$12)'
        ws[f"C{r}"] = f'=IF(OR(A{r}="",B{r}=""),"",IFERROR(INDEX({krng("F")},MATCH(A{r}&"|"&B{r},{krng("H")},0)),"Bu bölümde bu kod yok"))'
        ws[f"H{r}"] = f'=IF(OR(A{r}="",B{r}=""),"",A{r}&"|"&B{r})'
        ws[f"G{r}"] = (f'=IF(H{r}="","",IF(IFERROR(INDEX({krng("Q")},MATCH(H{r},{krng("H")},0)),0)=1,'
                       f'"Önemli sapma","Bilgi"))')
        for c in "ABDEF":
            ws[f"{c}{r}"].fill = INPUT_FILL
        ws[f"D{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    dv = DataValidation(type="list", formula1=f"{brng('C')}", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"A2:A{MAXA + 1}")
    ws["O1"] = ("Bölüm ve gider kodunu yazın, açıklamayı D sütununa girin. Bu dosyanın Giderler ve Önemli Sapmalar "
                "sayfaları buradan okur. Müdürlük dosyasında yazılan açıklamalar, dosya geri gönderildiğinde ana "
                "rapordaki Butce_Aciklama sorgusuyla toplanır. Gizli I–M sütunları otomatik doldurulur.")
    ws["O1"].font = NOTE
    ws["O1"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("O1:S2")
    ws["O3"] = "Açıklama bekleyen önemli sapmalar"
    ws["O3"].font = BOLD
    header(ws, 4, ["Bölüm", "Gider kodu", "Gider kalemi", "Gerçekleşen − bütçe (TL)", "Durum"], 15)
    for j in range(1, MAXM + 1):  # Önemli Sapmalar'dan, açıklaması olmayan veya onay bekleyenler
        r = j + 4
        k = f"IFERROR(MATCH({j},'Önemli Sapmalar'!$M$2:$M${MAXM + 1},0),\"\")"
        for col, src in zip("OPQRS", "BCDHK"):
            ws[f"{col}{r}"] = f"=IF({k}=\"\",\"\",INDEX('Önemli Sapmalar'!${src}$2:${src}${MAXM + 1},{k}))"
        ws[f"R{r}"].number_format = MONEY
    ws.column_dimensions["H"].hidden = True
    for c in "IJKLM":
        ws.column_dimensions[c].hidden = True
    widths(ws, {"A": 18, "B": 10, "C": 28, "D": 60, "E": 14, "F": 12, "G": 14, "O": 18, "P": 9, "Q": 28, "R": 18, "S": 16})
    ws.freeze_panes = "A2"


def build_onay(wb):
    """Açıklama Onayı: A..W Butce_Aciklama sorgusunun tablosu (A1'e yüklenir), Y..AA kontrol formülleri."""
    ws = wb.create_sheet("Açıklama Onayı")
    for i, t in enumerate(["kalem satırı", "Güncel gerçekleşen", "Güncel bütçe", "Kontrol"]):
        ws.cell(1, 24 + i, t)
    header(ws, 1, ["Güncel gerçekleşen", "Güncel bütçe", "Kontrol"], 25)
    A, C, D, I, J, K, Lc, O, Pc = (OC[n] for n in ("Anahtar", "Bölüm", "Gider kodu", "Kapanış", "Aylar",
                                                    "Gerçekleşen (dosyada)", "Bütçe (dosyada)", "Okuma notu", "Kaynakta"))
    for r in range(2, MAXO + 2):
        ws[f"X{r}"] = (f'=IF(OR(${A}{r}="",LEFT(${A}{r},6)="Dosya|"),"",'
                       f'IFERROR(MATCH(${C}{r}&"|"&${D}{r},{krng("H")},0),""))')
        ws[f"Y{r}"] = f'=IF(X{r}="","",INDEX({krng("K")},X{r}))'
        ws[f"Z{r}"] = f'=IF(X{r}="","",INDEX({krng("L")},X{r}))'
        changed = (f'OR(ISNUMBER(Y{r})<>ISNUMBER(${K}{r}),ISNUMBER(Z{r})<>ISNUMBER(${Lc}{r}),'
                   f'ABS(N(Y{r})-N(${K}{r}))>{TOLERANS},ABS(N(Z{r})-N(${Lc}{r}))>{TOLERANS})')
        ws[f"AA{r}"] = (f'=IF(${A}{r}="","",IF(${O}{r}<>"",${O}{r},IF(${I}{r}<>{H}!$B$21,"Önceki kapanış ("&${I}{r}&")",'
                        f'IF(${J}{r}<>{H}!$B$20,"Ay seçimi farklı (dosya: "&${J}{r}&")",'
                        f'IF(X{r}="","Bölüm/kod bu raporda yok",'
                        f'IF(AND(${Pc}{r}="Aktarıldı",${K}{r}="",${Lc}{r}=""),"Güncel (önceki dönemden aktarıldı)",'
                        f'IF({changed},"Tutar değişti","Güncel")))))))')
    fmt(ws, "YZ", 2, MAXO + 1, MONEY)
    last = MAXO + 1
    ws.conditional_formatting.add(f"AA2:AA{last}", FormulaRule(formula=['LEFT(AA2,6)="Güncel"'], fill=PatternFill("solid", bgColor="C6EFCE")))
    ws.conditional_formatting.add(f"AA2:AA{last}", FormulaRule(formula=['AND(AA2<>"",LEFT(AA2,6)<>"Güncel")'], fill=PatternFill("solid", bgColor="FFEB9C")))
    dv = DataValidation(type="list", formula1='"' + ",".join(DURUMLAR) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{OC['Durum']}2:{OC['Durum']}{last}")
    for c in (OC["Durum"], OC["Onay notu"]):
        for r in range(2, 202):
            ws[f"{c}{r}"].fill = INPUT_FILL
    ws["AC1"] = ("Bu sayfanın A1 hücresine Butce_Aciklama sorgusu yüklenir. Durum (sarı) ve Onay notu sütunlarını siz "
                 "doldurursunuz; Tümünü Yenile'de korunur ve işlem geçmişine Kontrol Paneli'ndeki adınızla yazılır. "
                 "Rapor, yalnız Kontrol = Güncel olan Onaylandı açıklamalarını gösterir (Bekliyor: [Onay bekliyor]). "
                 "Bu döneme aktar: önceki kapanışın açıklamasını bu kapanışa onaylı olarak taşır. "
                 "Kayıtlı durum sütununu değiştirmeyin.")
    ws["AC1"].font = NOTE
    ws["AC1"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("AC1:AH8")
    ws.column_dimensions["X"].hidden = True
    widths(ws, {"Y": 18, "Z": 18, "AA": 30, "AC": 16})
    ws.freeze_panes = "B2"


def build_dagitim(wb):
    ws = wb.create_sheet("Dağıtım")
    ws["A1"] = "Müdürlük dosyalarını hazırlama"
    ws["A1"].font = TITLE
    steps = [
        "Müdürlük dosyası yalnız o müdürlüğün verisini içerir; diğer bölümlerin tutarları dosyaya girmez.",
        "1. Bu raporu Dosya › Farklı Kaydet ile aşağıdaki önerilen adla kaydedin.",
        "2. Yeni dosyada Veri › Sorgular ve Bağlantılar › Butce_Veri › sağ tık › Düzenle › Gelişmiş Düzenleyici.",
        "3. MudurlukDosyasi satırına aşağıdaki değeri yazın › Bitti › Kapat ve Yükle. Kaydedip müdürlüğe gönderin.",
        "4. Müdürlük açıklamaları dosyadaki Açıklamalar sayfasına yazar ve dosyayı geri gönderir.",
        "5. Dönen dosyaları açıklama klasörüne koyun; ana raporda Tümünü Yenile › Açıklama Onayı sayfasında onaylayın.",
    ]
    for i, text in enumerate(steps, 2):
        ws.cell(i, 1, text).font = NOTE if i == 2 else Font(color="263746")
    header(ws, 9, ["Müdürlük", "Önerilen dosya adı", "Sorgudaki satır", "Bölüm sayısı"])
    for i in range(1, MAXD + 1):
        r = i + 9
        m = f"'_Liste'!$A${i + 1}"
        ws[f"A{r}"] = f'=IF({m}="","",{m})'
        ws[f"B{r}"] = f'=IF(A{r}="","",A{r}&"_"&{H}!$B$21&".xlsx")'
        ws[f"C{r}"] = f'=IF(A{r}="","","MudurlukDosyasi = "&CHAR(34)&A{r}&CHAR(34)&",")'
        ws[f"D{r}"] = f'=IF(A{r}="","",COUNTIF({brng("D")},A{r}))'
    widths(ws, {"A": 28, "B": 40, "C": 40, "D": 12})
    ws.sheet_view.showGridLines = False


def build_girdi(wb):
    ws = wb.create_sheet("Girdi Kaydı")
    ws["A1"] = "Raporun okuduğu dosyalar"
    ws["A1"].font = TITLE
    ws["A2"] = ("Her yenilemede güncellenir. Kaynak dosya sonradan değişirse boyut veya değiştirilme zamanı farklı olur; "
                "aylık raporu Farklı Kaydet ile saklamak, o ayın girdileriyle birlikte arşivlenmesini sağlar.")
    ws["A2"].font = NOTE
    header(ws, 4, ["Tür", "Dosya", "Dönem", "Boyut (KB)", "Değiştirilme zamanı", "Durum"])
    for k in range(1, MAXF + 1):
        r = k + 4
        f = find(f'"Girdi|{k}"')
        ws[f"G{r}"] = "=" + f
        g = f"$G{r}"
        ws[f"A{r}"] = f'=IF({g}="","",INDEX({vcol("Grup")},{g})&"")'
        ws[f"B{r}"] = f'=IF({g}="","",INDEX({vcol("Ad")},{g})&"")'
        ws[f"C{r}"] = (f'=IF({g}="","",IF(ISNUMBER(INDEX({vcol("Yıl")},{g})),INDEX({vcol("Yıl")},{g})&"-"&'
                       f'RIGHT("0"&INDEX({vcol("Sayı")},{g}),2),""))')
        ws[f"D{r}"] = f'=IF({g}="","",IF(ISNUMBER(INDEX({vcol("Kod")},{g})),INDEX({vcol("Kod")},{g})/1024,""))'
        ws[f"E{r}"] = f'=IF({g}="","",INDEX({vcol("Metin")},{g})&"")'
        ws[f"F{r}"] = f'=IF({g}="","",INDEX({vcol("Kaynak")},{g})&"")'
    fmt(ws, "D", 5, MAXF + 4, "#,##0.0")
    ws.column_dimensions["G"].hidden = True
    widths(ws, {"A": 12, "B": 40, "C": 10, "D": 12, "E": 20, "F": 18})
    ws.freeze_panes = "A5"


def build_ozet(wb):
    ws = wb.create_sheet("Yönetici Özeti", 1)
    Bo = brng
    ws["A1"] = "Yönetici özeti"
    ws["A1"].font = TITLE
    ws["A2"] = f'=IF({KAPYIL}="","Veri yok",{KAPSAM}&" · "&{SELCOUNT}&" ay ("&{YIL}&") · Kaynak kapanış "&{KAPYIL}&"-"&RIGHT("0"&{KAPAY},2))'
    ws["A2"].font = NOTE
    header(ws, 4, ["Gösterge", "Değer"])
    in_scope = f"{Bo('I')},1"
    actual = (f'IF(OR(COUNTIF({Bo("I")},1)=0,SUMIFS({Bo(BOLUM_COLS["eksik A"])},{in_scope})>0),"",SUMIFS({Bo("L")},{in_scope}))')
    budget = (f'IF(OR({BUTCEYILI}=0,COUNTIF({Bo("I")},1)=0,SUMIFS({Bo(BOLUM_COLS["eksik B"])},{in_scope})>0),"",SUMIFS({Bo("M")},{in_scope}))')
    top = f"MATCH(MAX({drng('I')}),{drng('I')},0)"
    rows = [
        ("Gerçekleşen (TL)", "=" + actual, MONEY),
        ("Bütçe (TL)", "=" + budget, MONEY),
        ("Gerçekleşen − bütçe (TL)", '=IF(AND(ISNUMBER(B5),ISNUMBER(B6)),B5-B6,"")', MONEY),
        ("Sapma oranı (%)", '=IF(AND(ISNUMBER(B7),ISNUMBER(B6)),IF(B6=0,"",B7/B6),"")', PCT),
        ("Yön", f'=IF({BUTCEYILI}=0,"Bütçe yılı dışında",IF(ISNUMBER(B7),IF(B7>0,"Aleyhte",IF(B7<0,"Lehte","Bütçede")),""))', None),
        ("Yıl sonu tahmini: gerçekleşen + kalan bütçe (TL)", f'=IF({BUTCEYILI}=0,"",IFERROR(INDEX({drng("V")},{top}),""))', MONEY),
        ("Yıllık bütçe (TL)", f'=IF({BUTCEYILI}=0,"",IFERROR(INDEX({drng("U")},{top}),""))', MONEY),
        ("Tahmin − yıllık bütçe (TL)", '=IF(AND(ISNUMBER(B10),ISNUMBER(B11)),B10-B11,"")', MONEY),
        ("Veri kontrolü", "=COUNTIF('Veri Kontrolü'!$H$4:$H$723,\"Doğrulandı\")&\" / \"&COUNTIF('Veri Kontrolü'!$H$4:$H$723,\"?*\")&\" kontrol doğrulandı\"", None),
        ("Uyarı sayısı", f"=COUNT('Uyarılar'!$A$2:$A${MAXW + MAXD + MAXG + 3})", "0"),
        ("Önemli sapma", f'=IF({BUTCEYILI}=0,"Uygulanmadı",COUNT({krng("S")})&" kalem ("&COUNT({krng("U")})&" aleyhte)")', None),
        ("Önemlilik eşiği", f'="≥ "&FIXED({ESIK},0)&" TL ve ≥ %"&{ORAN}*100', None),
    ]
    for i, (label, formula, number_format) in enumerate(rows, 5):
        ws.cell(i, 1, label)
        cell = ws.cell(i, 2, formula)
        if number_format:
            cell.number_format = number_format
    direction_colors(ws, "B9")
    cols = ["Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Gerçekleşen (TL)", "Bütçe (TL)", "Gerçekleşen − bütçe (TL)", "Sapma (%)"]
    r = 18
    for title, rank in (("En büyük aleyhte önemli sapmalar", "U"), ("En büyük lehte önemli sapmalar", "W")):
        ws.cell(r, 1, title).font = Font(bold=True, color="233C57")
        header(ws, r + 1, cols)
        for j in range(1, 11):
            rr = r + 1 + j
            k = f"IFERROR(MATCH({j},{krng(rank)},0),\"\")"
            ws[f"J{rr}"] = f"=IF({j}>{OZETN},\"\",{k})"
            for col, src in zip("ABCDEFGH", ["D", "C", "E", "F", "K", "L", "O", "P"]):
                ws[f"{col}{rr}"] = f'=IF($J{rr}="","",INDEX({krng(src)},$J{rr}))'
            ws[f"E{rr}"].number_format = ws[f"F{rr}"].number_format = ws[f"G{rr}"].number_format = MONEY
            ws[f"H{rr}"].number_format = PCT
        ws[f"A{r + 2}"] = (f'=IF($J{r + 2}<>"",INDEX({krng("D")},$J{r + 2}),IF({BUTCEYILI}=0,"Bütçe yılı dışında","Yok"))')
        r += 13
    ws.column_dimensions["J"].hidden = True
    widths(ws, {"A": 46, "B": 30, "C": 11, "D": 30, "E": 20, "F": 20, "G": 22, "H": 12})
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    from openpyxl.worksheet.properties import PageSetupProperties

    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_area = "A1:H44"


def build_pano(wb):
    ws = wb.create_sheet("Pano", 2)
    ws["A1"] = "Pano"
    ws["A1"].font = TITLE
    ws["A2"] = ("Filtreleri B4:B6 açılır listelerinden seçin. Kapsam ve aylar Kontrol Paneli'nden gelir. "
                "Eksik aylık tutarlar 0 sayılır; kesin toplamlar Genel Özet'tedir.")
    ws["A2"].font = NOTE
    for r, (label, col) in enumerate((("Müdürlük", "A"), ("Bölüm", "C"), ("Gider grubu", "B")), 4):
        ws.cell(r, 1, label).font = BOLD
        cell = ws.cell(r, 2, "Tümü")
        cell.fill = INPUT_FILL
        dv = DataValidation(type="list", formula1=f"'_Liste'!${col}$1:${col}${MAXD + 1}")
        ws.add_data_validation(dv)
        dv.add(cell)
    crit_m = f'IF($B$4="Tümü",IF({KAPSAM}="Tümü","*",{KAPSAM}),$B$4)'
    crit_d = 'IF($B$5="Tümü","*",$B$5)'
    crit_g = 'IF($B$6="Tümü","*",$B$6)'
    kal = f"{krng('D')},{crit_m},{krng('C')},{crit_d},{krng('G')},{crit_g}"
    kpis = [
        ("Gerçekleşen (TL)", f"=SUMIFS({krng('K')},{kal})", MONEY),
        ("Bütçe (TL)", f'=IF({BUTCEYILI}=1,SUMIFS({krng("L")},{kal}),"")', MONEY),
        ("Gerçekleşen − bütçe (TL)", '=IF(E5="","",E4-E5)', MONEY),
        ("Sapma (%)", '=IF(OR(E5="",E5=0),"",E6/E5)', PCT),
        ("Yön", '=IF(E6="","",IF(E6>0,"Aleyhte",IF(E6<0,"Lehte","Bütçede")))', None),
        ("Önceki yıl aynı aylar (TL)", f"=SUMIFS({krng('N')},{kal})", MONEY),
    ]
    for r, (label, formula, number_format) in enumerate(kpis, 4):
        ws.cell(r, 4, label).font = BOLD
        cell = ws.cell(r, 5, formula)
        if number_format:
            cell.number_format = number_format
    direction_colors(ws, "E8")
    header(ws, 11, ["Ay", "Gerçekleşen (TL)", "Bütçe (TL)", "Önceki yıl (TL)", "Gerçekleşen − bütçe (TL)"])
    base = f'{vcol("Tablo")},"Gider",{vcol("Müdürlük")},{crit_m},{vcol("Bölüm")},{crit_d},{vcol("Grup")},{crit_g}'
    for i in range(1, 13):
        r = 11 + i
        mi = f"INDEX({SELLIST},1,{i})"
        ws[f"A{r}"] = f'=IF({mi}="","",INDEX(\'_H\'!$C$19:$N$19,1,{mi})&" "&{YIL})'
        ws[f"B{r}"] = f'=IF({mi}="","",SUMIFS(INDEX({VG},0,{mi}),{base},{vcol("Yıl")},{YIL}))'
        ws[f"C{r}"] = f'=IF(OR({mi}="",{BUTCEYILI}=0),"",SUMIFS(INDEX({VB},0,{mi}),{base},{vcol("Yıl")},{YIL}))'
        ws[f"D{r}"] = f'=IF({mi}="","",SUMIFS(INDEX({VG},0,{mi}),{base},{vcol("Yıl")},{YIL}-1))'
        ws[f"E{r}"] = f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(C{r})),B{r}-C{r},"")'
    fmt(ws, "BCDE", 12, 23, MONEY)
    header(ws, 25, ["Gider grubu", "Gerçekleşen (TL)", "Bütçe (TL)", "Gerçekleşen − bütçe (TL)", "Sapma (%)"])
    kal_g = f"{krng('D')},{crit_m},{krng('C')},{crit_d}"
    for i in range(1, 21):
        r = 25 + i
        ws[f"A{r}"] = f"=IF(INDEX('_Liste'!$B$2:$B$61,{i})=\"\",\"\",INDEX('_Liste'!$B$2:$B$61,{i}))"
        ws[f"B{r}"] = f'=IF(A{r}="","",SUMIFS({krng("K")},{kal_g},{krng("G")},A{r}))'
        ws[f"C{r}"] = f'=IF(OR(A{r}="",{BUTCEYILI}=0),"",SUMIFS({krng("L")},{kal_g},{krng("G")},A{r}))'
        ws[f"D{r}"] = f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(C{r})),B{r}-C{r},"")'
        ws[f"E{r}"] = f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(C{r})),IF(C{r}=0,"",D{r}/C{r}),"")'
    fmt(ws, "BCD", 26, 45, MONEY)
    fmt(ws, "E", 26, 45, PCT)
    header(ws, 47, ["Bölüm", "Müdürlük", "Gerçekleşen (TL)", "Bütçe (TL)", "Gerçekleşen − bütçe (TL)", "Sapma (%)"])
    kal_d = f"{krng('G')},{crit_g}"
    for i in range(1, MAXD + 1):
        r = 47 + i
        name = f"INDEX('_Liste'!$C$2:$C${MAXD + 1},{i})"
        ws[f"A{r}"] = f'=IF({name}="","",{name})'
        ws[f"B{r}"] = f'=IF(A{r}="","",INDEX({brng("D")},MATCH(A{r},{brng("C")},0)))'
        ws[f"C{r}"] = f'=IF(A{r}="","",SUMIFS({krng("K")},{kal_d},{krng("C")},A{r}))'
        ws[f"D{r}"] = f'=IF(OR(A{r}="",{BUTCEYILI}=0),"",SUMIFS({krng("L")},{kal_d},{krng("C")},A{r}))'
        ws[f"E{r}"] = f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}-D{r},"")'
        ws[f"F{r}"] = f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(D{r})),IF(D{r}=0,"",E{r}/D{r}),"")'
    fmt(ws, "CDE", 48, 47 + MAXD, MONEY)
    fmt(ws, "F", 48, 47 + MAXD, PCT)
    widths(ws, {"A": 24, "B": 22, "C": 20, "D": 26, "E": 22, "F": 14})

    trend = BarChart()
    trend.type = "col"
    trend.title = "Aylık gerçekleşen ve bütçe (TL)"
    trend.add_data(Reference(ws, min_col=2, max_col=3, min_row=11, max_row=23), titles_from_data=True)
    trend.set_categories(Reference(ws, min_col=1, min_row=12, max_row=23))
    prior = LineChart()
    prior.add_data(Reference(ws, min_col=4, min_row=11, max_row=23), titles_from_data=True)
    trend += prior
    trend.height, trend.width = 8, 18
    ws.add_chart(trend, "H4")
    groups = BarChart()
    groups.type = "bar"
    groups.title = "Gider grubu: gerçekleşen − bütçe (TL)"
    groups.add_data(Reference(ws, min_col=4, min_row=25, max_row=45), titles_from_data=True)
    groups.set_categories(Reference(ws, min_col=1, min_row=26, max_row=45))
    groups.legend = None
    groups.height, groups.width = 8, 18
    ws.add_chart(groups, "H21")
    ws.sheet_view.showGridLines = False


def build(rows=None):
    wb = Workbook()
    build_panel(wb)
    veri = wb.create_sheet("Veri")
    if rows:  # örnek kopya; boş şablonda A1 sorgunun yükleneceği yer olarak boş kalır
        veri.append(COLUMNS)
        for row in rows:
            veri.append(row)
    veri["AL1"] = ("Bu sayfa Power Query sorgusunun (Butce_Veri) yükleme yeridir. Sorguyu buraya, A1 hücresine "
                   "yükleyin; elle değiştirmeyin.")
    build_helpers(wb)
    build_bolum(wb)
    build_duzen(wb)
    build_aciklama(wb)
    build_kalem(wb)
    build_satir(wb)
    build_kgrup(wb)
    last_uy = build_lists(wb)
    build_genel(wb)
    build_giderler(wb)
    build_birim(wb)
    build_onemli(wb)
    build_tahmin(wb)
    build_hacim(wb)
    build_kontrol(wb)
    build_revizyon(wb)
    build_uyarilar(wb, last_uy)
    build_ozet(wb)
    build_pano(wb)
    build_onay(wb)
    build_dagitim(wb)
    build_girdi(wb)
    order = ["Kontrol Paneli", "Yönetici Özeti", "Pano", "Genel Özet", "Giderler", "Birim Maliyet", "Önemli Sapmalar",
             "Yıl Sonu Tahmini", "Hacim ve Birim Maliyet", "Veri Kontrolü", "Revizyonlar", "Uyarılar", "Açıklamalar",
             "Açıklama Onayı", "Dağıtım", "Girdi Kaydı", "Veri"]
    wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 100)
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    return wb


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--ornek", type=Path, help="örnek veriyle dolu kopyanın yolu")
    parser.add_argument("--cikti", type=Path, default=ROOT / "powerquery" / "Butce_Rapor_Sablonu.xlsx")
    args = parser.parse_args()
    if args.ornek:
        from pq_veri import build as veri

        build(veri(ROOT / "python" / "ornek" / "Ayarlar.xlsx")).save(args.ornek)
        print("örnek:", args.ornek)
    else:
        build().save(args.cikti)
        print("şablon:", args.cikti)
