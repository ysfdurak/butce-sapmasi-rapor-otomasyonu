// Butce_Veri — makrosuz ve Python'suz yol: Excel'in yerleşik Power Query'si (lisans gerekmez, Excel 2016+).
// Kurulum: docs/POWER_QUERY.md. Bu sorgu Butce_Rapor_Sablonu.xlsx içindeki Veri sayfasına (A1) TEK tablo yükler;
// şablondaki raporlar ve panolar bu tablodan formüllerle hesaplanır. Veri › Tümünü Yenile = tek tuşla güncel rapor.
//
// Okunanlar (Python ve makro sürümüyle aynı kurallar):
//   Ayarlar.xlsx       Bölüm Eşleştirme, Gider Grupları, Gider Eşleştirme, Ortak Sayfalar, Dosya ve Konum Ayarları
//   maliyet klasörü    kapanış yılındaki tüm kapanış dosyaları (Ocak kapanışında önceki Aralık da) → revizyonlar
//   Kontrol dosyası    şirket cirosu, bütçe cirosu, direktörlük "Opex Toplamı" altındaki kontrol satırları
//   Endeks dosyası     "Bölüm Gider Endeksi" başlıklı ikinci tablo
// Tablonun her satırı bir bloğa aittir (Tablo sütunu); sözleşme ve Python benzetimi: tools/pq_veri.py.
// Konumlar Ayarlar'dan alınır; sayfalar başlıklara (Bütçe, ilk ardışık tarih, Bölüm Gider Endeksi) göre
// hizalandığından kaynağın boş üst satırları/sütunları sonucu etkilemez.
// Performans: tekrar kullanılan her şey Table.Buffer / List.Buffer ile belleğe alınır.
let
    // ===================== YALNIZ BU DÖRT SATIRI DÜZENLEYİN ====================
    AyarDosyasi = "C:\Butce\ornek\Ayarlar.xlsx",   // Ayarlar.xlsx dosyasının tam yolu
    KapanisDonemi = "",                             // boş = klasördeki en son kapanış; örnek: "2026-10"
    MudurlukDosyasi = "",                           // boş = direktörlük raporu; müdürlüğe gidecek dosya: "Depolama"
    HizliMod = false,                               // true = yalnız kapanış ve önceki ay okunur (revizyon yalnız son ay)
    // ============================================================================

    AyAdlari = {"ocak", "şubat", "mart", "nisan", "mayıs", "haziran", "temmuz", "ağustos", "eylül", "ekim", "kasım", "aralık"},
    GAdlari = List.Transform({1 .. 12}, each "G" & Text.From(_)),
    BAdlari = List.Transform({1 .. 12}, each "B" & Text.From(_)),
    RAdlari = List.Transform({1 .. 12}, each "R" & Text.From(_)),
    Sutunlar = {"Anahtar", "Tablo", "Direktörlük", "Müdürlük", "Bölüm", "Ad", "Kod", "Grup", "Yıl", "Kaynak", "Metin", "Sayı"}
        & GAdlari & BAdlari & RAdlari,
    ArdisikUyari = "Ardışık dosya eksik. Bu kapanış için revizyon karşılaştırması yapılamadı.",

    // ---------------------------------------------------------------- Ortak yardımcılar
    Kucuk = (v) => if v = null then "" else Text.Lower(Text.Trim(Text.From(v)), "tr-TR"),
    Bos = (v) => v = null or (Value.Is(v, type text) and Text.Trim(v) = ""),
    AyNo = (t) =>
        let
            k = Kucuk(t),
            tam = List.PositionOf(AyAdlari, k),
            kisa = List.PositionOf(List.Transform(AyAdlari, each Text.Start(_, 3)), k)
        in
            if tam >= 0 then tam + 1 else if kisa >= 0 then kisa + 1 else null,
    PYil = (p) => Number.IntegerDivide(p, 100),
    PAy = (p) => Number.Mod(p, 100),
    Sonraki = (p) => if PAy(p) = 12 then p + 89 else p + 1,
    Etiket = (p) => Text.From(PYil(p)) & "-" & Text.PadStart(Text.From(PAy(p)), 2, "0"),
    Harf = (c as number) as text =>
        let
            a = Number.IntegerDivide(c - 1, 26),
            b = Number.Mod(c - 1, 26),
            on = if a = 0 then "" else if a <= 26 then Character.FromNumber(64 + a)
                else Character.FromNumber(64 + Number.IntegerDivide(a - 1, 26)) & Character.FromNumber(65 + Number.Mod(a - 1, 26))
        in
            on & Character.FromNumber(65 + b),
    // Katı toplam: bir değer bile boşsa sonuç boş (Python total, VBA eşdeğeri).
    KatiToplam = (vs as list) => if List.IsEmpty(vs) or List.Contains(vs, null) then null else List.Sum(vs),
    Sayi = (v) => if v <> null and Value.Is(v, type number) then v else null,
    // Metin olarak girilmiş tutar: "1.234,56", "1234,56", "(1.234)", "1234.56", "1.234 TL". Değilse null.
    MetinSayi = (v as text) =>
        let
            t0 = Text.Remove(Text.Trim(v), {" ", "#(00A0)", "₺"}),
            t1 = if Text.EndsWith(Text.Upper(t0), "TL") then Text.Start(t0, Text.Length(t0) - 2) else t0,
            negatif = Text.Length(t1) > 2 and Text.StartsWith(t1, "(") and Text.EndsWith(t1, ")"),
            t = if negatif then Text.Middle(t1, 1, Text.Length(t1) - 2) else t1,
            isaretsiz = if Text.StartsWith(t, "-") or Text.StartsWith(t, "+") then Text.Middle(t, 1) else t,
            gecerli = isaretsiz <> "" and Text.Select(isaretsiz, {"0" .. "9", ".", ","}) = isaretsiz,
            parca = Text.Split(isaretsiz, "."),
            binlik = List.Count(parca) > 1 and List.AllTrue(List.Transform(List.Skip(parca), each Text.Length(Text.BeforeDelimiter(_, ",")) = 3)),
            duz = if Text.Contains(t, ",") or binlik then Text.Replace(Text.Remove(t, "."), ",", ".") else t,
            n = if gecerli then (try Number.FromText(duz, "en-US") otherwise null) else null
        in
            if n = null then null else if negatif then -n else n,
    // Başlık hücresindeki dönem (yyyymm): tarih, Excel seri numarası, "2026-03" veya "01.03.2026".
    TarihDonemi = (v) =>
        if v = null then null
        else if Value.Is(v, type datetime) or Value.Is(v, type date) then Date.Year(v) * 100 + Date.Month(v)
        else if Value.Is(v, type number) then
            (if v > 20000 and v < 100000 then let d = Date.From(v) in Date.Year(d) * 100 + Date.Month(d) else null)
        else if Value.Is(v, type text) then
            (try (
                let t = Text.Trim(v) in
                if Text.Length(t) >= 6 and Text.At(t, 4) = "-" then
                    (let y = Number.From(Text.Start(t, 4)), m = Number.From(Text.BeforeDelimiter(Text.Middle(t, 5), "-"))
                     in if m >= 1 and m <= 12 then y * 100 + m else null)
                else if Text.Contains(t, ".") then (let d = Date.FromText(t, "tr-TR") in Date.Year(d) * 100 + Date.Month(d))
                else null
            ) otherwise null)
        else null,
    Matris = (veri as table) as list => List.Buffer(List.Transform(Table.ToRows(Table.Buffer(veri)), List.Buffer)),
    Hucre = (adres as text) as record =>
        let
            a = Text.Upper(Text.Remove(adres, "$")),
            harf = Text.Select(a, {"A" .. "Z"}),
            sutun = List.Accumulate(Text.ToList(harf), 0, (s, c) => s * 26 + Character.ToNumber(c) - 64)
        in
            [Satir = Number.From(Text.Select(a, {"0" .. "9"})), Sutun = sutun],
    // Bir dönem dizisinden yıl × 12 ay listesi: {[Yil, Ay = 12 değer]}.
    Yillik = (donemler as list, degerler as list) as list =>
        let
            yillar = List.Sort(List.Distinct(List.Transform(donemler, PYil)))
        in
            List.Transform(yillar, (y) => [Yil = y, Ay = List.Transform({1 .. 12}, (m) =>
                let i = List.PositionOf(donemler, y * 100 + m) in if i < 0 then null else degerler{i})]),
    Aylar = (g as list, b as list) as record => Record.FromList(g, GAdlari) & Record.FromList(b, BAdlari),
    BosAy = List.Repeat({null}, 12),

    // ---------------------------------------------------------------- Ayarlar.xlsx
    Ayarlar = Excel.Workbook(File.Contents(AyarDosyasi), null, true),
    AyarKlasoru = Text.BeforeDelimiter(AyarDosyasi, "\", {0, RelativePosition.FromEnd}),
    Yol = (ad as text) => if Text.Contains(ad, ":") or Text.StartsWith(ad, "\\") then ad else AyarKlasoru & "\" & ad,
    Sayfa = (ad as text) as table => Ayarlar{[Item = ad, Kind = "Sheet"]}[Data],
    // Başlık satırını ilk sütundaki metne göre bulur, üstündeki açıklama satırlarını atlar.
    BaslikIle = (t as table, ilkBaslik as text) as table =>
        let
            atla = Table.Skip(t, each Kucuk([Column1]) <> Kucuk(ilkBaslik)),
            tablo = Table.PromoteHeaders(atla, [PromoteAllScalars = true])
        in
            Table.SelectRows(tablo, each not Bos(List.First(Record.FieldValues(_)))),
    Konum = Table.Buffer(BaslikIle(Sayfa("Dosya ve Konum Ayarları"), "Alan")),
    Ayar = (alan as text, kod as text, varsayilan) =>
        let
            r = Table.SelectRows(Konum, each [Alan] = alan and [#"Ayar kodu"] = kod)
        in
            if Table.RowCount(r) = 0 or Bos(r{0}[#"Değer"]) then varsayilan else r{0}[#"Değer"],

    TarihHucresi = Hucre(Ayar("Maliyet", "tarih_baslangici", "BF59")),
    ButceHucresi = Hucre(Ayar("Maliyet", "butce_etiketi", "EZ58")),
    IlkGider = Number.From(Ayar("Maliyet", "ilk_gider_satiri", null)),
    GiderSutunu = Number.From(Ayar("Maliyet", "gider_sutunu", TarihHucresi[Sutun] - 1)),
    AdetSatiri = Ayar("Maliyet", "islem_adedi_satiri", null),
    DipToplam = Ayar("Maliyet", "dip_toplam_satiri", null),
    // Revize bütçe: Bütçe başlığıyla aynı satırda bu etiket, altında kapanış yılının 12 ayı (yoksa revize boş kalır).
    RevizeEtiketi = Kucuk(Ayar("Maliyet", "revize_butce_etiketi", "Revize Bütçe")),
    KalemSayisi = Number.From(Ayar("Maliyet", "gider_kalem_sayisi",
        if DipToplam <> null and Number.From(DipToplam) > IlkGider then Number.From(DipToplam) - IlkGider else 27)),

    // Bölüm Eşleştirme: ayardaki sırayla bölüm kayıtları (Python read_settings varsayılanlarıyla).
    BolumTablosu = BaslikIle(Sayfa("Bölüm Eşleştirme"), "Maliyet sayfası / bölüm"),
    BolumSatirlari = List.Buffer(Table.ToRecords(BolumTablosu)),
    MetinAlani = (r as record, alan as text) => let v = Record.FieldOrDefault(r, alan, null) in if Bos(v) then null else Text.Trim(Text.From(v)),
    TumBolumKayit = List.Buffer(List.Transform(List.Positions(BolumSatirlari), (i) =>
        let
            r = BolumSatirlari{i},
            ad = MetinAlani(r, "Maliyet sayfası / bölüm"),
            dir = MetinAlani(r, "Kontrol direktörlüğü"),
            hedef = if MetinAlani(r, "Kontrol bölümü") = null then ad else MetinAlani(r, "Kontrol bölümü")
        in
            [Sira = i + 1, Bolum = ad, Anahtar = Kucuk(ad), Mudurluk = MetinAlani(r, "Müdürlük"), Direktorluk = dir,
             Endeks = MetinAlani(r, "Endeks dosyasındaki bölüm"),
             Lokasyon = if MetinAlani(r, "Lokasyon") = null then "Tanımsız" else MetinAlani(r, "Lokasyon"),
             Hedef = hedef, Grup = Kucuk(dir) & "|" & Kucuk(hedef)])),
    TumAdlari = List.Buffer(List.Transform(TumBolumKayit, each [Bolum])),
    TumAnahtarlari = List.Buffer(List.Transform(TumBolumKayit, each [Anahtar])),
    Direktorlukler = List.Buffer(List.Distinct(List.Transform(TumBolumKayit, each Kucuk([Direktorluk])))),
    // Müdürlük dosyası: yalnız o müdürlüğün bölümleri okunur ve tabloya yazılır (diğer bölümlerin verisi dosyaya girmez).
    MudurlukVar = MudurlukDosyasi <> "",
    SecilenBolumler = if MudurlukVar then List.Select(TumBolumKayit, each Kucuk([Mudurluk]) = Kucuk(MudurlukDosyasi)) else TumBolumKayit,
    BolumKayit = List.Buffer(
        if List.IsEmpty(SecilenBolumler) then error Error.Record("Ayar", "Müdürlük bulunamadı: " & MudurlukDosyasi)
        else List.Transform(List.Positions(SecilenBolumler), (i) => SecilenBolumler{i} & [Sira = i + 1])),
    BolumAdlari = List.Buffer(List.Transform(BolumKayit, each [Bolum])),
    BolumBilgi = Record.FromList(BolumKayit, BolumAdlari),
    // Kontrol grubu yalnız tüm üyeleri kapsamdaysa alınır; müdürlük dışına taşan birleşik grup direktörlük raporundadır.
    KontrolGruplari = List.Buffer(List.Select(List.Distinct(List.Transform(BolumKayit, each [Grup])), (g) =>
        List.AllTrue(List.Transform(List.Select(TumBolumKayit, each [Grup] = g), each List.Contains(BolumAdlari, [Bolum]))))),
    KapsamDisiGruplar = List.Sort(List.Distinct(List.Transform(
        List.Select(BolumKayit, each not List.Contains(KontrolGruplari, [Grup])), each [Hedef])), Comparer.Ordinal),

    Gruplar = List.Buffer(Table.ToRecords(BaslikIle(Sayfa("Gider Grupları"), "İlk indeks"))),
    GiderGrubu = (kod as number) as text =>
        let
            g = List.Select(Gruplar, each Number.From([#"İlk indeks"]) <= kod and kod <= Number.From([#"Son indeks"]))
        in
            if List.IsEmpty(g) then "Eşleşmeyen" else g{0}[#"Gider grubu"],
    GiderEslestirme =
        let
            t = try Table.Buffer(BaslikIle(Sayfa("Gider Eşleştirme"), "Kaynak sıra indeksi")) otherwise null
        in
            if t = null then {} else Table.ToRecords(t),
    // Kaynak sıra indeksi → {rapor indeksi, rapor gider adı}; satır başına tablo taraması yerine kayıt araması.
    EslesmeKaydi = Record.FromList(
        List.Transform(GiderEslestirme, each {Number.From([#"Rapor indeksi"]), Text.Trim(Text.From([#"Rapor gider adı"]))}),
        List.Transform(GiderEslestirme, each Text.From(Number.From([#"Kaynak sıra indeksi"])))),
    OrtakSayfalar =
        let
            t = try BaslikIle(Sayfa("Ortak Sayfalar"), "Sayfa adı") otherwise null
        in
            List.Buffer(if t = null then {} else List.Transform(Table.Column(t, Table.ColumnNames(t){0}), Kucuk)),

    // ---------------------------------------------------------------- Dosya dönemi ("Ekim 2026 - Dosya 1.xlsx")
    DosyaDonemi = (dosya as text) =>
        let
            govde = Text.BeforeDelimiter(dosya, ".", {0, RelativePosition.FromEnd}),
            parca = List.Select(Text.SplitAny(govde, " -_.,()[]"), each _ <> ""),
            yilMi = (t) => Text.Length(t) = 4 and Text.Select(t, {"0" .. "9"}) = t,
            adaylar = List.Transform({0 .. List.Count(parca) - 2}, (i) =>
                let
                    a = parca{i},
                    b = parca{i + 1},
                    y = if yilMi(a) then Number.From(a) else if yilMi(b) then Number.From(b) else null,
                    m = if yilMi(a) then AyNo(b) else if yilMi(b) then AyNo(a) else null
                in
                    if y <> null and m <> null then y * 100 + m else null),
            bulunan = List.Distinct(List.RemoveNulls(adaylar))
        in
            if List.Count(bulunan) = 1 then bulunan{0} else null,

    // ---------------------------------------------------------------- Bir bölüm sayfası → uzun satırlar
    // Dönüş: [Satirlar = {{Kod, Ad, Tür, Dönem, Tutar}}, Uyarilar = {metin}]; Tür: G, B, AG (adet), AB (bütçe adedi).
    OkuBolum = (veri as table, donem as number, sayfaAdi as text) as record =>
        let
            yil = PYil(donem),
            satirlar = Matris(veri),
            butceMi = (v) => List.Contains({"bütçe", "butce"}, Kucuk(v)),
            isaretSatiri = List.First(List.Select(List.Positions(satirlar), (i) => List.AnyTrue(List.Transform(satirlar{i}, butceMi))), null),
            satir = if isaretSatiri = null then error Error.Record("Kaynak", "Bütçe başlığı bulunamadı") else isaretSatiri,
            isaretler = List.Select(List.Positions(satirlar{satir}), (c) => butceMi(satirlar{satir}{c})),
            // Ayardaki Bütçe etiketi hücresi = sayfadaki ilk Bütçe başlığı; diğer konumlar buna göre kaydırılır.
            dSatir = satir - (ButceHucresi[Satir] - 1),
            dSutun = isaretler{0} - (ButceHucresi[Sutun] - 1),
            hucre = (r, c) =>
                let
                    rr = r - 1 + dSatir,
                    cc = c - 1 + dSutun
                in
                    if rr < 0 or rr >= List.Count(satirlar) or cc < 0 or cc >= List.Count(satirlar{rr}) then null else satirlar{rr}{cc},
            gercek = List.Buffer(List.FirstN(
                List.Transform({TarihHucresi[Sutun] .. ButceHucresi[Sutun] - 1},
                    (c) => [Sutun = c, Donem = TarihDonemi(hucre(TarihHucresi[Satir], c))]),
                each [Donem] <> null)),
            gercekDonem = List.Transform(gercek, each [Donem]),
            kapanisVar = if List.Contains(gercekDonem, donem) then true
                else error Error.Record("Kaynak", "Gerçekleşen başlıklarında kapanış ayı (" & Etiket(donem) & ") bulunamadı"),
            blok = (mc) => List.Transform({0 .. 11}, (i) =>
                let
                    v = hucre(ButceHucresi[Satir] + 1, mc + i),
                    p = TarihDonemi(v),
                    ay = if p = null and v <> null and Value.Is(v, type text) then AyNo(v) else null
                in
                    [Sutun = mc + i, Donem = if p <> null then p else if ay <> null then yil * 100 + ay else null]),
            uygun = List.Select(List.Transform(isaretler, (m) => blok(m - dSutun + 1)),
                (b) => List.Sort(List.Transform(b, each [Donem])) = List.Transform({1 .. 12}, each yil * 100 + _)),
            butce = if List.Count(uygun) = 1 then uygun{0} else error Error.Record("Kaynak", "Kapanış yılının tek 12 aylık bütçe bloğu bulunamadı"),
            butceDonem = List.Transform(butce, each [Donem]),
            revizeIsaret = List.Select(List.Positions(satirlar{satir}), (c) => Kucuk(satirlar{satir}{c}) = RevizeEtiketi),
            revizeUygun = List.Select(List.Transform(revizeIsaret, (m) => blok(m - dSutun + 1)),
                (b) => List.Sort(List.Transform(b, each [Donem])) = List.Transform({1 .. 12}, each yil * 100 + _)),
            revize = if List.IsEmpty(revizeUygun) then {} else revizeUygun{0},
            revizeDonem = List.Transform(revize, each [Donem]),
            kalemler = List.Transform({0 .. KalemSayisi - 1}, (k) =>
                let v = hucre(IlkGider + k, GiderSutunu)
                in [Kod = k + 1, Satir = IlkGider + k, Gider = if v <> null and Value.Is(v, type text) then Text.Trim(v) else ""]),
            bos = List.Select(kalemler, each [Gider] = ""),
            kontrol = if List.IsEmpty(bos) then kalemler
                else error Error.Record("Kaynak", "Gider adı bulunamadı: " & sayfaAdi & "!" & Harf(GiderSutunu) & Text.From(bos{0}[Satir])),
            // Tutar hücresi: sayı → kendisi; metin sayı → çevrilir (uyarı); çevrilemeyen metin → boş (uyarı).
            deger = (r, c) =>
                let v = hucre(r, c)
                in
                    if v = null then [V = null, D = 0]
                    else if Value.Is(v, type number) then [V = v, D = 0]
                    else if Value.Is(v, type text) and Text.Trim(v) <> "" then
                        (let n = MetinSayi(v) in [V = n, D = if n = null then 2 else 1, A = Harf(c) & Text.From(r)])
                    else [V = null, D = 0],
            okunan = List.Buffer(List.Transform(kontrol, (k) => [
                Kod = k[Kod], Ad = k[Gider],
                G = List.Buffer(List.Transform(gercek, each deger(k[Satir], [Sutun]))),
                B = List.Buffer(List.Transform(butce, each deger(k[Satir], [Sutun]))),
                R = List.Buffer(List.Transform(revize, each deger(k[Satir], [Sutun])))])),
            adetG = if AdetSatiri = null then {} else List.Buffer(List.Transform(gercek, each deger(Number.From(AdetSatiri), [Sutun]))),
            adetB = if AdetSatiri = null then {} else List.Transform(butce, each Sayi(hucre(Number.From(AdetSatiri), [Sutun]))),
            // Gider Eşleştirme: aynı rapor indeksine bağlanan kaynak giderler aylık katı toplamla birleşir.
            haritali = List.Transform(okunan, (k) =>
                let e = Record.FieldOrDefault(EslesmeKaydi, Text.From(k[Kod]), null)
                in k & [Hedef = if e = null then k[Kod] else e{0}, HedefAd = if e = null then k[Ad] else e{1}]),
            hedefler = List.Distinct(List.Transform(haritali, each [Hedef])),
            birlesik = List.Buffer(List.Transform(hedefler, (h) =>
                let
                    grup = List.Select(haritali, each [Hedef] = h),
                    adlar = List.Distinct(List.Transform(grup, each [HedefAd])),
                    topla = (alan) => List.Transform(List.Positions(Record.Field(grup{0}, alan)), (i) =>
                        KatiToplam(List.Transform(grup, each Record.Field(_, alan){i}[V])))
                in
                    if List.Count(adlar) <> 1 then error Error.Record("Kaynak", Text.From(h) & " rapor indeksinde gider adları farklı. Eşleştirmeyi tamamlayın.")
                    else [Kod = h, Ad = adlar{0}, G = topla("G"), B = topla("B"), R = topla("R")])),
            // Hataları burada tetikle: try bu fonksiyonun çağrısını yakalar, iç alanları değil.
            hazir = List.Count(butce) + List.Count(kontrol) + (if kapanisVar then 1 else 0)
                + List.Sum(List.Transform(birlesik, each List.Count([G]))),
            satirListe = List.Combine(List.Transform(birlesik, (k) =>
                List.Transform(List.Positions(gercekDonem), (i) => {k[Kod], k[Ad], "G", gercekDonem{i}, k[G]{i}})
                & List.Transform(List.Positions(butceDonem), (i) => {k[Kod], k[Ad], "B", butceDonem{i}, k[B]{i}})
                & List.Transform(List.Positions(revizeDonem), (i) => {k[Kod], k[Ad], "R", revizeDonem{i}, k[R]{i}})))
                & List.Transform(List.Positions(adetG), (i) => {null, null, "AG", gercekDonem{i}, adetG{i}[V]})
                & List.Transform(List.Positions(adetB), (i) => {null, null, "AB", butceDonem{i}, adetB{i}}),
            hucreler = List.Combine(List.Transform(okunan, each [G] & [B])) & adetG,
            ornek = (l) => Text.Combine(List.Transform(List.FirstN(l, 5), each [A]), ", ") & (if List.Count(l) > 5 then " …" else ""),
            cevrilen = List.Select(hucreler, each [D] = 1),
            okunamayan = List.Select(hucreler, each [D] = 2),
            uyarilar = (if List.IsEmpty(cevrilen) then {} else {"Metin olarak girilmiş " & Text.From(List.Count(cevrilen))
                    & " tutar sayıya çevrildi (" & sayfaAdi & ": " & ornek(cevrilen) & "). Kaynakta sayı biçimine çevirin."})
                & (if List.IsEmpty(okunamayan) then {} else {"Sayıya çevrilemeyen " & Text.From(List.Count(okunamayan))
                    & " hücre boş sayıldı (" & sayfaAdi & ": " & ornek(okunamayan) & ")."})
                & (if not List.IsEmpty(revizeIsaret) and List.IsEmpty(revizeUygun)
                    then {"Revize bütçe başlığı var ama altında kapanış yılının 12 ayı bulunamadı; revize bütçe boş bırakıldı."} else {})
        in
            if hazir <> null then [Satirlar = List.Buffer(satirListe), Uyarilar = uyarilar] else null,

    // ---------------------------------------------------------------- Kapanış dosyaları
    KlasorAyari = Text.From(Ayar("Dosyalar", "maliyet_klasoru", "maliyet")),
    Klasor = Yol(KlasorAyari),
    KlasorSonu = Kucuk(if Text.EndsWith(Klasor, "\") then Klasor else Klasor & "\"),
    Dosyalar = Table.Buffer(Table.Sort(Table.SelectRows(Folder.Files(Klasor), each
        Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$") and Kucuk([Folder Path]) = KlasorSonu), {"Name"})),
    DonemliDosyalar = Table.Buffer(Table.AddColumn(Dosyalar, "Donem", each DosyaDonemi([Name]))),
    Kapanis =
        if KapanisDonemi = "" then List.Max(List.RemoveNulls(DonemliDosyalar[Donem]))
        else Number.From(Text.BeforeDelimiter(KapanisDonemi, "-")) * 100 + Number.From(Text.AfterDelimiter(KapanisDonemi, "-")),
    KapYil = PYil(Kapanis),
    KapAy = PAy(Kapanis),
    OncekiAy = (p) => if PAy(p) = 1 then p - 89 else p - 1,
    // Tam: kapanış yılının tüm kapanışları (Ocak'ta önceki Aralık da). Hızlı: yalnız kapanış ve önceki ay.
    Kapsamda = (p) => p <> null and p <= Kapanis and (if HizliMod then p >= OncekiAy(Kapanis)
        else (PYil(p) = KapYil or (KapAy = 1 and p = (KapYil - 1) * 100 + 12))),
    DonemsizUyari = List.Transform(Table.ToRecords(Table.SelectRows(DonemliDosyalar, each [Donem] = null)), (f) =>
        [Ad = null, #"Bölüm" = null, Kaynak = f[Name], Metin = "Dosya adında yıl ve ay yan yana bulunamadı"]),
    KapsamDosyalari = Table.SelectRows(DonemliDosyalar, each Kapsamda([Donem])),

    // Her dosyanın sayfaları dosya adı ve sayfa sırasıyla; ortak sayfalar atlanır.
    SayfaListesi = List.Buffer(List.Combine(List.Transform(Table.ToRecords(KapsamDosyalari), (f) =>
        let
            kitap = Table.SelectRows(Excel.Workbook(f[Content], null, true), each [Kind] = "Sheet")
        in
            List.Select(
                List.Transform(Table.ToRecords(kitap), (s) => [Dosya = f[Name], Donem = f[Donem], Sayfa = s[Name], Data = s[Data],
                    Sira = List.PositionOf(TumAnahtarlari, Kucuk(s[Name]))]),
                each not List.Contains(OrtakSayfalar, Kucuk([Sayfa])))))),
    Eslesen = List.Buffer(List.Select(SayfaListesi, each [Sira] >= 0)),
    EslesenAnahtar = List.Buffer(List.Transform(Eslesen, each Text.From([Donem]) & "|" & Text.From([Sira]))),
    SayfaSonuclari = List.Buffer(List.Transform(List.Positions(SayfaListesi), (i) =>
        let
            s = SayfaListesi{i},
            etiket = Etiket(s[Donem]),
            ad = if s[Sira] >= 0 then TumAdlari{s[Sira]} else s[Sayfa],
            kapsamda = s[Sira] < 0 or List.Contains(BolumAdlari, ad),
            k = Text.From(s[Donem]) & "|" & Text.From(s[Sira]),
            adet = List.Count(List.Select(EslesenAnahtar, each _ = k)),
            sonuc = if s[Sira] < 0 or adet > 1 or not kapsamda then null else (try OkuBolum(s[Data], s[Donem], s[Sayfa])),
            uyari = (m) => [Ad = etiket, #"Bölüm" = ad, Kaynak = s[Dosya], Metin = m]
        in
            if not kapsamda then [Gecerli = false, Uyarilar = {}]
            else if s[Sira] < 0 then [Gecerli = false, Uyarilar = {uyari("Bölüm eşleştirmesi bulunamadı")}]
            else if adet > 1 then [Gecerli = false, Uyarilar = {uyari("Aynı bölüm ve kapanış için birden fazla kaynak")}]
            else if sonuc[HasError] then [Gecerli = false, Uyarilar = {uyari(Text.From(sonuc[Error][Message]))}]
            else [Gecerli = true, Donem = s[Donem], Bolum = ad, Kaynak = s[Dosya], Satirlar = sonuc[Value][Satirlar],
                  Uyarilar = List.Transform(sonuc[Value][Uyarilar], uyari)])),
    Gecerli = List.Buffer(List.Select(SayfaSonuclari, each [Gecerli])),
    // Okunan kapanış ("yyyymm|bölüm") → kaynak dosya adı.
    Anlik = Record.FromList(List.Transform(Gecerli, each [Kaynak]), List.Transform(Gecerli, each Text.From([Donem]) & "|" & [Bolum])),
    AnlikDosya = (p, bolum) => Record.FieldOrDefault(Anlik, Text.From(p) & "|" & bolum, null),
    Snap = Table.Buffer(Table.FromRows(List.Combine(List.Transform(Gecerli, (g) =>
            List.Transform(g[Satirlar], each {g[Donem], g[Bolum], g[Kaynak]} & _))),
        type table [Donem = number, Bolum = text, Kaynak = text, Kod = nullable number, Ad = nullable text,
            Tur = text, P = number, Tutar = nullable number])),
    KapanisSnap = Table.Buffer(Table.SelectRows(Snap, each [Donem] = Kapanis)),

    // ---------------------------------------------------------------- Aylık bloklar (Gider, Toplam, Adet)
    AySutunlari = GAdlari & BAdlari & RAdlari,
    Aylik = (t as table, alanlar as list) as table =>
        let
            t1 = Table.AddColumn(t, "Sutun", each (if [Tur] = "B" or [Tur] = "AB" then "B" else if [Tur] = "R" then "R" else "G")
                & Text.From(PAy([P]))),
            t2 = Table.AddColumn(t1, "Yil", each PYil([P])),
            t3 = Table.SelectColumns(t2, alanlar & {"Yil", "Sutun", "Tutar"})
        in
            Table.Pivot(t3, AySutunlari, "Sutun", "Tutar", each List.First(_)),
    BilgiAlanlari = (bolum as text) =>
        let b = Record.Field(BolumBilgi, bolum)
        in [#"Direktörlük" = b[Direktorluk], #"Müdürlük" = b[Mudurluk], #"Bölüm" = bolum],
    AyKaydi = (r as record) => Record.SelectFields(r, AySutunlari),

    GiderTutar = Table.SelectRows(KapanisSnap, each [Tur] = "G" or [Tur] = "B" or [Tur] = "R"),
    GiderBlok = List.Transform(Table.ToRecords(Aylik(GiderTutar, {"Bolum", "Kod", "Ad", "Kaynak"})), (r) =>
        [Anahtar = "Gider|" & r[Bolum] & "|" & Text.From(r[Kod]) & "|" & Text.From(r[Yil]), Tablo = "Gider"]
        & BilgiAlanlari(r[Bolum])
        & [Ad = r[Ad], Kod = r[Kod], Grup = GiderGrubu(r[Kod]), #"Yıl" = r[Yil], Kaynak = r[Kaynak]] & AyKaydi(r)),
    ToplamTutar = Table.Group(GiderTutar, {"Bolum", "Tur", "P"}, {{"Tutar", each KatiToplam([Tutar]), type nullable number}}),
    ToplamBlok = List.Transform(Table.ToRecords(Aylik(ToplamTutar, {"Bolum"})), (r) =>
        [Anahtar = "Toplam|" & r[Bolum] & "|" & Text.From(r[Yil]), Tablo = "Toplam"]
        & BilgiAlanlari(r[Bolum]) & [#"Yıl" = r[Yil]] & AyKaydi(r)),
    AdetTutar = Table.SelectRows(KapanisSnap, each [Tur] = "AG" or [Tur] = "AB"),
    AdetBlok = List.Transform(Table.ToRecords(Aylik(AdetTutar, {"Bolum"})), (r) =>
        [Anahtar = "Adet|" & r[Bolum] & "|" & Text.From(r[Yil]), Tablo = "Adet"]
        & BilgiAlanlari(r[Bolum]) & [#"Yıl" = r[Yil]] & AyKaydi(r)),

    // ---------------------------------------------------------------- Yapı blokları
    KapanisBlok = {[Anahtar = "Kapanis", Tablo = "Kapanis", #"Yıl" = KapYil, #"Sayı" = KapAy]},
    AyarBlok = List.Transform({
            {"tolerans", "Kontrol", 0.01}, {"onemlilik_tutari", "Rapor", 1000000}, {"onemlilik_orani", "Rapor", 0.1},
            {"ozet_kalem_sayisi", "Rapor", 5}, {"arsiv_saklama_adedi", "Rapor", 20}},
        (a) => [Anahtar = "Ayar|" & a{0}, Tablo = "Ayar", Ad = a{0}, #"Sayı" = Number.From(Ayar(a{1}, a{0}, a{2}))]),
    BolumBlok = List.Transform(BolumKayit, (b) =>
        [Anahtar = "Bolum|" & Text.From(b[Sira]), Tablo = "Bolum", #"Bölüm" = b[Bolum], #"Müdürlük" = b[Mudurluk],
         #"Direktörlük" = b[Direktorluk], Ad = if b[Endeks] = null then null else Kucuk(b[Endeks]),
         Kod = List.PositionOf(KontrolGruplari, b[Grup]) + 1, Grup = b[Hedef], Metin = b[Lokasyon],
         Kaynak = AnlikDosya(Kapanis, b[Bolum]), #"Sayı" = b[Sira]]),
    MudurlukCiftleri = List.Distinct(List.Transform(BolumKayit, each {[Mudurluk], [Direktorluk]})),
    MudurlukBlok = List.Transform(List.Positions(MudurlukCiftleri), (i) =>
        [Anahtar = "Mudurluk|" & Text.From(i + 1), Tablo = "Mudurluk", #"Müdürlük" = MudurlukCiftleri{i}{0},
         #"Direktörlük" = MudurlukCiftleri{i}{1}, #"Sayı" = i + 1]),
    // Hiyerarşi: bölümler, müdürlük toplamı, direktörlük toplamı; birden çok direktörlükte genel toplam.
    DuzenListesi = List.Combine(List.Transform(List.Distinct(List.Transform(BolumKayit, each [Direktorluk])), (d) =>
        let
            dBolum = List.Select(BolumKayit, each [Direktorluk] = d),
            mudurlukler = List.Distinct(List.Transform(dBolum, each [Mudurluk]))
        in
            List.Combine(List.Transform(mudurlukler, (m) =>
                List.Transform(List.Select(dBolum, each [Mudurluk] = m), (b) =>
                    [Ad = b[Bolum], Grup = "Bölüm", #"Bölüm" = b[Bolum], #"Müdürlük" = m, #"Direktörlük" = d])
                & {[Ad = m, Grup = "Müdürlük toplamı", #"Bölüm" = null, #"Müdürlük" = m, #"Direktörlük" = d]}))
            & {[Ad = d, Grup = "Direktörlük toplamı", #"Bölüm" = null, #"Müdürlük" = null, #"Direktörlük" = d]}))
        & (if List.Count(List.Distinct(List.Transform(BolumKayit, each [Direktorluk]))) > 1
            then {[Ad = "Tüm direktörlükler", Grup = "Genel toplam", #"Bölüm" = null, #"Müdürlük" = null, #"Direktörlük" = null]} else {}),
    DuzenBlok = List.Transform(List.Positions(DuzenListesi), (i) =>
        [Anahtar = "Duzen|" & Text.From(i + 1), Tablo = "Duzen"] & DuzenListesi{i} & [#"Sayı" = i + 1]),
    KGrupBlok = List.Transform(List.Positions(KontrolGruplari), (i) =>
        let
            uyeler = List.Select(BolumKayit, each [Grup] = KontrolGruplari{i})
        in
            [Anahtar = "KGrup|" & Text.From(i + 1), Tablo = "KGrup", Ad = uyeler{0}[Hedef], Metin = Kucuk(uyeler{0}[Hedef]),
             Grup = Kucuk(uyeler{0}[Direktorluk]), #"Direktörlük" = uyeler{0}[Direktorluk],
             #"Müdürlük" = Text.Combine(List.Sort(List.Distinct(List.Transform(uyeler, each [Mudurluk])), Comparer.Ordinal), ", "),
             #"Bölüm" = Text.Combine(List.Transform(uyeler, each [Bolum]), ", "), #"Sayı" = i + 1]),

    // Kapanış kalemleri: ayardaki bölüm sırası, sonra gider kodu.
    KalemTablosu = Table.Buffer(Table.Sort(
        Table.AddColumn(Table.Distinct(Table.SelectColumns(GiderTutar, {"Bolum", "Kod", "Ad"})), "Sira",
            each List.PositionOf(BolumAdlari, [Bolum])),
        {"Sira", "Kod"})),
    KalemKayit = List.Buffer(Table.ToRecords(KalemTablosu)),
    KalemBlok = List.Transform(List.Positions(KalemKayit), (i) =>
        let r = KalemKayit{i}
        in [Anahtar = "Kalem|" & Text.From(i + 1), Tablo = "Kalem", #"Bölüm" = r[Bolum],
            #"Müdürlük" = Record.Field(BolumBilgi, r[Bolum])[Mudurluk], Kod = r[Kod], Ad = r[Ad],
            Grup = GiderGrubu(r[Kod]), #"Sayı" = i + 1]),
    GrupAdlari = List.Distinct(List.Transform(KalemKayit, each GiderGrubu([Kod]))),
    GrupAdBlok = List.Transform(List.Positions(GrupAdlari), (i) =>
        [Anahtar = "GrupAd|" & Text.From(i + 1), Tablo = "GrupAd", Ad = GrupAdlari{i}, #"Sayı" = i + 1]),
    SatirListesi = List.Combine(List.Transform(GrupAdlari, (g) =>
        {[Metin = "grup", Grup = g, Kod = null, Ad = "Grup toplamı"]}
        & List.Transform(
            Table.ToRecords(Table.Sort(Table.Distinct(Table.SelectColumns(
                Table.SelectRows(KalemTablosu, each GiderGrubu([Kod]) = g), {"Kod", "Ad"})), {{"Kod", Order.Ascending}, {"Ad", Order.Ascending}})),
            (k) => [Metin = "kalem", Grup = g, Kod = k[Kod], Ad = k[Ad]])))
        & {[Metin = "toplam", Grup = "TOPLAM", Kod = null, Ad = "Tüm giderler"]},
    SatirBlok = List.Transform(List.Positions(SatirListesi), (i) =>
        [Anahtar = "Satir|" & Text.From(i + 1), Tablo = "Satir"] & SatirListesi{i} & [#"Sayı" = i + 1]),

    // ---------------------------------------------------------------- Kontrol dosyası (ciro ve kontrol satırları)
    KontrolDosyasi = Text.From(Ayar("Dosyalar", "kontrol_dosyasi", "Kontrol.xlsx")),
    KontrolOku = () as record =>
        let
            kitap = Excel.Workbook(File.Contents(Yol(KontrolDosyasi)), null, true),
            s = Matris(kitap{[Item = Text.From(Ayar("Kontrol", "sayfa", "Kontrol")), Kind = "Sheet"]}[Data]),
            nR = List.Count(s),
            nC = List.Count(s{0}),
            h = (r, c) => if r < 0 or r >= nR or c < 0 or c >= nC then null else s{r}{c},
            // Çapa: satır satır, kendisi ve sağındaki hücre ardışık iki ay olan ilk hücre (ciro tarih başlığı).
            adaylar = List.Transform(List.Positions(s), (r) =>
                let cs = List.Select({0 .. nC - 2}, (c) => let p = TarihDonemi(h(r, c)) in p <> null and TarihDonemi(h(r, c + 1)) = Sonraki(p))
                in if List.IsEmpty(cs) then null else [R = r, C = cs{0}]),
            capa0 = List.First(List.RemoveNulls(adaylar), null),
            capa = if capa0 = null then error Error.Record("Kontrol", "Ciro tarihleri bulunamadı") else capa0,
            ciroHucre = Hucre(Text.From(Ayar("Kontrol", "ciro_baslangici", "D18"))),
            kayma = ciroHucre[Sutun] - 1 - capa[C],
            etiketSutunu = Number.From(Ayar("Kontrol", "etiket_sutunu", 3)) - 1 - kayma,
            butceSol = Hucre(Text.From(Ayar("Kontrol", "butce_ciro_baslangici", "A1")))[Sutun] - 1 - kayma,
            gercekSutun = List.Buffer(List.FirstN(List.Transform({capa[C] .. nC - 1}, (c) => [C = c, V = h(capa[R], c)]), each not Bos([V]))),
            gercek = List.Buffer(List.Transform(gercekSutun, (x) =>
                let p = TarihDonemi(x[V]) in if p = null then error Error.Record("Kontrol", "Geçersiz ciro tarihi: " & Text.From(x[V])) else [C = x[C], P = p])),
            sonGercek = List.Last(gercek)[C],
            planAday = List.RemoveNulls(List.Transform({List.Max({sonGercek + 1, butceSol}) .. nC - 1}, (c) =>
                let
                    v = h(capa[R], c),
                    p0 = TarihDonemi(v),
                    p = if p0 <> null then p0 else if v <> null and Value.Is(v, type text) and AyNo(v) <> null then KapYil * 100 + AyNo(v) else null
                in
                    if p <> null and PYil(p) = KapYil then [C = c, P = p] else null)),
            plan = if List.Sort(List.Transform(planAday, each [P])) = List.Transform({1 .. 12}, each KapYil * 100 + _) then planAday
                else error Error.Record("Kontrol", "Ciro için gerçekleşenden sonraki tek 12 aylık bütçe bloğu bulunamadı"),
            dizi = (r, kolonlar) => List.Transform(kolonlar, each Sayi(h(r, [C]))),
            gDonem = List.Transform(gercek, each [P]),
            ciroG = Yillik(gDonem, dizi(capa[R] + 1, gercek)),
            ciroB = Yillik(List.Transform(plan, each [P]), dizi(capa[R] + 1, plan)),
            ciroYillar = List.Sort(List.Distinct(List.Transform(ciroG & ciroB, each [Yil]))),
            ciro = List.Transform(ciroYillar, (y) =>
                let
                    g = List.Select(ciroG, each [Yil] = y),
                    b = List.Select(ciroB, each [Yil] = y)
                in
                    [Anahtar = "Ciro|" & Text.From(y), Tablo = "Ciro", #"Yıl" = y]
                    & Aylar(if List.IsEmpty(g) then BosAy else g{0}[Ay], if List.IsEmpty(b) then BosAy else b{0}[Ay])),
            baslik = (r) => Kucuk(h(r, etiketSutunu)),
            // Direktörlük satırı: başlığı ayardaki bir direktörlük, iki alt satırı "Opex Toplamı".
            dirSatirlari = List.Select({capa[R] + 2 .. nR - 1}, (r) => List.Contains(Direktorlukler, baslik(r)) and baslik(r + 2) = "opex toplamı"),
            eslesmeler = List.Combine(List.Transform(dirSatirlari, (r) =>
                {[D = baslik(r), T = "opex toplamı", R = r + 2]}
                & List.Transform(
                    List.Generate(() => r + 3, (k) => k < nR and baslik(k) <> "" and not List.Contains(Direktorlukler, baslik(k)), (k) => k + 1),
                    (k) => [D = baslik(r), T = baslik(k), R = k]))),
            anahtarlar = List.Transform(eslesmeler, each [D] & "|" & [T]),
            // Müdürlük dosyasında yalnız kapsamdaki kontrol gruplarının satırları
            alinacak = List.Select(List.Positions(eslesmeler), (i) => not MudurlukVar or List.Contains(KontrolGruplari, anahtarlar{i})),
            kontrol = List.Combine(List.Transform(alinacak, (i) =>
                let
                    e = eslesmeler{i},
                    n = List.Count(List.Select(List.FirstN(anahtarlar, i + 1), each _ = anahtarlar{i}))
                in
                    List.Transform(Yillik(gDonem, dizi(e[R], gercek)), (y) =>
                        [Anahtar = "Kontrol|" & e[D] & "|" & e[T] & "|" & Text.From(n) & "|" & Text.From(y[Yil]), Tablo = "Kontrol",
                         Ad = e[T], Grup = e[D], Kod = n, #"Yıl" = y[Yil]] & Aylar(y[Ay], BosAy)))),
            hazir = List.Count(gercek) + List.Count(plan) + etiketSutunu
        in
            if hazir <> null then [Satirlar = List.Buffer(ciro & kontrol), Uyarilar = {}] else null,
    KontrolSonuc =
        let r = try KontrolOku()
        in if r[HasError] then [Satirlar = {}, Uyarilar = {[Ad = Etiket(Kapanis), #"Bölüm" = null, Kaynak = KontrolDosyasi,
                Metin = "Kontrol dosyası okunamadı: " & Text.From(r[Error][Message])]}]
           else r[Value],

    // ---------------------------------------------------------------- Endeks dosyası ("Bölüm Gider Endeksi" tablosu)
    EndeksDosyasi = Text.From(Ayar("Dosyalar", "endeks_dosyasi", "Endeks.xlsx")),
    EndeksOku = () as record =>
        let
            kitap = Excel.Workbook(File.Contents(Yol(EndeksDosyasi)), null, true),
            s = Matris(kitap{[Item = Text.From(Ayar("Endeks", "sayfa", "Endeks")), Kind = "Sheet"]}[Data]),
            nR = List.Count(s),
            nC = List.Count(s{0}),
            h = (r, c) => if r < 0 or r >= nR or c < 0 or c >= nC then null else s{r}{c},
            basliklar = {"bölüm adı", "bölüm gider endeksi"},
            adaylar = List.Transform(List.Positions(s), (r) =>
                let cs = List.Select({0 .. nC - 1}, (c) => List.Contains(basliklar, Kucuk(h(r, c))))
                in if List.IsEmpty(cs) then null else [R = r, C = cs{0}]),
            capa0 = List.First(List.RemoveNulls(adaylar), null),
            capa = if capa0 = null then error Error.Record("Endeks", "Bölüm Adı veya Bölüm Gider Endeksi başlığı bulunamadı") else capa0,
            ham = List.FirstN(List.Transform({capa[C] + 1 .. nC - 1}, (c) => [C = c, V = h(capa[R], c)]), each not Bos([V])),
            donemler = List.Buffer(List.Transform(ham, (x) =>
                let p = TarihDonemi(x[V]) in if p = null then error Error.Record("Endeks", "Geçersiz endeks tarihi: " & Text.From(x[V])) else [C = x[C], P = p])),
            ardisik = List.AllTrue(List.Transform(List.Skip(List.Positions(donemler)), (i) => donemler{i}[P] = Sonraki(donemler{i - 1}[P]))),
            dogru = if List.IsEmpty(donemler) then error Error.Record("Endeks", "İkinci endeks tablosunda ay başlığı bulunamadı")
                else if ardisik then true else error Error.Record("Endeks", "Endeks ayları ardışık değil"),
            pListe = List.Transform(donemler, each [P]),
            satirlar = List.Select({capa[R] + 1 .. nR - 1}, (r) => Kucuk(h(r, capa[C])) <> ""),
            istenen = List.Transform(List.Select(BolumKayit, each [Endeks] <> null), each Kucuk([Endeks])),
            adlar = List.Transform(satirlar, (r) => Kucuk(h(r, capa[C]))),
            tekil = if List.Count(List.Distinct(adlar)) = List.Count(adlar) then true else error Error.Record("Endeks", "Mükerrer endeks bölümü"),
            endeks = List.Combine(List.Transform(
                List.Select(satirlar, (r) => not MudurlukVar or List.Contains(istenen, Kucuk(h(r, capa[C])))), (r) =>
                List.Transform(Yillik(pListe, List.Transform(donemler, each Sayi(h(r, [C])))), (y) =>
                    [Anahtar = "Endeks|" & Kucuk(h(r, capa[C])) & "|" & Text.From(y[Yil]), Tablo = "Endeks",
                     Ad = Kucuk(h(r, capa[C])), #"Yıl" = y[Yil]] & Aylar(y[Ay], BosAy)))),
            hazir = List.Count(donemler) + (if dogru and tekil then 1 else 0)
        in
            if hazir <> null then [Satirlar = List.Buffer(endeks), Uyarilar = {}] else null,
    EndeksSonuc =
        let r = try EndeksOku()
        in if r[HasError] then [Satirlar = {}, Uyarilar = {[Ad = Etiket(Kapanis), #"Bölüm" = null, Kaynak = EndeksDosyasi,
                Metin = "Endeks dosyası okunamadı: " & Text.From(r[Error][Message])]}]
           else r[Value],

    // ---------------------------------------------------------------- Revizyonlar (ardışık kapanış dosyaları)
    // Ocak: önceki Aralık ↔ Ocak, geçen yılın 12 ayı gerçekleşen. Diğer aylar: (m−1) ↔ m; 1..m−1 gerçekleşen, 12 ay bütçe.
    TumCiftler = (if KapAy = 1 then {[Eski = (KapYil - 1) * 100 + 12, Yeni = Kapanis]} else {})
        & List.Transform({2 .. KapAy}, (m) => [Eski = KapYil * 100 + m - 1, Yeni = KapYil * 100 + m]),
    Ciftler = if HizliMod then List.LastN(TumCiftler, 1) else TumCiftler,
    Karsilastirmalar = List.Combine(List.Transform(List.Positions(Ciftler), (i) =>
        let c = Ciftler{i}
        in
            if PAy(c[Yeni]) = 1 then {c & [Cift = i, Tur = "G", Ad = "Gerçekleşen", Sira = 0, Donemler = List.Transform({1 .. 12}, each (KapYil - 1) * 100 + _)]}
            else {c & [Cift = i, Tur = "G", Ad = "Gerçekleşen", Sira = 0, Donemler = List.Transform({1 .. PAy(c[Yeni]) - 1}, each KapYil * 100 + _)],
                  c & [Cift = i, Tur = "B", Ad = "Bütçe", Sira = 1, Donemler = List.Transform({1 .. 12}, each KapYil * 100 + _)]})),
    Taraf = (p as number, tur as text, donemler as list, onEk as text) as table =>
        Table.RenameColumns(
            Table.SelectColumns(Table.SelectRows(Snap, each [Donem] = p and [Tur] = tur and List.Contains(donemler, [P])), {"Bolum", "Kod", "P", "Ad", "Tutar"}),
            {{"Bolum", onEk & "Bolum"}, {"Kod", onEk & "Kod"}, {"P", onEk & "P"}, {"Ad", onEk & "Ad"}, {"Tutar", onEk & "Tutar"}}),
    RevizyonListesi = List.Combine(List.Transform(Karsilastirmalar, (k) =>
        let
            j = Table.Join(Taraf(k[Eski], k[Tur], k[Donemler], "E"), {"EBolum", "EKod", "EP"},
                Taraf(k[Yeni], k[Tur], k[Donemler], "Y"), {"YBolum", "YKod", "YP"}, JoinKind.FullOuter),
            kayitlar = List.Transform(Table.ToRecords(j), (x) =>
                let
                    bolum = if x[YBolum] <> null then x[YBolum] else x[EBolum],
                    kod = if x[YKod] <> null then x[YKod] else x[EKod],
                    p = if x[YP] <> null then x[YP] else x[EP]
                in
                    [Bolum = bolum, Kod = kod, P = p, Ad = if x[YAd] <> null and x[YAd] <> "" then x[YAd] else x[EAd],
                     Eski = x[ETutar], Yeni = x[YTutar], Tur = k[Ad], Kapanis = k[Yeni],
                     EskiDosya = AnlikDosya(k[Eski], bolum), YeniDosya = AnlikDosya(k[Yeni], bolum),
                     // Sıra: bölüm, kapanış çifti, gider kodu, tür (gerçekleşen önce), etkilenen dönem.
                     Siralama = (((List.PositionOf(BolumAdlari, bolum) * 50 + k[Cift]) * 1000 + kod) * 2 + k[Sira]) * 1000000 + p])
        in
            List.Select(kayitlar, each [EskiDosya] <> null and [YeniDosya] <> null and [Eski] <> [Yeni]))),
    RevizyonSirali = List.Buffer(List.Sort(RevizyonListesi, each [Siralama])),
    RevizyonBlok = List.Transform(List.Positions(RevizyonSirali), (i) =>
        let r = RevizyonSirali{i}
        in [Anahtar = "Revizyon|" & Text.From(i + 1), Tablo = "Revizyon", Metin = Etiket(r[Kapanis]), Grup = r[Tur],
            #"Bölüm" = r[Bolum], #"Müdürlük" = Record.Field(BolumBilgi, r[Bolum])[Mudurluk], Kod = r[Kod], Ad = r[Ad],
            #"Yıl" = PYil(r[P]), #"Sayı" = PAy(r[P]), Kaynak = r[EskiDosya] & " → " & r[YeniDosya]]
            & Aylar({r[Eski], r[Yeni]} & List.Repeat({null}, 10), BosAy)),

    // ---------------------------------------------------------------- Uyarılar
    // "Kapanış dosyası yok" ve seçime bağlı uyarıları şablon kendisi üretir.
    ArdisikEksik = List.Combine(List.Transform(BolumKayit, (b) =>
        if AnlikDosya(Kapanis, b[Bolum]) = null then {}
        else List.Transform(
            List.Select(Ciftler, (c) => AnlikDosya(c[Eski], b[Bolum]) = null or AnlikDosya(c[Yeni], b[Bolum]) = null),
            (c) => [Ad = Etiket(c[Yeni]), #"Bölüm" = b[Bolum], Kaynak = null, Metin = ArdisikUyari]))),
    KapsamDisiUyari = List.Transform(KapsamDisiGruplar, (t) => [Ad = Etiket(Kapanis), #"Bölüm" = null, Kaynak = null,
        Metin = t & ": Birleşik bölüm kontrolü direktörlük raporunda yer alır"]),
    UyariListesi = DonemsizUyari & List.Combine(List.Transform(SayfaSonuclari, each [Uyarilar]))
        & KontrolSonuc[Uyarilar] & EndeksSonuc[Uyarilar] & ArdisikEksik & KapsamDisiUyari,
    UyariBlok = List.Transform(List.Positions(UyariListesi), (i) =>
        [Anahtar = "Uyari|" & Text.From(i + 1), Tablo = "Uyari"] & UyariListesi{i}),

    // ---------------------------------------------------------------- Kimlik ve girdi kaydı
    KimlikBlok = {[Anahtar = "Kimlik", Tablo = "Kimlik", Ad = if MudurlukVar then "Müdürlük" else "Direktörlük",
        Grup = if HizliMod then "Hızlı" else "Tam",
        #"Müdürlük" = if MudurlukVar then BolumKayit{0}[Mudurluk] else null,
        Metin = DateTime.ToText(DateTime.LocalNow(), "yyyy-MM-dd HH:mm")]},
    // Raporun okuduğu dosyalar: ad, tür, boyut, değiştirilme zamanı. Kaynak dosya sonradan değişirse buradan anlaşılır.
    DosyaBilgisi = (yol as text) =>
        let
            klasor = Text.BeforeDelimiter(yol, "\", {0, RelativePosition.FromEnd}),
            ad = Text.AfterDelimiter(yol, "\", {0, RelativePosition.FromEnd}),
            t = try Table.SelectRows(Folder.Contents(klasor), each Kucuk([Name]) = Kucuk(ad)) otherwise null
        in
            if t = null or Table.IsEmpty(t) then null else t{0},
    Girdi = (f, ad as text, tur as text, durum as text, p) =>
        [Ad = ad, Grup = tur, Kod = if f = null then null else f[Attributes][Size],
         Metin = if f = null then null else DateTime.ToText(f[Date modified], "yyyy-MM-dd HH:mm:ss"),
         Kaynak = if f = null then "Bulunamadı" else durum,
         #"Yıl" = if p = null then null else PYil(p), #"Sayı" = if p = null then null else PAy(p)],
    DosyaAdi = (yol as text) => Text.AfterDelimiter(yol, "\", {0, RelativePosition.FromEnd}),
    GirdiListesi =
        {Girdi(DosyaBilgisi(AyarDosyasi), DosyaAdi(AyarDosyasi), "Ayarlar", "Okundu", null),
         Girdi(DosyaBilgisi(Yol(KontrolDosyasi)), DosyaAdi(Yol(KontrolDosyasi)), "Kontrol",
            if List.IsEmpty(KontrolSonuc[Uyarilar]) then "Okundu" else "Okunamadı", null),
         Girdi(DosyaBilgisi(Yol(EndeksDosyasi)), DosyaAdi(Yol(EndeksDosyasi)), "Endeks",
            if List.IsEmpty(EndeksSonuc[Uyarilar]) then "Okundu" else "Okunamadı", null)}
        & List.Transform(Table.ToRecords(DonemliDosyalar), (f) => Girdi(f, f[Name], "Maliyet",
            if f[Donem] = null then "Dönem bulunamadı" else if Kapsamda(f[Donem]) then "Okundu" else "Kapsam dışı", f[Donem])),
    GirdiBlok = List.Transform(List.Positions(GirdiListesi), (i) =>
        [Anahtar = "Girdi|" & Text.From(i + 1), Tablo = "Girdi"] & GirdiListesi{i}),

    // ---------------------------------------------------------------- Tek tablo
    Hepsi = KapanisBlok & KimlikBlok & GirdiBlok & AyarBlok & BolumBlok & MudurlukBlok & DuzenBlok & KGrupBlok & KalemBlok & GrupAdBlok & SatirBlok
        & GiderBlok & ToplamBlok & AdetBlok & KontrolSonuc[Satirlar] & EndeksSonuc[Satirlar] & RevizyonBlok & UyariBlok,
    TekTablo = Table.FromRecords(Hepsi, Sutunlar, MissingField.UseNull),
    Sonuc = Table.TransformColumnTypes(TekTablo,
        {{"Anahtar", type text}, {"Tablo", type text}, {"Direktörlük", type text}, {"Müdürlük", type text}, {"Bölüm", type text},
         {"Ad", type text}, {"Kod", type number}, {"Grup", type text}, {"Yıl", type number}, {"Kaynak", type text},
         {"Metin", type text}, {"Sayı", type number}}
        & List.Transform(AySutunlari, each {_, type number}))
in
    Sonuc
