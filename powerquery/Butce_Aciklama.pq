// Butce_Aciklama — müdürlüklerden dönen açıklamaları toplar; onay durumunu ve işlem geçmişini korur.
// Kurulum: docs/POWER_QUERY.md › Açıklama toplama. Ana raporun "Açıklama Onayı" sayfasına (A1) tablo olarak yüklenir.
//
// Akış: müdürlük dosyası (MudurlukDosyasi ile hazırlanmış rapor) › müdürlük Açıklamalar sayfasına yazar › dosyayı geri
// gönderir › dosya AciklamaKlasoru'na konur › Veri › Tümünü Yenile.
// Onay: Açıklama Onayı sayfasında Durum sütunu (Bekliyor / Onaylandı / Reddedildi / Bu döneme aktar) ve Onay notu
// elle doldurulur. Sorgu yenilemede kendi önceki tablosunu okur (Excel.CurrentWorkbook): Durum ve notlar korunur,
// her değişiklik İşlem geçmişi sütununa zaman · kişi · işlem olarak eklenir (Kontrol Paneli › Onaylayan adı).
// Açıklama metni dosyada değişirse durum Bekliyor'a döner. Kaynak dosyası kaldırılan açıklama silinmez.
// Sözleşme ve Python benzetimi: tools/pq_aciklama.py.
let
    // ===================== YALNIZ BU İKİ SATIRI DÜZENLEYİN =====================
    AyarDosyasi = "C:\Butce\ornek\Ayarlar.xlsx",   // Butce_Veri sorgusundakiyle aynı yol
    AciklamaKlasoru = "pq_aciklama",                // dönen dosyalar: Ayarlar klasörüne göre alt klasör veya tam yol
    // ============================================================================

    Kucuk = (v) => if v = null then "" else Text.Lower(Text.Trim(Text.From(v)), "tr-TR"),
    Bos = (v) => v = null or (Value.Is(v, type text) and Text.Trim(v) = ""),
    Metin = (v) => if Bos(v) then null else Text.Trim(Text.From(v)),
    Sayi = (v) => if v <> null and Value.Is(v, type number) then v else null,
    TarihMetni = (v) =>
        if v = null then null
        else if Value.Is(v, type datetime) or Value.Is(v, type date) then Date.ToText(Date.From(v), "yyyy-MM-dd")
        else if Value.Is(v, type number) then (try Date.ToText(Date.From(v), "yyyy-MM-dd") otherwise Text.From(v))
        else Metin(v),
    Durumlar = {"Bekliyor", "Onaylandı", "Reddedildi", "Bu döneme aktar"},
    Sutunlar = {"Anahtar", "Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Açıklama", "Yazan", "Yazma tarihi",
        "Kapanış", "Aylar", "Gerçekleşen (dosyada)", "Bütçe (dosyada)", "Kaynak dosya", "Dosya tarihi", "Okuma notu",
        "Kaynakta", "Durum", "Onay notu", "Son işlem", "İşlemi yapan", "İşlem zamanı", "İşlem geçmişi", "Kayıtlı durum"},
    Simdi = DateTime.ToText(DateTime.LocalNow(), "yyyy-MM-dd HH:mm"),

    // ---------------------------------------------------------------- Bu çalışma kitabı: önceki tablo ve adlar
    CalismaKitabi = Excel.CurrentWorkbook(),
    AdDegeri = (ad as text) =>
        let t = Table.SelectRows(CalismaKitabi, each [Name] = ad)
        in if Table.IsEmpty(t) then null else (try Record.FieldValues(Table.First(t{0}[Content])){0} otherwise null),
    Kullanici = let k = AdDegeri("Kullanici") in if Bos(k) then "(Kontrol Paneli'nde ad yok)" else Text.Trim(Text.From(k)),
    RaporDonemi = Metin(AdDegeri("Rapor_Donemi")),
    RaporAylari = Metin(AdDegeri("Rapor_Aylari")),
    // Önceki yüklemenin tablosu, adından bağımsız olarak sütunlarından tanınır.
    OncekiAday = Table.SelectRows(CalismaKitabi, each
        (try List.ContainsAll(Table.ColumnNames([Content]), {"Anahtar", "Kayıtlı durum", "İşlem geçmişi"}) otherwise false)),
    OncekiTablo = if Table.IsEmpty(OncekiAday) then #table({"Anahtar"}, {}) else OncekiAday{0}[Content],
    Onceki = List.Buffer(List.Select(Table.ToRecords(Table.Distinct(OncekiTablo, {"Anahtar"})),
        each not Bos([Anahtar]) and not Text.StartsWith(Text.From([Anahtar]), "Dosya|"))),
    OncekiAnahtar = List.Buffer(List.Transform(Onceki, each Text.From([Anahtar]))),
    OncekiKayit = (k) => let i = List.PositionOf(OncekiAnahtar, k) in if i < 0 then null else Onceki{i},
    Alan = (r, ad) => if r = null then null else Record.FieldOrDefault(r, ad, null),

    // ---------------------------------------------------------------- Ayarlar: bölüm → müdürlük
    AyarKlasoru = Text.BeforeDelimiter(AyarDosyasi, "\", {0, RelativePosition.FromEnd}),
    Ayarlar = Excel.Workbook(File.Contents(AyarDosyasi), null, true),
    BolumSayfasi = Ayarlar{[Item = "Bölüm Eşleştirme", Kind = "Sheet"]}[Data],
    BolumTablosu = Table.PromoteHeaders(Table.Skip(BolumSayfasi, each Kucuk([Column1]) <> "maliyet sayfası / bölüm"),
        [PromoteAllScalars = true]),
    Bolumler = List.Buffer(List.Select(Table.ToRecords(BolumTablosu), each not Bos(Record.Field(_, "Maliyet sayfası / bölüm")))),
    BolumAnahtarlari = List.Buffer(List.Transform(Bolumler, each Kucuk(Record.Field(_, "Maliyet sayfası / bölüm")))),
    BolumBul = (ad) => let i = List.PositionOf(BolumAnahtarlari, Kucuk(ad)) in if i < 0 then null else Bolumler{i},

    // ---------------------------------------------------------------- Dönen dosyalar
    Klasor = if Text.Contains(AciklamaKlasoru, ":") or Text.StartsWith(AciklamaKlasoru, "\\") then AciklamaKlasoru
        else AyarKlasoru & "\" & AciklamaKlasoru,
    KlasorSonu = Kucuk(if Text.EndsWith(Klasor, "\") then Klasor else Klasor & "\"),
    Dosyalar = Table.Buffer(Table.Sort(Table.SelectRows(Folder.Files(Klasor), each Text.Lower([Extension]) = ".xlsx"
        and not Text.StartsWith([Name], "~$") and Kucuk([Folder Path]) = KlasorSonu), {"Name"})),
    // Bir dosyanın Açıklamalar sayfasındaki dolu satırlar (formül sütunlarının kayıtlı değerleriyle).
    OkuDosya = (f as record) as list =>
        let
            kitap = Excel.Workbook(f[Content], null, true),
            sayfa = Table.SelectRows(kitap, each [Kind] = "Sheet" and [Name] = "Açıklamalar"),
            tablo = if Table.IsEmpty(sayfa)
                then error Error.Record("Açıklama", "Açıklamalar sayfası yok; dosya rapor şablonundan hazırlanmamış")
                else Table.PromoteHeaders(sayfa{0}[Data], [PromoteAllScalars = true]),
            eksik = List.Difference({"Bölüm", "Gider kodu", "Açıklama", "Aylar", "Kapanış"}, Table.ColumnNames(tablo)),
            dogru = if List.IsEmpty(eksik) then tablo
                else error Error.Record("Açıklama", "Açıklamalar sayfasında sütun eksik: " & Text.Combine(eksik, ", ")),
            satirlar = List.Select(Table.ToRecords(dogru), each not Bos([Bölüm]) and not Bos([Açıklama])),
            dosyaTarihi = DateTime.ToText(f[Date modified], "yyyy-MM-dd HH:mm"),
            kayitlar = List.Transform(satirlar, (r) =>
                let
                    kod = try Number.From(Alan(r, "Gider kodu")) otherwise null,
                    bolum = BolumBul(r[Bölüm]),
                    bolumAdi = if bolum = null then Metin(r[Bölüm]) else Metin(Record.Field(bolum, "Maliyet sayfası / bölüm")),
                    mudurluk = if bolum = null then null else Metin(Record.FieldOrDefault(bolum, "Müdürlük", null)),
                    dosyaMud = Metin(Alan(r, "Dosya müdürlüğü")),
                    kalem = Metin(Alan(r, "Gider kalemi")),
                    kapanis = Metin(Alan(r, "Kapanış")),
                    aylar = Metin(Alan(r, "Aylar")),
                    notu =
                        if bolum = null then "Bölüm Ayarlar'da yok"
                        else if kod = null then "Gider kodu sayı değil"
                        else if kalem = "Bu bölümde bu kod yok" then "Bu bölümde bu kod yok"
                        else if dosyaMud <> null and Kucuk(dosyaMud) <> Kucuk(mudurluk) then "Bölüm dosyanın müdürlüğünde değil (" & dosyaMud & ")"
                        else if kapanis = null or aylar = null then "Kapanış/ay bilgisi yok: dosya hesaplanmadan kaydedilmiş"
                        else null
                in
                    [Anahtar = bolumAdi & "|" & (if kod = null then Text.From(Alan(r, "Gider kodu")) else Text.From(kod)) & "|"
                        & (if kapanis = null then "?" else kapanis) & "|" & (if aylar = null then "?" else aylar),
                     #"Müdürlük" = mudurluk, #"Bölüm" = bolumAdi, #"Gider kodu" = kod,
                     #"Gider kalemi" = if kalem = "Bu bölümde bu kod yok" then null else kalem,
                     #"Açıklama" = Metin(r[Açıklama]), Yazan = Metin(Alan(r, "Yazan")), #"Yazma tarihi" = TarihMetni(Alan(r, "Tarih")),
                     #"Kapanış" = kapanis, Aylar = aylar,
                     #"Gerçekleşen (dosyada)" = Sayi(Alan(r, "Gerçekleşen (seçili aylar)")),
                     #"Bütçe (dosyada)" = Sayi(Alan(r, "Bütçe (seçili aylar)")),
                     #"Kaynak dosya" = f[Name], #"Dosya tarihi" = dosyaTarihi, #"Okuma notu" = notu])
        in
            List.Buffer(kayitlar),
    DosyaSonuclari = List.Transform(Table.ToRecords(Dosyalar), (f) =>
        let r = try OkuDosya(f)
        in if r[HasError] then {[Anahtar = "Dosya|" & f[Name], #"Kaynak dosya" = f[Name],
                #"Dosya tarihi" = DateTime.ToText(f[Date modified], "yyyy-MM-dd HH:mm"),
                #"Okuma notu" = "Dosya okunamadı: " & Text.From(r[Error][Message]), Kaynakta = "Var"]}
           else r[Value]),
    Toplanan0 = List.Combine(DosyaSonuclari),
    // Aynı anahtar birden çok dosyada: en yeni dosya alınır; metinler farklıysa çakışma notu (onaylanamaz).
    YeniAnahtarlar = List.Distinct(List.Transform(Toplanan0, each [Anahtar])),
    Toplanan = List.Buffer(List.Transform(YeniAnahtarlar, (k) =>
        let
            grup = List.Select(Toplanan0, each [Anahtar] = k),
            secilen = List.Last(List.Sort(grup, each [#"Dosya tarihi"] & "|" & [#"Kaynak dosya"])),
            metinler = List.Distinct(List.Transform(grup, each Record.FieldOrDefault(_, "Açıklama", null))),
            dosyalar = Text.Combine(List.Transform(grup, each [#"Kaynak dosya"]), ", ")
        in
            if List.Count(metinler) > 1 and Bos(Record.FieldOrDefault(secilen, "Okuma notu", null))
            then secilen & [#"Okuma notu" = "Çakışma: farklı açıklamalar (" & dosyalar & "); en yeni dosya alındı, düzeltin"]
            else secilen)),

    // ---------------------------------------------------------------- Önceki tabloyla birleştirme ve işlem geçmişi
    Satir = (zaman, kisi, islem, notu) => zaman & " · " & kisi & " · " & islem & (if Bos(notu) then "" else " · " & Text.From(notu)),
    Ekle = (gecmis, satir) => if Bos(gecmis) then satir else Text.From(gecmis) & "#(lf)" & satir,
    Birlestir = (y, e) =>
        let
            dosyaHatasi = y <> null and Text.StartsWith(y[Anahtar], "Dosya|"),
            temel = if y <> null then y else e,
            kaynakta = if y <> null then "Var" else if Alan(e, "Kaynakta") = "Aktarıldı" then "Aktarıldı" else "Yok",
            eDurum = Metin(Alan(e, "Durum")),
            kayitli = Metin(Alan(e, "Kayıtlı durum")),
            metinDegisti = e <> null and y <> null and Metin(Alan(e, "Açıklama")) <> Metin(Alan(y, "Açıklama")),
            aktar = e <> null and eDurum = "Bu döneme aktar",
            kaldirildi = e <> null and y = null and Alan(e, "Kaynakta") = "Var",
            gecerliDurum = if eDurum <> null and List.Contains(Durumlar, eDurum) then eDurum else "Bekliyor",
            durum =
                if e = null or metinDegisti then "Bekliyor"
                else if aktar then (if kayitli = null then "Bekliyor" else kayitli)
                else gecerliDurum,
            islem =
                if e = null then "Yeni açıklama"
                else if metinDegisti then "Açıklama değişti, yeniden onay gerekir"
                else if aktar then "Bu döneme aktarıldı (" & (if RaporDonemi = null then "?" else RaporDonemi) & ")"
                else if kaldirildi then "Kaynak dosyadan kaldırıldı"
                else if gecerliDurum <> (if kayitli = null then "Bekliyor" else kayitli) then gecerliDurum
                else null,
            yapan =
                if e = null or metinDegisti then (if Bos(Alan(y, "Yazan")) then "Müdürlük" else Alan(y, "Yazan"))
                else if kaldirildi then "Sistem"
                else Kullanici,
            onayNotu = Alan(e, "Onay notu"),
            satir = if islem = null then null
                else Satir(Simdi, yapan, islem, if e <> null and not metinDegisti and not kaldirildi then onayNotu else null)
        in
            if dosyaHatasi then y
            else Record.SelectFields(temel, List.Intersect({Sutunlar, Record.FieldNames(temel)}))
                & [Kaynakta = kaynakta, Durum = durum, #"Onay notu" = onayNotu,
                   #"Son işlem" = if islem = null then Alan(e, "Son işlem") else islem,
                   #"İşlemi yapan" = if islem = null then Alan(e, "İşlemi yapan") else yapan,
                   #"İşlem zamanı" = if islem = null then Alan(e, "İşlem zamanı") else Simdi,
                   #"İşlem geçmişi" = if satir = null then Alan(e, "İşlem geçmişi") else Ekle(Alan(e, "İşlem geçmişi"), satir),
                   #"Kayıtlı durum" = durum],
    TumAnahtarlar = YeniAnahtarlar & List.Select(OncekiAnahtar, each not List.Contains(YeniAnahtarlar, _)),
    Birlesik = List.Transform(TumAnahtarlar, (k) =>
        let i = List.PositionOf(YeniAnahtarlar, k)
        in Birlestir(if i < 0 then null else Toplanan{i}, OncekiKayit(k))),
    // "Bu döneme aktar": önceki kapanışın açıklaması bu kapanış ve ay seçimi için onaylı satır olur.
    Aktarilan = List.RemoveNulls(List.Transform(Onceki, (e) =>
        let
            yeniAnahtar = Text.From(e[Bölüm]) & "|" & Text.From(e[#"Gider kodu"]) & "|" & RaporDonemi & "|" & RaporAylari,
            satir = Satir(Simdi, Kullanici, "Önceki dönemden aktarıldı ve onaylandı (" & Text.From(e[#"Kapanış"]) & ")", null)
        in
            if Metin(Alan(e, "Durum")) <> "Bu döneme aktar" or RaporDonemi = null or RaporAylari = null
                or List.Contains(TumAnahtarlar, yeniAnahtar) then null
            else Record.SelectFields(e, {"Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Açıklama", "Yazan", "Yazma tarihi"},
                    MissingField.UseNull)
                & [Anahtar = yeniAnahtar, #"Kapanış" = RaporDonemi, Aylar = RaporAylari,
                   #"Kaynak dosya" = "Önceki dönemden: " & Text.From(e[#"Kapanış"]), Kaynakta = "Aktarıldı",
                   Durum = "Onaylandı", #"Son işlem" = "Önceki dönemden aktarıldı ve onaylandı", #"İşlemi yapan" = Kullanici,
                   #"İşlem zamanı" = Simdi, #"İşlem geçmişi" = satir, #"Kayıtlı durum" = "Onaylandı"])),
    Tablo = Table.FromRecords(Birlesik & Aktarilan, Sutunlar, MissingField.UseNull),
    Sirali = Table.Sort(Tablo, {{"Kapanış", Order.Descending}, {"Müdürlük", Order.Ascending}, {"Bölüm", Order.Ascending},
        {"Gider kodu", Order.Ascending}, {"Anahtar", Order.Ascending}}),
    Sonuc = Table.TransformColumnTypes(Sirali,
        List.Transform(Sutunlar, each {_, if List.Contains({"Gider kodu", "Gerçekleşen (dosyada)", "Bütçe (dosyada)"}, _)
            then type number else type text}))
in
    Sonuc
