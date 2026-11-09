// Butce_Veri — makrosuz ve Python'suz yol: Excel'in yerleşik Power Query'si (lisans gerekmez).
// Kurulum: docs/POWER_QUERY.md. Bu sorgu Ayarlar.xlsx'i ve maliyet klasörünü okur, son kapanış dosyalarındaki
// bölüm sayfalarını tek tabloya çevirir: bölüm × gider kalemi × ay için Gerçekleşen, Bütçe ve Sapma.
// Konumlar Ayarlar.xlsx › Dosya ve Konum Ayarları'ndan alınır; tablo Bütçe başlığına göre bulunduğundan
// kaynak sayfasının boş üst satırları/sütunları sonucu etkilemez.
let
    // ===================== YALNIZ BU İKİ SATIRI DÜZENLEYİN =====================
    AyarDosyasi = "C:\Butce\ornek\Ayarlar.xlsx",   // Ayarlar.xlsx dosyasının tam yolu
    KapanisDonemi = "",                             // boş = klasördeki en son kapanış; örnek: "2026-10"
    // ============================================================================

    AyAdlari = {"ocak", "şubat", "mart", "nisan", "mayıs", "haziran", "temmuz", "ağustos", "eylül", "ekim", "kasım", "aralık"},
    Anahtar = (v) => if v = null then "" else Text.Lower(Text.Trim(Text.From(v)), "tr-TR"),
    AyNo = (t) =>
        let
            k = Anahtar(t),
            tam = List.PositionOf(AyAdlari, k),
            kisa = List.PositionOf(List.Transform(AyAdlari, each Text.Start(_, 3)), k)
        in
            if tam >= 0 then tam + 1 else if kisa >= 0 then kisa + 1 else null,

    // ---------------------------------------------------------------- Ayarlar.xlsx
    Ayarlar = Excel.Workbook(File.Contents(AyarDosyasi), null, true),
    AyarKlasoru = Text.BeforeDelimiter(AyarDosyasi, "\", {0, RelativePosition.FromEnd}),
    Sayfa = (ad as text) as table => Ayarlar{[Item = ad, Kind = "Sheet"]}[Data],
    // Başlık satırını ilk sütundaki metne göre bulur, üstündeki açıklama satırlarını atlar.
    BaslikIle = (t as table, ilkBaslik as text) as table =>
        let
            atla = Table.Skip(t, each Anahtar([Column1]) <> Anahtar(ilkBaslik)),
            tablo = Table.PromoteHeaders(atla, [PromoteAllScalars = true])
        in
            Table.SelectRows(tablo, each List.First(Record.FieldValues(_)) <> null),
    Konum = BaslikIle(Sayfa("Dosya ve Konum Ayarları"), "Alan"),
    Ayar = (alan as text, kod as text, varsayilan) =>
        let
            r = Table.SelectRows(Konum, each [Alan] = alan and [#"Ayar kodu"] = kod)
        in
            if Table.RowCount(r) = 0 then varsayilan else r{0}[#"Değer"],
    Hucre = (adres as text) as record =>
        let
            a = Text.Upper(Text.Remove(adres, "$")),
            harf = Text.Select(a, {"A" .. "Z"}),
            sutun = List.Accumulate(Text.ToList(harf), 0, (s, c) => s * 26 + Character.ToNumber(c) - 64)
        in
            [Satir = Number.From(Text.Select(a, {"0" .. "9"})), Sutun = sutun],

    TarihHucresi = Hucre(Ayar("Maliyet", "tarih_baslangici", "BF59")),
    ButceHucresi = Hucre(Ayar("Maliyet", "butce_etiketi", "EZ58")),
    IlkGider = Number.From(Ayar("Maliyet", "ilk_gider_satiri", null)),
    GiderSutunu = Number.From(Ayar("Maliyet", "gider_sutunu", null)),
    DipToplam = Ayar("Maliyet", "dip_toplam_satiri", null),
    KalemSayisi = Number.From(Ayar("Maliyet", "gider_kalem_sayisi",
        if DipToplam <> null and Number.From(DipToplam) > IlkGider then Number.From(DipToplam) - IlkGider else 27)),

    Bolumler = Table.AddColumn(
        BaslikIle(Sayfa("Bölüm Eşleştirme"), "Maliyet sayfası / bölüm"),
        "Anahtar", each Anahtar([#"Maliyet sayfası / bölüm"])),
    Gruplar = Table.ToRecords(BaslikIle(Sayfa("Gider Grupları"), "İlk indeks")),
    GiderGrubu = (kod as number) as text =>
        let
            g = List.Select(Gruplar, each Number.From([#"İlk indeks"]) <= kod and kod <= Number.From([#"Son indeks"]))
        in
            if List.IsEmpty(g) then "Eşleşmeyen" else g{0}[#"Gider grubu"],
    GiderEslestirme =
        let
            t = try BaslikIle(Sayfa("Gider Eşleştirme"), "Kaynak sıra indeksi") otherwise null
        in
            if t = null then #table({"Kaynak sıra indeksi", "Rapor indeksi", "Rapor gider adı"}, {}) else t,

    // ---------------------------------------------------------------- Dosya dönemi ("Ekim 2026 - Dosya 1.xlsx")
    Donem = (dosya as text) =>
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

    // Başlık hücresindeki dönem: tarih, Excel seri numarası, "2026-03" veya "01.03.2026".
    TarihDonemi = (v) =>
        if v = null then null
        else if Value.Is(v, type datetime) or Value.Is(v, type date) then Date.Year(v) * 100 + Date.Month(v)
        else if Value.Is(v, type number) then
            (if v > 20000 and v < 100000 then let d = Date.From(v) in Date.Year(d) * 100 + Date.Month(d) else null)
        else if Value.Is(v, type text) then
            (try (
                let t = Text.Trim(v) in
                if Text.Length(t) >= 7 and Text.At(t, 4) = "-" then Number.From(Text.Start(t, 4)) * 100 + Number.From(Text.Middle(t, 5, 2))
                else (let d = Date.FromText(t, "tr-TR") in Date.Year(d) * 100 + Date.Month(d))
            ) otherwise null)
        else null,
    Tutar = (v) =>
        if v = null then null
        else if Value.Is(v, type number) then v
        else if Value.Is(v, type text) then
            (try Number.FromText(Text.Replace(Text.Remove(Text.Trim(v), {" ", "₺"}), "TL", ""), "tr-TR") otherwise null)
        else null,

    // ---------------------------------------------------------------- Bir bölüm sayfası → uzun tablo
    Oku = (veri as table, yil as number) as table =>
        let
            satirlar = Table.ToRows(veri),
            butceMi = (v) => List.Contains({"bütçe", "butce"}, Anahtar(v)),
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
            gercek = List.FirstN(
                List.Transform({TarihHucresi[Sutun] .. ButceHucresi[Sutun] - 1}, (c) => [Sutun = c, Donem = TarihDonemi(hucre(TarihHucresi[Satir], c))]),
                each [Donem] <> null),
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
            kalemler = List.Transform({0 .. KalemSayisi - 1}, (k) =>
                let v = hucre(IlkGider + k, GiderSutunu)
                in [Kod = k + 1, Satir = IlkGider + k, Gider = if v = null then "" else Text.Trim(Text.From(v))]),
            bos = List.Select(kalemler, each [Gider] = ""),
            kontrol = if List.IsEmpty(bos) then kalemler else error Error.Record("Kaynak", "Gider adı bulunamadı: satır " & Text.From(bos{0}[Satir])),
            satirListe = List.Combine(List.Transform(kontrol, (k) =>
                List.Transform(gercek, (g) => {k[Kod], k[Gider], "Gerçekleşen", g[Donem], Tutar(hucre(k[Satir], g[Sutun]))})
                & List.Transform(butce, (b) => {k[Kod], k[Gider], "Bütçe", b[Donem], Tutar(hucre(k[Satir], b[Sutun]))})))
        in
            Table.FromRows(satirListe, {"Kaynak kodu", "Kaynak gider", "Tür", "Dönem", "Tutar"}),

    // ---------------------------------------------------------------- Kapanış dosyaları
    KlasorAyari = Text.From(Ayar("Dosyalar", "maliyet_klasoru", "maliyet")),
    Klasor = if Text.Contains(KlasorAyari, ":") or Text.StartsWith(KlasorAyari, "\\") then KlasorAyari else AyarKlasoru & "\" & KlasorAyari,
    KlasorSonu = Anahtar(if Text.EndsWith(Klasor, "\") then Klasor else Klasor & "\"),
    Dosyalar = Table.SelectRows(Folder.Files(Klasor), each
        Text.Lower([Extension]) = ".xlsx" and not Text.StartsWith([Name], "~$") and Anahtar([Folder Path]) = KlasorSonu),
    DonemliDosyalar = Table.SelectRows(Table.AddColumn(Dosyalar, "Donem", each Donem([Name])), each [Donem] <> null),
    Kapanis = if KapanisDonemi = "" then List.Max(DonemliDosyalar[Donem]) else Number.From(Text.Remove(KapanisDonemi, "-")),
    KapanisDosyalari = Table.SelectRows(DonemliDosyalar, each [Donem] = Kapanis),
    Sayfalar = Table.ExpandTableColumn(
        Table.AddColumn(Table.SelectColumns(KapanisDosyalari, {"Content", "Name", "Donem"}), "Kitap", each Excel.Workbook([Content], null, true)),
        "Kitap", {"Name", "Data", "Kind"}, {"Sayfa adı", "Data", "Kind"}),
    BolumSayfalari = Table.NestedJoin(
        Table.AddColumn(Table.SelectRows(Sayfalar, each [Kind] = "Sheet"), "Anahtar", each Anahtar([#"Sayfa adı"])),
        {"Anahtar"}, Bolumler, {"Anahtar"}, "Bolum", JoinKind.Inner),
    Okunan = Table.AddColumn(BolumSayfalari, "Satirlar", each
        let
            r = try Oku([Data], Number.IntegerDivide([Donem], 100))
        in
            if r[HasError] then error Error.Record("Kaynak okunamadı", [Name] & " / " & [#"Sayfa adı"] & ": " & r[Error][Message]) else r[Value]),
    Uzun = Table.ExpandTableColumn(
        Table.ExpandTableColumn(Table.SelectColumns(Okunan, {"Name", "Bolum", "Satirlar"}), "Bolum",
            {"Maliyet sayfası / bölüm", "Müdürlük", "Kontrol direktörlüğü", "Lokasyon"},
            {"Bölüm", "Müdürlük", "Direktörlük", "Lokasyon"}),
        "Satirlar", {"Kaynak kodu", "Kaynak gider", "Tür", "Dönem", "Tutar"}),

    // Gider Eşleştirme: aynı rapor indeksine bağlanan kaynak giderler toplanır.
    Eslesmeli = Table.AddColumn(
        Table.AddColumn(Uzun, "Gider kodu", each
            let e = Table.SelectRows(GiderEslestirme, (x) => Number.From(x[#"Kaynak sıra indeksi"]) = [Kaynak kodu])
            in if Table.RowCount(e) = 0 then [Kaynak kodu] else Number.From(e{0}[#"Rapor indeksi"]), Int64.Type),
        "Gider kalemi", each
            let e = Table.SelectRows(GiderEslestirme, (x) => Number.From(x[#"Kaynak sıra indeksi"]) = [Kaynak kodu])
            in if Table.RowCount(e) = 0 then [Kaynak gider] else Text.From(e{0}[#"Rapor gider adı"]), type text),
    Toplanan = Table.Group(Eslesmeli, {"Direktörlük", "Müdürlük", "Bölüm", "Lokasyon", "Gider kodu", "Gider kalemi", "Tür", "Dönem", "Name"},
        {{"Tutar", each List.Sum([Tutar]), type nullable number}}),
    Genis = Table.Pivot(Toplanan, {"Gerçekleşen", "Bütçe"}, "Tür", "Tutar", List.Sum),
    Sonuc0 = Table.AddColumn(Genis, "Gider grubu", each GiderGrubu([Gider kodu]), type text),
    Sonuc1 = Table.AddColumn(Sonuc0, "Tarih", each #date(Number.IntegerDivide([Dönem], 100), Number.Mod([Dönem], 100), 1), type date),
    Sonuc2 = Table.AddColumn(Sonuc1, "Yıl", each Number.IntegerDivide([Dönem], 100), Int64.Type),
    Sonuc3 = Table.AddColumn(Sonuc2, "Ay", each Text.Proper(AyAdlari{Number.Mod([Dönem], 100) - 1}, "tr-TR"), type text),
    Sonuc4 = Table.AddColumn(Sonuc3, "Ay no", each Number.Mod([Dönem], 100), Int64.Type),
    Sonuc5 = Table.AddColumn(Sonuc4, "Sapma (gerçekleşen − bütçe)", each
        if [Gerçekleşen] = null or [Bütçe] = null then null else [Gerçekleşen] - [Bütçe], type nullable number),
    Sonuc6 = Table.AddColumn(Sonuc5, "Yön", each
        let s = [#"Sapma (gerçekleşen − bütçe)"] in if s = null then null else if s > 0 then "Aleyhte" else if s < 0 then "Lehte" else "Bütçede", type nullable text),
    Sonuc = Table.TransformColumnTypes(
        Table.RenameColumns(
            Table.SelectColumns(Sonuc6, {"Direktörlük", "Müdürlük", "Bölüm", "Lokasyon", "Gider grubu", "Gider kodu", "Gider kalemi",
                "Tarih", "Yıl", "Ay no", "Ay", "Gerçekleşen", "Bütçe", "Sapma (gerçekleşen − bütçe)", "Yön", "Name"}, MissingField.UseNull),
            {{"Name", "Kaynak dosya"}}),
        {{"Gerçekleşen", type nullable number}, {"Bütçe", type nullable number}})
in
    Sonuc
