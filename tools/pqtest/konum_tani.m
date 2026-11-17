    let
        f = Table.ToRecords(Table.Sort(DonemliDosyalar, {{"Donem", Order.Descending}, {"Name", Order.Descending}})){0},
        k = Table.SelectRows(Excel.Workbook(f[Content], null, true), each [Kind] = "Sheet"),
        s = Matris(k{2}[Data]),
        rows = List.Select(List.Positions(s), (r) => List.AnyTrue(List.Transform(s{r}, (v) => Kucuk(v) = "bütçe"))),
        r0 = rows{0},
        c0 = List.PositionOf(List.Transform(s{r0}, Kucuk), "bütçe"),
        below = s{r0 + 1}{c0},
        tur = (v) => if v = null then "null" else if Value.Is(v, type datetime) then "datetime" else if Value.Is(v, type date) then "date"
            else if Value.Is(v, type number) then "number" else if Value.Is(v, type text) then "text" else "diger",
        info = [Eksik = Text.From(MaliyetEksik), Dosya = f[Name], Sayfalar = Text.Combine(k[Name], ","),
            Anahtarlar = Text.Combine(TumAnahtarlari, ","), Ortak = Text.Combine(OrtakSayfalar, ","), R = Text.From(r0), C = Text.From(c0),
            Alt = Text.From(below), AltTur = tur(below), AyMi = Text.From(try AyMi(below) otherwise "HATA"),
            Basliktaki = Text.Combine(List.Transform(List.FirstN(List.RemoveNulls(s{r0 + 1}), 4), each tur(_) & ":" & Text.From(_)), " | "),
            Bulunan = try Text.FromBinary(Json.FromValue(MaliyetBulunan)) otherwise "HATA: " & (try MaliyetBulunan)[Error][Message],
            Konum = try Text.FromBinary(Json.FromValue(MaliyetKonumu(s))) otherwise "HATA2: " & (try MaliyetKonumu(s))[Error][Message]]
    in
        Table.FromRecords({info})
