Attribute VB_Name = "MakroAnaliz"
Option Explicit

' Önemli sapmalar, yıl sonu tahmini, hacim/birim maliyet ayrıştırması ve yönetici özeti.
' Hesap kuralları Python analiz.py ile birebir aynıdır (docs/VERI_VE_HESAP.md):
'   Sapma = gerçekleşen - bütçe. Pozitif sapma Aleyhte (bütçe aşımı), negatif sapma Lehte.
'   Önemli sapma: |sapma| >= önemlilik tutarı VE |sapma / bütçe| >= önemlilik oranı (bütçe 0 ise yalnız tutar).
'   Yıl sonu tahmini: Ocak-kapanış gerçekleşen + kalan ayların bütçesi; ikinci yöntem aylık ortalama gidiş.
'   Hacim etkisi = (gerçekleşen adet - bütçe adet) x bütçe birim maliyeti;
'   birim maliyet etkisi = gerçekleşen gider - bütçe birim maliyeti x gerçekleşen adet.
' Not: VBA kaynak dosyaları Windows-1254 kodlamasıyla da içe aktarıldığından o kodlamada olmayan
' "büyük eşit" (U+2265) ve "eksi" (U+2212) işaretleri ChrW ile üretilir.

Private Function GE() As String
    GE = ChrW(8805)
End Function

Private Function MinusSign() As String
    MinusSign = ChrW(8722)
End Function

Public Function VarianceTitle() As String
    VarianceTitle = "Gerçekleşen " & MinusSign() & " bütçe (TL)"
End Function

Public Function RuleAmount() As Variant
    RuleAmount = CDec(Opt("Rapor", "onemlilik_tutari", 1000000))
End Function

Public Function RuleRate() As Variant
    RuleRate = CDec(Opt("Rapor", "onemlilik_orani", 0.1))
End Function

Public Function RuleTop() As Long
    RuleTop = CLng(Opt("Rapor", "ozet_kalem_sayisi", 5))
End Function

Public Function Direction(ByVal variance As Variant) As Variant
    Direction = Null
    If IsNull(variance) Then Exit Function
    If variance > 0 Then
        Direction = "Aleyhte"
    ElseIf variance < 0 Then
        Direction = "Lehte"
    Else
        Direction = "Bütçede"
    End If
End Function

Private Function Thousands(ByVal amount As Variant) As String
    ' Yerel ayardan bağımsız "1.000.000" biçimi.
    Dim digits As String, result As String, i As Long, negative As Boolean
    negative = (amount < 0)
    digits = CStr(Fix(Abs(CDec(amount)) + CDec(0.5)))
    For i = Len(digits) To 1 Step -1
        result = Mid$(digits, i, 1) & result
        If (Len(digits) - i + 1) Mod 3 = 0 And i > 1 Then result = "." & result
    Next i
    If negative Then result = "-" & result
    Thousands = result
End Function

Public Function ThresholdText() As String
    ' ">= 1.000.000 TL ve >= %10" — Python: analiz.threshold_text
    Dim rate As String
    rate = Trim$(Str$(RuleRate() * 100))
    If InStr(rate, ".") > 0 Then
        Do While Right$(rate, 1) = "0"
            rate = Left$(rate, Len(rate) - 1)
        Loop
        If Right$(rate, 1) = "." Then rate = Left$(rate, Len(rate) - 1)
    End If
    ThresholdText = GE() & " " & Thousands(RuleAmount()) & " TL ve " & GE() & " %" & Replace(rate, ".", ",")
End Function

Public Function IsMaterial(ByVal variance As Variant, ByVal budget As Variant) As Boolean
    Dim rate As Variant
    If IsNull(variance) Or IsNull(budget) Then Exit Function
    If Abs(variance) < RuleAmount() Then Exit Function
    rate = Divide(variance, budget)
    If IsNull(rate) Then
        IsMaterial = True
    Else
        IsMaterial = (Abs(rate) >= RuleRate())
    End If
End Function

Public Function MonthCaption() As String
    Dim names As Variant, m As Variant, s As String
    names = Split("Ocak Şubat Mart Nisan Mayıs Haziran Temmuz Ağustos Eylül Ekim Kasım Aralık", " ")
    For Each m In Picked.Keys
        If s <> "" Then s = s & ", "
        s = s & names(CLng(m) - 1)
    Next m
    MonthCaption = s & " " & ReportYear
End Function

' Önemli sapma satırları, |sapma| büyükten küçüğe. Her satır:
' (0 müdürlük, 1 bölüm, 2 gider kodu, 3 gider kalemi, 4 gider grubu, 5 gerçekleşen, 6 bütçe, 7 sapma, 8 oran, 9 yön, 10 bölüm anahtarı)
Public Function MaterialRows(ByVal scope As Object) As Collection
    Dim result As Collection, items As Object, n As Variant, code As Variant, local As Object, meta As Variant
    Dim a As Variant, b As Variant, v As Variant, row As Variant, i As Long, placed As Boolean
    Set result = New Collection
    If ReportYear <> CloseYear Then Set MaterialRows = result: Exit Function
    Set items = ItemNames()
    For Each n In scope.Keys
        meta = Depts(n): Set local = SingleScope(CStr(n))
        For Each code In items.Keys
            a = ExpenseTotal(local, "A", ReportYear, CStr(code))
            b = ExpenseTotal(local, "B", ReportYear, CStr(code))
            v = Minus(a, b)
            If IsMaterial(v, b) Then
                row = Array(meta(1), meta(0), CLng(code), items(code), GroupOf(CLng(code)), a, b, v, Divide(v, b), Direction(v), CStr(n))
                placed = False
                For i = 1 To result.Count
                    If RankBefore(row, result(i)) Then result.Add row, Before:=i: placed = True: Exit For
                Next i
                If Not placed Then result.Add row
            End If
        Next code
    Next n
    Set MaterialRows = result
End Function

Private Function RankBefore(ByVal x As Variant, ByVal y As Variant) As Boolean
    ' Python sıralaması: (-|sapma|, bölüm adı, gider kodu)
    If Abs(x(7)) <> Abs(y(7)) Then RankBefore = (Abs(x(7)) > Abs(y(7))): Exit Function
    If CStr(x(1)) <> CStr(y(1)) Then RankBefore = (StrComp(CStr(x(1)), CStr(y(1)), vbBinaryCompare) < 0): Exit Function
    RankBefore = (x(2) < y(2))
End Function

' Rapor düzeni: bölümler, müdürlük toplamları, (genel raporda) direktörlük toplamları ve gerekiyorsa genel toplam.
' Her eleman: Array(başlık, satır türü, kapsam sözlüğü). Python: butce.hierarchy
Public Function ReportLayout(ByVal scope As Object, ByVal general As Boolean) As Collection
    Dim result As Collection, directors As Object, director As Variant, managers As Object, manager As Variant
    Dim n As Variant, meta As Variant, local As Object, directorScope As Object, names As Object
    Set result = New Collection: Set directors = Dict(): Set names = Dict()
    For Each n In scope.Keys
        meta = Depts(n)
        If Not directors.Exists(K(meta(2))) Then directors.Add K(meta(2)), Dict(): names(K(meta(2))) = meta(2)
        Set managers = directors(K(meta(2)))
        If Not managers.Exists(K(meta(1))) Then managers.Add K(meta(1)), Dict(): names(K(meta(2)) & Sep & K(meta(1))) = meta(1)
        Set local = managers(K(meta(1))): local(n) = True
    Next n
    For Each director In directors.Keys
        Set managers = directors(director): Set directorScope = Dict()
        For Each manager In managers.Keys
            Set local = managers(manager)
            For Each n In local.Keys
                result.Add Array(Depts(n)(0), "Bölüm", SingleScope(CStr(n)))
                directorScope(n) = True
            Next n
            result.Add Array(names(director & Sep & manager), "Müdürlük toplamı", local)
        Next manager
        If general Then result.Add Array(names(director), "Direktörlük toplamı", directorScope)
    Next director
    If general And directors.Count > 1 Then result.Add Array("Tüm direktörlükler", "Genel toplam", scope)
    Set ReportLayout = result
End Function

Private Function SnapSum(ByVal scope As Object, ByVal prefix As String, ByVal y As Long, ByVal fromMonth As Long, ByVal toMonth As Long) As Variant
    Dim n As Variant, d As Object, m As Long, result As Variant
    result = CDec(0)
    If scope.Count = 0 Then SnapSum = Null: Exit Function
    For Each n In scope.Keys
        Set d = CurrentSnap(CStr(n))
        If d Is Nothing Then SnapSum = Null: Exit Function
        For m = fromMonth To toMonth
            result = Plus(result, ValAt(d, prefix & P(y, m)))
        Next m
    Next n
    SnapSum = result
End Function

' Yıl sonu satırı: (0 başlık, 1 tür, 2 YTD gerçekleşen, 3 YTD bütçe, 4 fark, 5 yıllık bütçe, 6 tahmin, 7 tahmin-yıllık, 8 yön, 9 ortalama gidiş, 10 gidiş-yıllık)
Public Function ForecastRow(ByVal title As String, ByVal level As String, ByVal scope As Object) As Variant
    Dim ytd As Variant, ytdBudget As Variant, remaining As Variant, annual As Variant, plan As Variant, pace As Variant
    ytd = SnapSum(scope, "AT|", CloseYear, 1, CloseMonth)
    ytdBudget = SnapSum(scope, "BT|", CloseYear, 1, CloseMonth)
    If CloseMonth = 12 Then
        remaining = CDec(0)
    Else
        remaining = SnapSum(scope, "BT|", CloseYear, CloseMonth + 1, 12)
    End If
    annual = Plus(ytdBudget, remaining): plan = Plus(ytd, remaining)
    pace = Null
    If Not IsNull(ytd) Then pace = CDec(ytd) * 12 / CloseMonth
    ForecastRow = Array(title, level, ytd, ytdBudget, Minus(ytd, ytdBudget), annual, plan, Minus(plan, annual), Direction(Minus(plan, annual)), pace, Minus(pace, annual))
End Function

' Hacim satırı: (0 başlık, 1 tür, 2 gerçekleşen, 3 bütçe, 4 fark, 5 adet, 6 bütçe adet, 7 BM, 8 bütçe BM, 9 hacim etkisi, 10 birim maliyet etkisi)
Public Function VolumeRow(ByVal title As String, ByVal level As String, ByVal scope As Object) As Variant
    Dim n As Variant, d As Object, a As Variant, b As Variant, q As Variant, qb As Variant, ucb As Variant
    Dim volume As Variant, price As Variant, sumA As Variant, sumB As Variant, sumV As Variant, sumP As Variant
    sumA = CDec(0): sumB = CDec(0): sumV = CDec(0): sumP = CDec(0)
    For Each n In scope.Keys
        Set d = CurrentSnap(CStr(n))
        If d Is Nothing Then
            sumA = Null: sumB = Null: sumV = Null: sumP = Null: a = Null: b = Null: q = Null: qb = Null: volume = Null: price = Null
        Else
            a = SumMonths(d, "AT|", ReportYear): b = SumMonths(d, "BT|", ReportYear)
            q = SumMonths(d, "AC|", ReportYear, "", True): qb = SumMonths(d, "BC|", ReportYear, "", True)
            ucb = Divide(b, qb): volume = Null: price = Null
            If Not IsNull(a) And Not IsNull(ucb) And Not IsNull(q) Then
                volume = Times(Minus(q, qb), ucb): price = Minus(a, Times(ucb, q))
            End If
            sumA = Plus(sumA, a): sumB = Plus(sumB, b): sumV = Plus(sumV, volume): sumP = Plus(sumP, price)
        End If
    Next n
    If level = "Bölüm" Then
        VolumeRow = Array(title, level, a, b, Minus(a, b), q, qb, Divide(a, q), Divide(b, qb), volume, price)
    Else
        VolumeRow = Array(title, level, sumA, sumB, Minus(sumA, sumB), Null, Null, Null, Null, sumV, sumP)
    End If
End Function

Private Function PlainSheet(ByVal wb As Workbook, ByVal name As String, ByVal headers As Variant) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    ws.Cells.Font.Name = "Calibri": ws.Cells.Font.Size = 11
    RowValues ws, 1, headers
    Header ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1))
    ws.Rows(1).RowHeight = 45
    Set PlainSheet = ws
End Function

Private Sub PaintDirection(ByVal cell As Range)
    If CStr(cell.Value2) = "Aleyhte" Then
        cell.Font.Color = RGB(190, 48, 58): cell.Font.Bold = True
    ElseIf CStr(cell.Value2) = "Lehte" Then
        cell.Font.Color = RGB(27, 117, 76): cell.Font.Bold = True
    End If
End Sub

Private Sub FinishPlain(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).WrapText = True
    ws.Columns(1).ColumnWidth = 30
    ws.Range(ws.Cells(1, 2), ws.Cells(1, lastCol)).EntireColumn.ColumnWidth = 18
    If lastRow > 1 Then ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False: ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.PrintTitleRows = "$1:$1"
    ws.Activate: ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
    ActiveWindow.SplitColumn = 0: ActiveWindow.SplitRow = 1: ActiveWindow.FreezePanes = True
End Sub

Private Sub StyleTotals(ByVal ws As Worksheet, ByVal r As Long, ByVal level As String, ByVal lastCol As Long)
    If level = "Bölüm" Then Exit Sub
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
        If level = "Müdürlük toplamı" Then .Interior.Color = RGB(239, 243, 245) Else .Interior.Color = RGB(216, 228, 232)
        .Font.Bold = True: .Font.Color = RGB(38, 55, 70)
    End With
End Sub

Public Sub WriteMaterial(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim ws As Worksheet, rows As Collection, row As Variant, r As Long, exp As Worksheet, codeRows As Object, cols As Object, noteCols As Object
    Dim address As String
    Set ws = PlainSheet(wb, "Önemli Sapmalar", Array("Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Gider grubu", "Gerçekleşen (TL)", "Bütçe (TL)", VarianceTitle(), "Sapma (%)", "Yön", "Açıklama durumu"))
    If ReportYear <> CloseYear Then
        ws.Range("A2").Value2 = "Seçili yıl bütçe yılı değil; önemlilik kontrolü uygulanmadı."
        FinishPlain ws, 2, 11: Exit Sub
    End If
    Set rows = MaterialRows(scope)
    If rows.Count = 0 Then
        ws.Range("A2").Value2 = "Önemlilik eşiğini aşan sapma yok."
        FinishPlain ws, 2, 11: Exit Sub
    End If
    If Not general Then Set exp = ExpensePositions(wb, codeRows, cols, noteCols)
    r = 2
    For Each row In rows
        RowValues ws, r, Array(row(0), row(1), row(2), row(3), row(4), row(5), row(6), row(7), row(8), row(9))
        PaintDirection ws.Cells(r, 10)
        address = ""
        If Not general Then
            If codeRows.Exists(CStr(row(2))) And noteCols.Exists(CStr(row(10))) Then
                address = exp.Cells(codeRows(CStr(row(2))), noteCols(CStr(row(10)))).Address(False, False)
            End If
        End If
        If address <> "" Then
            ws.Cells(r, 11).Formula = "=IF(LEN(TRIM(Giderler!" & address & "))>0,""Açıklandı"",""Açıklama gerekli"")"
            exp.Range(address).Interior.Color = RGB(248, 203, 173)
        Else
            ws.Cells(r, 11).Value2 = "Müdürlük dosyalarında açıklanır"
        End If
        r = r + 1
    Next row
    ws.Range("F2:H" & r - 1).NumberFormat = MoneyFmt
    ws.Range("I2:I" & r - 1).NumberFormat = PctFmt
    ws.Cells(r + 1, 1).Value2 = "Önemlilik eşiği: " & ThresholdText() & " (tutar ve oran birlikte). Açıklama gerekli hücreler müdürlük dosyalarının Giderler sayfasında turuncudur."
    ws.Cells(r + 1, 1).Font.Italic = True: ws.Cells(r + 1, 1).Font.Color = RGB(91, 107, 117)
    FinishPlain ws, r - 1, 11
    ws.Columns("D").ColumnWidth = 30: ws.Columns("K").ColumnWidth = 26
End Sub

Public Sub WriteForecast(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim ws As Worksheet, item As Variant, row As Variant, r As Long, span As String, names As Variant
    If ReportYear <> CloseYear Then Exit Sub
    names = Split("Ocak Şubat Mart Nisan Mayıs Haziran Temmuz Ağustos Eylül Ekim Kasım Aralık", " ")
    span = "Ocak" & ChrW(8211) & names(CloseMonth - 1) & " " & CloseYear
    Set ws = PlainSheet(wb, "Yıl Sonu Tahmini", Array("Bölüm / müdürlük / direktörlük", "Satır türü", "Gerçekleşen" & vbLf & span & " (TL)", "Bütçe" & vbLf & span & " (TL)", "Gerçekleşen " & MinusSign() & " bütçe" & vbLf & span & " (TL)", "Yıllık bütçe" & vbLf & CloseYear & " (TL)", "Yıl sonu tahmini" & vbLf & "gerçekleşen + kalan bütçe (TL)", "Tahmin " & MinusSign() & " yıllık bütçe (TL)", "Yön", "Yıl sonu tahmini" & vbLf & "aylık ortalama gidiş (TL)", "Ortalama gidiş " & MinusSign() & " yıllık bütçe (TL)"))
    r = 2
    For Each item In ReportLayout(scope, general)
        row = ForecastRow(CStr(item(0)), CStr(item(1)), item(2))
        RowValues ws, r, row
        PaintDirection ws.Cells(r, 9)
        StyleTotals ws, r, CStr(item(1)), 11
        r = r + 1
    Next item
    ws.Range("C2:H" & r - 1).NumberFormat = MoneyFmt
    ws.Range("J2:K" & r - 1).NumberFormat = MoneyFmt
    FinishPlain ws, r - 1, 11
End Sub

Public Sub WriteVolume(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim ws As Worksheet, item As Variant, r As Long
    If ReportYear <> CloseYear Then Exit Sub
    Set ws = PlainSheet(wb, "Hacim ve Birim Maliyet", Array("Bölüm / müdürlük / direktörlük", "Satır türü", "Gerçekleşen gider" & vbLf & MonthCaption() & " (TL)", "Bütçe gider" & vbLf & MonthCaption() & " (TL)", VarianceTitle(), "Gerçekleşen işlem adedi", "Bütçe işlem adedi", "Gerçekleşen birim maliyet (TL/işlem)", "Bütçe birim maliyet (TL/işlem)", "Hacim etkisi (TL)", "Birim maliyet etkisi (TL)"))
    r = 2
    For Each item In ReportLayout(scope, general)
        RowValues ws, r, VolumeRow(CStr(item(0)), CStr(item(1)), item(2))
        StyleTotals ws, r, CStr(item(1)), 11
        r = r + 1
    Next item
    ws.Range("C2:E" & r - 1).NumberFormat = MoneyFmt
    ws.Range("F2:G" & r - 1).NumberFormat = "#,##0"
    ws.Range("H2:K" & r - 1).NumberFormat = MoneyFmt
    ws.Cells(r + 1, 1).Value2 = "Hacim etkisi = (gerçekleşen adet " & MinusSign() & " bütçe adet) × bütçe birim maliyeti. Birim maliyet etkisi = gerçekleşen gider " & MinusSign() & " bütçe birim maliyeti × gerçekleşen adet. Toplamları gerçekleşen " & MinusSign() & " bütçe farkını verir."
    ws.Cells(r + 1, 1).Font.Italic = True: ws.Cells(r + 1, 1).Font.Color = RGB(91, 107, 117)
    FinishPlain ws, r - 1, 11
End Sub

Public Sub WriteExecutive(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim ws As Worksheet, rows As Collection, row As Variant, r As Long, budgetYear As Boolean, actual As Variant, planned As Variant
    Dim layout As Collection, top As Variant, verified As Long, total As Long, check As Variant, members As Object, n As Variant, include As Boolean
    Dim warningCount As Long, w As Variant, adverse As Long, labels As Variant, values As Variant, formats As Variant, i As Long
    Dim section As Variant, shown As Long, wanted As String, last As Variant
    budgetYear = (ReportYear = CloseYear)
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = "Yönetici Özeti"
    ws.Cells.Font.Name = "Calibri": ws.Cells.Font.Size = 11
    ws.Range("A1").Value2 = "Yönetici özeti"
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Color = RGB(35, 60, 87)
    ws.Range("A2").Value2 = MonthCaption() & " " & ChrW(183) & " Kaynak kapanış " & P(CloseYear, CloseMonth)
    ws.Range("A2").Font.Color = RGB(91, 107, 117)
    actual = ExpenseTotal(scope, "A", ReportYear): planned = Null
    If budgetYear Then planned = ExpenseTotal(scope, "B", ReportYear)
    Set layout = ReportLayout(scope, general)
    top = Array(Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null)
    If budgetYear And layout.Count > 0 Then
        last = layout.Item(layout.Count)
        top = ForecastRow(CStr(last(0)), CStr(last(1)), last(2))
    End If
    For Each check In Checks
        Set members = check(9): include = True
        For Each n In members.Keys
            If Not scope.Exists(n) Then include = False
        Next n
        If include Then
            total = total + 1
            If check(7) = "Doğrulandı" Then verified = verified + 1
        Else
            ' Python ile aynı: kapsamla kısmen kesişen birleşik kontrol, "direktörlük raporunda yer alır" uyarısı sayılır.
            For Each n In members.Keys
                If scope.Exists(n) Then warningCount = warningCount + 1: Exit For
            Next n
        End If
    Next check
    For Each w In Warnings
        If w(1) = "" Or scope.Exists(K(w(1))) Or Not Depts.Exists(K(w(1))) Then warningCount = warningCount + 1
    Next w
    Set rows = MaterialRows(scope)
    For Each row In rows
        If row(9) = "Aleyhte" Then adverse = adverse + 1
    Next row
    RowValues ws, 4, Array("Gösterge", "Değer")
    labels = Array("Gerçekleşen (TL)", "Bütçe (TL)", VarianceTitle(), "Sapma oranı (%)", "Yön", "Yıl sonu tahmini: gerçekleşen + kalan bütçe (TL)", "Yıllık bütçe (TL)", "Tahmin " & MinusSign() & " yıllık bütçe (TL)", "Veri kontrolü", "Uyarı sayısı", "Önemli sapma", "Önemlilik eşiği")
    If budgetYear Then
        values = Array(actual, planned, Minus(actual, planned), Divide(Minus(actual, planned), planned), Direction(Minus(actual, planned)), top(6), top(5), top(7), verified & " / " & total & " kontrol doğrulandı", warningCount, rows.Count & " kalem (" & adverse & " aleyhte)", ThresholdText())
    Else
        values = Array(actual, Null, Null, Null, "Bütçe yılı dışında", Null, Null, Null, verified & " / " & total & " kontrol doğrulandı", warningCount, "Uygulanmadı", ThresholdText())
    End If
    formats = Array(MoneyFmt, MoneyFmt, MoneyFmt, PctFmt, "", MoneyFmt, MoneyFmt, MoneyFmt, "", "0", "", "")
    For i = 0 To UBound(labels)
        RowValues ws, 5 + i, Array(labels(i), values(i))
        If formats(i) <> "" Then ws.Cells(5 + i, 2).NumberFormat = formats(i)
    Next i
    PaintDirection ws.Range("B9")
    Header ws.Range("A4:B4")
    r = 5 + UBound(labels) + 2
    For Each section In Array("Aleyhte", "Lehte")
        If section = "Aleyhte" Then ws.Cells(r, 1).Value2 = "En büyük aleyhte önemli sapmalar" Else ws.Cells(r, 1).Value2 = "En büyük lehte önemli sapmalar"
        ws.Cells(r, 1).Font.Bold = True: ws.Cells(r, 1).Font.Color = RGB(35, 60, 87): r = r + 1
        RowValues ws, r, Array("Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Gerçekleşen (TL)", "Bütçe (TL)", VarianceTitle(), "Sapma (%)")
        Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)): r = r + 1
        shown = 0: wanted = CStr(section)
        For Each row In rows
            If row(9) = wanted And shown < RuleTop() Then
                RowValues ws, r, Array(row(0), row(1), row(2), row(3), row(5), row(6), row(7), row(8))
                ws.Range(ws.Cells(r, 5), ws.Cells(r, 7)).NumberFormat = MoneyFmt: ws.Cells(r, 8).NumberFormat = PctFmt
                r = r + 1: shown = shown + 1
            End If
        Next row
        If shown = 0 Then
            If budgetYear Then ws.Cells(r, 1).Value2 = "Yok" Else ws.Cells(r, 1).Value2 = "Bütçe yılı dışında"
            r = r + 1
        End If
        r = r + 1
    Next section
    ws.Columns("A").ColumnWidth = 46: ws.Columns("B").ColumnWidth = 30: ws.Columns("C").ColumnWidth = 11
    ws.Columns("D").ColumnWidth = 32: ws.Columns("E:F").ColumnWidth = 20: ws.Columns("G").ColumnWidth = 22: ws.Columns("H").ColumnWidth = 12
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r - 1, 8)).Address
    End With
    ws.Activate: ActiveWindow.DisplayGridlines = False
End Sub
