Attribute VB_Name = "MakroYaz"
Option Explicit

' "Put" VBA'nın ayrılmış dosya komutu olduğundan prosedür adı SetCell'dir.
Public Sub SetCell(ByVal cell As Range, ByVal value As Variant)
    If IsNull(value) Or IsEmpty(value) Then
        cell.ClearContents
    Else
        cell.Value2 = value
    End If
End Sub

Public Sub RowValues(ByVal ws As Worksheet, ByVal row As Long, ByVal values As Variant, Optional ByVal firstCol As Long = 1)
    ' Satır tek seferde yazılır (hücre hücre yazmaya göre çok daha hızlı). Boş (Null) değer hücreyi boşaltır.
    Dim i As Long, buffer() As Variant, width As Long
    width = UBound(values) - LBound(values) + 1
    If width < 1 Then Exit Sub
    ReDim buffer(1 To 1, 1 To width)
    For i = LBound(values) To UBound(values)
        If IsNull(values(i)) Then
            buffer(1, i - LBound(values) + 1) = Empty
        ElseIf VarType(values(i)) = vbDecimal Then
            buffer(1, i - LBound(values) + 1) = CDbl(values(i))
        Else
            buffer(1, i - LBound(values) + 1) = values(i)
        End If
    Next i
    ws.Range(ws.Cells(row, firstCol), ws.Cells(row, firstCol + width - 1)).Value2 = buffer
End Sub

Public Function NewSheet(ByVal wb As Workbook, ByVal name As String, ByVal title As String) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = name
    ws.Cells.Font.Name = "Calibri": ws.Cells.Font.Size = 11
    ws.Range("A1:I2").Merge
    ws.Range("A1").Value2 = title
    ws.Range("A1:I2").Interior.Color = RGB(229, 239, 237)
    ws.Range("A1").Font.Size = 16: ws.Range("A1").Font.Bold = True
    ws.Range("A1").VerticalAlignment = xlCenter
    ws.Range("A3").Value2 = SelectionCaption() & " | Kaynak kapanış: " & P(CloseYear, CloseMonth)
    ws.Rows(3).RowHeight = 28
    Set NewSheet = ws
End Function

Public Sub Header(ByVal area As Range)
    With area
        .Interior.Color = RGB(237, 242, 246)
        .Font.Bold = True: .Font.Color = RGB(42, 60, 72)
        .WrapText = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
End Sub

Public Sub Finish(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    With ws
        .Range(.Cells(5, 1), .Cells(lastRow, lastCol)).VerticalAlignment = xlCenter
        .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Columns.AutoFit
        .Rows(5).RowHeight = 58
        .Columns(1).ColumnWidth = 31
        .Range(.Cells(5, 1), .Cells(lastRow, lastCol)).Borders(xlInsideHorizontal).Color = RGB(227, 232, 235)
        .PageSetup.Orientation = xlLandscape
        .PageSetup.Zoom = False: .PageSetup.FitToPagesWide = 1: .PageSetup.FitToPagesTall = False
        .PageSetup.PrintTitleRows = "$1:$5"
        .PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Address
        .Activate: ActiveWindow.DisplayGridlines = False
        ActiveWindow.FreezePanes = False: .Range("D6").Select: ActiveWindow.FreezePanes = True
    End With
End Sub

Public Sub VarianceColors(ByVal area As Range, Optional ByVal positiveGood As Boolean = True)
    ' Hücre değeri kuralı ve "0" eşiği dil bağımsızdır; İngilizce formül metni Türkçe Excel'de hata verebilirdi.
    ' Boş metin ("") hücreleri renklense de görünür bir şey yazmadığından etkisizdir.
    Dim rule As FormatCondition, good As Long, bad As Long
    good = RGB(27, 117, 76): bad = RGB(190, 48, 58)
    area.FormatConditions.Delete
    Set rule = area.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
    If positiveGood Then rule.Font.Color = good Else rule.Font.Color = bad
    Set rule = area.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    If positiveGood Then rule.Font.Color = bad Else rule.Font.Color = good
End Sub

Private Sub SummaryRow(ByVal ws As Worksheet, ByVal row As Long, ByVal label As String, ByVal level As String, ByVal scope As Object)
    Dim actual As Variant, prior As Variant, planned As Variant, rev As Variant, opex As Variant
    Dim eligible As Object, n As Variant, d As Object, prev As Variant
    actual = ExpenseTotal(scope, "A", ReportYear): prior = ExpenseTotal(scope, "A", ReportYear - 1)
    planned = Null
    If ReportYear = CloseYear Then planned = ExpenseTotal(scope, "B", ReportYear)
    Set eligible = Dict()
    For Each n In scope.Keys
        Set d = CurrentSnap(CStr(n))
        If Not d Is Nothing Then
            prev = SumMonths(d, "AT|", ReportYear - 1)
            If Positive(prev) Then eligible(n) = True
        End If
    Next n
    rev = SumMonths(Revenue, "", ReportYear)
    opex = Divide(ExpenseTotal(eligible, "A", ReportYear), rev)
    RowValues ws, row, Array(label, level, prior, actual, planned, Null, Null, opex)
    ws.Cells(row, 6).Formula = "=IF(COUNT(D" & row & ":E" & row & ")=2,E" & row & "-D" & row & "," & Chr$(34) & Chr$(34) & ")"
    ws.Cells(row, 7).Formula = "=IF(AND(COUNT(D" & row & ":E" & row & ")=2,E" & row & "<>0),(D" & row & "-E" & row & ")/E" & row & "," & Chr$(34) & Chr$(34) & ")"
    If level <> "Bölüm" Then
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 8)).Interior.Color = RGB(238, 244, 241)
        ws.Rows(row).Font.Bold = True
    End If
End Sub

Public Sub WriteSummary(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim ws As Worksheet, row As Long, manager As Variant, n As Variant, meta As Variant, local As Object, mgrs As Object
    Set ws = NewSheet(wb, "Genel Özet", "Gider özeti — TL")
    RowValues ws, 5, Array("Birim", "Düzey", SelectionCaption(ReportYear - 1) & " Gerçekleşen TL", SelectionCaption() & " Gerçekleşen TL", SelectionCaption() & " Bütçe TL", "Bütçe - Gerçekleşen TL", "Bütçe aşım oranı %", "OPEX / Şirket Cirosu %")
    Header ws.Range("A5:H5"): row = 6: Set mgrs = Managers(scope)
    For Each manager In mgrs.Keys
        Set local = Dict()
        For Each n In scope.Keys
            meta = Depts(n)
            If K(meta(1)) = K(manager) Then
                local(n) = True
                SummaryRow ws, row, CStr(meta(0)), "Bölüm", SingleScope(CStr(n))
                row = row + 1
            End If
        Next n
        SummaryRow ws, row, CStr(manager) & " toplamı", "Müdürlük", local: row = row + 1
    Next manager
    If general Then SummaryRow ws, row, "Direktörlük genel toplamı", "Direktörlük", scope: row = row + 1
    ws.Range("C6:F" & row - 1).NumberFormat = MoneyFmt
    ws.Range("G6:H" & row - 1).NumberFormat = PctFmt
    VarianceColors ws.Range("F6:F" & row - 1)
    VarianceColors ws.Range("G6:G" & row - 1), False
    Finish ws, row - 1, 8
    ws.Columns("C:F").ColumnWidth = 23: ws.Columns("G:H").ColumnWidth = 22
    ws.Range("A" & row + 2).Value2 = "OPEX/ciro: önceki yıl aynı seçili ayları bulunan bölümler; payda şirket cirosudur."
End Sub

Public Function SingleScope(ByVal n As String) As Object
    Dim d As Object
    Set d = Dict(): d(n) = True: Set SingleScope = d
End Function

Private Sub FourValues(ByVal ws As Worksheet, ByVal row As Long, ByVal col As Long, ByVal a As Variant, ByVal b As Variant)
    SetCell ws.Cells(row, col), a: SetCell ws.Cells(row, col + 1), b
    ws.Cells(row, col + 2).FormulaR1C1 = "=IF(COUNT(RC[-2]:RC[-1])=2,RC[-1]-RC[-2]," & Chr$(34) & Chr$(34) & ")"
    ws.Cells(row, col + 3).FormulaR1C1 = "=IF(AND(COUNT(RC[-3]:RC[-2])=2,RC[-2]<>0),(RC[-3]-RC[-2])/RC[-2]," & Chr$(34) & Chr$(34) & ")"
    ws.Range(ws.Cells(row, col), ws.Cells(row, col + 2)).NumberFormat = MoneyFmt
    ws.Cells(row, col + 3).NumberFormat = PctFmt
End Sub

Public Sub WriteExpenses(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim ws As Worksheet, names As Object, groupsUsed As Object, code As Variant, groupName As Variant
    Dim n As Variant, meta As Variant, col As Long, width As Long, row As Long, first As Long, subtotal As Long
    Dim a As Variant, b As Variant, item As String, last As Long, local As Object, rr As Long
    Set ws = NewSheet(wb, "Giderler", "Ana ve alt giderler — TL")
    Set names = ItemNames(): Set groupsUsed = Dict()
    For Each code In names.Keys
        groupsUsed(GroupOf(CLng(code))) = True
    Next code
    If general Then width = 4 Else width = 5
    RowValues ws, 5, Array("Ana gider", "Gider kodu", "Alt gider / toplam")
    col = 4
    For Each n In scope.Keys
        meta = Depts(n)
        ws.Range(ws.Cells(4, col), ws.Cells(4, col + width - 1)).Merge
        ws.Cells(4, col).Value2 = meta(0)
        Header ws.Range(ws.Cells(4, col), ws.Cells(4, col + width - 1))
        RowValues ws, 5, Array(SelectionCaption() & " Gerçekleşen TL", SelectionCaption() & " Bütçe TL", "Bütçe - Gerçekleşen TL", "Bütçe aşım oranı %"), col
        If Not general Then ws.Cells(5, col + 4).Value2 = "Açıklama"
        col = col + width
    Next n
    last = col - 1: Header ws.Range(ws.Cells(5, 1), ws.Cells(5, last)): row = 6
    ws.Outline.SummaryRow = xlBelow
    For Each groupName In groupsUsed.Keys
        first = row
        For Each code In names.Keys
            If GroupOf(CLng(code)) = groupName Then
                RowValues ws, row, Array(groupName, CLng(code), names(code))
                col = 4
                For Each n In scope.Keys
                    Set local = SingleScope(CStr(n)): a = ExpenseTotal(local, "A", ReportYear, CStr(code)): b = Null
                    If ReportYear = CloseYear Then b = ExpenseTotal(local, "B", ReportYear, CStr(code))
                    FourValues ws, row, col, a, b
                    col = col + width
                Next n
                row = row + 1
            End If
        Next code
        If row > first Then ws.Rows(first & ":" & row - 1).Group
        RowValues ws, row, Array(groupName, Null, CStr(groupName) & " toplamı")
        col = 4
        For Each n In scope.Keys
            Set local = SingleScope(CStr(n)): a = GroupTotal(local, CStr(groupName), "A", ReportYear, names): b = Null
            If ReportYear = CloseYear Then b = GroupTotal(local, CStr(groupName), "B", ReportYear, names)
            FourValues ws, row, col, a, b
            col = col + width
        Next n
        ws.Range(ws.Cells(row, 1), ws.Cells(row, last)).Interior.Color = RGB(237, 243, 239)
        ws.Rows(row).Font.Bold = True: row = row + 1
    Next groupName
    RowValues ws, row, Array("Genel toplam", Null, "Tüm ana giderler toplamı")
    col = 4
    For Each n In scope.Keys
        Set local = SingleScope(CStr(n)): a = ExpenseTotal(local, "A", ReportYear): b = Null
        If ReportYear = CloseYear Then b = ExpenseTotal(local, "B", ReportYear)
        FourValues ws, row, col, a, b
        VarianceColors ws.Range(ws.Cells(6, col + 2), ws.Cells(row, col + 2))
        VarianceColors ws.Range(ws.Cells(6, col + 3), ws.Cells(row, col + 3)), False
        ws.Range(ws.Cells(4, col), ws.Cells(row, col + width - 1)).BorderAround Weight:=xlMedium, Color:=RGB(99, 119, 125)
        col = col + width
    Next n
    ws.Rows(row).Font.Bold = True
    Finish ws, row, last
    ws.Columns("B").ColumnWidth = 10: ws.Columns("C").ColumnWidth = 34
    For col = 4 To last Step width
        ws.Range(ws.Cells(1, col), ws.Cells(1, col + 3)).EntireColumn.ColumnWidth = 23
        If Not general Then ws.Columns(col + 4).ColumnWidth = 38
    Next col
    ws.Outline.ShowLevels RowLevels:=1
End Sub

Private Sub UnitRow(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal level As String, ByVal scope As Object)
    Dim n As Variant, item As Variant, rate As Variant, a As Variant, b As Variant, x As Variant, y As Variant, adjustedCost As Variant
    Dim factorValue As Variant, plannedBM As Variant, managerWeight As Variant, directorWeight As Variant, totalCost As Variant, cnt As Long
    a = CDec(0): b = CDec(0): x = CDec(0): y = CDec(0): adjustedCost = CDec(0)
    rate = CDec(0): managerWeight = Null: directorWeight = Null: factorValue = Null: plannedBM = Null
    For Each n In scope.Keys
        If Units.Exists(n) Then
            item = Units(n): a = Plus(a, item(0)): x = Plus(x, item(1)): b = Plus(b, item(2)): y = Plus(y, item(3))
            adjustedCost = Plus(adjustedCost, Times(item(8), item(3))): cnt = cnt + 1
        End If
    Next n
    If cnt = 0 Then Exit Sub
    For Each n In scope.Keys
        If Units.Exists(n) Then
            item = Units(n)
            rate = Plus(rate, Times(Divide(item(0), a), item(9)))
            If level = "Bölüm" Then managerWeight = item(10): directorWeight = item(11): factorValue = item(7): plannedBM = item(6)
        End If
    Next n
    RowValues ws, r, Array(label, level, Divide(b, y), Divide(a, x), plannedBM, Minus(factorValue, 1), Divide(adjustedCost, y), rate, managerWeight, directorWeight, a, x)
    If ReportYear <> CloseYear Then
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 3)).ClearContents
        ws.Range(ws.Cells(r, 5), ws.Cells(r, 10)).ClearContents
    End If
    If level <> "Bölüm" Then
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Interior.Color = RGB(237, 243, 239)
        ws.Rows(r).Font.Bold = True
    End If
End Sub

Public Sub WriteUnits(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim ws As Worksheet, mgrs As Object, manager As Variant, n As Variant, meta As Variant, local As Object
    Dim r As Long, count As Long
    Set ws = NewSheet(wb, "Birim Maliyet", "Birim maliyet — TL / işlem")
    RowValues ws, 5, Array("Birim", "Düzey", CStr(ReportYear - 1) & " Gerçekleşen BM", CStr(ReportYear) & " Gerçekleşen BM", CStr(CloseYear) & " Bütçe BM", "Dönem enflasyonu %", "Endekslenmiş önceki yıl BM", "Reel iyileşme %", "Müdürlük ağırlığı %", "Direktörlük ağırlığı %", "Seçili dönem gideri TL", "Seçili dönem işlem adedi")
    Header ws.Range("A5:L5"): r = 6: Set mgrs = Managers(scope)
    ws.Range("L5").Value2 = "Seçili dönem " & IndicatorCaption(scope, 5)
    For Each manager In mgrs.Keys
        Set local = Dict()
        For Each n In scope.Keys
            meta = Depts(n)
            If K(meta(1)) = K(manager) And Units.Exists(n) Then
                local(n) = True
                UnitRow ws, r, CStr(meta(0)), "Bölüm", SingleScope(CStr(n)): r = r + 1
            End If
        Next n
        If local.Count > 0 Then UnitRow ws, r, CStr(manager) & " toplamı", "Müdürlük", local: r = r + 1
    Next manager
    If general And Units.Count > 0 Then UnitRow ws, r, "Direktörlük genel toplamı", "Direktörlük", scope: r = r + 1
    If r = 6 Then ws.Cells(r, 1).Value2 = "Seçilen aylar için karşılaştırılabilir eşleştirilmiş bölüm yok.": r = r + 1
    ws.Range("C6:E" & r - 1).NumberFormat = MoneyFmt
    ws.Range("F6:F" & r - 1).NumberFormat = PctFmt
    ws.Range("G6:G" & r - 1).NumberFormat = MoneyFmt
    ws.Range("H6:J" & r - 1).NumberFormat = PctFmt
    ws.Range("K6:K" & r - 1).NumberFormat = MoneyFmt
    ws.Range("L6:L" & r - 1).NumberFormat = "#,##0"
    VarianceColors ws.Range("H6:H" & r - 1)
    Finish ws, r - 1, 12: ws.Columns("C:L").ColumnWidth = 23
End Sub

Public Sub WriteAudit(ByVal wb As Workbook, ByVal scope As Object)
    Dim ws As Worksheet, r As Long, row As Variant, members As Object, n As Variant, include As Boolean
    Dim m As Variant, i As Long, meta As Variant, state As String, source As Object, indicator As Long, planned As Variant
    Set ws = NewSheet(wb, "Kontrol ve Revizyon", "Doğrulama ve geriye dönük değişiklikler")
    RowValues ws, 5, Array("Kontrol bölümü", "Müdürlük", "Direktörlük", "Dönem", "Kaynak toplamı TL", "Kontrol tutarı TL", "Fark TL", "Sonuç", "Birleşen bölümler")
    Header ws.Range("A5:I5"): r = 6
    For Each row In Checks
        Set members = row(9): include = True
        For Each n In members.Keys
            If Not scope.Exists(n) Then include = False
        Next n
        If include Then
            For i = 0 To 8
                SetCell ws.Cells(r, i + 1), row(i)
            Next i
            If row(7) <> "Doğrulandı" Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Font.Color = RGB(190, 48, 58)
            r = r + 1
        End If
    Next row
    If r > 6 Then ws.Range("E6:G" & r - 1).NumberFormat = MoneyFmt
    r = r + 2
    RowValues ws, r, Array("Yakalanan kapanış", "Etkilenen dönem", "Müdürlük", "Bölüm", "Alt gider kodu", "Alt gider kalemi", "Veri türü", "Önce TL", "Sonra TL", "Değişiklik TL", "Önceki kaynak", "Yeni kaynak")
    Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)): r = r + 1
    For Each row In Revisions
        If scope.Exists(CStr(row(12))) Then
            For i = 0 To 11
                SetCell ws.Cells(r, i + 1), row(i)
            Next i
            ws.Range(ws.Cells(r, 8), ws.Cells(r, 10)).NumberFormat = MoneyFmt: r = r + 1
        End If
    Next row
    r = r + 2: RowValues ws, r, Array("Dönem", "Şirket cirosu TL", "Bütçe cirosu TL")
    Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)): r = r + 1
    For Each m In Picked.Keys
        RowValues ws, r, Array(P(ReportYear, CLng(m)), ValAt(Revenue, P(ReportYear, CLng(m))), ValAt(RevenueBudget, P(ReportYear, CLng(m))))
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).NumberFormat = MoneyFmt: r = r + 1
    Next m
    r = r + 2: RowValues ws, r, Array("Kapanış", "Bölüm", "Kaynak", "Uyarı")
    Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)): r = r + 1
    For Each row In Warnings
        If row(1) = "" Or scope.Exists(K(row(1))) Or Not Depts.Exists(K(row(1))) Then RowValues ws, r, row: r = r + 1
    Next row
    r = r + 2
    RowValues ws, r, Array("Müdürlük", "Bölüm", "Gösterge sırası", "Kaynak Excel başlığı", CStr(ReportYear - 1) & " aynı aylar gerçekleşen adet", CStr(ReportYear) & " seçili aylar gerçekleşen adet", CStr(CloseYear) & " seçili aylar bütçe adet")
    Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)): r = r + 1
    For Each n In scope.Keys
        meta = Depts(n): Set source = CurrentSnap(CStr(n))
        If Not source Is Nothing Then
            For indicator = 1 To 5
                planned = Null
                If ReportYear = CloseYear Then planned = IndicatorBudget(source, indicator, ReportYear)
                RowValues ws, r, Array(meta(1), meta(0), indicator, IndicatorName(source, indicator), SumMonths(source, "AP" & indicator & Sep, ReportYear - 1), SumMonths(source, "AP" & indicator & Sep, ReportYear), planned)
                ws.Range(ws.Cells(r, 5), ws.Cells(r, 7)).NumberFormat = "#,##0": r = r + 1
            Next indicator
        End If
    Next n
    Finish ws, r, 12
    ws.Columns("A:L").ColumnWidth = 24: ws.Columns("C:D").ColumnWidth = 35
    ws.Columns("H:I").ColumnWidth = 30: ws.Columns("K:L").ColumnWidth = 45
    ws.UsedRange.WrapText = True: ws.UsedRange.Rows.AutoFit
End Sub

Public Function IndicatorBudget(ByVal source As Object, ByVal index As Long, ByVal year As Long) As Variant
    Dim m As Variant, value As Variant, first As Variant, second As Variant, result As Variant
    result = CDec(0)
    For Each m In Picked.Keys
        value = ValAt(source, "BP" & index & Sep & P(year, CLng(m)))
        If index = 3 Or index = 4 Then
            If IsNull(value) Then
                first = ValAt(source, "BP1|" & P(year, CLng(m))): second = ValAt(source, "BP2|" & P(year, CLng(m)))
                If Not IsNull(first) And Not IsNull(second) Then
                    If first = second Then value = second
                ElseIf Not IsNull(second) Then
                    value = second
                ElseIf Not IsNull(first) Then
                    value = first
                End If
            End If
        End If
        result = Plus(result, value)
    Next m
    IndicatorBudget = result
End Function

Public Sub WriteIdentity(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim ws As Worksheet, r As Long, n As Variant, meta As Variant, kind As String
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = "_Rapor Kimliği"
    If general Then kind = "Direktörlük" Else kind = "Müdürlük"
    RowValues ws, 1, Array("ButceSapmasi-1", kind, CStr(ReportYear) & "-" & Format$(CLng(Picked.Keys()(Picked.Count - 1)), "00"), P(CloseYear, CloseMonth), Signature())
    r = 2
    For Each n In scope.Keys
        meta = Depts(n): RowValues ws, r, Array(meta(0), meta(1), meta(2)): r = r + 1
    Next n
    ws.Visible = xlSheetVeryHidden
End Sub

Public Sub WriteOneReport(ByVal path As String, ByVal scope As Object, ByVal general As Boolean)
    Dim wb As Workbook, placeholder As Worksheet, errText As String, previousCalc As XlCalculation, sheet As Worksheet
    previousCalc = Application.Calculation
    On Error GoTo Bad
    Set wb = Application.Workbooks.Add(xlWBATWorksheet): Set placeholder = wb.Worksheets(1)
    WriteSummary wb, scope, general
    WriteExpenses wb, scope, general
    WriteUnits wb, scope, general
    WriteMaterial wb, scope, general
    WriteForecast wb, scope, general
    WriteVolume wb, scope, general
    WriteAudit wb, scope
    WriteIdentity wb, scope, general
    WriteMemoReport wb, scope, general
    placeholder.Delete
    WriteExecutive wb, scope, general
    WritePano wb, scope
    wb.CheckCompatibility = False
    Application.Calculation = xlCalculationAutomatic
    ' Yalnız bu rapor hesaplanır; açık tüm kitapları tam yeniden hesaplamak gereksiz yavaştı.
    For Each sheet In wb.Worksheets
        sheet.Calculate
    Next sheet
    wb.Worksheets(1).Activate
    wb.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    Application.Calculation = previousCalc
    wb.Close False
    Exit Sub
Bad:
    errText = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.Calculation = previousCalc
    On Error GoTo 0
    Fail errText
End Sub

Public Sub SaveReports()
    Dim fso As Object, stage As String, final As String, stamp As String, suffix As Long, mgrs As Object
    Dim scope As Object, manager As Variant, i As Long, errText As String, monthId As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(OutputPath) Then Fail "Çıktı klasörü yok; panelden klasör seçin."
    stamp = Format$(Now, "yyyymmdd_hhnnss")
    final = fso.BuildPath(OutputPath, "Raporlar_" & stamp)
    Do While fso.FolderExists(final)
        suffix = suffix + 1: final = fso.BuildPath(OutputPath, "Raporlar_" & stamp & "_" & suffix)
    Loop
    stage = final & "_HAZIRLANIYOR": fso.CreateFolder stage
    On Error GoTo Bad
    Set scope = DeptScope(): Set mgrs = Managers(scope)
    monthId = Replace(Replace(Signature(), "aylar:", ""), ",", "-")
    WriteOneReport fso.BuildPath(stage, "Direktorluk_Genel_" & ReportYear & "_Aylar_" & monthId & ".xlsx"), scope, True
    For Each manager In mgrs.Keys
        WorkGate "Müdürlük raporu: " & CStr(manager)
        i = i + 1: Set scope = DeptScope(CStr(manager))
        WriteOneReport fso.BuildPath(stage, "Mudurluk_" & Format$(i, "00") & "_" & CleanName(CStr(manager)) & "_" & ReportYear & "_Aylar_" & monthId & ".xlsx"), scope, False
    Next manager
    WorkGate "Tamamlanan raporlar yayımlanıyor"
    fso.MoveFolder stage, final: RunPath = final
    Exit Sub
Bad:
    errText = Err.Description
    On Error Resume Next
    fso.DeleteFolder stage, True
    On Error GoTo 0
    Fail "Raporlar tamamlanamadı; bu çalışmanın geçici çıktıları kaldırıldı. " & errText
End Sub
