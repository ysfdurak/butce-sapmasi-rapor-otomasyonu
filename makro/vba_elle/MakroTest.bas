Attribute VB_Name = "MakroTest"
Option Explicit

Private passed As Long
Private failures As String

Private Sub Expect(ByVal label As String, ByVal actual As Variant, ByVal expected As Variant, Optional ByVal tolerance As Double = 0.005)
    Dim ok As Boolean
    If IsNull(expected) Then
        ok = IsNull(actual)
    ElseIf IsNumeric(expected) And Not IsNull(actual) Then
        ok = Abs(CDbl(actual) - CDbl(expected)) <= tolerance
    ElseIf Not IsNull(actual) Then
        ok = (CStr(actual) = CStr(expected))
    End If
    If ok Then
        passed = passed + 1
    Else
        failures = failures & vbCrLf & label & " beklenen=" & ShowValue(expected) & " bulunan=" & ShowValue(actual)
    End If
End Sub

Private Function ShowValue(ByVal value As Variant) As String
    If IsNull(value) Then ShowValue = "BOÞ" Else ShowValue = CStr(value)
End Function

Public Sub SentetikTest()
    Dim security As Long, events As Boolean, scope As Object, n As Variant, row As Variant, a As Variant, totalRate As Variant
    Dim totalCost As Variant, i As Long, text As String, ws As Worksheet, source As Object, codes As Object
    Dim oldCount As Variant, oldCost As Variant, other As Object, testBook As Workbook, labels As Variant
    On Error GoTo Bad
    security = Application.AutomationSecurity: events = Application.EnableEvents
    Application.AutomationSecurity = 3: Application.EnableEvents = False
    passed = 0: failures = ""
    SettingsPath = ThisWorkbook.Path & "\ornek\Ayarlar.xlsx"
    CloseYear = 2026: CloseMonth = 10: ReportYear = 2026
    Set Picked = Dict(): Set Warnings = New Collection
    For i = 1 To 10
        Picked(CStr(i)) = True
    Next i
    LoadSettings: LoadSnapshots: LoadControl: LoadIndices
    Set scope = DeptScope(): CalculateUnits: CalculateChecks
    Expect "Bölüm sayýsý", Depts.Count, 12
    Expect "Müdürlük sayýsý", Managers(scope).Count, 4
    Expect "Gerçekleþen gider TL", ExpenseTotal(scope, "A", 2026), CDec(3449383390.31)
    Expect "Önceki yýl ayný aylar TL", ExpenseTotal(scope, "A", 2025), CDec(3004130754.88)
    Expect "Bütçe gider TL", ExpenseTotal(scope, "B", 2026), CDec(3440938980.13)
    Expect "Bütçe - gerçekleþen TL", Minus(ExpenseTotal(scope, "B", 2026), ExpenseTotal(scope, "A", 2026)), CDec(-8444410.18)
    Expect "Birim maliyete uygun bölüm", Units.Count, 12
    totalCost = CDec(0): totalRate = CDec(0)
    For Each n In Units.Keys
        row = Units(n): totalCost = Plus(totalCost, row(0))
    Next n
    For Each n In Units.Keys
        row = Units(n): totalRate = Plus(totalRate, Times(Divide(row(0), totalCost), row(9)))
    Next n
    Expect "Aðýrlýklý reel iyileþme", totalRate, -0.0152069853055713, 0.0000001
    Expect "Ocak ciro TL", ValAt(Revenue, "2026-01"), CDec(1768000000)
    Expect "Ocak bütçe ciro TL", ValAt(RevenueBudget, "2026-01"), CDec(1872000000)
    Set codes = ItemNames(): Expect "Birleþmiþ gider kalemleri", codes.Count, 26
    Expect "Revizyon kayýtlarý", Revisions.Count, 26
    For Each row In Checks
        Expect "Dip toplam kontrolü " & row(0) & " / " & row(3), row(7), "Doðrulandý"
    Next row
    Set Picked = Dict(): Picked("1") = True: Picked("3") = True: Picked("5") = True
    Expect "Ocak Mart Mayýs gerçekleþen TL", ExpenseTotal(scope, "A", 2026), CDec(913186126.67)
    Expect "Ocak Mart Mayýs bütçe TL", ExpenseTotal(scope, "B", 2026), CDec(1094635989.61)
    Set source = CurrentSnap("Tartým"): oldCount = ValAt(source, "AC|2025-01")
    Set other = CurrentSnap("Ýade Kabul"): oldCost = ValAt(other, "AT|2025-01")
    source("AC|2025-01") = Null: other("AT|2025-01") = Null
    CalculateUnits
    Expect "Eksik geçmiþ veride bölüm dýþlama", Units.Count, 10
    row = Units(K("Giriþ"))
    Expect "Müdürlükte kalan tek bölümün aðýrlýðý", row(10), 1, 0.0000001
    source("AC|2025-01") = oldCount: other("AT|2025-01") = oldCost
    CalculateUnits
    Expect "Eksik + deðer", Plus(Null, 10), Null
    Expect "Sýfýra bölme", Divide(10, 0), Null
    Expect "Dosya adý Þubat", FilenamePeriod("2026 Þubat - bölüm x.xlsx"), "2026-02"
    Expect "Dosya adý ters", FilenamePeriod("Ekim 2026 - Dosya 1.xlsx"), "2026-10"
    Expect "Tekrarlý dosya dönemi", FilenamePeriod("2026 Þubat 2025 Mart.xlsx"), ""
    Set testBook = Application.Workbooks.Add(xlWBATWorksheet)
    labels = Array("Kadrolu çalýþan", "Taþeron çalýþan", "Fiili çalýþan", "Planlý çalýþan", "Ýþlenen koli adedi")
    For i = 0 To 4
        testBook.Worksheets(1).Cells(60 + i, 57).Value2 = labels(i)
        Expect "Kaynak baþlýðý " & i + 1, SourceIndicatorTitle(testBook.Worksheets(1), i + 1), labels(i)
    Next i
    testBook.Worksheets(1).Cells(60, 57).ClearContents
    Expect "Eksik baþlýk açýk uyarý", SourceIndicatorTitle(testBook.Worksheets(1), 1), "Baþlýðý eksik gösterge (BE60)"
    testBook.Close False: Set testBook = Nothing
    If failures = "" Then
        text = CStr(passed) & " kontrol geçti. Rapor oluþtur ile dosya üretimini ve Excel görünümünü ayrýca kontrol edin."
    Else
        text = CStr(passed) & " kontrol geçti; baþarýsýz kontroller:" & failures
    End If
    Set ws = ThisWorkbook.Worksheets("Test Sonuçlarý")
    ws.Range("A1").Value2 = "Windows Excel sentetik hesap testi"
    ws.Range("A2").Value2 = Now
    ws.Range("A4").Value2 = text: ws.Range("A4").WrapText = True
    ws.Columns("A").ColumnWidth = 115: ws.Rows(4).AutoFit: ws.Activate
    MsgBox text, vbInformation
Done:
    Application.AutomationSecurity = security: Application.EnableEvents = events: Application.StatusBar = False
    Exit Sub
Bad:
    text = Err.Description
    On Error Resume Next
    If Not testBook Is Nothing Then testBook.Close False
    On Error GoTo 0
    MsgBox "Sentetik test çalýþtýrýlamadý: " & text, vbExclamation
    GoTo Done
End Sub
