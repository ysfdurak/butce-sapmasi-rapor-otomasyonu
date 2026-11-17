Attribute VB_Name = "MakroSurum"
Option Explicit

' Bütçe sürümleri, bütçe tablolarý, bütçe deðiþimi, trend ve girdi kaydý (V5.4).
' Python surum.py ve Power Query þablonuyla ayný kurallar (docs/VERI_VE_HESAP.md):
'   Güncel bütçe: kapanýþ dosyasý. Önceki ay bütçesi: ayný yýlýn bir önceki kapanýþ dosyasý.
'   Ýlk bütçe: bölümün kapanýþ yýlýnda okunan ilk kapanýþ dosyasý.
'   Ayarlar > Rapor > karsilastirma_butcesi = "Ýlk bütçe" ise raporlarýn bütçesi ilk bütçedir
'   (bütçe iþlem adedi ve bütçe cirosu güncel kalýr; revizyonlar kaynak dosyalardan hesaplanýr).
' Kapanýþ anlýk görüntüsünde özgün bütçe "OB|..." / "OBT|..." anahtarlarýnda saklanýr (HazirlaSurumler).

Private Function Mn() As String
    Mn = ChrW(8722)
End Function

Private Function Ge2() As String
    Ge2 = ChrW(8805)
End Function

Public Function FirstMode() As Boolean
    FirstMode = (Left$(K(Opt("Rapor", "karsilastirma_butcesi", "")), 3) = "ilk")
End Function

Public Function TrendThreshold() As Variant
    Dim v As Variant
    v = Opt("Rapor", "trend_aylik_esik", "")
    If CStr(v) = "" Then TrendThreshold = CDec(RuleAmount()) / 12 Else TrendThreshold = CDec(v)
End Function

Public Function TrendLeast() As Long
    Dim v As Variant
    v = Opt("Rapor", "trend_en_az_ay", "")
    If CStr(v) = "" Then TrendLeast = 3 Else TrendLeast = CLng(v)
End Function

Public Function FirstSnap(ByVal n As String) As Object
    Dim m As Long
    For m = 1 To CloseMonth
        Set FirstSnap = CurrentSnap(n, m)
        If Not FirstSnap Is Nothing Then Exit Function
    Next m
End Function

Public Function PrevSnap(ByVal n As String) As Object
    If CloseMonth > 1 Then Set PrevSnap = CurrentSnap(n, CloseMonth - 1)
End Function

' Revizyon ve kontroller hesaplandýktan sonra, birim maliyetten önce çaðrýlýr (MakroPanel.RaporOlustur).
Public Sub HazirlaSurumler()
    Dim n As Variant, d As Object, f As Object, keyText As Variant, k2 As String, prefix As String
    Dim total As Variant, code As Variant
    For Each n In Depts.Keys
        Set d = CurrentSnap(CStr(n))
        If Not d Is Nothing Then
            If Not d.Exists("SurumHazir") Then
                For Each keyText In d.Keys
                    k2 = CStr(keyText)
                    If Left$(k2, 2) = "B|" Or Left$(k2, 3) = "BT|" Then d("O" & k2) = d(k2)
                Next keyText
                d("SurumHazir") = True
            End If
            If FirstMode() Then
                Set f = FirstSnap(CStr(n))
                ' Kalem bütçesi ilk dosyadan alýnýr; toplam yalnýz kapanýþ dosyasýndaki kalemlerden, katý toplanýr.
                For Each keyText In d.Keys
                    k2 = CStr(keyText)
                    If Left$(k2, 2) = "B|" Then
                        If f Is Nothing Then
                            d(k2) = Null
                        Else
                            prefix = "B|": If f.Exists("SurumHazir") Then prefix = "OB|"
                            d(k2) = ValAt(f, prefix & Mid$(k2, 3))
                        End If
                    End If
                Next keyText
                For Each keyText In d.Keys
                    k2 = CStr(keyText)
                    If Left$(k2, 3) = "BT|" Then
                        total = CDec(0)
                        For Each code In d("names").Keys
                            total = Plus(total, ValAt(d, "B|" & Mid$(k2, 4) & Sep & code))
                        Next code
                        d(k2) = total
                    End If
                Next keyText
            End If
        End If
    Next n
End Sub

' kind: "A" gerçekleþen (rapor yýlý), "G" güncel bütçe, "I" ilk bütçe, "P" önceki ay bütçesi (kapanýþ yýlý),
' "T" gerçekleþen - karþýlaþtýrma bütçesi (kapanýþ yýlý, kapanýþ ayýna kadar)
Private Function DeptMonth(ByVal n As String, ByVal kind As String, ByVal m As Long) As Variant
    Dim d As Object, prefix As String
    DeptMonth = Null
    Select Case kind
        Case "A": Set d = CurrentSnap(n): If d Is Nothing Then Exit Function
            DeptMonth = ValAt(d, "AT|" & P(ReportYear, m))
        Case "G": Set d = CurrentSnap(n): If d Is Nothing Then Exit Function
            prefix = "OBT|": If Not d.Exists("SurumHazir") Then prefix = "BT|"
            DeptMonth = ValAt(d, prefix & P(CloseYear, m))
        Case "I": Set d = FirstSnap(n): If d Is Nothing Then Exit Function
            prefix = "BT|": If d.Exists("SurumHazir") Then prefix = "OBT|"
            DeptMonth = ValAt(d, prefix & P(CloseYear, m))
        Case "P": Set d = PrevSnap(n): If d Is Nothing Then Exit Function
            DeptMonth = ValAt(d, "BT|" & P(CloseYear, m))
        Case "T": Set d = CurrentSnap(n): If d Is Nothing Or m > CloseMonth Then Exit Function
            DeptMonth = Minus(ValAt(d, "AT|" & P(CloseYear, m)), ValAt(d, "BT|" & P(CloseYear, m)))
    End Select
End Function

Private Function GroupMonth(ByVal scope As Object, ByVal kind As String, ByVal m As Long) As Variant
    Dim n As Variant, result As Variant
    GroupMonth = Null
    If scope.Count = 0 Then Exit Function
    result = CDec(0)
    For Each n In scope.Keys
        result = Plus(result, DeptMonth(CStr(n), kind, m))
    Next n
    GroupMonth = result
End Function

Private Function GroupPicked(ByVal scope As Object, ByVal kind As String) As Variant
    Dim m As Variant, result As Variant
    result = CDec(0)
    For Each m In Picked.Keys
        result = Plus(result, GroupMonth(scope, kind, CLng(m)))
    Next m
    GroupPicked = result
End Function

Private Function YearTotal(ByVal scope As Object, ByVal kind As String) As Variant
    Dim m As Long, result As Variant, n As Variant, part As Variant
    ' Bölüm bazýnda 12 ay katý toplanýr, sonra bölümler katý toplanýr (Python surum._group ile ayný).
    YearTotal = Null
    If scope.Count = 0 Then Exit Function
    result = CDec(0)
    For Each n In scope.Keys
        part = CDec(0)
        For m = 1 To 12
            part = Plus(part, DeptMonth(CStr(n), kind, m))
        Next m
        result = Plus(result, part)
    Next n
    YearTotal = result
End Function

Private Function SortedCodes(ByVal names As Object) As Variant
    Dim codes() As Long, i As Long, j As Long, t As Long, c As Variant
    If names.Count = 0 Then SortedCodes = Array(): Exit Function
    ReDim codes(0 To names.Count - 1)
    For Each c In names.Keys
        codes(i) = CLng(c): i = i + 1
    Next c
    For i = 1 To UBound(codes)
        t = codes(i): j = i - 1
        Do While j >= 0
            If codes(j) <= t Then Exit Do
            codes(j + 1) = codes(j): j = j - 1
        Loop
        codes(j + 1) = t
    Next i
    SortedCodes = codes
End Function

Private Function Months3() As Variant
    Months3 = Array("Oca", "Þub", "Mar", "Nis", "May", "Haz", "Tem", "Aðu", "Eyl", "Eki", "Kas", "Ara")
End Function

Private Function Joined(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim result() As Variant, i As Long, n As Long
    ReDim result(0 To UBound(a) - LBound(a) + UBound(b) - LBound(b) + 1)
    For i = LBound(a) To UBound(a)
        result(n) = a(i): n = n + 1
    Next i
    For i = LBound(b) To UBound(b)
        result(n) = b(i): n = n + 1
    Next i
    Joined = result
End Function

Private Function VersionSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Set VersionSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    VersionSheet.Name = name
    VersionSheet.Cells.Font.Name = "Calibri": VersionSheet.Cells.Font.Size = 11
End Function

' ---------------------------------------------------------------- Bütçe Sürümleri
Public Sub WriteVersions(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim ws As Worksheet, item As Variant, r As Long, a As Variant, c As Variant, f As Variant, gap As Variant
    Set ws = VersionSheet(wb, "Bütçe Sürümleri")
    RowValues ws, 1, Array("Bölüm / müdürlük / direktörlük", "Satýr türü", "Gerçekleþen (TL)", "Güncel bütçe (TL)", "Ýlk bütçe (TL)", _
        "Güncel " & Mn() & " ilk bütçe (TL)", "Gerçekleþen " & Mn() & " güncel bütçe (TL)", "Gerçekleþen " & Mn() & " ilk bütçe (TL)", "Ýlk bütçeye göre sapma (%)")
    Header ws.Range("A1:I1")
    ws.Range("K1").Value2 = "Ýlk bütçe: bölümün bu yýl okunan ilk kapanýþ dosyasýndaki bütçe. Klasörde Ocak dosyasý yoksa ilk bulunan ay alýnýr (Girdi Kaydý)."
    r = 2
    If ReportYear <> CloseYear Then
        ws.Cells(2, 1).Value2 = "Seçili yýl bütçe yýlý deðil."
    Else
        For Each item In ReportLayout(scope, general)
            a = GroupPicked(item(2), "A"): c = GroupPicked(item(2), "G"): f = GroupPicked(item(2), "I")
            gap = Minus(a, f)
            If IsNull(f) Then
                RowValues ws, r, Array(item(0), item(1), a, c, f, Minus(c, f), Minus(a, c), gap, Null)
            ElseIf f = 0 Then
                RowValues ws, r, Array(item(0), item(1), a, c, f, Minus(c, f), Minus(a, c), gap, Null)
            Else
                RowValues ws, r, Array(item(0), item(1), a, c, f, Minus(c, f), Minus(a, c), gap, Divide(gap, f))
            End If
            r = r + 1
        Next item
    End If
    If r > 2 Then ws.Range("C2:H" & r - 1).NumberFormat = MoneyFmt: ws.Range("I2:I" & r - 1).NumberFormat = PctFmt
    ws.Columns("A").ColumnWidth = 34: ws.Columns("B").ColumnWidth = 20: ws.Columns("C:H").ColumnWidth = 21
    ws.Columns("I").ColumnWidth = 14: ws.Columns("K").ColumnWidth = 60
End Sub

' ---------------------------------------------------------------- Bütçe - Bölüm / Müdürlük / Direktörlük
Public Sub WriteBudgetTables(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim level As Variant, ws As Worksheet, item As Variant, r As Long, m As Long, values() As Variant, first As Long
    Dim total As Variant, initial As Variant, parent As String, wanted As Boolean, n As Variant, initialPart As Variant
    For Each level In Array("Bölüm", "Müdürlük", "Direktörlük")
        Set ws = VersionSheet(wb, "Bütçe - " & level)
        ws.Range("A1").Value2 = level & " bütçeleri " & ChrW(183) & " " & CloseYear & " (kapanýþ " & P(CloseYear, CloseMonth) & " dosyasýndaki güncel bütçe)"
        ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
        ws.Range("A2").Value2 = "Ay boþsa bütçe verisi eksiktir; yýl toplamý yalnýz 12 ay doluysa hesaplanýr."
        Select Case level
            Case "Bölüm": parent = "Müdürlük"
            Case "Müdürlük": parent = "Direktörlük"
            Case Else: parent = ""
        End Select
        If parent = "" Then
            RowValues ws, 4, Joined(Array(level), Joined(Months3(), Array("Yýl toplamý", "Ýlk bütçe yýl toplamý", "Güncel " & Mn() & " ilk (TL)")))
            first = 2
        Else
            RowValues ws, 4, Joined(Array(level, parent), Joined(Months3(), Array("Yýl toplamý", "Ýlk bütçe yýl toplamý", "Güncel " & Mn() & " ilk (TL)")))
            first = 3
        End If
        Header ws.Range(ws.Cells(4, 1), ws.Cells(4, first + 14))
        r = 5
        For Each item In ReportLayout(scope, general)
            Select Case level
                Case "Bölüm": wanted = (item(1) = "Bölüm")
                Case "Müdürlük": wanted = (item(1) = "Müdürlük toplamý")
                Case Else: wanted = (item(1) = "Direktörlük toplamý" Or item(1) = "Genel toplam")
            End Select
            If wanted Then
                ReDim values(0 To first + 14)
                values(0) = item(0)
                For Each n In item(2).Keys
                    If level = "Bölüm" Then values(1) = Depts(n)(1)
                    If level = "Müdürlük" Then values(1) = Depts(n)(2)
                    Exit For
                Next n
                total = CDec(0)
                For m = 1 To 12
                    values(first - 1 + m) = GroupMonth(item(2), "G", m)
                    total = Plus(total, values(first - 1 + m))
                Next m
                ' Ýlk bütçe yýl toplamý: her bölümün 12 ayý katý, sonra bölümler katý
                initial = CDec(0)
                For Each n In item(2).Keys
                    initialPart = CDec(0)
                    For m = 1 To 12
                        initialPart = Plus(initialPart, DeptMonth(CStr(n), "I", m))
                    Next m
                    initial = Plus(initial, initialPart)
                Next n
                values(first + 12) = total: values(first + 13) = initial: values(first + 14) = Minus(total, initial)
                RowValues ws, r, values
                r = r + 1
            End If
        Next item
        If r > 5 Then ws.Range(ws.Cells(5, first), ws.Cells(r - 1, first + 14)).NumberFormat = MoneyFmt
        ws.Columns(1).ColumnWidth = 30
        If first = 3 Then ws.Columns(2).ColumnWidth = 24
        ws.Range(ws.Cells(1, first), ws.Cells(1, first + 14)).EntireColumn.ColumnWidth = 15
    Next level
End Sub

' ---------------------------------------------------------------- Bütçe Deðiþimi
Public Function ChangeItems(ByVal scope As Object) As Collection
    ' Kalem satýrý: (müdürlük, bölüm, kod, kalem, deðiþen ay, önceki yýl, güncel yýl, deðiþim, 12 aylýk fark)
    Dim result As Collection, n As Variant, d As Object, o As Object, codes As Variant, i As Long, code As String
    Dim m As Long, g As Variant, b As Variant, diff As Variant, changed As Long, gy As Variant, py As Variant, tol As Double
    Dim row() As Variant
    Set result = New Collection
    Set ChangeItems = result
    If CloseMonth = 1 Then Exit Function
    tol = CDbl(Opt("Kontrol", "tolerans", 0.01))
    For Each n In scope.Keys
        Set d = CurrentSnap(CStr(n)): Set o = PrevSnap(CStr(n))
        If Not d Is Nothing And Not o Is Nothing Then
            codes = SortedCodes(d("names"))
            For i = LBound(codes) To UBound(codes)
                code = CStr(codes(i))
                If o("names").Exists(code) Then
                    ReDim row(0 To 19)
                    changed = 0: gy = CDec(0): py = CDec(0)
                    For m = 1 To 12
                        g = ValAt(d, "OB|" & P(CloseYear, m) & Sep & code)
                        If Not d.Exists("SurumHazir") Then g = ValAt(d, "B|" & P(CloseYear, m) & Sep & code)
                        b = ValAt(o, "B|" & P(CloseYear, m) & Sep & code)
                        diff = Minus(g, b)
                        If Not IsNull(diff) Then
                            If Abs(diff) > tol Then changed = changed + 1
                        End If
                        row(7 + m) = diff
                        gy = Plus(gy, g): py = Plus(py, b)
                    Next m
                    If changed > 0 Then
                        row(0) = Depts(n)(1): row(1) = Depts(n)(0): row(2) = CLng(code): row(3) = d("names")(code)
                        row(4) = changed: row(5) = py: row(6) = gy: row(7) = Minus(gy, py)
                        result.Add row
                    End If
                End If
            Next i
        End If
    Next n
End Function

Public Function HasPrevious(ByVal scope As Object) As Boolean
    Dim n As Variant
    If CloseMonth = 1 Then Exit Function
    For Each n In scope.Keys
        If Not PrevSnap(CStr(n)) Is Nothing Then HasPrevious = True: Exit Function
    Next n
End Function

Public Sub WriteChange(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim ws As Worksheet, item As Variant, r As Long, m As Long, values() As Variant, c As Variant, p0 As Variant
    Dim changed As Boolean, tol As Double, total As Variant, found As Boolean, py As Variant, gy As Variant, row As Variant
    Set ws = VersionSheet(wb, "Bütçe Deðiþimi")
    If CloseMonth > 1 Then
        ws.Range("A1").Value2 = "Bütçe deðiþimi: " & P(CloseYear, CloseMonth - 1) & " " & ChrW(8594) & " " & P(CloseYear, CloseMonth) & " kapanýþ dosyasý (güncel " & Mn() & " önceki)"
        ws.Range("A2").Value2 = "Yalnýz bütçesi deðiþen satýrlar. Önceki ay dosyasý olmayan bölüm karþýlaþtýrýlmaz."
    Else
        ws.Range("A1").Value2 = "Bütçe deðiþimi:  " & ChrW(8594) & " " & P(CloseYear, CloseMonth) & " kapanýþ dosyasý (güncel " & Mn() & " önceki)"
        ws.Range("A2").Value2 = "Ocak kapanýþýnda önceki ay farklý yýlýn bütçesidir; karþýlaþtýrma yapýlmaz."
    End If
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    RowValues ws, 4, Joined(Array("Bölüm / müdürlük / direktörlük", "Satýr türü"), Joined(Months3(), Array("Toplam deðiþim (TL)", "Önceki yýl bütçesi", "Güncel yýl bütçesi")))
    Header ws.Range("A4:Q4")
    ws.Cells(3, 21).Value2 = "Bütçesi deðiþen kalemler": ws.Cells(3, 21).Font.Bold = True
    RowValues ws, 4, Joined(Array("Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Deðiþen ay", "Önceki yýl bütçesi (TL)", "Güncel yýl bütçesi (TL)", "Deðiþim (TL)"), Months3()), 21
    Header ws.Range(ws.Cells(4, 21), ws.Cells(4, 40))
    r = 5
    If CloseMonth > 1 Then
        tol = CDbl(Opt("Kontrol", "tolerans", 0.01))
        For Each item In ReportLayout(scope, general)
            ReDim values(0 To 16)
            values(0) = item(0): values(1) = item(1)
            changed = False: total = CDec(0): found = False
            For m = 1 To 12
                c = GroupMonth(item(2), "G", m): p0 = GroupMonth(item(2), "P", m)
                values(1 + m) = Minus(c, p0)
                If Not IsNull(values(1 + m)) Then
                    found = True: total = total + values(1 + m)
                    If Abs(values(1 + m)) > tol Then changed = True
                End If
            Next m
            If changed Then
                values(14) = total
                py = YearTotal(item(2), "P"): gy = YearTotal(item(2), "G")
                values(15) = py: values(16) = gy
                RowValues ws, r, values
                r = r + 1
            End If
        Next item
        If r = 5 Then ws.Cells(5, 1).Value2 = "Bütçe deðiþmedi": r = 6
    End If
    If r > 5 Then ws.Range("C5:Q" & r - 1).NumberFormat = MoneyFmt
    r = 5
    For Each row In ChangeItems(scope)
        RowValues ws, r, row, 21
        r = r + 1
    Next row
    If CloseMonth > 1 And r = 5 Then ws.Cells(5, 21).Value2 = "Deðiþen kalem yok": r = 6
    If r > 5 Then
        ws.Range(ws.Cells(5, 26), ws.Cells(r - 1, 28)).NumberFormat = MoneyFmt
        ws.Range(ws.Cells(5, 29), ws.Cells(r - 1, 40)).NumberFormat = "#,##0;-#,##0;0"
    End If
    ws.Columns("A").ColumnWidth = 32: ws.Columns("B").ColumnWidth = 18
    ws.Columns("U:V").ColumnWidth = 16: ws.Columns("X").ColumnWidth = 26: ws.Columns("Z:AB").ColumnWidth = 18
End Sub

' ---------------------------------------------------------------- Trend
Public Function TrendItems(ByVal scope As Object) As Collection
    ' (müdürlük, bölüm, kod, kalem, üst üste aleyhte ay, toplam aleyhte ay, seri toplamý, kapanýþ ayý farký)
    Dim result As Collection, n As Variant, d As Object, codes As Variant, i As Long, code As String, m As Long
    Dim g As Variant, b As Variant, bad(1 To 12) As Boolean, streak As Long, badCount As Long, series As Variant
    Dim threshold As Variant, least As Long, row As Variant, j As Long, placed As Boolean
    Set result = New Collection
    Set TrendItems = result
    If ReportYear <> CloseYear Then Exit Function
    threshold = TrendThreshold(): least = TrendLeast()
    For Each n In scope.Keys
        Set d = CurrentSnap(CStr(n))
        If Not d Is Nothing Then
            codes = SortedCodes(d("names"))
            For i = LBound(codes) To UBound(codes)
                code = CStr(codes(i))
                For m = 1 To 12
                    g = ValAt(d, "A|" & P(CloseYear, m) & Sep & code): b = ValAt(d, "B|" & P(CloseYear, m) & Sep & code)
                    bad(m) = False
                    If Not IsNull(g) And Not IsNull(b) Then bad(m) = (CDec(g) - CDec(b) >= threshold)
                Next m
                streak = 0
                For m = CloseMonth To 1 Step -1
                    If Not bad(m) Then Exit For
                    streak = streak + 1
                Next m
                If streak >= least Then
                    badCount = 0: series = CDec(0)
                    For m = 1 To CloseMonth
                        If bad(m) Then badCount = badCount + 1
                    Next m
                    For m = CloseMonth - streak + 1 To CloseMonth
                        series = series + CDec(ValAt(d, "A|" & P(CloseYear, m) & Sep & code)) - CDec(ValAt(d, "B|" & P(CloseYear, m) & Sep & code))
                    Next m
                    row = Array(Depts(n)(1), Depts(n)(0), CLng(code), d("names")(code), streak, badCount, series, _
                        Minus(ValAt(d, "A|" & P(CloseYear, CloseMonth) & Sep & code), ValAt(d, "B|" & P(CloseYear, CloseMonth) & Sep & code)))
                    ' Sýra: seri uzunluðu azalan, sonra seri toplamý azalan; eþitlikte kalem sýrasý korunur.
                    placed = False
                    For j = 1 To result.Count
                        If row(4) > result(j)(4) Or (row(4) = result(j)(4) And row(6) > result(j)(6)) Then
                            result.Add row, Before:=j: placed = True: Exit For
                        End If
                    Next j
                    If Not placed Then result.Add row
                End If
            Next i
        End If
    Next n
End Function

Public Sub WriteTrend(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim ws As Worksheet, item As Variant, r As Long, m As Long, values() As Variant, row As Variant, budgetYear As Boolean
    Dim co As ChartObject
    budgetYear = (ReportYear = CloseYear)
    Set ws = VersionSheet(wb, "Trend")
    ws.Range("A1").Value2 = "Aylýk sapma trendi (gerçekleþen " & Mn() & " karþýlaþtýrma bütçesi)"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("B3").Value2 = "Aylýk eþik (TL)": ws.Range("C3").Value2 = CDbl(TrendThreshold()): ws.Range("C3").NumberFormat = "#,##0"
    ws.Range("B4").Value2 = "En az üst üste ay": ws.Range("C4").Value2 = TrendLeast()
    If Not budgetYear Then ws.Range("A5").Value2 = "Seçili yýl bütçe yýlý deðil; trend hesaplanmadý."
    RowValues ws, 7, Joined(Array("Bölüm / müdürlük / direktörlük", "Satýr türü"), Months3())
    Header ws.Range("A7:N7")
    ReDim values(0 To 13)
    values(0) = "Kapsam toplamý": values(1) = "Toplam"
    For m = 1 To 12
        values(1 + m) = Null
        If budgetYear And m <= CloseMonth Then values(1 + m) = GroupMonth(scope, "T", m)
    Next m
    RowValues ws, 8, values
    r = 9
    For Each item In ReportLayout(scope, general)
        ReDim values(0 To 13)
        values(0) = item(0): values(1) = item(1)
        For m = 1 To 12
            values(1 + m) = Null
            If budgetYear Then values(1 + m) = GroupMonth(item(2), "T", m)
        Next m
        RowValues ws, r, values
        r = r + 1
    Next item
    ws.Range("C8:N" & r - 1).NumberFormat = "#,##0;-#,##0;0"
    ws.Cells(18, 17).Value2 = "Üst üste aleyhte giden kalemler": ws.Cells(18, 17).Font.Bold = True
    RowValues ws, 19, Array("Müdürlük", "Bölüm", "Gider kodu", "Gider kalemi", "Üst üste aleyhte ay", "Toplam aleyhte ay", "Seri toplam sapma (TL)", "Kapanýþ ayý sapmasý (TL)"), 17
    Header ws.Range(ws.Cells(19, 17), ws.Cells(19, 24))
    r = 20
    For Each row In TrendItems(scope)
        RowValues ws, r, row, 17
        r = r + 1
    Next row
    If budgetYear And r = 20 Then ws.Cells(20, 17).Value2 = "Seri yok": r = 21
    If r > 20 Then ws.Range(ws.Cells(20, 23), ws.Cells(r - 1, 24)).NumberFormat = MoneyFmt
    ws.Columns("A").ColumnWidth = 32: ws.Columns("B").ColumnWidth = 18: ws.Columns("C:N").ColumnWidth = 13
    ws.Columns("Q").ColumnWidth = 18: ws.Columns("R").ColumnWidth = 16: ws.Columns("T").ColumnWidth = 26
    ws.Columns("W:X").ColumnWidth = 18
    Set co = ws.ChartObjects.Add(ws.Range("Q2").Left, ws.Range("Q2").Top, 520, 200)
    co.Chart.ChartType = xlLine
    co.Chart.SetSourceData ws.Range("C8:N8"), xlRows
    co.Chart.SeriesCollection(1).XValues = ws.Range("C7:N7")
    co.Chart.HasTitle = True: co.Chart.ChartTitle.Text = "Kapsam toplamý: aylýk sapma (TL)"
    co.Chart.HasLegend = False
End Sub

' ---------------------------------------------------------------- Girdi Kaydý
Private Sub InputRow(ByVal ws As Worksheet, ByVal r As Long, ByVal kind As String, ByVal path As String, ByVal status As String, ByVal period As String)
    Dim fso As Object, name As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    name = fso.GetFileName(path)
    If Not fso.FileExists(path) Then
        RowValues ws, r, Array(kind, name, Null, Null, Null, "Bulunamadý")
    Else
        If period = "" Then
            RowValues ws, r, Array(kind, name, Null, Round(FileLen(path) / 1024, 1), Format$(FileDateTime(path), "yyyy-mm-dd hh:nn:ss"), status)
        Else
            RowValues ws, r, Array(kind, name, period, Round(FileLen(path) / 1024, 1), Format$(FileDateTime(path), "yyyy-mm-dd hh:nn:ss"), status)
        End If
    End If
End Sub

Public Sub WriteInputs(ByVal wb As Workbook)
    Dim ws As Worksheet, folder As String, name As String, names As Collection, i As Long, j As Long, r As Long
    Dim list() As String, t As String, period As String
    Set ws = VersionSheet(wb, "Girdi Kaydý")
    ws.Range("A1").Value2 = "Raporun okuduðu dosyalar": ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value2 = "Kaynak dosya sonradan deðiþirse boyut veya deðiþtirilme zamaný farklý olur; aylýk raporu ayrý saklamak, o ayýn girdileriyle birlikte arþivlenmesini saðlar."
    RowValues ws, 4, Array("Tür", "Dosya", "Dönem", "Boyut (KB)", "Deðiþtirilme zamaný", "Durum")
    Header ws.Range("A4:F4")
    InputRow ws, 5, "Ayarlar", SettingsPath, "Okundu", ""
    InputRow ws, 6, "Kontrol", AbsPath(CStr(Opt("Dosyalar", "kontrol_dosyasi"))), "Okundu", ""
    InputRow ws, 7, "Endeks", AbsPath(CStr(Opt("Dosyalar", "endeks_dosyasi"))), "Okundu", ""
    folder = AbsPath(CStr(Opt("Dosyalar", "maliyet_klasoru")))
    Set names = New Collection
    name = Dir$(folder & "\*.xlsx")
    Do While name <> ""
        If Left$(name, 2) <> "~$" Then names.Add name
        name = Dir$()
    Loop
    r = 8
    If names.Count > 0 Then
        ReDim list(1 To names.Count)
        For i = 1 To names.Count
            list(i) = names(i)
        Next i
        For i = 2 To names.Count
            t = list(i): j = i - 1
            Do While j >= 1
                If StrComp(list(j), t, vbBinaryCompare) <= 0 Then Exit Do
                list(j + 1) = list(j): j = j - 1
            Loop
            list(j + 1) = t
        Next i
        For i = 1 To names.Count
            period = FilenamePeriod(folder & "\" & list(i))
            If period = "" Then
                InputRow ws, r, "Maliyet", folder & "\" & list(i), "Dönem bulunamadý", ""
            ElseIf InClosingScope(period) Then
                InputRow ws, r, "Maliyet", folder & "\" & list(i), "Okundu", period
            Else
                InputRow ws, r, "Maliyet", folder & "\" & list(i), "Kapsam dýþý", period
            End If
            r = r + 1
        Next i
    End If
    ws.Range("D5:D" & r).NumberFormat = "#,##0.0"
    ws.Columns("A").ColumnWidth = 12: ws.Columns("B").ColumnWidth = 40: ws.Columns("C:D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 20: ws.Columns("F").ColumnWidth = 18
End Sub

' Yönetici Özeti'ne eklenen iki satýrýn deðerleri (Python surum.facts)
Public Function ChangeFact(ByVal scope As Object) As String
    Dim row As Variant, effect As Variant, rows As Collection
    If Not HasPrevious(scope) Then ChangeFact = "Karþýlaþtýrma yok (Ocak kapanýþý veya önceki ay dosyasý yok)": Exit Function
    Set rows = ChangeItems(scope): effect = CDec(0)
    For Each row In rows
        If Not IsNull(row(7)) Then effect = effect + row(7)
    Next row
    ChangeFact = rows.Count & " kalemde deðiþti, yýllýk toplam etki " & Thousands(effect) & " TL"
End Function

Public Function TrendFact(ByVal scope As Object) As String
    If ReportYear <> CloseYear Then TrendFact = "Bütçe yýlý dýþýnda": Exit Function
    TrendFact = TrendItems(scope).Count & " kalem (" & Ge2() & " " & TrendLeast() & " ay, aylýk " & Ge2() & " " & Thousands(TrendThreshold()) & " TL)"
End Function
