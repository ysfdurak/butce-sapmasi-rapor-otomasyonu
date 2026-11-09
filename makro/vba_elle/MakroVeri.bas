Attribute VB_Name = "MakroVeri"
Option Explicit

Public Sub LoadSettings()
    Dim wb As Workbook, ws As Worksheet, r As Long, last As Long, n As String, id As String, meta As Variant
    Dim i As Long, lo As Long, hi As Long, g As Variant, target As String, groupName As String
    Dim errText As String
    On Error GoTo Bad
    Set Cfg = Dict(): Set Depts = Dict(): Set Groups = New Collection
    Set ExpenseMap = Dict(): Set Excluded = Dict()
    Set wb = ReadBook(SettingsPath)
    Set ws = wb.Worksheets("Dosya ve Konum Ayarlarý")
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To last
        If Application.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))) > 0 Then
            For i = 1 To 3
                If ws.Cells(r, i).HasFormula Or IsError(ws.Cells(r, i).Value) Or IsEmpty(ws.Cells(r, i).Value) Then Fail "Ayar satýrý eksik veya formüllü: " & r
            Next i
            id = CStr(ws.Cells(r, 1).Value2) & Sep & CStr(ws.Cells(r, 2).Value2)
            If Cfg.Exists(id) Then Fail "Mükerrer ayar: " & id
            Cfg.Add id, ws.Cells(r, 3).Value2
        End If
    Next r
    Set ws = wb.Worksheets("Ortak Sayfalar")
    For r = 5 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        n = Trim$(CStr(ws.Cells(r, 1).Value2))
        If n <> "" Then Excluded(K(n)) = True
    Next r
    Set ws = wb.Worksheets("Bölüm Eþleþtirme")
    For r = 5 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Application.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, 7))) > 0 Then
            For i = 1 To 7
                If ws.Cells(r, i).HasFormula Or IsError(ws.Cells(r, i).Value2) Then Fail "Bölüm eþleþtirmede sabit deðer kullanýn: " & r
                If i <= 3 And Len(Trim$(CStr(ws.Cells(r, i).Value2))) = 0 Then Fail "Bölüm/müdürlük/direktörlük eksik: " & r
            Next i
            n = K(ws.Cells(r, 1).Value2)
            If InStr(n, Sep) > 0 Or Depts.Exists(n) Or Excluded.Exists(n) Then Fail "Mükerrer, dýþlanan veya geçersiz bölüm: " & n
            meta = Array(Trim$(CStr(ws.Cells(r, 1).Value2)), Trim$(CStr(ws.Cells(r, 2).Value2)), Trim$(CStr(ws.Cells(r, 3).Value2)), Trim$(CStr(ws.Cells(r, 4).Value2)), Trim$(CStr(ws.Cells(r, 5).Value2)), Trim$(CStr(ws.Cells(r, 6).Value2)), Trim$(CStr(ws.Cells(r, 7).Value2)))
            If meta(5) = "" Then meta(5) = meta(0)
            Depts.Add n, meta
        End If
    Next r
    If Depts.Count = 0 Then Fail "En az bir bölüm eþleþtirmesi gerekiyor."
    Set ws = wb.Worksheets("Gider Gruplarý")
    For r = 5 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Application.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))) > 0 Then
            For i = 1 To 3
                If ws.Cells(r, i).HasFormula Or IsError(ws.Cells(r, i).Value2) Then Fail "Gider gruplarýnda sabit deðer kullanýn: " & r
            Next i
            lo = CLng(ws.Cells(r, 1).Value2): hi = CLng(ws.Cells(r, 2).Value2)
            If lo < 0 Or hi < lo Or lo <> ws.Cells(r, 1).Value2 Or hi <> ws.Cells(r, 2).Value2 Then Fail "Geçersiz gider aralýðý: " & r
            For Each g In Groups
                If lo <= g(1) And g(0) <= hi Then Fail "Gider grubu aralýklarý çakýþýyor: " & r
            Next g
            Groups.Add Array(lo, hi, CStr(ws.Cells(r, 3).Value2))
        End If
    Next r
    Set ws = wb.Worksheets("Gider Eþleþtirme")
    For r = 5 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Application.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))) > 0 Then
            For i = 1 To 3
                If ws.Cells(r, i).HasFormula Or IsError(ws.Cells(r, i).Value2) Then Fail "Gider eþleþtirmede sabit deðer kullanýn: " & r
            Next i
            id = CStr(CLng(ws.Cells(r, 1).Value2)): target = CStr(CLng(ws.Cells(r, 2).Value2))
            If ExpenseMap.Exists(id) Then Fail "Mükerrer kaynak gider: " & id
            groupName = CStr(ws.Cells(r, 3).Value2)
            If groupName = "" Then Fail "Hedef gider adý eksik: " & r
            ExpenseMap.Add id, Array(target, groupName)
        End If
    Next r
    wb.Close False: Set wb = Nothing
    For Each meta In Array("Dosyalar|maliyet_klasoru", "Dosyalar|kontrol_dosyasi", "Dosyalar|endeks_dosyasi", "Maliyet|ilk_gider_satiri", "Maliyet|gider_sutunu", "Maliyet|islem_adedi_satiri", "Kontrol|sayfa", "Endeks|sayfa")
        If Not Cfg.Exists(CStr(meta)) Then Fail "Eksik ayar: " & CStr(meta)
    Next meta
    For Each meta In Cfg.Keys
        n = CStr(meta)
        If Left$(n, 8) = "Maliyet|" Or Left$(n, 8) = "Kontrol|" Or Left$(n, 7) = "Endeks|" Then
            id = Split(n, Sep)(1)
            If InStr(id, "satiri") > 0 Or InStr(id, "sutunu") > 0 Or id = "gider_kalem_sayisi" Then
                If Not IsNumeric(Cfg(n)) Then Fail "Konum pozitif tam sayý olmalý: " & n
                If CDbl(Cfg(n)) < 1 Or CDbl(Cfg(n)) <> Fix(CDbl(Cfg(n))) Then Fail "Konum pozitif tam sayý olmalý: " & n
            End If
        End If
    Next meta
    If Not IsNumeric(Opt("Kontrol", "tolerans", 0.01)) Then Fail "Kontrol toleransý sayý olmalý."
    If CDbl(Opt("Kontrol", "tolerans", 0.01)) < 0 Then Fail "Kontrol toleransý negatif olamaz."
    If Groups.Count = 0 Then Fail "En az bir gider grubu tanýmlayýn."
    RootPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(SettingsPath)
    Exit Sub
Bad:
    errText = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Fail errText
End Sub

Public Sub CostColumns(ByVal ws As Worksheet, ByVal closing As String, ByRef actual As Object, ByRef budget As Object)
    Dim marker As Range, c As Long, endCol As Long, firstMarker As Long, b As Object, valid As Long
    Dim area As Range, v As Variant, pkey As String, block As Variant, y As Long, m As Long
    If CStr(Opt("Maliyet", "tarih_baslangici")) <> "" Then
        Set marker = ws.Range(CStr(Opt("Maliyet", "butce_etiketi")))
        endCol = ws.Cells(marker.Row, ws.Columns.Count).End(xlToLeft).Column
        For c = marker.Column To endCol
            If K(ws.Cells(marker.Row, c).Value2) = "butce" Then
                If firstMarker = 0 Then firstMarker = c
                Set b = ReadCols(ws, ws.Cells(marker.Row + 1, c).Address, c + 11, CloseYear)
                If FullBudget(b) Then valid = valid + 1: Set budget = b
            End If
        Next c
        If valid <> 1 Or firstMarker = 0 Then Fail "Tek 12 aylýk kapanýþ yýlý bütçe bloðu bulunamadý: " & ws.Name
        Set actual = ReadCols(ws, CStr(Opt("Maliyet", "tarih_baslangici")), firstMarker - 1)
    Else
        Set actual = Dict(): Set budget = Dict()
        For Each block In Array("gerceklesen_araligi", "butce_araligi")
            Set area = ws.Range(CStr(Opt("Maliyet", CStr(block))))
            For c = area.Column To area.Column + area.Columns.Count - 1
                v = ws.Cells(CLng(Opt("Maliyet", "yil_satiri")), c).Value2
                pkey = PeriodOf(v)
                If pkey = "" Then
                    y = CLng(v): m = MonthNum(ws.Cells(CLng(Opt("Maliyet", "ay_satiri")), c).Value2)
                    If m = 0 Then Fail "Ay baþlýðý geçersiz: " & ws.Name
                    pkey = P(y, m)
                End If
                If block = "gerceklesen_araligi" Then
                    If actual.Exists(pkey) Then Fail "Mükerrer gerçekleþen ayý: " & ws.Name
                    actual.Add pkey, c
                Else
                    If budget.Exists(pkey) Then Fail "Mükerrer bütçe ayý: " & ws.Name
                    budget.Add pkey, c
                End If
            Next c
        Next block
        If Not FullBudget(budget) Then Fail "Bütçe yýlý/aylarý eksik: " & ws.Name
    End If
    If Not actual.Exists(closing) Then Fail "Kaynak dosya kapanýþý gerçekleþen baþlýklarýnda yok: " & ws.Name
End Sub

Private Function ReadCost(ByVal ws As Worksheet, ByVal closing As String, ByVal source As String) As Object
    Dim d As Object, names As Object, actual As Object, budget As Object, raw As Object
    Dim c As Variant, id As Variant, code As Long, target As String, item As String, mapping As Variant
    Dim r As Long, first As Long, cnt As Long, slot As Variant, cols As Object, x As String, value As Variant
    Dim indicatorNames As Object, indicatorRow As Long, labelColumn As Long, indicator As Long, title As String
    Set d = Dict(): Set names = Dict(): Set raw = Dict()
    CostColumns ws, closing, actual, budget
    Set indicatorNames = Dict()
    indicatorRow = CLng(Opt("Maliyet", "islem_adedi_satiri"))
    labelColumn = CLng(Opt("Maliyet", "gosterge_adi_sutunu", Opt("Maliyet", "gider_sutunu")))
    For indicator = 1 To 5
        r = indicatorRow - 5 + indicator
        title = SourceIndicatorTitle(ws, indicator)
        If InStr(1, title, "Baþlýðý eksik", vbTextCompare) = 1 Then Warn ws.Name, source, title
        indicatorNames.Add CStr(indicator), title
        For Each slot In Array("A", "B")
            If slot = "A" Then Set cols = actual Else Set cols = budget
            For Each id In cols.Keys
                d(slot & "P" & indicator & Sep & id) = Num(ws.Cells(r, cols(id)).Value2)
            Next id
        Next slot
    Next indicator
    d.Add "IndicatorNames", indicatorNames
    first = CLng(Opt("Maliyet", "ilk_gider_satiri")): cnt = CLng(Opt("Maliyet", "gider_kalem_sayisi", CLng(Opt("Maliyet", "dip_toplam_satiri")) - first))
    If cnt < 1 Then Fail "Gider kalem sayýsý sýfýr olamaz."
    If CLng(Opt("Maliyet", "islem_adedi_satiri")) >= first Then Fail "Personel/iþlem adetleri gider aralýðýnda olamaz."
    For code = 1 To cnt
        r = first + code - 1: target = CStr(code)
        item = Trim$(CStr(ws.Cells(r, CLng(Opt("Maliyet", "gider_sutunu"))).Value2))
        If item = "" Then Fail "Gider adý eksik: " & ws.Name & " / " & r
        If ExpenseMap.Exists(CStr(code)) Then
            mapping = ExpenseMap(CStr(code)): target = mapping(0): item = mapping(1)
        End If
        If names.Exists(target) Then
            If names(target) <> item Then Fail "Birleþtirilen gider adlarý farklý: " & target
        Else
            names.Add target, item
        End If
        For Each slot In Array("A", "B")
            If slot = "A" Then Set cols = actual Else Set cols = budget
            For Each id In cols.Keys
                x = slot & Sep & id & Sep & target: value = Num(ws.Cells(r, cols(id)).Value2)
                If d.Exists(x) Then value = Plus(d(x), value)
                d(x) = value
                x = slot & "T" & Sep & id
                value = Num(ws.Cells(r, cols(id)).Value2)
                If d.Exists(x) Then value = Plus(d(x), value)
                d(x) = value
            Next id
        Next slot
    Next code
    For Each slot In Array("A", "B")
        If slot = "A" Then Set cols = actual Else Set cols = budget
        For Each id In cols.Keys
            d(slot & "C" & Sep & id) = Num(ws.Cells(CLng(Opt("Maliyet", "islem_adedi_satiri")), cols(id)).Value2)
        Next id
    Next slot
    d.Add "names", names: d("file") = source
    Set ReadCost = d
End Function

Public Sub LoadSnapshots()
    Dim files As Collection, file As Variant, name As String, folder As String, period As String
    Dim wb As Workbook, ws As Worksheet, d As Object, id As String, bad As Object, msg As String, n As Variant
    Set Snaps = Dict(): Set bad = Dict(): Set files = New Collection
    folder = AbsPath(CStr(Opt("Dosyalar", "maliyet_klasoru")))
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(folder) Then Fail "Maliyet klasörü yok: " & folder
    name = Dir$(folder & "\*.xlsx")
    Do While name <> ""
        If Left$(name, 2) <> "~$" Then files.Add folder & "\" & name
        name = Dir$()
    Loop
    For Each file In files
        WorkGate "Kaynak dosyasý: " & CStr(file)
        period = FilenamePeriod(CStr(file))
        If period = "" Then
            Warn "", CStr(file), "Dosya adýnda tek yýl ve ay bulunamadý; okunmadý."
        ElseIf Left$(period, 4) = CStr(CloseYear) And period <= P(CloseYear, CloseMonth) Then
            Application.StatusBar = "Kaynak okunuyor: " & file
            ReadSnapshotFile CStr(file), period, bad
        End If
    Next file
    For Each n In Depts.Keys
        Set d = CurrentSnap(CStr(n))
        If d Is Nothing Then Warn CStr(Depts(n)(0)), "", "Kapanýþ kaynaðý eksik/okunamadý; toplamlar eksikse boþ býrakýlýr."
    Next n
End Sub

Private Sub ReadSnapshotFile(ByVal file As String, ByVal closing As String, ByVal bad As Object)
    Dim wb As Workbook, ws As Worksheet, d As Object, id As String, msg As String, n As Variant
    On Error GoTo Failed
    Set wb = ReadBook(file)
    For Each ws In wb.Worksheets
        If Not Excluded.Exists(K(ws.Name)) Then
            If Not Depts.Exists(K(ws.Name)) Then
                Warn ws.Name, file, "Bölüm eþleþmedi; rapora alýnmadý."
            Else
                id = K(ws.Name) & Sep & closing
                If bad.Exists(id) Or Snaps.Exists(id) Then
                    If Snaps.Exists(id) Then Snaps.Remove id
                    bad(id) = True
                    Warn ws.Name, file, "Ayný bölüm/kapanýþ için birden fazla kaynak var; ilgili bölüm kullanýlmadý."
                Else
                    Set d = TryCost(ws, closing, file)
                    If d Is Nothing Then bad(id) = True Else Snaps.Add id, d
                End If
            End If
        End If
    Next ws
    wb.Close False
    Exit Sub
Failed:
    msg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    For Each n In Depts.Keys
        id = CStr(n) & Sep & closing
        If Snaps.Exists(id) Then
            Set d = Snaps(id)
            If d("file") = file Then Snaps.Remove id
        End If
    Next n
    Warn "", file, "Dosya okunamadý: " & msg
End Sub

Private Function TryCost(ByVal ws As Worksheet, ByVal closing As String, ByVal file As String) As Object
    On Error GoTo Failed
    Set TryCost = ReadCost(ws, closing, file)
    Exit Function
Failed:
    Warn ws.Name, file, "Kaynak okunamadý: " & Err.Description
End Function

Public Sub LoadControl()
    Dim wb As Workbook, ws As Worksheet, actual As Object, b As Object, planned As Object, directors As Object
    Dim start As Range, c As Long, lastCol As Long, count As Long, r As Long, krow As Long, id As Variant
    Dim header As Long, titleCol As Long, n As Variant, meta As Variant, director As String, name As String
    Dim d As Object, bucket As Collection, x As String, errText As String, maxActual As Long
    On Error GoTo Bad
    Set Revenue = Dict(): Set RevenueBudget = Dict(): Set Controls = Dict(): Set directors = Dict()
    For Each n In Depts.Keys
        meta = Depts(n): directors(K(meta(2))) = True
    Next n
    Set wb = ReadBook(AbsPath(CStr(Opt("Dosyalar", "kontrol_dosyasi"))))
    Set ws = wb.Worksheets(CStr(Opt("Kontrol", "sayfa")))
    Set start = ws.Range(CStr(Opt("Kontrol", "ciro_baslangici", "D18")))
    header = start.Row
    Set actual = ReadCols(ws, start.Address)
    If actual.Count = 0 Then Fail "Ciro dönemleri bulunamadý."
    For Each id In actual.Keys
        Revenue(id) = Num(ws.Cells(header + 1, actual(id)).Value2)
        If actual(id) > maxActual Then maxActual = actual(id)
    Next id
    c = ws.Range(CStr(Opt("Kontrol", "butce_ciro_baslangici", "CW18"))).Column
    If c <= maxActual Then c = maxActual + 1
    lastCol = ws.Cells(header, ws.Columns.Count).End(xlToLeft).Column
    Do While c <= lastCol - 11
        Set b = ReadCols(ws, ws.Cells(header, c).Address, c + 11, CloseYear)
        If FullBudget(b) Then count = count + 1: Set planned = b
        c = c + 1
    Loop
    If count <> 1 Then Fail "Ciroda tek 12 aylýk kapanýþ yýlý bütçe bloðu bulunamadý."
    For Each id In planned.Keys
        RevenueBudget(id) = Num(ws.Cells(header + 1, planned(id)).Value2)
    Next id
    titleCol = CLng(Opt("Kontrol", "etiket_sutunu", 3))
    For r = header + 2 To ws.Cells(ws.Rows.Count, titleCol).End(xlUp).Row
        director = K(ws.Cells(r, titleCol).Value2)
        If directors.Exists(director) And K(ws.Cells(r + 2, titleCol).Value2) = "opex toplami" Then
            krow = r + 3
            Do While krow <= ws.Cells(ws.Rows.Count, titleCol).End(xlUp).Row
                name = K(ws.Cells(krow, titleCol).Value2)
                If name = "" Or directors.Exists(name) Then Exit Do
                x = director & Sep & name
                If Not Controls.Exists(x) Then
                    Set bucket = New Collection: Controls.Add x, bucket
                End If
                Set bucket = Controls(x): Set d = Dict()
                For Each id In actual.Keys
                    d(id) = Num(ws.Cells(krow, actual(id)).Value2)
                Next id
                bucket.Add d
                krow = krow + 1
            Loop
        End If
    Next r
    wb.Close False
    Exit Sub
Bad:
    errText = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Fail "Kontrol/ciro dosyasý: " & errText
End Sub

Public Sub LoadIndices()
    Dim wb As Workbook, ws As Worksheet, top As Range, cols As Object, id As Variant, n As String
    Dim r As Long, c As Long, prev As String, expected As String, seen As Object, errText As String
    On Error GoTo Bad
    Set Indices = Dict(): Set seen = Dict()
    Set wb = ReadBook(AbsPath(CStr(Opt("Dosyalar", "endeks_dosyasi"))))
    Set ws = wb.Worksheets(CStr(Opt("Endeks", "sayfa")))
    Set top = ws.Range(CStr(Opt("Endeks", "endeks_araligi", "C" & CStr(Opt("Endeks", "baslik_satiri", 17))))).Cells(1, 1)
    If K(top.Value2) <> "bolum gider endeksi" And K(top.Value2) <> "bolum adi" Then Fail "Ýkinci endeks tablosunun baþlýðý bulunamadý: " & top.Address
    Set cols = ReadCols(ws, top.Offset(0, 1).Address)
    If cols.Count = 0 Then Fail "Endeks tarihleri yok."
    For Each id In cols.Keys
        If prev <> "" Then
            expected = Format$(DateAdd("m", 1, DateSerial(CLng(Left$(prev, 4)), CLng(Right$(prev, 2)), 1)), "yyyy-mm")
            If CStr(id) <> expected Then Fail "Endeks aylarý ardýþýk deðil: " & id
        End If
        prev = CStr(id)
    Next id
    For r = top.Row + 1 To ws.Cells(ws.Rows.Count, top.Column).End(xlUp).Row
        n = K(ws.Cells(r, top.Column).Value2)
        If n <> "" Then
            If seen.Exists(n) Then Fail "Mükerrer endeks bölümü: " & n
            seen.Add n, True
            For Each id In cols.Keys
                Indices(n & Sep & id) = Num(ws.Cells(r, cols(id)).Value2)
            Next id
        End If
    Next r
    wb.Close False
    Exit Sub
Bad:
    errText = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Fail "Endeks dosyasý: " & errText
End Sub
