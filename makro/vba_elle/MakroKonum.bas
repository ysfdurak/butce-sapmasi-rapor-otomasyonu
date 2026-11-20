Attribute VB_Name = "MakroKonum"
Option Explicit

' Kaynak dosyalarda konumlarýn otomatik bulunmasý (V5.5). Python: konum.py, Power Query: Butce_Veri.pq.
' Ayarlar'da konum satýrý varsa o kullanýlýr. Yoksa dosyadan bulunur ve Uyarýlar'a yazýlýr.
' Ayardaki konum bir maliyet sayfasýnda tutmazsa o sayfa için yeniden bulunur (MakroVeri.TryCost).
' Maliyet: "Bütçe" yazan hücrenin altý tarih/ay ise o satýr baþlýktýr; solundaki ilk tarih gerçekleþen
'          baþlangýcýdýr. Daha solda altýnda en az üç yazý olan en yakýn sütun gider adýdýr. "Ýþlem Adedi"
'          satýrýnýn bir altý ilk giderdir; giderler ilk boþ ada kadar sayýlýr.
' Kontrol: En az 12 tarihli ve altýnda sayý olan en üst satýr ciro baþlýðýdýr; tarihler kesildikten sonraki
'          ilk tarih bütçe cirosudur. "Opex Toplamý" yazan sütun etiket sütunudur.
' Endeks:  "Bölüm Gider Endeksi" veya "Bölüm Adý" yazan, saðýnda tarih olan hücre tablonun sol üstüdür.

Public AutoLayout As Collection   ' Array(alan, kaynak, açýklama)

Private Const MaxRows As Long = 400
Private Const MaxCols As Long = 1500

Private Function Grid(ByVal ws As Worksheet, ByRef nRows As Long, ByRef nCols As Long) As Variant
    Dim used As Range, v As Variant
    Set used = ws.UsedRange
    nRows = used.Row + used.Rows.Count - 1: nCols = used.Column + used.Columns.Count - 1
    If nRows > MaxRows Then nRows = MaxRows
    If nCols > MaxCols Then nCols = MaxCols
    If nRows < 2 Then nRows = 2
    If nCols < 2 Then nCols = 2
    Grid = ws.Range(ws.Cells(1, 1), ws.Cells(nRows, nCols)).Value
End Function

Public Function IsDateCell(ByVal v As Variant) As Boolean
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then IsDateCell = True: Exit Function
    If VarType(v) <> vbString Then Exit Function
    s = Trim$(v)
    IsDateCell = (s Like "####-#" Or s Like "####-##" Or s Like "####-##-##" Or s Like "#.#.####" Or s Like "##.#.####" Or s Like "#.##.####" Or s Like "##.##.####")
End Function

Private Function IsMonthCell(ByVal v As Variant) As Boolean
    If IsDateCell(v) Then IsMonthCell = True: Exit Function
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        IsMonthCell = (MonthNum(v) > 0 And Not IsNumeric(v))
    ElseIf IsNumeric(v) Then
        IsMonthCell = (CDbl(v) >= 1 And CDbl(v) <= 12 And CDbl(v) = Fix(CDbl(v)))
    End If
End Function

Private Function IsText(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then IsText = (Trim$(v) <> "")
End Function

Private Function IsAmount(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Or VarType(v) = vbDate Or VarType(v) = vbBoolean Then Exit Function
    IsAmount = IsNumeric(v)
End Function

Private Function CellRef(ByVal col As Long, ByVal row As Long) As String
    Dim letters As String, n As Long
    n = col
    Do While n > 0
        letters = Chr$(65 + (n - 1) Mod 26) & letters
        n = (n - 1) \ 26
    Loop
    CellRef = letters & row
End Function

Private Function ColumnOf(ByVal ref As String) As Long
    Dim i As Long, ch As String
    ref = UCase$(Replace(ref, "$", ""))
    For i = 1 To Len(ref)
        ch = Mid$(ref, i, 1)
        If ch < "A" Or ch > "Z" Then Exit For
        ColumnOf = ColumnOf * 26 + Asc(ch) - 64
    Next i
End Function

Private Function Safe(ByVal g As Variant, ByVal r As Long, ByVal c As Long, ByVal nRows As Long, ByVal nCols As Long) As Variant
    If r < 1 Or c < 1 Or r > nRows Or c > nCols Then Safe = Empty Else Safe = g(r, c)
End Function

Public Function CostLayout(ByVal ws As Worksheet) As Object
    Dim g As Variant, nRows As Long, nCols As Long, r As Long, c As Long, found As Object
    Dim markerRow As Long, markerCol As Long, header As Long, start As Long, nameCol As Long, texts As Long, x As Long
    Dim countRow As Long, size As Long, label As String
    Set found = Dict(): Set CostLayout = found
    g = Grid(ws, nRows, nCols)
    For r = 1 To nRows - 1
        For c = 1 To nCols
            If K(g(r, c)) = "butce" Then
                If IsMonthCell(g(r + 1, c)) Then
                    If markerRow = 0 Then
                        markerRow = r: markerCol = c
                    ElseIf r <> markerRow Then
                        found("neden") = "farklý satýrlarda ""Bütçe"" yazýyor": Exit Function
                    End If
                End If
            End If
        Next c
    Next r
    If markerRow = 0 Then found("neden") = "altý tarih/ay olan ""Bütçe"" hücresi yok": Exit Function
    header = markerRow + 1
    For c = 1 To markerCol - 1
        If IsDateCell(g(header, c)) Then start = c: Exit For
    Next c
    If start = 0 Then found("neden") = "Bütçe baþlýk satýrýnda solda tarih yok (" & CellRef(markerCol, header) & ")": Exit Function
    found("tarih_baslangici") = CellRef(start, header)
    found("butce_etiketi") = CellRef(markerCol, markerRow)
    For c = start - 1 To 1 Step -1
        texts = 0
        For x = header + 1 To header + 59
            If IsText(Safe(g, x, c, nRows, nCols)) Then texts = texts + 1
        Next x
        If texts >= 3 Then nameCol = c: Exit For
    Next c
    If nameCol = 0 Then found("neden") = "tarihlerin solunda gider adý sütunu yok": Exit Function
    found("gider_sutunu") = nameCol
    For x = header + 1 To header + 59
        label = K(Safe(g, x, nameCol, nRows, nCols))
        If InStr(label, "islem aded") > 0 Then countRow = x: Exit For
    Next x
    If countRow = 0 Then found("neden") = """Ýþlem Adedi"" satýrý yok": Exit Function
    found("islem_adedi_satiri") = countRow
    found("ilk_gider_satiri") = countRow + 1
    Do While IsText(Safe(g, countRow + 1 + size, nameCol, nRows, nCols))
        size = size + 1
    Loop
    If size > 0 Then found("gider_kalem_sayisi") = size
End Function

Public Function ControlLayout(ByVal ws As Worksheet) As Object
    Dim g As Variant, nRows As Long, nCols As Long, r As Long, c As Long, found As Object, n As Long, first As Long
    Dim header As Long, previous As Long, labelCol As Long, labels As Long
    Set found = Dict(): Set ControlLayout = found
    g = Grid(ws, nRows, nCols)
    For r = 1 To nRows - 1
        n = 0: first = 0
        For c = 1 To nCols
            If IsDateCell(g(r, c)) Then
                n = n + 1
                If first = 0 Then first = c
            End If
        Next c
        If n >= 12 Then
            If IsAmount(g(r + 1, first)) Then header = r: Exit For
        End If
    Next r
    If header = 0 Then Exit Function
    found("ciro_baslangici") = CellRef(first, header)
    previous = first
    For c = first + 1 To nCols
        If IsDateCell(g(header, c)) Then
            If c <> previous + 1 Then found("butce_ciro_baslangici") = CellRef(c, header): Exit For
            previous = c
        End If
    Next c
    For r = header + 2 To nRows
        For c = 1 To nCols
            If K(g(r, c)) = "opex toplami" Then
                If labelCol <> c Then labels = labels + 1
                labelCol = c
            End If
        Next c
    Next r
    If labels = 1 Then found("etiket_sutunu") = labelCol
End Function

Public Function IndexLayout(ByVal ws As Worksheet) As Object
    Dim g As Variant, nRows As Long, nCols As Long, r As Long, c As Long, found As Object, hits As Long, corner As String
    Set found = Dict(): Set IndexLayout = found
    g = Grid(ws, nRows, nCols)
    For r = 1 To nRows
        For c = 1 To nCols - 1
            If K(g(r, c)) = "bolum adi" Or K(g(r, c)) = "bolum gider endeksi" Then
                If IsDateCell(g(r, c + 1)) Then hits = hits + 1: corner = CellRef(c, r)
            End If
        Next c
    Next r
    If hits = 1 Then found("endeks_araligi") = corner
End Function

Public Function Describe(ByVal found As Object, Optional ByVal only As Object = Nothing) As String
    Dim k2 As Variant
    For Each k2 In found.Keys
        If only Is Nothing Then
            Describe = Describe & IIf(Describe = "", "", ", ") & k2 & " = " & found(k2)
        ElseIf only.Exists(k2) Then
            Describe = Describe & IIf(Describe = "", "", ", ") & k2 & " = " & found(k2)
        End If
    Next k2
End Function

Private Function Has(ByVal section As String, ByVal code As String) As Boolean
    If Cfg.Exists(section & Sep & code) Then Has = (CStr(Cfg(section & Sep & code)) <> "")
End Function

Private Function LegacyMode(ByVal section As String, ByVal positioned As String, ByVal legacy As Variant) As Boolean
    Dim code As Variant
    If Has(section, positioned) Then Exit Function
    For Each code In legacy
        If Has(section, CStr(code)) Then LegacyMode = True: Exit Function
    Next code
End Function

Private Function Fill(ByVal section As String, ByVal keys As Variant, ByVal found As Object) As Object
    ' Yalnýz Ayarlar'da olmayan anahtarlar eklenir; eklenenler döner.
    Dim code As Variant, added As Object
    Set added = Dict()
    For Each code In keys
        If Not Has(section, CStr(code)) And found.Exists(CStr(code)) Then
            Cfg(section & Sep & CStr(code)) = found(CStr(code)): added(CStr(code)) = found(CStr(code))
        End If
    Next code
    Set Fill = added
End Function

' MakroVeri.LoadSettings sonunda, zorunlu ayar denetiminden önce çaðrýlýr.
Public Sub KonumTamamla()
    Dim costKeys As Variant, folder As String, name As String, best As String, bestPeriod As String, period As String
    Dim wb As Workbook, ws As Worksheet, found As Object, added As Object, path As String, sheetName As String
    Dim lacking As String, code As Variant, costNote As String
    Set AutoLayout = New Collection
    costKeys = Array("tarih_baslangici", "butce_etiketi", "gider_sutunu", "islem_adedi_satiri", "ilk_gider_satiri", "gider_kalem_sayisi")
    If Not LegacyMode("Maliyet", "tarih_baslangici", Array("yil_satiri", "ay_satiri", "dip_toplam_satiri", "indeks_sutunu", "ilk_donem_sutunu", "gerceklesen_araligi", "butce_araligi")) Then
        If Not (Has("Maliyet", "tarih_baslangici") And Has("Maliyet", "butce_etiketi") And Has("Maliyet", "gider_sutunu") And Has("Maliyet", "islem_adedi_satiri") And Has("Maliyet", "ilk_gider_satiri") And Has("Maliyet", "gider_kalem_sayisi")) Then
            ' En yeni dönemli maliyet dosyasýnýn ilk bölüm sayfasýndan bulunur.
            folder = AbsPath(CStr(Opt("Dosyalar", "maliyet_klasoru")))
            If CreateObject("Scripting.FileSystemObject").FolderExists(folder) Then
                name = Dir$(folder & "\*.xlsx")
                Do While name <> ""
                    If Left$(name, 2) <> "~$" Then
                        period = FilenamePeriod(folder & "\" & name)
                        If period & "|" & name > bestPeriod & "|" & best Then bestPeriod = period: best = name
                    End If
                    name = Dir$()
                Loop
            End If
            If best <> "" Then
                costNote = best & ": bölüm sayfasý yok"
                Set wb = ReadBook(folder & "\" & best)
                For Each ws In wb.Worksheets
                    If Depts.Exists(K(ws.Name)) And Not Excluded.Exists(K(ws.Name)) Then
                        Set found = CostLayout(ws): sheetName = ws.Name
                        costNote = best & " > " & sheetName
                        If found.Exists("neden") Then costNote = costNote & ": " & found("neden"): found.Remove "neden"
                        Exit For
                    End If
                Next ws
                wb.Close False
                If Not found Is Nothing Then
                    If found.Exists("tarih_baslangici") Then
                        Set added = Fill("Maliyet", costKeys, found)
                        If added.Count > 0 Then AutoLayout.Add Array("Maliyet", best & " > " & sheetName, Describe(added))
                    End If
                End If
            ElseIf folder <> "" Then
                costNote = "maliyet klasöründe .xlsx yok: " & folder
            End If
        End If
        For Each code In Array("tarih_baslangici", "butce_etiketi", "islem_adedi_satiri", "ilk_gider_satiri")
            If Not Has("Maliyet", CStr(code)) Then lacking = lacking & IIf(lacking = "", "", ", ") & "Maliyet > " & code
        Next code
        If Not Has("Maliyet", "gider_sutunu") And Has("Maliyet", "tarih_baslangici") Then
            Cfg("Maliyet" & Sep & "gider_sutunu") = ColumnOf(CStr(Cfg("Maliyet|tarih_baslangici"))) - 1
        End If
    End If
    If Not LegacyMode("Kontrol", "ciro_baslangici", Array("ciro_satiri", "ilk_blok_satiri", "ilk_ay_sutunu", "ciro_araligi")) Then
        If Not (Has("Kontrol", "ciro_baslangici") And Has("Kontrol", "butce_ciro_baslangici") And Has("Kontrol", "etiket_sutunu")) Then
            path = AbsPath(CStr(Opt("Dosyalar", "kontrol_dosyasi")))
            If Len(Dir$(path)) > 0 Then
                Set wb = ReadBook(path)
                Set ws = Nothing
                On Error Resume Next
                Set ws = wb.Worksheets(CStr(Opt("Kontrol", "sayfa")))
                On Error GoTo 0
                If ws Is Nothing Then Set ws = wb.Worksheets(1)
                Set found = ControlLayout(ws)
                wb.Close False
                Set added = Fill("Kontrol", Array("ciro_baslangici", "butce_ciro_baslangici", "etiket_sutunu"), found)
                If added.Count > 0 Then AutoLayout.Add Array("Kontrol", CreateObject("Scripting.FileSystemObject").GetFileName(path), Describe(added))
            End If
        End If
        For Each code In Array("ciro_baslangici", "butce_ciro_baslangici", "etiket_sutunu")
            If Not Has("Kontrol", CStr(code)) Then lacking = lacking & IIf(lacking = "", "", ", ") & "Kontrol > " & code
        Next code
    End If
    If Not Has("Endeks", "baslik_satiri") And Not Has("Endeks", "endeks_araligi") And Not Has("Endeks", "ilk_veri_satiri") Then
        path = AbsPath(CStr(Opt("Dosyalar", "endeks_dosyasi")))
        If Len(Dir$(path)) > 0 Then
            Set wb = ReadBook(path)
            Set ws = Nothing
            On Error Resume Next
            Set ws = wb.Worksheets(CStr(Opt("Endeks", "sayfa")))
            On Error GoTo 0
            If ws Is Nothing Then Set ws = wb.Worksheets(1)
            Set found = IndexLayout(ws)
            wb.Close False
            Set added = Fill("Endeks", Array("endeks_araligi"), found)
            If added.Count > 0 Then AutoLayout.Add Array("Endeks", CreateObject("Scripting.FileSystemObject").GetFileName(path), Describe(added))
        End If
        If Not Has("Endeks", "endeks_araligi") Then lacking = lacking & IIf(lacking = "", "", ", ") & "Endeks > endeks_araligi"
    End If
    If lacking <> "" Then
        If costNote <> "" And InStr(lacking, "Maliyet >") > 0 Then lacking = lacking & " (" & costNote & ")"
        Fail "Konum dosyadan bulunamadý; Ayarlar > Dosya ve Konum Ayarlarý'na girin: " & lacking
    End If
End Sub

' LoadSnapshots baþýnda: dosyadan bulunan konumlar Uyarýlar'da görünür.
Public Sub KonumUyarilari()
    Dim note As Variant
    If AutoLayout Is Nothing Then Exit Sub
    For Each note In AutoLayout
        Warn "", CStr(note(1)), "Konum otomatik bulundu (" & note(0) & "): " & note(2)
    Next note
End Sub

' Ayardaki konum bu sayfada tutmadýysa sayfadan bulunan konumla yeniden dener (MakroVeri.TryCost).
' Baþarýlýysa Cfg eski hâline döner ve açýklama metni döner; deðilse "".
Public Function CostRetry(ByVal ws As Worksheet, ByRef found As Object) As Object
    Dim code As Variant, saved As Object, changed As Object
    Set found = CostLayout(ws)
    If found.Exists("neden") Then found.Remove "neden"
    If Not found.Exists("tarih_baslangici") Then Exit Function
    Set saved = Dict(): Set changed = Dict()
    For Each code In found.Keys
        If CStr(Opt("Maliyet", CStr(code))) <> CStr(found(code)) Then changed(code) = found(code)
        If Cfg.Exists("Maliyet" & Sep & code) Then saved(code) = Cfg("Maliyet" & Sep & code) Else saved(code) = Empty
    Next code
    If changed.Count = 0 Then Exit Function
    For Each code In found.Keys
        Cfg("Maliyet" & Sep & code) = found(code)
    Next code
    Set found = changed
    Set CostRetry = saved
End Function

Public Sub CostRestore(ByVal saved As Object)
    Dim code As Variant
    For Each code In saved.Keys
        If IsEmpty(saved(code)) Then
            If Cfg.Exists("Maliyet" & Sep & code) Then Cfg.Remove "Maliyet" & Sep & code
        Else
            Cfg("Maliyet" & Sep & code) = saved(code)
        End If
    Next code
End Sub
