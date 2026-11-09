Attribute VB_Name = "MakroAyarlar"
Option Explicit

Public Function SettingsSignature(ByVal path As String) As String
    ' Adler checksum + exact byte length detects accidental changes since preview.
    Dim handle As Integer, raw() As Byte, length As Long, i As Long, a As Double, b As Double, opened As Boolean, message As String
    On Error GoTo Failed
    handle = FreeFile
    Open path For Binary Access Read As #handle
    opened = True: length = LOF(handle)
    If length > 0 Then
        ReDim raw(0 To length - 1): Get #handle, , raw
    End If
    Close #handle: opened = False
    a = 1: b = 0
    For i = 0 To length - 1
        a = (a + raw(i)) Mod 65521
        b = (b + a) Mod 65521
    Next i
    SettingsSignature = CStr(length) & ":" & Format$(a, "0") & ":" & Format$(b, "0")
    Exit Function
Failed:
    message = Err.Description
    On Error Resume Next
    If opened Then Close #handle
    On Error GoTo 0
    Fail "Ayar dosyasý okunamadý: " & message
End Function

Public Sub AyarKaynakOku()
    Dim source As Worksheet, message As String, security As Long
    If WorkRunning Then Exit Sub
    security = Application.AutomationSecurity
    On Error GoTo Failed
    Application.AutomationSecurity = 3
    Set source = ThisWorkbook.Worksheets("Kaynak Önizleme")
    SettingsPath = AbsPath(CStr(ThisWorkbook.Worksheets("Kontrol Paneli").Range("D7").Value2), ThisWorkbook.Path)
    Set Warnings = New Collection: Set Snaps = Dict()
    LoadSettings
    With ThisWorkbook.Worksheets("Kontrol Paneli")
        If IsNumeric(.Range("D8").Value2) Then CloseYear = CLng(.Range("D8").Value2)
        If IsNumeric(.Range("D9").Value2) Then CloseMonth = CLng(.Range("D9").Value2)
    End With
    PaintSources False
    ThisWorkbook.Worksheets("Kontrol Paneli").Range("D7").Value2 = SettingsPath
    source.Activate
Done:
    Application.AutomationSecurity = security
    Exit Sub
Failed:
    message = Err.Description: MsgBox "Ayarlar okunamadý: " & message, vbExclamation
    Resume Done
End Sub

Public Sub PaintSources(ByVal loaded As Boolean)
    Dim ws As Worksheet, n As Variant, meta As Variant, g As Variant, code As Variant, mapping As Variant
    Dim r As Long, name As String, id As String, state As String, d As Object, i As Long, row As Variant, button As Object
    Dim actions As Variant, labels As Variant, cells As Variant, path As String
    Set ws = ThisWorkbook.Worksheets("Kaynak Önizleme")
    ws.Range("A15:L" & Application.Max(16, ws.UsedRange.Row + ws.UsedRange.Rows.Count)).Clear
    ws.Range("D4").Value2 = SettingsPath
    ws.Range("D6").Value2 = AbsPath(CStr(Opt("Dosyalar", "maliyet_klasoru")))
    ws.Range("D7").Value2 = AbsPath(CStr(Opt("Dosyalar", "kontrol_dosyasi")))
    ws.Range("D8").Value2 = AbsPath(CStr(Opt("Dosyalar", "endeks_dosyasi")))
    ws.Range("N4").Value2 = SettingsPath: ws.Range("N5").Value2 = SettingsSignature(SettingsPath)
    ws.Columns("N").Hidden = True
    ws.Range("A12").Value2 = Depts.Count & " bölüm; " & Managers(DeptScope()).Count & " müdürlük. Tüm kayýtlar ayar Excel'inden okundu."
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "MRA_" Then ws.Shapes(i).Delete
    Next i
    actions = Array("MaliyetKlasoruSec", "KontrolDosyasiSec", "EndeksDosyasiSec", "AyarKaynakOku", "KonumlariKaydet", "KaynakOnizle", "PaneliKaydet")
    labels = Array("Klasör seç", "Dosya seç", "Dosya seç", "Ayarlarý oku", "Konumlarý kaydet", "Kaynaklarý yenile", "Seçimleri kaydet")
    cells = Array("J6", "J7", "J8", "A10", "D10", "G10", "J10")
    For i = 0 To UBound(actions)
        Set button = ws.Buttons.Add(ws.Range(cells(i)).Left, ws.Range(cells(i)).Top, 150, 24)
        button.Name = "MRA_Dugme" & i: button.Caption = labels(i)
        button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & actions(i)
    Next i
    r = 15
    RowValues ws, r, Array("Bölüm", "Müdürlük", "Direktörlük", "Müdürlük altý", "Lokasyon", "Kontrol bölümü", "Endeks adý", "Kaynak durumu")
    Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)): r = r + 1
    For Each n In Depts.Keys
        meta = Depts(n): state = "Kaynak önizlemesi bekleniyor"
        If loaded Then
            Set d = CurrentSnap(CStr(n))
            If d Is Nothing Then state = "Kapanýþ kaynaðý eksik/okunamadý" Else state = "Okundu"
        End If
        RowValues ws, r, Array(meta(0), meta(1), meta(2), meta(6), meta(4), meta(5), meta(3), state): r = r + 1
    Next n
    r = r + 2: RowValues ws, r, Array("Ýlk indeks", "Son indeks", "Ana gider grubu")
    Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)): r = r + 1
    For Each g In Groups
        RowValues ws, r, g: r = r + 1
    Next g
    r = r + 2: RowValues ws, r, Array("Kaynak gider kodu", "Rapor gider kodu", "Birleþtirilmiþ gider adý")
    Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)): r = r + 1
    For Each code In ExpenseMap.Keys
        mapping = ExpenseMap(code): RowValues ws, r, Array(CLng(code), CLng(mapping(0)), mapping(1)): r = r + 1
    Next code
    r = r + 2: ws.Cells(r, 1).Value2 = "Okunmayacak ortak sayfalar": Header ws.Cells(r, 1): r = r + 1
    For Each n In Excluded.Keys
        ws.Cells(r, 1).Value2 = n: r = r + 1
    Next n
    r = r + 2: RowValues ws, r, Array("Dosya", "Dosya adýndaki dönem", "Seçim sonucu")
    Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)): r = r + 1
    path = AbsPath(CStr(Opt("Dosyalar", "maliyet_klasoru")))
    If CreateObject("Scripting.FileSystemObject").FolderExists(path) Then
        name = Dir$(path & "\*.xlsx")
        Do While name <> ""
            If Left$(name, 2) <> "~$" Then
                id = FilenamePeriod(name): state = "Seçim dýþýnda"
                If id = "" Then
                    state = "Tek yýl/ay bulunamadý; okunmayacak"
                ElseIf id = P(CloseYear, CloseMonth) Then
                    state = "Rapor kaynaðý"
                ElseIf Left$(id, 4) = CStr(CloseYear) And id < P(CloseYear, CloseMonth) Then
                    state = "Revizyon karþýlaþtýrmasý"
                End If
                RowValues ws, r, Array(name, id, state): r = r + 1
            End If
            name = Dir$()
        Loop
    Else
        ws.Cells(r, 1).Value2 = "Maliyet klasörünü seçip konumlarý kaydedin.": r = r + 1
    End If
    If loaded Then
        r = r + 2: RowValues ws, r, Array("Bölüm", "Gösterge sýrasý", "Kaynak Excel baþlýðý", "Kaynak dosya")
        Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)): r = r + 1
        For Each n In Depts.Keys
            meta = Depts(n): Set d = CurrentSnap(CStr(n))
            If Not d Is Nothing Then
                For i = 1 To 5
                    RowValues ws, r, Array(meta(0), i, IndicatorName(d, i), d("file")): r = r + 1
                Next i
            End If
        Next n
        r = r + 2: RowValues ws, r, Array("Kapanýþ", "Bölüm", "Dosya", "Kaynak uyarýsý")
        Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)): r = r + 1
        For Each row In Warnings
            RowValues ws, r, row: r = r + 1
        Next row
    End If
    ws.Range("A15:L" & r).Font.Name = "Calibri": ws.Range("A15:L" & r).Font.Size = 11
    ws.Range("A15:L" & r).WrapText = True: ws.Rows("15:" & r).AutoFit
    ws.Activate: ActiveWindow.DisplayGridlines = False
End Sub

Public Sub MaliyetKlasoruSec()
    If WorkRunning Then Exit Sub
    With Application.FileDialog(4)
        .Title = "Maliyet Excel dosyalarýnýn klasörünü seçin"
        If .Show = -1 Then ThisWorkbook.Worksheets("Kaynak Önizleme").Range("D6").Value2 = .SelectedItems(1)
    End With
End Sub

Public Sub KontrolDosyasiSec()
    If WorkRunning Then Exit Sub
    PickSourceFile "D7", "Ciro ve tutar kontrol dosyasýný seçin"
End Sub

Public Sub EndeksDosyasiSec()
    If WorkRunning Then Exit Sub
    PickSourceFile "D8", "Endeks dosyasýný seçin"
End Sub

Private Sub PickSourceFile(ByVal cell As String, ByVal title As String)
    Dim path As Variant
    path = Application.GetOpenFilename("Excel (*.xlsx),*.xlsx", , title)
    If VarType(path) <> vbBoolean Then ThisWorkbook.Worksheets("Kaynak Önizleme").Range(cell).Value2 = CStr(path)
End Sub

Public Function SaveController() As Boolean
    On Error GoTo Failed
    If ThisWorkbook.ReadOnly Then Exit Function
    ThisWorkbook.Save: SaveController = True
    Exit Function
Failed:
    SaveController = False
End Function

Public Sub PaneliKaydet()
    If SaveController() Then
        MsgBox "Ayar yolu, çýktý yolu ve ay seçimleri kaydedildi.", vbInformation
    Else
        MsgBox "Kontrol dosyasý kaydedilemedi. Yazýlabilir bir .xlsm kopyasý kaydedin; sonra Ctrl+S kullanýn.", vbExclamation
    End If
End Sub

Public Sub KonumlariKaydet()
    Dim ws As Worksheet, target As String, signature As String, folder As String, control As String, indexFile As String
    Dim fso As Object, wb As Workbook, sheet As Worksheet, row As Long, code As String, seen As Object, names As Object
    Dim temporary As String, rollback As String, backup As String, stamp As String, message As String
    Dim movedOriginal As Boolean, committed As Boolean, display As Boolean, events As Boolean, security As Long
    If WorkRunning Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    display = Application.DisplayAlerts: events = Application.EnableEvents: security = Application.AutomationSecurity
    On Error GoTo Failed
    Application.DisplayAlerts = False: Application.EnableEvents = False: Application.AutomationSecurity = 3
    Set ws = ThisWorkbook.Worksheets("Kaynak Önizleme")
    target = AbsPath(CStr(ThisWorkbook.Worksheets("Kontrol Paneli").Range("D7").Value2), ThisWorkbook.Path)
    If target <> CStr(ws.Range("N4").Value2) Then Fail "Ayar dosyasý deðiþti. Önce Ayarlarý oku düðmesine basýn."
    signature = CStr(ws.Range("N5").Value2)
    If signature = "" Or SettingsSignature(target) <> signature Then Fail "Ayar Excel'i son okumadan sonra deðiþti. Ayarlarý yeniden okuyun."
    SettingsPath = target: RootPath = fso.GetParentFolderName(target)
    folder = Trim$(CStr(ws.Range("D6").Value2)): control = Trim$(CStr(ws.Range("D7").Value2)): indexFile = Trim$(CStr(ws.Range("D8").Value2))
    If folder = "" Or control = "" Or indexFile = "" Then Fail "Üç kaynak konumunu birlikte seçin."
    folder = AbsPath(folder): control = AbsPath(control): indexFile = AbsPath(indexFile)
    If Not fso.FolderExists(folder) Then Fail "Maliyet klasörü bulunamadý: " & folder
    If Not fso.FileExists(control) Or LCase$(fso.GetExtensionName(control)) <> "xlsx" Then Fail "Kontrol Excel dosyasý bulunamadý: " & control
    If Not fso.FileExists(indexFile) Or LCase$(fso.GetExtensionName(indexFile)) <> "xlsx" Then Fail "Endeks Excel dosyasý bulunamadý: " & indexFile
    ReadPanel
    stamp = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 1000))
    temporary = fso.BuildPath(RootPath, ".butce_ayar_" & stamp & ".xlsx")
    rollback = fso.BuildPath(RootPath, ".butce_onceki_" & stamp & ".xlsx")
    If fso.FileExists(temporary) Or fso.FileExists(rollback) Then Fail "Geçici ad çakýþtý; iþlemi yeniden baþlatýn."
    fso.CopyFile target, temporary, False
    Set wb = Application.Workbooks.Open(Filename:=temporary, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    If wb.ReadOnly Then Fail "Geçici ayar dosyasý yazýlabilir açýlamadý."
    Set sheet = wb.Worksheets("Dosya ve Konum Ayarlarý"): Set seen = Dict()
    For row = 5 To sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
        If CStr(sheet.Cells(row, 1).Value2) = "Dosyalar" Then
            code = CStr(sheet.Cells(row, 2).Value2)
            Select Case code
                Case "maliyet_klasoru": sheet.Cells(row, 3).Value2 = folder
                Case "kontrol_dosyasi": sheet.Cells(row, 3).Value2 = control
                Case "endeks_dosyasi": sheet.Cells(row, 3).Value2 = indexFile
            End Select
            If code = "maliyet_klasoru" Or code = "kontrol_dosyasi" Or code = "endeks_dosyasi" Then
                If seen.Exists(code) Then Fail "Mükerrer konum ayarý: " & code
                seen.Add code, True
            End If
        End If
    Next row
    If seen.Count <> 3 Then Fail "Üç kaynak konumu ayar Excel'inde bulunamadý."
    wb.Save: wb.Close False: Set wb = Nothing
    SettingsPath = temporary
    LoadSettings: LoadSnapshots: LoadControl: LoadIndices
    Set names = ItemNames()
    If names.Count = 0 Then Fail "Yeni konumlarda kapanýþ giderleri okunamadý."
    If SettingsSignature(target) <> signature Then Fail "Ayar dosyasý iþlem sýrasýnda deðiþti; üzerine yazýlmadý."
    folder = fso.BuildPath(RootPath, "Ayar_Yedekleri")
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    backup = fso.BuildPath(folder, fso.GetBaseName(target) & "_" & stamp & ".xlsx")
    fso.CopyFile target, backup, False
    fso.MoveFile target, rollback: movedOriginal = True
    fso.MoveFile temporary, target: temporary = "": committed = True
    SettingsPath = target
    On Error Resume Next
    fso.DeleteFile rollback, True: movedOriginal = False
    ThisWorkbook.Worksheets("Kontrol Paneli").Range("D7").Value2 = target
    PaintSources True
    ws.Range("A12").Value2 = "Konumlar kaydedildi. Ayar yedeði: " & backup
    If Not SaveController() Then ws.Range("A12").Value2 = "Konumlar ayar Excel'ine kaydedildi. Kontrol dosyasýný ayrýca Ctrl+S ile kaydedin. Yedek: " & backup
    On Error GoTo 0
Done:
    Application.DisplayAlerts = display: Application.EnableEvents = events: Application.AutomationSecurity = security
    If committed Then MsgBox "Kaynak konumlarý kaydedildi; tüm eþleþtirmeler yenilendi." & vbCrLf & "Yedek: " & backup, vbInformation
    If message <> "" Then MsgBox "Konumlar kaydedilemedi: " & message, vbExclamation
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If movedOriginal And Not committed Then
        If Not fso.FileExists(target) Then fso.MoveFile rollback, target
        If fso.FileExists(rollback) Then message = message & vbCrLf & "Önceki ayar dosyasý korundu: " & rollback & vbCrLf & "Yedek: " & backup
    End If
    If temporary <> "" Then
        If Not movedOriginal Or fso.FileExists(target) Then
            If fso.FileExists(temporary) Then fso.DeleteFile temporary, True
        Else
            message = message & vbCrLf & "Doðrulanmýþ geçici ayar: " & temporary
        End If
    End If
    SettingsPath = target
    Set Warnings = New Collection: LoadSettings
    On Error GoTo 0
    GoTo Done
End Sub
