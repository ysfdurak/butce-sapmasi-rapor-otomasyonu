Attribute VB_Name = "MakroPanel"
Option Explicit

Public Sub PanelHazirla()
    Dim ws As Worksheet, lb As Object, button As Object, i As Long, names As Variant, item As Variant
    Dim labels As Variant, actions As Variant, j As Long, old As Object, keep(1 To 12) As Boolean, hadList As Boolean, limit As Long
    If WorkRunning Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Kontrol Paneli")
    On Error Resume Next
    Set old = ws.ListBoxes("MRV_Aylar")
    On Error GoTo 0
    hadList = Not old Is Nothing
    If hadList Then
        For i = 1 To 12
            keep(i) = old.Selected(i)
        Next i
    End If
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "MRV_" Then ws.Shapes(i).Delete
    Next i
    names = Split("Ocak Şubat Mart Nisan Mayıs Haziran Temmuz Ağustos Eylül Ekim Kasım Aralık", " ")
    Set lb = ws.ListBoxes.Add(ws.Range("H7").Left, ws.Range("H7").Top, 170, 215)
    lb.Name = "MRV_Aylar": lb.MultiSelect = xlSimple
    For i = 0 To 11
        lb.AddItem Format$(i + 1, "00") & " " & names(i)
        If hadList Then
            lb.Selected(i + 1) = keep(i + 1)
        Else
            limit = 12
            If IsNumeric(ws.Range("D9").Value2) Then limit = CLng(ws.Range("D9").Value2)
            lb.Selected(i + 1) = (i < limit)
        End If
    Next i
    labels = Array("Ayar dosyası seç", "Kaynaklar ve eşleştirmeler", "Çıktı klasörü seç", "Kaynakları önizle", "Rapor oluştur", "Açıklamaları birleştir", "Sentetik hesap testi", "Son çıktı klasörünü aç")
    actions = Array("AyarSec", "AyarKaynakOku", "KlasorSec", "KaynakOnizle", "RaporOlustur", "AciklamaBirlestir", "SentetikTest", "SonKlasoruAc")
    For i = 0 To UBound(labels)
        j = 17 + i * 2
        Set button = ws.Buttons.Add(ws.Cells(j, 2).Left, ws.Cells(j, 2).Top, 280, 24)
        button.Name = "MRV_Dugme" & i: button.Caption = labels(i)
        button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & actions(i)
    Next i
    AyarKaynakOku
    MerkezHazirla
    WorkspaceSheet("Çalışma Merkezi").Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("B34").Value2 = "Hazır. Ayları seçin, kaynakları kontrol edin ve Rapor oluştur'a basın."
End Sub

Public Sub AyarSec()
    Dim path As Variant
    If WorkRunning Then Exit Sub
    path = Application.GetOpenFilename("Excel ayar dosyası (*.xlsx),*.xlsx", , "Ayarlar Excel'ini seçin")
    If VarType(path) <> vbBoolean Then
        ThisWorkbook.Worksheets("Kontrol Paneli").Range("D7").Value2 = CStr(path)
        AyarKaynakOku
        If Not SaveController() Then MsgBox "Ayar yolu seçildi. Kontrol dosyasını Ctrl+S ile kaydedin.", vbInformation
    End If
End Sub

Public Sub KlasorSec()
    If WorkRunning Then Exit Sub
    With Application.FileDialog(4)
        .Title = "Raporların kaydedileceği klasörü seçin"
        If .Show = -1 Then
            ThisWorkbook.Worksheets("Kontrol Paneli").Range("D11").Value2 = .SelectedItems(1)
            If Not SaveController() Then MsgBox "Çıktı yolu seçildi. Kontrol dosyasını Ctrl+S ile kaydedin.", vbInformation
        End If
    End With
End Sub

Public Sub SonKlasoruAc()
    Dim path As String
    path = CStr(ThisWorkbook.Worksheets("Kontrol Paneli").Range("D36").Value2)
    If path = "" Then MsgBox "Henüz çıktı üretilmedi.", vbInformation: Exit Sub
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(path) Then MsgBox "Klasör bulunamadı: " & path, vbExclamation: Exit Sub
    CreateObject("Shell.Application").Open path
End Sub

Public Sub ReadPanel()
    Dim ws As Worksheet, lb As Object, i As Long, folder As String, normalized As String
    Set ws = ThisWorkbook.Worksheets("Kontrol Paneli")
    RootPath = ThisWorkbook.Path
    SettingsPath = AbsPath(Trim$(CStr(ws.Range("D7").Value2)), ThisWorkbook.Path)
    CloseYear = ValidInteger(ws.Range("D8").Value2, 1900, 9998, "Kapanış yılı")
    CloseMonth = ValidInteger(ws.Range("D9").Value2, 1, 12, "Kapanış ayı")
    ReportYear = ValidInteger(ws.Range("D10").Value2, 1901, CloseYear, "Rapor yılı")
    OutputPath = AbsPath(Trim$(CStr(ws.Range("D11").Value2)), ThisWorkbook.Path)
    Set Picked = Dict(): Set Warnings = New Collection
    On Error Resume Next
    Set lb = ws.ListBoxes("MRV_Aylar")
    On Error GoTo 0
    If lb Is Nothing Then Fail "Ay seçimi listesi yok. Alt+F8 -> PanelHazirla çalıştırın."
    For i = 1 To 12
        If lb.Selected(i) Then
            If ReportYear = CloseYear And i > CloseMonth Then Fail "Kapanıştan sonraki ay seçilemez: " & i
            Picked(CStr(i)) = True
        End If
    Next i
    If Picked.Count = 0 Then Fail "En az bir ay seçin."
    LoadSettings
    folder = AbsPath(CStr(Opt("Dosyalar", "maliyet_klasoru")))
    normalized = LCase$(OutputPath & "\")
    If Left$(normalized, Len(folder & "\")) = LCase$(folder & "\") Then Fail "Çıktı klasörü maliyet kaynak klasörünün içinde olamaz."
    If ReportYear <> CloseYear Then Warn "", "", "Seçili yıl bütçe yılı değil; bütçe karşılaştırmaları boş bırakılacak."
End Sub

Private Function ValidInteger(ByVal value As Variant, ByVal low As Long, ByVal high As Long, ByVal label As String) As Long
    If IsError(value) Or IsEmpty(value) Or Not IsNumeric(value) Then Fail label & " tam sayı olmalı."
    If CDbl(value) <> Fix(CDbl(value)) Or CDbl(value) < low Or CDbl(value) > high Then Fail label & " sınır dışında."
    ValidInteger = CLng(value)
End Function

Public Sub KaynakOnizle()
    Dim ws As Worksheet, folder As String, name As String, period As String, r As Long
    Dim security As Long, errText As String
    If WorkRunning Then Exit Sub
    On Error GoTo Bad
    security = Application.AutomationSecurity: Application.AutomationSecurity = 3
    ReadPanel
    LoadSnapshots
    PaintSources True
Done:
    Application.AutomationSecurity = security
    Exit Sub
Bad:
    errText = Err.Description: MsgBox errText, vbExclamation
    Resume Done
End Sub

Public Sub RaporOlustur()
    Dim calc As XlCalculation, events As Boolean, display As Boolean, security As Long, screen As Boolean
    Dim errText As String, names As Object, success As Boolean, ws As Worksheet
    If WorkRunning Then Exit Sub
    WorkRunning = True: StopRequested = False
    calc = Application.Calculation: events = Application.EnableEvents: display = Application.DisplayAlerts
    screen = Application.ScreenUpdating: security = Application.AutomationSecurity
    On Error GoTo Bad
    Application.EnableEvents = False: Application.DisplayAlerts = False
    Application.ScreenUpdating = False: Application.AutomationSecurity = 3
    Application.Calculation = xlCalculationManual
    Set ws = ThisWorkbook.Worksheets("Kontrol Paneli")
    ws.Range("B34").Value2 = "Kaynaklar okunuyor…"
    YorumKaydedilmemisKontrol
    ReadPanel
    WorkGate "Girdi kopyaları hazırlanıyor"
    ArchiveBegin
    LoadSnapshots
    WorkGate "Ciro kontrolü"
    LoadControl
    WorkGate "Endeks"
    LoadIndices
    Set names = ItemNames()
    If names.Count = 0 Then Fail "Seçilen kapanışta raporlanabilir gider bulunamadı. Kaynakları ve eşleştirmeleri kontrol edin."
    CalculateUnits: CalculateChecks
    ValidateArchive ArchivePath
    SaveReports
    ws.Range("D36").Value2 = RunPath
    ws.Range("B34").Value2 = "Tamamlandı. " & Managers(DeptScope()).Count + 1 & " dosya; " & Warnings.Count & " uyarı."
    success = True
    ArchiveEnd True, ""
    If Not SaveController() Then ws.Range("B34").Value2 = ws.Range("B34").Value2 & " Kontrol dosyasını Ctrl+S ile kaydedin."
Done:
    If Not success Then ArchiveEnd False, errText
    WorkRunning = False
    Application.StatusBar = False: Application.Calculation = calc: Application.EnableEvents = events
    Application.DisplayAlerts = display: Application.ScreenUpdating = screen: Application.AutomationSecurity = security
    ThisWorkbook.Activate
    If success Then
        MsgBox "Raporlar kaydedildi:" & vbCrLf & RunPath & vbCrLf & vbCrLf & "Uyarı sayısı: " & Warnings.Count & ". Kontrol ve Revizyon sayfasını inceleyin.", vbInformation
    Else
        MsgBox "Rapor oluşturulamadı:" & vbCrLf & errText, vbExclamation
    End If
    Exit Sub
Bad:
    errText = Err.Description
    On Error Resume Next
    ThisWorkbook.Worksheets("Kontrol Paneli").Range("B34").Value2 = "Hata: " & errText
    On Error GoTo 0
    GoTo Done
End Sub
