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
    names = Split("Ocak Þubat Mart Nisan Mayýs Haziran Temmuz Aðustos Eylül Ekim Kasým Aralýk", " ")
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
    labels = Array("Ayar dosyasý seç", "Kaynaklar ve eþleþtirmeler", "Çýktý klasörü seç", "Kaynaklarý önizle", "Rapor oluþtur", "Açýklamalarý birleþtir", "Sentetik hesap testi", "Son çýktý klasörünü aç")
    actions = Array("AyarSec", "AyarKaynakOku", "KlasorSec", "KaynakOnizle", "RaporOlustur", "AciklamaBirlestir", "SentetikTest", "SonKlasoruAc")
    For i = 0 To UBound(labels)
        j = 17 + i * 2
        Set button = ws.Buttons.Add(ws.Cells(j, 2).Left, ws.Cells(j, 2).Top, 280, 24)
        button.Name = "MRV_Dugme" & i: button.Caption = labels(i)
        button.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & actions(i)
    Next i
    AyarKaynakOku
    MerkezHazirla
    WorkspaceSheet("Çalýþma Merkezi").Activate
    ActiveWindow.DisplayGridlines = False
    ws.Range("B34").Value2 = "Hazýr. Aylarý seçin, kaynaklarý kontrol edin ve Rapor oluþtur'a basýn."
End Sub

Public Sub AyarSec()
    Dim path As Variant
    If WorkRunning Then Exit Sub
    path = Application.GetOpenFilename("Excel ayar dosyasý (*.xlsx),*.xlsx", , "Ayarlar Excel'ini seçin")
    If VarType(path) <> vbBoolean Then
        ThisWorkbook.Worksheets("Kontrol Paneli").Range("D7").Value2 = CStr(path)
        AyarKaynakOku
        If Not SaveController() Then MsgBox "Ayar yolu seçildi. Kontrol dosyasýný Ctrl+S ile kaydedin.", vbInformation
    End If
End Sub

Public Sub KlasorSec()
    If WorkRunning Then Exit Sub
    With Application.FileDialog(4)
        .Title = "Raporlarýn kaydedileceði klasörü seçin"
        If .Show = -1 Then
            ThisWorkbook.Worksheets("Kontrol Paneli").Range("D11").Value2 = .SelectedItems(1)
            If Not SaveController() Then MsgBox "Çýktý yolu seçildi. Kontrol dosyasýný Ctrl+S ile kaydedin.", vbInformation
        End If
    End With
End Sub

Public Sub SonKlasoruAc()
    Dim path As String
    path = CStr(ThisWorkbook.Worksheets("Kontrol Paneli").Range("D36").Value2)
    If path = "" Then MsgBox "Henüz çýktý üretilmedi.", vbInformation: Exit Sub
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(path) Then MsgBox "Klasör bulunamadý: " & path, vbExclamation: Exit Sub
    CreateObject("Shell.Application").Open path
End Sub

Public Sub ReadPanel()
    Dim ws As Worksheet, lb As Object, i As Long, folder As String, normalized As String
    Set ws = ThisWorkbook.Worksheets("Kontrol Paneli")
    RootPath = ThisWorkbook.Path
    SettingsPath = AbsPath(Trim$(CStr(ws.Range("D7").Value2)), ThisWorkbook.Path)
    CloseYear = ValidInteger(ws.Range("D8").Value2, 1900, 9998, "Kapanýþ yýlý")
    CloseMonth = ValidInteger(ws.Range("D9").Value2, 1, 12, "Kapanýþ ayý")
    ReportYear = ValidInteger(ws.Range("D10").Value2, 1901, CloseYear, "Rapor yýlý")
    OutputPath = AbsPath(Trim$(CStr(ws.Range("D11").Value2)), ThisWorkbook.Path)
    Set Picked = Dict(): Set Warnings = New Collection
    On Error Resume Next
    Set lb = ws.ListBoxes("MRV_Aylar")
    On Error GoTo 0
    If lb Is Nothing Then Fail "Ay seçimi listesi yok. Alt+F8 -> PanelHazirla çalýþtýrýn."
    For i = 1 To 12
        If lb.Selected(i) Then
            If ReportYear = CloseYear And i > CloseMonth Then Fail "Kapanýþtan sonraki ay seçilemez: " & i
            Picked(CStr(i)) = True
        End If
    Next i
    If Picked.Count = 0 Then Fail "En az bir ay seçin."
    LoadSettings
    folder = AbsPath(CStr(Opt("Dosyalar", "maliyet_klasoru")))
    normalized = LCase$(OutputPath & "\")
    If Left$(normalized, Len(folder & "\")) = LCase$(folder & "\") Then Fail "Çýktý klasörü maliyet kaynak klasörünün içinde olamaz."
    If ReportYear <> CloseYear Then Warn "", "", "Seçili yýl bütçe yýlý deðil; bütçe karþýlaþtýrmalarý boþ býrakýlacak."
End Sub

Private Function ValidInteger(ByVal value As Variant, ByVal low As Long, ByVal high As Long, ByVal label As String) As Long
    If IsError(value) Or IsEmpty(value) Or Not IsNumeric(value) Then Fail label & " tam sayý olmalý."
    If CDbl(value) <> Fix(CDbl(value)) Or CDbl(value) < low Or CDbl(value) > high Then Fail label & " sýnýr dýþýnda."
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
    WorkGate "Girdi kopyalarý hazýrlanýyor"
    ArchiveBegin
    LoadSnapshots
    WorkGate "Ciro kontrolü"
    LoadControl
    WorkGate "Endeks"
    LoadIndices
    Set names = ItemNames()
    If names.Count = 0 Then Fail "Seçilen kapanýþta raporlanabilir gider bulunamadý. Kaynaklarý ve eþleþtirmeleri kontrol edin."
    CalculateUnits: CalculateChecks
    ValidateArchive ArchivePath
    SaveReports
    ws.Range("D36").Value2 = RunPath
    ws.Range("B34").Value2 = "Tamamlandý. " & Managers(DeptScope()).Count + 1 & " dosya; " & Warnings.Count & " uyarý."
    success = True
    ArchiveEnd True, ""
    If Not SaveController() Then ws.Range("B34").Value2 = ws.Range("B34").Value2 & " Kontrol dosyasýný Ctrl+S ile kaydedin."
Done:
    ' Geri yükleme adýmlarý hata verse bile Excel ayarlarý mutlaka eski hâline döner.
    On Error Resume Next
    If Not success Then ArchiveEnd False, errText
    WorkRunning = False
    Application.StatusBar = False: Application.Calculation = calc: Application.EnableEvents = events
    Application.DisplayAlerts = display: Application.ScreenUpdating = screen: Application.AutomationSecurity = security
    ThisWorkbook.Activate
    If success Then
        MsgBox "Raporlar kaydedildi:" & vbCrLf & RunPath & vbCrLf & vbCrLf & "Uyarý sayýsý: " & Warnings.Count & ". Kontrol ve Revizyon sayfasýný inceleyin.", vbInformation
    Else
        MsgBox "Rapor oluþturulamadý:" & vbCrLf & errText, vbExclamation
    End If
    Exit Sub
Bad:
    errText = Err.Description
    ' Resume hata durumunu temizler; GoTo ile çýkýlsaydý Done içindeki ikinci bir hata
    ' yakalanmaz ve hesaplama/ekran ayarlarý kapalý kalabilirdi.
    Resume Failed
Failed:
    On Error Resume Next
    ThisWorkbook.Worksheets("Kontrol Paneli").Range("B34").Value2 = "Hata: " & errText
    GoTo Done
End Sub
