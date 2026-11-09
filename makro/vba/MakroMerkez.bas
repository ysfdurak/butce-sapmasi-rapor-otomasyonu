Attribute VB_Name = "MakroMerkez"
Option Explicit

Public WorkRunning As Boolean, StopRequested As Boolean, ReplayInput As String
Public ArchivePath As String, HistoryRow As Long
Private OriginalSettings As String, OriginalRoot As String

Public Function WorkspaceSheet(ByVal title As String) As Worksheet
    On Error Resume Next
    Set WorkspaceSheet = ThisWorkbook.Worksheets(title)
    On Error GoTo 0
    If WorkspaceSheet Is Nothing Then
        Set WorkspaceSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        WorkspaceSheet.Name = title
    End If
End Function

Public Sub WorkGate(ByVal message As String)
    Application.StatusBar = message
    ThisWorkbook.Worksheets("Kontrol Paneli").Range("B34").Value2 = message
    DoEvents
    If StopRequested Then Fail "Kullanıcı isteğiyle güvenli adım sınırında durduruldu. Önceki raporlar korundu."
End Sub

Public Sub Durdur()
    If WorkRunning Then
        StopRequested = True
        Application.StatusBar = "Durdurma istendi; mevcut dosya adımının bitmesi bekleniyor."
    End If
End Sub

Public Sub MerkezHazirla()
    Dim ws As Worksheet, names As Variant, actions As Variant, labels As Variant, i As Long, b As Object, old As Object
    Set ws = WorkspaceSheet("Çalışma Merkezi")
    ws.Cells.Clear
    ws.Range("A1").Value2 = "Bütçe Sapması | Çalışma merkezi"
    ws.Range("A3").Value2 = "Kaynakları kontrol et > sorunları incele > rapor oluştur. Rapor kapsamı ayar Excel'indeki tüm bölümlerdir."
    actions = Array("MerkezKontrol", "EslesmeIncele", "HesapIzle", "DuzenOnizle", "GorunumKaydet", "GorunumAc", "GecmiseGit", "RaporOlustur", "Durdur")
    labels = Array("Kontrolleri yenile", "Eşleştirme değişiklikleri", "Hesap nasıl oluştu?", "Tablo düzeni", "Görünümü kaydet", "Görünümü aç", "Çalıştırma geçmişi", "Rapor oluştur", "Güvenli durdur")
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    For i = 0 To UBound(actions)
        Set b = ws.Buttons.Add(ws.Cells(5 + (i \ 3) * 2, 1 + (i Mod 3) * 3).Left, ws.Cells(5 + (i \ 3) * 2, 1 + (i Mod 3) * 3).Top, 185, 26)
        b.Caption = labels(i): b.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & actions(i)
    Next i
    AddWorkButton ws, "Dönem ve çıktı seçimi", "DonemPanelineGit", "J5"
    AddWorkButton ws, "Önceki ayın yorumlarını getir", "YorumGetir", "J7"
    AddWorkButton ws, "Açıklamaları toplu düzenle", "YorumListele", "J9"
    ws.Columns("A:I").ColumnWidth = 24: ws.Rows("3:3").RowHeight = 42
    ws.Range("A3:I3").WrapText = True
    ws.Range("A1:I1").Interior.Color = RGB(225, 236, 239): ws.Range("A1").Font.Size = 19
    Set old = WorkspaceSheet("Kayıtlı Görünümler")
    If old.Range("A1").Value2 = "" Then RowValues old, 1, Array("Görünüm adı", "Ayar dosyası", "Kapanış yılı", "Kapanış ayı", "Rapor yılı", "Çıktı konumu", "Seçili aylar")
    Set old = WorkspaceSheet("Çalıştırma Geçmişi")
    If old.Range("A1").Value2 = "" Then RowValues old, 1, Array("Başlangıç", "Durum", "Ayar dosyası", "Kapanış yılı", "Kapanış ayı", "Rapor yılı", "Aylar", "Girdi arşivi", "Çıktı / hata")
    Set old = WorkspaceSheet("Eşleştirme Kaydı")
    If old.Range("A1").Value2 = "" Then RowValues old, 1, Array("Ayar yolu", "Kalıcı kimlik", "Bölüm", "Müdürlük", "Direktörlük", "Endeks", "Lokasyon", "Kontrol bölümü", "Müdürlük altı", "Durum", "Son onay")
    ws.Activate
End Sub

Private Sub FullRead()
    ReadPanel
    WorkGate "Maliyet dosyaları okunuyor"
    LoadSnapshots
    WorkGate "Ciro ve tutar kontrolleri okunuyor"
    LoadControl
    WorkGate "Endeks okunuyor"
    LoadIndices
    WorkGate "Birim maliyet ve revizyonlar hesaplanıyor"
    CalculateUnits: CalculateChecks
End Sub

Public Sub MerkezKontrol()
    Dim ws As Worksheet, r As Long, row As Variant, n As Variant, d As Object, meta As Variant
    Dim security As Long, message As String
    If WorkRunning Then Exit Sub
    WorkRunning = True: StopRequested = False
    security = Application.AutomationSecurity: Application.AutomationSecurity = 3
    On Error GoTo Bad
    FullRead
    Set ws = WorkspaceSheet("Çalışma Merkezi")
    ws.Range("A12:J" & Application.Max(13, ws.UsedRange.Rows.Count)).Clear
    ws.Range("A12").Value2 = Depts.Count & " eşleştirilmiş bölüm | " & Warnings.Count & " uyarı | " & Revisions.Count & " revizyon"
    RowValues ws, 14, Array("Tür", "Bölüm", "Dönem", "Dosya", "Sorun", "Yapılacak işlem")
    Header ws.Range("A14:F14"): r = 15
    For Each row In Warnings
        RowValues ws, r, Array("Kaynak", row(1), row(0), row(2), row(3), "Kaynak Önizleme ve ayar Excel'indeki eşleştirmeyi inceleyin."): r = r + 1
    Next row
    For Each row In Checks
        If row(7) <> "Doğrulandı" Then
            RowValues ws, r, Array("Dip toplam", row(0), row(3), AbsPath(CStr(Opt("Dosyalar", "kontrol_dosyasi"))), row(7), "Bölüm birleştirmelerini ve kaynak/kontrol tutarlarını karşılaştırın. Fark TL: " & CStr(IIf(IsNull(row(6)), "Eksik", row(6)))): r = r + 1
        End If
    Next row
    r = r + 2: RowValues ws, r, Array("Bölüm", "Hesap", "Kapsama alınmama gerekçesi"): Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)): r = r + 1
    For Each n In Depts.Keys
        meta = Depts(n): Set d = CurrentSnap(CStr(n))
        If Not Units.Exists(n) Then RowValues ws, r, Array(meta(0), "Birim maliyet", "Endeks / önceki yıl / pozitif işlem adedi / kaynak eksik"): r = r + 1
        If Not OpexEligible(CStr(n)) Then RowValues ws, r, Array(meta(0), "OPEX / ciro", "Önceki yıl gideri veya pozitif şirket cirosu eksik"): r = r + 1
    Next n
    ws.Range("A14:F" & r).WrapText = True: ws.Rows("14:" & r).AutoFit
    ws.Columns("D:F").ColumnWidth = 50
    ws.Activate
Done:
    WorkRunning = False: Application.StatusBar = False: Application.AutomationSecurity = security
    If message <> "" Then MsgBox message, vbExclamation
    Exit Sub
Bad:
    message = Err.Description: Resume Done
End Sub

Public Sub EslesmeIncele()
    Dim ws As Worksheet, reg As Worksheet, n As Variant, meta As Variant, old As Object, i As Long, r As Long
    Dim state As String, text As String, j As Long
    If WorkRunning Then Exit Sub
    On Error GoTo Bad
    ReadPanel
    Set ws = WorkspaceSheet("Eşleştirme İncelemesi"): Set reg = WorkspaceSheet("Eşleştirme Kaydı")
    ws.Cells.Clear: Set old = Dict()
    For i = 2 To reg.Cells(reg.Rows.Count, 1).End(xlUp).Row
        If reg.Cells(i, 1).Value2 = SettingsPath And reg.Cells(i, 10).Value2 = "Etkin" Then old(K(reg.Cells(i, 3).Value2)) = i
    Next i
    ws.Range("A1").Value2 = "Yeni / kaldırılan / değişen eşleştirmeler"
    ws.Range("A2").Value2 = "İsim değişikliği ise D sütununa eski bölüm adını yazın. Yalnız açıkça onaylanan adlar aynı kalıcı kimliğe bağlanır."
    ws.Range("K1").Value2 = SettingsPath: ws.Range("K2").Value2 = SettingsSignature(SettingsPath): ws.Columns("K").Hidden = True
    RowValues ws, 4, Array("Durum", "Bölüm", "Açıklama", "Eski bölüm adı (isteğe bağlı)")
    r = 5
    For Each n In Depts.Keys
        meta = Depts(n): state = "Eklendi"
        If old.Exists(n) Then
            state = "Aynı"
            For j = 1 To 6
                If CStr(reg.Cells(old(n), j + 3).Value2) <> CStr(meta(j)) Then state = "Eşleştirme değişti"
            Next j
        End If
        RowValues ws, r, Array(state, meta(0), "Ayar Excel'indeki kayıt; kaynak ve organizasyonu kontrol edin.", ""): r = r + 1
    Next n
    For Each n In old.Keys
        If Not Depts.Exists(n) Then RowValues ws, r, Array("Kaldırıldı", reg.Cells(old(n), 3).Value2, "Eski kayıt geçmişte korunur.", ""): r = r + 1
    Next n
    AddWorkButton ws, "İnceledim, kaydet", "EslesmeOnayla", "F4"
    ws.Columns("A:D").ColumnWidth = 38: ws.Columns("C").ColumnWidth = 65: ws.UsedRange.WrapText = True: ws.UsedRange.Rows.AutoFit
    ws.Activate: Exit Sub
Bad:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub EslesmeOnayla()
    Dim ws As Worksheet, reg As Worksheet, old As Object, renames As Object, used As Object, n As Variant, meta As Variant
    Dim r As Long, i As Long, target As String, previous As String, uid As String, dest As Long
    If WorkRunning Then Exit Sub
    On Error GoTo Bad
    Set ws = WorkspaceSheet("Eşleştirme İncelemesi"): Set reg = WorkspaceSheet("Eşleştirme Kaydı")
    ReadPanel
    If ws.Range("K1").Value2 <> SettingsPath Or ws.Range("K2").Value2 <> SettingsSignature(SettingsPath) Then Fail "Ayar değişti; eşleştirmeyi yeniden inceleyin."
    Set old = Dict(): Set renames = Dict(): Set used = Dict()
    For r = 2 To reg.Cells(reg.Rows.Count, 1).End(xlUp).Row
        If reg.Cells(r, 1).Value2 = SettingsPath And reg.Cells(r, 10).Value2 = "Etkin" Then old(K(reg.Cells(r, 3).Value2)) = r
    Next r
    For r = 5 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        previous = K(ws.Cells(r, 4).Value2): target = K(ws.Cells(r, 2).Value2)
        If previous <> "" Then
            If ws.Cells(r, 1).Value2 <> "Eklendi" Or Not old.Exists(previous) Or Depts.Exists(previous) Or used.Exists(previous) Then Fail "Eski/yeni bölüm çifti geçersiz veya tekrar ediyor: " & target
            renames(target) = previous: used(previous) = True
        End If
    Next r
    For r = 2 To reg.Cells(reg.Rows.Count, 1).End(xlUp).Row
        If reg.Cells(r, 1).Value2 = SettingsPath Then reg.Cells(r, 10).Value2 = "Geçmiş"
    Next r
    For Each n In Depts.Keys
        meta = Depts(n): uid = "B-" & Format$(Now, "yyyymmddhhnnss") & "-" & CStr(reg.Cells(reg.Rows.Count, 1).End(xlUp).Row + 1)
        If old.Exists(n) Then uid = CStr(reg.Cells(old(n), 2).Value2)
        If renames.Exists(n) Then uid = CStr(reg.Cells(old(renames(n)), 2).Value2)
        dest = reg.Cells(reg.Rows.Count, 1).End(xlUp).Row + 1
        RowValues reg, dest, Array(SettingsPath, uid, meta(0), meta(1), meta(2), meta(3), meta(4), meta(5), meta(6), "Etkin", Now)
    Next n
    If Not SaveController() Then Fail "Onay bellekte tutuldu; kontrol dosyasını Ctrl+S ile kaydedin."
    MsgBox "Eşleştirmeler ve kalıcı bölüm kimlikleri kaydedildi. Eski raporlar korunur; açıklama birleştirmede aynı kimliğin onaylı adları tanınır.", vbInformation
    Exit Sub
Bad:
    MsgBox Err.Description, vbExclamation
End Sub

Public Function CanonicalDept(ByVal name As String) As String
    Dim reg As Worksheet, r As Long, uid As String, path As String
    CanonicalDept = K(name): Set reg = WorkspaceSheet("Eşleştirme Kaydı")
    path = AbsPath(CStr(ThisWorkbook.Worksheets("Kontrol Paneli").Range("D7").Value2), ThisWorkbook.Path)
    For r = reg.Cells(reg.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If reg.Cells(r, 1).Value2 = path And K(reg.Cells(r, 3).Value2) = K(name) Then uid = CStr(reg.Cells(r, 2).Value2): Exit For
    Next r
    If uid = "" Then Exit Function
    For r = reg.Cells(reg.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If reg.Cells(r, 1).Value2 = path And reg.Cells(r, 2).Value2 = uid And reg.Cells(r, 10).Value2 = "Etkin" Then CanonicalDept = K(reg.Cells(r, 3).Value2): Exit For
    Next r
End Function

Public Sub AddWorkButton(ByVal ws As Worksheet, ByVal caption As String, ByVal action As String, ByVal address As String)
    Dim b As Object, i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name = "MC_" & action Then ws.Shapes(i).Delete
    Next i
    Set b = ws.Buttons.Add(ws.Range(address).Left, ws.Range(address).Top, 200, 25)
    b.Name = "MC_" & action: b.Caption = caption: b.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & action
End Sub

Public Sub GorunumKaydet()
    Dim ws As Worksheet, title As String, r As Long
    If WorkRunning Then Exit Sub
    On Error GoTo Bad
    ReadPanel
    title = Trim$(InputBox("Görünüm adı (rapor yılı, aylar, kaynak ve çıktı konumu kaydedilir):"))
    If title = "" Then Exit Sub
    Set ws = WorkspaceSheet("Kayıtlı Görünümler")
    r = 2
    Do While ws.Cells(r, 1).Value2 <> "" And ws.Cells(r, 1).Value2 <> title
        r = r + 1
    Loop
    RowValues ws, r, Array(title, SettingsPath, CloseYear, CloseMonth, ReportYear, OutputPath, Signature())
    ws.Columns("A:G").AutoFit
    If Not SaveController() Then MsgBox "Ctrl+S ile kaydedin.", vbInformation
    Exit Sub
Bad:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub GorunumAc()
    Dim ws As Worksheet, panel As Worksheet, title As String, r As Long, i As Long, lb As Object, chosen As String
    If WorkRunning Then Exit Sub
    Set ws = WorkspaceSheet("Kayıtlı Görünümler"): ws.Activate
    title = Trim$(InputBox("Açılacak görünümün A sütunundaki adı:"))
    If title = "" Then Exit Sub
    On Error GoTo Bad
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(r, 1).Value2 = title Then
            Set panel = ThisWorkbook.Worksheets("Kontrol Paneli")
            panel.Range("D7").Value2 = ws.Cells(r, 2).Value2
            For i = 3 To 6
                panel.Cells(i + 5, 4).Value2 = ws.Cells(r, i).Value2
            Next i
            Set lb = panel.ListBoxes("MRV_Aylar"): chosen = "," & Replace(CStr(ws.Cells(r, 7).Value2), "aylar:", "") & ","
            For i = 1 To 12
                lb.Selected(i) = (InStr(chosen, "," & Format$(i, "00") & ",") > 0)
            Next i
            AyarKaynakOku: panel.Activate: Exit Sub
        End If
    Next r
    Fail "Görünüm bulunamadı."
Bad:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub GecmiseGit()
    Dim ws As Worksheet
    Set ws = WorkspaceSheet("Çalıştırma Geçmişi")
    AddWorkButton ws, "Seçili çalışmayı yeniden üret", "GecmisiYenidenUret", "K2"
    AddWorkButton ws, "Kaynak değişikliklerini göster", "KaynakFarklari", "K4"
    ws.Columns("A:I").ColumnWidth = 27: ws.Columns("H:I").ColumnWidth = 55
    ws.UsedRange.WrapText = True: ws.UsedRange.Rows.AutoFit: ws.Activate
End Sub

Public Sub GecmisiYenidenUret()
    Dim ws As Worksheet, r As Long, panel As Worksheet, lb As Object, i As Long, chosen As String
    If WorkRunning Then Exit Sub
    On Error GoTo Bad
    Set ws = WorkspaceSheet("Çalıştırma Geçmişi")
    If ActiveSheet.Name <> ws.Name Or ActiveCell.Row < 2 Then Fail "Geçmişte yeniden üretilecek satırı seçin."
    r = ActiveCell.Row: ReplayInput = CStr(ws.Cells(r, 8).Value2)
    ValidateArchive ReplayInput
    Set panel = ThisWorkbook.Worksheets("Kontrol Paneli")
    panel.Range("D8").Value2 = ws.Cells(r, 4).Value2: panel.Range("D9").Value2 = ws.Cells(r, 5).Value2: panel.Range("D10").Value2 = ws.Cells(r, 6).Value2
    Set lb = panel.ListBoxes("MRV_Aylar"): chosen = "," & Replace(CStr(ws.Cells(r, 7).Value2), "aylar:", "") & ","
    For i = 1 To 12
        lb.Selected(i) = (InStr(chosen, "," & Format$(i, "00") & ",") > 0)
    Next i
    RaporOlustur: Exit Sub
Bad:
    ReplayInput = "": MsgBox Err.Description, vbExclamation
End Sub

Public Sub ArchiveBegin()
    Dim fso As Object, folder As String, name As String, period As String, wb As Workbook, ws As Worksheet
    Dim r As Long, source As String, dest As String, list As Object, i As Long, output As Object, fileKey As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    OriginalSettings = SettingsPath: OriginalRoot = RootPath
    ArchivePath = "": HistoryRow = 0
    If ReplayInput <> "" Then
        ValidateArchive ReplayInput
        ArchivePath = ReplayInput
    Else
        folder = fso.BuildPath(ThisWorkbook.Path, "Calisma_Arsivi")
        If Not fso.FolderExists(folder) Then fso.CreateFolder folder
        ArchivePath = fso.BuildPath(folder, Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 100)))
        fso.CreateFolder ArchivePath: fso.CreateFolder ArchivePath & "\maliyet"
        Set list = Dict()
        list.Add SettingsPath, "original_settings.xlsx"
        list.Add AbsPath(CStr(Opt("Dosyalar", "kontrol_dosyasi"))), "control.xlsx"
        list.Add AbsPath(CStr(Opt("Dosyalar", "endeks_dosyasi"))), "index.xlsx"
        folder = AbsPath(CStr(Opt("Dosyalar", "maliyet_klasoru")))
        name = Dir$(folder & "\*.xlsx")
        Do While name <> ""
            period = FilenamePeriod(name)
            If Left$(name, 2) <> "~$" And Left$(period, 4) = CStr(CloseYear) And period <= P(CloseYear, CloseMonth) Then list.Add folder & "\" & name, "maliyet\" & name
            name = Dir$()
        Loop
        For Each fileKey In list.Keys
            WorkGate "Girdi kopyalanıyor: " & fso.GetFileName(CStr(fileKey))
            source = CStr(fileKey): dest = ArchivePath & "\" & list(fileKey)
            name = SettingsSignature(source): fso.CopyFile source, dest, False
            If SettingsSignature(dest) <> name Or SettingsSignature(source) <> name Then Fail "Kaynak kopyalanırken değişti: " & source
        Next fileKey
        Set wb = Application.Workbooks.Open(ArchivePath & "\original_settings.xlsx", UpdateLinks:=0, ReadOnly:=False)
        On Error GoTo CopyFailed
        Set ws = wb.Worksheets("Dosya ve Konum Ayarları")
        For r = 5 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If ws.Cells(r, 1).Value2 = "Dosyalar" Then
                Select Case CStr(ws.Cells(r, 2).Value2)
                    Case "maliyet_klasoru": ws.Cells(r, 3).Value2 = "maliyet"
                    Case "kontrol_dosyasi": ws.Cells(r, 3).Value2 = "control.xlsx"
                    Case "endeks_dosyasi": ws.Cells(r, 3).Value2 = "index.xlsx"
                End Select
            End If
        Next r
        wb.SaveAs ArchivePath & "\Ayarlar.xlsx", xlOpenXMLWorkbook: wb.Close False: Set wb = Nothing
        Set output = fso.CreateTextFile(ArchivePath & "\Imzalar.tsv", False, True)
        output.WriteLine "Ayarlar.xlsx" & vbTab & SettingsSignature(ArchivePath & "\Ayarlar.xlsx")
        For Each fileKey In list.Items
            output.WriteLine CStr(fileKey) & vbTab & SettingsSignature(ArchivePath & "\" & fileKey)
        Next fileKey
        output.Close
    End If
    Set ws = WorkspaceSheet("Çalıştırma Geçmişi")
    HistoryRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    RowValues ws, HistoryRow, Array(Now, "Çalışıyor", OriginalSettings, CloseYear, CloseMonth, ReportYear, Signature(), ArchivePath, "")
    SettingsPath = ArchivePath & "\Ayarlar.xlsx": LoadSettings
    Exit Sub
CopyFailed:
    source = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Fail source
End Sub

Public Sub ValidateArchive(ByVal folder As String)
    Dim fso As Object, stream As Object, parts As Variant, line As String, message As String
    Dim expected As Object, filename As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo Bad
    Set expected = Dict()
    Set stream = fso.OpenTextFile(folder & "\Imzalar.tsv", 1, False, -1)
    Do While Not stream.AtEndOfStream
        line = stream.ReadLine: parts = Split(line, vbTab)
        If UBound(parts) <> 1 Then Fail "Arşiv imza listesi geçersiz."
        If InStr(parts(0), "..") > 0 Or InStr(parts(0), ":") > 0 Then Fail "Arşiv yolu geçersiz."
        expected(CStr(parts(0))) = True
        If SettingsSignature(folder & "\" & parts(0)) <> parts(1) Then Fail "Saklanan girdi değişmiş: " & parts(0)
    Loop
    stream.Close
    filename = Dir$(folder & "\maliyet\*.xlsx")
    Do While filename <> ""
        If Not expected.Exists("maliyet\" & filename) Then Fail "Arşive sonradan dosya eklenmiş: " & filename
        filename = Dir$()
    Loop
    Exit Sub
Bad:
    message = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    Fail message
End Sub

Public Sub ArchiveEnd(ByVal success As Boolean, ByVal message As String)
    Dim ws As Worksheet
    If HistoryRow > 0 Then
        Set ws = WorkspaceSheet("Çalıştırma Geçmişi")
        If success Then
            ws.Cells(HistoryRow, 2).Value2 = "Tamamlandı": ws.Cells(HistoryRow, 9).Value2 = RunPath
        Else
            ws.Cells(HistoryRow, 2).Value2 = "Hata / durduruldu": ws.Cells(HistoryRow, 9).Value2 = message
        End If
    End If
    If OriginalSettings <> "" Then SettingsPath = OriginalSettings: RootPath = OriginalRoot
    ReplayInput = "": WorkRunning = False: StopRequested = False
    HistoryRow = 0: OriginalSettings = "": OriginalRoot = ""
End Sub

Public Function OpexEligible(ByVal n As String) As Boolean
    Dim source As Object, a As Variant, b As Variant
    Set source = CurrentSnap(n)
    If source Is Nothing Then Exit Function
    a = SumMonths(source, "AT|", ReportYear): b = SumMonths(source, "AT|", ReportYear - 1)
    OpexEligible = Nonnegative(a) And Positive(b) And Positive(SumDictionary(Revenue, ReportYear)) And Positive(SumDictionary(Revenue, ReportYear - 1))
End Function

Private Function SumDictionary(ByVal source As Object, ByVal year As Long) As Variant
    Dim m As Variant, result As Variant
    result = CDec(0)
    For Each m In Picked.Keys
        result = Plus(result, ValAt(source, P(year, CLng(m))))
    Next m
    SumDictionary = result
End Function

Public Sub HesapIzle()
    Dim ws As Worksheet, wb As Workbook, src As Worksheet, source As Object, actual As Object, budget As Object, cols As Object
    Dim n As Variant, m As Variant, slot As Variant, meta As Variant, mapping As Variant, code As Long, target As String
    Dim r As Long, rawRow As Long, count As Long, first As Long, id As String, value As Variant, filterDept As String, filterCode As String
    Dim security As Long, message As String, name As String, unit As Variant, indexRow As Long, button As Object
    If WorkRunning Then Exit Sub
    WorkRunning = True: StopRequested = False
    security = Application.AutomationSecurity: Application.AutomationSecurity = 3
    On Error GoTo Bad
    FullRead
    Set ws = WorkspaceSheet("Hesap İzleme")
    filterDept = K(ws.Range("D4").Value2): filterCode = Trim$(CStr(ws.Range("D5").Value2))
    ws.Range("A7:L" & Application.Max(8, ws.UsedRange.Rows.Count)).Clear
    ws.Range("A1").Value2 = "Hesap nasıl oluştu? | Kaynak hücreleri"
    ws.Range("A3").Value2 = "D4: bölüm adı (boş = tümü). D5: rapor gider kodu (boş = tümü). Değiştirip İncele'ye basın."
    ws.Range("A4").Value2 = "Bölüm": ws.Range("A5").Value2 = "Gider kodu"
    ws.Range("D4:D5").Interior.Color = RGB(255, 246, 222)
    AddWorkButton ws, "İncele / yenile", "HesapIzle", "G4"
    RowValues ws, 7, Array("Bölüm", "Müdürlük", "Gider kodu", "Gider kalemi", "Tür", "Dönem", "Dosya", "Sayfa", "Hücre", "Değer TL")
    Header ws.Range("A7:J7"): r = 8
    first = CLng(Opt("Maliyet", "ilk_gider_satiri")): count = CLng(Opt("Maliyet", "gider_kalem_sayisi", 27))
    For Each n In Depts.Keys
        If filterDept = "" Or K(n) = filterDept Then
            Set source = CurrentSnap(CStr(n))
            If Not source Is Nothing Then
                WorkGate "Hesap izi: " & Depts(n)(0)
                meta = Depts(n): Set wb = ReadBook(CStr(source("file"))): Set src = wb.Worksheets(CStr(meta(0)))
                CostColumns src, P(CloseYear, CloseMonth), actual, budget
                For code = 1 To count
                    target = CStr(code): rawRow = first + code - 1: name = CStr(src.Cells(rawRow, CLng(Opt("Maliyet", "gider_sutunu"))).Value2)
                    If ExpenseMap.Exists(CStr(code)) Then mapping = ExpenseMap(CStr(code)): target = CStr(mapping(0)): name = CStr(mapping(1))
                    If filterCode = "" Or filterCode = target Then
                        For Each m In Picked.Keys
                            id = P(ReportYear, CLng(m))
                            For Each slot In Array("A", "B")
                                If slot = "A" Or ReportYear = CloseYear Then
                                    If slot = "A" Then Set cols = actual Else Set cols = budget
                                    name = CStr(source("names")(target))
                                    If cols.Exists(id) Then
                                        value = Num(src.Cells(rawRow, cols(id)).Value2)
                                        RowValues ws, r, Array(meta(0), meta(1), CLng(target), name, IIf(slot = "A", "Gerçekleşen", "Bütçe"), id, source("file"), src.Name, src.Cells(rawRow, cols(id)).Address(False, False), value)
                                    Else
                                        RowValues ws, r, Array(meta(0), meta(1), CLng(target), name, slot, id, source("file"), src.Name, "", Null)
                                    End If
                                    r = r + 1
                                End If
                            Next slot
                        Next m
                    End If
                Next code
                wb.Close False: Set wb = Nothing
            End If
        End If
    Next n
    ws.Range("J8:J" & r).NumberFormat = MoneyFmt
    r = r + 2: ws.Cells(r, 1).Value2 = "Gider toplamı yalnız gider satırlarından oluşur; ilk 4 personel ve 5. işlem adedi dahil edilmez.": r = r + 2
    ws.Cells(r, 1).Value2 = "BM = gider / işlem. Endeks = seçili aylardaki yıllık endeks oranlarının ortalaması. İyileşme = (endekslenmiş önceki BM - güncel BM) / endekslenmiş önceki BM.": r = r + 2
    RowValues ws, r, Array("Bölüm", "Gider TL", "İşlem adedi", "BM", "Önceki BM", "Endeks faktörü", "Endekslenmiş BM", "Reel iyileşme", "Müdürlük ağırlığı", "Direktörlük ağırlığı")
    Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)): r = r + 1
    For Each n In Units.Keys
        If filterDept = "" Or K(n) = filterDept Then
            unit = Units(n): RowValues ws, r, Array(Depts(n)(0), unit(0), unit(1), unit(4), unit(5), unit(7), unit(8), unit(9), unit(10), unit(11))
            ws.Range(ws.Cells(r, 8), ws.Cells(r, 10)).NumberFormat = PctFmt: r = r + 1
        End If
    Next n
    r = r + 2: RowValues ws, r, Array("Bölüm", "Ay", "Bu yıl endeks", "Önceki yıl endeks", "Oran")
    Header ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)): r = r + 1
    For Each n In Depts.Keys
        If filterDept = "" Or K(n) = filterDept Then
            meta = Depts(n)
            For Each m In Picked.Keys
                value = ValAt(Indices, K(meta(3)) & Sep & P(ReportYear, CLng(m)))
                unit = ValAt(Indices, K(meta(3)) & Sep & P(ReportYear - 1, CLng(m)))
                RowValues ws, r, Array(meta(0), CLng(m), value, unit, Divide(value, unit)): r = r + 1
            Next m
        End If
    Next n
    ws.Columns("A:J").ColumnWidth = 25: ws.Columns("D").ColumnWidth = 38: ws.Columns("G").ColumnWidth = 55
    ws.Range("A7:J" & r).WrapText = True: ws.Rows("7:" & r).AutoFit: ws.Activate
Done:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    WorkRunning = False: Application.StatusBar = False: Application.AutomationSecurity = security
    On Error GoTo 0
    If message <> "" Then MsgBox message, vbExclamation
    Exit Sub
Bad:
    message = Err.Description: Resume Done
End Sub

Public Sub DuzenOnizle()
    Dim ws As Worksheet, book As Workbook, src As Worksheet, r As Long, dest As Long, n As Variant, source As Object, first As Long, c As Long, j As Long
    Dim security As Long, message As String, x As Range
    If WorkRunning Then Exit Sub
    WorkRunning = True: StopRequested = False
    security = Application.AutomationSecurity: Application.AutomationSecurity = 3
    On Error GoTo Bad
    ReadPanel: LoadSnapshots
    Set ws = WorkspaceSheet("Tablo Düzeni"): ws.Cells.Clear
    ws.Range("A1").Value2 = "Parametrik kaynak düzeni"
    ws.Range("A2").Value2 = "D sütununda yeni değerleri girin. Kaydetmeden önce tüm kaynaklar okunur; ayar yedeği alınır."
    ws.Range("H1").Value2 = SettingsPath: ws.Range("H2").Value2 = SettingsSignature(SettingsPath): ws.Columns("H").Hidden = True
    RowValues ws, 4, Array("Alan", "Parametre", "Mevcut değer", "Yeni değer"): Header ws.Range("A4:D4")
    Set book = ReadBook(SettingsPath): Set src = book.Worksheets("Dosya ve Konum Ayarları"): dest = 5
    For r = 5 To src.Cells(src.Rows.Count, 1).End(xlUp).Row
        If src.Cells(r, 1).Value2 <> "Dosyalar" Then
            RowValues ws, dest, Array(src.Cells(r, 1).Value2, src.Cells(r, 2).Value2, src.Cells(r, 3).Value2, src.Cells(r, 3).Value2): dest = dest + 1
        End If
    Next r
    book.Close False: Set book = Nothing
    ws.Range("D5:D" & dest - 1).Interior.Color = RGB(255, 246, 222)
    ws.Range("J1").Value2 = dest - 1: ws.Columns("J").Hidden = True
    dest = dest + 3: ws.Cells(dest, 1).Value2 = "Kaynak alanı örneği (hücre : değer)": dest = dest + 1
    For Each n In Depts.Keys
        Set source = CurrentSnap(CStr(n))
        If Not source Is Nothing Then
            Set book = ReadBook(CStr(source("file"))): Set src = book.Worksheets(CStr(Depts(n)(0)))
            ws.Cells(dest, 1).Value2 = source("file"): dest = dest + 1
            first = CLng(Opt("Maliyet", "islem_adedi_satiri")) - 4
            Set x = src.Range(CStr(Opt("Maliyet", "tarih_baslangici")))
            For r = first To first + 6
                ws.Cells(dest, 1).Value2 = src.Cells(r, CLng(Opt("Maliyet", "gider_sutunu"))).Address(False, False) & ": " & CStr(src.Cells(r, CLng(Opt("Maliyet", "gider_sutunu"))).Value2)
                For j = 0 To 2
                    c = x.Column + j
                    ws.Cells(dest, j + 2).Value2 = src.Cells(r, c).Address(False, False) & ": " & CStr(src.Cells(r, c).Value2)
                Next j
                dest = dest + 1
            Next r
            book.Close False: Set book = Nothing: Exit For
        End If
    Next n
    AddWorkButton ws, "Düzeni doğrula ve kaydet", "DuzenKaydet", "F4"
    AddWorkButton ws, "Kaynak hücresinden seç", "KaynakHucresindenSec", "F6"
    ws.Columns("A:D").ColumnWidth = 37: ws.UsedRange.WrapText = True: ws.UsedRange.Rows.AutoFit: ws.Activate
Done:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    WorkRunning = False: Application.AutomationSecurity = security: Application.StatusBar = False
    On Error GoTo 0
    If message <> "" Then MsgBox message, vbExclamation
    Exit Sub
Bad:
    message = Err.Description: Resume Done
End Sub

Public Sub DuzenKaydet()
    Dim ws As Worksheet, book As Workbook, src As Worksheet, keys As Object, seen As Object, r As Long, id As String, row As Long
    Dim fso As Object, target As String, temporary As String, rollback As String, backup As String, signature As String, security As Long
    Dim message As String, moved As Boolean, committed As Boolean, originalRootPath As String
    If WorkRunning Then Exit Sub
    WorkRunning = True: StopRequested = False
    security = Application.AutomationSecurity: Application.AutomationSecurity = 3
    On Error GoTo Bad
    ReadPanel: originalRootPath = RootPath
    Set ws = WorkspaceSheet("Tablo Düzeni"): target = SettingsPath: signature = CStr(ws.Range("H2").Value2)
    If ws.Range("H1").Value2 <> target Or SettingsSignature(target) <> signature Then Fail "Ayar değişti; tablo düzenini yeniden okuyun."
    Set fso = CreateObject("Scripting.FileSystemObject")
    temporary = fso.BuildPath(RootPath, "Duzen_" & Format$(Now, "yyyymmddhhnnss") & ".xlsx")
    rollback = temporary & ".rollback": fso.CopyFile target, temporary, False
    Set book = Application.Workbooks.Open(temporary, UpdateLinks:=0, ReadOnly:=False)
    Set src = book.Worksheets("Dosya ve Konum Ayarları"): Set keys = Dict(): Set seen = Dict()
    For r = 5 To src.Cells(src.Rows.Count, 1).End(xlUp).Row
        If src.Cells(r, 1).Value2 <> "Dosyalar" Then keys(CStr(src.Cells(r, 1).Value2) & Sep & CStr(src.Cells(r, 2).Value2)) = r
    Next r
    For r = 5 To CLng(ws.Range("J1").Value2)
        id = CStr(ws.Cells(r, 1).Value2) & Sep & CStr(ws.Cells(r, 2).Value2)
        If Not keys.Exists(id) Or seen.Exists(id) Or ws.Cells(r, 4).HasFormula Then Fail "Düzen parametresi geçersiz veya tekrar ediyor: " & id
        seen(id) = True: src.Cells(keys(id), 3).Value2 = ws.Cells(r, 4).Value2
    Next r
    book.Save: book.Close False: Set book = Nothing
    SettingsPath = temporary: LoadSettings: LoadSnapshots: LoadControl: LoadIndices
    If ItemNames().Count = 0 Then Fail "Yeni düzende gider bulunamadı."
    WorkGate "Tablo düzeni doğrulandı"
    If SettingsSignature(target) <> signature Then Fail "Ayar işlem sırasında değişti."
    backup = fso.BuildPath(originalRootPath, "Ayar_Yedekleri")
    If Not fso.FolderExists(backup) Then fso.CreateFolder backup
    backup = fso.BuildPath(backup, "Duzen_" & Format$(Now, "yyyymmddhhnnss") & ".xlsx")
    fso.CopyFile target, backup, False
    fso.MoveFile target, rollback: moved = True
    fso.MoveFile temporary, target: committed = True
    fso.DeleteFile rollback, True
Done:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If moved And Not committed And Not fso.FileExists(target) Then fso.MoveFile rollback, target
    If temporary <> "" Then
        If fso.FileExists(temporary) Then fso.DeleteFile temporary, True
    End If
    If target <> "" Then SettingsPath = target: LoadSettings
    WorkRunning = False: Application.AutomationSecurity = security: Application.StatusBar = False
    On Error GoTo 0
    If committed Then MsgBox "Tablo düzeni kaydedildi. Yedek: " & backup, vbInformation
    If message <> "" Then MsgBox "Kaydedilemedi: " & message, vbExclamation
    Exit Sub
Bad:
    message = Err.Description: Resume Done
End Sub

Private Function ArchiveSignatures(ByVal folder As String) As Object
    Dim result As Object, file As Object, parts As Variant, message As String
    Set result = Dict()
    On Error GoTo Bad
    Set file = CreateObject("Scripting.FileSystemObject").OpenTextFile(folder & "\Imzalar.tsv", 1, False, -1)
    Do While Not file.AtEndOfStream
        parts = Split(file.ReadLine, vbTab)
        If UBound(parts) = 1 Then
            If parts(0) <> "Ayarlar.xlsx" Then result(CStr(parts(0))) = CStr(parts(1))
        End If
    Loop
    file.Close: Set ArchiveSignatures = result: Exit Function
Bad:
    message = Err.Description
    On Error Resume Next
    If Not file Is Nothing Then file.Close
    On Error GoTo 0
    Fail "Arşiv karşılaştırılamadı: " & message
End Function

Public Sub KaynakFarklari()
    Dim history As Worksheet, ws As Worksheet, current As Object, prior As Object, keys As Object
    Dim r As Long, previous As Long, i As Long, n As Variant, state As String, a As String, b As String
    If WorkRunning Then Exit Sub
    On Error GoTo Bad
    Set history = WorkspaceSheet("Çalıştırma Geçmişi")
    If ActiveSheet.Name <> history.Name Or ActiveCell.Row < 2 Then Fail "Geçmişten bir çalışma satırı seçin."
    r = ActiveCell.Row: Set current = ArchiveSignatures(CStr(history.Cells(r, 8).Value2))
    Set prior = Dict(): Set keys = Dict()
    For i = r - 1 To 2 Step -1
        If history.Cells(i, 3).Value2 = history.Cells(r, 3).Value2 Then previous = i: Exit For
    Next i
    If previous > 0 Then Set prior = ArchiveSignatures(CStr(history.Cells(previous, 8).Value2))
    For Each n In current.Keys
        keys(n) = True
    Next n
    For Each n In prior.Keys
        keys(n) = True
    Next n
    Set ws = WorkspaceSheet("Kaynak Değişiklikleri"): ws.Cells.Clear
    RowValues ws, 1, Array("Girdi", "Durum", "Önceki imza", "Seçili çalışma imzası")
    i = 2
    For Each n In keys.Keys
        a = "": b = "": state = "Aynı"
        If prior.Exists(n) Then a = prior(n)
        If current.Exists(n) Then b = current(n)
        If a = "" Then state = "Eklendi"
        If b = "" Then state = "Kaldırıldı"
        If a <> "" And b <> "" And a <> b Then state = "Değişti"
        RowValues ws, i, Array(n, state, a, b): i = i + 1
    Next n
    Header ws.Range("A1:D1"): ws.Columns("A:D").AutoFit: ws.Activate: Exit Sub
Bad:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub KaynakHucresindenSec()
    Dim target As Worksheet, wb As Workbook, picked As Range, source As String, section As String, code As String
    Dim r As Long, file As Variant, security As Long, message As String, value As Variant
    If WorkRunning Then Exit Sub
    security = Application.AutomationSecurity
    On Error GoTo Bad
    Set target = WorkspaceSheet("Tablo Düzeni")
    If ActiveSheet.Name <> target.Name Then Fail "Tablo Düzeni sayfasında bir parametre satırı seçin."
    r = ActiveCell.Row
    If r < 5 Or r > CLng(target.Range("J1").Value2) Then Fail "Bir düzen parametresi satırı seçin."
    section = CStr(target.Cells(r, 1).Value2): code = CStr(target.Cells(r, 2).Value2)
    If code = "tolerans" Then Fail "Toleransı doğrudan D sütununa yazın."
    file = Application.GetOpenFilename("Excel (*.xlsx),*.xlsx", , "İlgili kaynak Excel dosyasını seçin")
    If VarType(file) = vbBoolean Then Exit Sub
    WorkRunning = True: StopRequested = False
    security = Application.AutomationSecurity: Application.AutomationSecurity = 3
    Set wb = ReadBook(CStr(file))
    On Error Resume Next
    Set picked = Application.InputBox("Kaynak sayfadan hücre veya aralığı seçin: " & section & " / " & code, "Kaynak alanını işaretle", Type:=8)
    On Error GoTo Bad
    If Not picked Is Nothing Then
        If Not (picked.Parent.Parent Is wb) Then Fail "Seçimi açılan kaynak dosyasında yapın."
        If InStr(code, "satiri") > 0 Then
            value = picked.Row
        ElseIf InStr(code, "sutunu") > 0 Then
            value = picked.Column
        ElseIf code = "gider_kalem_sayisi" Then
            value = picked.Rows.Count
        ElseIf code = "sayfa" Then
            value = picked.Parent.Name
        Else
            value = picked.Address(False, False)
        End If
        target.Cells(r, 4).Value2 = value
    End If
Done:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.AutomationSecurity = security: WorkRunning = False: target.Activate
    On Error GoTo 0
    If message <> "" Then MsgBox message, vbExclamation
    Exit Sub
Bad:
    message = Err.Description: Resume Done
End Sub

Public Sub DonemPanelineGit()
    If WorkRunning Then Exit Sub
    ThisWorkbook.Worksheets("Kontrol Paneli").Activate
End Sub
