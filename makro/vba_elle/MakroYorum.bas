Attribute VB_Name = "MakroYorum"
Option Explicit

Public Const PendingLabel As String = "[Kontrol bekliyor — önceki dönemden]"
Private Const ReviewStart As Long = 11

Private Function MemoStore() As Worksheet
    Set MemoStore = WorkspaceSheet("Yorum Kaydý")
    If MemoStore.Range("A1").Value2 = "" Then
        RowValues MemoStore, 1, Array("Ayar yolu", "Kapanýþ", "Ay seçimi", "Bölüm anahtarý", "Gider kodu", "Önceki kapanýþ", "Önceki ay seçimi", "Önceki yorum", "Taslak", "Durum", "Önceki gerçekleþen", "Önceki bütçe", "Güncel gerçekleþen", "Güncel bütçe", "Kaynak", "Son iþlem", "Gider", "Müdürlük")
    End If
    ' V5: onay izi sütunlarý (eski kayýtlara da eklenir).
    If MemoStore.Range("S1").Value2 = "" Then MemoStore.Range("S1:T1").Value2 = Array("Ýþlemi yapan", "Son iþlem türü")
    MemoStore.Visible = xlSheetVeryHidden
End Function

' Açýklama iþlemlerinin silinmeyen kaydý: kim, ne zaman, hangi iþlem. Python: note_review.audit
Private Sub LogMemo(ByVal dept As String, ByVal code As String, ByVal action As String, ByVal text As String)
    Dim ws As Worksheet, r As Long
    Set ws = WorkspaceSheet("Yorum Günlüðü")
    If ws.Range("A1").Value2 = "" Then RowValues ws, 1, Array("Ayar yolu", "Kapanýþ", "Ay seçimi", "Bölüm", "Gider kodu", "Ýþlem", "Ýþlemi yapan", "Zaman", "Metin")
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Range(ws.Cells(r, 9), ws.Cells(r, 9)).NumberFormat = "@"
    RowValues ws, r, Array(MemoDataset(), P(CloseYear, CloseMonth), Signature(), dept, CLng(code), action, CurrentUser(), Format$(Now, "yyyy-mm-dd hh:nn:ss"), text)
    ws.Visible = xlSheetVeryHidden
End Sub

Private Function MemoDataset() As String
    MemoDataset = AbsPath(CStr(ThisWorkbook.Worksheets("Kontrol Paneli").Range("D7").Value2), ThisWorkbook.Path)
End Function

Private Function PrevClose() As String
    Dim d As Date
    d = DateAdd("m", -1, DateSerial(CloseYear, CloseMonth, 1))
    PrevClose = P(Year(d), Month(d))
End Function

Private Function MemoKey(ByVal dept As String, ByVal code As String) As String
    MemoKey = K(dept) & Sep & code
End Function

Private Function CurrentRows(ByVal ws As Worksheet, Optional ByVal closing As String = "", Optional ByVal monthset As String = "") As Object
    Dim result As Object, r As Long, id As String
    Set result = Dict()
    If closing = "" Then closing = P(CloseYear, CloseMonth)
    If monthset = "" Then monthset = Signature()
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(r, 1).Value2 = MemoDataset() And ws.Cells(r, 2).Value2 = closing And ws.Cells(r, 3).Value2 = monthset Then
            id = MemoKey(CanonicalDept(CStr(ws.Cells(r, 4).Value2)), CStr(ws.Cells(r, 5).Value2))
            If result.Exists(id) Then Fail "Yorum kaydýnda mükerrer kimlik var: " & id
            result(id) = r
        End If
    Next r
    Set CurrentRows = result
End Function

Private Sub MemoContext()
    ReadPanel
    If ReportYear <> CloseYear Then Fail "Yorumlar için bütçe yýlýný seçin."
    LoadSnapshots
End Sub

Private Function MemoAmount(ByVal dept As String, ByVal code As String, ByVal slot As String) As Variant
    MemoAmount = ExpenseTotal(SingleScope(CanonicalDept(dept)), slot, ReportYear, code)
End Function

Private Sub NewMemo(ByVal ws As Worksheet, ByVal dept As String, ByVal code As String, ByVal text As String, ByVal oldmonths As String, ByVal olda As Variant, ByVal oldb As Variant, ByVal source As String, Optional ByVal status As String = "Bekliyor")
    Dim r As Long, names As Object, meta As Variant
    Set names = ItemNames(): meta = Depts(dept)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Range(ws.Cells(r, 8), ws.Cells(r, 9)).NumberFormat = "@"
    RowValues ws, r, Array(MemoDataset(), P(CloseYear, CloseMonth), Signature(), dept, CLng(code), IIf(status = "Bu dönem", P(CloseYear, CloseMonth), PrevClose()), oldmonths, text, text, status, olda, oldb, MemoAmount(dept, code, "A"), MemoAmount(dept, code, "B"), source, Format$(Now, "yyyy-mm-dd hh:nn:ss"), names(code), meta(1))
    ws.Cells(r, 8).NumberFormat = "@": ws.Cells(r, 9).NumberFormat = "@"
    ws.Cells(r, 8).Value2 = text: ws.Cells(r, 9).Value2 = text
    If status = "Bu dönem" Then
        ws.Cells(r, 20).Value2 = "Kaydedildi: " & source
    Else
        ws.Cells(r, 20).Value2 = "Önceki dönemden getirildi (" & PrevClose() & ")"
    End If
    ws.Cells(r, 19).Value2 = CurrentUser()
    LogMemo dept, code, CStr(ws.Cells(r, 20).Value2), text
End Sub

Public Sub YorumGetir()
    Dim store As Worksheet, existing As Object, candidates As Object, ambiguous As Object, r As Long, id As String, dept As String, code As String, created As Long
    Dim security As Long, message As String, sourceRow As Variant, names As Object
    If WorkRunning Then Exit Sub
    security = Application.AutomationSecurity: Application.AutomationSecurity = 3
    WorkRunning = True: StopRequested = False
    On Error GoTo Bad
    YorumKaydedilmemisKontrol
    MemoContext
    Set store = MemoStore(): Set existing = CurrentRows(store): Set candidates = Dict(): Set ambiguous = Dict(): Set names = ItemNames()
    For r = 2 To store.Cells(store.Rows.Count, 1).End(xlUp).Row
        If store.Cells(r, 1).Value2 = MemoDataset() And store.Cells(r, 2).Value2 = PrevClose() And (store.Cells(r, 10).Value2 = "Onaylý" Or store.Cells(r, 10).Value2 = "Bu dönem") Then
            dept = CanonicalDept(CStr(store.Cells(r, 4).Value2)): code = CStr(store.Cells(r, 5).Value2): id = MemoKey(dept, code)
            If candidates.Exists(id) Then
                If store.Cells(r, 3).Value2 = Signature() Then
                    candidates(id) = r: ambiguous(id) = False
                ElseIf store.Cells(candidates(id), 3).Value2 <> Signature() Then
                    ambiguous(id) = True
                End If
            Else
                candidates(id) = r: ambiguous(id) = False
            End If
        End If
    Next r
    For Each sourceRow In candidates.Keys
        id = CStr(sourceRow): r = candidates(id): dept = CanonicalDept(CStr(store.Cells(r, 4).Value2)): code = CStr(store.Cells(r, 5).Value2)
        If Not ambiguous(id) And Not existing.Exists(id) And Depts.Exists(dept) And names.Exists(code) Then
            If names(code) <> CStr(store.Cells(r, 17).Value2) Then Fail "Gider adý deðiþmiþ; önceki rapor dosyasýndan eþleþtirerek yükleyin."
            NewMemo store, dept, code, CStr(store.Cells(r, 9).Value2), CStr(store.Cells(r, 3).Value2), Num(store.Cells(r, 13).Value2), Num(store.Cells(r, 14).Value2), "Yerel yorum geçmiþi"
            created = created + 1
        End If
    Next sourceRow
    WorkRunning = False: YorumListele
    MsgBox created & " geçmiþ yorum getirildi. Mevcut kayýtlar korundu. Yerel geçmiþte yoksa Önceki dosyalarý seç düðmesini kullanýn. Birden fazla ay seçimi varsa ilgili dosyayý seçin.", vbInformation
Done:
    WorkRunning = False: Application.AutomationSecurity = security: Application.StatusBar = False
    If message <> "" Then MsgBox message, vbExclamation
    Exit Sub
Bad:
    message = Err.Description: Resume Done
End Sub

Public Sub YorumDosyadanGetir()
    ImportMemoFiles False
End Sub

Public Sub YorumMevcutYukle()
    ImportMemoFiles True
End Sub

Private Sub ImportMemoFiles(ByVal current As Boolean)
    Dim files As Variant, file As Variant, book As Workbook, src As Worksheet, meta As Worksheet, rows As Object, cols As Object, noteCols As Object
    Dim store As Worksheet, existing As Object, staged As Object, row As Variant, item As Variant, dept As Variant, mapped As String, id As String, text As String
    Dim period As String, monthset As String, commonMonths As String, names As Object, currentItem As String, r As Long, count As Long, skipped As Long, unknown As Long
    Dim security As Long, message As String, status As String, a As Variant, b As Variant, backup As Variant, touched As Boolean, sourcePending As Boolean
    If WorkRunning Then Exit Sub
    security = Application.AutomationSecurity
    If MemoDirty() Then MsgBox "Önce düzenlenen taslaklarý kaydedin veya onaylayýn.", vbInformation: Exit Sub
    files = Application.GetOpenFilename("Excel raporu (*.xlsx),*.xlsx", , IIf(current, "Bu dönemin açýklamalý müdürlük dosyalarý", "Bir önceki kapanýþýn açýklamalý müdürlük dosyalarý"), MultiSelect:=True)
    If VarType(files) = vbBoolean Then Exit Sub
    WorkRunning = True: StopRequested = False: Application.AutomationSecurity = 3
    On Error GoTo Bad
    MemoContext
    Set store = MemoStore(): Set existing = CurrentRows(store): Set staged = Dict(): Set names = ItemNames()
    For Each file In files
        WorkGate "Yorum dosyasý: " & CStr(file)
        Set book = ReadBook(CStr(file)): Set meta = Nothing
        On Error Resume Next
        Set meta = book.Worksheets("_Rapor Kimliði")
        On Error GoTo Bad
        If meta Is Nothing Then Fail "Rapor kimliði yok: " & book.Name & ". Beklenen senaryo çalýþma dosyasýný seçmeyin; açýklamalý müdürlük raporlarýný seçin."
        If meta.Range("A1").Value2 <> "ButceSapmasi-1" Or meta.Range("B1").Value2 <> "Müdürlük" Then Fail "Açýklamalý müdürlük raporu bekleniyor."
        period = CStr(meta.Range("D1").Value2): monthset = CStr(meta.Range("E1").Value2)
        If current Then
            If period <> P(CloseYear, CloseMonth) Or monthset <> Signature() Then Fail "Bu dönem raporunun kapanýþ/ay seçimi farklý."
        Else
            If period <> PrevClose() Or Left$(CStr(meta.Range("C1").Value2), 4) <> Left$(period, 4) Then Fail "Yalnýz önceki kapanýþ raporlarý seçilebilir: " & PrevClose()
        End If
        If commonMonths <> "" And commonMonths <> monthset Then Fail "Seçilen raporlarýn ay kümeleri farklý."
        commonMonths = monthset
        Set src = ExpensePositions(book, rows, cols, noteCols)
        For Each dept In cols.Keys
            mapped = CanonicalDept(CStr(dept))
            For Each item In rows.Keys
                If noteCols.Exists(dept) Then
                    r = rows(item)
                    If src.Cells(r, noteCols(dept)).HasFormula Or IsError(src.Cells(r, noteCols(dept)).Value2) Then Fail "Açýklama sabit metin olmalý."
                    text = Trim$(CStr(src.Cells(r, noteCols(dept)).Value2))
                    sourcePending = (InStr(text, "[Kontrol bekliyor") = 1)
                    If sourcePending And InStr(text, vbLf) > 0 Then text = Mid$(text, InStr(text, vbLf) + 1)
                    If text <> "" Then
                        If Depts.Exists(mapped) And names.Exists(item) Then
                            If names(item) <> CStr(src.Cells(r, 3).Value2) Then Fail "Gider adý farklý: " & item
                            If Len(text) > 32000 Then Fail "Açýklama en çok 32.000 karakter olmalý."
                            a = Num(src.Cells(r, cols(dept)).Value2): b = Num(src.Cells(r, cols(dept) + 1).Value2)
                            If current Then
                                If Changed(a, MemoAmount(mapped, CStr(item), "A")) Or Changed(b, MemoAmount(mapped, CStr(item), "B")) Then Fail "Bu dönem rapor tutarý farklý: " & mapped & " / " & item
                            End If
                            id = MemoKey(mapped, CStr(item))
                            If staged.Exists(id) Then
                                row = staged(id)
                                If row(2) <> text Then Fail "Ayný bölüm/gidere farklý açýklamalar seçildi; hiçbir kayýt yüklenmedi."
                            Else
                                staged.Add id, Array(mapped, CStr(item), text, monthset, a, b, CStr(file), sourcePending)
                            End If
                        Else
                            unknown = unknown + 1
                        End If
                    End If
                End If
            Next item
        Next dept
        book.Close False: Set book = Nothing
    Next file
    backup = store.UsedRange.Value2
    For Each item In staged.Keys
        row = staged(item)
        If existing.Exists(item) Then
            skipped = skipped + 1
        Else
            status = "Bekliyor"
            If current And Not CBool(row(7)) Then status = "Bu dönem"
            touched = True
            NewMemo store, CStr(row(0)), CStr(row(1)), CStr(row(2)), CStr(row(3)), row(4), row(5), CStr(row(6)), status
            count = count + 1
        End If
    Next item
    WorkRunning = False: YorumListele
    If Not SaveController() Then MsgBox "Yorumlar bellekte; Ctrl+S ile kaydedin.", vbInformation
    MsgBox count & " yorum getirildi; " & skipped & " mevcut kayýt korundu; " & unknown & " eþleþmeyen kayýt alýnmadý.", vbInformation
Done:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    WorkRunning = False: Application.AutomationSecurity = security: Application.StatusBar = False
    On Error GoTo 0
    If message <> "" Then MsgBox message, vbExclamation
    Exit Sub
Bad:
    message = Err.Description
    On Error Resume Next
    If touched Then store.Cells.Clear: store.Range("A1").Resize(UBound(backup, 1), UBound(backup, 2)).Value2 = backup
    On Error GoTo 0
    Resume Done
End Sub

Private Function MemoToken(ByVal store As Worksheet, ByVal r As Long, ByVal a As Variant, ByVal b As Variant) As String
    Dim text As String, i As Long, j As Long, x As Double, y As Double, ch As Long
    For i = 1 To 18
        text = text & CStr(store.Cells(r, i).Value2) & ChrW$(30)
    Next i
    If Not IsNull(a) Then text = text & CStr(a)
    text = text & ChrW$(30)
    If Not IsNull(b) Then text = text & CStr(b)
    x = 1: y = 0
    For j = 1 To Len(text)
        ch = AscW(Mid$(text, j, 1)) And &HFFFF&
        x = (x + ch) Mod 65521: y = (y + x) Mod 65521
    Next j
    MemoToken = CStr(Len(text)) & ":" & CStr(x) & ":" & CStr(y)
End Function

Public Sub YorumListele()
    Dim store As Worksheet, ws As Worksheet, records As Object, id As Variant, r As Long, out As Long, dept As String, code As String, meta As Variant
    Dim a As Variant, b As Variant, changedText As String, security As Long, message As String, names As Object
    Dim buttonActions As Variant, button As Object, i As Long
    If WorkRunning Then Exit Sub
    If MemoDirty() Then
        If MsgBox("Kaydedilmemiþ taslaklar var. Kaydetmeden liste yenilensin mi?", vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then Exit Sub
    End If
    security = Application.AutomationSecurity: Application.AutomationSecurity = 3
    WorkRunning = True: StopRequested = False
    On Error GoTo Bad
    MemoContext
    Set store = MemoStore(): Set records = CurrentRows(store): Set ws = WorkspaceSheet("Yorum Ýnceleme"): Set names = ItemNames()
    ws.Cells.Clear
    ws.Range("A1").Value2 = "Önceki ayýn yorumlarý | Toplu inceleme"
    ws.Range("A2").Value2 = P(CloseYear, CloseMonth) & " | " & Signature() & ". A sütununa Evet yazýn. G sütununu doðrudan düzenleyin; seçileni topluca onaylayýn. Mevcut bu dönem yorumlarý korunur."
    ws.Range("B3:M4").Merge: ws.Range("B3:M4").NumberFormat = "@": ws.Range("B3").Value2 = "": ws.Range("B3:M4").Interior.Color = RGB(255, 246, 222)
    ws.Range("A3").Value2 = "Ortak metin:"
    AddWorkButton ws, "Önceki dosyalarý seç", "YorumDosyadanGetir", "A6"
    AddWorkButton ws, "Görünürleri seç", "YorumGorunurSec", "D6"
    AddWorkButton ws, "Seçileni onayla", "YorumOnayla", "G6"
    AddWorkButton ws, "Seçileni geçersiz say", "YorumGecersiz", "J6"
    AddWorkButton ws, "Taslaklarý kaydet", "YorumTaslakKaydet", "A8"
    AddWorkButton ws, "Ortak metni seçilene uygula", "YorumOrtakMetin", "D8"
    AddWorkButton ws, "Bekleyenleri göster", "YorumBekleyenGoster", "G8"
    AddWorkButton ws, "Tümünü göster", "YorumHepsiniGoster", "J8"
    AddWorkButton ws, "Bu dönem yorumlarýný yükle", "YorumMevcutYukle", "J1"
    RowValues ws, 10, Array("Seç (Evet)", "Müdürlük", "Bölüm", "Gider kodu", "Alt gider", "Önceki yorum", "Bu dönem taslaðý", "Önceki kapanýþ", "Önceki aylar", "Durum", "Veri kontrolü", "Gerçekleþen TL", "Bütçe TL", "Kimlik", "Kontrol imzasý", "Açýlan taslak")
    Header ws.Range("A10:M10"): out = ReviewStart
    For Each id In records.Keys
        r = records(id): dept = CanonicalDept(CStr(store.Cells(r, 4).Value2)): code = CStr(store.Cells(r, 5).Value2)
        If Depts.Exists(dept) And names.Exists(code) Then
            meta = Depts(dept): a = MemoAmount(dept, code, "A"): b = MemoAmount(dept, code, "B")
            changedText = "Tutarlar ayný"
            If Changed(a, Num(store.Cells(r, 11).Value2)) Or Changed(b, Num(store.Cells(r, 12).Value2)) Then changedText = "Tutar deðiþmiþ"
            If store.Cells(r, 7).Value2 <> Signature() Then changedText = "Ay seçimi farklý; doðrudan kýyaslanamaz"
            ws.Range(ws.Cells(out, 6), ws.Cells(out, 7)).NumberFormat = "@"
            ws.Cells(out, 16).NumberFormat = "@"
            RowValues ws, out, Array("", meta(1), meta(0), CLng(code), names(code), store.Cells(r, 8).Value2, store.Cells(r, 9).Value2, store.Cells(r, 6).Value2, store.Cells(r, 7).Value2, store.Cells(r, 10).Value2, changedText, a, b, CStr(id), MemoToken(store, r, a, b), store.Cells(r, 9).Value2)
            ws.Cells(out, 6).NumberFormat = "@": ws.Cells(out, 7).NumberFormat = "@"
            ws.Cells(out, 6).Font.Color = RGB(113, 129, 142)
            If store.Cells(r, 10).Value2 = "Bekliyor" Then ws.Cells(out, 7).Font.Color = RGB(190, 48, 58)
            If store.Cells(r, 10).Value2 = "Bu dönem" Then ws.Range(ws.Cells(out, 1), ws.Cells(out, 13)).Interior.Color = RGB(239, 242, 243)
            out = out + 1
        End If
    Next id
    ws.Range("U1").Value2 = SettingsPath: ws.Range("U2").Value2 = SettingsSignature(SettingsPath)
    ws.Range("U3").Value2 = P(CloseYear, CloseMonth): ws.Range("U4").Value2 = Signature()
    ws.Columns("N:U").Hidden = True
    ws.Columns("A").ColumnWidth = 8: ws.Columns("B").ColumnWidth = 14: ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 8: ws.Columns("E").ColumnWidth = 25
    ws.Columns("F:G").ColumnWidth = 42: ws.Columns("H").ColumnWidth = 14: ws.Columns("I").ColumnWidth = 20
    ws.Columns("J:K").ColumnWidth = 25: ws.Columns("L:M").ColumnWidth = 20
    ws.Range("L11:M" & Application.Max(11, out)).NumberFormat = MoneyFmt
    ws.Range("A10:M" & Application.Max(11, out)).WrapText = True: ws.Rows("10:" & Application.Max(11, out)).AutoFit
    ws.Rows("6:8").RowHeight = 27
    ws.Range("A2:M2").WrapText = True: ws.Rows(2).RowHeight = 40
    buttonActions = Array("YorumDosyadanGetir", "YorumGorunurSec", "YorumOnayla", "YorumGecersiz", "YorumTaslakKaydet", "YorumOrtakMetin", "YorumBekleyenGoster", "YorumHepsiniGoster")
    For i = 0 To 7
        Set button = ws.Shapes("MC_" & buttonActions(i))
        button.Left = ws.Range("A6").Left + (i Mod 4) * 215
        button.Top = ws.Cells(IIf(i < 4, 6, 8), 1).Top: button.Width = 200: button.Height = 25
    Next i
    Set button = ws.Shapes("MC_YorumMevcutYukle")
    button.Left = ws.Range("A1").Left + 690: button.Top = ws.Range("A1").Top: button.Width = 200: button.Height = 25
    ws.Rows(1).RowHeight = 30
    ws.Activate: ActiveWindow.FreezePanes = False: ws.Range("F11").Select: ActiveWindow.FreezePanes = True
Done:
    WorkRunning = False: Application.AutomationSecurity = security: Application.StatusBar = False
    If message <> "" Then MsgBox message, vbExclamation
    Exit Sub
Bad:
    message = Err.Description: Resume Done
End Sub

Public Sub YorumGorunurSec()
    Dim ws As Worksheet, r As Long
    If WorkRunning Then Exit Sub
    Set ws = WorkspaceSheet("Yorum Ýnceleme")
    For r = ReviewStart To ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
        If Not ws.Rows(r).Hidden And ws.Cells(r, 10).Value2 <> "Bu dönem" Then ws.Cells(r, 1).Value2 = "Evet"
    Next r
End Sub

Public Sub YorumBekleyenGoster()
    Dim ws As Worksheet, r As Long
    If WorkRunning Then Exit Sub
    Set ws = WorkspaceSheet("Yorum Ýnceleme")
    For r = ReviewStart To ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
        ws.Rows(r).Hidden = (ws.Cells(r, 10).Value2 <> "Bekliyor")
        ws.Cells(r, 1).Value2 = ""
    Next r
End Sub

Public Sub YorumHepsiniGoster()
    If WorkRunning Then Exit Sub
    WorkspaceSheet("Yorum Ýnceleme").Rows.Hidden = False
End Sub

Public Sub YorumOrtakMetin()
    Dim ws As Worksheet, r As Long, text As String
    If WorkRunning Then Exit Sub
    Set ws = WorkspaceSheet("Yorum Ýnceleme"): text = CStr(ws.Range("B3").Value2)
    If Trim$(text) = "" Then MsgBox "B3 alanýna ortak metni yazýn.", vbInformation: Exit Sub
    For r = ReviewStart To ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
        If Not ws.Rows(r).Hidden And K(ws.Cells(r, 1).Value2) = "evet" And ws.Cells(r, 10).Value2 <> "Bu dönem" Then
            ws.Cells(r, 7).NumberFormat = "@": ws.Cells(r, 7).Value2 = text
            ws.Cells(r, 7).Font.Color = RGB(190, 48, 58)
        End If
    Next r
End Sub

Public Sub YorumOnayla()
    ApplyMemos "Onaylý"
End Sub
Public Sub YorumGecersiz()
    ApplyMemos "Geçersiz"
End Sub
Public Sub YorumTaslakKaydet()
    ApplyMemos "Bekliyor"
End Sub

Private Sub ApplyMemos(ByVal action As String)
    Dim ws As Worksheet, store As Worksheet, records As Object, staged As Object, id As String, r As Long, target As Long, text As String
    Dim a As Variant, b As Variant, entry As Variant, n As Variant, security As Long, message As String, backup As Variant, touched As Boolean
    If WorkRunning Then Exit Sub
    security = Application.AutomationSecurity: Application.AutomationSecurity = 3
    WorkRunning = True: StopRequested = False
    On Error GoTo Bad
    MemoContext
    Set ws = WorkspaceSheet("Yorum Ýnceleme"): Set store = MemoStore(): Set records = CurrentRows(store): Set staged = Dict()
    If ws.Range("U1").Value2 <> SettingsPath Or ws.Range("U2").Value2 <> SettingsSignature(SettingsPath) Or ws.Range("U3").Value2 <> P(CloseYear, CloseMonth) Or ws.Range("U4").Value2 <> Signature() Then Fail "Dönem, ay seçimi veya ayar deðiþti. Listeyi yenileyin; hiçbir yorum deðiþtirilmedi."
    For r = ReviewStart To ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
        If (action = "Bekliyor" And ws.Cells(r, 7).Value2 <> ws.Cells(r, 16).Value2) Or (action <> "Bekliyor" And Not ws.Rows(r).Hidden And K(ws.Cells(r, 1).Value2) = "evet") Then
            id = CStr(ws.Cells(r, 14).Value2)
            If Not records.Exists(id) Or staged.Exists(id) Then Fail "Yorum kimliði geçersiz; listeyi yenileyin."
            target = records(id)
            If store.Cells(target, 10).Value2 = "Bu dönem" Then Fail "Mevcut bu dönem açýklamasý korunuyor."
            a = MemoAmount(CStr(store.Cells(target, 4).Value2), CStr(store.Cells(target, 5).Value2), "A")
            b = MemoAmount(CStr(store.Cells(target, 4).Value2), CStr(store.Cells(target, 5).Value2), "B")
            If MemoToken(store, target, a, b) <> CStr(ws.Cells(r, 15).Value2) Then Fail "Yorum veya tutarlar deðiþti. Listeyi yenileyin; hiçbir yorum deðiþtirilmedi."
            If ws.Cells(r, 7).HasFormula Or IsError(ws.Cells(r, 7).Value2) Then Fail "Açýklama sabit metin olmalý."
            text = CStr(ws.Cells(r, 7).Value2)
            If Len(text) > 32000 Then Fail "Açýklama en çok 32.000 karakter olmalý."
            If action = "Onaylý" And (Trim$(text) = "" Or IsNull(a) Or IsNull(b)) Then Fail "Boþ açýklama veya eksik tutar onaylanamaz."
            If action = "Bekliyor" And store.Cells(target, 10).Value2 = "Onaylý" Then Fail "Onaylý metni düzenleyip yeniden onaylayýn."
            staged.Add id, Array(target, text, a, b)
        End If
    Next r
    If staged.Count = 0 Then Fail "Ýþlem yapýlacak satýr seçin veya taslak düzenleyin."
    backup = store.UsedRange.Value2
    For Each n In staged.Keys
        entry = staged(n): target = entry(0): touched = True
        store.Cells(target, 9).NumberFormat = "@": store.Cells(target, 9).Value2 = entry(1)
        store.Cells(target, 10).Value2 = action
        SetCell store.Cells(target, 13), entry(2): SetCell store.Cells(target, 14), entry(3)
        store.Cells(target, 16).Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        store.Cells(target, 19).Value2 = CurrentUser()
        Select Case action
            Case "Onaylý": store.Cells(target, 20).Value2 = "Onaylandý"
            Case "Geçersiz": store.Cells(target, 20).Value2 = "Geçersiz sayýldý"
            Case Else: store.Cells(target, 20).Value2 = "Taslak kaydedildi"
        End Select
        LogMemo CStr(store.Cells(target, 4).Value2), CStr(store.Cells(target, 5).Value2), CStr(store.Cells(target, 20).Value2), CStr(entry(1))
    Next n
    ' Keep other unsaved grid drafts in place; only committed rows get fresh tokens.
    For r = ReviewStart To ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
        id = CStr(ws.Cells(r, 14).Value2)
        If staged.Exists(id) Then
            entry = staged(id): target = entry(0): ws.Cells(r, 10).Value2 = action
            ws.Cells(r, 15).Value2 = MemoToken(store, target, entry(2), entry(3)): ws.Cells(r, 16).Value2 = entry(1): ws.Cells(r, 1).Value2 = ""
            If action = "Bekliyor" Then ws.Cells(r, 7).Font.Color = RGB(190, 48, 58) Else ws.Cells(r, 7).Font.Color = RGB(35, 55, 70)
        End If
    Next r
    If Not SaveController() Then MsgBox "Yorumlar bellekte; Ctrl+S ile kaydedin.", vbInformation
    MsgBox staged.Count & " yorum iþlendi. Diðer taslaklar korundu.", vbInformation
Done:
    WorkRunning = False: Application.AutomationSecurity = security: Application.StatusBar = False
    If message <> "" Then MsgBox message, vbExclamation
    Exit Sub
Bad:
    message = Err.Description
    On Error Resume Next
    If touched Then store.Cells.Clear: store.Range("A1").Resize(UBound(backup, 1), UBound(backup, 2)).Value2 = backup
    On Error GoTo 0
    Resume Done
End Sub

Public Sub WriteMemoReport(ByVal wb As Workbook, ByVal scope As Object, ByVal general As Boolean)
    Dim store As Worksheet, records As Object, audit As Worksheet, summary As Worksheet, exp As Worksheet, rows As Object, cols As Object, noteCols As Object
    Dim code As Variant, dept As Variant, names As Object, id As String, r As Long, out As Long, line As String, combined As String, pending As Boolean, isPending As Boolean
    Dim a As Variant, b As Variant, state As String, text As String, hasComment As Boolean
    If ReportYear <> CloseYear Or ReplayInput <> "" Then Exit Sub
    Set store = MemoStore(): Set records = CurrentRows(store)
    If records.Count = 0 Then Exit Sub
    Set names = ItemNames()
    For Each dept In scope.Keys
        For Each code In names.Keys
            id = MemoKey(CStr(dept), CStr(code))
            If records.Exists(id) Then
                r = records(id)
                If store.Cells(r, 10).Value2 <> "Geçersiz" And Trim$(CStr(store.Cells(r, 9).Value2)) <> "" Then hasComment = True
            End If
        Next code
    Next dept
    If Not hasComment Then Exit Sub
    Set audit = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): audit.Name = "Açýklama Durumu"
    RowValues audit, 1, Array("Bölüm", "Gider kodu", "Gider", "Durum", "Önceki kapanýþ", "Önceki ay seçimi", "Kaynak", "Açýklama", "Son iþlem", "Ýþlemi yapan", "Ýþlem zamaný")
    Set exp = ExpensePositions(wb, rows, cols, noteCols): Set names = ItemNames()
    If general Then
        Set summary = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): summary.Name = "Alt Gider Açýklamalarý"
        RowValues summary, 1, Array("Ana gider", "Alt gider kodu", "Alt gider", "Bölüm açýklamalarý")
    End If
    out = 2
    For Each code In names.Keys
        combined = "": pending = False
        For Each dept In scope.Keys
            id = MemoKey(CStr(dept), CStr(code))
            If records.Exists(id) Then
                r = records(id): state = CStr(store.Cells(r, 10).Value2): text = CStr(store.Cells(r, 9).Value2)
                If state <> "Geçersiz" And Trim$(text) <> "" Then
                    a = MemoAmount(CStr(dept), CStr(code), "A"): b = MemoAmount(CStr(dept), CStr(code), "B")
                    isPending = (state = "Bekliyor" Or Changed(a, Num(store.Cells(r, 13).Value2)) Or Changed(b, Num(store.Cells(r, 14).Value2)))
                    line = text
                    If isPending Then
                        line = PendingLabel & " " & CStr(store.Cells(r, 6).Value2) & " | " & CStr(store.Cells(r, 7).Value2) & vbLf & text
                        If state = "Bu dönem" Then line = "[Kontrol bekliyor — veri deðiþti] " & P(CloseYear, CloseMonth) & vbLf & text
                    End If
                    If Len(line) > 32767 Then Fail "Açýklama hücre sýnýrýný aþýyor."
                    If general Then
                        If combined <> "" Then combined = combined & vbLf
                        combined = combined & Depts(dept)(0) & ": " & line
                        If isPending Then pending = True
                    Else
                        If rows.Exists(code) And noteCols.Exists(dept) Then
                            exp.Cells(rows(code), noteCols(dept)).NumberFormat = "@": exp.Cells(rows(code), noteCols(dept)).Value2 = line
                            If isPending Then exp.Cells(rows(code), noteCols(dept)).Font.Color = RGB(190, 48, 58)
                        End If
                    End If
                    audit.Cells(out, 8).NumberFormat = "@"
                    RowValues audit, out, Array(Depts(dept)(0), CLng(code), names(code), IIf(isPending, "Kontrol bekliyor", "Onaylý / bu dönem"), store.Cells(r, 6).Value2, store.Cells(r, 7).Value2, store.Cells(r, 15).Value2, text)
                    audit.Cells(out, 8).NumberFormat = "@": audit.Cells(out, 8).Value2 = text
                    RowValues audit, out, Array(store.Cells(r, 20).Value2, store.Cells(r, 19).Value2, store.Cells(r, 16).Value2), 9
                    If isPending Then audit.Rows(out).Font.Color = RGB(190, 48, 58)
                    out = out + 1
                End If
            End If
        Next dept
        If general And combined <> "" Then
            If Len(combined) > 32767 Then Fail "Birleþik açýklama hücre sýnýrýný aþýyor."
            r = summary.Cells(summary.Rows.Count, 1).End(xlUp).Row + 1
            summary.Cells(r, 4).NumberFormat = "@"
            RowValues summary, r, Array(GroupOf(CLng(code)), CLng(code), names(code), combined)
            summary.Cells(r, 4).NumberFormat = "@": summary.Cells(r, 4).Value2 = combined
            If pending Then summary.Cells(r, 4).Font.Color = RGB(190, 48, 58)
        End If
    Next code
    Header audit.Range("A1:K1"): audit.Columns("A:G").ColumnWidth = 24: audit.Columns("H").ColumnWidth = 80: audit.Columns("I:K").ColumnWidth = 24: audit.UsedRange.WrapText = True: audit.UsedRange.Rows.AutoFit
    If general Then
        Header summary.Range("A1:D1"): summary.Columns("A:C").ColumnWidth = 30: summary.Columns("D").ColumnWidth = 90
        summary.UsedRange.WrapText = True: summary.UsedRange.Rows.AutoFit
    End If
End Sub

Private Function MemoDirty() As Boolean
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Yorum Ýnceleme")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For r = ReviewStart To ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
        If IsError(ws.Cells(r, 7).Value2) Or ws.Cells(r, 7).HasFormula Then MemoDirty = True: Exit Function
        If CStr(ws.Cells(r, 7).Value2) <> CStr(ws.Cells(r, 16).Value2) Then MemoDirty = True: Exit Function
    Next r
End Function

Public Sub YorumKaydedilmemisKontrol()
    If MemoDirty() Then Fail "Kaydedilmemiþ yorum taslaklarý var. Önce taslaklarý kaydedin veya onaylayýn."
End Sub
