Attribute VB_Name = "MakroAciklama"
Option Explicit

Private Function Identity(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Set ws = wb.Worksheets("_Rapor Kimliği")
    If ws.Range("A1").Value2 <> "ButceSapmasi-1" Then Fail "Rapor kimliği tanınmadı: " & wb.Name
    Identity = CStr(ws.Range("C1").Value2) & Sep & CStr(ws.Range("D1").Value2) & Sep & CStr(ws.Range("E1").Value2)
End Function

Public Function ExpensePositions(ByVal wb As Workbook, ByRef rows As Object, ByRef columns As Object, ByRef notes As Object) As Worksheet
    Dim ws As Worksheet, r As Long, c As Long, endCol As Long, n As String, code As String, title As String
    Set rows = Dict(): Set columns = Dict(): Set notes = Dict()
    Set ws = wb.Worksheets("Giderler")
    For r = 6 To ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
        If Len(CStr(ws.Cells(r, 2).Value2)) > 0 And IsNumeric(ws.Cells(r, 2).Value2) Then
            code = CStr(CLng(ws.Cells(r, 2).Value2))
            If rows.Exists(code) Then Fail "Mükerrer gider kodu: " & wb.Name
            rows.Add code, r
        End If
    Next r
    endCol = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column
    For c = 4 To endCol
        title = K(ws.Cells(5, c).Value2)
        If InStr(title, "gerceklesen") > 0 Then
            n = CanonicalDept(CStr(ws.Cells(4, c).MergeArea.Cells(1, 1).Value2))
            If n = "" Or columns.Exists(n) Then Fail "Geçersiz bölüm başlığı: " & wb.Name
            columns.Add n, c
        ElseIf title = "aciklama" Then
            n = CanonicalDept(CStr(ws.Cells(4, c).MergeArea.Cells(1, 1).Value2))
            If notes.Exists(n) Then Fail "Mükerrer açıklama alanı: " & wb.Name
            notes.Add n, c
        End If
    Next c
    Set ExpensePositions = ws
End Function

Public Sub AciklamaBirlestir()
    Dim generalPath As Variant, fd As Object, file As Variant, general As Workbook, manager As Workbook
    Dim gs As Worksheet, ms As Worksheet, gr As Object, gc As Object, gn As Object, mr As Object, mc As Object, mn As Object
    Dim merged As Object, names As Object, code As Variant, dept As Variant, id As String, text As String, signature As String
    Dim r As Long, c As Long, i As Long, value As Variant, note As String, errText As String, target As Variant
    Dim output As Workbook, ws As Worksheet, display As Boolean, security As Long, didSave As Boolean, accepted As Object
    Dim frozen As Boolean, savePath As String, meta As Worksheet, exists As Boolean, fso As Object
    If WorkRunning Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    display = Application.DisplayAlerts: security = Application.AutomationSecurity
    On Error GoTo Bad
    generalPath = Application.GetOpenFilename("Excel raporu (*.xlsx),*.xlsx", , "Açıklamalar eklenecek direktörlük raporunu seçin")
    If VarType(generalPath) = vbBoolean Then Exit Sub
    Set fd = Application.FileDialog(3): fd.AllowMultiSelect = True
    fd.Title = "Açıklama yazılan müdürlük raporlarını seçin"
    fd.Filters.Clear: fd.Filters.Add "Excel", "*.xlsx"
    If fd.Show <> -1 Then Exit Sub
    Application.AutomationSecurity = 3: Application.DisplayAlerts = False
    Set general = ReadBook(CStr(generalPath)): signature = Identity(general)
    If general.Worksheets("_Rapor Kimliği").Range("B1").Value2 <> "Direktörlük" Then Fail "İlk seçilen dosya direktörlük raporu olmalı."
    Set gs = ExpensePositions(general, gr, gc, gn)
    Set merged = Dict(): Set accepted = Dict(): Set names = Dict()
    Set meta = general.Worksheets("_Rapor Kimliği")
    For r = 2 To meta.Cells(meta.Rows.Count, 1).End(xlUp).Row
        names(CanonicalDept(CStr(meta.Cells(r, 1).Value2))) = CStr(meta.Cells(r, 1).Value2)
    Next r
    For Each file In fd.SelectedItems
        If LCase$(CStr(file)) = LCase$(CStr(generalPath)) Then Fail "Genel raporu müdürlük dosyası olarak seçmeyin."
        Set manager = ReadBook(CStr(file))
        If Identity(manager) <> signature Then Fail "Dönem, kapanış veya ay seçimi farklı: " & manager.Name
        If manager.Worksheets("_Rapor Kimliği").Range("B1").Value2 <> "Müdürlük" Then Fail "Müdürlük raporu bekleniyor: " & manager.Name
        Set ms = ExpensePositions(manager, mr, mc, mn)
        If mn.Count <> mc.Count Or mr.Count <> gr.Count Then Fail "Gider/açıklama yapısı farklı: " & manager.Name
        For Each dept In mc.Keys
            If Not gc.Exists(dept) Or Not names.Exists(dept) Then Fail "Bölüm genel raporda yok: " & dept
            If accepted.Exists(dept) Then Fail "Bölüm birden fazla müdürlük dosyasında seçildi: " & dept
            accepted.Add dept, True
            For Each code In gr.Keys
                If Not mr.Exists(code) Then Fail "Alt gider kodu eksik: " & code
                If ms.Cells(mr(code), 3).Value2 <> gs.Cells(gr(code), 3).Value2 Then Fail "Alt gider adı farklı: " & code
                For i = 0 To 1
                    value = Num(ms.Cells(mr(code), mc(dept) + i).Value2)
                    If Changed(value, Num(gs.Cells(gr(code), gc(dept) + i).Value2)) Then Fail "Rapor tutarları farklı: " & dept & " / " & code
                Next i
                If ms.Cells(mr(code), mn(dept)).HasFormula Then Fail "Açıklama sabit metin olmalı: " & dept & " / " & code
                If IsError(ms.Cells(mr(code), mn(dept)).Value2) Then Fail "Açıklama hücresinde hata var."
                note = Trim$(CStr(ms.Cells(mr(code), mn(dept)).Value2))
                If note <> "" Then merged(CStr(code) & Sep & dept) = note
            Next code
        Next dept
        manager.Close False: Set manager = Nothing
    Next file
    If accepted.Count <> gc.Count Then Fail "Tüm bölümleri kapsayan müdürlük raporlarını seçin. Eksik bölüm sayısı: " & gc.Count - accepted.Count
    target = Application.GetSaveAsFilename(InitialFileName:=fso.BuildPath(fso.GetParentFolderName(CStr(generalPath)), fso.GetBaseName(CStr(generalPath)) & "_Aciklamali_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"), FileFilter:="Excel (*.xlsx),*.xlsx", Title:="Birleştirilmiş yeni raporu kaydedin")
    If VarType(target) = vbBoolean Then GoTo Done
    savePath = CStr(target)
    If LCase$(Right$(savePath, 5)) <> ".xlsx" Then savePath = savePath & ".xlsx"
    If fso.FileExists(savePath) Then Fail "Var olan dosyanın üzerine yazılmayacak; yeni dosya adı seçin."
    ' Validate before creating a copy. Source workbooks are never edited.
    general.SaveCopyAs savePath
    Set output = Application.Workbooks.Open(Filename:=savePath, UpdateLinks:=0)
    For Each ws In output.Worksheets
        If ws.Name = "Alt Gider Açıklamaları" Then ws.Delete: Exit For
    Next ws
    On Error Resume Next
    output.Worksheets("Açıklama Durumu").Delete
    On Error GoTo Bad
    Set ws = output.Worksheets.Add(After:=output.Worksheets(output.Worksheets.Count))
    ws.Name = "Alt Gider Açıklamaları"
    RowValues ws, 1, Array("Ana gider", "Alt gider kodu", "Alt gider", "Bölümlere ait açıklamalar")
    r = 2
    For Each code In gr.Keys
        text = ""
        For Each dept In gc.Keys
            id = CStr(code) & Sep & dept
            If merged.Exists(id) Then
                If text <> "" Then text = text & vbLf
                text = text & names(dept) & ": " & merged(id)
            End If
        Next dept
        If Len(text) > 32767 Then Fail "Açıklama Excel hücre sınırını aşıyor: " & code
        RowValues ws, r, Array(gs.Cells(gr(code), 1).Value2, CLng(code), gs.Cells(gr(code), 3).Value2, text)
        If InStr(text, "[Kontrol bekliyor") > 0 Then ws.Cells(r, 4).Font.Color = RGB(190, 48, 58)
        r = r + 1
    Next code
    Header ws.Range("A1:D1"): ws.Columns("A:C").AutoFit: ws.Columns("D").ColumnWidth = 100
    ws.UsedRange.WrapText = True: ws.UsedRange.Rows.AutoFit
    output.Save: output.Close False: Set output = Nothing: didSave = True
Done:
    On Error Resume Next
    If Not manager Is Nothing Then manager.Close False
    If Not general Is Nothing Then general.Close False
    If Not output Is Nothing Then output.Close False
    If Not didSave And savePath <> "" Then
        If fso.FileExists(savePath) Then fso.DeleteFile savePath, True
    End If
    Application.DisplayAlerts = display: Application.AutomationSecurity = security
    On Error GoTo 0
    If didSave Then
        MsgBox "Açıklamalar bölüm adlarıyla birleştirildi:" & vbCrLf & savePath, vbInformation
    ElseIf errText <> "" Then
        MsgBox "Açıklamalar birleştirilemedi: " & errText, vbExclamation
    End If
    Exit Sub
Bad:
    errText = Err.Description
    Resume Done
End Sub
