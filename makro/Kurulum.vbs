Option Explicit
Dim fso, root, target, excel, book, project, moduleFile, component, stream, code, lines, i, modules, oldSecurity
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetParentFolderName(WScript.ScriptFullName)
target = fso.BuildPath(root, "Butce_Kontrol_Makro_V4.xlsm")
If fso.FileExists(target) Then
    CreateObject("Shell.Application").Open target
    WScript.Quit 0
End If
On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then StopInstall "Windows desktop Excel is required: " & Err.Description
excel.DisplayAlerts = False
Set book = excel.Workbooks.Open(fso.BuildPath(root, "Kontrol_Sablonu.xlsx"), 0, False)
If Err.Number <> 0 Then StopInstall "Template could not be opened: " & Err.Description
Set project = book.VBProject
If Err.Number <> 0 Then StopInstall "Excel denied access to the VBA project. Use the manual installation in KULLANIM.md. No security settings were changed. " & Err.Description
modules = Array("MakroOrtak", "MakroVeri", "MakroHesap", "MakroYaz", "MakroPanel", "MakroAciklama", "MakroTest", "MakroAyarlar", "MakroMerkez", "MakroYorum")
For Each moduleFile In modules
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 : stream.Charset = "utf-8" : stream.Open
    stream.LoadFromFile fso.BuildPath(root, "vba\" & moduleFile & ".bas")
    code = stream.ReadText : stream.Close
    If Err.Number <> 0 Then StopInstall "Module could not be read: " & moduleFile & " / " & Err.Description
    lines = Split(Replace(code, vbCrLf, vbLf), vbLf)
    code = ""
    For i = 0 To UBound(lines)
        If Left(Trim(lines(i)), 10) <> "Attribute " Then code = code & lines(i) & vbCrLf
    Next
    Set component = project.VBComponents.Add(1)
    component.Name = moduleFile
    component.CodeModule.AddFromString code
    If Err.Number <> 0 Then StopInstall "VBA module import failed: " & moduleFile & " / " & Err.Description
Next
book.SaveAs target, 52
If Err.Number <> 0 Then StopInstall "XLSM could not be saved: " & Err.Description
excel.DisplayAlerts = True : excel.Visible = True
excel.Run "'" & Replace(book.Name, "'", "''") & "'!PanelHazirla"
If Err.Number <> 0 Then
    MsgBox "XLSM was created. Enable macros according to your company rules, then Alt+F8 > PanelHazirla. " & Err.Description, 48, "Budget macro setup"
Else
    book.Save
    MsgBox "Ready: Butce_Kontrol_Makro_V4.xlsm. Select sources, months and click Rapor olustur. See KULLANIM.md for the synthetic test.", 64, "Budget macro setup"
End If
WScript.Quit 0

Sub StopInstall(message)
    Dim errorText
    errorText = message
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If Not excel Is Nothing Then excel.Quit
    MsgBox errorText, 48, "Budget macro setup"
    WScript.Quit 1
End Sub
