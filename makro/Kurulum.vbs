' Butce Sapmasi Excel makro kurulumu / guncellemesi (V5).
' Mesajlar Turkce karakter kullanmaz: VBScript dosyasi sistem kod sayfasiyla okunur.
Option Explicit
Dim fso, shell, root, target, legacy, excel, book, project, moduleFile, component, stream, code, lines, i, modules
Dim mode, answer, backupFolder, stamp, acp, removed
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
root = fso.GetParentFolderName(WScript.ScriptFullName)
target = fso.BuildPath(root, "Butce_Kontrol_Makro_V5.xlsm")
legacy = fso.BuildPath(root, "Butce_Kontrol_Makro_V4.xlsm")
modules = Array("MakroOrtak", "MakroVeri", "MakroHesap", "MakroYaz", "MakroPanel", "MakroAciklama", "MakroTest", "MakroAyarlar", "MakroMerkez", "MakroYorum", "MakroAnaliz")
stamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)

' 1) Kod sayfasi: VBA duzenleyicisi Unicode metni sistemin ANSI kod sayfasina cevirir.
On Error Resume Next
acp = shell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Nls\CodePage\ACP")
On Error GoTo 0
If CStr(acp) <> "" And CStr(acp) <> "1254" Then
    answer = MsgBox("Windows kod sayfasi " & acp & " (Turkce degil)." & vbCrLf & vbCrLf & _
        "VBA modullerindeki Turkce karakterler (sayfa adlari, mesajlar) bozulabilir ve raporlar hatali olabilir." & vbCrLf & _
        "Oneri: Denetim Masasi > Bolge > Yonetimsel > Unicode olmayan programlarin dili = Turkce (Turkiye)." & vbCrLf & vbCrLf & _
        "Yine de devam edilsin mi?", vbYesNo + vbExclamation, "Butce makro kurulumu")
    If answer <> vbYes Then WScript.Quit 1
End If

' 2) Kurulum turu: yeni, mevcut V5 guncellemesi veya V4'ten yukseltme.
mode = "new"
If fso.FileExists(target) Then
    answer = MsgBox("Butce_Kontrol_Makro_V5.xlsm zaten var." & vbCrLf & vbCrLf & _
        "Evet: makro modullerini bu paketteki surumle guncelle (once yedek alinir; paneldeki secimler, yorumlar ve gecmis korunur)." & vbCrLf & _
        "Hayir: dosyayi degistirmeden ac." & vbCrLf & "Iptal: cik.", vbYesNoCancel + vbQuestion, "Butce makro kurulumu")
    If answer = vbCancel Then WScript.Quit 0
    If answer = vbNo Then
        shell.Run """" & target & """"
        WScript.Quit 0
    End If
    mode = "upgrade"
ElseIf fso.FileExists(legacy) Then
    answer = MsgBox("Butce_Kontrol_Makro_V4.xlsm bulundu." & vbCrLf & vbCrLf & _
        "Evet: V4 dosyasinin kopyasi V5 olarak guncellensin (V4 dosyasi oldugu gibi kalir; yorumlar ve gecmis tasinir)." & vbCrLf & _
        "Hayir: bos sablondan yeni V5 olustur.", vbYesNoCancel + vbQuestion, "Butce makro kurulumu")
    If answer = vbCancel Then WScript.Quit 0
    If answer = vbYes Then
        fso.CopyFile legacy, target, False
        mode = "upgrade"
    End If
End If

If mode = "upgrade" Then
    backupFolder = fso.BuildPath(root, "Kurulum_Yedekleri")
    If Not fso.FolderExists(backupFolder) Then fso.CreateFolder backupFolder
    fso.CopyFile target, fso.BuildPath(backupFolder, "Butce_Kontrol_Makro_V5_" & stamp & ".xlsm"), False
End If

On Error Resume Next
Set excel = CreateObject("Excel.Application")
If Err.Number <> 0 Then StopInstall "Masaustu Windows Excel gerekli: " & Err.Description
excel.DisplayAlerts = False
excel.EnableEvents = False  ' mevcut xlsm'deki olay kodu kurulum sirasinda calismasin
If mode = "upgrade" Then
    Set book = excel.Workbooks.Open(target, 0, False)
Else
    Set book = excel.Workbooks.Open(fso.BuildPath(root, "Kontrol_Sablonu.xlsx"), 0, False)
End If
If Err.Number <> 0 Then StopInstall "Calisma kitabi acilamadi: " & Err.Description
Set project = book.VBProject
If Err.Number <> 0 Then StopInstall "Excel VBA projesine erisime izin vermedi. KULLANIM.md'deki elle kurulumu kullanin " & _
    "(vba_elle klasoru). Hicbir guvenlik ayari degistirilmedi. " & Err.Description

' 3) Guncellemede eski Makro* modulleri kaldirilir; veri sayfalari dokunulmadan kalir.
If mode = "upgrade" Then
    removed = 0
    For i = project.VBComponents.Count To 1 Step -1
        Set component = project.VBComponents(i)
        If component.Type = 1 And Left(component.Name, 5) = "Makro" Then
            project.VBComponents.Remove component
            removed = removed + 1
        End If
    Next
    If Err.Number <> 0 Then StopInstall "Eski moduller kaldirilamadi: " & Err.Description
End If

For Each moduleFile In modules
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 : stream.Charset = "utf-8" : stream.Open
    stream.LoadFromFile fso.BuildPath(root, "vba\" & moduleFile & ".bas")
    code = stream.ReadText : stream.Close
    If Err.Number <> 0 Then StopInstall "Modul okunamadi: " & moduleFile & " / " & Err.Description
    lines = Split(Replace(code, vbCrLf, vbLf), vbLf)
    code = ""
    For i = 0 To UBound(lines)
        If Left(Trim(lines(i)), 10) <> "Attribute " Then code = code & lines(i) & vbCrLf
    Next
    Set component = project.VBComponents.Add(1)
    component.Name = moduleFile
    component.CodeModule.AddFromString code
    If Err.Number <> 0 Then StopInstall "VBA modulu eklenemedi: " & moduleFile & " / " & Err.Description
Next

If mode = "upgrade" Then
    book.Save
Else
    book.SaveAs target, 52
End If
If Err.Number <> 0 Then StopInstall "XLSM kaydedilemedi: " & Err.Description
excel.DisplayAlerts = True : excel.EnableEvents = True : excel.Visible = True
excel.Run "'" & Replace(book.Name, "'", "''") & "'!PanelHazirla"
If Err.Number <> 0 Then
    MsgBox "XLSM hazir. Sirket kurallarina gore makrolari etkinlestirin, ardindan Alt+F8 > PanelHazirla calistirin. " & Err.Description, 48, "Butce makro kurulumu"
Else
    book.Save
    If mode = "upgrade" Then
        MsgBox "Guncellendi: Butce_Kontrol_Makro_V5.xlsm (" & removed & " eski modul degistirildi). Yedek: Kurulum_Yedekleri klasoru." & vbCrLf & _
            "Once Sentetik hesap testi dugmesini calistirin.", 64, "Butce makro kurulumu"
    Else
        MsgBox "Hazir: Butce_Kontrol_Makro_V5.xlsm. Once Sentetik hesap testi dugmesini, sonra Kaynaklari onizle ve Rapor olustur adimlarini kullanin.", 64, "Butce makro kurulumu"
    End If
End If
WScript.Quit 0

Sub StopInstall(message)
    Dim errorText
    errorText = message
    On Error Resume Next
    If IsObject(book) Then book.Close False
    If IsObject(excel) Then excel.Quit
    MsgBox errorText, 48, "Butce makro kurulumu"
    WScript.Quit 1
End Sub
