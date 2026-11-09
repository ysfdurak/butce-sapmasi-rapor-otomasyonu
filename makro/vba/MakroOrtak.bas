Attribute VB_Name = "MakroOrtak"
Option Explicit

' Excel 2016+ / Windows. No Python, PowerShell or external references.
Public Cfg As Object, Depts As Object, Groups As Collection, ExpenseMap As Object
Public Excluded As Object, Snaps As Object, Indices As Object, Revenue As Object
Public RevenueBudget As Object, Controls As Object, Warnings As Collection
Public Revisions As Collection, Checks As Collection, Units As Object, Results As Object
Public CloseYear As Long, CloseMonth As Long, ReportYear As Long, Picked As Object
Public SettingsPath As String, RootPath As String, OutputPath As String, RunPath As String
Public Const Sep As String = "|"
Public Const MoneyFmt As String = "#,##0.00;-#,##0.00;0.00"
Public Const PctFmt As String = "0.00%;-0.00%;0.00%"

Public Function Dict() As Object
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = vbTextCompare
End Function

Public Function K(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    Dim s As String
    s = Trim$(CStr(value))
    s = Replace(s, "İ", "i"): s = Replace(s, "I", "ı")
    s = LCase$(s)
    s = Replace(s, "ı", "i"): s = Replace(s, "ş", "s")
    s = Replace(s, "ğ", "g"): s = Replace(s, "ü", "u")
    s = Replace(s, "ö", "o"): s = Replace(s, "ç", "c")
    K = s
End Function

Public Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 2101, "Bütçe Raporu", message
End Sub

Public Function Num(ByVal v As Variant) As Variant
    Num = Null
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Or VarType(v) = vbString Then Exit Function
    If IsNumeric(v) Then Num = CDec(v)
End Function

' Metin olarak girilmiş tutar: "1.234,56", "1234,56", "(1.234)", "1234.56". Değilse Null.
' Nokta + üçlü basamak grubu Türkçe binlik ayırıcıdır ("1.234" = 1234). Python: butce.text_number.
Public Function TextNumber(ByVal v As Variant) As Variant
    Dim s As String, negative As Boolean, re As Object, intPart As String, fracPart As String, parts As Variant
    TextNumber = Null
    If VarType(v) <> vbString Then Exit Function
    s = Trim$(CStr(v))
    s = Replace(s, ChrW(160), ""): s = Replace(s, " ", ""): s = Replace(s, ChrW(8378), "")
    If UCase$(Right$(s, 2)) = "TL" Then s = Left$(s, Len(s) - 2)
    If Len(s) > 2 And Left$(s, 1) = "(" And Right$(s, 1) = ")" Then negative = True: s = Mid$(s, 2, Len(s) - 2)
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[-+]?([0-9]{1,3}(\.[0-9]{3})+|[0-9]+)(,[0-9]+)?$"
    If re.Test(s) Then
        s = Replace(s, ".", "")
    Else
        re.Pattern = "^[-+]?[0-9]+\.[0-9]+$"
        If Not re.Test(s) Then Exit Function
        s = Replace(s, ".", ",")
    End If
    ' Burada yalnız rakam, işaret ve tek virgül (ondalık) kaldı; yerel ayardan bağımsız çevrilir.
    parts = Split(s & ",", ",")
    intPart = parts(0): fracPart = parts(1)
    TextNumber = CDec(0)
    If intPart <> "" And intPart <> "-" And intPart <> "+" Then TextNumber = CDec(Abs(CDec(Val(Replace(intPart, "+", "")))))
    If fracPart <> "" Then TextNumber = TextNumber + CDec(Val(fracPart)) / CDec(10 ^ Len(fracPart))
    If Left$(intPart, 1) = "-" Then TextNumber = -TextNumber
    If negative Then TextNumber = -TextNumber
End Function

' Gider kalemi sayısı: ayar yoksa dip toplam satırına kadar, o da yoksa 27. Python: butce.item_count.
Public Function ItemCount() As Long
    Dim first As Long
    first = CLng(Opt("Maliyet", "ilk_gider_satiri"))
    If CStr(Opt("Maliyet", "gider_kalem_sayisi")) <> "" Then
        ItemCount = CLng(Opt("Maliyet", "gider_kalem_sayisi"))
    ElseIf CStr(Opt("Maliyet", "dip_toplam_satiri")) <> "" Then
        ItemCount = CLng(Opt("Maliyet", "dip_toplam_satiri")) - first
        If ItemCount < 1 Then ItemCount = 27
    Else
        ItemCount = 27
    End If
End Function

' Sayfayı büyük/küçük harf ve Türkçe harf farkı gözetmeden bulur (yükleme ile aynı eşleştirme).
Public Function SheetByKey(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If K(ws.Name) = K(name) Then Set SheetByKey = ws: Exit Function
    Next ws
    Fail "Kaynak sayfası bulunamadı: " & name & " / " & wb.Name
End Function

' Kapanış yılının <= kapanış dosyaları; Ocak kapanışında yıl dönümü revizyonu için önceki Aralık da.
Public Function InClosingScope(ByVal period As String) As Boolean
    If period = "" Then Exit Function
    If period > P(CloseYear, CloseMonth) Then Exit Function
    InClosingScope = (Left$(period, 4) = CStr(CloseYear)) Or (CloseMonth = 1 And period = P(CloseYear - 1, 12))
End Function

Public Function CurrentUser() As String
    CurrentUser = Environ$("USERNAME")
    If CurrentUser = "" Then CurrentUser = Application.UserName
    If CurrentUser = "" Then CurrentUser = "Bilinmiyor"
End Function

Public Function ValAt(ByVal d As Object, ByVal id As String) As Variant
    ValAt = Null
    If d.Exists(id) Then ValAt = d(id)
End Function

Public Function Plus(ByVal a As Variant, ByVal b As Variant) As Variant
    Plus = Null
    If Not IsNull(a) And Not IsNull(b) Then Plus = CDec(a) + CDec(b)
End Function

Public Function Minus(ByVal a As Variant, ByVal b As Variant) As Variant
    Minus = Null
    If Not IsNull(a) And Not IsNull(b) Then Minus = CDec(a) - CDec(b)
End Function

Public Function Times(ByVal a As Variant, ByVal b As Variant) As Variant
    Times = Null
    If Not IsNull(a) And Not IsNull(b) Then Times = CDec(a) * CDec(b)
End Function

Public Function Divide(ByVal a As Variant, ByVal b As Variant) As Variant
    Divide = Null
    If IsNull(a) Or IsNull(b) Then Exit Function
    If b <> 0 Then Divide = CDec(a) / CDec(b)
End Function

Public Function Positive(ByVal a As Variant) As Boolean
    If Not IsNull(a) Then Positive = (a > 0)
End Function

Public Function Nonnegative(ByVal a As Variant) As Boolean
    If Not IsNull(a) Then Nonnegative = (a >= 0)
End Function

Public Function Changed(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNull(a) Or IsNull(b) Then
        Changed = (IsNull(a) Xor IsNull(b))
    Else
        Changed = (a <> b)
    End If
End Function

Public Function P(ByVal y As Long, ByVal m As Long) As String
    P = Format$(y, "0000") & "-" & Format$(m, "00")
End Function

Public Function Opt(ByVal section As String, ByVal code As String, Optional ByVal fallback As Variant = "") As Variant
    Opt = fallback
    If Cfg.Exists(section & Sep & code) Then Opt = Cfg(section & Sep & code)
End Function

Public Function AbsPath(ByVal relative As String, Optional ByVal base As String = "") As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If base = "" Then base = RootPath
    If Left$(relative, 2) <> "\\" And Mid$(relative, 2, 1) <> ":" Then relative = fso.BuildPath(base, relative)
    AbsPath = fso.GetAbsolutePathName(relative)
End Function

Public Function ReadBook(ByVal path As String) As Workbook
    If Len(Dir$(path)) = 0 Then Fail "Dosya bulunamadı: " & path
    Set ReadBook = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
End Function

Public Sub Warn(ByVal dept As String, ByVal file As String, ByVal message As String)
    Warnings.Add Array(P(CloseYear, CloseMonth), dept, file, message)
End Sub

Public Function PeriodOf(ByVal value As Variant, Optional ByVal budgetYear As Long = 0) As String
    Dim dt As Date, s As String, a As Variant, y As Long, m As Long
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    On Error GoTo Invalid
    If VarType(value) = vbDate Then
        dt = CDate(value): PeriodOf = P(Year(dt), Month(dt)): Exit Function
    ElseIf IsNumeric(value) And VarType(value) <> vbString Then
        If CDbl(value) > 20000 And CDbl(value) < 100000 Then
            dt = DateSerial(1899, 12, 30) + CDbl(value)
            PeriodOf = P(Year(dt), Month(dt)): Exit Function
        End If
    End If
    s = Trim$(CStr(value))
    If Len(s) >= 7 And Mid$(s, 5, 1) = "-" Then
        y = CLng(Left$(s, 4)): m = CLng(Mid$(s, 6, 2))
    ElseIf InStr(s, ".") > 0 Then
        a = Split(s, "."): y = CLng(a(2)): m = CLng(a(1))
    ElseIf IsDate(s) Then
        dt = CDate(s): y = Year(dt): m = Month(dt)
    ElseIf budgetYear > 0 Then
        y = budgetYear: m = MonthNum(s)
    End If
    If y >= 1900 And y <= 9999 And m >= 1 And m <= 12 Then PeriodOf = P(y, m)
Invalid:
End Function

Public Function MonthNum(ByVal value As Variant) As Long
    Dim names As Variant, i As Long, s As String
    s = K(value): names = Split("ocak subat mart nisan mayis haziran temmuz agustos eylul ekim kasim aralik", " ")
    If IsNumeric(s) Then
        If Val(s) >= 1 And Val(s) <= 12 And Val(s) = Fix(Val(s)) Then MonthNum = CLng(s)
    Else
        For i = 0 To 11
            If s = names(i) Or s = Left$(names(i), 3) Then MonthNum = i + 1: Exit Function
        Next i
    End If
End Function

Public Function FilenamePeriod(ByVal file As String) As String
    Dim re As Object, matches As Object, found As Object, i As Long, a As String, b As String, y As Long, m As Long
    Set re = CreateObject("VBScript.RegExp"): re.Global = True: re.Pattern = "[a-z0-9]+"
    Set matches = re.Execute(K(CreateObject("Scripting.FileSystemObject").GetBaseName(file)))
    Set found = Dict()
    For i = 0 To matches.Count - 2
        a = matches(i).value: b = matches(i + 1).value: y = 0: m = 0
        If Len(a) = 4 And IsNumeric(a) Then
            y = CLng(a): m = MonthNum(b)
        ElseIf Len(b) = 4 And IsNumeric(b) Then
            y = CLng(b): m = MonthNum(a)
        End If
        If y >= 1900 And y <= 9999 And m > 0 Then found(P(y, m)) = True
    Next i
    If found.Count = 1 Then FilenamePeriod = found.Keys()(0)
End Function

Public Function ReadCols(ByVal ws As Worksheet, ByVal address As String, Optional ByVal lastCol As Long = 0, Optional ByVal onlyYear As Long = 0) As Object
    Dim d As Object, first As Range, c As Long, v As Variant, id As String
    Set d = Dict(): Set first = ws.Range(address).Cells(1, 1)
    If lastCol = 0 Then lastCol = ws.Cells(first.Row, ws.Columns.Count).End(xlToLeft).Column
    For c = first.Column To lastCol
        v = ws.Cells(first.Row, c).Value2
        If IsEmpty(v) Or CStr(v) = "" Then
            If onlyYear = 0 Then Exit For
        Else
            id = PeriodOf(v, onlyYear)
            If id = "" And onlyYear = 0 Then Fail "Geçersiz tarih: " & ws.Name & "!" & ws.Cells(first.Row, c).Address
            If id <> "" Then
                If onlyYear = 0 Or Left$(id, 4) = CStr(onlyYear) Then
                    If d.Exists(id) Then Fail "Mükerrer dönem: " & id & " / " & ws.Name
                    d.Add id, c
                End If
            End If
        End If
    Next c
    Set ReadCols = d
End Function

Public Function FullBudget(ByVal cols As Object, Optional ByVal y As Long = 0) As Boolean
    ' y: bütçe yılı; verilmezse kapanış yılı. Yıl dönümünde önceki Aralık dosyası kendi yılıyla okunur.
    Dim m As Long
    If y = 0 Then y = CloseYear
    If cols.Count <> 12 Then Exit Function
    For m = 1 To 12
        If Not cols.Exists(P(y, m)) Then Exit Function
    Next m
    FullBudget = True
End Function

Public Function SumMonths(ByVal d As Object, ByVal prefix As String, ByVal y As Long, Optional ByVal suffix As String = "", Optional ByVal positiveOnly As Boolean = False) As Variant
    Dim m As Variant, v As Variant, result As Variant
    result = CDec(0)
    For Each m In Picked.Keys
        v = ValAt(d, prefix & P(y, CLng(m)) & suffix)
        If positiveOnly Then
            If Not Positive(v) Then v = Null
        End If
        result = Plus(result, v)
    Next m
    SumMonths = result
End Function

Public Function GroupOf(ByVal code As Long) As String
    Dim g As Variant
    GroupOf = "Eşleşmeyen"
    For Each g In Groups
        If code >= g(0) And code <= g(1) Then GroupOf = g(2): Exit Function
    Next g
End Function

Public Function CurrentSnap(ByVal dept As String, Optional ByVal closingMonth As Long = 0) As Object
    Dim id As String
    If closingMonth = 0 Then closingMonth = CloseMonth
    id = K(dept) & Sep & P(CloseYear, closingMonth)
    If Snaps.Exists(id) Then Set CurrentSnap = Snaps(id)
End Function

Public Function DeptScope(Optional ByVal manager As String = "", Optional ByVal director As String = "") As Object
    Dim d As Object, n As Variant, meta As Variant
    Set d = Dict()
    For Each n In Depts.Keys
        meta = Depts(n)
        If (manager = "" Or K(meta(1)) = K(manager)) And (director = "" Or K(meta(2)) = K(director)) Then d(n) = True
    Next n
    Set DeptScope = d
End Function

Public Function Managers(ByVal scope As Object) As Object
    ' Müdürlük adları K() ile karşılaştırılır; yazım farkı (büyük/küçük, Türkçe harf) ayrı müdürlük üretmez.
    Dim d As Object, seen As Object, n As Variant, meta As Variant
    Set d = Dict(): Set seen = Dict()
    For Each n In scope.Keys
        meta = Depts(n)
        If Not seen.Exists(K(meta(1))) Then seen(K(meta(1))) = True: d(meta(1)) = True
    Next n
    Set Managers = d
End Function

Public Function CleanName(ByVal s As String) As String
    Dim c As Variant
    For Each c In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|", "[", "]")
        s = Replace(s, c, "_")
    Next c
    CleanName = Left$(Trim$(s), 70)
End Function

Public Function SelectionCaption(Optional ByVal y As Long = 0) As String
    Dim m As Variant, s As String
    If y = 0 Then y = ReportYear
    For Each m In Picked.Keys
        If s <> "" Then s = s & ","
        s = s & Format$(CLng(m), "00")
    Next m
    SelectionCaption = CStr(y) & " / Aylar " & s
End Function

Public Function Signature() As String
    Dim m As Variant, s As String
    For Each m In Picked.Keys
        If s <> "" Then s = s & ","
        s = s & Format$(CLng(m), "00")
    Next m
    Signature = "aylar:" & s
End Function

Public Function IndicatorName(ByVal snapshot As Object, ByVal index As Long) As String
    Dim names As Object
    IndicatorName = "Kaynak başlığı eksik (" & index & ". gösterge)"
    If snapshot Is Nothing Then Exit Function
    If snapshot.Exists("IndicatorNames") Then
        Set names = snapshot("IndicatorNames")
        If names.Exists(CStr(index)) Then IndicatorName = CStr(names(CStr(index)))
    End If
End Function

Public Function SourceIndicatorTitle(ByVal sheet As Worksheet, ByVal index As Long) As String
    Dim row As Long, col As Long, value As Variant
    row = CLng(Opt("Maliyet", "islem_adedi_satiri")) - 5 + index
    col = CLng(Opt("Maliyet", "gosterge_adi_sutunu", Opt("Maliyet", "gider_sutunu")))
    value = sheet.Cells(row, col).Value2
    If Not IsError(value) Then
        If VarType(value) = vbString Then SourceIndicatorTitle = Trim$(CStr(value))
    End If
    If SourceIndicatorTitle = "" Then SourceIndicatorTitle = "Başlığı eksik gösterge (" & sheet.Cells(row, col).Address(False, False) & ")"
End Function

Public Function IndicatorCaption(ByVal scope As Object, ByVal index As Long) As String
    Dim names As Object, n As Variant, d As Object, title As Variant, value As String
    Set names = Dict()
    For Each n In scope.Keys
        Set d = CurrentSnap(CStr(n))
        If Not d Is Nothing Then names(IndicatorName(d, index)) = True
    Next n
    For Each title In names.Keys
        If value <> "" Then value = value & " / "
        value = value & CStr(title)
    Next title
    If value = "" Then value = "Kaynak başlığı bulunamadı"
    IndicatorCaption = value
End Function
