Attribute VB_Name = "MakroHesap"
Option Explicit

Public Function Factor(ByVal alias As String) As Variant
    Dim m As Variant, a As Variant, b As Variant, result As Variant
    Factor = Null
    If alias = "" Then Exit Function
    result = CDec(0)
    For Each m In Picked.Keys
        a = ValAt(Indices, K(alias) & Sep & P(ReportYear, CLng(m)))
        b = ValAt(Indices, K(alias) & Sep & P(ReportYear - 1, CLng(m)))
        If Not Positive(a) Or Not Positive(b) Then Exit Function
        result = Plus(result, Divide(a, b))
    Next m
    Factor = Divide(result, Picked.Count)
End Function

Public Function ExpenseTotal(ByVal scope As Object, ByVal slot As String, ByVal y As Long, Optional ByVal code As String = "") As Variant
    Dim n As Variant, d As Object, result As Variant, v As Variant
    result = CDec(0)
    If scope.Count = 0 Then ExpenseTotal = Null: Exit Function
    For Each n In scope.Keys
        Set d = CurrentSnap(CStr(n))
        If d Is Nothing Then
            v = Null
        ElseIf code = "" Then
            v = SumMonths(d, slot & "T" & Sep, y)
        Else
            v = SumMonths(d, slot & Sep, y, Sep & code)
        End If
        result = Plus(result, v)
    Next n
    ExpenseTotal = result
End Function

Public Function ItemNames() As Object
    Dim names As Object, d As Object, n As Variant, code As Variant, local As Object, a As Variant
    Dim sorted As Object, keys As Variant, i As Long, j As Long, tmp As Variant
    Set names = Dict(): Set sorted = Dict()
    For Each n In Depts.Keys
        Set d = CurrentSnap(CStr(n))
        If Not d Is Nothing Then
            Set local = d("names")
            For Each code In local.Keys
                If names.Exists(code) Then
                    If names(code) <> local(code) Then Fail "Bölümler arasında gider adı farklı; gider eşleştirmesini tamamlayın: " & code
                Else
                    names.Add code, local(code)
                End If
            Next code
        End If
    Next n
    keys = names.Keys
    For i = 0 To names.Count - 2
        For j = i + 1 To names.Count - 1
            If CLng(keys(j)) < CLng(keys(i)) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next j
    Next i
    For i = 0 To names.Count - 1
        sorted.Add CStr(keys(i)), names(keys(i))
    Next i
    Set ItemNames = sorted
End Function

Public Function GroupTotal(ByVal scope As Object, ByVal groupName As String, ByVal slot As String, ByVal y As Long, ByVal names As Object) As Variant
    Dim code As Variant, result As Variant, count As Long
    result = CDec(0)
    For Each code In names.Keys
        If GroupOf(CLng(code)) = groupName Then
            count = count + 1: result = Plus(result, ExpenseTotal(scope, slot, y, CStr(code)))
        End If
    Next code
    If count = 0 Then result = Null
    GroupTotal = result
End Function

Public Sub CalculateUnits()
    Dim n As Variant, meta As Variant, d As Object, a As Variant, b As Variant, x As Variant, y As Variant
    Dim f As Variant, budgetBM As Variant, adjusted As Variant, rate As Variant, row As Variant
    Dim groupCosts As Object, directors As Object, eligible As Boolean, managerKey As String
    Set Units = Dict(): Set groupCosts = Dict(): Set directors = Dict()
    For Each n In Depts.Keys
        meta = Depts(n): Set d = CurrentSnap(CStr(n))
        If meta(3) <> "" And Not d Is Nothing Then
            a = SumMonths(d, "AT|", ReportYear): b = SumMonths(d, "AT|", ReportYear - 1)
            x = SumMonths(d, "AC|", ReportYear, "", True): y = SumMonths(d, "AC|", ReportYear - 1, "", True)
            f = Factor(CStr(meta(3)))
            If ReportYear = CloseYear Then
                eligible = Nonnegative(a) And Positive(b) And Positive(x) And Positive(y) And Positive(f)
            Else
                eligible = Nonnegative(a) And Positive(x)
            End If
            If eligible Then
                adjusted = Times(Divide(b, y), f)
                rate = Divide(Minus(adjusted, Divide(a, x)), adjusted)
                budgetBM = Null
                If ReportYear = CloseYear Then budgetBM = Divide(SumMonths(d, "BT|", ReportYear), SumMonths(d, "BC|", ReportYear, "", True))
                ' A, countA, B, countB, currentBM, priorBM, budgetBM, factor, adjustedBM, realRate, managerWeight, directorWeight
                row = Array(a, x, b, y, Divide(a, x), Divide(b, y), budgetBM, f, adjusted, rate, Null, Null)
                Units.Add n, row
                managerKey = K(meta(2)) & Sep & K(meta(1))
                If Not groupCosts.Exists(managerKey) Then groupCosts(managerKey) = CDec(0)
                If Not directors.Exists(K(meta(2))) Then directors(K(meta(2))) = CDec(0)
                groupCosts(managerKey) = Plus(groupCosts(managerKey), a)
                directors(K(meta(2))) = Plus(directors(K(meta(2))), a)
            Else
                Warn CStr(meta(0)), CStr(d("file")), "Birim maliyet kapsamı dışında: gider/pozitif işlem adedi/önceki yıl veya endeks eksik."
            End If
        ElseIf meta(3) = "" Then
            Warn CStr(meta(0)), "", "Endeks eşleşmesi yok; birim maliyet kapsamına alınmadı."
        End If
    Next n
    For Each n In Units.Keys
        meta = Depts(n): row = Units(n): managerKey = K(meta(2)) & Sep & K(meta(1))
        row(10) = Divide(row(0), groupCosts(managerKey))
        row(11) = Divide(row(0), directors(K(meta(2))))
        Units(n) = row
    Next n
End Sub

Public Sub CalculateChecks()
    Dim clusters As Object, members As Object, n As Variant, meta As Variant, cluster As Variant
    Dim id As String, bucket As Collection, target As Object, m As Variant, d As Object
    Dim a As Variant, b As Variant, diff As Variant, state As String, labels As String, manager As String
    Dim row As Variant, cm As Long, old As Object, fresh As Object, codes As Object, code As Variant
    Dim local As Object, slot As Variant, affected As Long, last As Long, before As Variant, after As Variant
    Set Checks = New Collection: Set Revisions = New Collection: Set clusters = Dict()
    For Each n In Depts.Keys
        meta = Depts(n): id = K(meta(2)) & Sep & K(meta(5))
        Set d = CurrentSnap(CStr(n))
        If Not d Is Nothing Then
            For Each m In Picked.Keys
                If IsNull(ValAt(d, "AT|" & P(ReportYear, CLng(m)))) Then Warn CStr(meta(0)), CStr(d("file")), P(ReportYear, CLng(m)) & ": Gerçekleşen giderlerden en az biri eksik."
                If IsNull(ValAt(d, "AT|" & P(ReportYear - 1, CLng(m)))) Then Warn CStr(meta(0)), CStr(d("file")), P(ReportYear - 1, CLng(m)) & ": Önceki yıl giderlerinden en az biri eksik."
                If ReportYear = CloseYear Then
                    If IsNull(ValAt(d, "BT|" & P(ReportYear, CLng(m)))) Then Warn CStr(meta(0)), CStr(d("file")), P(ReportYear, CLng(m)) & ": Bütçe giderlerinden en az biri eksik."
                End If
            Next m
        End If
        If Not clusters.Exists(id) Then
            Set members = Dict(): clusters.Add id, members
        End If
        Set members = clusters(id): members.Add n, True
    Next n
    For Each m In Picked.Keys
        If IsNull(ValAt(Revenue, P(ReportYear, CLng(m)))) Then Warn "", "", P(ReportYear, CLng(m)) & ": Şirket cirosu eksik."
        If ReportYear = CloseYear Then
            If IsNull(ValAt(RevenueBudget, P(ReportYear, CLng(m)))) Then Warn "", "", P(ReportYear, CLng(m)) & ": Bütçe cirosu eksik."
        End If
    Next m
    For Each cluster In clusters.Keys
        Set members = clusters(cluster): n = members.Keys()(0): meta = Depts(n)
        labels = "": manager = ""
        For Each n In members.Keys
            row = Depts(n)
            If labels <> "" Then labels = labels & ", "
            labels = labels & row(0)
            If manager <> "" Then manager = manager & ", "
            manager = manager & row(1)
        Next n
        For Each m In Picked.Keys
            a = CDec(0): b = Null: state = ""
            For Each n In members.Keys
                Set d = CurrentSnap(CStr(n))
                If d Is Nothing Then
                    a = Null
                Else
                    a = Plus(a, ValAt(d, "AT|" & P(ReportYear, CLng(m))))
                End If
            Next n
            If Not Controls.Exists(cluster) Then
                state = "Bölüm eşleşmedi"
            Else
                Set bucket = Controls(cluster)
                If bucket.Count <> 1 Then
                    state = "Birden fazla kontrol satırı eşleşti"
                Else
                    Set target = bucket(1): b = ValAt(target, P(ReportYear, CLng(m)))
                End If
            End If
            diff = Minus(a, b)
            If state = "" Then
                If IsNull(a) Then
                    state = "Kaynak toplamı eksik"
                ElseIf IsNull(b) Then
                    state = "Kontrol tutarı/ayı eksik"
                ElseIf Abs(diff) > CDec(Opt("Kontrol", "tolerans", 0.01)) Then
                    state = "Veri doğrulanamadı: tutar farklı"
                Else
                    state = "Doğrulandı"
                End If
            End If
            Checks.Add Array(CStr(meta(5)), manager, CStr(meta(2)), P(ReportYear, CLng(m)), a, b, diff, state, labels, members)
            If state <> "Doğrulandı" Then Warn labels, "", P(ReportYear, CLng(m)) & ": " & state & " / Kontrol: " & meta(5)
        Next m
    Next cluster
    For Each n In Depts.Keys
        meta = Depts(n)
        For cm = 2 To CloseMonth
            Set old = CurrentSnap(CStr(n), cm - 1): Set fresh = CurrentSnap(CStr(n), cm)
            If old Is Nothing Or fresh Is Nothing Then
                Warn CStr(meta(0)), "", P(CloseYear, cm) & ": Ardışık kaynak eksik; revizyon karşılaştırılamadı."
            Else
                Set codes = Dict(): Set local = old("names")
                For Each code In local.Keys
                    codes(code) = local(code)
                Next code
                Set local = fresh("names")
                For Each code In local.Keys
                    codes(code) = local(code)
                Next code
                For Each code In codes.Keys
                    For Each slot In Array("A", "B")
                        If slot = "A" Then last = cm - 1 Else last = 12
                        For affected = 1 To last
                            id = slot & Sep & P(CloseYear, affected) & Sep & code
                            before = ValAt(old, id): after = ValAt(fresh, id)
                            If Changed(before, after) Then
                                If slot = "A" Then state = "Gerçekleşen" Else state = "Bütçe"
                                Revisions.Add Array(P(CloseYear, cm), P(CloseYear, affected), meta(1), meta(0), CLng(code), codes(code), state, before, after, Minus(after, before), old("file"), fresh("file"), CStr(n))
                            End If
                        Next affected
                    Next slot
                Next code
            End If
        Next cm
    Next n
End Sub
