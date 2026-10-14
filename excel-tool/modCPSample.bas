Attribute VB_Name = "modCPSample"
Option Explicit

' ==========================================================================
' COUNTERPARTY SAMPLING
'
' Picks the case's sample counterparties from the alerted transactions and
' writes them to Sheet1: names in J18:J23, addresses in T18:T23.
'
' Data: ThisWorkbook's ConsolidatedData sheet - the export copies the case's
' "Alerted Transaction" sheet there, so the figures match the "Alerted
' Transaction Pivot" (Alert Information > Counterparty, sum of Transaction Amount).
'
' Selection:
'   - each distinct "Alert Information" value is a rule group (a combined
'     label such as "Rule A, Rule B" is its own group);
'   - value = a counterparty's total Transaction Amount within that group
'     (CR and DR added together);
'   - the top counterparty of every group is picked (one per group, groups in
'     order of their top value, at most 6);
'   - if that gives fewer than 3, the highest values left across all groups
'     fill up to 3;
'   - every pick is a different counterparty; blank counterparties are skipped
'     (and reported when one would have been a group's top).
'
' Address: taken from the counterparty's side of each of its rows - DR rows
' use "Beneficiary Address", CR rows "Originator Address" (the same rule the
' Counterparty column itself is built with). The most frequent address is used;
' one that is only a country (e.g. "US") is left blank and reported.
'
' Names written to J18:J23 are cleaned for file names and merging (CleanCpName):
' "&" -> "and", characters Windows forbids in file names -> space, other
' punctuation removed, spaces collapsed. Two names that clean to the same text
' count as one counterparty. Amounts and addresses still use ECM's spelling.
'
' Never overwrites: if J18:J23 already has any name, nothing is written.
' Run it on its own with Alt+F8 > Sample_Counterparties.
' ==========================================================================

Private Const PWD As String = "p7ss"
Private Const MIN_PICKS As Long = 3
Private Const MAX_PICKS As Long = 6           ' Sheet1 has CounterParty 1-6

Public Sub Sample_Counterparties()
    Dim note As String
    note = FillCounterparties()
    MsgBox IIf(note = "", "No counterparties picked.", note), vbInformation, "Sample counterparties"
End Sub

' Picks and writes the counterparties. Returns a readable summary (or the
' reason nothing was written). Never raises an error.
Public Function FillCounterparties() As String
    Dim home As Worksheet, cd As Worksheet
    Dim cRule As Long, cCp As Long, cAmt As Long, cDir As Long
    Dim cBenAddr As Long, cBenCty As Long, cOrgAddr As Long, cOrgCty As Long
    Dim lastRow As Long, r As Long, i As Long, j As Long
    Dim rule As String, cp As String, amt As Double, key As Variant
    Dim pairVal As Object, ruleTop As Object, blankTop As Object, rules As Collection
    Dim picks As New Collection, pickRule As New Collection, picked As Object, origs As New Collection, cpClean As String
    Dim keys() As String, vals() As Double, n As Long, tmpS As String, tmpD As Double
    Dim note As String, addrNote As String, addr As String

    On Error GoTo Fail
    Set home = ThisWorkbook.Sheets("Sheet1")
    For r = 18 To 23
        If Trim$(CStr(home.Cells(r, "J").Value)) <> "" Then
            FillCounterparties = "Counterparties not changed - J18:J23 already has names."
            Exit Function
        End If
    Next r

    Set cd = ThisWorkbook.Sheets("ConsolidatedData")
    cRule = HeaderCol(cd, "Alert Information")
    cCp = HeaderCol(cd, "Counterparty")
    cAmt = HeaderCol(cd, "Transaction Amount")
    cDir = HeaderCol(cd, "Dr Cr")
    If cRule = 0 Or cCp = 0 Or cAmt = 0 Then
        FillCounterparties = "Counterparties not picked - ConsolidatedData has no " & _
            IIf(cRule = 0, "'Alert Information' ", "") & IIf(cCp = 0, "'Counterparty' ", "") & _
            IIf(cAmt = 0, "'Transaction Amount' ", "") & "column."
        Exit Function
    End If
    cBenAddr = HeaderCol(cd, "Beneficiary Address")
    cBenCty = HeaderCol(cd, "Beneficiary Country")
    cOrgAddr = HeaderCol(cd, "Originator Address")
    cOrgCty = HeaderCol(cd, "Originator Country")

    lastRow = cd.Cells(cd.Rows.Count, cCp).End(xlUp).Row
    If cd.Cells(cd.Rows.Count, cAmt).End(xlUp).Row > lastRow Then lastRow = cd.Cells(cd.Rows.Count, cAmt).End(xlUp).Row
    If lastRow < 2 Then FillCounterparties = "Counterparties not picked - no alerted transactions in ConsolidatedData.": Exit Function

    ' ---- totals per (rule group, counterparty) ----
    Set pairVal = CreateObject("Scripting.Dictionary"): pairVal.CompareMode = vbTextCompare
    Set blankTop = CreateObject("Scripting.Dictionary"): blankTop.CompareMode = vbTextCompare
    Set rules = New Collection
    For r = 2 To lastRow
        rule = Trim$(CStr(cd.Cells(r, cRule).Value))
        cp = Trim$(CStr(cd.Cells(r, cCp).Value))
        amt = NumOf(cd.Cells(r, cAmt).Value)
        If rule = "" Then rule = "(no rule)"
        If cp = "" Then
            blankTop(rule) = blankTop(rule) + amt
        Else
            key = rule & vbTab & cp
            If Not pairVal.Exists(key) Then pairVal(key) = 0
            pairVal(key) = pairVal(key) + amt
        End If
    Next r
    If pairVal.Count = 0 Then FillCounterparties = "Counterparties not picked - every alerted transaction has a blank counterparty.": Exit Function

    ' ---- all pairs, highest value first ----
    n = pairVal.Count
    ReDim keys(1 To n): ReDim vals(1 To n)
    i = 0
    For Each key In pairVal.keys
        i = i + 1: keys(i) = key: vals(i) = pairVal(key)
    Next key
    For i = 1 To n - 1
        For j = i + 1 To n
            If vals(j) > vals(i) Then
                tmpD = vals(i): vals(i) = vals(j): vals(j) = tmpD
                tmpS = keys(i): keys(i) = keys(j): keys(j) = tmpS
            End If
        Next j
    Next i

    ' ---- one per rule group (groups in order of their top value) ----
    Set picked = CreateObject("Scripting.Dictionary"): picked.CompareMode = vbTextCompare
    Set ruleTop = CreateObject("Scripting.Dictionary"): ruleTop.CompareMode = vbTextCompare
    For i = 1 To n                                   ' highest value first, so the first
        rule = Split(keys(i), vbTab)(0)              ' unpicked pair of a rule is its top
        cp = Split(keys(i), vbTab)(1)
        cpClean = CleanCpName(cp)
        If cpClean <> "" And Not ruleTop.Exists(rule) And picks.Count < MAX_PICKS Then
            If Not picked.Exists(cpClean) Then
                ruleTop(rule) = cp
                picked(cpClean) = True
                picks.Add cpClean: origs.Add cp: pickRule.Add rule
            End If
        End If
    Next i

    ' ---- fill up to MIN_PICKS by value ----
    For i = 1 To n
        If picks.Count >= MIN_PICKS Then Exit For
        cp = Split(keys(i), vbTab)(1)
        cpClean = CleanCpName(cp)
        If cpClean <> "" Then
            If Not picked.Exists(cpClean) Then
                picked(cpClean) = True
                picks.Add cpClean: origs.Add cp: pickRule.Add "(by value) " & Split(keys(i), vbTab)(0)
            End If
        End If
    Next i

    ' ---- write names and addresses ----
    Application.EnableEvents = True
    For i = 1 To picks.Count
        addr = CounterpartyAddress(cd, lastRow, origs(i), cCp, cDir, cBenAddr, cBenCty, cOrgAddr, cOrgCty, addrNote)
        PutCell home, "J" & (17 + i), picks(i)
        If addr <> "" Then PutCell home, "T" & (17 + i), addr
        note = note & "   " & i & ". " & picks(i) & "  " & FormatCurrency(PairTotal(pairVal, origs(i)), 2) & _
               IIf(picks(i) <> origs(i), "  (ECM: " & origs(i) & ")", "") & _
               IIf(addr <> "", "  |  " & addr, "") & IIf(addrNote <> "", "  (" & addrNote & ")", "") & vbCrLf
    Next i
    ProtectHome home

    ' ---- blank counterparties that would have topped a group ----
    For Each key In blankTop.keys
        If ruleTop.Exists(key) Then
            If blankTop(key) > pairVal(key & vbTab & ruleTop(key)) Then
                note = note & "   Note: " & FormatCurrency(blankTop(key), 2) & " in '" & key & _
                       "' has no counterparty name - check manually." & vbCrLf
            End If
        End If
    Next key

    FillCounterparties = "Counterparties picked (" & ruleTop.Count & " rule group" & IIf(ruleTop.Count = 1, "", "s") & "):" & vbCrLf & note
    Exit Function

Fail:
    FillCounterparties = "Counterparties not picked - error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    ProtectHome ThisWorkbook.Sheets("Sheet1")
    Application.EnableEvents = True
End Function

' Most frequent address on the counterparty's side of its rows: DR -> Beneficiary,
' otherwise Originator. note gets "country only (US)" / "n other addresses seen".
Private Function CounterpartyAddress(ByVal cd As Worksheet, ByVal lastRow As Long, ByVal cp As String, _
        ByVal cCp As Long, ByVal cDir As Long, ByVal cBenAddr As Long, ByVal cBenCty As Long, _
        ByVal cOrgAddr As Long, ByVal cOrgCty As Long, ByRef note As String) As String
    Dim seen As Object, r As Long, isDr As Boolean, a As String, cty As String, key As Variant
    Dim best As String, bestN As Long, countryOnly As String
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    note = ""
    For r = 2 To lastRow
        If StrComp(Trim$(CStr(cd.Cells(r, cCp).Value)), cp, vbTextCompare) = 0 Then
            isDr = False
            If cDir > 0 Then isDr = (UCase$(Trim$(CStr(cd.Cells(r, cDir).Value))) = "DR")
            If isDr Then
                a = CellText(cd, r, cBenAddr): cty = CellText(cd, r, cBenCty)
            Else
                a = CellText(cd, r, cOrgAddr): cty = CellText(cd, r, cOrgCty)
            End If
            If a <> "" And cty <> "" And UCase$(Right$(a, Len(cty))) <> UCase$(cty) Then a = a & " " & cty
            If a = "" Then a = cty
            If a <> "" Then
                If IsCountryOnly(a, cty) Then
                    countryOnly = a
                Else
                    seen(a) = seen(a) + 1
                End If
            End If
        End If
    Next r
    For Each key In seen.keys
        If seen(key) > bestN Then best = key: bestN = seen(key)
    Next key
    If best = "" Then
        If countryOnly <> "" Then note = "address: country only (" & countryOnly & ")" Else note = "no address in the data"
    ElseIf seen.Count > 1 Then
        note = (seen.Count - 1) & " other address" & IIf(seen.Count = 2, "", "es") & " seen"
    End If
    CounterpartyAddress = best
End Function

' Counterparty name safe for file names, PDF merging and searches: "&" -> "and",
' characters Windows forbids in file names -> space, other punctuation removed,
' letters (accented too), digits, spaces and hyphens kept; spaces collapsed.
Public Function CleanCpName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    s = Replace(s, "&", " and ")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9-]" Or ch = " " Then
            out = out & ch
        ElseIf AscW(ch) >= 192 And (UCase$(ch) <> LCase$(ch) Or AscW(ch) >= 12288) Then
            out = out & ch                           ' accented letters, CJK - not curly quotes or symbols
        ElseIf InStr("\/:*?""<>|" & vbTab & vbCr & vbLf, ch) > 0 Then
            out = out & " "
        End If                                       ' anything else is dropped
    Next i
    out = Application.WorksheetFunction.Trim(out)
    Do While Left$(out, 1) = "-": out = Trim$(Mid$(out, 2)): Loop
    Do While Right$(out, 1) = "-": out = Trim$(Left$(out, Len(out) - 1)): Loop
    CleanCpName = out
End Function

' An address with no street - just a country code or the country itself.
Private Function IsCountryOnly(ByVal a As String, ByVal cty As String) As Boolean
    a = Trim$(a)
    IsCountryOnly = (Len(a) <= 3) Or (cty <> "" And StrComp(a, cty, vbTextCompare) = 0)
End Function

Private Function PairTotal(ByVal pairVal As Object, ByVal cp As String) As Double
    Dim key As Variant
    For Each key In pairVal.keys
        If StrComp(Split(key, vbTab)(1), cp, vbTextCompare) = 0 Then PairTotal = PairTotal + pairVal(key)
    Next key
End Function

' Column of a row-1 header: exact name first, then one that starts with it.
Private Function HeaderCol(ByVal ws As Worksheet, ByVal name As String) As Long
    Dim c As Long, lastC As Long, h As String
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        If StrComp(Trim$(CStr(ws.Cells(1, c).Value)), name, vbTextCompare) = 0 Then HeaderCol = c: Exit Function
    Next c
    For c = 1 To lastC
        h = Trim$(CStr(ws.Cells(1, c).Value))
        If Len(h) > Len(name) And StrComp(Left$(h, Len(name)), name, vbTextCompare) = 0 Then HeaderCol = c: Exit Function
    Next c
End Function

Private Function CellText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    If c = 0 Then Exit Function
    CellText = Application.WorksheetFunction.Trim(Replace(Replace(CStr(ws.Cells(r, c).Value), vbCr, " "), vbLf, " "))
End Function

' Amount cell to a number ("$1,234.50", "(12.00)" and blanks handled).
Private Function NumOf(ByVal v As Variant) As Double
    Dim s As String, neg As Boolean
    If IsNumeric(v) And Not IsEmpty(v) Then NumOf = CDbl(v): Exit Function
    s = Trim$(CStr(v))
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then neg = True: s = Mid$(s, 2, Len(s) - 2)
    s = Replace(Replace(Replace(s, "$", ""), ",", ""), " ", "")
    If IsNumeric(s) Then NumOf = CDbl(s)
    If neg Then NumOf = -NumOf
End Function

' Sheet1's change handler re-protects after every edit, so unprotect before each write.
Private Sub PutCell(ByVal ws As Worksheet, ByVal cellAddr As String, ByVal v As String)
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=PWD
    ws.Unprotect Password:=PWD
    On Error GoTo 0
    ws.Range(cellAddr).Value = v
End Sub

Private Sub ProtectHome(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=PWD
    ThisWorkbook.Protect Password:=PWD, Structure:=True
End Sub
