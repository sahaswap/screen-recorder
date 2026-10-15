Attribute VB_Name = "Module3"
' Auto narratives (set by modECMImport): no "Export Complete!" message; the saved
' file's full path goes to gNarrLastFile.
Public gNarrQuiet As Boolean
Public gNarrLastFile As String

Sub ExportToWord()
Dim wdApp As Object
Dim wdDoc As Object
Dim ws As Worksheet
Dim lastRow As Long
Dim i As Long
Dim tagStr As String
Dim valStr As String
Dim masterShell As String
Dim shellRow As Long
Dim rng As Object

' --- NEW VARIABLES FOR SAVING ---
Dim ecmID As String
Dim AlertID As String
Dim desktopPath As String
Dim folderPath As String
Dim wordFileName As String
Dim CustName As String

' --- NEW VARIABLES FOR RFI TABLE & HEADER/FOOTER ---
Dim isRFI As Boolean
Dim isEscalation As Boolean
Dim decValue As String
Dim sec As Object
Dim hdrRange As Object
Dim ftrRange As Object
Dim docText As String
Dim infoPos As Long
Dim promptPos As Long
Dim questionsText As String
Dim qlist() As String
Dim qArray() As String
Dim qCount As Long
Dim k As Long
Dim lineText As String
Dim matchNum As Boolean
Dim dotPos As Long
Dim numPart As String
Dim tableRng As Object
Dim wdTable As Object
Dim rIdx As Long
Dim c1 As Object
Dim c2 As Object

' 1. Point the macro to Sheet7 (where your mapping table lives)
Set ws = ThisWorkbook.Sheets("Sheet7")

' Recompute [Rule Name] (K23) fresh, right before it's used below - see
' RefreshRuleNameTag for why this replaced the old in-sheet array formula.
RefreshRuleNameTag

' --- CAPTURE ID's FROM SHEET1 FOR NAMING ---
ecmID = ThisWorkbook.Sheets("Sheet1").Range("J9").Value
AlertID = ThisWorkbook.Sheets("Sheet1").Range("J10").Value

lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).row

' 2. Hunt for [MASTER_SHELL] in Column J
shellRow = 0
For i = 1 To lastRow
    If Trim(ws.Cells(i, "J").Value) = "[MASTER_SHELL]" Then
        shellRow = i
        Exit For
    End If
Next i

If shellRow = 0 Then
    MsgBox "Could not find [MASTER_SHELL] in Column J!", vbCritical
    Exit Sub
End If

' 3. Grab the template text from Column K of that row
masterShell = ws.Cells(shellRow, "K").Value

If masterShell = "" Or masterShell = "Select Template" Then
    MsgBox "Please select a template type on Sheet1 first!", vbExclamation
    Exit Sub
End If

' 4. Open Word
On Error Resume Next
Set wdApp = GetObject(, "Word.Application")
If wdApp Is Nothing Then
    Set wdApp = CreateObject("Word.Application")
End If
On Error GoTo 0

wdApp.Visible = True
' Keep the window VISIBLE so you can watch the narrative build live (as
' it did before). Only turn off Word's background spell-check / grammar-
' check / repagination - those passes, not the on-screen editing, are
' what made it slow. Screen updating stays ON. All restored at the end.
Dim prevSpell As Boolean, prevGram As Boolean, prevPag As Boolean
On Error Resume Next
prevSpell = wdApp.Options.CheckSpellingAsYouType
prevGram = wdApp.Options.CheckGrammarAsYouType
prevPag = wdApp.Options.Pagination
wdApp.Options.CheckSpellingAsYouType = False
wdApp.Options.CheckGrammarAsYouType = False
wdApp.Options.Pagination = False
On Error GoTo 0
Set wdDoc = wdApp.Documents.Add

' 5. Paste the template into Word
wdDoc.content.text = masterShell

' 6. Loop through Column J (Tags) and replace with Column K (Data)
For i = 1 To lastRow
    tagStr = ws.Cells(i, "J").Value
    
    If tagStr = "[Program Description]" Then
        valStr = ws.Cells(i, "K").Value
    Else
        valStr = ws.Cells(i, "K").text
    End If

    If tagStr <> "" And tagStr <> "[MASTER_SHELL]" And tagStr <> "Template Tag" Then
        Set rng = wdDoc.content
        
        With rng.Find
            .ClearFormatting
            .text = tagStr
            .Forward = True
            .Wrap = 0
            
            Do While .Execute = True
                rng.text = valStr
                rng.Collapse Direction:=0
            Loop
        End With
    End If
Next i

' Apply formatting to all narrative text: Justify alignment, 1.15 line spacing
With wdDoc.content.ParagraphFormat
    .Alignment = 3 ' wdAlignParagraphJustify
    .LineSpacingRule = 5 ' wdLineSpaceMultiple
    .LineSpacing = wdApp.LinesToPoints(1.15)
    .SpaceBefore = 0
    .SpaceAfter = 0
End With

' --- NEW: RFI FORMATTING DELEGATED TO LOCAL SUBROUTINE ---
' J5 is the Decision cell (RFI / Escalation / Non-Escalation / ...).
decValue = UCase(Trim(ThisWorkbook.Sheets("Sheet1").Range("J5").Value))
isRFI = (decValue = "RFI")
' Escalation only - deliberately excludes "Non-Escalation".
isEscalation = (InStr(decValue, "ESCALAT") > 0 And InStr(decValue, "NON") = 0)
If isRFI Then
    FormatRFIDocument wdApp, wdDoc
End If

' --------------------------------------------------------
' 7. AUTO-CREATE FOLDER & SAVE WORD (.docx)
' --------------------------------------------------------
' Using your original shell path logic with backslash fixes
desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
folderPath = desktopPath & ecmID & "\"

' Create folder if it doesn't exist
If Len(Dir(folderPath, vbDirectory)) = 0 Then
    MkDir folderPath
End If

' Define file name
CustName = ThisWorkbook.Sheets("Sheet1").Range("J13").Value

If isRFI Then
    wordFileName = ecmID & "_" & AlertID & "_" & CustName & "_RFI QUESTIONS.docx"
ElseIf isEscalation Then
    wordFileName = ecmID & "_" & AlertID & "_" & CustName & "_Escalation Narrative.docx"
Else
    wordFileName = ecmID & "_" & AlertID & "_" & CustName & "_Alert Write-Up.docx"
End If

' Save the file
On Error Resume Next
wdDoc.SaveAs2 fileName:=folderPath & wordFileName, FileFormat:=12
If Err.Number <> 0 Then
    MsgBox "Warning: Word Auto-save failed. Check if file is already open.", vbExclamation
    Err.Clear
Else
    gNarrLastFile = folderPath & wordFileName
    ' Auto mode: close the saved document so the case folder can be zipped
    ' (Word keeps open documents locked).
    If gNarrQuiet Then wdDoc.Close SaveChanges:=False
End If
On Error GoTo 0

' --- RFI decisions also get a Pre RFI Alert Write-Up companion doc ---
' A fill-in narrative shell (customer OSDD profile + counterparty
' review) saved alongside the questions doc for the analyst to
' complete from their OSDD results.
If isRFI Then
    GeneratePreRFIWriteUp wdApp, folderPath, ecmID, AlertID, CustName
End If

' ==========================================
' 8. UPDATE SHARED MASTER TRACKER (SHEET3) - DEFERRED
'    Scheduled to run ~1s later in the background (modTrackerPush) so the
'    analyst gets their documents immediately instead of waiting on the
'    slow ghost-Excel / OneDrive write. The row still gets written.
' ==========================================
On Error Resume Next
Application.OnTime Now + TimeSerial(0, 0, 1), "PushRFITracker_Deferred"
On Error GoTo 0
' ==========================================

' Restore Word's background checks now that the documents are built
' (the window stayed visible throughout, so nothing to reveal here).
On Error Resume Next
wdApp.Options.CheckSpellingAsYouType = prevSpell
wdApp.Options.CheckGrammarAsYouType = prevGram
wdApp.Options.Pagination = prevPag
wdApp.Activate
On Error GoTo 0

If Not gNarrQuiet Then MsgBox "Export Complete! File saved to: " & folderPath, vbInformation

Set wdDoc = Nothing
Set wdApp = Nothing
End Sub

Public Sub FormatRFIDocument(ByVal wdApp As Object, ByVal wdDoc As Object)
    Dim sec As Object
    Dim hdrRange As Object
    Dim ftrRange As Object
    Dim docText As String
    Dim infoPos As Long
    Dim promptPos As Long
    Dim questionsText As String
    Dim qlist() As String
    Dim qArray() As String
    Dim qCount As Long
    Dim k As Long
    Dim lineText As String
    Dim matchNum As Boolean
    Dim dotPos As Long
    Dim numPart As String
    Dim tableRng As Object
    Dim wdTable As Object
    Dim rIdx As Long
    Dim qRow As Long
    Dim aRow As Long
    Dim c1 As Object
    Dim c2 As Object
    Dim bodyRng As Object
    Dim endRng As Object
    Dim titleRng As Object
    Dim p As Object
    Dim pIdx As Long
    Dim pText As String
    Dim r As Object
    Dim firstChar As String
    
    ' The Common template omits the page header + footer that Wise /
    ' Airwallex carry. Everything else is formatted identically.
    Dim isCommonTpl As Boolean
    isCommonTpl = (UCase(Trim(ThisWorkbook.Sheets("Sheet1").Range("J6").Value)) = "COMMON")

    ' 0. Set Page Margins to 1.0 Inch (72 points)
    With wdDoc.PageSetup
        .TopMargin = 72
        .BottomMargin = 72
        .LeftMargin = 72
        .RightMargin = 72
    End With
    
    ' Set Corporate Styling for all body text: Arial 11 (Line spacing and alignment set globally)
    With wdDoc.content.ParagraphFormat
        .LineSpacingRule = 5 ' wdLineSpaceMultiple
        .LineSpacing = wdApp.LinesToPoints(1.15)
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
    wdDoc.content.Font.Name = "Arial"
    wdDoc.content.Font.Size = 11
    wdDoc.content.Font.Color = RGB(0, 0, 0)
    
    ' 1. Set page header and footer. The Common template gets an
    '    explicitly EMPTY header + footer (no compliance banner, no page
    '    numbers). We still touch every section's header/footer rather
    '    than skip it, so the section is properly initialized - just
    '    skipping it left the first page laid out so the top of the body
    '    (title / To / Cc / Subject) rendered up in the margin, out of
    '    view. The other templates get the standard header/footer.
    For Each sec In wdDoc.Sections
        If isCommonTpl Then
            sec.Headers(1).Range.text = ""
            sec.Footers(1).Range.text = ""
        Else
            ' Right-aligned header
            Set hdrRange = sec.Headers(1).Range
            hdrRange.text = "Financial Crimes Compliance Unit" & vbCrLf & "AML / TM Department"
            hdrRange.ParagraphFormat.Alignment = 2 ' wdAlignParagraphRight
            hdrRange.ParagraphFormat.LineSpacingRule = 0
            hdrRange.ParagraphFormat.SpaceAfter = 0
            hdrRange.Font.Name = "Arial"
            hdrRange.Font.Size = 9
            hdrRange.Font.Color = RGB(0, 0, 0)

            ' Centered page numbers in footer
            Set ftrRange = sec.Footers(1).Range
            ftrRange.text = ""
            ftrRange.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
            ftrRange.ParagraphFormat.LineSpacingRule = 0
            ftrRange.ParagraphFormat.SpaceAfter = 0
            ftrRange.Font.Name = "Arial"
            ftrRange.Font.Size = 10
            ftrRange.Font.Color = RGB(0, 0, 0)

            ' Add page fields dynamically (33 = wdFieldPage, 26 = wdFieldNumPages)
            ftrRange.Fields.Add Range:=ftrRange, Type:=33

            Set ftrRange = sec.Footers(1).Range
            ftrRange.Collapse Direction:=0
            ftrRange.text = " of "

            Set ftrRange = sec.Footers(1).Range
            ftrRange.Collapse Direction:=0
            ftrRange.Fields.Add Range:=ftrRange, Type:=26

            Set ftrRange = sec.Footers(1).Range
            ftrRange.InsertBefore "Page "
        End If
    Next sec

    ' 2. Strip compliance header from document body (including surrounding whitespace/tabs)
    Set bodyRng = wdDoc.content
    With bodyRng.Find
        .ClearFormatting
        .text = "Financial Crimes Compliance Unit"
        .REPLACEMENT.text = ""
        .Execute Replace:=2 ' wdReplaceAll
    End With
    
    Set bodyRng = wdDoc.content
    With bodyRng.Find
        .ClearFormatting
        .text = "AML / TM Department"
        .REPLACEMENT.text = ""
        .Execute Replace:=2
    End With
    
    ' Clean up starting whitespace in body (tabs, spaces, carriage returns)
    Do While Len(wdDoc.content.text) > 0 And (Left(wdDoc.content.text, 1) = vbCr Or Left(wdDoc.content.text, 1) = vbLf Or Left(wdDoc.content.text, 1) = vbTab Or Left(wdDoc.content.text, 1) = " ")
        wdDoc.Range(0, 1).Delete
    Loop
    
    ' 3. Insert and Format Title centered and underlined at top.
    '    Applies to all templates (Common included): the letter opens with
    '    the "Request for Information (RFI)" banner, then To / Cc / Subject.
    If InStr(wdDoc.content.text, "Request for Information (RFI)") <> 1 Then
        Set titleRng = wdDoc.Range(0, 0)
        titleRng.text = "Request for Information (RFI)" & vbCrLf
    Else
        Set titleRng = wdDoc.Paragraphs(1).Range
    End If
    titleRng.ParagraphFormat.Alignment = 1 ' Center
    titleRng.ParagraphFormat.SpaceBefore = 0
    titleRng.ParagraphFormat.SpaceAfter = 12
    titleRng.ParagraphFormat.LineSpacingRule = 0 ' Single spacing
    titleRng.Font.Name = "Arial"
    titleRng.Font.Size = 11
    titleRng.Font.bold = True
    titleRng.Font.Underline = 1 ' Underline
    titleRng.Font.Color = RGB(0, 0, 0)
    
    ' 4. Clean up mail routing keywords
    Set bodyRng = wdDoc.content
    With bodyRng.Find
        .ClearFormatting
        .text = "send cfsbencrypt"
        .REPLACEMENT.text = ""
        .Execute Replace:=2
    End With
    
    ' 5. Convert text bullets to native Word bullet points
    For pIdx = wdDoc.Paragraphs.count To 1 Step -1
        Set p = wdDoc.Paragraphs(pIdx)
        pText = Trim(p.Range.text)
        If Len(pText) > 0 Then
            firstChar = Left(pText, 1)
            If firstChar = ChrW(8226) Or firstChar = Chr(149) Then
                Set r = p.Range
                Do While r.Characters.count > 0
                    firstChar = r.Characters(1).text
                    If firstChar = ChrW(8226) Or firstChar = Chr(149) Or firstChar = " " Or firstChar = vbTab Then
                        r.Characters(1).Delete
                    Else
                        Exit Do
                    End If
                Loop
                p.Range.ListFormat.ApplyBulletDefault
                ' Apply styling: Arial 11, line spacing 1
                p.Range.Font.Name = "Arial"
                p.Range.Font.Size = 11
                p.Range.ParagraphFormat.LineSpacingRule = 0
                p.Range.ParagraphFormat.SpaceAfter = 0
            End If
        End If
    Next pIdx
    
    ' 6. Extract questions list and build the Word Table
    docText = wdDoc.content.text
    
    infoPos = InStr(docText, "Information Requested:")
    promptPos = InStr(docText, "Your prompt response")
    
    If infoPos > 0 And promptPos > infoPos Then
        questionsText = Mid(docText, infoPos, promptPos - infoPos)
        
        ' Normalize line breaks to vbCr before splitting
        questionsText = Replace(questionsText, vbCrLf, vbCr)
        questionsText = Replace(questionsText, vbLf, vbCr)
        questionsText = Replace(questionsText, Chr(11), vbCr)
        
        ' Parse questions
        qlist = Split(questionsText, vbCr)
        
        ReDim qArray(1 To 1)
        qCount = 0
        
        For k = LBound(qlist) To UBound(qlist)
            lineText = Trim(qlist(k))
            lineText = Replace(lineText, vbLf, "")
            lineText = Replace(lineText, vbTab, "")
            
            matchNum = False
            If Len(lineText) > 2 Then
                dotPos = InStr(lineText, ".")
                If dotPos > 1 And dotPos < 5 Then
                    numPart = Left(lineText, dotPos - 1)
                    If IsNumeric(numPart) Then
                        matchNum = True
                        lineText = Trim(Mid(lineText, dotPos + 1))
                    End If
                End If
            End If
            
            If matchNum And lineText <> "" Then
                qCount = qCount + 1
                ReDim Preserve qArray(1 To qCount)
                
                ' Replace slash with newline ONLY for Entity / Individual
                If InStr(lineText, "Entity:") > 0 And InStr(lineText, "Individual:") > 0 Then
                    lineText = Replace(lineText, " / Individual:", vbCrLf & "Individual:")
                    lineText = Replace(lineText, "/ Individual:", vbCrLf & "Individual:")
                    lineText = Replace(lineText, " /Individual:", vbCrLf & "Individual:")
                    lineText = Replace(lineText, "/Individual:", vbCrLf & "Individual:")
                End If
                
                qArray(qCount) = lineText
            End If
        Next k
        
        If qCount > 0 Then
            ' Inject table replacing the raw text list
            Set tableRng = wdDoc.content
            
            With tableRng.Find
                .ClearFormatting
                .text = "Information Requested:"
                If .Execute Then
                    Set endRng = wdDoc.content
                    With endRng.Find
                        .ClearFormatting
                        .text = "Your prompt response"
                        If .Execute Then
                            tableRng.End = endRng.Start
                            tableRng.text = ""
                            
                            ' Create table with double rows (1 Question + 1 Answer row per question)
                            Set wdTable = wdDoc.Tables.Add(Range:=tableRng, NumRows:=qCount * 2, NumColumns:=2)
                            
                            ' Table Border Formatting (full grid in gray)
                            wdTable.Borders.Enable = True
                            wdTable.Borders(-1).Color = RGB(127, 127, 127) ' Top
                            wdTable.Borders(-2).Color = RGB(127, 127, 127) ' Left
                            wdTable.Borders(-3).Color = RGB(127, 127, 127) ' Bottom
                            wdTable.Borders(-4).Color = RGB(127, 127, 127) ' Right
                            wdTable.Borders(-5).Color = RGB(127, 127, 127) ' Inside Horizontal
                            wdTable.Borders(-6).Color = RGB(127, 127, 127) ' Inside Vertical
                            wdTable.Borders(-6).LineStyle = 1              ' Visible vertical lines
                            
                            ' Columns widths: (Col 1 = 0.5 inches / 36 points, Col 2 = 6.0 inches / 432 points)
                            wdTable.Columns(1).Width = 36
                            wdTable.Columns(2).Width = 432
                            
                            ' Padding
                            wdTable.TopPadding = 5
                            wdTable.BottomPadding = 5
                            wdTable.LeftPadding = 8
                            wdTable.RightPadding = 8
                            
                            ' Populate Cells
                            For rIdx = 1 To qCount
                                qRow = 2 * rIdx - 1
                                aRow = 2 * rIdx
                                
                                ' -- Question Row --
                                Set c1 = wdTable.cell(qRow, 1)
                                Set c2 = wdTable.cell(qRow, 2)
                                
                                c1.Range.text = "(" & rIdx & ")"
                                c1.Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                                c1.Range.ParagraphFormat.SpaceAfter = 0
                                c1.Range.ParagraphFormat.LineSpacingRule = 0
                                c1.Range.Font.bold = True
                                c1.Range.Font.Name = "Arial"
                                c1.Range.Font.Size = 11
                                c1.Range.Font.Color = RGB(0, 0, 0)
                                c1.Shading.BackgroundPatternColor = RGB(218, 233, 247) ' #DAE9F7 exact blue
                                
                                c2.Range.text = qArray(rIdx)
                                c2.Range.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
                                c2.Range.ParagraphFormat.SpaceAfter = 0
                                c2.Range.ParagraphFormat.LineSpacingRule = 0
                                c2.Range.Font.bold = False
                                c2.Range.Font.Name = "Arial"
                                c2.Range.Font.Size = 11
                                c2.Range.Font.Color = RGB(0, 0, 0)
                                c2.Shading.BackgroundPatternColor = RGB(218, 233, 247) ' #DAE9F7 exact blue
                                
                                ' -- Answer Row --
                                Set c1 = wdTable.cell(aRow, 1)
                                Set c2 = wdTable.cell(aRow, 2)
                                
                                c1.Range.text = ""
                                c1.Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                                c1.Range.ParagraphFormat.SpaceAfter = 0
                                c1.Range.ParagraphFormat.LineSpacingRule = 0
                                c1.Range.Font.Name = "Arial"
                                c1.Range.Font.Size = 11
                                c1.Shading.BackgroundPatternColor = RGB(255, 255, 255) ' White
                                
                                c2.Range.text = ""
                                c2.Range.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
                                c2.Range.ParagraphFormat.SpaceAfter = 0
                                c2.Range.ParagraphFormat.LineSpacingRule = 0
                                c2.Range.Font.Name = "Arial"
                                c2.Range.Font.Size = 11
                                c2.Range.Font.Color = RGB(0, 0, 0)
                                c2.Shading.BackgroundPatternColor = RGB(255, 255, 255) ' White
                            Next rIdx
                        End If
                    End With
                End If
            End With
        End If
    End If

    ' Leave the cursor at the very top so the letter opens showing
    ' To / Cc / Subject, not scrolled down to wherever building ended
    ' (that scroll position is why the Common doc "started" at Good
    ' Morning with the top lines out of view).
    On Error Resume Next
    wdDoc.Activate
    wdApp.Selection.HomeKey Unit:=6      ' wdStory - top of document
    wdApp.ActiveWindow.ScrollIntoView wdDoc.Range(0, 0), True
    On Error GoTo 0
End Sub

' ================================================================
' Pre RFI Alert Write-Up companion document.
'
' Generated only when the Decision (J5) is RFI, alongside the RFI
' questions doc. It's a fill-in narrative shell: the "xxxxx"
' placeholders and CP1/CP2/CP3 lines are meant to be completed by
' the analyst from their OSDD search results. Saved as
'   {ECM}_{AlertID}_{CustName}_Pre RFI Alert Write-Up.docx
' in the same case folder as everything else.
' ================================================================
Private Sub GeneratePreRFIWriteUp(ByVal wdApp As Object, ByVal folderPath As String, _
                                  ByVal ecmID As String, ByVal AlertID As String, _
                                  ByVal CustName As String)
    On Error Resume Next
    Dim wdDoc2 As Object
    Set wdDoc2 = wdApp.Documents.Add
    If wdDoc2 Is Nothing Then Exit Sub

    ' Sheet7 tags are embedded so the shared tag->value loop below fills
    ' them from the sheet, exactly like the questions doc:
    '   [Subject Name] -> customer/subject searched
    '   [CP n]         -> that counterparty's name
    ' The remaining xxxxx / xxxxxxx spots stay as free-text placeholders
    ' for the analyst to complete from their OSDD results.
    Dim body As String
    body = "CUSTOMER PROFILE (OSDD search results)" & vbCr & _
           "An internet search for [Subject Name] returned results with an exact match found and it is a xxxxx company specializing in xxxxxxx. The company provides xxxxxx, Website (if available)" & vbCr & _
           "Internal records confirmed the KYC details as well (if applicable)." & vbCr & _
           "A negative news search identified xxxxxxxxxxxxx information."

    ' COUNTERPARTY REVIEW is built dynamically from Sheet1!J18:J23 - one
    ' line per counterparty actually present, in slot order (so gaps are
    ' handled). If no counterparties are listed, the whole section is
    ' left out rather than showing empty CP lines.
    Dim wsHome As Worksheet
    Set wsHome = ThisWorkbook.Sheets("Sheet1")
    Dim cpLines As String, rowIdx As Long, slot As Long
    For rowIdx = 18 To 23
        If Trim$(CStr(wsHome.Range("J" & rowIdx).Value)) <> "" Then
            slot = rowIdx - 17                      ' J18 -> CP 1 ... J23 -> CP 6
            If cpLines <> "" Then cpLines = cpLines & vbCr
            cpLines = cpLines & "[CP " & slot & "] (OSDD results)"
        End If
    Next rowIdx

    If cpLines <> "" Then
        body = body & vbCr & vbCr & _
               "COUNTERPARTY REVIEW" & vbCr & _
               "The following counterparties were selected for review based on transaction volume and monetary value observed in the alerted transaction activity:" & vbCr & _
               cpLines
    End If

    wdDoc2.content.text = body

    ' Resolve embedded Sheet7 tags into their values - same Column J/K
    ' mapping the questions doc uses, so [Subject Name]/[CP 1..3] fill in.
    ApplySheet7Tags wdDoc2

    ' Same body formatting as the main write-up: justified, 1.15 spacing.
    With wdDoc2.content.ParagraphFormat
        .Alignment = 3          ' wdAlignParagraphJustify
        .LineSpacingRule = 5    ' wdLineSpaceMultiple
        .LineSpacing = wdApp.LinesToPoints(1.15)
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With

    ' Bold the two section headers so the shell reads like a write-up.
    BoldHeadingInDoc wdDoc2, "CUSTOMER PROFILE (OSDD search results)"
    BoldHeadingInDoc wdDoc2, "COUNTERPARTY REVIEW"

    Dim writeUpName As String
    writeUpName = ecmID & "_" & AlertID & "_" & CustName & "_Pre RFI Alert Write-Up.docx"
    wdDoc2.SaveAs2 fileName:=folderPath & writeUpName, FileFormat:=12
    If Err.Number <> 0 Then
        MsgBox "Warning: Pre RFI Write-Up auto-save failed. Check if the file is already open.", vbExclamation
        Err.Clear
    End If

    On Error GoTo 0
End Sub

' Bolds the first occurrence of headingText in the given Word doc.
Private Sub BoldHeadingInDoc(ByVal wdDoc2 As Object, ByVal headingText As String)
    On Error Resume Next
    Dim rng As Object
    Set rng = wdDoc2.content
    With rng.Find
        .ClearFormatting
        .text = headingText
        .Forward = True
        .Wrap = 0
        If .Execute Then rng.Font.bold = True
    End With
End Sub

' Runs Sheet7's Column J (tag) -> Column K (value) replacement over the
' given Word doc - the exact same mapping ExportToWord applies to the
' questions doc. Kept as its own sub so the Pre RFI Write-Up resolves
' any Sheet7 tag ([Subject Name], [CP 1], [CP 2], [CP 3], etc.) the same
' way. Word's Find is case-insensitive, but spacing must match the sheet.
Private Sub ApplySheet7Tags(ByVal wdTarget As Object)
    On Error Resume Next
    Dim ws7 As Worksheet
    Set ws7 = ThisWorkbook.Sheets("Sheet7")
    If ws7 Is Nothing Then Exit Sub

    Dim lastRow As Long, i As Long
    Dim tagStr As String, valStr As String, rng As Object
    lastRow = ws7.Cells(ws7.Rows.count, "J").End(xlUp).row

    For i = 1 To lastRow
        tagStr = ws7.Cells(i, "J").Value

        If tagStr = "[Program Description]" Then
            valStr = ws7.Cells(i, "K").Value
        Else
            valStr = ws7.Cells(i, "K").text
        End If

        If tagStr <> "" And tagStr <> "[MASTER_SHELL]" And tagStr <> "Template Tag" Then
            Set rng = wdTarget.content
            With rng.Find
                .ClearFormatting
                .text = tagStr
                .Forward = True
                .Wrap = 0
                Do While .Execute = True
                    rng.text = valStr
                    rng.Collapse Direction:=0
                Loop
            End With
        End If
    Next i
    On Error GoTo 0
End Sub




' ==========================================================
' RefreshRuleNameTag - recomputes Sheet7's [Rule Name] tag (the cell in
' column K next to "[Rule Name]" in column J) directly in VBA, replacing
' the fragile in-sheet array formula that used to live there:
'
'   =LET(AllAlerts, TEXTJOIN(",",TRUE,FILTER(ConsolidatedData!C2:C1000,
'        ConsolidatedData!C2:C1000<>"")),
'        SplitNames, UNIQUE(TRIM(TEXTSPLIT(AllAlerts,,","))),
'        RuleCodes, XLOOKUP(SplitNames, Rule_DB!C:C, Rule_DB!B:B,
'        "Code Not Found"),
'        TEXTJOIN(", ",TRUE, RuleCodes & " - " & SplitNames))
'
' That formula had two real bugs:
'   1. It hardcoded ConsolidatedData column C rather than finding the
'      "Alert Information" header by name, so it silently pointed at the
'      wrong column if a source file's column order ever shifted.
'   2. When TWO OR MORE distinct alert scenarios were present in one
'      case, XLOOKUP's array result did not reliably zip one-to-one
'      against the array of distinct names - the second (and later)
'      scenario could end up paired with the FIRST scenario's rule code
'      instead of its own. A genuine text mismatch against Rule_DB also
'      surfaced as the raw "Code Not Found" fallback, which then showed
'      up as a leaked calculation-error number (e.g. "-2363053") once
'      that error propagated through the rest of the array formula.
'
' This version resolves each unique alert scenario ONE AT A TIME in a
' plain loop - no array-broadcasting is possible - matches Rule_DB
' robustly (trimmed, case-insensitive) instead of a brittle exact array
' match, and writes a clearly human-readable flag for anything that
' still doesn't match, instead of letting a raw error leak into the
' generated narrative. Assigning a plain .Value to K23 replaces the old
' formula outright the first time this runs - nothing needs to be
' deleted by hand in the sheet.
'
' One ConsolidatedData row can also list MORE THAN ONE triggered rule in
' its Alert Information cell, comma-separated (e.g. "Rule A, Rule B"),
' when a single transaction trips multiple rules at once. SplitAlertCell
' below breaks each cell into its real individual rule names before
' matching, instead of treating a multi-rule cell as one giant unmatched
' name - which is what previously produced BOTH a "Rule Not Found" AND a
' duplicate-looking second copy of a rule that actually did match fine
' on its own elsewhere.
' ==========================================================
Public Sub RefreshRuleNameTag()
    On Error Resume Next
    Dim wsData As Worksheet, wsRule As Worksheet, wsTag As Worksheet
    Set wsData = ThisWorkbook.Sheets("ConsolidatedData")
    Set wsRule = ThisWorkbook.Sheets("Rule_DB")
    Set wsTag = ThisWorkbook.Sheets("Sheet7")
    If wsData Is Nothing Or wsRule Is Nothing Or wsTag Is Nothing Then Exit Sub

    ' locate "Alert Information" in ConsolidatedData by HEADER NAME, not a
    ' hardcoded column letter - this is what made the old formula fragile.
    Dim AlertCol As Long, LastDataRow As Long
    AlertCol = 0
    AlertCol = wsData.Rows(1).Find(What:="Alert Information", LookAt:=xlPart).Column
    If AlertCol = 0 Then Exit Sub
    LastDataRow = wsData.Cells(wsData.Rows.count, AlertCol).End(xlUp).row
    If LastDataRow < 2 Then Exit Sub

    ' locate Rule_DB's Rule Code / Rule Name columns by header name too.
    Dim codeCol As Long, nameCol As Long, lastRuleRow As Long
    codeCol = 0: nameCol = 0
    codeCol = wsRule.Rows(1).Find(What:="Rule Code", LookAt:=xlPart).Column
    nameCol = wsRule.Rows(1).Find(What:="Rule Name", LookAt:=xlPart).Column
    If codeCol = 0 Or nameCol = 0 Then Exit Sub
    lastRuleRow = wsRule.Cells(wsRule.Rows.count, nameCol).End(xlUp).row

    ' ---- 0. build a lookup of every VALID Rule_DB name up front (case-
    ' insensitive), so a raw Alert Information cell can be tested against
    ' real known rule names instead of guessing where to split it. ----
    Dim validNames As Object
    Set validNames = CreateObject("Scripting.Dictionary")
    Dim rr As Long, ruleName As String
    For rr = 2 To lastRuleRow
        ruleName = Trim$(CStr(wsRule.Cells(rr, nameCol).Value))
        If ruleName <> "" Then
            If Not validNames.Exists(UCase$(ruleName)) Then validNames.Add UCase$(ruleName), ruleName
        End If
    Next rr

    ' ---- 1. collect the DISTINCT set of individual alert-rule names.
    ' SplitAlertCell tries each cell WHOLE first (covers the common
    ' single-rule case, and protects a rule name that itself contains a
    ' comma, e.g. "...just below $10,000..."); only if that fails does it
    ' fall back to splitting the cell as a comma-separated list. ----
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim r As Long, v As String, dictKey As String
    Dim tok As Variant
    For r = 2 To LastDataRow
        v = Trim$(CStr(wsData.Cells(r, AlertCol).Value))
        If v <> "" Then
            For Each tok In SplitAlertCell(v, validNames)
                dictKey = UCase$(CStr(tok))
                If dictKey <> "" Then
                    If Not seen.Exists(dictKey) Then seen.Add dictKey, CStr(tok)   ' keep first-seen casing
                End If
            Next tok
        End If
    Next r
    If seen.count = 0 Then Exit Sub

    ' ---- 2. match each one against Rule_DB, ONE AT A TIME - no array zip,
    ' so there is no way for a second entry to inherit the first entry's
    ' code. Matching is trimmed + case-insensitive (StrComp vbTextCompare)
    ' rather than a brittle exact-text array match. ----
    Dim result As String, oneName As Variant, matched As Boolean
    Dim ruleCode As String
    result = ""
    For Each oneName In seen.Items
        matched = False
        For rr = 2 To lastRuleRow
            ruleName = Trim$(CStr(wsRule.Cells(rr, nameCol).Value))
            If StrComp(ruleName, CStr(oneName), vbTextCompare) = 0 Then
                ruleCode = Trim$(CStr(wsRule.Cells(rr, codeCol).Value))
                matched = True
                Exit For
            End If
        Next rr

        ' Clear, human-readable flag instead of a leaked error value - a
        ' reviewer will immediately notice this needs a manual check,
        ' rather than a cryptic number appearing in the narrative.
        If Not matched Then ruleCode = "Rule Not Found"

        If result <> "" Then result = result & ", "
        result = result & ruleCode & " - " & CStr(oneName)
    Next oneName

    ' ---- 3. write the corrected value into Sheet7's [Rule Name] row.
    ' Found by searching column J for the tag text (not a hardcoded row
    ' number), so this keeps working even if Sheet7's rows are ever
    ' reordered. Assigning .Value to a cell that currently holds a
    ' formula REPLACES the formula outright. ----
    Dim tagRow As Long
    tagRow = 0
    tagRow = wsTag.Columns("J").Find(What:="[Rule Name]", LookAt:=xlWhole).row
    If tagRow > 0 Then wsTag.Cells(tagRow, "K").Value = result

    On Error GoTo 0
End Sub

' Splits ONE raw Alert Information cell value into its individual
' triggered-rule names. Tries the WHOLE trimmed string against the known
' rule-name list first; only if that fails does it treat the cell as a
' comma-separated list, matching the LONGEST possible run of comma-joined
' pieces against a known rule name before falling back to a single
' unmatched piece - so a genuine list separator is never confused with a
' comma that's actually part of one rule's own name.
Private Function SplitAlertCell(ByVal v As String, ByVal validNames As Object) As Collection
    Dim outCol As New Collection
    v = Trim$(v)

    If validNames.Exists(UCase$(v)) Then
        outCol.Add v
        Set SplitAlertCell = outCol
        Exit Function
    End If

    Dim parts() As String
    parts = Split(v, ",")

    Dim i As Long, k As Long, m As Long, consumedTo As Long
    Dim candidate As String, gotMatch As Boolean
    i = LBound(parts)
    Do While i <= UBound(parts)
        gotMatch = False
        consumedTo = i
        For k = UBound(parts) To i Step -1
            candidate = parts(i)
            For m = i + 1 To k
                candidate = candidate & "," & parts(m)
            Next m
            candidate = Trim$(candidate)
            If validNames.Exists(UCase$(candidate)) Then
                outCol.Add candidate
                consumedTo = k
                gotMatch = True
                Exit For
            End If
        Next k
        If Not gotMatch Then outCol.Add Trim$(parts(i))
        i = consumedTo + 1
    Loop

    Set SplitAlertCell = outCol
End Function

