Attribute VB_Name = "Module9"
Option Explicit

' Set by modECMImport (Import ECM Info / auto-import) to run this export in a
' fixed mode ("LEGACY" / "EN" / "PIVOT") without showing the picker. Cleared on use.
Public gExportModeOverride As String
' gExportQuiet (auto-import batch): no completion or empty-window message, and
' the EN export's two saved files are closed instead of left open.
' gExportLastResult is set to "OK" when an export finishes.
Public gExportQuiet As Boolean
Public gExportLastResult As String

Sub Consolidated_AML_Workflow()

' ==========================================
' THE "JACKPOT" CONSOLIDATION WORKFLOW
' ==========================================
Application.EnableCancelKey = xlErrorHandler
On Error GoTo CancelHandler

' Safe default so CancelHandler can always restore Calculation even
' if an error fires before the real capture below ever runs.
Dim origCalc As XlCalculation
origCalc = xlCalculationAutomatic

Dim WbSource As Workbook, WsMaster As Worksheet, wsSource As Worksheet, wsHome As Worksheet
Dim wsRealCD As Worksheet
Dim LastRowSource As Long, LastRowMaster As Long, lastCol As Long
Dim HeaderCopied As Boolean

Dim HeaderCell As Range, HeaderRow As Long, HeadCol As Long
Dim DateHeader As Range, DateRange As Range, AmtHeader As Range, AmtRange As Range
Dim drCrCell As Range, benNameCell As Range, origNameCell As Range
Dim drCrCol As Long, benNameCol As Long, origNameCol As Long
Dim WsRawTemp As Worksheet, newWb As Workbook, wsExport As Worksheet, ws As Worksheet
Dim TransCol As Long, AlertCol As Long
Dim desktopPath As String, excelFileName As String, saveFolderPath As String, folderPath As String
Dim finalSavePath As String
Dim ecmID As String, AlertID As String
Dim wsPivot As Worksheet, ptCache As PivotCache, pt As PivotTable, ptRange As Range
Dim lastRowCP As Long, lastColCP As Long
Dim slash As String

Dim FSO As Object, objFolder As Object, objFile As Object
Dim fileFound As Boolean

' --- THE ULTIMATE PATH FIX ---
slash = Application.PathSeparator

' Break sheet grouping if active
On Error Resume Next
If Not ActiveWindow Is Nothing Then
    If ActiveWindow.SelectedSheets.count > 1 Then
        ActiveSheet.Select
    End If
End If
On Error GoTo CancelHandler

' ==========================================
' 1. SETUP & THE PROVEN FOLDER CONNECTION
' ==========================================

ThisWorkbook.Unprotect Password:="p7ss"
On Error Resume Next
ThisWorkbook.Sheets("Sheet1").Unprotect Password:="p7ss"
ThisWorkbook.Sheets("ConsolidatedData").Unprotect Password:="p7ss"
On Error GoTo CancelHandler

Set wsHome = ActiveWorkbook.Sheets("Sheet1")
ecmID = Trim(wsHome.Range("J9").Value)
AlertID = Trim(wsHome.Range("J10").Value)
If AlertID = "" Then AlertID = "ALERT"

If ecmID = "" Then
    MsgBox "Action Denied: ECM ID is missing in J9.", vbCritical, "Missing ID"
    Exit Sub
End If

' ==========================================
' 1b. EXPORT FORMAT PICKER - frmExportMode
' ==========================================
Dim exportMode As String
If gExportModeOverride <> "" Then
    exportMode = gExportModeOverride
    gExportModeOverride = ""
Else
    frmExportMode.Show vbModal

    If frmExportMode.userCancelled Then
        Unload frmExportMode
        Exit Sub
    End If
    exportMode = frmExportMode.SelectedMode   ' "LEGACY" / "EN" / "PIVOT"
    Unload frmExportMode
End If

Dim sourceFolderName As String
If exportMode = "PIVOT" Then
    sourceFolderName = "Pivot"
Else
    sourceFolderName = "Transaction Files"
End If

desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
saveFolderPath = desktopPath & slash & ecmID
folderPath = saveFolderPath & slash & sourceFolderName

Set FSO = CreateObject("Scripting.FileSystemObject")

' CHECK IF FOLDER EXISTS
If Not FSO.FolderExists(folderPath) Then
    MsgBox "Folder structure missing!" & vbCrLf & "Please click the START button to generate your '" & ecmID & "" & sourceFolderName & "' folder first.", vbCritical, "Folder Not Found"
    Exit Sub
End If

' CHECK IF FILES EXIST
Set objFolder = FSO.GetFolder(folderPath)
fileFound = False

For Each objFile In objFolder.Files
    If (InStr(1, objFile.Name, ".xls", vbTextCompare) > 0) And (Left(objFile.Name, 2) <> "~$") And (objFile.Name <> ThisWorkbook.Name) Then
        fileFound = True
        Exit For
    End If
Next objFile

If Not fileFound Then
    MsgBox "No Excel files found in the target folder!", vbExclamation, "Folder is Empty"
    Exit Sub
End If

' ==========================================
' 1c. PIVOT ANALYSIS EXPORT (own source folder: \Pivot)
' ==========================================
If exportMode = "PIVOT" Then
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    origCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wsPivScratch As Worksheet
    Dim pFile As Object, pWb As Workbook, pWs As Worksheet
    Dim pHeaderCell As Range, pHeaderRow As Long, pHeadCol As Long
    Dim pLastRowSource As Long, pLastRowMaster As Long, pHeaderCopied As Boolean

    SafeDeleteSheet ThisWorkbook, "TempPivotScratch"
    Set wsPivScratch = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsPivScratch.Name = "TempPivotScratch"

    pHeaderCopied = False
    For Each pFile In objFolder.Files
        If (InStr(1, pFile.Name, ".xls", vbTextCompare) > 0) And (Left(pFile.Name, 2) <> "~$") And (pFile.Name <> ThisWorkbook.Name) Then
            Set pWb = Workbooks.Open(pFile.path, ReadOnly:=True, UpdateLinks:=False)
            On Error Resume Next
            Set pWs = pWb.Sheets(1)
            On Error GoTo CancelHandler
            If Not pWs Is Nothing Then
                With pWs
                    Set pHeaderCell = .Cells.Find(What:="Transaction ID", LookIn:=xlValues, LookAt:=xlWhole)
                    If Not pHeaderCell Is Nothing Then
                        pHeaderRow = pHeaderCell.row
                        pHeadCol = pHeaderCell.Column
                        pLastRowSource = .Cells(.Rows.count, pHeadCol).End(xlUp).row
                        If pLastRowSource >= pHeaderRow Then
                            If Not pHeaderCopied Then
                                .Range(.Cells(pHeaderRow, pHeadCol), .UsedRange.SpecialCells(xlCellTypeLastCell)).Copy Destination:=wsPivScratch.Range("A1")
                                pHeaderCopied = True
                            Else
                                If pLastRowSource > pHeaderRow Then
                                    pLastRowMaster = wsPivScratch.Cells(wsPivScratch.Rows.count, "A").End(xlUp).row + 1
                                    .Range(.Cells(pHeaderRow + 1, pHeadCol), .UsedRange.SpecialCells(xlCellTypeLastCell)).Copy Destination:=wsPivScratch.Range("A" & pLastRowMaster)
                                End If
                            End If
                        End If
                    End If
                End With
            End If
            Application.CutCopyMode = False
            pWb.Close SaveChanges:=False
        End If
    Next pFile

    CleanTransactionData wsPivScratch

    Dim newWbPiv As Workbook
    Set newWbPiv = Workbooks.Add
    newWbPiv.Sheets(1).Name = "Pivot Data"
    If wsPivScratch.UsedRange.Cells.count > 0 Then
        wsPivScratch.UsedRange.Copy Destination:=newWbPiv.Sheets("Pivot Data").Range("A1")
    End If

    BuildEnPivots newWbPiv, "Pivot Data", "Pivot", "Pivot Data"

    Dim pSh As Long
    For pSh = newWbPiv.Sheets.count To 1 Step -1
        Select Case newWbPiv.Sheets(pSh).Name
            Case "Pivot Data", "Pivot"
                ' keep
            Case Else
                SafeDeleteSheet newWbPiv, newWbPiv.Sheets(pSh).Name
        End Select
    Next pSh

    SafeDeleteSheet ThisWorkbook, "TempPivotScratch"

    ' Same tidy pass every other mode's data sheets already get (Legacy's
    ' Raw Transactions/CP Selection/DeDupe, EN Network's 3 sheets,
    ' Lookback Transactions) - "Pivot Data" here was the one sheet still
    ' missing it.
    With newWbPiv.Sheets("Pivot Data").Cells
        .WrapText = False: .EntireColumn.AutoFit: .WrapText = True
        .EntireRow.AutoFit: .VerticalAlignment = xlTop
    End With

    Dim pivotFileName As String, pivotSavePath As String
    pivotFileName = ecmID & "_" & AlertID & "_Pivot Analysis.xlsx"
    pivotSavePath = folderPath & slash & pivotFileName
    CloseIfAlreadyOpen pivotSavePath
    SetSheetZoom85 newWbPiv, Array("Pivot Data", "Pivot")
    newWbPiv.SaveAs fileName:=pivotSavePath, FileFormat:=51

    newWbPiv.Sheets("Pivot Data").Activate

    Application.EnableCancelKey = xlInterrupt
    Application.Calculation = origCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    ThisWorkbook.Sheets("ConsolidatedData").Protect Password:="p7ss"
    ThisWorkbook.Sheets("Sheet1").Protect Password:="p7ss"
    ThisWorkbook.Protect Password:="p7ss", Structure:=True, Windows:=False
    Application.OnTime Now + TimeSerial(0, 0, 1), "PushTrxTracker_Deferred"
    On Error GoTo 0

    gExportLastResult = "OK"
    MsgBox "Pivot Analysis export complete!" & vbCrLf & _
        "ConsolidatedData was not touched." & vbCrLf & _
        "Saved to:" & vbCrLf & pivotSavePath, vbInformation, "Success"
    Exit Sub
End If

' Initialize the REAL ConsolidatedData sheet
On Error Resume Next
Set wsRealCD = wsHome.Parent.Sheets("ConsolidatedData")
On Error GoTo CancelHandler

If wsRealCD Is Nothing Then
    Set wsRealCD = wsHome.Parent.Sheets.Add(After:=wsHome.Parent.Sheets(wsHome.Parent.Sheets.count))
    wsRealCD.Name = "ConsolidatedData"
End If

' Disposable SCRATCH copy
wsHome.Parent.Unprotect Password:="p7ss"
SafeDeleteSheet wsHome.Parent, "TempConsolidatedScratch"
Set WsMaster = wsHome.Parent.Sheets.Add(After:=wsHome.Parent.Sheets(wsHome.Parent.Sheets.count))
WsMaster.Name = "TempConsolidatedScratch"
WsMaster.Visible = xlSheetVeryHidden

origCalc = Application.Calculation
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Application.ScreenUpdating = False
Application.DisplayAlerts = False
HeaderCopied = False

' ==========================================
' 2. COMBINE FILES
' ==========================================
For Each objFile In objFolder.Files
    If (InStr(1, objFile.Name, ".xls", vbTextCompare) > 0) And (Left(objFile.Name, 2) <> "~$") And (objFile.Name <> ThisWorkbook.Name) Then
        Set WbSource = Workbooks.Open(objFile.path, ReadOnly:=True, UpdateLinks:=False)
        On Error Resume Next
        Set wsSource = WbSource.Sheets(1)
        On Error GoTo CancelHandler

        If Not wsSource Is Nothing Then
            With wsSource
                Set HeaderCell = .Cells.Find(What:="Transaction ID", LookIn:=xlValues, LookAt:=xlWhole)
                If Not HeaderCell Is Nothing Then
                    HeaderRow = HeaderCell.row
                    HeadCol = HeaderCell.Column
                    LastRowSource = .Cells(.Rows.count, HeadCol).End(xlUp).row

                    If LastRowSource >= HeaderRow Then
                        If Not HeaderCopied Then
                            .Range(.Cells(HeaderRow, HeadCol), .UsedRange.SpecialCells(xlCellTypeLastCell)).Copy Destination:=WsMaster.Range("A1")
                            HeaderCopied = True
                        Else
                            If LastRowSource > HeaderRow Then
                                LastRowMaster = WsMaster.Cells(WsMaster.Rows.count, "A").End(xlUp).row + 1
                                .Range(.Cells(HeaderRow + 1, HeadCol), .UsedRange.SpecialCells(xlCellTypeLastCell)).Copy Destination:=WsMaster.Range("A" & LastRowMaster)
                            End If
                        End If
                    End If
                End If
            End With
        End If
        Application.CutCopyMode = False
        WbSource.Close SaveChanges:=False
    End If
Next objFile

Application.CutCopyMode = False

If WsMaster.UsedRange.Cells.count > 0 Then WsMaster.UsedRange.Value = WsMaster.UsedRange.Value

' ==========================================
' 2.5 SNAPSHOT RAW DATA (ROCK-SOLID RANGE CLONE)
' ==========================================
SafeDeleteSheet wsHome.Parent, "TempRawBackup"
Set WsRawTemp = wsHome.Parent.Sheets.Add(After:=wsHome)
WsRawTemp.Name = "TempRawBackup"

If WsMaster.UsedRange.Cells.count > 0 Then
    WsMaster.UsedRange.Copy Destination:=WsRawTemp.Range("A1")
End If
WsRawTemp.Visible = xlSheetVeryHidden

' ==========================================
' 3. AGGRESSIVE DATA CLEANUP
' ==========================================
Set DateHeader = WsMaster.Rows(1).Find(What:="Transaction Date", LookIn:=xlValues, LookAt:=xlPart)
If Not DateHeader Is Nothing Then
    LastRowMaster = WsMaster.Cells(WsMaster.Rows.count, DateHeader.Column).End(xlUp).row
    If LastRowMaster > 1 Then
        Set DateRange = WsMaster.Range(WsMaster.Cells(2, DateHeader.Column), WsMaster.Cells(LastRowMaster, DateHeader.Column))
        DateRange.TextToColumns Destination:=DateRange.Cells(1, 1), DataType:=xlDelimited, FieldInfo:=Array(Array(1, 3))
        DateRange.NumberFormat = "dddd, mmmm d, yyyy"
    End If
End If

Set AmtHeader = WsMaster.Rows(1).Find(What:="Transaction Amount", LookIn:=xlValues, LookAt:=xlPart)
If Not AmtHeader Is Nothing Then
    LastRowMaster = WsMaster.Cells(WsMaster.Rows.count, AmtHeader.Column).End(xlUp).row
    If LastRowMaster > 1 Then
        Set AmtRange = WsMaster.Range(WsMaster.Cells(2, AmtHeader.Column), WsMaster.Cells(LastRowMaster, AmtHeader.Column))
        AmtRange.Value = AmtRange.Value
        AmtRange.NumberFormat = "$#,##0.00"
    End If
End If

Set drCrCell = WsMaster.Rows(1).Find(What:="Dr Cr", LookAt:=xlPart)
Set benNameCell = WsMaster.Rows(1).Find(What:="Beneficiary Name", LookAt:=xlPart)
Set origNameCell = WsMaster.Rows(1).Find(What:="Originator Name", LookAt:=xlPart)

If Not drCrCell Is Nothing And Not benNameCell Is Nothing And Not origNameCell Is Nothing Then
    drCrCol = drCrCell.Column
    benNameCol = benNameCell.Column
    origNameCol = origNameCell.Column

    LastRowMaster = WsMaster.Cells(WsMaster.Rows.count, "A").End(xlUp).row
    lastCol = WsMaster.Cells(1, WsMaster.Columns.count).End(xlToLeft).Column + 1

    If LastRowMaster > 1 Then
        WsMaster.Cells(1, lastCol).Value = "Counterparty"
        WsMaster.Range(WsMaster.Cells(2, lastCol), WsMaster.Cells(LastRowMaster, lastCol)).FormulaR1C1 = _
        "=IF(RC" & drCrCol & "=""DR"", IF(RC" & benNameCol & "="""","""",RC" & benNameCol & "), IF(RC" & origNameCol & "="""","""",RC" & origNameCol & "))"

        WsMaster.Cells(1, lastCol - 1).Copy
        WsMaster.Cells(1, lastCol).PasteSpecial Paste:=xlPasteFormats
        WsMaster.Range(WsMaster.Cells(2, lastCol - 1), WsMaster.Cells(LastRowMaster, lastCol - 1)).Copy
        WsMaster.Range(WsMaster.Cells(2, lastCol), WsMaster.Cells(LastRowMaster, lastCol)).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    End If
End If

' ==========================================
' 4-EN. EN NETWORK EXPORT
' ==========================================
If exportMode = "EN" Then
    Dim aColLB As Long, aDateColLB As Long, scanLastLB As Long, rLB As Long
    Dim cvLB As Variant
    Dim lastAlertedLB As Date, haveAlertedLB As Boolean, lbStart As Date, lbEnd As Date
    Dim wsLB As Worksheet
    Dim lbSavedPath As String

    aColLB = 0: aDateColLB = 0
    On Error Resume Next
    aColLB = WsMaster.Rows(1).Find(What:="Is Alerted Transaction?", LookAt:=xlWhole).Column
    aDateColLB = WsMaster.Rows(1).Find(What:="Transaction Date", LookAt:=xlPart).Column
    On Error GoTo CancelHandler
    If aColLB = 0 Then
        MsgBox "EN Network export needs an 'Is Alerted Transaction?' column (exact name), but it wasn't found in the data.", vbCritical, "Column Not Found"
        GoTo CancelHandler
    End If

    Dim sArrLB As Variant, dArrLB As Variant
    scanLastLB = LastDataRow(WsMaster)
    haveAlertedLB = False
    If scanLastLB > 1 Then
        sArrLB = WsMaster.Range(WsMaster.Cells(2, aColLB), WsMaster.Cells(scanLastLB, aColLB)).Value
        If aDateColLB > 0 Then _
            dArrLB = WsMaster.Range(WsMaster.Cells(2, aDateColLB), WsMaster.Cells(scanLastLB, aDateColLB)).Value
        For rLB = 1 To scanLastLB - 1
            If Trim(CStr(ArrCell(sArrLB, rLB))) = "Yes" Then
                cvLB = ArrCell(dArrLB, rLB)
                If IsDate(cvLB) Then
                    If Not haveAlertedLB Then
                        lastAlertedLB = CDate(cvLB): haveAlertedLB = True
                    ElseIf CDate(cvLB) > lastAlertedLB Then
                        lastAlertedLB = CDate(cvLB)
                    End If
                End If
            End If
        Next rLB
    End If
    If Not haveAlertedLB Then
        MsgBox "No dated 'Yes' alerted transactions were found, so the EN Network export can't be built.", vbCritical, "No Alerted Rows"
        GoTo CancelHandler
    End If
    lbStart = DateSerial(Year(lastAlertedLB) - 1, Month(lastAlertedLB), 1)

    ' End of the MONTH the last alerted transaction falls in, not the
    ' alerted date itself. Last alerted 09/10 -> window ends 09/30.
    '
    ' Day 0 of the following month is the last day of this one, and
    ' DateSerial rolls a month of 13 over into January of the next year,
    ' so a December alert correctly gives 12/31 rather than erroring.
    ' This is the same idiom the Non Alerted window already uses below.
    '
    ' Previously this was the alerted date itself, which meant a
    ' non-alerted transaction later in the same month fell outside the
    ' lookback even though the Non Alerted export included it.
    lbEnd = DateSerial(Year(lastAlertedLB), Month(lastAlertedLB) + 1, 0)

    Set newWb = Workbooks.Add
    Set wsLB = newWb.Sheets(1)
    wsLB.Name = "Lookback Transactions"
    If WsMaster.UsedRange.Cells.count > 0 Then
        WsMaster.UsedRange.Copy Destination:=wsLB.Range("A1")
    End If

    FilterRowsFast wsLB, 0, "", aDateColLB, True, lbStart, lbEnd
    With wsLB.Cells
        .WrapText = False: .EntireColumn.AutoFit: .WrapText = True
        .EntireRow.AutoFit: .VerticalAlignment = xlTop
    End With

    BuildEnPivots newWb, "Lookback Transactions", "Pivot", "Lookback Transactions"

    For rLB = newWb.Sheets.count To 1 Step -1
        Select Case newWb.Sheets(rLB).Name
            Case "Lookback Transactions", "Pivot"
                ' keep
            Case Else
                SafeDeleteSheet newWb, newWb.Sheets(rLB).Name
        End Select
    Next rLB

    excelFileName = ecmID & "_" & AlertID & "_Lookback Transactions (" & _
        Format$(lbStart, "mm.dd.yyyy") & " to " & Format$(lbEnd, "mm.dd.yyyy") & ").xlsx"
    finalSavePath = saveFolderPath & slash & excelFileName
    CloseIfAlreadyOpen finalSavePath
    SetSheetZoom85 newWb, Array("Lookback Transactions", "Pivot")
    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=finalSavePath, FileFormat:=51
    Application.DisplayAlerts = True
    lbSavedPath = finalSavePath

    ' ---------------------------------------------------------------
    ' 4-AN. Alerted / Non-Alerted Trx File
    ' ---------------------------------------------------------------
    Dim wsAlertedEN As Worksheet, wsNonEN As Worksheet, wsRawEN As Worksheet
    Dim aColEN As Long, aDateColEN As Long
    Dim rEN As Long, scanLastEN As Long, wsTidy As Variant, cvEN As Variant
    Dim minAlerted As Date, maxAlerted As Date, haveAlerted As Boolean
    Dim winStartEN As Date, winEndEN As Date

    aColEN = 0: aDateColEN = 0
    On Error Resume Next
    aColEN = WsMaster.Rows(1).Find(What:="Is Alerted Transaction?", LookAt:=xlWhole).Column
    aDateColEN = WsMaster.Rows(1).Find(What:="Transaction Date", LookAt:=xlPart).Column
    On Error GoTo CancelHandler
    If aColEN = 0 Then
        MsgBox "EN Network export needs an 'Is Alerted Transaction?' column (exact name), but it wasn't found in the data.", vbCritical, "Column Not Found"
        GoTo CancelHandler
    End If

    Dim sArrEN As Variant, dArrEN As Variant
    scanLastEN = LastDataRow(WsMaster)
    haveAlerted = False
    If scanLastEN > 1 Then
        sArrEN = WsMaster.Range(WsMaster.Cells(2, aColEN), WsMaster.Cells(scanLastEN, aColEN)).Value
        If aDateColEN > 0 Then _
            dArrEN = WsMaster.Range(WsMaster.Cells(2, aDateColEN), WsMaster.Cells(scanLastEN, aDateColEN)).Value
        For rEN = 1 To scanLastEN - 1
            If Trim(CStr(ArrCell(sArrEN, rEN))) = "Yes" Then
                cvEN = ArrCell(dArrEN, rEN)
                If IsDate(cvEN) Then
                    If Not haveAlerted Then
                        minAlerted = CDate(cvEN): maxAlerted = CDate(cvEN): haveAlerted = True
                    Else
                        If CDate(cvEN) < minAlerted Then minAlerted = CDate(cvEN)
                        If CDate(cvEN) > maxAlerted Then maxAlerted = CDate(cvEN)
                    End If
                End If
            End If
        Next rEN
    End If
    If haveAlerted Then
        winStartEN = DateSerial(Year(minAlerted), Month(minAlerted), 1)
        winEndEN = DateSerial(Year(maxAlerted), Month(maxAlerted) + 1, 0)
    Else
        If Not gExportQuiet Then MsgBox "No 'Yes' alerted transactions were found, so the Non Alerted window can't be built. The Non Alerted sheet will be empty.", vbExclamation, "No Alerted Rows"
    End If

    Set newWb = Workbooks.Add
    Set wsRawEN = newWb.Sheets(1)
    wsRawEN.Name = "Raw Transactions"
    If WsRawTemp.UsedRange.Cells.count > 0 Then
        WsRawTemp.UsedRange.Copy Destination:=wsRawEN.Range("A1")
    End If
    CleanTransactionData wsRawEN

    Set wsAlertedEN = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.count))
    wsAlertedEN.Name = "Alerted Transaction"
    If WsMaster.UsedRange.Cells.count > 0 Then
        WsMaster.UsedRange.Copy Destination:=wsAlertedEN.Range("A1")
    End If
    FilterRowsFast wsAlertedEN, aColEN, "Yes", 0, False, 0, 0

    Dim hasNoInWin As Boolean
    hasNoInWin = False
    If haveAlerted And scanLastEN > 1 Then
        For rEN = 1 To scanLastEN - 1
            If Trim(CStr(ArrCell(sArrEN, rEN))) = "No" Then
                cvEN = ArrCell(dArrEN, rEN)
                If IsDate(cvEN) Then
                    If CDate(cvEN) >= winStartEN And CDate(cvEN) <= winEndEN Then hasNoInWin = True: Exit For
                End If
            End If
        Next rEN
    End If

    Set wsNonEN = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.count))
    wsNonEN.Name = "Non Alerted Transaction"
    If WsMaster.UsedRange.Cells.count > 0 Then
        WsMaster.UsedRange.Copy Destination:=wsNonEN.Range("A1")
    End If

    If Not haveAlerted Then
        wsNonEN.Cells.Clear
        wsNonEN.Range("A1").Value = "No alerted transactions were found, so the Non Alerted window could not be determined."
    ElseIf Not hasNoInWin Then
        wsNonEN.Cells.Clear
        wsNonEN.Range("A1").Value = "There were 0 non-alerted transactions during the alerted month(s) " & _
            Format$(winStartEN, "mm/dd/yyyy") & " to " & Format$(winEndEN, "mm/dd/yyyy") & "."
    Else
        FilterRowsFast wsNonEN, aColEN, "No", aDateColEN, True, winStartEN, winEndEN
    End If

    Dim dateColTidy As Long, lastRTidy As Long
    For Each wsTidy In Array("Raw Transactions", "Alerted Transaction", "Non Alerted Transaction")
        With newWb.Sheets(CStr(wsTidy)).Cells
            .WrapText = False: .EntireColumn.AutoFit: .WrapText = True
            .EntireRow.AutoFit: .VerticalAlignment = xlTop
        End With
        On Error Resume Next
        dateColTidy = 0
        dateColTidy = newWb.Sheets(CStr(wsTidy)).Rows(1).Find(What:="Transaction Date", LookAt:=xlPart).Column
        If dateColTidy > 0 Then
            lastRTidy = newWb.Sheets(CStr(wsTidy)).Cells(newWb.Sheets(CStr(wsTidy)).Rows.count, dateColTidy).End(xlUp).row
            If lastRTidy > 1 Then
                newWb.Sheets(CStr(wsTidy)).Range(newWb.Sheets(CStr(wsTidy)).Cells(2, dateColTidy), _
                    newWb.Sheets(CStr(wsTidy)).Cells(lastRTidy, dateColTidy)).NumberFormat = "m/d/yyyy"
            End If
        End If
        On Error GoTo CancelHandler
    Next wsTidy

    BuildEnPivots newWb, "Alerted Transaction", "Alerted Transaction Pivot", "Alerted Transaction"

    For rEN = newWb.Sheets.count To 1 Step -1
        Select Case newWb.Sheets(rEN).Name
            Case "Raw Transactions", "Alerted Transaction", "Alerted Transaction Pivot", "Non Alerted Transaction"
                ' keep
            Case Else
                SafeDeleteSheet newWb, newWb.Sheets(rEN).Name
        End Select
    Next rEN

    newWb.Sheets("Non Alerted Transaction").Move After:=newWb.Sheets(newWb.Sheets.count)

    SafeDeleteSheet wsHome.Parent, "TempRawBackup"

    excelFileName = ecmID & "_" & AlertID & "_Combined Alerted & Non Alerted Transactions.xlsx"
    finalSavePath = saveFolderPath & slash & excelFileName
    CloseIfAlreadyOpen finalSavePath
    SetSheetZoom85 newWb, Array("Raw Transactions", "Alerted Transaction", "Alerted Transaction Pivot", "Non Alerted Transaction")
    Application.DisplayAlerts = False
    newWb.SaveAs fileName:=finalSavePath, FileFormat:=51
    Application.DisplayAlerts = True

    wsRealCD.Cells.Clear
    newWb.Sheets("Alerted Transaction").UsedRange.Copy Destination:=wsRealCD.Range("A1")

    On Error Resume Next
    Module3.RefreshRuleNameTag
    On Error GoTo CancelHandler

    SafeDeleteSheet wsHome.Parent, "TempConsolidatedScratch"

    newWb.Sheets("Raw Transactions").Activate

    Application.EnableCancelKey = xlInterrupt
    Application.Calculation = origCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error Resume Next
    ThisWorkbook.Sheets("ConsolidatedData").Protect Password:="p7ss"
    ThisWorkbook.Sheets("Sheet1").Protect Password:="p7ss"
    ThisWorkbook.Protect Password:="p7ss", Structure:=True, Windows:=False
    Application.OnTime Now + TimeSerial(0, 0, 1), "PushTrxTracker_Deferred"
    On Error GoTo 0

    gExportLastResult = "OK"
    If gExportQuiet Then
        CloseSavedExport lbSavedPath
        CloseSavedExport finalSavePath
        Exit Sub
    End If
    MsgBox "EN Network export complete! Both files were generated:" & vbCrLf & vbCrLf & _
        "Lookback Transactions:" & vbCrLf & lbSavedPath & vbCrLf & vbCrLf & _
        "Alerted / Non-Alerted Transactions:" & vbCrLf & finalSavePath, vbInformation, "Success"
    Exit Sub
End If

' ==========================================
' 4. TWO-LAYER EXPORT & DEDUPE (LEGACY)
' ==========================================
Set newWb = Workbooks.Add

Set ws = newWb.Sheets(1)
ws.Name = "Raw Transactions"
If WsRawTemp.UsedRange.Cells.count > 0 Then
    WsRawTemp.UsedRange.Copy Destination:=ws.Range("A1")
End If

Set wsExport = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.count))
wsExport.Name = "CP Selection"
If WsMaster.UsedRange.Cells.count > 0 Then
    WsMaster.UsedRange.Copy Destination:=wsExport.Range("A1")
End If

TransCol = 0: AlertCol = 0
On Error Resume Next
TransCol = wsExport.Rows(1).Find(What:="Transaction ID", LookAt:=xlPart).Column
AlertCol = wsExport.Rows(1).Find(What:="Alert Information", LookAt:=xlPart).Column
On Error GoTo CancelHandler

If TransCol > 0 And AlertCol > 0 Then
    wsExport.UsedRange.RemoveDuplicates Columns:=Array(TransCol, AlertCol), Header:=xlYes
End If

Dim wsDeDupe As Worksheet
Set wsDeDupe = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.count))
wsDeDupe.Name = "DeDupe"
If wsExport.UsedRange.Cells.count > 0 Then
    wsExport.UsedRange.Copy Destination:=wsDeDupe.Range("A1")
End If

Set wsExport = wsDeDupe
TransCol = 0
On Error Resume Next
TransCol = wsExport.Rows(1).Find(What:="Transaction ID", LookAt:=xlPart).Column
On Error GoTo CancelHandler

If TransCol > 0 Then wsExport.UsedRange.RemoveDuplicates Columns:=Array(TransCol), Header:=xlYes

For Each ws In newWb.Sheets
    If ws.Name = "Raw Transactions" Or ws.Name = "CP Selection" Or ws.Name = "DeDupe" Then
        With ws.Cells
            .WrapText = False
            .EntireColumn.AutoFit
            .WrapText = True
            .EntireRow.AutoFit
            .VerticalAlignment = xlTop
        End With
    Else
        SafeDeleteSheet newWb, ws.Name
    End If
Next ws

' ==========================================
' 4.5 PIVOT TABLES (MULTI-SOURCE)
' ==========================================
Set wsExport = newWb.Sheets("CP Selection")
lastRowCP = wsExport.Cells(wsExport.Rows.count, "A").End(xlUp).row
lastColCP = wsExport.Cells(1, wsExport.Columns.count).End(xlToLeft).Column

If lastRowCP > 1 Then
    Set ptRange = wsExport.Range(wsExport.Cells(1, 1), wsExport.Cells(lastRowCP, lastColCP))
    Set wsPivot = newWb.Sheets.Add(After:=newWb.Sheets("Raw Transactions"))
    wsPivot.Name = "Pivot"

    Set ptCache = newWb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ptRange)

    ' ------------------------------------------
    ' PIVOT 1: SCENARIO PIVOT
    ' ------------------------------------------
    Set pt = ptCache.CreatePivotTable(TableDestination:=wsPivot.Range("A3"), TableName:="ScenarioPivot")
    On Error Resume Next
    With pt
        .TableStyle2 = "PivotStyleLight16"
        With .PivotFields("Alert Information"): .Orientation = xlRowField: .Position = 1: End With
        With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 2: End With
        With .PivotFields("Counterparty"): .Orientation = xlRowField: .Position = 3: End With
        .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount", xlSum
        .PivotFields("Sum of Transaction Amount").NumberFormat = "$#,#00.00"
        .AddDataField .PivotFields("Transaction Amount"), "Count of Transaction Amount", xlCount
        .RowAxisLayout xlCompactRow
        .PivotFields("Count of Transaction Amount").NumberFormat = "0"
        .PivotFields("Alert Information").AutoSort xlDescending, "Sum of Transaction Amount"
        .PivotFields("Counterparty").AutoSort xlDescending, "Sum of Transaction Amount"
    End With
    On Error GoTo CancelHandler

    ' ------------------------------------------
    ' SETUP SOURCE 2: "DeDupe"
    ' ------------------------------------------
    Dim lastRowDD As Long, lastColDD As Long, dtColDD As Long
    Dim ptRangeDeDupe As Range, ptCacheDeDupe As PivotCache

    Set wsDeDupe = newWb.Sheets("DeDupe")

    On Error Resume Next
    dtColDD = wsDeDupe.Rows(1).Find(What:="Transaction Date", LookAt:=xlPart).Column
    lastRowDD = wsDeDupe.Cells(wsDeDupe.Rows.count, dtColDD).End(xlUp).row
    If lastRowDD > 1 And dtColDD > 0 Then
        wsDeDupe.Range(wsDeDupe.Cells(2, dtColDD), wsDeDupe.Cells(lastRowDD, dtColDD)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
        wsDeDupe.Range(wsDeDupe.Cells(2, dtColDD), wsDeDupe.Cells(lastRowDD, dtColDD)).NumberFormat = "m/d/yyyy"
    End If
    On Error GoTo CancelHandler

    lastRowDD = wsDeDupe.Cells(wsDeDupe.Rows.count, "A").End(xlUp).row
    lastColDD = wsDeDupe.Cells(1, wsDeDupe.Columns.count).End(xlToLeft).Column

    If lastRowDD > 1 Then
        Set ptRangeDeDupe = wsDeDupe.Range(wsDeDupe.Cells(1, 1), wsDeDupe.Cells(lastRowDD, lastColDD))
        Set ptCacheDeDupe = newWb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=ptRangeDeDupe)

        ' PIVOT 2: TEMPORAL PIVOT
        Set pt = ptCacheDeDupe.CreatePivotTable(TableDestination:=wsPivot.Range("F3"), TableName:="TemporalPivot")
        On Error Resume Next
        With pt
            .TableStyle2 = "PivotStyleLight16"
            With .PivotFields("Transaction Date"): .Orientation = xlRowField: .Position = 1: End With
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount ", xlSum
            .PivotFields("Sum of Transaction Amount ").NumberFormat = "$#,#00.00"
            .AddDataField .PivotFields("Transaction Amount"), "Count of Transaction Amount ", xlCount
            .PivotFields("Count of Transaction Amount ").NumberFormat = "0"
        End With
        wsPivot.Range("F4").Group Start:=True, End:=True, Periods:=Array(False, False, False, True, True, False, True)

        On Error Resume Next
        pt.PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        On Error GoTo CancelHandler

        ' PIVOT 3: ENHANCED TEMPORAL PIVOT
        pt.TableRange2.Copy Destination:=wsPivot.Range("K3")
        Dim pt3 As PivotTable
        Set pt3 = wsPivot.Range("K3").PivotTable
        pt3.Name = "TemporalPivot_Enhanced"

        On Error Resume Next
        With pt3
            .PivotFields("Sum of Transaction Amount ").Orientation = xlHidden
            .PivotFields("Count of Transaction Amount ").Orientation = xlHidden
            With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 4: End With
            .AddDataField .PivotFields("Transaction Amount"), "No of Trx  ", xlCount
            .PivotFields("No of Trx  ").NumberFormat = "0"
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount  ", xlSum
            .PivotFields("Sum of Transaction Amount  ").NumberFormat = "$#,#00.00"
            .PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        End With
        On Error GoTo CancelHandler

        ' PIVOT 4: DR/CR INVERTED PIVOT
        pt.TableRange2.Copy Destination:=wsPivot.Range("Q3")
        Dim pt4 As PivotTable
        Set pt4 = wsPivot.Range("Q3").PivotTable
        pt4.Name = "DrCrTemporalPivot"

        On Error Resume Next
        With pt4
            .PivotFields("Sum of Transaction Amount ").Orientation = xlHidden
            .PivotFields("Count of Transaction Amount ").Orientation = xlHidden
            With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 1: End With
            .AddDataField .PivotFields("Transaction Amount"), "No of Trx   ", xlCount
            .PivotFields("No of Trx   ").NumberFormat = "0"
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount   ", xlSum
            .PivotFields("Sum of Transaction Amount   ").NumberFormat = "$#,#00.00"
            .PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        End With
        On Error GoTo CancelHandler

        On Error Resume Next
        BulletproofDateFormat wsPivot
        HighlightDrCrRows pt3, wsPivot
        HighlightDrCrRows pt4, wsPivot
        On Error GoTo CancelHandler
    End If

    wsPivot.Columns("A:W").AutoFit
End If

' ==========================================
' 5. FINALIZE MASTER TAB & SAVE
' ==========================================
SafeDeleteSheet wsHome.Parent, "TempRawBackup"

Dim fileTag As String
fileTag = "Alerted"
excelFileName = ecmID & "_" & AlertID & "_Combined_" & fileTag & "_Transaction.xlsx"

finalSavePath = saveFolderPath & slash & excelFileName
CloseIfAlreadyOpen finalSavePath
' "Pivot" only exists if lastRowCP > 1 above - SetSheetZoom85's own error
' handling silently skips it otherwise, same as any other missing name.
SetSheetZoom85 newWb, Array("Raw Transactions", "CP Selection", "DeDupe", "Pivot")

Application.DisplayAlerts = False
newWb.SaveAs fileName:=finalSavePath, FileFormat:=51
Application.DisplayAlerts = True

wsRealCD.Cells.Clear
newWb.Sheets("DeDupe").UsedRange.Copy Destination:=wsRealCD.Range("A1")

With wsRealCD.Cells
    .WrapText = False
    .EntireColumn.AutoFit
    .WrapText = True
    .EntireRow.AutoFit
    .VerticalAlignment = xlTop
End With

On Error Resume Next
Module3.RefreshRuleNameTag
On Error GoTo CancelHandler

SafeDeleteSheet wsHome.Parent, "TempConsolidatedScratch"

newWb.Sheets("Raw Transactions").Activate

Application.EnableCancelKey = xlInterrupt
Application.Calculation = origCalc
Application.EnableEvents = True
Application.ScreenUpdating = True

On Error Resume Next
ThisWorkbook.Sheets("ConsolidatedData").Protect Password:="p7ss"
ThisWorkbook.Sheets("Sheet1").Protect Password:="p7ss"
ThisWorkbook.Protect Password:="p7ss", Structure:=True, Windows:=False
On Error GoTo 0

' ==========================================
' 6. UPDATE SHARED MASTER TRACKER (SHEET2) - DEFERRED
' ==========================================
On Error Resume Next
Application.OnTime Now + TimeSerial(0, 0, 1), "PushTrxTracker_Deferred"
On Error GoTo 0

gExportLastResult = "OK"
MsgBox "Workflow Complete!" & vbCrLf & _
"Exported file inside the folder exactly to: " & vbCrLf & finalSavePath, vbInformation, "Success"
Exit Sub

CancelHandler:
Dim savedErrNum As Long, savedErrDesc As String
savedErrNum = Err.Number
savedErrDesc = Err.Description

Application.Calculation = origCalc
Application.EnableCancelKey = xlInterrupt
Application.ScreenUpdating = True
Application.DisplayAlerts = True

On Error Resume Next
If Not wsHome Is Nothing Then
    SafeDeleteSheet wsHome.Parent, "TempConsolidatedScratch"
    SafeDeleteSheet wsHome.Parent, "TempRawBackup"
    SafeDeleteSheet wsHome.Parent, "TempPivotScratch"
End If

ThisWorkbook.Sheets("ConsolidatedData").Protect Password:="p7ss"
ThisWorkbook.Sheets("Sheet1").Protect Password:="p7ss"
ThisWorkbook.Protect Password:="p7ss", Structure:=True, Windows:=False

Application.EnableEvents = True
On Error GoTo 0

If savedErrNum = 18 Then
    MsgBox "Process Safely Cancelled.", vbInformation, "Aborted"
ElseIf savedErrNum <> 0 Then
    MsgBox "An unexpected error occurred:" & vbCrLf & savedErrDesc, vbCritical, "Error " & savedErrNum
End If
End Sub

' ==========================================================
' SafeDeleteSheet - 100% immune to Subscript out of range (Err 9)
' and Workbook Protection / VeryHidden deletion crashes (Err 1004)
' ==========================================================
Private Sub SafeDeleteSheet(ByVal wb As Workbook, ByVal sheetName As String)
    On Error Resume Next
    If wb Is Nothing Then Exit Sub

    wb.Unprotect Password:="p7ss"
    wb.Unprotect

    Dim wsTarget As Worksheet
    Set wsTarget = Nothing
    Set wsTarget = wb.Sheets(sheetName)

    If Not wsTarget Is Nothing Then
        Application.DisplayAlerts = False
        wsTarget.Visible = xlSheetVisible
        wsTarget.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub

' Scans a sheet's UsedRange for date-shaped values and forces mm/dd/yyyy
Private Sub BulletproofDateFormat(ByVal ws As Worksheet)
On Error Resume Next

Dim rUsed As Range
Set rUsed = ws.UsedRange

If rUsed.Cells.count = 1 Then
    If IsDate(rUsed.Value) And Not IsEmpty(rUsed.Value) Then
        If Year(CDate(rUsed.Value)) > 1950 Then rUsed.NumberFormat = "mm/dd/yyyy"
    End If
    Exit Sub
End If

Dim arr As Variant, r As Long, c As Long
Dim hitRange As Range
arr = rUsed.Value

For r = 1 To UBound(arr, 1)
    For c = 1 To UBound(arr, 2)
        If IsDate(arr(r, c)) And Not IsEmpty(arr(r, c)) Then
            If Year(CDate(arr(r, c))) > 1950 Then
                If hitRange Is Nothing Then
                    Set hitRange = rUsed.Cells(r, c)
                Else
                    Set hitRange = Union(hitRange, rUsed.Cells(r, c))
                End If
            End If
        End If
    Next c
Next r

If Not hitRange Is Nothing Then hitRange.NumberFormat = "mm/dd/yyyy"
End Sub

' Highlights the Sum-column cell for every CR/DR row in the pivot's RowRange
Private Sub HighlightDrCrRows(ByVal pt As PivotTable, ByVal ws As Worksheet)
On Error Resume Next

Dim sumColIndex As Long, baseRow As Long
sumColIndex = pt.DataBodyRange.Columns(2).Column
baseRow = pt.RowRange.row

Dim rowArr As Variant, rIdx As Long, cellVal As String
Dim hitRange As Range
rowArr = pt.RowRange.Value

If IsArray(rowArr) Then
    For rIdx = 1 To UBound(rowArr, 1)
        cellVal = Trim(UCase(CStr(rowArr(rIdx, 1))))
        If cellVal = "CR" Or cellVal = "DR" Then
            If hitRange Is Nothing Then
                Set hitRange = ws.Cells(baseRow + rIdx - 1, sumColIndex)
            Else
                Set hitRange = Union(hitRange, ws.Cells(baseRow + rIdx - 1, sumColIndex))
            End If
        End If
    Next rIdx
Else
    cellVal = Trim(UCase(CStr(rowArr)))
    If cellVal = "CR" Or cellVal = "DR" Then
        Set hitRange = ws.Cells(baseRow, sumColIndex)
    End If
End If

If Not hitRange Is Nothing Then
    hitRange.Interior.Color = RGB(255, 199, 206)   ' Light Red Fill
    hitRange.Font.Color = RGB(156, 0, 6)           ' Dark Red Text
End If
End Sub

' ==========================================================
' BuildEnPivots - builds the 4 "Legacy" pivots on a fresh pivot sheet
' ==========================================================
Private Sub BuildEnPivots(ByVal wb As Workbook, ByVal dataSheet As String, _
    ByVal pivotSheet As String, ByVal insertAfter As String)
    Dim wsData As Worksheet, wsPv As Worksheet, wsPvScratch As Worksheet
    Dim dCol As Long, lastRow As Long, lastCol As Long, ddLast As Long
    Dim rngScn As Range, cacheScn As PivotCache, rngTmp As Range, cacheTmp As PivotCache
    Dim ptx As PivotTable, ptE As PivotTable, ptD As PivotTable

    On Error Resume Next
    Set wsData = wb.Sheets(dataSheet)
    On Error GoTo 0
    If wsData Is Nothing Then Exit Sub

    dCol = 0
    On Error Resume Next
    dCol = wsData.Rows(1).Find(What:="Transaction Date", LookAt:=xlPart).Column
    On Error GoTo 0

    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).row
    lastCol = wsData.Cells(1, wsData.Columns.count).End(xlToLeft).Column
    If lastRow <= 1 Then Exit Sub

    Set wsPv = wb.Sheets.Add(After:=wb.Sheets(insertAfter))
    wsPv.Name = pivotSheet

    ' PIVOT 1: SCENARIO
    Set rngScn = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))
    Set cacheScn = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngScn)
    Set ptx = cacheScn.CreatePivotTable(TableDestination:=wsPv.Range("A3"), TableName:="ScenarioPivot")
    On Error Resume Next
    With ptx
        .TableStyle2 = "PivotStyleLight16"
        With .PivotFields("Alert Information"): .Orientation = xlRowField: .Position = 1: End With
        With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 2: End With
        With .PivotFields("Counterparty"): .Orientation = xlRowField: .Position = 3: End With
        .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount", xlSum
        .PivotFields("Sum of Transaction Amount").NumberFormat = "$#,#00.00"
        .AddDataField .PivotFields("Transaction Amount"), "Count of Transaction Amount", xlCount
        .RowAxisLayout xlCompactRow
        .PivotFields("Count of Transaction Amount").NumberFormat = "0"
        .PivotFields("Alert Information").AutoSort xlDescending, "Sum of Transaction Amount"
        .PivotFields("Counterparty").AutoSort xlDescending, "Sum of Transaction Amount"
    End With
    On Error GoTo 0

    ' Build scratch copy for date-grouped pivots using Sheets.Add + UsedRange.Copy
    On Error Resume Next
    Set wsPvScratch = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.count))
    wsPvScratch.Visible = xlSheetVeryHidden
    If wsData.UsedRange.Cells.count > 0 Then
        wsData.UsedRange.Copy Destination:=wsPvScratch.Range("A1")
    End If
    On Error GoTo 0
    If wsPvScratch Is Nothing Then GoTo SkipDateGroupedPivots

    On Error Resume Next
    If dCol > 0 Then
        ddLast = wsPvScratch.Cells(wsPvScratch.Rows.count, dCol).End(xlUp).row
        If ddLast > 1 Then
            wsPvScratch.Range(wsPvScratch.Cells(2, dCol), wsPvScratch.Cells(ddLast, dCol)).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
            wsPvScratch.Range(wsPvScratch.Cells(2, dCol), wsPvScratch.Cells(ddLast, dCol)).NumberFormat = "m/d/yyyy"
        End If
    End If
    On Error GoTo 0

    lastRow = wsPvScratch.Cells(wsPvScratch.Rows.count, "A").End(xlUp).row
    lastCol = wsPvScratch.Cells(1, wsPvScratch.Columns.count).End(xlToLeft).Column
    If lastRow > 1 Then
        Set rngTmp = wsPvScratch.Range(wsPvScratch.Cells(1, 1), wsPvScratch.Cells(lastRow, lastCol))
        Set cacheTmp = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngTmp)

        ' PIVOT 2: TEMPORAL
        Set ptx = cacheTmp.CreatePivotTable(TableDestination:=wsPv.Range("F3"), TableName:="TemporalPivot")
        On Error Resume Next
        With ptx
            .TableStyle2 = "PivotStyleLight16"
            With .PivotFields("Transaction Date"): .Orientation = xlRowField: .Position = 1: End With
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount ", xlSum
            .PivotFields("Sum of Transaction Amount ").NumberFormat = "$#,#00.00"
            .AddDataField .PivotFields("Transaction Amount"), "Count of Transaction Amount ", xlCount
            .PivotFields("Count of Transaction Amount ").NumberFormat = "0"
        End With
        wsPv.Range("F4").Group Start:=True, End:=True, Periods:=Array(False, False, False, True, True, False, True)
        On Error Resume Next
        ptx.PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        On Error GoTo 0

        ' PIVOT 3: ENHANCED TEMPORAL
        ptx.TableRange2.Copy Destination:=wsPv.Range("K3")
        Set ptE = wsPv.Range("K3").PivotTable
        ptE.Name = "TemporalPivot_Enhanced"
        On Error Resume Next
        With ptE
            .PivotFields("Sum of Transaction Amount ").Orientation = xlHidden
            .PivotFields("Count of Transaction Amount ").Orientation = xlHidden
            With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 4: End With
            .AddDataField .PivotFields("Transaction Amount"), "No of Trx  ", xlCount
            .PivotFields("No of Trx  ").NumberFormat = "0"
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount  ", xlSum
            .PivotFields("Sum of Transaction Amount  ").NumberFormat = "$#,#00.00"
            .PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        End With
        On Error GoTo 0

        ' PIVOT 4: DR/CR INVERTED
        ptx.TableRange2.Copy Destination:=wsPv.Range("Q3")
        Set ptD = wsPv.Range("Q3").PivotTable
        ptD.Name = "DrCrTemporalPivot"
        On Error Resume Next
        With ptD
            .PivotFields("Sum of Transaction Amount ").Orientation = xlHidden
            .PivotFields("Count of Transaction Amount ").Orientation = xlHidden
            With .PivotFields("Dr Cr"): .Orientation = xlRowField: .Position = 1: End With
            .AddDataField .PivotFields("Transaction Amount"), "No of Trx   ", xlCount
            .PivotFields("No of Trx   ").NumberFormat = "0"
            .AddDataField .PivotFields("Transaction Amount"), "Sum of Transaction Amount   ", xlSum
            .PivotFields("Sum of Transaction Amount   ").NumberFormat = "$#,#00.00"
            .PivotFields("Transaction Date").NumberFormat = "mm/dd/yyyy"
        End With
        On Error GoTo 0

        On Error Resume Next
        BulletproofDateFormat wsPv
        HighlightDrCrRows ptE, wsPv
        HighlightDrCrRows ptD, wsPv
        On Error GoTo 0
    End If

SkipDateGroupedPivots:
    SafeDeleteSheet wb, "TempPivotScratch"
    wsPv.Columns("A:W").AutoFit
End Sub

' ==========================================================
' SetSheetZoom85 - forces each named sheet's saved window zoom to 85%,
' matching ConsolidatedData's OWN saved zoom in the live workbook (its
' sheetView zoomScale is 85, not 100 - confirmed directly in the file's
' XML). A sheet built via Sheets.Add inside a brand-new Workbooks.Add
' workbook defaults to 100% instead of inheriting whatever zoom the
' analyst's own window happened to be at - which is what the OLD
' Worksheet.Copy-based approach used to carry over silently, and is the
' most likely reason exported sheets used to look "different" even
' though the font itself was never actually changed by that rewrite.
'
' Zoom is a WINDOW property in VBA (Window.Zoom), not a range/cell
' property, so each sheet must be briefly activated to set it. This
' MUST run BEFORE SaveAs - Zoom set after a workbook is already written
' to disk never makes it into that saved file.
' ==========================================================
Private Sub SetSheetZoom85(ByVal wb As Workbook, ByVal sheetNames As Variant)
    On Error Resume Next
    Dim nm As Variant
    For Each nm In sheetNames
        wb.Sheets(CStr(nm)).Activate
        ActiveWindow.Zoom = 85
    Next nm
    On Error GoTo 0
End Sub

' ==========================================================
' CloseIfAlreadyOpen - safely closes an open target file prior to SaveAs
' ==========================================================
Private Sub CloseIfAlreadyOpen(ByVal targetPath As String)
    On Error Resume Next
    Dim targetName As String, wb As Workbook
    targetName = Mid$(targetPath, InStrRev(targetPath, Application.PathSeparator) + 1)
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            If StrComp(wb.Name, targetName, vbTextCompare) = 0 Then
                Application.DisplayAlerts = False
                wb.Close SaveChanges:=False
                Application.DisplayAlerts = True
                Exit For
            End If
        End If
    Next wb
    On Error GoTo 0
End Sub

' ==========================================================
' LastDataRow - finds the true last data row across all columns
' ==========================================================
Private Function LastDataRow(ByVal ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If c Is Nothing Then
        LastDataRow = 1
    Else
        LastDataRow = c.row
    End If
End Function

Private Function ArrCell(ByVal v As Variant, ByVal idx As Long) As Variant
    If IsArray(v) Then
        ArrCell = v(idx, 1)
    Else
        ArrCell = v
    End If
End Function

' ==========================================================
' FilterRowsFast - high-performance bulk row deletion via AutoFilter
' ==========================================================
Private Sub FilterRowsFast(ByVal ws As Worksheet, ByVal splitCol As Long, _
    ByVal wantVal As String, ByVal dateCol As Long, ByVal useWindow As Boolean, _
    ByVal winStart As Date, ByVal winEnd As Date)

    Dim lastRow As Long, lastCol As Long, helperCol As Long, r As Long
    Dim sVals As Variant, dVals As Variant, dv As Variant
    Dim keep As Boolean, anyDelete As Boolean
    Dim flags() As Variant
    Dim delRange As Range

    lastRow = LastDataRow(ws)
    If lastRow < 2 Then Exit Sub

    lastCol = 1
    On Error Resume Next
    lastCol = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    On Error GoTo 0
    helperCol = lastCol + 1

    If splitCol > 0 Then sVals = ws.Range(ws.Cells(2, splitCol), ws.Cells(lastRow, splitCol)).Value
    If dateCol > 0 Then dVals = ws.Range(ws.Cells(2, dateCol), ws.Cells(lastRow, dateCol)).Value

    ReDim flags(1 To lastRow - 1, 1 To 1)
    anyDelete = False
    For r = 1 To lastRow - 1
        keep = True
        If splitCol > 0 And Len(wantVal) > 0 Then
            If Trim(CStr(ArrCell(sVals, r))) <> wantVal Then keep = False
        End If
        If keep And useWindow Then
            dv = ArrCell(dVals, r)
            If IsDate(dv) Then
                If CDate(dv) < winStart Or CDate(dv) > winEnd Then keep = False
            Else
                keep = False
            End If
        End If
        If keep Then
            flags(r, 1) = "K"
        Else
            flags(r, 1) = "D"
            anyDelete = True
        End If
    Next r

    If Not anyDelete Then Exit Sub

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0

    ws.Cells(1, helperCol).Value = "_flag"
    ws.Range(ws.Cells(2, helperCol), ws.Cells(lastRow, helperCol)).Value = flags
    ws.Range(ws.Cells(1, helperCol), ws.Cells(lastRow, helperCol)).AutoFilter Field:=1, Criteria1:="D"

    On Error Resume Next
    Set delRange = ws.Range(ws.Cells(2, helperCol), ws.Cells(lastRow, helperCol)).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not delRange Is Nothing Then delRange.EntireRow.Delete

    On Error Resume Next
    ws.AutoFilterMode = False
    On Error GoTo 0
    ws.Columns(helperCol).Delete
End Sub

' ==========================================================
' CleanTransactionData - applies standardized cleaning and Counterparty formula
' ==========================================================
Private Sub CleanTransactionData(ByVal ws As Worksheet)
    Dim dH As Range, dLast As Long, dRng As Range
    Dim aH As Range, aLast As Long, aRng As Range
    Dim drC As Range, benC As Range, orgC As Range
    Dim drCol As Long, benCol As Long, orgCol As Long
    Dim lastR As Long, lastC As Long

    On Error Resume Next

    Set dH = ws.Rows(1).Find(What:="Transaction Date", LookIn:=xlValues, LookAt:=xlPart)
    If Not dH Is Nothing Then
        dLast = ws.Cells(ws.Rows.count, dH.Column).End(xlUp).row
        If dLast > 1 Then
            Set dRng = ws.Range(ws.Cells(2, dH.Column), ws.Cells(dLast, dH.Column))
            dRng.TextToColumns Destination:=dRng.Cells(1, 1), DataType:=xlDelimited, FieldInfo:=Array(Array(1, 3))
            dRng.NumberFormat = "m/d/yyyy"
        End If
    End If

    Set aH = ws.Rows(1).Find(What:="Transaction Amount", LookIn:=xlValues, LookAt:=xlPart)
    If Not aH Is Nothing Then
        aLast = ws.Cells(ws.Rows.count, aH.Column).End(xlUp).row
        If aLast > 1 Then
            Set aRng = ws.Range(ws.Cells(2, aH.Column), ws.Cells(aLast, aH.Column))
            aRng.Value = aRng.Value
            aRng.NumberFormat = "$#,##0.00"
        End If
    End If

    Set drC = ws.Rows(1).Find(What:="Dr Cr", LookAt:=xlPart)
    Set benC = ws.Rows(1).Find(What:="Beneficiary Name", LookAt:=xlPart)
    Set orgC = ws.Rows(1).Find(What:="Originator Name", LookAt:=xlPart)
    If Not drC Is Nothing And Not benC Is Nothing And Not orgC Is Nothing Then
        drCol = drC.Column: benCol = benC.Column: orgCol = orgC.Column
        lastR = ws.Cells(ws.Rows.count, "A").End(xlUp).row
        lastC = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column + 1
        If lastR > 1 Then
            ws.Cells(1, lastC).Value = "Counterparty"
            ws.Range(ws.Cells(2, lastC), ws.Cells(lastR, lastC)).FormulaR1C1 = _
                "=IF(RC" & drCol & "=""DR"", IF(RC" & benCol & "="""","""",RC" & benCol & "), IF(RC" & orgCol & "="""","""",RC" & orgCol & "))"
            ws.Cells(1, lastC - 1).Copy
            ws.Cells(1, lastC).PasteSpecial Paste:=xlPasteFormats
            ws.Range(ws.Cells(2, lastC - 1), ws.Cells(lastR, lastC - 1)).Copy
            ws.Range(ws.Cells(2, lastC), ws.Cells(lastR, lastC)).PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
        End If
    End If

    On Error GoTo 0
End Sub

