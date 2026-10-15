Attribute VB_Name = "modECMImport"
Option Explicit

' ==========================================================================
' IMPORT ECM INFO
'
' Takes the ECM_Alert_Info_<date>_<time>.csv that the "ECM Export" browser
' button saves to Downloads (one row per alert) and, for an alert:
'   1. fills Sheet1:
'        J9  = ECM Case           <- ECM ID
'        J10 = Alert ID           <- Velocity Alert ID
'        J13 = Customer Name      <- Customer Name
'        J14 = Customer Address   <- Address (street, city, state zip, country)
'        J7  = Program Name       <- ECM Focal Entity "ProgramDescription" (the program
'                                    manager) matched to Program_Library!A1:A97 - see
'                                    MatchProgram; ProgramName is tried if it does not match
'        L27 = ISO Code 2         <- country code of the address; Sheet1's own
'                                    change handler fills G27 / O27 and the risk
'                                    formulas follow (filled here if it does not)
'      Cells are written one at a time with events on, so Sheet1 applies its
'      usual auto-capitalisation exactly as if they had been typed;
'   2. creates the case folders (same rule as START);
'   3. copies the alert's transaction file from Downloads into
'      <Desktop>\<ECM>\Transaction Files;
'   4. runs "Export Trx File" in AUTO_EXPORT_MODE (EN Network), no picker;
'   5. samples the counterparties into J18:J23 / T18:T23 (modCPSample);
'   6. runs the OSDD search with Option 2 (Module2 auto mode: name prompts take
'      their defaults, Sigma "additional address" = No) - AUTO_OSDD;
'   7. generates both narratives with the Wise template - Escalation, then
'      Non-Escalation (Module3 quiet mode) - AUTO_NARRATIVES. Decision/Template
'      are cleared afterwards so the analyst still makes the call; J7 is kept;
'   8. merges every entity's OSDD PDFs through Power Automate - AUTO_MERGE:
'      MERGE_MODE "WATCH" leaves a request for the "Merge Watcher" flow (running
'      all day, no confirmation per case); "URL" starts the "Merge All OSDD" flow
'      by its run link in Backend_Settings!B5 (Power Automate asks to Continue).
'
' Failsafe - nothing already made for a case is made again:
'   - a Transaction Files folder that already has an Excel file is not copied into;
'   - if the case folder already has its EN "Combined Alerted & Non Alerted
'     Transactions" file, the export is not re-run; instead ConsolidatedData is
'     reloaded from that file's "Alerted Transaction" sheet (what the export
'     itself would put there), so the narrative still matches the case.
'   (The Export Trx File button itself is unchanged and always rebuilds.)
'
' When a queued case is ready an on-top "case ready" message appears above
' every window (NOTIFY_WHEN_READY); clicking OK leaves you in the tool on Sheet1.
'
' Two ways in:
'   AUTOMATIC  The workbook watches Downloads (started on open by Auto_Open)
'              and works as a queue, ONE CASE AT A TIME: whenever the form is
'              empty (after Reset) it loads the next waiting alert - steps 1-4,
'              no prompts - and then leaves it for you to work. Press Reset
'              when you are done and the next alert loads within ~10 seconds.
'              Waiting alerts come from every ECM_Alert_Info csv in Downloads
'              (moved to Downloads\ECM_Imported when picked up), oldest first.
'              Each csv keeps its own list of alerts already loaded
'              (<csv name>_loaded.txt), so every new export is a fresh queue
'              whatever is already on the Desktop. A csv whose alerts have all
'              been loaded is renamed Done_<name>.
'              Reset (ClearForm) calls ECM_AfterReset, which loads the next
'              alert straight away and restarts the watcher if Excel lost it.
'              ECM_Queue_Status shows what is waiting.
'              Self-healing: every entry point traps its own errors and logs
'              them; a check stuck for over 5 minutes is released; an alert
'              whose transaction file has not finished downloading is retried
'              for up to 2 minutes before it is reported; the watcher restarts
'              from Reset, from START and (with the ThisWorkbook snippet) when
'              the tool window is activated.
'              StopEcmWatcher / StartEcmWatcher pause and resume it.
'   MANUAL     START with an empty ECM ID, or Alt+F8 > Import_ECM_Info:
'              pick the csv and the alert.
' ==========================================================================

Private Const PWD As String = "p7ss"
Private Const ECM_IMPORT_VERSION As String = "2026-10-15e (merge requests without BOM)"
Private Const NOTIFY_WHEN_READY As Boolean = True ' on-top "case ready" message when a queued case loads
' On-top messages use the Windows MessageBox directly (VBA's MsgBox does not
' accept the "topmost" flag) - the same way the OSDD search prompts do, so the
' message shows above Edge even though Excel is in the background.
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpText As LongPtr, _
                                                         ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
Private Const MB_ICONWARNING As Long = &H30
Private Const MB_ICONINFO As Long = &H40
Private Const MB_ON_TOP As Long = &H1000 Or &H10000 Or &H40000   ' system modal + set foreground + topmost
Private Const MAX_LISTED As Long = 15
Private Const AUTO_EXPORT_MODE As String = "EN"   ' "EN" / "LEGACY" / "PIVOT" / "" = do not export
Private Const AUTO_OSDD As Boolean = True          ' run the OSDD search (Option 2) after sampling
Private Const AUTO_NARRATIVES As Boolean = True    ' generate the Escalation and Non-Escalation narratives
Private Const NARR_TEMPLATE As String = "Wise"     ' template (J6) used for both narratives
Private Const AUTO_MERGE As Boolean = True         ' start the "Merge All OSDD" Power Automate flow at the end
Private Const MERGE_URL_CELL As String = "B5"      ' Backend_Settings cell holding that flow's run link (URL mode)
Private Const MERGE_MODE As String = "WATCH"       ' "WATCH" = Merge Watcher flow, "URL" = start Merge All OSDD by link
Private Const WATCH_SECONDS As Long = 10          ' how often Downloads is checked
Private Const BUSY_TIMEOUT_MIN As Long = 45       ' a check running longer than this is treated as stuck (OSDD can take a while)
Private Const FILE_WAIT_SECONDS As Long = 120     ' how long to wait for a transaction file still downloading
Private Const IMPORTED_FOLDER As String = "ECM_Imported"
Private Const DONE_PREFIX As String = "Done_"
Private Const LOG_NAME As String = "ECM_import_log.txt"

' Loaded csv
Private mRows As Collection
Private mCEcm As Long, mCVel As Long, mCCust As Long, mCAddr As Long, mCCountry As Long, mCFile As Long
Private mCProgDesc As Long, mCProgName As Long
Private mProgramNote As String      ' last import's program line, for the "case ready" message
Private mProgMatchWord As String    ' set by MatchProgram when only a shared word matched
Private mProgramValue As String     ' library name matched at import - reused by the narratives

' Watcher state
Private mWatching As Boolean
Private mNextTick As Date
Private mBusy As Boolean
Private mBusySince As Date
Private mWaitEcm As String          ' alert waiting for its transaction file
Private mWaitSince As Date

' ==========================================================================
' MANUAL
' ==========================================================================
Public Sub Import_ECM_Info()
    ECM_EnsureWatcher
    Dim ws As Worksheet, csvPath As String, loadErr As String
    Dim pick As Long, ecmID As String, note As String, copied As String, result As String

    Set ws = ThisWorkbook.Sheets("Sheet1")

    csvPath = ChooseCsv()
    If csvPath = "" Then Exit Sub

    On Error GoTo Fail
    loadErr = LoadCsv(csvPath)
    If loadErr <> "" Then MsgBox loadErr, vbExclamation, "Import ECM Info": Exit Sub

    pick = ChooseAlert(mRows, mCEcm, mCVel, mCCust, Trim$(CStr(ws.Range("J9").Value)))
    If pick < 2 Then Exit Sub
    ecmID = Field(mRows(pick), mCEcm)

    If Trim$(CStr(ws.Range("J9").Value)) <> "" And Trim$(CStr(ws.Range("J9").Value)) <> ecmID Then
        If MsgBox("Sheet1 already holds ECM case " & ws.Range("J9").Value & "." & vbCrLf & vbCrLf & _
                  "Replace the case details with ECM case " & ecmID & "?" & vbCrLf & _
                  "(Tip: press Reset first to clear the rest of the form.)", _
                  vbYesNo + vbQuestion, "Import ECM Info") <> vbYes Then Exit Sub
    End If

    copied = ImportRow(ws, pick, csvPath, note)

    If copied <> "" And AUTO_EXPORT_MODE <> "" Then
        ' Straight on to the export - no prompt. Its own completion message ends the run.
        Application.StatusBar = "ECM case " & ecmID & " imported - running Export Trx File (" & _
                                ExportModeName(AUTO_EXPORT_MODE) & ")..."
        result = RunAutoExport(ecmID, False)
        Application.StatusBar = False
        If Left$(result, 2) = "OK" Then
            Dim cpNote As String
            cpNote = FillCounterparties()
            If AUTO_OSDD Then cpNote = cpNote & vbCrLf & RunAutoOsdd()
            If AUTO_NARRATIVES Then cpNote = cpNote & vbCrLf & RunAutoNarratives(ws)
            If AUTO_MERGE Then cpNote = cpNote & vbCrLf & RunAutoMerge(ws)
            Application.StatusBar = Replace(Split(cpNote & vbCrLf, vbCrLf)(0), vbTab, " ")
        End If
        If result <> "OK" And Left$(result, 2) = "OK" Then
            MsgBox note & vbCrLf & "EN Network files already exist for this case - " & Mid$(result, 6) & ".", _
                   vbInformation, "Import ECM Info"
        ElseIf result <> "OK" And result <> "" Then
            MsgBox note & vbCrLf & result, vbExclamation, "Import ECM Info"
        End If
    Else
        MsgBox note & vbCrLf & "Next: choose Decision and Template, then click Export Trx File.", _
               vbInformation, "Import ECM Info"
    End If
    Exit Sub

Fail:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number: errDesc = Err.Description
    On Error Resume Next
    ProtectSheet ws
    Application.EnableEvents = True
    On Error GoTo 0
    MsgBox "Import stopped with an error:" & vbCrLf & vbCrLf & "Err " & errNum & ": " & errDesc, _
           vbCritical, "Import ECM Info"
End Sub

' ==========================================================================
' AUTOMATIC - Downloads watcher, one case at a time
' ==========================================================================
' Every entry point below traps its own errors: the watcher must never put up
' a VBA error box, whatever state Excel is in.
Public Sub Auto_Open()
    On Error Resume Next
    StartEcmWatcher
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    StopEcmWatcher
End Sub

Public Sub StartEcmWatcher()
    On Error Resume Next
    MergeRequestFolder              ' so the Merge Watcher flow always finds its folder
    mWatching = True
    ScheduleTick
    Application.StatusBar = "ECM auto-import is on - the next waiting alert loads whenever the form is empty."
End Sub

Public Sub StopEcmWatcher()
    mWatching = False
    On Error Resume Next
    Application.OnTime mNextTick, MacroRef("ECM_Watch_Tick"), , False
    Application.StatusBar = False
End Sub

' Excel refuses OnTime (error 1004) while a Protected View window is active -
' e.g. an ECM file opened from Downloads. So: close Protected View windows of
' ECM export files first, then retry with the plain macro name, then log.
Private Sub ScheduleTick()
    On Error Resume Next
    mNextTick = Now + TimeSerial(0, 0, WATCH_SECONDS)
    Application.OnTime mNextTick, MacroRef("ECM_Watch_Tick")
    If Err.Number <> 0 Then
        Err.Clear
        CloseEcmProtectedViews
        Application.OnTime mNextTick, MacroRef("ECM_Watch_Tick")
    End If
    If Err.Number <> 0 Then
        Err.Clear
        Application.OnTime mNextTick, "ECM_Watch_Tick"
    End If
    If Err.Number <> 0 Then
        mWatching = False        ' so Reset / START / switching to the tool restart it
        WriteLog "-", "watcher timer could not be set (" & Err.Number & ": " & Err.Description & _
                      ") - it restarts on Reset, START or when you switch to the tool"
    End If
End Sub

' Closes Protected View windows showing ECM export files (ECM_Alert_Info csv,
' <ECM ID>_*.xlsx, anything opened from Downloads\ECM_Imported). Nothing else
' is touched. Returns how many were closed.
Private Function CloseEcmProtectedViews() As Long
    Dim i As Long, nm As String, pth As String
    On Error Resume Next
    For i = Application.ProtectedViewWindows.Count To 1 Step -1
        nm = "": pth = ""
        nm = Application.ProtectedViewWindows(i).SourceName
        pth = Application.ProtectedViewWindows(i).SourcePath
        If nm Like "*ECM_Alert_Info_*.csv" Or nm Like "#*_ALERT*.xls*" Or _
           InStr(1, pth, "\" & IMPORTED_FOLDER, vbTextCompare) > 0 Then
            Application.ProtectedViewWindows(i).Close
            CloseEcmProtectedViews = CloseEcmProtectedViews + 1
        End If
    Next i
End Function

' OnTime needs the workbook name when several workbooks are open.
Private Function MacroRef(ByVal procName As String) As String
    MacroRef = "'" & ThisWorkbook.Name & "'!" & procName
End Function

Public Sub ECM_Watch_Tick()
    On Error Resume Next          ' a failed check must never stop the watcher
    If Not mWatching Then Exit Sub
    CloseEcmProtectedViews
    RunCheck
    If mWatching Then ScheduleTick
End Sub

' One queue check, never two at once. A check marked busy for longer than
' BUSY_TIMEOUT_MIN is assumed stuck (e.g. interrupted) and released.
Private Sub RunCheck()
    On Error GoTo Logged
    If mBusy Then
        If Now - mBusySince < TimeSerial(0, BUSY_TIMEOUT_MIN, 0) Then Exit Sub
        WriteLog "-", "previous check looked stuck - released"
    End If
    mBusy = True
    mBusySince = Now
    CheckQueue
    mBusy = False
    Exit Sub
Logged:
    mBusy = False
    WriteLog "-", "check failed (" & Err.Number & ": " & Err.Description & ") - will retry"
    On Error Resume Next
    Application.EnableEvents = True
    ProtectSheet ThisWorkbook.Sheets("Sheet1")
End Sub

' Restarts the watcher if Excel lost it (e.g. after a VBA reset). Safe to call often.
Public Sub ECM_EnsureWatcher()
    On Error Resume Next
    If Not mWatching Then StartEcmWatcher
End Sub

' Picks up new exports, then - only if the form is empty - loads the next
' waiting alert. A case on the form (J9 filled) is never replaced.
Private Sub CheckQueue()
    Dim ws As Worksheet, csv As String, idx As Long, waiting As Long
    Set ws = ThisWorkbook.Sheets("Sheet1")

    PickUpNewExports
    waiting = NextWaiting(csv, idx)

    If Trim$(CStr(ws.Range("J9").Value)) <> "" Then
        If waiting > 0 Then
            Application.StatusBar = waiting & " ECM alert(s) waiting - press Reset when you finish case " & _
                                    ws.Range("J9").Value & " to load the next one."
        End If
        Exit Sub
    End If
    If waiting = 0 Then Exit Sub

    LoadQueued ws, csv, idx, waiting - 1
End Sub

' Moves finished ECM_Alert_Info csvs from Downloads into ECM_Imported.
' A file still being written (changed in the last 3 seconds) is left for the next check.
Private Sub PickUpNewExports()
    Dim fso As Object, folder As String, destFolder As String, f As String, dest As String
    Dim found As New Collection, item As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = DownloadsFolder()
    destFolder = folder & "\" & IMPORTED_FOLDER

    f = Dir$(folder & "\ECM_Alert_Info_*.csv")
    Do While f <> ""
        If Now - FileDateTime(folder & "\" & f) > TimeSerial(0, 0, 3) Then found.Add f
        f = Dir$()
    Loop
    If found.Count = 0 Then Exit Sub

    If Not fso.FolderExists(destFolder) Then fso.CreateFolder destFolder
    For Each item In found
        dest = destFolder & "\" & item
        If fso.FileExists(dest) Then dest = destFolder & "\" & fso.GetBaseName(item) & "_" & Format$(Now, "hhnnss") & ".csv"
        fso.MoveFile folder & "\" & item, dest
    Next item
End Sub

' Finds the first waiting alert (oldest csv first, then file order): one not in
' that csv's loaded list. Sets csv / idx to it and returns how many alerts are
' waiting in total. Csvs with nothing left are renamed Done_.
Private Function NextWaiting(ByRef csv As String, ByRef idx As Long) As Long
    Dim fso As Object, folder As String, f As String, files As New Collection
    Dim i As Long, j As Long, r As Long, pending As Long, total As Long
    Dim paths() As String, times() As Date, tmpS As String, tmpT As Date, loadedList As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = DownloadsFolder() & "\" & IMPORTED_FOLDER
    csv = "": idx = 0
    If Not fso.FolderExists(folder) Then Exit Function

    f = Dir$(folder & "\ECM_Alert_Info_*.csv")
    Do While f <> ""
        files.Add folder & "\" & f
        f = Dir$()
    Loop
    If files.Count = 0 Then Exit Function

    ' oldest first
    ReDim paths(1 To files.Count): ReDim times(1 To files.Count)
    For i = 1 To files.Count
        paths(i) = files(i): times(i) = FileDateTime(files(i))
    Next i
    For i = 1 To files.Count - 1
        For j = i + 1 To files.Count
            If times(j) < times(i) Then
                tmpS = paths(i): paths(i) = paths(j): paths(j) = tmpS
                tmpT = times(i): times(i) = times(j): times(j) = tmpT
            End If
        Next j
    Next i

    For i = 1 To files.Count
        If LoadCsv(paths(i)) = "" Then
            pending = 0
            loadedList = ReadLoaded(paths(i))
            For r = 2 To mRows.Count
                If InStr(loadedList, "|" & Field(mRows(r), mCEcm) & "|") = 0 Then
                    pending = pending + 1
                    If csv = "" Then csv = paths(i): idx = r
                End If
            Next r
            If pending = 0 Then
                fso.MoveFile paths(i), folder & "\" & DONE_PREFIX & fso.GetFileName(paths(i))
                If fso.FileExists(LoadedPath(paths(i))) Then _
                    fso.MoveFile LoadedPath(paths(i)), folder & "\" & DONE_PREFIX & fso.GetFileName(LoadedPath(paths(i)))
            End If
            total = total + pending
        End If
    Next i
    NextWaiting = total
End Function

' Steps 1-4 for one queued alert, with no prompts. Leaves the case on Sheet1
' for the analyst. Only a failure shows a message.
Private Sub LoadQueued(ByVal ws As Worksheet, ByVal csv As String, ByVal idx As Long, ByVal stillWaiting As Long)
    Dim ecmID As String, note As String, copied As String, result As String, cpNote As String
    Dim osddNote As String, narrNote As String, mergeNote As String

    If LoadCsv(csv) <> "" Then Exit Sub
    ecmID = Field(mRows(idx), mCEcm)

    ' Self-healing: if its transaction file is not there yet (still downloading),
    ' leave the form alone and try again on the next check, for up to FILE_WAIT_SECONDS.
    If Not TransactionFileReady(csv, Field(mRows(idx), mCFile), ecmID) Then
        If mWaitEcm <> ecmID Then mWaitEcm = ecmID: mWaitSince = Now
        If Now - mWaitSince < TimeSerial(0, 0, FILE_WAIT_SECONDS) Then
            Application.StatusBar = "ECM auto-import: waiting for ECM " & ecmID & "'s transaction file to finish downloading..."
            Exit Sub
        End If
    End If
    mWaitEcm = ""
    MarkLoaded csv, ecmID          ' before loading, so a failing alert can never block the queue

    ThisWorkbook.Activate
    ws.Activate
    Application.StatusBar = "ECM auto-import: loading ECM " & ecmID & "..."

    On Error GoTo Failed
    copied = ImportRow(ws, idx, csv, note)
    WriteLog ecmID, "step program: " & IIf(mProgramNote = "", "(no program line)", mProgramNote)
    If copied = "" Then
        result = "transaction file not found in Downloads"
    ElseIf AUTO_EXPORT_MODE = "" Then
        result = "OK"
    Else
        result = RunAutoExport(ecmID, True)
        If result = "" Then result = "export did not finish (see its message)"
    End If
    WriteLog ecmID, "step export: " & result
    If Left$(result, 2) = "OK" Then
        cpNote = FillCounterparties()
        WriteLog ecmID, "step counterparties: " & FirstLine(cpNote)
        If AUTO_OSDD Then
            Application.StatusBar = "ECM " & ecmID & ": running OSDD search (Option 2)..."
            WriteLog ecmID, "step OSDD: started"
            osddNote = RunAutoOsdd()
            cpNote = cpNote & vbCrLf & osddNote
            WriteLog ecmID, "step OSDD: " & FirstLine(osddNote)
        End If
        If AUTO_NARRATIVES Then
            Application.StatusBar = "ECM " & ecmID & ": generating narratives..."
            WriteLog ecmID, "step narratives: started"
            narrNote = RunAutoNarratives(ws)
            cpNote = cpNote & vbCrLf & narrNote
            WriteLog ecmID, "step narratives: " & Replace(Replace(narrNote, vbCrLf, " "), "   ", " ")
        End If
        If AUTO_MERGE Then
            mergeNote = RunAutoMerge(ws)
            cpNote = cpNote & vbCrLf & mergeNote
            WriteLog ecmID, "step merge: " & Replace(mergeNote, vbCrLf, " ")
        End If
    Else
        WriteLog ecmID, "steps after the export skipped (export not OK)"
    End If
    GoTo Report

Failed:
    result = "error " & Err.Number & ": " & Err.Description
    ProtectSheet ws
    Application.EnableEvents = True
    Resume Report

Report:
    WriteLog ecmID, result
    PrepareToolWindow ws
    If Left$(result, 2) = "OK" Then
        Application.StatusBar = "ECM " & ecmID & IIf(result = "OK", " loaded and exported. ", " loaded (EN files already existed, not recreated). ") & _
            IIf(stillWaiting > 0, stillWaiting & " more waiting - press Reset when done to load the next.", "No more alerts waiting.")
        If NOTIFY_WHEN_READY Then
            TopMessage "ECM case " & ecmID & " is ready in the tool." & vbCrLf & _
                       IIf(stillWaiting > 0, stillWaiting & " more waiting after this one.", "No more alerts waiting.") & vbCrLf & vbCrLf & _
                       IIf(mProgramNote <> "", mProgramNote & vbCrLf & vbCrLf, "") & _
                       IIf(cpNote <> "", cpNote & vbCrLf, "") & _
                       "Click OK to work it.", "ECM case ready", MB_ICONINFO
        End If
    Else
        Application.StatusBar = False
        TopMessage "ECM auto-import - case " & ecmID & ":" & vbCrLf & result & vbCrLf & vbCrLf & note, _
                   "ECM auto-import", MB_ICONWARNING
    End If
End Sub

' Message box that sits above every window. Falls back to a normal MsgBox.
Private Sub TopMessage(ByVal text As String, ByVal title As String, ByVal icon As Long)
    On Error GoTo Plain
    MessageBoxW 0, StrPtr(text), StrPtr(title), icon Or MB_ON_TOP
    Exit Sub
Plain:
    On Error Resume Next
    MsgBox text, IIf(icon = MB_ICONWARNING, vbExclamation, vbInformation), title
End Sub

' Makes the tool the active workbook on Sheet1 and un-minimises it, so that
' when the on-top message is closed you land straight in the case.
Private Sub PrepareToolWindow(ByVal ws As Worksheet)
    Const WIN_MINIMIZED As Long = -4140, WIN_NORMAL As Long = -4143   ' xlMinimized, xlNormal
    On Error Resume Next
    ThisWorkbook.Activate
    ThisWorkbook.Windows(1).Activate
    If ThisWorkbook.Windows(1).WindowState = WIN_MINIMIZED Then ThisWorkbook.Windows(1).WindowState = WIN_NORMAL
    If Application.WindowState = WIN_MINIMIZED Then Application.WindowState = WIN_NORMAL
    ws.Activate
End Sub

' Per-csv list of alerts the queue has already loaded: "<csv name>_loaded.txt".
Private Function LoadedPath(ByVal csv As String) As String
    LoadedPath = Left$(csv, Len(csv) - 4) & "_loaded.txt"
End Function

' The loaded list as "|id|id|...|" for quick InStr checks.
Private Function ReadLoaded(ByVal csv As String) As String
    Dim fso As Object, ts As Object, t As String
    ReadLoaded = "|"
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(LoadedPath(csv)) Then Exit Function
    Set ts = fso.OpenTextFile(LoadedPath(csv), 1)
    t = ts.ReadAll
    ts.Close
    ReadLoaded = "|" & Replace(Replace(t, vbCr, ""), vbLf, "|")
    If Right$(ReadLoaded, 1) <> "|" Then ReadLoaded = ReadLoaded & "|"
End Function

Private Sub MarkLoaded(ByVal csv As String, ByVal ecmID As String)
    Dim ts As Object
    On Error Resume Next
    Set ts = CreateObject("Scripting.FileSystemObject").OpenTextFile(LoadedPath(csv), 8, True)
    ts.WriteLine ecmID
    ts.Close
End Sub

' Called by Reset (ClearForm): load the next waiting alert now instead of at
' the next 10-second check, and restart the watcher if Excel lost it.
Public Sub ECM_AfterReset()
    On Error Resume Next
    ECM_EnsureWatcher
    RunCheck
End Sub

' Alt+F8 > ECM_Queue_Status: what the queue sees right now.
Public Sub ECM_Queue_Status()
    Dim csv As String, idx As Long, waiting As Long, msg As String
    On Error Resume Next
    PickUpNewExports
    waiting = NextWaiting(csv, idx)
    msg = "Version: " & ECM_IMPORT_VERSION & vbCrLf & _
          "Watcher: " & IIf(mWatching, "ON", "OFF (run StartEcmWatcher, or press Reset)") & vbCrLf & _
          "Protected View windows open: " & Application.ProtectedViewWindows.Count & vbCrLf & _
          "Case on the form (J9): " & IIf(Trim$(CStr(ThisWorkbook.Sheets("Sheet1").Range("J9").Value)) = "", "none", _
                                          ThisWorkbook.Sheets("Sheet1").Range("J9").Value) & vbCrLf & _
          "Alerts waiting: " & waiting & vbCrLf
    If waiting > 0 Then
        If LoadCsv(csv) = "" Then msg = msg & "Next: ECM " & Field(mRows(idx), mCEcm) & "  (from " & Mid$(csv, InStrRev(csv, "\") + 1) & ")" & vbCrLf
    End If
    msg = msg & vbCrLf & "Exports are read from: " & DownloadsFolder() & "\" & IMPORTED_FOLDER
    MsgBox msg, vbInformation, "ECM queue"
End Sub

Private Sub WriteLog(ByVal ecmID As String, ByVal result As String)
    Dim fso As Object, ts As Object, path As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = DownloadsFolder() & "\" & IMPORTED_FOLDER & "\" & LOG_NAME
    Set ts = fso.OpenTextFile(path, 8, True)     ' 8 = append, create if missing
    ts.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & ecmID & vbTab & result
    ts.Close
End Sub

' ==========================================================================
' SHARED
' ==========================================================================

' Reads the csv into mRows and finds the columns. Returns "" or an error text.
Private Function LoadCsv(ByVal csvPath As String) As String
    Set mRows = ParseCsv(ReadUtf8(csvPath))
    If mRows.Count < 2 Then LoadCsv = "That file has no alert rows:" & vbCrLf & csvPath: Exit Function
    mCEcm = ColIdx(mRows(1), "ECM ID")
    mCVel = ColIdx(mRows(1), "Velocity Alert ID")
    mCCust = ColIdx(mRows(1), "Customer Name")
    mCAddr = ColIdx(mRows(1), "Address")
    mCCountry = ColIdx(mRows(1), "Country")
    mCFile = ColIdx(mRows(1), "Transaction File")
    mCProgDesc = ColIdx(mRows(1), "Program Description")
    mCProgName = ColIdx(mRows(1), "Program Name")
    If mCEcm < 0 Then LoadCsv = "This does not look like an ECM export file (no 'ECM ID' column):" & vbCrLf & csvPath
End Function

' Steps 1-3 for row idx of mRows: fill Sheet1, create the folders, copy the
' transaction file. Returns the copied file's path ("" if not found); note
' gets a readable summary.
Private Function ImportRow(ByVal ws As Worksheet, ByVal idx As Long, ByVal csvPath As String, _
                           ByRef note As String) As String
    Dim rec As Variant, ecmID As String, velID As String, custName As String
    Dim addr As String, iso2 As String, caseFolder As String, copied As String

    rec = mRows(idx)
    ecmID = Field(rec, mCEcm)
    velID = Field(rec, mCVel)
    custName = Field(rec, mCCust)
    addr = Field(rec, mCAddr)
    iso2 = CountryCode(Field(rec, mCCountry))

    Application.EnableEvents = True
    ' Clear what the previous case put there, so nothing carries over.
    If Trim$(CStr(ws.Range("J9").Value)) <> ecmID Then
        Dim r As Long
        For r = 18 To 23
            PutValue ws, "J" & r, ""
            PutValue ws, "T" & r, ""
        Next r
        PutValue ws, "J7", ""
    End If
    PutValue ws, "J13", ""
    PutValue ws, "J14", ""
    PutValue ws, "L27", ""
    PutValue ws, "G27", ""
    PutValue ws, "O27", ""

    PutValue ws, "J9", ecmID
    PutValue ws, "J10", velID
    If custName <> "" Then PutValue ws, "J13", custName
    If addr <> "" Then PutValue ws, "J14", addr
    mProgramNote = FillProgram(ws, Field(rec, mCProgDesc), Field(rec, mCProgName))
    If iso2 <> "" Then
        PutValue ws, "L27", iso2
        ' Sheet1's handler normally fills G27/O27 from L27; fill them here if it did not.
        If Trim$(CStr(ws.Range("G27").Value)) = "" Then FillCountryFromIso2 ws, iso2
    End If
    ProtectSheet ws

    note = "Filled Sheet1 for ECM case " & ecmID & ":" & vbCrLf & _
           "   ECM Case (J9), Alert ID (J10)" & _
           IIf(custName <> "", ", Customer Name (J13)", "") & _
           IIf(addr <> "", ", Customer Address (J14)", "") & vbCrLf & _
           IIf(mProgramNote <> "", "   " & mProgramNote & vbCrLf, "")
    If iso2 <> "" Then
        If Trim$(CStr(ws.Range("G27").Value)) <> "" Then
            note = note & "   Country (G27): " & ws.Range("G27").Value & vbCrLf
        Else
            note = note & "   Country code " & iso2 & " is not in the country table - pick the country in G27." & vbCrLf
        End If
    End If

    caseFolder = MakeCaseFolders(ecmID)
    copied = FirstExcelIn(caseFolder & "\Transaction Files")
    If copied <> "" Then
        note = note & vbCrLf & "Transaction file already in the case folder (not copied again):" & vbCrLf & _
               "   " & copied & vbCrLf
        ImportRow = copied
        Exit Function
    End If
    copied = CopyTransactionFile(csvPath, Field(rec, mCFile), ecmID, caseFolder & "\Transaction Files")
    If copied <> "" Then
        note = note & vbCrLf & "Transaction file copied to:" & vbCrLf & "   " & copied & vbCrLf
    Else
        note = note & vbCrLf & "Transaction file NOT found in Downloads - copy it into:" & vbCrLf & _
               "   " & caseFolder & "\Transaction Files" & vbCrLf
    End If
    ImportRow = copied
End Function

' Step 4. quiet = no completion message and the export's files are closed
' (queue mode). Returns "OK" when the export finished, "OK - <what>" when the
' EN files already existed and were reused, "" or an error text otherwise.
Private Function RunAutoExport(ByVal ecmID As String, ByVal quiet As Boolean) As String
    Dim existing As String
    If UCase$(AUTO_EXPORT_MODE) = "EN" Then
        existing = ExistingEnFile(ecmID)
        If existing <> "" Then
            RunAutoExport = ReuseEnFile(existing)
            Exit Function
        End If
    End If

    gExportLastResult = ""
    gExportQuiet = quiet
    gExportModeOverride = AUTO_EXPORT_MODE
    Consolidated_AML_Workflow
    gExportModeOverride = ""
    gExportQuiet = False
    RunAutoExport = gExportLastResult
End Function

' The case's EN "Combined Alerted & Non Alerted Transactions" file, if one was made before.
Private Function ExistingEnFile(ByVal ecmID As String) As String
    Dim folder As String, f As String
    folder = CaseFolderPath(ecmID)
    f = Dir$(folder & "\" & ecmID & "_*Combined Alerted & Non Alerted Transactions*.xls*")
    Do While f <> ""
        If Left$(f, 2) <> "~$" Then ExistingEnFile = folder & "\" & f: Exit Function
        f = Dir$()
    Loop
End Function

' Instead of rebuilding: fill ConsolidatedData from the existing file's
' "Alerted Transaction" sheet, exactly as the EN export does, and refresh the
' rule-name tag. The file itself is opened read-only and left untouched.
Private Function ReuseEnFile(ByVal path As String) As String
    Dim wb As Workbook, cd As Worksheet, wasOpen As Boolean, prevUpd As Boolean
    prevUpd = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then wasOpen = True: Exit For
    Next wb
    If Not wasOpen Then Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=False)

    Set cd = ThisWorkbook.Sheets("ConsolidatedData")
    ThisWorkbook.Unprotect Password:=PWD
    cd.Unprotect Password:=PWD
    cd.Cells.Clear
    wb.Sheets("Alerted Transaction").UsedRange.Copy Destination:=cd.Range("A1")
    If Not wasOpen Then wb.Close SaveChanges:=False

    On Error Resume Next
    RefreshRuleNameTag
    cd.Protect Password:=PWD
    ThisWorkbook.Protect Password:=PWD, Structure:=True
    ThisWorkbook.Sheets("Sheet1").Activate
    On Error GoTo 0
    Application.ScreenUpdating = prevUpd
    ReuseEnFile = "OK - existing EN file reused, not recreated"
    Exit Function

Fail:
    ReuseEnFile = "could not reuse existing EN file (" & Err.Description & "): " & path
    On Error Resume Next
    If Not wasOpen And Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Protect Password:=PWD, Structure:=True
    Application.ScreenUpdating = prevUpd
End Function

' Step 6: the tool's OSDD search (Module2) in auto mode - Option 2, defaults for
' every name prompt, Sigma = No, no start/success messages. Returns a short
' summary for the "case ready" message. Never raises an error.
Private Function RunAutoOsdd() As String
    Dim lines() As String, i As Long, out As String
    On Error Resume Next
    ThisWorkbook.Activate
    ThisWorkbook.Sheets("Sheet1").Activate
    gOsddSummary = ""
    gOsddAuto = True
    SearchAndSavePDF_Direct
    gOsddAuto = False
    If Err.Number <> 0 Then
        RunAutoOsdd = "OSDD search stopped with an error (" & Err.Number & ": " & Err.Description & ") - run it from the OSDD Search button."
        Err.Clear
        Exit Function
    End If
    If gOsddSummary = "" Then
        RunAutoOsdd = "OSDD search did not finish - check its message, or run it from the OSDD Search button."
        Exit Function
    End If
    lines = Split(Replace(gOsddSummary, vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        If Trim$(lines(i)) <> "" Then out = out & "   " & Trim$(lines(i)) & vbCrLf
        If i >= 7 Then out = out & "   ...": Exit For
    Next i
    RunAutoOsdd = "OSDD search (Option 2):" & vbCrLf & out
End Function

' Alt+F8 > ECM_Generate_Narratives: both narratives for the case now on Sheet1 -
' the same step the queue runs, on its own, to check it.
Public Sub ECM_Generate_Narratives()
    Dim ws As Worksheet, msg As String
    Set ws = ThisWorkbook.Sheets("Sheet1")
    If Trim$(CStr(ws.Range("J9").Value)) = "" Then MsgBox "No case on Sheet1 (J9 is empty).", vbExclamation, "Narratives": Exit Sub
    msg = RunAutoNarratives(ws)
    WriteLog CStr(ws.Range("J9").Value), "manual narratives: " & Replace(msg, vbCrLf, " ")
    MsgBox msg, vbInformation, "Narratives"
End Sub

Private Function FirstLine(ByVal s As String) As String
    FirstLine = Split(Replace(s, vbCr, "") & vbLf, vbLf)(0)
End Function

' Step 7: both narratives with the Wise template. For each decision: J5 = decision,
' J6 = template, then J7 is written again - Sheet1 clears J7 whenever J5 changes -
' recalculate, and run Generate Narrative (Module3 ExportToWord) quietly. The files
' are the tool's usual "..._Escalation Narrative.docx" and "..._Alert Write-Up.docx"
' in Desktop\<ECM>. Afterwards J5/J6 are cleared (the decision stays with the
' analyst) and J7 is restored. Never raises an error.
Private Function RunAutoNarratives(ByVal ws As Worksheet) As String
    Dim prog As String, d As Variant, out As String, f As String
    On Error GoTo Fail
    prog = CStr(ws.Range("J7").Value)
    ' J7 can be emptied after the import (Sheet1 clears it whenever J5 changes),
    ' so fall back to the program matched at import.
    If Trim$(prog) = "" Then prog = mProgramValue
    For Each d In Array("Escalation", "Non-Escalation")
        Application.EnableEvents = True
        PutValue ws, "J5", CStr(d)
        PutValue ws, "J6", NARR_TEMPLATE
        If prog <> "" Then PutValue ws, "J7", prog
        ProtectSheet ws
        Application.CalculateFull
        gNarrLastFile = ""
        gNarrQuiet = True
        ExportToWord
        gNarrQuiet = False
        f = gNarrLastFile
        out = out & "   " & d & ": " & IIf(f <> "", Mid$(f, InStrRev(f, "\") + 1), "not saved - see its message") & vbCrLf
    Next d
    RestoreDecision ws, prog
    RunAutoNarratives = "Narratives (" & NARR_TEMPLATE & " template" & IIf(prog = "", ", NO program name - check [Program Description]", "") & "):" & vbCrLf & out
    Exit Function

Fail:
    RunAutoNarratives = "Narratives stopped with an error (" & Err.Number & ": " & Err.Description & ") - " & _
                        "use Generate Narrative." & vbCrLf & out
    gNarrQuiet = False
    RestoreDecision ws, prog
End Function

' After the narratives: Decision and Template cleared, Program Name put back.
Private Sub RestoreDecision(ByVal ws As Worksheet, ByVal prog As String)
    On Error Resume Next
    PutValue ws, "J5", ""
    PutValue ws, "J6", ""
    If prog <> "" Then PutValue ws, "J7", prog
    ProtectSheet ws
    Application.CalculateFull
End Sub

' Step 8: start the "Merge All OSDD" Power Automate flow once for the case. Its
' inputs: ECMCase, AlertID, OsddFolder (Desktop\<ECM>\OSDD Searches) and Entities,
' "Customer|<customer name>;CP1|CP1_<name>;CP2|CP2_<name>..." - for each entry the
' flow merges the OSDD PDFs whose names contain "_Customer_" / "_CP1_" ... into
' <case folder>\<ECM>_<Alert>_<second part>_MergedOSDD.pdf (same names as the
' PDF Merge button). The flow runs on its own; Excel does not wait for it.
Private Function RunAutoMerge(ByVal ws As Worksheet) As String
    Dim url As String, ents As String, r As Long, nm As String, ecm As String, json As String
    On Error GoTo Fail
    ecm = Trim$(CStr(ws.Range("J9").Value))
    nm = SafeEntity(CStr(ws.Range("J13").Value))
    If nm <> "" Then ents = "Customer|" & nm
    For r = 18 To 23
        nm = SafeEntity(CStr(ws.Cells(r, "J").Value))
        If nm <> "" Then ents = ents & IIf(ents = "", "", ";") & "CP" & (r - 17) & "|CP" & (r - 17) & "_" & nm
    Next r
    If ents = "" Then RunAutoMerge = "PDF merge not started - no customer or counterparty names on Sheet1.": Exit Function

    If UCase$(MERGE_MODE) = "WATCH" Then
        RunAutoMerge = QueueMergeRequest(ecm, Trim$(CStr(ws.Range("J10").Value)), CaseFolderPath(ecm) & "\OSDD Searches", ents)
        Exit Function
    End If
    url = Trim$(CStr(ThisWorkbook.Sheets("Backend_Settings").Range(MERGE_URL_CELL).Value))
    If url = "" Then
        RunAutoMerge = "PDF merge not started - put the ""Merge All OSDD"" flow's run link in Backend_Settings!" & MERGE_URL_CELL & "."
        Exit Function
    End If

    json = "{""ECMCase"":""" & JsonText(ecm) & """,""AlertID"":""" & JsonText(CStr(ws.Range("J10").Value)) & _
           """,""OsddFolder"":""" & JsonText(CaseFolderPath(ecm) & "\OSDD Searches") & _
           """,""Entities"":""" & JsonText(ents) & """}"
    url = url & IIf(InStr(url, "?") > 0, "&", "?") & "inputArguments=" & UrlEncodeAll(json)
    CreateObject("Shell.Application").ShellExecute url
    RunAutoMerge = "PDF merge started in Power Automate for: " & Replace(Replace(ents, ";", ", "), "|", " -> ")
    Exit Function
Fail:
    RunAutoMerge = "PDF merge not started (" & Err.Description & ") - use the PDF Merge button."
End Function

' MERGE_MODE "WATCH": hand the merge to the "Merge Watcher" flow, which runs all
' day in Power Automate (started once from its console - no confirmation). One
' request file per entity that has OSDD PDFs:
'   Downloads\ECM_Imported\MergeRequests\<ECM>_<Customer|CP1..CP6>.txt
' with three lines - the OSDD Searches folder, the file filter (*_CP1_*.pdf) and
' the merged PDF's full path - so the flow only has to read, merge and delete.
' Each file is written as .tmp and then renamed, so the flow never reads a
' half-written one.
Private Function QueueMergeRequest(ByVal ecm As String, ByVal alertId As String, _
                                   ByVal osddFolder As String, ByVal ents As String) As String
    Dim fso As Object, dirPath As String, e As Variant, pre As String, outName As String
    Dim filt As String, merged As String, queued As String, skipped As String
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    dirPath = MergeRequestFolder()
    For Each e In Split(ents, ";")
        pre = Split(e, "|")(0): outName = Split(e, "|")(1)
        filt = "*_" & pre & "_*.pdf"
        If Dir$(osddFolder & "\" & filt) = "" Then
            skipped = skipped & IIf(skipped = "", "", ", ") & pre
        Else
            merged = fso.GetParentFolderName(osddFolder) & "\" & ecm & "_" & alertId & "_" & outName & "_MergedOSDD.pdf"
            WriteRequest dirPath, ecm & "_" & pre, osddFolder & vbCrLf & filt & vbCrLf & merged & vbCrLf
            queued = queued & IIf(queued = "", "", ", ") & outName
        End If
    Next e
    If queued = "" Then
        QueueMergeRequest = "PDF merge: no OSDD PDFs found in " & osddFolder & " - nothing to merge."
    Else
        QueueMergeRequest = "PDF merge queued for the Merge Watcher flow: " & queued & _
                            IIf(skipped <> "", " (no PDFs for " & skipped & ")", "")
    End If
    Exit Function
Fail:
    QueueMergeRequest = "PDF merge request could not be written (" & Err.Description & ") - use the PDF Merge button."
End Function

' UTF-8 WITHOUT the byte-order mark, so line 1 (the folder path) has no hidden
' character in front of it when Power Automate reads the file.
Private Sub WriteRequest(ByVal dirPath As String, ByVal baseName As String, ByVal body As String)
    Dim fso As Object, st As Object, bin As Object, tmp As String, fin As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmp = dirPath & "\" & baseName & ".tmp"
    fin = dirPath & "\" & baseName & ".txt"
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText body
    st.Position = 0
    st.Type = 1                       ' switch to bytes and skip the 3-byte BOM
    st.Position = 3
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1: bin.Open
    st.CopyTo bin
    bin.SaveToFile tmp, 2
    bin.Close
    st.Close
    If fso.FileExists(fin) Then fso.DeleteFile fin, True
    fso.MoveFile tmp, fin
End Sub

' Downloads\ECM_Imported\MergeRequests, created if missing.
Private Function MergeRequestFolder() As String
    Dim fso As Object, p As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    p = DownloadsFolder() & "\" & IMPORTED_FOLDER
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    p = p & "\MergeRequests"
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    MergeRequestFolder = p
End Function

' Name usable in a file name and in the Entities list.
Private Function SafeEntity(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|", ";", vbCr, vbLf, vbTab)
        s = Replace(s, ch, " ")
    Next ch
    SafeEntity = Application.WorksheetFunction.Trim(s)
End Function

Private Function JsonText(ByVal s As String) As String
    JsonText = Replace(Replace(Trim$(s), "\", "\\"), """", "\""")
End Function

' Percent-encodes everything except letters and digits (UTF-8).
Private Function UrlEncodeAll(ByVal s As String) As String
    Dim i As Long, code As Long, out As String
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1)) And 65535
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then
            out = out & ChrW(code)
        ElseIf code < 128 Then
            out = out & "%" & Right$("0" & Hex(code), 2)
        ElseIf code < 2048 Then
            out = out & "%" & Hex(192 + code \ 64) & "%" & Hex(128 + (code And 63))
        Else
            out = out & "%" & Hex(224 + code \ 4096) & "%" & Hex(128 + ((code \ 64) And 63)) & "%" & Hex(128 + (code And 63))
        End If
    Next i
    UrlEncodeAll = out
End Function

' Closes an open workbook by full path, without saving (it was just saved).
Public Sub CloseSavedExport(ByVal fullPath As String)
    Dim wb As Workbook
    If fullPath = "" Then Exit Sub
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=False
            Exit Sub
        End If
    Next wb
End Sub

Private Function DownloadsFolder() As String
    DownloadsFolder = Environ$("USERPROFILE") & "\Downloads"
End Function

' --------------------------------------------------------------------------
' Newest ECM_Alert_Info_*.csv in Downloads (with a Yes/No to use it), or a
' file picker. Returns "" if cancelled.
' --------------------------------------------------------------------------
Private Function ChooseCsv() As String
    Dim folder As String, newest As String, imported As String, ans As VbMsgBoxResult
    Dim picked As Variant

    folder = DownloadsFolder()
    newest = NewestCsv(folder)
    imported = NewestCsv(folder & "\" & IMPORTED_FOLDER)
    If imported <> "" Then
        If newest = "" Then
            newest = imported
        ElseIf FileDateTime(imported) > FileDateTime(newest) Then
            newest = imported
        End If
    End If

    If newest <> "" Then
        ans = MsgBox("Use the latest ECM export file?" & vbCrLf & vbCrLf & _
                     Mid$(newest, InStrRev(newest, "\") + 1) & vbCrLf & _
                     "saved " & Format$(FileDateTime(newest), "dd mmm yyyy hh:nn") & vbCrLf & vbCrLf & _
                     "Yes = use it    No = choose another file", _
                     vbYesNoCancel + vbQuestion, "Import ECM Info")
        If ans = vbYes Then ChooseCsv = newest: Exit Function
        If ans = vbCancel Then Exit Function
    End If

    On Error Resume Next
    ChDrive Left$(folder, 1)
    ChDir folder
    On Error GoTo 0
    picked = Application.GetOpenFilename("ECM export (*.csv),*.csv", , "Choose the ECM_Alert_Info file")
    If VarType(picked) = vbBoolean Then Exit Function
    ChooseCsv = CStr(picked)
End Function

Private Function NewestCsv(ByVal folder As String) As String
    Dim f As String, best As String, bestT As Date, t As Date
    On Error Resume Next
    f = Dir$(folder & "\*ECM_Alert_Info_*.csv")       ' includes Done_ files
    Do While f <> ""
        t = FileDateTime(folder & "\" & f)
        If best = "" Or t > bestT Then best = folder & "\" & f: bestT = t
        f = Dir$()
    Loop
    NewestCsv = best
End Function

' --------------------------------------------------------------------------
' Pick one alert row. One row: taken as is. Several: numbered list, the row
' matching the ECM ID already in J9 is the default. Returns the row index in
' rows (2..n), or 0 if cancelled.
' --------------------------------------------------------------------------
Private Function ChooseAlert(ByVal rows As Collection, ByVal cEcm As Long, ByVal cVel As Long, _
                             ByVal cCust As Long, ByVal currentEcm As String) As Long
    Dim i As Long, listText As String, def As String, ans As String

    If rows.Count = 2 Then ChooseAlert = 2: Exit Function

    def = "1"
    For i = 2 To rows.Count
        If currentEcm <> "" And Field(rows(i), cEcm) = currentEcm Then def = CStr(i - 1)
        If i - 1 <= MAX_LISTED Then
            listText = listText & (i - 1) & ")  " & Field(rows(i), cEcm) & "   " & Field(rows(i), cVel) & _
                       "   " & Left$(Field(rows(i), cCust), 28) & vbCrLf
        End If
    Next i
    If rows.Count - 1 > MAX_LISTED Then
        listText = listText & "... and " & (rows.Count - 1 - MAX_LISTED) & " more (type the ECM ID)" & vbCrLf
    End If

    Do
        ans = Trim$(InputBox("Which alert? Type its number or ECM ID." & vbCrLf & vbCrLf & listText, _
                             "Import ECM Info", def))
        If ans = "" Then Exit Function
        If IsNumeric(ans) And Len(ans) <= 3 Then
            If CLng(ans) >= 1 And CLng(ans) <= rows.Count - 1 Then
                ChooseAlert = CLng(ans) + 1
                Exit Function
            End If
        End If
        For i = 2 To rows.Count
            If Field(rows(i), cEcm) = ans Then ChooseAlert = i: Exit Function
        Next i
        MsgBox "No alert '" & ans & "' in this file.", vbExclamation, "Import ECM Info"
    Loop
End Function

' --------------------------------------------------------------------------
' Sheet1 helpers. Sheet1's change handler re-protects the sheet and the
' workbook after every edit, so unprotect before each write.
' --------------------------------------------------------------------------
Private Sub PutValue(ByVal ws As Worksheet, ByVal cellAddr As String, ByVal v As String)
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=PWD
    ws.Unprotect Password:=PWD
    On Error GoTo 0
    ws.Range(cellAddr).Value = v
End Sub

Private Sub ProtectSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=PWD
    ThisWorkbook.Protect Password:=PWD, Structure:=True
    On Error GoTo 0
End Sub

' Country table on Sheet1: O37:O290 = name, P = ISO2, Q = ISO3.
Private Sub FillCountryFromIso2(ByVal ws As Worksheet, ByVal iso2 As String)
    Dim m As Variant
    m = Application.Match(iso2, ws.Range("P37:P290"), 0)
    If IsError(m) Then Exit Sub
    PutValue ws, "G27", CStr(ws.Range("O37:O290").Cells(CLng(m), 1).Value)
    PutValue ws, "O27", CStr(ws.Range("Q37:Q290").Cells(CLng(m), 1).Value)
End Sub

Private Function ExportModeName(ByVal mode As String) As String
    Select Case UCase$(mode)
        Case "EN": ExportModeName = "EN Network"
        Case "PIVOT": ExportModeName = "Pivot Analysis"
        Case Else: ExportModeName = "Legacy"
    End Select
End Function

' ---- Program Name (J7) ----------------------------------------------------
' Writes the Program_Library name matching ECM's ProgramDescription (or, failing
' that, ProgramName) into J7, exactly as the library spells it so J7's dropdown
' accepts it. Leaves J7 alone when nothing matches. Returns a one-line note.
Private Function FillProgram(ByVal ws As Worksheet, ByVal progDesc As String, ByVal progName As String) As String
    Dim hit As String, src As String
    On Error GoTo Fail
    mProgramValue = ""
    If progDesc = "" And progName = "" Then
        If mCProgDesc < 0 And mCProgName < 0 Then
            FillProgram = "Program: this export file has no program columns (exported before the Focal Entity update) - pick Program Name (J7) yourself."
        Else
            FillProgram = "Program: ECM Focal Entity info had no ProgramDescription / ProgramName - pick Program Name (J7) yourself."
        End If
        Exit Function
    End If
    mProgMatchWord = ""
    hit = MatchProgram(progDesc): src = progDesc
    If hit = "" Then mProgMatchWord = "": hit = MatchProgram(progName): src = progName
    If hit = "" Then
        FillProgram = "Program: '" & IIf(progDesc <> "", progDesc, progName) & _
                      "' is not in Program_Library - pick Program Name (J7) yourself."
        Exit Function
    End If
    PutValue ws, "J7", hit
    mProgramValue = hit
    FillProgram = "Program Name (J7): " & Trim$(hit) & IIf(NormName(src) <> NormName(hit), "  (ECM: " & src & ")", "") & _
                  IIf(mProgMatchWord <> "", " - matched on the word '" & mProgMatchWord & "', please check", "")
    Exit Function
Fail:
    FillProgram = "Program: could not fill J7 (" & Err.Description & ")."
End Function

' Best Program_Library!A1:A97 entry for an ECM program value, or "".
'   1. same name, ignoring case, spaces, punctuation and legal endings (Inc, Ltd, ...);
'   2. same as one of the entry's alternative names - the parts in brackets or
'      split by "/" (e.g. "Pomelo (Lemonade/LemFi)" also answers to LemFi);
'   3. otherwise the LONGEST entry (or alternative) that the value starts with, or
'      that starts with the value, as whole words - "Airwallex Hong Kong Limited"
'      picks "Airwallex Hong Kong Ltd" over "Airwallex";
'   4. otherwise any entry sharing a distinctive word (not a generic word such as
'      "global" or "payments", see GenericWord): most shared words wins, then the
'      entry with the fewest other words. mProgMatchWord then names the word.
Private Function MatchProgram(ByVal v As String) As String
    Dim lib As Range, c As Range, alts As Variant, a As Variant, nv As String, na As String
    Dim bestLen As Long, best As String
    nv = NormName(v)
    If nv = "" Then Exit Function
    On Error Resume Next
    Set lib = ThisWorkbook.Sheets("Program_Library").Range("A1:A97")
    On Error GoTo 0
    If lib Is Nothing Then Exit Function

    For Each c In lib.Cells                                   ' 1. exact
        If Trim$(CStr(c.Value)) <> "" Then
            If NormName(CStr(c.Value)) = nv Then MatchProgram = CStr(c.Value): Exit Function
        End If
    Next c
    For Each c In lib.Cells                                   ' 2. alternative names
        If Trim$(CStr(c.Value)) <> "" Then
            alts = AltNames(CStr(c.Value))
            For Each a In alts
                If NormName(CStr(a)) = nv And NormName(CStr(a)) <> "" Then MatchProgram = CStr(c.Value): Exit Function
            Next a
        End If
    Next c
    For Each c In lib.Cells                                   ' 3. longest whole-word prefix
        If Trim$(CStr(c.Value)) <> "" Then
            alts = AltNames(CStr(c.Value))
            For Each a In alts
                na = NormName(CStr(a))
                If Len(na) >= 3 Then
                    If (nv & " ") Like (na & " *") Or (na & " ") Like (nv & " *") Then
                        If Len(na) > bestLen Then bestLen = Len(na): best = CStr(c.Value)
                    End If
                End If
            Next a
        End If
    Next c
    If best <> "" Then MatchProgram = best: Exit Function

    Dim vw As Variant, lw As Variant, nShared As Long, extra As Long, bestNShared As Long, bestExtra As Long, word As String, bestWord As String
    bestExtra = 9999
    For Each c In lib.Cells                                   ' 4. shared distinctive word
        If Trim$(CStr(c.Value)) <> "" Then
            nShared = 0: extra = 0: word = ""
            For Each lw In Split(NormName(Replace(Replace(Replace(CStr(c.Value), "(", " "), ")", " "), "/", " ")), " ")
                If Not GenericWord(CStr(lw)) Then
                    extra = extra + 1
                    For Each vw In Split(nv, " ")
                        If CStr(vw) = CStr(lw) Then nShared = nShared + 1: extra = extra - 1: word = CStr(lw): Exit For
                    Next vw
                End If
            Next lw
            If nShared > 0 Then
                If nShared > bestNShared Or (nShared = bestNShared And extra < bestExtra) Then
                    bestNShared = nShared: bestExtra = extra: best = CStr(c.Value): bestWord = word
                End If
            End If
        End If
    Next c
    If best <> "" Then mProgMatchWord = bestWord
    MatchProgram = best
End Function

' Words too common in program names to identify one (also anything under 3 letters).
Private Function GenericWord(ByVal w As String) As Boolean
    If Len(w) < 3 Then GenericWord = True: Exit Function
    Select Case w
        Case "global", "money", "pay", "payment", "payments", "financial", "finance", "technology", "technologies", _
             "international", "bank", "banking", "transfer", "transfers", "remit", "remittance", "exchange", _
             "credit", "union", "card", "cards", "cash", "services", "service", "solutions", "solution", _
             "securities", "markets", "market", "cooperative", "platform", "digital", "network", "systems", _
             "group", "holdings", "company", "fintech", "trading", "capital", "usa", "europe", "evolution", "origination"
            GenericWord = True
    End Select
End Function

' The full name plus the parts in brackets / split by "/".
Private Function AltNames(ByVal full As String) As Variant
    Dim parts As String, p As Long, q As Long, outside As String, inside As String
    p = InStr(full, "(")
    If p > 0 Then
        q = InStr(p, full, ")")
        If q = 0 Then q = Len(full) + 1
        outside = Left$(full, p - 1) & " " & Mid$(full, q + 1)
        inside = Mid$(full, p + 1, q - p - 1)
    Else
        outside = full
    End If
    parts = full & "|" & Replace(outside, "/", "|") & "|" & Replace(inside, "/", "|")
    AltNames = Split(parts, "|")
End Function

' Lower case, letters and digits only (single spaces), legal endings removed.
Private Function NormName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, w As Variant, words As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[a-z0-9]" Then out = out & ch Else out = out & " "
    Next i
    For Each w In Split(Application.WorksheetFunction.Trim(out), " ")
        Select Case w
            Case "inc", "ltd", "llc", "corp", "corporation", "co", "pte", "limited", "plc", "gmbh", "sa", "ag", "bv", "lp", "llp", "the"
            Case Else
                If w <> "" Then words = words & " " & w
        End Select
    Next w
    NormName = Trim$(words)
End Function

' "US - United States" -> "US"; a bare 2-letter code is returned as is.
Private Function CountryCode(ByVal v As String) As String
    Dim p As Long
    v = Trim$(v)
    p = InStr(v, " - ")
    If p > 0 Then v = Trim$(Left$(v, p - 1))
    If Len(v) = 2 Then CountryCode = UCase$(v)
End Function

' --------------------------------------------------------------------------
' Case folders - same location rule as START (Start_Button_Create_Folders).
' --------------------------------------------------------------------------
Private Function CaseFolderPath(ByVal ecmID As String) As String
    Dim fso As Object, userProfile As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    userProfile = Environ$("USERPROFILE")
    If fso.FolderExists(userProfile & "\OneDrive - Community Federal Savings Bank\Desktop") Then
        CaseFolderPath = userProfile & "\OneDrive - Community Federal Savings Bank\Desktop\" & ecmID
    Else
        CaseFolderPath = userProfile & "\Desktop\" & ecmID
    End If
End Function

Private Function MakeCaseFolders(ByVal ecmID As String) As String
    Dim fso As Object, mainFolder As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    mainFolder = CaseFolderPath(ecmID)
    If Not fso.FolderExists(mainFolder) Then fso.CreateFolder mainFolder
    If Not fso.FolderExists(mainFolder & "\Transaction Files") Then fso.CreateFolder mainFolder & "\Transaction Files"
    If Not fso.FolderExists(mainFolder & "\Pivot") Then fso.CreateFolder mainFolder & "\Pivot"
    MakeCaseFolders = mainFolder
End Function

' First Excel file in a folder (ignores Excel's ~$ lock files), or "".
Private Function FirstExcelIn(ByVal folder As String) As String
    Dim f As String
    f = Dir$(folder & "\*.xls*")
    Do While f <> ""
        If Left$(f, 2) <> "~$" Then FirstExcelIn = folder & "\" & f: Exit Function
        f = Dir$()
    Loop
End Function

' The alert's transaction file is in its case folder already, or in Downloads.
Private Function TransactionFileReady(ByVal csvPath As String, ByVal trxFile As String, ByVal ecmID As String) As Boolean
    Dim fso As Object, folder As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If FirstExcelIn(CaseFolderPath(ecmID) & "\Transaction Files") <> "" Then TransactionFileReady = True: Exit Function
    If trxFile <> "" Then
        folder = Left$(csvPath, InStrRev(csvPath, "\") - 1)
        If fso.FileExists(folder & "\" & trxFile) Or fso.FileExists(DownloadsFolder() & "\" & trxFile) Then
            TransactionFileReady = True
            Exit Function
        End If
    End If
    TransactionFileReady = (Dir$(DownloadsFolder() & "\" & ecmID & "_*.xls*") <> "")
End Function

' Copies the alert's transaction file (named in the CSV, else the newest
' "<ECM ID>_*.xls*") into destFolder, looking next to the CSV and in
' Downloads. Returns the new path, or "".
Private Function CopyTransactionFile(ByVal csvPath As String, ByVal trxFile As String, _
                                     ByVal ecmID As String, ByVal destFolder As String) As String
    Dim fso As Object, folders(1 To 2) As String, src As String, f As String
    Dim bestT As Date, t As Date, k As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    folders(1) = Left$(csvPath, InStrRev(csvPath, "\") - 1)
    folders(2) = DownloadsFolder()

    For k = 1 To 2
        If trxFile <> "" And src = "" Then
            If fso.FileExists(folders(k) & "\" & trxFile) Then src = folders(k) & "\" & trxFile
        End If
    Next k
    If src = "" Then
        For k = 1 To 2
            f = Dir$(folders(k) & "\" & ecmID & "_*.xls*")
            Do While f <> ""
                t = FileDateTime(folders(k) & "\" & f)
                If src = "" Or t > bestT Then src = folders(k) & "\" & f: bestT = t
                f = Dir$()
            Loop
        Next k
    End If
    If src = "" Then Exit Function

    fso.CopyFile src, destFolder & "\", False     ' never overwrite
    CopyTransactionFile = destFolder & "\" & fso.GetFileName(src)
End Function

' --------------------------------------------------------------------------
' CSV reading
' --------------------------------------------------------------------------
Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object, s As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2            ' text
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    s = st.ReadText
    st.Close
    If Left$(s, 1) = ChrW$(65279) Then s = Mid$(s, 2)
    ReadUtf8 = s
End Function

' RFC 4180 CSV: quoted fields, "" inside quotes, commas and line breaks
' inside quotes. Returns a Collection of 0-based string arrays.
Private Function ParseCsv(ByVal txt As String) As Collection
    Dim result As New Collection, fields As Collection
    Dim i As Long, n As Long, ch As String, cur As String, inQ As Boolean

    Set fields = New Collection
    n = Len(txt)
    For i = 1 To n
        ch = Mid$(txt, i, 1)
        If inQ Then
            If ch = """" Then
                If i < n And Mid$(txt, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQ = True
                Case ","
                    fields.Add cur: cur = ""
                Case vbCr
                    ' ignored; vbLf ends the row
                Case vbLf
                    fields.Add cur: cur = ""
                    result.Add ToArray(fields)
                    Set fields = New Collection
                Case Else
                    cur = cur & ch
            End Select
        End If
    Next i
    If cur <> "" Or fields.Count > 0 Then
        fields.Add cur
        result.Add ToArray(fields)
    End If
    Set ParseCsv = result
End Function

Private Function ToArray(ByVal c As Collection) As Variant
    Dim a() As String, i As Long
    ReDim a(0 To c.Count - 1)
    For i = 1 To c.Count
        a(i - 1) = c(i)
    Next i
    ToArray = a
End Function

Private Function ColIdx(ByVal hdr As Variant, ByVal name As String) As Long
    Dim i As Long
    ColIdx = -1
    For i = LBound(hdr) To UBound(hdr)
        If StrComp(Trim$(hdr(i)), name, vbTextCompare) = 0 Then ColIdx = i: Exit Function
    Next i
End Function

' Field i of a row, trimmed; "" when the column is missing. Removes the
' leading apostrophe the export adds to values starting with = + - @.
Private Function Field(ByVal rec As Variant, ByVal i As Long) As String
    Dim v As String
    If i < 0 Then Exit Function
    If i > UBound(rec) Then Exit Function
    v = Trim$(rec(i))
    If Len(v) > 1 And Left$(v, 1) = "'" And InStr("=+-@", Mid$(v, 2, 1)) > 0 Then v = Mid$(v, 2)
    Field = v
End Function
