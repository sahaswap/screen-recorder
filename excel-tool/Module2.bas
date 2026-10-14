Attribute VB_Name = "Module2"
'==================================================================
' OSINT Automation Tool
'
' Runs a batch of Google OSINT searches for a customer and their
' counterparties, saves each result page as a PDF straight into
' the ECM case folder, and writes an audit trail back to the
' workbook and the shared SharePoint tracker.
'
' Headless Edge does the heavy lifting. A small worker pool lets
' a couple of searches run in parallel so a typical 20-30 search
' job finishes in a few minutes instead of the better part of ten.
'
' v3.5 adds a search-mode prompt right at the start of the run:
'   Fast      - headless, 2 workers, modest jitter. Quickest.
'   Optimised - headless, 2 workers, much wider jitter. Same
'               fingerprint as Fast, far less bursty request
'               pattern - trades run time for fewer CAPTCHA hits.
'   Visible   - real, visible Edge windows, no headless
'               fingerprint at all. Slowest; keep hands off mouse/
'               keyboard while it runs.
'
' To run: assign "SearchAndSavePDF_Direct" to your macro button.
'==================================================================

Option Explicit

' Auto OSDD (set by modECMImport after a queued case is loaded): Option 2 (Fast)
' without the picker, every name prompt takes its default (as if OK was
' pressed), Sigma "additional address?" = No, no "Starting searches" or
' "Success" message (the summary goes to gOsddSummary instead), and the case
' folder is not opened. CAPTCHA / blocked-search prompts still appear.
Public gOsddAuto As Boolean
Public gOsddSummary As String

' Direct Win32 imports - we reach past the VBA wrappers for things
' like foreground windows and modal dialogs because the wrappers
' don't behave well on locked-down / VDI hosts.
#If VBA7 Then
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As LongPtr
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
(ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function AttachThreadInput Lib "user32" _
(ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _
ByVal hwnd As LongPtr, ByVal lpText As LongPtr, _
ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, _
ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Declare PtrSafe Function SendMessageGetText Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, _
ByVal lParam As String) As LongPtr
Private Declare PtrSafe Function SendMessageLen Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, _
ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
(ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
#Else
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" _
(ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function AttachThreadInput Lib "user32" _
(ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function MessageBoxW Lib "user32" ( _
ByVal hwnd As Long, ByVal lpText As Long, _
ByVal lpCaption As Long, ByVal uType As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWndParent As Long, ByVal hWndChildAfter As Long, _
ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function SendMessageGetText Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
ByVal lParam As String) As Long
Private Declare Function SendMessageLen Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
(ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
#End If

' Win32 MessageBox flags used by TopMostMsgBox.
Private Const MB_OK          As Long = &H0
Private Const MB_YESNO       As Long = &H4
Private Const MB_ICONERROR   As Long = &H10
Private Const MB_ICONQUESTION As Long = &H20
Private Const MB_ICONWARNING As Long = &H30
Private Const MB_ICONINFO    As Long = &H40
Private Const MB_TOPMOST     As Long = &H40000
Private Const MB_SETFOREGROUND As Long = &H10000
Private Const MB_SYSTEMMODAL As Long = &H1000
Private Const IDOK     As Long = 1
Private Const IDYES    As Long = 6
Private Const IDNO     As Long = 7

' Window-message constants for reading the Save dialog filename
' field back (true paste verification).
Private Const WM_GETTEXT As Long = &HD
Private Const WM_GETTEXTLENGTH As Long = &HE

' ---- Tunables ----
' Anything above MIN_PDF_SIZE_BYTES is treated as a real PDF.
' Anything below CAPTCHA_SIZE_HINT is suspiciously small and flagged
' for a retry, since Google's interstitial/CAPTCHA pages tend to
' render tiny compared to a normal 100-result SERP.
Private Const TOOL_VERSION As String = "3.5"
Private Const MIN_PDF_SIZE_BYTES As Long = 5120
Private Const CAPTCHA_SIZE_HINT As Long = 80000

' Size below which we open the human-solve windows.
'
' NOTE (from real logs): on this environment real CAPTCHA pages come
' in at ~45-73 KB, while valid result pages are ~148 KB+. So there's
' a clean gap and 80 KB separates them well - a LOW value like 30 KB
' was wrong (it missed the real 45-73 KB CAPTCHAs). Kept equal to
' CAPTCHA_SIZE_HINT so the solve fires on any sub-80 KB page, which
' matched observed real CAPTCHAs with no false positives in testing.
' Lower it only if you start seeing the solve open on pages that were
' actually valid results (check the "HUMAN-SOLVE armed: PDF NNNNB"
' line against what the window showed).
Private Const CAPTCHA_HARD_SIZE_HINT As Long = 80000

' ---- Content-based CAPTCHA confirmation (v4.1) ----
' Size alone can't tell a 73 KB CAPTCHA page from a 73 KB small-but-
' valid results page. When a PDF comes back under CAPTCHA_SIZE_HINT,
' we now do a quick headless --dump-dom of the same query and READ
' the page: real CAPTCHA pages carry CAPTCHA markers ("unusual
' traffic", "/sorry/", "recaptcha"...), valid pages carry results
' markers ("result-stats", id="search"...). Only a content-confirmed
' CAPTCHA is retried / flagged / opens the solve window; a confirmed
' results page is accepted even if small. Safety-first: if the DOM
' can't be read (timeout/error), we DEFAULT to treating it as a
' CAPTCHA, so a real block is never mistaken for a clean result.
' Flip to False to fall back to pure size-based behaviour.
'
' DISABLED (Aug 2026): the --dump-dom probe launches an EXTRA headless
' Edge process per small PDF, and every Edge launch from Excel trips the
' VDI's "blocked by your administrator" ASR popup. Turning the probe off
' removes those extra launches. Trade-off: without content confirmation
' we fall back to pure size (any sub-CAPTCHA_SIZE_HINT PDF is treated as
' a CAPTCHA), so small-but-valid results pages can be flagged / trigger
' the human-solve window again. Tune CAPTCHA_SIZE_HINT if that gets noisy.
Private Const CAPTCHA_CONTENT_DETECT_ENABLED As Boolean = False
Private Const CAPTCHA_DETECT_TIMEOUT_SEC As Single = 12

Private USE_HEADLESS As Boolean
Private Const TIMING_PROFILE As String = "FAST"         ' FAST | NORMAL | SLOW

' ---- Search mode (v3.5) ----
' Chosen once per run via PromptForSearchMode, right where the old
' non-English Yes/No prompt used to sit. Replaces that binary prompt
' entirely - Visible mode covers the "I need to see/translate the
' page" need that prompt existed for, so there's no separate
' question for it any more.
'   smFast      - headless, 2 workers, modest jitter. Quickest,
'                 same basic CAPTCHA exposure as v3.4.
'   smOptimised - headless, 2 workers, much wider jitter. Same
'                 fingerprint as Fast, far less bursty request
'                 pattern - trades run time for fewer CAPTCHA hits.
'   smVisible   - routes every task through the existing
'                 PrintSearchToPDF_Interactive engine (real, visible
'                 Edge windows, serial). No headless fingerprint at
'                 all. Slowest; needs hands off mouse/keyboard like
'                 the old non-English fallback did.
Private Enum OsintSearchMode
    smFast = 0
    smOptimised = 1
    smVisible = 2
End Enum
Private m_searchMode As OsintSearchMode

' Parallel dispatch. Two workers is the sweet spot on a 2-vCPU VDI -
' a third worker just fights the others for CPU and pushes tasks
' past the timeout.
Private Const MAX_PARALLEL As Long = 2
Private Const LAUNCH_STAGGER_SEC As Single = 0.3        ' breathing room between launches
Private Const POLL_INTERVAL_SEC As Single = 0.3         ' how often we check for a finished PDF
' Lowered from 75/90 (Aug 2026): logs show every real search finishes
' in 6-9s, so a task still empty at 35s is a hung CAPTCHA page, not a
' slow-but-valid one. Catching it sooner requeues / escalates it to the
' human-solve far quicker instead of burning 75-90s per hang.
' Lowered 35->20 / 45->30 after the spawn-reduction flags tightened render
' times (max render ~16s, edgeStart mostly <=9s). A launch with no file by
' 20s is now genuinely hung, so we catch the "NO FILE EVER" hangs 15s
' sooner instead of waiting out the old 35s.
Private Const TASK_TIMEOUT_SEC As Long = 20             ' kill a task that is clearly stuck
Private Const RETRY_TIMEOUT_SEC As Long = 30            ' give retries a little more rope
Private Const RETRY_COOLDOWN_SEC As Single = 5          ' unused since v2.5.10 (inline retries)

' CAPTCHA-specific retry tuning. A CAPTCHA/interstitial page is
' almost always much smaller than a real 100-result SERP, so a task
' whose PDF lands under CAPTCHA_SIZE_HINT is NEVER accepted under
' its expected filename. It either gets requeued with a growing
' backoff (so we don't immediately re-trip the same block), or -
' once every attempt is burned - gets flagged for a human instead
' of quietly being counted as a clean result.
Private Const CAPTCHA_MAX_ATTEMPTS_MAIN As Long = 8     ' total tries in the main pass
Private Const CAPTCHA_MAX_ATTEMPTS_RESCUE As Long = 3   ' extra tries in the rescue pass
Private Const CAPTCHA_BACKOFF_BASE_SEC As Single = 5    ' first retry waits ~5s
Private Const CAPTCHA_BACKOFF_MAX_SEC As Single = 45    ' backoff never grows past this
Private Const CAPTCHA_FLAG_SUFFIX As String = "_CAPTCHA_UNRESOLVED"

' ---- Global CAPTCHA circuit-breaker (v3.8) ----
' The per-task backoff above only delays the ONE task that got
' blocked - the pool immediately fills that worker with a different
' task and keeps firing at full rate. So the moment Google starts
' rate-limiting us, we respond by continuing to hammer it, which
' deepens the block and turns a brief throttle into a batch-wide
' failure. That is the single worst behaviour left in the dispatcher.
'
' This adds a pool-wide brake: once THRESHOLD blocks land inside
' WINDOW seconds, ALL launching stops for PAUSE_SEC so the rate
' limit can decay, then resumes. In-flight tasks are left alone.
' Costs wall-clock on a bad batch; on a good batch it never fires.
Private Const CAPTCHA_CIRCUIT_ENABLED As Boolean = True
Private Const CAPTCHA_CIRCUIT_THRESHOLD As Long = 3      ' blocks needed to trip
Private Const CAPTCHA_CIRCUIT_WINDOW_SEC As Long = 120   ' ...within this window
Private Const CAPTCHA_CIRCUIT_PAUSE_SEC As Long = 180    ' ...pauses the pool this long
Private Const CAPTCHA_CIRCUIT_MAX_TRIPS As Long = 4      ' give up after this many trips

' ---- Human-in-the-loop CAPTCHA solve (v4.0) ----
' Instead of the blind timed pause above, when a CAPTCHA is hit the
' tool can PAUSE, open a visible Edge window on each worker's profile
' (so the challenge shows), let the analyst click "I'm not a robot",
' and then resume the headless searches - which now reuse the
' exemption cookie the solve just stored in each profile. Applies to
' ALL headless modes (Fast and Optimised). Triggers on the FIRST
' CAPTCHA, with a cooldown so a just-solved profile gets a chance to
' run before it can prompt again.
Private Const CAPTCHA_HUMAN_SOLVE_ENABLED As Boolean = True
Private Const HUMAN_SOLVE_COOLDOWN_SEC As Long = 30      ' min gap between solve prompts
Private Const HUMAN_SOLVE_FLUSH_SEC As Long = 3          ' let Edge write cookies before we close it
' Don't disrupt the analyst on the FIRST CAPTCHA - a lot of them clear
' on a plain retry (fresh Edge/cookies). Only arm the visible solve
' window once a task has already been retried this many times and is
' STILL blocked. In Fast mode the between-retry backoff is 0s, so with
' 2 retries the window opens after ~2 renders (~10-15s) on a genuinely
' stuck search, not on the first transient hit.
Private Const HUMAN_SOLVE_RETRY_THRESHOLD As Long = 2    ' retries before opening solve tabs

' Cookie warm-up: before the real headless batch, briefly open each
' worker's Edge profile VISIBLY against a plain google.com homepage
' (not a search) so it picks up real session cookies (NID, consent,
' etc.) from a genuine browser context instead of showing up to the
' headless requests with a bare, cookie-less profile. This is a
' mitigation, not a guarantee - it stacks with the CAPTCHA backoff
' logic above, it doesn't replace it. Flip to False to disable
' without touching the rest of the code.
Private Const WARM_PROFILE_COOKIES_ENABLED As Boolean = True
Private Const WARM_COOKIE_SECONDS As Long = 8

' Random inter-launch delay: after a worker finishes one search, it
' waits a random few seconds before grabbing its next one instead of
' firing back-to-back the instant a task completes. This staggers
' each worker's OWN request cadence over time - it does NOT delay
' the two workers' initial simultaneous start against each other,
' so parallelism is unaffected. Additive to, not a replacement for,
' the CAPTCHA backoff and cookie warm-up above.
'
' v3.5: the jitter window is now per-search-mode instead of one
' fixed pair of constants. MAX_PARALLEL stays at 2 workers in both
' Fast and Optimised - only how long each worker waits between its
' own tasks changes. Fast is tuned for throughput and accepts a
' higher CAPTCHA rate; Optimised trades run time for a much less
' bursty request pattern. If Fast is getting blocked too often,
' that's the signal to run Optimised rather than to widen Fast.
Private Const RANDOM_LAUNCH_DELAY_ENABLED As Boolean = True
' v3.9: Fast is now stripped for raw speed - ZERO jitter, and all
' the anti-CAPTCHA barricades (circuit-breaker, cookie warm-up,
' retry backoff) are turned OFF for it (see m_protectionsOn). It
' fires searches back-to-back with nothing in the way - fastest
' possible, highest CAPTCHA risk. Optimised inherits what Fast used
' to be: a modest 2-6s jitter WITH all the protections on. Use
' Optimised when the IP is getting challenged; use Fast when it
' isn't and you just want speed.
' Fast was 0-0 (zero jitter). That fired launches back-to-back, which on a
' 2-vCPU VDI let Edge child processes pile up faster than they self-exit,
' starving the CPU and causing "NO FILE EVER" launch hangs (each costing a
' full task-timeout). A small 2-3s gap - like the old fast build had - paces
' launches just enough for Edge to self-clean between them, which is a NET
' speed-up because it avoids the hangs.
Private Const FAST_LAUNCH_DELAY_MIN_SEC As Single = 2
Private Const FAST_LAUNCH_DELAY_MAX_SEC As Single = 3
Private Const OPTIMISED_LAUNCH_DELAY_MIN_SEC As Single = 2
Private Const OPTIMISED_LAUNCH_DELAY_MAX_SEC As Single = 6

' Reap orphan Edge processes every Nth completed task instead of after every
' single one. The per-task WMI kill costs ~1-2s and blocks the (single-thread)
' dispatcher - doing it every task added ~40-55s of pure overhead to a run.
' With the launch pacing above letting Edge self-clean, a periodic sweep is
' enough to contain the pileup. Stuck tasks are still reaped immediately on
' timeout, regardless of this counter.
Private Const REAP_EVERY_N_TASKS As Long = 3

' Set once per run by ApplySearchMode; SetWorkerCooldown reads these
' instead of a fixed const pair.
Private m_launchDelayMinSec As Single
Private m_launchDelayMaxSec As Single

' v3.9 master toggle for the anti-CAPTCHA barricades, set per mode by
' ApplySearchMode. True (Optimised/Visible) = circuit-breaker,
' cookie warm-up and retry backoff all active. False (Fast) = all of
' them stripped out for maximum speed.
Private m_protectionsOn As Boolean

' Leave the drive alone if it is running out of space.
Private Const DISK_ABORT_GB As Double = 2
Private Const DISK_WARN_GB As Double = 5

' While the macro runs we briefly drop the priority of the usual CPU
' hogs (Teams, Outlook, OneDrive) so Edge gets a cleaner run at the
' 2 vCPUs. Flip to False if it ever causes trouble on a given VDI.
Private Const TUNE_BACKGROUND_PRIORITIES As Boolean = True

' Shared tracker path/name come from modConfig (single source of
' truth) - TrackerFile() for the full path, TrackerFileName() for
' the "is it already open?" check.

' Module state. bAbort is flipped to True when the user presses ESC
' or any branch wants to unwind cleanly.
Private bAbort As Boolean

Private m_origCalc As XlCalculation
Private m_origScreenUpdate As Boolean
Private m_origEvents As Boolean
Private m_origAlerts As Boolean
Private m_origCursor As XlMousePointer
Private m_origPrintComm As Boolean
Private m_runtimeApplied As Boolean

' PIDs of processes we temporarily knocked down to BelowNormal, so
' we know which ones to lift back to Normal at the end of the run.
Private m_tunedPids As Collection

' How many searches, across the whole run, ended up permanently
' CAPTCHA-blocked even after every retry. Reset per-run in
' RunSearchBatch; read back by the caller for the summary/audit.
Private m_captchaFlaggedCount As Long

' Circuit-breaker state. All reset per-run in RunSearchBatch.
'   m_captchaWindowStart  - start of the current counting window
'   m_captchaHitsInWindow - blocks seen since that window opened
'   m_globalPauseUntil    - no worker may launch before this time
'   m_circuitTripCount    - how many times we've tripped this run
Private m_captchaWindowStart As Date
Private m_captchaHitsInWindow As Long
Private m_globalPauseUntil As Date
Private m_circuitTripCount As Long

' Human-solve state (v4.0).
'   m_humanSolvePending - a CAPTCHA hit; do the solve at the next
'                         dispatch cycle boundary.
'   m_lastHumanSolve    - when we last prompted, for the cooldown.
Private m_humanSolvePending As Boolean
Private m_lastHumanSolve As Date

' Entry point. Wire this up to your macro button.
Sub SearchAndSavePDF_Direct()

On Error GoTo ErrorHandler

' Start every run with a fresh debug log so the tail is always this session.
On Error Resume Next
Kill Environ("TEMP") & "\osint_debug.log"
On Error GoTo ErrorHandler
LogStep "=== Macro started (v" & TOOL_VERSION & ") ==="

' Self-heal entry guard. If the previous run died before its
' AutoRestoreRuntime fired (VBE Reset, "End" on a runtime error,
' Excel killed mid-batch), Excel is still in our muted mode. We
' recover here before doing anything else. No-op if state is clean.
ForceClearStaleOSINTState

' ---- Declarations ----
Dim ws As Worksheet, masterWs As Worksheet
Dim WshShell As Object, FSO As Object, ghostApp As Object
Dim masterWb As Workbook, wb As Workbook

Dim i As Long, fileCounter As Long, totalCPs As Long
Dim entityCount As Long, actualSearchCount As Long
Dim cpLoopPosition As Long, hdrIdx As Long
Dim mRow As Long
Dim wasAlreadyOpen As Boolean, headersOK As Boolean

Dim ecmCase As String, AlertID As String
Dim CustName As String, custAddr As String
Dim custNameNegNews As String, custNameNegNewsNoMiddle As String
Dim custNameNegNewsNoLegal As String       ' negnews name minus a legal suffix (LLC/Inc/Ltd/...)
Dim additionalCustAddr As String
Dim cpName As String, cpAddr As String
Dim cpNameNegNews(18 To 23) As String
Dim cpNameNegNewsNoMiddle(18 To 23) As String
Dim cpNameNegNewsNoLegal(18 To 23) As String

Dim baseFileName As String, desktopPath As String
Dim caseFolderPath As String, mainFolderPath As String
Dim negWords As String
Dim userProfile As String
Dim custLabel As String, cpLabel As String

Dim StartTime As Date
Dim ElapsedSeconds As Long, Minutes As Long, Seconds As Long
Dim timeString As String

Dim nameParts() As String
Dim addAddrResponse As VbMsgBoxResult

Dim expectedHeaders As Variant, headerMsg As String
Dim spSyncStatus As String      ' "OK" / "FAILED: ..." / "SKIPPED"
Dim allTasks As Collection

'--------------------------------------------------------------
' INITIALIZATION
'--------------------------------------------------------------
bAbort = False
additionalCustAddr = ""
custNameNegNewsNoMiddle = ""
wasAlreadyOpen = False
totalCPs = 0
spSyncStatus = "SKIPPED"
actualSearchCount = 0

If Not SheetExists(ThisWorkbook, "Sheet1") Then
MsgBox "Required sheet 'Sheet1' is missing. Please use the original template.", _
vbCritical, "Template Error"
Exit Sub
End If
Set ws = ThisWorkbook.Sheets("Sheet1")

Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' Figure out where this machine keeps its Desktop - it might be
' under the OneDrive-synced profile or the plain local one.
userProfile = Environ("USERPROFILE")
If FSO.FolderExists(userProfile & "\OneDrive - Community Federal Savings Bank\Desktop") Then
desktopPath = userProfile & "\OneDrive - Community Federal Savings Bank\Desktop"
Else
desktopPath = userProfile & "\Desktop"
End If

ecmCase = SanitizeFileNamePart(Trim(CStr(ws.Range("J9").Value)))
AlertID = SanitizeFileNamePart(Trim(CStr(ws.Range("J10").Value)))
CustName = SanitizeFileNamePart(Trim(CStr(ws.Range("J13").Value)))
custAddr = SanitizeFileNamePart(Trim(CStr(ws.Range("J14").Value)))

If CustName = "" And Application.CountA(ws.Range("J18:J23")) = 0 Then
MsgBox "I don't see a Customer Name or any Counterparties to search. " & _
"Please fill them in first!", vbExclamation, "Nothing to Search"
Exit Sub
End If

If ecmCase = "" Then
MsgBox "I don't see an ECM Case ID in cell J9. Please enter it before running.", _
vbExclamation, "Missing Case ID"
Exit Sub
End If

negWords = " AND (arrest OR corruption OR sentencing OR money laundering OR AML " & _
"OR launder OR embezzle OR evas OR evad OR crime OR corrupt OR bribe OR " & _
"theft OR extort OR drug OR traffic OR trafficking OR felony OR sanction " & _
"OR counterfeit OR terror)"

baseFileName = ecmCase & "_" & AlertID
baseFileName = Replace(baseFileName, " ", "_")
Do While InStr(baseFileName, "__") > 0
baseFileName = Replace(baseFileName, "__", "_")
Loop
' Trim any stray underscores the sanitiser left at the edges.
Do While Len(baseFileName) > 0 And Right(baseFileName, 1) = "_"
baseFileName = Left(baseFileName, Len(baseFileName) - 1)
Loop
Do While Len(baseFileName) > 0 And Left(baseFileName, 1) = "_"
baseFileName = Mid(baseFileName, 2)
Loop

caseFolderPath = desktopPath & "\" & ecmCase
If Not FSO.FolderExists(caseFolderPath) Then FSO.CreateFolder caseFolderPath

mainFolderPath = caseFolderPath & "\OSDD Searches"
If Not FSO.FolderExists(mainFolderPath) Then FSO.CreateFolder mainFolderPath

totalCPs = Application.CountA(ws.Range("J18:J23"))

' Ask which search mode to run this batch under (v3.5). Replaces
' the old binary non-English Yes/No prompt - Visible mode below
' covers that need directly (real windows are inherently readable/
' translatable), so there's no separate question for it any more.
LogStep "STEP 2: before search-mode prompt"
Application.StatusBar = "OSINT: Waiting for search mode selection..."
If Not PromptForSearchMode() Then Exit Sub   ' user cancelled - abort the run
LogStep "STEP 2: search-mode prompt returned mode=" & modeName(m_searchMode) & _
" (USE_HEADLESS=" & CStr(USE_HEADLESS) & ")"

' Kick off DNS/TLS warmup early so the handshake is already
' done by the time the user is through the prompts.
LogStep "STEP 2: before AutoPrewarmDNS"
AutoPrewarmDNS
LogStep "STEP 2: after AutoPrewarmDNS"

' Profile pre-warm: launch a tiny headless Edge against each
' worker's user-data-dir so the profile is populated and any
' first-run cost is paid while the analyst is still answering
' prompts. Fire-and-forget; DrainPrewarmAndCleanup picks up the
' pieces just before the real dispatch begins.
LogStep "STEP 2: before PrewarmEdgeProfiles"
PrewarmEdgeProfiles
LogStep "STEP 2: after PrewarmEdgeProfiles"

' Collect the small bits of input we need from the analyst:
' negative-news variants and any extra addresses.
If CustName <> "" Then
LogStep "STEP 3: before customer negnews InputBox"
Application.StatusBar = "OSINT: Waiting for Customer Negative News input..."
Beep
ForcePromptToFront
custNameNegNews = OsddInput( _
"Do you want to run a Negative News search for the Customer? " & _
"You can edit the name below, or delete it completely to skip.", _
"Customer Negative News", CustName)
LogStep "STEP 3: customer negnews returned, len=" & Len(custNameNegNews)
If StrPtr(custNameNegNews) = 0 Then Exit Sub

' Collapse runs of spaces so "John  Smith" doesn't look like a middle-name case.
custNameNegNews = NormalizeSpaces(custNameNegNews)
LogStep "STEP 3: after NormalizeSpaces, len=" & Len(custNameNegNews)

If Trim(custNameNegNews) <> "" Then
' An entity (legal suffix present) is NOT an individual, so it gets the
' legal-extension variant and NO middle-name search. An individual gets
' the middle-name variant instead. The two are mutually exclusive.
Dim custStripLegal As String
custStripLegal = StripLegalExtension(custNameNegNews)
Dim custIsEntity As Boolean
custIsEntity = (custStripLegal <> "" And StrComp(custStripLegal, custNameNegNews, vbTextCompare) <> 0)

If custIsEntity Then
' ---- ENTITY: extra Negative News WITHOUT the legal extension ----
LogStep "STEP 3: before customer legal-extension InputBox"
Application.StatusBar = "OSINT: Waiting for Legal Extension confirmation..."
Beep
ForcePromptToFront
custNameNegNewsNoLegal = OsddInput( _
"This name looks like it has a legal extension (" & custNameNegNews & ")." & _
vbCrLf & vbCrLf & "To ALSO run a Negative News search WITHOUT the legal " & _
"extension, confirm/edit the base name below." & vbCrLf & _
"(Clear the text or click Cancel to skip this extra search)", _
"Legal Extension Detected", custStripLegal)
If StrPtr(custNameNegNewsNoLegal) = 0 Then custNameNegNewsNoLegal = ""
custNameNegNewsNoLegal = NormalizeSpaces(custNameNegNewsNoLegal)
Else
' ---- INDIVIDUAL: extra Negative News WITHOUT the middle name ----
nameParts = Split(Trim(custNameNegNews), " ")
LogStep "STEP 3: Split into " & (UBound(nameParts) + 1) & " parts"
If UBound(nameParts) >= 2 Then
custNameNegNewsNoMiddle = nameParts(0) & " " & nameParts(UBound(nameParts))
LogStep "STEP 3: before middle-name InputBox"
Application.StatusBar = "OSINT: Waiting for Middle Name confirmation..."
Beep
ForcePromptToFront
custNameNegNewsNoMiddle = OsddInput( _
"I noticed the customer might have a middle name (" & custNameNegNews & ")." & _
vbCrLf & vbCrLf & "If you want to run an additional search WITHOUT the middle " & _
"name, you can edit/confirm the name below." & vbCrLf & _
"(Clear the text or click Cancel to skip this extra search)", _
"Middle Name Detected", custNameNegNewsNoMiddle)
LogStep "STEP 3: middle-name InputBox returned"
If StrPtr(custNameNegNewsNoMiddle) = 0 Then custNameNegNewsNoMiddle = ""
custNameNegNewsNoMiddle = NormalizeSpaces(custNameNegNewsNoMiddle)
End If
End If
End If

LogStep "STEP 3: before Sigma address MsgBox"
Application.StatusBar = "OSINT: Waiting for Sigma Address decision..."
Dim sigmaResp As Long
If gOsddAuto Then
    sigmaResp = IDNO                 ' auto OSDD: no additional Sigma address
Else
sigmaResp = TopMostMsgBox( _
"Did you identify any additional addresses for the Customer via Sigma?", _
"Sigma Address Check", MB_YESNO Or MB_ICONQUESTION Or MB_TOPMOST)
End If
LogStep "STEP 3: Sigma TopMostMsgBox returned " & sigmaResp
If sigmaResp = IDYES Then
LogStep "STEP 3: before additional-address InputBox"
Beep
ForcePromptToFront
additionalCustAddr = OsddInput( _
"Please paste the additional address identified via Sigma below:", _
"Enter Additional Address")
LogStep "STEP 3: additional-address returned"
If StrPtr(additionalCustAddr) = 0 Then Exit Sub
additionalCustAddr = NormalizeSpaces(additionalCustAddr)
End If
End If
LogStep "STEP 3: customer prompts complete"

For i = 18 To 23
cpName = SanitizeFileNamePart(Trim(CStr(ws.Range("J" & i).Value)))
If cpName <> "" Then
Application.StatusBar = "OSINT: Waiting for CP '" & cpName & "' Negative News input..."
Beep
ForcePromptToFront
cpNameNegNews(i) = OsddInput( _
"Do you want to run a Negative News search for this Counterparty? " & _
"You can edit the name or clear it to skip.", _
"Counterparty: " & cpName, cpName)
If StrPtr(cpNameNegNews(i)) = 0 Then Exit Sub
cpNameNegNews(i) = NormalizeSpaces(cpNameNegNews(i))

If Trim(cpNameNegNews(i)) <> "" Then
' Entity -> legal-extension variant, NO middle-name search.
' Individual -> middle-name variant. Mutually exclusive.
Dim cpStripLegal As String
cpStripLegal = StripLegalExtension(cpNameNegNews(i))
Dim cpIsEntity As Boolean
cpIsEntity = (cpStripLegal <> "" And StrComp(cpStripLegal, cpNameNegNews(i), vbTextCompare) <> 0)

If cpIsEntity Then
' ---- ENTITY: extra Negative News WITHOUT the legal extension ----
Application.StatusBar = "OSINT: Waiting for CP '" & cpName & "' Legal Extension confirmation..."
Beep
ForcePromptToFront
cpNameNegNewsNoLegal(i) = OsddInput( _
"This counterparty name looks like it has a legal extension (" & _
cpNameNegNews(i) & ")." & vbCrLf & vbCrLf & _
"To ALSO run a Negative News search WITHOUT the legal extension, " & _
"confirm/edit the base name below." & vbCrLf & _
"(Clear the text or click Cancel to skip this extra search)", _
"Legal Extension Detected", cpStripLegal)
If StrPtr(cpNameNegNewsNoLegal(i)) = 0 Then cpNameNegNewsNoLegal(i) = ""
cpNameNegNewsNoLegal(i) = NormalizeSpaces(cpNameNegNewsNoLegal(i))
Else
' ---- INDIVIDUAL: extra Negative News WITHOUT the middle name ----
nameParts = Split(Trim(cpNameNegNews(i)), " ")
If UBound(nameParts) >= 2 Then
cpNameNegNewsNoMiddle(i) = nameParts(0) & " " & nameParts(UBound(nameParts))
Application.StatusBar = "OSINT: Waiting for CP '" & cpName & "' Middle Name confirmation..."
Beep
ForcePromptToFront
cpNameNegNewsNoMiddle(i) = OsddInput( _
"I noticed the counterparty might have a middle name (" & _
cpNameNegNews(i) & ")." & vbCrLf & vbCrLf & _
"If you want to run an additional search WITHOUT the middle name, " & _
"you can edit/confirm the name below." & vbCrLf & _
"(Clear the text or click Cancel to skip this extra search)", _
"Middle Name Detected", cpNameNegNewsNoMiddle(i))
If StrPtr(cpNameNegNewsNoMiddle(i)) = 0 Then cpNameNegNewsNoMiddle(i) = ""
cpNameNegNewsNoMiddle(i) = NormalizeSpaces(cpNameNegNewsNoMiddle(i))
End If
End If
End If
End If
Next i

LogStep "STEP 3: before Starting Searches MsgBox"
Application.StatusBar = "OSINT: Waiting for Starting-Searches confirmation..."
If Not gOsddAuto Then
TopMostMsgBox _
"I'm going to run the searches now! Please keep your hands off the mouse " & _
"and keyboard so I don't get interrupted." & vbCrLf & vbCrLf & _
"(If things go wrong, just hold down the ESC key for 2 seconds to stop me).", _
"Starting Searches", MB_OK Or MB_ICONINFO
End If
LogStep "STEP 3: Starting Searches MsgBox dismissed -> proceeding to searches"

' Bail out if the drive is dangerously full, then flip Excel into
' its fast-mode (calc manual, no screen update, etc.).
If Not CheckDiskSpace() Then Exit Sub

AutoPrepareRuntime
StartTime = Now
LogStep "SEARCH PHASE start - total entities to process: " & (IIf(CustName <> "", 1, 0) + totalCPs)

' Build one flat queue for the customer + every counterparty.
' Keeping all tasks in a single pool means the workers never sit
' idle waiting for the next entity's batch to start.
Set allTasks = New Collection
Application.StatusBar = "OSINT: Queueing searches..."

' ---- Customer tasks ----
If CustName <> "" Then
custLabel = "Customer (" & CustName & ")"
fileCounter = 1

allTasks.Add Array(CustName, _
mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Google.pdf", _
1, "Google name", custLabel)
fileCounter = fileCounter + 1

If custAddr <> "" Then
allTasks.Add Array(custAddr, _
mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Address.pdf", _
1, "Address", custLabel)
fileCounter = fileCounter + 1

allTasks.Add Array(CustName & " + " & custAddr, _
mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Google+Address.pdf", _
1, "Name + Address", custLabel)
fileCounter = fileCounter + 1
End If

If Trim(additionalCustAddr) <> "" Then
allTasks.Add Array(CustName & " + " & additionalCustAddr, _
mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Name+Additional Address.pdf", _
1, "Name + Additional Address", custLabel)
fileCounter = fileCounter + 1
End If

If Trim(custNameNegNews) <> "" Then
allTasks.Add Array("""" & custNameNegNews & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Negative News 1.pdf", _
1, "Negative News page 1", custLabel)
fileCounter = fileCounter + 1
allTasks.Add Array("""" & custNameNegNews & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Negative News 2.pdf", _
2, "Negative News page 2", custLabel)
fileCounter = fileCounter + 1
End If

If Trim(custNameNegNewsNoMiddle) <> "" And _
StrComp(custNameNegNewsNoMiddle, custNameNegNews, vbTextCompare) <> 0 Then
allTasks.Add Array("""" & custNameNegNewsNoMiddle & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Negative News 3.pdf", _
1, "Negative News w/o middle, page 1", custLabel)
fileCounter = fileCounter + 1
allTasks.Add Array("""" & custNameNegNewsNoMiddle & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Negative News 4.pdf", _
2, "Negative News w/o middle, page 2", custLabel)
fileCounter = fileCounter + 1
End If

' Negative News WITHOUT the legal extension (2 pages), only if it differs.
If Trim(custNameNegNewsNoLegal) <> "" And _
StrComp(custNameNegNewsNoLegal, custNameNegNews, vbTextCompare) <> 0 Then
allTasks.Add Array("""" & custNameNegNewsNoLegal & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Negative News 5.pdf", _
1, "Negative News w/o legal ext, page 1", custLabel)
fileCounter = fileCounter + 1
allTasks.Add Array("""" & custNameNegNewsNoLegal & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_Customer_" & fileCounter & "_Negative News 6.pdf", _
2, "Negative News w/o legal ext, page 2", custLabel)
fileCounter = fileCounter + 1
End If
End If

' ---- Counterparty tasks (appended to same queue) ----
cpLoopPosition = 0
For i = 18 To 23
cpName = SanitizeFileNamePart(Trim(CStr(ws.Range("J" & i).Value)))
cpAddr = SanitizeFileNamePart(Trim(CStr(ws.Range("T" & i).Value)))

If cpName <> "" Then
cpLoopPosition = cpLoopPosition + 1
cpLabel = "CP " & cpLoopPosition & " of " & totalCPs & " (" & cpName & ")"
fileCounter = 1

allTasks.Add Array(cpName, _
mainFolderPath & "\" & baseFileName & "_CP" & (i - 17) & "_" & fileCounter & "_Google.pdf", _
1, "Google name", cpLabel)
fileCounter = fileCounter + 1

If cpAddr <> "" Then
allTasks.Add Array(cpAddr, _
mainFolderPath & "\" & baseFileName & "_CP" & (i - 17) & "_" & fileCounter & "_Address.pdf", _
1, "Address", cpLabel)
fileCounter = fileCounter + 1

allTasks.Add Array(cpName & " + " & cpAddr, _
mainFolderPath & "\" & baseFileName & "_CP" & (i - 17) & "_" & fileCounter & "_Google+Address.pdf", _
1, "Name + Address", cpLabel)
fileCounter = fileCounter + 1
End If

If Trim(cpNameNegNews(i)) <> "" Then
allTasks.Add Array("""" & cpNameNegNews(i) & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_CP" & (i - 17) & "_" & fileCounter & "_Negative News 1.pdf", _
1, "Negative News page 1", cpLabel)
fileCounter = fileCounter + 1
allTasks.Add Array("""" & cpNameNegNews(i) & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_CP" & (i - 17) & "_" & fileCounter & "_Negative News 2.pdf", _
2, "Negative News page 2", cpLabel)
fileCounter = fileCounter + 1
End If

If Trim(cpNameNegNewsNoMiddle(i)) <> "" And _
StrComp(cpNameNegNewsNoMiddle(i), cpNameNegNews(i), vbTextCompare) <> 0 Then
allTasks.Add Array("""" & cpNameNegNewsNoMiddle(i) & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_CP" & (i - 17) & "_" & fileCounter & "_Negative News 3.pdf", _
1, "Negative News w/o middle, page 1", cpLabel)
fileCounter = fileCounter + 1
allTasks.Add Array("""" & cpNameNegNewsNoMiddle(i) & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_CP" & (i - 17) & "_" & fileCounter & "_Negative News 4.pdf", _
2, "Negative News w/o middle, page 2", cpLabel)
fileCounter = fileCounter + 1
End If

' Negative News WITHOUT the legal extension (2 pages), only if it differs.
If Trim(cpNameNegNewsNoLegal(i)) <> "" And _
StrComp(cpNameNegNewsNoLegal(i), cpNameNegNews(i), vbTextCompare) <> 0 Then
allTasks.Add Array("""" & cpNameNegNewsNoLegal(i) & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_CP" & (i - 17) & "_" & fileCounter & "_Negative News 5.pdf", _
1, "Negative News w/o legal ext, page 1", cpLabel)
fileCounter = fileCounter + 1
allTasks.Add Array("""" & cpNameNegNewsNoLegal(i) & """" & negWords, _
mainFolderPath & "\" & baseFileName & "_CP" & (i - 17) & "_" & fileCounter & "_Negative News 6.pdf", _
2, "Negative News w/o legal ext, page 2", cpLabel)
fileCounter = fileCounter + 1
End If
End If
Next i

' ---- Gap-fill: offer to run ONLY what's missing/failed ----
' Only asked when the case folder already holds results from a
' previous attempt, so a first run is never interrupted by it.
If allTasks.count > 0 Then
Set allTasks = MaybeFilterToFailedOnly(allTasks)
If allTasks.count = 0 Then
AutoRestoreRuntime
MsgBox "Every search for this case already has a good result on disk - " & _
"there's nothing left to re-run.", vbInformation, "Nothing To Do"
GoTo CleanExit
End If
End If

' ---- Single dispatch across a shared worker pool ----
If allTasks.count > 0 Then
If USE_HEADLESS Then
Set allTasks = SortTasksHeaviestFirst(allTasks)
End If

Application.StatusBar = "OSINT: Queued " & allTasks.count & " searches - dispatching..."
LogStep "SEARCH PHASE - queued " & allTasks.count & " tasks across all entities (heaviest first)"
actualSearchCount = RunSearchBatch(allTasks, WshShell)
If bAbort Then GoTo AbortProcess

' ---- Fast-mode auto-rescue ----
' After a Fast run, offer to finish the CAPTCHA-blocked searches in
' Visible mode. Ones that succeed are saved under the normal name and
' their leftover _CAPTCHA_UNRESOLVED files are deleted automatically.
If m_searchMode = smFast Then
Dim fastRescue As Collection
Set fastRescue = New Collection
Dim rt As Long, rescTask As Variant
For rt = 1 To allTasks.count
If Not TaskHasGoodResult(allTasks(rt)) Then fastRescue.Add allTasks(rt)
Next rt

If fastRescue.count > 0 Then
Dim rescueResp As Long
rescueResp = TopMostMsgBox( _
fastRescue.count & " search(es) hit a CAPTCHA in Fast mode and were saved with a '" & _
CAPTCHA_FLAG_SUFFIX & "' tag." & vbCrLf & vbCrLf & _
"Re-run just those " & fastRescue.count & " now in VISIBLE mode?" & vbCrLf & vbCrLf & _
"Browser windows will open - keep your hands off the mouse and keyboard. Ones that " & _
"succeed are saved under the normal name and their flagged files are deleted.", _
"Finish Blocked Searches?", MB_YESNO Or MB_ICONQUESTION Or MB_TOPMOST)

If rescueResp = IDYES Then
USE_HEADLESS = False   ' visible/interactive path, for the rescue only
Application.StatusBar = "OSINT: finishing " & fastRescue.count & _
" blocked search(es) in Visible mode..."
LogStep "RESCUE (Fast->Visible): re-running " & fastRescue.count & " blocked search(es) visibly"

Dim rescuedOK As Long
rescuedOK = RunSearchBatch(fastRescue, WshShell)
If bAbort Then GoTo AbortProcess

' Delete the _CAPTCHA_UNRESOLVED sidecar ONLY for tasks whose visible
' re-run produced a genuinely full-size results page under the normal
' name. The old check was > MIN_PDF_SIZE_BYTES (5 KB), so a re-run that
' hit the SAME CAPTCHA (~71 KB) still counted as "rescued" and had its
' flag deleted - leaving a block page sitting under a clean name. Must
' match TaskHasGoodResult's bar: at least CAPTCHA_SIZE_HINT.
Dim cleaned As Long, spRes As String, fpRes As String
For rt = 1 To fastRescue.count
rescTask = fastRescue(rt)
spRes = CStr(rescTask(1))
On Error Resume Next
If Dir(spRes) <> "" Then
If FileLen(spRes) >= CAPTCHA_SIZE_HINT Then
fpRes = FlagPathWithSuffix(spRes, CAPTCHA_FLAG_SUFFIX)
If Dir(fpRes) <> "" Then
Kill fpRes
cleaned = cleaned + 1
End If
End If
End If
On Error GoTo ErrorHandler
Next rt

actualSearchCount = actualSearchCount + rescuedOK

' Recount what's STILL unresolved so the summary is accurate.
m_captchaFlaggedCount = 0
For rt = 1 To allTasks.count
If Not TaskHasGoodResult(allTasks(rt)) Then _
m_captchaFlaggedCount = m_captchaFlaggedCount + 1
Next rt

LogStep "RESCUE (Fast->Visible): " & rescuedOK & " saved, " & cleaned & _
" flagged file(s) deleted, " & m_captchaFlaggedCount & " still unresolved"
End If
End If
End If
End If

' -------- Final CAPTCHA-size safety sweep --------
' Guarantee, no matter what: a CAPTCHA-size PDF is never left sitting
' under a clean name. Whatever route produced it - a content-detect
' false-accept, a rescue that cleared the flag, anything - any expected
' output still under CAPTCHA_SIZE_HINT gets the _CAPTCHA_UNRESOLVED
' suffix here, so a 71 KB block page can't pass as a real result.
If Not allTasks Is Nothing Then
Dim swT As Long, swTask As Variant, swSave As String, swFlag As String, swept As Long, unflagged As Long
For swT = 1 To allTasks.count
swTask = allTasks(swT)
swSave = CStr(swTask(1))
swFlag = FlagPathWithSuffix(swSave, CAPTCHA_FLAG_SUFFIX)
On Error Resume Next
If Dir(swSave) <> "" Then
If FileLen(swSave) < CAPTCHA_SIZE_HINT Then
' Small/blocked page left under a clean name -> flag it.
If Dir(swFlag) <> "" Then Kill swFlag   ' replace any stale flag file
Name swSave As swFlag
swept = swept + 1
Else
' Full-size clean result present -> remove any leftover flag
' sidecar so this search isn't reported both saved AND unresolved.
If Dir(swFlag) <> "" Then
Kill swFlag
unflagged = unflagged + 1
End If
End If
End If
On Error GoTo ErrorHandler
Next swT
If swept > 0 Or unflagged > 0 Then LogStep "SAFETY SWEEP: flagged " & swept & _
" CAPTCHA-size file(s) under a clean name; cleared " & unflagged & _
" stale flag(s) where a clean full-size result exists"
End If

' Final authoritative unresolved count - after the rescue AND the sweep,
' so the summary dialog matches what's actually on disk (a search that
' succeeded on a late/retry write is no longer counted as blocked).
If Not allTasks Is Nothing Then
m_captchaFlaggedCount = 0
Dim fcT As Long
For fcT = 1 To allTasks.count
If Not TaskHasGoodResult(allTasks(fcT)) Then _
m_captchaFlaggedCount = m_captchaFlaggedCount + 1
Next fcT
End If

' Elapsed time - using DateDiff instead of subtraction so a run
' that straddles midnight doesn't come out negative.
ElapsedSeconds = DateDiff("s", StartTime, Now)
Minutes = ElapsedSeconds \ 60
Seconds = ElapsedSeconds Mod 60
LogStep "SEARCH PHASE end - elapsed=" & ElapsedSeconds & "s, searches=" & actualSearchCount
If Minutes > 0 Then
timeString = Minutes & " minute(s) and " & Seconds & " second(s)"
Else
timeString = Seconds & " second(s)"
End If

entityCount = IIf(CustName <> "" Or custNameNegNews <> "", 1, 0) + totalCPs
' actualSearchCount is the authoritative successful-PDF counter.
' We trust this over re-scanning the folder - OneDrive can lie
' about what's on disk while a sync is still in flight.

' Also append this run to the LOCAL "Audit_Log" sheet (the v3.4 OSINT
' run-summary log) - this is what re-activates that dormant tab. One row
' per OSDD run: entities / searches / time taken / tool version / CAPTCHA
' flagged. Best-effort and self-contained, so it never disrupts the run.
WriteLocalAuditLog ecmCase, entityCount, actualSearchCount, timeString, m_captchaFlaggedCount, "None"

' Hand Excel back to the user. Folder open is deferred until
' after they dismiss the success dialog (analyst-requested UX).
AutoRestoreRuntime

' Push a row to the shared SharePoint tracker so the team can
' see this run. Failures here are non-fatal - the local audit
' log already has the authoritative record.
Application.StatusBar = "OSINT: Pushing run to SharePoint master tracker..."

For Each wb In Application.Workbooks
If wb.Name = TrackerFileName() Then
Set masterWb = wb
wasAlreadyOpen = True
Exit For
End If
Next wb

If masterWb Is Nothing Then
Set ghostApp = CreateObject("Excel.Application")
ghostApp.Visible = False
ghostApp.DisplayAlerts = False
ghostApp.EnableEvents = False

On Error Resume Next
Err.Clear
Set masterWb = ghostApp.Workbooks.Open(fileName:=TrackerFile(), UpdateLinks:=False)
If masterWb Is Nothing Then
spSyncStatus = "FAILED: " & IIf(Err.Number <> 0, _
"Err " & Err.Number & " - " & Err.Description, _
"WebDAV open returned Nothing (offline, auth, or file renamed?)")
End If
On Error GoTo ErrorHandler
End If

If Not masterWb Is Nothing Then
If Not masterWb.ReadOnly Then
Set masterWs = masterWb.Sheets("Sheet1")
expectedHeaders = Array("Date & Time", "Analyst ID", "ECM Case ID", _
"Total Entities", "Total Searches", "Time Taken", "Tool Version")
headersOK = True
headerMsg = ""

For hdrIdx = LBound(expectedHeaders) To UBound(expectedHeaders)
If masterWs.Cells(1, hdrIdx + 1).Value <> expectedHeaders(hdrIdx) Then
headersOK = False
headerMsg = headerMsg & "- Col " & _
Split(masterWs.Cells(1, hdrIdx + 1).Address, "$")(1) & _
" expected '" & expectedHeaders(hdrIdx) & "' but found '" & _
masterWs.Cells(1, hdrIdx + 1).Value & "'" & vbCrLf
End If
Next hdrIdx

If headersOK Then
mRow = masterWs.Cells(masterWs.Rows.count, "A").End(xlUp).row + 1
masterWs.Cells(mRow, 1).Value = Now
masterWs.Cells(mRow, 2).Value = Environ("USERNAME")
masterWs.Cells(mRow, 3).Value = ecmCase
masterWs.Cells(mRow, 4).Value = entityCount
masterWs.Cells(mRow, 5).Value = actualSearchCount
masterWs.Cells(mRow, 6).Value = timeString
masterWs.Cells(mRow, 7).Value = TOOL_VERSION

If wasAlreadyOpen Then
masterWb.Save
Else
masterWb.Close SaveChanges:=True
End If
spSyncStatus = "OK"
Else
spSyncStatus = "FAILED: header mismatch -" & vbCrLf & headerMsg
If Not wasAlreadyOpen Then masterWb.Close SaveChanges:=False
End If
Else
spSyncStatus = "FAILED: master tracker is read-only (another user has it open?)"
If Not wasAlreadyOpen Then masterWb.Close SaveChanges:=False
End If
End If

If Not ghostApp Is Nothing Then
On Error Resume Next
ghostApp.Quit
Set ghostApp = Nothing
On Error GoTo ErrorHandler
End If

Application.StatusBar = False

'--------------------------------------------------------------
' STEP 11: Final summary dialog -> folder open on OK
'--------------------------------------------------------------
AppActivate Application.caption

Dim summary As String
summary = "All done! I saved " & actualSearchCount & " PDF(s) directly into:" & vbCrLf & _
"  " & ecmCase & "\OSDD Searches" & vbCrLf & vbCrLf & _
"Search mode: " & modeName(m_searchMode) & vbCrLf & _
"Total search time: " & timeString

If m_captchaFlaggedCount > 0 Then
summary = summary & vbCrLf & vbCrLf & _
"Heads up: " & m_captchaFlaggedCount & " search(es) kept hitting what looks like a Google " & _
"CAPTCHA even after repeated retries. Those are saved with a '" & CAPTCHA_FLAG_SUFFIX & _
"' suffix instead of the normal filename so they can't be mistaken for a clean result - " & _
"take a look and re-run just those if needed."
End If

If gOsddAuto Then
    gOsddSummary = summary           ' shown in the "case ready" message instead
Else
MsgBox summary, vbInformation, "Success"

' User dismissed the dialog - now pop the case folder so they
' land on the freshly-saved PDFs.
On Error Resume Next
Shell "explorer.exe """ & mainFolderPath & """", vbNormalFocus
On Error GoTo ErrorHandler
End If

Set WshShell = Nothing
Set FSO = Nothing
GoTo CleanExit

'------------------------------------------------------------------
' ABORT PATH (user pressed ESC)
'------------------------------------------------------------------
AbortProcess:
AutoRestoreRuntime

On Error Resume Next
AppActivate Application.caption
On Error GoTo 0
MsgBox "You safely canceled the process. Any PDFs already downloaded " & _
"are in the case folder." & vbCrLf & vbCrLf & _
"Note: any Edge processes still running will finish on their own " & _
"and their output will also land in the folder.", _
vbInformation, "Canceled"
GoTo CleanExit

' Catch-all for anything unexpected. We log the error on the audit
' sheet and show the analyst a dialog instead of dying silently.
ErrorHandler:
Dim errN As Long, errD As String, errS As String
errN = Err.Number
errD = Err.Description
errS = Err.source

On Error Resume Next
LogStep "ERROR Err " & errN & ": " & errD & " [src: " & errS & "]"

AutoRestoreRuntime
Application.StatusBar = False

MsgBox "Something went wrong and I had to stop." & vbCrLf & vbCrLf & _
"Error " & errN & ": " & errD & vbCrLf & _
"Source: " & errS & vbCrLf & vbCrLf & _
"Details are logged to " & Environ("TEMP") & "\osint_debug.log. " & _
"Any PDFs already saved are safe in the case folder. " & _
"You can re-run the macro; already-downloaded searches will be " & _
"overwritten with fresh copies.", _
vbCritical + vbSystemModal, "OSINT Error"
On Error GoTo 0
' fall through to CleanExit

CleanExit:
On Error Resume Next
If Not ghostApp Is Nothing Then
ghostApp.Quit
Set ghostApp = Nothing
End If
AutoRestoreRuntime
Application.StatusBar = False
On Error GoTo 0

End Sub

' Refuses to run if the user's drive is close to full. Batch PDF
' runs can easily eat a few hundred MB and a half-written case
' folder is worse than no folder at all.
Private Function CheckDiskSpace() As Boolean
On Error Resume Next
Dim FSO As Object, drv As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Set drv = FSO.GetDrive(FSO.GetDriveName(Environ("USERPROFILE")))

Dim freeGB As Double
freeGB = drv.FreeSpace / (1024 ^ 3)
On Error GoTo 0

If freeGB < DISK_ABORT_GB Then
MsgBox "CRITICAL: Only " & Format(freeGB, "0.0") & " GB free on your VDI." & vbCrLf & _
"The tool needs at least " & DISK_ABORT_GB & " GB to run safely. Aborting." & vbCrLf & vbCrLf & _
"Please empty Recycle Bin, clear Downloads, or contact IT to free up space.", _
vbCritical, "Disk Full"
CheckDiskSpace = False
ElseIf freeGB < DISK_WARN_GB Then
Dim r As VbMsgBoxResult
r = MsgBox("WARNING: Only " & Format(freeGB, "0.0") & " GB free on your VDI." & vbCrLf & _
"A typical 20-search run uses ~700 MB. Continue anyway?", _
vbYesNo + vbExclamation, "Low Disk Space")
CheckDiskSpace = (r = vbYes)
Else
CheckDiskSpace = True
End If
End Function

' Small helper for the serial path - pokes the status bar so the
' analyst can see which search is currently running.
Private Sub SetSearchStatus(entityLabel As String, searchType As String)
Application.StatusBar = "OSINT: " & entityLabel & " -> " & searchType & "..."
End Sub

' Flips Excel into its fastest, least-distracting state while the
' macro runs (no calc, no screen redraw, no dialogs) and saves
' the original values so we can restore them on the way out.
Public Sub AutoPrepareRuntime()
On Error Resume Next
If m_runtimeApplied Then Exit Sub

With Application
m_origCalc = .Calculation
m_origScreenUpdate = .ScreenUpdating
m_origEvents = .EnableEvents
m_origAlerts = .DisplayAlerts
m_origCursor = .Cursor
m_origPrintComm = .PrintCommunication

.Calculation = xlCalculationManual
.ScreenUpdating = False
.EnableEvents = False
.DisplayAlerts = False
.Cursor = xlWait
.PrintCommunication = False
End With

m_runtimeApplied = True
On Error GoTo 0

TuneBackgroundApps
End Sub

Public Sub AutoRestoreRuntime()
On Error Resume Next

' Always try to restore priorities, even if Excel state wasn't
' applied - a half-way failure still might have tuned a few PIDs.
RestoreBackgroundApps

If Not m_runtimeApplied Then
Application.StatusBar = False
Exit Sub
End If

With Application
.Calculation = m_origCalc
.ScreenUpdating = m_origScreenUpdate
.EnableEvents = m_origEvents
.DisplayAlerts = m_origAlerts
.Cursor = m_origCursor
.PrintCommunication = m_origPrintComm
.StatusBar = False
End With

m_runtimeApplied = False
On Error GoTo 0
End Sub

' If a previous OSINT run died before AutoRestoreRuntime fired
' (VBE Reset, "End" on a runtime error, Excel killed mid-batch...)
' Excel stays in our muted mode: Calculation = manual, EnableEvents
' off, DisplayAlerts off, etc. In Excel 365 that shows up on every
' formula as a yellow "Stale" badge with the value struck through,
' which is alarming on hidden backend sheets the analyst can't fix
' from the ribbon.
'
' This sub force-clears that residue. It only intervenes when it
' sees clear OSINT-shaped fingerprints (our status-bar prefix, or
' the trifecta of muted Excel toggles) so it never accidentally
' overrides a user who genuinely keeps Excel on manual calc.
Private Sub ForceClearStaleOSINTState()
On Error Resume Next

Dim looksStale As Boolean
looksStale = False

' Smoking gun #1: our own status-bar message is still up.
Dim sb As Variant
sb = Application.StatusBar
If VarType(sb) = vbString Then
If InStr(1, CStr(sb), "OSINT", vbTextCompare) > 0 Then looksStale = True
End If

' Smoking gun #2: the trifecta of toggles we mute together.
' A normal user would never have all three flipped at once.
If Not looksStale Then
If Application.Calculation = xlCalculationManual And _
Application.EnableEvents = False And _
Application.DisplayAlerts = False Then
looksStale = True
End If
End If

If Not looksStale Then Exit Sub

LogStep "RECOVER: stale OSINT state detected on entry - force-clearing " & _
"(calc/events/alerts/cursor/screen/print/statusbar)"

With Application
.Calculation = xlAutomatic
.EnableEvents = True
.ScreenUpdating = True
.DisplayAlerts = True
.Cursor = xlDefault
.PrintCommunication = True
.StatusBar = False
End With

' Lift any background-app priorities the dead run lowered.
' Safe no-op if nothing is currently tuned down.
RestoreBackgroundApps

' Reset our flag so the new run's AutoPrepareRuntime starts
' from a clean baseline (it'll save the freshly-restored state
' as the "original" to put back later).
m_runtimeApplied = False

On Error GoTo 0
End Sub

' Public, manual-recovery entry point. Bind it to a button or run
' it from Alt+F8 if the workbook ever ends up with "Stale" cell
' indicators / events not firing / formulas not recalculating
' after an OSINT run that didn't tear down cleanly.
'
' Safe to run any time - it's a no-op if there's nothing to clean.
Public Sub OSINT_RecoverState()
On Error Resume Next

Dim took As String

' Authoritative path: we have saved originals from a still-live
' AutoPrepareRuntime, restore them exactly.
If m_runtimeApplied Then
AutoRestoreRuntime
took = "Restored Excel state from saved originals."
Else
' Fallback path: no saved originals, but state may still
' look stuck. ForceClearStaleOSINTState only fires if it
' sees real OSINT residue.
ForceClearStaleOSINTState
took = "Force-cleared any leftover OSINT state. " & _
"(If Excel was already healthy, nothing changed.)"
End If

' Always-safe housekeeping regardless of which path ran.
Application.StatusBar = False
RestoreBackgroundApps

' Nudge a recalc so any "Stale" badges clear immediately
' instead of waiting for the next edit.
On Error Resume Next
Application.CalculateFull
On Error GoTo 0

MsgBox took & vbCrLf & vbCrLf & _
"Calculation, events, alerts, screen updates, status bar and " & _
"background-app priorities are back to defaults. " & _
"Stale-cell indicators have been cleared.", _
vbInformation, "OSINT Recover"

On Error GoTo 0
End Sub

' Drop Teams / Outlook / OneDrive to BelowNormal so they stop
' fighting Edge for the 2 vCPUs. Remembers the PIDs we actually
' managed to change, so RestoreBackgroundApps can put them back.
' Uses WMI, which is usable on locked-down VDIs without admin.
Private Sub TuneBackgroundApps()
If Not TUNE_BACKGROUND_PRIORITIES Then Exit Sub

Set m_tunedPids = New Collection
On Error Resume Next

Dim wmi As Object
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
If wmi Is Nothing Then Exit Sub

' Process names we want to de-prioritise. Upper/lower case
' matches what Task Manager / WMI actually report.
Dim targets As Variant
targets = Array( _
"Teams.exe", "ms-teams.exe", "msteams.exe", _
"msedgewebview2.exe", _
"OUTLOOK.EXE", _
"OneDrive.exe")

Const PRIORITY_BELOW_NORMAL As Long = 16384
Dim target As Variant
Dim procs As Object, p As Object
Dim lowered As Long

For Each target In targets
Set procs = wmi.ExecQuery( _
"SELECT ProcessId FROM Win32_Process WHERE Name = '" & _
Replace(CStr(target), "'", "''") & "'")
If Not procs Is Nothing Then
For Each p In procs
If p.SetPriority(PRIORITY_BELOW_NORMAL) = 0 Then
m_tunedPids.Add CLng(p.ProcessId)
lowered = lowered + 1
End If
Next p
End If
Next target

If lowered > 0 Then
LogStep "BG TUNE: lowered " & lowered & " background process(es) to BelowNormal"
End If
End Sub

' Lifts every PID we tuned back to Normal. Processes that have
' exited in the meantime are just ignored.
Private Sub RestoreBackgroundApps()
If m_tunedPids Is Nothing Then Exit Sub
If m_tunedPids.count = 0 Then
Set m_tunedPids = Nothing
Exit Sub
End If

On Error Resume Next

Dim wmi As Object
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
If wmi Is Nothing Then
Set m_tunedPids = Nothing
Exit Sub
End If

Const PRIORITY_NORMAL As Long = 32
Dim pid As Variant
Dim procs As Object, p As Object
Dim restored As Long

For Each pid In m_tunedPids
Set procs = wmi.ExecQuery( _
"SELECT ProcessId FROM Win32_Process WHERE ProcessId = " & CLng(pid))
If Not procs Is Nothing Then
For Each p In procs
If p.SetPriority(PRIORITY_NORMAL) = 0 Then
restored = restored + 1
End If
Next p
End If
Next pid

LogStep "BG TUNE: restored " & restored & " background process(es) to Normal"
Set m_tunedPids = Nothing
End Sub

' Fires a throwaway HEAD request to Google so the DNS + TLS
' handshake is already cached by the time Edge gets here. We fire
' and forget - there's no point waiting for the response.
Public Sub AutoPrewarmDNS()
On Error Resume Next
Dim http As Object
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.SetTimeouts 1500, 2000, 3000, 3000
http.Open "HEAD", "https://www.google.com/", True   ' True = async
http.Send
' Intentionally NOT calling WaitForResponse - fire and forget.
Set http = Nothing
On Error GoTo 0
End Sub

' Pre-warms each worker's headless Edge profile by firing a tiny
' --print-to-pdf of about:blank against every worker's user-data-dir.
' That forces Edge to populate the first-run profile state so the
' first real search doesn't pay the cold-start tax.
Public Sub PrewarmEdgeProfiles()
On Error Resume Next

If Not USE_HEADLESS Then Exit Sub

Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

Dim w As Long, profilePath As String, prewarmPdf As String, cmd As String
Dim browserExe As String
Dim fired As Long: fired = 0

For w = 0 To MAX_PARALLEL - 1
' v3.7: resolved per worker, since each slot may be a different
' browser now. Skip this slot if its exe isn't actually present.
browserExe = GetEdgePath()
If Len(browserExe) > 0 Then
If Dir(browserExe) <> "" Then

profilePath = GetWorkerProfilePath(w)
If Not FSO.FolderExists(profilePath) Then FSO.CreateFolder profilePath

prewarmPdf = Environ("TEMP") & "\osint_prewarm_w" & w & ".pdf"
If Dir(prewarmPdf) <> "" Then Kill prewarmPdf

' Stripped-down version of the real headless cmd: same flags
' that touch the profile dir, none of the search-specific bits.
cmd = """" & browserExe & """" & _
" --headless=new" & _
" --disable-gpu" & _
" --user-data-dir=""" & profilePath & """" & _
" --no-first-run" & _
" --no-default-browser-check" & _
" --disable-logging --log-level=3" & _
" --disable-background-networking" & _
" --disable-sync" & _
" --disable-default-apps" & _
" --disable-component-update" & _
" --disable-client-side-phishing-detection" & _
" --disable-extensions" & _
" --metrics-recording-only" & _
" --mute-audio" & _
" --print-to-pdf=""" & prewarmPdf & """" & _
" about:blank"

wsh.Run cmd, 0, False
fired = fired + 1
End If
End If
Next w

If fired > 0 Then
LogStep "PREWARM: fired " & fired & " Edge profile warmer(s) - " & _
"running in background while user is prompted"
End If

On Error GoTo 0
End Sub

' Warms each worker's Edge profile against a plain google.com
' homepage so the profile picks up real session cookies (NID,
' consent, etc.) before the headless search batch - instead of the
' searches being the very first thing that profile says to Google.
'
' v3.9: runs HEADLESS (hidden) now, not in visible windows. It fires
' a throwaway --print-to-pdf of google.com, which loads the page
' (storing its cookies) and exits on its own - no windows pop up.
' Trade-off: a hidden warm-up is less convincing than a real visible
' one (it's another headless hit rather than a genuine browsing
' context), so it seeds cookies but carries less anti-CAPTCHA weight.
Public Sub WarmEdgeProfileCookies()
On Error Resume Next

Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")

Dim w As Long, profilePath As String, cmd As String, warmPdf As String
Dim browserExe As String
Dim warmed As Long: warmed = 0

For w = 0 To MAX_PARALLEL - 1
browserExe = GetEdgePath()
If Len(browserExe) > 0 Then
If Dir(browserExe) <> "" Then

profilePath = GetWorkerProfilePath(w)
If Not FSO.FolderExists(profilePath) Then FSO.CreateFolder profilePath

warmPdf = Environ("TEMP") & "\osint_cookiewarm_w" & w & ".pdf"
If Dir(warmPdf) <> "" Then Kill warmPdf

' Headless visit to google.com (not a search) - loads the page,
' stores its cookies in the profile, self-exits after the print.
cmd = """" & browserExe & """" & _
" --headless=new" & _
" --disable-gpu" & _
" --user-data-dir=""" & profilePath & """" & _
" --no-first-run" & _
" --no-default-browser-check" & _
" --disable-logging --log-level=3" & _
" --disable-sync" & _
" --disable-extensions" & _
" --mute-audio" & _
" --print-to-pdf=""" & warmPdf & """" & _
" ""https://www.google.com/"""

wsh.Run cmd, 0, False   ' 0 = hidden (headless has no window anyway)
warmed = warmed + 1
End If
End If
Next w

If warmed = 0 Then Exit Sub

LogStep "COOKIE WARM: fired " & warmed & " HEADLESS (hidden) warm-up(s) to pick up Google session cookies"
Application.StatusBar = "OSINT: warming up Google session cookies..."

' Give the page (and any cookie-setting redirects) time to fully
' settle before we hand the profile back to the search dispatcher.
SmartWait CSng(WARM_COOKIE_SECONDS), True

' Belt-and-braces: kill any straggler that didn't self-exit, and
' drop the throwaway warm-up PDFs.
For w = 0 To MAX_PARALLEL - 1
KillEdgeWorkerProcesses w
warmPdf = Environ("TEMP") & "\osint_cookiewarm_w" & w & ".pdf"
If Dir(warmPdf) <> "" Then Kill warmPdf
Next w

LogStep "COOKIE WARM: done after " & WARM_COOKIE_SECONDS & "s (headless)"
Application.StatusBar = False

On Error GoTo 0
End Sub

' Called just before the real dispatch starts. Usually the prewarm
' Edge processes have already exited on their own; if not, we wait
' a short grace period and then kill any straggler holding a lock.
Private Sub DrainPrewarmAndCleanup()
On Error Resume Next

Dim w As Long, prewarmPdf As String
Dim drainStartT As Single: drainStartT = Timer
Dim allDone As Boolean

Do
allDone = True
For w = 0 To MAX_PARALLEL - 1
prewarmPdf = Environ("TEMP") & "\osint_prewarm_w" & w & ".pdf"
If Dir(prewarmPdf) = "" Then
allDone = False
Exit For
End If
Next w
If allDone Then Exit Do
' 2 s ceiling is plenty - print-to-pdf of about:blank is sub-second.
If Timer - drainStartT > 2 Then Exit Do
SmartWait 0.1, True
Loop

' Belt-and-braces: kill any straggler msedge.exe still bound to
' our worker profiles. No-op if prewarm already exited cleanly.
For w = 0 To MAX_PARALLEL - 1
KillEdgeWorkerProcesses w
Next w

' Drop the throwaway prewarm PDFs.
For w = 0 To MAX_PARALLEL - 1
prewarmPdf = Environ("TEMP") & "\osint_prewarm_w" & w & ".pdf"
If Dir(prewarmPdf) <> "" Then Kill prewarmPdf
Next w

LogStep "PREWARM: drain done in " & Format(Timer - drainStartT, "0.00") & "s" & _
" (allDone=" & allDone & ")"

On Error GoTo 0
End Sub

' Nudge Excel to the top so the next InputBox / MsgBox isn't hidden
' behind something.
Private Sub ForcePromptToFront()
On Error Resume Next
AppActivate Application.caption
DoEvents
On Error GoTo 0
End Sub

' Direct Win32 MessageBoxW call with TOPMOST + SETFOREGROUND.
' VBA's own MsgBox (even with vbSystemModal) loses the z-order
' fight on Citrix / V2 Cloud, so we go around it and talk to
' user32 directly. Returns Win32 IDOK / IDYES / IDNO.
' InputBox, except in auto OSDD where it returns the default unchanged
' (what pressing OK would give).
Private Function OsddInput(ByVal prompt As String, Optional ByVal Title As String = "", _
                           Optional ByVal dflt As String = "") As String
If gOsddAuto Then
    OsddInput = dflt
Else
    OsddInput = InputBox(prompt, Title, dflt)
End If
End Function

Private Function TopMostMsgBox(ByVal msg As String, _
ByVal Title As String, _
ByVal flags As Long) As Long
Beep   ' audible cue for the analyst
On Error Resume Next
AppActivate Application.caption
DoEvents
On Error GoTo 0
TopMostMsgBox = MessageBoxW(0, StrPtr(msg), StrPtr(Title), _
flags Or MB_TOPMOST Or MB_SETFOREGROUND Or MB_SYSTEMMODAL)
End Function

'------------------------------------------------------------------
' Search mode selection (v3.6)
'------------------------------------------------------------------
' Real 3-button chooser (frmSearchMode) - Fast / Optimised / Visible
' / Cancel, same pattern as the existing UserForm1 (Transaction
' Files / Non Alerted / Cancel). Requires frmSearchMode to have been
' built in the VBA IDE's form designer first (its button layout
' lives in a binary blob, not something pasteable as plain code) -
' see frmSearchMode.frm for the code-behind and build steps.
'
' Returns False if the analyst cancelled (caller should abort the
' run), True once m_searchMode/USE_HEADLESS/jitter are all set.
Private Function PromptForSearchMode() As Boolean
If gOsddAuto Then
    ApplySearchMode smFast           ' Option 2
    PromptForSearchMode = True
    Exit Function
End If
frmSearchMode.Show vbModal

If frmSearchMode.userCancelled Then
PromptForSearchMode = False
Else
ApplySearchMode CLng(frmSearchMode.SelectedMode)
PromptForSearchMode = True
End If

Unload frmSearchMode
End Function

' Sets USE_HEADLESS and the per-mode jitter window for the rest of
' the run. Called once, right after the analyst picks a mode.
Private Sub ApplySearchMode(ByVal mode As OsintSearchMode)
m_searchMode = mode
Select Case mode
Case smFast
' Raw speed: zero jitter, every barricade off.
USE_HEADLESS = True
m_launchDelayMinSec = FAST_LAUNCH_DELAY_MIN_SEC
m_launchDelayMaxSec = FAST_LAUNCH_DELAY_MAX_SEC
m_protectionsOn = False
Case smOptimised
' What Fast used to be: modest jitter WITH all protections on.
USE_HEADLESS = True
m_launchDelayMinSec = OPTIMISED_LAUNCH_DELAY_MIN_SEC
m_launchDelayMaxSec = OPTIMISED_LAUNCH_DELAY_MAX_SEC
m_protectionsOn = True
Case smVisible
USE_HEADLESS = False
' Jitter doesn't apply on this path - RunSearchBatchSerial /
' PrintSearchToPDF_Interactive is inherently serial and already
' paced by real dialog choreography. Values set anyway so
' SetWorkerCooldown has something sane if it's ever reached.
m_launchDelayMinSec = OPTIMISED_LAUNCH_DELAY_MIN_SEC
m_launchDelayMaxSec = OPTIMISED_LAUNCH_DELAY_MAX_SEC
m_protectionsOn = True
End Select
LogStep "MODE: search mode set to " & modeName(mode) & _
" (USE_HEADLESS=" & USE_HEADLESS & ", jitter=" & _
m_launchDelayMinSec & "-" & m_launchDelayMaxSec & "s, protections=" & _
m_protectionsOn & ")"
End Sub

Private Function modeName(ByVal mode As OsintSearchMode) As String
Select Case mode
Case smFast: modeName = "Fast"
Case smOptimised: modeName = "Optimised"
Case smVisible: modeName = "Visible"
Case Else: modeName = "Unknown"
End Select
End Function

'------------------------------------------------------------------
' File-based step log. Writes one timestamped line per call to
' %TEMP%\osint_debug.log. Survives VBA crashes, hangs, Excel kills.
'------------------------------------------------------------------
Public Sub LogStep(ByVal msg As String)
On Error Resume Next
Dim fnum As Integer, logPath As String, attempt As Long
logPath = Environ("TEMP") & "\osint_debug.log"
' Retry up to 3x with Shared lock so Notepad/Notepad++ holding the
' file open cannot silently swallow log lines.
For attempt = 1 To 3
fnum = FreeFile
Err.Clear
Open logPath For Append Shared As #fnum
If Err.Number = 0 Then
Print #fnum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & msg
Close #fnum
Exit For
End If
Next attempt
On Error GoTo 0
End Sub

Public Function BuildEdgeArgs(ByVal url As String) As String
' Only the interactive fallback (USE_HEADLESS = False) uses this.
' Flags picked to play nicely with V2 Cloud's no-GPU environment.
BuildEdgeArgs = _
"--disable-gpu --disable-gpu-compositing --disable-software-rasterizer " & _
"--disable-background-timer-throttling --disable-renderer-backgrounding " & _
"--disable-backgrounding-occluded-windows " & _
"--disable-features=CalculateNativeWinOcclusion,VizDisplayCompositor " & _
"--renderer-process-limit=2 " & _
"--no-first-run --no-default-browser-check --disable-sync --disable-extensions " & _
"--lang=en-US """ & url & """"
End Function


' ---- Gap-fill filter (v3.8) ----
' A re-run used to re-fire EVERY search in the case, even the ones
' that already came back clean. On a batch where 5 of 25 got
' CAPTCHA-blocked that's 20 needless requests against the very rate
' limit doing the blocking - each retry making the next retry more
' likely to fail. This offers to queue only the searches that don't
' already have a good PDF on disk.
'
' Stays silent on a first run (nothing on disk yet = nothing to
' skip), so it never adds a click to the normal path.
'
' Returns either the original collection or a filtered copy.
Private Function MaybeFilterToFailedOnly(ByVal tasks As Collection) As Collection
On Error GoTo Fallback

Dim i As Long, goodCount As Long
For i = 1 To tasks.count
If TaskHasGoodResult(tasks(i)) Then goodCount = goodCount + 1
Next i

' First run (or everything previously failed) - nothing to offer.
If goodCount = 0 Then
Set MaybeFilterToFailedOnly = tasks
Exit Function
End If

Dim pendingCount As Long
pendingCount = tasks.count - goodCount

Dim resp As Long
resp = TopMostMsgBox( _
"This case already has " & goodCount & " good result(s) saved from a " & _
"previous run, and " & pendingCount & " that are missing or were " & _
"CAPTCHA-blocked." & vbCrLf & vbCrLf & _
"YES = re-run ONLY the " & pendingCount & " missing/failed search(es)." & vbCrLf & _
"   Recommended - far fewer requests, so much less likely to trip" & vbCrLf & _
"   the block again. Existing good results are left untouched." & vbCrLf & vbCrLf & _
"NO = re-run all " & tasks.count & " from scratch (overwrites everything).", _
"Re-run Failed Only?", MB_YESNO Or MB_ICONQUESTION Or MB_TOPMOST)

If resp <> IDYES Then
LogStep "GAP-FILL: analyst chose FULL re-run of all " & tasks.count & " task(s)"
Set MaybeFilterToFailedOnly = tasks
Exit Function
End If

Dim filtered As New Collection
For i = 1 To tasks.count
If Not TaskHasGoodResult(tasks(i)) Then filtered.Add tasks(i)
Next i

LogStep "GAP-FILL: re-running " & filtered.count & " missing/failed task(s), " & _
"skipping " & goodCount & " already-good (saved " & goodCount & " request(s))"
Set MaybeFilterToFailedOnly = filtered
Exit Function

Fallback:
' Any trouble inspecting the folder - fail open to a normal full run
' rather than silently skipping work.
LogStep "GAP-FILL: filter failed (err " & Err.Number & ") - running full batch"
Set MaybeFilterToFailedOnly = tasks
End Function

' True only if this task's expected PDF is present AND big enough to
' be a real SERP. Anything missing, truncated, CAPTCHA-sized, or
' carrying a leftover _CAPTCHA_UNRESOLVED sidecar counts as needing
' another attempt.
Private Function TaskHasGoodResult(ByVal task As Variant) As Boolean
On Error GoTo NotGood

Dim savePath As String
savePath = CStr(task(1))

If Dir(savePath) = "" Then GoTo NotGood
If FileLen(savePath) < CAPTCHA_SIZE_HINT Then GoTo NotGood

' A full-size PDF (>= CAPTCHA_SIZE_HINT) at the expected name IS a real
' result - the search succeeded. A leftover _CAPTCHA_UNRESOLVED sidecar
' from an earlier blocked attempt does NOT change that: a late/retry
' browser write can land a clean page after the block was flagged, so
' both files end up on disk. The clean full-size file wins; the flag is
' stale (the safety sweep deletes it). Do NOT report a saved full-size
' result as unresolved - that double-counted it and fired a bogus re-run
' prompt for a search that had actually succeeded.
TaskHasGoodResult = True
Exit Function

NotGood:
TaskHasGoodResult = False
End Function

' Reorders the task queue so the statistically heavier searches
' run first. Based on past run logs the fat-tail outliers are
' almost always "Negative News" queries (20-70 s each, vs 10-15 s
' for a plain name search), so if they land at the end of the
' queue one of the workers is stuck on a 60 s task while the other
' sits idle. Launching them first means by the time we get to the
' tail both workers are on short tasks and finish together.
'
' Stable sort - within a weight bucket, original order is kept so
' per-entity ordering is mostly preserved.
Private Function SortTasksHeaviestFirst(ByVal tasks As Collection) As Collection
Dim N As Long: N = tasks.count
If N <= 1 Then
Set SortTasksHeaviestFirst = tasks
Exit Function
End If

Dim weights() As Long
Dim ordered() As Variant
ReDim weights(1 To N)
ReDim ordered(1 To N)

Dim i As Long, A As Variant, desc As String
For i = 1 To N
A = tasks(i)
desc = LCase$(CStr(A(3)))
Select Case True
Case InStr(desc, "negative news") > 0: weights(i) = 10
Case InStr(desc, "name + address") > 0: weights(i) = 5
Case Else:                              weights(i) = 1
End Select
ordered(i) = A
Next i

' Insertion sort - descending, stable. n is small (under ~30).
Dim j As Long, wKey As Long, aKey As Variant
For i = 2 To N
wKey = weights(i)
aKey = ordered(i)
j = i - 1
Do While j >= 1
If weights(j) < wKey Then
weights(j + 1) = weights(j)
ordered(j + 1) = ordered(j)
j = j - 1
Else
Exit Do
End If
Loop
weights(j + 1) = wKey
ordered(j + 1) = aKey
Next i

Dim result As New Collection
For i = 1 To N
result.Add ordered(i)
Next i

LogStep "SORT: reordered " & N & " tasks (heavy first) - heaviest weight=" & _
weights(1) & ", lightest=" & weights(N)

Set SortTasksHeaviestFirst = result
End Function

' ============================ Search engine ============================
'
' Runs a flat collection of search tasks. Each task is
' Array(query, savePath, pageNum, description, entityLabel).
'
' In headless mode we keep up to MAX_PARALLEL Edge processes busy at
' once across the whole batch - there's no per-entity boundary, so
' workers are never idle while some other entity's tasks are still
' queued. Anything that comes back suspiciously small (probably a
' CAPTCHA) is retried inline, with a growing backoff, until it either
' comes back clean or burns through every allowed attempt - it is
' NEVER accepted under its expected filename while still CAPTCHA-sized.
' Anything that times out gets the same inline-retry treatment.
'
' In interactive mode we just run them one at a time.
'
' Returns the number of tasks that produced a valid, non-CAPTCHA PDF
' on disk.
Private Function RunSearchBatch(tasks As Collection, wsh As Object) As Long
If tasks Is Nothing Then Exit Function
If tasks.count = 0 Then Exit Function

' Reset the per-run CAPTCHA counter here so it's correct regardless
' of which path (headless/interactive) actually runs.
m_captchaFlaggedCount = 0

' Reset circuit-breaker state for this run.
m_captchaWindowStart = Now
m_captchaHitsInWindow = 0
m_globalPauseUntil = #12:00:00 AM#
m_circuitTripCount = 0

' Reset human-solve state for this run.
m_humanSolvePending = False
m_lastHumanSolve = #12:00:00 AM#

If USE_HEADLESS Then
RunSearchBatch = RunSearchBatchParallel(tasks, wsh)
Else
RunSearchBatch = RunSearchBatchSerial(tasks, wsh)
End If
End Function

Private Function RunSearchBatchParallel(tasks As Collection, wsh As Object) As Long
On Error Resume Next

Dim N As Long: N = tasks.count
If N = 0 Then Exit Function

Dim batchStartT As Single: batchStartT = Timer
LogStep "BATCH ALL start, n=" & N & ", maxParallel=" & MAX_PARALLEL

' Make sure the profile pre-warm Edge processes have either
' exited cleanly or been killed before we try to grab the
' worker user-data-dirs. Usually they're long gone by now -
' this is a sub-50 ms no-op in the typical case.
DrainPrewarmAndCleanup

' Cookie warm-up (see WARM_PROFILE_COOKIES_ENABLED / WarmEdgeProfileCookies).
' Logged explicitly here - greppable in osint_debug.log - so runs can be
' correlated against the "captchaFlagged=" count logged at BATCH ALL end
' to see whether this actually moves the needle over time.
' v3.9: warm-up now runs in ALL headless modes (Fast included), and
' it's hidden/headless so no windows pop up.
If WARM_PROFILE_COOKIES_ENABLED Then
WarmEdgeProfileCookies
LogStep "BATCH ALL: cookie warm-up RAN before main pass"
Else
LogStep "BATCH ALL: cookie warm-up SKIPPED (disabled)"
End If

' Single main pass with inline retries. A task whose PDF lands at
' CAPTCHA size is requeued into the same pool (with a growing
' backoff so we don't immediately re-trip the same block) instead
' of running as a separate sequential pass, so the retries overlap
' the tail of the initial dispatch.
Dim mainSizes() As Long
DispatchTaskPass tasks, wsh, "main", mainSizes, False, TASK_TIMEOUT_SEC, CAPTCHA_MAX_ATTEMPTS_MAIN - 1

' Rescue check: did anything still end up missing (including
' anything that got flagged as CAPTCHA-exhausted and renamed out
' of the way)? Give those a further round of attempts at the
' longer retry timeout, with its own backoff-retry budget, before
' finally giving up on them.
Dim rescueList As Collection
Set rescueList = New Collection
Dim i As Long, A As Variant
For i = 1 To N
A = tasks(i)
On Error Resume Next
If Dir(CStr(A(1))) = "" Or FileLen(CStr(A(1))) <= MIN_PDF_SIZE_BYTES Then
rescueList.Add tasks(i)
End If
On Error GoTo 0
Next i

If rescueList.count > 0 And Not bAbort Then
LogStep "RESCUE PASS: " & rescueList.count & " files still missing - final retry"
Application.StatusBar = "OSINT: final rescue pass for " & rescueList.count & " missing file(s)..."
Dim rescueSizes() As Long
DispatchTaskPass rescueList, wsh, "rescue", rescueSizes, True, RETRY_TIMEOUT_SEC, CAPTCHA_MAX_ATTEMPTS_RESCUE - 1
End If

' Final tally is taken straight from disk - that way any retry
' getting the file to a good state still gets counted. Anything
' still CAPTCHA-flagged was renamed out from under this path by
' DispatchTaskPass, so it correctly falls out of this count.
Dim finalSuccess As Long
For i = 1 To N
A = tasks(i)
On Error Resume Next
If FileLen(CStr(A(1))) > MIN_PDF_SIZE_BYTES Then
finalSuccess = finalSuccess + 1
End If
On Error GoTo 0
Next i

LogStep "BATCH ALL end, success=" & finalSuccess & "/" & N & _
" wall=" & CLng(Timer - batchStartT) & "s rescue=" & rescueList.count & _
" captchaFlagged=" & m_captchaFlaggedCount

RunSearchBatchParallel = finalSuccess
End Function

' Runs one pass over a task list using the parallel worker pool.
' Fills finalSizes(1..n) with the size each task's PDF ends up at
' on disk, so the caller can decide if a retry is warranted.
'
' When preserveExisting is True (used on retry / rescue passes) we
' don't delete the existing PDF before relaunching Edge - instead
' we rename it to a .retrybak sidecar. If the retry produces a
' good file we drop the backup; if it fails we put the old file
' back. That way we never regress a file that was already working.
Private Sub DispatchTaskPass(tasks As Collection, _
wsh As Object, _
ByVal passLabel As String, _
ByRef finalSizes() As Long, _
Optional ByVal preserveExisting As Boolean = False, _
Optional ByVal timeoutSeconds As Long = 0, _
Optional ByVal maxRetries As Long = 0)
On Error GoTo Fail

Dim N As Long: N = tasks.count
If N = 0 Then
ReDim finalSizes(1 To 1)   ' guard against uninitialised array in caller
Exit Sub
End If

Randomize   ' seed Rnd() so CAPTCHA backoff jitter isn't identical every run

Dim qry() As String, pth() As String, pg() As Long
Dim dsc() As String, lbl() As String
If timeoutSeconds <= 0 Then timeoutSeconds = TASK_TIMEOUT_SEC

' Per-task bookkeeping:
'   st         - 0 pending / 1 running / 2 done / 3 failed
'   launchT    - when we kicked Edge off
'   firstSeenT - when the PDF first appeared on disk with any bytes
'   wk         - which worker slot the task is using
'   bak        - path of the preserved old PDF, if any
'   lastSize / stableHits - watch the file size settle before we
'                           call the task done, so we never catch
'                           Edge mid-write and mark a partial PDF.
'   durSecArr  - captured duration per task for the end-of-pass stats.
'   nextLaunchT - earliest time this task may be (re)launched. Lets
'                 a CAPTCHA backoff delay just this one task instead
'                 of stalling the whole worker pool.
Dim st() As Long
Dim launchT() As Date
Dim firstSeenT() As Date
Dim wk() As Long
Dim bak() As String
Dim hadBackup() As Boolean
Dim lastSize() As Long
Dim stableHits() As Long
Dim durSecArr() As Long
Dim retriesUsed() As Long      ' how many retry attempts already burned per task
Dim nextLaunchT() As Date
Dim backoffSec As Single
ReDim qry(1 To N), pth(1 To N), pg(1 To N), dsc(1 To N), lbl(1 To N)
ReDim st(1 To N), launchT(1 To N), firstSeenT(1 To N), wk(1 To N)
ReDim bak(1 To N), hadBackup(1 To N), lastSize(1 To N), stableHits(1 To N)
ReDim durSecArr(1 To N), retriesUsed(1 To N), nextLaunchT(1 To N)
ReDim finalSizes(1 To N)

' Track how much of the pass each worker was actually busy -
' lets us spot a lone worker holding up the whole batch at the tail.
Dim workerBusyT() As Date
Dim workerBusySec() As Long
Dim workerCooldownUntil() As Date
ReDim workerBusyT(0 To MAX_PARALLEL - 1)
ReDim workerBusySec(0 To MAX_PARALLEL - 1)
ReDim workerCooldownUntil(0 To MAX_PARALLEL - 1)
Dim queueStartT As Date: queueStartT = Now

Dim i As Long, A As Variant
For i = 1 To N
A = tasks(i)
qry(i) = CStr(A(0))
pth(i) = CStr(A(1))
pg(i) = CLng(A(2))
dsc(i) = CStr(A(3))
If UBound(A) >= 4 Then lbl(i) = CStr(A(4)) Else lbl(i) = ""
Next i

Dim busy() As Boolean
ReDim busy(0 To MAX_PARALLEL - 1)

Dim doneCount As Long, runningCount As Long, successCount As Long
Dim w As Long, pendingFound As Boolean
Dim cmd As String
Dim sz As Long
Dim passStartT As Single: passStartT = Timer
Dim taskDurSec As Long
Dim currLabel As String

' Totals for "how long did the dispatcher sit idle" - lets us see
' whether the sleep/stagger knobs are well-tuned.
Dim launchCount As Long: launchCount = 0
Dim retryCount As Long: retryCount = 0    ' inline-retry firings inside this pass
Dim somethingCompletedThisCycle As Boolean
Dim idleSleepMs As Long: idleSleepMs = 0
Dim taskTimeoutSec As Long
Dim reapCounter As Long          ' counts completions toward the periodic orphan sweep
reapCounter = 0

LogStep "PASS [" & passLabel & "] start, n=" & N & ", maxRetries=" & maxRetries

Do
somethingCompletedThisCycle = False

' Human-solve: if a CAPTCHA flagged this, pause here, let the analyst
' solve it in visible windows, then resume. Requeues in-flight tasks
' (their headless Edge gets killed to free the profiles) so they
' re-run with the freshly-solved exemption cookies.
If m_humanSolvePending Then
DoHumanCaptchaSolve wsh, N, st, wk, busy, runningCount
End If

' Fill any free worker slots with pending tasks that are eligible
' to launch right now (a task backing off after a CAPTCHA hit
' won't be picked up again until its nextLaunchT has passed - it
' just sits pending while other tasks keep the workers busy). The
' launch stagger only fires for the initial ramp-up: once every
' worker has been launched once we stop staggering, so a
' mid-batch completion gets an instant relaunch.
' Pool-wide brake: while the circuit is open, no worker takes a
' new task. Anything already in flight is left to finish.
If CircuitIsOpen() Then
Application.StatusBar = "OSINT [" & passLabel & "]: rate-limited - " & _
"pausing " & CLng(DateDiff("s", Now, m_globalPauseUntil)) & "s " & _
"before resuming (" & doneCount & "/" & N & " saved)"
Else

For w = 0 To MAX_PARALLEL - 1
If bAbort Then Exit For
If Not busy(w) And Now >= workerCooldownUntil(w) Then
pendingFound = False
For i = 1 To N
If bAbort Then Exit For
If st(i) = 0 And Now >= nextLaunchT(i) Then
On Error Resume Next
If preserveExisting Then
bak(i) = pth(i) & ".retrybak"
If Dir(bak(i)) <> "" Then Kill bak(i)
If Dir(pth(i)) <> "" Then
Name pth(i) As bak(i)
hadBackup(i) = True
End If
Else
If Dir(pth(i)) <> "" Then Kill pth(i)
End If
On Error GoTo Fail

cmd = BuildHeadlessEdgeCmd(qry(i), pth(i), pg(i), w)
wsh.Run cmd, 0, False

st(i) = 1
launchT(i) = Now
workerBusyT(w) = Now         ' worker utilisation clock starts
wk(i) = w
busy(w) = True
runningCount = runningCount + 1
pendingFound = True
launchCount = launchCount + 1

LogStep "  TASK launch [" & i & "/" & N & "] w" & w & _
" running=" & runningCount & _
" [" & lbl(i) & "] " & dsc(i)

' Stagger only during initial fill, not mid-batch.
If launchCount <= MAX_PARALLEL Then
SmartWait t(LAUNCH_STAGGER_SEC), True
End If
Exit For
End If
Next i
If Not pendingFound Then Exit For
End If
Next w

End If   ' circuit closed - normal launching

' Poll for completions or timeouts
For i = 1 To N
If st(i) = 1 Then
sz = 0
If Dir(pth(i)) <> "" Then
On Error Resume Next
sz = FileLen(pth(i))
On Error GoTo Fail
End If

' Remember the moment the PDF first shows up with any
' bytes - that's our proxy for Edge having finished
' booting and navigated to the SERP.
If sz > 0 And firstSeenT(i) = #12:00:00 AM# Then
firstSeenT(i) = Now
End If

If sz > MIN_PDF_SIZE_BYTES Then
If sz = lastSize(i) Then
stableHits(i) = stableHits(i) + 1
Else
lastSize(i) = sz
stableHits(i) = 0
End If

If stableHits(i) < 1 Then GoTo ContinuePollingTask

' Tiny PDF = almost certainly a CAPTCHA / interstitial.
' We never accept this under the expected filename:
' either we still have attempts left and requeue with a
' growing backoff (so we don't hammer Google right back
' into the same block), or we've burned every attempt and
' flag the file for a human instead of quietly reporting
' it as a good result.
If sz < CAPTCHA_SIZE_HINT Then
' Size says "maybe CAPTCHA" - now CONFIRM by reading the page
' content (v4.1). Only a content-confirmed CAPTCHA proceeds to
' retry / flag / solve; a small-but-valid results page falls
' through to the success path below.
If IsCaptchaContent(qry(i), pg(i), wk(i), wsh) Then
' Count this against the pool-wide brake before deciding what
' to do with the individual task - a cluster of these means we
' should stop launching entirely, not just delay this one.
RegisterCaptchaHit

' Human-solve: flag that the analyst should solve a CAPTCHA. The
' dispatch loop picks this up at the next cycle boundary. Cooldown
' stops it re-prompting while a just-solved profile settles. We only
' arm once this task has already been retried HUMAN_SOLVE_RETRY_THRESHOLD
' times and is STILL blocked - a first-hit CAPTCHA that clears on retry
' never opens the disruptive visible window.
If CAPTCHA_HUMAN_SOLVE_ENABLED And Not m_humanSolvePending _
And retriesUsed(i) >= HUMAN_SOLVE_RETRY_THRESHOLD Then
If DateDiff("s", m_lastHumanSolve, Now) > HUMAN_SOLVE_COOLDOWN_SEC Then
m_humanSolvePending = True
LogStep "  HUMAN-SOLVE armed after " & retriesUsed(i) & " retr(ies) (" & _
IIf(CAPTCHA_CONTENT_DETECT_ENABLED, "content-confirmed", "size-based") & _
" CAPTCHA, PDF " & sz & "B)"
End If
End If

If retriesUsed(i) < maxRetries Then
On Error Resume Next
If Dir(pth(i)) <> "" Then Kill pth(i)
On Error GoTo Fail

workerBusySec(wk(i)) = workerBusySec(wk(i)) + _
DateDiff("s", workerBusyT(wk(i)), Now)
st(i) = 0                       ' back to pending
retriesUsed(i) = retriesUsed(i) + 1
runningCount = runningCount - 1
busy(wk(i)) = False
SetWorkerCooldown wk(i), workerCooldownUntil
somethingCompletedThisCycle = True
retryCount = retryCount + 1
firstSeenT(i) = #12:00:00 AM#   ' reset bookkeeping
lastSize(i) = 0
stableHits(i) = 0

' Fast mode (protections off) skips the backoff wait entirely and
' relaunches immediately; Optimised/Visible use the growing backoff.
If m_protectionsOn Then
backoffSec = CAPTCHA_BACKOFF_BASE_SEC * (2 ^ (retriesUsed(i) - 1))
If backoffSec > CAPTCHA_BACKOFF_MAX_SEC Then backoffSec = CAPTCHA_BACKOFF_MAX_SEC
backoffSec = backoffSec + CSng(Rnd() * 2#)   ' jitter so tasks don't relaunch in lockstep
Else
backoffSec = 0
End If
nextLaunchT(i) = DateAdd("s", CDbl(backoffSec), Now)

LogStep "  TASK retry [" & i & "/" & N & "] CAPTCHA-size " & sz & _
"B - requeued (attempt " & (retriesUsed(i) + 1) & "/" & (maxRetries + 1) & _
"), backing off " & Format(backoffSec, "0.0") & "s before relaunch"
Else
' Out of attempts. Very likely a persistent block (it
' could occasionally be a genuinely thin result page,
' but size alone can't tell us that, so we err toward
' flagging rather than silently trusting it). Rename
' it out of the way so it can never be mistaken for a
' clean result, and do NOT count it as a success.
workerBusySec(wk(i)) = workerBusySec(wk(i)) + _
DateDiff("s", workerBusyT(wk(i)), Now)

Dim flagPath As String
flagPath = FlagPathWithSuffix(pth(i), CAPTCHA_FLAG_SUFFIX)
On Error Resume Next
If Dir(flagPath) <> "" Then Kill flagPath
If Dir(pth(i)) <> "" Then Name pth(i) As flagPath
On Error GoTo Fail

If hadBackup(i) Then
On Error Resume Next
If Dir(bak(i)) <> "" Then Name bak(i) As pth(i)
On Error GoTo Fail
End If

st(i) = 3
finalSizes(i) = 0
doneCount = doneCount + 1
runningCount = runningCount - 1
busy(wk(i)) = False
SetWorkerCooldown wk(i), workerCooldownUntil
somethingCompletedThisCycle = True
m_captchaFlaggedCount = m_captchaFlaggedCount + 1

LogStep "  TASK CAPTCHA-EXHAUSTED [" & i & "/" & N & "] w" & wk(i) & _
" after " & (retriesUsed(i) + 1) & " attempt(s), last size=" & sz & _
"B - flagged as '" & flagPath & "' (NOT counted as success) - [" & _
lbl(i) & "] " & dsc(i)
End If
GoTo ContinuePollingTask
Else
' Content shows a real results page that just happens to be small -
' accept it instead of flagging/retrying. Falls through to success.
LogStep "  CONTENT-DETECT [" & i & "/" & N & "] small PDF " & sz & _
"B is a VALID results page - accepting (not a CAPTCHA)"
End If
End If

st(i) = 2
finalSizes(i) = sz
successCount = successCount + 1
doneCount = doneCount + 1
runningCount = runningCount - 1
busy(wk(i)) = False
SetWorkerCooldown wk(i), workerCooldownUntil
somethingCompletedThisCycle = True

' Reap orphan Edge processes PERIODICALLY (every REAP_EVERY_N_TASKS
' completions) rather than after every single task. A headless launch
' leaves ~12-23 child processes that don't all self-exit; on a 2-vCPU VDI
' they can pile up and cause "NO FILE EVER" hangs. But the WMI kill costs
' ~1-2s and blocks the single-threaded dispatcher, so doing it every task
' added ~40-55s of overhead per run. With the launch pacing letting Edge
' self-clean between launches, a periodic sweep contains the pileup at a
' fraction of the cost. (Stuck tasks are still reaped immediately on
' timeout, below - that path is unconditional.)
reapCounter = reapCounter + 1
If reapCounter >= REAP_EVERY_N_TASKS Then
KillEdgeWorkerProcesses wk(i)
reapCounter = 0
End If

taskDurSec = DateDiff("s", launchT(i), Now)
durSecArr(i) = taskDurSec
workerBusySec(wk(i)) = workerBusySec(wk(i)) + _
DateDiff("s", workerBusyT(wk(i)), Now)
currLabel = lbl(i)

' Slice total duration into wait / Edge-start / render
' so the log tells us which phase ate the time.
Dim edgeStartSec As Long, renderSec As Long, waitSec As Long
waitSec = DateDiff("s", queueStartT, launchT(i))
If firstSeenT(i) <> #12:00:00 AM# Then
edgeStartSec = DateDiff("s", launchT(i), firstSeenT(i))
renderSec = DateDiff("s", firstSeenT(i), Now)
Else
edgeStartSec = -1
renderSec = -1
End If

LogStep "  TASK done  [" & i & "/" & N & "] w" & wk(i) & _
" wait=" & waitSec & "s edgeStart=" & edgeStartSec & _
"s render=" & renderSec & "s dur=" & taskDurSec & _
"s size=" & sz & "B running=" & runningCount

If hadBackup(i) Then
On Error Resume Next
If Dir(bak(i)) <> "" Then Kill bak(i)
On Error GoTo Fail
End If

Else
' Adaptive timeout: a task that's already been retried
' once gets the longer RETRY_TIMEOUT_SEC because
' those tend to be the genuinely-slow Google queries.
If retriesUsed(i) > 0 Then
taskTimeoutSec = RETRY_TIMEOUT_SEC
Else
taskTimeoutSec = timeoutSeconds
End If

If DateDiff("s", launchT(i), Now) > taskTimeoutSec Then
' Always free the worker first - the orphan
' Edge process needs killing whether we retry
' or give up.
KillEdgeWorkerProcesses wk(i)
workerBusySec(wk(i)) = workerBusySec(wk(i)) + _
DateDiff("s", workerBusyT(wk(i)), Now)
durSecArr(i) = DateDiff("s", launchT(i), Now)

Dim toFileState As String
If firstSeenT(i) = #12:00:00 AM# Then
toFileState = "NO FILE EVER"
Else
toFileState = "file seen after " & _
DateDiff("s", launchT(i), firstSeenT(i)) & "s"
End If

If retriesUsed(i) < maxRetries Then
' Inline retry: drop straight back into pending,
' the next free worker will pick it up.
On Error Resume Next
If Not preserveExisting Then
If Dir(pth(i)) <> "" Then Kill pth(i)
End If
On Error GoTo Fail

st(i) = 0
retriesUsed(i) = retriesUsed(i) + 1
runningCount = runningCount - 1
busy(wk(i)) = False
SetWorkerCooldown wk(i), workerCooldownUntil
somethingCompletedThisCycle = True
retryCount = retryCount + 1
firstSeenT(i) = #12:00:00 AM#
lastSize(i) = 0
stableHits(i) = 0
LogStep "  TASK retry [" & i & "/" & N & "] TIMEOUT after " & _
taskTimeoutSec & "s (" & toFileState & _
") - requeued (attempt " & (retriesUsed(i) + 1) & ")"
Else
' Out of retries - mark as final failure.
st(i) = 3
finalSizes(i) = 0
doneCount = doneCount + 1
runningCount = runningCount - 1
busy(wk(i)) = False
SetWorkerCooldown wk(i), workerCooldownUntil
somethingCompletedThisCycle = True
If hadBackup(i) Then
On Error Resume Next
If Dir(pth(i)) <> "" Then Kill pth(i)
If Dir(bak(i)) <> "" Then Name bak(i) As pth(i)
If Dir(pth(i)) <> "" Then finalSizes(i) = FileLen(pth(i))
On Error GoTo Fail
End If
LogStep "  TASK TIMEOUT [" & i & "/" & N & "] w" & wk(i) & _
" after " & taskTimeoutSec & "s (" & toFileState & _
") - [" & lbl(i) & "] " & dsc(i)
End If
End If
End If
ContinuePollingTask:
End If
Next i

' Status bar: show pass + most-recent entity + progress
If currLabel <> "" Then
Application.StatusBar = "OSINT [" & passLabel & "]: " & currLabel & " - " & _
doneCount & "/" & N & " saved, " & runningCount & " running"
Else
Application.StatusBar = "OSINT [" & passLabel & "]: " & _
doneCount & "/" & N & " saved, " & runningCount & " running"
End If

If bAbort Then Exit Do

' If a task just finished, loop back immediately so the launch
' phase can push work onto the now-free worker without waiting
' out a poll cycle. Only sleep when nothing changed this round.
If Not somethingCompletedThisCycle Then
SmartWait t(POLL_INTERVAL_SEC)
idleSleepMs = idleSleepMs + CLng(t(POLL_INTERVAL_SEC) * 1000)
End If
Loop While doneCount < N

' End-of-pass stats: how were the task durations distributed,
' and how saturated was each worker?
Dim passWallSec As Long: passWallSec = CLng(Timer - passStartT)
Dim minD As Long, maxD As Long, sumD As Long, countD As Long
Dim b0_10 As Long, b10_20 As Long, b20_30 As Long, b30_45 As Long, b45_60 As Long, b60p As Long
minD = 99999
For i = 1 To N
If durSecArr(i) > 0 Then
countD = countD + 1
sumD = sumD + durSecArr(i)
If durSecArr(i) < minD Then minD = durSecArr(i)
If durSecArr(i) > maxD Then maxD = durSecArr(i)
Select Case durSecArr(i)
Case Is < 10:  b0_10 = b0_10 + 1
Case Is < 20:  b10_20 = b10_20 + 1
Case Is < 30:  b20_30 = b20_30 + 1
Case Is < 45:  b30_45 = b30_45 + 1
Case Is < 60:  b45_60 = b45_60 + 1
Case Else:     b60p = b60p + 1
End Select
End If
Next i
If countD = 0 Then minD = 0

LogStep "PASS [" & passLabel & "] end, success=" & successCount & "/" & N & _
" wall=" & passWallSec & "s"
If countD > 0 Then
LogStep "  DIST [" & passLabel & "] min=" & minD & "s avg=" & (sumD \ countD) & _
"s max=" & maxD & "s | <10s=" & b0_10 & " <20s=" & b10_20 & _
" <30s=" & b20_30 & " <45s=" & b30_45 & " <60s=" & b45_60 & _
" 60+s=" & b60p
End If
For w = 0 To MAX_PARALLEL - 1
Dim utilPct As Long
If passWallSec > 0 Then utilPct = (workerBusySec(w) * 100) \ passWallSec
LogStep "  UTIL w" & w & " busy=" & workerBusySec(w) & "s/" & passWallSec & _
"s = " & utilPct & "%"
' Final safety sweep: guarantee no worker Edge process survives the
' end of this pass, even if one slipped past the per-task reap above.
KillEdgeWorkerProcesses w
Next w
LogStep "  DISP [" & passLabel & "] launches=" & launchCount & _
" inlineRetries=" & retryCount & _
" dispatcher-sleep=" & (idleSleepMs \ 1000) & "." & _
Format((idleSleepMs Mod 1000) \ 100, "0") & "s"
Exit Sub
Fail:
LogStep "PASS [" & passLabel & "] FAIL err=" & Err.Number & " desc=" & Err.Description
End Sub

' Sets a short random cooldown (m_launchDelayMinSec to
' m_launchDelayMaxSec, set per search mode by ApplySearchMode)
' before worker w is allowed to grab its next task - so each
' worker's own request cadence has natural human-ish gaps instead
' of firing back-to-back the instant a task completes. No-op if
' RANDOM_LAUNCH_DELAY_ENABLED is False.
' ---- Global CAPTCHA circuit-breaker (v3.8) ----
' Records one block. If CAPTCHA_CIRCUIT_THRESHOLD of them land
' inside CAPTCHA_CIRCUIT_WINDOW_SEC, the whole pool stops launching
' for CAPTCHA_CIRCUIT_PAUSE_SEC. Counting is windowed, so blocks
' scattered thinly across a long batch never trip it - only a burst
' does, which is exactly the shape of being actively rate-limited.
Private Sub RegisterCaptchaHit()
On Error Resume Next
If Not (CAPTCHA_CIRCUIT_ENABLED And m_protectionsOn) Then Exit Sub

' Roll the window forward if the last one has aged out.
If DateDiff("s", m_captchaWindowStart, Now) > CAPTCHA_CIRCUIT_WINDOW_SEC Then
m_captchaWindowStart = Now
m_captchaHitsInWindow = 0
End If

m_captchaHitsInWindow = m_captchaHitsInWindow + 1
If m_captchaHitsInWindow < CAPTCHA_CIRCUIT_THRESHOLD Then Exit Sub

' Already paused from an earlier trip - don't stack another.
If Now < m_globalPauseUntil Then Exit Sub

m_circuitTripCount = m_circuitTripCount + 1

If m_circuitTripCount > CAPTCHA_CIRCUIT_MAX_TRIPS Then
' Repeated trips mean pausing isn't clearing it - this is a
' sustained block, not a transient throttle. Stop burning time
' on it; remaining tasks will exhaust their retries and get
' flagged for a human, which is the honest outcome.
LogStep "CIRCUIT: tripped " & (m_circuitTripCount - 1) & " time(s) already and " & _
"blocks are still coming - giving up on pausing. Remaining searches " & _
"will be flagged rather than retried indefinitely."
Exit Sub
End If

m_globalPauseUntil = DateAdd("s", CAPTCHA_CIRCUIT_PAUSE_SEC, Now)
m_captchaHitsInWindow = 0
m_captchaWindowStart = Now

LogStep "CIRCUIT OPEN (trip " & m_circuitTripCount & "/" & CAPTCHA_CIRCUIT_MAX_TRIPS & _
"): " & CAPTCHA_CIRCUIT_THRESHOLD & " blocks within " & CAPTCHA_CIRCUIT_WINDOW_SEC & _
"s - pausing ALL launches for " & CAPTCHA_CIRCUIT_PAUSE_SEC & "s to let the " & _
"rate limit decay. In-flight tasks continue."
On Error GoTo 0
End Sub

' True while the pool is paused. Logs once on close so the resume
' point is visible in the log without spamming every poll cycle.
Private Function CircuitIsOpen() As Boolean
On Error Resume Next
If Not (CAPTCHA_CIRCUIT_ENABLED And m_protectionsOn) Then Exit Function
If m_globalPauseUntil = #12:00:00 AM# Then Exit Function

If Now < m_globalPauseUntil Then
CircuitIsOpen = True
Else
' Just expired - reset and note it once.
m_globalPauseUntil = #12:00:00 AM#
m_captchaWindowStart = Now
m_captchaHitsInWindow = 0
LogStep "CIRCUIT CLOSED - pause elapsed, resuming launches"
CircuitIsOpen = False
End If
On Error GoTo 0
End Function

' ---- Human-in-the-loop CAPTCHA solve (v4.0) ----
' Pauses the batch, opens one VISIBLE Edge per worker profile at a
' Google search (so the challenge appears), lets the analyst solve
' "I'm not a robot", then resumes the headless searches with the
' exemption cookies the solve just stored in each profile. Called
' from the dispatch loop; requeues any in-flight tasks so they re-run
' cleanly afterwards.
Private Sub DoHumanCaptchaSolve(wsh As Object, ByVal N As Long, _
ByRef st() As Long, ByRef wk() As Long, ByRef busy() As Boolean, _
ByRef runningCount As Long)
On Error Resume Next

m_humanSolvePending = False   ' consume the flag up front
LogStep "HUMAN-SOLVE: CAPTCHA hit - pausing for analyst to solve"

Dim w As Long, i As Long

' 1. Kill the headless workers so their profiles are free for a
' visible Edge (can't share a --user-data-dir between two processes).
For w = 0 To MAX_PARALLEL - 1
KillEdgeWorkerProcesses w
Next w
SmartWait 1, True   ' let them die + release the profile locks

' 2. Requeue anything that was in flight - it was about to fail on
' the CAPTCHA anyway; it'll re-run after the solve.
For i = 1 To N
If st(i) = 1 Then
st(i) = 0
If wk(i) >= 0 And wk(i) <= MAX_PARALLEL - 1 Then busy(wk(i)) = False
runningCount = runningCount - 1
End If
Next i

' 3. Open one VISIBLE Edge per worker profile at a Google search so
' the challenge (if any) is shown for the analyst.
Dim edgeExe As String, cmd As String, profilePath As String
edgeExe = GetEdgePath()
For w = 0 To MAX_PARALLEL - 1
profilePath = GetWorkerProfilePath(w)
cmd = """" & edgeExe & """" & _
" --user-data-dir=""" & profilePath & """" & _
" --no-first-run --no-default-browser-check --disable-sync" & _
" --window-size=1000,800" & _
" --window-position=" & (w * 90) & ",60" & _
" ""https://www.google.com/search?q=test&num=100&hl=en"""
wsh.Run cmd, 1, False   ' 1 = visible window
Next w
SmartWait 2, True   ' let the windows open

' 4. Wait for the analyst. Human-paced - they solve in each window,
' let the results load, then click OK.
Application.StatusBar = "OSINT: solve the CAPTCHA in the browser window(s), then click OK"
TopMostMsgBox _
"Google is challenging the searches with a CAPTCHA." & vbCrLf & vbCrLf & _
MAX_PARALLEL & " browser window(s) just opened. In EACH one, complete the " & _
"'I'm not a robot' check and let the results page load - then click OK here." & vbCrLf & vbCrLf & _
"(If a window already shows normal results with no challenge, just click OK.)", _
"Solve CAPTCHA to Continue", MB_OK Or MB_ICONWARNING Or MB_TOPMOST

' 5. Give Edge a moment to write the exemption cookie to disk BEFORE
' we close it - a hard kill before the flush would lose the solve.
SmartWait CSng(HUMAN_SOLVE_FLUSH_SEC), True

' 6. Close the visible windows so the profiles are free for headless.
For w = 0 To MAX_PARALLEL - 1
KillEdgeWorkerProcesses w
Next w
SmartWait 1, True

' 7. Reset gates so the batch resumes right away.
m_lastHumanSolve = Now
m_globalPauseUntil = #12:00:00 AM#
m_captchaHitsInWindow = 0
m_captchaWindowStart = Now

LogStep "HUMAN-SOLVE: analyst confirmed - resuming (in-flight tasks requeued)"
Application.StatusBar = "OSINT: resuming searches after CAPTCHA solve..."
On Error GoTo 0
End Sub

Private Sub SetWorkerCooldown(ByVal w As Long, ByRef workerCooldownUntil() As Date)
On Error Resume Next
If Not RANDOM_LAUNCH_DELAY_ENABLED Then Exit Sub
Dim delaySec As Single
delaySec = m_launchDelayMinSec + Rnd() * (m_launchDelayMaxSec - m_launchDelayMinSec)
workerCooldownUntil(w) = DateAdd("s", CDbl(delaySec), Now)
LogStep "  WORKER w" & w & " [" & modeName(m_searchMode) & "] cooldown " & _
Format(delaySec, "0.0") & "s before next task"
On Error GoTo 0
End Sub

Private Function RunSearchBatchSerial(tasks As Collection, wsh As Object) As Long
Dim N As Long: N = tasks.count
Dim i As Long, successCount As Long
Dim A As Variant
Dim currLabel As String, currDesc As String

For i = 1 To N
A = tasks(i)
currDesc = CStr(A(3))
If UBound(A) >= 4 Then currLabel = CStr(A(4)) Else currLabel = "Search"

SetSearchStatus currLabel, currDesc & " (" & i & "/" & N & ")"
If PrintSearchToPDF(CStr(A(0)), CStr(A(1)), wsh, CLng(A(2))) Then
successCount = successCount + 1
End If
If bAbort Then Exit For
Next i

RunSearchBatchSerial = successCount
End Function

' ---- Single-task search helpers ----
' Used by the interactive fallback path and anywhere else we just
' want to run one search synchronously. The parallel dispatcher
' above drives Edge directly.
Function PrintSearchToPDF(query As String, savePath As String, wsh As Object, _
Optional PageNum As Long = 1) As Boolean
If USE_HEADLESS Then
PrintSearchToPDF = PrintSearchToPDF_Headless(query, savePath, PageNum, wsh, 0)
Else
PrintSearchToPDF = PrintSearchToPDF_Interactive(query, savePath, wsh, PageNum)
End If
End Function

' Synchronous single-shot headless run. Only used by the fallback
' paths - the parallel dispatcher talks to Edge directly so it
' can keep several workers busy at once.
Private Function PrintSearchToPDF_Headless(query As String, _
savePath As String, _
ByVal PageNum As Long, _
wsh As Object, _
ByVal workerIdx As Long) As Boolean
On Error GoTo FuncFail

Dim cmd As String

On Error Resume Next
If Dir(savePath) <> "" Then Kill savePath
On Error GoTo FuncFail

cmd = BuildHeadlessEdgeCmd(query, savePath, PageNum, workerIdx)

' Synchronous when called from the single-task path.
' Parallel dispatch calls BuildHeadlessEdgeCmd directly and uses Run False.
wsh.Run cmd, 0, True

Dim resultSize As Long
If Dir(savePath) <> "" Then resultSize = FileLen(savePath)

If resultSize < CAPTCHA_SIZE_HINT And resultSize > MIN_PDF_SIZE_BYTES Then
Debug.Print "WARN: small headless PDF (" & resultSize & " B) - possible CAPTCHA: " & savePath
End If

PrintSearchToPDF_Headless = (resultSize > MIN_PDF_SIZE_BYTES)
Exit Function
FuncFail:
PrintSearchToPDF_Headless = False
End Function

' Builds the command line for one headless Edge instance.
' workerIdx picks a unique --user-data-dir so parallel Edge
' processes each have their own profile and don't fight over
' a single profile lock.
Private Function BuildHeadlessEdgeCmd(ByVal query As String, _
ByVal savePath As String, _
ByVal PageNum As Long, _
ByVal workerIdx As Long) As String
Dim baseURL As String, browserExe As String, profilePath As String
Dim s As String

baseURL = "https://www.google.com/search?q=" & URLEncode(query) & "&num=100&hl=en"
If PageNum > 1 Then baseURL = baseURL & "&start=" & ((PageNum - 1) * 10)

browserExe = GetEdgePath()
profilePath = GetWorkerProfilePath(workerIdx)

' Flag set tuned for a no-GPU VDI. Headers/footers are left enabled
' on purpose so the PDF shows the date, URL and page number for audit.
'
' v3.8 - the --user-agent override was REMOVED here, deliberately.
' It was pinned to "Chrome/131 ... Edg/131", which caused two
' problems that made us easier to spot, not harder:
'
'   1. Stale version. Edge is far past 131 by now, so we were
'      announcing a browser build that's long out of support -
'      unusual enough on its own to be worth scoring against.
'   2. Client Hints mismatch. Chromium derives its Sec-CH-UA
'      headers from the REAL build and --user-agent does not
'      update them. So the UA header claimed 131 while the
'      Sec-CH-UA headers sent the actual installed version on the
'      very same request - a flat self-contradiction that's
'      trivial to check for and that no genuine browser produces.
'
' Sending no override at all means Edge presents its real, current,
' internally consistent identity. That is strictly better than any
' hand-maintained UA string, and it can never go stale.
s = """" & browserExe & """"
s = s & " --headless=new"
s = s & " --disable-gpu"
' Spawn-reduction flags for the 2-vCPU VDI. A headless launch normally
' forks 12-23 child processes (zygote helpers, crashpad handler, extra
' renderers) that pile up and starve the CPU - the root cause of the slow
' edgeStart times and the "NO FILE EVER" launch hangs. These cut the
' process footprint at the source; none affect the rendered PDF:
'   --no-zygote            : no fork-a-helper-per-tab model
'   --disable-crash-reporter: no crashpad_handler processes
'   --renderer-process-limit=1 : one renderer (plenty for one search page)
s = s & " --no-zygote"
s = s & " --disable-crash-reporter"
s = s & " --renderer-process-limit=1"
s = s & " --disable-blink-features=AutomationControlled"
s = s & " --disable-extensions"
s = s & " --window-size=1920,1080"
' --lang drives navigator.languages; --accept-lang drives the
' Accept-Language header. Setting only the latter (as before) left
' the two disagreeing, which is itself a checked signal.
s = s & " --lang=en-US"
s = s & " --accept-lang=""en-US,en;q=0.9"""
s = s & " --user-data-dir=""" & profilePath & """"
s = s & " --no-first-run"
s = s & " --no-default-browser-check"
s = s & " --disable-logging --log-level=3"
s = s & " --disable-background-networking"
s = s & " --disable-sync"
s = s & " --disable-default-apps"
s = s & " --disable-component-update"
s = s & " --disable-client-side-phishing-detection"
s = s & " --metrics-recording-only"
s = s & " --mute-audio"
s = s & " --hide-scrollbars"
s = s & " --no-pings"
s = s & " --print-to-pdf=""" & savePath & """"
s = s & " """ & baseURL & """"

BuildHeadlessEdgeCmd = s
End Function

' Profile dir name for a worker slot. Doubles as the unique string
' the orphan-killer matches on in a process command line.
Private Function GetWorkerProfileName(ByVal workerIdx As Long) As String
GetWorkerProfileName = "OSINT_EdgeWorker_" & workerIdx
End Function

Private Function GetWorkerProfilePath(ByVal workerIdx As Long) As String
GetWorkerProfilePath = Environ("LOCALAPPDATA") & "\" & GetWorkerProfileName(workerIdx)
End Function

' Inserts a suffix just before the file extension so a flagged PDF
' ("..._Google.pdf" -> "..._Google_CAPTCHA_UNRESOLVED.pdf") is
' visually unmistakable in the case folder and can never collide
' with the plain expected filename a downstream success-check
' looks for.
Private Function FlagPathWithSuffix(ByVal originalPath As String, ByVal suffix As String) As String
Dim dotPos As Long
dotPos = InStrRev(originalPath, ".")
If dotPos > 0 Then
FlagPathWithSuffix = Left$(originalPath, dotPos - 1) & suffix & Mid$(originalPath, dotPos)
Else
FlagPathWithSuffix = originalPath & suffix
End If
End Function

' Kill any msedge.exe whose command line references this worker's
' profile dir (OSINT_EdgeWorker<idx>). Called after a timeout so
' the worker's user-data-dir lock
' is released and the next task can re-use the slot cleanly. Uses
' WMI, which is available on locked-down VDIs without admin rights;
' if WMI is blocked we just fall through.
'
Private Sub KillEdgeWorkerProcesses(ByVal workerIdx As Long)
On Error Resume Next
Dim wmi As Object, procs As Object, p As Object
Dim profileName As String, cmdLine As String
Dim killed As Long

profileName = GetWorkerProfileName(workerIdx)

Set wmi = GetObject("winmgmts:\\.\root\cimv2")
If wmi Is Nothing Then Exit Sub

Set procs = wmi.ExecQuery( _
"SELECT ProcessId, CommandLine FROM Win32_Process " & _
"WHERE Name = 'msedge.exe'")
If procs Is Nothing Then Exit Sub

killed = 0
For Each p In procs
cmdLine = ""
cmdLine = CStr(p.CommandLine)
If Len(cmdLine) > 0 And _
InStr(1, cmdLine, profileName, vbTextCompare) > 0 Then
p.Terminate
killed = killed + 1
End If
Next p

If killed > 0 Then
LogStep "  KILL w" & workerIdx & " - terminated " & killed & _
" orphan msedge.exe process(es)"
End If
End Sub

Private Function GetEdgePath() As String
Dim candidates As Variant, p As Variant
candidates = Array( _
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", _
"C:\Program Files\Microsoft\Edge\Application\msedge.exe", _
Environ("LOCALAPPDATA") & "\Microsoft\Edge\Application\msedge.exe")

For Each p In candidates
If Dir(CStr(p)) <> "" Then
GetEdgePath = CStr(p)
Exit Function
End If
Next p
GetEdgePath = "msedge.exe"
End Function

' ---- Content-based CAPTCHA confirmation (v4.1) ----
' Re-fetches the same query HEADLESS with --dump-dom (DOM to a text
' file) and reads it to decide whether a small page is a real CAPTCHA
' or a valid results page. Returns TRUE if it's (or is assumed to be)
' a CAPTCHA. Only called on pages already below CAPTCHA_SIZE_HINT, so
' it's rare - not an every-search cost.
'
' Safety-first: TRUE unless we can positively read a results page.
' A read failure/timeout defaults to TRUE, so a genuine block is
' never accepted as a clean result.
Private Function IsCaptchaContent(ByVal query As String, ByVal PageNum As Long, _
ByVal workerIdx As Long, wsh As Object) As Boolean
IsCaptchaContent = True   ' safe default

If Not CAPTCHA_CONTENT_DETECT_ENABLED Then Exit Function

On Error GoTo Fallback

Dim edgeExe As String, profilePath As String, outPath As String
Dim baseURL As String, cmd As String, Q As String
Q = Chr$(34)   ' one double-quote

edgeExe = GetEdgePath()
If Len(edgeExe) = 0 Then Exit Function
profilePath = GetWorkerProfilePath(workerIdx)
outPath = Environ$("TEMP") & "\osint_capdetect_w" & workerIdx & ".html"

' Free this worker's profile + clear any old probe output.
KillEdgeWorkerProcesses workerIdx
On Error Resume Next
If Dir(outPath) <> "" Then Kill outPath
On Error GoTo Fallback

baseURL = "https://www.google.com/search?q=" & URLEncode(query) & "&num=100&hl=en"
If PageNum > 1 Then baseURL = baseURL & "&start=" & ((PageNum - 1) * 10)

' Headless --dump-dom, stdout redirected to a file via cmd. Hidden.
cmd = "cmd.exe /c " & Q & Q & edgeExe & Q & _
" --headless=new --disable-gpu" & _
" --user-data-dir=" & Q & profilePath & Q & _
" --no-first-run --no-default-browser-check --disable-extensions" & _
" --disable-sync --dump-dom " & Q & baseURL & Q & _
" > " & Q & outPath & Q & " 2>nul" & Q
wsh.Run cmd, 0, False   ' 0 = hidden, don't wait (we poll below)

' Poll for the DOM file to be written (the redirect flushes when Edge
' exits after dumping), bounded by CAPTCHA_DETECT_TIMEOUT_SEC.
Dim waited As Single
waited = 0
Do
SmartWait 0.4, True
waited = waited + 0.4
If Dir(outPath) <> "" Then
If FileLen(outPath) > 200 Then Exit Do
End If
If waited >= CAPTCHA_DETECT_TIMEOUT_SEC Then Exit Do
Loop

KillEdgeWorkerProcesses workerIdx   ' clean the probe process

If Dir(outPath) = "" Then
LogStep "  CONTENT-DETECT w" & workerIdx & ": no DOM captured - assuming CAPTCHA (safe default)"
IsCaptchaContent = True
GoTo CleanUp
End If
If FileLen(outPath) < 100 Then
LogStep "  CONTENT-DETECT w" & workerIdx & ": empty DOM - assuming CAPTCHA (safe default)"
IsCaptchaContent = True
GoTo CleanUp
End If

Dim html As String
html = LCase$(ReadTextFile(outPath))

Dim isCap As Boolean, isResults As Boolean
isCap = (InStr(html, "unusual traffic") > 0) _
Or (InStr(html, "/sorry/") > 0) _
Or (InStr(html, "g-recaptcha") > 0) _
Or (InStr(html, "recaptcha") > 0) _
Or (InStr(html, "id=" & Q & "recaptcha" & Q) > 0) _
Or (InStr(html, "not a robot") > 0)

isResults = (InStr(html, "result-stats") > 0) _
Or (InStr(html, "id=" & Q & "search" & Q) > 0) _
Or (InStr(html, "id=" & Q & "rso" & Q) > 0) _
Or (InStr(html, "id=" & Q & "result-stats" & Q) > 0)

If isCap Then
IsCaptchaContent = True
LogStep "  CONTENT-DETECT w" & workerIdx & ": CAPTCHA markers found -> CAPTCHA"
ElseIf isResults Then
IsCaptchaContent = False
LogStep "  CONTENT-DETECT w" & workerIdx & ": results markers found -> VALID"
Else
IsCaptchaContent = True
LogStep "  CONTENT-DETECT w" & workerIdx & ": inconclusive DOM - assuming CAPTCHA (safe default)"
End If

CleanUp:
On Error Resume Next
If Dir(outPath) <> "" Then Kill outPath
On Error GoTo 0
Exit Function

Fallback:
IsCaptchaContent = True
On Error Resume Next
If Dir(outPath) <> "" Then Kill outPath
On Error GoTo 0
End Function

' Reads a whole text file into a string (used for the DOM probe).
Private Function ReadTextFile(ByVal path As String) As String
On Error Resume Next
Dim FSO As Object, ts As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ts = FSO.OpenTextFile(path, 1)   ' 1 = ForReading
ReadTextFile = ts.ReadAll
ts.Close
On Error GoTo 0
End Function


'------------------------------------------------------------------
' INTERACTIVE MODE - fallback if USE_HEADLESS = False.
' User's Edge print dialog controls headers/footers.
'------------------------------------------------------------------
Private Function PrintSearchToPDF_Interactive(query As String, savePath As String, _
wsh As Object, ByVal PageNum As Long) As Boolean
On Error GoTo FuncFail

Dim baseURL As String, edgeArgs As String
Dim waitCounter As Long, retryPrint As Long, saveAttempts As Long
Dim isFileReady As Boolean, saveWindowReady As Boolean
Dim fileSizeCheck As Long
Dim retryResponse As VbMsgBoxResult
#If VBA7 Then
Dim hwndDialog As LongPtr, hEdge As LongPtr
#Else
Dim hwndDialog As Long, hEdge As Long
#End If

PrintSearchToPDF_Interactive = False
isFileReady = False

baseURL = "https://www.google.com/search?q=" & URLEncode(query) & "&num=100&hl=en"
If PageNum > 1 Then baseURL = baseURL & "&start=" & ((PageNum - 1) * 10)
StartSearch:
On Error Resume Next
If Dir(savePath) <> "" Then Kill savePath
On Error GoTo FuncFail

edgeArgs = BuildEdgeArgs(baseURL)
ShellExecute 0, "open", "msedge.exe", edgeArgs, "", 3

For waitCounter = 1 To 30
SmartWait t(0.5)
hEdge = FindEdgeBrowserWindow()
If hEdge <> 0 Then Exit For
If bAbort Then GoTo EmergencyClose
Next waitCounter
' Fallback: if the title-based match never fired (odd Edge build /
' locale), fall back to the first window of the class so we still try.
If hEdge = 0 Then hEdge = FindWindow("Chrome_WidgetWin_1", vbNullString)
SmartWait t(1.5)

' Edge finding its window handle does NOT mean Edge has focus -
' Windows' foreground-lock can leave Excel as the active window
' (especially right after SafeCloseBrowserTab last handed focus
' back), in which case the keystrokes below would land on Excel
' instead of the browser. Force Edge to the foreground before we
' ever send it a keystroke.
If hEdge <> 0 Then ForceForeground hEdge
SmartWait t(0.3), True

retryPrint = 0
AttemptPrint:
If bAbort Then GoTo EmergencyClose

' Re-assert foreground AND verify before Ctrl+P - the old code just
' called ForceForeground and assumed it worked; if it didn't, Ctrl+P
' (and the save-dialog keys after) landed on Excel. Now we confirm
' Edge is actually the active window first, and if we can't, we
' retry/abort instead of firing keystrokes at Excel.
Dim pAttempt As Long, pFore As Boolean
pFore = False
For pAttempt = 1 To 6
If hEdge = 0 Then hEdge = FindEdgeBrowserWindow()
If hEdge <> 0 Then ForceForeground hEdge
SmartWait t(0.35), True
' Accept success when ANY Edge browser window is the foreground -
' robust to the handle differing from the exact one we grabbed.
If ForegroundIsEdge() Then pFore = True: Exit For
Next pAttempt

If Not pFore Then
retryPrint = retryPrint + 1
If retryPrint <= 2 Then
SmartWait t(1)
GoTo AttemptPrint
Else
GoTo TriggerRetryPrompt
End If
End If

wsh.SendKeys "^p"
SmartWait t(3)
wsh.SendKeys "{ENTER}"

saveWindowReady = False
For waitCounter = 1 To 30
hwndDialog = FindWindow(vbNullString, "Save As")
If hwndDialog = 0 Then hwndDialog = FindWindow(vbNullString, "Save Print Output As")
If hwndDialog <> 0 Then saveWindowReady = True: Exit For
SmartWait t(0.4)
If bAbort Then GoTo EmergencyClose
Next waitCounter

If Not saveWindowReady Then
retryPrint = retryPrint + 1
If retryPrint <= 2 Then
SmartWait t(2)
GoTo AttemptPrint
Else
GoTo TriggerRetryPrompt
End If
End If

SmartWait t(0.8)
CopyToClipboard savePath
SmartWait t(0.4)

saveAttempts = 0
Do While saveAttempts < 5
hwndDialog = FindWindow(vbNullString, "Save As")
If hwndDialog = 0 Then hwndDialog = FindWindow(vbNullString, "Save Print Output As")
If hwndDialog <> 0 Then ForceForeground hwndDialog

SmartWait t(1.2)

wsh.SendKeys "^a"
SmartWait t(0.3)
wsh.SendKeys "{DEL}"
SmartWait t(0.3)
wsh.SendKeys "+{INSERT}"
' Hard 1s floor (NOT scaled by TIMING_PROFILE) so the pasted
' filename has time to settle in the Save dialog field before we
' read it back.
SmartWait 1

' True verification: read the dialog's filename field back via
' Win32 (WM_GETTEXT on the file-name Edit control) and re-paste
' until it matches the path we intended. Catches the rare case
' where the paste didn't land - clipboard race, dialog not yet
' focused - before we commit the save with Alt+S.
Dim verifyTries As Long, fieldText As String, nameMatched As Boolean
nameMatched = False
For verifyTries = 1 To 4
fieldText = ReadSaveDialogFileName(hwndDialog)
If FileNameFieldMatches(fieldText, savePath) Then
nameMatched = True
Exit For
End If
If bAbort Then Exit For
' Mismatch (or couldn't read) - clear the field and paste again.
If hwndDialog <> 0 Then ForceForeground hwndDialog
wsh.SendKeys "^a"
SmartWait t(0.3)
wsh.SendKeys "{DEL}"
SmartWait t(0.3)
CopyToClipboard savePath
SmartWait t(0.3)
wsh.SendKeys "+{INSERT}"
SmartWait 1
Next verifyTries

If nameMatched Then
LogStep "  SAVE - filename field confirmed after " & verifyTries & " check(s)"
Else
LogStep "  SAVE warn - filename field never confirmed; proceeding anyway. field='" & _
fieldText & "'"
End If

wsh.SendKeys "%s"
SmartWait t(0.9)

hwndDialog = FindWindow(vbNullString, "Save As")
If hwndDialog = 0 Then hwndDialog = FindWindow(vbNullString, "Save Print Output As")
If hwndDialog = 0 Then Exit Do

If hwndDialog <> 0 Then ForceForeground hwndDialog
wsh.SendKeys "{ENTER}"
SmartWait t(0.9)

hwndDialog = FindWindow(vbNullString, "Save As")
If hwndDialog = 0 Then hwndDialog = FindWindow(vbNullString, "Save Print Output As")
If hwndDialog = 0 Then Exit Do

saveAttempts = saveAttempts + 1
Loop

For waitCounter = 1 To 40
SmartWait t(0.5)
If bAbort Then GoTo EmergencyClose

If Dir(savePath) <> "" Then
fileSizeCheck = 0
On Error Resume Next
fileSizeCheck = FileLen(savePath)
On Error GoTo FuncFail

If fileSizeCheck > MIN_PDF_SIZE_BYTES Then
isFileReady = True
Exit For
End If
End If
Next waitCounter

If isFileReady Then
PrintSearchToPDF_Interactive = True
SmartWait t(0.5)
SafeCloseBrowserTab wsh
Exit Function
End If
TriggerRetryPrompt:
SafeCloseBrowserTab wsh
retryResponse = MsgBox("Failed to detect the saved PDF." & vbCrLf & vbCrLf & _
"VBA was looking for the file exactly here:" & vbCrLf & savePath & _
vbCrLf & vbCrLf & "Click Retry to restart this specific search.", _
vbRetryCancel + vbCritical + vbSystemModal, "Save/Naming Error")

If retryResponse = vbRetry Then
SmartWait t(0.5)
GoTo StartSearch
Else
PrintSearchToPDF_Interactive = False
Exit Function
End If
EmergencyClose:
SafeCloseBrowserTab wsh
SmartWait t(0.5)
Exit Function

FuncFail:
On Error Resume Next
SafeCloseBrowserTab wsh
PrintSearchToPDF_Interactive = False
End Function

' ---- Window / SendKeys helpers (interactive fallback only) ----
#If VBA7 Then
Private Sub ForceForeground(ByVal hTarget As LongPtr)
#Else
Private Sub ForceForeground(ByVal hTarget As Long)
#End If
On Error Resume Next
Dim tidMe As Long, tidTarget As Long, dummyPid As Long
tidMe = GetCurrentThreadId()
tidTarget = GetWindowThreadProcessId(hTarget, dummyPid)
AttachThreadInput tidTarget, tidMe, 1
BringWindowToTop hTarget
SetForegroundWindow hTarget
AttachThreadInput tidTarget, tidMe, 0
On Error GoTo 0
End Sub

' Reads a window's title bar text.
#If VBA7 Then
Private Function WindowTitleText(ByVal h As LongPtr) As String
#Else
Private Function WindowTitleText(ByVal h As Long) As String
#End If
On Error Resume Next
WindowTitleText = ""
If h = 0 Then Exit Function
Dim buf As String, N As Long
buf = String$(512, vbNullChar)
N = GetWindowText(h, buf, 512)
If N > 0 Then WindowTitleText = Left$(buf, N)
On Error GoTo 0
End Function

' Finds the REAL, visible Edge BROWSER window - not the first window of
' class "Chrome_WidgetWin_1" (which on Win10/11 is shared by every
' WebView2 surface: Widgets, Copilot, Outlook, Teams, the headless
' warm-up Edge...). Grabbing one of those ghost windows is why the
' foreground check never confirmed and the print keystrokes were never
' sent. We enumerate top-level windows of that class and return the
' first VISIBLE one whose title identifies it as Microsoft Edge.
#If VBA7 Then
Private Function FindEdgeBrowserWindow() As LongPtr
Dim h As LongPtr
#Else
Private Function FindEdgeBrowserWindow() As Long
Dim h As Long
#End If
On Error Resume Next
h = 0
Do
h = FindWindowEx(0, h, "Chrome_WidgetWin_1", vbNullString)
If h = 0 Then Exit Do
If IsWindowVisible(h) <> 0 Then
If InStr(1, WindowTitleText(h), "Microsoft", vbTextCompare) > 0 _
And InStr(1, WindowTitleText(h), "Edge", vbTextCompare) > 0 Then
FindEdgeBrowserWindow = h
Exit Function
End If
End If
Loop
On Error GoTo 0
End Function

' True if the CURRENT foreground window is an Edge browser window. Used
' instead of an exact-handle match so the keystrokes fire as long as an
' Edge browser (not Excel, not a dialog) genuinely holds focus.
Private Function ForegroundIsEdge() As Boolean
Dim ttl As String
ttl = WindowTitleText(GetForegroundWindow())
ForegroundIsEdge = (InStr(1, ttl, "Microsoft", vbTextCompare) > 0 _
And InStr(1, ttl, "Edge", vbTextCompare) > 0)
End Function

' ---- Save-dialog filename verification (interactive fallback) ----
' Reads the text currently sitting in the Save dialog's file-name
' Edit control so we can confirm our pasted path actually landed
' before pressing Alt+S. Returns "" if the control can't be found
' or has no text.
#If VBA7 Then
Private Function ReadSaveDialogFileName(ByVal hDlg As LongPtr) As String
Dim hEdit As LongPtr
#Else
Private Function ReadSaveDialogFileName(ByVal hDlg As Long) As String
Dim hEdit As Long
#End If
On Error Resume Next
ReadSaveDialogFileName = ""
If hDlg = 0 Then Exit Function

hEdit = FindFileNameEdit(hDlg)
If hEdit = 0 Then Exit Function

Dim N As Long
N = CLng(SendMessageLen(hEdit, WM_GETTEXTLENGTH, 0, 0))
If N <= 0 Then Exit Function

Dim buf As String
buf = String$(N + 1, vbNullChar)
SendMessageGetText hEdit, WM_GETTEXT, N + 1, buf

Dim z As Long
z = InStr(buf, vbNullChar)
If z > 0 Then buf = Left$(buf, z - 1)
ReadSaveDialogFileName = buf
On Error GoTo 0
End Function

' Locates the file-name Edit control inside the Save dialog. Tries
' the modern Win10/11 common-item-dialog hierarchy first, then the
' classic dialog layout, then a last-ditch direct Edit child.
#If VBA7 Then
Private Function FindFileNameEdit(ByVal hDlg As LongPtr) As LongPtr
Dim h As LongPtr, hEdit As LongPtr
#Else
Private Function FindFileNameEdit(ByVal hDlg As Long) As Long
Dim h As Long, hEdit As Long
#End If
On Error Resume Next

' Strategy 1 - modern: DUIViewWndClassName > DirectUIHWND >
' FloatNotifySink > ComboBox > Edit.
h = FindWindowEx(hDlg, 0, "DUIViewWndClassName", vbNullString)
If h <> 0 Then h = FindWindowEx(h, 0, "DirectUIHWND", vbNullString)
If h <> 0 Then h = FindWindowEx(h, 0, "FloatNotifySink", vbNullString)
If h <> 0 Then h = FindWindowEx(h, 0, "ComboBox", vbNullString)
If h <> 0 Then hEdit = FindWindowEx(h, 0, "Edit", vbNullString)
If hEdit <> 0 Then
FindFileNameEdit = hEdit
Exit Function
End If

' Strategy 2 - classic: ComboBoxEx32 > ComboBox > Edit.
h = FindWindowEx(hDlg, 0, "ComboBoxEx32", vbNullString)
If h <> 0 Then h = FindWindowEx(h, 0, "ComboBox", vbNullString)
If h <> 0 Then hEdit = FindWindowEx(h, 0, "Edit", vbNullString)
If hEdit <> 0 Then
FindFileNameEdit = hEdit
Exit Function
End If

' Strategy 3 - last ditch: first direct Edit child of the dialog.
FindFileNameEdit = FindWindowEx(hDlg, 0, "Edit", vbNullString)
On Error GoTo 0
End Function

' True if the dialog field text matches the path we intended to
' save to. Accepts a full-path match (the usual case, since we
' paste the whole path) and falls back to a base-filename match
' with or without the extension, since some dialogs show only the
' name.
Private Function FileNameFieldMatches(ByVal fieldText As String, _
ByVal fullPath As String) As Boolean
Dim A As String, b As String, baseName As String, baseNoExt As String
A = Trim$(fieldText)
b = Trim$(fullPath)
If Len(A) = 0 Then Exit Function

If StrComp(A, b, vbTextCompare) = 0 Then
FileNameFieldMatches = True
Exit Function
End If

A = Replace(A, """", "")
baseName = Mid$(b, InStrRev(b, "\") + 1)
If StrComp(A, baseName, vbTextCompare) = 0 Then
FileNameFieldMatches = True
Exit Function
End If

If InStrRev(baseName, ".") > 0 Then
baseNoExt = Left$(baseName, InStrRev(baseName, ".") - 1)
If StrComp(A, baseNoExt, vbTextCompare) = 0 Then FileNameFieldMatches = True
End If
End Function

Sub SafeCloseBrowserTab(wsh As Object)
Application.EnableCancelKey = xlDisabled
SmartWait t(1.2), True

#If VBA7 Then
Dim hwndBrowser As LongPtr
#Else
Dim hwndBrowser As Long
#End If

hwndBrowser = FindEdgeBrowserWindow()
If hwndBrowser = 0 Then hwndBrowser = FindWindow("Chrome_WidgetWin_1", vbNullString)

' CRITICAL SAFETY: never send Ctrl+W unless we've CONFIRMED Edge is
' the foreground window. If no Edge window is found, or we can't
' bring it to the foreground, we must NOT fire the keystroke - it
' would land on Excel, and Ctrl+W in Excel closes the workbook (the
' exact "the tool gets closed instead of the Edge tab" symptom).
If hwndBrowser = 0 Then
' Nothing to close, and firing keys now would hit Excel.
LogStep "  CLOSE-TAB skipped - no Edge window found (not sending Ctrl+W at Excel)"
Application.EnableCancelKey = xlInterrupt
Exit Sub
End If

' Try to bring Edge forward and VERIFY it actually is, retrying a few
' times. Only proceed to the close keystroke once it's confirmed.
Dim attempt As Long, isFore As Boolean
isFore = False
For attempt = 1 To 6
ForceForeground hwndBrowser
SmartWait t(0.35), True
If ForegroundIsEdge() Then
isFore = True
Exit For
End If
Next attempt

If Not isFore Then
' Could not confirm Edge is active - do NOT send Ctrl+W. Leaving a
' stray Edge tab open is harmless; closing the workbook is not.
LogStep "  CLOSE-TAB skipped - could not confirm Edge foreground after " & _
attempt & " tries (leaving tab open rather than risking Excel)"
Application.EnableCancelKey = xlInterrupt
Exit Sub
End If

' Edge confirmed foreground - safe to close the tab.
wsh.SendKeys "{ESC}"
SmartWait t(0.3), True
' Re-verify immediately before the destructive keystroke.
If ForegroundIsEdge() Then
wsh.SendKeys "^w"
SmartWait t(0.5), True
Else
LogStep "  CLOSE-TAB aborted - Edge lost foreground just before Ctrl+W"
End If

' Deliberately NOT calling AppActivate Application.Caption here - doing
' so hands focus back to Excel and Windows' foreground-lock then keeps
' Excel active into the next iteration's keystrokes. Edge is
' re-foregrounded by the verify loop above instead.

Application.EnableCancelKey = xlInterrupt
End Sub

' Put text on the clipboard and verify it stuck. Some VDIs lose
' clipboard writes silently, so we read the value back; if that
' fails, we fall through to the htmlfile/IE trick as a backup.
Sub CopyToClipboard(ByVal text As String)
Dim objData As Object, verifyText As String
Dim clipboardSet As Boolean: clipboardSet = False

On Error Resume Next
Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
If Not objData Is Nothing Then
objData.SetText text
objData.PutInClipboard
objData.GetFromClipboard
verifyText = objData.GetText
If verifyText = text Then clipboardSet = True
End If
On Error GoTo 0

If Not clipboardSet Then
On Error Resume Next
Dim ieClip As Object
Set ieClip = CreateObject("htmlfile")
ieClip.ParentWindow.ClipboardData.SetData "text", text
Set ieClip = Nothing
On Error GoTo 0
End If

SmartWait t(0.3)
Set objData = Nothing
End Sub

' ---- Small string / data helpers ----
Function URLEncode(ByVal text As String) As String
Dim i As Long, Char As String, EncodedText As String
EncodedText = ""
For i = 1 To Len(text)
Char = Mid(text, i, 1)
Select Case Asc(Char)
Case 48 To 57, 65 To 90, 97 To 122
EncodedText = EncodedText & Char
Case 32
EncodedText = EncodedText & "+"
Case Else
' Zero-pad hex so single-digit byte values stay
' two characters wide (%0A, not %A). Some servers
' choke on the unpadded form.
EncodedText = EncodedText & "%" & Right("0" & Hex(Asc(Char)), 2)
End Select
Next i
URLEncode = EncodedText
End Function

Function SanitizeFileNamePart(ByVal text As String) As String
Dim invalidChars As Variant, i As Long

text = Replace(text, Chr(10), " ")
text = Replace(text, Chr(13), " ")
text = Replace(text, Chr(9), " ")

invalidChars = Array("""", "/", "\", ":", "*", "?", "<", ">", "|")
For i = LBound(invalidChars) To UBound(invalidChars)
text = Replace(text, invalidChars(i), " ")
Next i

Do While InStr(text, "  ") > 0
text = Replace(text, "  ", " ")
Loop

text = Trim(text)
If IsReservedFilename(text) Then text = "_" & text

SanitizeFileNamePart = text
End Function

Private Function NormalizeSpaces(ByVal text As String) As String
' Collapse runs of whitespace into a single space. Unlike
' SanitizeFileNamePart this doesn't strip any characters - it
' just normalises spacing so "John  Smith" becomes "John Smith".
' We use a bounded For loop instead of Do While so there is no
' chance of an infinite loop on weird input.
LogStep "  NormalizeSpaces entry, inLen=" & Len(text)
On Error Resume Next
text = Replace(text, Chr(9), " ")      ' tab
text = Replace(text, Chr(10), " ")     ' LF
text = Replace(text, Chr(11), " ")     ' VT
text = Replace(text, Chr(12), " ")     ' FF
text = Replace(text, Chr(13), " ")     ' CR
text = Replace(text, Chr(160), " ")    ' non-breaking space
Dim i As Long
For i = 1 To 50
If InStr(text, "  ") = 0 Then Exit For
text = Replace(text, "  ", " ")
Next i
NormalizeSpaces = Trim$(text)
LogStep "  NormalizeSpaces exit, outLen=" & Len(NormalizeSpaces) & " iter=" & (i - 1)
On Error GoTo 0
End Function

' Returns the name with a trailing legal/business suffix removed (e.g.
' "Acme Trading LLC" -> "Acme Trading"), or "" if the name doesn't end
' in one of the known suffixes. A leading comma before the suffix
' ("Acme Trading, Inc.") is dropped too. Only used to build the extra
' Negative News "without legal extension" search - every other search
' keeps the full name as entered.
Private Function StripLegalExtension(ByVal Name As String) As String
    On Error Resume Next
    Dim s As String
    s = Trim$(Name)
    If s = "" Then Exit Function

    ' Known suffixes. Comparison is case-insensitive and ignores a
    ' trailing period on the last token, so "Inc" and "Inc." both match.
    ' Multiword suffixes MUST come first so "Pvt Ltd" is matched whole
    ' before the single-word "Ltd" strips only half of it.
    Dim suffixes As Variant
    suffixes = Array("PRIVATE LIMITED", "PVT LTD", _
                     "LLC", "L.L.C.", "LLP", "LP", "PLLC", "PLC", _
                     "INC", "INCORPORATED", "CORP", "CORPORATION", _
                     "LTD", "LIMITED", "CO", "COMPANY", _
                     "GMBH", "AG", "SARL", "SA", "NV", "BV", "PTY", "PVT")

    ' Compare on a period-stripped, upper-cased copy of the trailing word(s).
    Dim words() As String
    words = Split(s, " ")
    Dim nWords As Long
    nWords = UBound(words) + 1
    If nWords < 2 Then Exit Function   ' a lone token isn't "name + suffix"

    Dim j As Long, suf As String, sufWords As Long, tail As String
    For j = LBound(suffixes) To UBound(suffixes)
        suf = CStr(suffixes(j))
        sufWords = UBound(Split(suf, " ")) + 1
        If nWords > sufWords Then          ' must leave at least one base word
            tail = JoinLastWords(words, sufWords)
            If NormalizeSuffixToken(tail) = NormalizeSuffixToken(suf) Then
                StripLegalExtension = JoinFirstWords(words, nWords - sufWords)
                ' Drop a trailing comma the suffix was hanging off of.
                Do While Len(StripLegalExtension) > 0 And _
                         Right$(StripLegalExtension, 1) = ","
                    StripLegalExtension = Left$(StripLegalExtension, Len(StripLegalExtension) - 1)
                Loop
                StripLegalExtension = Trim$(StripLegalExtension)
                Exit Function
            End If
        End If
    Next j
    On Error GoTo 0
End Function

' Upper-case a token and strip periods/commas so "L.L.C.," = "LLC".
Private Function NormalizeSuffixToken(ByVal t As String) As String
    t = UCase$(Trim$(t))
    t = Replace(t, ".", "")
    t = Replace(t, ",", "")
    NormalizeSuffixToken = Trim$(t)
End Function

Private Function JoinLastWords(ByRef words() As String, ByVal N As Long) As String
    Dim k As Long, out As String
    For k = UBound(words) - N + 1 To UBound(words)
        out = out & IIf(out = "", "", " ") & words(k)
    Next k
    JoinLastWords = out
End Function

Private Function JoinFirstWords(ByRef words() As String, ByVal N As Long) As String
    Dim k As Long, out As String
    For k = LBound(words) To LBound(words) + N - 1
        out = out & IIf(out = "", "", " ") & words(k)
    Next k
    JoinFirstWords = out
End Function

' Appends one OSDD run's summary to the LOCAL "Audit_Log" sheet - the
' v3.4 OSINT run log (Date&Time | Analyst ID | ECM | Total Entities |
' Total Searches | Time Taken | Tool Version | Error Diagnostic |
' CAPTCHA Flagged).
' Best-effort and fully self-contained: it creates the sheet (with the
' v3.4 headers) if missing, respects sheet/workbook protection, and
' swallows any error so it can never disrupt the search run.
Private Sub WriteLocalAuditLog(ByVal ecmCase As String, ByVal entityCount As Long, _
    ByVal searchCount As Long, ByVal timeString As String, _
    ByVal captchaFlagged As Long, ByVal errDiag As String)
    On Error Resume Next
    Dim ws As Worksheet, wasProt As Boolean, wbWasProt As Boolean, r As Long

    Set ws = ThisWorkbook.Sheets("Audit_Log")
    If ws Is Nothing Then
        ' Recreate the tab with its v3.4 header row if it's ever missing.
        wbWasProt = ThisWorkbook.ProtectStructure
        If wbWasProt Then ThisWorkbook.Unprotect Password:="p7ss"
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = "Audit_Log"
        ws.Range("A1:I1").Value = Array("Date & Time", "Analyst ID", "ECM Case ID", _
            "Total Entities", "Total Searches", "Time Taken", "Tool Version", _
            "Error Diagnostic", "CAPTCHA Flagged")
        If wbWasProt Then ThisWorkbook.Protect Password:="p7ss", Structure:=True
    End If
    If ws Is Nothing Then Exit Sub

    wasProt = ws.ProtectContents
    If wasProt Then ws.Unprotect Password:="p7ss"

    r = ws.Cells(ws.Rows.count, "A").End(xlUp).row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = Environ$("USERNAME")
    ws.Cells(r, 3).Value = ecmCase
    ws.Cells(r, 4).Value = entityCount
    ws.Cells(r, 5).Value = searchCount
    ws.Cells(r, 6).Value = timeString
    ws.Cells(r, 7).Value = TOOL_VERSION
    ws.Cells(r, 8).Value = errDiag
    ws.Cells(r, 9).Value = captchaFlagged

    If wasProt Then ws.Protect Password:="p7ss"
    On Error GoTo 0
End Sub

Private Function IsReservedFilename(ByVal Name As String) As Boolean
Dim reserved As Variant, r As Variant, upper As String
reserved = Array("CON", "PRN", "AUX", "NUL", _
"COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
"LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
upper = UCase(Name)
For Each r In reserved
If upper = r Then IsReservedFilename = True: Exit Function
Next r
End Function

Private Function SheetExists(wb As Workbook, sheetName As String) As Boolean
Dim s As Worksheet
On Error Resume Next
Set s = wb.Sheets(sheetName)
SheetExists = Not s Is Nothing
On Error GoTo 0
End Function

' TIMING_PROFILE scales every internal wait by a single factor so
' we can dial things tighter on a fast machine or looser on a
' struggling VDI without touching individual sleep calls.
Private Function t(BaseSeconds As Single) As Single
Select Case TIMING_PROFILE
Case "FAST": t = BaseSeconds * 0.7
Case "SLOW": t = BaseSeconds * 1.5
Case Else:   t = BaseSeconds
End Select
End Function

' DoEvents-friendly sleep. Keeps Excel responsive, listens for the
' ESC key so the analyst can cancel, and survives a midnight
' rollover (Timer resets to 0 after 86 400 seconds).
Sub SmartWait(Seconds As Single, Optional IgnoreESC As Boolean = False)
Dim EndTime As Single
EndTime = Timer + Seconds
Do While Timer < EndTime
DoEvents
If Not IgnoreESC Then
If GetAsyncKeyState(27) <> 0 Then
bAbort = True
Exit Sub
End If
End If
If Timer < EndTime - 86000 Then EndTime = EndTime - 86400
Loop

' Drain any ESC presses we swallowed while ignoring them.
If IgnoreESC Then GetAsyncKeyState 27
End Sub

