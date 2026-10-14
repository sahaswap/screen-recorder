Attribute VB_Name = "Module1"
Sub ClearForm()
Dim ws As Worksheet
Dim restoreErrNum As Long, restoreErrDesc As String

' Bind to Sheet1 by name so Reset always clears the form sheet,
' regardless of which sheet happens to be active when it's clicked.
Set ws = ThisWorkbook.Sheets("Sheet1")

' Mute Excel for speed during the clear.
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

On Error GoTo Restore

' UNLOCK AND CLEAR Sheet1 input fields only.
' ConsolidatedData is intentionally left intact - Sheet7's
' formulas pull from there.
ws.Unprotect Password:="p7ss"
' Input map on Sheet1:
'   J5        = Decision (RFI/Escalation/...)   - now cleared too
'   J6:J11    = case header fields (ECM J9, Alert J10, etc.)
'   J13       = Customer Name  (was being SKIPPED before)
'   J14:J15   = Customer Address block
'   J18:J23   = Counterparty names   (matches GetCounterpartyList)
'   T18:T23   = Counterparty second field (addresses)
'   G27       = country; L27:N27 = ISO2 (full merge); O27 = ISO3.
' ISO2/ISO3 are input cells now (no longer XLOOKUP formulas), so they
' must be cleared explicitly - clearing the country alone won't blank them.
ws.Range("J5:J11, J13:J15, J18:J23, T18:T23, G27, L27:N27, O27").Value = ""
ws.Protect Password:="p7ss"
Restore:
restoreErrNum = Err.Number
restoreErrDesc = Err.Description

On Error Resume Next
' Best-effort re-protect in case the clear bailed mid-way.
ws.Protect Password:="p7ss"
On Error GoTo 0

' --- THE KEY CHANGE ---
' Force Automatic mode unconditionally. This corrects any leaked
' Manual-mode state from prior macros (OSINT, etc.) regardless
' of what the workbook came in with.
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True

On Error Resume Next
Application.CalculateFull
' Reset runs with events OFF, so Worksheet_Change never fires and the
' Search Matrix "Open URL" hyperlinks (macro-built, static objects) would
' otherwise keep pointing at the previous case. Rebuild them now: column C
' is blank post-reset, so this wipes E and adds nothing back - clean slate.
RefreshSearchMatrixHyperlinks
On Error GoTo 0

DoEvents

If restoreErrNum <> 0 Then
    MsgBox "ClearForm hit an error and stopped early:" & vbCrLf & vbCrLf & _
           "Err " & restoreErrNum & ": " & restoreErrDesc & vbCrLf & vbCrLf & _
           "Sheet1 may need to be re-protected manually.", _
           vbCritical, "ClearForm"
Else
    MsgBox "Form cleared.", vbInformation
    ' Load the next waiting ECM alert straight away (modECMImport), if any.
    ECM_AfterReset
End If
End Sub



