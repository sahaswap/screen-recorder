Attribute VB_Name = "Module8"
Sub Start_Button_Create_Folders()
Dim wsHome As Worksheet
Dim ecmID As String
Dim desktopPath As String, mainFolderPath As String
Dim subFolderPath As String, otherFolderPath As String
Dim FSO As Object
Dim userProfile As String

Set wsHome = ThisWorkbook.Sheets("Sheet1")
ecmID = Trim(wsHome.Range("J9").Value)

If ecmID = "" Then
' No ECM ID yet: offer to fill the case from the ECM export file
' (modECMImport), which also creates these folders and copies the
' transaction file in.
If MsgBox("ECM ID is empty." & vbCrLf & vbCrLf & _
          "Fill the case from the ECM export file (ECM_Alert_Info_....csv) in Downloads?", _
          vbYesNo + vbQuestion, "START") = vbYes Then
    Import_ECM_Info
Else
    MsgBox "Action Denied: ECM ID is missing." & vbCrLf & "Please fill in the ECM ID before clicking START.", vbCritical, "Missing ID"
End If
Exit Sub
End If

Set FSO = CreateObject("Scripting.FileSystemObject")
userProfile = Environ("USERPROFILE")

' STRICT PATHING MATCHING THE CONSOLIDATION MACRO
If FSO.FolderExists(userProfile & "\OneDrive - Community Federal Savings Bank\Desktop") Then
desktopPath = userProfile & "\OneDrive - Community Federal Savings Bank\Desktop"
Else
desktopPath = userProfile & "\Desktop"
End If

mainFolderPath = desktopPath & "\" & ecmID
subFolderPath = mainFolderPath & "\Transaction Files"
otherFolderPath = mainFolderPath & "\Pivot"

' Safely create Main Folder if it doesn't exist
If Not FSO.FolderExists(mainFolderPath) Then
FSO.CreateFolder mainFolderPath
End If

' Safely create the Transaction Files sub-folder if it doesn't exist.
' (Non Alerted / Lookback folders are no longer created here.)
If Not FSO.FolderExists(subFolderPath) Then
FSO.CreateFolder subFolderPath
End If

' Safely create the Pivot sub-folder too.
If Not FSO.FolderExists(otherFolderPath) Then
FSO.CreateFolder otherFolderPath
End If

' Open the main case folder so the analyst sees both sub-folders.
Shell "explorer.exe """ & mainFolderPath & """", vbNormalFocus
End Sub



