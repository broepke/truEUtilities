Attribute VB_Name = "ModError"
Sub ErrorHandler(ErrorNumber, Optional ErrorFile)

Dim Msg As String
Dim MsgVars As String

If ErrorNumber <> 0 Then

Select Case ErrorNumber

    Case 6
        Msg = "You have reached the maximum number of files that can be displayed." & vbCrLf & _
            "You will need to repeat this operation to clean all files."
        MsgVars = vbCritical + vbOKOnly
    Case 7
        Msg = "You have run out of memory on your system." & vbCrLf & _
            "Please close any application before continuing."
        MsgVars = vbCritical + vbOKOnly
    Case 53
        Msg = "File not found"
        MsgVars = vbCritical + vbOKOnly
    Case 58
        Msg = ErrorFile & " already exists."
        MsgVars = vbCritical + vbOKOnly
    Case 61
        Msg = "You disk is full."
        MsgVars = vbCritical + vbOKOnly
    Case 68
        Msg = UCase(ErrorFile) & "\ is not accessible." & vbcerlf & vbCrLf & "The device is not ready."
        MsgVars = vbCritical + vbOKOnly 'vbAbortRetryIgnore
    Case 70
        Msg = "Permission was denied to " & ErrorFile & vbCrLf & _
            "This series of files will not be purged correctly."
        MsgVars = vbCritical + vbOKOnly
    Case 71
        Msg = "Disk not ready."
        MsgVars = vbCritical + vbOKOnly
    Case 74
        Msg = "You can't rename to a different drive than the file is on."
        MsgVars = vbCritical + vbOKOnly
    Case 75
        Msg = "There was an error accessing the file " & ErrorFile & vbCrLf & _
            "The program will resume without changing the file."
        MsgVars = vbCritical + vbOKOnly
    Case 76
        Msg = "The Directry path not found."
        MsgVars = vbCritical + vbOKOnly
    Case Else
        Msg = "truEUtilities has experienced an undetermined error." & vbCrLf & _
                "Please click ok to continue."
        MsgVars = vbCritical + vbOKOnly
    End Select
End If

MsgBox Msg, MsgVars, "Error Message"

End Sub



