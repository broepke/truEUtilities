Attribute VB_Name = "ModOpenFile"
     Function StartDoc(DocName As String)
      On Error GoTo StartDoc_Error

        StartDoc = ShellExecute(Application.hWndAccessApp, "Open", DocName, "", "C:\", SW_SHOWNORMAL)
      Exit Function

StartDoc_Error:
         MsgBox "Error: " & Err & " " & Error
         Exit Function
      End Function
