Attribute VB_Name = "ModGlobal"
Public Sub subRecurse(initPath As String)

'Recursive search for subdirectories call from sub by Call "subRecurse(Directory)"
'Returns DirArray(DirArrayCounter)

Dim OldCounter As Long
Dim OrgCounter As Long
Dim x As Long
Dim vCurDirArray As String
Dim fs, f, f1, fc, s


DirArrayCounter = 0
OldCounter = 0
OrgCounter = 0
ReDim DirArray(0)

'Set the top level dir to the first element in the array
DirArray(DirArrayCounter) = initPath
DirArrayCounter = DirArrayCounter + 1
ReDim Preserve DirArray(DirArrayCounter)

On Error GoTo ErrorHandler

'Loop thu all sub dirs until no further subs are found
    Do
    
        OrgCounter = DirArrayCounter
        
        For x = OldCounter To OrgCounter
                
            vCurDirArray = DirArray(x)
            
            If vCurDirArray <> "" Then
                
               Set fs = CreateObject("Scripting.FileSystemObject")
                Set f = fs.GetFolder(vCurDirArray)
                Set fc = f.SubFolders
                For Each f1 In fc
                    DirArray(DirArrayCounter) = f1.Path
                    DirArrayCounter = DirArrayCounter + 1
                    ReDim Preserve DirArray(DirArrayCounter)
                Next
                                            
StartHere:

            OldCounter = x + 1
            
            End If
  

        
        Next
    
    Loop Until OldCounter >= DirArrayCounter

Exit Sub

'Make sure that if the user has no permissons to the directory that
'They can get out of the loop. - remove the dir from the array
ErrorHandler:
    Resume StartHere
    
End Sub
