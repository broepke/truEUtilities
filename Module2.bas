Attribute VB_Name = "ModJunk"
Option Explicit
Public JunkVal As String
Public ExtArray() As String
Public ExtArrayCounter As Long
Public JunkFileArray() As String
Public JunkFileArrayCounter As Long
Public spString As String

Public Function JunkArray(argJunkVal As String) As String

Dim JunkExt As String
Dim Position As Long
Dim counter As Long
Dim PositionOld As Long
Dim LastCommaPOS As Long
Dim LastComma As String
    
ExtArrayCounter = 0
PositionOld = 0
Position = 1
counter = 0
ReDim ExtArray(ExtArrayCounter)
    
       
    Do While InStr(Position, argJunkVal, ",")
        Position = InStr(Position, argJunkVal, ",") + 1
        JunkExt = Left(argJunkVal, Position - 2)
        If counter = 0 Then
        Else
            JunkExt = Right(JunkExt, Position - PositionOld)
        End If
        ExtArray(ExtArrayCounter) = JunkExt
        ExtArrayCounter = ExtArrayCounter + 1
        ReDim Preserve ExtArray(ExtArrayCounter)
        PositionOld = Position + 1
        counter = counter + 1
         
    Loop
    

End Function

Public Sub subFindJunk(Optional fDirArray, Optional fDirCounter As Long)

Dim f As Long   'Place holder for each element in the array
Dim e As Long
Dim CurExt As String
Dim CurJunkFile As String
Dim fName As String
Dim fPath As String
    
'Loop thru all directories and for each file call the "AddToFileArray"
'This is done one directory at a time to keep all files separate when in dir dirs

JunkFileArrayCounter = 0
ReDim Preserve JunkFileArray(JunkFileArrayCounter)

On Error GoTo ErrorHandler

For f = 0 To fDirCounter
    
frmMain.ProgressBar1.Value = f
    
    fPath = fDirArray(f) & "\"
    If Right(fPath, 2) = "\\" Then
        fPath = Left(fPath, InStr(1, fPath, "\", vbTextCompare))
    End If
    
        If fPath <> "" Then
        
                    
            For e = 0 To ExtArrayCounter
                CurExt = ExtArray(e)
                
                If CurExt <> "" Then
                
                CurJunkFile = Dir(fPath & "*." & CurExt & "*", vbReadOnly)
                    Do While CurJunkFile <> ""
                        If CurExt <> "." And CurExt <> ".." Then
                                                        
                            JunkFileArray(JunkFileArrayCounter) = fPath & CurJunkFile
                            JunkFileArrayCounter = JunkFileArrayCounter + 1
                            ReDim Preserve JunkFileArray(JunkFileArrayCounter)
                        
                        End If
                        
                        CurJunkFile = Dir
                        
                    Loop
                
                End If
            Next e
        End If
    Next f

Exit Sub

ErrorHandler:
    Call ErrorHandler(Err.Number)
    Exit Sub

End Sub
