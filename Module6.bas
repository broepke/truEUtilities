Attribute VB_Name = "ModDirPath"
Option Explicit
Public FinalDirPath As String
Public Function DirPath(DirPathOld, DirDisplayLoc As Double, FromForm As Form, DirLen As Integer) As String

'Code to fill text with a directory Path that is shortened to 25 charaters max
'Returns - FinalDirPath

Dim WorkDirLen As Long
Dim SmallDirPath As String
Dim DrivePath As String
Dim SmallDirPathPos As Long
Dim WorkDirLenNum As Double
Dim DrivePathWidth As Double
Dim RemainLblWidth As Double
Dim SmallDirPathLen As Double

'Convert the with of the path where the string in to be placed to pixels
If FromForm.ScaleMode = vbTwips Then
    DirDisplayLoc = DirDisplayLoc / Screen.TwipsPerPixelX  ' if twips change to pixels
End If

    'Find the lenght of the path in twips
    WorkDirLen = FromForm.TextWidth(DirPathOld & "  ")
    WorkDirLenNum = Len(DirPathOld)
    'Change to length in pixels
    If FromForm.ScaleMode = vbTwips Then
        WorkDirLen = WorkDirLen / Screen.TwipsPerPixelX
    End If
    
'    If WorkDirLen > DirDisplayLoc Then
        DrivePath = Left(DirPathOld, 3)
        
        'Find the width of the drive path in pixels
        DrivePathWidth = FromForm.TextWidth(DrivePath & "..")
        If FromForm.ScaleMode = vbTwips Then
            DrivePathWidth = DrivePathWidth / Screen.TwipsPerPixelX
        End If
        'Find the width of the remaining lable that can be used afer the drive letter
        RemainLblWidth = DirDisplayLoc - DrivePathWidth
        
        SmallDirPathPos = InStrRev(DirPathOld, "\", WorkDirLenNum - DirLen)
        SmallDirPath = Right(DirPathOld, WorkDirLenNum - SmallDirPathPos + 1)
               
        SmallDirPathLen = FromForm.TextWidth(SmallDirPath)
        If FromForm.ScaleMode = vbTwips Then
            SmallDirPathLen = SmallDirPathLen / Screen.TwipsPerPixelX
        End If
            
            If SmallDirPathLen > RemainLblWidth Then
                FinalDirPath = DrivePath & ".."
            Else
                FinalDirPath = DrivePath & ".." & SmallDirPath
            End If
    
    'Else
        'DirPath = FinalDirPath
'    End If
    
    'Return the value for the function
    DirPath = FinalDirPath

End Function





