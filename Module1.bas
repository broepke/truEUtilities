Attribute VB_Name = "ModPurge"
Option Explicit


Private Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Boolean
  hNameMappings As Long
  lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias _
  "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40


Public KillCounter As Long              'Keep track of how many files were deleted
Public RenameCounter As Long            'Keep track of how many files were deleted
Public FileArray(10000, 1) As String    '2D Array to store file names and extensions
Public rPath As String                  'Name of current path (rename)
Public rName As String                  'Name of current file or subdirectory (rename)
Public DirArray() As String             'Array to hold all directories
Public DirArrayCounter As Long          'Counter for directory array
Public FileArrayCounter As Long         'Counter for file array
Public DotCount As Long                 'Number of "." in the string
Public BrowsePath As String             'Used in the options for purging (custom path)
Public PurgeLoc As String               'OptionButton value for location to move files
Public RecycleVal As Long               'value based on operating system
Public RecyclePath As String            'Path set based on operating system


Public Sub subPurge(Optional DirArray, Optional PurgeCounter As Long)
    
Dim x As Long           ' Place holder for each element in the array
Dim fs, f, f1, fc, s    ' File System Object variables
Dim PurgeDir As String ' The Current file found the the dir
    
On Error GoTo ErrorHandler
    
'Loop thru all directories and for each file call the "AddToFileArray" sub
'This is done one directory at a time to keep all files separate when in dir dirs
    For x = 0 To PurgeCounter
    
    frmMain.ProgressBar1.Value = x
    
        PurgeDir = DirArray(x)

        If PurgeDir <> "" Then
        
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFolder(PurgeDir)
            Set fc = f.Files
                For Each f1 In fc
                    Call AddToFileArray(f1.Path)
                    
                Next
                    
        End If
        
    Call subKill(PurgeDir)
StartHere:
    Next x
    
Exit Sub

'Make sure that if the user has no permissons to the directory that
'They can get out of the loop.
ErrorHandler:
    Call ErrorHandler(Err.Number, f.Path)
    Resume StartHere
    
End Sub

Public Sub AddToFileArray(argFileName As String)

Dim POS As Integer              'Variable to hold the position of the right most "."
Dim CurVal As Integer           'The current value of the revision of the file
Dim FileNameWOEXT As String     'The name of the file with out the number extension
Dim FOUND As Boolean            'Flag that states wether the file being is added to the array
Dim x As Integer                'Variable for the for loop
Dim ExtStr As String            'Get the extension of the file
Dim ExtCheck                    'Value returned when checking to see if the ext is a number


'Check the current file to see if the extension is a number
POS = InStrRev(argFileName, ".", -1)
ExtStr = Right(argFileName, Len(argFileName) - POS)
ExtCheck = IsNumeric(ExtStr)

'Call a function to check how many dots exist in a file name
Call DotCheck(argFileName)

If ExtCheck = True And DotCount = 2 Then

'Find the numeric value of the extension and file with out extension.
CurVal = Val(Right(argFileName, Len(argFileName) - POS))
FileNameWOEXT = Left(argFileName, POS - 1)

    
'Add the first file that comes here to the array
    If FileArrayCounter = 0 Then
        FileArray(FileArrayCounter, 0) = FileNameWOEXT
        FileArray(FileArrayCounter, 1) = CurVal
        FileArrayCounter = FileArrayCounter + 1
    
    Else
    
'Check to see if the file is in the array and if it is check to see if it has a higher ext.
        For x = 0 To FileArrayCounter
            
            If FileArray(x, 0) = FileNameWOEXT Then
                
                FOUND = True
                
                If FileArray(x, 1) < CurVal Then
                    FileArray(x, 1) = CurVal
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
            
           
        Next x
        
'If the file is not found yet add it and the ext to the array
        If Not FOUND Then
            FileArray(FileArrayCounter, 0) = FileNameWOEXT
            FileArray(FileArrayCounter, 1) = CurVal
    
            FileArrayCounter = FileArrayCounter + 1
           
        End If
        
    End If

Else
    Exit Sub
End If

End Sub

Public Sub subKill(KillPath As String)

Dim FOUND As Boolean           'Value to tell whether or not the file is in the array
Dim POS As Integer              'Find the position of the "."
Dim x As Integer                'Place holder for looping thru the array
Dim ExtStr As String            'String for the extension (check to see if it is a num)
Dim ExtCheck                    'Value returned when checking to see if the ext is a number
Dim CurVal As Integer           'The current value of the revision of the file
Dim FileNameWOEXT As String     'The name of the file with out the number extension
Dim SkipFile As String          'File used if user is moveing files and gets error 58
Dim DeleteFile As String
Dim fs, f, f1, fc, s

Dim ErrArray()                  'Files with errors in directory
Dim ErrArrayCounter As Long     'Counter for array
Dim ErrMsg
Dim e

BrowsePath = GetSetting(App.Title, "UserSettings", "BrowsePath", "")
FOUND = False
x = 0
ErrArrayCounter = 0
ReDim ErrArray(ErrArrayCounter)

On Error GoTo ErrorHandler

    If KillPath <> "" Then
    
        ' Get the files in the current dir - one at a time
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFolder(KillPath)
        Set fc = f.Files
        For Each f1 In fc
    
                ' Check to see if the extension is a number
                POS = InStrRev(f1.Path, ".", -1)
                ExtStr = Right(f1.Path, Len(f1.Path) - POS)
                ExtCheck = IsNumeric(ExtStr)
                
                ' Check for two "." in the name of the file
                Call DotCheck(f1.Path)
                
                If ExtCheck = True And DotCount = 2 Then
                        
                ' Get the extension value and file name without extension
                CurVal = Val(Right(f1.Path, Len(f1.Path) - POS))
                FileNameWOEXT = Left(f1.Path, POS - 1)
                
                If FileNameWOEXT = SkipFile Then
                    GoTo SkipLocation
                End If
                        
                    ' Loop thru all higest files and check for a match
                    For x = 0 To FileArrayCounter
                        If FileArray(x, 0) & "." & FileArray(x, 1) = _
                                FileNameWOEXT & "." & CurVal Then
                            ' Set equal to true if the file is found
                            FOUND = True
                        End If
                    Next
                        
                                           
                    If FOUND = False Then
                        
                        'Delete the file permanently
                        If PurgeLoc = 0 Then
                            f1.Delete
                            KillCounter = KillCounter + 1
                        
                        'Move the files to the recycling bin
                        ElseIf PurgeLoc = 1 Then
                            Dim FileOperation As SHFILEOPSTRUCT
                            Dim lReturn As Long
                            Dim sTempFilename As String
                            Dim sSendMeToTheBin As String
                            sTempFilename = f1
                            sSendMeToTheBin = sTempFilename
                            With FileOperation
                               .wFunc = FO_DELETE
                               .pFrom = sSendMeToTheBin
                               .fFlags = FOF_ALLOWUNDO
                            End With
                            lReturn = SHFileOperation(FileOperation)
                            KillCounter = KillCounter + 1
                        
                        'Move the file to the Browse Path
                        ElseIf PurgeLoc = 2 Then
                            f1.Move BrowsePath & "\"
                            KillCounter = KillCounter + 1
                        End If
                    End If
                                              
                
                Else
                    FOUND = False
                End If
                
StartHere:
SkipLocation:
            FOUND = False
        Next
            
        End If
        
    If ErrArrayCounter > 0 Then
        ErrMsg = MsgBox(ErrArrayCounter & " file(s) were not accessible in " & f & vbCrLf & _
        "Click Ok to see a list of the file(s)", vbOKCancel + vbInformation)
        If ErrMsg = vbOK Then
            ErrMsg = ""
            For e = 0 To ErrArrayCounter
                ErrMsg = ErrMsg & ErrArray(e)
                ErrMsg = ErrMsg & vbCrLf
            Next e
            MsgBox (ErrMsg)
        End If
    End If
    
    ' Reset the counters to speed up the process and keep the dirs independent
    FileArrayCounter = 0
    Erase FileArray
Exit Sub

ErrorHandler:
     
    ErrArray(ErrArrayCounter) = f1.Path
    ErrArrayCounter = ErrArrayCounter + 1
    ReDim Preserve ErrArray(ErrArrayCounter)
    
    Resume StartHere
    
End Sub

Public Function DotCheck(argDotFile As String) As Long

    Dim Position, Count
    Position = 1
    Count = 0
    DotCount = 0
    Do While InStr(Position, argDotFile, ".")
            Position = InStr(Position, argDotFile, ".") + 1
            Count = Count + 1
        Loop
    DotCount = Count
End Function
