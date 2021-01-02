Attribute VB_Name = "ModRename"
Option Explicit
Public RenameArray() As String          'Array to check for like files during rename
Public RenameArrayCounter As Long       'Counter for rename array
Public RenameVal As String              'Rename value
Public RenameChoice As Double           'Radio button to determine how to rename files
Public NOExtArray() As String           'Array to check for like files during rename
Public MultiFOUND As Boolean            'More than one file with the same name in a dir

Public Sub subRename(Optional rDirArray, Optional rCounter As Long)  'Subroutine to rename files


Dim r As Long               ' Place holder for each element in the array
Dim fs, f, f1, fc           ' File System Object variables
Dim PurgeDir As String      ' The Current file found the the dir
Dim x As Long
Dim rPath As String
    
On Error GoTo ErrorHandler
    
'Loop thru all directories and for each file call the "AddToFileArray" sub
'This is done one directory at a time to keep all files separate when in dir dirs
    For x = 0 To rCounter
    
    'Display the steps of the progress bar
    frmMain.ProgressBar1.Value = x
    
        rPath = rDirArray(x)

        If rPath <> "" Then
        
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFolder(rPath)
            Set fc = f.Files
                For Each f1 In fc
                    Call AddToRenameArray(f1.Path, f1.Name)
                    
                    If MultiFOUND = True And rCounter = 0 Then
                        MsgBox "You have not purged your directory yet. " & vbCrLf & _
                                "Please do so before running truERename.", vbCritical
                        Exit Sub
                    ElseIf MultiFOUND = True And rCounter > 0 Then
                        MsgBox "All of your directories have not been purged yet. " & vbCrLf & _
                                "Please do so and run truERename again.", vbCritical
                        Exit Sub
                    Else
                    End If
                    
                Next
                    
        End If
        
    Call subRemExt

StartHere:
    
    Next x
    
Exit Sub

'Make sure that if the user has no permissons to the directory that
'They can get out of the loop.
ErrorHandler:
    Call ErrorHandler(Err.Number, f.Path)
    Resume StartHere
    
End Sub

Sub AddToRenameArray(argFileName As String, argShortName As String)

Dim POS As Integer              'Variable to hold the position of the right most "."
Dim CurVal As Integer           'The current value of the revision of the file
Dim FileNameWOEXT As String     'The name of the file with out the number extension
Dim rFOUND As Boolean           'Flag that states wether the file being is added to the array
Dim x As Integer                'Variable for the for loop
Dim ExtStr As String            'Get the extension of the file
Dim ExtCheck                    'Value returned when checking to see if the ext is a number

MultiFOUND = False

ReDim Preserve RenameArray(RenameArrayCounter)
ReDim Preserve NOExtArray(RenameArrayCounter)

'Check the current file to see if the extension is a number
POS = InStrRev(argFileName, ".", -1)
ExtStr = Right(argFileName, Len(argFileName) - POS)
ExtCheck = IsNumeric(ExtStr)

Call DotCheck(argShortName)

If ExtCheck = True And DotCount = 2 Then

'Extract the file name without the extension
FileNameWOEXT = Left(argFileName, POS - 1)
    
'Add the first file that comes here to the array
    If RenameArrayCounter = 0 Then
        RenameArray(RenameArrayCounter) = argFileName
        NOExtArray(RenameArrayCounter) = FileNameWOEXT
        RenameArrayCounter = RenameArrayCounter + 1
        ReDim Preserve RenameArray(RenameArrayCounter)
        ReDim Preserve NOExtArray(RenameArrayCounter)
    
    Else
    
'Check to see if the file is in the array and if it is check to see if it has a higher ext.
        For x = 0 To RenameArrayCounter
            
            If NOExtArray(x) = FileNameWOEXT Then
                
                rFOUND = True
                MultiFOUND = True
                Exit Sub
                
            End If
            
           
        Next x
        
'If the file is not found yet add it and the ext to the array
        If Not rFOUND Then
            RenameArray(RenameArrayCounter) = argFileName
            NOExtArray(RenameArrayCounter) = FileNameWOEXT
            RenameArrayCounter = RenameArrayCounter + 1
            ReDim Preserve RenameArray(RenameArrayCounter)
            ReDim Preserve NOExtArray(RenameArrayCounter)
           
        End If
        
    End If

Else
    Exit Sub
End If

End Sub

Public Sub subRemExt()

Dim reFOUND As Boolean          'Value to tell whether or not the file is in the array
Dim POS As Integer              'Find the position of the "."
Dim re As Integer               'Place holder for looping thru the array
Dim ExtStr As String            'String for the extension (check to see if it is a num)
Dim ExtCheck                    'Value returned when checking to see if the ext is a number
Dim CurVal As Integer           'The current value of the revision of the file
Dim FileNameWOEXT As String     'The name of the file with out the number extension
Dim RenameFile As String        'File to be renamed
Dim fs, f, f1, fc               'File System Object variables
Dim CurFolder As String             'Place holder for current folder location
Dim ErrArray()                  'Files with errors in directory
Dim ErrArrayCounter As Long     'Counter for array
Dim ErrMsg
Dim e

reFOUND = False
re = 0
ErrArrayCounter = 0
ReDim ErrArray(ErrArrayCounter)

On Error GoTo ErrorHandler

    If RenameArrayCounter <> 0 Then
    
        For re = 0 To RenameArrayCounter
            
            RenameFile = RenameArray(re)

        
            If RenameFile <> "" Then
                                    
                POS = InStrRev(RenameFile, ".", -1)
                CurVal = Val(Right(RenameFile, Len(RenameFile) - POS))
                
                FileNameWOEXT = Left(RenameFile, POS - 1)
                
                If RenameChoice = 0 Then
                                        
                    'Rename the file to the proper extension
                    If CurVal <> RenameVal Then
                        Name RenameFile As FileNameWOEXT & "." & RenameVal
                        RenameCounter = RenameCounter + 1
                    End If
                
                'Strip the extension off of the file
                ElseIf RenameChoice = 1 Then
                    
                    'Check to see if the new file will interfere with another file
                    'if it already exists delete the old file and keep the new
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    Set f = fs.GetFolder(Left(FileNameWOEXT, (InStrRev(FileNameWOEXT, "\"))))
                    Set fc = f.Files
                        For Each f1 In fc
                        
                        If f1 = FileNameWOEXT Then
                            
                            Kill (f1)
                            
                        End If
                        Next
                        
                    Name RenameFile As FileNameWOEXT
                    RenameCounter = RenameCounter + 1
                
                End If
                
                
                
            End If
        
        Next re
       
    End If
    
    'Handle all file that can't be accessed.
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


                
'Reset the counters to speed up the process and keep the dirs independent
    RenameArrayCounter = 0
    Erase RenameArray

On Error GoTo 0

Exit Sub

ErrorHandler:
    ErrArray(ErrArrayCounter) = RenameFile
    ErrArrayCounter = ErrArrayCounter + 1
    ReDim Preserve ErrArray(ErrArrayCounter)
    
    RenameCounter = RenameCounter - 1
    
    Resume Next
   
End Sub

