VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " truEUtilities - Main"
   ClientHeight    =   4425
   ClientLeft      =   5295
   ClientTop       =   3705
   ClientWidth     =   6405
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   4200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton btnPurge 
      Caption         =   "truEPurge"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton btnJunk 
      Caption         =   "truEClean..."
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "truEExit"
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton btnRename 
      Caption         =   "truERename"
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CheckBox chkPurgeRec 
      Caption         =   "Sub Folders"
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkRename 
      Caption         =   "Sub Folders"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkJunk 
      Caption         =   "Sub Folders"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3120
      TabIndex        =   0
      Top             =   2100
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Drives:"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   1838
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Folders:"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "File Names:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuPreferences 
         Caption         =   "&Preferences..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnutruEWeb 
         Caption         =   "truEInnovations on the &Web..."
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuFileBox 
      Caption         =   "File Box Menu"
      Visible         =   0   'False
      Begin VB.Menu Open 
         Caption         =   "Open with Editor"
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" _
    (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
    
Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnJunk_Click()

'Start the junk file operation

Dim LastChar As String      'Last character in the string (checking for root dir)
Dim NewPathLen As String    'New path length if the initial dir is a root dir
Dim NewPath As String       'New path based on the NewPathLen value returned for root dirs
Dim initPath As String      'Initial path to purge
Dim x As Long
Dim JunkMsg As String
Dim JunkFile As String
Dim Response

Erase ExtArray
ExtArrayCounter = 0
Call JunkArray(JunkVal)

'sbStatusBar.SimpleText = "Finding files to clean..."
'sbStatusBar.Refresh    ' You must refresh to see the Simple text.

    If chkJunk = 1 Then
        
        'Find all the subdirectories
        Call subRecurse(Dir1.Path)
        
        'Progress bar control
        ProgressBar1.Min = LBound(DirArray)
        ProgressBar1.Max = UBound(DirArray)
        ProgressBar1.Visible = True
        
        'Set the Progress's Value to Min.
        ProgressBar1.Value = ProgressBar1.Min
        
        'Call routine to find the junk files in the array of subdirs
        Call subFindJunk(DirArray, DirArrayCounter)
    
    Else
    
        Erase DirArray
        DirArrayCounter = 0
        ReDim DirArray(counter)
        
        initPath = Dir1.Path
        DirArray(counter) = initPath
        counter = counter + 1
        
        ReDim Preserve DirArray(counter)
        Call subFindJunk(DirArray, DirArrayCounter)
    
    End If
    
    If JunkFileArrayCounter <> 0 Then
        frmJunk.Show vbModal, Me
    Else
        MsgBox "There were no files to clean.", vbInformation
    End If
    
    'Do not display the progress bar
    ProgressBar1.Visible = False
    
    File1.Refresh
    'sbStatusBar.SimpleText = ""
           
End Sub

Private Sub btnPurge_Click()

Dim LastChar As String      'Last character in the string (checking for root dir)
Dim NewPathLen As String    'New path length if the initial dir is a root dir
Dim NewPath As String       'New path based on the NewPathLen value returned for root dirs
Dim initPath As String      'Initial path to purge
Dim PurgeWORD As String     'Verbage for deleting files (msgbox)


' Show the status in the status bar
'sbStatusBar.SimpleText = "Purging files..."
'sbStatusBar.Refresh


    ' Code to recurse all sub directories and run the purge sub
    If chkPurgeRec = 1 Then
    
            ' Set the counters to 0
            Erase FileArray
            Erase DirArray
            counter = 0
            KillCounter = 0
            DirArrayCounter = 0
            FileArrayCounter = 0
            
            Call subRecurse(Dir1.Path)
            
            'Progress bar control
            ProgressBar1.Min = LBound(DirArray)
            ProgressBar1.Max = UBound(DirArray)
            ProgressBar1.Visible = True
            
            'Set the Progress's Value to Min.
            ProgressBar1.Value = ProgressBar1.Min

            Call subPurge(DirArray, DirArrayCounter)
                           
        Else


            ' Set the counters to 0
            Erase FileArray
            Erase DirArray
            counter = 0
            KillCounter = 0
            DirArrayCounter = 0
            FileArrayCounter = 0
            ReDim DirArray(counter)
            
            initPath = Dir1.Path
            
            DirArray(counter) = initPath
            counter = counter + 1
            
            ReDim Preserve DirArray(counter)
            Call subPurge(DirArray, DirArrayCounter)
            
        End If
        
    'Do not display the progress bar
    ProgressBar1.Visible = False
        
    File1.Refresh
    'sbStatusBar.SimpleText = ""
    
    If PurgeLoc = 0 Then
        PurgeWORD = "deleted"
    Else
        PurgeWORD = "removed"
    End If
           
    If DirArrayCounter <= 1 And KillCounter > 1 Then
        MsgBox KillCounter & " Files were " & PurgeWORD & " from your directory.", vbInformation
    ElseIf DirArrayCounter <= 1 And KillCounter = 0 Then
        MsgBox KillCounter & " Files were " & PurgeWORD & " from your directory.", vbInformation
    ElseIf DirArrayCounter <= 1 And KillCounter = 1 Then
        MsgBox KillCounter & " File was " & PurgeWORD & " from your directory.", vbInformation
    ElseIf DirArrayCounter > 1 And KillCounter > 1 Then
        MsgBox KillCounter & " Files were " & PurgeWORD & " from " & DirArrayCounter & " directories." _
            , vbInformation
    ElseIf DirArrayCounter > 1 And KillCounter = 0 Then
        MsgBox KillCounter & " Files were " & PurgeWORD & " from " & DirArrayCounter & " directories." _
            , vbInformation
    ElseIf DirArrayCounter > 1 And KillCounter = 1 Then
        MsgBox KillCounter & " File was " & PurgeWORD & " from " & DirArrayCounter & " directories." _
            , vbInformation
    Else
        MsgBox KillCounter & " File was " & PurgeWORD & " from your directory", vbInformation
    End If
    

End Sub

Private Sub btnRename_Click()

Dim LastChar As String      'Last character in the string (checking for root dir)
Dim NewPathLen As String    'New path length if the initial dir is a root dir
Dim NewPath As String       'New path based on the NewPathLen value returned for root dirs
Dim initPath As String      'Initial path to purge


If RenameChoice <> 0 Then
    textMsg = "You have selected to remove all of the numeric extensions." & vbCrLf & _
    "Are you sure you want to do this?"
    Msg = MsgBox(textMsg, vbOKCancel + vbQuestion)
        If Msg = 2 Then
            Exit Sub
        Else
            GoTo StartRenameHere:
        End If
Else
    GoTo StartRenameHere:
End If
        
StartRenameHere:

    If chkRename = 1 Then
    
            Erase RenameArray
            Erase DirArray
            DirArrayCounter = 0
            RenameCounter = 0
            RenameArrayCounter = 0
            counter = 0
        
            Call subRecurse(Dir1.Path)
            
            'Progress bar control
            ProgressBar1.Min = LBound(DirArray)
            ProgressBar1.Max = UBound(DirArray)
            ProgressBar1.Visible = True
            
            'Set the Progress's Value to Min.
            ProgressBar1.Value = ProgressBar1.Min
                
            Call subRename(DirArray, DirArrayCounter)
                           
        Else

            Erase RenameArray
            Erase DirArray
            DirArrayCounter = 0
            RenameCounter = 0
            RenameArrayCounter = 0
            counter = 0
            ReDim DirArray(counter)

            initPath = Dir1.Path
            DirArray(counter) = initPath
            counter = counter + 1
            ReDim Preserve DirArray(counter)
            Call subRename(DirArray, DirArrayCounter)
            
        End If
    
    'Do not display the progress bar
    ProgressBar1.Visible = False
    
    File1.Refresh
    


        If DirArrayCounter <= 1 And RenameCounter > 1 Then
            MsgBox RenameCounter & " Files were renamed in your directory.", vbInformation
        ElseIf DirArrayCounter <= 1 And RenameCounter = 0 Then
            MsgBox RenameCounter & " Files were renamed in your directory.", vbInformation
        ElseIf DirArrayCounter <= 1 And RenameCounter = 1 Then
            MsgBox RenameCounter & " File was renamed in your directory.", vbInformation
    
        ElseIf DirArrayCounter > 1 And RenameCounter > 1 Then
            MsgBox RenameCounter & " Files were renamed in " & DirArrayCounter & " directories." _
                , vbInformation
        ElseIf DirArrayCounter > 1 And RenameCounter = 0 Then
            MsgBox RenameCounter & " Files were renamed in " & DirArrayCounter & " directories." _
                , vbInformation
        ElseIf DirArrayCounter > 1 And RenameCounter = 1 Then
            MsgBox RenameCounter & " File was renamed in " & DirArrayCounter & " directories." _
                , vbInformation
        Else
            MsgBox RenameCounter & " File was renamed in your directory", vbInformation
        End If


End Sub

Private Sub Dir1_Change()
    
    File1.Path = Dir1.Path
    
    Label3.Caption = DirPath(Dir1.Path, Label3.Width, frmMain, 0)
    
    File1.Refresh
    
End Sub

Private Sub Drive1_Change()

On Error GoTo ErrorHandler

    ChDrive Drive1.Drive
    Dir1.Path = Drive1.Drive
    
Exit Sub
    
ErrorHandler:
    Call ErrorHandler(Err.Number, Drive1.Drive)
    Exit Sub
       
End Sub

'Private Sub File1_MouseUp(Button As Integer, Shift As _
'    Integer, x As Single, y As Single)
    ' Check if right mouse button was clicked.
'    If Button = 2 Then
    ' Display the File menu as a pop-up menu.
'    PopupMenu mnuFileBox
'   End If
'End Sub



Private Sub Form_Load()
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 5125)
    JunkVal = GetSetting(App.Title, "UserSettings", "JunkFiles", "als,bde,bdi,bdm,bom,crc,ers,ger,inf,dat,memb,m_p,ptd,pls,")
    RenameVal = GetSetting(App.Title, "UserSettings", "RenameValue", 1)
    PurgeLoc = GetSetting(App.Title, "UserSettings", "PurgeLocation", 0)
    chkPurgeRec.Value = GetSetting(App.Title, "Usersettings", "RecursePurge", 0)
    chkRename.Value = GetSetting(App.Title, "Usersettings", "RecurseRename", 0)
    chkJunk.Value = GetSetting(App.Title, "Usersettings", "RecurseClean", 0)
    RenameChoice = GetSetting(App.Title, "UserSettings", "RenameChoice", 0)
    RenameChoice = 0
    
    'Progress bar declarations
    ProgressBar1.Align = vbAlignBottom
    ProgressBar1.Visible = False
    
    Label3.Caption = DirPath(Dir1.Path, Label3.Width, frmMain, 0)
        
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
        
    RenameChoice = 0
    
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
        SaveSetting App.Title, "UserSettings", "JunkFiles", JunkVal
        SaveSetting App.Title, "UserSettings", "RenameValue", RenameVal
        SaveSetting App.Title, "UserSettings", "PurgeLocation", PurgeLoc
        SaveSetting App.Title, "UserSettings", "WorkingDirectory", Dir1.Path
        SaveSetting App.Title, "UserSettings", "RecursePurge", chkPurgeRec.Value
        SaveSetting App.Title, "UserSettings", "RecurseRename", chkRename.Value
        SaveSetting App.Title, "UserSettings", "RecurseClean", chkJunk.Value
        SaveSetting App.Title, "UserSettings", "RenameChoice", RenameChoice
    End If
    
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub



Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuPreferences_Click()
    frmOptions.Show vbModal, Me
End Sub


Private Sub mnutruEWeb_Click()

On Error GoTo lblWeb_Click_Error
Dim StartDoc As Long
Dim SiteURL As String

SiteURL = "http://www.trueinnovations.com"

            If Not IsNull(SiteURL) Then
              StartDoc = ShellExecute(Me.hwnd, "open", SiteURL, "", "C:\", SW_SHOWNORMAL)
                End If
Exit Sub
lblWeb_Click_Error:
    MsgBox "Error: " & Err & " " & Error
    
Exit Sub

End Sub

Private Sub Open_Click()

Call StartDoc(Dir1.Path & "\" & File1.FileName)

End Sub
