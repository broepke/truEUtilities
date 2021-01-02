VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "truEUtilities - Preferences"
   ClientHeight    =   4908
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6084
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4908
   ScaleWidth      =   6084
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3708
      Index           =   0
      Left            =   6240
      ScaleHeight     =   3767.807
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.TextBox Text0 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse"
         Height          =   369
         Left            =   4440
         TabIndex        =   19
         Top             =   1518
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Move file to a selected location"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Move files to the &Recycle Bin"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Destroy files permanently"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Caption         =   "Purge Options"
         Height          =   1935
         Left            =   0
         TabIndex        =   24
         Top             =   120
         Width           =   5655
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3708
      Index           =   2
      Left            =   240
      ScaleHeight     =   3767.807
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5400
      Width           =   5685
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   495
         Left            =   2280
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear All"
         Height          =   495
         Left            =   2280
         TabIndex        =   13
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   495
         Left            =   2280
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.ListBox lstExt 
         Height          =   2160
         Left            =   360
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblDisplay 
         BorderStyle     =   1  'Fixed Single
         Height          =   252
         Left            =   2280
         TabIndex        =   16
         Top             =   3360
         Width           =   612
      End
      Begin VB.Label lblClients 
         Caption         =   "Total Extensions:"
         Height          =   252
         Left            =   360
         TabIndex        =   15
         Top             =   3360
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "&Extension to add"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3708
      Index           =   1
      Left            =   6240
      ScaleHeight     =   3767.807
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5400
      Width           =   5685
      Begin VB.OptionButton Option2 
         Caption         =   "Strip Extensions from Files"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Rename File Extensions to Number - "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   295
         Left            =   3240
         TabIndex        =   18
         Top             =   460
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rename Options"
         Height          =   1215
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   3735
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10393
      _ExtentY        =   7493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Purge"
            Object.ToolTipText     =   "Use this tab to set preferences for Purging"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Rename"
            Object.ToolTipText     =   "Change options for renaming files"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Clean"
            Object.ToolTipText     =   "Use this tab to set Junk File Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Tag             =   "&Apply"
      Top             =   4455
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" _
                                  (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                  (ByVal pidList As Long, _
                                  ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                  (ByVal lpString1 As String, ByVal _
                                  lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type


Private Sub Form_Load()

Dim JunkExt As String
Dim Position As Long
Dim counter As Long
Dim PositionOld As Long
Dim LastCommaPOS As Long
Dim LastComma As String
Dim argJunkVal As String
Dim BrowsePath As String
Dim i As Integer

'Get all user defined options from the registry
BrowsePath = GetSetting(App.Title, "UserSettings", "BrowsePath", "")
argJunkVal = GetSetting(App.Title, "UserSettings", "JunkFiles", "als,bde,bdi,bdm,bom,crc,ers,ger,inf,dat,memb,m_p,ptd,pls,")
RenameVal = GetSetting(App.Title, "UserSettings", "RenameValue", 1)
PurgeLoc = GetSetting(App.Title, "UserSettings", "PurgeLocation", 0)
RenameChoice = GetSetting(App.Title, "UserSettings", "RenameChoice", 0)

    
'Set the optionbutton to the correct one
Option1(PurgeLoc).Value = True
Text0.Text = BrowsePath
Text1.Text = RenameVal
Option2(RenameChoice).Value = True

If PurgeLoc = 0 Then
    Command1.Enabled = False
    Text0.Enabled = False
ElseIf PurgeLoc = 1 Then
    Command1.Enabled = False
    Text0.Enabled = False
Else
    Command1.Enabled = True
    Text0.Enabled = True
End If
    

ExtArrayCounter = 0
PositionOld = 0
Position = 1
counter = 0
ReDim ExtArray(ExtArrayCounter)
        
      
'Add the junk file extensions to the list box
    Do While InStr(Position, argJunkVal, ",")
        Position = InStr(Position, argJunkVal, ",") + 1
        JunkExt = Left(argJunkVal, Position - 2)
        If counter = 0 Then
        Else
            JunkExt = Right(JunkExt, Position - PositionOld)
        End If
        lstExt.AddItem JunkExt
        PositionOld = Position + 1
        counter = counter + 1
         
    Loop


   For i = 0 To picOptions.Count - 1
   With picOptions(i)
      .Move tbsOptions.ClientLeft, _
      tbsOptions.ClientTop, _
      tbsOptions.ClientWidth, _
      tbsOptions.ClientHeight
   End With
   Next i

End Sub

Private Sub cmdApply_Click()
  

    SaveSetting App.Title, "UserSettings", "JunkFiles", JunkVal
    SaveSetting App.Title, "UserSettings", "RenameValue", RenameVal
    SaveSetting App.Title, "UserSettings", "BrowsePath", BrowsePath
    SaveSetting App.Title, "UserSettings", "PurgeLocation", PurgeLoc
    SaveSetting App.Title, "UserSettings", "RenameChoice", RenameChoice
   

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    
    SaveSetting App.Title, "UserSettings", "JunkFiles", JunkVal
    SaveSetting App.Title, "UserSettings", "RenameValue", RenameVal
    SaveSetting App.Title, "UserSettings", "RenameChoice", RenameChoice
    If Option1(1).Value = True Then
        SaveSetting App.Title, "UserSettings", "BrowsePath", BrowsePath
    End If
    SaveSetting App.Title, "UserSettings", "PurgeLocation", PurgeLoc
    Unload Me
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If
End Sub

Private Sub Option2_Click(Index As Integer)
    
    'Turn off the rename value box if the radio button for
    'stripping the extensions is selected
    If Option2(1).Value = True Then
        Text1.Enabled = False
        RenameChoice = 1
    Else
        Text1.Enabled = True
        RenameChoice = 0
    End If
    
End Sub

Private Sub Text1_Change()

Dim Msg As String 'Message box for extension existing in list

    If IsNumeric(Text1.Text) = True Then
        RenameVal = Text1.Text
    Else
        Msg = MsgBox("You must enter a numeric extension only", vbExclamation)
        Text1.Text = "1"
    End If

End Sub

Private Sub Option1_Click(Index As Integer)

Dim i As Integer

    For i = 0 To Option1.Count - 1
        If Option1(i).Value = True Then
            PurgeLoc = i
        End If
    
    Next i
        
        If PurgeLoc = 0 Then
            Command1.Enabled = False
            Text0.Enabled = False
        ElseIf PurgeLoc = 1 Then
            Command1.Enabled = False
            Text0.Enabled = False
        Else
            Command1.Enabled = True
            Text0.Enabled = True
        End If

    
    

End Sub

Private Sub tbsOptions_Click()

    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 180
            picOptions(i).Top = 440
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -2000000
            picOptions(i).Top = -2000000
            picOptions(i).Enabled = False
        End If
    Next
    

End Sub

Private Sub Command1_Click()
    
    'Old way of doing it using frmBrowse
    'frmBrowse.Show vbModal, Me

'Opens a Treeview control that displays the directories in a computer

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

szTitle = "This is the title"
With tBrowseInfo
   .hWndOwner = Me.hwnd
   .lpszTitle = lstrcat(szTitle, "")
   .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With

lpIDList = SHBrowseForFolder(tBrowseInfo)

If (lpIDList) Then
   sBuffer = Space(MAX_PATH)
   SHGetPathFromIDList lpIDList, sBuffer
   sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
   BrowsePath = sBuffer
End If

    Text0.Text = BrowsePath
    
End Sub

Private Sub cmdAdd_Click()

Dim ExtFoundPos As Long 'Position of the ext in the Junk String from Registry
Dim ExtDot As Long 'Position of "." in the extension
Dim Msg As String 'Message box for extension existing in list

ExtFoundPos = 0

'Check to see if the extension already exists.
ExtFoundPos = InStr(JunkVal, txtName.Text)
If ExtFoundPos <> 0 Then
    'Display a message if the extension is in the list
    Msg = MsgBox("That extension is already in the list.", vbInformation + vbOKOnly)
    txtName.Text = ""
    Exit Sub
End If

'Check to see if there are any "." in the extension name.
ExtDot = InStr(txtName.Text, ".")
If ExtDot <> 0 Then
    'Display a message if the extension is in the list
    Msg = MsgBox("A ""."" is not required in the extension.", vbInformation + vbOKOnly)
    txtName.Text = ""
    Exit Sub
End If

    'Add the extension to the list
    lstExt.AddItem txtName.Text   ' Add to list.
    JunkVal = JunkVal & txtName.Text & "," 'concatonate the ext to the string
    txtName.Text = "" 'Clear text box.
    txtName.SetFocus 'Focus on the text box
    lblDisplay.Caption = lstExt.ListCount ' Display number.

End Sub

Private Sub cmdRemove_Click()
    
Dim Ind As Integer
Dim JunkExt As String

JunkExt = lstExt.Text

   Ind = lstExt.ListIndex   ' Get index.
   ' Make sure list item is selected.
   If Ind >= 0 Then
      ' Remove it from list box.
      lstExt.RemoveItem Ind
      
        'Find the extension and remove it from the Registry String
        Position = InStr(1, JunkVal, JunkExt)
        JunkVal = Left(JunkVal, Position - 1) & Right(JunkVal, Len(JunkVal) - Len(JunkExt) - Position)
      
      ' Display number.
      lblDisplay.Caption = lstExt.ListCount
   Else
      Beep
   End If
   ' Disable button if no entries in list.
   cmdRemove.Enabled = (lstExt.ListIndex <> -1)
End Sub

Private Sub cmdClear_Click()
    ' Clear the string for the registry save
    JunkVal = ""
    ' Empty list box.
    lstExt.Clear
    ' Disable Remove button.
    cmdRemove.Enabled = False
    ' Display number.
    lblDisplay.Caption = lstExt.ListCount
End Sub

Private Sub lstExt_Click()
   cmdRemove.Enabled = lstExt.ListIndex <> -1
End Sub

Private Sub txtName_Change()
' Enable the Add button if at least one character
' in the name.
cmdAdd.Enabled = (Len(txtName.Text) > 0)
End Sub


Private Sub VScroll1_Change()

VScroll1.LargeChange = 5
VScroll1.SmallChange = 1

VScroll1.Value = RenameVal

RenameVal = RenameVal + 1
Text1.Text = RenameVal

End Sub


