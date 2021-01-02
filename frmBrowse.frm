VERSION 5.00
Begin VB.Form frmBrowse 
   Caption         =   "truEUtilities - Browse"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   2760
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Get window position information from the Registry
    Me.Left = GetSetting(App.Title, "Settings", "BrowseLeft", 5500)
    Me.Top = GetSetting(App.Title, "Settings", "BrowseTop", 4750)
    Me.Width = GetSetting(App.Title, "Settings", "BrowseWidth", 4800)
    Me.Height = GetSetting(App.Title, "Settings", "BrowseHeight", 4250)
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    'Save the browsepath location to "BrowsePath"
    BrowsePath = Dir1.Path
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    File1.Refresh
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Unload(Cancel As Integer)
        'Save all window postions settings to the registry
        SaveSetting App.Title, "Settings", "BrowseLeft", Me.Left
        SaveSetting App.Title, "Settings", "BrowseTop", Me.Top
        SaveSetting App.Title, "Settings", "BrowseWidth", Me.Width
        SaveSetting App.Title, "Settings", "BrowseHeight", Me.Height

End Sub
