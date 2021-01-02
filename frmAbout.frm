VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "truEUtilities - About "
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5865
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About Project1"
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2265
      Left            =   120
      Picture         =   "frmAbout.frx":0A02
      ScaleHeight     =   2265
      ScaleMode       =   0  'User
      ScaleWidth      =   1125
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2520
      Width           =   1467
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4200
      TabIndex        =   1
      Tag             =   "&System Info..."
      Top             =   3000
      Width           =   1452
   End
   Begin VB.Label lblFileDescription 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label lblMail 
      Caption         =   "support@trueinnovations.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1410
      MouseIcon       =   "frmAbout.frx":90C0
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1920
      Width           =   3735
   End
   Begin VB.Label lblwarning 
      Caption         =   "Warning..."
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Tag             =   "Warning"
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label lblWeb 
      Caption         =   "http://www.trueinnovations.com"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1410
      MouseIcon       =   "frmAbout.frx":9212
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "App Description"
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label lblTitle 
      Caption         =   "Title..."
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1410
      TabIndex        =   4
      Tag             =   "Application Title"
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version..."
      Height          =   225
      Left            =   1410
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   780
      Width           =   3735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Reg Key Security Options...
Const KEY_ALL_ACCESS = &H2003F
                                          

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.CompanyName & " " & App.Title & vbCrLf & "For 32 bit Windows operating systems"
    lblFileDescription = App.FileDescription
    lblwarning = App.LegalCopyright
    lblContact = App.Comments
End Sub

Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
        Unload Me
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' Try To Get System Info Program Path\Name From Registry...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Validate Existance Of Known 32 Bit File Version
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Error - File Can Not Be Found...
                Else
                        GoTo SysInfoErr
                End If
        ' Error - Registry Entry Can Not Be Found...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Loop Counter
        Dim rc As Long                                          ' Return Code
        Dim hKey As Long                                        ' Handle To An Open Registry Key
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Data Type Of A Registry Key
        Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
        Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
        '------------------------------------------------------------
        ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
        

        tmpVal = String$(1024, 0)                             ' Allocate Variable Space
        KeyValSize = 1024                                       ' Mark Variable Size
        

        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Search Data Types...
        Case REG_SZ                                             ' String Registry Key Data Type
                KeyVal = tmpVal                                     ' Copy String Value
        Case REG_DWORD                                          ' Double Word Registry Key Data Type
                For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
        End Select
        

        GetKeyValue = True                                      ' Return Success
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
        Exit Function                                           ' Exit
        

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub lblMail_Click()

On Error GoTo lblMail_Click_Error
Dim StartDoc As Long
Dim SiteURL As String

SiteURL = "mailto:support@trueinnovations.com"

            If Not IsNull(SiteURL) Then
              StartDoc = ShellExecute(Me.hwnd, "open", SiteURL, _
                "", "C:\", SW_SHOWNORMAL)
                End If
Exit Sub
lblMail_Click_Error:
    MsgBox "Error: " & Err & " " & Error
Exit Sub

End Sub

Private Sub lblWeb_Click()

On Error GoTo lblWeb_Click_Error
Dim StartDoc As Long
Dim SiteURL As String

SiteURL = "http://www.trueinnovations.com"

            If Not IsNull(SiteURL) Then
              StartDoc = ShellExecute(Me.hwnd, "open", SiteURL, _
                "", "C:\", SW_SHOWNORMAL)
                End If
Exit Sub
lblWeb_Click_Error:
    MsgBox "Error: " & Err & " " & Error
Exit Sub

End Sub
