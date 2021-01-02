VERSION 5.00
Begin VB.Form frmJunk 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "truEutilities - Junk File Selection"
   ClientHeight    =   3435
   ClientLeft      =   5955
   ClientTop       =   4680
   ClientWidth     =   4845
   Icon            =   "frmJunk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4845
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdJunk 
      Caption         =   "Clean Selected Files"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdToggleAll 
      Caption         =   "Toggle &All"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2085
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmJunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Add a horizontal scroll bar to the listbox
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hwnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
Public chkindex As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdToggleAll_Click()

Dim i As Double

For i = 0 To List1.ListCount - 1
    If List1.Selected(i) = True Then
        List1.Selected(i) = False
    Else
        List1.Selected(i) = True
    End If
Next i

End Sub

Private Sub cmdJunk_Click()

Dim i As Double
Dim Msg As String
Dim ErrArray()                  'Files with errors in directory
Dim ErrArrayCounter As Long     'Counter for array
Dim ErrMsg
Dim e

ErrArrayCounter = 0
ReDim ErrArray(ErrArrayCounter)

On Error GoTo ErrorHandler
    
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            Msg = List1.List(i)
            Kill (Msg)
        End If

StartHere:
    
    Next i

'If there were any files that can't be accessed - display them
If ErrArrayCounter > 0 Then
    ErrMsg = MsgBox(ErrArrayCounter & " file(s) were not accessible" & vbCrLf & _
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


Unload Me

Exit Sub

ErrorHandler:

    ErrArray(ErrArrayCounter) = Msg
    ErrArrayCounter = ErrArrayCounter + 1
    ReDim Preserve ErrArray(ErrArrayCounter)

    'removed to do mass ok for 2.4.5
    'Call ErrorHandler(Err.Number, Msg)
    
    Resume StartHere

End Sub

Private Sub Form_Load()

Dim JunkChkCnt As Long
Dim JunkFileLength As Long
Dim JunkFileLenghtMax As Long
Dim x As Double
Static y As Double

If JunkFileArrayCounter <> 0 Then
      
    For x = 0 To JunkFileArrayCounter - 1
        
        'Add to the list box from the array
        List1.AddItem (JunkFileArray(x))
        
        'Add a horizontal scroll bar to the box if needed
        If y < TextWidth(JunkFileArray(x) & "         ") / Screen.TwipsPerPixelX Then
            y = TextWidth(JunkFileArray(x) & "         ")
            If ScaleMode = vbTwips Then
                y = y / Screen.TwipsPerPixelX ' if twips change to pixels
                SendMessageByNum List1.hwnd, LB_SETHORIZONTALEXTENT, y, 0
            End If
        End If
        
        
    Next x
           
End If
   
          
End Sub
