Attribute VB_Name = "ModHyperlink"
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation _
    As String, ByVal lpFile As String, ByVal lpParameters _
    As String, ByVal lpDirectory As String, ByVal nShowCmd _
    As Long) As Long

Global Const SW_SHOWNORMAL = 1

