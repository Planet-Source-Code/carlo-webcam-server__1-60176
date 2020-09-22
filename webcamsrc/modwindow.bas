Attribute VB_Name = "modwindow"
' window module
' ---------------
' when called, ActiveCaption returns the entire caption of the
' foreground window

Option Explicit
Global fCaption As String
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" _
    () As Long
   
Function ActiveCaption() As String

    Dim sText As String * 255
    Dim Ret As Long
    
    Ret = GetWindowText(GetForegroundWindow, sText, 255)
    ActiveCaption = Left(sText, InStr(1, sText, Chr(0)) - 1) ' strip caption out of buffer

End Function



