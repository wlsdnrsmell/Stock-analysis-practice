Attribute VB_Name = "upForm"
Option Explicit

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32" _
  (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
  ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Sub SetTopMostWindow(hWnd As Long, TopMost As Boolean)

  If TopMost = True Then
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
  Else
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
  End If
End Sub



