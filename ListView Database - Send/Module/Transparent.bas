Attribute VB_Name = "Transparent"
Option Explicit

'FORM ENTER FADE
Public Const LWA_ALPHA = 2
Public Const LWA_COLORKEY = 1
Public Const LWA_BOTH = 3
Public Const GWL_EXSTYLE = -20
Public Const WS_EX_LAYERED = &H80000
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal color As Long, ByVal x As Byte, ByVal alpha As Long) As Boolean

'FORM ENTER FADE
Public Sub ofFrm(hWnd As Long, Transval As Integer)
    Dim gt As Long
    gt = GetWindowLong(hWnd, GWL_EXSTYLE)
    SetWindowLong hWnd, GWL_EXSTYLE, gt Or WS_EX_LAYERED
    SetLayeredWindowAttributes hWnd, RGB(255, 255, 0), Transval, LWA_ALPHA
    Exit Sub
End Sub
