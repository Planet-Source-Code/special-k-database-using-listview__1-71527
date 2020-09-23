Attribute VB_Name = "Module1"
Option Explicit

'PROGRAM TO RUN ONLY ONE AT A TIME
Public Const GW_HWNDPREV = 3
Public Const SW_SHOWDEFAULT = 10
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'PROGRAM TO RUN ONLY ONE AT A TIME
Sub ActivatePrevInstance()

   Dim OldTitle As String
   Dim PrevHndl As Long
   Dim result As Long
   OldTitle = App.Title
   App.Title = "unwanted instance"
   PrevHndl = FindWindow("ThunderRTMain", OldTitle)
   If PrevHndl = 0 Then
      PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
   End If
   If PrevHndl = 0 Then
        PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
   End If
   If PrevHndl = 0 Then
      Exit Sub
   End If
   PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
   result = OpenIcon(PrevHndl)
   result = SetForegroundWindow(PrevHndl)
   End

End Sub

'CHILD FORM TO CENTER IN MDIFORM
Sub CenterForm(frm As Form)
        
        frm.Left = (Screen.Width - frm.Width) \ 2
           
        frm.Top = (Screen.Height - frm.Height) \ 2
    
End Sub

