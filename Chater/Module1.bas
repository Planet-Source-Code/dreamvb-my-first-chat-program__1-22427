Attribute VB_Name = "Module1"
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Private Const HWND_NOTOPMOST = -2
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2

Public Play_Sounds As Boolean
Public LoginOk As Boolean


Public ServerPort As Integer
Public ServerIp As String
Public Nick_name As String
Public Function DrawBar(Bar As PictureBox)

Dim X, Y, Red, Blue, Green As Integer

X = Bar.ScaleWidth
Y = Bar.ScaleHeight

Red = 255
Blue = 255
Green = 255
   
Do Until Red = 1
    X = X - Bar.Width / 255
    Red = Red - 1
    Bar.Line (0, 0)-(X, Y), RGB(Red, Red, Red), BF
Loop

End Function

Public Sub FlatBorder(ByVal hwnd As Long)
  Dim TFlat As Long
  TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
  TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
  SetWindowLong hwnd, GWL_EXSTYLE, TFlat
  SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
  
End Sub
Function FileHere(LzFilename As String) As Boolean
    If Dir(LzFilename) = "" Then FileHere = False Else FileHere = True
    
End Function
Function AddSlash(LzPath As String) As String
    If Right(LzPath, 1) = "\" Then AddSlash = LzPath Else AddSlash = LzPath & "\"
    
End Function

Function CenterForm(Frm As Form)
    With Frm
        .Top = (Screen.Height - Frm.Height) / 2
        .Left = (Screen.Width - Frm.Width) / 2
    End With
    
End Function
Function MoveForm(mHwnd As Form)
    ReleaseCapture
    SendMessage mHwnd.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 1

End Function

