VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6405
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   8985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "ChatMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   8985
      TabIndex        =   12
      Top             =   0
      Width           =   8985
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dreams Chat House Beta 2 Build 2.00.1456"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   90
         TabIndex        =   13
         Top             =   30
         Width           =   3795
      End
   End
   Begin Project1.Bevel Bevel2 
      Height          =   270
      Left            =   2205
      TabIndex        =   10
      Top             =   5460
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   476
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H80000018&
      Caption         =   "Option1"
      Height          =   210
      Left            =   1710
      TabIndex        =   7
      Top             =   5790
      Width           =   210
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000018&
      Caption         =   "Option1"
      Height          =   210
      Left            =   1410
      TabIndex        =   6
      Top             =   5790
      Width           =   210
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000018&
      Caption         =   "Option1"
      Height          =   210
      Left            =   1050
      TabIndex        =   5
      Top             =   5790
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000018&
      Caption         =   "Option1"
      Height          =   210
      Left            =   720
      TabIndex        =   4
      Top             =   5790
      Width           =   240
   End
   Begin Project1.Bevel Bevel1 
      Height          =   3990
      Left            =   75
      TabIndex        =   3
      Top             =   780
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7038
   End
   Begin SHDocVwCtl.WebBrowser ChatWin 
      Height          =   4005
      Left            =   60
      TabIndex        =   2
      Top             =   765
      Width           =   8850
      ExtentX         =   15610
      ExtentY         =   7064
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINNT\system32\shdoclc.dll/dnserror.htm#http:///"
   End
   Begin VB.ComboBox cboMesstype 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2190
      TabIndex        =   1
      Top             =   5445
      Width           =   1740
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7410
      Top             =   6645
   End
   Begin MSWinsockLib.Winsock IChat 
      Left            =   7920
      Top             =   6660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox ChatSend 
      Height          =   345
      Left            =   45
      MaxLength       =   255
      TabIndex        =   0
      Top             =   4860
      Width           =   5115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text Formatting >"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   2
      Left            =   4365
      TabIndex        =   16
      Top             =   5310
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   0
      X2              =   1125
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      Index           =   2
      X1              =   0
      X2              =   1125
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "&Options"
      Height          =   195
      Left            =   525
      TabIndex        =   15
      Top             =   420
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&File"
      Height          =   195
      Left            =   75
      TabIndex        =   14
      Top             =   420
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      Index           =   1
      X1              =   -90
      X2              =   1035
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   -90
      X2              =   1035
      Y1              =   345
      Y2              =   345
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   2
      Left            =   5070
      Picture         =   "ChatMain.frx":0442
      Top             =   5595
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   1
      Left            =   4770
      Picture         =   "ChatMain.frx":04D4
      Top             =   5595
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   285
      Index           =   0
      Left            =   4470
      Picture         =   "ChatMain.frx":055E
      Top             =   5595
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Messages"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   1
      Left            =   2325
      TabIndex        =   11
      Top             =   5805
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status >"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   5610
      Width           =   600
   End
   Begin VB.Image imgclear 
      Height          =   285
      Index           =   2
      Left            =   1590
      Picture         =   "ChatMain.frx":05EE
      Top             =   7080
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Image imgclear 
      Height          =   285
      Index           =   1
      Left            =   1590
      Picture         =   "ChatMain.frx":075D
      Top             =   6855
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Image imgclear 
      Height          =   285
      Index           =   0
      Left            =   7110
      Picture         =   "ChatMain.frx":08CA
      Top             =   4890
      Width           =   1560
   End
   Begin VB.Image imgsend 
      Height          =   285
      Index           =   2
      Left            =   -15
      Picture         =   "ChatMain.frx":0A37
      Top             =   6840
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Image imgsend 
      Height          =   270
      Index           =   1
      Left            =   0
      Picture         =   "ChatMain.frx":0BA2
      Top             =   7020
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Image imgsend 
      Height          =   285
      Index           =   0
      Left            =   5370
      Picture         =   "ChatMain.frx":0D0B
      Top             =   4890
      Width           =   1560
   End
   Begin VB.Image img4 
      Height          =   225
      Left            =   1710
      Picture         =   "ChatMain.frx":0E76
      Top             =   5460
      Width           =   225
   End
   Begin VB.Image img3 
      Height          =   225
      Left            =   1395
      Picture         =   "ChatMain.frx":0FF4
      Top             =   5460
      Width           =   225
   End
   Begin VB.Image img2 
      Height          =   225
      Left            =   1050
      Picture         =   "ChatMain.frx":116A
      Top             =   5460
      Width           =   225
   End
   Begin VB.Image img1 
      Height          =   225
      Left            =   720
      Picture         =   "ChatMain.frx":12F0
      Top             =   5460
      Width           =   225
   End
   Begin VB.Label lblbar 
      BackColor       =   &H80000018&
      Height          =   855
      Left            =   60
      TabIndex        =   8
      Top             =   5280
      Width           =   8790
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UserNum As String
Public Sub KillHtm()
On Error Resume Next
    Kill AddSlash(App.Path) & "chat.htm"
    If Err Then Err.Clear
    
End Sub
Sub Send()
Dim StrSend_text As String
Dim StrText As String
Dim HLine As String

    imgsend(0).Picture = imgsend(1).Picture
    HLine = "<hr size=""0"" noshade>"
    StrSend_text = ChatSend.Text
    
    If Not IChat.State = sckConnected Then
        MsgBox "Your not conected please try agian", vbInformation
        ChatSend.Text = ""
        Exit Sub
    Else
        If Len(ChatSend.Text) = 0 Then
            Exit Sub
        Else

            StrText = UCase(Left(ChatSend.Text, 6))
            If StrText = "\CLEAR" Then
                Kill AddSlash(App.Path) & "chat.htm"
                ChatWin.Navigate "about:"
                ChatSend.Text = ""
                Exit Sub
            End If
    '///////////////////////////////////////////////////////////////////////////
    If InStr(StrSend_text, ":)") Then
        a = "<img src=""images/icon10.gif"" width=""15"" height=""15"">"
        StrSend_text = Replace(StrSend_text, ":)", a)
        a = ""
    '///////////////////////////////////////////////////////////////////////////
    ElseIf InStr(StrSend_text, ":(") Then
        a = "<img src=""images/icon9.gif"" width=""15"" height=""15"">"
        StrSend_text = Replace(StrSend_text, ":(", a)
    '///////////////////////////////////////////////////////////////////////////
    ElseIf InStr(StrSend_text, "?") Then
        a = "<img src=""images/icon5.gif"" width=""15"" height=""15"">"
        StrSend_text = Replace(StrSend_text, "?", a)
    '///////////////////////////////////////////////////////////////////////////
    ElseIf InStr(StrSend_text, ":cool:") Then
        a = "<img src=""images/icon6.gif"" width=""15"" height=""15"">"
        StrSend_text = Replace(StrSend_text, ":cool:", a)
    End If
            IChat.SendData "105 " & "<b>" & Module1.Nick_name & "</b>" & " > " & StrSend_text & HLine
            ChatSend.Text = ""
            If Play_Sounds = False Then
            Exit Sub
        Else
            sndPlaySound AddSlash(App.Path) & "send.wav", SND_ASYNC Or SND_NODEFAULT
        End If
        End If
    End If
    a = ""
    HLine = ""
    StrSend_text = ""

End Sub
Private Sub cboMesstype_Click()
    If IChat.State = sckConnected Then
        IChat.SendData "106" & "<b>" & Module1.Nick_name & " > " & "</b>" & cboMesstype.Text
    Else
        cboMesstype.Enabled = False
        
    End If

End Sub

Private Sub ChatSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
    
End Sub


Private Sub ChatSend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Send
        imgsend(0).Picture = imgsend(2).Picture
        
    End If
    
End Sub

Private Sub Command1_Click()
    frmenu.Show
    frmenu.Move Command1.Left + Command1.Width + 290, Command1.Top + 1500
    
    
    
    
End Sub



Private Sub Form_Load()
On Error Resume Next
    KillHtm
    Module1.ServerIp = "127.0.0.1"
    Module1.ServerPort = 69
    Module1.Nick_name = "Ben"
    mnuDis.Enabled = False
    CenterForm Form1
    Play_Sounds = True
    
    cboMesstype.AddItem "I'll be right back"
    cboMesstype.AddItem "Back Now"
    cboMesstype.AddItem "Busy at the moment"
    cboMesstype.AddItem "On the phone"
    cboMesstype.AddItem "Someone at the door"
    cboMesstype.AddItem "Getting a Drink"
    cboMesstype.AddItem "Auto Away"
    cboMesstype.AddItem "Going know"
    cboMesstype.AddItem "Not at my Computer"
    cboMesstype.AddItem "Left the Room"
    Bevel3.BevelSytle VbRaised
    
    
    Module1.FlatBorder ChatSend.hwnd
    ChatWin.Navigate "about:"
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = vbBlack
    Label5.ForeColor = vbBlack

End Sub

Private Sub Form_Resize()
'    Bevel3.BevelSytle VbLowered
    DrawBar Picture1
    Line1(0).X2 = Form1.ScaleWidth
    Line1(1).X2 = Form1.ScaleWidth
    Line1(2).X2 = Form1.ScaleWidth
    Line1(3).X2 = Form1.ScaleWidth
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    KillHtm
    If Err Then Err.Clear
    
End Sub

Private Sub IChat_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
On Error Resume Next
    IChat.GetData Data, vbString
    a = "<img src=""images/icon6.gif"" width=""15"" height=""15"">"
    Data = Replace(Data, ":cool:", a)
    
    If Left(Data, 3) = "101" Then
    If Play_Sounds Then
            sndPlaySound AddSlash(App.Path) & "login.wav", SND_ASYNC Or SND_NODEFAULT
    End If
        cboMesstype.Enabled = True
        MsgBox Trim(Right(Data, Len(Data) - 3)), vbInformation
        Data = ""
        Exit Sub
    End If
    
    If Left(Data, 3) = "100" Then
        SetCursorPos 100, 100
        Exit Sub
    End If

    Open AddSlash(App.Path) & "chat.htm" For Append As #1
        Print #1, Data & "<p>" & vbCrLf;
        ChatWin.Navigate AddSlash(App.Path) & "chat.htm"
    Close #1

    Data = ""

    
    
End Sub

Private Sub IChat_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Select Case Number
        Case 10061
            sndPlaySound AddSlash(App.Path) & "error.wav", SND_ASYNC Or SND_NODEFAULT
            MsgBox Description, vbInformation
            frmenu.mnuDis.Enabled = False
            frmenu.mnuCon.Enabled = True
            IChat.Close
    End Select
    
    
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IChat.State = sckConnected Then
        MsgBox "Your not conected please try agian", vbInformation
        ChatSend.Text = ""
        Exit Sub
    Else
        Image1(Index).BorderStyle = 1
    End If
    
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1(Index).BorderStyle = 0
    Select Case Index
        Case 0
            If Len(ChatSend.Text) < 1 Then
                Exit Sub
            Else
                ChatSend.SelText = "<b>" & ChatSend.SelText & "</b>"
            End If
            '
        Case 1
            If Len(ChatSend.Text) < 1 Then
                Exit Sub
            Else
                ChatSend.SelText = "<i>" & ChatSend.SelText & "</i>"
            End If
        '
        Case 2
            If Len(ChatSend.Text) < 1 Then
                Exit Sub
            Else
                ChatSend.SelText = "<u>" & ChatSend.SelText & "</u>"
            End If

    End Select
    
End Sub

Private Sub imgclear_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IChat.State = sckConnected Then
        MsgBox "Your not conected please try agian", vbInformation
        ChatSend.Text = ""
        Exit Sub
    Else
        imgclear(0).Picture = imgclear(2).Picture
    End If
    
End Sub

Private Sub imgclear_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    imgclear(0).Picture = imgclear(1).Picture
    Kill AddSlash(App.Path) & "chat.htm"
    ChatWin.Navigate "about:"
    ChatSend.Text = ""

End Sub

Private Sub imgsend_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Send
    
End Sub

Private Sub imgsend_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgsend(0).Picture = imgsend(2).Picture

End Sub








Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        MoveForm Form1
    End If
    
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = vbBlack
    Label5.ForeColor = vbBlack
    
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = vbRed
    Label5.ForeColor = vbBlack
    
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PopupMenu frmenu.mnufile, , Label4.Left - 10, Label4.Top + Label4.Height
    
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.ForeColor = vbRed
    Label4.ForeColor = vbBlack
    
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PopupMenu frmenu.mnuOp, , Label5.Left, Label5.Top + Label5.Height

End Sub

Private Sub Option1_Click()
    If Not IChat.State = sckConnected Then
        MsgBox "Your not conected please try agian", vbInformation
        ChatSend.Text = ""
        Exit Sub
    Else
        ChatSend.SelText = " " & ":)" & " "
    End If
    
End Sub

Private Sub Option2_Click()
    If Not IChat.State = sckConnected Then
        MsgBox "Your not conected please try agian", vbInformation
        ChatSend.Text = ""
        Exit Sub
    Else
        ChatSend.SelText = " " & ":(" & " "
    End If
    
End Sub

Private Sub Option3_Click()
    If Not IChat.State = sckConnected Then
        MsgBox "Your not conected please try agian", vbInformation
        ChatSend.Text = ""
        Exit Sub
    Else
        ChatSend.SelText = " " & "?" & " "
    End If
    
End Sub

Private Sub Option4_Click()
    If Not IChat.State = sckConnected Then
        MsgBox "Your not conected please try agian", vbInformation
        ChatSend.Text = ""
        Exit Sub
    Else
        ChatSend.SelText = " " & ":cool:" & " "
    End If
    
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        MoveForm Form1
    End If
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = vbBlack
    Label5.ForeColor = vbBlack

End Sub

Private Sub Timer1_Timer()
    If IChat.State = sckConnected Then
        IChat.SendData "103" & Nick_name
        Timer1.Enabled = False
    End If
    
End Sub
