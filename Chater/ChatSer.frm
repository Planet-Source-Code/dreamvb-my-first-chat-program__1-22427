VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My First Chat Program Beta 2"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Enabled         =   0   'False
      Height          =   1230
      Left            =   3060
      TabIndex        =   12
      Top             =   2595
      Width           =   1845
   End
   Begin MSWinsockLib.Winsock IpChat 
      Index           =   0
      Left            =   4095
      Top             =   4710
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   5100
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   5100
      Begin VB.CommandButton Command1 
         Caption         =   "&Exit"
         Height          =   330
         Left            =   3060
         TabIndex        =   10
         Top             =   4500
         Width           =   1425
      End
      Begin VB.ListBox List2 
         Enabled         =   0   'False
         Height          =   1230
         Left            =   1665
         TabIndex        =   9
         Top             =   2670
         Width           =   1320
      End
      Begin VB.TextBox Text2 
         Height          =   1335
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Text            =   "ChatSer.frx":0000
         Top             =   960
         Width           =   4830
      End
      Begin VB.CommandButton cmdSStop 
         Caption         =   "S&top Server"
         Height          =   330
         Left            =   1605
         TabIndex        =   6
         Top             =   4500
         Width           =   1425
      End
      Begin VB.CommandButton cmdSStart 
         Caption         =   "&Start Server"
         Height          =   330
         Left            =   165
         TabIndex        =   5
         Top             =   4500
         Width           =   1425
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   105
         TabIndex        =   4
         Top             =   2670
         Width           =   1485
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   585
         TabIndex        =   2
         Text            =   "69"
         Top             =   330
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Users"
         Height          =   195
         Index           =   1
         Left            =   2025
         TabIndex        =   11
         Top             =   2355
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Welcome Message"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   705
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Current Conections"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   2355
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         TabIndex        =   1
         Top             =   345
         Width           =   360
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UCount As Integer
Dim mIndex As Integer
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Const LB_FINDSTRINGEXACT = &H1A2
Sub LeftRoom(Nickname As String)
  For Cnt = 1 To IpChat.UBound
            If IpChat(Cnt).State = sckConnected Then
                IpChat(Cnt).SendData Nickname & " has left the room"
            End If
            DoEvents
        Next
        Cnt = 0
        
End Sub
Sub Messtype(StrMess As String)
  For Cnt = 1 To IpChat.UBound
            If IpChat(Cnt).State = sckConnected Then
                IpChat(Cnt).SendData StrMess
            End If
            DoEvents
        Next
        Cnt = 0
        
End Sub

Function RemoveDupeItems(LstBox As ListBox) As Integer
Dim Counter As Integer
Dim StrLine As String
Dim XPos As Integer
Dim YPos As Integer

For Counter = 0 To LstBox.ListCount - 1
Do
    DoEvents
    StrLine = LstBox.List(Counter)
    XPos = SendMessageByString(LstBox.hWnd, LB_FINDSTRINGEXACT, Counter, StrLine)
    YPos = SendMessageByString(LstBox.hWnd, LB_FINDSTRINGEXACT, XPos + 1, StrLine)
        If YPos = -1 Or YPos = XPos Then Exit Do
            LstBox.RemoveItem YPos
Loop
Next Counter

    End Function

Private Sub cmdSStart_Click()
    If Len(txtPort.Text) = 0 Then
        MsgBox "You Port Number has been Entered please do so now", vbInformation
        Exit Sub
    Else
        cmdSStart.Enabled = False
        cmdSStop.Enabled = True
        Form1.Caption = "Server started now serving on port " & txtPort.Text
        IpChat(0).Close
        IpChat(0).LocalPort = Val(txtPort.Text)
        IpChat(0).Listen
    End If
End Sub

Private Sub cmdSStop_Click()
  For Cnt = 1 To IpChat.UBound
            If IpChat(Cnt).State = sckConnected Then
                IpChat(Cnt).SendData "Server shutdown timed at " & Time
            End If
            DoEvents
        Next
        cmdSStart.Enabled = True
        cmdSStop.Enabled = False
        Form1.Caption = "Server stoped"
        IpChat(0).Close
    
End Sub







Private Sub Command1_Click()
    Unload Form1
    End
    
End Sub

Private Sub Form_Load()
    Form1.Caption = "Your IP is  [" & IpChat(0).LocalIP & "]"
    List3.AddItem "Funny Mouse"
    List3.AddItem "Dialog message"
    
End Sub

Private Sub IpChat_Close(Index As Integer)
    IpChat(Index).Close
    
End Sub

Private Sub IpChat_ConnectionRequest(Index As Integer, ByVal requestID As Long)
 On Error Resume Next
    UCount = UCount + 1
    Load IpChat(UCount)
    'IpChat(UCount).LocalPort = 69
    IpChat(UCount).Accept requestID
    List1.AddItem IpChat(UCount).RemoteHostIP
    RemoveDupeItems List1
    IpChat(UCount).SendData Text2.Text
    If Err Then Err.Clear
    
End Sub

Private Sub IpChat_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim User As String
Dim Data As String
Dim Cnt As Integer
On Error Resume Next
    Data = ""
    IpChat(Index).GetData Data, vbString
    If Left(Data, 3) = "108" Then LeftRoom Trim((Right(Data, Len(Data) - 3))): Data = "": Exit Sub
    If Left(Data, 3) = "106" Then Messtype Trim((Right(Data, Len(Data) - 3))): Data = "": Exit Sub
        For Cnt = 1 To IpChat.UBound
            If IpChat(Cnt).State = sckConnected Then
            '//////////////////////////////////////////////////////////////////////////
            If Left(Data, 3) = "103" Then
                List2.AddItem Right(Data, Len(Data) - 3)
                IpChat(Index).SendData "101" & "Welcome " & Right(Data, Len(Data) - 3) & " to the chat room you are vister " & "[" & UCount & "]"
                Data = ""
                Exit Sub
            End If
            '//////////////////////////////////////////////////////////////////////////
            If Left(Data, 3) = "105" Then
                IpChat(Cnt).SendData Trim(Right(Data, Len(Data) - 3))
            Else
                IpChat(Index).SendData "You need to update your software"
                Data = ""
            End If
        End If
        DoEvents
        Next
        If Err Then Err.Clear
        Data = ""
        
End Sub

Private Sub List1_Click()
    mIndex = List1.ListIndex + 1
    List2.Enabled = True
    List3.Enabled = True
    
End Sub

Private Sub List2_Click()
Dim Message As String
Dim Ans
    On Error Resume Next
    If List2.ListIndex = 1 Then
        Message = InputBox("Type in your Message below", "Send User Message")
        If Len(Message) = 0 Then
            Exit Sub
        Else
            IpChat(mIndex).SendData "Admin: " & Message
       End If
    End If
       
       If List2.ListIndex = 0 Then
            Ans = _
            MsgBox("Are you sure you want to disconect this user", _
            vbYesNo)
            If Ans = vbNo Then
                Exit Sub
            Else
                IpChat(mIndex).SendData "CLOSEUSER"
            End If
        End If
        
End Sub


