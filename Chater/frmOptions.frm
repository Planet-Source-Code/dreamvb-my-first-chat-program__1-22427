VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   1200
      TabIndex        =   8
      Top             =   1665
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   330
      Left            =   75
      TabIndex        =   4
      Top             =   1665
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Height          =   1470
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4815
      Begin VB.TextBox txtIp 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   690
         TabIndex        =   7
         Top             =   375
         Width           =   1320
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "69"
         Top             =   375
         Width           =   390
      End
      Begin VB.TextBox txtNickname 
         Height          =   285
         Left            =   1005
         TabIndex        =   5
         Top             =   795
         Width           =   2145
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nick name"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   825
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Port No."
         Height          =   195
         Left            =   2100
         TabIndex        =   2
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IP No."
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   420
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
    If Len(txtIp) = 0 Then
        MsgBox "You must enter the servers IP number", vbInformation
        Exit Sub
        ElseIf Len(txtPort) = 0 Then
            MsgBox "You must enter the servers Port", vbInformation
            Exit Sub
        ElseIf Len(txtNickname) = 0 Then
            MsgBox "You must enter a nickname to use", vbInformation
            Exit Sub
        Else
        ServerIp = txtIp.Text
        ServerPort = Val(txtPort.Text)
        Nick_name = txtNickname.Text
    End If
    Unload frmOptions
    Form1.Show
    
    
End Sub

Private Sub Command2_Click()
    Unload frmOptions: Form1.Show
    
End Sub


Private Sub Form_Load()
Dim CompName As String
    CompName = String(255, Chr(0))
    GetComputerName CompName, 255
    CompName = Left(CompName, InStr(1, CompName, Chr(0)) - 1)
    txtNickname = StrConv(CompName, vbProperCase)
    
End Sub
