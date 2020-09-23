VERSION 5.00
Begin VB.Form frmenu 
   ClientHeight    =   2745
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   1560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2745
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuCon 
         Caption         =   "&Conect"
      End
      Begin VB.Menu mnuDis 
         Caption         =   "&Disconect"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOp 
      Caption         =   "&Options"
      Begin VB.Menu mnuConf 
         Caption         =   "&Config"
      End
      Begin VB.Menu mnuSOn 
         Caption         =   "&Sound Effects On"
      End
      Begin VB.Menu mnuSOff 
         Caption         =   "&Sound Effects Off"
      End
   End
End
Attribute VB_Name = "frmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAbout_Click()
    MsgBox "Dreams's Chat House by Ben Jones Beta 2" & vbCrLf & "Email: Dreamvb@yahoo.com", vbInformation
    
End Sub

Private Sub mnuCon_Click()
On Error Resume Next
    Form1.KillHtm
    mnuCon.Enabled = False
    mnuDis.Enabled = True
    Form1.IChat.Connect ServerIp, ServerPort
    Form1.Timer1.Enabled = True
    If Err Then Err.Clear

End Sub

Private Sub mnuConf_Click()
    Form1.Hide
    frmOptions.Show

End Sub

Private Sub mnuDis_Click()
On Error Resume Next
    Form1.IChat.Close
    Form1.KillHtm
    mnuDis.Enabled = False
    mnuCon.Enabled = True
    If Play_Sounds = False Then
        Exit Sub
    Else
        sndPlaySound AddSlash(App.Path) & "dis.wav", SND_ASYNC Or SND_NODEFAULT
        
    End If

End Sub

Private Sub mnuExit_Click()
On Error Resume Next
    Form1.KillHtm
    If Form1.IChat.State = sckConnected Then
        MsgBox "You are still conected to the server please disconect before exiting", vbInformation, "Exit program........."
        Exit Sub
    Else
        Kill AddSlash(App.Path) & "chat.htm"
        Unload Form1: End
    End If

End Sub

Private Sub mnuSOff_Click()
    Play_Sounds = False
    mnuSOff.Enabled = False
    mnuSOn.Enabled = True

End Sub

Private Sub mnuSOn_Click()
    Play_Sounds = True
    mnuSOn.Enabled = False
    mnuSOff.Enabled = True

End Sub
