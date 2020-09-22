VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Menu mnu_shell 
      Caption         =   "Shell Menu"
      Begin VB.Menu itm_refresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu itm_hide 
         Caption         =   "Hide"
      End
      Begin VB.Menu itm_quit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub itm_hide_Click()
    If itm_hide.Caption = "Hide" Then
        itm_hide.Caption = "Show"
        Form1.Hide
        Form1.Timer1.Enabled = False
        winvis = False
     Else
        itm_hide.Caption = "Hide"
        Form1.Timer1.Enabled = True
     End If
End Sub

Private Sub itm_quit_Click()
    DeleteDC LogoDC
    DeleteDC SpriteDC
    DeleteDC BackDC
    Shell_NotifyIcon NIM_DELETE, nidProgramData
    Unload Form1
    Unload Form2
    End
End Sub

Private Sub itm_refresh_Click()
   Form1.Hide
   winvis = False
End Sub
