VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "TV.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   2640
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   645
      Left            =   2040
      Picture         =   "TV.frx":0442
      ScaleHeight     =   585
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
   'add to system tray
   With nidProgramData
        .cbSize = Len(nidProgramData)
        .hwnd = Form1.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Form1.Icon
        .szTip = "We're On TV" & vbNullChar
    End With
    
  Shell_NotifyIcon NIM_ADD, nidProgramData
  
  oldfocus = GetForegroundWindow()             ' get currently active window
  
  frmx = Screen.Width / Screen.TwipsPerPixelX - (Picture1.Width / 1.5)    ' calculate screen positions
  frmy = 0
  
  'set form1 to be top window
  tmpval = SetWindowPos(Form1.hwnd, HWND_NOTOPMOST, frmx, frmy, Picture1.Width, Picture1.Height, SWP_SHOWME)
  
  'set variables
  angle_x = 0 'logo x angle
  angle_y = 0 'logo y angle
  speed = 5    'spin speed
  i = 0               'general dogs body variable... ;)
   
  screendc = CreateDC("DISPLAY", "", "", 0&) 'get screen device context
     
  LogoDC = NewDC(Form1.hDC, Picture1.Width, Picture1.Height) 'create work areas for logo
  BackDC = NewDC(Form1.hDC, Picture1.Width, Picture1.Height)
  StageDC = NewDC(Form1.hDC, Picture1.Width, Picture1.Height)
    
  tmpval = SelectObject(LogoDC, Picture1) ' copy logo
  tmpval = BitBlt(BackDC, 0, 0, Picture1.Width, Picture1.Height, screendc, frmx, frmy, SRCCOPY) ' set background of work area
       
  'for use later when setting window visible but not active
  currWinP.Length = Len(currWinP)
  currWinP.flags = 0&
  currWinP.showCmd = SW_SHOWNOACTIVATE
       
  Form1.Hide   ' hide form1
  winvis = False   'form1 hidden
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If X = WM_RBUTTONUP Then  ' was the r-button clicked on system tray
        Form1.PopupMenu Form2.mnu_shell 'Yep, display popup menu
     End If
End Sub

Private Sub Timer1_Timer()

    tmpval = GetCursorPos(mouse)   'get mouse position
    curfocus = GetForegroundWindow()    ' get currently active window
 
    If (mouse.X > frmx And mouse.Y < Picture1.Height) Or (curfocus <> oldfocus) Then      'is the mouse on the form or has the active window changed????
        oldfocus = curfocus   ' update old active window
        If winvis = True Then    ' if were displaying then ....
            currWinP.showCmd = SW_HIDE         ' set window to hide
            tmpval = SetWindowPlacement(Form1.hwnd, currWinP)    ' hide window
            winvis = False    ' window hidden
        End If
    Else
        If mouse.X < frmx Or mouse.Y > frmy + Picture1.Height Then  ' is mouse on form1?
            If winvis = False Then ' is form1 hidden?
                tmpval = BitBlt(BackDC, 0, 0, Picture1.Width, Picture1.Height, screendc, frmx, frmy, SRCCOPY) 'as form1 is hidden grab screen
                currWinP.showCmd = SW_SHOWNOACTIVATE ' set form1 to show but not activate
                tmpval = SetWindowPlacement(Form1.hwnd, currWinP) 'send to form1
                tmpval = SetWindowPos(Form1.hwnd, HWND_TOPMOST, frmx, frmy, Picture1.Width, Picture1.Height, SWP_SHOWME) ' show form1
                winvis = True 'form1 visible
            End If
        End If
   
        tmpval = BitBlt(StageDC, 0, 0, Picture1.Width, Picture1.Height, BackDC, 0, 0, SRCCOPY) 'copy background to stage area
        
        For i = Form1.Picture1.Width To 1 Step -1
            tmpval = BitBlt(StageDC, Cos(degtorad(angle_x + i)) * (Picture1.Width / 3.2) + (Picture1.Width / 3.2), Sin(degtorad(angle_y + i)) * (5) + 2.5, 1, Picture1.Height, LogoDC, i, 0, SRCAND) ' put spinning logo onto stage area
        Next i
        tmpval = BitBlt(Form1.hDC, 0, 0, Picture1.Width, Picture1.Height, StageDC, 0, 0, SRCCOPY) ' copy stage to form1
        
        angle_x = angle_x + speed ' rotate logo x
        angle_y = angle_y + speed ' rotate logo y
        
        If angle_x >= 360 Then  ' have we done a full rotation 360o??
            angle_x = 0  ' Yep, reset angle
        End If
        If angle_y >= 360 Then
            angle_y = 0
        End If
    End If
End Sub
