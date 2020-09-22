Attribute VB_Name = "Module1"
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

' Constants required by system tray
Public Enum enm_NIM_Shell
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
    NIF_MESSAGE = &H1
    NIF_ICON = &H2
    NIF_TIP = &H4
    WM_MOUSEMOVE = &H200
End Enum

Type Rect
     rleft As Long
     rtop As Long
     rright As Long
     rbot As Long
End Type

Type POINTAPI
        X As Long
        Y As Long
End Type


Public Type WINDOWPLACEMENT
    Length  As Long
    flags             As Long
    showCmd           As Long
    ptMinPosition     As POINTAPI
    ptMaxPosition     As POINTAPI
    rcNormalPosition  As Rect
End Type

Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As enm_NIM_Shell, pnid As NOTIFYICONDATA) As Boolean
Declare Function GetForegroundWindow& Lib "user32" ()
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function CreateDC& Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName$, ByVal lpDeviceName$, ByVal lpOutput$, ByVal lpInitData&)
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Any) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest

 'Behaviour over system tray
Public Const WM_MOUSEISMOVING = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_SETHOTKEY = &H32

' ShowWindow() Commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4

Public nidProgramData As NOTIFYICONDATA
Public Const HWND_TOPMOST = -1 ' sets a window to top
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOACTIVATE
Public currWinP As WINDOWPLACEMENT
Public mouse As POINTAPI 'store mouse position
Public StageDC As Long      'The staging area for the graphics
Public LogoDC As Long     'The sprite bitmap storage area
Public BackDC As Long       'The background bitmap storage
Public screendc As Long      'store screen DC
Public frmx, frmy                    As Integer
Public curfocus, oldfocus, tmpval As Long
Public angle_x, angle_y, speed, i, dir As Integer
Public winvis As Boolean

Function NewDC(hdcScreen As Long, HorRes As Long, VerRes As Long) As Long
    Dim hdcCompatible As Long
    Dim hbmScreen As Long
    hdcCompatible = CreateCompatibleDC(hdcScreen)                   'Create the DC
    hbmScreen = CreateCompatibleBitmap(hdcScreen, HorRes, VerRes)   'Temporary bitmap
    If SelectObject(hdcCompatible, hbmScreen) = vbNull Then         'If the function fails
        NewDC = vbNull                                              ' return null
    Else                                                            'If it succeeds
        NewDC = hdcCompatible                                       ' return the DC
    End If
End Function
Public Function degtorad(ang)
   degtorad = ang / 180 * 3.14159265359
End Function

