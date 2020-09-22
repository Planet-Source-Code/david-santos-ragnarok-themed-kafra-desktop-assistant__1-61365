Attribute VB_Name = "modAPI"
Option Explicit

'API Imports
'===================================================
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function LoadCursorFromFile Lib "user32.dll" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fWinIni As Long) As Long

Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long

'Constants
'=============================================================
Public Const HWND_TOPMOST = -1&
Public Const HWND_NOTOPMOST = -2&

Public Const SWP_NOACTIVATE = &H10&
Public Const SWP_SHOWWINDOW = &H40&

Public Const SW_HIDE = 0&
Public Const SW_MAX = 10&
Public Const SW_MAXIMIZE = 3&
Public Const SW_MINIMIZE = 6&
Public Const SW_NORMAL = 1&
Public Const SW_SHOWMAXIMIZED = 3&
Public Const SW_SHOWMINIMIZED = 2&
Public Const SW_SHOWMINNOACTIVE = 7&
Public Const SW_SHOWNOACTIVATE = 4&
Public Const SW_SHOWNORMAL = 1&

Public Const GW_CHILD = 5&
Public Const GW_HWNDNEXT = 2&
Public Const WS_VISIBLE = &H10000000
Public Const GWL_STYLE = (-16&)

Public Const GCL_HCURSOR = (-12&)

Public Const NIF_ICON = &H2&
Public Const NIF_MESSAGE = &H1&
Public Const NIF_TIP = &H4&

Public Const NIM_ADD = &H0&
Public Const NIM_MODIFY = &H1&
Public Const NIM_DELETE = &H2&
Public Const GWL_WNDPROC = (-4&)

Public Const PK_TRAYICON = &H401&

Public Const RGN_OR = 2
Public Const RGN_DIFF = 4

Public Const SC_CLICKMOVE = &HF012&

Public Const SRCCOPY = &HCC0020                             ' (DWORD) dest = source

'Windows Message Constants
'===================================================
Public Const WM_SYSCOMMAND = &H112&
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_LBUTTONDBLCLK = &H203&
Public Const WM_LBUTTONDOWN = &H201&
Public Const WM_LBUTTONUP = &H202&
Public Const WM_RBUTTONDBLCLK = &H206&
Public Const WM_RBUTTONDOWN = &H204&
Public Const WM_RBUTTONUP = &H205&
Public Const WM_CLOSE = &H10&


'Types
'===================================================
Public Type POINT_TYPE
    x As Long
    y As Long
End Type


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64                                    ' Windows 2000: make this String * 128
    ' The following data members are only valid in Windows 2000!
    ' (uncomment the following lines to use them)
    'dwState As Long
    'dwStateMask As Long
    'szInfo As String * 256
    'uTimeoutOrVersion As Long
    'szInfoTitle As String * 64
    'dwInfoFlags As Long
End Type
