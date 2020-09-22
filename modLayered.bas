Attribute VB_Name = "modLayered"
'For Windows 2000/XP only
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
   (ByVal hWnd As Long, ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
'Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

'Requires Windows 2000 or later:
Private Const WS_EX_LAYERED = &H80000

Private Declare Function SetLayeredWindowAttributes Lib "user32" _
   (ByVal hWnd As Long, ByVal crKey As Long, _
   ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
   
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2


Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Public OS_2000 As Boolean

Public Sub CheckOSVersion()
Dim os As OSVERSIONINFO  ' receives version information
Dim retval As Long  ' return value

os.dwOSVersionInfoSize = Len(os)  ' set the size of the structure
retval = GetVersionEx(os)  ' read Windows's version information
        
    If os.dwPlatformId > 1 Then OS_2000 = True
        
End Sub

Public Function SetWindowEffects( _
      ByVal lhWnd As Long, _
      Optional ByVal transparentColor As Long = -1, _
      Optional ByVal alpha As Byte = 255 _
   )
Dim crKey As Long
Dim bAlpha As Byte
Dim lFlags As Long

    If Not OS_2000 Then Exit Function

   If (transparentColor = -1) Then
      transparentColor = 0
   Else
      crKey = transparentColor
      lFlags = lFlags Or LWA_COLORKEY
   End If
   
   bAlpha = alpha
   
   If (alpha < 255) Then
      lFlags = lFlags Or LWA_ALPHA
   End If
   
   SetLayeredWindowAttributes lhWnd, _
       transparentColor, bAlpha, lFlags

End Function

Public Sub SetLayered(ByVal hWnd As Long, ByVal bState As Boolean)
   Dim lStyle As Long
    
    If Not OS_2000 Then Exit Sub
   
   lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
   If bState Then
      lStyle = lStyle Or WS_EX_LAYERED
   Else
      lStyle = lStyle And Not WS_EX_LAYERED
   End If
   SetWindowLong hWnd, GWL_EXSTYLE, lStyle
End Sub


