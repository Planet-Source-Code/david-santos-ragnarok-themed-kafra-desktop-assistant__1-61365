Attribute VB_Name = "modWndProc"
Option Explicit

Public pOldProc As Long

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
           
    Case PK_TRAYICON
        
        Select Case lParam
        Case WM_RBUTTONUP
            mnuForm.PopupMenu mnuForm.mnu1
        
        Case WM_LBUTTONDBLCLK
        
            Debug.Print frmMain.Visible
        
            If frmMain.Hidden Then
                frmMain.ShowKafra
            Else
                frmMain.HideKafra
            End If
        Case Else
         
        End Select
        WindowProc = 1  ' this return value doesn't really matter
    
    Case Else
        ' Pass the message to the procedure Visual Basic provided.
        WindowProc = CallWindowProc(pOldProc, hWnd, uMsg, wParam, lParam)
    
    End Select
    
End Function
