Attribute VB_Name = "modAFS"
Public TaskBarHeight As Long
Public MovingMe As Boolean
Public DraggingMe As Boolean
Public DraggingPop As Boolean

Public KafraName As String
Public curPath As String

Dim hcurPointer As Long
Dim hcurHandOpen As Long
Dim hcurHandClose As Long
Dim hcurTarget As Long

Public scrWidth As Long
Public scrHeight As Long

Enum RagCursors
    CUR_POINT
    CUR_HANDOPEN
    CUR_HANDCLOSE
    CUR_TARGET
End Enum

#Const XPFORMS = False

Public Sub SnapToScreen(curx As Long, cury As Long, r As RECT, snapped As Boolean, Optional snapwidth As Long = 10)
Dim lWidth As Long
Dim lHeight As Long
    
    'If snapped Then Exit Sub

    With r
        lWidth = .Right - .Left
        lHeight = .Bottom - .Top
        
        If Abs(curx) < snapwidth Then
            .Left = 0
            snapped = True
        ElseIf Abs((scrWidth - lWidth) - curx) < snapwidth Then
            .Left = scrWidth - lWidth
            snapped = True
        Else
            .Left = curx
            snapped = False
        End If
        
        If Abs(cury) < snapwidth Then
            .Top = 0
            snapped = True
        ElseIf Abs((scrHeight - lHeight) - cury) < snapwidth Then
            .Top = scrHeight - lHeight
            snapped = True
        Else
            .Top = cury
            snapped = False
        End If
     
    End With

End Sub

Public Sub SnapToForm(curx As Long, cury As Long, dx As Long, dy As Long, r As RECT, TargetForm As Form, snapped As Boolean, Optional snapwidth As Long = 10)
Dim lWidth As Long
Dim lHeight As Long

    'doesn not support snapping to right and bottom edges
    '(if one form is smaller than other form)
    'requires r coordinates to be modified and re-wriiten
    'into the SendMessage function

    If snapped Then Exit Sub

    lWidth = r.Right - r.Left
    lHeight = r.Bottom - r.Top

    If TargetForm.Visible Then
        
        If Abs(r.Top - TargetForm.Top / Screen.TwipsPerPixelY) < TargetForm.Height / Screen.TwipsPerPixelY Then
            'snap left to right side
            If (Abs(r.Left - (TargetForm.Left + TargetForm.Width) / Screen.TwipsPerPixelX) < 10) Then
                r.Left = (TargetForm.Left + TargetForm.Width) / Screen.TwipsPerPixelX
                snapped = True
            End If
        End If
        
        If Abs(r.Top - TargetForm.Top / Screen.TwipsPerPixelY) < lHeight Then
            'snap right to left side
            If (Abs(r.Left - (TargetForm.Left / Screen.TwipsPerPixelX) + lWidth) < 10) Then
                r.Left = (TargetForm.Left) / Screen.TwipsPerPixelX - lWidth
                snapped = True
            End If
        End If
        
        If Abs(r.Left - TargetForm.Left / Screen.TwipsPerPixelY) < TargetForm.Width / Screen.TwipsPerPixelY Then
            'snap top to bottom
            If (Abs(r.Top - (TargetForm.Top + TargetForm.Height) / Screen.TwipsPerPixelY) < 10) Then
                r.Top = (TargetForm.Top + TargetForm.Height) / Screen.TwipsPerPixelY
                snapped = True
            End If
        End If
        
        If Abs(r.Left - TargetForm.Left / Screen.TwipsPerPixelY) < lWidth Then
            'snap bottom to top
            If (Abs(r.Top - (TargetForm.Top / Screen.TwipsPerPixelY) + lHeight) < 10) Then
                r.Top = (TargetForm.Top) / Screen.TwipsPerPixelY - lHeight
                snapped = True
            End If
        End If
        
        If (r.Left - (TargetForm.Left / Screen.TwipsPerPixelX)) <= TargetForm.Width / Screen.TwipsPerPixelY And _
           (r.Left - (TargetForm.Left / Screen.TwipsPerPixelX)) >= 0 Then
            'snap top to top
            If (Abs(r.Top - TargetForm.Top / Screen.TwipsPerPixelY) < 10) Then
                r.Top = TargetForm.Top / Screen.TwipsPerPixelY
                snapped = True
            End If
        End If
        
        
        If (r.Top - (TargetForm.Top / Screen.TwipsPerPixelY)) <= TargetForm.Height / Screen.TwipsPerPixelY And _
           (r.Top - (TargetForm.Top / Screen.TwipsPerPixelY)) >= 0 Then
            'snap left to left
            If (Abs(r.Left - TargetForm.Left / Screen.TwipsPerPixelX) < 10) Then
                r.Left = TargetForm.Left / Screen.TwipsPerPixelX
                snapped = True
            End If
        End If
        
            
        If ((TargetForm.Left / Screen.TwipsPerPixelX) - r.Left) <= lWidth And _
           ((TargetForm.Left / Screen.TwipsPerPixelX) - r.Left) >= 0 Then
            'snap top to top
            If (Abs(r.Top - TargetForm.Top / Screen.TwipsPerPixelY) < 10) Then
                r.Top = TargetForm.Top / Screen.TwipsPerPixelY
                snapped = True
            End If
        End If
        
        
        If ((TargetForm.Top / Screen.TwipsPerPixelY) - r.Top) <= lHeight And _
           ((TargetForm.Top / Screen.TwipsPerPixelY) - r.Top) >= 0 Then
            'snap left to left
            If (Abs(r.Left - TargetForm.Left / Screen.TwipsPerPixelX) < 10) Then
                r.Left = TargetForm.Left / Screen.TwipsPerPixelX
                snapped = True
            End If
        End If
        
        
        
    
    End If


End Sub


Public Function CornerFormShape(ByRef bG As Form)
Dim CurRgn As Long
bG.ScaleMode = vbPixels
Dim W As Long
Dim H As Long

W = bG.ScaleWidth
H = bG.ScaleHeight

CurRgn = CreateRectRgn(0, 0, W, H)  ' Create base region which is the current whole window
RemoveArea CurRgn, 0, 0, 2, 1
RemoveArea CurRgn, 0, 1, 1, 2
RemoveArea CurRgn, W - 2, 0, W, 1
RemoveArea CurRgn, W - 1, 1, W, 2
RemoveArea CurRgn, 0, H - 1, 2, H
RemoveArea CurRgn, 0, H - 2, 1, H - 1
RemoveArea CurRgn, W - 2, H - 1, W, H
RemoveArea CurRgn, W - 1, H - 2, W, H - 1
success = SetWindowRgn(bG.hWnd, CurRgn, True)  ' Finally set the windows region to the final product

DeleteObject (CurRgn)  ' Delete the now un-needed base region and free resources

Dim lastcursor As Long

Set src = Nothing
End Function

Private Sub RemoveArea(ByRef CurRgn As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
Dim success As Long
    TempRgn = CreateRectRgn(X1, Y1, X2, Y2) ' Create a temporary pixel region for this pixel
    success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)  ' Combine temp pixel region with base region using RGN_DIFF to extract the pixel and make it transparent
    DeleteObject (TempRgn)  ' Delete the temporary pixel region and clear up very important resources
End Sub

Public Sub SetRagCursor(curtype As RagCursors)
    Select Case curtype
    Case RagCursors.CUR_POINT
        If hcurPointer <> 0 Then SetCursor hcurPointer
    Case RagCursors.CUR_HANDOPEN
        If hcurHandOpen <> 0 Then SetCursor hcurHandOpen
    Case RagCursors.CUR_HANDCLOSE
        If hcurHandClose <> 0 Then SetCursor hcurHandClose
    Case RagCursors.CUR_TARGET
        If hcurTarget <> 0 Then SetCursor hcurTarget
    End Select
End Sub

Public Sub SetAlwaysOnTop(target As Form, mode As Long)
    SetWindowPos target.hWnd, mode, target.Left / 15, _
                        target.Top / 15, target.Width / 15, _
                        target.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Public Sub LoadCursors()
    hcurPointer = LoadCursorFromFile(App.path & "\Cursors\rag_cur.ani")
    hcurHandOpen = LoadCursorFromFile(App.path & "\Cursors\handopen.cur")
    hcurHandClose = LoadCursorFromFile(App.path & "\Cursors\handclosed.cur")
    hcurTarget = LoadCursorFromFile(App.path & "\Cursors\rag_target.ani")
End Sub
