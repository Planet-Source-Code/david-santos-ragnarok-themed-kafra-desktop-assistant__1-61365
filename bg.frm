VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Kafra Desktop Assistant"
   ClientHeight    =   2835
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   2940
   Icon            =   "bg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   189
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   196
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const XPFORMS = True

Option Explicit

Dim frameno As Integer
Dim dX As Integer

Dim Kafra As Long
Public Hidden As Boolean
Dim MovingMe As Boolean

Private Type KafraAniData
    x As Long
    y As Long
    W As Long
    H As Long
End Type

Dim myDIB As New cDIBSection
Dim myRGN As New cDIBSectionRegion

Dim myKafraAniData(6) As KafraAniData
Dim mTrans As Integer

Dim MainSubClass As New cSubclass
Implements WinSubHook.iSubclass

Private Sub Form_Activate()
    If mnuForm.mnuAlwaysOnTop.Checked Then
        SetAlwaysOnTop Me, HWND_TOPMOST
    End If
End Sub

Private Sub Form_DblClick()
    If MovingMe Then Exit Sub
    frmPopup.Show
End Sub

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
     SetHexKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Kafra X", (Me.Left \ Screen.TwipsPerPixelX)
End Sub

Private Sub Form_Load()
Dim tempstr As String
    
    CheckOSVersion
    
    If App.PrevInstance Then End
    frameno = 0
    dX = 1
    
    If Dir(App.path & "\ani.txt") <> "" Then
        Open App.path & "\ani.txt" For Input As 1
        While Not EOF(1)
        Line Input #1, tempstr
        With myKafraAniData(Split(tempstr, ",")(0) - 1)
            .W = Split(tempstr, ",")(1)
            .H = Split(tempstr, ",")(2)
            .x = Split(tempstr, ",")(3)
            .y = Split(tempstr, ",")(4)
        End With
        Wend
        Close 1
    End If
    
    tempstr = ""
    
    Dim myRect As RECT
    SystemParametersInfo 48&, 0&, myRect, 0&

    If Dir(App.path & "\Storage", vbDirectory) = "" Then
        MkDir App.path & "\Storage"
    End If
    
    LoadCursors
    
    Load frmPopup
    Load frmTray
    
    TaskBarHeight = Screen.Height - (myRect.Bottom * Screen.TwipsPerPixelY)
    
    mnuForm.mnuAutorun.Checked = Not (GetStringKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Desktop Kafra") = "")
    mnuForm.mnuAlwaysOnTop.Checked = GetStringKey(HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Always On Top") = "Yes"
    
    Me.Top = Screen.Height - Me.Height - TaskBarHeight
    
    Dim formLeft As Long
    formLeft = GetHexKey(HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Kafra X") * Screen.TwipsPerPixelX
    Me.Left = IIf(formLeft = 0, Screen.Width - Me.Width, formLeft)
    
    Randomize Timer
    Kafra = GetHexKey(HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Current Kafra")
    
    mnuForm.mnuKafraSwitch(Kafra).Checked = True
    
    If mnuForm.mnuAlwaysOnTop.Checked Then
        SetAlwaysOnTop Me, HWND_TOPMOST
    End If
    
    MainSubClass.Subclass Me.hWnd, Me
    MainSubClass.AddMsg WM_MOVING, MSG_BEFORE
    MainSubClass.AddMsg WM_EXITSIZEMOVE, MSG_BEFORE
    
    
    NewForm
    
    frmBlurb.Show , Me
    frmBlurb.SetCaption "You have " & frmPopup.Messages & " messages in the memopad."

    frmPopup.Top = Me.Top - 500
    frmPopup.Left = Me.Left - 300
End Sub

Private Sub NewForm()
    frmMain.Hide
    
    SetHexKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Current Kafra", Kafra

    Dim i As Integer
    
    For i = 0 To mnuForm.mnuKafraSwitch.UBound
        mnuForm.mnuKafraSwitch(i).Checked = False
    Next
    
    mnuForm.mnuKafraSwitch(Kafra).Checked = True
    
    frmMain.Hide
    SetLayered Me.hWnd, False
    
    KafraName = LoadResString(Kafra + 101)
    
    frmMain.Picture = LoadResPicture(Kafra + 111, 0)
    
    myDIB.CreateFromPicture frmMain.Picture
    Me.Height = myDIB.Height * Screen.TwipsPerPixelY
    Me.Width = myDIB.Width * Screen.TwipsPerPixelX
    
    myRGN.LoadFromResource Kafra + 101
    myRGN.Applied(Me.hWnd) = True
    
    SetLayered Me.hWnd, True
    SetWindowEffects Me.hWnd, , 255
    frmMain.Show
    
    Me.Top = Screen.Height - Me.Height - TaskBarHeight
    'If Me.Left > Screen.Width - Me.Width Then Me.Left = Screen.Width - Me.Width
    frmMain.Refresh
    
End Sub

Public Sub SwitchKafra(Index As Integer)
    Kafra = Index
    NewForm
End Sub

Public Sub RandomKafra()
    Dim oldKafra As Integer
    oldKafra = Kafra
    'make sure we never get the same kafra during a random switch
    While Kafra = oldKafra
        Kafra = Int(Rnd(1) * 7)
    Wend
    
    NewForm
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        SetRagCursor CUR_HANDCLOSE
        ReleaseCapture
        SendMessage Me.hWnd, WM_SYSCOMMAND, SC_CLICKMOVE, 0
    End If
    If Button = 2 Then
        PopupMenu mnuForm.mnu1
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MovingMe Then Exit Sub
    If Button = 0 Then
        SetRagCursor RagCursors.CUR_POINT
        If Me.Top > Screen.Height - Me.Height * 2 / 3 Then
            HideKafra
        Else
            Me.Top = Screen.Height - Me.Height - TaskBarHeight
        End If
    Else
        SetRagCursor RagCursors.CUR_HANDOPEN
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MovingMe Then Exit Sub
    If Button = 1 Then
        SetRagCursor RagCursors.CUR_HANDOPEN
        Me.Top = Screen.Height - Me.Height - TaskBarHeight
    End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo errhandler
Dim mFile As New FileSystemObject
Dim i As Integer

    If MovingMe Then Exit Sub
    
    For i = 1 To Data.Files.Count
        If GetAttr(Data.Files(i)) = vbDirectory Then
            mFile.MoveFolder Data.Files(i), curPath & "\" & mFile.GetFileName(Data.Files(i))
        Else
            mFile.MoveFile Data.Files(i), curPath & "\" & mFile.GetFileName(Data.Files(i))
        End If
    Next
    
    frmPopup.ShowItems curPath
    Exit Sub

errhandler:
    MsgBox "One or more of the items could not be moved.", vbInformation
End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    SetRagCursor CUR_TARGET
End Sub

' This function removes the tray icon and returns the old window procedure to use.
Private Sub Form_Unload(Cancel As Integer)
    
    MainSubClass.DelMsg WM_EXITSIZEMOVE, MSG_BEFORE
    MainSubClass.DelMsg WM_MOVING, MSG_BEFORE
    MainSubClass.UnSubclass
    Set MainSubClass = Nothing
    
    Me.Hide
    
    myDIB.ClearUp
    myRGN.Destroy
    Set myRGN = Nothing
    
    SetHexKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Kafra X", (Me.Left \ Screen.TwipsPerPixelX)
    Unload frmTray
    Unload frmAbout
    Unload frmPopup
    Unload mnuForm
End Sub

Public Sub HideKafra()
Dim spd As Single
    SetHexKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Kafra X", (Me.Left \ Screen.TwipsPerPixelX)
    SetAlwaysOnTop Me, HWND_NOTOPMOST
    Dim i As Integer
    If MovingMe Then Exit Sub
    
    If Hidden Then Exit Sub
    MovingMe = True
    
    If OS_2000 Then
        For i = 255 To 0 Step -5
            SetWindowEffects Me.hWnd, , CByte(i)
            DoEvents
        Next
    Else
        While (Screen.Height - Me.Top) > 10
            Me.Top = Me.Top + spd
            DoEvents
            Sleep 10
            spd = (Screen.Height - Me.Top) / 12
        Wend
    End If
    
    Me.Hide
    MovingMe = False
    Hidden = True
     
    #If XPFORMS Then
        SetWindowEffects Me.hWnd, , 254
    #End If
End Sub

Public Sub ShowKafra()
Dim spd As Single
    SetHexKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Kafra X", (Me.Left \ Screen.TwipsPerPixelX)
    Dim i As Integer
    If MovingMe Then Exit Sub
    
    If Not Hidden Then Exit Sub
    MovingMe = True
    
    If OS_2000 Then
        Me.Top = Screen.Height - Me.Height - TaskBarHeight
        For i = 0 To 255 Step 5
            SetWindowEffects Me.hWnd, , CByte(i)
            Me.Show
            DoEvents
        Next
    Else
        Me.Show
        While Me.Top > (Screen.Height - Me.Height - TaskBarHeight)
            Me.Top = Me.Top + spd
            DoEvents
            Sleep 10
            spd = (Me.Top - Screen.Height - Me.Height) / 16
        Wend
    End If
    
    MovingMe = False
    Me.Top = Screen.Height - Me.Height - TaskBarHeight
    
    SetAlwaysOnTop Me, HWND_TOPMOST
    Hidden = False

    #If XPFORMS Then
        SetWindowEffects Me.hWnd, , 254
    #End If
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
'
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
'Dim lR As Long
Dim tR As RECT
Dim lWidth As Long
Dim lHeight As Long

Dim scrWidth As Long
Dim scrHeight As Long

        scrWidth = Screen.Width / Screen.TwipsPerPixelX
        scrHeight = Screen.Height / Screen.TwipsPerPixelY
        
    Select Case uMsg
    Case WM_MOVING
        If mTrans > 128 Then
            SetWindowEffects hWnd, , mTrans
            mTrans = mTrans - 16
        End If
       ' Form is moving:
       CopyMemory tR, ByVal lParam, Len(tR)
       lWidth = tR.Right - tR.Left
       lHeight = tR.Bottom - tR.Top
               
       With tR
        If .Top < scrHeight - lHeight - TaskBarHeight / Screen.TwipsPerPixelY Then tR.Top = scrHeight - lHeight - TaskBarHeight / Screen.TwipsPerPixelY
       End With
       
       tR.Right = tR.Left + lWidth
       tR.Bottom = tR.Top + lHeight
       CopyMemory ByVal lParam, tR, Len(tR)
       
    Case WM_EXITSIZEMOVE
        While mTrans < 254
            If mTrans > 254 Then mTrans = 254
            SetWindowEffects Me.hWnd, , mTrans
            mTrans = mTrans + 32
            DoEvents
        Wend
        mTrans = 254
        SetWindowEffects Me.hWnd, , mTrans
    
        If Me.Top > Screen.Height - Me.Height * 2 / 3 Then
            HideKafra
        Else
            Me.Top = Screen.Height - Me.Height - TaskBarHeight
        End If
    
    End Select
End Sub

