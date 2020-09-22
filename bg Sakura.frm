VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Kafra Desktop Assistant"
   ClientHeight    =   2550
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3015
   Icon            =   "bg Sakura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   201
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picKafra 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   1005
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   67
      ImageHeight     =   26
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":22A2
            Key             =   "kafra1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":37AE
            Key             =   "kafra1a"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":4C52
            Key             =   "kafra1b"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":60F6
            Key             =   "kafra1c"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":759A
            Key             =   "kafra2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":8D22
            Key             =   "kafra2a"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":A4AA
            Key             =   "kafra2b"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":BC32
            Key             =   "kafra2c"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":D3BA
            Key             =   "kafra3"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":ECCE
            Key             =   "kafra3a"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":105E2
            Key             =   "kafra3b"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "bg Sakura.frx":11EF6
            Key             =   "kafra3c"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picAni 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   3
      Left            =   1800
      ScaleHeight     =   390
      ScaleWidth      =   990
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox picAni 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   2
      Left            =   1800
      ScaleHeight     =   390
      ScaleWidth      =   990
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox picAni 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   1800
      ScaleHeight     =   390
      ScaleWidth      =   990
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox picAni 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   1800
      ScaleHeight     =   390
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   1680
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const XPFORMS = True

Option Explicit

Dim frameno As Integer
Dim dx As Integer

Dim Kafra As Long
Public Hidden As Boolean

Private Type KafraAniData
    X As Long
    Y As Long
    W As Long
    H As Long
End Type

Dim myDIB As New cDIBSection
Dim myRGN As New cDIBSectionRegion

Dim myKafraAniData As KafraAniData
Dim mTrans As Integer

Dim MainSubClass As New cSubclass
Implements WinSubHook.iSubclass

Private Sub Form_DblClick()
    If MovingMe Then Exit Sub
    frmPopup.Show
End Sub

Private Sub Form_Load()
    CheckOSVersion
    
    
    scrWidth = Screen.Width / Screen.TwipsPerPixelX
    scrHeight = Screen.Height / Screen.TwipsPerPixelY
    
    ' Set Class Cursor to NULL so that window
    ' doesn't try to reset the cursor everytime
    ' the window moves
    ' also seems to affect child windows
    
    SetClassLong Me.hWnd, GCL_HCURSOR, 0&
    
    If App.PrevInstance Then End
    
    frameno = 0
    dx = 1
  
    Dim myRect As RECT
    SystemParametersInfo 48&, 0&, myRect, 0&

    If Dir(App.path & "\Storage", vbDirectory) = "" Then
        MkDir App.path & "\Storage"
    End If
    
    LoadCursors
    
    Load frmPopup
    Load frmTray
    
    TaskBarHeight = Screen.Height - (myRect.Bottom * Screen.TwipsPerPixelY)
    Me.Top = Screen.Height - Me.Height - TaskBarHeight
    
    mnuForm.mnuAutorun.Checked = Not (GetStringKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Desktop Kafra") = "")
    mnuForm.mnuAlwaysOnTop.Checked = GetStringKey(HKEY_CURRENT_USER, "SOFTWARE\EnderSoft\DesktopKafra\", "Always On Top") = "Yes"
    
    Dim formLeft As Long
    formLeft = GetHexKey(HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Kafra X") * Screen.TwipsPerPixelX
    Me.Left = IIf(formLeft = 0, Screen.Width - Me.Width, formLeft)
    
    Randomize Timer
    Kafra = GetHexKey(HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Current Kafra")
    
    mnuForm.mnuKafraSwitch(Kafra).Checked = True
    
    If mnuForm.mnuAlwaysOnTop.Checked Then
        SetAlwaysOnTop Me, HWND_TOPMOST
    End If
    
    NewForm
    
    mTrans = 254
    
    MainSubClass.Subclass Me.hWnd, Me
    MainSubClass.AddMsg WM_MOVING, MSG_BEFORE
    MainSubClass.AddMsg WM_EXITSIZEMOVE, MSG_BEFORE
    
    frmBlurb.Show , Me
    If frmPopup.Messages > 0 Then
        frmBlurb.SetCaption "You have " & frmPopup.Messages & " memo" & IIf(frmPopup.Messages > 1, "s", "") & " in the memopad."
    Else
        frmBlurb.SetCaption "Hi, I'm your Kafra Desktop Assistant!  Let me help you create memos or store files!"
    End If
        
End Sub

Private Sub NewForm()
    SetHexKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Current Kafra", Kafra
    
    'for now, skip any images we don't have yet
    On Error GoTo errNewForm
    picAni(0).Picture = ImageList2.ListImages("kafra" & Kafra + 1).Picture
    picAni(1).Picture = ImageList2.ListImages("kafra" & Kafra + 1 & "a").Picture
    picAni(2).Picture = ImageList2.ListImages("kafra" & Kafra + 1 & "b").Picture
    picAni(3).Picture = ImageList2.ListImages("kafra" & Kafra + 1 & "c").Picture
    
    Timer1.Enabled = True
    GoTo noErr

errNewForm:
    Timer1.Enabled = False

noErr:
    
    On Error GoTo 0
    
    Dim i As Integer
    
    For i = 0 To mnuForm.mnuKafraSwitch.UBound
        mnuForm.mnuKafraSwitch(i).Checked = False
    Next
    
    mnuForm.mnuKafraSwitch(Kafra).Checked = True
    
    frmMain.Hide
    SetLayered Me.hWnd, False
    
    KafraName = LoadResString(Kafra + 101)
    
    Dim tempstr As String
    tempstr = LoadResString(Kafra + 111)
    With myKafraAniData
        .W = Split(tempstr, ",")(0)
        .H = Split(tempstr, ",")(1)
        .X = Split(tempstr, ",")(2)
        .Y = Split(tempstr, ",")(3)
    End With

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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SetRagCursor CUR_HANDCLOSE
        ReleaseCapture
        SendMessage Me.hWnd, WM_SYSCOMMAND, SC_CLICKMOVE, 0
    End If
    
    If Button = 2 Then
        SetAlwaysOnTop Me, HWND_NOTOPMOST
        PopupMenu mnuForm.mnu1
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MovingMe Then Exit Sub
    
    If Button = 0 Then
        SetRagCursor RagCursors.CUR_POINT
        
    Else
        SetRagCursor RagCursors.CUR_HANDOPEN
    End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    SetRagCursor CUR_TARGET
End Sub

' This function removes the tray icon and returns the old window procedure to use.
Private Sub Form_Unload(Cancel As Integer)
        
    Me.Hide
    
    MainSubClass.DelMsg WM_EXITSIZEMOVE, MSG_BEFORE
    MainSubClass.DelMsg WM_MOVING, MSG_BEFORE
    MainSubClass.UnSubclass
    Set MainSubClass = Nothing
    
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
     
    SetWindowEffects Me.hWnd, , 254
End Sub

Public Sub ShowKafra()
Dim spd As Single
    SetHexKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Kafra X", (Me.Left \ Screen.TwipsPerPixelX)
    Dim i As Integer
    If MovingMe Then Exit Sub
    
    MovingMe = True
    
    If OS_2000 Then
        If Hidden Then
            Me.Top = Screen.Height - Me.Height - TaskBarHeight
            For i = 0 To 255 Step 5
                SetWindowEffects Me.hWnd, , CByte(i)
                Me.Show
                DoEvents
            Next
        Else
            Me.Show
        End If
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
    
    'SetAlwaysOnTop Me, HWND_TOPMOST
    Hidden = False

    SetWindowEffects Me.hWnd, , 254
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

Private Sub Timer1_Timer()
    Timer1.Interval = 25
        
    If frameno = 3 Then
        dx = -1
        If Kafra = 2 Then
            Timer1.Interval = (1 + Int(Rnd() * 14)) * 500
        Else
            Timer1.Interval = 100
        End If
    End If
    
    If frameno = -1 Then
        dx = 1
        If Kafra = 2 Then
            Timer1.Interval = 1000
        Else
            Timer1.Interval = (1 + Int(Rnd() * 14)) * 500
        End If
    End If
    
    If frameno >= 0 Then
        With myKafraAniData
            BitBlt frmMain.hdc, .X, .Y, .W, .H, picAni(frameno).hdc, 0, 0, SRCCOPY
        End With
        
        frmMain.Refresh
    Else
        frmMain.Refresh
    End If
    
    frameno = frameno + dx
End Sub
