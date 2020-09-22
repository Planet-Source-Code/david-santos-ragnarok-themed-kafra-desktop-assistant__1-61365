VERSION 5.00
Begin VB.Form frmPopupTools 
   BorderStyle     =   0  'None
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   Icon            =   "frmPopupTools.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmPopupTools.frx":000C
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin KafraDesk.RagButtonResource RagButtonResource1 
      Left            =   120
      Top             =   4200
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin KafraDesk.SuperRagButton RagButtonClose 
      Height          =   300
      Left            =   3480
      TabIndex        =   8
      Top             =   4320
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
      Caption         =   "cancel"
      ResourceControl =   "RagButtonResource1"
   End
   Begin KafraDesk.SuperRagButton RagButtonOK 
      Height          =   300
      Left            =   2760
      TabIndex        =   7
      Top             =   4320
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
      ResourceControl =   "RagButtonResource1"
   End
   Begin KafraDesk.RagListBox RagListBox1 
      Height          =   3705
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6535
   End
   Begin VB.PictureBox picBtnMini 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3720
      Picture         =   "frmPopupTools.frx":1610C
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   45
      Width           =   135
   End
   Begin VB.PictureBox picBtnExit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3960
      Picture         =   "frmPopupTools.frx":1624C
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   3
      Top             =   45
      Width           =   135
   End
   Begin VB.PictureBox picBtnNothing 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   75
      Picture         =   "frmPopupTools.frx":1638C
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   45
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kafra Desktop Tools"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   15
      Width           =   3255
   End
   Begin VB.Label lblStorage 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   2325
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2565
   End
End
Attribute VB_Name = "frmPopupTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim curMsg As Integer
'Dim lastMsg As Integer
Dim winstate As Integer
'Dim curPointer As Integer

'Dim moveForm As Boolean
Dim pdx As Long
Dim pdy As Long
Dim mTrans As Integer

Dim mySubClass As New cSubclass
Implements WinSubHook.iSubclass

Private Sub Form_Activate()
        SetAlwaysOnTop Me, HWND_TOPMOST
        SetLayered Me.hWnd, True
        SetWindowEffects Me.hWnd, , 254
End Sub

Private Sub Form_Load()

    CornerFormShape Me

    Me.Left = GetHexKey(HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Calendar X") * Screen.TwipsPerPixelX
    Me.Top = GetHexKey(HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Calendar Y") * Screen.TwipsPerPixelY
    
    mySubClass.Subclass Me.hWnd, Me
    mySubClass.AddMsg WM_EXITSIZEMOVE, MSG_BEFORE
    mySubClass.AddMsg WM_MOVING, MSG_BEFORE
    mySubClass.AddMsg WM_LBUTTONDOWN, MSG_BEFORE
    
    mTrans = 255
    
    RagListBox1.Add "Run Ragnarok Online"
    RagListBox1.Add "-------------------"
    RagListBox1.Add "Free Memory"
    RagListBox1.Add "Shutdown Timer"
    RagListBox1.Add "-------------------"
    RagListBox1.Add "Run Scandisk"
    RagListBox1.Add "Run Defrag"
    RagListBox1.Add "-------------------"
    
End Sub

Private Sub Form_LostFocus()
    SetHexKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Calendar X", (Me.Left \ Screen.TwipsPerPixelX)
    SetHexKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Calendar Y", (Me.Top \ Screen.TwipsPerPixelY)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SetRagCursor CUR_HANDCLOSE
        ReleaseCapture
        SendMessage Me.hWnd, WM_SYSCOMMAND, SC_CLICKMOVE, 0
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        SetRagCursor CUR_POINT
        picBtnExit.Picture = frmPopup.ImageList1.ListImages("closeoff").Picture
        picBtnMini.Picture = frmPopup.ImageList1.ListImages("minioff").Picture
        picBtnNothing.Picture = frmPopup.ImageList1.ListImages("buttonoff").Picture
    
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mySubClass.DelMsg WM_EXITSIZEMOVE, MSG_BEFORE
    mySubClass.DelMsg WM_MOVING, MSG_BEFORE
    mySubClass.UnSubclass
    Set mySubClass = Nothing
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
''
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
'Dim lR As Long
Dim tR As RECT
Dim lWidth As Long
Dim lHeight As Long
Dim cursorloc As POINT_TYPE
Dim dx As Long
Dim dy As Long

    
    Select Case uMsg
    Case WM_LBUTTONDOWN
        GetCursorPos cursorloc
        pdx = cursorloc.X
        pdy = cursorloc.Y
    
    Case WM_EXITSIZEMOVE
       ' Exit modal sizing/moving loop
       
        While mTrans < 254
            If mTrans > 254 Then mTrans = 254
            SetWindowEffects Me.hWnd, , mTrans
            mTrans = mTrans + 16
            DoEvents
        Wend
        mTrans = 254
        SetWindowEffects Me.hWnd, , mTrans
       
    Case WM_MOVING
        If mTrans > 128 Then
            SetWindowEffects Me.hWnd, , mTrans
            mTrans = mTrans - 8
        End If
       
       ' Form is moving:
       CopyMemory tR, ByVal lParam, Len(tR)
       lWidth = tR.Right - tR.Left
       lHeight = tR.Bottom - tR.Top
       
        GetCursorPos cursorloc
        dx = cursorloc.X - pdx
        dy = cursorloc.Y - pdy
               
        Dim snapped As Boolean
        
        SnapToScreen dx, dy, tR, snapped
        'SnapToForm dx, dy, tR, frmPopupCal, snapped
        'SnapToForm dx, dy, tR, frmPopup, snapped
       
       
       tR.Right = tR.Left + lWidth
       tR.Bottom = tR.Top + lHeight
       CopyMemory ByVal lParam, tR, Len(tR)
    
    End Select
    
End Sub

Private Sub Label1_DblClick()
    picBtnMini_Click
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetRagCursor CUR_HANDCLOSE
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetRagCursor CUR_HANDOPEN
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetRagCursor CUR_HANDOPEN
End Sub

Private Sub picBtnExit_Click()
    SetAlwaysOnTop Me, HWND_NOTOPMOST
    'Me.Hide
    Unload Me
End Sub

Private Sub picBtnExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnExit.Picture = frmPopup.ImageList1.ListImages("closeon").Picture
    picBtnMini.Picture = frmPopup.ImageList1.ListImages("minioff").Picture
    picBtnNothing.Picture = frmPopup.ImageList1.ListImages("buttonoff").Picture
End Sub

Private Sub picBtnMini_Click()
    winstate = winstate + 1
    If winstate > 1 Then winstate = 0
    If winstate = 0 Then
        Me.Height = 320 * Screen.TwipsPerPixelY
        If Me.Top > Screen.Height - Me.Height Then
            Me.Top = Screen.Height - Me.Height
        End If
    Else
        Me.Height = 17 * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub picBtnMini_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnExit.Picture = frmPopup.ImageList1.ListImages("closeoff").Picture
    picBtnMini.Picture = frmPopup.ImageList1.ListImages("minion").Picture
    picBtnNothing.Picture = frmPopup.ImageList1.ListImages("buttonoff").Picture
End Sub

Private Sub picBtnNothing_Click()
    PopupMenu mnuForm.mnuOptions
End Sub

Private Sub picBtnNothing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnExit.Picture = frmPopup.ImageList1.ListImages("closeoff").Picture
    picBtnMini.Picture = frmPopup.ImageList1.ListImages("minioff").Picture
    picBtnNothing.Picture = frmPopup.ImageList1.ListImages("buttonon").Picture
End Sub

Private Sub RagButtonClose_Click()
    SetAlwaysOnTop Me, HWND_NOTOPMOST
    Unload Me
End Sub

Private Sub ragButtonOK_Click()
    SetAlwaysOnTop Me, HWND_NOTOPMOST
    Unload Me
End Sub
