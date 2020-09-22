VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout Sakura.frx":0000
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7F7F7&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmAbout Sakura.frx":A198
      Top             =   360
      Width           =   3975
   End
   Begin VB.PictureBox picBtnNothing 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   75
      Picture         =   "frmAbout Sakura.frx":A24B
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   2
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
      Left            =   4005
      Picture         =   "frmAbout Sakura.frx":A38B
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   1
      Top             =   45
      Width           =   135
   End
   Begin VB.PictureBox picBtnMini 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3765
      Picture         =   "frmAbout Sakura.frx":A4CB
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   0
      Top             =   45
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "About KDA"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   30
      Width           =   3135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim winstate As Integer
Dim abtTrans As Integer

Private Sub Form_Load()
    CornerFormShape Me
    

    
    Me.Top = frmPopup.Top
    Me.Left = frmPopup.Left - Me.Width - 200
    If Me.Left < 0 Then Me.Left = frmPopup.Left + frmPopup.Width + 200
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        ReleaseCapture  ' This releases the mouse communication with the form so it can communicate with the operating system to move the form
        Result& = SendMessage(Me.hwnd, WM_SYSCOMMAND, SC_CLICKMOVE, 0)   ' This tells the OS to pick up the form to be moved
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 1 Then
        If abtTrans > 128 Then
            SetWindowEffects hwnd, , abtTrans
            abtTrans = abtTrans - 16
        End If
    Else
       '#If XPFORMS Then
        While abtTrans < 254
            SetWindowEffects Me.hwnd, , abtTrans
            abtTrans = abtTrans + 16
            DoEvents
        Wend
        abtTrans = 254
        SetWindowEffects Me.hwnd, , abtTrans
      '#End If
    End If
    
    picBtnExit.Picture = frmPopup.ImageList1.ListImages("closeoff").Picture
    picBtnMini.Picture = frmPopup.ImageList1.ListImages("minioff").Picture
    picBtnNothing.Picture = frmPopup.ImageList1.ListImages("buttonoff").Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'SetLayered Me.hWnd, False
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetRagCursor CUR_HANDCLOSE
    Form_MouseDown Button, Shift, x, y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetRagCursor CUR_HANDOPEN
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetRagCursor CUR_HANDOPEN
End Sub

Private Sub picBtnExit_Click()
    Unload Me
End Sub

Private Sub picBtnExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
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


Private Sub picBtnMini_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picBtnExit.Picture = frmPopup.ImageList1.ListImages("closeoff").Picture
    picBtnMini.Picture = frmPopup.ImageList1.ListImages("minion").Picture
    picBtnNothing.Picture = frmPopup.ImageList1.ListImages("buttonoff").Picture
End Sub


Private Sub picBtnNothing_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picBtnExit.Picture = frmPopup.ImageList1.ListImages("closeoff").Picture
    picBtnMini.Picture = frmPopup.ImageList1.ListImages("minioff").Picture
    picBtnNothing.Picture = frmPopup.ImageList1.ListImages("buttonon").Picture
End Sub

