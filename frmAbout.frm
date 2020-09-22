VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
      TabIndex        =   4
      Text            =   "frmAbout.frx":0000
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
      Picture         =   "frmAbout.frx":00A7
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   3
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
      Picture         =   "frmAbout.frx":01E7
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   2
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
      Picture         =   "frmAbout.frx":0327
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   1
      Top             =   45
      Width           =   135
   End
   Begin VB.PictureBox picPopup 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   0
      Picture         =   "frmAbout.frx":0467
      ScaleHeight     =   2160
      ScaleWidth      =   4200
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4200
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   4080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   9
         ImageHeight     =   9
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAbout.frx":0F6D
               Key             =   "closeon"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAbout.frx":10BD
               Key             =   "closeoff"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAbout.frx":120D
               Key             =   "minion"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAbout.frx":135D
               Key             =   "minioff"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAbout.frx":14AD
               Key             =   "buttonon"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAbout.frx":15FD
               Key             =   "buttonoff"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "About KDA"
      Height          =   255
      Left            =   240
      TabIndex        =   5
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

Private Sub Form_Load()
    CornerFormShape Me
    
    Me.Picture = picPopup.Picture
    Me.Top = frmPopup.Top
    Me.Left = frmPopup.Left - Me.Width - 200
    If Me.Left < 0 Then Me.Left = frmPopup.Left + frmPopup.Width + 200
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        ReleaseCapture  ' This releases the mouse communication with the form so it can communicate with the operating system to move the form
        Result& = SendMessage(Me.hWnd, WM_SYSCOMMAND, SC_CLICKMOVE, 0)   ' This tells the OS to pick up the form to be moved
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picBtnExit.Picture = ImageList1.ListImages("closeoff").Picture
    picBtnMini.Picture = ImageList1.ListImages("minioff").Picture
    picBtnNothing.Picture = ImageList1.ListImages("buttonoff").Picture
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
    picBtnExit.Picture = ImageList1.ListImages("closeon").Picture
    picBtnMini.Picture = ImageList1.ListImages("minioff").Picture
    picBtnNothing.Picture = ImageList1.ListImages("buttonoff").Picture
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
    picBtnExit.Picture = ImageList1.ListImages("closeoff").Picture
    picBtnMini.Picture = ImageList1.ListImages("minion").Picture
    picBtnNothing.Picture = ImageList1.ListImages("buttonoff").Picture
End Sub


Private Sub picBtnNothing_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picBtnExit.Picture = ImageList1.ListImages("closeoff").Picture
    picBtnMini.Picture = ImageList1.ListImages("minioff").Picture
    picBtnNothing.Picture = ImageList1.ListImages("buttonon").Picture
End Sub
