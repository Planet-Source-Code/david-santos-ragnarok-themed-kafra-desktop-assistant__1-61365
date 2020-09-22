VERSION 5.00
Begin VB.Form frmNewTrunk 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   Picture         =   "frmNewTrunk.frx":0000
   ScaleHeight     =   144
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   Begin KafraDesk.RagButtonResource RagButtonResource1 
      Left            =   120
      Top             =   1560
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin KafraDesk.SuperRagButton RagButtonClose 
      Height          =   300
      Left            =   3360
      TabIndex        =   7
      Top             =   1680
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
      Caption         =   "cancel"
      ResourceControl =   "RagButtonResource1"
   End
   Begin KafraDesk.SuperRagButton RagButtonOK 
      Height          =   300
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
      ResourceControl =   "RagButtonResource1"
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7F7F7&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   3735
   End
   Begin VB.PictureBox picBtnMini 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3765
      Picture         =   "frmNewTrunk.frx":A198
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
      Picture         =   "frmNewTrunk.frx":A2D8
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   2
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
      Picture         =   "frmNewTrunk.frx":A418
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   1
      Top             =   45
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the name of the new folder you would like to create:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   3720
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "New Storage Folder"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmNewTrunk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim winstate As Integer

Private Sub Form_Activate()
    SetAlwaysOnTop Me, HWND_TOPMOST
    Me.Top = frmPopup.Top + frmPopup.Height / 2 - Me.Height / 2
    Me.Left = frmPopup.Left
End Sub

Private Sub Form_Load()
    CornerFormShape Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ReleaseCapture  ' This releases the mouse communication with the form so it can communicate with the operating system to move the form
        Result& = SendMessage(Me.hWnd, WM_SYSCOMMAND, SC_CLICKMOVE, 0)  ' This tells the OS to pick up the form to be moved
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnExit.Picture = frmPopup.ImageList1.ListImages("closeoff").Picture
    picBtnMini.Picture = frmPopup.ImageList1.ListImages("minioff").Picture
    picBtnNothing.Picture = frmPopup.ImageList1.ListImages("buttonoff").Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'SetLayered Me.hWnd, False
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetRagCursor CUR_HANDCLOSE
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetRagCursor CUR_HANDOPEN
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetRagCursor CUR_HANDOPEN
End Sub

Private Sub picBtnExit_Click()
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


Private Sub picBtnNothing_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnExit.Picture = frmPopup.ImageList1.ListImages("closeoff").Picture
    picBtnMini.Picture = frmPopup.ImageList1.ListImages("minioff").Picture
    picBtnNothing.Picture = frmPopup.ImageList1.ListImages("buttonon").Picture
End Sub

Private Sub RagButtonClose_Click()
    Unload Me
End Sub

Private Sub ragButtonOK_Click()
Dim path As String
    path = Trim(txtPath.Text)
    If path = "" Then Exit Sub

    notallowed = "/\*?""<>|"
    For i = 1 To Len(notallowed)
        If InStr(1, path, Mid(notallowed, i, 1)) > 0 Then
            MsgBox "Invalid character in path!  The following characters are not allowed: " & vbCrLf & vbCrLf & "        """ & notallowed & """"
            Exit Sub
        End If
    Next
    
    MkDir curPath & "\" & path
    
    Unload Me
End Sub
