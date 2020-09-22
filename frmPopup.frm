VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPopup 
   BorderStyle     =   0  'None
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   Icon            =   "frmPopup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmPopup.frx":000C
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Drop files and folders here to store them."
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6376
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList3"
      SmallIcons      =   "ImageList3"
      ForeColor       =   -2147483640
      BackColor       =   16250871
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item Name"
         Object.Width           =   6350
      EndProperty
   End
   Begin KafraDesk.RagButtonResource RagButtonResource1 
      Left            =   120
      Top             =   4200
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin KafraDesk.SuperRagButton RagButtonNext 
      Height          =   300
      Left            =   840
      TabIndex        =   16
      Top             =   3840
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
      Caption         =   ">>"
      ResourceControl =   "RagButtonResource1"
   End
   Begin KafraDesk.SuperRagButton RagButtonClose 
      Height          =   300
      Left            =   3360
      TabIndex        =   18
      Top             =   4320
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
      Caption         =   "close"
      ResourceControl =   "RagButtonResource1"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2910
      Picture         =   "frmPopup.frx":1610C
      ScaleHeight     =   300
      ScaleWidth      =   105
      TabIndex        =   12
      Top             =   3840
      Width           =   105
   End
   Begin VB.PictureBox picBtnAdd 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3000
      Picture         =   "frmPopup.frx":165EE
      ScaleHeight     =   300
      ScaleWidth      =   390
      TabIndex        =   7
      Tag             =   "Add"
      ToolTipText     =   "Write a new memo"
      Top             =   3840
      Width           =   390
   End
   Begin KafraDesk.SuperRagButton RagButtonMemo 
      Height          =   300
      Left            =   2160
      TabIndex        =   17
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1720
      _ExtentY        =   529
      Caption         =   "storage"
      ResourceControl =   "RagButtonResource1"
   End
   Begin KafraDesk.SuperRagButton RagButtonPrev 
      Height          =   300
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
      Caption         =   "<<"
      ResourceControl =   "RagButtonResource1"
   End
   Begin VB.PictureBox picToolLeft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2430
      Picture         =   "frmPopup.frx":1693E
      ScaleHeight     =   300
      ScaleWidth      =   105
      TabIndex        =   9
      Top             =   3840
      Width           =   105
   End
   Begin VB.PictureBox picBtnAbout 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2520
      Picture         =   "frmPopup.frx":16B02
      ScaleHeight     =   300
      ScaleWidth      =   390
      TabIndex        =   13
      Tag             =   "Add"
      ToolTipText     =   "About Kafra Desktop Assistant"
      Top             =   3840
      Width           =   390
   End
   Begin VB.PictureBox picToolRight 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3840
      Picture         =   "frmPopup.frx":16E5A
      ScaleHeight     =   300
      ScaleWidth      =   105
      TabIndex        =   11
      Top             =   3840
      Width           =   105
   End
   Begin VB.PictureBox picToolMid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3390
      Picture         =   "frmPopup.frx":1707E
      ScaleHeight     =   300
      ScaleWidth      =   105
      TabIndex        =   10
      Top             =   3840
      Width           =   105
   End
   Begin VB.PictureBox picBtnDelete 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3480
      Picture         =   "frmPopup.frx":17560
      ScaleHeight     =   300
      ScaleWidth      =   390
      TabIndex        =   8
      Tag             =   "Delete"
      ToolTipText     =   "Delete current memo"
      Top             =   3840
      Width           =   390
   End
   Begin VB.PictureBox picBtnMini 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3720
      Picture         =   "frmPopup.frx":178E8
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   6
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
      Picture         =   "frmPopup.frx":17A28
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   5
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
      Picture         =   "frmPopup.frx":17B68
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   45
      Width           =   135
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   120
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":17CA8
            Key             =   "redpotion"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1833C
            Key             =   "apple"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":189D0
            Key             =   "sword"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":19064
            Key             =   "trunk"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":196F8
            Key             =   "box"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":19D8C
            Key             =   "book"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1A420
            Key             =   "scroll"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1AAB4
            Key             =   "note"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1B148
            Key             =   "openbook"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1B7DC
            Key             =   "lute"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1BE70
            Key             =   "specs"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1C504
            Key             =   "specs2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   120
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   26
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1CB98
            Key             =   "addoff"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1CEF8
            Key             =   "addover"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1D320
            Key             =   "adddown"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1D704
            Key             =   "deleteoff"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1DA9C
            Key             =   "deleteover"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1DF60
            Key             =   "deletedown"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1E3B4
            Key             =   "aboutoff"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1E71C
            Key             =   "aboutover"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1EB60
            Key             =   "aboutdown"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   3600
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
            Picture         =   "frmPopup.frx":1EF24
            Key             =   "closeon"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1F074
            Key             =   "closeoff"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1F1C4
            Key             =   "minion"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1F314
            Key             =   "minioff"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1F464
            Key             =   "buttonon"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPopup.frx":1F5B4
            Key             =   "buttonoff"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00F7F7F7&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmPopup.frx":1F704
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kafra Desktop Assistant"
      Height          =   255
      Left            =   360
      TabIndex        =   14
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
      TabIndex        =   3
      Top             =   360
      Width           =   2325
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2565
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curMsg As Integer
Dim lastMsg As Integer
Dim memo() As String
Dim winstate As Integer
Dim curPointer As Integer

'Dim moveForm As Boolean
Dim pdx As Long
Dim pdy As Long

Dim mTrans As Integer

Dim mySubClass As New cSubclass
Implements WinSubHook.iSubclass
Dim memomode As Boolean

Property Get Messages() As Integer
    Messages = lastMsg + 1
End Property

Private Sub SaveMsg()
Dim myfile As Long
Dim i As Integer
Dim savememo As String

    myfile = FreeFile
    If lastMsg < 0 Then Exit Sub
    Open App.path & "\memos.txt" For Output As myfile
    For i = 0 To lastMsg
        savememo = Replace(memo(i), vbCrLf, "|")
        Print #myfile, savememo
    Next
    Close myfile
End Sub

Private Sub LoadMsg()
Dim myfile As Long
Dim savememo As String

    myfile = FreeFile

    If Dir(App.path & "\memos.txt") = "" Then
        Exit Sub
    End If
    
    Open App.path & "\memos.txt" For Input As myfile
    
    lastMsg = -1
    While Not EOF(myfile)
        Line Input #myfile, savememo
        If savememo <> "" Then
            lastMsg = lastMsg + 1
            ReDim Preserve memo(lastMsg)
            memo(lastMsg) = Replace(savememo, "|", vbCrLf)
        End If
    Wend
    
    Close myfile
End Sub

Public Sub ShowItems(dirpath As String)
    Dim i As Integer
    Dim mIcon As String
    Dim a As String
    Dim fName As String
    
    ListView1.ListItems.Clear
    
    a = Dir(dirpath & "\*.*", vbDirectory)
    While a <> ""
        If GetAttr(dirpath & "\" & a) = vbDirectory Then
            fName = UCase(Left(a, 1)) & LCase(Right(a, Len(a) - 1))
            ListView1.ListItems.Add , , "\" & fName, "trunk", "trunk"
            i = i + 1
        End If
        a = Dir
    Wend
    
    
    a = Dir(dirpath & "\*.*", vbNormal)
    While a <> ""
        Select Case LCase(Right(a, 4))
        Case ".exe", ".lnk"
            mIcon = "apple"
        Case ".zip", ".rar"
            mIcon = "box"
        Case ".doc", ".txt", ".rtf", ".log"
            mIcon = "scroll"
        Case ".mp3", ".m3u", ".wav", ".mid", ".wma", ".spc", ".psf", ".wma"
            mIcon = "lute"
        Case ".bmp", ".jpg", ".jpe", ".gif", ".tga", ".tif", ".psd", ".pcx", ".png", ".wmf"
            mIcon = "specs"
        Case ".mpg", ".mpe", ".avi", ".mov", ".asf", ".wmv"
            mIcon = "specs2"
        Case Else
            mIcon = "book"
        End Select
        fName = UCase(Left(a, 1)) & LCase(Right(a, Len(a) - 1))
        ListView1.ListItems.Add , , fName, mIcon, mIcon
        i = i + 1
        a = Dir
    Wend
    
    ListView1.Arrange = lvwAutoTop
    'ListView1.SortKey = 0
    'ListView1.Sorted = True
    
    lblStorage.Caption = i & " item(s) in storage."
End Sub

Private Sub Form_Activate()
    If picBtnAdd.Tag = "Add" Then
        SetAlwaysOnTop Me, HWND_TOPMOST
        
        SetLayered Me.hWnd, True
        SetWindowEffects Me.hWnd, , 254
        'ShowIntro
        curMsg = lastMsg
        UpdateMsgs
        ShowItems curPath
    End If

    If Me.Left > Screen.Width Then
        Me.Left = Screen.Width - Me.Width
    End If

    If Me.Top > Screen.Height Then
        Me.Top = Screen.Height - Me.Height
    End If
End Sub

Private Sub ShowIntro()
    txtInfo.Text = "[Kafra]" & vbCrLf
    txtInfo.Text = txtInfo.Text & "I'm " & KafraName & ", your Kafra Desktop Assistant!" & vbCrLf
    txtInfo.Text = txtInfo.Text & "We will stay with you wherever you go!" & vbCrLf & vbCrLf
    txtInfo.Text = txtInfo.Text & "You have " & lastMsg + 1 & " memo" & IIf(lastMsg + 1 = 1, "", "s") & "."
    lblMemo.Caption = "Kafra Memo Assistant"
End Sub

Private Sub Form_Load()
    
    CornerFormShape Me

    
    curPath = App.path & "\Storage"
    curPointer = 1
    
    curMsg = -1
    lastMsg = -1
    
    LoadMsg
    
    mySubClass.Subclass Me.hWnd, Me
    mySubClass.AddMsg WM_EXITSIZEMOVE, MSG_AFTER
    mySubClass.AddMsg WM_EXITSIZEMOVE, MSG_BEFORE
    mySubClass.AddMsg WM_MOVING, MSG_BEFORE
    mySubClass.AddMsg WM_LBUTTONDOWN, MSG_BEFORE
    
    mTrans = 255

    memomode = True
    'RagButtonMemo.HighLight memomode
    RagButtonMemo.Caption = "storage"
    ListView1.Visible = Not memomode

    Dim newleft As Single
    Dim newtop As Single
    
    newleft = GetHexKey(HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "MemoPad X") * Screen.TwipsPerPixelX
    newtop = GetHexKey(HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "MemoPad Y") * Screen.TwipsPerPixelY

    If newleft > Screen.Width Then
        newleft = Screen.Width - Me.Width
    End If

    If newtop > Screen.Height Then
        newtop = Screen.Height - Me.Height
    End If


    Me.Move newleft, newtop


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
        
        picBtnAdd.Picture = frmPopup.ImageList2.ListImages("addoff").Picture
        picBtnDelete.Picture = frmPopup.ImageList2.ListImages("deleteoff").Picture
        picBtnAbout.Picture = frmPopup.ImageList2.ListImages("aboutoff").Picture
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveMsg
    mySubClass.UnSubclass
    Set mySubClass = Nothing
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
Dim tR As RECT
Dim lWidth As Long
Dim lHeight As Long
Dim cursorloc As POINT_TYPE
Dim dx As Long
Dim dy As Long


    Select Case uMsg
    Case WM_EXITSIZEMOVE
       ' Exit modal sizing/moving loop
        SetHexKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "MemoPad X", (Me.Left \ Screen.TwipsPerPixelX)
        SetHexKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "MemoPad Y", (Me.Top \ Screen.TwipsPerPixelY)
    End Select



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
        pdx = cursorloc.X - Me.Left / Screen.TwipsPerPixelX
        pdy = cursorloc.Y - Me.Top / Screen.TwipsPerPixelY
        
    Case WM_EXITSIZEMOVE
       ' Exit modal sizing/moving loop
                
        pdx = 0
        pdy = 0
        
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
        
        Dim snapped As Boolean
        
        GetCursorPos cursorloc
        'dx = cursorloc.X - pdx
        'dy = cursorloc.Y - pdy
        
        SnapToScreen cursorloc.X - pdx, cursorloc.Y - pdy, tR, snapped, 10
        
        'SnapToForm dx, dy, tR, frmPopupCal, snapped
        'SnapToForm dx, dy, tR, frmPopupTools, snapped
       
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

Private Sub ListView1_DblClick()
Dim newPath As String
    If Left(ListView1.SelectedItem.Text, 1) = "\" Then
        newPath = ListView1.SelectedItem.Text
        ShowItems curPath & ListView1.SelectedItem.Text
        curPath = curPath & newPath
    Else
        ShellExecute Me.hWnd, "open", curPath & "\" & ListView1.SelectedItem.Text, "", curPath, 1
    End If
End Sub

Public Sub OpenFolder()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
        ShellExecute Me.hWnd, "explore", curPath & "\" & ListView1.SelectedItem.Text, "", "", 1
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuForm.mnuFolder
End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandler
Dim mFile As New FileSystemObject
Dim i As Integer

    For i = 1 To Data.Files.Count
        If GetAttr(Data.Files(i)) = vbDirectory Then
            mFile.MoveFolder Data.Files(i), curPath & "\" & mFile.GetFileName(Data.Files(i))
        Else
            mFile.MoveFile Data.Files(i), curPath & "\" & mFile.GetFileName(Data.Files(i))
        End If
    Next
    ShowItems curPath
Exit Sub
errhandler:
    MsgBox "One or more of the items could not be moved.", vbInformation, "DesKafra"
End Sub

Private Sub ListView1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    SetRagCursor CUR_TARGET
End Sub

Private Sub picBtnAbout_Click()
    Randomize Timer
    frmAbout.Show
End Sub

Private Sub picBtnAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnAbout.Picture = frmPopup.ImageList2.ListImages("aboutdown").Picture
End Sub

Private Sub picBtnAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnAdd.Picture = frmPopup.ImageList2.ListImages("addoff").Picture
    picBtnDelete.Picture = frmPopup.ImageList2.ListImages("deleteoff").Picture
    picBtnAbout.Picture = frmPopup.ImageList2.ListImages("aboutover").Picture
End Sub

Private Sub picBtnAbout_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnAbout.Picture = frmPopup.ImageList2.ListImages("aboutover").Picture
End Sub

Private Sub picBtnAdd_Click()
    If picBtnAdd.Tag = "Add" Then
        picBtnAdd.Tag = "OK"
        picBtnDelete.Tag = "Cancel"
        
        picBtnAdd.ToolTipText = "Accept new memo"
        picBtnDelete.ToolTipText = "Cancel new memo"
        
        txtInfo.Locked = False
        lblMemo.Caption = "Writing Memo #" & lastMsg + 1
        lblMemo.ForeColor = RGB(255, 0, 0)
        txtInfo.SetFocus
        txtInfo.Text = Format(Date, "mmm d, yyyy") & " " & Time & vbCrLf & vbCrLf
        
        txtInfo.SelStart = Len(txtInfo.Text)
        txtInfo.SelLength = 0
    
    ElseIf picBtnAdd.Tag = "OK" Then
        If Trim(txtInfo.Text) = "" Then Exit Sub
        
        lastMsg = lastMsg + 1
        curMsg = lastMsg
        
        ReDim Preserve memo(curMsg)
        memo(curMsg) = txtInfo.Text
        picBtnAdd.Tag = "Add"
        picBtnDelete.Tag = "Delete"
        
        lblMemo.Caption = "Memo #" & curMsg + 1
        
        lblMemo.ForeColor = RGB(0, 0, 128)
        picBtnAdd.ToolTipText = "Write a new memo"
        picBtnDelete.ToolTipText = "Delete current memo"
                
        SaveMsg
        
        txtInfo.Locked = True
    End If
End Sub

Private Sub picBtnAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnAdd.Picture = frmPopup.ImageList2.ListImages("adddown").Picture
End Sub

Private Sub picBtnAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetRagCursor CUR_POINT
    picBtnAdd.Picture = frmPopup.ImageList2.ListImages("addover").Picture
    picBtnDelete.Picture = frmPopup.ImageList2.ListImages("deleteoff").Picture
    picBtnAbout.Picture = frmPopup.ImageList2.ListImages("aboutoff").Picture
End Sub

Private Sub picBtnAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnAdd.Picture = frmPopup.ImageList2.ListImages("addover").Picture
End Sub

Private Sub picBtnDelete_Click()
Dim i As Integer

    If picBtnAdd.Tag = "Add" Then
        'we're in view mode
        If curMsg = -1 Then Exit Sub
        
        SetAlwaysOnTop Me, HWND_NOTOPMOST
        
        If MsgBox("Delete this memo?", vbQuestion + vbYesNo, "Kafra Memopad") = vbYes Then
            lastMsg = lastMsg - 1
            For i = curMsg To lastMsg
                memo(i) = memo(i + 1)
            Next
            
            If lastMsg >= 0 Then ReDim Preserve memo(lastMsg)
            ragButtonNext_Click
        End If
        
        SetAlwaysOnTop Me, HWND_TOPMOST
        
    ElseIf picBtnAdd.Tag = "OK" Then
        'we're in ADD mode
        'goto last message
        lblMemo.ForeColor = RGB(0, 0, 128)
        If curMsg > 0 Then
            lblMemo.Caption = "Memo #" & curMsg
            txtInfo.Text = memo(curMsg)
        Else
            ShowIntro
        End If
        txtInfo.Locked = True
        picBtnAdd.Tag = "Add"
        picBtnDelete.Tag = "Delete"
        picBtnAdd.ToolTipText = "Write a new memo"
        picBtnDelete.ToolTipText = "Delete current memo"
    End If
End Sub

Private Sub picBtnDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnDelete.Picture = frmPopup.ImageList2.ListImages("deletedown").Picture
End Sub

Private Sub picBtnDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnAdd.Picture = frmPopup.ImageList2.ListImages("addoff").Picture
    picBtnDelete.Picture = frmPopup.ImageList2.ListImages("deleteover").Picture
    picBtnAbout.Picture = frmPopup.ImageList2.ListImages("aboutoff").Picture
End Sub

Private Sub picBtnDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtnDelete.Picture = frmPopup.ImageList2.ListImages("deleteover").Picture
End Sub

Private Sub picBtnExit_Click()
    SetAlwaysOnTop Me, HWND_NOTOPMOST
    Me.Hide
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

Public Sub ResetStorage()
    curPath = App.path & "\Storage"
    ShowItems curPath
End Sub

Private Sub UpdateMsgs()
    If curMsg < 0 Then
        ShowIntro
    Else
        txtInfo.Text = memo(curMsg)
        lblMemo.Caption = "Memo #" & curMsg + 1
    End If
End Sub

Private Sub RagButtonClose_Click()
    SetAlwaysOnTop Me, HWND_NOTOPMOST
    Me.Hide
End Sub

Private Sub RagButtonMemo_Click()
    memomode = Not memomode
    If memomode Then
        RagButtonMemo.Caption = "storage"
    Else
        RagButtonMemo.Caption = "memo"
    End If
    'RagButtonMemo.HighLight memomode
    ListView1.Visible = Not memomode
End Sub

Private Sub ragButtonNext_Click()
    If picBtnAdd.Tag = "OK" Then Exit Sub

    curMsg = curMsg + 1
    If curMsg > lastMsg Then curMsg = lastMsg
    
    UpdateMsgs
End Sub

Private Sub ragButtonPrev_Click()
    If picBtnAdd.Tag = "OK" Then Exit Sub
    
    curMsg = curMsg - 1

    UpdateMsgs
End Sub


