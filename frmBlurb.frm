VERSION 5.00
Begin VB.Form frmBlurb 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   Picture         =   "frmBlurb.frx":0000
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close this Window... (10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2520
      MouseIcon       =   "frmBlurb.frx":762C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   840
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open the MemoPad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      MouseIcon       =   "frmBlurb.frx":777E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmBlurb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const XPFORMS = True

Dim cnt As Integer
Dim myRGN As New cDIBSectionRegion

Private Sub Form_Load()
        
    myRGN.LoadFromResource 108
    myRGN.Applied(Me.hwnd) = True
    
    #If XPFORMS Then
        SetLayered Me.hwnd, True
        SetWindowEffects Me.hwnd, , 191
    #End If
    
    Me.Top = frmMain.Top - 50
    Me.Left = frmMain.Left + (frmMain.Width / 2) - Me.Width - 50
    cnt = 9
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

    #If XPFORMS Then
        For i = 191 To 0 Step -6
            SetWindowEffects Me.hwnd, , CByte(i)
            DoEvents
        Next
    #End If
    
    myRGN.Destroy
    Set myRGN = Nothing
End Sub

Private Sub Label1_Click()
    Unload Me
    frmPopup.Show , frmMain
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Label2.Caption = "Close this Window... (" & cnt & ")"
    cnt = cnt - 1
    If cnt < 0 Then Unload Me
End Sub

Public Sub SetCaption(NewCaption As String)
    lblCaption.Caption = NewCaption
End Sub
