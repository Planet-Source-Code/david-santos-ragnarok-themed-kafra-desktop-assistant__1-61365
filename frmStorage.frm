VERSION 5.00
Begin VB.Form frmStorage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin KafraDesk.RagListBox RagList1 
      Height          =   2955
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5159
   End
   Begin KafraDesk.RagButton RagButton1 
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin KafraDesk.RagButton RagButton2 
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   3240
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
      Style           =   "close"
   End
End
Attribute VB_Name = "frmStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    RagList1.Clear
End Sub

Private Sub Command2_Click()
    If RagList1.ListIndex > -1 Then
        RagList1.Remove RagList1.ListIndex
    End If
End Sub

Private Sub Form_Load()
    Me.Show
    RagList1.Clear
    For i = 0 To 99
        RagList1.Add "Item " & i
    Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RagButton1.Deactivate
    RagButton2.Deactivate
End Sub

Private Sub RagButton1_Click()
    RagList1.Add "Item " & RagList1.ListCount
End Sub

Private Sub RagList1_DblClick()
    Me.Caption = RagList1.List(RagList1.ListIndex)
End Sub
