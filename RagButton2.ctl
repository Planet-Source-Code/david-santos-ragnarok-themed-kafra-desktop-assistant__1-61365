VERSION 5.00
Begin VB.UserControl SuperRagButton 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1260
   PropertyPages   =   "RagButton2.ctx":0000
   ScaleHeight     =   675
   ScaleWidth      =   1260
   ToolboxBitmap   =   "RagButton2.ctx":0014
   Begin VB.PictureBox picButtonMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   0
      Tag             =   "prev"
      Top             =   0
      Width           =   630
   End
End
Attribute VB_Name = "SuperRagButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Private Const WM_MOUSEHOVER = &H2A1&
Private Const WM_MOUSELEAVE = &H2A3&

Private Const TME_HOVER = &H1&
Private Const TME_LEAVE = &H2&
'Private Const TME_QUERY = &H40000000
'Private Const TME_CANCEL = &H80000000

Private Const HOVER_DEFAULT = &HFFFFFFFF

Private Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Declare Function TrackMouseEvent Lib "user32" _
   (lpEventTrack As tagTRACKMOUSEEVENT) As Long

Public Event Click()
Public Event DblClick()
Public Event MouseMove()
Public Event MouseDown()
Public Event MouseUp()

'Dim mCount As Long
Dim mStyle As String
Dim myRsrc As String

Dim mySubClass As cSubclass
Implements iSubclass
Dim m_bTracking As Boolean
Dim myImageList As ImageList
Dim m_Parent As Form
Dim m_RsrcControl As RagButtonResource
Dim HighLightMode As Boolean

'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
'
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
Dim tR As RECT
    
    Select Case uMsg
    Case WM_SIZING
        Stop
    Case WM_MOUSELEAVE
        'picButtonMain.Picture = ImageList1.ListImages(mStyle).Picture
        HighLightMode = False
        m_bTracking = False
        DrawBG HighLightMode
        WriteText
    End Select

End Sub

Private Sub WriteText(Optional down As Boolean = False)
Dim tWidth As Long
Dim tHeight As Long
Dim xPos As Long
Dim yPos As Long


    tWidth = CLng(picButtonMain.TextWidth(mStyle))
    tHeight = CLng(picButtonMain.TextHeight(mStyle))
    xPos = (picButtonMain.Width / 15) / 2 - tWidth / 15 / 2
    yPos = (picButtonMain.Height / 15) / 2 - tHeight / 15 / 2
    
    'If HighLightMode = True Then
    '    picButtonMain.ForeColor = RGB(138, 158, 196)
    'Else
    '    picButtonMain.ForeColor = RGB(186, 186, 186)
    'End If
    
    'TextOut picButtonMain.hdc, xPos + 1, yPos + 1, mStyle, Len(mStyle)
    
    If down Then
        picButtonMain.ForeColor = RGB(255, 255, 255)
        TextOut picButtonMain.hdc, xPos, yPos, mStyle, Len(mStyle)
        picButtonMain.ForeColor = RGB(0, 0, 0)
        TextOut picButtonMain.hdc, xPos + 1, yPos + 1, mStyle, Len(mStyle)
    Else
        picButtonMain.ForeColor = RGB(0, 0, 0)
        TextOut picButtonMain.hdc, xPos, yPos, mStyle, Len(mStyle)
    End If
    
    picButtonMain.Refresh

End Sub

Private Sub picButtonMain_Click()
    RaiseEvent Click
End Sub

Private Sub picButtonMain_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picButtonMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X <= picButtonMain.Width And X >= 0 And Y >= 0 And Y <= picButtonMain.Height Then
        DrawBG True
    Else
        DrawBG False
    End If
    WriteText (Button > 0)
    
    RaiseEvent MouseDown
End Sub

Private Sub picButtonMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tET As tagTRACKMOUSEEVENT
Dim lR As Long

On Error GoTo ErrorHandler
    
    If Not (m_bTracking) Then
        tET.cbSize = Len(tET)
        tET.dwFlags = TME_HOVER Or TME_LEAVE
        tET.dwHoverTime = HOVER_DEFAULT
        tET.hwndTrack = picButtonMain.hWnd
        lR = TrackMouseEvent(tET)
        HighLightMode = True
        m_bTracking = True
    End If
   
    DrawBG HighLightMode
    WriteText (Button > 0)
     
    RaiseEvent MouseMove
    
    Exit Sub

ErrorHandler:
   m_bTracking = False
End Sub

Private Sub picButtonMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp
End Sub

Property Get Parent() As Form
    Set Parent = m_Parent
End Property

Private Sub UserControl_Initialize()
    Set mySubClass = New cSubclass
 
    mySubClass.Subclass picButtonMain.hWnd, Me
    mySubClass.AddMsg WM_MOUSELEAVE, MSG_BEFORE
    mySubClass.AddMsg WM_SIZING, MSG_BEFORE
    
    mStyle = "ok"
    
    picButtonMain.Font.Name = "Arial"
    picButtonMain.Font.Size = 9
    
    
    DrawBG False
    
    WriteText
End Sub

Public Property Let Caption(aStyle As String)
    mStyle = aStyle
    DrawBG HighLightMode
    WriteText
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Caption = mStyle
End Property

Public Property Get ResourceControl() As String
    ResourceControl = myRsrc
End Property

Public Property Let ResourceControl(resourcename As String)
    myRsrc = resourcename
    On Error Resume Next
    Set m_RsrcControl = Nothing
    Set m_RsrcControl = UserControl.Parent.Controls(myRsrc)
    On Error GoTo 0
    DrawBG False
    WriteText
End Property

Private Sub UserControl_InitProperties()
    Set m_Parent = UserControl.Parent
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set m_Parent = UserControl.Parent
    mStyle = LCase(PropBag.ReadProperty("Caption", "ok"))

    myRsrc = PropBag.ReadProperty("ResourceControl", "")
    
    On Error Resume Next
    Set m_RsrcControl = Nothing
    Set m_RsrcControl = UserControl.Parent.Controls(myRsrc)
    On Error GoTo 0
    
    DrawBG False
    
    WriteText
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "Caption", mStyle, "ok"
   PropBag.WriteProperty "ResourceControl", myRsrc, ""
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = picButtonMain.Height
    If UserControl.Width < 25 Then UserControl.Width = 25
    picButtonMain.Width = UserControl.Width
    DrawBG False
    WriteText
End Sub

Private Sub DrawBG(Highlight As Boolean)
Dim ext As String
    ext = "up"
    If Highlight Then ext = ""
    picButtonMain.Cls
    
    If Not m_RsrcControl Is Nothing Then
        Set myImageList = m_RsrcControl.ImageResource
        myImageList.ListImages("left" & ext).Draw picButtonMain.hdc, 0, 0
        For i = 0 To (UserControl.Width / 15) - 16
            myImageList.ListImages("mid" & ext).Draw picButtonMain.hdc, (8 + i) * 15, 0
        Next
        myImageList.ListImages("right" & ext).Draw picButtonMain.hdc, ((UserControl.Width / 15) - 8) * 15, 0
        picButtonMain.Refresh
        Set myImageList = Nothing
    End If
End Sub

Private Sub UserControl_Terminate()
    mySubClass.DelMsg WM_MOUSELEAVE, MSG_BEFORE
    mySubClass.DelMsg WM_SIZING, MSG_BEFORE
    mySubClass.UnSubclass
    Set mySubClass = Nothing
End Sub

