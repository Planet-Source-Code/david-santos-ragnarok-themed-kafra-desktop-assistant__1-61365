VERSION 5.00
Begin VB.Form frmTray 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1050
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3900
   Icon            =   "frmTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   3900
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This form exists only to act as the system tray icon holder. "
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim nid As NOTIFYICONDATA  ' icon information
    Dim retval As Long  ' return value
    
    ' Put the icon settings into the structure.
    With nid
      .cbSize = Len(nid)  ' size of structure
      .hWnd = Me.hWnd  ' owner of the icon and processor of its messages
      .uID = 1  ' unique identifier for the window's tray icons
      .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP  ' provide icon, message, and tool tip text
      .uCallbackMessage = PK_TRAYICON  ' message to use for icon events
      .hIcon = Me.Icon  ' handle to the icon to actually display in the tray
      .szTip = "Kafra Desktop Assistant" & vbNullChar  ' tool tip text for icon
    End With
    
    ' Add the icon to the system tray.
    retval = Shell_NotifyIcon(NIM_ADD, nid)
    ' Set the new window procedure for window Form1.
    pOldProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim nid As NOTIFYICONDATA  ' icon information
Dim retval As Long  ' return value
    
    ' Load the structure with just the identifying information.
    With nid
        .cbSize = Len(nid)  ' size of structure
        .hWnd = Me.hWnd  ' handle of owning window
        .uID = 1  ' unique identifier
    End With
    retval = Shell_NotifyIcon(NIM_DELETE, nid)
    
    ' Make the old window procedure the current window procedure.
    retval = SetWindowLong(Me.hWnd, GWL_WNDPROC, pOldProc)
End Sub
