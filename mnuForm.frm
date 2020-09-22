VERSION 5.00
Begin VB.Form mnuForm 
   ClientHeight    =   3285
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnu1 
      Caption         =   "Select1"
      Begin VB.Menu mnuMemoPad 
         Caption         =   "&Memo Pad"
      End
      Begin VB.Menu mnuBlankA 
         Caption         =   "-"
      End
      Begin VB.Menu mnunext 
         Caption         =   "&Switch Kafra"
         Begin VB.Menu mnuKafraSwitch 
            Caption         =   "&Pavianne"
            Index           =   0
         End
         Begin VB.Menu mnuKafraSwitch 
            Caption         =   "&Blossom"
            Index           =   1
         End
         Begin VB.Menu mnuKafraSwitch 
            Caption         =   "&Jasmine"
            Index           =   2
         End
         Begin VB.Menu mnuKafraSwitch 
            Caption         =   "&Roxie"
            Index           =   3
         End
         Begin VB.Menu mnuKafraSwitch 
            Caption         =   "&Leila"
            Index           =   4
         End
         Begin VB.Menu mnuKafraSwitch 
            Caption         =   "&Curly Sue"
            Index           =   5
         End
         Begin VB.Menu mnuKafraSwitch 
            Caption         =   "&Sampaguita"
            Index           =   6
         End
         Begin VB.Menu mnuBlank6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRandom 
            Caption         =   "&Random Kafra"
         End
      End
      Begin VB.Menu mnuBlank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKafraSettings 
         Caption         =   "&Options"
         Begin VB.Menu mnuAutorun 
            Caption         =   "&Autorun"
         End
         Begin VB.Menu mnuAlwaysOnTop 
            Caption         =   "Always on &Top"
         End
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "&Show Kafra"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "&Hide Kafra"
      End
      Begin VB.Menu mnuBlank0 
         Caption         =   "-"
      End
      Begin VB.Menu mnux 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Select2"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset Storage"
      End
   End
   Begin VB.Menu mnuFolder 
      Caption         =   "Select3"
      Begin VB.Menu mnuNewTrunk 
         Caption         =   "&New Trunk"
      End
      Begin VB.Menu mnuExplore 
         Caption         =   "&Explore"
      End
   End
End
Attribute VB_Name = "mnuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuAlwaysOnTop_Click()
    mnuAlwaysOnTop.Checked = Not mnuAlwaysOnTop.Checked
    
    If mnuAlwaysOnTop.Checked Then
        SetStringKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Always On Top", "Yes"
    Else
        SetStringKey HKEY_CURRENT_USER, "Software\EnderSoft\DesktopKafra\", "Always On Top", "No"
    End If
    
    SetAlwaysOnTop frmMain, IIf(mnuAlwaysOnTop.Checked, HWND_TOPMOST, HWND_NOTOPMOST)
End Sub

Private Sub mnuAutorun_Click()
    mnuAutorun.Checked = Not mnuAutorun.Checked
    SetAutoRun mnuAutorun.Checked
End Sub

Private Sub mnuCalendar_Click()
    frmPopupCal.Show
End Sub

Private Sub mnuExplore_Click()
    frmPopup.OpenFolder
End Sub

Private Sub mnuHide_Click()
    frmMain.HideKafra
End Sub

Private Sub mnuKafraSwitch_Click(Index As Integer)
    frmMain.SwitchKafra Index
End Sub

Private Sub mnuMemoPad_Click()
    frmPopup.Show
End Sub

Private Sub mnuNewTrunk_Click()
    frmNewTrunk.Show vbModal, frmPopup
End Sub

Private Sub mnuRandom_Click()
    frmMain.RandomKafra
End Sub

Private Sub mnuReset_Click()
    frmPopup.ResetStorage
End Sub

Private Sub mnuShow_Click()
    frmMain.ShowKafra
End Sub

Private Sub mnuTools_Click()
    frmPopupTools.Show
End Sub

Private Sub mnux_Click()
    Unload frmMain
End Sub
