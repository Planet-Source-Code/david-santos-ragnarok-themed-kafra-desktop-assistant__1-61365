VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl RagButtonResource 
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   840
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   840
   ScaleWidth      =   840
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RagButtonResource.ctx":0000
            Key             =   "left"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RagButtonResource.ctx":0234
            Key             =   "mid"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RagButtonResource.ctx":0328
            Key             =   "right"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RagButtonResource.ctx":055C
            Key             =   "leftup"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RagButtonResource.ctx":0790
            Key             =   "midup"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RagButtonResource.ctx":0884
            Key             =   "rightup"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "RagButtonResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Resize()
    If UserControl.Width < 500 Then UserControl.Width = 500
    If UserControl.Width > 500 Then UserControl.Width = 500
    If UserControl.Height < 500 Then UserControl.Height = 500
    If UserControl.Height > 500 Then UserControl.Height = 500
   
End Sub

Property Get ImageResource()
    Set ImageResource = ImageList1
End Property
