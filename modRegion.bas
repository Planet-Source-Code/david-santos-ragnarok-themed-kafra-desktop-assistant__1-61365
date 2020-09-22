Attribute VB_Name = "modRegion"
Option Explicit

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public myDIBObj As cDIBSection
Public myRegion As cDIBSectionRegion
