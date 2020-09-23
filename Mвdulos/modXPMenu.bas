Attribute VB_Name = "modXPMenu"
Option Explicit
 
 Public Const MAX_ITEMS As Integer = 64
 
 Public Type XPMenuItem
  Items(MAX_ITEMS)  As String
  Status(MAX_ITEMS) As String
 End Type

 Public Type RGBQUAD
  rgbBlue      As Byte
  rgbGreen     As Byte
  rgbRed       As Byte
  rgbReserved  As Byte
 End Type
 
 Public Type POINTAPI
  X As Long
  Y As Long
 End Type
  
 Public mnuName           As String
 
 Public Color1            As Long
 Public Color2            As Long
 Public ContFont          As Long
 Public vIndex            As Long
 
 Public dGradient         As Boolean
 Public SetBackGround     As Boolean
 Public myMargin          As Boolean
 
 Public BackPicture       As New StdPicture
 Public myMenu            As OptionMenu
 Public TypeGradient      As TipoGradiente
 
 '* Image List
 Public ImageLst          As ImageList
  
 Public MenuXpItem        As XPMenuItem
 
 Public frmMenu           As New frmXPMenu
 
 Public myBackColor       As OLE_COLOR
 Public myHighColor       As OLE_COLOR
 Public myBorderColor     As OLE_COLOR
 Public mySeparatorColor  As OLE_COLOR
 Public myMarginColor     As OLE_COLOR
 Public myForeColor       As OLE_COLOR
 Public myHighBorderColor As OLE_COLOR
 Public mySelForeColor    As OLE_COLOR
 Public myShadowColor     As OLE_COLOR
 Public m_clrArrowNormal  As OLE_COLOR
 Public m_clrArrowSelect  As OLE_COLOR
 
 Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
 Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
  
Public Sub HandleClick(ByVal MenuName As String, ByVal ItemNum As Integer, ByVal StrItemText As String)
 Call MsgBox("Menu Name: " & MenuName & vbCrLf & "Item Number: " & ItemNum & vbCrLf & "Item Text: " & StrItemText)
End Sub

Public Sub MenuOver(ByVal MenuText As String)
 Dim MSG As String
 Dim Item As Integer
   
 For Item = 0 To MAX_ITEMS
  If (MenuXpItem.Items(Item) = MenuText) Then MSG = MenuXpItem.Status(Item)
 Next
 frmTestForm.sbr1.Panels(1).Text = MSG
End Sub

Public Sub ClearStatus()
 frmTestForm.sbr1.Panels(1).Text = ""
End Sub

Public Sub MenuExit()
 Dim f As Form
    
 For Each f In Forms
  Call Unload(f)
 Next
 End
End Sub

Public Sub Long2RGB(ByVal LngColor As Long, ByRef RgbColor As RGBQUAD)
 Dim Aux As Byte
 
 Call CopyMemory(RgbColor, LngColor, 4)
 Aux = RgbColor.rgbBlue
 RgbColor.rgbBlue = RgbColor.rgbRed
 RgbColor.rgbRed = Aux
End Sub
