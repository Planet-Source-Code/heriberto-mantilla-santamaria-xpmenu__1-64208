VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXPMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
 '???  ???  ????????  ??????  ???  ???  ???????  ???????  ???????
 '????????  ???  ???  ???     ??? ????  ???  ??  ???  ??  ??   ??
 '????????  ????????  ???     ?????     ???????  ??????   ??   ??
 '???  ???  ???  ???  ??????  ???  ???  ???      ???  ??  ???????
'****************************************************************

'* See the original post: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=29342&lngWId=1
'* Modificado por HACKPRO © TM 2004. Programación a su alcance.

Option Explicit
 
 '* Types
 Private Type TypMenuItem
  IconNum         As Long
  ProcAddr        As Long
  FontColor       As Long
  FontSize        As Long
  Text            As String
  FontName        As String
  ShortCut        As String
  bPopUpMenu      As Boolean
  bSeperator      As Boolean
  Check           As Boolean
  Enabled         As Boolean
  Visible         As Boolean
  FontBold        As Boolean
  FontItalic      As Boolean
  FontStrikethru  As Boolean
  FontUnderline   As Boolean
  ShadowText      As Boolean
  MnuSubMenu      As clsXPMenu
 End Type
 
 Public Enum OptionMenu
  [None] = 0
  [BackGround] = 1
  [Gradient] = 2
 End Enum
 
 Public Enum TipoGradiente
  [Horizontal] = 0
  [Vertical] = 1
  [Diagonal] = 2
 End Enum
 
 'LeftMargin = &HD1D8D8 _
               &H8000000F _
               &HD6BD7B _
               &HD2BDB6 _
               &H6A240A _
               &HFFDFEA, &H990033, &HEA5A00
 
 '* Menu Constants
 Private Const XBuffer            As Long = 4
 Private Const YBuffer            As Long = 4
 Private Const Clr_BackGround     As Long = &HFFFFFF
 Private Const Clr_LeftMargin     As Long = &HDEEDEF
 Private Const FontFlecha         As Long = &H9900FF
 Private Const FontHighLight      As Long = &HFE0099
 Private Const FontHighLightF     As Long = &HBE8100
 Private Const Clr_Seperator      As Long = &HE0E0E0
 Private Const Clr_MenuBorder     As Long = &H6A240A
 Private Const MenuBar            As Long = vbGrayText
 Private Const Clr_HiLightBack    As Long = &HD2BDB6
 Private Const Clr_HiLightBorder  As Long = &H6A240A
 Private Const Dim_MarginWidth    As Long = 23
 Private Const Font_MenuItem      As String = "Tahoma"
 
 Private Const PS_SOLID = 0
 Private Const DST_TEXT = &H2
 Private Const DST_ICON = &H3
 Private Const DSS_NORMAL = &H0
 Private Const DSS_DISABLED = &H20
 Private Const DST_BITMAP = &H3
 Private Const DSS_MONO = &H80
 
 Private kText           As String
 Private cText           As String
 
 Private MnuWidth        As Long
 Private TheTextHeight   As Long
 Private YHiLight        As Long
 Private MenuItemCount   As Long
 Private TextItemCnt     As Long
 Private SepItemCnt      As Long
 Private HiLightedItem   As Long
 Private Maskara         As Long
 
 Private frmMenu         As New frmXPMenu
 Private ActivePopup     As New clsXPMenu
 
 Private bVisible        As Boolean
 Private bPopupShown     As Boolean
 
 '* Menu Array
 Private MenuItems()     As TypMenuItem
      
 '* API's Declaration
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
 Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
 Private Declare Function DrawStateString Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lpString As String, ByVal cbStringLen As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
 Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
 Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Function GetItemProc(ByVal ItemNum As Integer) As String
 If (ItemNum > MenuItemCount) Then
  GetItemProc = 0
  Exit Function
 End If
 GetItemProc = MenuItems(ItemNum).ProcAddr
End Function

Private Sub SetMyShortCut(Optional ByVal KeyMask As ShiftConstants, Optional ByVal KeyCode As Long)
 Select Case KeyMask
  Case 1:    cText = "Shift"
  Case 2:    cText = "Ctrl"
  Case 3:    cText = "Ctrl+Shift"
  Case 4:    cText = "Alt"
  Case 5:    cText = "Shift+Alt"
  Case 6:    cText = "Ctrl+Alt"
  Case 7:    cText = "Ctrl+Shift+Alt"
  Case Else: cText = ""
 End Select
On Error Resume Next
 If (KeyCode <> 0) Then
  Select Case KeyCode
   Case 112:  kText = "F1"
   Case 113:  kText = "F2"
   Case 114:  kText = "F3"
   Case 115:  kText = "F4"
   Case 116:  kText = "F5"
   Case 117:  kText = "F6"
   Case 118:  kText = "F7"
   Case 119:  kText = "F8"
   Case 120:  kText = "F9"
   Case 121:  kText = "F10"
   Case 122:  kText = "F11"
   Case 123:  kText = "F12"
   Case Else: kText = Chr$(KeyCode)
  End Select
 End If
End Sub

Public Sub AddItem(Optional ByVal IconNum As Long = 0, Optional ByVal Text As String = "", Optional ByVal FontName As String = "Tahoma", Optional ByVal FontSize As Long = 8, Optional ByVal FontColor As OLE_COLOR = vbBlack, Optional ByVal FontBold As Boolean = False, Optional ByVal FontItalic As Boolean = False, Optional ByVal FontStrikethru As Boolean = False, Optional ByVal FontUnderline As Boolean = False, Optional ByVal bPopUpMenu As Boolean = False, Optional ByVal bSeperator As Boolean = False, Optional ByVal Check As Boolean = False, Optional ByVal Enabled As Boolean = True, Optional ByVal Visible As Boolean = True, Optional ByVal ShadowText As Boolean = False, Optional ByVal KeyMask As ShiftConstants, Optional ByVal KeyCode As Long, Optional ByVal MnuSubItem As clsXPMenu = Nothing, Optional ByVal ProcAddr As Long = 0&)
 MenuItemCount = MenuItemCount + 1
 ReDim Preserve MenuItems(1 To MenuItemCount) As TypMenuItem
 Call SetMyShortCut(KeyMask, KeyCode)
 With MenuItems(MenuItemCount)
  .IconNum = IconNum
  .bSeperator = bSeperator
  If (.bSeperator = False) Then .Text = Text
  .bPopUpMenu = bPopUpMenu
  If (cText <> "") And (kText <> "") And (.bPopUpMenu = False) And (.bSeperator = False) Then
   .ShortCut = Space$(8) & cText & "+" & kText
  Else
   .ShortCut = ""
  End If
  If (.ShortCut <> "") Then Call RegisterShortCut(.ShortCut)
  .Check = Check
  .Enabled = Enabled
  .ShadowText = ShadowText
  .Visible = Visible
  .ProcAddr = ProcAddr
  If (FontColor <= 0) Then FontColor = myForeColor
 On Error GoTo myErr
  .FontName = FontName
  .FontSize = FontSize
  If (FontSize > 15) Then FontSize = 8
  .FontColor = FontColor
  .FontBold = FontBold
  .FontItalic = FontItalic
  .FontStrikethru = FontStrikethru
  .FontUnderline = FontUnderline
  If Not (MnuSubItem Is Nothing) Then Set .MnuSubMenu = MnuSubItem
  GoTo NoErr
myErr:
  .FontName = "Tahoma"
  .FontSize = "8"
  .FontColor = vbBlack
  .FontBold = False
  .FontItalic = False
  .FontStrikethru = False
  .FontUnderline = False
 End With
NoErr:
 If (bSeperator = True) Then SepItemCnt = SepItemCnt + 1 Else TextItemCnt = TextItemCnt + 1
 Call GetWidth(MenuItemCount, bPopUpMenu)
End Sub

Public Function GetHilightNum() As Integer
 GetHilightNum = HiLightedItem
End Function

Public Function GetItemText(ByVal ItemNum As Integer) As String
 If (ItemNum > MenuItemCount) Then
  GetItemText = ""
  Exit Function
 End If
 GetItemText = MenuItems(ItemNum).Text
End Function

Public Function GetMenuName() As Variant
 GetMenuName = mnuName
End Function

Public Function IsTextItem(ByVal ItemNum As Integer) As Boolean
 If (ItemNum > MenuItemCount) Then
  IsTextItem = False
  Exit Function
 End If
 If (ItemNum = 0) Then Exit Function
 If (MenuItems(ItemNum).bPopUpMenu) Or (MenuItems(ItemNum).bSeperator) Or (MenuItems(ItemNum).Enabled = False) Then
  IsTextItem = False
 Else
  IsTextItem = True
 End If
End Function

Public Function IsVisible() As Boolean
 IsVisible = bVisible
End Function

Public Sub KillAllMenus()
 Dim Frm As Form
 
 For Each Frm In Forms
  If (Frm.Tag = "XPMenu") Then
   Frm.XPMenuClass.KillPopupMenus
   Frm.XPMenuClass.UnLoadMenu
  End If
 Next
End Sub

Public Sub KillPopupMenus()
 Dim I As Long
    
 For I = 1 To MenuItemCount
  If (MenuItems(I).bPopUpMenu) Then
   MenuItems(I).MnuSubMenu.KillPopupMenus
   MenuItems(I).MnuSubMenu.UnLoadMenu
  End If
 Next
End Sub

Public Function PopUpShown() As Boolean
 PopUpShown = bPopupShown
End Function

Public Sub ShowMenu(ByVal X As Long, ByVal Y As Long)
 If (bVisible = True) Then Exit Sub
 frmMenu.Left = X * 15
 frmMenu.Top = Y * 15
 Call DrawMenu
 bVisible = True
 bPopupShown = False
 frmMenu.tmrActive.Enabled = True
 frmMenu.tmrHover.Enabled = True
 frmMenu.Tag = "XPMenu"
End Sub

Private Sub DrawMenu()
 Dim oldFont       As String, oldSize       As Integer, MSG     As String
 Dim oldItalic     As Boolean, oldColor     As Long, cMsg       As String
 Dim oldStrikethru As Boolean, oldUnderLine As Boolean, oldBold As Boolean
 Dim Xcur          As Long, wEna            As Boolean, Index   As Integer
 Dim Ycur          As Long
 
 With frmMenu.picMenuBuffer
  .Cls
  .Height = GetHeight()
  .Width = MnuWidth
  If (myMenu = 1) Then
   Call Tile_Main(frmMenu.picMenuBuffer, BackPicture, 0, 0, frmMenu.picMenuBuffer.Width, frmMenu.picMenuBuffer.Height)
  ElseIf (myMenu = 2) Then
   Call Gradiente(frmMenu.picMenuBuffer, Color1, Color2, TypeGradient)
  Else
   .BackColor = IIf(myBackColor <> 0, myBackColor, Clr_BackGround) '* BackGround Color.
  End If
   '* Set Border.
   If (myBorderColor <> 0) Then
    frmMenu.picMenuBuffer.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), myBorderColor, B
   Else
    frmMenu.picMenuBuffer.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), Clr_MenuBorder, B
   End If
  If (myMargin = False) Then
   '* Set Margin.
   If (myMarginColor <> 0) Then
    frmMenu.picMenuBuffer.Line (1, 1)-(Dim_MarginWidth + 1, .ScaleHeight - 2), myMarginColor, BF
   Else
    frmMenu.picMenuBuffer.Line (1, 1)-(Dim_MarginWidth + 1, .ScaleHeight - 2), Clr_LeftMargin, BF
   End If
  End If
  '* X, Y Info.
  Ycur = 3
  For Index = 1 To MenuItemCount
   wEna = False
   MSG = MenuItems(Index).ShortCut
   cMsg = MenuItems(Index).Text
   wEna = MenuItems(Index).Enabled
   frmMenu.picIcon.Cls
   frmMenu.picResult.Cls
   frmMenu.picIconD.Cls
   If (MenuItems(Index).Visible = True) Then
    Xcur = Dim_MarginWidth + (XBuffer * 2) + 1 '* 1 For the Border's.
    '* ¿If HilightedItem?.
    If (HiLightedItem = Index) And (MenuItems(Index).bSeperator = False) And (wEna = True) Then
     YHiLight = Ycur
     If (myHighColor <> 0) Then
      frmMenu.picMenuBuffer.Line (2, Ycur - 1)-(.ScaleWidth - 3, Ycur + (YBuffer * 2) + TheTextHeight), myHighColor, BF
     Else
      frmMenu.picMenuBuffer.Line (2, Ycur - 1)-(.ScaleWidth - 3, Ycur + (YBuffer * 2) + TheTextHeight), Clr_HiLightBack, BF
     End If
     If (myHighBorderColor <> 0) Then
      frmMenu.picMenuBuffer.Line (2, Ycur - 1)-(.ScaleWidth - 3, Ycur + (YBuffer * 2) + TheTextHeight), myHighBorderColor, B
     Else
      frmMenu.picMenuBuffer.Line (2, Ycur - 1)-(.ScaleWidth - 3, Ycur + (YBuffer * 2) + TheTextHeight), Clr_HiLightBorder, B
     End If
    Else
    End If
    '* Bit Icon.
    If (MenuItems(Index).Check = True) And (MenuItems(Index).bPopUpMenu = False) And (MenuItems(Index).bSeperator = False) Then
     frmMenu.picIcon.Picture = frmMenu.ImgCheck.Picture
     frmMenu.picIconD.Picture = frmMenu.picIcon.Picture
     If (wEna = False) Then Call picDisabled(frmMenu.picIconD)
     If (HiLightedItem = Index) And (MenuItems(Index).bSeperator = False) And (wEna = True) Then
      Call picShadow(frmMenu.picIcon, frmMenu.picIcon.Picture)
      Call IconBackground((Dim_MarginWidth - 16) \ 2 + 1, Ycur + (((TheTextHeight + (YBuffer * 2)) - 16) \ 2), , True)
      frmMenu.picIcon.Picture = frmMenu.ImgCheck.Picture
      Call IconBackground((Dim_MarginWidth - 16) \ 2, Ycur + (((TheTextHeight + (YBuffer * 2)) - 16) \ 2) - 1)
     Else
      Call IconBackground((Dim_MarginWidth - 16) \ 2 + 1, Ycur + (((TheTextHeight + (YBuffer * 2)) - 16) \ 2), Not (wEna))
     End If
    End If
    If Not (ImageLst Is Nothing) Then
     If (MenuItems(Index).IconNum <> 0) And (MenuItems(Index).Check = False) And (MenuItems(Index).Visible = True) And (MenuItems(Index).bSeperator = False) Then
      frmMenu.picIcon.Picture = ImageLst.ListImages(MenuItems(Index).IconNum).Picture
      frmMenu.picIconD.Picture = frmMenu.picIcon.Picture
      If (wEna = False) Then picDisabled frmMenu.picIconD
      If (HiLightedItem = Index) And (MenuItems(Index).bSeperator = False) And (wEna = True) Then
       Call picShadow(frmMenu.picIcon, frmMenu.picIcon.Picture)
       Call IconBackground((Dim_MarginWidth - 16) \ 2 + 1, Ycur + (((TheTextHeight + (YBuffer * 2)) - 16) \ 2), , True)
       frmMenu.picIcon.Picture = ImageLst.ListImages(MenuItems(Index).IconNum).Picture
       Call IconBackground((Dim_MarginWidth - 16) \ 2, Ycur + (((TheTextHeight + (YBuffer * 2)) - 16) \ 2) - 1)
      Else
       Call IconBackground((Dim_MarginWidth - 16) \ 2 + 1, Ycur + (((TheTextHeight + (YBuffer * 2)) - 16) \ 2), Not (wEna))
      End If
     End If
    End If
    '* PopUp Menu Show.
    oldFont = MenuItems(Index).FontName
    oldSize = MenuItems(Index).FontSize
    oldBold = MenuItems(Index).FontBold
    oldUnderLine = MenuItems(Index).FontUnderline
    oldItalic = MenuItems(Index).FontItalic
    oldColor = MenuItems(Index).FontColor
    oldStrikethru = MenuItems(Index).FontStrikethru
    If (MenuItems(Index).bPopUpMenu) Then
     frmMenu.picMenuBuffer.FontName = "Marlett"
     frmMenu.picMenuBuffer.FontSize = 10
     frmMenu.picMenuBuffer.FontItalic = False
     frmMenu.picMenuBuffer.FontStrikethru = False
     frmMenu.picMenuBuffer.FontUnderline = False
     frmMenu.picMenuBuffer.FontBold = False
     '* Shadow Text.
     If (MenuItems(Index).ShadowText = True) And (MenuItems(Index).Enabled = True) And (HiLightedItem = Index) Then
      Call SetTextColor(.hDC, vbWhite)
      Call DrawStateString(.hDC, 0, 0, "4", Len("4"), .ScaleWidth - .TextHeight("4") - XBuffer, Ycur + (((TheTextHeight + (YBuffer * 2)) - 16) \ 2) + 2, 0, 0, DST_TEXT Or IIf(wEna, DSS_NORMAL, DSS_DISABLED))
     End If
     If (wEna = True) Then
      If (HiLightedItem = Index) Then
       Call SetTextColor(.hDC, IIf(m_clrArrowSelect <> 0, m_clrArrowSelect, FontFlecha))
      Else
       Call SetTextColor(.hDC, IIf(m_clrArrowNormal <> 0, m_clrArrowNormal, FontHighLightF))
      End If
     End If
     Call DrawStateString(.hDC, 0, 0, "4", Len("4"), .ScaleWidth - .TextHeight("4") - XBuffer, Ycur + (((TheTextHeight + (YBuffer * 2)) - 16) \ 2) + 2 - 1, 0, 0, DST_TEXT Or IIf(wEna, DSS_NORMAL, DSS_DISABLED))
    End If
    frmMenu.picMenuBuffer.FontName = oldFont
    frmMenu.picMenuBuffer.FontSize = oldSize
    frmMenu.picMenuBuffer.FontBold = oldBold
    frmMenu.picMenuBuffer.FontItalic = oldItalic
    frmMenu.picMenuBuffer.FontStrikethru = oldStrikethru
    frmMenu.picMenuBuffer.FontUnderline = oldUnderLine
    '* Draw Item Select.
    If (MenuItems(Index).bSeperator) Then
     If (SeparatorColor <> 0) Then
      frmMenu.picMenuBuffer.Line (Dim_MarginWidth + 7, Ycur + YBuffer)-(.ScaleWidth - 1, Ycur + YBuffer), SeparatorColor
      frmMenu.picMenuBuffer.Line (Dim_MarginWidth + 8, Ycur + YBuffer + 1)-(.ScaleWidth - 1, Ycur + YBuffer + 1), MenuBar
     Else
      frmMenu.picMenuBuffer.Line (Dim_MarginWidth + 7, Ycur + YBuffer)-(.ScaleWidth - 1, Ycur + YBuffer), Clr_Seperator
      frmMenu.picMenuBuffer.Line (Dim_MarginWidth + 8, Ycur + YBuffer + 1)-(.ScaleWidth - 1, Ycur + YBuffer + 1), MenuBar
     End If
     Ycur = Ycur + 1 + (XBuffer * 2)
    Else
     '* Shadow Text.
     If (MenuItems(Index).ShadowText = True) And (MenuItems(Index).Enabled = True) Then
      If (myShadowColor <> 0) Then
       Call SetTextColor(.hDC, myShadowColor)
      Else
       Call SetTextColor(.hDC, &HC0C0C0)
      End If
      Call DrawStateString(.hDC, 0, 0, cMsg, Len(cMsg), Xcur + 1, Ycur + YBuffer + 1, 0, 0, DST_TEXT Or IIf(wEna, DSS_NORMAL, DSS_DISABLED))
      Call DrawStateString(.hDC, 0, 0, MSG, Len(MSG), .ScaleWidth - (.TextHeight(cMsg) + .TextWidth(MSG) + Xcur - 33) + 1, Ycur + (((TheTextHeight + (YBuffer * 2)) - 16) \ 2) + 3, 0, 0, DST_TEXT Or IIf(wEna, DSS_NORMAL, DSS_DISABLED))
     End If
     If (wEna = True) Then
      If (HiLightedItem = Index) Then
       Call SetTextColor(.hDC, IIf(mySelForeColor <> 0, mySelForeColor, FontHighLight))
      Else
       Call SetTextColor(.hDC, oldColor)
      End If
     End If
     Call DrawStateString(.hDC, 0, 0, cMsg, Len(cMsg), Xcur, Ycur + YBuffer, 0, 0, DST_TEXT Or IIf(wEna, DSS_NORMAL, DSS_DISABLED))
     Call DrawStateString(.hDC, 0, 0, MSG, Len(MSG), .ScaleWidth - (.TextHeight(cMsg) + .TextWidth(MSG) + Xcur - 33), Ycur + (((TheTextHeight + (YBuffer * 2)) - 16) \ 2) + 2, 0, 0, DST_TEXT Or IIf(wEna, DSS_NORMAL, DSS_DISABLED))
     Ycur = Ycur + TheTextHeight + (YBuffer * 2)
    End If
   End If
  Next
 End With
 frmMenu.Width = frmMenu.picMenuBuffer.Width * 15
 frmMenu.Height = frmMenu.picMenuBuffer.Height * 15
 frmMenu.Picture = frmMenu.picMenuBuffer.Image
 frmMenu.Show
End Sub

Private Sub GetWidth(ByVal Index As Long, Optional bPopUpMenu As Boolean = False)
 Dim TheWidth As Integer
 
 With frmMenu
  .lblWidth.FontSize = MenuItems(Index).FontSize
  .lblWidth.FontBold = MenuItems(Index).FontBold
  .lblWidth.FontName = MenuItems(Index).FontName
  .lblWidth.FontItalic = MenuItems(Index).FontItalic
  .lblWidth.FontStrikethru = MenuItems(Index).FontStrikethru
  .lblWidth.FontUnderline = MenuItems(Index).FontUnderline
  .lblWidth.Caption = MenuItems(MenuItemCount).Text & MenuItems(MenuItemCount).ShortCut
  TheWidth = .lblWidth.Width + XBuffer + Dim_MarginWidth + 20
  If (bPopUpMenu = True) Then TheWidth = TheWidth + (XBuffer * 2) + frmMenu.picPopUp.TextWidth("4")
  If (TheWidth > MnuWidth) Then MnuWidth = TheWidth
 End With
End Sub

Private Function GetHeight() As Long
 Dim LngHeight As Long, Index As Integer
 Dim Ycur      As Long
 
 With frmMenu.picMenuBuffer
  Ycur = 3
  For Index = 1 To MenuItemCount
   '* Draw Item.
   If (MenuItems(Index).Visible = True) Then
    If (MenuItems(Index).bSeperator) Then
     Ycur = Ycur + 1 + (XBuffer * 2)
    Else
     Ycur = Ycur + TheTextHeight + (YBuffer * 2)
    End If
   End If
  Next
 End With
 LngHeight = Ycur + 4
 GetHeight = LngHeight
End Function

Public Function GetHiLightedItem(ByVal Y As Single) As Integer
 Dim Ycur As Long, Index As Integer
 
On Error GoTo EndD
 With frmMenu.picMenuBuffer
  '* X, Y info.
  Ycur = 3
  For Index = 1 To MenuItemCount
   If (MenuItems(Index).Visible = True) Then
    If (MenuItems(Index).bSeperator) Then
     If (Y >= Ycur) And (Y <= Ycur + (YBuffer * 2) + 1) Then
      GetHiLightedItem = Index
      Exit Function
     End If
     Ycur = Ycur + 1 + (XBuffer * 2)
    Else
     If (Y >= Ycur) And (Y <= Ycur + TheTextHeight + (YBuffer * 2)) Then
      GetHiLightedItem = Index
      Exit Function
     End If
     Ycur = Ycur + TheTextHeight + (YBuffer * 2)
    End If
   End If
  Next
 End With
 Exit Function
EndD:
End Function

Public Sub Init(ByVal StrMenuName As String)
 mnuName = StrMenuName
 Set frmMenu.XPMenuClass = Me
 frmMenu.FontName = Font_MenuItem
 frmMenu.picMenuBuffer.FontName = Font_MenuItem
 TheTextHeight = frmMenu.picMenuBuffer.TextHeight("gW")
 MenuItemCount = 0
 SepItemCnt = 0
 TextItemCnt = 0
 HiLightedItem = 0
End Sub

Public Sub SetHilightedItem(ByVal Item As Integer)
 If (Item = 0) Or (HiLightedItem = Item) Then Exit Sub
 If (Item = -1) Then
  HiLightedItem = -1
  Call DrawMenu
  Exit Sub
 End If
 HiLightedItem = Item
 Call KillPopupMenus
 bPopupShown = False
 Call DrawMenu
 If (MenuItems(Item).bPopUpMenu) And (MenuItems(Item).Enabled = True) Then
  bPopupShown = True
  Set ActivePopup = MenuItems(Item).MnuSubMenu
  If (ActivePopup.IsVisible) Then Exit Sub
  Call ActivePopup.ShowMenu(frmMenu.Left \ 15 + frmMenu.Width \ 15, frmMenu.Top \ 15 + YHiLight)
 End If
End Sub

Public Sub UnLoadMenu()
 Call Unload(frmMenu)
 bVisible = False
 HiLightedItem = 0
 frmMenu.tmrActive.Enabled = False
 frmMenu.tmrHover.Enabled = False
 Call ClearStatus
End Sub

Private Sub Gradiente(ByRef Objeto As Object, ByVal LngColor1 As Long, ByVal LngColor2 As Long, Optional ByVal Tipo As Long = 0)
 Dim RgbColor1 As RGBQUAD, RgbColor2 As RGBQUAD
 Dim ColorRojo As Double, ColorVerde As Double, ColorAzul As Double
 Dim CDiffRed  As Double, CDiffGreen As Double, CDiffBlue As Double
 Dim CFadeRed  As Double, CFadeGreen As Double, CFadeBlue As Double
 Dim Fade      As Double, Tamaño     As Double, Longitud  As Double
 Dim hNewPen   As Long, hOldPen      As Long, Escala      As Long
  
 On Error Resume Next
  Escala = Objeto.ScaleMode
  Objeto.ScaleMode = vbPixels
  Select Case Tipo
   Case Horizontal
    Tamaño = Objeto.ScaleWidth
    Longitud = Objeto.ScaleHeight
   Case Vertical
    Tamaño = Objeto.ScaleHeight
    Longitud = Objeto.ScaleWidth
   Case Diagonal
    Tamaño = Objeto.ScaleWidth
    Longitud = Objeto.ScaleHeight
  End Select
  Call Long2RGB(LngColor1, RgbColor1)
  Call Long2RGB(LngColor2, RgbColor2)
  CDiffRed = -(CLng(RgbColor1.rgbRed) - CLng(RgbColor2.rgbRed))
  CDiffGreen = -(CLng(RgbColor1.rgbGreen) - CLng(RgbColor2.rgbGreen))
  CDiffBlue = -(CLng(RgbColor1.rgbBlue) - CLng(RgbColor2.rgbBlue))
  ColorRojo = RgbColor1.rgbRed: ColorVerde = RgbColor1.rgbGreen: ColorAzul = RgbColor1.rgbBlue
  CFadeRed = CDiffRed / Tamaño
  CFadeGreen = CDiffGreen / Tamaño
  CFadeBlue = CDiffBlue / Tamaño
  If (Tipo = Diagonal) Then Tamaño = Tamaño + Longitud
  For Fade = 0 To Tamaño
   hNewPen = CreatePen(PS_SOLID, 0, RGB(Abs(ColorRojo), Abs(ColorVerde), Abs(ColorAzul)))
   hOldPen = SelectObject(Objeto.hDC, hNewPen)
   Select Case Tipo
    Case Horizontal
     Call MoveToEx(Objeto.hDC, Fade, 0&, 0&)
     Call LineTo(Objeto.hDC, Fade, Longitud)
    Case Vertical
     Call MoveToEx(Objeto.hDC, 0&, Fade, 0&)
     Call LineTo(Objeto.hDC, Longitud, Fade)
    Case Diagonal
     Call MoveToEx(Objeto.hDC, Fade, 0&, 0&)
     Call LineTo(Objeto.hDC, 0&, Fade)
   End Select
   hNewPen = SelectObject(Objeto.hDC, hOldPen)
   Call DeleteObject(hNewPen)
   ColorRojo = ColorRojo + CFadeRed
   ColorVerde = ColorVerde + CFadeGreen
   ColorAzul = ColorAzul + CFadeBlue
  Next
 Objeto.Picture = Objeto.Image
 Objeto.ScaleMode = Escala
End Sub

Private Sub picDisabled(ByRef picTo As Object)
 Dim sTMPpathFName As String
 
 Call DrawState(picTo.hDC, 0, 0, picTo.Picture, 0, 0, 0, 16, 16, DST_BITMAP Or DSS_DISABLED)
 sTMPpathFName = App.Path + "\~ConvIcon2Bmp.tmp"
 Call SavePicture(picTo.Image, sTMPpathFName)
 Set picTo.Picture = LoadPicture(sTMPpathFName)
 Call Kill(sTMPpathFName)
 picTo.Refresh
End Sub

Private Sub picShadow(ByRef picTo As Object, ByVal isPicture As StdPicture)
 Dim lValue As Long, sTMPpathFName As String
 Dim isType As Long
 
 Select Case isPicture.Type
  Case 1, 4: isType = DST_BITMAP
  Case 3:    isType = DST_ICON
 End Select
 lValue = CreateSolidBrush(RGB(136, 141, 157))
 Call DrawState(picTo.hDC, lValue, 0, isPicture.Handle, 0, 0, 0, 16, 16, isType Or DSS_MONO)
 'sTMPpathFName = App.Path + "\~ConvIcon2Bmp.tmp"
 'Call SavePicture(picTo.Image, sTMPpathFName)
 'Set picTo.Picture = LoadPicture(sTMPpathFName)
 'Call Kill(sTMPpathFName)
 picTo.Refresh
End Sub

Public Sub MenuItem(ByVal MenuIndex As Long, ByVal Text As String, Optional ByVal KeyMask As ShiftConstants, Optional ByVal KeyCode As Long, Optional ByVal bPopUpMenu As Boolean = False, Optional ByVal bSeperator As Boolean = False, Optional ByVal FontName As String = "Tahoma", Optional ByVal FontSize As Long = 8, Optional ByVal FontColor As OLE_COLOR = vbBlack, Optional ByVal FontBold As Boolean = False, Optional ByVal FontItalic As Boolean = False, Optional ByVal FontStrikethru As Boolean = False, Optional ByVal FontUnderline As Boolean = False)
 '* Cambiar el Texto de un Menú.
On Error GoTo myErr
 MenuItems(MenuIndex).Text = Text
 Call SetMyShortCut(KeyMask, KeyCode)
 If (cText <> "") And (kText <> "") And (bPopUpMenu = False) And (bSeperator = False) Then
  MenuItems(MenuIndex).ShortCut = Space$(3) & cText & "+" & kText
 Else
  MenuItems(MenuIndex).ShortCut = ""
 End If
 If (FontSize > 15) Then FontSize = 8
 MenuItems(MenuIndex).FontBold = FontBold
 MenuItems(MenuIndex).FontItalic = FontItalic
 MenuItems(MenuIndex).FontName = FontName
 MenuItems(MenuIndex).FontUnderline = FontUnderline
 MenuItems(MenuIndex).FontSize = FontSize
 If (FontColor <= 0) Then FontColor = myForeColor
 MenuItems(MenuIndex).FontColor = FontColor
 MenuItems(MenuIndex).FontStrikethru = FontStrikethru
 Call GetWidth(MenuIndex, bPopUpMenu)
 Exit Sub
myErr: iniFont MenuIndex
End Sub

Public Sub MenuEnabled(ByVal MenuIndex As Long, ByVal Enabled As Boolean)
 '* Habilitar ó Deshabilitar un Menú.
On Error Resume Next
 MenuItems(MenuIndex).Enabled = Enabled
End Sub

Public Sub MenuVisible(ByVal MenuIndex As Long, ByVal Visible As Boolean)
 '* Visible un Menú.
On Error Resume Next
 MenuItems(MenuIndex).Visible = Visible
End Sub

Public Sub MenuCheck(ByVal MenuIndex As Long, ByVal Check As Boolean)
 '* Check un Menú.
On Error Resume Next
 MenuItems(MenuIndex).Check = Check
End Sub

Public Sub MenuPicture(ByVal MenuIndex As Long, ByVal IconNum As Long)
 '* Picture en el Menú.
On Error Resume Next
 MenuItems(MenuIndex).IconNum = IconNum
End Sub

Public Sub MenuShadow(ByVal MenuIndex As Long, ByVal State As Boolean)
 '* Colocar el Texto con Sombra.
On Error Resume Next
 MenuItems(MenuIndex).ShadowText = State
End Sub

Private Sub RegisterShortCut(ByVal ShortCut As String)
 '* Registrar el ShortCut para el Item del Menú que lo tenga.
End Sub

Public Sub MaskColor(Optional ByVal Mask As Long = &HFF00FF)
 Maskara = Mask
End Sub

Private Sub SetMaskColor(ByVal cTrue As Boolean)
 Dim BrsH As Long
 
 If (cTrue = True) Then BrsH = CreateSolidBrush(Maskara) Else Call DeleteObject(BrsH)
End Sub

Public Property Get MenuBackColor() As OLE_COLOR
 MenuBackColor = myBackColor
End Property

Public Property Let MenuBackColor(ByVal myColor As OLE_COLOR)
 myBackColor = myColor
End Property

Public Property Get MenuHighLightColor() As OLE_COLOR
 MenuHighLightColor = myHighColor
End Property

Public Property Let MenuHighLightColor(ByVal myColor As OLE_COLOR)
 myHighColor = myColor
End Property

Public Property Get MenuHighLightBorderColor() As OLE_COLOR
 MenuHighLightBorderColor = myHighBorderColor
End Property

Public Property Let MenuHighLightBorderColor(ByVal myColor As OLE_COLOR)
 myHighBorderColor = myColor
End Property

Public Property Get MenuMarginColor() As OLE_COLOR
 MenuMarginColor = myMarginColor
End Property

Public Property Let MenuMarginColor(ByVal myColor As OLE_COLOR)
 myMarginColor = myColor
End Property

Public Property Get MenuShadowColor() As OLE_COLOR
 MenuShadowColor = myShadowColor
End Property

Public Property Let MenuShadowColor(ByVal myColor As OLE_COLOR)
 myShadowColor = myColor
End Property

Public Property Get MenuBorderColor() As OLE_COLOR
 MenuBorderColor = myBorderColor
End Property

Public Property Let MenuBorderColor(ByVal myColor As OLE_COLOR)
 myBorderColor = myColor
End Property

Public Property Get SeparatorColor() As OLE_COLOR
 SeparatorColor = mySeparatorColor
End Property

Public Property Let SeparatorColor(ByVal myColor As OLE_COLOR)
 mySeparatorColor = myColor
End Property

Public Property Get MenuForeColor() As OLE_COLOR
 MenuForeColor = myForeColor
End Property

Public Property Let MenuForeColor(ByVal myColor As OLE_COLOR)
 myForeColor = myColor
End Property

Public Property Get MenuSelForeColor() As OLE_COLOR
 MenuSelForeColor = mySelForeColor
End Property

Public Property Let MenuSelForeColor(ByVal myColor As OLE_COLOR)
 mySelForeColor = myColor
End Property

Public Property Get MenuArrowNormalColor() As OLE_COLOR
 MenuArrowNormalColor = m_clrArrowNormal
End Property

Public Property Let MenuArrowNormalColor(ByVal myColor As OLE_COLOR)
 m_clrArrowNormal = myColor
End Property

Public Property Get MenuArrowSelectColor() As OLE_COLOR
 MenuArrowSelectColor = m_clrArrowSelect
End Property

Public Property Let MenuArrowSelectColor(ByVal myColor As OLE_COLOR)
 m_clrArrowSelect = myColor
End Property

Private Sub iniFont(ByVal iPos)
On Error GoTo myErr
 With MenuItems(iPos)
  MenuItems(iPos).FontName = "Tahoma"
  MenuItems(iPos).FontSize = "8"
  MenuItems(iPos).FontColor = myForeColor
  MenuItems(iPos).FontBold = False
  MenuItems(iPos).FontItalic = False
  MenuItems(iPos).FontStrikethru = False
  MenuItems(iPos).FontUnderline = False
 End With
 Exit Sub
myErr:
End Sub

Public Sub BackGroundPicture(ByVal ImagePicture As StdPicture)
 
End Sub

Private Sub Tile_Main(ByRef TileObject As Object, ByRef TilePicture As Picture, ByVal TileLeft As Single, ByVal TileTop As Single, ByVal TileWidth As Single, ByVal TileHeight As Single)
 Dim ImageTop          As Single, C          As Long
 Dim ImageHeight       As Single, ImageWidth As Single
 Dim Max_Images_Width  As Integer, I         As Long
 Dim Max_Images_Height As Integer, ImageLeft As Single
 
On Error GoTo Cancel
 ImageTop = TileTop
 ImageLeft = TileLeft
 ImageWidth = TileObject.ScaleX(TilePicture.Width, vbHimetric, TileObject.ScaleMode)
 ImageHeight = TileObject.ScaleY(TilePicture.Height, vbHimetric, TileObject.ScaleMode)
 Max_Images_Width = CInt(TileWidth \ ImageWidth) + 1
 Max_Images_Height = IIf(CInt(TileHeight \ ImageHeight) = 0, 1, CInt(TileHeight \ ImageHeight) + 1)
 TileObject.AutoRedraw = True
 If (Max_Images_Width > 1) Or (Max_Images_Height > 1) Then
  For I = 1 To Max_Images_Height + 1
   For C = 0 To Max_Images_Width
    TileObject.PaintPicture TilePicture, ImageLeft, ImageTop, ImageWidth, ImageHeight
    ImageLeft = ImageLeft + ImageWidth
   Next
   ImageLeft = TileLeft
   ImageWidth = TileObject.ScaleX(TilePicture.Width, vbHimetric, TileObject.ScaleMode)
   ImageTop = ImageTop + ImageHeight
  Next
 Else
  TileObject.Picture = TilePicture
 End If
 Exit Sub
Cancel:
End Sub

Public Sub ListImages(ByRef ListImage As ImageList)
 If Not (ListImage Is Nothing) Then Set ImageLst = ListImage
End Sub

Public Sub StyleMenu(Optional ByVal Style As OptionMenu, Optional ImagePicture As StdPicture = Nothing, Optional ByVal xColor As OLE_COLOR = &HEBE2E0, Optional ByVal yColor As OLE_COLOR = &HDEEDEF, Optional ByVal tGradient As TipoGradiente = 0)
 If (Style = Gradient) Then
  Color1 = xColor
  Color2 = yColor
  TypeGradient = tGradient
 ElseIf (Style = BackGround) Then
  Set BackPicture = Nothing
 On Error GoTo myErr
  Set BackPicture = ImagePicture
  GoTo NoErr:
myErr: Exit Sub
 End If
NoErr:
 myMenu = Style
End Sub

Private Sub IconBackground(ByVal X1 As Long, ByVal Y1 As Long, Optional ByVal Disabled As Boolean = False, Optional ByVal Shadow As Boolean = False)
 Dim sTMPpathFName As String
 
 If (Disabled = False) Then
  frmMenu.picResult.Picture = frmMenu.picIcon.Picture
 Else
  frmMenu.picResult.Picture = frmMenu.picIconD.Picture
 End If
 '*** ---- ***'
 sTMPpathFName = App.Path + "\~ConvIcon2Bmp.tmp"
 Call SavePicture(frmMenu.picResult.Image, sTMPpathFName)
 Set frmMenu.picResult.Picture = LoadPicture(sTMPpathFName)
 Call Kill(sTMPpathFName)
 frmMenu.picResult.Refresh
 frmMenu.picIcon.Refresh
 With frmMenu.picMenuBuffer
  Call StretchBlt(.hDC, X1, Y1, 16, 16, frmMenu.picResult.hDC, 0, 0, frmMenu.picResult.ScaleWidth, frmMenu.picResult.ScaleHeight, vbSrcAnd)
  If (Disabled = False) Then Call StretchBlt(.hDC, X1, Y1, 16, 16, frmMenu.picIcon.hDC, 0, 0, frmMenu.picIcon.ScaleWidth, frmMenu.picIcon.ScaleHeight, vbSrcPaint)
  .Refresh
 End With
End Sub

Private Sub Espera(ByVal Segundos As Single)
 Dim ComienzoSeg As Single, FinSeg As Single
 
 ComienzoSeg = Timer
 FinSeg = ComienzoSeg + Segundos
 Do While FinSeg > Timer
  DoEvents
  If ComienzoSeg > Timer Then
   FinSeg = FinSeg - 24 * 60 * 60
  End If
 Loop
End Sub

Public Property Get SetMargin() As Boolean
 SetMargin = myMargin
End Property

Public Property Let SetMargin(ByVal LetMargin As Boolean)
 myMargin = LetMargin
End Property
