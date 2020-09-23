VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTestForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test XP Menu - HACKPRO TM"
   ClientHeight    =   4965
   ClientLeft      =   5700
   ClientTop       =   2715
   ClientWidth     =   6615
   Icon            =   "frmTestForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdDialog 
      Left            =   90
      Top             =   1050
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cmbGradient 
      Height          =   315
      ItemData        =   "frmTestForm.frx":0442
      Left            =   2520
      List            =   "frmTestForm.frx":044F
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2025
      Width           =   1230
   End
   Begin VB.PictureBox picGradient 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1950
      MouseIcon       =   "frmTestForm.frx":0473
      MousePointer    =   99  'Custom
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1620
      Width           =   285
   End
   Begin VB.PictureBox picGradient 
      BackColor       =   &H00D2C8BE&
      Height          =   285
      Index           =   0
      Left            =   885
      MouseIcon       =   "frmTestForm.frx":077D
      MousePointer    =   99  'Custom
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1620
      Width           =   285
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      Left            =   1410
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1140
      Width           =   1155
   End
   Begin VB.ComboBox cmbShortCut 
      Height          =   315
      Index           =   1
      ItemData        =   "frmTestForm.frx":0A87
      Left            =   1410
      List            =   "frmTestForm.frx":0AA0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   1155
   End
   Begin VB.ComboBox cmbShortCut 
      Height          =   315
      Index           =   0
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   780
      Width           =   1155
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5355
      MouseIcon       =   "frmTestForm.frx":0AE7
      MousePointer    =   99  'Custom
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   840
      Width           =   1140
   End
   Begin VB.TextBox txtItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H008A4500&
      Height          =   285
      Index           =   2
      Left            =   1410
      TabIndex        =   0
      Text            =   "8"
      Top             =   90
      Width           =   1120
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FontUnderLine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008A4500&
      Height          =   180
      Index           =   3
      Left            =   2985
      TabIndex        =   7
      Top             =   855
      Width           =   1350
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FontStrikethru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H008A4500&
      Height          =   180
      Index           =   2
      Left            =   2985
      TabIndex        =   6
      Top             =   615
      Width           =   1365
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FontItalic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008A4500&
      Height          =   180
      Index           =   1
      Left            =   2985
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FontBold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008A4500&
      Height          =   180
      Index           =   0
      Left            =   2985
      TabIndex        =   4
      Top             =   105
      Width           =   1140
   End
   Begin VB.TextBox txtItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H008A4500&
      Height          =   285
      Index           =   1
      Left            =   5355
      TabIndex        =   9
      Text            =   "Cool Menu"
      Top             =   465
      Width           =   1140
   End
   Begin VB.TextBox txtItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H008A4500&
      Height          =   285
      Index           =   0
      Left            =   5355
      TabIndex        =   8
      Text            =   "3"
      Top             =   90
      Width           =   1140
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   345
      Left            =   2760
      MouseIcon       =   "frmTestForm.frx":0DF1
      TabIndex        =   14
      Top             =   1545
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Set Margin"
      ForeColor       =   &H008A4500&
      Height          =   255
      Left            =   210
      TabIndex        =   18
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1110
   End
   Begin VB.ListBox lstProperty 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H008A4500&
      Height          =   2985
      ItemData        =   "frmTestForm.frx":10FB
      Left            =   3810
      List            =   "frmTestForm.frx":113E
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   1545
      Width           =   2715
   End
   Begin VB.CommandButton cmdMenuPicture 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menu &Picture"
      Height          =   510
      Left            =   105
      MouseIcon       =   "frmTestForm.frx":1236
      TabIndex        =   17
      Top             =   3375
      Width           =   1305
   End
   Begin VB.CommandButton cmdMenuGradient 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menu &Gradient"
      Height          =   510
      Left            =   105
      MouseIcon       =   "frmTestForm.frx":1540
      TabIndex        =   16
      Top             =   2745
      Width           =   1305
   End
   Begin VB.CommandButton cmdViewMenu 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&See Menu"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   105
      MouseIcon       =   "frmTestForm.frx":184A
      TabIndex        =   15
      Top             =   2115
      Width           =   1305
   End
   Begin MSComctlLib.StatusBar sbr1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   20
      Top             =   4635
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4701
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1094
            MinWidth        =   1094
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1094
            MinWidth        =   1094
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1085
            MinWidth        =   1094
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "09:34 p.m."
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "02/02/06"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLstIcon 
      Left            =   1995
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":1B54
            Key             =   "Color"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":1CB0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":224C
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":25E8
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":2984
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":2D20
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":437C
            Key             =   "New"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":4918
            Key             =   "ClipBoard"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":4CB4
            Key             =   "Mayus"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":5050
            Key             =   "SelPal"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":53EC
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":5988
            Key             =   "Minus"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":5D24
            Key             =   "Erase"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":6A00
            Key             =   "SelAll"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":6B5C
            Key             =   "SelLine"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":6CB8
            Key             =   "PickPaste"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":7054
            Key             =   "Element"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":71B4
            Key             =   "DelAll"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":731C
            Key             =   "End"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":76B8
            Key             =   "Recent"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":7C54
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":7DB0
            Key             =   "Error"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":834C
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":88E8
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTestForm.frx":8C84
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Style:"
      ForeColor       =   &H008A4500&
      Height          =   195
      Index           =   12
      Left            =   2070
      TabIndex        =   33
      Top             =   2070
      Width           =   390
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Picture"
      ForeColor       =   &H008A4500&
      Height          =   195
      Index           =   11
      Left            =   2580
      TabIndex        =   32
      Top             =   4305
      Width           =   495
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Gradien2:"
      ForeColor       =   &H008A4500&
      Height          =   195
      Index           =   10
      Left            =   1215
      TabIndex        =   31
      Top             =   1650
      Width           =   690
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Gradient1:"
      ForeColor       =   &H008A4500&
      Height          =   195
      Index           =   9
      Left            =   105
      TabIndex        =   30
      Top             =   1650
      Width           =   735
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Font:"
      ForeColor       =   &H008A4500&
      Height          =   195
      Index           =   8
      Left            =   960
      TabIndex        =   29
      Top             =   1185
      Width           =   360
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ShortMask:"
      ForeColor       =   &H008A4500&
      Height          =   195
      Index           =   6
      Left            =   525
      TabIndex        =   28
      Top             =   465
      Width           =   810
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FontSize:"
      ForeColor       =   &H008A4500&
      Height          =   195
      Index           =   7
      Left            =   675
      TabIndex        =   27
      Top             =   120
      Width           =   660
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ShortCode:"
      ForeColor       =   &H008A4500&
      Height          =   195
      Index           =   5
      Left            =   525
      TabIndex        =   26
      Top             =   825
      Width           =   795
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   184
      X2              =   437
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      Index           =   0
      X1              =   184
      X2              =   437
      Y1              =   79
      Y2              =   79
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Color:"
      ForeColor       =   &H008A4500&
      Height          =   195
      Index           =   4
      Left            =   4875
      TabIndex        =   25
      Top             =   855
      Width           =   405
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Caption:"
      ForeColor       =   &H008A4500&
      Height          =   195
      Index           =   3
      Left            =   4695
      TabIndex        =   24
      Top             =   480
      Width           =   585
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Item:"
      ForeColor       =   &H008A4500&
      Height          =   195
      Index           =   2
      Left            =   4950
      TabIndex        =   23
      Top             =   105
      Width           =   345
   End
   Begin VB.Image imgPicture 
      Height          =   1800
      Left            =   1860
      MouseIcon       =   "frmTestForm.frx":8DE0
      MousePointer    =   99  'Custom
      Picture         =   "frmTestForm.frx":90EA
      Stretch         =   -1  'True
      Top             =   2490
      Width           =   1800
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Properties and Methods"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   0
      Left            =   3900
      TabIndex        =   21
      Top             =   1275
      Width           =   2490
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Properties and Methods"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   240
      Index           =   1
      Left            =   3885
      TabIndex        =   22
      Top             =   1260
      Width           =   2490
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuCaption 
         Caption         =   "&Caption"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuEnabled 
         Caption         =   "Enabled"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmTestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Private XPMenu     As New clsXPMenu
 Private XPM_EFNet  As New clsXPMenu
 Private XPMenu2    As New clsXPMenu
 Private XPM_DALNet As New clsXPMenu
 Private I          As Integer
 Private V          As String

Private Sub cmdApply_Click()
 XPMenu.MenuItem txtItem(0).Text, txtItem(1).Text, cmbShortCut(1).ListIndex + 1, cmbShortCut(0).ListIndex + 33, , , cmbFont.Text, txtItem(2).Text, picColor.BackColor, chk1(0).Value, chk1(1).Value, chk1(2).Value, chk1(3).Value
End Sub

Private Sub cmdViewMenu_Click()
 Dim Pos As POINTAPI
    
 If (Check1.Value = 0) Then XPMenu.SetMargin = True Else XPMenu.SetMargin = False
 Call GetCursorPos(Pos)
 Call XPMenu.ShowMenu(Pos.X, Pos.Y)
End Sub

Private Sub cmdMenuGradient_Click()
 If (cmdMenuGradient.Caption = "Menu &Gradient") Then
  Call XPMenu.StyleMenu(Gradient, , picGradient(0).BackColor, picGradient(1).BackColor, cmbGradient.ListIndex)
  cmdMenuGradient.Caption = "No &Gradient"
  cmdMenuPicture.Caption = "Menu &Picture"
 Else
  XPMenu.StyleMenu None
  cmdMenuGradient.Caption = "Menu &Gradient"
 End If
End Sub

Private Sub cmdMenuPicture_Click()
 If (cmdMenuPicture.Caption = "No &Picture") Then
  cmdMenuPicture.Caption = "Menu &Picture"
  XPMenu.StyleMenu
 Else
  cmdMenuPicture.Caption = "No &Picture"
  Call XPMenu.StyleMenu(BackGround, imgPicture.Picture)
  cmdMenuGradient.Caption = "Menu &Gradient"
 End If
End Sub

Private Sub Form_Load()
 Call XPMenu.ListImages(imgLstIcon)
 MenuXpItem.Items(0) = "&Edición"
 MenuXpItem.Status(0) = "Menú Edición"
 MenuXpItem.Items(1) = "&Copiar"
 MenuXpItem.Status(1) = "Copia el Texto Seleccionado"
 MenuXpItem.Items(2) = "&Deshacer"
 MenuXpItem.Status(2) = "Deshace la ultima Opción"
 MenuXpItem.Items(3) = "Se&leccionar Todo"
 MenuXpItem.Status(3) = "Selecciona todo el Texto"
 MenuXpItem.Items(4) = "&Salir"
 MenuXpItem.Status(4) = "Sale del Programa"
   
 MenuXpItem.Items(5) = "C&omplemento"
 MenuXpItem.Status(5) = "Menú Complemento"
 MenuXpItem.Items(6) = "&Proyecto"
 MenuXpItem.Status(6) = "Crea un Nuevo Proyecto"
   
 MenuXpItem.Items(7) = "Agregar &Formulario"
 MenuXpItem.Status(7) = "Agrega un Nuevo Formulario"
 MenuXpItem.Items(8) = "Agregar Formulario &MDI"
 MenuXpItem.Status(8) = "Agrega un Nuevo Formulario MDI"
 MenuXpItem.Items(9) = "Agregar &Módulo"
 MenuXpItem.Status(9) = "Agrega un Nuevo Módulo"
 MenuXpItem.Items(10) = "Agregar Módulo de Clase"
 MenuXpItem.Status(10) = "Agrega un Nuevo Módulo de Clase"
 MenuXpItem.Items(11) = "Agregar Control de &usuario"
 MenuXpItem.Status(11) = "Agrega un Nuevo Control de Usuario"
  
 MenuXpItem.Items(12) = "Servidor &1"
 MenuXpItem.Status(12) = "Primer Servidor"
 MenuXpItem.Items(13) = "Servidor &2"
 MenuXpItem.Status(13) = "Segundo Servidor"
 MenuXpItem.Items(14) = "Servidor &3"
 MenuXpItem.Status(14) = "Tercer Servidor"
 MenuXpItem.Items(15) = "Random"
 MenuXpItem.Status(15) = "Random Servidor"
 iniLoad
 For I = 33 To 255
  Call cmbShortCut(0).AddItem(Chr$(I))
 Next
 cmbShortCut(0).ListIndex = 34
 cmbShortCut(1).ListIndex = 0
 For I = 0 To Printer.FontCount - 1      '* Determina el número de fuentes.
  Call cmbFont.AddItem(Printer.Fonts(I)) '* Las mueve al cuadro de lista.
 Next
 cmbFont.ListIndex = 0
 cmbGradient.ListIndex = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If (Button = 2) Then Call cmdViewMenu_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim Frm As Form
 
 For Each Frm In Forms
  Call Unload(Frm)
 Next
 End
End Sub

Private Sub iniLoad()
 XPMenu.MenuArrowNormalColor = &H9900FF
 XPMenu.MenuArrowSelectColor = &HC56A31
 XPMenu.MenuBackColor = vbWhite
 XPMenu.MenuBorderColor = &H73B0BB
 XPMenu.MenuForeColor = &HD18410
 XPMenu.MenuHighLightBorderColor = &HC56A31
 XPMenu.MenuHighLightColor = &HEED5C4
 XPMenu.MenuMarginColor = &HDEEDEF
 XPMenu.MenuSelForeColor = &HC56A31
 XPMenu.MenuShadowColor = &HE0E0E0
 Call XPM_DALNet.Init("DALNet")
 Call XPM_DALNet.AddItem(5, MenuXpItem.Items(12))
 Call XPM_DALNet.AddItem(7, MenuXpItem.Items(13))
 Call XPM_DALNet.AddItem(21, MenuXpItem.Items(14))
 Call XPM_DALNet.AddItem(, , , , , , , , , False, True)
 Call XPM_DALNet.AddItem(12, MenuXpItem.Items(15))
    
 Call XPM_EFNet.Init("EFNet")
 Call XPM_EFNet.AddItem(16, MenuXpItem.Items(7), "Verdana", , vbBlue, , , , , , , True)
 Call XPM_EFNet.AddItem(18, MenuXpItem.Items(9))
 Call XPM_EFNet.AddItem(, , , , , , , , , False, True)
 Call XPM_EFNet.AddItem(21, MenuXpItem.Items(10))
 Call XPM_EFNet.AddItem(, , , , , , , , , False, True)
 Call XPM_EFNet.AddItem(11, MenuXpItem.Items(8), , , , , , , , , , , False, , True, vbAltMask, vbKeyF5)
 Call XPM_EFNet.AddItem(11, MenuXpItem.Items(11))
     
 Call XPMenu2.Init("Servers")
 Call XPMenu2.AddItem(17, MenuXpItem.Items(5), , , , , , , , True, , , False, , , , , XPM_DALNet)
 Call XPMenu2.AddItem(10, MenuXpItem.Items(6), , , , , , , , True, , , , , , , , XPM_EFNet)
  
 Call XPMenu.Init("Connect")
 Call XPMenu.AddItem(2, MenuXpItem.Items(0), , , , , , , , True, , , , , , , , XPMenu2)
 Call XPMenu.AddItem(, , , , , , , , , False, True)
 Call XPMenu.AddItem(4, MenuXpItem.Items(1))
 Call XPMenu.AddItem(6, MenuXpItem.Items(2))
 Call XPMenu.AddItem(, , , , , , , , , False, True)
 Call XPMenu.AddItem(14, MenuXpItem.Items(3), "Verdana", , vbRed, , True, , , , , , , , , vbCtrlMask + vbShiftMask, vbKeyF)
 Call XPMenu.AddItem(, "", , , , , , , , False, True)
 Call XPMenu.AddItem(19, MenuXpItem.Items(4), , , , , , , , , , , , , , , , , AddressOf MenuExit)
End Sub

Private Sub imgPicture_Click()
 Set imgPicture.Picture = LoadPicture(ShowDialog(True), vbResBitmap)
 Call XPMenu.StyleMenu(BackGround, imgPicture.Picture)
 cmdMenuPicture.Caption = "No &Picture"
 cmdMenuGradient.Caption = "Menu &Gradient"
End Sub

Private Sub picColor_Click()
 picColor.BackColor = ShowDialog
End Sub

Private Function ShowDialog(Optional ByVal WhatIts As Boolean = False) As Variant
 If (WhatIts = True) Then
  With cdDialog
   .Filter = "Todas las Imagenes|*.bmp;*.gif;*.jpg;*.jpeg;*.wmf"
   .flags = &H4
   .ShowOpen
   If (.CancelError Or .FileName = "") Then Exit Function
   ShowDialog = .FileName
  End With
 Else
  cdDialog.ShowColor
  ShowDialog = cdDialog.Color
 End If
End Function

Private Sub picGradient_Click(Index As Integer)
 picGradient(Index).BackColor = ShowDialog
End Sub
