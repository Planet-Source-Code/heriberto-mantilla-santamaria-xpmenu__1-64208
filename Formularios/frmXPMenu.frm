VERSION 5.00
Begin VB.Form frmXPMenu 
   BorderStyle     =   0  'None
   ClientHeight    =   3930
   ClientLeft      =   4410
   ClientTop       =   5955
   ClientWidth     =   3150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   262
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   210
   ShowInTaskbar   =   0   'False
   Tag             =   "XPMenu"
   Begin VB.PictureBox picPopUp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2550
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   5
      Top             =   2070
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picIconD 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2790
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3630
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picResult 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2355
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3405
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrHover 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   645
      Top             =   3225
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1515
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3405
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox ImgCheck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1215
      Picture         =   "frmXPMenu.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3345
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrActive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   60
      Top             =   3225
   End
   Begin VB.PictureBox picMenuBuffer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      Height          =   3165
      Left            =   0
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label lblWidth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   210
      TabIndex        =   6
      Top             =   3765
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmXPMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Public XPMenuClass As New clsXPMenu
 Public UpY         As Single
 
 Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Private Declare Function GetActiveWindow Lib "user32" () As Long
 Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Sub Form_Click()
 Dim SelectedItem As Long

 SelectedItem = XPMenuClass.GetHiLightedItem(UpY)
 If (XPMenuClass.IsTextItem(CInt(SelectedItem))) Then
  Call XPMenuClass.KillAllMenus
  If (XPMenuClass.GetItemProc(CInt(SelectedItem)) <> 0) Then
   Call CallWindowProc(XPMenuClass.GetItemProc(CInt(SelectedItem)), 0&, 0&, 0&, 0&)
  Else
   Call HandleClick(XPMenuClass.GetMenuName(), CInt(SelectedItem), XPMenuClass.GetItemText(CInt(SelectedItem)))
  End If
 End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim GetHiLight As Long

 GetHiLight = XPMenuClass.GetHiLightedItem(Y)
 If (GetHiLight = XPMenuClass.GetHilightNum) Then
  If (GetHiLight <> 0) Then Call MenuOver(XPMenuClass.GetItemText(CInt(GetHiLight)))
  Exit Sub
 End If
 Call XPMenuClass.SetHilightedItem(CInt(GetHiLight))
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 UpY = Y
End Sub

Private Sub tmrActive_Timer()
 Dim Frm As Form
    
 For Each Frm In Forms
  If (Frm.Tag = "XPMenu") And (GetActiveWindow() = Frm.hWnd) Then Exit Sub
 Next
 XPMenuClass.KillPopupMenus
 XPMenuClass.UnLoadMenu
End Sub

Private Sub tmrHover_Timer()
 Dim Pt   As POINTAPI, hWnd As Long
 
 Call GetCursorPos(Pt)
 hWnd = WindowFromPoint(Pt.X, Pt.Y)
 If (hWnd <> Me.hWnd) Then If (XPMenuClass.PopUpShown() = False) Then XPMenuClass.SetHilightedItem -1
End Sub
