VERSION 5.00
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   5250
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timTime 
      Interval        =   1
      Left            =   840
      Top             =   2520
   End
   Begin VB.PictureBox Maskk 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   240
      Picture         =   "frmMain.frx":452D
      ScaleHeight     =   570
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Galaxy BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      TabIndex        =   4
      Text            =   "Password"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Galaxy BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   480
      TabIndex        =   3
      Text            =   "User Name"
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Image imgGreenP 
      Height          =   225
      Left            =   3210
      Picture         =   "frmMain.frx":4498F
      Top             =   1635
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgRedP 
      Height          =   225
      Left            =   3210
      Picture         =   "frmMain.frx":44C5E
      Top             =   1635
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblCheck 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Checking..."
      BeginProperty Font 
         Name            =   "Galaxy BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgRedN 
      Height          =   225
      Left            =   3210
      Picture         =   "frmMain.frx":44F45
      Top             =   1155
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgGreenN 
      Height          =   225
      Left            =   3210
      Picture         =   "frmMain.frx":4522C
      Top             =   1155
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbltime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Galaxy BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label MoveForm 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblClear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "C&lear"
      BeginProperty Font 
         Name            =   "Galaxy BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Galaxy BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Galaxy BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const RGN_OR = 2
Private lngRegion As Long
Dim Trfl
Dim Trfl2

Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long
  Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
  Dim lngRgnFinal As Long, lngRgnTmp As Long
  Dim lngStart As Long, lngRow As Long
  Dim lngCol As Long
  If lngTransColor& < 1 Then
    lngTransColor& = GetPixel(picSource.hDC, 0, 0)
  End If
  lngHeight& = picSource.Height / Screen.TwipsPerPixelY
  lngWidth& = picSource.Width / Screen.TwipsPerPixelX
  lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
  For lngRow& = 0 To lngHeight& - 1
    lngCol& = 0
    Do While lngCol& < lngWidth&
      Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) = lngTransColor&
        lngCol& = lngCol& + 1
      Loop
      If lngCol& < lngWidth& Then
        lngStart& = lngCol&
        Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) <> lngTransColor&
          lngCol& = lngCol& + 1
        Loop
        If lngCol& > lngWidth& Then lngCol& = lngWidth&
        lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
        lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
        DeleteObject (lngRgnTmp&)
      End If
    Loop
  Next
  RegionFromBitmap& = lngRgnFinal&
  
End Function
Sub ChangeMask()
'On Error Resume Next ' In case of error
' This is also part of Dos's Dos-Shape example. To update if the skin is changed
  Dim lngRetr As Long
  lngRegion& = RegionFromBitmap(Maskk)
  lngRetr& = SetWindowRgn(Me.hWnd, lngRegion&, True)
End Sub

Private Sub Form_Load()
Maskk.AutoSize = True
MoveForm.Height = frmmain.Maskk.Height ' the move label
MoveForm.Width = frmmain.Maskk.Width ' the move label
MoveForm.Top = 0
MoveForm.Left = 0
MoveForm.ZOrder 1
Call frmmain.ChangeMask
End Sub

Private Sub lblCancel_Click()
End
End Sub

Private Sub lblCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel.ForeColor = vbBlack
End Sub

Private Sub lblCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblCancel.ForeColor = &HC0&
End Sub

Private Sub lblClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClear.ForeColor = vbBlack
End Sub

Private Sub lblClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClear.ForeColor = &HC0&
End Sub

Private Sub lblOK_Click()
Static a As Long
If Trfl = True And Trfl2 = True Then
End
Else
    If a = 3 Then
    ''do the boobo
    a = 0
    Else
    a = a + 1
    End If
End If
End Sub

Private Sub lblOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOK.ForeColor = vbBlack
End Sub

Private Sub lblOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOK.ForeColor = &HC0&
End Sub

Private Sub MoveForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then ' Left button
    ReleaseCapture
    Call SendMessage(Me.hWnd, &HA1, 2, 0)
End If
End Sub

Private Sub timTime_Timer()
lbltime.Caption = Time
End Sub

Private Sub txtName_Change()
lblCheck.Visible = True
Trfl = txtName.Text Like "[<]?[Z]?[E]?[X]?[X]?[>]"
If Trfl = True Then
imgRedN.Visible = False
imgGreenN.Visible = True
lblCheck.Visible = False
txtPass.SetFocus
Else
imgRedN.Visible = True
imgGreenN.Visible = False
End If
End Sub

Private Sub txtName_Click()
txtName.Text = ""
End Sub

Private Sub txtName_LostFocus()
If txtName.Text = "" Then
txtName.Text = "User Name"
Else
End If
End Sub

Private Sub txtPass_Change()
lblCheck.Visible = True
Trfl2 = txtPass.Text Like "[3]?[1]?[8]?[5]?[9]?[6]?[2]"
If Trfl2 = True Then
imgRedP.Visible = False
imgGreenP.Visible = True
lblCheck.Visible = False
Else
imgRedP.Visible = True
imgGreenP.Visible = False
End If
End Sub

Private Sub txtPass_Click()
txtPass.Text = ""
txtPass.PasswordChar = "*"
End Sub

Private Sub txtPass_GotFocus()
txtPass.Text = ""
txtPass.PasswordChar = "*"
End Sub

Private Sub txtPass_LostFocus()
If txtPass.Text = "" Then
txtPass.Text = "Password"
txtPass.PasswordChar = ""
Else
txtPass.PasswordChar = "*"
End If
End Sub
