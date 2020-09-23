VERSION 5.00
Begin VB.Form frmProgress 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCounter 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "99%"
      Top             =   630
      Width           =   495
   End
   Begin VB.TextBox txtHider 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3420
      TabIndex        =   1
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "soorya@vsnl.com"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   6660
      TabIndex        =   4
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Image imgProgressBar 
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Left            =   60
      Picture         =   "frmProgress.frx":0000
      Stretch         =   -1  'True
      Top             =   570
      Width           =   8190
   End
   Begin VB.Label lblPath 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Path: "
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   8055
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oldX As Long, _
    oldY As Long, _
    isMoving As Boolean

'***********************************************************************
'api for on top
' Declare functions obtained from the API Text Viewer
Private Declare Function SetWindowPos Lib _
    "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
'***********************************************************************

Private Sub cmdCancel_Click()
    lblPath.ForeColor = vbRed
    lblPath.Caption = "Please wait... Cancelation Process is going on..."
    frmMain.pStoped = True
    'Unload Me
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'for ontop form
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

'&&&&&&&&&&&&&&&&&&&&& FOR MOUSE DRAG &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isMoving = True
    oldX = X '* 15
    oldY = Y '* 15
End Sub

Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If isMoving = True Then
        X = X '* 15
        Y = Y '* 15
        Me.Left = Me.Left - (oldX - X)
        Me.Top = Me.Top - (oldY - Y)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isMoving = False
End Sub




Private Sub lblPath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isMoving = True
    oldX = X '* 15
    oldY = Y '* 15
End Sub

Private Sub lblPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If isMoving = True Then
        X = X '* 15
        Y = Y '* 15
        Me.Left = Me.Left - (oldX - X)
        Me.Top = Me.Top - (oldY - Y)
    End If
End Sub

Private Sub lblPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isMoving = False
End Sub
