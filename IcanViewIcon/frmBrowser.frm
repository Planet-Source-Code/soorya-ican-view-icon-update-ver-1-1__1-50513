VERSION 5.00
Begin VB.Form frmBrowser 
   BackColor       =   &H00C0C0C0&
   Caption         =   "BlackMagic Browser"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1200
      TabIndex        =   3
      Top             =   5430
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   2820
      TabIndex        =   2
      Top             =   5430
      Width           =   1395
   End
   Begin VB.DirListBox Dir1 
      Height          =   4590
      Left            =   180
      TabIndex        =   1
      Top             =   690
      Width           =   5025
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   5025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "soorya@vsnl.com"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   3930
      TabIndex        =   5
      Top             =   60
      Width           =   1260
   End
   Begin VB.Label lblSelect 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Folder for Icons"
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   60
      Width           =   2295
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'storing the last surfed path, curDir does,t work here
    frmMain.LastPath = Dir1.List(Dir1.ListIndex)
    'clearing the clipboard
    Clipboard.Clear
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'storing the last surfed path, curDir does,t work here
    frmMain.LastPath = Dir1.List(Dir1.ListIndex)
    'copy the path to clipboard
    Clipboard.Clear
    Clipboard.SetText Dir1.List(Dir1.ListIndex)
    Unload Me
End Sub

Private Sub Dir1_Change()
    frmMain.txtBrowse.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Dir1.Path = frmMain.LastPath
End Sub


