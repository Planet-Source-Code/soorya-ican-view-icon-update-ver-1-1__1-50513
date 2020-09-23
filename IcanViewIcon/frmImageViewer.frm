VERSION 5.00
Begin VB.Form frmImageViewer 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enlarged View"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2160
   Icon            =   "frmImageViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   2160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picViewer 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   30
      ScaleHeight     =   405
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image imgViewer 
      Height          =   480
      Left            =   750
      Top             =   450
      Width           =   480
   End
End
Attribute VB_Name = "frmImageViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fName As String

Private Sub Form_Load()
    Call LoadImage
End Sub

Private Sub LoadImage()
    imgViewer.Left = 0
    imgViewer.Top = 0
    
    'checking errored picture and replace it with error image
    If frmMain.isFileOK(fName) = True Then
        picViewer.Picture = LoadPicture(fName)
        imgViewer.Picture = LoadPicture(fName)
    Else
        picViewer.Picture = frmMain.imgERROR.Picture
    End If
    
    Me.Width = picViewer.Width + 50
    Me.Height = picViewer.Height + 300 '300 for control bar
    imgViewer.Width = Me.Width
    imgViewer.Height = Me.Height
    imgViewer.Picture = picViewer.Picture
    Exit Sub
End Sub
