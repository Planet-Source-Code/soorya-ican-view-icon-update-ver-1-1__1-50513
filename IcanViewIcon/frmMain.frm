VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Ican VIEW Icon 1.1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   8610
      ScaleHeight     =   945
      ScaleWidth      =   1485
      TabIndex        =   13
      Top             =   60
      Width           =   1485
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse a Folder"
         Height          =   315
         Left            =   60
         TabIndex        =   16
         Top             =   570
         Width           =   1365
      End
      Begin VB.OptionButton optIcon 
         BackColor       =   &H00000000&
         Caption         =   "Icon"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   30
         TabIndex        =   15
         Top             =   210
         Value           =   -1  'True
         Width           =   645
      End
      Begin VB.OptionButton optBmp 
         BackColor       =   &H00000000&
         Caption         =   "Image"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   690
         TabIndex        =   14
         Top             =   210
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   10140
      ScaleHeight     =   945
      ScaleWidth      =   1515
      TabIndex        =   9
      Top             =   60
      Width           =   1515
      Begin VB.OptionButton optIconCount 
         BackColor       =   &H00000000&
         Caption         =   "4001 to 6000"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   12
         Top             =   660
         Width           =   1275
      End
      Begin VB.OptionButton optIconCount 
         BackColor       =   &H00000000&
         Caption         =   "2001 to 4000"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton optIconCount 
         BackColor       =   &H00000000&
         Caption         =   "1 to 2000"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   60
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   180
      TabIndex        =   6
      Top             =   8220
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtFilePath 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   300
      Width           =   8415
   End
   Begin VB.TextBox txtBrowse 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   660
      Width           =   8415
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      Height          =   6825
      Left            =   180
      ScaleHeight     =   6765
      ScaleWidth      =   11445
      TabIndex        =   0
      Top             =   1050
      Width           =   11505
      Begin VB.PictureBox picScroller 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   6675
         Left            =   30
         ScaleHeight     =   6675
         ScaleWidth      =   11055
         TabIndex        =   4
         Top             =   30
         Width           =   11055
         Begin VB.Image imgIcon 
            Height          =   480
            Index           =   0
            Left            =   60
            Picture         =   "frmMain.frx":0CCA
            Stretch         =   -1  'True
            Top             =   90
            Width           =   480
         End
         Begin VB.Label lblIcon 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "9999"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   5
            Top             =   660
            Width           =   405
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   6765
         Left            =   11130
         TabIndex        =   3
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.Label lblViewingNow 
      BackStyle       =   0  'Transparent
      Caption         =   "Viewing Now :"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   4800
      TabIndex        =   18
      Top             =   60
      Width           =   2085
   End
   Begin VB.Label lblTotalFiles 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Files :"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   2640
      TabIndex        =   17
      Top             =   60
      Width           =   1965
   End
   Begin VB.Image imgErrorCheck 
      Height          =   480
      Left            =   11130
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgERROR 
      Height          =   480
      Left            =   10650
      Picture         =   "frmMain.frx":1594
      Top             =   8040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "soorya@vsnl.com"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   8
      Top             =   60
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":225E
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   0
      Left            =   750
      TabIndex        =   7
      Top             =   7950
      Width           =   9885
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LastPath As String
'NOTE:
'very often when we make programs. finding proper icons and bitmaps
'is a boring job..
'so..
'i made this utility in 3 hrs... in the middle of my new project.
'hope all the programmers can put this exe in the desktop and use it often.
'ur precious feed back will make me to send more stuff...
'ENJOY !!!

'soorya@vsnl.com
'19-12-2003
Dim i, X As Long
Dim iw, ih, m  As Integer
Public pStoped As Boolean
Dim FileEXTENTION As String
Dim totIcons As Integer
Dim ICONFrom, ICONTill As Integer


Private Sub Form_Load()
    'init variables
    iw = 480 'width of each img
    ih = 480  'width of each img
    m = 100  'margin spaceing between each img
    imgIcon(0).Visible = False
    lblIcon(0).Visible = False
    
    'init size
    imgIcon(0).Width = iw
    imgIcon(0).Height = ih
    VScroll1.LargeChange = iw / 15
    VScroll1.SmallChange = iw / 15
    
    ICONFrom = 0
    'init file catagory
    FileEXTENTION = ".ico"
End Sub


Private Sub cmdBrowse_Click()
    txtBrowse.Text = ""
    'here vbmodal is used cos' user can't proceed without browsing
    frmBrowser.Show vbModal
    If Clipboard.GetText = "" Then Exit Sub
    'paste the clipboard path into text box
    txtBrowse.Text = Clipboard.GetText
    Call PopulateIcons
End Sub

Private Sub PopulateIcons()
    Dim s As String
    On Error Resume Next
    
    'reset text boxes and ...
    Call ResetAll
        
    'init the pos
    imgIcon(0).Left = m
    imgIcon(0).Top = m
    lblIcon(0).Left = imgIcon(0).Left
    lblIcon(0).Top = ih + lblIcon(0).Height
    picScroller.Height = ih + lblIcon(0).Height
    
    pStoped = False
    
    If optIcon.Value = True Then
        FileEXTENTION = "*.ico"
    Else
        FileEXTENTION = "*.bmp;*.jpg;*.jpeg;*.gif"
    End If
    
    'pattern to get only .ico or .bmp files
    File1.Pattern = FileEXTENTION

    'If Right(txtBrowse.Text, 1) <> "\" Then txtBrowse.Text = txtBrowse.Text & "\"
    'file control is used to get all the icon/bmp files in alphabatic order
    File1.Path = txtBrowse.Text
    
    lblTotalFiles = "Total Files : " & File1.ListCount
    
    If File1.ListCount < 1 Then
        If optIcon.Value = True Then
            MsgBox "There is no Icon File in the selected path !!!"
        Else
            MsgBox "There is no Image File in the selected path !!!"
        End If
            
        'reset text boxes and ...
        Call ResetAll
    End If
    
    'showing progress bar form
    frmProgress.Show
    i = 0
    totIcons = 0
    
    ICONTill = IIf(File1.ListCount - ICONFrom > 2000, 2000, File1.ListCount - ICONFrom)

    For X = ICONFrom To (ICONFrom + ICONTill) - 1 'File1.ListCount - 1
        'watch user pressed cancel button and reset all
        If pStoped = True Then Exit For
        
        'we have the imgIcon(0) already
        If i > 0 Then
            'populating image box and labels at run time as per the file count
            Load imgIcon(i)
            Load lblIcon(i)
            
            'just to put the icons in proper rows and colomns
            If imgIcon(i - 1).Left + iw + m < picScroller.Width Then
                'same row
                imgIcon(i).Left = imgIcon(i - 1).Left + iw + m '+ vs
                imgIcon(i).Top = imgIcon(i - 1).Top
            Else
                'next new row
                imgIcon(i).Left = m
                imgIcon(i).Top = imgIcon(i - 1).Top + ih + m + lblIcon(i - 1).Height
            End If
        End If
        
        'positioning the labels
        lblIcon(i).Left = imgIcon(i).Left
        lblIcon(i).Top = imgIcon(i).Top + ih
        
        'showing is must
        imgIcon(i).Visible = True
        lblIcon(i).Visible = True
        
        If isFileOK(File1.Path & "\" & File1.List(X)) = True Then
            imgIcon(i).Picture = LoadPicture(File1.Path & "\" & File1.List(X)) 'Picture1.Image
        Else
            imgIcon(i).Picture = imgERROR.Picture 'Picture1.Image
        End If

        'loading icons and setting of other params
        imgIcon(i).ToolTipText = File1.List(X) 'Left(File1.List(x), InStrRev(File1.List(x), ".") - 1) 'only filename without ext
        lblIcon(i).ToolTipText = File1.List(X) 'Left(File1.List(x), InStrRev(File1.List(x), ".") - 1) 'only filename without ext
        lblIcon(i).Caption = X + 1
        
        
        'here is my funny and simple progress bar, used to avoid extra controls
        pbr frmProgress, ICONTill - 1, i + 1
        
        'showing the path currently being populated
        frmProgress.lblPath = File1.List(X)
        lblViewingNow.Caption = "Viewing Now : " & i + 1
        
        picScroller.Height = imgIcon(i).Top + ih + m + lblIcon(i).Height
        VScroll1.Max = picScroller.Height / 15 - picContainer.Height / 15
        
        'put this inside the loop...
        Screen.MousePointer = vbHourglass
        totIcons = i
        i = i + 1
        
        'must so that windowzzz can do other events too...
        DoEvents
    Next
    
    'watch user pressed cancel button and reset all
    If pStoped = True Then Call ResetAll
    
    Unload frmProgress
    Screen.MousePointer = vbDefault
    picScroller.Refresh
    Exit Sub
End Sub

Private Sub ResetAll()
    picScroller.Height = imgIcon(0).Height
    imgIcon(0).Visible = False
    lblIcon(0).Visible = False
    
    For i = 1 To imgIcon.Count - 1
        'kill old stuff
        Unload imgIcon(i)
        Unload lblIcon(i)
        Screen.MousePointer = vbHourglass
        'must so that windowzzz can do other events too...
        DoEvents
    Next
    
    lblTotalFiles.Caption = "Total Files : " & 0
    lblViewingNow.Caption = "Viewing Now : " & "0"
    
    txtFilePath.Text = ""
    lblIcon(0).Caption = 1
    
    Screen.MousePointer = vbDefault
    Unload frmProgress
    picScroller.Refresh
End Sub


Public Function isFileOK(ByRef fName As String) As Boolean
    On Error GoTo errr
    imgErrorCheck.Picture = LoadPicture(fName)
    isFileOK = True
    Exit Function
errr:
    isFileOK = False
End Function


Private Sub imgIcon_Click(Index As Integer)
    txtFilePath.Text = txtBrowse.Text & "\" & imgIcon(Index).ToolTipText '& Mid(FileEXTENTION, 2)
    'lblIconFileName.Caption = "File Name:  " & imgIcon(Index).ToolTipText & Mid(FileEXTENTION, 2)
    lblViewingNow.Caption = "Viewing Now : " & totIcons
    'to get the path in the clipboard
    Clipboard.Clear
    Clipboard.SetText txtFilePath.Text
End Sub


Private Sub imgIcon_DblClick(Index As Integer)
    If optIcon.Value = True Then
        MsgBox "Sorry, You can't view this Icon in Enlarged Mode"
    Else
        frmImageViewer.fName = txtFilePath.Text
        frmImageViewer.Caption = imgIcon(Index).ToolTipText
        frmImageViewer.Show
    End If
End Sub

Private Sub imgIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then 'right mouse
        Dim sPath, dPath As String
        sPath = txtBrowse.Text & "\" & imgIcon(Index).ToolTipText '& Mid(FileEXTENTION, 2)
        
        On Error GoTo er
        If Button = 2 Then
            Clipboard.Clear
            frmBrowser.lblSelect.Caption = "Select a Folder where to Copy..."
            frmBrowser.Show vbModal
            dPath = Clipboard.GetText
            FileCopy sPath, dPath & "\" & imgIcon(Index).ToolTipText '& Mid(FileEXTENTION, 2)
            MsgBox "Copy Done Successfully"
        End If
    End If
    Exit Sub
er:
    Clipboard.Clear
    MsgBox "Copy Failed, Try again..."
End Sub

Private Sub optIconCount_Click(Index As Integer)
    If Index = 0 Then ICONFrom = 0
    If Index = 1 Then ICONFrom = 2000
    If Index = 2 Then ICONFrom = 4000
    
    If txtBrowse.Text <> "" Then Call PopulateIcons
End Sub

Private Sub VScroll1_Change()
    Call VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
    'If picScroller.Top > picScroller.Height - picContainer.Height Then
        picScroller.Top = -CLng(VScroll1.Value) * 15
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'must when using modules in our project
    End
End Sub

