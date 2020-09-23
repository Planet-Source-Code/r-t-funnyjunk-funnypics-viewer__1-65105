VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.UserControl ctlDownImgForm 
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   ScaleHeight     =   4140
   ScaleWidth      =   6705
   Begin MSComDlg.CommonDialog cdl1 
      Left            =   120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2025
      ScaleWidth      =   3705
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Idle"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3615
      End
      Begin VB.Image imgMain 
         Appearance      =   0  'Flat
         Height          =   975
         Left            =   1080
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ctlDownImgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Status As String
Public StatusVisible As Boolean
Dim imageFilter As String

Public Sub DisplayImage(sPath As String)
On Error GoTo Hell
    If Ambient.UserMode Then
        AsyncRead sPath, vbAsyncTypePicture
    End If
    
    StatusVisible = False
    imgMain.Visible = True
    RenderMe
Exit Sub
Hell:
    Status = "Error: Could not display image " & sPath
    RenderMe
End Sub

Public Sub HideImage()
    imgMain.Visible = False
    RenderMe
End Sub

Public Sub SaveImgToFile(Optional sPath As String)
    On Error GoTo Hell
    
    Dim ret
    
    cdl1.FileName = ""
    cdl1.Filter = "JPEG Image (*.jpg)|*.jpg|GIF Image (*.gif)|*.gif|Bitmap Image (*.bmp)|*.bmp|Meta Picture file (*.wmf)|*.wmf|Aldus Corporation format(*.tiff)|*.tiff|WordPerfect image format (*.WPG)|*.WPG|Paint Shop Pro format (*.PSP)|*.PSP|GEM Paint format (*IMG)|*.IMG"
    cdl1.CancelError = True
    cdl1.ShowSave
    
    If Dir(cdl1.FileName) = "" Then
        SavePicture imgMain.Picture, cdl1.FileName
    Else
        ret = MsgBox("A file with the name " & cdl1.FileTitle & " allready exists." & vbCrLf & "Do you want to overwrite it?", vbYesNo + vbInformation, "File Allready Exists")
        If ret = vbYes Then
            SavePicture imgMain.Picture, cdl1.FileName
        Else
            Exit Sub
        End If
    End If
    Exit Sub
Hell:
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    On Error Resume Next
    imgMain.Picture = AsyncProp.Value
    imageFilter = Right(AsyncProp.Status, 4)
    imgMain.Refresh
'    UserControl.Height = imgmain.Height
'    UserControl.Width = imgmain.Width + 50
    LogicalSize picBack, imgMain, 0
End Sub

Private Sub UserControl_Initialize()
    RenderMe
End Sub

Private Sub UserControl_Resize()
    RenderMe
End Sub

Private Sub RenderMe()
    picBack.Width = UserControl.Width
    picBack.Height = UserControl.Height
    If StatusVisible Then
        lblStatus.Caption = Status
        lblStatus.Visible = True
        lblStatus.Top = UserControl.Height / 2 - 100
        lblStatus.Left = UserControl.Width / 2 - lblStatus.Width / 2
    Else
        lblStatus.Visible = False
    End If
End Sub
