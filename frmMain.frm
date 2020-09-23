VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "FunnyJunk Pic View"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5280
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ComboBox Cbo 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   9855
   End
   Begin FunnyPics.ctlDownImgForm Img 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11668
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim colURL As New Collection

Private Sub Cbo_Click()
    Img.Status = "Retrieving Picture"
    Img.StatusVisible = True
    'Img.HideImage
    Img.DisplayImage ParseForPic(colURL(Cbo.ListIndex + 1))
End Sub

Private Sub Form_Activate()
    Cbo.ListIndex = 0
End Sub

Private Sub Form_Load()
Dim sHTML As String, iStart As Long, iStop As Long
    sHTML = WebGetHTML("http://www.funnyjunk.com")
    iStart = InStrRev(sHTML, "Funny Pictures<BR><BR></BIG></BIG>#1 ") + 34

    If iStart > 0 Then
        iStop = InStrRev(sHTML, "</A><BR><br>")

        If iStop > 0 Then
            Text1 = Mid(sHTML, iStart, iStop - iStart)
        End If
    End If

    Text1 = Replace(Text1, "<B>", "")
    Text1 = Replace(Text1, "</B>", "")
    Text1 = Replace(Text1, "<BR>", "")
    Text1 = Replace(Text1, vbNewLine, "")

Dim sPics() As String, i As Integer
Dim sCur As String, sURL As String, sTitle As String
    sPics = Split(Text1, "#")
    
    For i = 1 To UBound(sPics)
        sCur = sPics(i)
        
        iStart = InStr(1, sCur, "/")
        iStop = InStr(iStart, sCur, """")
        sURL = Mid(sCur, iStart, iStop - iStart)
        
        colURL.Add "http://www.funnyjunk.com" & sURL
        
        iStop = Len(sCur) - 3
        iStart = InStrRev(sCur, ">", iStop) + 1
        sTitle = Mid(sCur, iStart, iStop - iStart)
        Cbo.AddItem sTitle
    Next i
End Sub

Private Function ParseForPic(ByRef sURL As String) As String
Dim sHTML As String, iStart As Long, iStop As Long
    sHTML = WebGetHTML(sURL)
    iStop = InStrRev(sHTML, "border=""0"" alt=""Funny Pictures""") - 2
    iStart = InStrRev(sHTML, "src=", iStop) + 5
    ParseForPic = Mid(sHTML, iStart, iStop - iStart)
End Function

Private Sub Form_Resize()
    Cbo.Width = Me.ScaleWidth
    Img.Width = Cbo.Width
    Img.Height = Me.ScaleHeight - Cbo.Height
End Sub
