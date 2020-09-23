Attribute VB_Name = "modImgSize"
Option Explicit

Public sAccessCode As String
Public iMaxEntries As Integer

Public Sub LogicalSize(ContainerObj As Object, ImgObj As Object, ByVal Cushion As Integer)
    Dim VertChg, HorzChg As Integer
    Dim iRatio As Double
    Dim ActualH, ActualW As Integer
    Dim ContH, ContW As Integer
    On Error GoTo LogicErr


    With ImgObj 'hide picture While changing size
        .Visible = False
        .Stretch = False 'set actual size
    End With

    VertChg = 0: HorzChg = 0
    ActualH = ImgObj.Height 'actual picture height
    ActualW = ImgObj.Width 'actual picture width
    ContH = ContainerObj.Height - Cushion 'set max. picture height
    ContW = ContainerObj.Width - Cushion 'set max. picture width
    CenterCTL ContainerObj, ImgObj 'center picture
    
    If ImgObj.Top < Cushion Or ImgObj.Left < Cushion Then 'is picture larger than container
        If ActualH <> ActualW Then 'picture is Not square
            If ActualH > ActualW Then 'height is greater
                iRatio = (ActualH / ActualW) 'get ratio between height and width
                HorzChg = 10 'scale down by 10 units per Loop
                VertChg = CInt(Format(iRatio * 10, "####"))
            Else 'width is greater
                iRatio = (ActualW / ActualH) 'get ratio between height and width
                VertChg = 10 'scale down by 10 units per Loop
                HorzChg = CInt(Format(iRatio * 10, "####")) 'round number
            End If
            
        Else 'picture is square
            VertChg = 10 'scale both height and width equally
            HorzChg = 10
        End If
        
        Do Until ActualH <= ContH And ActualW <= ContW
            ActualH = ActualH - VertChg 'scale height down
            ActualW = ActualW - HorzChg 'scale width down
            
            If ActualH < 100 Then
                ActualH = 100 'set min. picture height=100
                Exit Do
            ElseIf ActualW < 100 Then
                ActualW = 100 'set min. picture width=100
                Exit Do
            End If
        Loop
        
        With ImgObj 'set new height and width
            .Stretch = True
            .Height = ActualH
            .Width = ActualW
        End With

    End If

    CenterCTL ContainerObj, ImgObj 'center picture in container
    ImgObj.Visible = True 'show picture
    Exit Sub
LogicErr:
    MsgBox "An Error occured While rescaling this image. Image size maybe invalid.", vbSystemModal + vbExclamation, "Resize Error!"
End Sub

Public Sub CenterCTL(FRMObj As Object, OBJ As Control)
    With OBJ
        .Top = (FRMObj.Height / 2) - (OBJ.Height / 2)
        .Left = (FRMObj.Width / 2) - (OBJ.Width / 2)
        .ZOrder
    End With
End Sub

