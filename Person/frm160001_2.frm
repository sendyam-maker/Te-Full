VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm160001_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工相片"
   ClientHeight    =   3465
   ClientLeft      =   150
   ClientTop       =   990
   ClientWidth     =   3135
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   3135
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   1850
      TabIndex        =   8
      Top             =   120
      Width           =   1200
   End
   Begin VB.PictureBox G_SeekPicColor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      Height          =   500
      Left            =   1320
      ScaleHeight     =   29
      ScaleMode       =   3  '像素
      ScaleWidth      =   29
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   500
   End
   Begin VB.PictureBox tmpPic 
      AutoRedraw      =   -1  'True
      DragMode        =   1  '自動
      Height          =   2300
      Left            =   600
      ScaleHeight     =   2235
      ScaleWidth      =   1785
      TabIndex        =   0
      Top             =   960
      Width           =   1840
      Begin VB.Image tmpImg 
         BorderStyle     =   1  '單線固定
         Height          =   960
         Left            =   420
         Stretch         =   -1  'True
         Top             =   390
         Visible         =   0   'False
         Width           =   1020
      End
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   7
      Top             =   690
      Width           =   690
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   390
      Width           =   690
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   90
      Width           =   690
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部　　門 ："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   690
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工姓名 ："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   390
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號 ："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   945
   End
End
Attribute VB_Name = "frm160001_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/17 Form2.0已修改
'Add by Amy 2017/07/21
Option Explicit

Public UpForm As Form


Private Sub cmdok_Click()
    '回前畫面
    Unload Me
    UpForm.cmdState = 5
    UpForm.PubShowNextData
End Sub

Private Sub Form_Load()
     MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm160001_2 = Nothing
End Sub

'載入照片
Public Function ReadPhoto(ByVal stST01 As String) As Boolean
    Dim PicRs As New ADODB.Recordset
    Dim file_num As Integer
    Dim bytes() As Byte
    Dim IsWmf As Boolean
    Dim pWidth As Integer '圖片寬度
    Dim pHeight As Integer '圖片高度
    Dim dblTmp As Double
    Dim sW As Integer, sH As Integer
    Dim stAttachFile As String

    ReadPhoto = False
    '清圖片
    tmpPic.Picture = LoadPicture()
    tmpImg.Picture = LoadPicture()
    G_SeekPicColor.Picture = LoadPicture()
    G_SeekPicColor.Width = 0
    G_SeekPicColor.Height = 0
    
    DoEvents
    Set PicRs = New ADODB.Recordset
    PicRs.CursorLocation = adUseClient
    PicRs.Open "select ImgByteFile.*,S1.st02 as Cst02,s2.st02 as Ust02 from ImgByteFile,staff S1,staff S2 where ibf05='3' and ibf01='000' and ibf02='" & stST01 & "' and ibf03='0' and ibf04='00' and ibf07=s1.st01(+) and ibf10=s2.st01(+) ", cnnConnection, adOpenStatic, adLockOptimistic
    If PicRs.RecordCount <> 0 Then
        ReadPhoto = True
        PicRs.MoveFirst
         If CheckStr(PicRs.Fields("ibf06")) = "3" Or CheckStr(PicRs.Fields("ibf06")) = "4" Or CheckStr(PicRs.Fields("ibf06")) = "6" Then
            IsWmf = True
            stAttachFile = App.path & "\NowPic.wmf"
         Else
            IsWmf = False
            stAttachFile = App.path & "\NowPic.jpg"
         End If
         
         'Add By Sindy 2017/8/10
'         If "" & PicRs.Fields("IBF15") <> "" Then
            If PUB_GetFtpFile(PicRs.Fields("IBF15"), stAttachFile, UCase("ImgByteFile")) = False Then
               Screen.MousePointer = vbDefault
               Exit Function
            End If
'         Else
'         '2017/8/10 END
'            ReDim bytes(Val(PicRs.Fields("ibf13").Value))
'            bytes() = PicRs.Fields("ibf14").GetChunk(Val(PicRs.Fields("ibf13").Value))
'            file_num = FreeFile
'            If IsWmf = False Then
'                Open App.path & "\NowPic.jpg" For Binary Access Write As #file_num
'            Else
'                Open App.path & "\NowPic.wmf" For Binary Access Write As #file_num
'            End If
'            Put #file_num, , bytes()
'            Close #file_num
'         End If
         
        G_SeekPicColor.Picture = LoadPicture(App.path & "\NowPic.jpg")
        pWidth = G_SeekPicColor.Width
        pHeight = G_SeekPicColor.Height
        sH = 0: sW = 0
        If pWidth < pHeight Then '以高的比例
            dblTmp = pHeight / tmpPic.Height
            sH = tmpPic.Height
        Else '以寬的比例
            dblTmp = pWidth / tmpPic.Width
            sW = tmpPic.Width
        End If
        If sW = 0 Then
            sW = pWidth / dblTmp
            If sW > tmpPic.Width Then
                '寬度等比例縮小後還是大於圖框寬,再以寬的比例縮放
                dblTmp = sW / tmpPic.Width
                sW = tmpPic.Width
                sH = sH / dblTmp
            End If
        ElseIf sH = 0 Then
            sH = pHeight / dblTmp
            If sH > tmpPic.Height Then
                '高度等比例縮小後還是大於圖框高,再以高的比例縮放
                dblTmp = sH / tmpPic.Height
                sH = tmpPic.Height
                sW = sW / dblTmp
            End If
        End If
        tmpImg.Width = sW: tmpImg.Height = sH
        Set tmpImg.Picture = G_SeekPicColor
        tmpPic.PaintPicture G_SeekPicColor, IIf(tmpPic.ScaleWidth / 2 - (sW / 2) < 0, 0, tmpPic.ScaleWidth / 2 - (sW / 2)), IIf(tmpPic.ScaleHeight / 2 - (sH / 2) < 0, 0, tmpPic.ScaleHeight / 2 - (sH / 2)), sW, sH
        Set tmpPic.Picture = tmpPic.Image
        
        If Dir(App.path & "\NowPic.jpg") <> "" Then
            Kill App.path & "\NowPic.jpg"
        End If
        If Dir(App.path & "\NowPic.wmf") <> "" Then
            Kill App.path & "\NowPic.wmf"
        End If
    End If
    Screen.MousePointer = vbDefault
End Function
