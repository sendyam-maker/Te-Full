VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPic001 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  '雙線固定對話方塊
   ClientHeight    =   5136
   ClientLeft      =   840
   ClientTop       =   1416
   ClientWidth     =   8148
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   428
   ScaleMode       =   3  '像素
   ScaleWidth      =   679
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdok 
      Caption         =   "刪除(&D)"
      Enabled         =   0   'False
      Height          =   360
      Index           =   10
      Left            =   6975
      TabIndex        =   59
      Top             =   2910
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox G_SeekPicColor2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   3015
      Left            =   8400
      ScaleHeight     =   247
      ScaleMode       =   3  '像素
      ScaleWidth      =   264
      TabIndex        =   56
      Top             =   3480
      Width           =   3210
   End
   Begin VB.PictureBox pic2 
      BackColor       =   &H80000005&
      Height          =   585
      Left            =   8280
      ScaleHeight     =   45
      ScaleMode       =   3  '像素
      ScaleWidth      =   56
      TabIndex        =   54
      Top             =   720
      Width           =   720
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "本所案號匯入"
      Height          =   360
      Index           =   8
      Left            =   2190
      TabIndex        =   52
      Top             =   270
      Width           =   1350
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Height          =   4200
      Left            =   120
      TabIndex        =   44
      Top             =   2880
      Width           =   4350
      Begin VB.CommandButton Cmd1 
         Caption         =   "關閉"
         Height          =   360
         Index           =   2
         Left            =   3530
         TabIndex        =   55
         Top             =   360
         Width           =   700
      End
      Begin VB.CommandButton Cmd1 
         Caption         =   "完成"
         Height          =   360
         Index           =   1
         Left            =   2800
         TabIndex        =   53
         Top             =   360
         Width           =   700
      End
      Begin VB.PictureBox tmpPic2 
         Height          =   3200
         Left            =   0
         ScaleHeight     =   263
         ScaleMode       =   3  '像素
         ScaleWidth      =   295
         TabIndex        =   48
         Top             =   840
         Width           =   3588
         Begin VB.Image tmpImg2 
            Height          =   1770
            Left            =   1425
            Stretch         =   -1  'True
            Top             =   255
            Width           =   1890
         End
      End
      Begin VB.CommandButton Cmd1 
         Caption         =   "搜尋"
         Height          =   360
         Index           =   0
         Left            =   2085
         TabIndex        =   51
         Top             =   360
         Width           =   700
      End
      Begin VB.TextBox txtSystem 
         Height          =   264
         Left            =   0
         MaxLength       =   3
         TabIndex        =   46
         Top             =   360
         Width           =   520
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   50
         Top             =   360
         Width           =   400
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   49
         Top             =   360
         Width           =   300
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   530
         MaxLength       =   6
         TabIndex        =   47
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Lbl10 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   45
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.CheckBox chk1 
      Caption         =   "無圖式"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      TabIndex        =   43
      Top             =   720
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "複製(&C)"
      Height          =   360
      Index           =   7
      Left            =   3555
      TabIndex        =   38
      Top             =   270
      Width           =   780
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H80000005&
      Height          =   585
      Left            =   8220
      ScaleHeight     =   45
      ScaleMode       =   3  '像素
      ScaleWidth      =   56
      TabIndex        =   37
      Top             =   0
      Width           =   720
   End
   Begin VB.PictureBox tmpPic 
      Height          =   4455
      Left            =   0
      ScaleHeight     =   367
      ScaleMode       =   3  '像素
      ScaleWidth      =   295
      TabIndex        =   36
      Top             =   660
      Width           =   3588
      Begin VB.Image tmpImg 
         Height          =   1770
         Left            =   1425
         Stretch         =   -1  'True
         Top             =   1095
         Width           =   1890
      End
   End
   Begin VB.PictureBox G_SeekPicColor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   3015
      Left            =   8400
      ScaleHeight     =   247
      ScaleMode       =   3  '像素
      ScaleWidth      =   264
      TabIndex        =   35
      Top             =   360
      Width           =   3210
   End
   Begin VB.PictureBox CropPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   210
      Left            =   7500
      ScaleHeight     =   14
      ScaleMode       =   3  '像素
      ScaleWidth      =   29
      TabIndex        =   34
      Top             =   3870
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "灰階掃描(&A)"
      Enabled         =   0   'False
      Height          =   360
      Index           =   6
      Left            =   1170
      TabIndex        =   1
      Top             =   270
      Width           =   1110
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   3765
      TabIndex        =   31
      Top             =   2940
      Width           =   2160
      Begin VB.OptionButton optColor 
         Caption         =   "彩色"
         Height          =   180
         Index           =   1
         Left            =   1380
         TabIndex        =   33
         Top             =   45
         Width           =   675
      End
      Begin VB.OptionButton optColor 
         Caption         =   "黑白/灰階"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   32
         Top             =   30
         Value           =   -1  'True
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "灰階 切換"
      Height          =   360
      Index           =   9
      Left            =   30
      TabIndex        =   0
      Top             =   270
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   810
      Left            =   3765
      TabIndex        =   26
      Top             =   3225
      Visible         =   0   'False
      Width           =   3525
      Begin VB.TextBox txtUserNo 
         Height          =   270
         Left            =   900
         MaxLength       =   6
         TabIndex        =   27
         Top             =   60
         Width           =   732
      End
      Begin VB.TextBox txtPassword 
         Height          =   270
         IMEMode         =   3  '暫止
         Left            =   900
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   28
         Top             =   435
         Width           =   1575
      End
      Begin MSForms.Label lblUserName 
         Height          =   255
         Left            =   1710
         TabIndex        =   60
         Top             =   60
         Width           =   1350
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Size            =   "2381;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號"
         Height          =   270
         Index           =   2
         Left            =   45
         TabIndex        =   30
         Top             =   60
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "密碼"
         Height          =   270
         Index           =   1
         Left            =   45
         TabIndex        =   29
         Top             =   435
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7440
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      Height          =   345
      Left            =   -60
      ScaleHeight     =   300
      ScaleWidth      =   13644
      TabIndex        =   24
      Top             =   -75
      Width           =   13695
      Begin VB.CommandButton cmdok 
         Caption         =   "'選擇裝置(H)"
         Enabled         =   0   'False
         Height          =   360
         Index           =   5
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "黑白掃描(D)"
         Enabled         =   0   'False
         Height          =   360
         Index           =   4
         Left            =   1200
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "複製、貼上只對圖框有效"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   6000
         TabIndex        =   39
         Top             =   60
         Width           =   2145
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "專利代表圖"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   90
         TabIndex        =   25
         Top             =   75
         Width           =   900
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7470
      Top             =   3300
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "繼續(&X)"
      Height          =   360
      Index           =   3
      Left            =   7050
      TabIndex        =   5
      Top             =   270
      Width           =   1080
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "存檔(&S)"
      Height          =   360
      Index           =   2
      Left            =   6255
      TabIndex        =   4
      Top             =   270
      Width           =   780
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "選擇檔案(&F)"
      Default         =   -1  'True
      Height          =   360
      Index           =   1
      Left            =   5130
      TabIndex        =   3
      Top             =   270
      Width           =   1110
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "貼上(&P)"
      Height          =   360
      Index           =   0
      Left            =   4335
      TabIndex        =   2
      Top             =   270
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   3
      Left            =   3885
      TabIndex        =   42
      Top             =   990
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   8
      Left            =   4830
      TabIndex        =   41
      Top             =   990
      Width           =   3255
      VariousPropertyBits=   27
      Size            =   "5741;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      Caption         =   $"frmPic001.frx":0000
      ForeColor       =   &H000000C0&
      Height          =   795
      Left            =   4350
      TabIndex        =   40
      Top             =   4290
      Width           =   3585
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   7
      Left            =   4830
      TabIndex        =   23
      Top             =   2715
      Width           =   2685
      VariousPropertyBits=   27
      Size            =   "4736;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   6
      Left            =   4830
      TabIndex        =   22
      Top             =   2490
      Width           =   2685
      VariousPropertyBits=   27
      Size            =   "4736;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   5
      Left            =   4830
      TabIndex        =   21
      Top             =   2250
      Width           =   2685
      VariousPropertyBits=   27
      Size            =   "4736;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   4
      Left            =   4830
      TabIndex        =   20
      Top             =   2010
      Width           =   2685
      VariousPropertyBits=   27
      Size            =   "4736;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   3
      Left            =   4830
      TabIndex        =   19
      Top             =   1740
      Width           =   2685
      VariousPropertyBits=   27
      Size            =   "4736;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   2
      Left            =   4830
      TabIndex        =   18
      Top             =   1500
      Width           =   2685
      VariousPropertyBits=   27
      Size            =   "4736;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   1
      Left            =   4830
      TabIndex        =   17
      Top             =   1260
      Width           =   2685
      VariousPropertyBits=   27
      Size            =   "4736;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "檔案大小："
      Height          =   180
      Left            =   3885
      TabIndex        =   16
      Top             =   2715
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Time："
      Height          =   180
      Left            =   4245
      TabIndex        =   15
      Top             =   2505
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Date："
      Height          =   180
      Left            =   4290
      TabIndex        =   14
      Top             =   2265
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "ID："
      Height          =   180
      Left            =   4425
      TabIndex        =   13
      Top             =   2010
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Update："
      Height          =   180
      Left            =   3765
      TabIndex        =   12
      Top             =   1995
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Time："
      Height          =   180
      Left            =   4245
      TabIndex        =   11
      Top             =   1740
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Create："
      Height          =   180
      Left            =   3765
      TabIndex        =   10
      Top             =   1245
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Date："
      Height          =   180
      Left            =   4290
      TabIndex        =   9
      Top             =   1515
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ID："
      Height          =   180
      Index           =   0
      Left            =   4425
      TabIndex        =   8
      Top             =   1260
      Width           =   360
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   0
      Left            =   4830
      TabIndex        =   7
      Top             =   735
      Width           =   1709
      VariousPropertyBits=   27
      Size            =   "3014;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   3885
      TabIndex        =   6
      Top             =   735
      Width           =   900
   End
End
Attribute VB_Name = "frmPic001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/11/30 改成Form2.0 ;lblUserName、lbl1(index)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'create by nickc 2005/10/28
Option Explicit

'定義輸出輸入的本所案號
Public oCP01 As String
Public oCP02 As String
Public oCP03 As String
Public oCP04 As String
'add by nickc 2007/11/16 加入查名單，使用圖片
Public oPic As PictureBox
Public oImg As Image
Public oRtPic As Boolean
Public UpForm As Form '前一畫面

Dim PicRs As New ADODB.Recordset
Dim file_num As Integer
Dim bytes() As Byte
'Public WithEvents DIBDither As cDIBDither         ' DIB Dither object  (1, 4, 8 bpp)
'Public WithEvents DIBFilter As cDIBFilter                  ' DIB Filter object  (24 bpp)
'Dim DIBPal               As New cDIBPal                 ' DIB Palette object (1, 4, 8 bpp)
'Dim DIBSave              As New cDIBSave             ' Save object (BMP)  (1, 4, 8, 24 bpp)
'Dim DIBbpp               As Byte                             ' Current color depth

'Dim DIB As New cDIB
'Dim PicFrame As New cFrame
Dim IsSave As Boolean
'add by nickc 2005/12/21
Dim G_PicTemp As StdPicture      '暫存的圖檔
'========Sizing grip staff===========
'Dim XGrip(2) As Long, YGrip(2) As Long
'Dim bMoving As Boolean, bSizing As Boolean
'Dim xStart As Long, yStart As Long
'Const GripSize = 90
'Dim G_SeekPic As StdPicture
''====================================
Dim tmpObject As Variant
Dim i As Integer
Private Type SeekState
    Caption As String
    Enabled As Boolean
    Visible As Boolean
End Type
Dim SeekCmdok(0 To 9) As SeekState 'Modify by Amy 2018/07/02 原:7
Dim IsWmf As Boolean
Dim lIdxNew As Long
'add by nickc 2006/05/24
Const TwipsPerPixel       As Long = 15    'Is this ever not true?
Dim m_Image                    As New cImage
Dim m_Jpeg  As cJpeg
'Add By Sindy 2012/6/14
Public strWorkType As String '1:為員工照片;否則為代表圖
Public bolQuery As Boolean
Dim m_ibf05 As String
'2012/6/14 End
'Added by Lydia 2015/8/5 查名單電子化
Public m_TMQ As String '商標查名單0:圖形,1:文字1,2:文字2 'Modified by Lydia 2024/09/06 查名單-網中：A
'Added by Lydia 2015/11/06 查名單電子化
'Dim strLoadPath As String '讀取前次設定路徑 'Remove by Lydia 2016/05/26
Dim bolTMQUpd As Boolean '更新查名單輸入的圖片
Public bolNoMsg As Boolean 'Add by Amy 2017/01/20 不show訊息
'Add by Amy 2018/07/02
Dim oLbl As Object
Dim CountPhoto As Integer '代表圖圖示數量
Dim bolCall As Boolean 'Added by Morgan 2023/7/27

'add by sonia 2014/4/8 勾選無圖式
Private Sub Chk1_Click()
Dim BytesS() As Byte
Dim BytesVal As String
Dim PicRs As New ADODB.Recordset
Dim p_FileName As String, strFtpPath As String
   
   If Chk1.Value = vbChecked Then
      
      PUB_DelFtpFile2 oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04 & "-1", "", UCase("ImgByteFile") 'Add By Sindy 2017/8/10 檔案改放 FTP,必須在DB資料刪除前執行
      strSql = "delete ImgByteFile where ibf01='" & oCP01 & "' and ibf02='" & oCP02 & "' and ibf03='" & oCP03 & "' and ibf04='" & oCP04 & "'"
      cnnConnection.Execute strSql, intI
      
      strSql = "select * from ImgByteFile where ibf01='000' and ibf02='000000' and ibf03='0' and ibf04='01'  "
      Set PicRs = New ADODB.Recordset
      PicRs.CursorLocation = adUseClient
      PicRs.Open strSql, cnnConnection, adOpenStatic, adLockOptimistic
      If PicRs.RecordCount <> 0 Then
         BytesVal = PicRs.Fields("ibf13").Value
'         ReDim BytesS(Val(BytesVal))
'         BytesS() = PicRs.Fields("ibf14").GetChunk(Val(BytesVal))
         'Add By Sindy 2017/8/10 下載檔案
         p_FileName = App.path & "\TempFile"
         RidFile p_FileName
         If "" & PicRs.Fields("IBF15") <> "" Then
            Call PUB_GetFtpFile(PicRs.Fields("IBF15"), p_FileName, UCase("ImgByteFile"))
         End If
         '2017/8/10 END
         PicRs.AddNew
         PicRs.Fields("ibf07").Value = strUserNum
         PicRs.Fields("ibf08").Value = Val(strSrvDate(1))
         PicRs.Fields("ibf09").Value = Val(Format(time, "HHMM"))
         PicRs.Fields("ibf01").Value = oCP01
         PicRs.Fields("ibf02").Value = oCP02
         PicRs.Fields("ibf03").Value = oCP03
         PicRs.Fields("ibf04").Value = oCP04
         PicRs.Fields("ibf05").Value = "1"
         PicRs.Fields("ibf06").Value = "6"
         PicRs.Fields("ibf13").Value = BytesVal
'         PicRs.Fields("ibf14").Value = Null
'         PicRs.Fields("ibf14").AppendChunk BytesS()
         'Modify By Sindy 2017/8/10
         '檔案改放FTP
         If FileExists(p_FileName) Then
            PUB_PutFtpFile p_FileName, oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04 & "-1", oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04 & "-1", strFtpPath, UCase("imgbytefile")
            If strFtpPath <> "" Then
               PicRs.Fields("ibf15") = strFtpPath
            End If
         End If
         '2017/8/10 END
         PicRs.UPDATE
      End If
      IsSave = True
      StrMenu
   End If
End Sub
'2014/4/8 end

Private Sub Cmd1_Click(Index As Integer)
    Select Case Index
        Case 0 '搜尋
            If txtSystem & txtCode(0) = MsgText(601) Then Exit Sub
            Call GetCopyPhoto
        Case 1 '完成
            If G_SeekPicColor2.Picture = 0 Then Exit Sub
            InitAll
            Set G_SeekPicColor.Picture = G_SeekPicColor2.Picture
            Set tmpImg.Picture = G_SeekPicColor2.Picture
            IsSave = False
            Cmd1_Click (2)
        Case 2 '關閉
            Frame3.Visible = False
            InitAll2
            MainBTEnabled (False)
    End Select
End Sub

'Private Type RGBQUAD
'    B As Byte
'    G As Byte
'    R As Byte
'    A As Byte
'End Type
'
'Private Type SAFEARRAYBOUND
'    cElements As Long
'    lLbound   As Long
'End Type
'
'Private Type SAFEARRAY2D
'    cDims      As Integer
'    fFeatures  As Integer
'    cbElements As Long
'    cLocks     As Long
'    pvData     As Long
'    Bounds(1)  As SAFEARRAYBOUND
'End Type
'Private Type PALETTEENTRY
'    peR     As Byte
'    peG     As Byte
'    peB     As Byte
'    peFlags As Byte
'End Type
'
'Private Type LOGPALETTE002
'    palVersion       As Integer
'    palNumEntries    As Integer
'    palPalEntry(1)   As PALETTEENTRY
'End Type
'Private m_hPal     As Long
'Private m_tPal()   As RGBQUAD
'Private logPal002  As LOGPALETTE002
'Private m_Pow2(31) As Long
'Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As Any) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
'Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As RGBQUAD) As Long
'Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Dim OriginalTable(1 To 256) As RGBQUAD
'Dim GrayTable(1 To 256) As RGBQUAD
'Dim RedTable(1 To 256) As RGBQUAD
'Dim BlueTable(1 To 256) As RGBQUAD
'Dim GreenTable(1 To 256) As RGBQUAD
'Dim InvertTable(1 To 256) As RGBQUAD

'Add By Sindy 2021/1/29
Private Function ButtonChkLimit() As Boolean
   ButtonChkLimit = True
   '代表圖之本所案號匯入及貼上功能權限修改:
   '操作人員為SXX部門及M71、W10部門人員不可操作
   'Add By Sindy 2021/2/3 排除查名單
   'Add By Sindy 2021/2/4 排除接洽單
   'Modify by Amy 2022/12/30 +frm090801_New
   If Len(m_TMQ) = 0 And UCase(TypeName(Me.UpForm)) <> UCase("frm090801") And UCase(TypeName(Me.UpForm)) <> UCase("frm090801_New") Then
   '2021/2/3 END
      'Modify By Sindy 2024/4/18 W10不可操作改為W部門僅W20可操作，其他W部門都不可操作。
'      If Mid(Pub_StrUserSt03, 1, 1) = "S" Or _
'         Pub_StrUserSt03 = "M71" Or _
'         Pub_StrUserSt03 = "W10" Then
      If Mid(Pub_StrUserSt03, 1, 1) = "S" Or _
         Pub_StrUserSt03 = "M71" Or _
         Pub_StrUserSt03 = "W20" Then
         ButtonChkLimit = False
         MsgBox "權限不足，不可操作！", vbInformation
      End If
   End If
End Function

Public Sub cmdok_Click(Index As Integer)
    Select Case Index
    'Modify by Amy 2018/07/02 按鈕程式寫至Sub
    Case 0 '貼上
        If ButtonChkLimit = False Then cmdOK(0).Enabled = False: Exit Sub 'Add By Sindy 2021/1/29
        Screen.MousePointer = vbHourglass
        Call PhotoPaste
        Screen.MousePointer = vbDefault
    Case 1 '選擇檔案
        Call OpenAtt
    Case 2 '存檔
        Call PhotoSave
    Case 3 '繼續
        Call PContinue
    'Memo 2018/07/02 黑白掃描/選擇裝置/灰階掃描 鈕 拿掉不用-薛經理
'    Case 4 '黑白掃描(D)
'        Call GeneralScan
'    'add by nickc 2005/12/16
'    Case 5 '選擇裝置(H)
'         Call PopupSelectSourceDialog
'         Label11.Caption = "專利代表圖--" & GetDefTwainDev
'    Case 6 '灰階掃描(A)
'        Call GrayScan
    Case 7    'edit by nickc 2007/07/25 改複製
        Call PhotoCopy
    'Add by Amy 2018/07/02
    Case 8 '本所案號代表圖複製
        If ButtonChkLimit = False Then cmdOK(8).Enabled = False: Exit Sub 'Add By Sindy 2021/1/29
        Frame3.Visible = True
        MainBTEnabled (True)
        txtSystem = "": txtCode(0) = "": txtCode(1) = "": txtCode(2) = ""
        txtSystem.SetFocus
        txtSystem_GotFocus
    Case 9 '彩色/灰階切換
        Call ShowBt
        
   'Added by Morgan 2018/8/24
   Case 10 '刪除
      If LBL1(7).Caption = "" Then
         MsgBox "無代表圖可刪！", vbExclamation
      ElseIf MsgBox("是否確定要刪除？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
         Screen.MousePointer = vbHourglass
         If FormDelete() = True Then
            Screen.MousePointer = vbDefault
            Unload Me
         Else
            MsgBox "刪除失敗！", vbCritical
            Screen.MousePointer = vbDefault
         End If
      End If
      
    Case Else
    End Select
    'end 2018/07/02
End Sub

'Added by Lydia 2017/12/01
Private Sub Form_Activate()
Dim intD As Double
  '因為表單.ScaleMode從1-Twips 改成 3-像素,造成有些表單只顯示部份按鈕的位置計算錯誤
  'Modify by Amy 2018/07/02 改判斷按鈕caption,因下列位移只需查名使用
  'If Me.cmdOK(4).Visible = False And Me.cmdOK(5).Visible = False And Me.cmdOK(6).Visible = False Then
  If Me.cmdOK(3).Caption <> "繼續(&X)" Then
     intD = Me.cmdOK(0).Left
     Me.cmdOK(0).Left = 0
     Me.cmdOK(1).Left = Me.cmdOK(1).Left - intD
     Me.cmdOK(2).Left = Me.cmdOK(2).Left - intD
     Me.cmdOK(3).Left = Me.cmdOK(3).Left - intD
  End If
  
End Sub

'Private Sub DIBDither_ProgressEnd()
'   Repaint
'End Sub
'
'Private Sub DIBFilter_ProgressEnd()
'Repaint
'End Sub

Private Sub Form_Load()
frmpic002.Show
frmpic002.ZOrder 0
'Dim GpInput As mGDIpEx.GdiplusStartupInput
'    GpInput.GdiplusVersion = 1
'    If (mGDIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
'        cmdok(0).Enabled = False
'        cmdok(1).Enabled = False
'        cmdok(2).Enabled = False
'        Call MsgBox("沒有安裝 GDI+，請通知電腦中心安裝！", vbCritical)
'    Else
         Pub_Can_Copy_Pic = True
'         cmdok(0).Enabled = False
'         InitAll
'    End If
    cmdOK(4).Visible = False
    cmdOK(4).Enabled = False
    cmdOK(5).Visible = False
    cmdOK(5).Enabled = False
    cmdOK(6).Visible = False
    cmdOK(6).Enabled = False
   
    'Modify by Amy 2018/07/02 案件代表圖才出現 本所案號代表圖複製鈕
    Frame3.Visible = False
    cmdOK(8).Visible = False
    cmdOK(9).Visible = False
    cmdOK(9).Enabled = False
    '原設定SeekCmdok()寫至SetSeekCmdok
    'end 2018/07/02
   MoveFormToCenter Me
   Make
   '鎖定高跟寬
   'pic1.Height = 220 '3300
   'pic1.Width = 280 '4200
'   PicMain.Height = 220 '3300
'   PicMain.Width = 280 '4200
   'tmpPic.Height = 220 * 1.9
   'tmpPic.Width = 280 * 1.9
   tmpPic.AutoRedraw = True
   pic1.AutoRedraw = True
   pic1.ScaleMode = 3
   pic1.BorderStyle = 0
   pic1.Height = tmpPic.Height
   pic1.Width = tmpPic.Width
   IsSave = True
   
   If m_ibf05 = "" Then m_ibf05 = "1" 'Add By Sindy 2012/6/14
   
'   InitGrip
'    Set DIBDither = New cDIBDither
'    Set DIBFilter = New cDIBFilter
'    Set DIBPal = New cDIBPal
'    Set DIBSave = New cDIBSave
'    Set DIB = New cDIB
'    Set PicFrame = New cFrame

   'add by sonia 2014/4/8 M51才出現
   If strWorkType <> "1" And Pub_StrUserSt03 = "M51" Then
      Chk1.Visible = True
      Chk1.Enabled = True
      'Added by Morgan 2018/8/24 代表圖加刪除功能
      If Len(m_TMQ) = 0 Then
         cmdOK(10).Visible = True
         cmdOK(10).Enabled = True
      End If
      'end 2018/8/24
   End If
   '2014/4/8 end
   
   'Added by Lydia 2015/11/06 查名單電子化-記錄檔檔路徑
   'Remove by Lydia 2016/05/26
   'If m_TMQ <> "" Then
   '   strLoadPath = UpForm.strLoadPath
   '   If strLoadPath = "" Then
   '        strLoadPath = PUB_Getdesktop
   '   End If
   'End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

'Modified by Lydia 2016/03/23 查名單電子化,回傳圖片
'If Len(m_TMQ) > 0 And IsSave = True Then
If Len(m_TMQ) > 0 And IsSave = True And bolTMQUpd = True Then
    'Modified by Lydia 2016/03/28 統一調整查名單輸入作業時,圖片顯示符合畫面大小
    'Set oPic.Picture = G_SeekPicColor.Picture
    'Set oImg.Picture = tmpImg.Picture
    'oImg.Height = tmpImg.Height
    'oImg.Width = tmpImg.Width
    ''圖片查詢-置中(與共同表單大小相仿)
    'If m_TMQ = TMQ_AkindPic Then
    '    oImg.Top = tmpImg.Top
    '    oImg.Left = tmpImg.Left
    'End If
    Dim t_hd As Double
    Dim t_wd As Double
    t_hd = G_SeekPicColor.ScaleHeight / oPic.ScaleHeight
    t_wd = G_SeekPicColor.ScaleWidth / oPic.ScaleWidth
    If t_hd > t_wd Then
        t_wd = G_SeekPicColor.ScaleWidth / t_hd
        t_hd = G_SeekPicColor.ScaleHeight / t_hd
    Else
        t_hd = G_SeekPicColor.ScaleHeight / t_wd
        t_wd = G_SeekPicColor.ScaleWidth / t_wd
    End If
    oImg.Width = t_wd
    oImg.Height = t_hd
    oImg.Move (oPic.ScaleWidth - oImg.Width) / 2, (oPic.ScaleHeight - oImg.Height) / 2, t_wd, t_hd
    
    Set oImg.Picture = G_SeekPicColor.Picture
    
    'UpForm.IsWmf = IsWmf  'Mark by Lydia 2024/10/16 經過確認，不需回傳
    Unload Me
End If
'end 2016/03/28

Pub_Can_Copy_Pic = False
'Call mGDIpEx.GdiplusShutdown(m_GDIpToken)
'add by nickc 2006/05/24
Unload frmpic002
Set frmpic002 = Nothing
Set m_Image = Nothing
Set frmPic001 = Nothing
End Sub

'允許掃描
Sub CanScan()
'   cmdok(4).Visible = True
'   cmdok(6).Visible = True
'   'add by nickc 2005/12/16
'   If CheckAnyTwainDev = True Then
'      cmdok(4).Enabled = True
'      cmdok(6).Enabled = True
'      If GetTwainCounts > 1 Then
'         cmdok(5).Visible = True
'         cmdok(5).Enabled = True
'      End If
'      Label11.Caption = "專利代表圖--" & GetDefTwainDev
'   Else
'      cmdok(4).Enabled = False
'      cmdok(5).Visible = False
'      cmdok(5).Enabled = False
'      cmdok(6).Enabled = False
'   End If
End Sub

Sub StrMenu_Old()
'Mark by Amy 2018/07/02 增加彩色代表圖
'Dim bSuccess As Boolean
'
'     'Add By Sindy 2012/6/14
'     'Label13.Visible = False
'     If strWorkType = "1" Then
'        Label1(0) = "員工代號："
'        m_ibf05 = "3"
'        Frame2.Visible = False '黑色或彩色
'        If bolQuery = True Then
'           cmdok(1).Enabled = False '選擇檔案
'           cmdok(2).Enabled = False '存檔
'        End If
'        Label13.Caption = "備註：影像尺寸請縮小至480(W)*640(H)以下　　　檔案大小限 50 KB"
'        'Label13.Visible = True
'        'Add By Sindy 2013/7/2
'        '案件名稱
'        Label1(3).Visible = False
'        lbl1(8).Visible = False
'        '2013/7/2 END
'     Else
'        Label1(0) = "本所案號："
'        m_ibf05 = "1"
'        If InStr(oCP01, "T") > 0 Then
'           Label11.Caption = "商標代表圖"
'        Else
'           Label11.Caption = "專利代表圖"
'        End If
'        Label13.Caption = "備註：檔案大小限 300 KB"
'        'Add By Sindy 2013/7/2
'        '案件名稱
'        Label1(3).Visible = True
'        lbl1(8).Visible = True
'        '2013/7/2 END
'     End If
'     '2012/6/14 End
'     If Trim(oCP01 & oCP02 & oCP03 & oCP04) = "" Then
'        'Modify By Sindy 2012/6/14
'        If strWorkType = "1" Then
'           lbl1(0).Caption = "錯誤員工代號！"
'        Else
'        '2012/6/14 End
'           lbl1(0).Caption = "錯誤本所案號！"
'        End If
'        cmdok(1).Enabled = False
'        cmdok(2).Enabled = False
'     Else
'        'Modify By Sindy 2012/6/14
'        If strWorkType = "1" Then
'           lbl1(0).Caption = oCP02 & "  " & GetPrjSalesNM(oCP02)
'        Else
'        '2012/6/14 End
'           lbl1(0).Caption = oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04
'           'Add By Sindy 2013/7/2
'           '案件名稱
'           lbl1(8).Caption = GetPrjName(oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04)
'           '2013/7/2 END
'        End If
'     End If
'
'   Screen.MousePointer = vbHourglass
'   DoEvents
'   'InitAll
'   'Set pic1.Picture = LoadPicture()
'   Set PicRs = New ADODB.Recordset
'   PicRs.CursorLocation = adUseClient
'   'Modify By Sindy 2012/6/14 原ibf05='1'改為ibf05='" & m_ibf05 & "'
'   PicRs.Open "select ImgByteFile.*,S1.st02 as Cst02,s2.st02 as Ust02 from ImgByteFile,staff S1,staff S2 where ibf05='" & m_ibf05 & "' and ibf01='" & oCP01 & "' and ibf02='" & oCP02 & "' and ibf03='" & oCP03 & "' and ibf04='" & oCP04 & "' and ibf07=s1.st01(+) and ibf10=s2.st01(+) ", cnnConnection, adOpenStatic, adLockOptimistic
'   If PicRs.RecordCount <> 0 Then
'        PicRs.MoveFirst
'        lbl1(1).Caption = CheckStr(PicRs.Fields("ibf07")) & "  " & CheckStr(PicRs.Fields("Cst02"))
'        lbl1(2).Caption = ChangeWStringToTDateString(CheckStr(PicRs.Fields("ibf08").Value))
'        lbl1(3).Caption = Format(CheckStr(PicRs.Fields("ibf09")), "00:00")
'        lbl1(4).Caption = CheckStr(PicRs.Fields("ibf10")) & "  " & CheckStr(PicRs.Fields("Ust02"))
'        lbl1(5).Caption = ChangeWStringToTDateString(CheckStr(PicRs.Fields("ibf11").Value))
'        lbl1(6).Caption = Format(CheckStr(PicRs.Fields("ibf12")), "00:00")
'        lbl1(7).Caption = Format(CheckStr(PicRs.Fields("ibf13")), "###,###,###,###,##0") & " 位元組"
'        If CheckStr(PicRs.Fields("ibf06")) = "1" Or CheckStr(PicRs.Fields("ibf06")) = "3" Then
'            optColor(0).Value = True
'        Else
'            optColor(1).Value = True
'        End If
'        'edit by nickc 2007/11/29 加入無圖式的格式
'        'If CheckStr(PicRs.Fields("ibf06")) = "3" Or CheckStr(PicRs.Fields("ibf06")) = "4" Then IsWmf = True Else IsWmf = False
'        If CheckStr(PicRs.Fields("ibf06")) = "3" Or CheckStr(PicRs.Fields("ibf06")) = "4" Or CheckStr(PicRs.Fields("ibf06")) = "6" Then
'            IsWmf = True
'        Else
'            IsWmf = False
'        End If
'        'Add By Sindy 2017/8/10 下載檔案
''        If "" & PicRs.Fields("IBF15") <> "" Then
'            If IsWmf = False Then
'               Call PUB_GetFtpFile(PicRs.Fields("IBF15"), App.path & "\NowPic.jpg", UCase("ImgByteFile"))
'            Else
'               Call PUB_GetFtpFile(PicRs.Fields("IBF15"), App.path & "\NowPic.wmf", UCase("ImgByteFile"))
'            End If
''        Else
''        '2017/8/10 END
''            ReDim bytes(Val(PicRs.Fields("ibf13").Value))
''            bytes() = PicRs.Fields("ibf14").GetChunk(Val(PicRs.Fields("ibf13").Value))
''            file_num = FreeFile
''            If IsWmf = False Then
''                Open App.path & "\NowPic.jpg" For Binary Access Write As #file_num
''            Else
''                Open App.path & "\NowPic.wmf" For Binary Access Write As #file_num
''            End If
''            Put #file_num, , bytes()
''            Close #file_num
''        End If
'        'Frame2.Enabled = False
'        Frame2.Enabled = True 'Modify By Sindy 98/04/09
'        'pic1.Move -20, -20, PicMain.Width, PicMain.Height
'        If IsWmf = False Then
'            'G_SeekPicColor.Picture = pvGetStdPicture(Trim(App.Path & "\tmp.jpg"))
'            PicToObj Trim(App.path & "\NowPic.jpg")
'        Else
'            'G_SeekPicColor.Picture = pvGetStdPicture(Trim(App.Path & "\tmp.wmf"))
'            PicToObj Trim(App.path & "\NowPic.wmf")
'        End If
''        Call pvSetDIBPicture(G_SeekPicColor.Picture, optColor(1).Value)
''        ResizeImage True
'        If Dir(App.path & "\NowPic.jpg") <> "" Then
'            Kill App.path & "\NowPic.jpg"
'        End If
'        If Dir(App.path & "\tmp.tif") <> "" Then
'            Kill App.path & "\tmp.tif"
'        End If
'        If Dir(App.path & "\NowPic.wmf") <> "" Then
'            Kill App.path & "\NowPic.wmf"
'        End If
'    Else
''        pic1.Font.Size = 32
''        pic1.Font = "標楷體"
''        pic1.ForeColor = &HFF&
''        pic1.CurrentX = (pic1.ScaleWidth / 2) - (pic1.TextWidth("未設定圖片")) / 2
''        pic1.CurrentY = (pic1.ScaleHeight / 2) - (pic1.TextHeight("未設定圖片")) / 2
''        pic1.Print "未設定圖片"
''        Set G_SeekPicColor.Picture = pic1.Image
''        Set pic1.Picture = G_SeekPicColor.Picture
''
''        Set tmpImg.Picture = G_SeekPicColor.Picture
''        tmpImg.Move 0, 0, tmpPic.ScaleWidth, tmpPic.ScaleHeight
'           Set PicRs = New ADODB.Recordset
'           If PicRs.State = 1 Then PicRs.Close
'           PicRs.CursorLocation = adUseClient
'           'Modify By Sindy 2012/6/14 原ibf05='1'改為ibf05='" & m_ibf05 & "'
'           PicRs.Open "select ImgByteFile.* from ImgByteFile where ibf05='" & m_ibf05 & "' and ibf01='000' and ibf02='000000' and ibf03='0' and ibf04='00' ", cnnConnection, adOpenStatic, adLockOptimistic
'           If PicRs.RecordCount <> 0 Then
'                PicRs.MoveFirst
'                lbl1(1).Caption = ""
'                lbl1(2).Caption = ""
'                lbl1(3).Caption = ""
'                lbl1(4).Caption = ""
'                lbl1(5).Caption = ""
'                lbl1(6).Caption = ""
'                lbl1(7).Caption = ""
'                If CheckStr(PicRs.Fields("ibf06")) = "1" Or CheckStr(PicRs.Fields("ibf06")) = "3" Then optColor(0).Value = True Else optColor(1).Value = True
'                'edit by nickc 2007/11/29 加入無圖式的格式
'                'If CheckStr(PicRs.Fields("ibf06")) = "3" Or CheckStr(PicRs.Fields("ibf06")) = "4" Then IsWmf = True Else IsWmf = False
'               If CheckStr(PicRs.Fields("ibf06")) = "3" Or CheckStr(PicRs.Fields("ibf06")) = "4" Or CheckStr(PicRs.Fields("ibf06")) = "6" Then
'                    IsWmf = True
'               Else
'                    IsWmf = False
'               End If
'               'Add By Sindy 2017/8/10 下載檔案
''               If "" & PicRs.Fields("IBF15") <> "" Then
'                  If IsWmf = False Then
'                     Call PUB_GetFtpFile(PicRs.Fields("IBF15"), App.path & "\NowPic.jpg", UCase("ImgByteFile"))
'                  Else
'                     Call PUB_GetFtpFile(PicRs.Fields("IBF15"), App.path & "\NowPic.wmf", UCase("ImgByteFile"))
'                  End If
''               Else
''               '2017/8/10 END
''                  ReDim bytes(Val(PicRs.Fields("ibf13").Value))
''                  bytes() = PicRs.Fields("ibf14").GetChunk(Val(PicRs.Fields("ibf13").Value))
''                  file_num = FreeFile
''                  If IsWmf = False Then
''                      Open App.path & "\NowPic.jpg" For Binary Access Write As #file_num
''                  Else
''                      Open App.path & "\NowPic.wmf" For Binary Access Write As #file_num
''                  End If
''                  Put #file_num, , bytes()
''                  Close #file_num
''               End If
'                'Frame2.Enabled = False
'                Frame2.Enabled = True 'Modify By Sindy 98/04/09
'                If IsWmf = False Then
'                    PicToObj Trim(App.path & "\NowPic.jpg")
'                Else
'                    PicToObj Trim(App.path & "\NowPic.wmf")
'                End If
'
'                If Dir(App.path & "\NowPic.jpg") <> "" Then
'                    Kill App.path & "\NowPic.jpg"
'                End If
'                If Dir(App.path & "\tmp.tif") <> "" Then
'                    Kill App.path & "\tmp.tif"
'                End If
'                If Dir(App.path & "\NowPic.wmf") <> "" Then
'                    Kill App.path & "\NowPic.wmf"
'                End If
'            End If
'    End If
'    Screen.MousePointer = vbDefault
End Sub

Private Sub optColor_Click(Index As Integer)
Screen.MousePointer = vbHourglass
'ResizeImage
Screen.MousePointer = vbDefault
End Sub

'Private Sub pic1_Resize()
'Img1.Width = pic1.Width
'Img1.Height = pic1.Height
'Img1.Top = pic1.Top
'Img1.Left = pic1.Left
'End Sub

Private Sub Timer1_Timer()
'edit by nickc 2007/12/07 加入查名單可以貼圖
'If Trim(oCP01 & oCP02 & oCP03 & oCP04) <> "" And (Clipboard.GetFormat(2) = True Or Clipboard.GetFormat(3) = True) Then
If (Trim(oCP01 & oCP02 & oCP03 & oCP04) <> "" Or Width = 3800) And (Clipboard.GetFormat(2) = True Or Clipboard.GetFormat(3) = True) Then
   cmdOK(0).Enabled = True
'Add by Morgan 2010/9/17
ElseIf Not PUB_PicCopy Is Nothing Then
   cmdOK(0).Enabled = True
Else
   cmdOK(0).Enabled = False
End If
End Sub

Private Sub Make()
Dim x As Integer, g1 As Integer, g2 As Integer, g3 As Integer, r1 As Integer, r2 As Integer, r3 As Integer, b1 As Integer, b2 As Integer, b3 As Integer, MixRate As Integer, i As Integer
Dim j As Integer, DrawWidth As Integer
x = 0: g1 = 0: g2 = 153: g3 = 0: r1 = 0: r3 = 0: r2 = 54: b1 = 120: b3 = 120: b2 = 216: MixRate = 15

Picture1.Cls
DrawWidth = Picture1.Width / 255
For i = 0 To 255
    If r3 = r2 Then GoTo 30
    If r1 > r2 Then r3 = r3 - 1
    If r1 < r2 Then r3 = r3 + 1
30
    If g3 = g2 Then GoTo 20
    If g1 > g2 Then g3 = g3 - 1
    If g1 < g2 Then g3 = g3 + 1
20
    If b3 = b2 Then GoTo 10
    If b1 > b2 Then b3 = b3 - 1
    If b1 < b2 Then b3 = b3 + 1
10
    For j = 1 To DrawWidth
      Picture1.Line (x + j, 0)-(x + j, Picture1.Height), RGB(r3, g3, b3)
      If x + j = Picture1.Width Then Exit Sub
    Next j
    x = x + DrawWidth

Next i
    Picture1.Refresh

End Sub

'Private Sub pvSetDIBPicture(Image As StdPicture, IsColor As Boolean)
'
'    Static lstW As Long
'    Static lstH As Long
'
'    If (Not Image Is Nothing) Then
'        If Image <> 0 Then
'            lstW = DIB.Width
'            lstH = DIB.Height
'
'            Call DIBPal.Clear
'
'            DIBbpp = DIB.CreateFromStdPicture(Image, DIBPal, DIBDither) ', IsColor)
'
'            PicFrame.Clear
'            Repaint
'
'        End If
'    End If
'End Sub

Public Function IsAuthorized(ByVal strID As String, Optional ByVal strPWD As String, Optional ByVal bolChkPwd As Boolean = True) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, strSQLpwd As String
      
   IsAuthorized = False
   strSQLpwd = ""
   If bolChkPwd = True Then
      If (strPWD = "" Or strID = strPWD) Then
         MsgBox "密碼不可空白或與員工代號相同！", vbCritical
         Exit Function
      Else
         strSQLpwd = " and sp03='" & Encrypt(strPWD, True) & "'"
      End If
   End If
   
   strSql = "Select 1 as Msg From staff_pwd where sp03<>'" & Encrypt(strID, True) & "' and sp03 is not null" & _
      " and sp01='" & strID & "'" & strSQLpwd
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      IsAuthorized = True
   ElseIf (bolChkPwd = True) Then
      MsgBox "請輸入登入資料！"
   End If
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
End Function

Private Sub txtCode_GotFocus(Index As Integer)
    TextInverse txtCode(Index)
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPassword_GotFocus()
    TextInverse txtPassword
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_GotFocus()
    txtSystem.SetFocus
    TextInverse txtSystem
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUserNo_GotFocus()
   TextInverse txtUserNo
End Sub

Private Function getUserName(strID As String) As String

   Dim strSql As String, rsQuery As New ADODB.Recordset
      
   strSql = "Select ST02 From STAFF where ST01='" & strID & "'"
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      getUserName = "" & rsQuery.Fields(0)
   Else
      getUserName = ""
   End If
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
End Function

'Add By Sindy 2010/11/25
Private Sub txtUserNo_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUserNo_LostFocus()

   lblUserName = getUserName(txtUserNo)
   
End Sub

'Public Sub ResizeImage(Optional IsSeek As Boolean = False)
'
'    Dim t_wd As Double
'    Dim t_hd As Double
'    Dim tmpPB As PictureBox
'
'If G_SeekPicColor.Picture = 0 Then Exit Sub
'    If optColor(1).Value = True Or IsWmf = True Then
'        t_hd = G_SeekPicColor.ScaleHeight / PicMain.ScaleHeight
'        t_wd = G_SeekPicColor.ScaleWidth / PicMain.ScaleWidth
'        If t_hd > t_wd Then
'            t_wd = G_SeekPicColor.ScaleWidth / t_hd
'            t_hd = G_SeekPicColor.ScaleHeight / t_hd
'        Else
'            t_hd = G_SeekPicColor.ScaleHeight / t_wd
'            t_wd = G_SeekPicColor.ScaleWidth / t_wd
'        End If
'        Set tmpImg.Picture = G_SeekPicColor.Picture
'        tmpImg.Move (PicMain.ScaleWidth - t_wd) / 2, (PicMain.ScaleHeight - t_hd) / 2, t_wd, t_hd
'    Else
'        t_hd = G_SeekPicColor.ScaleHeight / PicMain.ScaleHeight
'        t_wd = G_SeekPicColor.ScaleWidth / PicMain.ScaleWidth
'        If t_hd > t_wd Then
'            t_wd = G_SeekPicColor.ScaleWidth / t_hd
'            t_hd = G_SeekPicColor.ScaleHeight / t_hd
'        Else
'            t_hd = G_SeekPicColor.ScaleHeight / t_wd
'            t_wd = G_SeekPicColor.ScaleWidth / t_wd
'        End If
''        Repaint
''        DrawGrayBitmap tmpPic.hDC, 0, 0, PicMain.ScaleWidth, PicMain.ScaleHeight
'        tmpPic.Refresh
'        tmpImg.Visible = False
'        tmpImg.Width = tmpPic.Width
'        tmpImg.Height = tmpPic.Height
'        tmpImg.Top = 0
'        tmpImg.Left = 0
'        Set tmpImg.Picture = tmpPic.Image
'        tmpImg.Visible = True
'        t_wd = PicMain.Width
'        t_hd = PicMain.Height
'        Set tmpPic.Picture = LoadPicture()
'        tmpImg.Move (tmpPic.Width - tmpImg.Width) / 2, (tmpPic.Height - tmpImg.Height) / 2, t_wd, t_hd
'    End If
'
'End Sub
'
''add by nickc 2005/12/21 轉黑白
'Function ChgPicToBW(oImage As StdPicture) As StdPicture
'Exit Function
'Dim bm As BITMAP
'Dim lhOldBmp As Long
'Dim lHDC As Long
'If IsWmf = False Then
'    Screen.MousePointer = vbHourglass
'    Set G_SeekPicBW.Picture = G_SeekPicColor.Picture
'    If optColor(0).Enabled = True Then
''        DrawGrayBitmap G_SeekPicBW.hDC, 0, 0, (G_SeekPicBW.ScaleWidth / Screen.TwipsPerPixelX) + 2, (G_SeekPicBW.ScaleHeight / Screen.TwipsPerPixelY) + 2
'    End If
'    Screen.MousePointer = vbDefault
'Else
'    Set G_SeekPicBW.Picture = G_SeekPicColor.Picture
'End If
'End Function

''=============Sizing grip staff==============
'Private Sub InitGrip()
'    Dim i As Integer
'    lblGrip(0).Width = GripSize
'    lblGrip(0).Height = GripSize
'    For i = 1 To 7
'        Load lblGrip(i)
'        lblGrip(i).MousePointer = i + 4 * Int((9 - i) / 4)
'    Next i
'    lblGrip(0).MousePointer = 8
'    ShowGrip False
'End Sub
'
'Private Sub lblGrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = vbLeftButton Then
'        bSizing = True
'        xStart = (x / Screen.TwipsPerPixelX): yStart = (y / Screen.TwipsPerPixelY)
'        lblShape.Enabled = False
'    End If
'End Sub
'
'Private Sub lblGrip_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim lft As Long, tp As Long, wdt As Long, hgt As Long
'    If bSizing And Button = 1 Then
'        Select Case Index
'            Case 0
'                lft = lblShape.Left + x - xStart
'                tp = lblShape.Top + y - yStart
'                wdt = lblShape.Width - x + xStart
'                hgt = lblShape.Height - y + yStart
'            Case 1
'                lft = lblShape.Left + x - xStart
'                tp = lblShape.Top
'                wdt = lblShape.Width - x + xStart
'                hgt = lblShape.Height
'            Case 2
'                lft = lblShape.Left + x - xStart
'                tp = lblShape.Top
'                wdt = lblShape.Width - x + xStart
'                hgt = lblShape.Height + y - yStart
'            Case 3
'                lft = lblShape.Left
'                tp = lblShape.Top
'                wdt = lblShape.Width
'                hgt = lblShape.Height + y - yStart
'            Case 4
'                lft = lblShape.Left
'                tp = lblShape.Top
'                wdt = lblShape.Width + x - xStart
'                hgt = lblShape.Height + y - yStart
'            Case 5
'                lft = lblShape.Left
'                tp = lblShape.Top
'                wdt = lblShape.Width + x - xStart
'                hgt = lblShape.Height
'            Case 6
'                lft = lblShape.Left
'                tp = lblShape.Top + y - yStart
'                wdt = lblShape.Width + x - xStart
'                hgt = lblShape.Height - y + yStart
'            Case 7
'                lft = lblShape.Left
'                tp = lblShape.Top + y - yStart
'                wdt = lblShape.Width
'                hgt = lblShape.Height - y + yStart
'        End Select
'        If wdt < 0 Or hgt < 0 Or lft < 0 Or tp < 0 Or lft + wdt > pic1.Width Or tp + hgt > pic1.Height Then Exit Sub
'        lblShape.Move lft, tp, wdt, hgt
'        MoveGrips
'    End If
'End Sub
'
'Private Sub lblGrip_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'    bSizing = False
'    lblShape.Enabled = True
'End Sub
'
'Private Sub lblShape_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = vbLeftButton Then
'        bMoving = True
'        xStart = x: yStart = y
'        lblShape.MousePointer = 5
'    End If
'End Sub
'
'Private Sub lblShape_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Dim lft As Long, tp As Long
'    If bMoving Then
'        lft = lblShape.Left + x - xStart
'        tp = lblShape.Top + y - yStart
'        If lft <= 0 Then lft = 0
'        If tp <= 0 Then tp = 0
'        If lft > pic1.Width - lblShape.Width Then lft = pic1.Width - lblShape.Width
'        If tp > pic1.Height - lblShape.Height Then tp = pic1.Height - lblShape.Height
'        lblShape.Move lft, tp
'        MoveGrips
'    End If
'End Sub
'
'Private Sub lblShape_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    bMoving = False
'    lblShape.MousePointer = 0
'End Sub
'
'Private Sub MoveGrips()
'    XGrip(0) = lblShape.Left - GripSize
'    XGrip(1) = lblShape.Left + lblShape.Width / 2 - GripSize / 2
'    XGrip(2) = lblShape.Left + lblShape.Width
'    YGrip(0) = lblShape.Top - GripSize
'    YGrip(1) = lblShape.Top + lblShape.Height / 2 - GripSize / 2
'    YGrip(2) = lblShape.Top + lblShape.Height
'    lblGrip(0).Move XGrip(0), YGrip(0)
'    lblGrip(1).Move XGrip(0), YGrip(1)
'    lblGrip(2).Move XGrip(0), YGrip(2)
'    lblGrip(3).Move XGrip(1), YGrip(2)
'    lblGrip(4).Move XGrip(2), YGrip(2)
'    lblGrip(5).Move XGrip(2), YGrip(1)
'    lblGrip(6).Move XGrip(2), YGrip(0)
'    lblGrip(7).Move XGrip(1), YGrip(0)
'End Sub
'
'Private Sub ShowGrip(bShow As Boolean)
'    Dim i As Integer
'    lblShape.Move (pic1.ScaleWidth - (pic1.ScaleWidth * 2 / 3)) / 2, (pic1.ScaleHeight - (pic1.ScaleHeight * 2 / 3)) / 2, (pic1.ScaleWidth * 2 / 3), (pic1.ScaleHeight * 2 / 3)
'    lblShape.Visible = bShow
'    For i = 0 To 7
'        lblGrip(i).Visible = bShow
'    Next i
'    MoveGrips
'End Sub
'
'Public Sub Crop(ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long)
'    Dim sPic As StdPicture
'    Dim oldPicW As Long
'    Dim oldPicH As Long
'    Dim pi As Double
'    Dim pX As Double
'    Dim pY As Double
'    Dim pdx As Double
'    Dim pdy As Double
'    Dim pj As Double
'    Dim bSuccess As Boolean
'    CropPic.Picture = LoadPicture()   '剪下來的圖暫存
'
'    Set sPic = G_SeekPicColor.Picture
'    pi = pic1.ScaleWidth / G_SeekPicColor.ScaleWidth
'    pj = pic1.ScaleHeight / G_SeekPicColor.ScaleHeight
'    pX = x / pi
'    pY = y / pj
'    pdx = dx / pi
'    pdy = dy / pj
'
'
'    CropPic.Width = pdx
'    CropPic.Height = pdy
'    CropPic = LoadPicture()
'    CropPic.PaintPicture G_SeekPicColor.Picture, 0, 0, pdx, pdy, pX, pY, pdx, pdy
'    '先存檔，再讀，因為要改變大小
'    SavePicture CropPic.Image, App.Path & "\tmp.bmp"
'    G_SeekPicColor.Picture = LoadPicture(App.Path & "\tmp.bmp")
'    Set G_SeekPicBW = LoadPicture()
'    Call pvSetDIBPicture(G_SeekPicColor.Picture, True)
'    Call mGDIpEx.SaveDIB(DIB, App.Path & "\tmp.jpg", [ImageJPEG], 100, , 2)
'    Set G_SeekPicColor.Picture = pvGetStdPicture(App.Path & "\tmp.jpg", bSuccess)
'    If bSuccess Then
'        ResizeImage True
'    End If
'    If Dir(App.Path & "\tmp.bmp") <> "" Then
'        Kill App.Path & "\tmp.bmp"
'    End If
'    If Dir(App.Path & "\tmp.jpg") <> "" Then
'        Kill App.Path & "\tmp.jpg"
'    End If
'
'    Set sPic = Nothing
'End Sub

'Sub Repaint()
'
'    Dim xOff As Long, yOff As Long
'    Dim wDst As Double, hDst As Double
'    Dim xSrc As Long, ySrc As Long
'    Dim wSrc As Long, hSrc As Long
'
'    If (DIB.hDIB <> 0) Then
'        hDst = (G_SeekPicColor.ScaleHeight) / (PicMain.ScaleHeight)
'        wDst = (G_SeekPicColor.ScaleWidth) / (PicMain.ScaleWidth)
'        If hDst > wDst Then
'            wDst = (G_SeekPicColor.ScaleWidth) / hDst
'            hDst = (G_SeekPicColor.ScaleHeight) / hDst
'        Else
'            hDst = (G_SeekPicColor.ScaleHeight) / wDst
'            wDst = (G_SeekPicColor.ScaleWidth) / wDst
'        End If
'        xOff = (PicMain.ScaleWidth - wDst) / 2
'        xSrc = 0
'        wSrc = DIB.Width
'        yOff = (PicMain.ScaleHeight - hDst) / 2
'        ySrc = 0
'        hSrc = DIB.Height
'        Call DIB.Stretch(tmpPic.hDC, xOff, yOff, wDst, hDst, xSrc, ySrc, wSrc, hSrc)
'        Call PicFrame.PaintToDC(tmpPic.hDC)
'        tmpPic.Refresh
'    End If
'End Sub

'Function GetRValue&(ByVal rgbColor&)
'    GetRValue = rgbColor And &HFF
'End Function
'
'Function GetGValue&(ByVal rgbColor&)
'    GetGValue = (rgbColor And &HFF00&) / &HFF&
'End Function
'
'Function GetBValue&(ByVal rgbColor&)
'    GetBValue = (rgbColor& And &HFF0000) / &HFF00&
'End Function
'
'Sub ChangetoGray(ByVal SrcDC&, _
'                 ByVal nX&, _
'                 ByVal ny&, _
'                 Optional ByVal nMaskColor& = -1)
'    Dim rgbColor&, Gray&
'    Dim RValue&, GValue&, BValue&
'    Dim dl&
'
'    rgbColor = GetPixel(SrcDC, nX, ny)
'
'    If rgbColor = nMaskColor Then GoTo Release:
'
'    RValue = GetRValue(rgbColor)
'    GValue = GetGValue(rgbColor)
'    BValue = GetBValue(rgbColor)
'
'    Gray = (9798 * RValue + 19235 * GValue + 3735 * BValue) / 32768 'Change wffs
'
'    rgbColor = RGB(Gray, Gray, Gray)
'
'    dl& = SetPixelV(SrcDC, nX, ny, rgbColor)
'
'Release:
'    rgbColor = 0: Gray = 0
'    RValue = 0: GValue = 0: BValue = 0
'    dl = 0
'End Sub


'Sub DrawGrayBitmap(ByVal hDC&, _
'                   ByVal nX&, _
'                   ByVal ny&, _
'                   ByVal nWidth&, _
'                   ByVal nHeight&, _
'                   Optional ByVal nMaskColor& = -1)
'    Dim i&, j&
'
'    For i = nX To nWidth
'        For j = ny To nHeight
'            ChangetoGray hDC, i, j, nMaskColor
'            DoEvents
'        Next j
'    Next i
'End Sub

'Private Sub pvSetPalMode(ByVal bpp As Long)
'
'
'    Select Case bpp
'        Case 1                                             '-- 2 colors / Black and White
'            lIdxNew = IIf(DIBPal.IsGreyScale, 0, 4)
'        Case 4                                             '-- 16 colors / 16 greys
'            lIdxNew = IIf(DIBPal.IsGreyScale, 1, 5)
'        Case 8                                             '-- 256 colors / 256 greys
'            lIdxNew = IIf(DIBPal.IsGreyScale, 2, 6)
'        Case 24                                            '-- True color
'            lIdxNew = 8
'        Case Else
'            Exit Sub
'    End Select
'
'End Sub

Sub InitAll()
'    DIB.Destroy
'    DIBPal.Clear
'    PicFrame.Clear
    Set tmpPic.Picture = LoadPicture()
    Set tmpImg.Picture = LoadPicture()
    Set G_SeekPicColor.Picture = LoadPicture()
'    Set G_SeekPicBW.Picture = LoadPicture()
'    Set Me.CropPic.Picture = LoadPicture()
'    Set PicMain.Picture = LoadPicture()
    Set pic1.Picture = LoadPicture()
End Sub

'Modify by Amy 2018/07/02 +bolGrayScale: 轉灰階
Sub PicToObj(oFileNameAndPath As String, Optional ByVal bolGrayScale As Boolean = False)
    Dim objImg As StdPicture
    Dim nSrcWidth, nSrcHeight, nWidth, nHeight
    Dim tBI      As BITMAP
    InitAll
    On Error GoTo BE
    If UCase(pvGetExt(oFileNameAndPath)) = "WMF" Or UCase(pvGetExt(oFileNameAndPath)) = "EMF" Then
        Set G_SeekPicColor.Picture = LoadPicture(App.path & "\NowPic.wmf")
        'tmpImg.BackColor = &H80000009
    Else
        'tmpPic.BackColor = &H8000000A
        Set objImg = pvGetStdPicture(oFileNameAndPath)
        Set m_Image = New cImage
        Set m_Jpeg = New cJpeg
        Set G_SeekPicColor.Picture = objImg
        If G_SeekPicColor.Picture <> 0 Then
            Call GetObject(objImg.handle, Len(tBI), tBI)
            If tBI.bmWidth = 0 Or tBI.bmHeight = 0 Then
                MsgBox "發生錯誤！", vbExclamation, "圖檔格式錯誤"
                'edit by nickc 2007/11/19
                If oRtPic = False Then
                    StrMenu
                End If
                Exit Sub
            End If
            If tBI.bmWidth > 2000 Or tBI.bmHeight > 2000 Then
                nSrcWidth = tBI.bmWidth
                nSrcHeight = tBI.bmHeight
                If nSrcWidth > nSrcHeight Then
                    nWidth = 1200
                    nHeight = nSrcHeight / (nSrcWidth / nWidth)
                ElseIf nSrcWidth < nSrcHeight Then
                    nHeight = 1200
                    nWidth = nSrcWidth / (nSrcHeight / nHeight)
                Else
                    nHeight = 1200
                    nWidth = 1200
                End If
                pic1.Width = nWidth
                pic1.Height = nHeight
                pic1.BackColor = &H8000000A
                '重新定義大小
                pic1.Scale (0, 0)-(nWidth, nHeight)
                '縮小
                pic1.PaintPicture objImg, 0, 0, nWidth, nHeight, , , , , vbSrcCopy
                '存檔
                SavePicture pic1.Image, App.path & "\NowPic.bmp"
                Set objImg = pvGetStdPicture(App.path & "\NowPic.bmp")
            'Added by Lydia 2024/06/03 查名單-網中：圖片強制變大，網中系統要求提供最小解析度224x224，過小則無法AI相似度運算。
            ElseIf Len(m_TMQ) > 0 Then
               '參考智慧局的「以圖搜圖」建議圖片裁切為正方形尺寸 242px * 242px 以上-->先以智慧局的規格來管制
                       '另外DPI設定的圖片5cmX 5cm, 在224 dpi X 224 dpi, 所需像素441 x 441;在300 dpi X 300 dpi, 所需像素591 x 591
                'Modified by Lydia 2025/04/24 在網站輸入查名發現圖片的高度和寬度都要符合224 x 224
                'If tBI.bmWidth < 242 And tBI.bmHeight < 242 Then
                If tBI.bmWidth < 242 Or tBI.bmHeight < 242 Then
                   nSrcWidth = tBI.bmWidth
                   nSrcHeight = tBI.bmHeight
                   If nSrcWidth < nSrcHeight Then
                      nWidth = 242
                      nHeight = nSrcHeight / (nSrcWidth / nWidth)
                   ElseIf nSrcWidth > nSrcHeight Then
                      nHeight = 242
                      nWidth = nSrcWidth / (nSrcHeight / nHeight)
                   Else
                      nHeight = 242
                      nWidth = 242
                   End If
                   pic1.Width = nWidth
                   pic1.Height = nHeight
                   pic1.BackColor = &H8000000A
                   '重新定義大小
                   pic1.Scale (0, 0)-(nWidth, nHeight)
                   '縮小
                   pic1.PaintPicture objImg, 0, 0, nWidth, nHeight, , , , , vbSrcCopy
                   '存檔
                   SavePicture pic1.Image, App.path & "\NowPic.bmp"
                   Set objImg = pvGetStdPicture(App.path & "\NowPic.bmp")
                End If
            'end 2024/06/03
            End If
            m_Image.CopyStdPicture objImg
            'Added by Lydia 2015/06/24 查名單電子化:查名單的圖片用彩色
            'Modify by Amy 2018/07/02 +bolGrayScale 彩圖存檔後才再轉灰階
            If strWorkType <> "1" And Len(m_TMQ) = 0 And bolGrayScale = True Then 'Add By Sindy 2012/6/19 +if
               m_Jpeg.SetSamplingFrequencies 1, 1, 0, 0, 0, 0 '轉灰階
            End If
            m_Jpeg.Quality = 75
            m_Jpeg.SampleHDC m_Image.hDC, m_Image.Width, m_Image.Height
            RidFile App.path & "\NowPic.jpg"
            m_Jpeg.SaveFile App.path & "\NowPic.jpg"
            m_Image.CopyStdPicture pvGetStdPicture(App.path & "\NowPic.jpg")
            Set G_SeekPicColor.Picture = LoadPicture(App.path & "\NowPic.jpg")

        End If
    End If
    Dim t_hd As Double
    Dim t_wd As Double
    t_hd = G_SeekPicColor.ScaleHeight / tmpPic.ScaleHeight
    t_wd = G_SeekPicColor.ScaleWidth / tmpPic.ScaleWidth
    If t_hd > t_wd Then
        t_wd = G_SeekPicColor.ScaleWidth / t_hd
        t_hd = G_SeekPicColor.ScaleHeight / t_hd
    Else
        t_hd = G_SeekPicColor.ScaleHeight / t_wd
        t_wd = G_SeekPicColor.ScaleWidth / t_wd
    End If
        tmpImg.Width = t_wd
        tmpImg.Height = t_hd
        tmpImg.Move (tmpPic.ScaleWidth - tmpImg.Width) / 2, (tmpPic.ScaleHeight - tmpImg.Height) / 2, t_wd, t_hd
    
    Set tmpImg.Picture = G_SeekPicColor.Picture
    Set objImg = Nothing

    Exit Sub
BE:
    Resume Next
End Sub

'Add by Amy 2018/07/02 各按鈕拆出寫成Sub
Private Sub PhotoPaste()
    Dim objImg2 As StdPicture
    Dim tBI2      As BITMAP
    
On Error GoTo 0
On Error GoTo ErrHand
       bolTMQUpd = False 'Added by Lydia 2015/11/06
      'edit by nickc 2007/07/25 用貼上的，一率不以WMF 格式存
      'If Clipboard.GetFormat(vbCFBitmap) And Clipboard.GetFormat(vbCFMetafile) And Clipboard.GetFormat(vbCFDIB) Then
      '  IsWmf = True
      'Else
        IsWmf = False
      'End If
      Set G_PicTemp = Clipboard.GetData()
      Clipboard.Clear
      'add by nickc 2006/05/08
      
      'Add by Morgan 2010/9/17 加判斷系統內複製的圖
      If G_PicTemp <> 0 Then
         Set PUB_PicCopy = G_PicTemp
      ElseIf Not PUB_PicCopy Is Nothing Then
         Set G_PicTemp = PUB_PicCopy
      End If
      'end 2010/9/17
      
      If G_PicTemp <> 0 Then
         SavePicture G_PicTemp, App.path & "\tmp.bmp"
         Set objImg2 = pvGetStdPicture(App.path & "\tmp.bmp")
         Call GetObject(objImg2.handle, Len(tBI2), tBI2)
         If tBI2.bmWidth = 0 Or tBI2.bmWidth = 0 Then
            IsWmf = True
            SavePicture G_PicTemp, App.path & "\NowPic.wmf"
            PicToObj App.path & "\NowPic.wmf"
         Else
            PicToObj App.path & "\tmp.bmp"
         End If
        
         
         'InitAll
         'Set G_SeekPicColor.Picture = pvGetStdPicture(App.Path & "\tmp.bmp")
             optColor(0).Value = True
             Frame2.Enabled = True
             If Dir(App.path & "\tmp.bmp") <> "" Then
               Kill App.path & "\tmp.bmp"
             End If
'              Call pvSetDIBPicture(G_SeekPicColor.Picture, optColor(1).Value)
'             ResizeImage True

             'Remove by Morgan 2010/7/27 不用了
             'If Val(strUserNum) < 65001 Then
             '    Frame1.Visible = True
             '   txtUserNo.SetFocus
             'End If
             'end 2010/7/27
             
             IsSave = False
      Else
        'Modify by Amy 2017/01/20 +不秀訊息判斷
        If bolNoMsg = False Then MsgBox "不正確的圖片！", , "錯誤！"
        Screen.MousePointer = vbDefault
        Exit Sub
      End If
      'Modify by Amy 2017/01/20 +不秀訊息判斷
      If bolNoMsg = False Then MsgBox "貼圖成功", vbInformation, "恭喜！"
      Exit Sub

ErrHand:
    MsgBox "發生錯誤！", vbExclamation, "未知錯誤！"
    Screen.MousePointer = vbDefault
End Sub

Private Sub OpenAtt()
    Dim sFile 'Added by Lydia 2015/11/06 查名單電子化
    
On Error GoTo 0
On Error GoTo ErrHand
         bolTMQUpd = False 'Added by Lydia 2015/11/06
         cd1.FileName = ""
         'Modify by Amy 2018/07/02 .png/.tif/.tiff 加入會錯
         'cd1.Filter = "Supported files|*.bmp;*.gif;*.jpg;*.png;*.wmf;*.tif|Bitmap files (*.bmp)|*.bmp|GIF files (*.gif)|*.gif|JPEG files (*.jpg)|*.jpg|PNG files (*.png)|*.png|TIFF files (*.tif)|*.tif|WMF files (*.wmf)|*.wmf"
         cd1.Filter = "Supported files|*.bmp;*.gif;*.jpg;*.wmf|Bitmap files (*.bmp)|*.bmp|GIF files (*.gif)|*.gif|JPEG files (*.jpg)|*.jpg|WMF files (*.wmf)|*.wmf"
         'Added by Lydia 2015/11/06 查名單電子化:讀取前次路徑
         'Modified by Lydia 2016/05/30
         'If m_TMQ <> "" Then cd1.InitDir = strLoadPath
         strSql = ""
         If m_TMQ <> "" Then strSql = GetSetting("TAIE", TMQ_查名作業, UCase(UpForm.Name) & "Dir", "")
         
         If strSql <> "" Then
            cd1.InitDir = strSql
         Else
            cd1.InitDir = PUB_Getdesktop
         End If
         'end 2016/05/30
         
         cd1.FilterIndex = 0
         cd1.ShowOpen
         If Trim(cd1.FileName) <> "" Then
            Screen.MousePointer = vbHourglass
            'Added by Lydia 2015/11/06 查名單電子化:記錄路徑
            If m_TMQ <> "" Then
                sFile = Split(cd1.FileName, ChrW$(0))
                'Modified by Lydia 2016/05/26 記錄路徑只到資料夾位置
                'SaveSetting "TAIE", TMQ_查名作業, UCase(UpForm.Name) & "Dir", sFile(0)
                'UpForm.strLoadPath = sFile(0)
                SaveSetting "TAIE", TMQ_查名作業, UCase(UpForm.Name) & "Dir", Left(sFile(0), InStrRev(sFile(0), "\") - 1)
            End If
            'end 2015/11/06
            
            'InitAll
            'pic1.Move -20, -20, PicMain.Width, PicMain.Height
            'Set G_SeekPicColor.Picture = pvGetStdPicture(cd1.FileName, bSuccess, G_SeekPicBW)
            'Set G_SeekPicColor.Picture = pvGetStdPicture(cd1.FileName)
            If UCase(pvGetExt(cd1.FileName)) = "WMF" Or UCase(pvGetExt(cd1.FileName)) = "EMF" Then
                SavePicture pvGetStdPicture(cd1.FileName), App.path & "\NowPic.wmf"
                PicToObj App.path & "\NowPic.wmf"
                IsWmf = True
            Else
                PicToObj cd1.FileName
                IsWmf = False
            End If
'                IsWmf = False
'                Frame2.Enabled = True
'                If UCase(pvGetExt(cd1.FileName)) = "WMF" Or UCase(pvGetExt(cd1.FileName)) = "EMF" Then
'                   DoEvents
'                   IsWmf = True
'                   optColor(1).Enabled = False
'                Else
'                    optColor(1).Enabled = True
'                    optColor(0).Enabled = True
'                End If
'                optColor(0).Value = True
'                Call pvSetDIBPicture(G_SeekPicColor.Picture, optColor(1).Value)
'                ResizeImage True
                
                'Remove by Morgan 2010/7/27 不用了
                'If Val(strUserNum) < 65001 Then
                '    Frame1.Visible = True
                '    txtUserNo.SetFocus
                'End If
                'end 2010/7/27
                
                If Dir(App.path & "\tmp.bmp") <> "" Then
                    Kill App.path & "\tmp.bmp"
                End If
                If Dir(App.path & "\NowPic.wmf") <> "" Then
                    Kill App.path & "\NowPic.wmf"
                End If
                IsSave = False
                'Add by Amy 2018/07/02
                Frame2.Enabled = True '按過本所案號代表圖複製 需再開放
'            End If
            Screen.MousePointer = vbDefault
         End If
         Exit Sub

ErrHand:
    MsgBox "發生錯誤！", vbExclamation, "未知錯誤！"
    Screen.MousePointer = vbDefault
End Sub

'複製原程式改寫法
Private Sub PhotoSave()
    Dim intRun As Integer
    Dim stFileName As String, strFtpPath As String
    Dim intStatus As Integer '0不處理/1新增/2修改
    Dim bolChkColor As Boolean '判斷彩色代表圖
    
    '回傳圖片
    If oRtPic = True Then
        Screen.MousePointer = vbHourglass
        Set oPic.Picture = G_SeekPicColor.Picture
        Set oImg.Picture = tmpImg.Picture
        oImg.Height = tmpImg.Height
        oImg.Width = tmpImg.Width
        oImg.Top = tmpImg.Top
        oImg.Left = tmpImg.Left
        UpForm.IsWmf = IsWmf
        Screen.MousePointer = vbDefault
        Unload Me
        Exit Sub
    End If
    
    If FormCheck = False Then Exit Sub
    
    If Len(m_TMQ) > 0 Then stFileName = App.path & "\NowPic.jpg"
    Get #file_num, , bytes()
    '存檔
    If Len(m_TMQ) = 0 Then
        intRun = 1
        '案件
        'Modified by Morgan 2023/7/27
        'If Frame2.Visible = True Then
        If Frame2.Visible = True Or strWorkType = "" Then
        'end 2023/7/27
            If optColor(1).Value = True Then
                intRun = 2
            Else
                '選擇黑白圖需判斷是否有彩圖,避免彩圖切黑白圖再更新圖(彩圖未刪除)
                bolChkColor = True
            End If
            cnnConnection.BeginTrans
            For i = intRun To 1 Step -1
                intStatus = 0
                '彩色圖需判斷彩圖及灰階圖是否存在
                If ChkExists(i) = True Then
                    intStatus = 2 '修改
                '檔案不存在
                Else
                    intStatus = 1 '新增
                End If
                If intStatus > 0 Then
                    If InsUpdImg(intStatus, i, Trim(LOF(file_num)), strFtpPath, bolChkColor) = False Then
                        cnnConnection.RollbackTrans
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    '彩圖存檔後再轉換成黑白存入系統中
                    ElseIf i = 2 And intRun = 2 Then
                        file_num = FreeFile
                        If IsWmf = False Then
                            Call PicToObj(App.path & "\NowPic.bmp", True)
                            Open App.path & "\NowPic.jpg" For Binary Access Read As #file_num
                        Else
                            Call PicToObj(App.path & "\NowPic.wmf", True)
                            Open App.path & "\NowPic.wmf" For Binary Access Read As #file_num
                        End If
                        ReDim bytes(LOF(file_num))
                    End If
                End If
            Next i
            cnnConnection.CommitTrans
        '人事
        Else
            intStatus = 1
            If ChkExists(m_ibf05) = True Then
                intStatus = 2
            End If
            If InsUpdImg(intStatus, m_ibf05, Trim(LOF(file_num)), strFtpPath) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    '查名
    Else
        'Added by Lydia 2024/09/06 查名單-網中
        If m_TMQ = "A" Then
           If PUB_TMQAppFileSave(False, oCP01, oCP03, oCP04, stFileName) = False Then
              Screen.MousePointer = vbDefault
              Exit Sub
           End If
        Else
        'end 2024/09/06
           If PUB_TMQAFileSave(oCP01, oCP02, oCP03, Format(oCP04, TMQ_附件F04), "JPG", stFileName) = False Then
              Screen.MousePointer = vbDefault
              Exit Sub
           End If
        End If 'Added by Lydia 2024/09/06
    End If
    Close #file_num
    
    '刪除檔案
    If Dir(App.path & "\tmp.jpg") <> "" Then Kill App.path & "\tmp.jpg"
    If Dir(App.path & "\tmp.tif") <> "" Then Kill App.path & "\tmp.tif"
    If Dir(App.path & "\tmp1.tif") <> "" Then Kill App.path & "\tmp1.tif"
    If Dir(App.path & "\tmp.bmp") <> "" Then Kill App.path & "\tmp.bmp"
    If Dir(App.path & "\NowPic.wmf") <> "" Then Kill App.path & "\NowPic.wmf"
    If Dir(App.path & "\NowPic.jpg") <> "" Then Kill App.path & "\NowPic.jpg"
    If Dir(App.path & "\NowPic.bmp") <> "" Then Kill App.path & "\NowPic.bmp"
    
    If bolCall Then Exit Sub 'Added by Morgan 2023/7/27
    
    '重抓資料
    If Len(m_TMQ) = 0 Then '非查名
        Call StrMenu
        MsgBox "存檔完成！", vbInformation, "恭喜！"
    Else
        Call ReadTMQ
        bolTMQUpd = True
    End If
    Frame2.Enabled = True
    IsSave = True
    Screen.MousePointer = vbDefault
    If Len(m_TMQ) > 0 Then Unload Me

End Sub

Private Sub PhotoSave_Old()
'    Dim stFileName As String 'Added by Lydia 2016/06/23
'    Dim strFtpPath As String
'
'On Error GoTo 0
'On Error GoTo ErrHand
'        'add by nickc 2007/11/19 加入回傳圖片
'        If oRtPic = True Then
'            Screen.MousePointer = vbHourglass
'            Set oPic.Picture = G_SeekPicColor.Picture
'            Set oImg.Picture = tmpImg.Picture
'            oImg.Height = tmpImg.Height
'            oImg.Width = tmpImg.Width
'            oImg.Top = tmpImg.Top
'            oImg.Left = tmpImg.Left
'            UpForm.IsWmf = IsWmf
'            Screen.MousePointer = vbDefault
'            Unload Me
'        Else
'             'add by nickc 2005/12/12
'             If Frame1.Visible = True Then
'                If IsAuthorized(txtUserNo, txtPassword) = False Then
'                    txtUserNo.SetFocus
'                    Exit Sub
'                End If
'             Else
'                If IsSave = True Then
'                    MsgBox "沒有變更！", vbExclamation, "錯誤！"
'                    Exit Sub
'                End If
'             End If
'          'Added by Lydia 2015/8/5 查名單電子化:商標查名單0:圖形,1:文字1,2:文字2
'          If Len(m_TMQ) = 0 Then '非查名
'             Screen.MousePointer = vbHourglass
'             DoEvents
'             Set PicRs = New ADODB.Recordset
'             PicRs.CursorLocation = adUseClient
'
'             'Modify By Sindy 2012/6/14 原ibf05='1'改為ibf05='" & m_ibf05 & "'
'             PicRs.Open "select ImgByteFile.*,S1.st02 as Cst02,s2.st02 as Ust02 from ImgByteFile,staff S1,staff S2 where ibf05='" & m_ibf05 & "' and ibf01='" & oCP01 & "' and ibf02='" & oCP02 & "' and ibf03='" & oCP03 & "' and ibf04='" & oCP04 & "' and ibf07=s1.st01(+) and ibf10=s2.st01(+) ", cnnConnection, adOpenStatic, adLockOptimistic
'             If IsWmf = False Then
'    '            If optColor(1).Value = True Then
'    ''               Call pvSetDIBPicture(G_SeekPicColor.Picture, True)
'    ''               Call mGDIpEx.SaveDIB(DIB, App.Path & "\tmp.jpg", [ImageJPEG], , , 2)
'    '            Else
'    ''               Call pvSetDIBPicture(G_SeekPicBW.Picture, True)
'    ''               Call mGDIpEx.SaveDIB(DIB, App.Path & "\tmp.jpg", [ImageJPEG], , , 2)
'    '            End If
'                If Dir(App.path & "\NowPic.jpg") <> "" Then Kill App.path & "\NowPic.jpg"
'                If Dir(App.path & "\NowPic.bmp") <> "" Then Kill App.path & "\NowPic.bmp"
'                SavePicture G_SeekPicColor.Picture, App.path & "\NowPic.bmp"
'                PicToObj App.path & "\NowPic.bmp"
'             Else
'    '            SavePicture G_SeekPicColor.Picture, App.Path & "\tmp.wmf"
'                If Dir(App.path & "\NowPic.wmf") <> "" Then Kill App.path & "\NowPic.wmf"
'                SavePicture G_SeekPicColor.Picture, App.path & "\NowPic.wmf"
'                PicToObj App.path & "\NowPic.wmf"
'             End If
'             file_num = FreeFile
'             If IsWmf = False Then
'                Open App.path & "\NowPic.jpg" For Binary Access Read As #file_num
'    '            If optColor(1).Value = True Then
'    '               Open App.Path & "\tmp.jpg" For Binary Access Read As #file_num
'    '            Else
'    '               Open App.Path & "\tmp.jpg" For Binary Access Read As #file_num
'    '            End If
'             Else
'                Open App.path & "\NowPic.wmf" For Binary Access Read As #file_num
'             End If
'
'         ReDim bytes(LOF(file_num))
'
'            'Modify By Sindy 2012/6/21
'            If strWorkType = "1" Then '人事照片
'               '超過 50K 秀錯誤
'               If LOF(file_num) > 51200 Then
'                  If Pub_StrUserSt03 = "M51" Then
'                      If MsgBox("檔案太大！" & vbCrLf & "請去除不必要的陰影及複雜線條，或縮小檔案！" & vbCrLf & "是否還是要存入??", vbCritical + vbYesNo, "發生錯誤！") = vbNo Then
'                          Screen.MousePointer = vbDefault
'                          Close #file_num
'                          Exit Sub
'                      End If
'                  Else
'                      MsgBox "檔案太大！" & vbCrLf & "請去除不必要的陰影及複雜線條，或縮小檔案！", vbCritical, "發生錯誤！"
'                      Screen.MousePointer = vbDefault
'                      Close #file_num
'                      Exit Sub
'                  End If
'               End If
'            Else
'            '2012/6/21 End
'               'add by nickc 2007/09/27 超過 300K 秀錯誤
'               If LOF(file_num) > 307200 Then
'                  If Pub_StrUserSt03 = "M51" Then
'                      If MsgBox("檔案太大！" & vbCrLf & "請去除不必要的陰影及複雜線條，或縮小檔案！" & vbCrLf & "是否還是要存入??", vbCritical + vbYesNo, "發生錯誤！") = vbNo Then
'                          Screen.MousePointer = vbDefault
'                          Close #file_num
'                          Exit Sub
'                      End If
'                  Else
'                      MsgBox "檔案太大！" & vbCrLf & "請去除不必要的陰影及複雜線條，或縮小檔案！", vbCritical, "發生錯誤！"
'                      Screen.MousePointer = vbDefault
'                      Close #file_num
'                      Exit Sub
'                  End If
'               End If
'            End If
'            Get #file_num, , bytes()
'
'            'Add By Sindy 2017/8/10 檔案改放 FTP,必須在DB資料刪除前執行
'            PUB_DelFtpFile2 Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-" & m_ibf05, , UCase("ImgByteFile")
'
'            If PicRs.RecordCount = 0 Then
'                  PicRs.AddNew
'                  'Modify by Morgan 2010/7/27
'                  'PicRs.Fields("ibf07").Value = IIf(Val(strUserNum) < 65001, txtUserNo, strUserNum)
'                  PicRs.Fields("ibf07").Value = strUserNum
'                  'end 2010/7/27
'                  PicRs.Fields("ibf08").Value = Val(strSrvDate(1))
'                  PicRs.Fields("ibf09").Value = Val(Format(time, "HHMM"))
'            Else
'                  'Modify by Morgan 2010/7/27
'                  'PicRs.Fields("ibf10").Value = IIf(Val(strUserNum) < 65001, txtUserNo, strUserNum)
'                  PicRs.Fields("ibf10").Value = strUserNum
'                  'end 2010/7/27
'                  PicRs.Fields("ibf11").Value = Val(strSrvDate(1))
'                  PicRs.Fields("ibf12").Value = Val(Format(time, "HHMM"))
'            End If
'            PicRs.Fields("ibf01").Value = Trim(oCP01)
'            PicRs.Fields("ibf02").Value = Trim(oCP02)
'            PicRs.Fields("ibf03").Value = Trim(oCP03)
'            PicRs.Fields("ibf04").Value = Trim(oCP04)
'            'Modify By Sindy 2012/6/14 原ibf05='1'改為ibf05 = m_ibf05
'            PicRs.Fields("ibf05").Value = m_ibf05
'            PicRs.Fields("ibf06").Value = IIf(optColor(1).Value = True, IIf(IsWmf = True, "4", "2"), IIf(IsWmf = True, "3", "1"))
'            PicRs.Fields("ibf13").Value = Trim(LOF(file_num))
''            PicRs.Fields("ibf14").Value = Null
''            PicRs.Fields("ibf14").AppendChunk bytes()
'            Close #file_num
'            'Modify By Sindy 2017/8/10
'            '檔案改放FTP
'            If IsWmf = False Then
'               PUB_PutFtpFile App.path & "\NowPic.jpg", Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-" & m_ibf05, Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-" & m_ibf05, strFtpPath, UCase("imgbytefile")
'            Else
'               PUB_PutFtpFile App.path & "\NowPic.wmf", Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-" & m_ibf05, Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-" & m_ibf05, strFtpPath, UCase("imgbytefile")
'            End If
'            If strFtpPath <> "" Then
'               PicRs.Fields("ibf15") = strFtpPath
'            End If
'            '2017/8/10 END
'            PicRs.UPDATE
'
'            If Dir(App.path & "\tmp.jpg") <> "" Then
'               Kill App.path & "\tmp.jpg"
'            End If
'            If Dir(App.path & "\tmp.tif") <> "" Then
'               Kill App.path & "\tmp.tif"
'            End If
'            If Dir(App.path & "\tmp1.tif") <> "" Then
'               Kill App.path & "\tmp1.tif"
'            End If
'            If Dir(App.path & "\tmp.bmp") <> "" Then
'               Kill App.path & "\tmp.bmp"
'            End If
'            If Dir(App.path & "\NowPic.wmf") <> "" Then
'               Kill App.path & "\NowPic.wmf"
'            End If
'            If Dir(App.path & "\NowPic.jpg") <> "" Then
'               Kill App.path & "\NowPic.jpg"
'            End If
'            If Dir(App.path & "\NowPic.bmp") <> "" Then
'                Kill App.path & "\NowPic.bmp"
'            End If
'            PicRs.Requery
'            lbl1(1).Caption = CheckStr(PicRs.Fields("ibf07")) & "  " & CheckStr(PicRs.Fields("Cst02"))
'            lbl1(2).Caption = ChangeWStringToTDateString(CheckStr(PicRs.Fields("ibf08").Value))
'            lbl1(3).Caption = Format(CheckStr(PicRs.Fields("ibf09")), "00:00")
'            lbl1(4).Caption = CheckStr(PicRs.Fields("ibf10")) & "  " & CheckStr(PicRs.Fields("Ust02"))
'            lbl1(5).Caption = ChangeWStringToTDateString(CheckStr(PicRs.Fields("ibf11").Value))
'            lbl1(6).Caption = Format(CheckStr(PicRs.Fields("ibf12")), "00:00")
'            lbl1(7).Caption = Format(CheckStr(PicRs.Fields("ibf13")), "###,###,###,###,##0") & " 位元組"
'            'Frame2.Enabled = False
'            Frame2.Enabled = True 'Modify By Sindy 98/04/09
'            MsgBox "存檔完成！", vbInformation, "恭喜！"
'            IsSave = True
'            Screen.MousePointer = vbDefault
'          Else
'          'Added by Lydia 2015/8/5 商標查名單0:圖形,1:文字1,2:文字2
'             Screen.MousePointer = vbHourglass
'             DoEvents
'             Set PicRs = New ADODB.Recordset
'             PicRs.CursorLocation = adUseClient
'             'Modified by Lydia 2016/03/22
'             'PicRs.Open "select * from TMQFile where TQF01='" & oCP01 & "' AND TQF02='" & oCP02 & "' AND TQF03='" & oCP03 & "' AND TQF04='" & Format(oCP04, "00") & "' ", cnnConnection, adOpenStatic, adLockOptimistic
'             'Remove by Lydia 2016/07/07
'             'If strSrvDate(1) < TMQFileFTP Then
'             '   PicRs.Open "select * from TMQFile where TQF01='" & oCP01 & "' AND TQF02='" & oCP02 & "' AND TQF03='" & oCP03 & "' AND TQF04='" & Format(oCP04, TMQ_附件F04) & "' ", cnnConnection, adOpenStatic, adLockOptimistic
'             'End If
'             If IsWmf = False Then
'                If Dir(App.path & "\NowPic.jpg") <> "" Then Kill App.path & "\NowPic.jpg"
'                If Dir(App.path & "\NowPic.bmp") <> "" Then Kill App.path & "\NowPic.bmp"
'                       SavePicture G_SeekPicColor.Picture, App.path & "\NowPic.bmp"
'                       PicToObj App.path & "\NowPic.bmp"
'             Else
'                If Dir(App.path & "\NowPic.wmf") <> "" Then Kill App.path & "\NowPic.wmf"
'                       SavePicture G_SeekPicColor.Picture, App.path & "\NowPic.wmf"
'                       PicToObj App.path & "\NowPic.wmf"
'             End If
'             file_num = FreeFile
'             If IsWmf = False Then
'                Open App.path & "\NowPic.jpg" For Binary Access Read As #file_num
'                stFileName = App.path & "\NowPic.jpg" 'Added by Lydia 2016/06/23
'             Else
'                Open App.path & "\NowPic.wmf" For Binary Access Read As #file_num
'                stFileName = App.path & "\NowPic.jpg" 'Added by Lydia 2016/06/23
'             End If
'
'            ReDim bytes(LOF(file_num))
'
'            If LOF(file_num) > 307200 Then '超過300K提示
'               If Pub_StrUserSt03 = "M51" Then
'                   If MsgBox("檔案太大！" & vbCrLf & "請去除不必要的陰影及複雜線條，或縮小檔案！" & vbCrLf & "是否還是要存入??", vbCritical + vbYesNo, "發生錯誤！") = vbNo Then
'                       Screen.MousePointer = vbDefault
'                       Close #file_num
'                       Exit Sub
'                   End If
'               Else
'                   MsgBox "檔案太大！" & vbCrLf & "請去除不必要的陰影及複雜線條，或縮小檔案！", vbCritical, "發生錯誤！"
'                   Screen.MousePointer = vbDefault
'                   Close #file_num
'                   Exit Sub
'               End If
'            End If
'
'            Get #file_num, , bytes()
'            'Modified by Lydia 2016/06/23 改成模組
'            'Remove by Lydia 2016/07/07
'            'If strSrvDate(1) < TMQFileFTP Then
'            '    If PicRs.RecordCount = 0 Then PicRs.AddNew
'            '
'            '    PicRs.Fields("TQF01").Value = Trim(oCP01)
'            '    PicRs.Fields("TQF02").Value = Trim(oCP02)
'            '    PicRs.Fields("TQF03").Value = Trim(oCP03)
'            '    'Modified by Lydia 2016/03/22
'            '    'PicRs.Fields("TQF04").Value = Format(oCP04, "00")
'            '    PicRs.Fields("TQF04").Value = Format(oCP04, TMQ_附件F04)
'            '    '直接存檔案類型
'            '    PicRs.Fields("TQF05").Value = "JPG"
'            '    PicRs.Fields("TQF06").Value = Trim(LOF(file_num))
'            '    PicRs.Fields("TQF07").Value = Null
'            '    PicRs.Fields("TQF07").AppendChunk bytes()
'            '    PicRs.Fields("TQF08").Value = strUserNum
'            '    PicRs.Fields("TQF09").Value = Val(strSrvDate(1))
'            '    PicRs.Fields("TQF10").Value = Val(Format(time, "HHMM"))
'            '    PicRs.Fields("TQF11").Value = Null
'            '    PicRs.UPDATE
'            '    Close #file_num
'            'Else
'                Close #file_num
'                If PUB_TMQAFileSave(oCP01, oCP02, oCP03, Format(oCP04, TMQ_附件F04), "JPG", stFileName) Then
'                End If
'            'End If
'            'end 2016/06/23
'
'            If Dir(App.path & "\tmp.jpg") <> "" Then
'               Kill App.path & "\tmp.jpg"
'            End If
'            If Dir(App.path & "\tmp.tif") <> "" Then
'               Kill App.path & "\tmp.tif"
'            End If
'            If Dir(App.path & "\tmp1.tif") <> "" Then
'               Kill App.path & "\tmp1.tif"
'            End If
'            If Dir(App.path & "\tmp.bmp") <> "" Then
'               Kill App.path & "\tmp.bmp"
'            End If
'            If Dir(App.path & "\NowPic.wmf") <> "" Then
'               Kill App.path & "\NowPic.wmf"
'            End If
'            If Dir(App.path & "\NowPic.jpg") <> "" Then
'               Kill App.path & "\NowPic.jpg"
'            End If
'            If Dir(App.path & "\NowPic.bmp") <> "" Then
'                Kill App.path & "\NowPic.bmp"
'            End If
'            'Modified by Lydia 2016/06/23
'            'Remove by Lydia 2016/07/07
'            'If strSrvDate(1) < TMQFileFTP Then
'            '    PicRs.ReQuery
'            'Else
'                Set PicRs = New ADODB.Recordset
'                PicRs.CursorLocation = adUseClient
'                PicRs.Open "select TQF06,TQF08,TQF09,TQF10 from TMQFile where TQF01='" & oCP01 & "' AND TQF02='" & oCP02 & "' AND TQF03='" & oCP03 & "' AND TQF04='" & Format(oCP04, TMQ_附件F04) & "' ", cnnConnection, adOpenStatic, adLockOptimistic
'            'End If
'            ''end 2016/06/23
'
'            lbl1(1).Caption = CheckStr(PicRs.Fields("TQF08"))
'            lbl1(2).Caption = ChangeWStringToTDateString(CheckStr(PicRs.Fields("TQF09").Value))
'            lbl1(3).Caption = Format(CheckStr(PicRs.Fields("TQF10")), "00:00")
'            lbl1(7).Caption = Format(CheckStr(PicRs.Fields("TQF06")), "###,###,###,###,##0") & " 位元組"
'            Frame2.Enabled = True
'           'Modified by Lydia 2015/11/06
'           'MsgBox "存檔完成！", vbInformation, "恭喜！"
'            bolTMQUpd = True
'            IsSave = True
'            Screen.MousePointer = vbDefault
'            Unload Me
'          End If
'        End If
'        Exit Sub
'
'ErrHand:
'    MsgBox "發生錯誤！", vbExclamation, "未知錯誤！"
'    Screen.MousePointer = vbDefault
End Sub

Private Sub PContinue()
On Error GoTo 0
On Error GoTo ErrHand
     'Added by Lydia 2015/8/5 查名單電子化:商標查名單0:圖形,1:文字1,2:文字2
        If Len(m_TMQ) = 0 Then '非查名
            'add by nickc 2007/11/19 加入回傳圖片
            If oRtPic = True Then
                Unload Me
            Else
                If cmdOK(3).Caption = "取消(&X)" Then
            '        ShowGrip False
                    For i = 0 To UBound(SeekCmdok())
                        cmdOK(i).Caption = SeekCmdok(i).Caption
                        cmdOK(i).Enabled = SeekCmdok(i).Enabled
                        cmdOK(i).Visible = SeekCmdok(i).Visible
                    Next i
                    Frame2.Enabled = True
                Else
                    If IsSave = False Then
                        If MsgBox("還沒儲存圖檔，是否要儲存(Yes/No)？", vbYesNo + vbCritical, "警告！") = vbYes Then
                            cmdok_Click 2
                            If IsSave = True Then
                                Unload Me
                            End If
                        Else
                            Unload Me
                        End If
                    Else
                        Unload Me
                    End If
                End If
            End If
        Else
        'Added by Lydia 2015/8/5 商標查名單0:圖形,1:文字1,2:文字2
            If IsSave = False Then
                If MsgBox("還沒儲存圖檔，是否要儲存(Yes/No)？", vbYesNo + vbCritical, "警告！") = vbYes Then
                    cmdok_Click 2
                    If IsSave = True Then
                        Unload Me
                    End If
                Else
                    Unload Me
                End If
            Else
                Unload Me
            End If
        End If
        Exit Sub

ErrHand:
    MsgBox "發生錯誤！", vbExclamation, "未知錯誤！"
    Screen.MousePointer = vbDefault
End Sub

Private Sub GeneralScan()
Dim lRtn As Long

On Error GoTo 0
On Error GoTo ErrHand
    If Dir(App.path & "\tmp.bmp") <> "" Then
      Kill App.path & "\tmp.bmp"
    End If
    Screen.MousePointer = vbHourglass
    'edit by nickc 2006/05/01
    'lRtn = mdlTwain.TransferWithoutUI(200, to_BW, 1, 1, 0, 0, App.Path & "\tmp.bmp")
    'edit by nickc 2006/05/24 再改回 200 dpi
    'lRtn = mdlTwain.TransferWithoutUI(150, to_BW, 1, 1, 0, 0, App.Path & "\tmp.bmp")
    lRtn = mdlTwain.TransferWithoutUI(200, to_BW, 1, 1, 0, 0, App.path & "\tmp.bmp")
    Screen.MousePointer = vbDefault
    If lRtn = 0 Then
        'InitAll
        'Set G_SeekPicColor.Picture = pvGetStdPicture(App.Path & "\tmp.bmp", bSuccess, G_SeekPicBW)
        'Set G_SeekPicColor.Picture = pvGetStdPicture(App.Path & "\tmp.bmp")
        PicToObj App.path & "\tmp.bmp"
        'If (bSuccess) Then
        '    IsWmf = False
            '僅限黑白
            'Frame2.Enabled = False
            Frame2.Enabled = True 'Modify By Sindy 98/04/09
            optColor(0).Value = True
            pic1.AutoSize = False
            pic1.AutoRedraw = True
'            Call pvSetDIBPicture(G_SeekPicColor.Picture, optColor(1).Value)
'            ResizeImage True
            
            'Remove by Morgan 2010/7/27
            'If Val(strUserNum) < 65001 Then
            '    Frame1.Visible = True
            '    txtUserNo.SetFocus
            'End If
            'end 2010/7/27
            
            IsSave = False
            If Dir(App.path & "\tmp.bmp") <> "" Then
                Kill App.path & "\tmp.bmp"
            End If
        'End If
    Else
        MsgBox "掃描失敗！" & vbCrLf & "請檢查裝置是否正確！", vbCritical, "錯誤！"
    End If
    Exit Sub

ErrHand:
    MsgBox "發生錯誤！", vbExclamation, "未知錯誤！"
    Screen.MousePointer = vbDefault
End Sub

Private Sub GrayScan()
Dim lRtn As Long

On Error GoTo 0
On Error GoTo ErrHand

    If Dir(App.path & "\tmp.bmp") <> "" Then
      Kill App.path & "\tmp.bmp"
    End If
    Screen.MousePointer = vbHourglass
    'edit by nickc 2006/05/26
    'lRtn = mdlTwain.TransferWithUI(App.Path & "\tmp.bmp")
    lRtn = mdlTwain.TransferWithoutUI(200, to_GREY, 1, 1, 0, 0, App.path & "\tmp.bmp")
    Screen.MousePointer = vbDefault
    If lRtn = 0 Then
        'InitAll
        'Set G_SeekPicColor.Picture = pvGetStdPicture(App.Path & "\tmp.bmp")
        'If (bSuccess) Then
        PicToObj App.path & "\tmp.bmp"
        '    IsWmf = False
            Frame2.Enabled = True
            optColor(0).Value = True
'            Call pvSetDIBPicture(G_SeekPicColor.Picture, optColor(1).Value)
'            ResizeImage True
            
            'Remove by Morgan 2010/7/27
            'If Val(strUserNum) < 65001 Then
            '    Frame1.Visible = True
            '    txtUserNo.SetFocus
            'End If
            'end 2010/7/27
            
            IsSave = False
        'End If
    Else
        MsgBox "掃描失敗！" & vbCrLf & "請檢查裝置是否正確！", vbCritical, "錯誤！"
    End If
    Exit Sub

ErrHand:
    MsgBox "發生錯誤！", vbExclamation, "未知錯誤！"
    Screen.MousePointer = vbDefault
End Sub

Private Sub PhotoCopy()
On Error GoTo 0
On Error GoTo ErrHand
'    If cmdok(7).Caption = "窗選(&A)" Then
'        For i = 0 To UBound(SeekCmdok())
'            SeekCmdok(i).Caption = cmdok(i).Caption
'            SeekCmdok(i).Enabled = cmdok(i).Enabled
'            SeekCmdok(i).Visible = cmdok(i).Visible
'        Next i
'        cmdok(7).Caption = "裁切(&A)"
'        cmdok(3).Caption = "取消(&X)"
'        cmdok(5).Enabled = False
'        cmdok(6).Enabled = False
'        cmdok(4).Enabled = False
'        cmdok(0).Enabled = False
'        cmdok(1).Enabled = False
'        cmdok(2).Enabled = False
'        Frame2.Enabled = False
'        ShowGrip True
'    Else
'        Crop lblShape.Left, lblShape.Top, lblShape.Width, lblShape.Height
'        ShowGrip False
'        For i = 0 To UBound(SeekCmdok())
'            cmdok(i).Caption = SeekCmdok(i).Caption
'            cmdok(i).Enabled = SeekCmdok(i).Enabled
'            cmdok(i).Visible = SeekCmdok(i).Visible
'        Next i
'        Frame2.Enabled = True
'    End If
    Clipboard.Clear
    Clipboard.SetData G_SeekPicColor.Picture
    Set PUB_PicCopy = G_SeekPicColor.Picture 'Add by Morgan 2010/9/17
    
    'Modify by Amy 2017/01/20 +不秀訊息判斷
    If bolNoMsg = False Then MsgBox "複製完成！", vbInformation, "恭喜！"
    Exit Sub

ErrHand:
    MsgBox "發生錯誤！", vbExclamation, "未知錯誤！"
    Screen.MousePointer = vbDefault
End Sub

Private Sub ShowBt(Optional ByVal IsFirst As Boolean = False)
    If IsFirst = True Then
        cmdOK(9).Visible = False
        cmdOK(9).Enabled = False
        If CountPhoto > 1 Then
            cmdOK(9).Visible = True
            cmdOK(9).Enabled = True
            cmdOK(9).Caption = "灰階 切換"
        End If
        SeekCmdok(9).Enabled = cmdOK(9).Enabled
        SeekCmdok(9).Visible = cmdOK(9).Visible
        Exit Sub
    End If
    
    '目前為彩色,顯示灰階 切換
    If cmdOK(9).Caption = "灰階 切換" Then
        m_ibf05 = "1"
        Call GetPhoto(False)
        cmdOK(9).Caption = "彩色 切換"
    '目前為灰階,顯示彩色 切換
    Else
        m_ibf05 = "2"
        Call GetPhoto(False)
        cmdOK(9).Caption = "灰階 切換"
    End If
End Sub

'增加彩色代表圖,原程式放至StrMenu_Old Mark
Sub StrMenu()
Dim strA1 As String, strA2 As String, strCon As String 'Added by Lydia 2024/05/10
    FormClear
    cmdOK(1).Enabled = False '選擇檔案
    cmdOK(2).Enabled = False '存檔
    
    '員工照片
    If strWorkType = "1" Then
        'Added by Lydia 2024/05/10 聯絡人相片
        'Modified by Lydia 2024/05/14 +潛在客戶R
        If Left(oCP01, 1) = "X" Or Left(oCP01, 1) = "Y" Or Left(oCP01, 1) = "R" Then
            Label1(0) = "聯絡人編號："
            If Trim(oCP01 & oCP02 & oCP03 & oCP04) = "" Then
                LBL1(0).Caption = "錯誤聯絡人編號！"
            Else
                strA1 = Mid(oCP01 & oCP02 & oCP03, 1, 8)
                strA2 = Mid(oCP01 & oCP02 & oCP03, 9, 2)
                LBL1(0).Caption = strA1 & "-" & strA2
                If bolQuery = False Then
                    cmdOK(1).Enabled = True  '選擇檔案
                    cmdOK(2).Enabled = True  '存檔
                End If
            End If
            Label1(3).Visible = True
            LBL1(8).Visible = True
            Label1(0).Left = Label4.Left
            Label1(3).Left = Label4.Left
            Label1(3).Caption = "聯絡人名稱："
            If ClsPDGetContact(LBL1(0).Caption, strCon) = True Then
               LBL1(8).Caption = strCon
            End If
        Else
        'end 2024/05/10
            Label1(0) = "員工代號："
            If Trim(oCP01 & oCP02 & oCP03 & oCP04) = "" Then
                LBL1(0).Caption = "錯誤員工代號！"
            Else
                LBL1(0).Caption = oCP02 & "  " & GetPrjSalesNM(oCP02)
                If bolQuery = False Then
                    cmdOK(1).Enabled = True  '選擇檔案
                    cmdOK(2).Enabled = True  '存檔
                End If
            End If
            'Move by Lydia 2024/05/10 從下面移上來
            Label1(3).Visible = False '案件名稱
            LBL1(8).Visible = False
            'end 2024/05/10
        End If
        m_ibf05 = "3"
        Frame2.Visible = False '黑色或彩色
        Label13.Caption = "備註：影像尺寸請縮小至480(W)*640(H)以下　　　檔案大小限 50 KB"

    '案件
    Else
        Label1(0) = "本所案號："
        If Trim(oCP01 & oCP02 & oCP03 & oCP04) = "" Then
            LBL1(0).Caption = "錯誤本所案號！"
        Else
            LBL1(0).Caption = oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04
            LBL1(8).Caption = GetPrjName(oCP01 & "-" & oCP02 & "-" & oCP03 & "-" & oCP04) '案件名稱
            cmdOK(1).Enabled = True  '選擇檔案
            cmdOK(2).Enabled = True  '存檔
        End If
        If InStr(oCP01, "T") > 0 Then
            Label11.Caption = "商標代表圖"
        Else
            Label11.Caption = "專利代表圖"
        End If
        Label13.Caption = "備註：彩色圖檔案大小1000KB／灰階圖檔案大小限 300 KB"
        Label1(3).Visible = True '案件名稱
        LBL1(8).Visible = True
    End If
    Screen.MousePointer = vbHourglass
    DoEvents
    If GetPhoto(False, IIf(strWorkType = "1", True, False)) = False Then
        '無圖示
        Call GetPhoto(True)
    End If
    Call ShowBt(True)
    Screen.MousePointer = vbDefault
End Sub

'取得圖示
'bolNoPhoto:True為無圖示
'NotCaseAll:True不抓案件所有代表圖
Private Function GetPhoto(ByVal bolNoPhoto As Boolean, Optional ByVal NotCaseAll As Boolean = True) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    
    GetPhoto = False
    If bolNoPhoto = False Then
        'Modify By Sindy 2018/10/31 TF馬德里商標圖檔-子案的圖同母案
        '固定都以IBF01=tm01 AND IBF02=substr(tm02,1,5)||'0' AND IBF03='0' AND IBF04='00' 去抓代表圖
        If oCP01 = "TF" Then
            strQ = "ibf01='" & oCP01 & "' and ibf02='" & Mid(oCP02, 1, 5) & "0" & "' and ibf03='0' And ibf04='00' and ibf07=s1.st01(+) and ibf10=s2.st01(+) "
        Else
        '2018/10/31 END
            strQ = "ibf01='" & oCP01 & "' and ibf02='" & oCP02 & "' and ibf03='" & oCP03 & "' and ibf04='" & oCP04 & "' and ibf07=s1.st01(+) and ibf10=s2.st01(+) "
        End If
        '案件所有代表圖(有彩色圖先顯示)
        If NotCaseAll = False Then
            strQ = "Select ImgByteFile.*,S1.st02 as Cst02,s2.st02 as Ust02,1 as sort From ImgByteFile,Staff S1,Staff S2 Where " & strQ & " And ibf05='2' " & _
             "Union Select ImgByteFile.*,S1.st02 as Cst02,s2.st02 as Ust02,2 as sort From ImgByteFile,Staff S1,Staff S2 Where " & strQ & " And ibf05='1' " & _
             "Order by sort"
        '員工照片/只顯示彩色代表圖或灰階代表圖
        Else
            strQ = "Select ImgByteFile.*,S1.st02 as Cst02,s2.st02 as Ust02 From ImgByteFile,Staff S1,Staff S2 " & _
                      "Where " & strQ & " And ibf05='" & m_ibf05 & "' "
        End If
    Else
        '無圖式
        strQ = "Select ImgByteFile.* From ImgByteFile " & _
                  "Where ibf05='" & m_ibf05 & "' and ibf01='000' and ibf02='000000' and ibf03='0' and ibf04='00' "
    End If
  
    If RsQ.State = adStateOpen Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockOptimistic
    If RsQ.RecordCount <> 0 Then
        RsQ.MoveFirst
        If NotCaseAll = False Then CountPhoto = RsQ.RecordCount
        GetPhoto = True
        If bolNoPhoto = False Then
            LBL1(1).Caption = CheckStr(RsQ.Fields("ibf07")) & "  " & CheckStr(RsQ.Fields("Cst02"))
            LBL1(2).Caption = ChangeWStringToTDateString(CheckStr(RsQ.Fields("ibf08").Value))
            LBL1(3).Caption = Format(CheckStr(RsQ.Fields("ibf09")), "00:00")
            LBL1(4).Caption = CheckStr(RsQ.Fields("ibf10")) & "  " & CheckStr(RsQ.Fields("Ust02"))
            LBL1(5).Caption = ChangeWStringToTDateString(CheckStr(RsQ.Fields("ibf11").Value))
            LBL1(6).Caption = Format(CheckStr(RsQ.Fields("ibf12")), "00:00")
            LBL1(7).Caption = Format(CheckStr(RsQ.Fields("ibf13")), "###,###,###,###,##0") & " 位元組"
            m_ibf05 = "" & RsQ.Fields("ibf05") 'Added by Morgan 2021/1/26
        End If
        '格式 1.黑白位圖/2.彩色位圖/3.黑白向量圖/4.彩色向量圖/6.無圖式
        If CheckStr(RsQ.Fields("ibf06")) = "1" Or CheckStr(RsQ.Fields("ibf06")) = "3" Then
            optColor(0).Value = True
        Else
            optColor(1).Value = True
        End If
        If CheckStr(RsQ.Fields("ibf06")) = "3" Or CheckStr(RsQ.Fields("ibf06")) = "4" Or CheckStr(RsQ.Fields("ibf06")) = "6" Then
            IsWmf = True
        Else
            IsWmf = False
        End If
        '下載檔案
        If IsWmf = False Then
            Call PUB_GetFtpFile(RsQ.Fields("IBF15"), App.path & "\NowPic.jpg", UCase("ImgByteFile"))
        Else
            Call PUB_GetFtpFile(RsQ.Fields("IBF15"), App.path & "\NowPic.wmf", UCase("ImgByteFile"))
        End If
        Frame2.Enabled = True
        If IsWmf = False Then
            PicToObj Trim(App.path & "\NowPic.jpg")
        Else
            PicToObj Trim(App.path & "\NowPic.wmf")
        End If
        If Dir(App.path & "\NowPic.jpg") <> "" Then
            Kill App.path & "\NowPic.jpg"
        End If
        If Dir(App.path & "\tmp.tif") <> "" Then
            Kill App.path & "\tmp.tif"
        End If
        If Dir(App.path & "\NowPic.wmf") <> "" Then
            Kill App.path & "\NowPic.wmf"
        End If
    End If
    RsQ.Close
End Function

Private Function GetCopyPhoto() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
  
    GetCopyPhoto = False
    If txtCode(1) = MsgText(601) Then txtCode(1) = "0"
    If txtCode(2) = MsgText(601) Then txtCode(2) = "00"
    'Modify By Sindy 2018/10/31 TF馬德里商標圖檔-子案的圖同母案
    '固定都以IBF01=tm01 AND IBF02=substr(tm02,1,5)||'0' AND IBF03='0' AND IBF04='00' 去抓代表圖
    If txtSystem = "TF" Then
      strQ = "Select * From ImgByteFile Where ibf05='" & IIf(optColor(0).Value = True, 1, 2) & "' " & _
                "and ibf01='" & txtSystem & "' and ibf02='" & Mid(txtCode(0), 1, 5) & "0" & "' and ibf03='0' And ibf04='00' "
    Else
    '2018/10/31 END
      strQ = "Select * From ImgByteFile Where ibf05='" & IIf(optColor(0).Value = True, 1, 2) & "' " & _
                "and ibf01='" & txtSystem & "' and ibf02='" & txtCode(0) & "' and ibf03='" & txtCode(1) & "' and ibf04='" & txtCode(2) & "' "
    End If
    If RsQ.State = adStateOpen Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockOptimistic
    If RsQ.RecordCount = 0 Then
        MsgBox "此案件無" & IIf(optColor(1).Value = True, "彩色", "") & "代表圖！"
        Exit Function
    Else
        RsQ.MoveFirst
        If CheckStr(RsQ.Fields("ibf06")) = "3" Or CheckStr(RsQ.Fields("ibf06")) = "4" Then
            IsWmf = True
        Else
            IsWmf = False
        End If
        '下載檔案
        If IsWmf = False Then
            Call PUB_GetFtpFile(RsQ.Fields("IBF15"), App.path & "\TempPic.jpg", UCase("ImgByteFile"))
        Else
            Call PUB_GetFtpFile(RsQ.Fields("IBF15"), App.path & "\TempPic.wmf", UCase("ImgByteFile"))
        End If
        
        If IsWmf = False Then
            PicToObj2 Trim(App.path & "\TempPic.jpg")
        Else
            PicToObj2 Trim(App.path & "\TempPic.wmf")
        End If
        If Dir(App.path & "\TempPic.jpg") <> "" Then
            Kill App.path & "\TempPic.jpg"
        End If
        If Dir(App.path & "\TempPic.wmf") <> "" Then
            Kill App.path & "\TempPic.wmf"
        End If
    End If
    RsQ.Close
End Function

Private Function FormCheck() As Boolean
    Dim strMsg As String
    
    FormCheck = False
    If Frame1.Visible = True Then
        If IsAuthorized(txtUserNo, txtPassword) = False Then
            txtUserNo.SetFocus
            Exit Function
        End If
    Else
        If IsSave = True Then
            MsgBox "沒有變更！", vbExclamation, "錯誤！"
            Exit Function
        End If
    End If
    '讀檔案
    If IsWmf = False Then
        If Dir(App.path & "\NowPic.jpg") <> "" Then Kill App.path & "\NowPic.jpg"
        If Dir(App.path & "\NowPic.bmp") <> "" Then Kill App.path & "\NowPic.bmp"
        SavePicture G_SeekPicColor.Picture, App.path & "\NowPic.bmp"
        PicToObj App.path & "\NowPic.bmp"
    Else
        If Dir(App.path & "\NowPic.wmf") <> "" Then Kill App.path & "\NowPic.wmf"
        SavePicture G_SeekPicColor.Picture, App.path & "\NowPic.wmf"
        PicToObj App.path & "\NowPic.wmf"
    End If
    file_num = FreeFile
    If IsWmf = False Then
        Open App.path & "\NowPic.jpg" For Binary Access Read As #file_num
    Else
        Open App.path & "\NowPic.wmf" For Binary Access Read As #file_num
    End If
    ReDim bytes(LOF(file_num))
    
    strMsg = "檔案太大！" & vbCrLf & "請去除不必要的陰影及複雜線條，或縮小檔案！"
    
    '人事照片
    If strWorkType = "1" Then
        '超過 50K 秀錯誤
        If LOF(file_num) > 51200 Then
            If Pub_StrUserSt03 = "M51" Then
                If MsgBox(strMsg & vbCrLf & "是否還是要存入??", vbCritical + vbYesNo, "發生錯誤！") = vbNo Then
                    Screen.MousePointer = vbDefault
                    Close #file_num
                    Exit Function
                End If
            Else
                MsgBox strMsg, vbCritical, "發生錯誤！"
                Screen.MousePointer = vbDefault
                Close #file_num
                Exit Function
            End If
        End If
    '案件彩色代表圖
    ElseIf optColor(1).Value = True Then
        '超過 1MB 秀錯誤
        If LOF(file_num) > 1048576 Then
            If Pub_StrUserSt03 = "M51" Then
                If MsgBox(strMsg & vbCrLf & "是否還是要存入??", vbCritical + vbYesNo, "發生錯誤！") = vbNo Then
                    Screen.MousePointer = vbDefault
                    Close #file_num
                    Exit Function
                End If
            Else
                MsgBox strMsg, vbCritical, "發生錯誤！"
                Screen.MousePointer = vbDefault
                Close #file_num
                Exit Function
            End If
        End If
    '案件灰階代表圖或查名超過 300K 秀錯誤
    ElseIf LOF(file_num) > 307200 Then
        If Pub_StrUserSt03 = "M51" Then
            If MsgBox(strMsg & vbCrLf & "是否還是要存入??", vbCritical + vbYesNo, "發生錯誤！") = vbNo Then
                Screen.MousePointer = vbDefault
                Close #file_num
                Exit Function
            End If
        Else
            MsgBox strMsg, vbCritical, "發生錯誤！"
            Screen.MousePointer = vbDefault
            Close #file_num
            Exit Function
        End If
    End If
    FormCheck = True
End Function

'intStatus=1.新增2.修改
Private Function InsUpdImg(ByVal intStatus As Integer, ByVal stIBF05 As String, ByVal stIBF13 As String, ByRef stIBF15 As String, Optional ByVal bolChkColor As Boolean = False) As Boolean
    Dim stSQL As String

On Error GoTo ErrHand
    InsUpdImg = False
    '黑白圖存檔需判斷是否有彩色圖,有需刪彩色圖
    If bolChkColor = True Then
        If ChkImgByteFile(Trim(oCP01), Trim(oCP02), Trim(oCP03), Trim(oCP04), , , "2") = True Then
            '刪除彩色代表圖
            PUB_DelFtpFile2 Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-2", , UCase("ImgByteFile")
            '刪除DB資料
            stSQL = "Delete ImgByteFile Where ibf01='" & oCP01 & "' and ibf02='" & oCP02 & "' and ibf03='" & oCP03 & "' and ibf04='" & oCP04 & "' and ibf05='2' "
            cnnConnection.Execute stSQL
        End If
    End If
    '檔案改放 FTP,必須在DB資料刪除前執行
    PUB_DelFtpFile2 Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-" & stIBF05, , UCase("ImgByteFile")
    If intStatus = 1 Then
        PicRs.AddNew
        PicRs.Fields("ibf07").Value = strUserNum
        PicRs.Fields("ibf08").Value = Val(strSrvDate(1))
        PicRs.Fields("ibf09").Value = Val(Format(time, "HHMM"))
    Else
        PicRs.Fields("ibf10").Value = strUserNum
        PicRs.Fields("ibf11").Value = Val(strSrvDate(1))
        PicRs.Fields("ibf12").Value = Val(Format(time, "HHMM"))
    End If
    PicRs.Fields("ibf01").Value = Trim(oCP01)
    PicRs.Fields("ibf02").Value = Trim(oCP02)
    PicRs.Fields("ibf03").Value = Trim(oCP03)
    PicRs.Fields("ibf04").Value = Trim(oCP04)
    PicRs.Fields("ibf05").Value = stIBF05
    PicRs.Fields("ibf06").Value = IIf(stIBF05 = "2", IIf(IsWmf = True, "4", "2"), IIf(IsWmf = True, "3", "1"))
    PicRs.Fields("ibf13").Value = stIBF13
    Close #file_num
    '檔案放FTP
    stIBF15 = ""
    If IsWmf = False Then
        PUB_PutFtpFile App.path & "\NowPic.jpg", Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-" & m_ibf05, Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-" & stIBF05, stIBF15, UCase("imgbytefile")
    Else
        PUB_PutFtpFile App.path & "\NowPic.wmf", Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-" & m_ibf05, Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-" & stIBF05, stIBF15, UCase("imgbytefile")
    End If
    If stIBF15 <> "" Then
        PicRs.Fields("ibf15") = stIBF15
    End If
    PicRs.UPDATE
    InsUpdImg = True
    Exit Function
    
ErrHand:
    MsgBox "存檔(InsUpdImg)發生錯誤！-" & Err.Description, vbExclamation, "未知錯誤！"
    Screen.MousePointer = vbDefault
End Function

Private Sub ReadTMQ()
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
   
    strQ = "Select TQF06,TQF08,TQF09,TQF10 From TMQFile " & _
               "Where TQF01='" & oCP01 & "' AND TQF02='" & oCP02 & "' AND TQF03='" & oCP03 & "' AND TQF04='" & Format(oCP04, TMQ_附件F04) & "' "
               
    If RsQ.State = adStateOpen Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockOptimistic
    If RsQ.RecordCount <> 0 Then
        RsQ.MoveFirst
        LBL1(1).Caption = CheckStr(RsQ.Fields("TQF08"))
        LBL1(2).Caption = ChangeWStringToTDateString(CheckStr(RsQ.Fields("TQF09").Value))
        LBL1(3).Caption = Format(CheckStr(RsQ.Fields("TQF10")), "00:00")
        LBL1(7).Caption = Format(CheckStr(RsQ.Fields("TQF06")), "###,###,###,###,##0") & " 位元組"
    End If
    RsQ.Close
End Sub

Sub PicToObj2(oFileNameAndPath As String)
    Dim objImg As StdPicture
    Dim nSrcWidth, nSrcHeight, nWidth, nHeight
    Dim tBI      As BITMAP
    InitAll2
    On Error GoTo BE
    If UCase(pvGetExt(oFileNameAndPath)) = "WMF" Or UCase(pvGetExt(oFileNameAndPath)) = "EMF" Then
        Set G_SeekPicColor2.Picture = LoadPicture(App.path & "\TempPic.wmf")
    Else
        Set objImg = pvGetStdPicture(oFileNameAndPath)
        Set m_Image = New cImage
        Set m_Jpeg = New cJpeg
        Set G_SeekPicColor2.Picture = objImg
        If G_SeekPicColor2.Picture <> 0 Then
            Call GetObject(objImg.handle, Len(tBI), tBI)
            If tBI.bmWidth = 0 Or tBI.bmHeight = 0 Then
                MsgBox "發生錯誤！", vbExclamation, "圖檔格式錯誤"
                If oRtPic = False Then
                    StrMenu
                End If
                Exit Sub
            End If
            If tBI.bmWidth > 2000 Or tBI.bmHeight > 2000 Then
                nSrcWidth = tBI.bmWidth
                nSrcHeight = tBI.bmHeight
                If nSrcWidth > nSrcHeight Then
                    nWidth = 1200
                    nHeight = nSrcHeight / (nSrcWidth / nWidth)
                ElseIf nSrcWidth < nSrcHeight Then
                    nHeight = 1200
                    nWidth = nSrcWidth / (nSrcHeight / nHeight)
                Else
                    nHeight = 1200
                    nWidth = 1200
                End If
                pic2.Width = nWidth
                pic2.Height = nHeight
                pic2.BackColor = &H8000000A
                '重新定義大小
                pic2.Scale (0, 0)-(nWidth, nHeight)
                '縮小
                pic2.PaintPicture objImg, 0, 0, nWidth, nHeight, , , , , vbSrcCopy
                '存檔
                SavePicture pic2.Image, App.path & "\TempPic.bmp"
                Set objImg = pvGetStdPicture(App.path & "\TempPic.bmp")
            End If
            m_Image.CopyStdPicture objImg
            m_Jpeg.Quality = 75
            m_Jpeg.SampleHDC m_Image.hDC, m_Image.Width, m_Image.Height
            RidFile App.path & "\TempPic.jpg"
            m_Jpeg.SaveFile App.path & "\TempPic.jpg"
            m_Image.CopyStdPicture pvGetStdPicture(App.path & "\TempPic.jpg")
            Set G_SeekPicColor2.Picture = LoadPicture(App.path & "\TempPic.jpg")

        End If
    End If
    Dim t_hd As Double
    Dim t_wd As Double
    t_hd = G_SeekPicColor2.ScaleHeight / tmpPic2.ScaleHeight
    t_wd = G_SeekPicColor2.ScaleWidth / tmpPic2.ScaleWidth
    If t_hd > t_wd Then
        t_wd = G_SeekPicColor2.ScaleWidth / t_hd
        t_hd = G_SeekPicColor2.ScaleHeight / t_hd
    Else
        t_hd = G_SeekPicColor2.ScaleHeight / t_wd
        t_wd = G_SeekPicColor2.ScaleWidth / t_wd
    End If
        tmpImg2.Width = t_wd
        tmpImg2.Height = t_hd
        tmpImg2.Move (tmpPic2.ScaleWidth - tmpImg2.Width) / 2, (tmpPic2.ScaleHeight - tmpImg2.Height) / 2, t_wd, t_hd
    
    Set tmpImg2.Picture = G_SeekPicColor2.Picture
    Set objImg = Nothing

    Exit Sub
BE:
    Resume Next
End Sub

Sub InitAll2()
    Set tmpPic2.Picture = LoadPicture()
    Set tmpImg2.Picture = LoadPicture()
    Set G_SeekPicColor2.Picture = LoadPicture()
    Set pic2.Picture = LoadPicture()
End Sub

Sub MainBTEnabled(ByVal bolEnabled As Boolean)
    Dim oCmd As CommandButton
    
    If bolEnabled = True Then
        For Each oCmd In cmdOK
            oCmd.Enabled = False
        Next
    Else
        For i = 0 To UBound(SeekCmdok())
            If i >= 4 And i <= 6 Then
                '目前不使用
            Else
                cmdOK(i).Enabled = SeekCmdok(i).Enabled
            End If
        Next i
    End If
End Sub

Private Sub FormClear()
    For Each oLbl In LBL1
        oLbl.Caption = ""
    Next
End Sub

'設定按鈕
Public Sub SetSeekCmdok()
    '查名單輸入m_TMQ會有值;接洽單cmdOK(2)=確定
    If strWorkType <> "1" And Trim(m_TMQ) = MsgText(601) And cmdOK(2).Enabled = True And cmdOK(2).Caption = "存檔(&S)" Then
        cmdOK(8).Visible = True
        cmdOK(8).Enabled = True
        Frame3.Top = 48
        Frame3.Left = 248
    End If
    For i = 0 To UBound(SeekCmdok())
        SeekCmdok(i).Caption = cmdOK(i).Caption
        SeekCmdok(i).Enabled = cmdOK(i).Enabled
        SeekCmdok(i).Visible = cmdOK(i).Visible
    Next i
End Sub

Private Function ChkExists(ByVal stIBF05 As String) As Boolean
    Dim strQ As String
    
    ChkExists = False
    'Modify By Sindy 2018/10/31 TF馬德里商標圖檔-子案的圖同母案
    '固定都以IBF01=tm01 AND IBF02=substr(tm02,1,5)||'0' AND IBF03='0' AND IBF04='00' 去抓代表圖
    If oCP01 = "TF" Then
      strQ = "Select * From ImgByteFile " & _
                 "Where ibf01='" & oCP01 & "' and ibf02='" & Mid(oCP02, 1, 5) & "0" & "' and ibf03='0' And ibf04='00' " & _
                 "And ibf05='" & stIBF05 & "'"
    Else
    '2018/10/31 END
      strQ = "Select * From ImgByteFile " & _
                 "Where ibf01='" & oCP01 & "' and ibf02='" & oCP02 & "' and ibf03='" & oCP03 & "' and ibf04='" & oCP04 & "' " & _
                 "And ibf05='" & stIBF05 & "'"
    End If
    Set PicRs = New ADODB.Recordset
    If PicRs.State = adStateOpen Then PicRs.Close
    PicRs.CursorLocation = adUseClient
    PicRs.Open strQ, cnnConnection, adOpenStatic, adLockOptimistic
    If PicRs.RecordCount > 0 Then
        ChkExists = True
    End If
End Function

'Added by Morgan 2018/8/24
Private Function FormDelete() As Boolean
On Error GoTo ErrHnd
   
   If PUB_DelFtpFile2(Trim(oCP01) & "-" & Trim(oCP02) & "-" & Trim(oCP03) & "-" & Trim(oCP04) & "-" & m_ibf05, , UCase("ImgByteFile")) = True Then
      strSql = "DELETE IMGBYTEFILE WHERE IBF01='" & oCP01 & "' AND IBF02='" & oCP02 & "' AND IBF03='" & oCP03 & "' AND IBF04='" & oCP04 & "' AND IBF05='" & m_ibf05 & "'"
      Pub_SeekTbLog strSql 'Added by Lydia 2024/11/11
      cnnConnection.Execute strSql, intI
      FormDelete = True
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

'Added by Morgan 2023/7/27
Public Function UploadFile(pFilePath As String, pColor As Boolean) As Boolean
On Error GoTo ErrHnd
   If Dir(pFilePath) <> "" Then
      bolCall = True
      If pColor Then
         optColor(1).Value = True
      Else
         optColor(0).Value = True
      End If
      PicToObj pFilePath
      IsSave = False
      IsWmf = False
      PhotoSave
      UploadFile = True
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function
