VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090702_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員工作量查詢"
   ClientHeight    =   5715
   ClientLeft      =   135
   ClientTop       =   975
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton cmd 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7752
      TabIndex        =   0
      Top             =   48
      Width           =   1200
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3060
      TabIndex        =   66
      Top             =   1395
      Width           =   930
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   29
      Left            =   5760
      TabIndex        =   65
      Top             =   4185
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   28
      Left            =   5760
      TabIndex        =   64
      Top             =   3870
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   13
      Left            =   1470
      TabIndex        =   63
      Top             =   4500
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "墨圖作業天數："
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   62
      Top             =   4498
      Width           =   1308
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   15
      Left            =   1470
      TabIndex        =   61
      Top             =   5130
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   14
      Left            =   1290
      TabIndex        =   60
      Top             =   4815
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   12
      Left            =   1470
      TabIndex        =   59
      Top             =   4185
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   16
      Left            =   1650
      TabIndex        =   58
      Top             =   5430
      Width           =   1560
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2752;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   6
      Left            =   1080
      TabIndex        =   57
      Top             =   2325
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   11
      Left            =   1080
      TabIndex        =   56
      Top             =   3870
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   9
      Left            =   1080
      TabIndex        =   55
      Top             =   3270
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   54
      Top             =   2955
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   7
      Left            =   720
      TabIndex        =   53
      Top             =   2640
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   5
      Left            =   1470
      TabIndex        =   52
      Top             =   2010
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   4
      Left            =   1065
      TabIndex        =   51
      Top             =   1710
      Width           =   2010
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3545;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   50
      Top             =   1410
      Width           =   1965
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "3466;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   2
      Left            =   900
      TabIndex        =   49
      Top             =   1095
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   48
      Top             =   465
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   47
      Top             =   780
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   10
      Left            =   900
      TabIndex        =   46
      Top             =   3570
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "草圖作業天數："
      Height          =   180
      Index           =   33
      Left            =   108
      TabIndex        =   45
      Top             =   4188
      Width           =   1308
   End
   Begin VB.Label Label1 
      Caption         =   "取消收文日："
      Height          =   180
      Index           =   29
      Left            =   108
      TabIndex        =   44
      Top             =   4808
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   180
      Index           =   21
      Left            =   108
      TabIndex        =   43
      Top             =   778
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "收文日："
      Height          =   180
      Index           =   20
      Left            =   108
      TabIndex        =   42
      Top             =   1088
      Width           =   744
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   19
      Left            =   108
      TabIndex        =   41
      Top             =   1398
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   18
      Left            =   108
      TabIndex        =   40
      Top             =   1708
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "專利/商標種類："
      Height          =   180
      Index           =   16
      Left            =   108
      TabIndex        =   39
      Top             =   2018
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   15
      Left            =   108
      TabIndex        =   38
      Top             =   2328
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限："
      Height          =   180
      Index           =   14
      Left            =   108
      TabIndex        =   37
      Top             =   2948
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "法定期限："
      Height          =   180
      Index           =   13
      Left            =   108
      TabIndex        =   36
      Top             =   3258
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   12
      Left            =   105
      TabIndex        =   35
      Top             =   3885
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "點數："
      Height          =   180
      Index           =   11
      Left            =   108
      TabIndex        =   34
      Top             =   2638
      Width           =   564
   End
   Begin VB.Label Label1 
      Caption         =   "國外案本所案號："
      Height          =   180
      Index           =   10
      Left            =   108
      TabIndex        =   33
      Top             =   5436
      Width           =   1512
   End
   Begin VB.Label Label1 
      Caption         =   "國外案承辦人："
      Height          =   180
      Index           =   9
      Left            =   108
      TabIndex        =   32
      Top             =   5118
      Width           =   1308
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   8
      Left            =   108
      TabIndex        =   31
      Top             =   3568
      Width           =   744
   End
   Begin VB.Label Label1 
      Caption         =   "繪圖人員："
      Height          =   180
      Index           =   4
      Left            =   108
      TabIndex        =   30
      Top             =   468
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "3."
      Height          =   180
      Index           =   27
      Left            =   5376
      TabIndex        =   29
      Top             =   4188
      Width           =   348
   End
   Begin VB.Label Label1 
      Caption         =   "2."
      Height          =   180
      Index           =   25
      Left            =   5376
      TabIndex        =   28
      Top             =   3878
      Width           =   348
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   26
      Left            =   6150
      TabIndex        =   27
      Top             =   3255
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "墨圖："
      Height          =   180
      Index           =   24
      Left            =   5412
      TabIndex        =   26
      Top             =   3258
      Width           =   696
   End
   Begin VB.Label Label1 
      Caption         =   "草圖："
      Height          =   180
      Index           =   5
      Left            =   5448
      TabIndex        =   25
      Top             =   2948
      Width           =   696
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   21
      Left            =   5685
      TabIndex        =   24
      Top             =   1710
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "墨圖張數："
      Height          =   180
      Index           =   34
      Left            =   4512
      TabIndex        =   23
      Top             =   2328
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "墨圖完稿日："
      Height          =   180
      Index           =   28
      Left            =   4512
      TabIndex        =   22
      Top             =   2018
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "墨圖齊備日："
      Height          =   180
      Index           =   1
      Left            =   4512
      TabIndex        =   21
      Top             =   1708
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "承辦期限："
      Height          =   180
      Index           =   7
      Left            =   4512
      TabIndex        =   20
      Top             =   468
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "是否算案件數："
      Height          =   180
      Index           =   17
      Left            =   4512
      TabIndex        =   19
      Top             =   2638
      Width           =   1308
   End
   Begin VB.Label Label1 
      Caption         =   "草圖齊備日："
      Height          =   180
      Index           =   23
      Left            =   4512
      TabIndex        =   18
      Top             =   778
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "草圖完稿日："
      Height          =   180
      Index           =   26
      Left            =   4512
      TabIndex        =   17
      Top             =   1088
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "草圖張數："
      Height          =   180
      Index           =   2
      Left            =   4512
      TabIndex        =   16
      Top             =   1398
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "修改時數："
      Height          =   180
      Index           =   3
      Left            =   4512
      TabIndex        =   15
      Top             =   3568
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "承辦時數："
      Height          =   180
      Index           =   6
      Left            =   4512
      TabIndex        =   14
      Top             =   2948
      Width           =   924
   End
   Begin VB.Label Label1 
      Caption         =   "備註："
      Height          =   180
      Index           =   31
      Left            =   4512
      TabIndex        =   13
      Top             =   4498
      Width           =   552
   End
   Begin VB.Label Label1 
      Caption         =   "(N：不算)"
      Height          =   180
      Index           =   32
      Left            =   6516
      TabIndex        =   12
      Top             =   2638
      Width           =   1068
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   17
      Left            =   5475
      TabIndex        =   11
      Top             =   465
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   18
      Left            =   5685
      TabIndex        =   10
      Top             =   780
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   19
      Left            =   5685
      TabIndex        =   9
      Top             =   1095
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   24
      Left            =   5865
      TabIndex        =   8
      Top             =   2640
      Width           =   660
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1164;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   22
      Left            =   5685
      TabIndex        =   7
      Top             =   2010
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   23
      Left            =   5475
      TabIndex        =   6
      Top             =   2325
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   20
      Left            =   5475
      TabIndex        =   5
      Top             =   1410
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   25
      Left            =   6180
      TabIndex        =   4
      Top             =   2955
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   27
      Left            =   5790
      TabIndex        =   3
      Top             =   3570
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2805;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   1100
      Index           =   30
      Left            =   5115
      TabIndex        =   2
      Top             =   4500
      Width           =   3780
      Caption         =   "lblFM2"
      Size            =   "6667;1931"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "1."
      Height          =   180
      Index           =   22
      Left            =   5412
      TabIndex        =   1
      Top             =   3568
      Width           =   348
   End
End
Attribute VB_Name = "frm090702_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/07 改成Form2.0 ; lbl1(index)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim i As Integer

Private Sub cmd_Click()
Me.Hide
frm090702_2.Show
'Unload Me
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
StrMenu
End Sub

Sub StrMenu()
'Modify By Cheng 2002/04/29
'引進是否閉卷欄
'strSQL = "SELECT S1.ST02,cp09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),cp18," & SQLDate("cp06") & "," & SQLDate("cp07") & ",S2.ST02,s3.st02,0,0," & SQLDate("CP57") & ",'',''," & SQLDate("cP48") & "," & SQLDate("eP14") & "," & SQLDate("eP15") & ",ep16," & SQLDate("ep17") & "," & SQLDate("ep18") & ",ep19,ep20,ep21,ep22,ep23,ep24,ep25,ep26 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE EP02=CP09(+) AND  PA01=CP01 AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND ep13=S1.ST01(+) AND eP05=S2.ST01(+) AND cP13=S3.ST01(+) AND EP02='" & frm090702_2.StrForm1 & "' "
'92.04.03 nick add left join
'strSQL = "SELECT S1.ST02,cp09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),cp18," & SQLDate("cp06") & "," & SQLDate("cp07") & ",S2.ST02,s3.st02,0,0," & SQLDate("CP57") & ",'',''," & SQLDate("cP48") & "," & SQLDate("eP14") & "," & SQLDate("eP15") & ",ep16," & SQLDate("ep17") & "," & SQLDate("ep18") & ",ep19,ep20,ep21,ep22,ep23,ep24,ep25,ep26,PA57 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE EP02=CP09(+) AND  PA01=CP01 AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND ep13=S1.ST01(+) AND eP05=S2.ST01(+) AND cP13=S3.ST01(+) AND EP02='" & frm090702_2.StrForm1 & "' "
strSql = "SELECT S1.ST02,cp09," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),cp18," & SQLDate("cp06") & "," & SQLDate("cp07") & ",S2.ST02,s3.st02,0,0," & SQLDate("CP57") & ",'',''," & SQLDate("cP48") & "," & SQLDate("eP14") & "," & SQLDate("eP15") & ",ep16," & SQLDate("ep17") & "," & SQLDate("ep18") & ",ep19,ep20,ep21,ep22,ep23,ep24,ep25,ep26,PA57 FROM ENGINEERPROGRESS,CASEPROGRESS,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,PATENTTRADEMARKMAP,PATENT WHERE EP02=CP09(+) AND  CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND ep13=S1.ST01(+) AND eP05=S2.ST01(+) AND cP13=S3.ST01(+) AND EP02='" & frm090702_2.StrForm1 & "' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        For i = 0 To 30
            lbl1(i) = CheckStr(.Fields(i))
        Next i
        'Add By Cheng 2002/04/29
'        If IsNull(.Fields(33).Value) Then
        If IsNull(.Fields(31).Value) Then
            Me.lblClose.Caption = ""
        Else
            Me.lblClose.Caption = "已閉卷"
        End If
        '92.04.03 nick add left join
        'strSQL = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01 AND CP02=CM02 AND CP03=CM03 AND CP04=CM04 AND CP14=ST01(+) AND CP31='Y' and cp09='" & frm090702_2.StrForm1 & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
        strSql = "SELECT CM01||'-'||CM02||'-'||CM03||'-'||CM04,ST02 FROM CASEMAP,CASEPROGRESS,STAFF WHERE CP01=CM01(+) AND CP02=CM02(+) AND CP03=CM03(+) AND CP04=CM04(+) AND CP14=ST01(+) AND CP31='Y' and cp09='" & frm090702_2.StrForm1 & "' order by CM01||'-'||CM02||'-'||CM03||'-'||CM04 "
        CheckOC2
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            lbl1(16) = CheckStr(adoRecordset1.Fields(0))
            lbl1(15) = CheckStr(adoRecordset1.Fields(1))
        Else
            lbl1(16) = ""
            lbl1(15) = ""
        End If
        CheckOC2
        '計算草圖作業天數
        If Len(lbl1(18)) <> 0 And Len(lbl1(19)) <> 0 And Val(lbl1(18)) <> 0 And Val(lbl1(19)) <> 0 Then
            lbl1(12) = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(lbl1(19))), ChangeTStringToWString(ChangeTDateStringToTString(lbl1(18))))
        End If
        '計算墨圖作業天數
        If Len(lbl1(21)) <> 0 And Len(lbl1(22)) <> 0 And Val(lbl1(21)) <> 0 And Val(lbl1(22)) <> 0 Then
            lbl1(13) = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(lbl1(22))), ChangeTStringToWString(ChangeTDateStringToTString(lbl1(21))))
        End If

    End If
End With
CheckOC
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090702_3 = Nothing
End Sub

