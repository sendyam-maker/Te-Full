VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040104_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專發文-授權"
   ClientHeight    =   5724
   ClientLeft      =   672
   ClientTop       =   996
   ClientWidth     =   8544
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5724
   ScaleWidth      =   8544
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   4500
      MaxLength       =   4
      TabIndex        =   18
      Top             =   5340
      Width           =   540
   End
   Begin VB.TextBox txtCP118 
      Height          =   270
      Left            =   5985
      MaxLength       =   1
      TabIndex        =   73
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox txtPayToday 
      Height          =   270
      Left            =   7605
      MaxLength       =   1
      TabIndex        =   15
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox txtChkRltDate 
      Height          =   270
      Left            =   7140
      MaxLength       =   8
      TabIndex        =   19
      Top             =   5370
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   5520
      MaxLength       =   1
      TabIndex        =   68
      Top             =   2700
      Width           =   255
   End
   Begin VB.TextBox txtCP84 
      Height          =   285
      Left            =   3780
      TabIndex        =   4
      Top             =   2700
      Width           =   1092
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   0
      Left            =   2940
      MaxLength       =   60
      TabIndex        =   8
      Top             =   3300
      Width           =   5415
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   1
      Left            =   2940
      MaxLength       =   60
      TabIndex        =   9
      Top             =   3540
      Width           =   5415
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   2
      Left            =   2940
      MaxLength       =   60
      TabIndex        =   10
      Top             =   3780
      Width           =   5415
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   4800
      TabIndex        =   13
      Top             =   4110
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3036
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1560
      TabIndex        =   17
      Top             =   5352
      Width           =   255
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Index           =   1
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   12
      Top             =   4116
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   38
      Top             =   768
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   37
      Top             =   768
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   36
      Top             =   768
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   35
      Top             =   768
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   3
      Left            =   4368
      TabIndex        =   21
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6396
      TabIndex        =   23
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5580
      TabIndex        =   22
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7620
      TabIndex        =   24
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Index           =   0
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   11
      Top             =   4116
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   14
      Top             =   4416
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1200
      MaxLength       =   9
      TabIndex        =   7
      Top             =   3336
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   1020
      MaxLength       =   9
      TabIndex        =   20
      Top             =   5940
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2700
      Width           =   1095
   End
   Begin MSForms.TextBox Text14 
      Height          =   615
      Left            =   1200
      TabIndex        =   16
      Top             =   4710
      Width           =   7215
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "12726;1085"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   600
      Left            =   6888
      TabIndex        =   5
      Top             =   2700
      Width           =   1476
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2603;1058"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   0
      Left            =   1395
      TabIndex        =   0
      Top             =   1875
      Width           =   6930
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "12224;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   1
      Left            =   1395
      TabIndex        =   1
      Top             =   2115
      Width           =   6930
      VariousPropertyBits=   671107099
      MaxLength       =   250
      Size            =   "12224;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   2
      Left            =   1395
      TabIndex        =   2
      Top             =   2355
      Width           =   6930
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "12224;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   18
      Left            =   3600
      TabIndex        =   76
      Top             =   5385
      Width           =   765
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:         (Y:是)"
      Height          =   180
      Index           =   2
      Left            =   4815
      TabIndex        =   75
      Top             =   3030
      Width           =   1995
   End
   Begin VB.Label lblPayToday 
      AutoSize        =   -1  'True
      Caption         =   "電子送件是否當日扣款:         (Y/N)"
      Height          =   180
      Left            =   5670
      TabIndex        =   74
      Top             =   4470
      Width           =   2655
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   6300
      TabIndex        =   72
      Top             =   5385
      Width           =   765
   End
   Begin VB.Label lblCaseFee 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   8130
      TabIndex        =   70
      Tag             =   "Y"
      Top             =   5310
      Width           =   255
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   5970
      TabIndex        =   69
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label lblCP84 
      AutoSize        =   -1  'True
      Caption         =   "發文規費:"
      Height          =   180
      Left            =   2880
      TabIndex        =   67
      Top             =   2745
      Width           =   765
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(中):"
      Height          =   180
      Left            =   240
      TabIndex        =   66
      Top             =   1890
      Width           =   1065
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(英):"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   65
      Top             =   2130
      Width           =   1065
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(日):"
      Height          =   180
      Left            =   240
      TabIndex        =   64
      Top             =   2370
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   8340
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   240
      X2              =   8340
      Y1              =   1704
      Y2              =   1704
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   12
      Left            =   4920
      TabIndex        =   63
      Top             =   540
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2646;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "代理人:"
      Height          =   180
      Index           =   1
      Left            =   4080
      TabIndex        =   62
      Top             =   4110
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   11
      Left            =   6300
      TabIndex        =   61
      Top             =   4110
      Width           =   2070
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3651;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家"
      Height          =   180
      Index           =   1
      Left            =   4080
      TabIndex        =   60
      Top             =   540
      Width           =   720
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "列印通知函:       (N:不印 1:申請人 2:被授權人 3:二者皆印)"
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   59
      Top             =   3030
      Width           =   4425
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否列印指示信         (Y)"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   58
      Top             =   5370
      Width           =   2145
   End
   Begin MSForms.Label Label2 
      Height          =   270
      Index           =   10
      Left            =   2430
      TabIndex        =   57
      Top             =   4440
      Width           =   3090
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5450;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   7080
      TabIndex        =   56
      Top             =   540
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2328;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Left            =   240
      TabIndex        =   55
      Top             =   1500
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Left            =   4080
      TabIndex        =   54
      Top             =   1500
      Width           =   675
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Left            =   240
      TabIndex        =   53
      Top             =   1275
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Left            =   4080
      TabIndex        =   52
      Top             =   1275
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Left            =   240
      TabIndex        =   51
      Top             =   1065
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   4080
      TabIndex        =   50
      Top             =   765
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   49
      Top             =   765
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   0
      Left            =   4080
      TabIndex        =   48
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   240
      TabIndex        =   47
      Top             =   540
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   46
      Top             =   540
      Width           =   2850
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5027;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   4920
      TabIndex        =   45
      Top             =   1065
      Width           =   3450
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "6085;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   2
      Left            =   4920
      TabIndex        =   44
      Top             =   765
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2646;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   3
      Left            =   1080
      TabIndex        =   43
      Top             =   1065
      Width           =   2850
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5027;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   4
      Left            =   4920
      TabIndex        =   42
      Top             =   1275
      Width           =   3450
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "6085;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   5
      Left            =   1080
      TabIndex        =   41
      Top             =   1275
      Width           =   2850
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5027;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   6
      Left            =   4920
      TabIndex        =   40
      Top             =   1500
      Width           =   3450
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "6085;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   7
      Left            =   1080
      TabIndex        =   39
      Top             =   1500
      Width           =   2850
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5027;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   34
      Top             =   4710
      Width           =   765
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "年費通知人:"
      Height          =   180
      Left            =   240
      TabIndex        =   33
      Top             =   4410
      Width           =   945
   End
   Begin VB.Label Label28 
      Caption         =   "~"
      Height          =   255
      Left            =   2400
      TabIndex        =   32
      Top             =   4110
      Width           =   135
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "授權期間:"
      Height          =   180
      Left            =   240
      TabIndex        =   31
      Top             =   4110
      Width           =   765
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "日:"
      Height          =   180
      Left            =   2580
      TabIndex        =   30
      Top             =   3810
      Width           =   225
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "英:"
      Height          =   180
      Left            =   2580
      TabIndex        =   29
      Top             =   3570
      Width           =   225
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "中:"
      Height          =   180
      Left            =   2580
      TabIndex        =   28
      Top             =   3330
      Width           =   225
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "被授權人:"
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   3330
      Width           =   765
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Index           =   0
      Left            =   6480
      TabIndex        =   26
      Top             =   540
      Width           =   585
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   2745
      Width           =   585
   End
   Begin VB.Label lblCaseFees 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   8175
      TabIndex        =   71
      Top             =   5370
      Width           =   255
   End
End
Attribute VB_Name = "frm040104_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (Text15,lstNameAgent,Label2..)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'2005/7/8整理
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/7/14 改用動態陣列
'Dim pa(T_PA) As String, cp(T_CP) As String
Dim pa() As String, cp() As String

Dim intWhere As Integer
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/23 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_Subject As String 'Add by Morgan 2016/5/19
Dim m_bolFMP As Boolean 'Added by Lydia 2023/06/20 是否為FMP案
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/20 是否為寰華

Private Function Process(Index As Integer) As Boolean
'Added by Lydia 2019/12/05
Dim strFilePath As String '記錄智慧局收文文號
Dim strNewCP64 As String '保留進度備註
Dim bolUp As Boolean 'Added by Lydia 2020/03/23 是否需要上傳檔案到卷宗區

   Dim bolChk As Boolean, strTmp As String
   '檢查輸入資料的完整性
   If CheckDataIntegrity = False Then Exit Function
   'Add By Cheng 2002/05/22
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Function
   
    strNewCP64 = Text14.Text 'Added by Lydia 2019/12/05 保留進度備註
    
   'Add by Morgan 2009/3/23 設定是否算發文室案件
   If pa(9) = "000" Then
   
      'Add by Morgan 2015/7/7 電子送件
      If txtCP118 = "Y" Then
         m_CP123s = ""
         'Added by Morgan 2016/5/16 電子送件也要記錄主管機關
         If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text9, , True) = False Then
            Exit Function
         End If
         'end 2016/5/16
         
         strExc(0) = InputBox("請輸入智慧局收文文號!!")
         If strExc(0) = "" Then
            Exit Function
         Else
            'Modified by Lydia 2019/12/05
            'Text14 = "智慧局收文文號:" & strExc(0) & ";" & Text14
            strFilePath = strExc(0)  '記錄智慧局收文文號
            strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & Text14 '保留進度備註
            'end 2019/12/05
         End If
         'end 2012/12/26
         
      Else
      'end 2015/7/7
      
         'Add by Morgan 2009/4/28
         If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text9) = False Then
            Exit Function
         End If
         If m_CP123s = "Y" Then
         'end 2009/4/28
            'modify by sonia 2014/6/23 加傳發文規費, P-108903
            If ModifyDispatch(cp(9), m_CP09s, m_CP123s, txtCP84, Text9) = False Then
                Exit Function
            End If
            
         End If 'Add by Morgan 2015/7/7
         
      End If
      'Add by Amy 2014/10/14 P台灣案發文控制
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
        If pa(1) = "P" And cp(9) < "C" Then
            If cp(9) < "B" Then
                'A類一定要有接洽單才可發文
                'Modify by Amy 2014/11/27 取消ChkOneDayHasCP27判斷,接洽單改檢查,因考慮可能同時發文其他案件性質情形
                'If PUB_CheckPDF2(cp(9), 0, True, strExc(0)) = False And ChkOneDayHasCP27(pa(1), pa(2), pa(3), pa(4), cp(5) + 19110000) = False Then
                If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
                    Exit Function
                End If
            End If
            'AB類申請書確認檢查,符合條件才可發文
            'Modified by Morgan 2015/3/17
            'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" Then
               If PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            'end 2015/3/17
                  MsgBox "無申請書PDF檔 ,不可發文!", vbInformation
                  Exit Function
               End If 'Added by Morgan 2015/3/17
            End If
        End If
      End If
      'end 2014/10/14

   'Added by Morgan 2016/6/29 非臺灣案電子化
   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      If cp(9) < "B" And Left(cp(12), 1) <> "F" Then
          If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
              Exit Function
          End If
      End If
   'end 2016/6/29
   End If
   
    'Added by Lydia 2019/12/05 檢查是否有電子送件的檔案
    bolUp = False 'Added by Lydia 2020/03/23
    If txtCP118.Text = "Y" And strFilePath <> "" And pa(9) = "000" Then
        strExc(1) = cp(82)
        If Val(cp(82)) > 0 Then
            If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                 strExc(1) = ""
            End If
        End If
        If Val(strExc(1)) = 0 Then
           'Modified by Lydia 2020/03/23 改成先判斷是否上傳檔案; ex.P-124220發明申請 因為上傳檔案在FormSave前,所以沒抓到出名代理,造成無POA檔仍設電子檔案齊備CP121=Y
           'If Pub_AutoEsetToCppByP(True, pa(1), pa(2), pa(3), pa(4), pa(8), Label2(0).Caption, cp(10), strFilePath, Text9.Text) = False Then
           If Pub_AutoEsetToCppByP(True, pa(1), pa(2), pa(3), pa(4), pa(8), "", cp(10), strFilePath, Text9.Text) = False Then
                 Exit Function
           End If
           bolUp = True 'Added by Lydia 2020/04/08
        End If
    End If
   
   If pa(9) = 台灣國家代號 Then 'Added by Lydia 2020/01/17 限台灣案
       Text14.Text = strNewCP64 '檢查完畢，更新備註欄位
   End If
   'end 2019/12/05
   
   If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Function
   
   'Move by Lydia 2020/04/07 移到最下方
   
   Process = True
   
   'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
   PUB_CheckEMail Combo2
   PUB_CheckEMail pa(75), pa(144)
   If pa(145) <> "" Then
      PUB_CheckEMail pa(75), pa(145)
   End If
   'end 2008/2/20
   
   'Add by Morgan 2007/6/14
   If pa(9) = "000" Then
      PUB_ReAsignInform pa(1), pa(2), pa(3), pa(4), strReceiveNo
   End If
   
   '2012/7/23 add by sonia
   '台灣案發文規費與收文規費不符時,mail給智權人員
   If txtCP84.Enabled = True And pa(9) = "000" And Val(Me.txtCP84.Text) <> Val(cp(17)) Then
      '2013/7/2 modify by sonia 改用共用module
      'Modified by Morgan 2015/7/7 +傳strCP118參數
      PUB_ChkOfficialFee cp(9), Me.txtCP84.Text, IIf(txtCP118 = "Y", "A", "")
   End If
   '2012/7/23 end
   
   Select Case Text13.Text '通知函
      Case "", " ", "1" '印申請人
          'Modify By Cheng 2003/01/30
          '修改定稿種類
   '               NowPrint strReceiveNo, "2", "00", False, strUserNum, 0
         NowPrint strReceiveNo, "02", "00", False, strUserNum, 0, , , , , , , , , , , , strReceiveNo
      Case "2" '印被授權人
          'Modify By Cheng 2003/01/30
          '修改定稿種類
   '               NowPrint strReceiveNo, "2", "05", False, strUserNum, 0
         NowPrint strReceiveNo, "02", "05", False, strUserNum, 0, , , , , , , , , , , , strReceiveNo
      Case "3" '印申請人及被授權人
          'Modify By Cheng 2003/01/30
          '修改定稿種類
   '               NowPrint strReceiveNo, "2", "00", False, strUserNum, 0
   '               NowPrint strReceiveNo, "2", "05", False, strUserNum, 0
         NowPrint strReceiveNo, "02", "00", False, strUserNum, 0, , , , , , , , , , , , strReceiveNo
         NowPrint strReceiveNo, "02", "05", False, strUserNum, 0
         
         'Added by Morgan 2016/5/26 一個收文號只能對應一封通知函，故提醒User手動上傳給被授權人的通知函(因目前非申請人通知函沒有副本問題)
         If 內專全面電子化啟用日 <= Val(strSrvDate(1)) Then
            If Left(Pub_StrUserSt03, 1) <> "F" Then
               MsgBox "請自行存放被授權人通知函到卷宗區！", vbExclamation
            End If
         End If
         'end 2016/5/26
         
   End Select
   'end 2014/08/21
   
   'Modify By Cheng 2003/01/30
   '若為"Y"要印指示信
   '         If Text5 <> "Y" Then '指示信
   'Modify by Amy 2014/08/21 台灣案電子化 +傳strLetterRecNo
   If Text5 = "Y" Then '指示信
      'Modified by Morgan 2016/5/19
      '指示信電子化
      'NowPrint strReceiveNo, "02", "30", False, strUserNum
      If Left(Pub_StrUserSt03, 1) = "F" Then
         NowPrint strReceiveNo, "02", "30", False, strUserNum
      Else
         NowPrint strReceiveNo, "02", "30", True, strUserNum, , , , , , , , , , , , , strReceiveNo
         frm1105_1.m_RecNo = strReceiveNo
         frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".DATA.PDF"
         frm1105_1.m_Subject = m_Subject
         frm1105_1.Show
      End If
   End If
   
   'Add By Cheng 2002/04/30
   '若有未發文資料顯示警告
   PUB_GetCPunIssueDatas "" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)

   'Added by Lydia 2020/03/23 是否可以上傳檔案,前面已判斷
   'Move by Lydia 2020/04/07 從中間移下來
   If bolUp = True Then
      If Pub_AutoEsetToCppByP(False, pa(1), pa(2), pa(3), pa(4), pa(8), Label2(0).Caption, cp(10), strFilePath, Text9.Text) = False Then
           Exit Function
      End If
   End If
   'end 2020/03/23
   
End Function

Private Sub cmdok_Click(Index As Integer)
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   
   Select Case Index
      Case 0, 3
         'Modify by Morgan 2010/2/10 改呼叫函數方式以便鎖定按鍵
         cmdOK(Index).Enabled = False
         If Not Process(Index) Then
            cmdOK(Index).Enabled = True
         Else
            If Index = 0 Then
               ' 90.07.11 modify by louis (回第一個畫面清除)
               'Add By Sindy 2013/5/20
               If frm040104_1.bolIsEMPFlow = True Then
                  Unload frm040104_1
                  frm090202_4.Show
                  frm090202_4.QueryData
               Else
               '2013/5/20 End
                  frm040104_1.Show
                  frm040104_1.Clear
               End If
            Else
               ' 90.07.11 modify by louis (回第一個畫面重新查詢)
               'Add By Sindy 2013/5/20
               If frm040104_1.bolIsEMPFlow = True Then
                  frm090202_4.QueryData
               End If
               '2013/5/20 End
               frm040104_1.Show
               frm040104_1.ReQuery
            End If
            Unload Me
         End If
         
      Case 1
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            frm040104_1.Show
         End If
         Unload Me
         
      Case 2
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            Unload frm040104_1
         End If
         Unload Me
   End Select
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
End Sub

Private Function FormSave() As Boolean

   Dim i As Integer, strTxt(1 To 20) As String, ii As Integer
   Dim stCP118 As String, stCP152 As String 'Added by Morgan 2015/7/7
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrorHandler


   'Added by Morgan 2013/6/7 自 lstNameAgent_Validate 移來,否則若觸發 Form_Activate 事件會跑 ReadPatent 導致 cp(110) 被清除
   cp(110) = ""
   If lstNameAgent.Visible = True Then
      For ii = 0 To lstNameAgent.ListCount - 1
         If lstNameAgent.Selected(ii) = True Then
            'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
            'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
            'Modified by Morgan 2021/12/14f Forms2.0 改用模組
            'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
            cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         End If
      Next
      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   End If
   'end 2013/6/7
   
   'Modified by Morgan 2017/1/11 從下面移上來(更新pa76前就要補滿)
   ' 90.07.18 modify by louis (被授權人補滿九碼)
   If IsEmptyText(Text6) = False Then
      Text6 = Text6 & String(9 - Len(Text6), "0")
   End If
   If IsEmptyText(Text11) = False Then
      Text11 = Text11 & String(9 - Len(Text11), "0")
   End If
   'end 2017/1/11
   
   strTxt(1) = "UPDATE PATENT SET PA05=" & CNULL(Text15(0)) & ",PA06=" & CNULL(ChgSQL(Text15(1))) & _
      ",PA07=" & CNULL(Text15(2)) & ",PA76=" & CNULL(Text11) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
    'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(1)
             
   If Combo2 <> "" Then
      'Modify by Morgan 2008/2/22
      'cp(44) = ChangeCustomerL(Combo2)
      intI = InStr(Combo2, "-")
      If intI > 0 Then
         cp(44) = Left(Combo2, intI - 1)
         cp(116) = Mid(Combo2, intI + 1)
      Else
         cp(44) = Combo2
         cp(116) = ""
      End If
      cp(44) = ChangeCustomerL(cp(44))
      'end 2008/2/22
      
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.GetCaseThatCode(cp) Then cp(45) = ""
      If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   Else
      cp(44) = ""
      cp(116) = ""
      cp(45) = ""
   End If
   

      
   'Added by Morgan 2015/7/7
   stCP118 = txtCP118
   stCP152 = ""
   If txtCP118 = "Y" And Val(txtCP84) > 0 And pa(9) = "000" Then
      If txtPayToday <> "" Then
         stCP118 = "A"
         If txtPayToday = "Y" Then
            stCP152 = CompWorkDay(2, DBDATE(Text9))
         Else
            stCP152 = CompWorkDay(3, DBDATE(Text9))
         End If
      End If
   End If
   'end 2015/7/7
      
   ' 91.03.25 modify by louis (單引號)
   'Modify by morgan 2004/8/11 加 cp84
   'Modify by Morgan 2005/7/14 加 cp110
   'Modify by Morgan 2008/2/22 +cp116
   'Modified by Morgan 2015/7/7 +CP118,CP152
   'Modified by Lydia 2021/05/25 +CP113工作時數
   strTxt(2) = "UPDATE CASEPROGRESS SET CP22=" & CNULL(ChgSQL(Text7)) & ",CP27=" & TransDate(Text9, 2) & "," & _
      "cp14=" & CNULL(Text10) & ",CP44=" & CNULL(cp(44)) & ",CP116=" & CNULL(cp(116)) & ",CP45=" & CNULL(cp(45)) & "," & _
      "cp50=" & CNULL(Text8(0)) & ",CP51=" & CNULL(ChgSQL(Text8(1))) & ",CP52=" & CNULL(Text8(2)) & "," & _
      "cp64=" & CNULL(ChgSQL(Text14)) & ",cp72=" & CNULL(ChgSQL(Text6)) & ",cp53=" & TransDate(Text12(0), 2) & "," & _
      "cp54=" & TransDate(Text12(1), 2) & ", cp84=" & Format(Val(txtCP84.Text)) & ",cp110=" & CNULL(cp(110)) & _
      ",cp118='" & stCP118 & "',cp152=" & CNULL(stCP152, True) & " ,cp113=" & CNULL(txtCP113, True) & _
      " where cp09='" & strReceiveNo & "'"
    'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(2)
      
   i = 2
   strExc(0) = "SELECT CF23 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & cp(10) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields(0)) Or RsTemp.Fields(0) <> 0 Then
         i = 3
            'Modify By Cheng 2003/12/08
            '若本所期限非工作天則抓最近的工作天
'         strTxt(i) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'            "NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & _
'            "','" & pa(3) & "','" & pa(4) & "'," & 收達 & "," & _
'            CompDate(2, rsTemp.Fields(0), TransDate(Text9, 2)) & "," & _
'            CompDate(2, rsTemp.Fields(0), TransDate(Text9, 2)) & ",'" & _
'            strUserNum & "'," & objPublicData.GetNextProgressNo & ")"
         'edit by nickc 2007/02/02 不用 dll 了
         'strTxt(i) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
            "NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pA(1) & "','" & pA(2) & _
            "','" & pA(3) & "','" & pA(4) & "'," & 收達 & "," & _
            PUB_GetWorkDay1(CompDate(2, rsTemp.Fields(0), TransDate(Text9, 2)), True) & "," & _
            CompDate(2, rsTemp.Fields(0), TransDate(Text9, 2)) & ",'" & _
            strUserNum & "'," & objPublicData.GetNextProgressNo & ")"
         strTxt(i) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
            "NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & _
            "','" & pa(3) & "','" & pa(4) & "'," & 收達 & "," & _
            PUB_GetWorkDay1(CompDate(2, RsTemp.Fields(0), TransDate(Text9, 2)), True) & "," & _
            CompDate(2, RsTemp.Fields(0), TransDate(Text9, 2)) & ",'" & _
            strUserNum & "'," & GetNextProgressNo & ")"
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(i)
      End If
   End If
   
   
    'Modify by Amy 2014/09/09 for 台灣案電子化
    If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
        If cp(9) < "C" And pa(9) = 台灣國家代號 Then
            cnnConnection.Execute "delete LetterProgress where lp01='" & strReceiveNo & "'", intI 'Added by Morgan 2016/2/26 可能會重新發文
            '*沒出客戶通知函
            If Trim(Text13) = "N" Then
                'Modify by Amy 2015/02/13 原:同一天 沒 其他有規費的發文
                  '1.    電子送件且規費>0,有收據
                  '2.非電子送件且經發文室要計件,有回執
                'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'                strExc(1) = PUB_GetLetterJudge(pa(1), cp(10))
'                If cp(118) = "Y" Then
'                     If Val(txtCP84) > 0 Then
'                        PUB_AddLetterProgress strReceiveNo, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                     End If
'                Else
'                    If Left(m_CP123s, 1) = "Y" Then
'                        PUB_AddLetterProgress strReceiveNo, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                    End If
'                End If
                
            '*有出客戶通知函
            Else
                'Modified by Morgan 2018/8/1
                'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , , pa(1), pa(2), pa(3), pa(4))
                strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10))
                'Modify by Amy 2015/02/13 修改、整理判斷條件
                  '1.　電子送件有規費的有收據；無規費的無回執
                  '2.非電子送件要計件的有回執；不計件的無回執
                'Modify by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'                If cp(118) = "Y" Then
'                    If Val(txtCP84) > 0 Then
'                        PUB_AddLetterProgress strReceiveNo, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
'                    Else
'                        PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
'                    End If
'                Else
                    If Left(m_CP123s, 1) = "Y" Then
                        PUB_AddLetterProgress strReceiveNo, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
                    Else
                        PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
                    End If
'                End If
'                'end 2015/02/13
                'end 2015/03/06
            End If
         'Added by Morgan 2016/5/18
         '指示信電子化
         ElseIf pa(9) <> 台灣國家代號 And Left(Pub_StrUserSt03, 1) <> "F" Then
            If Text5 = "Y" Then
               If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
                  m_Subject = "請代為提出" & GetPrjState4(cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4), cp(10)) & "申請" & IIf(cp(45) <> "", " Y/R:" & cp(45) & ";", "") & " O/R:" & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
                  'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
                  strExc(2) = PUB_GetLetterJudgeNew("2", pa(1), cp(10), pa(9))
                  PUB_AddAppForm strReceiveNo, True, strExc(2), m_Subject '不轉檔,自行判發
               End If
            End If
            'end 2016/5/18
            
            'Added by Morgan 2016/5/26
            '客戶通知函
            If 內專全面電子化啟用日 <= Val(strSrvDate(1)) Then
               If Text13 <> "N" Then
                  'Modified by Morgan 2018/8/1
                  'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , pa(9), pa(1), pa(2), pa(3), pa(4))
                  strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10), pa(9), , , IIf(Left(cp(12), 1) = "F", True, False))
                  PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75)
               End If
            End If
            'end 2016/5/26
            
         End If
      End If
      'end 2014/09/09
   
   'Add by Morgan 2009/3/23
   If pa(9) = "000" Then
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130
      'Add by Amy 2015/02/13 更新收據/回執設定
      'Modify by Amy 2015/3/06 +發文日參數
      PUB_UpdateLP19 cp(1), cp(2), cp(3), cp(4), m_CP09s, m_CP123s, Text9
   End If
   
   'Add by Morgan 2009/8/4
   If txtChkRltDate <> "" Then
      PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9), cp(10), cp(43)
   End If

   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   
End Function

Private Sub Combo2_Click()
   Combo2_Validate False
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Dim strNo As String, iPos As Integer
   If Combo2.Text = "" Then
      If pa(9) <> 台灣國家代號 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Cancel = True
         Exit Sub
      End If
      
   ElseIf Not ChgType(12) Then
      Cancel = True
      
   Else
      strNo = Combo2.Text
      
      'Add by Morgan 2008/2/22 加聯絡人判斷
      iPos = InStr(strNo, "-")
      If iPos > 0 Then
         strNo = Left(strNo, iPos - 1)
      End If
      'end 2008/2/22
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If PUB_CheckStatus(strNo) = False Then Cancel = True
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   With frm040104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
    'Add by Morgan 2005/7/14
    ReDim pa(TF_PA)
    ReDim cp(TF_CP)
    
   ReadPatent
   
   cp(110) = "" '要清空,否則若重新發文會殘留前次發文資料,當新案有改出名人而本程序未改選將會造成不一致 Added by Morgan 2012/9/7
   
   'Add by Morgan 2005/7/14
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   Text7.Visible = False
   lstNameAgent.Clear
   If pa(9) = "000" Then
      PUB_SetOurAgent lstNameAgent, pa(), cp(110), , True  'Modified by Morgan 2021/12/14 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2005/7/14 END
   
   Label2(0) = strReceiveNo
   
   'Added by Morgan 2017/1/11
   '專利處人員操作時年費通知人欄位鎖住以避免不小心改到(目前只有外專人員會設定)
   If Left(Pub_StrUserSt03, 1) = "P" Then
      Text11.Locked = True
   End If
   'end 2017/1/11
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2009/8/17
   'Set frm040104_6 = Nothing 'Removed by Morgan 2021/12/14 form2.0會有問題，改在呼叫時清除記憶體變數

End Sub

Private Sub ReadPatent()
Dim Lbl As Object, txt As Object, i As Integer
Dim m_Fee As String         '銷帳服務費 2012/8/1 add by sonia
Dim m_Official As String    '銷帳規費   2012/8/1 add by sonia
   
   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      For i = 3 To 7
         If pa(i + 23) <> "" Then ChgType (i)
      Next
      Text15(0) = pa(5)
      Text15(1) = pa(6)
      Text15(2) = pa(7)
      
      If pa(76) <> "" Then Text11 = pa(76): ChgType (11)
      
      If pa(9) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(pA(9), strExc(0)) Then Label2(12) = strExc(0)
         If ClsPDGetNation(pa(9), strExc(0)) Then Label2(12) = strExc(0)
      End If
   End If
   
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(13) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(13), strExc(0)) Then Label2(1) = strExc(0)
         If ClsPDGetStaff(cp(13), strExc(0)) Then Label2(1) = strExc(0)
      End If
      Label2(2) = cp(6)
      If cp(27) = "" Then
         Text9 = strSrvDate(2)
      Else
         Text9 = cp(27)
      End If
      'Added by Lydia 2023/06/20 判斷FCP案,寰華案
      If Left(cp(12), 1) = "F" And pa(9) <> "000" Then
         m_bolFMP = True
      Else
         m_bolFMP = False
      End If
      m_bolFMP2 = False
      If m_bolFMP = True Then
         m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4))
      End If
      'end 2023/06/20
      If cp(14) <> "" Then Text10 = cp(14): ChgType (10)
      Text14 = cp(64)
      If cp(72) <> "" Then Text6 = cp(72): ChgType (8)
      '2005/7/8 MODIFY BY SONIA
      'Text12(0) = cp(53)
      'Text12(1) = cp(54)
      Text8(0) = cp(50)
      Text8(1) = cp(51)
      Text8(2) = cp(52)
      If pa(9) = "000" Then
         Text12(0) = cp(53)
         Text12(1) = cp(54)
      Else
         Text12(0) = TransDate(cp(53), 2)
         Text12(1) = TransDate(cp(54), 2)
      End If
      '2005/7/8 END
      Text7 = cp(22)
      
      'Added by Morgan 2015/7/7
      txtCP118 = ""
      If cp(118) <> "" Then txtCP118 = "Y"
      
      'Added by Morgan 2024/1/19 大陸案預設電子送件--郭
      'Removed by Morgan 2024/1/30 改分案預設--郭
      'If pa(9) = "020" Then
      '   'Modified by Morgan 2024/1/25 有設定大陸P案要公文正本者預設紙本送件--郭
      '   If PUB_GetCustomerValue(pa(26), "CU182") = "Y" Then
      '      txtCP118 = ""
      '   Else
      '      txtCP118 = "Y"
      '   End If
      'End If
      'end 2024/1/19
      
      txtPayToday = ""
      If txtCP118 = "Y" And pa(9) = "000" Then
         If Val(ServerTime) <= 153000 Then
            txtPayToday = "Y"
         End If
      End If
      'end 2015/7/7
'13
      'Modify by Morgan 2008/10/16 +若進度檔已有代理人則預設
      'Modified by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設 => cp(9), pa(9), pa(26)
      AddAgent Combo2, cp, , cp(44), cp(116), cp(9), pa(9), pa(26)
   End If
   '2012/8/1 add by sonia 若有銷帳則要扣除銷帳規費
   If Val(cp(77)) > 0 Then
      If GetCP77Detail(cp(9), m_Fee, m_Official) = True Then
         cp(17) = cp(17) - m_Official
      End If
   End If
   '2012/8/1 end
    
   'Add by Morgan 2004/9/8 檢查是否有延期
   If PUB_ChkDelay(cp(9)) = True Then cp(17) = "0"

   'Add by Morgan 2004/8/11
   txtCP84.Tag = cp(17)
   
   'Add by Morgan 2004/11/18
   If pa(9) = "020" Then
      Text13.Text = "1"
   Else
      Text13.Text = ""
   End If
   
   'Add by Morgan 2009/8/5
   If Text9 <> "" Then
      PUB_SetChkResultDate pa(1), pa(9), cp(10), Text9, txtChkRltDate, cp, pa(8)
      Text9.Tag = Text9
   End If
   
    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 0
         '2011/12/8 MODIFY BY SONIA 發文日可輸系統日的下一個工作日
         'If Not ChkDate(Text9) Or Val(Text9.Text) > Val(strSrvDate(2)) Then
         '   MsgBox "發文日期不正確或發文日大於系統日，請重新輸入 !", vbCritical
         If Not ChkDate(Text9) Or DBDATE(Val(Text9.Text)) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
            MsgBox "發文日期不正確或發文日大於系統日下一個工作日，請重新輸入 !", vbCritical
         '2011/12/8 END
         Else
            ChgType = True
         End If
      Case 3, 4, 5, 6, 7
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(pa(i + 23), strTempName) Then
         If ClsLawLawGetName(pa(i + 23), strTempName) Then
            Label2(i) = strTempName
            ChgType = True
         End If
      Case 8
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.GetCusCAJnam(Text6.Text, strExc(0), strExc(1), strExc(2)) = True Then
         If ClsLawGetCusCAJnam(Text6.Text, strExc(0), strExc(1), strExc(2)) = True Then
            For i = 0 To 2
               Text8(i) = strExc(i)
            Next
            ChgType = True
         End If
      Case 10
         'Added by Lydia 2023/06/20 寰華案:承辦人為外專程序時,改為操作人員
         If m_bolFMP2 = True Then
            Text10 = GetFCPUser(Text10)
         End If
         'end 2023/06/20
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(Text10, strTempName) Then
         If ClsPDGetStaff(Text10, strTempName) Then
            Label2(9) = strTempName
            ChgType = True
         End If
      Case 11
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(Text11, strTempName) Then
         If ClsLawLawGetName(Text11, strTempName) Then
            Label2(10) = strTempName
            ChgType = True
         End If
      Case 12 '代理人
         strExc(1) = Combo2.Text
         'Add by Morgan 2008/2/22 加判斷是否為聯絡人
         If InStr(strExc(1), "-") > 0 Then
            If ClsPDGetContact(strExc(1), strTempName) Then
               Combo2 = strExc(1)
               Label2(11) = strTempName
               ChgType = True
            End If
            
         '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
         ElseIf PUB_GetAgentName(pa(1), strExc(1), strTempName) = True Then
            Combo2.Text = strExc(1)
            Label2(11).Caption = strTempName
            ChgType = True
         Else
            Label2(11).Caption = ""
         End If
   End Select
End Function

Private Sub Text10_GotFocus()
  TextInverse Text10
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then
      If Not ChgType(10) Then Cancel = True
   End If
End Sub

Private Sub Text11_GotFocus()
  TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   '2005/7/8 MODIFY BY SONIA 加判斷有輸入才檢查
   If Text11 <> "" Then
      If Not ChgType(11) Then Cancel = True
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If Cancel = False Then
         If PUB_CheckStatus(Text11.Text) = False Then Cancel = True
      End If
   End If
End Sub

Private Sub Text12_GotFocus(Index As Integer)
   TextInverse Text12(Index)
End Sub

Private Sub Text12_Validate(Index As Integer, Cancel As Boolean)
   
   If Text12(Index) = "" Then
      'Modify by Morgan 2004/11/18 改在存檔前檢查以免輸入錯誤時無法跳離
      'MsgBox "授權期間不可空白，請重新輸入 !", vbCritical
      'Cancel = True
   Else
      If Not ChkDate(Text12(Index)) Then
         MsgBox "授權期間不正確，請重新輸入 !", vbCritical
         Cancel = True
      Else
         If Index = 1 Then
            If pa(25) = "" Then
               MsgBox "專用期間止日不正確，請重新輸入 !", vbCritical
               Cancel = True
            Else
               'Modify by Morgan 2004/11/18 改西元
               'If Val(Text12(1)) > Val(pa(25)) Then
               If Val(DBDATE(Text12(1))) > Val(DBDATE(pa(25))) Then
                  MsgBox "授權期間止日大於專用期間止日，請重新輸入 !", vbCritical
                  Cancel = True
               Else
                  If ChkRange(Text12(0), Text12(1), "授權期間") = False Then Cancel = True
               End If
            End If
         Else
            If pa(24) = "" Then
               MsgBox "專用期間起日不正確，請重新輸入 !", vbCritical
               Cancel = True
            Else
               'Modify by Morgan 2004/11/18 改西元
               'If Val(pa(24)) > Val(Text12(0)) Then
               If Val(DBDATE(pa(24))) > Val(DBDATE(Text12(0))) Then
                  MsgBox "授權期間起日小於專用期間起日，請重新輸入 !", vbCritical
                  Cancel = True
               End If
            End If
         End If
      End If
   End If
   If Cancel = True Then TextInverse Text12(Index)
End Sub

Private Sub Text13_GotFocus()
  TextInverse Text13
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text14_GotFocus()
  TextInverse Text14
End Sub

Private Sub Text15_GotFocus(Index As Integer)
    'Modify By Cheng 2002/10/28
    Select Case Index
    Case 0
        Me.Text15(Index).SelStart = 0
        Me.Text15(Index).SelLength = 0
    Case Else
        TextInverse Text15(Index)
    End Select
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 <> "" Then If Not ChgType(8) Then Cancel = True
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text8_GotFocus(Index As Integer)
  TextInverse Text8(Index)
End Sub

Private Sub Text8_Validate(Index As Integer, Cancel As Boolean)
   If Index = 2 Then
      If Text8(0) = "" And Text8(1) = "" And Text8(2) = "" Then
         MsgBox "被授權人名稱不可同時空白 !"
         Cancel = True
      End If
   End If
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> "" Then
      If Not ChgType(0) Then
         Cancel = True
      Else
         'Add by Morgan 2009/8/5
         If Text9.Tag <> Text9 Then
            PUB_SetChkResultDate pa(1), pa(9), cp(10), Text9, txtChkRltDate, cp, pa(8)
            Text9.Tag = Text9
            'Added by Morgan 2015/7/7
            '當發文日有改時,電子送件案要人工輸入是否當日扣款
            If txtCP118 = "Y" Then
               txtPayToday.Text = ""
            End If
            'end 2015/7/7
         End If
      End If
   Else
      MsgBox "發文日不可空白 !", vbCritical
      Cancel = True
   End If
End Sub

'Add By Cheng 2002/03/08
Private Function CheckDataIntegrity() As Boolean
Dim Cancel As Boolean
   'add by nickc 2008/05/01
   If IsDebt(pa(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        GoTo IntegrityOrNot
   End If
   
Cancel = False

'檢查代理人欄位
Combo2_Validate Cancel
If Cancel = True Then
   Me.Combo2.SetFocus
   GoTo IntegrityOrNot
End If

CheckDataIntegrity = True
Exit Function

IntegrityOrNot:
CheckDataIntegrity = False
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False


   'Added by Morgan 2021/12/14 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/14
   
If Me.Text10.Enabled = True Then
   Cancel = False
   Text10_Validate Cancel
   If Cancel = True Then
      Me.Text10.SetFocus
      Text10_GotFocus
      Exit Function
   End If
End If

If Me.Text11.Enabled = True Then
   Cancel = False
   Text11_Validate Cancel
   If Cancel = True Then
      Me.Text11.SetFocus
      Text11_GotFocus
      Exit Function
   End If
End If

For Each objTxt In Text12
   If objTxt.Enabled = True Then
      Cancel = False
      Text12_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.Text12(objTxt.Index).SetFocus
         Text12_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next

If Me.Text6.Enabled = True Then
   Cancel = False
   Text6_Validate Cancel
   If Cancel = True Then
      Me.Text6.SetFocus
      Text6_GotFocus
      Exit Function
   End If
End If

For Each objTxt In Text8
   If objTxt.Enabled = True Then
      Cancel = False
      Text8_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.Text8(objTxt.Index).SetFocus
         Text8_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next

If Me.Text9.Enabled = True Then
   Cancel = False
   Text9_Validate Cancel
   If Cancel = True Then
      Me.Text9.SetFocus
      Text9_GotFocus
      Exit Function
   End If
End If

'Add by Morgan 2004/8/11
If txtCP84.Enabled = True Then
   Cancel = False
   txtCP84_Validate Cancel
   If Cancel = True Then
      txtCP84.SetFocus
      txtCP84_GotFocus
      Exit Function
   End If
End If

'Add by Morgan 2004/9/14
If Combo2.Enabled = True Then
   Cancel = False
   Combo2_Validate Cancel
   If Cancel = True Then
      Combo2.SetFocus
      Exit Function
   End If
End If

'Add by Morgan 2004/11/19
If Text12(0) = "" Then
   MsgBox "授權期間不可空白，請重新輸入 !", vbCritical
   Text12(0).SetFocus
   Exit Function
End If
If Text12(1) = "" Then
   MsgBox "授權期間不可空白，請重新輸入 !", vbCritical
   Text12(1).SetFocus
   Exit Function
End If
'2004/11/9 end

   'Add by Morgan 2005/7/14
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   '2005/7/14 END

'Added by Morgan 2015/7/7
If txtCP118 = "Y" And pa(9) = "000" Then
   If Text7 = "N" Then
      MsgBox "電子送件不可不出名！", vbCritical
      Exit Function
   ElseIf Val(txtCP84) > 0 Then
      If txtPayToday = "" Then
         MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
         txtPayToday.SetFocus
         Exit Function
      End If
   End If
End If
'end 2015/7/7
   
'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
If Pub_ChkACS112isNull(pa(1), pa(2), pa(3), pa(4), txtCP113) = True Then
      txtCP113.SetFocus
      txtCP113_GotFocus
      Exit Function
End If
'end 2021/05/25
   
'Added by Morgan 2024/1/19
If pa(9) = "020" And txtCP118 = "" Then
   If MsgBox("請確認本案是否為紙本送件？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      txtCP118.SetFocus
      Exit Function
   End If
End If
'end 2024/1/19

TxtValidate = True
End Function

'Added by Lydia 2021/05/25
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtCP118_GotFocus()
   TextInverse txtCP118
   CloseIme
End Sub

Private Sub txtCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

'Add by Morgan 2004/8/11
Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_Validate(Cancel As Boolean)
   '台灣
   If pa(9) = "000" Then
      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            txtCP84.Tag = txtCP84.Text
         Else
            txtCP84_GotFocus
            Cancel = True
         End If
      End If
   End If
End Sub
'Add by Morgan 2005/7/14
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   cp(110) = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         'Modified by Morgan 2021/12/14f Forms2.0 改用模組
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         
         bolCheck = True
      End If
   Next
   If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   If bolCheck = True Then
      Text7 = ""
   Else
      Text7 = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub
'Add by Morgan 2009/8/4
Private Sub lblCaseFee_Click()
   frm12040102_2.txtCF(1) = cp(1)
   frm12040102_2.txtCF(2) = pa(9)
   frm12040102_2.txtCF(3) = cp(10)
   frm12040102_2.Show vbModal
   If Val(Text9) > 0 Then
      PUB_SetChkResultDate pa(1), pa(9), cp(10), Text9, txtChkRltDate, cp, pa(8)
   End If
End Sub

Private Sub lblCaseFee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblCaseFee, lblCaseFees
End Sub

Private Sub lblCaseFee_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblCaseFee, lblCaseFees
End Sub

Private Sub txtChkRltDate_Validate(Cancel As Boolean)
   If txtChkRltDate <> "" Then
      If ChkDate(txtChkRltDate) = False Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtPayToday_GotFocus()
   TextInverse txtPayToday
   CloseIme
End Sub

Private Sub txtPayToday_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
