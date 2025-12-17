VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180104 
   BorderStyle     =   1  '單線固定
   Caption         =   "簽核人員異動作業"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Tag             =   "加班資料"
   Begin VB.CommandButton cmdSave 
      Caption         =   "修改簽核人員(&S)"
      Height          =   360
      Left            =   7440
      TabIndex        =   32
      Top             =   960
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   645
      Left            =   420
      TabIndex        =   71
      Top             =   2100
      Visible         =   0   'False
      Width           =   2535
      Begin VB.ComboBox cboETime 
         Height          =   300
         ItemData        =   "frm180104.frx":0000
         Left            =   1410
         List            =   "frm180104.frx":0002
         Style           =   2  '單純下拉式
         TabIndex        =   73
         Top             =   330
         Width           =   1005
      End
      Begin VB.ComboBox cboSTime 
         Height          =   300
         ItemData        =   "frm180104.frx":0004
         Left            =   1410
         List            =   "frm180104.frx":0006
         Style           =   2  '單純下拉式
         TabIndex        =   72
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "迄日下班時段："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   75
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "起日上班時段："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   74
         Top             =   30
         Width           =   1260
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   70
      Text            =   "(藍色字標題的欄位為可查詢的條件)"
      Top             =   120
      Width           =   2985
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清除(&C)"
      Height          =   330
      Left            =   6420
      TabIndex        =   30
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   330
      Left            =   5580
      TabIndex        =   29
      Top             =   30
      Width           =   800
   End
   Begin VB.CheckBox Chk1Day 
      Caption         =   "非整日"
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame02 
      BorderStyle     =   0  '沒有框線
      Height          =   705
      Left            =   60
      TabIndex        =   53
      Top             =   4290
      Visible         =   0   'False
      Width           =   1875
      Begin VB.TextBox txtB101213 
         Enabled         =   0   'False
         Height          =   315
         Left            =   750
         MaxLength       =   4
         TabIndex        =   26
         Top             =   360
         Width           =   705
      End
      Begin VB.TextBox txtB1030 
         Enabled         =   0   'False
         Height          =   315
         Left            =   750
         MaxLength       =   4
         TabIndex        =   25
         Top             =   30
         Width           =   705
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "假日-共                     時"
         Height          =   180
         Left            =   60
         TabIndex        =   55
         Top             =   390
         Width           =   1725
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "實際時數"
         Height          =   180
         Left            =   60
         TabIndex        =   54
         Top             =   90
         Width           =   720
      End
   End
   Begin VB.TextBox txtB1005_1 
      Height          =   300
      Left            =   960
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1004 
      Height          =   300
      Left            =   1410
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1290
      Width           =   945
   End
   Begin VB.TextBox txtB1005_2 
      Height          =   300
      Left            =   1770
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1007_2 
      Height          =   300
      Left            =   4170
      MaxLength       =   2
      TabIndex        =   12
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1007_1 
      Height          =   300
      Left            =   3300
      MaxLength       =   2
      TabIndex        =   11
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1001 
      Height          =   300
      Left            =   960
      MaxLength       =   10
      TabIndex        =   0
      Top             =   90
      Width           =   945
   End
   Begin VB.TextBox txtB1018 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   420
      Width           =   1845
   End
   Begin VB.TextBox txtB1006 
      Height          =   300
      Left            =   3810
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1290
      Width           =   945
   End
   Begin VB.Frame Frame03 
      BorderStyle     =   0  '沒有框線
      Height          =   885
      Left            =   5190
      TabIndex        =   49
      Top             =   990
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txtB1014 
         Height          =   315
         Left            =   540
         MaxLength       =   1
         TabIndex        =   8
         Top             =   120
         Width           =   225
      End
      Begin MSForms.TextBox txtB1015 
         Height          =   285
         Left            =   540
         TabIndex        =   9
         Top             =   450
         Width           =   2535
         VariousPropertyBits=   679495707
         ScrollBars      =   3
         Size            =   "4471;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "差程："
         Height          =   180
         Left            =   30
         TabIndex        =   52
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "(1:長程 2:短程 3:大陸 4:國外)"
         Height          =   180
         Left            =   780
         TabIndex        =   51
         Top             =   180
         Width           =   2235
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "地點："
         Height          =   180
         Left            =   30
         TabIndex        =   50
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.ComboBox CboB1002 
      Height          =   300
      ItemData        =   "frm180104.frx":0008
      Left            =   960
      List            =   "frm180104.frx":000A
      TabIndex        =   2
      Top             =   690
      Width           =   1695
   End
   Begin VB.TextBox txtB1003 
      Height          =   300
      Left            =   960
      MaxLength       =   6
      TabIndex        =   1
      Top             =   390
      Width           =   645
   End
   Begin VB.ComboBox CboB1008 
      Height          =   300
      ItemData        =   "frm180104.frx":000C
      Left            =   3300
      List            =   "frm180104.frx":000E
      TabIndex        =   3
      Top             =   690
      Width           =   1515
   End
   Begin VB.CommandButton cmdagainSend 
      Caption         =   "重送(&R)"
      Height          =   330
      Left            =   7260
      TabIndex        =   31
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "刪除(&D)"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7620
      TabIndex        =   28
      Top             =   420
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改(&M)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8130
      TabIndex        =   27
      Top             =   1770
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8100
      TabIndex        =   33
      Top             =   30
      Width           =   800
   End
   Begin VB.Frame Frame01 
      BorderStyle     =   0  '沒有框線
      Height          =   495
      Left            =   3090
      TabIndex        =   56
      Top             =   2070
      Width           =   1965
      Begin VB.TextBox txtB1010 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1110
         MaxLength       =   5
         TabIndex        =   14
         Top             =   30
         Width           =   525
      End
      Begin VB.TextBox txtB1009 
         Enabled         =   0   'False
         Height          =   315
         Left            =   270
         MaxLength       =   3
         TabIndex        =   13
         Top             =   30
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "共              日               時"
         Height          =   180
         Left            =   60
         TabIndex        =   57
         Top             =   90
         Width           =   1845
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm180104.frx":0010
      Height          =   1485
      Left            =   4710
      TabIndex        =   34
      Top             =   3960
      Width           =   4215
      _ExtentX        =   7444
      _ExtentY        =   2611
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSForms.Label Label26 
      Height          =   195
      Left            =   240
      TabIndex        =   77
      Top             =   5520
      Width           =   8325
      VariousPropertyBits=   27
      Caption         =   "CREATE :                                                    UPDATE : "
      Size            =   "14684;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBoss 
      Height          =   285
      Index           =   4
      Left            =   3750
      TabIndex        =   23
      Top             =   3630
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBoss 
      Height          =   285
      Index           =   2
      Left            =   6510
      TabIndex        =   21
      Top             =   3330
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBoss 
      Height          =   285
      Index           =   1
      Left            =   3750
      TabIndex        =   20
      Top             =   3330
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBoss 
      Height          =   285
      Index           =   0
      Left            =   1050
      TabIndex        =   19
      Top             =   3330
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBoss 
      Height          =   285
      Index           =   3
      Left            =   1050
      TabIndex        =   22
      Top             =   3630
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboEmp 
      Height          =   285
      Index           =   2
      Left            =   6480
      TabIndex        =   17
      Top             =   2520
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboEmp 
      Height          =   285
      Index           =   1
      Left            =   6480
      TabIndex        =   16
      Top             =   2220
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboEmp 
      Height          =   285
      Index           =   0
      Left            =   6480
      TabIndex        =   15
      Top             =   1920
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1207 
      Height          =   1515
      Left            =   990
      TabIndex        =   24
      Top             =   3960
      Width           =   3675
      VariousPropertyBits=   -1466939361
      ScrollBars      =   3
      Size            =   "6482;2672"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1011 
      Height          =   465
      Left            =   1050
      TabIndex        =   18
      Top             =   2820
      Width           =   7875
      VariousPropertyBits=   -1466939365
      ScrollBars      =   3
      Size            =   "13891;820"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1003_2 
      Height          =   285
      Left            =   1650
      TabIndex        =   76
      Top             =   420
      Width           =   1605
      VariousPropertyBits=   679495711
      ScrollBars      =   3
      Size            =   "2831;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "審核主管5："
      Height          =   180
      Left            =   2760
      TabIndex        =   69
      Top             =   3660
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(共7碼)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   5
      Left            =   2400
      TabIndex        =   68
      Top             =   1350
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人(3)："
      Height          =   180
      Left            =   5190
      TabIndex        =   67
      Top             =   2550
      Width           =   1290
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "簽核意見："
      Height          =   180
      Left            =   30
      TabIndex        =   66
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "表單編號："
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   65
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人(2)："
      Height          =   180
      Left            =   5190
      TabIndex        =   64
      Top             =   2250
      Width           =   1290
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "目前表單狀態："
      Height          =   180
      Left            =   3555
      TabIndex        =   63
      Top             =   420
      Width           =   1260
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   3150
      X2              =   4950
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2850
      TabIndex        =   62
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "時"
      Height          =   180
      Left            =   3930
      TabIndex        =   61
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      Height          =   180
      Index           =   2
      Left            =   3390
      TabIndex        =   60
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "            迄"
      Height          =   180
      Left            =   3570
      TabIndex        =   59
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "分"
      Height          =   180
      Left            =   4770
      TabIndex        =   58
      Top             =   1710
      Width           =   180
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   930
      X2              =   4980
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "表單類別："
      Height          =   180
      Left            =   30
      TabIndex        =   48
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "請假人員："
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   47
      Top             =   420
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "假別："
      Height          =   180
      Left            =   2730
      TabIndex        =   46
      Top             =   750
      Width           =   540
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   960
      X2              =   2760
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "分"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   2400
      TabIndex        =   45
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "時間：                 起                    "
      Height          =   180
      Left            =   390
      TabIndex        =   44
      Top             =   1020
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   17
      Left            =   990
      TabIndex        =   43
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "時"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   1560
      TabIndex        =   42
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人(1)："
      Height          =   180
      Left            =   5190
      TabIndex        =   41
      Top             =   1980
      Width           =   1290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "事由："
      Height          =   180
      Left            =   390
      TabIndex        =   40
      Top             =   2850
      Width           =   540
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "審核主管1："
      Height          =   180
      Left            =   30
      TabIndex        =   39
      Top             =   3360
      Width           =   990
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "審核主管2："
      Height          =   180
      Left            =   2760
      TabIndex        =   38
      Top             =   3360
      Width           =   990
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "審核主管3："
      Height          =   180
      Left            =   5490
      TabIndex        =   37
      Top             =   3360
      Width           =   990
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "審核主管4："
      Height          =   180
      Left            =   30
      TabIndex        =   36
      Top             =   3660
      Width           =   990
   End
End
Attribute VB_Name = "frm180104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/8/8
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
Dim m_B1003 As String
Public m_B1017 As String
Dim m_B1018 As String
Dim m_B1019 As String
Dim m_B1023 As String
Dim m_ABS001_1 As String
Dim m_ABS001_2 As String
Dim m_ABS001_3 As String
Dim m_cboEmp(3) As String
Dim m_cboBoss(5) As String
Dim i As Integer, j As Integer, k As Integer
Dim strSubject As String, strContent As String
Dim m_B1004 As String, m_B1005_1 As String, m_B1005_2 As String, m_B1006 As String
Dim m_B1007_1 As String, m_B1007_2 As String, m_B1028 As String, m_B1029 As String
Dim m_B1008 As String, m_B1014 As String
Dim bolChk As Boolean
Dim dblPrevRow As Double
Dim m_BossNum As Integer


Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("簽核人員", "身份", "日期", "時間", "簽核結果", "B1104")
   arrGridHeadWidth = Array(1050, 600, 800, 800, 800, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Public Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCon As String
   
   If txtB1001 = "" Then
      If txtB1003 = "" Then
         If txtB1003 = "" Or txtB1004 = "" Then
            MsgBox "請輸入查詢條件！"
            txtB1003.SetFocus
            Exit Sub
         End If
      End If
   End If
   
   m_EditMode = 0
   Screen.MousePointer = vbHourglass
   
   If txtB1001 <> "" Then
      strCon = " and B1001='" & Me.txtB1001 & "' "
   ElseIf txtB1003 <> "" And txtB1004 <> "" And txtB1005_1 <> "" And txtB1005_2 <> "" Then
      strCon = " and B1003='" & txtB1003 & "' " & _
               " and B1004=" & DBDATE(txtB1004) & _
               " and B1005=" & txtB1005_1 & Format("00" & txtB1005_2, "00")
   ElseIf txtB1003 <> "" And txtB1004 <> "" Then
      strCon = " and B1003='" & txtB1003 & "' " & _
               " and B1004=" & DBDATE(txtB1004)
   ElseIf txtB1003 <> "" Then
      strCon = " and B1003='" & Me.txtB1003 & "' "
   End If
   
   txtB1001.Tag = "" 'Add By Sindy 2023/2/24
   '出缺勤電子簽核主檔，只查詢簽核中的表單
   'Modify By Sindy 2016/12/27 +,B1030
   strSql = "Select B1001,B1002,B1003,B1004,substr(ltrim(to_char('0000'||to_char(B1005),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),3,2) B1005,B1006,substr(ltrim(to_char('0000'||to_char(B1007),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),3,2) B1007,B1008||' '||AC03 B1008,B1009,B1010,B1011,B1012,B1013,B1014,B1015,B1016,B1017," & B1018CName & " B1018,B1019,B1020,B1021,B1022,B1023,B1024,B1025,B1026,B1027,substr(ltrim(to_char('0000'||to_char(B1028),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1028),'0000')),3,2) B1028,substr(ltrim(to_char('0000'||to_char(B1029),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1029),'0000')),3,2) B1029,B1030 " & _
            "From ABS010, allcode " & _
            "Where ac01(+)='04' and B1008=ac02(+) and B1019 is null " & strCon & _
            " order by b1026,b1027,b1001 asc "
'            "and B1001='" & Me.txtB1001 & "'  "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   m_B1003 = "": m_B1017 = "": m_B1018 = "": m_B1019 = ""
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("B1001")) Then
         txtB1001 = rsTmp.Fields("B1001")
         txtB1001.Tag = txtB1001.Text 'Add By Sindy 2023/2/24
      End If
      If Not IsNull(rsTmp.Fields("B1002")) Then CboB1002 = GetB1002Value(rsTmp.Fields("B1002"))
      Call CboB1002_Click
      If Not IsNull(rsTmp.Fields("B1003")) Then txtB1003 = rsTmp.Fields("B1003"): m_B1003 = rsTmp.Fields("B1003"): txtB1003_2 = GetPrjSalesNM(rsTmp.Fields("B1003"))
      If Not IsNull(rsTmp.Fields("B1004")) Then txtB1004 = ChangeWStringToTString(rsTmp.Fields("B1004"))
      If Not IsNull(rsTmp.Fields("B1005")) Then txtB1005_1 = Left(rsTmp.Fields("B1005"), 2): txtB1005_2 = Right(rsTmp.Fields("B1005"), 2)
      If Not IsNull(rsTmp.Fields("B1006")) Then txtB1006 = ChangeWStringToTString(rsTmp.Fields("B1006"))
      If Not IsNull(rsTmp.Fields("B1007")) Then txtB1007_1 = Left(rsTmp.Fields("B1007"), 2): txtB1007_2 = Right(rsTmp.Fields("B1007"), 2)
      If Not IsNull(rsTmp.Fields("B1008")) Then CboB1008 = Trim(rsTmp.Fields("B1008"))
      If Not IsNull(rsTmp.Fields("B1009")) Then txtB1009 = rsTmp.Fields("B1009")
      If Not IsNull(rsTmp.Fields("B1010")) Then txtB1010 = rsTmp.Fields("B1010")
      If Not IsNull(rsTmp.Fields("B1011")) Then txtB1011 = rsTmp.Fields("B1011")
'      If Not IsNull(rsTmp.Fields("B1012")) Then txtB1012 = rsTmp.Fields("B1012")
'      If Not IsNull(rsTmp.Fields("B1013")) Then txtB1013 = rsTmp.Fields("B1013")
      
      'Add By Sindy 2021/8/13
      SetB102829Combo cboSTime, 1, txtB1004, txtB1003
      SetB102829Combo cboETime, 2, txtB1004, txtB1003
      '2021/8/13 END
      
      'Add By Sindy 2016/12/26
      If Not IsNull(rsTmp.Fields("B1012")) Then
         Label16.Caption = "平日-共                     時"
         txtB101213.Text = rsTmp.Fields("B1012")
      ElseIf Not IsNull(rsTmp.Fields("B1013")) Then
         Label16.Caption = "假日-共                     時"
         txtB101213.Text = rsTmp.Fields("B1013")
      End If
      If Not IsNull(rsTmp.Fields("B1030")) Then
         txtB1030 = rsTmp.Fields("B1030")
      Else
         txtB1030 = txtB101213
      End If
      '2016/12/26 END
      If Not IsNull(rsTmp.Fields("B1014")) Then txtB1014 = rsTmp.Fields("B1014")
      If Not IsNull(rsTmp.Fields("B1015")) Then txtB1015 = rsTmp.Fields("B1015")
      'If Not IsNull(rsTmp.Fields("B1016")) Then m_B1016 = rsTmp.Fields("B1016")
      If Not IsNull(rsTmp.Fields("B1017")) Then m_B1017 = rsTmp.Fields("B1017")
      If Not IsNull(rsTmp.Fields("B1018")) Then txtB1018 = rsTmp.Fields("B1018"): Call GetB1018CodeOrCName(m_B1018, rsTmp.Fields("B1018"))
      If Not IsNull(rsTmp.Fields("B1028")) And rsTmp.Fields("B1028") <> "00:00" Then
         For i = 0 To cboSTime.ListCount - 1
            If cboSTime.List(i) = Format(Format(rsTmp.Fields("B1028"), "hhmm"), "00:00") Then
               cboSTime.ListIndex = i
               Exit For
            End If
         Next i
'         Label1(3).Visible = True
'         cboSTime.Visible = True
         Frame1.Visible = True
         Chk1Day.Value = 1
      Else
'         Label1(3).Visible = False
'         cboSTime.Visible = False
         Frame1.Visible = False
         Chk1Day.Value = 0
      End If
      If Not IsNull(rsTmp.Fields("B1029")) And rsTmp.Fields("B1029") <> "00:00" Then
         For i = 0 To cboETime.ListCount - 1
            If cboETime.List(i) = Format(Format(rsTmp.Fields("B1029"), "hhmm"), "00:00") Then
               cboETime.ListIndex = i
               Exit For
            End If
         Next i
'         Label1(4).Visible = True
'         cboETime.Visible = True
         Frame1.Visible = True
         Chk1Day.Value = 1
      Else
'         Label1(4).Visible = False
'         cboETime.Visible = False
         Frame1.Visible = False
         Chk1Day.Value = 0
      End If
'      If (cboSTime.Visible = True And cboETime.Visible = True) Or rsTmp.Fields("B1002") = "02" Or Left(rsTmp.Fields("B1008"), 2) = "08" Then
'         Chk1Day.Value = 1
'      Else
'         Chk1Day.Value = 0
'      End If
      'Modify By Sindy 2012/4/13
      'If rsTmp.Fields("B1002") = 表單類別_加班 Or rsTmp.Fields("B1002") = 表單類別_出差 Then
      If rsTmp.Fields("B1002") = 表單類別_加班 Then
         Chk1Day.Value = 1
      End If
      
      If IsNull(rsTmp.Fields("B1019")) Then
'         txtB1008_2.Visible = True
'         'txtB1008_2 = GetCurrSpecRestDay(Trim(txtB1003))
      Else
         m_B1019 = rsTmp.Fields("B1019")
'         txtB1008_2.Visible = False
      End If
      
      'Add By Sindy 2015/11/19
      '記錄計算完畢當時的日期及時間,方便比對是否有需要重新計算
      m_B1004 = Val(txtB1004)
      m_B1005_1 = Val(txtB1005_1)
      m_B1005_2 = Val(txtB1005_2)
      m_B1006 = Val(txtB1006)
      m_B1007_1 = Val(txtB1007_1)
      m_B1007_2 = Val(txtB1007_2)
      m_B1014 = txtB1014
      m_B1028 = Val(Format(cboSTime.Text, "hhmm"))
      m_B1029 = Val(Format(cboETime.Text, "hhmm"))
      m_B1008 = Left(CboB1008, 2)
      '2015/11/19 END
      Call UpdateCUID(rsTmp)
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Call cmdClear_Click
      Exit Sub
   End If
   rsTmp.Close
   
   Call SetCtrlReadOnly(False, False)
   SetABS001_1Combo txtB1003
   SetABS001_2Combo txtB1003
   
   '先清空欄位值
   Call ClearFieldCbo
   
   m_BossNum = 0
   If Trim(txtB1001) <> "" Then
      '表單流程備註檔
      SetABS012TextBox txtB1207, txtB1001
      '表單簽核檔
      strSql = "SELECT ST02||nvl(B1108,'') 簽核人員," & B1102CName & " 身份,sqldateT(B1105) 日期,sqltime6(B1106) 時間," & B1107CName & " 簽核結果,B1104 FROM ABS011,Staff WHERE B1101='" & txtB1001 & "' and B1104=ST01(+) order by B1102,B1103 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         Set GRD1.Recordset = rsTmp
         For i = 1 To GRD1.Rows - 1
            If GRD1.TextMatrix(i, 1) = "職代" Then
               For j = 0 To CboEmp.UBound
                  If CboEmp(j).Text = "" Then
                     CboEmp(j).Text = SetCboStaffName(GRD1.TextMatrix(i, 5)): m_cboEmp(j) = GRD1.TextMatrix(i, 5)
                     'If GRD1.TextMatrix(i, 2) <> "" Then CboEmp(j).Enabled = False
                     If GRD1.TextMatrix(i, 2) <> "" And GRD1.TextMatrix(i, 2) <> "11/11/11" Then CboEmp(j).Enabled = False
                     Exit For
                  End If
               Next j
            ElseIf GRD1.TextMatrix(i, 1) = "審核主管" Then
               m_BossNum = m_BossNum + 1
               For j = 0 To CboBoss.UBound
                  If CboBoss(j).Text = "" Then
                     CboBoss(j).Text = SetCboStaffName(GRD1.TextMatrix(i, 5)): m_cboBoss(j) = GRD1.TextMatrix(i, 5)
                     'If GRD1.TextMatrix(i, 2) <> "" Then CboBoss(j).Enabled = False
                     If GRD1.TextMatrix(i, 2) <> "" And GRD1.TextMatrix(i, 2) <> "11/11/11" Then CboBoss(j).Enabled = False
                     Exit For
                  End If
               Next j
            End If
         Next i
      End If
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
'   If rsTmp.RecordCount > 0 Then
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
'   Me.Enabled = True
   
   'Add By Sindy 2023/2/24
   cmdDel.Enabled = False
   cmdDel.Visible = False
   '開放人事處可以刪除未簽核完畢的假單
   If GetStaffDepartment(strUserNum) = "M21" And m_B1019 = "" Then
      cmdDel.Enabled = True
      cmdDel.Visible = True
   End If
   '2023/2/24 END
   
'   '人事處已簽收,簽核完畢不可再異動資料
'   If m_B1019 <> "" Then
'      cmdModify.Enabled = False
'      cmdDel.Enabled = False
'      cmdSend.Enabled = False
'      cmdagainSend.Enabled = False
'   Else
      '尚未送簽核
'      If m_B1017 = "" Then
'         cmdModify.Enabled = False
'         If m_B1018 = 主管代填 Then
'            cmdDel.Enabled = True
'         Else
'            cmdDel.Enabled = False
'         End If
'         cmdSend.Enabled = True
'         cmdagainSend.Enabled = False
'         m_EditMode = 1 '新增
'         Call SetCtrlReadOnly(True, False)
'      Else
'         '下一處理人員=自己
'         If m_B1017 = strUserNum And m_B1003 = strUserNum Then
'            cmdModify.Enabled = True
'            cmdDel.Enabled = True
'         '下一處理人員<>自己
'         Else
'            cmdModify.Enabled = False
'            cmdDel.Enabled = False
'         End If
'         cmdSend.Enabled = False
'         If m_B1018 = 送人事處簽收 Then
'            cmdagainSend.Enabled = False
'         Else
'            'Modify By Sindy 2011/10/11
'            'cmdagainSend.Enabled = True
'            cmdagainSend.Enabled = False
'         End If
'      End If
'   End If
'
'   '檢查人事系統裡是否已有表單編號
'   If ChkPerSysB1001Exist(txtB1001, txtB1003) = True Then
'      cmdModify.Enabled = False
'      cmdDel.Enabled = False
'      Call SetCtrlReadOnly(False, False)
'      m_EditMode = 0
'   End If
   
'   If Left(Trim(CboB1002), 2) <> "01" Then txtB1008_2.Visible = False
   
'   'Modify By Sindy 2011/10/11 鎖住職代及審核主管
'   For i = 0 To CboEmp.UBound
'      CboEmp(i).Enabled = False
'   Next i
'   For i = 0 To CboBoss.UBound
'      CboBoss(i).Enabled = False
'   Next i
   
   cmdagainSend.Enabled = True
'   'Add By Sindy 2012/12/13 #10105257 游經理想指定職務代理人
'   '檢查簽核流程資料, 有簽核的最後一道流程結果=退回時, 職代及審核主管欄位均可修改
'   '且重送按鍵鎖住, 修改簽核人員按鍵亮起來
'   cmdSave.Enabled = False
'   If rsTmp.State = adStateOpen Then
'      rsTmp.Close
'   End If
'   strSql = "select * from abs011 where b1101='" & txtB1001 & "' and b1105 is not null order by b1102 desc,b1103 desc"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      If "" & rsTmp.Fields("b1107") = "2" Then '退回當事者
'         For j = 0 To CboEmp.UBound
'            CboEmp(j).Enabled = True
'         Next j
'         For j = 0 To CboBoss.UBound
'            CboBoss(j).Enabled = True
'         Next j
'         cmdagainSend.Enabled = False
'         cmdSave.Enabled = True
'      End If
'   End If
'   '2012/12/13 End
   
   '檢查是否有權限
   If ChkLimitsIsOk() = False Then
      'Add By Sindy 2012/11/15
      If (strUserNum = Left(Trim(CboBoss(0).Text), 5) And CboBoss(0).Enabled = True) Or _
         (strUserNum = Left(Trim(CboBoss(1).Text), 5) And CboBoss(1).Enabled = True) Or _
         (strUserNum = Left(Trim(CboBoss(2).Text), 5) And CboBoss(2).Enabled = True) Or _
         (strUserNum = Left(Trim(CboBoss(3).Text), 5) And CboBoss(3).Enabled = True) Or _
         (strUserNum = Left(Trim(CboBoss(4).Text), 5) And CboBoss(4).Enabled = True) Then
         '該表單的簽核主管也有權限修改 (例如:簽核主管的職代,但事後要改回原簽核主管)
      Else
      '2012/11/15 End
         MsgBox "無權限異動此人員表單!", vbExclamation + vbOKOnly
         cmdagainSend.Enabled = False
      End If
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub ClearFieldCbo()
   For i = 0 To CboEmp.UBound
      m_cboEmp(i) = Empty
      CboEmp(i) = Empty
      CboEmp(i).Enabled = True
   Next i
   For i = 0 To CboBoss.UBound
      m_cboBoss(i) = Empty
      CboBoss(i) = Empty
      CboBoss(i).Enabled = True
   Next i
End Sub

Private Sub ClearField()
   txtB1001 = Empty
   CboB1002 = Empty
   txtB1003 = Empty 'strUserNum
   txtB1003_2 = Empty 'strUserName
   txtB1004 = Empty
   txtB1006 = Empty
'   Chk1Day.Value = 0 '整日
'   Call Chk1Day_Click
   CboB1008 = Empty
'   txtB1008_2 = Empty
   txtB1009 = Empty
   txtB1010 = Empty
   txtB1011 = Empty
   txtB1030 = Empty
   txtB101213 = Empty
   txtB1014 = Empty
   txtB1015 = Empty
   txtB1018 = Empty
   txtB1207 = Empty
   GRD1.Clear
   For i = GRD1.Rows - 1 To 2 Step -1
      GRD1.RemoveItem i
   Next i
   SetGrd
   Call ClearFieldCbo
   
   txtB1005_1 = Empty
   txtB1005_2 = Empty
   txtB1007_1 = Empty
   txtB1007_2 = Empty
   txtB1004.Enabled = True
   txtB1005_1.Enabled = True
   txtB1005_2.Enabled = True
   cmdagainSend.Enabled = False
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean, bolModify As Boolean)
   CboB1002.Enabled = bEnable
   CboB1008.Enabled = bEnable
   txtB1004.Enabled = bEnable
   txtB1005_1.Enabled = bEnable
   txtB1005_2.Enabled = bEnable
   txtB1006.Enabled = bEnable
   txtB1007_1.Enabled = bEnable
   txtB1007_2.Enabled = bEnable
'   txtB1009.Enabled = bEnable
'   txtB1010.Enabled = bEnable
   txtB1011.Enabled = bEnable
   txtB1014.Enabled = bEnable
   txtB1015.Enabled = bEnable
   Chk1Day.Enabled = bEnable
   cboSTime.Enabled = bEnable
   cboETime.Enabled = bEnable
   'Modify By Sindy 2011/10/11
'   If bolModify = False Then
'      For i = 0 To CboEmp.UBound
'         CboEmp(i).Enabled = True
'      Next i
'      For i = 0 To CboBoss.UBound
'         CboBoss(i).Enabled = True
'      Next i
'   End If
'   If bEnable = True Then
'      Call Chk1Day_Click
'   End If
End Sub

Private Sub Chk1Day_LostFocus()
   Call GetCountDayHour(False)
End Sub

Private Sub cmdClear_Click()
   '清空欄位值
   ClearField
End Sub

'Add By Sindy 2023/2/24
Private Sub cmdDel_Click()
Dim strTo As String
Dim rsTmp As New ADODB.Recordset
Dim bolConn As Boolean
   
On Error GoTo ErrHand
   
   If txtB1001.Tag <> txtB1001.Text Then
      MsgBox "查詢有誤！" & vbCrLf & "畫面上表單編號(" & txtB1001.Text & ")和查詢結果的資料(" & txtB1001.Tag & ")不一致！" & vbCrLf & "請重新查詢，再刪除！", vbExclamation
      Exit Sub
   End If
   If MsgBox("確定是否要刪除資料？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then Exit Sub
   
   strTo = GetBossB1107_All(txtB1001)
   strTo = strTo & IIf(strTo <> "", ";", "") & m_B1017 'Add By Sindy 2024/6/14
   strContent = GetEMailContent(txtB1001, strSubject, "刪除", , , , , , , 3)
   
   Screen.MousePointer = vbHourglass
   
   cmdDel.Enabled = False
   cnnConnection.BeginTrans: bolConn = True
   
   '出缺勤主檔
   strSql = "DELETE FROM ABS010 WHERE B1001='" & txtB1001 & "' "
   Pub_SeekTbLog strSql '記錄刪除Log
   cnnConnection.Execute strSql
   
   '簽核檔
   strSql = "SELECT * FROM ABS011 WHERE B1101='" & txtB1001 & "' order by B1102,B1103 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            strSql = "DELETE FROM ABS011 WHERE B1101='" & txtB1001 & "' and B1102='" & rsTmp.Fields("B1102") & "' and B1103=" & rsTmp.Fields("B1103")
            Pub_SeekTbLog strSql '記錄刪除Log
            cnnConnection.Execute strSql
            .MoveNext
         Loop
      End With
   End If
   rsTmp.Close
   
   '流程備註檔
   strSql = "SELECT * FROM ABS012 WHERE B1201='" & txtB1001 & "' order by B1202 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            strSql = "DELETE FROM ABS012 WHERE B1201='" & txtB1001 & "' and B1202=" & rsTmp.Fields("B1202")
            Pub_SeekTbLog strSql '記錄刪除Log
            cnnConnection.Execute strSql
            .MoveNext
         Loop
      End With
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
   cnnConnection.CommitTrans: bolConn = False
   
   Screen.MousePointer = vbDefault
   Call cmdClear_Click
   cmdDel.Visible = False
   
   '發E-Mail通知已簽核的職代和審核主管
   'Add By Sindy 2024/6/14 加發下一處理人員
   If strTo <> "" Then
      PUB_SendMail strUserNum, txtB1003 & ";" & strTo, "", strSubject, strContent, , , , , , , , , , True
   End If
   
   Exit Sub
   
ErrHand:
   cmdDel.Enabled = True
   If bolConn = True Then cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdExit_Click()
'   frm180101.Hide
'   frm180101.QueryData
'   frm180101.Show
   Unload Me
End Sub

Private Sub cmdModify_Click()
   m_EditMode = 2 '修改
   Call SetCtrlReadOnly(True, False)
   cmdagainSend.Enabled = True 'Add By Sindy 2011/10/11
End Sub

Private Sub cmdQuery_Click()
   Call QueryData
End Sub

'Add By Sindy 2012/12/13 單純只是修改簽核人員
Private Sub cmdSave_Click()
Dim strOldB1017 As String
Dim strUpdDate As String, strUpdTime As String, strB1207 As String
Dim strText As String
Dim bolSave As Boolean
   
On Error GoTo ErrHand
   
   bolSave = False
   '檢查條件
   If TxtValidate = False Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   
   cnnConnection.BeginTrans
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   strText = ""
   
   '2.審核主管
   For i = CboBoss.UBound To 0 Step -1
      If CboBoss(i).Enabled = True Then '欄位為可使用狀態
         If m_cboBoss(i) <> Left(Trim(CboBoss(i)), 5) Then '比對原資料與目前資料是否相同,不同時才更新
            strText = "；" & IIf(m_cboBoss(i) = "", "新增", GetPrjSalesNM(m_cboBoss(i)) & "->") & GetPrjSalesNM(Left(Trim(CboBoss(i)), 5)) & strText
            If m_cboBoss(i) <> "" Then '原始資料有值時,先刪除
               bolSave = True
               strSql = "delete From ABS011 where B1101=" & CNULL(txtB1001) & " and B1102='2' and B1104=" & CNULL(m_cboBoss(i))
               cnnConnection.Execute strSql
            End If
            If Left(CboBoss(i), 5) <> "" Then '新增目前資料
               bolSave = True
               'Modify By Sindy 2015/11/13 + B1109
               strSql = "insert into ABS011 (B1101,B1102,B1103,B1104,B1109) values(" & CNULL(txtB1001) & ",'2'," & (i + 1) & "," & CNULL(Left(CboBoss(i), 5)) & "," & CNULL(Left(CboBoss(i), 5)) & ")"
               cnnConnection.Execute strSql
            End If
         End If
      End If
   Next i
   '1.職代
   For i = CboEmp.UBound To 0 Step -1
      If CboEmp(i).Enabled = True Then '欄位為可使用狀態
         If m_cboEmp(i) <> Left(Trim(CboEmp(i)), 5) Then '比對原資料與目前資料是否相同,不同時才更新
            strText = "；" & IIf(m_cboEmp(i) = "", "新增", GetPrjSalesNM(m_cboEmp(i)) & "->") & GetPrjSalesNM(Left(Trim(CboEmp(i)), 5)) & strText
            If m_cboEmp(i) <> "" Then '原始資料有值時,先刪除
               bolSave = True
               strSql = "delete From ABS011 where B1101=" & CNULL(txtB1001) & " and B1102='1' and B1104=" & CNULL(m_cboEmp(i))
               cnnConnection.Execute strSql
            End If
            If Left(CboEmp(i), 5) <> "" Then  '新增目前資料
               bolSave = True
               strSql = "insert into ABS011 (B1101,B1102,B1103,B1104,B1108) values(" & CNULL(txtB1001) & ",'1'," & (i + 1) & "," & CNULL(Left(CboEmp(i), 5)) & ",'(代" & GetPersonSeqno(txtB1003, Left(CboEmp(i), 5)) & ")')"
               cnnConnection.Execute strSql
            End If
         End If
      End If
   Next i
   
   '記錄修改訊息
   If bolSave = True And strText <> "" Then
      strB1207 = "修改簽核人員" & Right(strText, Len(strText) - 1)
      strSql = GetInsertABS012Sql(Trim(txtB1001), strUserNum, strUpdDate, strUpdTime, "", strB1207)
      cnnConnection.Execute strSql
      cnnConnection.CommitTrans
   End If
   
   If bolSave = True Then
      MsgBox "執行完畢！", vbExclamation + vbOKOnly
      Call cmdClear_Click
   Else
      MsgBox "無異動資料！", vbExclamation + vbOKOnly
   End If
   
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox " 修改簽核人員失敗！" & vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
'   Me.txtB1001.BackColor = &H8000000F
'   Me.txtB1003.BackColor = &H8000000F
   Me.txtB1003_2.BackColor = &H8000000F
   Me.txtB1018.BackColor = &H8000000F
'   Me.txtB1008_2.BackColor = &H8000000F
   Me.txtB1030.BackColor = &H8000000F
   Me.txtB101213.BackColor = &H8000000F
   Me.txtB1009.BackColor = &H8000000F
   Me.txtB1010.BackColor = &H8000000F
   
   Me.CboB1002.BackColor = &H8000000F
   Me.CboB1008.BackColor = &H8000000F
   Me.txtB1006.BackColor = &H8000000F
   Me.txtB1007_1.BackColor = &H8000000F
   Me.txtB1007_2.BackColor = &H8000000F
   Me.txtB1011.BackColor = &H8000000F
   Me.txtB1014.BackColor = &H8000000F
   Me.txtB1015.BackColor = &H8000000F
   Me.Chk1Day.BackColor = &H8000000F
   Me.cboSTime.BackColor = &H8000000F
   Me.cboETime.BackColor = &H8000000F
   
   Me.CboB1002.Locked = True
   Me.CboB1008.Locked = True
   Me.txtB1006.Locked = True
   Me.txtB1007_1.Locked = True
   Me.txtB1007_2.Locked = True
   Me.txtB1011.Locked = True
   Me.txtB1014.Locked = True
   Me.txtB1015.Locked = True
   Me.Chk1Day.Enabled = False
   Me.cboSTime.Locked = True
   Me.cboETime.Locked = True
   
   '清空欄位值
   ClearField
      
   '預設值
'   SetB1002Combo CboB1002
'   SetB1008Combo CboB1008
''   For i = 0 To CboEmp.UBound
'      SetABS001_1Combo strUserNum
''   Next i
''   For i = 0 To CboBoss.UBound
'      SetABS001_2Combo strUserNum
''   Next i
'
'   Me.cmdModify.Enabled = False
'   Me.cmdDel.Enabled = False
'   Me.cmdSend.Enabled = True
'   Me.cmdagainSend.Enabled = False
'   m_EditMode = 1 '新增
'
'   Me.txtB1008_2 = GetCurrSpecRestDay(strUserNum)
   Call CboB1002_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180104 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 2
      GRD1.row = dblPrevRow
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   dblPrevRow = GRD1.row
   For i = 0 To GRD1.Cols - 1
      GRD1.col = i
      GRD1.CellBackColor = &HFFC0C0
   Next i
End If
GRD1.Visible = True
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("B1022")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1022")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("B1022"), True)
      End If
   End If
   m_B1023 = ""
   If IsNull(rsSrcTmp.Fields("B1023")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1023")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("B1023"))
         strCDate = Format(strTemp, "###/##/##")
         m_B1023 = rsSrcTmp.Fields("B1023")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1024")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1024")) = False Then
         strTemp = rsSrcTmp.Fields("B1024")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1025")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1025")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("B1025"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1026")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1026")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("B1026"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1027")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1027")) = False Then
         strTemp = rsSrcTmp.Fields("B1027")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label26.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

'Private Sub cmdSend_Click()
'
'On Error GoTo ErrHand
'
'   '檢查條件
'   If TxtValidate = False Then Exit Sub
'
'   cnnConnection.BeginTrans
'
'   If txtB1001 = "" Then
'      '表單編號自動給號
'      txtB1001 = AutoNo_ABS("ABS", 5)
'   Else
'      '主管代填時,不可再給號
'   End If
'
'   If SaveABS011() = False Then Exit Sub
'   If SaveABS010() = False Then Exit Sub
'
'   '送呈下一處理人員
'   If GetSendNextPerson(Trim(txtB1001), Trim(txtB1003), m_B1017, strUserNum) = False Then GoTo ErrHand
'
'   cnnConnection.CommitTrans
'
'   '發E-Mail通知下一處理人員
'   strContent = GetEMailContent(txtB1001, strSubject)
'   PUB_SendMail strUserNum, m_B1017, "", strSubject, strContent, , , , , , , , , , True
'
'   frm180101.Hide
'   frm180101.QueryData
'   frm180101.Show
'   Unload Me
'   Exit Sub
'
'ErrHand:
'   cnnConnection.RollbackTrans
'   MsgBox " 送出失敗！" & vbCrLf & Err.Description
'End Sub

''更新出缺勤電子簽核主檔
'Private Function SaveABS010() As Boolean
'Dim strB1008 As String, strB1014 As String, strB1015 As String
'Dim strB1006 As String, strB1009 As String, strB1010 As String
'Dim strB1012 As String, strB1013 As String
'
'On Error GoTo ErrHand
'
'   PUB_FilterFormText Me 'Add by Sindy 2011/10/14 修正畫面所有含跳行符號的文字框
'   SaveABS010 = True
'
'   '假別
'   If CboB1008.Visible = True Then
'      strB1008 = Left(CboB1008, 2)
'   End If
'   '迄止日期
'   If txtB1006.Visible = True Then
'      strB1006 = DBDATE(txtB1006)
'   End If
'   '日,時
'   If Frame01.Visible = True Then
'      strB1009 = txtB1009
'      strB1010 = txtB1010
'   End If
'   '時數-平日,假日
'   If Left(CboB1002, 2) = 表單類別_加班 Then
'      If txtB1012 <> "" Then strB1012 = txtB1012
'      If txtB1013 <> "" Then strB1013 = txtB1013
'   End If
'   '差程,地點
'   If Left(CboB1002, 2) = 表單類別_出差 Then
'      strB1014 = txtB1014
'      strB1015 = txtB1015
'   End If
'
'   strSql = "SELECT * FROM ABS010 WHERE B1001=" & CNULL(txtB1001)
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      '修改
'      strSql = "update ABS010 set " & _
'               "B1002= " & CNULL(Left(CboB1002, 2)) & _
'               ",B1004= " & CNULL(DBDATE(txtB1004)) & _
'               ",B1005= " & CNULL(txtB1005_1 & Format("00" & txtB1005_2, "00")) & _
'               ",B1006= " & CNULL(strB1006) & _
'               ",B1007= " & CNULL(txtB1007_1 & Format("00" & txtB1007_2, "00")) & _
'               ",B1008= " & CNULL(strB1008) & _
'               ",B1009= " & CNULL(strB1009) & _
'               ",B1010= " & CNULL(strB1010) & _
'               ",B1011= " & CNULL(Trim(txtB1011)) & _
'               ",B1012= " & CNULL(strB1012) & _
'               ",B1013= " & CNULL(strB1013) & _
'               ",B1014= " & CNULL(strB1014) & _
'               ",B1015= " & CNULL(strB1015) & _
'               ",B1028= " & CNULL(IIf(cboSTime.Visible = True And cboSTime.Text <> "", Format(cboSTime.Text, "hhmm"), "")) & _
'               ",B1029= " & CNULL(IIf(cboETime.Visible = True And cboETime.Text <> "", Format(cboETime.Text, "hhmm"), "")) & _
'               " where B1001=" & CNULL(txtB1001)
'   Else
'      '新增
'      strSql = "insert into ABS010(B1001,B1002,B1003,B1004,B1005,B1006,B1007,B1008,B1009,B1010,B1011,B1012,B1013,B1014,B1015,B1016,B1017,B1018,B1028,B1029) " & _
'               "values(" & CNULL(txtB1001) & "," & CNULL(Left(CboB1002, 2)) & "," & CNULL(txtB1003) & "," & _
'               CNULL(DBDATE(txtB1004)) & "," & CNULL(txtB1005_1 & Format("00" & txtB1005_2, "00")) & "," & CNULL(strB1006) & "," & _
'               CNULL(txtB1007_1 & Format("00" & txtB1007_2, "00")) & "," & CNULL(strB1008) & "," & CNULL(strB1009) & "," & _
'               CNULL(strB1010) & "," & CNULL(Trim(txtB1011)) & "," & CNULL(strB1012) & "," & _
'               CNULL(strB1013) & "," & CNULL(strB1014) & "," & CNULL(strB1015) & "," & _
'               CNULL(strUserNum) & "," & CNULL(m_B1017) & "," & CNULL(m_B1018) & "," & _
'               CNULL(IIf(cboSTime.Visible = True And cboSTime.Text <> "", Format(cboSTime.Text, "hhmm"), "")) & "," & _
'               CNULL(IIf(cboETime.Visible = True And cboETime.Text <> "", Format(cboETime.Text, "hhmm"), "")) & ") "
'   End If
'   cnnConnection.Execute strSql
'   Exit Function
'
'ErrHand:
'   SaveABS010 = False
'   cnnConnection.RollbackTrans
'   MsgBox " 新增ABS010失敗！" & vbCrLf & Err.Description
'End Function

''更新表單簽核檔
'Private Function SaveABS011() As Boolean
'
'On Error GoTo ErrHand
'
'   SaveABS011 = True
'
'   strSql = "delete From ABS011 where B1101=" & CNULL(txtB1001)
'   cnnConnection.Execute strSql
'   'If strType = "1" Then '送出
'      '1.職代
'      For i = 0 To CboEmp.UBound
'         If CboEmp(i) <> "" And CboEmp(i).Visible = True Then
'            strSql = "insert into ABS011 (B1101,B1102,B1103,B1104,B1108) values(" & CNULL(txtB1001) & ",'1'," & (i + 1) & "," & CNULL(Left(CboEmp(i), 5)) & ",'(代" & GetPersonSeqno(txtB1003, Left(CboEmp(i), 5)) & ")')"
'            cnnConnection.Execute strSql
'         End If
'      Next i
'      '2.審核主管
'      For i = 0 To CboBoss.UBound
'         If CboBoss(i) <> "" Then
'            strSql = "insert into ABS011 (B1101,B1102,B1103,B1104) values(" & CNULL(txtB1001) & ",'2'," & (i + 1) & "," & CNULL(Left(CboBoss(i), 5)) & ")"
'            cnnConnection.Execute strSql
'         End If
'      Next i
''   Else '重送
''      strSql = "update ABS011 set " & _
''               "B1105=null" & _
''               ",B1106=null" & _
''               ",B1107=null" & _
''               " where B1101=" & CNULL(txtB1001)
''      cnnConnection.Execute strSql
''   End If
'
'   Exit Function
'
'ErrHand:
'   SaveABS011 = False
'   cnnConnection.RollbackTrans
'   MsgBox " 新增ABS011失敗！" & vbCrLf & Err.Description
'End Function

Private Sub cmdagainSend_Click()
Dim strOldB1017 As String
Dim strUpdDate As String, strUpdTime As String, strB1207 As String
Dim strText As String
Dim bolSave As Boolean
'Add By Sindy 2012/1/4
Dim strB1001 As String, strB1002 As String, strB1003 As String
Dim strB1004 As String, strB1005 As String
Dim strB1006 As String, strB1007 As String
Dim strB1008 As String, strB1009 As String, strB1010 As String
Dim strB101213 As String
Dim strB1014 As String, strB1015 As String
Dim strB1028 As String, strB1029 As String
'2012/1/4 End
Dim strB1030 As String 'Add By Sindy 2016/12/30
   
On Error GoTo ErrHand
   
   bolSave = False
   strOldB1017 = m_B1017 '記錄原下一處理人員
   '檢查條件
   If TxtValidate = False Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   
   cnnConnection.BeginTrans
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
'   '表單退回到當事人時,重送
'   If cmdModify.Enabled = True Then
'      If SaveABS011() = False Then Exit Sub
'      If SaveABS010() = False Then Exit Sub
'      '送呈下一處理人員
'      If GetSendNextPerson(Trim(txtB1001), Trim(txtB1003), m_B1017, strUserNum) = False Then GoTo ErrHand
'
'      '記錄10.重送訊息
'      strSql = GetInsertABS012Sql(Trim(txtB1001), strUserNum, strUpdDate, strUpdTime, "07", "")
'      cnnConnection.Execute strSql
'
'      '發E-Mail通知下一處理人員
'      strContent = GetEMailContent(txtB1001, strSubject, 重送)
'      PUB_SendMail strUserNum, m_B1017, "", strSubject, strContent, , , , , , , , , ,  True
'
'   '表單簽核中,異動簽核人員(卡單)
'   Else
      '2.審核主管
      For i = CboBoss.UBound To 0 Step -1
         If CboBoss(i).Enabled = True Then '欄位為可使用狀態
            If m_cboBoss(i) <> Left(Trim(CboBoss(i)), 5) Then '比對原資料與目前資料是否相同,不同時才更新
               If m_cboBoss(i) <> "" Then '原始資料有值時,先刪除
                  bolSave = True
                  strSql = "delete From ABS011 where B1101=" & CNULL(txtB1001) & " and B1102='2' and B1104=" & CNULL(m_cboBoss(i)) & " and B1107 is null "
                  Pub_SeekTbLog strSql '記錄刪除Log Add By Sindy 2023/9/15
                  cnnConnection.Execute strSql
               End If
               If Left(CboBoss(i), 5) <> "" Then '新增目前資料
                  bolSave = True
                  'Modify By Sindy 2015/11/13 + B1109
                  strSql = "INSERT INTO ABS011 (B1101,B1102,B1103,B1104,B1109) VALUES(" & CNULL(txtB1001) & ",'2'," & (i + 1) & "," & CNULL(Left(CboBoss(i), 5)) & "," & CNULL(Left(CboBoss(i), 5)) & ")"
                  Pub_SeekTbLog strSql '記錄刪除Log Add By Sindy 2023/9/15
                  cnnConnection.Execute strSql
               End If
            End If
         End If
      Next i
      '1.職代
      For i = CboEmp.UBound To 0 Step -1
         If CboEmp(i).Enabled = True Then '欄位為可使用狀態
            If m_cboEmp(i) <> Left(Trim(CboEmp(i)), 5) Then '比對原資料與目前資料是否相同,不同時才更新
               If m_cboEmp(i) <> "" Then '原始資料有值時,先刪除
                  bolSave = True
                  strSql = "delete From ABS011 where B1101=" & CNULL(txtB1001) & " and B1102='1' and B1104=" & CNULL(m_cboEmp(i)) & " and B1107 is null "
                  Pub_SeekTbLog strSql '記錄刪除Log Add By Sindy 2023/9/15
                  cnnConnection.Execute strSql
               End If
               If Left(CboEmp(i), 5) <> "" Then '新增目前資料
                  bolSave = True
                  strSql = "INSERT INTO ABS011 (B1101,B1102,B1103,B1104,B1108) VALUES(" & CNULL(txtB1001) & ",'1'," & (i + 1) & "," & CNULL(Left(CboEmp(i), 5)) & ",'(代" & GetPersonSeqno(txtB1003, Left(CboEmp(i), 5)) & ")')"
                  Pub_SeekTbLog strSql '記錄刪除Log Add By Sindy 2023/9/15
                  cnnConnection.Execute strSql
               End If
            End If
         End If
      Next i
      
      '送呈下一處理人員
      'If GetSendNextPerson(Trim(txtB1001), Trim(txtB1003), m_B1017, strUserNum) = False Then GoTo ErrHand
      '讀取下一處理人員
      If GetNextProPerson(Trim(txtB1001), Trim(txtB1003), m_B1017, strUserNum) = False Then GoTo ErrHand
      'Modify By Sindy 2012/1/4 [寫成共用函數] 人事處不簽收,最高審核主管簽核完畢,系統自動簽收進人事系統
      If m_B1017 = "M21" Then
         strB1001 = Trim(txtB1001)
         strB1002 = Trim(Left(CboB1002, 2))
         strB1003 = Trim(txtB1003)
         strB1004 = DBDATE(txtB1004)
         strB1005 = txtB1005_1 & Format("00" & txtB1005_2, "00")
         strB1006 = DBDATE(txtB1006)
         strB1007 = txtB1007_1 & Format("00" & txtB1007_2, "00")
         strB1008 = Trim(Left(CboB1008, 2))
         strB1009 = txtB1009
         strB1010 = txtB1010
         strB1030 = txtB1030
         strB101213 = txtB101213
         strB1014 = txtB1014
         strB1015 = txtB1015
         strB1028 = IIf(cboSTime.Visible = False, "", Format(cboSTime, "hhmm"))
         strB1029 = IIf(cboETime.Visible = False, "", Format(cboETime, "hhmm"))
         If PUB_AutoM21Receive(strUserNum, strUpdDate, strUpdTime, strB1001, strB1002, strB1003, strB1004, strB1005, strB1006, strB1007, strB1008, strB1009, strB1010, strB1030, strB101213, strB1014, strB1015, strB1028, strB1029, strSubject, strContent) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      '2012/1/4 End
      Else
         If strOldB1017 <> m_B1017 Then
            If ChkStaffST04(strOldB1017, False) = True Then
               strText = "人員離職，"
            Else
               'strText = "人員休假，"
               strText = ""
            End If
            
            '記錄10.重送訊息
            strB1207 = strText & "更改簽核人員" & GetPrjSalesNM(strOldB1017) & "->" & GetPrjSalesNM(m_B1017)
            strSql = GetInsertABS012Sql(Trim(txtB1001), strUserNum, strUpdDate, strUpdTime, "07", strB1207)
            cnnConnection.Execute strSql
            
            '有異動下一處理人員時,發E-Mail通知下一處理人員及原下一處理人員
            If m_B1017 <> "" Then
               strContent = GetEMailContent(txtB1001, strSubject)
               PUB_SendMail strUserNum, m_B1017, "", strSubject, strContent, , , , , , , , , , True
            End If
            If strOldB1017 <> strUserNum And strOldB1017 <> "" And strText <> "人員離職" Then
               strContent = GetEMailContent(txtB1001, strSubject, 重送更改通知, "，" & strText & "更改簽核人員為" & GetPrjSalesNM(m_B1017))
               PUB_SendMail strUserNum, strOldB1017, "", strSubject, strContent, , , , , , , , , , True
            End If
         End If
         cnnConnection.CommitTrans
      End If
'   End If
   
   If bolSave = True Then
      'MsgBox "已重送成功！", vbExclamation + vbOKOnly
      MsgBox "執行完畢！", vbExclamation + vbOKOnly
      Call cmdClear_Click
   Else
      MsgBox "無異動資料！", vbExclamation + vbOKOnly
   End If
   
   Screen.MousePointer = vbDefault
'   frm180101.Hide
'   frm180101.QueryData
'   frm180101.Show
'   Unload Me
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox " 重送失敗！" & vbCrLf & Err.Description
End Sub

Public Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim strST20 As String
Dim intBossNum As Integer

TxtValidate = False

If CboEmp(0).Visible = True Then
   If CboEmp(0) = "" Then
      MsgBox "職務代理人不可以空白！", vbExclamation
      CboEmp(0).SetFocus
      Exit Function
   End If
   For i = 0 To CboEmp.UBound
      If Me.CboEmp(i).Enabled = True Then
         Cancel = False
         bolChk = False
         cboEmp_Validate i, Cancel
         If Cancel = True Then
            CboEmp(i).SetFocus
            Exit Function
         End If
      End If
   Next i
End If

'所長可以無審核主管
strSql = "SELECT st20 FROM staff WHERE ST01='" & txtB1003 & "' "
intI = 1: strST20 = ""
Set RsTemp = ClsLawReadRstMsg(intI, strSql)
If intI = 1 Then
   If Not IsNull(RsTemp("ST20")) Then strST20 = RsTemp("ST20")
End If
'Modify By Sindy 2022/7/18 + 15.名譽所長
If strST20 <> "11" And strST20 <> "15" And CboBoss(0) = "" Then
   MsgBox "審核主管不可以空白！", vbExclamation
   CboBoss(0).SetFocus
   Exit Function
End If

intBossNum = 0
For i = 0 To CboBoss.UBound
   If Me.CboBoss(i).Enabled = True Then
      Cancel = False
      CboBoss_Validate i, Cancel
      If Cancel = True Then
         CboBoss(i).SetFocus
         Exit Function
      End If
   End If
   If Me.CboBoss(i).Text <> "" Then
      intBossNum = intBossNum + 1
   End If
Next i

''增加判斷權責主管人數是否足夠，若不足，則不可簽核
'If Val(m_BossNum) > 0 Then
'   If intBossNum < m_BossNum Then
'      MsgBox "審核主管人數應為" & m_BossNum & "人，人數不足不可簽核！", vbExclamation
'      Exit Function
'   End If
'End If

'Add by Sindy 2021/5/28 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me) = False Then
   Exit Function
End If
'2021/5/28 END

TxtValidate = True
End Function

Private Sub CboB1002_Click()
   If Left(CboB1002.Text, 2) = 表單類別_請假 Or Left(CboB1002.Text, 2) = "" Then
      'txtB1008_2.Visible = True
      Label10.Visible = True
      CboB1008.Visible = True
      Label1(2).Visible = True
      txtB1006.Visible = True
      Frame01.Visible = True
      Frame02.Visible = False
      Frame03.Visible = False
      Label13.Visible = True
      CboEmp(0).Visible = True
      Label25.Visible = True
      CboEmp(1).Visible = True
      Label27.Visible = True
      CboEmp(2).Visible = True
      Chk1Day.Visible = True: If txtB1001 = "" Then Chk1Day.Value = 0
   '      Label1(3).Visible = True
   '      cboSTime.Visible = True
   '      Label1(4).Visible = True
   '      cboETime.Visible = True
   ElseIf Left(CboB1002.Text, 2) = 表單類別_加班 Then
      'txtB1008_2.Visible = False
      Label10.Visible = False
      CboB1008.Visible = False
      CboB1008.Text = ""
      Label1(2).Visible = False
      txtB1006.Visible = False
      Frame01.Visible = False
      Frame02.Visible = True
      Frame03.Visible = False
      Frame02.Left = 3090 '900
      Frame02.Top = 2070
      Label13.Visible = False
      CboEmp(0).Visible = False
      Label25.Visible = False
      CboEmp(1).Visible = False
      Label27.Visible = False
      CboEmp(2).Visible = False
      Chk1Day.Visible = False: Chk1Day.Value = 1
'      Label1(3).Visible = False
'      cboSTime.Visible = False
'      Label1(4).Visible = False
'      cboETime.Visible = False
      Frame1.Visible = False
   ElseIf Left(CboB1002.Text, 2) = 表單類別_出差 Then
      'txtB1008_2.Visible = False
      Label10.Visible = False
      CboB1008.Visible = False
      CboB1008.Text = ""
      Label1(2).Visible = True
      txtB1006.Visible = True
      Frame01.Visible = True
      Frame02.Visible = False
      Frame03.Visible = True
      Label13.Visible = True
      CboEmp(0).Visible = True
      Label25.Visible = True
      CboEmp(1).Visible = True
      Label27.Visible = True
      CboEmp(2).Visible = True
      'Modify By Sindy 2012/4/13
      Chk1Day.Visible = True: If txtB1001 = "" Then Chk1Day.Value = 0
      'Chk1Day.Visible = False: Chk1Day.Value = 1
      '2012/4/13 End
   '      Label1(3).Visible = True
   '      cboSTime.Visible = True
   '      Label1(4).Visible = True
   '      cboETime.Visible = True
      Frame1.Visible = False
   End If
   'Modify By Sindy 2012/7/9 +尤春彬
   If strUserNum = "99029" Or strUserNum = "84043" Then '伊恩
      Chk1Day.Visible = False: Chk1Day.Value = 1
   End If
'   Call Chk1Day_Click
End Sub

Private Sub CboB1002_GotFocus()
   InverseTextBox CboB1002
End Sub

Private Sub CboB1002_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboB1002_LostFocus()
   If Left(CboB1002, 2) = 表單類別_請假 Then
      If CboB1008.Enabled = True Then CboB1008.SetFocus
   ElseIf Left(CboB1002, 2) = 表單類別_加班 Then
      If txtB1004.Enabled = True Then txtB1004.SetFocus
   ElseIf Left(CboB1002, 2) = 表單類別_出差 Then
      If txtB1014.Enabled = True Then txtB1014.SetFocus
   End If
   If CboB1002.Text > "" Then
      For i = 0 To CboB1002.ListCount - 1
         If Left(CboB1002.List(i), 2) = CboB1002.Text Then CboB1002.Text = CboB1002.List(i): Exit For
      Next i
   End If
End Sub

Private Sub CboB1002_Validate(Cancel As Boolean)
Dim bolComp As Boolean
   
'   If CboB1002 <> "" Then
'      bolComp = False
'      For i = 0 To CboB1002.ListCount
'         If Left(CboB1002, 2) = Left(CboB1002.List(i), 2) Then
'            bolComp = True
'            Exit For
'         End If
'      Next i
'      If bolComp = False Then
'         MsgBox "表單類別有誤!!!", vbExclamation + vbOKOnly
'         Call CboB1002_GotFocus
'         Cancel = True
'         Exit Sub
'      End If
'   Else
'      MsgBox "表單類別不可以空白！", vbExclamation
'      Call CboB1002_GotFocus
'      Cancel = True
'      Exit Sub
'   End If
End Sub

Private Sub CboB1008_GotFocus()
   InverseTextBox CboB1008
End Sub

Private Sub CboB1008_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboB1008_LostFocus()
'   Call Chk1Day_Click
   If CboB1008.Text > "" Then
      For i = 0 To CboB1008.ListCount - 1
         If Left(CboB1008.List(i), 2) = CboB1008.Text Then CboB1008.Text = CboB1008.List(i): Exit For
      Next i
      If GetCountDayHour(False) Then
         Call CboB1008_GotFocus
         Exit Sub
      End If
      If Left(CboB1008.Text, 2) = "08" Then txtB1004.SetFocus
   End If
End Sub

Private Sub CboB1008_Validate(Cancel As Boolean)
Dim MyRs As New ADODB.Recordset
Dim MyArr As Variant
   
'   If CboB1008.Text <> "" Then
'      MyArr = Split(CboB1008, " ")
'      Set MyRs = New ADODB.Recordset
'      If MyRs.State = 1 Then MyRs.Close
'      ' 排除不須要的代碼 : 01.忘打卡 02.遲到 03.曠職 04.出差 16.加班 17.扣年終產假 18.扣年終流產假
'      strSql = "select ac02||' '||ac03 from allcode where ac01='04' and ac02='" & MyArr(0) & "' and ac02 not in ('01','02','03','04','16','17','18') order by ac02"
'      MyRs.CursorLocation = adUseClient
'      MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If MyRs.RecordCount <> 0 Then
'         CboB1008.Text = "" & MyRs.Fields(0).Value
'      Else
'         MsgBox "假別代號輸入錯誤!!!", vbExclamation + vbOKOnly
'         Call CboB1008_GotFocus
'         Cancel = True
'         Exit Sub
'      End If
'   Else
'      MsgBox "假別不可以空白！", vbExclamation
'      Call CboB1008_GotFocus
'      Cancel = True
'      Exit Sub
'   End If
End Sub

'Add By Sindy 2023/2/24
Private Sub txtB1001_LostFocus()
   If txtB1001 <> "" Then
      cmdQuery_Click
   End If
End Sub

Private Sub txtB1003_GotFocus()
   InverseTextBox txtB1003
End Sub

Private Sub txtB1003_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtB1003_LostFocus()
   If txtB1003.Text <> "" Then
      '抓取員工姓名
      txtB1003_2.Text = GetPrjSalesNM(txtB1003.Text)
   End If
End Sub

Private Sub txtB1003_Validate(Cancel As Boolean)
   If txtB1003.Text = "" Then txtB1003_2.Text = ""
   
   'If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If txtB1003 <> "" Then
      '檢查人員是否存在或離職
      If ChkStaffST04(txtB1003) = True Then
         Call txtB1003_GotFocus
         Cancel = True
         Exit Sub
      End If
      '檢查 員工不可為”不寄信”
      If ChkStaffST14(txtB1003) = True Then
         Call txtB1003_GotFocus
         Cancel = True
         Exit Sub
      End If
      '檢查是否有權限
      If ChkLimitsIsOk() = False Then
'         Call txtB1003_GotFocus
         MsgBox "無權限異動此人員表單!", vbExclamation + vbOKOnly
'         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtB1004_GotFocus()
   InverseTextBox txtB1004
   CloseIme
End Sub

Private Sub txtB1004_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'Private Sub txtB1004_LostFocus()
'   If txtB1006.Visible = True Then
'      If txtB1004 <> "" And txtB1006 <> "" And txtB1004 <> txtB1006 Then
'         Call Chk1Day_Click
'      End If
'   End If
'End Sub

Private Sub txtB1004_Validate(Cancel As Boolean)
'Dim strTime As String
'
'If txtB1006.Visible = True Then
'   If txtB1004 <> "" And txtB1006 <> "" And txtB1004 <> txtB1006 Then
'      Call Chk1Day_Click
'   End If
'End If

If txtB1004 <> "" Then
   If CheckIsTaiwanDate(txtB1004, False) = False Then
      MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
      Call txtB1004_GotFocus
      Cancel = True
      Exit Sub
   End If
   If Left(CboB1002, 2) <> 表單類別_加班 Then
      If ChkWorkDay(DBDATE(txtB1004)) = False Then
         MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
         Call txtB1004_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
'   If txtB1004 <> "" And txtB1006 <> "" Then
'      If Val(txtB1004) > Val(txtB1006) Then
'         txtB1006 = ""
'      Else
'         If RunNick2(txtB1004, txtB1006) Then
'            Call txtB1004_GotFocus
'            Cancel = True
'            Exit Sub
'         End If
'      End If
'   End If
'   If Left(CboB1008, 2) = "08" Then '特別假必須要提早1天前請(指工作天)
'      If txtB1001 <> "" Then '有表單編號時
'         If DBDATE(txtB1004) <= CompWorkDay(2, DBDATE(m_B1023), 0) Then
'            Call txtB1004_GotFocus
'            Cancel = True
'            MsgBox "特別假須提早1個工作天！", vbInformation, "輸入日期錯誤"
'            Exit Sub
'         End If
'      Else
'         strTime = Right("000000" & ServerTime, 6)
'         If DBDATE(txtB1004) <= CompWorkDay(2, DBDATE(strSrvDate(1)), 0) Then
'            Call txtB1004_GotFocus
'            Cancel = True
'            MsgBox "特別假須提早1個工作天！", vbInformation, "輸入日期錯誤"
'            Exit Sub
'         ElseIf DBDATE(txtB1004) = CompWorkDay(3, DBDATE(strSrvDate(1)), 0) And Val(Left(strTime, Len(strTime) - 2)) >= 1800 Then
'            Call txtB1004_GotFocus
'            Cancel = True
'            MsgBox "特別假已超出可以請假的時間！", vbInformation, "輸入日期錯誤"
'            Exit Sub
'         End If
'      End If
'   End If
'   If GetCountDayHour(True) = False Then
'      Call txtB1004_GotFocus
'      Cancel = True
'      Exit Sub
'   End If
End If
End Sub

Private Sub txtB1005_1_GotFocus()
   InverseTextBox txtB1005_1
End Sub

Private Sub txtB1005_1_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1005_1_Validate(Cancel As Boolean)
'If txtB1005_1 = "" Then txtB1005_1 = "00"

If txtB1005_1 <> "" Then
   If CheckLengthIsOK(txtB1005_1, txtB1005_1.MaxLength) = False Then
      Call txtB1005_1_GotFocus
      Cancel = True
      Exit Sub
   End If
   If Val(txtB1005_1.Text) = 0 And txtB1004 <> "" Then
      MsgBox "請輸入時分!", vbExclamation + vbOKOnly
      Call txtB1005_1_GotFocus
      Cancel = True
      Exit Sub
   End If
   If txtB1005_1.Text > 24 Then
      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
      Call txtB1005_1_GotFocus
      Cancel = True
      Exit Sub
   End If
'   If GetCountDayHour(True) = False Then
'      Call txtB1005_1_GotFocus
'      Cancel = True
'      Exit Sub
'   End If
End If
CloseIme
End Sub

Private Sub txtB1005_2_GotFocus()
   InverseTextBox txtB1005_2
End Sub

Private Sub txtB1005_2_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1005_2_Validate(Cancel As Boolean)
'If txtB1005_2 = "" Then txtB1005_2 = "00"

If txtB1005_2 <> "" Then
   If CheckLengthIsOK(txtB1005_2, txtB1005_2.MaxLength) = False Then
      Call txtB1005_2_GotFocus
      Cancel = True
      Exit Sub
   End If
   If txtB1005_2.Text > 59 Then
      Call txtB1005_2_GotFocus
      MsgBox "不可超過59分!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
'   If GetCountDayHour(True) = False Then
'      Call txtB1005_2_GotFocus
'      Cancel = True
'      Exit Sub
'   End If
End If
CloseIme
End Sub

Private Sub txtB1006_GotFocus()
   InverseTextBox txtB1006
End Sub

Private Sub txtB1006_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'Private Sub txtB1006_LostFocus()
'   If txtB1006.Visible = True Then
'      If txtB1004 <> "" And txtB1006 <> "" And txtB1004 <> txtB1006 Then
'         Call Chk1Day_Click
'      End If
'   End If
'End Sub

Private Sub txtB1006_Validate(Cancel As Boolean)
Dim strTime As String

'If txtB1006.Visible = True Then
'   If txtB1004 <> "" And txtB1006 <> "" And txtB1004 <> txtB1006 Then
'      Call Chk1Day_Click
'   End If
'End If

If txtB1006 <> "" Then
   If CheckIsTaiwanDate(txtB1006, False) = False Then
      MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
      Call txtB1006_GotFocus
      Cancel = True
      Exit Sub
   End If
   If Left(CboB1002, 2) = 表單類別_請假 Then
      If ChkWorkDay(DBDATE(txtB1006)) = False Then
         Call txtB1006_GotFocus
         Cancel = True
         MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
         Exit Sub
      End If
   End If
   If txtB1004 <> "" And txtB1006 <> "" Then
      If RunNick2(txtB1004, txtB1006) Then
         Call txtB1006_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
'   If txtB1007_2.Enabled = False Then '特別假時,若輸入迄止日期即可計算天,時
'      Cancel = False
'      txtB1007_2_Validate Cancel
'      If Cancel = True Then
'         Exit Sub
'      End If
'   End If
   If Left(CboB1008, 2) = "08" Then '特別假必須要提早1天前請(指工作天)
      If txtB1001 <> "" Then '有表單編號時
         If DBDATE(txtB1006) <= CompWorkDay(2, DBDATE(m_B1023), 0) Then
            Call txtB1006_GotFocus
            Cancel = True
            MsgBox "特別假必須要提早1天前請(指工作天)！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
      Else
         strTime = Right("000000" & ServerTime, 6)
         If DBDATE(txtB1006) <= CompWorkDay(2, DBDATE(strSrvDate(1)), 0) Then
            Call txtB1006_GotFocus
            Cancel = True
            MsgBox "特別假必須要提早1天前請(指工作天)！", vbInformation, "輸入日期錯誤"
            Exit Sub
         ElseIf DBDATE(txtB1006) = CompWorkDay(3, DBDATE(strSrvDate(1)), 0) And Val(Left(strTime, Len(strTime) - 2)) >= 1800 Then
            Call txtB1006_GotFocus
            Cancel = True
            MsgBox "特別假已超出可以請假的時間！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
      End If
   End If
   If GetCountDayHour(True) = False Then
      Call txtB1006_GotFocus
      Cancel = True
      Exit Sub
   End If
End If
End Sub

Private Sub txtB1007_1_GotFocus()
   InverseTextBox txtB1007_1
End Sub

Private Sub txtB1007_1_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1007_1_Validate(Cancel As Boolean)
If txtB1007_1 = "" Then txtB1007_1 = "00"

If txtB1007_1 <> "" Then
   If CheckLengthIsOK(txtB1007_1, txtB1007_1.MaxLength) = False Then
      Call txtB1007_1_GotFocus
      Cancel = True
      Exit Sub
   End If
   If Val(txtB1007_1.Text) = 0 And _
      ((txtB1004 <> "" And Left(CboB1002, 2) = 表單類別_加班) Or (txtB1006 <> "" And Left(CboB1002, 2) <> 表單類別_加班)) Then
      Call txtB1007_1_GotFocus
      MsgBox "請輸入時分!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   If txtB1007_1.Text > 24 Then
      Call txtB1007_1_GotFocus
      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   If GetCountDayHour(True) = False Then
      Call txtB1007_1_GotFocus
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
End Sub

Private Sub txtB1007_2_GotFocus()
   InverseTextBox txtB1007_2
End Sub

Private Sub txtB1007_2_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1007_2_Validate(Cancel As Boolean)
If txtB1007_2 = "" Then txtB1007_2 = "00"

If txtB1007_2 <> "" Then
   If CheckLengthIsOK(txtB1007_2, txtB1007_2.MaxLength) = False Then
      Call txtB1007_2_GotFocus
      Cancel = True
      Exit Sub
   End If
   If txtB1007_2.Text > 59 Then
      Call txtB1007_2_GotFocus
      MsgBox "不可超過59分!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   'If Trim(txtB1004) <> "" And Trim(txtB1005_1) <> "" And Trim(txtB1005_2) <> "" And Trim(txtB1006) <> "" And Trim(txtB1007_1) <> "" And Trim(txtB1007_2) <> "" Then
      If CheckIsTaiwanDate(txtB1004, False) = True And CheckIsTaiwanDate(txtB1006, False) = True Then
         If CompDateTime(txtB1004 & Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), txtB1006 & Format(txtB1007_1, "00") & Format(txtB1007_2, "00")) = False Then
            Call txtB1007_2_GotFocus
            MsgBox "日期時間設定錯誤！", vbInformation, "輸入錯誤！"
            'Cancel = True
            Exit Sub
         End If
      End If
   'End If
   If GetCountDayHour(True) = False Then
      Call txtB1007_2_GotFocus
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
End Sub

Private Function GetCountDayHour(bolChkExist As Boolean) As Boolean
Dim dblSTime As Double, dblETime As Double, strB1008 As String
   
   GetCountDayHour = True
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If Left(CboB1002, 2) = 表單類別_加班 Then
         '欄位值尚未輸入完整
         If Val(txtB1004) = 0 Or Val(txtB1005_1) = 0 Or _
            Val(txtB1007_1) = 0 Then
            Exit Function
         End If
         '無異動欄位值
         If Val(txtB1004) = Val(m_B1004) And Val(txtB1005_1) = Val(m_B1005_1) And Val(txtB1005_2) = Val(m_B1005_2) And _
            Val(txtB1007_1) = Val(m_B1007_1) And Val(txtB1007_2) = Val(m_B1007_2) And _
            CboBoss(0) <> "" Then
            Exit Function
         End If
         If bolChkExist = True Then
            If IsRecordExist(txtB1003, DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1004), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
               GetCountDayHour = False
               Exit Function
            End If
         End If
         'If txtB1012 = "" And txtB1013 = "" Then Call AutoCount
      Else
         '欄位值尚未輸入完整
         If Val(txtB1004) = 0 Or Val(txtB1005_1) = 0 Or _
            Val(txtB1006) = 0 Or Val(txtB1007_1) = 0 Or _
            (cboSTime.Visible = True And cboSTime = "" And Chk1Day.Value = 1 And txtB1004 <> txtB1006) Or _
            (cboETime.Visible = True And cboETime = "" And Chk1Day.Value = 1 And txtB1004 <> txtB1006) Then
            Exit Function
         End If
         '無異動欄位值
         If cboSTime.Visible = True And cboSTime.Text <> "" Then dblSTime = Val(Format(cboSTime.Text, "hhmm"))
         If cboETime.Visible = True And cboETime.Text <> "" Then dblETime = Val(Format(cboETime.Text, "hhmm"))
         If CboB1008.Visible = True And CboB1008.Text <> "" Then strB1008 = Left(CboB1008, 2)
         If Val(txtB1004) = Val(m_B1004) And Val(txtB1005_1) = Val(m_B1005_1) And Val(txtB1005_2) = Val(m_B1005_2) And _
            Val(txtB1006) = Val(m_B1006) And Val(txtB1007_1) = Val(m_B1007_1) And Val(txtB1007_2) = Val(m_B1007_2) And _
            dblSTime = Val(m_B1028) And dblETime = Val(m_B1029) And strB1008 = m_B1008 And _
            CboEmp(0) <> "" And CboBoss(0) <> "" Then
            Exit Function
         End If
         If bolChkExist = True Then
            If IsRecordExist(txtB1003, DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1006), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
               GetCountDayHour = False
               Exit Function
            End If
         End If
         'If (txtB1009 = "" And txtB1010 = "") Or (txtB1009 = "0" And txtB1010 = "0") Then Call AutoCount
      End If
'      If AutoCount = False Then GetCountDayHour = False: Exit Function
   End If
End Function

Private Sub txtB1009_GotFocus()
   InverseTextBox txtB1009
End Sub

Private Sub txtB1009_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1009_Validate(Cancel As Boolean)
If txtB1009 <> "" Then
   If CheckLengthIsOK(txtB1009, txtB1009.MaxLength) = False Then
      Call txtB1009_GotFocus
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
End Sub

Private Sub txtB1010_GotFocus()
   InverseTextBox txtB1010
End Sub

Private Sub txtB1010_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub txtB1010_Validate(Cancel As Boolean)
If txtB1010 <> "" Then
   If CheckLengthIsOK(txtB1010, txtB1010.MaxLength) = False Then
      Call txtB1010_GotFocus
      Cancel = True
      Exit Sub
   End If
   'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
   'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
   'Modify By Sindy 2012/7/9 上班時數為特殊者
   Call Pub_GetSpecWorkHour(txtB1003, txtB1004)
'   If txtB1003 = "99029" Then
'      If txtB1010.Text >= 5 Then
'         Call txtB1010_GotFocus
'         MsgBox "請假時數-共(時)不可超過5小時!!!", vbExclamation + vbOKOnly
'         Cancel = True
'         Exit Sub
'      End If
   If Val(txtB1010.Text) >= Val(PUB_intWkHour) Then
      Call txtB1010_GotFocus
      MsgBox "請假時數-共(時)不可超過" & PUB_intWkHour & "小時!!!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
End Sub

Private Sub txtB1011_GotFocus()
   InverseTextBox txtB1011
   OpenIme
End Sub

'Add By Sindy 2021/5/31
Private Sub txtB1011_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtB1011
End Sub

Private Sub txtB1011_Validate(Cancel As Boolean)
If txtB1011 <> "" Then
   If CheckLengthIsOK(txtB1011, txtB1011.MaxLength) = False Then
      Call txtB1011_GotFocus
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
End Sub

'Private Sub txtB1012_GotFocus()
'   InverseTextBox txtB1012
'   CloseIme
'End Sub
'
'Private Sub txtB1012_KeyPress(KeyAscii As Integer)
'   KeyAscii = Pub_NumAscii(KeyAscii, True)
'End Sub
'
'Private Sub txtB1012_Validate(Cancel As Boolean)
'If txtB1012 <> "" Then
'   If CheckLengthIsOK(txtB1012, txtB1012.MaxLength) = False Then
'      Call txtB1012_GotFocus
'      Cancel = True
'      Exit Sub
'   End If
'End If
'CloseIme
'End Sub
'
'Private Sub txtB1013_GotFocus()
'   InverseTextBox txtB1013
'   CloseIme
'End Sub
'
'Private Sub txtB1013_KeyPress(KeyAscii As Integer)
'   KeyAscii = Pub_NumAscii(KeyAscii, True)
'End Sub
'
'Private Sub txtB1013_Validate(Cancel As Boolean)
'If txtB1013 <> "" Then
'   If CheckLengthIsOK(txtB1013, txtB1013.MaxLength) = False Then
'      Call txtB1013_GotFocus
'      Cancel = True
'      Exit Sub
'   End If
'End If
'CloseIme
'End Sub

Private Sub txtB1014_GotFocus()
   InverseTextBox txtB1014
End Sub

Private Sub txtB1014_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1014_Validate(Cancel As Boolean)
   If txtB1014 <> "" Then
      If CheckLengthIsOK(txtB1014, txtB1014.MaxLength) = False Then
          Call txtB1014_GotFocus
          Cancel = True
          Exit Sub
      End If
      If Trim(txtB1014) <> "" Then
        If txtB1014 <> "1" And txtB1014 <> "2" And txtB1014 <> "3" And txtB1014 <> "4" Then
           MsgBox "差程代碼有誤!!!", vbExclamation + vbOKOnly
           Call txtB1014_GotFocus
           Cancel = True
           Exit Sub
        End If
      End If
   Else
      MsgBox "差程不可以空白！", vbExclamation
      Call txtB1014_GotFocus
      Cancel = True
      Exit Sub
   End If
CloseIme
End Sub

Private Sub txtB1015_GotFocus()
   InverseTextBox txtB1015
   OpenIme
End Sub

'Add By Sindy 2021/5/31
Private Sub txtB1015_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtB1015
End Sub

Private Sub txtB1015_Validate(Cancel As Boolean)
   If txtB1015 <> "" Then
      If CheckLengthIsOK(txtB1015, txtB1015.MaxLength) = False Then
         Call txtB1015_GotFocus
         Cancel = True
         Exit Sub
      End If
'   Else
'      MsgBox "出差地點不可以空白！", vbExclamation
'      Call txtB1015_GotFocus
'      Cancel = True
'      Exit Sub
   End If
CloseIme
End Sub

Private Sub CboEmp_GotFocus(Index As Integer)
   InverseTextBox CboEmp(Index)
   bolChk = True
End Sub

Private Sub CboEmp_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboEmp_LostFocus(Index As Integer)
   If CboEmp(Index).Text > "" And Len(Trim(CboEmp(Index).Text)) = 5 Then
      '抓取員工姓名
      CboEmp(Index).Text = SetCboStaffName(CboEmp(Index).Text)
   End If
End Sub

Private Sub cboEmp_Validate(Index As Integer, Cancel As Boolean)
Dim strMsgText As String
   
   If CboEmp(Index) <> "" Then
      If Left(CboEmp(Index), 5) = txtB1003 Then
         MsgBox "不可為本人！", vbExclamation
         CboEmp(Index).SetFocus
         Call CboEmp_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(CboEmp(Index), 5)) = True Then
         Call CboEmp_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查 員工不可為”不寄信”
      If ChkStaffST14(Left(CboEmp(Index), 5)) = True Then
         Call CboEmp_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查職代輸入順序
      If (Trim(CboEmp(1)) <> "" And Trim(CboEmp(0)) = "") Or _
         (Trim(CboEmp(2)) <> "" And Trim(CboEmp(1)) = "") Then
         MsgBox "請依序輸入職務代理人！", vbExclamation
         CboEmp(Index).SetFocus
         Call CboEmp_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      If (CboEmp(1) <> "" And Left(CboEmp(1), 5) = Left(CboEmp(0), 5) And CboEmp(0).Enabled = True) Or _
         (CboEmp(2) <> "" And Left(CboEmp(2), 5) = Left(CboEmp(1), 5) And CboEmp(1).Enabled = True) Then
         MsgBox "資料重覆！", vbExclamation
         CboEmp(Index).SetFocus
         Call CboEmp_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      If bolChk = True Then
         bolChk = False
         strMsgText = ""
         '檢查取得的職代和表單當事人是否有相同的請假區間
         'Modify By Sindy 2017/1/10
         If CheckIsPersonRestSectorSame(Left(CboEmp(Index), 5), txtB1004, Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), txtB1006, Trim(txtB1007_1.Text & ":" & txtB1007_2.Text), txtB1001) = True Then
            MsgBox "該請假區間" & Trim(Mid(CboEmp(Index), 6)) & "正在休假，不可為職代！", vbExclamation
            CboEmp(Index).SetFocus
            Call CboEmp_GotFocus(Index)
            Cancel = True
            Exit Sub
         ElseIf CheckIsPersonRestSector(Left(CboEmp(Index), 5), txtB1004, Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), txtB1006, Trim(txtB1007_1.Text & ":" & txtB1007_2.Text), txtB1001) = True Then
'            strMsgText = "該請假區間此人員休假"
            'MsgBox "該請假區間此人員休假，不可為職代！", vbExclamation
            If MsgBox("該請假區間" & Trim(Mid(CboEmp(Index), 6)) & "也有休假，確定要選為職代嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
               CboEmp(Index).SetFocus
               Call CboEmp_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         '2017/1/10 END
         If CheckIsPersonRest(Left(CboEmp(Index), 5), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = True Then
'            If strMsgText <> "" Then strMsgText = strMsgText & "，並且"
            'strMsgText = strMsgText & "此人員今日休假，會延後簽核"
            MsgBox Trim(Mid(CboEmp(Index), 6)) & "今日休假，不可為職代！", vbExclamation
            CboEmp(Index).SetFocus
            Call CboEmp_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
'         If strMsgText <> "" Then
'            'MsgBox "此人員休假，不可為職代！", vbExclamation
'            If MsgBox(strMsgText & "，確定為職代嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then Exit Sub
'            CboEmp(Index).SetFocus
'            Call CboEmp_GotFocus(Index)
'            Cancel = True
'            Exit Sub
'         End If
      End If
   End If
End Sub

Private Sub CboBoss_GotFocus(Index As Integer)
   InverseTextBox CboBoss(Index)
End Sub

Private Sub CboBoss_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboBoss_LostFocus(Index As Integer)
   If CboBoss(Index).Text > "" And Len(Trim(CboBoss(Index).Text)) = 5 Then
      '抓取員工姓名
      CboBoss(Index).Text = SetCboStaffName(CboBoss(Index).Text)
   End If
End Sub

Private Sub CboBoss_Validate(Index As Integer, Cancel As Boolean)
   If CboBoss(Index) <> "" Then
      If Left(CboBoss(Index), 5) = txtB1003 Then
         MsgBox "不可為本人！", vbExclamation
         CboBoss(Index).SetFocus
         Call CboBoss_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(CboBoss(Index), 5)) = True Then
         Call CboBoss_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查 員工不可為”不寄信”
      If ChkStaffST14(Left(CboBoss(Index), 5)) = True Then
         Call CboBoss_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      If (Trim(CboBoss(1)) <> "" And Trim(CboBoss(0)) = "") Or _
         (Trim(CboBoss(2)) <> "" And Trim(CboBoss(1)) = "") Or _
         (Trim(CboBoss(3)) <> "" And Trim(CboBoss(2)) = "") Or _
         (Trim(CboBoss(4)) <> "" And Trim(CboBoss(3)) = "") Then
         MsgBox "請依序輸入審核主管！", vbExclamation
         CboBoss(Index).SetFocus
         Call CboBoss_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      If (CboBoss(1) <> "" And Left(CboBoss(1), 5) = Left(CboBoss(0), 5) And CboBoss(0).Enabled = True) Or _
         (CboBoss(2) <> "" And Left(CboBoss(2), 5) = Left(CboBoss(1), 5) And CboBoss(1).Enabled = True) Or _
         (CboBoss(3) <> "" And Left(CboBoss(3), 5) = Left(CboBoss(2), 5) And CboBoss(2).Enabled = True) Or _
         (CboBoss(4) <> "" And Left(CboBoss(4), 5) = Left(CboBoss(3), 5) And CboBoss(3).Enabled = True) Then
         MsgBox "資料重覆！", vbExclamation
         CboBoss(Index).SetFocus
         Call CboBoss_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String, ByVal strKEY04 As String, ByVal strKEY05 As String) As Boolean
   IsRecordExist = False
   
   If IsNull(strKEY01) Or strKEY01 = "" Then Exit Function
   If IsNull(strKEY02) Or strKEY02 = "" Then Exit Function
   If IsNull(strKEY03) Or strKEY03 = "" Then Exit Function
   If IsNull(strKEY04) Or strKEY04 = "" Then
      strKEY04 = strKEY02
      strKEY05 = strKEY03
   End If
   
   If CheckIsAbsenceExist(strKEY01, strKEY02, strKEY03, strKEY04, strKEY05, txtB1001, Left(Trim(CboB1002), 2)) = True Then IsRecordExist = True
   If IsRecordExist = True Then
      MsgBox "該筆記錄已存在", vbOKOnly, "新增資料"
      '先清空欄位值
      txtB1009 = Empty
      txtB1010 = Empty
      txtB1030 = Empty
      txtB101213 = Empty
      Call ClearFieldCbo
   End If
End Function

Private Sub cboSTime_GotFocus()
'   InverseTextBox cboSTime
End Sub

Private Sub cboSTime_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub cboSTime_Validate(Cancel As Boolean)
If cboSTime.Visible = True And cboSTime <> "" Then
   If Val(Format(cboSTime.Text, "hhmm")) > Val(Right("00" & txtB1005_1, 2) & Right("00" & txtB1005_2, 2)) Then
      Call cboSTime_GotFocus
      MsgBox "起日上班時段必須小於或等於起日請假時間!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
'   If Val(Format(cboSTime.Text, "hhmm")) > 2400 Then
'      Call cboSTime_GotFocus
'      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
'      Cancel = True
'      Exit Sub
'   End If
   If GetCountDayHour(False) = False Then
      Call cboSTime_GotFocus
      Cancel = True
      Exit Sub
   End If
Else
   If Chk1Day.Value = 1 Then '非整日
      If txtB1004 <> txtB1006 Then '跨日
         If cboSTime = "" Then
            Call cboSTime_GotFocus
            MsgBox "請輸入起日上班時段!", vbExclamation + vbOKOnly
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End If
CloseIme
End Sub

Private Sub cboETime_GotFocus()
'   InverseTextBox cboETime
End Sub

Private Sub cboETime_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub cboETime_Validate(Cancel As Boolean)
If cboETime.Visible = True And cboETime <> "" Then
   If Val(Format(cboETime.Text, "hhmm")) < Val(Right("00" & txtB1007_1, 2) & Right("00" & txtB1007_2, 2)) Then
      Call cboETime_GotFocus
      MsgBox "迄日下班時段必須大於或等於迄日請假時間!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
'   If Val(Format(cboETime.Text, "hhmm")) > 2400 Then
'      Call cboETime_GotFocus
'      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
'      Cancel = True
'      Exit Sub
'   End If
   If GetCountDayHour(False) = False Then
      Call cboETime_GotFocus
      Cancel = True
      Exit Sub
   End If
Else
   If Chk1Day.Value = 1 Then '非整日
      If txtB1004 <> txtB1006 Then '跨日
         If cboETime = "" Then
            Call cboETime_GotFocus
            MsgBox "請輸入迄日下班時段!", vbExclamation + vbOKOnly
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End If
CloseIme
End Sub

'設定職務代理人的下拉式選單
Private Sub SetABS001_1Combo(strST01 As String)
Dim strText As String
Dim kk As Integer
   
   For i = 0 To CboEmp.UBound
      CboEmp(i).Clear
      CboEmp(i).AddItem ""
   Next i
   strSql = "SELECT B0102,1 FROM ABS001,Staff WHERE B0101='" & strST01 & "' AND B0102 is not null AND B0102=ST01(+) AND ST04='1' " & _
      "Union SELECT B0103,2 FROM ABS001,Staff WHERE B0101='" & strST01 & "' AND B0103 is not null AND B0103=ST01(+) AND ST04='1' " & _
      "Union SELECT B0104,3 FROM ABS001,Staff WHERE B0101='" & strST01 & "' AND B0104 is not null AND B0104=ST01(+) AND ST04='1' " & _
      "Union SELECT B0105,4 FROM ABS001,Staff WHERE B0101='" & strST01 & "' AND B0105 is not null AND B0105=ST01(+) AND ST04='1' " & _
      "Union SELECT B0106,5 FROM ABS001,Staff WHERE B0101='" & strST01 & "' AND B0106 is not null AND B0106=ST01(+) AND ST04='1' " & _
      "Union SELECT B0107,6 FROM ABS001,Staff WHERE B0101='" & strST01 & "' AND B0107 is not null AND B0107=ST01(+) AND ST04='1' " & _
      "order by 2 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            If Not IsNull(RsTemp.Fields(0)) Then
               strText = SetCboStaffName(RsTemp.Fields(0))
               For i = 0 To CboEmp.UBound
                  CboEmp(i).AddItem strText
               Next i
            End If
            .MoveNext
         Loop
      End With
   End If
   For i = 0 To CboEmp.UBound
      If CboEmp(i).ListCount > 0 Then CboEmp(i).ListIndex = 0
   Next i
   
   'Modify By Sindy 2017/1/10 王副總提若職代的請假時間含蓋了請假人的請假時間, 則不可以出現
'   If (cmdSend.Visible = True And cmdSend.Enabled = True) Or _
'      (cmdagainSend.Visible = True And cmdagainSend.Enabled = True) Then
      If txtB1004 <> "" And txtB1005_1 <> "" And txtB1005_2 <> "" And _
         txtB1006 <> "" And txtB1007_1 <> "" And txtB1007_2 <> "" Then
         For i = 0 To CboEmp.UBound
            For kk = CboEmp(i).ListCount - 1 To 0 Step -1
               If Trim(CboEmp(i).List(kk)) <> "" Then
                  If CheckIsPersonRestSectorSame(CStr(Left(Trim(CboEmp(i).List(kk)), 5)), txtB1004, Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), txtB1006, Trim(txtB1007_1.Text & ":" & txtB1007_2.Text), txtB1001) = True Then
                     CboEmp(i).RemoveItem (kk)
                  End If
               End If
            Next kk
         Next i
      End If
'   End If
   '2017/1/10 END
End Sub

'設定審核主管的下拉式選單
Private Sub SetABS001_2Combo(strST01 As String)
Dim strText As String
   
   For i = 0 To CboBoss.UBound
      CboBoss(i).Clear
      CboBoss(i).AddItem ""
   Next i
   strSql = "SELECT B0108,1 FROM ABS001,Staff WHERE B0101='" & strST01 & "' AND B0108 is not null AND B0108=ST01(+) AND ST04='1' " & _
      "Union SELECT B0109,2 FROM ABS001,Staff WHERE B0101='" & strST01 & "' AND B0109 is not null AND B0109=ST01(+) AND ST04='1' " & _
      "Union SELECT B0110,3 FROM ABS001,Staff WHERE B0101='" & strST01 & "' AND B0110 is not null AND B0110=ST01(+) AND ST04='1' " & _
      "Union SELECT B0111,4 FROM ABS001,Staff WHERE B0101='" & strST01 & "' AND B0111 is not null AND B0111=ST01(+) AND ST04='1' " & _
      "order by 2 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         .MoveFirst
         Do While Not .EOF
            If Not IsNull(RsTemp.Fields(0)) Then
               strText = SetCboStaffName(RsTemp.Fields(0))
               For i = 0 To CboBoss.UBound
                  CboBoss(i).AddItem strText
               Next i
            End If
            .MoveNext
         Loop
      End With
   End If
   For i = 0 To CboBoss.UBound
      If CboBoss(i).ListCount > 0 Then CboBoss(i).ListIndex = 0
   Next i
End Sub

'檢查是否有權限
Private Function ChkLimitsIsOk() As Boolean
Dim rsTmp As New ADODB.Recordset
   
   ChkLimitsIsOk = False
   
   '開放人事處及電腦中心權限
   If GetStaffDepartment(strUserNum) = "M51" Or _
      GetStaffDepartment(strUserNum) = "M21" Then
      ChkLimitsIsOk = True
      Exit Function
   End If
   
   '開放當事人的審核主管才有權限
   strSql = "SELECT * FROM ABS001 " & _
            "WHERE B0101='" & txtB1003 & "' " & _
            "and (B0108='" & strUserNum & "' or B0109='" & strUserNum & "' or B0110='" & strUserNum & "' or B0111='" & strUserNum & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ChkLimitsIsOk = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
