VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100104_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "以收／發文日查詢"
   ClientHeight    =   5730
   ClientLeft      =   5850
   ClientTop       =   1550
   ClientWidth     =   8910
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8910
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   38
      Left            =   6555
      MaxLength       =   1
      TabIndex        =   43
      Top             =   750
      Width           =   492
   End
   Begin VB.CheckBox ChkCP159 
      Caption         =   "是否含已取消收文案件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   4140
      TabIndex        =   44
      Top             =   135
      Width           =   2325
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Index           =   1
      ItemData        =   "frm100104_1.frx":0000
      Left            =   7215
      List            =   "frm100104_1.frx":0010
      TabIndex        =   39
      Top             =   4950
      Width           =   1245
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Index           =   0
      ItemData        =   "frm100104_1.frx":0034
      Left            =   5580
      List            =   "frm100104_1.frx":0044
      TabIndex        =   38
      Top             =   4950
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   40
      Left            =   5580
      MaxLength       =   7
      TabIndex        =   42
      Top             =   5250
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   37
      Left            =   5960
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1050
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   36
      Left            =   6350
      MaxLength       =   1
      TabIndex        =   21
      Top             =   3168
      Width           =   492
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm100104_1.frx":0068
      Left            =   6350
      List            =   "frm100104_1.frx":0075
      TabIndex        =   24
      Top             =   3432
      Width           =   1785
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm100104_1.frx":0099
      Left            =   6350
      List            =   "frm100104_1.frx":00A9
      TabIndex        =   26
      Top             =   3732
      Width           =   1785
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   28
      Left            =   6350
      MaxLength       =   1
      TabIndex        =   19
      Top             =   2850
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   35
      Left            =   6345
      MaxLength       =   1
      TabIndex        =   13
      Top             =   1950
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   34
      Left            =   5580
      MaxLength       =   5
      TabIndex        =   29
      Top             =   4050
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   32
      Left            =   5580
      MaxLength       =   4
      TabIndex        =   34
      Top             =   4660
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   33
      Left            =   7215
      MaxLength       =   4
      TabIndex        =   35
      Top             =   4660
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   31
      Left            =   5085
      MaxLength       =   1
      TabIndex        =   17
      Top             =   2550
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   30
      Left            =   2800
      MaxLength       =   6
      TabIndex        =   41
      Text            =   " "
      Top             =   5250
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   29
      Left            =   1120
      MaxLength       =   6
      TabIndex        =   40
      Text            =   " "
      Top             =   5250
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   27
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   20
      Top             =   3150
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   26
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   18
      Top             =   2850
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   25
      Left            =   2800
      MaxLength       =   3
      TabIndex        =   8
      Top             =   1350
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   24
      Left            =   1120
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1350
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   23
      Left            =   2610
      MaxLength       =   1
      TabIndex        =   16
      Top             =   2550
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   22
      Left            =   2265
      MaxLength       =   1
      TabIndex        =   23
      Top             =   3450
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   21
      Left            =   1425
      MaxLength       =   1
      TabIndex        =   22
      Top             =   3450
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   20
      Left            =   1120
      TabIndex        =   25
      Top             =   3750
      Width           =   2930
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   19
      Left            =   2800
      MaxLength       =   9
      TabIndex        =   37
      Top             =   4950
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   18
      Left            =   1120
      MaxLength       =   9
      TabIndex        =   36
      Top             =   4950
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   17
      Left            =   2800
      MaxLength       =   9
      TabIndex        =   33
      Top             =   4650
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   16
      Left            =   1120
      MaxLength       =   9
      TabIndex        =   32
      Top             =   4650
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   15
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   14
      Top             =   2250
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   2800
      MaxLength       =   4
      TabIndex        =   31
      Top             =   4350
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   13
      Left            =   1120
      MaxLength       =   4
      TabIndex        =   30
      Top             =   4350
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   2800
      MaxLength       =   4
      TabIndex        =   28
      Top             =   4050
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   1120
      MaxLength       =   4
      TabIndex        =   27
      Top             =   4050
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   6350
      MaxLength       =   1
      TabIndex        =   15
      Top             =   2250
      Width           =   492
   End
   Begin VB.CommandButton CmdOk 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7605
      Style           =   1  '圖片外觀
      TabIndex        =   46
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6810
      Style           =   1  '圖片外觀
      TabIndex        =   45
      Top             =   60
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   5580
      MaxLength       =   6
      TabIndex        =   11
      Top             =   1650
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1120
      MaxLength       =   6
      TabIndex        =   10
      Top             =   1650
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   5580
      TabIndex        =   9
      Top             =   1350
      Width           =   3015
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2800
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1050
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1120
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1050
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1120
      TabIndex        =   3
      Top             =   750
      Width           =   2930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2800
      MaxLength       =   7
      TabIndex        =   2
      Top             =   450
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1120
      MaxLength       =   7
      TabIndex        =   1
      Top             =   450
      Width           =   1245
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1120
      MaxLength       =   1
      TabIndex        =   0
      Top             =   150
      Width           =   492
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1120
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "1"
      Top             =   1950
      Width           =   492
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "主管簽核：          （1.一般  2.特例）簽核"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   5640
      TabIndex        =   84
      Top             =   810
      Width           =   3150
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   0
      Left            =   6870
      TabIndex        =   56
      Top             =   1700
      Width           =   1575
      Size            =   "2778;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Index           =   1
      Left            =   2400
      TabIndex        =   57
      Top             =   1700
      Width           =   1575
      Size            =   "2778;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line8 
      X1              =   6930
      X2              =   7050
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "智慧局扣款日："
      Height          =   180
      Left            =   4305
      TabIndex        =   83
      Top             =   5295
      Width           =   1260
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "特殊商標："
      Height          =   180
      Left            =   4665
      TabIndex        =   82
      Top             =   5010
      Width           =   900
   End
   Begin VB.Line Line9 
      X1              =   6930
      X2              =   7050
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否只查詢新申請案：              Y: 僅查詢新申請案(含改請) "
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   9
      Left            =   4110
      TabIndex        =   81
      Top             =   1095
      Width           =   4605
   End
   Begin VB.Label lblName 
      Height          =   180
      Left            =   6960
      TabIndex        =   80
      Top             =   3210
      Width           =   1440
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "FCP工程師組別："
      Height          =   180
      Index           =   8
      Left            =   4950
      TabIndex        =   79
      Top             =   3210
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "台灣設計案件屬性："
      Height          =   180
      Index           =   1
      Left            =   4710
      TabIndex        =   78
      Top             =   3792
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利發明/新型案件屬性："
      Height          =   180
      Index           =   168
      Left            =   4305
      TabIndex        =   77
      Top             =   3492
      Width           =   2025
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否顯示PCT 案欄：            （Y：是, 僅查詢時有效 ）"
      Height          =   180
      Index           =   7
      Left            =   4680
      TabIndex        =   76
      Top             =   1992
      Width           =   4215
   End
   Begin VB.Label Label21 
      Caption         =   "(國際分類只查詢台灣已審定之發明新型專利案件)"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   4680
      TabIndex        =   75
      Top             =   4347
      Width           =   4080
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "國際分類："
      Height          =   180
      Left            =   4680
      TabIndex        =   74
      Top             =   4092
      Width           =   900
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人國籍："
      Height          =   180
      Left            =   4275
      TabIndex        =   73
      Top             =   4702
      Width           =   1290
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "送件方式：            （空白：全部 1：電子送件 2：紙本送件）"
      Height          =   180
      Index           =   6
      Left            =   4140
      TabIndex        =   72
      Top             =   2610
      Width           =   4770
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   2500
      X2              =   2620
      Y1              =   5350
      Y2              =   5350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FCP管制人："
      Height          =   180
      Index           =   8
      Left            =   60
      TabIndex        =   71
      Top             =   5310
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否顯示工作時數欄：　        （Y：是）"
      Height          =   180
      Index           =   5
      Left            =   4505
      TabIndex        =   70
      Top             =   2892
      Width           =   3180
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "PCT進入國家階段：                 （Y：國家階段）"
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   69
      Top             =   3210
      Width           =   3720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否顯示發文規費欄：            （Y：是）"
      Height          =   180
      Index           =   3
      Left            =   60
      TabIndex        =   68
      Top             =   2892
      Width           =   3180
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2505
      X2              =   2625
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "業  務  區："
      Height          =   180
      Left            =   60
      TabIndex        =   67
      Top             =   1392
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否只考慮有期限的案件資料：            （Y：是）"
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   66
      Top             =   2610
      Width           =   3900
   End
   Begin VB.Line Line7 
      X1              =   2025
      X2              =   2145
      Y1              =   3550
      Y2              =   3550
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "專利/商標種類："
      Height          =   180
      Left            =   60
      TabIndex        =   65
      Top             =   3510
      Width           =   1305
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "商品類別："
      Height          =   180
      Left            =   60
      TabIndex        =   64
      Top             =   3810
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "客    　戶："
      Height          =   180
      Left            =   60
      TabIndex        =   63
      Top             =   5010
      Width           =   900
   End
   Begin VB.Line Line6 
      X1              =   2500
      X2              =   2620
      Y1              =   5050
      Y2              =   5050
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "代  理  人："
      Height          =   180
      Left            =   60
      TabIndex        =   62
      Top             =   4710
      Width           =   900
   End
   Begin VB.Line Line5 
      X1              =   2500
      X2              =   2620
      Y1              =   4750
      Y2              =   4750
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否含內部收文資料：            （Y：含內部收文）"
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   61
      Top             =   2292
      Width           =   3900
   End
   Begin VB.Line Line4 
      X1              =   2500
      X2              =   2620
      Y1              =   4450
      Y2              =   4450
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "申請人國籍："
      Height          =   180
      Left            =   60
      TabIndex        =   60
      Top             =   4410
      Width           =   1080
   End
   Begin VB.Line Line3 
      X1              =   2500
      X2              =   2620
      Y1              =   4150
      Y2              =   4150
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   60
      TabIndex        =   59
      Top             =   4110
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "(請輸入民國年月日)"
      Height          =   180
      Left            =   4110
      TabIndex        =   58
      Top             =   510
      Width           =   1560
   End
   Begin VB.Line Line2 
      X1              =   2500
      X2              =   2620
      Y1              =   1150
      Y2              =   1150
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2500
      X2              =   2620
      Y1              =   550
      Y2              =   550
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否含來函資料：            （Y：含來函）"
      Height          =   180
      Index           =   0
      Left            =   4865
      TabIndex        =   55
      Top             =   2292
      Width           =   3180
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "列  印  別：              （1.查詢  2.印表）"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   60
      TabIndex        =   54
      Top             =   1992
      Width           =   2970
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   60
      TabIndex        =   53
      Top             =   1692
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "承  辦  人："
      Height          =   180
      Left            =   4680
      TabIndex        =   52
      Top             =   1692
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "不顯示案件性質："
      Height          =   180
      Left            =   4110
      TabIndex        =   51
      Top             =   1392
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   60
      TabIndex        =   50
      Top             =   1092
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "系統類別：                                                                        (ALL ：全部)"
      Height          =   180
      Left            =   60
      TabIndex        =   49
      Top             =   810
      Width           =   5175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日        期："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   60
      TabIndex        =   48
      Top             =   510
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查  詢  別：              （1.收文  2.發文）"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   47
      Top             =   210
      Width           =   2970
   End
End
Attribute VB_Name = "frm100104_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/18 改成Form2.0(lbl1(0),lbl1(1))
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/10 日期欄已修改
Option Explicit

Dim s As Integer, i As Integer, j As Integer
Dim StrTag As String, strSql As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Add By Cheng 2003/08/07
Public m_blnNoData As Boolean '是否有資料
'Add by Morgan 2005/11/9
Dim m_bolFinalCheck As Boolean '最後檢查控制
'Added by Lydia 2015/11/04
'列印用
Dim iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 16, ciFontSize = 11
Private Const ciStartX = 500
Private Const ciStartY = 500
Private Const ciColGap = 150
Private Const iCols = 13 '12
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim PLeft(0 To iCols + 1) As Integer
Dim m_Column(0 To iCols) As String
Dim strTemp(0 To iCols) As String
Public mQueryStr As String
Public mQueryStrEsc As String 'Added by Lydia 2019/11/01 利益衝突案件：排除限閱案件


'92.04.16 nick
Public Sub PubShowNextData()

Dim ii As Integer
Dim oText As TextBox

Select Case cmdState
Case 0 '確定
    cmdState = -1
    If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
        txt1_GotFocus 1
        Exit Sub
    End If
    
    'Add by Morgan 2011/4/1
    If DBDATE(txt1(1)) < "19221111" Then
      MsgBox "查詢起始日期輸入錯誤！", vbExclamation
      txt1(1).SetFocus
      txt1_GotFocus 1
      Exit Sub
    End If
    
    If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
        txt1_GotFocus 2
        Exit Sub
    End If
    'Add By Cheng 2003/11/14
    If Me.txt1(3).Text = "" Then
        MsgBox "請輸入系統類別!!!", vbExclamation + vbOKOnly
        Me.txt1(3).SetFocus
        Exit Sub
    End If
    If Len(Trim(txt1(0))) = 0 Or Len(Trim(txt1(1))) = 0 Or Len(Trim(txt1(2))) = 0 Or Len(Trim(txt1(9))) = 0 Then
        s = MsgBox("請檢查是否有必要條件忘了輸入．．．．", , "輸入條件不足")
        Exit Sub
    End If
    
    'Add by Morgan 2003/12/18
    If Not (txt1(24) = "" And txt1(25) = "") Then
      If txt1(24) = "" Then
         MsgBox "若要以業務區條件篩選則起迄值皆需輸入！"
         txt1(24).SetFocus
         Exit Sub
      ElseIf txt1(25) = "" Then
         MsgBox "若要以業務區條件篩選則起迄值皆需輸入！"
         txt1(25).SetFocus
         Exit Sub
      ElseIf (txt1(25) < txt1(24)) Then
         MsgBox "業務區條件迄值不可小於起值！", vbCritical
         txt1(25).SetFocus
         txt1_GotFocus (25)
         Exit Sub
      End If
    End If
    'Add end 2003/12/18
    
    For ii = 16 To 19
        If CheckKeyIn(ii) <> 0 Then
            Me.txt1(ii).SetFocus
            txt1_GotFocus ii
            Exit Sub
        End If
    Next ii
    
    m_bolFinalCheck = True
    For Each oText In txt1
      txt1_LostFocus oText.Index
      If m_bolFinalCheck = False Then
         Exit Sub
      End If
    Next
    
    'Add By Sindy 2019/1/21
    If Me.Combo3(0).Text <> "" Or Me.Combo3(1).Text <> "" Then
       If Val(Left(Me.Combo3(1).Text, 1)) < Val(Left(Me.Combo3(0).Text, 1)) Then
           MsgBox "特殊商標區間輸入錯誤!!!", vbExclamation + vbOKOnly
           Me.Combo3(0).SetFocus
           m_bolFinalCheck = False
           Exit Sub
       End If
    End If
    
    Me.Enabled = False
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/01/22 清除查詢印表記錄檔欄位
    '執行查詢功能
    If Trim(txt1(9)) = "1" Then
        If fnSaveParentForm(Me) = False Then
            Me.Enabled = True
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        frm100104_2.Label1(0).Caption = IIf(Me.txt1(0).Text = "1", "收文期間", "發文期間")
        frm100104_2.Show
        frm100104_2.StrMenu
        Screen.MousePointer = vbDefault
        Me.Enabled = True
        Exit Sub
    '執行列印功能
    Else
        'Modified by Lydia 2015/11/04 改成Print
'        m_blnNoData = True
        Screen.MousePointer = vbHourglass
        frm100104_2.Label1(0).Caption = IIf(Me.txt1(0).Text = "1", "收文期間", "發文期間")
        frm100104_2.Show
        frm100104_2.Hide
        frm100104_2.StrMenu1
'        'Modify By Cheng 2003/08/07
'        If m_blnNoData = False Then
'            Printer.Orientation = 2
'            If DataEnvironment1.Connection1.State = 1 Then
'                DataEnvironment1.Connection1.Close
'            End If
'            DataEnvironment1.Connections.Item(1).ConnectionString = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=M51CON;Persist Security Info=True"
'            DataEnvironment1.Connections.Item(1).Open
'            DataEnvironment1.rsCommand1.CursorLocation = adUseClient
'            Select Case Trim(txt1(0))
'            Case "1" '收文
'                If Me.txt1(23).Text <> "" Then
'                    '2010/9/10 MODIFY BY SONIA 改百年日期排序問題
'                    'strSQL = "SELECT R01002 AS 收文日,R01003 AS 本所案號,R01004 AS 案件名稱,R01005 AS 案件性質,R01006 AS 承辦人,R01007 AS 智權人員,R01008 AS 本所期限,R01009 AS 法定期限,R01010 AS 發文日,R01011 AS 是否出名,R01012 AS 點數,R01013 AS 申請人,R01014 AS 取消收文日 FROM R100104 where id='" & strUserNum & "' ORDER BY R01008,R01003 "
'                    strSql = "SELECT SUBSTR(' '||R01002,-9) AS 收文日,R01003 AS 本所案號,R01004 AS 案件名稱,R01005 AS 案件性質,R01006 AS 承辦人,R01007 AS 智權人員,SUBSTR(' '||R01008,-9) AS 本所期限,SUBSTR(' '||R01009,-9) AS 法定期限,SUBSTR(' '||R01010,-9) AS 發文日,R01011 AS 是否出名,R01012 AS 點數,R01013 AS 申請人,SUBSTR(' '||R01014,-9) AS 取消收文日 FROM R100104 where id='" & strUserNum & "' ORDER BY 本所期限,R01003 "
'                Else
'                    '2010/9/10 MODIFY BY SONIA 改百年日期排序問題
'                    'strSQL = "SELECT R01002 AS 收文日,R01003 AS 本所案號,R01004 AS 案件名稱,R01005 AS 案件性質,R01006 AS 承辦人,R01007 AS 智權人員,R01008 AS 本所期限,R01009 AS 法定期限,R01010 AS 發文日,R01011 AS 是否出名,R01012 AS 點數,R01013 AS 申請人,R01014 AS 取消收文日 FROM R100104 where id='" & strUserNum & "' ORDER BY R01002,R01003 "
'                    strSql = "SELECT  SUBSTR(' '||R01002,-9) AS 收文日,R01003 AS 本所案號,R01004 AS 案件名稱,R01005 AS 案件性質,R01006 AS 承辦人,R01007 AS 智權人員,SUBSTR(' '||R01008,-9) AS 本所期限,SUBSTR(' '||R01009,-9) AS 法定期限,SUBSTR(' '||R01010,-9) AS 發文日,R01011 AS 是否出名,R01012 AS 點數,R01013 AS 申請人,SUBSTR(' '||R01014,-9) AS 取消收文日 FROM R100104 where id='" & strUserNum & "' ORDER BY 收文日,R01003 "
'                End If
'            Case "2" '發文
'                If Me.txt1(23).Text <> "" Then
'                    '2010/9/10 MODIFY BY SONIA 改百年日期排序問題
'                    'strSQL = "SELECT R01002 AS 收文日,R01003 AS 本所案號,R01004 AS 案件名稱,R01005 AS 案件性質,R01006 AS 承辦人,R01007 AS 智權人員,R01008 AS 本所期限,R01009 AS 法定期限,R01010 AS 發文日,R01011 AS 是否出名,R01012 AS 點數,R01013 AS 申請人,R01014 AS 取消收文日 FROM R100104 where id='" & strUserNum & "' ORDER BY R01008,R01003 "
'                    strSql = "SELECT  SUBSTR(' '||R01002,-9) AS 收文日,R01003 AS 本所案號,R01004 AS 案件名稱,R01005 AS 案件性質,R01006 AS 承辦人,R01007 AS 智權人員,SUBSTR(' '||R01008,-9) AS 本所期限,SUBSTR(' '||R01009,-9) AS 法定期限,SUBSTR(' '||R01010,-9) AS 發文日,R01011 AS 是否出名,R01012 AS 點數,R01013 AS 申請人,SUBSTR(' '||R01014,-9) AS 取消收文日 FROM R100104 where id='" & strUserNum & "' ORDER BY 本所期限,R01003 "
'                Else
'                    '2010/9/10 MODIFY BY SONIA 改百年日期排序問題
'                    'strSQL = "SELECT R01002 AS 收文日,R01003 AS 本所案號,R01004 AS 案件名稱,R01005 AS 案件性質,R01006 AS 承辦人,R01007 AS 智權人員,R01008 AS 本所期限,R01009 AS 法定期限,R01010 AS 發文日,R01011 AS 是否出名,R01012 AS 點數,R01013 AS 申請人,R01014 AS 取消收文日 FROM R100104 where id='" & strUserNum & "' ORDER BY R01010,R01003 "
'                    strSql = "SELECT  SUBSTR(' '||R01002,-9) AS 收文日,R01003 AS 本所案號,R01004 AS 案件名稱,R01005 AS 案件性質,R01006 AS 承辦人,R01007 AS 智權人員,SUBSTR(' '||R01008,-9) AS 本所期限,SUBSTR(' '||R01009,-9) AS 法定期限,SUBSTR(' '||R01010,-9) AS 發文日,R01011 AS 是否出名,R01012 AS 點數,R01013 AS 申請人,SUBSTR(' '||R01014,-9) AS 取消收文日 FROM R100104 where id='" & strUserNum & "' ORDER BY 發文日,R01003 "
'                End If
'            Case Else
'            End Select
'            DataEnvironment1.rsCommand1.Open strSql, DataEnvironment1.Connections.Item(1), adOpenStatic, adLockReadOnly
'            DR100104.Orientation = rptOrientLandscape
'            DR100104.Sections(2).Controls("lbl1").Caption = strUserName
'            'edit by nickc 2006/03/17
'            'DR100104.Sections(2).Controls("lbl2").Caption = ChangeTStringToTDateString(ChangeWDateStringToTString(Date))
'            DR100104.Sections(2).Controls("lbl2").Caption = ChangeTStringToTDateString(strSrvDate(2))
'            DR100104.Sections(2).Controls("LBL3").Caption = Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@") & "-" & ChangeTStringToTDateString(txt1(2))
'            Select Case Trim(txt1(0))
'            Case "1"
'                DR100104.Sections(2).Controls("Label19").Caption = "收文日："
'                DR100104.Sections(2).Controls("Label1").Caption = "收文日"
'                DR100104.Sections(2).Controls("Label9").Caption = "發文日"
'            Case "2"
'                DR100104.Sections(2).Controls("Label19").Caption = "發文日："
'                DR100104.Sections(2).Controls("Label1").Caption = "發文日"
'                DR100104.Sections(2).Controls("Label9").Caption = "收文日"
'            Case Else
'            End Select
'
'            PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
'
'            DR100104.PrintReport: DoEvents
'            'Add By Cheng 2003/08/20
'            '直到DataReport的列印動作完畢時, 才載出DataReport
'            While DR100104.AsyncCount > 0
'                DoEvents
'            Wend
'
'            PUB_RestoreOsPrt 'Add by Morgan 2010/2/23
'
'            Unload DR100104
'        End If
        PrintData
        Unload frm100104_2
        'If Me.m_blnNoData = False Then ShowPrintOk
        'end 2015/11/04
        Screen.MousePointer = vbDefault
    End If
    Me.Enabled = True
    Me.Show
Case 1 '結束
     fnCloseAllFrm100
Case Else
End Select
End Sub


Private Sub cmdOK_Click(Index As Integer)
'add by nickc 2007/01/12
If Len(Trim(Me.txt1(3).Text)) = 0 Then
    Me.txt1(3).Text = "ALL"
End If
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
''92.04.16 nick 以下無效
''Add By Cheng 2002/04/24
'Dim ii As Integer
'
'Select Case Index
'Case 0
'      'Add By Cheng 2002/03/15
'      If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
'         txt1_GotFocus 1
'         Exit Sub
'      End If
'      If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
'         txt1_GotFocus 2
'         Exit Sub
'      End If
'      'Modify By Cheng 2002/03/14
''    'Add By Cheng 2002/01/31
''    txt1_LostFocus 3
'    If Len(Trim(txt1(0))) = 0 Or Len(Trim(txt1(1))) = 0 Or Len(Trim(txt1(2))) = 0 Or Len(Trim(txt1(9))) = 0 Then
'       s = MsgBox("請檢查是否有必要條件忘了輸入．．．．", , "輸入條件不足")
'       Exit Sub
'    End If
'    'Add By Cheng 2002/04/24
'    For ii = 16 To 19
'      If CheckKeyIn(ii) <> 0 Then
'         Me.txt1(ii).SetFocus
'         txt1_GotFocus ii
'         Exit Sub
'      End If
'    Next ii
'
'    Me.Enabled = False
'    '執行查詢功能
'    If Trim(txt1(9)) = "1" Then
'        Screen.MousePointer = vbHourglass
'         'Add By Cheng 2002/02/08
'         frm100104_2.Label1(0).Caption = IIf(Me.txt1(0).Text = "1", "收文期間", "發文期間")
'        frm100104_2.Show
'        'frm100104_2.Hide
'
'        frm100104_2.StrMenu
'        Screen.MousePointer = vbDefault
'        Me.Hide
'        'frm100104_2.Show
'        Do
'        DoEvents
'        If bolToEndByNick = True Then Unload Me: Exit Sub
'        Loop Until Not frm100104_2.Visible
'        Unload frm100104_2
'    '執行列印功能
'    Else
'        Screen.MousePointer = vbHourglass
'         'Add By Cheng 2002/02/08
'         frm100104_2.Label1(0).Caption = IIf(Me.txt1(0).Text = "1", "收文期間", "發文期間")
'        frm100104_2.Show
'        frm100104_2.Hide
'        frm100104_2.StrMenu1
'        Printer.Orientation = 2
'        If DataEnvironment1.Connection1.State = 1 Then
'            DataEnvironment1.Connection1.Close
'        End If
'        DataEnvironment1.Connections.Item(1).ConnectionString = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=M51CON;Persist Security Info=True"
'        DataEnvironment1.Connections.Item(1).Open
'        DataEnvironment1.rsCommand1.CursorLocation = adUseClient
'        Select Case Trim(txt1(0))
'        Case "1"
'            strSql = "SELECT  R01002 AS 收文日,R01003 AS 本所案號,R01004 AS 案件名稱,R01005 AS 案件性質,R01006 AS 承辦人,R01007 AS 智權人員,R01008 AS 本所期限,R01009 AS 法定期限,R01010 AS 發文日,R01011 AS 是否出名,R01012 AS 點數,R01013 AS 申請人,R01014 AS 取消收文日 FROM R100104 where id='" & strUserNum & "' ORDER BY R01002,R01003 "
'        Case "2"
'            strSql = "SELECT  R01002 AS 收文日,R01003 AS 本所案號,R01004 AS 案件名稱,R01005 AS 案件性質,R01006 AS 承辦人,R01007 AS 智權人員,R01008 AS 本所期限,R01009 AS 法定期限,R01010 AS 發文日,R01011 AS 是否出名,R01012 AS 點數,R01013 AS 申請人,R01014 AS 取消收文日 FROM R100104 where id='" & strUserNum & "' ORDER BY R01002,R01003 "
'        Case Else
'        End Select
'        DataEnvironment1.rsCommand1.Open strSql, DataEnvironment1.Connections.Item(1), adOpenStatic, adLockReadOnly
'        DR100104.Orientation = rptOrientLandscape
'        DR100104.Sections(2).Controls("lbl1").Caption = strUserName
'        'edit by nickc 2006/03/17
'        'DR100104.Sections(2).Controls("lbl2").Caption = ChangeTStringToTDateString(ChangeWDateStringToTString(Date))
'        DR100104.Sections(2).Controls("lbl2").Caption = ChangeTStringToTDateString(strSrvDate(2))
'        DR100104.Sections(2).Controls("LBL3").Caption = Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@") & "-" & ChangeTStringToTDateString(txt1(2))
'        Select Case Trim(txt1(0))
'        Case "1"
'            DR100104.Sections(2).Controls("Label19").Caption = "收文日："
'            DR100104.Sections(2).Controls("Label1").Caption = "收文日"
'            DR100104.Sections(2).Controls("Label9").Caption = "發文日"
'        Case "2"
'            DR100104.Sections(2).Controls("Label19").Caption = "發文日："
'            DR100104.Sections(2).Controls("Label1").Caption = "發文日"
'            DR100104.Sections(2).Controls("Label9").Caption = "收文日"
'        Case Else
'        End Select
'
'        PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
'        DR100104.PrintReport
'        PUB_RestoreOsPrt 'Add by Morgan 2010/2/23
'
'        Unload frm100104_2
'        ShowPrintOk
'        Screen.MousePointer = vbDefault
'    End If
'    Me.Enabled = True
'    Me.Show
'Case 1
'     Unload Me
'Case Else
'End Select
End Sub

'Add By Sindy 2014/7/9
Private Sub Combo1_Validate(Cancel As Boolean)
   If Combo1 <> "" Then
      Combo1 = Left(Combo1, 1) + "." + PUB_GetCaseAttributeName(Left(Combo1, 1), "1")
      If Combo1 = Left(Combo1, 1) + "." Then
         Combo1 = Left(Combo1, 1)
         Cancel = True
         Combo1.SetFocus
      End If
   End If
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
   If Combo2 <> "" Then
      Combo2 = Left(Combo2, 1) + "." + PUB_GetCaseAttributeName(Left(Combo2, 1), "3")
      If Combo2 = Left(Combo2, 1) + "." Then
         Combo2 = Left(Combo2, 1)
         Cancel = True
         Combo2.SetFocus
      End If
   End If
End Sub
'2014/7/9 End

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   bolToEndByNick = False
   MoveFormToCenter Me
   txt1(3) = Systemkind_g
   
   bolToEndByNick = False
   '92.04.16 nick
   cmdState = -1
   
   'Add By Sindy 2019/7/11 特殊商標改用拉下式選單
   StrSQLa = "Select * From SpecialPatentTrademark Where SPT01='2' order by spt02 asc"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      rsA.MoveFirst
      Combo3(0).Clear
      Combo3(1).Clear
      Do While Not rsA.EOF
         Combo3(0).AddItem rsA.Fields("SPT02") & " " & rsA.Fields("SPT03")
         Combo3(1).AddItem rsA.Fields("SPT02") & " " & rsA.Fields("SPT03")
         rsA.MoveNext
      Loop
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100104_1 = Nothing
End Sub

Private Sub txt1_Change(Index As Integer)
   Select Case Index
      Case 16
         If Len(Me.txt1(Index).Text) > 6 Then Exit Sub
         Me.txt1(17).Text = Me.txt1(Index).Text
      Case 18
         If Len(Me.txt1(Index).Text) > 6 Then Exit Sub
         Me.txt1(19).Text = Me.txt1(Index).Text
      Case 36
          'Add by Lydia 2014/11/18 增加FCP工程師組別
           lblName = PUB_GetFCPGrpName(txt1(Index))
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    Select Case Index
    'Add by Morgan 2004/9/30
    'Modified by Morgan 2019/11/6 +31
    'Modified by Sindy 2019/11/6 +38
    Case 9, 31, 38
      If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
            KeyAscii = 0
      End If
      
    Case 23 '是否只考慮有期限的資料
        If KeyAscii <> 8 And KeyAscii <> 89 Then
            KeyAscii = 0
        End If
    'add by nick 2004/08/23
    'edit by nick 2005/02/04
    'Case 26
    Case 26, 27
        If KeyAscii <> 8 And KeyAscii <> 89 Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   Select Case Index
   Case 0
         If InStr(1, "12 ", txt1(0)) = 0 Then
            s = MsgBox("請輸入 1 或 2 !!", , "輸入錯誤")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            m_bolFinalCheck = False
            Exit Sub
         End If
   'Modified by Lydia 2019/05/16 智慧局扣款日txt1(40)
   Case 1, 2, 40
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            m_bolFinalCheck = False
            Exit Sub
         End If
         If Index = 2 Then
           If Not nickChgRan(txt1(1), txt1(2), "日期") Then
               txt1(1).SetFocus
               txt1_GotFocus (1)
               m_bolFinalCheck = False
               Exit Sub
           End If
         'Added by Lydia 2019/05/16 智慧局扣款日
         ElseIf Index = 40 Then
             If Trim(txt1(Index).Text) <> "" Then txt1(26).Text = "Y" '是否顯示發文規費欄
         End If
   Case 3 '系統類別
         'Modify By Cheng 2002/03/14
   '      'Add By Cheng 2002/01/07
   '      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(3))
        'Added by Lydia 2016/02/24 檢查跨部門權限
        txt1(Index) = Replace(txt1(Index), " ", "")
        If Len(Me.txt1(Index)) > 0 And Me.txt1(Index) <> "ALL" Then
           If PUB_CheckSKAddCross(strUserNum, Systemkind_g, True, Me.txt1(Index)) = False Then
               txt1(Index).SetFocus
               txt1_GotFocus Index
               m_bolFinalCheck = False
               Exit Sub
           End If
        End If
        'end 2016/02/24
        
   'Modify By Sindy 2012/2/22 +33
   Case 5, 12, 14, 17, 19, 33
        If Index = 17 Or Index = 19 Then
           If Mid(txt1(Index - 1), 1, 6) <> Mid(txt1(Index), 1, 6) Then
               s = MsgBox("前6碼必須相同！", , "錯誤！")
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               m_bolFinalCheck = False
               Exit Sub
           End If
        End If
         If RunNick(txt1(Index - 1), txt1(Index)) Then
            txt1(Index - 1).SetFocus
            txt1_GotFocus (Index - 1)
            m_bolFinalCheck = False
            Exit Sub
         End If
   Case 6
   Case 7 '承辦人
         If Len(txt1(7)) <> 0 Then
            strSql = "SELECT ST02 FROM STAFF WHERE ST01='" & txt1(7) & "'"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                If Not IsNull(adoRecordset.Fields(0)) Then
                     Lbl1(0).Caption = adoRecordset.Fields(0)
                Else
                     Lbl1(0).Caption = ""
                End If
            Else
                Lbl1(0).Caption = ""
                s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
                txt1(Index).SetFocus
                txt1_GotFocus (Index)
                m_bolFinalCheck = False
                Exit Sub
            End If
            CheckOC
            'Add By Cheng 2002/03/05
            'Me.txt1(15).Text = "Y"     '2008/1/28 cancel by sonia
         Else
            Lbl1(0).Caption = ""
         End If
         
   Case 8 '智權人員
         If Len(txt1(8)) <> 0 Then
            strSql = "SELECT ST02 FROM STAFF WHERE ST01='" & txt1(8) & "'"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                If Not IsNull(adoRecordset.Fields(0)) Then
                     Lbl1(1).Caption = adoRecordset.Fields(0)
                Else
                     Lbl1(1).Caption = ""
                End If
               'Add By Cheng 2002/03/05
               'Me.txt1(15).Text = Empty   '2008/1/28 cancel by sonia
            Else
                Lbl1(1).Caption = ""
                s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
                txt1(Index).SetFocus
                txt1_GotFocus (Index)
                m_bolFinalCheck = False
                Exit Sub
            End If
            CheckOC
         
         Else
              Lbl1(1).Caption = ""
         End If
   'Modified by Morgan 2019/11/6 +31
   Case 9, 31
        If InStr(1, "12 ", txt1(9)) = 0 Then
            s = MsgBox("請輸入 1 或 2 !!", , "輸入錯誤")
            txt1(9).SetFocus
            txt1(9).SelStart = 0
            txt1(9).SelLength = Len(txt1(9))
            m_bolFinalCheck = False
            Exit Sub
         End If
   'Modify By Cheng 2002/03/05
   'Case 10
   '     If InStr(1, "Yy ", txt1(10)) = 0 Then
   '         s = MsgBox("請輸入 Y 或空白!!", , "輸入錯誤")
   '         txt1(10).SetFocus
   '         txt1(10).SelStart = 0
   '         txt1(10).SelLength = Len(txt1(10))
   '     End If
   'Add By Sindy 2012/3/7 +35
   'Add by Lydia 2015/02/12 + 37 (+原本未+23,26,27,31,28)
   Case 10, 15, 35, 37, 23, 26, 27, 28 '是否含來函資料, 是否含內部收文資料
        If InStr(1, "Yy ", txt1(Index)) = 0 Then
            s = MsgBox("請輸入 Y 或空白!!", , "輸入錯誤")
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            m_bolFinalCheck = False
        End If
   'Add By Cheng 2003/06/02
   '檢查專利/商標種類區間
   Case 22
       If Me.txt1(21).Text <> "" And Me.txt1(22).Text <> "" Then
           If Me.txt1(22).Text < Me.txt1(21).Text Then
               MsgBox "專利/商標種類區間輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.txt1(21).SetFocus
               txt1_GotFocus 21
               m_bolFinalCheck = False
               Exit Sub
           End If
       End If
   Case Else
   End Select
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   'If Len(Trim(txt1(0))) <> 0 And Len(Trim(txt1(1))) <> 0 And Len(Trim(txt1(2))) <> 0 And Len(Trim(txt1(9))) <> 0 Then
   '   'Me.Enabled = True
   '    CmdOk(0).Enabled = True
   '    'cmdok(0).SetFocus
   'Else
   '    CmdOk(0).Enabled = False
   'End If
End Sub

Private Function CheckKeyIn(Index) As Integer
CheckKeyIn = 0
Select Case Index
'add by nick 加強欄位判斷檢查
Case 0
      If InStr(1, "12 ", txt1(Index)) = 0 Then
         s = MsgBox("請輸入 1 或 2 !!", , "輸入錯誤")
         CheckKeyIn = -1
         Exit Function
      End If
'Add by Lydia 2015/02/12 + 37 (+原本未+23,26,27,31,28)
'Case 10, 15 '是否含來函資料, 是否含內部收文資料
Case 10, 15, 35, 37, 23, 26, 27, 31, 28
'end 2015/02/12
     If InStr(1, "Yy ", txt1(Index)) = 0 Then
         s = MsgBox("請輸入 Y 或空白!!", , "輸入錯誤")
         CheckKeyIn = -1
         Exit Function
     End If
Case 9
     If InStr(1, "12 ", txt1(Index)) = 0 Then
         s = MsgBox("請輸入 1 或 2 !!", , "輸入錯誤")
         CheckKeyIn = -1
         Exit Function
      End If
'add end
Case 16 '代理人(起)
   If Len(Me.txt1(Index).Text) > 0 Then
      If Left(Me.txt1(Index).Text, 1) <> "Y" Then
         s = MsgBox("代理人代碼輸入錯誤!!!", , "輸入錯誤")
         CheckKeyIn = -1
         Exit Function
      End If
      If Len(Me.txt1(17).Text) > 0 Then
         If Left(Me.txt1(17).Text, 6) <> Left(Me.txt1(Index).Text, 6) Then
            s = MsgBox("代理人代碼前六碼必須相同!!!", , "輸入錯誤")
            CheckKeyIn = -1
            Exit Function
         End If
      End If
   End If
Case 17 '代理人(迄)
   If Len(Me.txt1(Index).Text) > 0 Then
      If Left(Me.txt1(Index).Text, 1) <> "Y" Then
         s = MsgBox("代理人代碼輸入錯誤!!!", , "輸入錯誤")
         CheckKeyIn = -1
         Exit Function
      End If
      If Len(Me.txt1(16).Text) > 0 Then
         If Left(Me.txt1(16).Text, 6) <> Left(Me.txt1(Index).Text, 6) Then
            s = MsgBox("代理人代碼前六碼必須相同!!!", , "輸入錯誤")
            CheckKeyIn = -1
            Exit Function
         End If
      End If
   End If
Case 18 '客戶(起)
   If Len(Me.txt1(Index).Text) > 0 Then
      If Left(Me.txt1(Index).Text, 1) <> "X" Then
         s = MsgBox("客戶代碼輸入錯誤!!!", , "輸入錯誤")
         CheckKeyIn = -1
         Exit Function
      End If
      If Len(Me.txt1(19).Text) > 0 Then
         If Left(Me.txt1(19).Text, 6) <> Left(Me.txt1(Index).Text, 6) Then
            s = MsgBox("客戶代碼前六碼必須相同!!!", , "輸入錯誤")
            CheckKeyIn = -1
            Exit Function
         End If
      End If
   End If
Case 19 '客戶(迄)
   If Len(Me.txt1(Index).Text) > 0 Then
      If Left(Me.txt1(Index).Text, 1) <> "X" Then
         s = MsgBox("客戶代碼輸入錯誤!!!", , "輸入錯誤")
         CheckKeyIn = -1
         Exit Function
      End If
      If Len(Me.txt1(18).Text) > 0 Then
         If Left(Me.txt1(18).Text, 6) <> Left(Me.txt1(Index).Text, 6) Then
            s = MsgBox("客戶代碼前六碼必須相同!!!", , "輸入錯誤")
            CheckKeyIn = -1
            Exit Function
         End If
      End If
   End If
End Select
End Function

'Add by Morgan 2003/12/18
Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
      
   Select Case Index
      Case 25  '業務區迄
         If txt1(25) <> "" And txt1(25) < txt1(24) Then
            MsgBox "業務區條件迄值不可小於起值！", vbCritical
            txt1_GotFocus (Index)
            Cancel = True
         End If
   End Select
End Sub
'Added by Lydia 2015/11/04 改成Print
Private Sub PrintData()
Dim rsRe As New ADODB.Recordset
Dim intR As Integer
    intI = 0
    'Modify By Sindy 2023/4/14 +,簽
    strSql = "select 收文日,本所案號,案件名稱,案件性質,承辦人,智權人員,簽,本所期限,發文日,法定期限,出名,點數,substr(申請人,1,6),取消收文日 from (" & mQueryStr & ") x1 "
    'Added by Lydia 2019/11/01 利益衝突案件：排除限閱案件
    If Trim(mQueryStrEsc) <> "" Then
        strSql = strSql & " where 本所案號 not in (" & GetAddStr(mQueryStrEsc) & ") "
    End If
    strSql = strSql & " order by 1"
    'end 2019/11/01
    Set rsRe = ClsLawReadRstMsg(intI, strSql)
    InsertQueryLog (rsRe.RecordCount)
    If intI = 1 Then
        Printer.PaperSize = 9 'A4
        Printer.Orientation = 2 '橫印
        Printer.Font.Name = "新細明體"
        Printer.Font.Size = ciFontSize
        Printer.Font.Bold = False
        Printer.Font.Underline = False
        lngPageHeight = Printer.ScaleHeight
        lngPageWidth = Printer.ScaleWidth
        lngLineHeight = 300
        Call GetPleft
        Call SetColumnName
        
        Erase strTemp
        iPage = 1
        rsRe.MoveFirst
        PrintHeader '列印表頭
        Do While Not rsRe.EOF
           '限定資料長度
            For intR = 0 To iCols
               strTemp(intR) = "" & rsRe.Fields(intR)
               Select Case intR
                   Case 2 '案件名稱
                       strTemp(intR) = PUB_StrToStr(strTemp(intR), 22)
                   Case 3 '案件性質
                       strTemp(intR) = PUB_StrToStr(strTemp(intR), 8)
                   Case 11 '申請人
                       strTemp(intR) = PUB_StrToStr(strTemp(intR), 8)
               End Select
            Next intR
           '列印明細
            For intR = 0 To iCols
               Printer.CurrentX = PLeft(intR)
               Printer.CurrentY = iPrint
               Printer.Print strTemp(intR)
            Next intR
            PrintNewLine
            rsRe.MoveNext
        Loop
        Printer.EndDoc
        ShowPrintOk
    End If
End Sub
Private Sub GetPleft() '明細表邊界
Dim intX As Integer

Erase PLeft

   PLeft(0) = ciStartX
   PLeft(1) = PLeft(0) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(7, "　")) + ciColGap
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(12, "　")) + ciColGap
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(4, "　")) + ciColGap 'Add By Sindy 2023/4/14
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(1, "　")) + ciColGap
   PLeft(8) = PLeft(7) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(9) = PLeft(8) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(10) = PLeft(9) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(11) = PLeft(10) + Printer.TextWidth(String(2, "　")) + ciColGap
   PLeft(12) = PLeft(11) + Printer.TextWidth(String(3, "　")) + ciColGap
   PLeft(13) = PLeft(12) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(14) = lngPageWidth - ciStartX
End Sub
'明細資料的抬頭
Private Sub SetColumnName()
Dim tmpArr As Variant
Dim tmpStr As String
   
   'Modify By Sindy 2023/4/14 +|簽
   tmpStr = "收文日|本所案號|案件名稱|案件性質|承辦人|智權人員|簽|本所期限|發文日|法定期限|出名|點數|申請人|取消收文日"
   tmpArr = Split(tmpStr, "|")
   
   For intI = 0 To UBound(tmpArr)
       m_Column(intI) = tmpArr(intI)
   Next intI
   
End Sub
Private Sub PrintHeader()
Dim strPTmp As String
Dim pa1 As Integer
iPrint = ciStartY
Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.Font.Underline = True
strPTmp = "以收/發文日查詢"
pa1 = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentX = pa1
Printer.CurrentY = iPrint
Printer.Print strPTmp
 
'title line = 2
PrintNewLine
iPrint = iPrint + 150
Printer.Font.Size = ciFontSize
Printer.Font.Bold = False
Printer.Font.Underline = False
strPTmp = IIf(txt1(0) = "1", "收", "發") & "文日：" & CFDate(txt1(1)) & " -" & CFDate(txt1(2))
Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
Printer.CurrentY = iPrint
Printer.Print strPTmp
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "製表日期：" & CFDate(strSrvDate(2))
'title line = 3
PrintNewLine
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "列表人：" & strUserName
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & Printer.Page
    
'title line = 4
PrintNewLine
PrintLine

For pa1 = 0 To iCols
   Printer.CurrentX = PLeft(pa1)
   Printer.CurrentY = iPrint
   Printer.Print m_Column(pa1)
Next pa1

PrintNewLine
PrintLine 2

End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint

      iPage = iPage + 1
      Printer.NewPage
      PrintHeader
   End If
End Sub
Private Sub PrintLine(Optional iNum As Integer = 1)
    
    Printer.Line (PLeft(0), iPrint)-(lngPageWidth - 300, iPrint)
    If iNum = 1 Then
       iPrint = iPrint + 150
    ElseIf iNum = 2 Then
          iPrint = iPrint + 50
          Printer.Line (PLeft(0), iPrint)-(lngPageWidth - 300, iPrint)
          iPrint = iPrint + 150
    End If
End Sub
'end 2015/11/04
