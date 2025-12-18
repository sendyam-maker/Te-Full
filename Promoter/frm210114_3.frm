VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210114_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "委任契約書-T"
   ClientHeight    =   6900
   ClientLeft      =   1836
   ClientTop       =   2400
   ClientWidth     =   9264
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9264
   Begin VB.CheckBox Chk2 
      Caption         =   "延展"
      Height          =   240
      Index           =   2
      Left            =   4605
      TabIndex        =   79
      Top             =   2580
      Width           =   1560
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "變更"
      Height          =   240
      Index           =   3
      Left            =   6165
      TabIndex        =   78
      Top             =   2580
      Width           =   1620
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "申請"
      Height          =   240
      Index           =   1
      Left            =   2610
      TabIndex        =   77
      Top             =   2580
      Width           =   1965
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "查名"
      Height          =   240
      Index           =   0
      Left            =   1035
      TabIndex        =   76
      Top             =   2580
      Width           =   1035
   End
   Begin VB.CheckBox ChkSeal 
      Caption         =   "用印"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   75
      Top             =   6570
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0FFFF&
      Caption         =   "空白列印"
      Height          =   330
      Index           =   5
      Left            =   3600
      Style           =   1  '圖片外觀
      TabIndex        =   74
      Top             =   60
      Width           =   920
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm210114_3.frx":0000
      Left            =   6750
      List            =   "frm210114_3.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   43
      Top             =   6540
      Width           =   2475
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋委任人(&Q)"
      Height          =   330
      Left            =   5460
      TabIndex        =   72
      Top             =   4680
      Width           =   1365
   End
   Begin VB.TextBox txtPCnt 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Left            =   7680
      MaxLength       =   1
      TabIndex        =   44
      Text            =   "2"
      Top             =   90
      Width           =   270
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印        份"
      Height          =   330
      Index           =   0
      Left            =   7200
      TabIndex        =   49
      Top             =   60
      Width           =   1100
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面"
      Height          =   330
      Index           =   1
      Left            =   8300
      TabIndex        =   50
      Top             =   60
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "清空資料"
      Height          =   330
      Index           =   2
      Left            =   6300
      TabIndex        =   48
      Top             =   60
      Width           =   920
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   360
      Left            =   30
      TabIndex        =   70
      Top             =   30
      Width           =   3645
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   660
         Style           =   2  '單純下拉式
         TabIndex        =   45
         Top             =   30
         Width           =   2840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "印表機"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   71
         Top             =   90
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "儲存文檔"
      Height          =   330
      Index           =   3
      Left            =   4500
      TabIndex        =   46
      Top             =   60
      Width           =   920
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "讀取文檔"
      Height          =   330
      Index           =   4
      Left            =   5400
      TabIndex        =   47
      Top             =   60
      Width           =   920
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "補充理由、補充答辯"
      Height          =   240
      Index           =   17
      Left            =   2610
      TabIndex        =   22
      Top             =   3540
      Width           =   1965
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "廢止"
      Height          =   240
      Index           =   14
      Left            =   4605
      TabIndex        =   19
      Top             =   3300
      Width           =   1560
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "再授權"
      Height          =   240
      Index           =   8
      Left            =   1035
      TabIndex        =   13
      Top             =   3060
      Width           =   1035
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "分割"
      Height          =   240
      Index           =   4
      Left            =   1035
      TabIndex        =   9
      Top             =   2820
      Width           =   1035
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "其他："
      Height          =   240
      Index           =   22
      Left            =   4605
      TabIndex        =   27
      Top             =   3780
      Width           =   900
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "移轉"
      Height          =   240
      Index           =   5
      Left            =   2610
      TabIndex        =   10
      Top             =   2820
      Width           =   1965
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "註冊費"
      Height          =   240
      Index           =   21
      Left            =   2610
      TabIndex        =   26
      Top             =   3780
      Width           =   1965
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "異議"
      Height          =   240
      Index           =   12
      Left            =   1035
      TabIndex        =   17
      Top             =   3300
      Width           =   1035
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "訴願"
      Height          =   240
      Index           =   15
      Left            =   6165
      TabIndex        =   20
      Top             =   3300
      Width           =   1620
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "補呈文件"
      Height          =   240
      Index           =   20
      Left            =   1035
      TabIndex        =   25
      Top             =   3780
      Width           =   1035
   End
   Begin VB.CheckBox Chk2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "上訴審行政訴訟"
      Height          =   240
      Index           =   19
      Left            =   480
      TabIndex        =   24
      Top             =   3900
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "商標"
      Height          =   240
      Index           =   0
      Left            =   1050
      TabIndex        =   2
      Top             =   1095
      Width           =   1035
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "團體標章"
      Height          =   240
      Index           =   1
      Left            =   2145
      TabIndex        =   3
      Top             =   1095
      Width           =   1035
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "證明標章"
      Height          =   240
      Index           =   2
      Left            =   3240
      TabIndex        =   4
      Top             =   1095
      Width           =   1035
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "授權"
      Height          =   240
      Index           =   6
      Left            =   4605
      TabIndex        =   11
      Top             =   2820
      Width           =   1560
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "英文證明"
      Height          =   240
      Index           =   10
      Left            =   4605
      TabIndex        =   15
      Top             =   3060
      Width           =   1560
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "評定"
      Height          =   240
      Index           =   13
      Left            =   2610
      TabIndex        =   18
      Top             =   3300
      Width           =   1965
   End
   Begin VB.CheckBox Chk2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "第一審行政訴訟"
      Height          =   240
      Index           =   18
      Left            =   4590
      TabIndex        =   23
      Top             =   3540
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "質權登記"
      Height          =   240
      Index           =   9
      Left            =   2610
      TabIndex        =   14
      Top             =   3060
      Width           =   1965
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "補證"
      Height          =   240
      Index           =   7
      Left            =   6165
      TabIndex        =   12
      Top             =   2820
      Width           =   1620
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "申請理由"
      Height          =   240
      Index           =   11
      Left            =   6165
      TabIndex        =   16
      Top             =   3060
      Width           =   1620
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "答辯"
      Height          =   240
      Index           =   16
      Left            =   1035
      TabIndex        =   21
      Top             =   3540
      Width           =   1035
   End
   Begin VB.OptionButton opt1 
      Caption         =   "會稿"
      Height          =   210
      Index           =   0
      Left            =   825
      TabIndex        =   31
      Top             =   4755
      Width           =   1080
   End
   Begin VB.OptionButton opt1 
      Caption         =   "不會稿"
      Height          =   210
      Index           =   1
      Left            =   1965
      TabIndex        =   32
      Top             =   4755
      Width           =   1080
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7590
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   5520
      TabIndex        =   28
      Top             =   3750
      Width           =   1410
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "2487;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   435
      Width           =   6945
      VariousPropertyBits=   671105051
      MaxLength       =   52
      Size            =   "12250;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   750
      Width           =   9180
      VariousPropertyBits=   671105051
      MaxLength       =   72
      Size            =   "16192;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   2
      Left            =   1395
      TabIndex        =   5
      Top             =   1350
      Width           =   7830
      VariousPropertyBits=   671105051
      MaxLength       =   56
      Size            =   "13811;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1395
      TabIndex        =   6
      Top             =   1650
      Width           =   7830
      VariousPropertyBits=   671105051
      MaxLength       =   58
      Size            =   "13811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   1395
      TabIndex        =   7
      Top             =   1950
      Width           =   7830
      VariousPropertyBits=   671105051
      MaxLength       =   58
      Size            =   "13811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1380
      TabIndex        =   8
      Top             =   2250
      Width           =   7830
      VariousPropertyBits=   671105051
      MaxLength       =   58
      Size            =   "13811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1155
      TabIndex        =   29
      Top             =   4095
      Width           =   2460
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "4339;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   1155
      TabIndex        =   30
      Top             =   4410
      Width           =   2460
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "4339;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   1410
      TabIndex        =   33
      Top             =   5025
      Width           =   7830
      VariousPropertyBits=   671105051
      MaxLength       =   54
      Size            =   "13811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   1410
      TabIndex        =   34
      Top             =   5331
      Width           =   2280
      VariousPropertyBits=   671105051
      MaxLength       =   18
      Size            =   "4022;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   11
      Left            =   4440
      TabIndex        =   35
      Top             =   5331
      Width           =   4785
      VariousPropertyBits=   671105051
      MaxLength       =   26
      Size            =   "8440;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   12
      Left            =   1410
      TabIndex        =   36
      Top             =   5637
      Width           =   7830
      VariousPropertyBits=   671105051
      MaxLength       =   54
      Size            =   "13811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   13
      Left            =   1410
      TabIndex        =   37
      Top             =   5943
      Width           =   2280
      VariousPropertyBits=   671105051
      MaxLength       =   18
      Size            =   "4022;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   14
      Left            =   4440
      TabIndex        =   38
      Top             =   5943
      Width           =   4785
      VariousPropertyBits=   671105051
      MaxLength       =   26
      Size            =   "8440;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   15
      Left            =   1410
      TabIndex        =   39
      Top             =   6249
      Width           =   7830
      VariousPropertyBits=   671105051
      MaxLength       =   54
      Size            =   "13811;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   16
      Left            =   1410
      TabIndex        =   40
      Top             =   6555
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1235;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   17
      Left            =   2520
      TabIndex        =   41
      Top             =   6570
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1235;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   300
      Index           =   18
      Left            =   3600
      TabIndex        =   42
      Top             =   6555
      Width           =   705
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "1235;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "受任人："
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   6000
      TabIndex        =   73
      Top             =   6600
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "請輸入數字"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4155
      TabIndex        =   69
      Top             =   4470
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "請輸入數字"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4140
      TabIndex        =   68
      Top             =   4140
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "委辦案件內容或案件名稱："
      Height          =   180
      Left            =   45
      TabIndex        =   67
      Top             =   465
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "商品/服務類別："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   66
      Top             =   1410
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "商品/服務名稱："
      Height          =   180
      Left            =   45
      TabIndex        =   65
      Top             =   1710
      Width           =   1305
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請種類："
      Height          =   180
      Left            =   45
      TabIndex        =   64
      Top             =   1125
      Width           =   900
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "性　　質："
      Height          =   180
      Left            =   75
      TabIndex        =   63
      Top             =   2610
      Width           =   900
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "前   酬   金："
      Height          =   180
      Left            =   45
      TabIndex        =   62
      Top             =   4140
      Width           =   990
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "後   酬   金："
      Height          =   180
      Left            =   45
      TabIndex        =   61
      Top             =   4455
      Width           =   990
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "元整"
      Height          =   180
      Left            =   3750
      TabIndex        =   60
      Top             =   4140
      Width           =   360
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "元整"
      Height          =   180
      Left            =   3750
      TabIndex        =   59
      Top             =   4455
      Width           =   360
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "甲方：委任人："
      Height          =   180
      Left            =   45
      TabIndex        =   58
      Top             =   5040
      Width           =   1260
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "ID  NO.："
      Height          =   180
      Left            =   585
      TabIndex        =   57
      Top             =   5391
      Width           =   735
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代表人："
      Height          =   180
      Left            =   3735
      TabIndex        =   56
      Top             =   5391
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "地　址："
      Height          =   180
      Left            =   600
      TabIndex        =   55
      Top             =   5670
      Width           =   720
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "電　話："
      Height          =   180
      Left            =   600
      TabIndex        =   54
      Top             =   6003
      Width           =   720
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "傳　真："
      Height          =   180
      Left            =   3735
      TabIndex        =   53
      Top             =   6003
      Width           =   720
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "乙方：經手人："
      Height          =   180
      Left            =   45
      TabIndex        =   52
      Top             =   6300
      Width           =   1260
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "中　華　民　國　　　　　年　　　　　月　　　　　日"
      Height          =   180
      Left            =   45
      TabIndex        =   51
      Top             =   6600
      Width           =   4500
   End
End
Attribute VB_Name = "frm210114_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/22 改成Form2.0 ; txt1(index)、Printer改成Word列印
'Memo by Lydia 2019/07/01 表單名稱:案件委任契約書=>委任契約書
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim SeekPrint As Integer, SeekPrintL As Integer
Dim iCount As Integer
'Add By Sindy 2010/4/30
Public m_strCustCode As String
Public m_blnOneRec As Boolean
'2010/4/30 End
Dim strNowCustNo As String 'Add by Amy 2016/08/19 客戶編號
Dim iPrintC As Integer 'Added by Lydia 2017/03/28 目前列印第幾份
Dim bolAddSeal As Boolean 'Added by Lydia 2017/03/28 是否用印
Dim d_Left As Double, d_Top As Double 'Added by Lydia 2017/04/25 印表機實際列印的左邊界、右邊界
Dim strPrinter As String 'Added by Lydia 2017/04/28
Dim strDetail As String 'Move by Lydia 2017/05/16 記錄內容(從StrMenu移出來)
Dim strCompSeal As String 'Added by Lydia 2020/03/25 記錄"公司名稱|用印編號",用,區隔
'Added by Lydia 2022/01/22  加入圖片用(Word)
Const msoFalse = 0
Const msoLineSolid = 1
Const msoLineSingle = 1
Const msoTrue = -1
Const msoPictureAutomatic = 1
'end 2022/01/22
Dim m_TempPDF As String 'Added by Lydia 2022/01/22
Dim m_TempFN As String 'Added by Lydia 2022/01/22

Private Sub Chk1_Click(Index As Integer)
'先清空
Dim i As Integer

   If Chk1(Index).Value = vbChecked Then
       For i = 0 To 2
           If i <> Index Then
               Chk1(i).Value = vbUnchecked
           End If
       Next i
   End If
End Sub

'Add By Sindy 2010/4/29
Private Sub cmdFind_Click()
   Dim strCmpName As String, strMsg As String 'Add by Amy 2016/08/19
   
   If Me.txt1(9).Text = "" Then
      MsgBox "請輸入委任人中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txt1(9).SetFocus
      Exit Sub
   End If

   frm090801_1.m_DouChk = False '可複選  'add by Lydia 2014/9/22
   Set frm090801_1.m_frm0908A = Me
   
   frm090801_1.m_strCustChnName = Me.txt1(9).Text
   frm090801_1.lblName.Caption = Me.txt1(9).Text
   m_blnOneRec = False
   m_strCustCode = ""
   If frm090801_1.StrMenu = True Then
      If frm090801_1.m_blnOneRec = False Then
         frm090801_1.Show vbModal
      End If
      m_blnOneRec = frm090801_1.m_blnOneRec
      m_strCustCode = frm090801_1.m_strCustCode
      Unload frm090801_1
   Else
      Unload frm090801_1
   End If
    Combo2.Tag = "": strNowCustNo = "" 'Add by Amy 2016/08/19
   If m_blnOneRec = True And m_strCustCode <> "" Then
     'Add by Amy 2016/08/19 記錄收據公司別(放於SetCustTxt前避免m_strCustCode被清空)
      strNowCustNo = m_strCustCode
      strCmpName = "Y"
      Combo2.Tag = GetReceiptCmp(Left(strNowCustNo, 8), Mid(strNowCustNo, 9, 1), "T", "000", False, strCmpName, Me.Name)
      If Combo2.Tag <> MsgText(601) And Combo2 <> MsgText(601) And Combo2.Tag <> frm210114_1.GetComp(Combo2) Then
        strMsg = "您輸入之收據公司別「" & Combo2 & "」與客戶檔設定值「" & strCmpName & "」不同" & vbCrLf & _
                     "是否依客戶檔設定覆蓋您的輸入值？"
        If MsgBox(strMsg, vbYesNo + vbCritical) = vbYes Then
            'Modified by Lydia 2024/08/06
            'Combo2 = strCmpName
            Call Pub_SetCboListIdx(Me.Combo2, strCmpName)
        End If
      ElseIf strCmpName = MsgText(601) Then
        Combo2.ListIndex = 0
      Else
        'Modified by Lydia 2024/08/06
        'Combo2 = strCmpName
        Call Pub_SetCboListIdx(Me.Combo2, strCmpName)
      End If
      'end 2016/08/19
      'Modify by Amy 2021/05/13 +if 讀取文檔要保留原文檔內容
      If Me.ActiveControl.Name = "cmdFind" Then
        Call SetCustTxt(m_strCustCode)
      End If
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim tb As Control
Dim op As OptionButton
Dim ck As CheckBox
Dim fN As Integer
Dim strBuffer As String
Dim AllObj(0 To 48) As String
Dim AllObjV As Variant
'Add by Amy 2016/08/19 目前收據公司別
Dim strNowCmp As String

   Select Case Index
      Case 0
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(0)) = "" And Trim(txt1(1)) = "" Then
              MsgBox "委辦內容或案件名稱不可空白！", vbInformation, "錯誤！"
              txt1(0).SetFocus
              txt1_GotFocus 0
              Exit Sub
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(2)) = "" Then
              MsgBox "商品/服務類別不可空白！", vbInformation, "錯誤！"
              txt1(2).SetFocus
              txt1_GotFocus 2
              Exit Sub
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(3)) = "" And Trim(txt1(4)) = "" And Trim(txt1(5)) = "" Then
              MsgBox "商品/服務名稱不可空白！", vbInformation, "錯誤！"
              txt1(3).SetFocus
              txt1_GotFocus 3
              Exit Sub
          End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(7)) = "" And Trim(txt1(8)) = "" Then
              MsgBox "費用最少輸一個不可空白！", vbInformation, "錯誤！"
              txt1(7).SetFocus
              txt1_GotFocus 7
              Exit Sub
          End If
          If opt1(0).Value <> True And opt1(1).Value <> True Then
              MsgBox "會不會稿要選擇一項！", vbInformation, "錯誤！"
              Exit Sub
          End If
      '    If txt1(9) = "" Then
      '        MsgBox "委任人不可空白！", vbInformation, "錯誤！"
      '        txt1(9).SetFocus
      '        txt1_GotFocus 9
      '        Exit Sub
      '    End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(15)) = "" Then
              MsgBox "經手人不可空白！", vbInformation, "錯誤！"
              txt1(15).SetFocus
              txt1_GotFocus 15
              Exit Sub
          End If
          
          If Chk1(0).Value = vbUnchecked And Chk1(1).Value = vbUnchecked And Chk1(2).Value = vbUnchecked Then
              MsgBox "申請種類最少選一個！", vbInformation, "錯誤！"
              Exit Sub
          End If
          If Chk2(0).Value = vbUnchecked And Chk2(1).Value = vbUnchecked And Chk2(2).Value = vbUnchecked And Chk2(3).Value = vbUnchecked And Chk2(4).Value = vbUnchecked And Chk2(5).Value = vbUnchecked And Chk2(6).Value = vbUnchecked And Chk2(7).Value = vbUnchecked And Chk2(8).Value = vbUnchecked And Chk2(9).Value = vbUnchecked And Chk2(10).Value = vbUnchecked And Chk2(11).Value = vbUnchecked And _
             Chk2(12).Value = vbUnchecked And Chk2(13).Value = vbUnchecked And Chk2(14).Value = vbUnchecked And Chk2(15).Value = vbUnchecked And Chk2(16).Value = vbUnchecked And Chk2(17).Value = vbUnchecked And Chk2(18).Value = vbUnchecked And Chk2(19).Value = vbUnchecked And Chk2(20).Value = vbUnchecked And Chk2(21).Value = vbUnchecked And Chk2(22).Value = vbUnchecked Then
              MsgBox "案件性質最少選一個！", vbInformation, "錯誤！"
              Exit Sub
          End If
          If Chk2(22).Value = vbChecked Then
              If Trim(txt1(6)) = "" Then
                  MsgBox "請輸入其他案件性質！", vbInformation, "錯誤！"
                  txt1(6).SetFocus
                  txt1_GotFocus 6
                  Exit Sub
              End If
          End If
                
         '2011/10/18 ADD BY SONIA 檢查四縣市地址
         If txt1(12) <> "" Then
           If CheckTaiwanAddr(txt1(12), "000", "甲方委任人地址") = False Then
              txt1(12).SetFocus
              txt1_GotFocus (12)
              Exit Sub
           End If
         End If
         '2011/10/18 END
         'Add by Amy 2016/08/19 +受任人不可為空
         If Combo2 = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
         End If
          '2009/11/13 MODIFY BY SONIA 杜副總提出
      '    If txt1(16) = "" Or txt1(17) = "" Or txt1(18) = "" Then
      '        MsgBox "日期需要正確！", vbInformation, "錯誤！"
      '        txt1(16).SetFocus
      '        txt1_GotFocus 16
      '        Exit Sub
      '    End If
          'Modified by Lydia 2017/03/28 +Trim清除空白鍵
          If Trim(txt1(16)) = "" Or Trim(txt1(17)) = "" Or Trim(txt1(18)) = "" Then
             If MsgBox("契約書日期不完整，是否確定？", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
               txt1(16).SetFocus
               txt1_GotFocus 16
               Exit Sub
             End If
          End If
      '2009/11/13 END
                
          'Added by Lydia 2017/03/28
          If ChkSeal.Value = 1 Then
             'Modified by Lydia 2017/04/27 PDF印表機不需詢問,並且份數改為1份
             If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
             Else
                If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                   Exit Sub
                End If
             End If
             'end 2017/04/27
             bolAddSeal = True
          Else
             bolAddSeal = False
          End If
          'end 2017/03/28
          
          'Modified by Lydia 2017/04/13
'          For iCount = 1 To Val(txtPCnt) 'edit by nickc 2006/09/27 2
'              'add by nickc 2006/06/05
'              Set Printer = Printers(Combo1.ListIndex)
'              Screen.MousePointer = vbHourglass
'              DoEvents
'              StrMenu
'          Next iCount
          'Modified by Lydia 2022/01/25 改成Word直接印
          'Call Print2PDF(False)
          Call runWordProc(False)
          PUB_SetOsDefaultPrinter strPrinter
          'end 2022/01/25
          
          'Add by Amy 2016/08/19 畫面與客戶檔收據公司別不同更新客戶檔
          strNowCmp = frm210114_1.GetComp(Combo2)
          If Combo2.Tag <> strNowCmp Then
             Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          End If
          'end 2016/08/19
          Screen.MousePointer = vbDefault
          Call RunEndProc(True) 'Added by Lydia 2022/01/22 刪除暫存檔
          'Modified by Lydia 2022/01/22 判斷是否有列印
          'ShowPrintOk 'Added by Lydia 2017/04/11
          If m_TempPDF <> "" Then ShowPrintOk
      Case 1
          frm210114.Show
          Unload Me
      Case 2
          For Each tb In txt1
              tb.Text = Empty
          Next
          For Each op In opt1
              op.Value = False
          Next
          For Each ck In Chk1
              ck.Value = vbUnchecked
          Next
          For Each ck In Chk2
              ck.Value = vbUnchecked
          Next
      Case 3
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowSave
          If cd1.FileName <> "" Then
              AllObj(0) = "案件委任契約書-T"
              For iCount = 1 To 19
                  AllObj(iCount) = txt1(iCount - 1).Text
              Next iCount
              AllObj(20) = Chk1(0).Value
              AllObj(21) = Chk1(1).Value
              AllObj(22) = Chk1(2).Value
              AllObj(23) = Chk2(0).Value
              AllObj(24) = Chk2(1).Value
              AllObj(25) = Chk2(2).Value
              AllObj(26) = Chk2(3).Value
              AllObj(27) = Chk2(4).Value
              AllObj(28) = Chk2(5).Value
              AllObj(29) = Chk2(6).Value
              AllObj(30) = Chk2(7).Value
              AllObj(31) = Chk2(8).Value
              AllObj(32) = Chk2(9).Value
              AllObj(33) = Chk2(10).Value
              AllObj(34) = Chk2(11).Value
              AllObj(35) = Chk2(12).Value
              AllObj(36) = Chk2(13).Value
              AllObj(37) = Chk2(14).Value
              AllObj(38) = Chk2(15).Value
              AllObj(39) = Chk2(16).Value
              AllObj(40) = Chk2(17).Value
              AllObj(41) = Chk2(18).Value
              AllObj(42) = Chk2(19).Value
              AllObj(43) = Chk2(20).Value
              AllObj(44) = Chk2(21).Value
              AllObj(45) = Chk2(22).Value
              AllObj(46) = IIf(opt1(0).Value = True, "0", "1")
              AllObj(47) = IIf(opt1(1).Value = True, "0", "1")
              AllObj(48) = Combo2.Text 'Add By Sindy 2011/3/23
              strBuffer = Join(AllObj, Chr(30))
              strBuffer = StrEncrypt(strBuffer)
              fN = FreeFile
              Open cd1.FileName For Output As fN
              Print #fN, strBuffer
              Close #fN
          End If
          'Add by Amy 2016/08/19 畫面與客戶檔收據公司別不同更新客戶檔
          strNowCmp = frm210114_1.GetComp(Combo2)
          If Combo2 <> MsgText(601) And Combo2.Tag <> strNowCmp Then
             Call UpdReceiptCmp(strNowCustNo, strNowCmp)
          End If
          'end 2016/08/19
      Case 4
          cd1.Filter = "Contract Files(*.Con)|*.Con"
          cd1.InitDir = GetMyDocPath
          On Error GoTo DialogCancel
          cd1.CancelError = True
          cd1.ShowOpen
          If cd1.FileName <> "" Then
              fN = FreeFile
              Open cd1.FileName For Input As fN
              Input #fN, strBuffer
              Close #fN
              strBuffer = StrDecrypt(strBuffer)
              AllObjV = Split(strBuffer, Chr(30))
              If AllObjV(0) = "案件委任契約書-T" Then
                  cmdOK_Click 2
                  For iCount = 1 To 19
                       txt1(iCount - 1).Text = AllObjV(iCount)
                  Next iCount
                  Chk1(0).Value = AllObjV(20)
                  Chk1(1).Value = AllObjV(21)
                  Chk1(2).Value = AllObjV(22)
                  Chk2(0).Value = AllObjV(23)
                  Chk2(1).Value = AllObjV(24)
                  Chk2(2).Value = AllObjV(25)
                  Chk2(3).Value = AllObjV(26)
                  Chk2(4).Value = AllObjV(27)
                  Chk2(5).Value = AllObjV(28)
                  Chk2(6).Value = AllObjV(29)
                  Chk2(7).Value = AllObjV(30)
                  Chk2(8).Value = AllObjV(31)
                  Chk2(9).Value = AllObjV(32)
                  Chk2(10).Value = AllObjV(33)
                  Chk2(11).Value = AllObjV(34)
                  Chk2(12).Value = AllObjV(35)
                  Chk2(13).Value = AllObjV(36)
                  Chk2(14).Value = AllObjV(37)
                  Chk2(15).Value = AllObjV(38)
                  Chk2(16).Value = AllObjV(39)
                  Chk2(17).Value = AllObjV(40)
                  Chk2(18).Value = AllObjV(41)
                  Chk2(19).Value = AllObjV(42)
                  Chk2(20).Value = AllObjV(43)
                  Chk2(21).Value = AllObjV(44)
                  Chk2(22).Value = AllObjV(45)
                  opt1(0).Value = IIf(Val(AllObjV(46)) = 0, True, False)
                  opt1(1).Value = IIf(Val(AllObjV(47)) = 0, True, False)
                  'Modify by Amy 2016/08/19 避免空值會Error
                  If AllObjV(48) = MsgText(601) Then
                    Combo2.ListIndex = 0
                  Else
                    Combo2.Text = AllObjV(48) 'Add By Sindy 2011/3/23
                  End If
                  'end 2016/08/19
                  
                  'Add By Sindy 2011/1/21 檢查地址欄
                  '委任人地址
                  If txt1(9).Text <> "" And txt1(12).Text <> "" Then
                     If CheckCustomerAddr(1, Trim(txt1(9).Text), Trim(txt1(12).Text), "委任人", True) = False Then
                        txt1(12).SetFocus
                     End If
                  End If
                  '2011/1/21 End
                  'Add by Amy 2016/08/19 讀取收據公司別
                  cmdFind_Click
              Else
                  MsgBox "錯誤格式，此份內容並非 內商 格式！", vbExclamation
              End If
          End If
      'Added by Lydia 2017/03/28 空白委任書
      Case 5
          If Trim(Combo2) = "" Then
             MsgBox "受任人不可為空白！", vbInformation, "錯誤！"
             Combo2.SetFocus
             Exit Sub
          End If
          'Modified by Lydia 2017/04/17 文雄表示用印由下方勾選,可直接空白列印
          If ChkSeal.Value = 1 Then
            If (InStr(UCase(Combo1.Text), "BATCH") > 0 Or InStr(UCase(Combo1.Text), "WRITER") > 0 Or InStr(UCase(Combo1.Text), "PDF") > 0) And Pub_StrUserSt03 <> "M51" Then
               MsgBox "空白用印的印表機不可選擇PDF列印！", vbInformation, "錯誤！"
               Combo1.SetFocus
               Exit Sub
            End If
            'Modified by Lydia 2017/04/27 PDF印表機不需詢問,並且份數改為1份
            If InStr(UCase(Combo1.Text), "PDF") > 0 Then
                txtPCnt = "1"
            Else
               If MsgBox("用印的委任書需選擇彩色印表機，是否已選擇？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If
             'end 2017/04/27
            bolAddSeal = True
          End If
          'end 2017/04/17

          Call cmdOK_Click(2) '清空資料
          'Modified by Lydia 2022/01/25 改成Word直接印
          'Call Print2PDF(True)
          Call runWordProc(True)
          PUB_SetOsDefaultPrinter strPrinter
          'end 2022/01/25
          
          m_strCustCode = ""
          bolAddSeal = False
          Screen.MousePointer = vbDefault
          Call RunEndProc(False) 'Added by Lydia 2022/01/22 刪除暫存檔
          'Modified by Lydia 2022/01/22 判斷是否有列印
          'ShowPrintOk 'Added by Lydia 2017/04/11
          If m_TempPDF <> "" Then ShowPrintOk
      'end 2017/03/28
      Case Else
   End Select
   Exit Sub
DialogCancel:
End Sub

'Add by Morgan 2011/2/24 只要鍵盤有動作就不斷線
Private Sub Form_KeyPress(KeyAscii As Integer)
   If UCase(Forms(0).Name) = "MDIMAIN" Then Forms(0).tmrConnect.Tag = 0
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
   
   PUB_InitForm210114 Forms(0), Me 'Added by Lydia 2017/05/19 委任契約書表單大於主表單，控制主表單放大。
   MoveFormToCenter Me
   'Modified by Lydia 2017/04/28 改用模組
   'strSql = Printer.DeviceName
   'SeekPrintL = Printer.Orientation
   'For i = 0 To Printers.Count - 1
   '    Set Printer = Printers(i)
   '    Combo1.AddItem Printer.DeviceName, j
   '    j = j + 1
   '    If Printer.DeviceName = strSql Then
   '        SeekPrint = i
   '    End If
   'Next i
   'Set Printer = Printers(SeekPrint)
   'Combo1.Text = Combo1.List(SeekPrint)
   PUB_SetPrinter Me.Name, Me.Combo1, strPrinter, , , , , True 'Modified by Morgan 2020/10/30 +只顯示有效的印表機參數

    'Added by Lydia 2017/04/17 先用模組抓所有印表機後,排除特定印表機
    'Remove by Lydia 2017/06/07 改直接列印
    'For i = 0 To Combo1.ListCount - 1
    '   If InStr(UCase(Combo1.List(i)), "PDFCREATOR") > 0 And Trim(Combo1.List(i)) <> "" Then
    '      Combo1.RemoveItem i
    '      'If i = SeekPrint Then Combo1.Text = Combo1.List(0) 'Remove by Lydia 2017/04/28
    '   End If
    'Next
    'end 2017/04/17
    'end 2017/06/07
    
   'Added by Lydia 2020/03/25 設定公司別下拉選項
   Call PUB_SetCboTofrm210114(Me.Name, Me.Combo2, strCompSeal)
   
   'Modify by Amy 2016/08/19
   'Combo2.Text = Combo2.List(1) 'Add By Sindy 2011/3/23
    Combo2.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '還原預設印表機
   'Modified by Lydia 2017/04/28 記錄表單的印表機
   'Set Printer = Printers(SeekPrint)
   'Printer.Orientation = SeekPrintL
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   'end 2017/04/28
   
   Call RunEndProc(False) 'Added by Lydia 2022/01/22 刪除暫存檔
   
   Set frm210114_3 = Nothing
End Sub

'Modified by Lydia 2017/03/28
'Sub StrMenu()
Sub StrMenu(Optional ByVal bolSpace As Boolean = False)
Dim iY As Integer
Dim tmpI As Integer
'Modify By Sindy 2010/5/12
'Dim iStr(1 To 49) As String
Dim iStr(1 To 48) As String
'2010/5/12 End
Dim tBoxTop As Integer
'Added by Lydia 2017/03/28
Dim tObj As New StdPicture
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String
'end 2017/03/28

   iStr(1) = "商標案件委任契約書"
   iStr(2) = "委任人(甲方)茲委任受任人(乙方)辦理國內商標案件，雙方同意條件如下："
   iStr(3) = "第一條　委辦範圍："
   iStr(4) = "    一、委辦內容或案件名稱：" & StrToStr(txt1(0) & String(52, " "), 26)
   iStr(5) = "　　　　" & StrToStr(txt1(1) & String(72, " "), 36)
   iStr(6) = "　　二、申請種類：" & IIf(Chk1(0).Value = 1, "■", "□") & "商標  " & IIf(Chk1(1).Value = 1, "■", "□") & "團體標章  " & IIf(Chk1(2).Value = 1, "■", "□") & "證明標章"
   iStr(7) = "　　　　商品/服務類別：" & StrToStr(txt1(2) & String(56, " "), 28) & "類"
   iStr(8) = "　　　　商品/服務名稱：" & StrToStr(txt1(3) & String(58, " "), 29)
   iStr(9) = "　　　　　　　　　　　 " & StrToStr(txt1(4) & String(58, " "), 29)
   iStr(10) = "　　　　　　　　　　　 " & StrToStr(txt1(5) & String(58, " "), 29)
   'Modify By Sindy 2010/5/12
   'iStr(11) = "　　三、案件性質：" & IIf(Chk2(0).Value = 1, "■", "□") & "查名　  　" & IIf(Chk2(1).Value = 1, "■", "□") & "申請　　　　　　　　" & IIf(Chk2(2).Value = 1, "■", "□") & "延展　　　　　　" & IIf(Chk2(3).Value = 1, "■", "□") & "變更　　　　　　"
   'iStr(12) = "　　　　　　　　　" & IIf(Chk2(4).Value = 1, "■", "□") & "分割  　　" & IIf(Chk2(5).Value = 1, "■", "□") & "移轉　　　　　　　　" & IIf(Chk2(6).Value = 1, "■", "□") & "授權　　　　　　" & IIf(Chk2(7).Value = 1, "■", "□") & "補證　　　　　　"
   'iStr(13) = "　　　　　　　　　" & IIf(Chk2(8).Value = 1, "■", "□") & "再授權　　" & IIf(Chk2(9).Value = 1, "■", "□") & "質權登記　　　　　　" & IIf(Chk2(10).Value = 1, "■", "□") & "英文證明　　　　" & IIf(Chk2(11).Value = 1, "■", "□") & "申請理由　　　　"
   'iStr(14) = "　　　　　　　　　" & IIf(Chk2(12).Value = 1, "■", "□") & "異議　　  " & IIf(Chk2(13).Value = 1, "■", "□") & "評定　　　　　　　　" & IIf(Chk2(14).Value = 1, "■", "□") & "廢止　　　　　　" & IIf(Chk2(15).Value = 1, "■", "□") & "訴願　　　　　　"
   'iStr(15) = "　　　　　　　　　" & IIf(Chk2(16).Value = 1, "■", "□") & "答辯　　  " & IIf(Chk2(17).Value = 1, "■", "□") & "補充理由、補充答辯　" & IIf(Chk2(18).Value = 1, "■", "□") & "第一審行政訴訟　" & IIf(Chk2(19).Value = 1, "■", "□") & "上訴審行政訴訟　"
   'iStr(16) = "　　　　　　　　　" & IIf(Chk2(20).Value = 1, "■", "□") & "補呈文件  " & IIf(Chk2(21).Value = 1, "■", "□") & "調卷　　　　　　　　" & IIf(Chk2(22).Value = 1, "■", "□") & "其他：" & StrToStr(txt1(6).Text & String(20, " "), 10)
   iStr(11) = "　　三、案件性質："
   iStr(12) = "　　　　" & IIf(Chk2(0).Value = 1, "■", "□") & "查名　　" & IIf(Chk2(1).Value = 1, "■", "□") & "申請　　" & IIf(Chk2(2).Value = 1, "■", "□") & "延展　　 " & IIf(Chk2(3).Value = 1, "■", "□") & "變更　　 " & IIf(Chk2(4).Value = 1, "■", "□") & "分割　　 " & IIf(Chk2(5).Value = 1, "■", "□") & "移轉　　" & IIf(Chk2(6).Value = 1, "■", "□") & "授權　　"
   iStr(13) = "　　　　" & IIf(Chk2(7).Value = 1, "■", "□") & "補證　　" & IIf(Chk2(8).Value = 1, "■", "□") & "再授權　" & IIf(Chk2(9).Value = 1, "■", "□") & "質權登記 " & IIf(Chk2(10).Value = 1, "■", "□") & "英文證明 " & IIf(Chk2(11).Value = 1, "■", "□") & "申請理由 " & IIf(Chk2(12).Value = 1, "■", "□") & "異議　　" & IIf(Chk2(13).Value = 1, "■", "□") & "評定　　"
   iStr(14) = "　　　　" & IIf(Chk2(14).Value = 1, "■", "□") & "廢止　　" & IIf(Chk2(15).Value = 1, "■", "□") & "訴願　　" & IIf(Chk2(16).Value = 1, "■", "□") & "答辯　　 " & IIf(Chk2(17).Value = 1, "■", "□") & "補充理由、補充答辯  " & IIf(Chk2(18).Value = 1, "■", "□") & "第一審行政訴訟　"
   iStr(15) = "　　　　" & IIf(Chk2(19).Value = 1, "■", "□") & "上訴審行政訴訟　　" & IIf(Chk2(20).Value = 1, "■", "□") & "補呈文件 " & IIf(Chk2(21).Value = 1, "■", "□") & "註冊費　 " & IIf(Chk2(22).Value = 1, "■", "□") & "其他：" & StrToStr(txt1(6).Text & String(20, " "), 10)
   iStr(16) = "　　　　乙方根據甲方所提供資料，依前項約定之範圍，代撰必要書件向本程序主管機關提"
   iStr(17) = "　　　　出，並代為收受有關文件。"
   iStr(18) = "第二條　委辦費用："
   'Modified by Lydia 2017/03/28 空白委任書要保留費用
   'If Trim(txt1(7)) = "" Then
   If Val(Trim(txt1(7))) = 0 And bolSpace = False Then
       iStr(19) = ""
       iStr(20) = ""
   Else
       'Added by Lydia 2017/03/28
       If Val(Trim(txt1(7))) = 0 Then
          strSpaceAmt = String(12, "　") & "元整"
       Else
          strSpaceAmt = ChangeNumber(txt1(7))
       End If
       'Modified by Lydia 2017/03/28 ChangeNumber(txt1(7)) => ChangeNumber(strSpaceAmt)
       iStr(19) = StrToStr("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40)
       iStr(20) = "　　　　" & Replace("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", StrToStr("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40), "")
       'Added by Lydia 2017/03/28 金額直接併入說明
       strExc(3) = StrToStr("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40)
       strExc(4) = "　　　　" & Replace("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", StrToStr("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40), "")
   End If
   'Modified by Lydia 2017/03/28 空白委任書要保留費用
   'If Trim(txt1(8)) = "" Then
   If Val(Trim(txt1(8))) = 0 And bolSpace = False Then
       iStr(21) = ""
       iStr(22) = ""
   Else
       'Added by Lydia 2017/03/28
       If Val(Trim(txt1(8))) = 0 Then
          strSpaceAmt = String(12, "　") & "元整"
       Else
          strSpaceAmt = ChangeNumber(txt1(8))
       End If
       'Modified by Lydia 2017/03/28 ChangeNumber(txt1(8)) => ChangeNumber(strSpaceAmt)
       iStr(21) = StrToStr("　　二、後酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本程序終結時，由甲方一次付清。", 40)
       iStr(22) = "　　　　" & Replace("　　二、後酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本程序終結時，由甲方一次付清。", StrToStr("　　二、後酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本程序終結時，由甲方一次付清。", 40), "")
       'Added by Lydia 2017/03/28 金額直接併入說明
       strExc(5) = StrToStr("　　二、後酬金新台幣　" & strSpaceAmt & "，於本程序終結時，由甲方一次付清。", 40)
       strExc(6) = "　　　　" & Replace("　　二、後酬金新台幣　" & strSpaceAmt & "，於本程序終結時，由甲方一次付清。", StrToStr("　　二、後酬金新台幣　" & strSpaceAmt & "，於本程序終結時，由甲方一次付清。", 40), "")
   End If
   iStr(23) = "第三條　乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，並不得發生足以"
   iStr(24) = "　　　　影響甲方權益之疏誤，否則應對甲方負損害賠償責任。但以不超過第二條所載前酬"
   iStr(25) = "　　　　金金額之三倍為限。"
   iStr(26) = "第四條　甲方確保所交付予乙方之資料均無虛偽情事，如因不實致生損害或法律責任時，概"
   iStr(27) = "　　　　由甲方負責，與乙方無關。"
   iStr(28) = "第五條　乙方於辦理過程中，應隨時將辦理經過如申請日期、案號及其他重要函件，儘速通"
   iStr(29) = "　　　　知或交付甲方。但甲方於簽約後變更連絡處所，未即時通知乙方，因而連絡不及致"
   iStr(30) = "　　　　延誤時限者，乙方不負責任。"
   iStr(31) = "第六條　凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時限者，乙方不負責。"
   iStr(32) = "　　　　經乙方通知甲方繳費而未依限繳納者，亦同。"
   iStr(33) = "第七條　甲方如逕自撤回所委辦程序，或未經乙方同意終止契約時，所約定之費用，仍應全"
   iStr(34) = "　　　　數給付。"
   iStr(35) = "第八條　本契約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，但有增刪修改時，需"
   iStr(36) = "　　　　甲乙雙方於更動處蓋章始生效力，並由雙方各執乙份為憑。"
   iStr(37) = "           "
   iStr(38) = "　　　　　　甲方：委任人：" & StrToStr(txt1(9) & String(54, " "), 27)
   iStr(39) = "　　　　　　　　　ID NO.：" & StrToStr(txt1(10) & String(20, " "), 10) & "代表人：" & StrToStr(txt1(11) & String(26, " "), 13)
   iStr(40) = "　　　　　　      地  址：" & StrToStr(txt1(12) & String(54, " "), 27)
   iStr(41) = "　　　　　　　　　電  話：" & StrToStr(txt1(13) & String(20, " "), 10) & "傳  真：" & StrToStr(txt1(14) & String(26, " "), 13)
   iStr(42) = "　　　　　　乙方：受任人：" & Combo2.Text 'Add By Sindy 2011/3/23 台一國際專利商標事務所                               "
   iStr(43) = "　　　　　　　　　經手人：" & StrToStr(txt1(15) & String(54, " "), 27)
   'Modified by Lydia 2020/04/09 改用模組控制
   'iStr(44) = "　　　　　　　　　地　址：台北市長安東路二段一一二號九樓"
   iStr(44) = "　　　　　　　　　地　址：" & PUB_SetAddrTofrm210114(Combo2.Text)
   iStr(45) = "　　　　　　　　　電  話：(02)25061023(總機)   FAX:(02)25011666"
   iStr(46) = "　　　　　　　　　網  址：www.taie.com.tw"
   iStr(47) = "　　　　　　　　　E-mail：ipdept@taie.com.tw"  'modify by sonia 2020/4/8 原為lawoffice
   iStr(48) = "  中    華    民    國 " & String((10 - LenB(StrConv((txt1(16)), vbFromUnicode))) / 2, " ") & txt1(16) & String((10 - LenB(StrConv((txt1(16)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(17)), vbFromUnicode))) / 2, " ") & txt1(17) & String((10 - LenB(StrConv((txt1(17)), vbFromUnicode))) / 2, " ") & "月" & String((10 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & txt1(18) & String((10 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & "日"
   
   'Added by Lydia 2017/03/28 有用印就記錄列印內容
   If iPrintC = 1 And bolAddSeal = True Then
           strDetail = ""
           For intI = 1 To UBound(iStr)
              If Trim(iStr(intI)) <> "" Then
                 If intI >= 19 And intI <= 22 Then
                    Select Case intI
                        Case 19: strDetail = strDetail & vbCrLf & RTrim(strExc(3))
                        Case 20: strDetail = strDetail & vbCrLf & RTrim(strExc(4))
                        Case 21: strDetail = strDetail & vbCrLf & RTrim(strExc(5))
                        Case 22: strDetail = strDetail & vbCrLf & RTrim(strExc(6))
                    End Select
                 Else
                    If intI <= 18 Or (intI >= 38 And intI <= 43) Or intI = 48 Then
                        If intI = 38 Then
                          strDetail = strDetail & vbCrLf & vbCrLf & IIf(opt1(0).Value = True, "會稿", "不會稿") & vbCrLf
                        End If
                       strDetail = strDetail & vbCrLf & RTrim(iStr(intI))
                    End If
                 End If
              End If
           Next
        'Modified by Lydia 2017/04/17 空白用印改由勾選項目控制
        'If PUB_AddRecSeal("3", txtPCnt.Text, IIf(ChkSeal.Value = 1, "", "Y"), strDetail, Combo2.Text) Then
        'Remove by Lydia 2017/05/16 用印記錄移到pdf建立
        'If PUB_AddRecSeal("3", txtPCnt.Text, IIf(bolSpace = True, "Y", ""), strDetail, Combo2.Text) Then
        'End If
   End If
   'end 2017/03/28
   
   iY = 0
   Printer.PaperSize = 9

   'add by nickc 2007/05/04
   'edit by nickc 2007/07/12 試著解決第二頁的格子線會不見的問題
   'If iCount = 1 Then
       Printer.Orientation = 1
   'End If
   Printer.FontName = "標楷體"
   Printer.FontSize = 20
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(iStr(1))) / 2
   iY = iY + Printer.TextHeight(iStr(1))
   Printer.CurrentY = iY
   iY = iY + ((Printer.TextHeight(iStr(1)) / 3) * 4)
   Printer.Print iStr(1)
   Printer.FontSize = 12
   'Added by Lydia 2017/03/28 同步用印
   If bolAddSeal = True Then
      '列印座置抓乙方資料的起始
      'X軸
      strExc(1) = 1000 + (Printer.TextWidth("　") * 30)
      'Y軸
      intI = 0
      If bolSpace = True Then
         If Val(Trim(txt1(8))) = 0 Then
            intI = 4
         Else
            intI = 3
         End If
      Else
         '顯示費用，資料列數不同
         If Val(Trim(txt1(7))) > 0 Then intI = intI + 2
         If Val(Trim(txt1(8))) > 0 Then intI = intI + 2
      End If
      strExc(2) = iY + ((Printer.TextHeight("　") / 3) * 4) * (36 + intI) - 50
      'Added by Lydia 2017/04/25 圖片尺寸
      strExc(3) = 1600 'width
      strExc(4) = 1600 'height
      
      'Added by Lydia 2020/03/25 已記錄公司名稱|用印編號
      intI = InStr(strCompSeal, Combo2.Text)
      If intI > 0 Then
         strExc(9) = Mid(strCompSeal, intI + Len(Combo2.Text))
         If InStr(strExc(9), ",") > 0 Then
             strExc(9) = Mid(strExc(9), 2, InStr(strExc(9), ",") - 2)
         Else
             strExc(9) = Mid(strExc(9), 2)
         End If
          If PUB_ReadDB2File(strSealFile, Val(strExc(9))) Then
             Set tObj = pvGetStdPicture(strSealFile)
             Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
          End If
      Else
      'end 2020/03/25
            If InStr(Combo2.Text, "專利法律") > 0 Then
              If PUB_ReadDB2File(strSealFile, 51) Then
                 Set tObj = pvGetStdPicture(strSealFile)
                 'Modified by Lydia 2017/04/25
                 'Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), 1570, 1570
                 Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
              End If
            End If
            If InStr(Combo2.Text, "專利商標") > 0 Then
              If PUB_ReadDB2File(strSealFile, 52) Then
                 Set tObj = pvGetStdPicture(strSealFile)
                 'Modified by Lydia 2017/04/25
                 'Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), 1570, 1570
                 Printer.PaintPicture tObj, Val(strExc(1)), Val(strExc(2)), Val(strExc(3)), Val(strExc(4))
              End If
            End If
      End If 'Added by Lydia 2020/03/25
   End If
   'end 2017/03/28
   For tmpI = 2 To UBound(iStr) - 1
       If iStr(tmpI) <> "" Then
           If tmpI = 39 Then
               tBoxTop = iY
           End If
           Printer.CurrentX = 1000
           Printer.CurrentY = iY
           Printer.Print iStr(tmpI)
           If tmpI = 37 Then
               iY = iY - (((Printer.TextHeight(iStr(1)) / 3) * 4) / 2)
           End If
           If tmpI = 19 Then
               Printer.FontBold = True
               Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 11) - 50
               Printer.CurrentY = iY
               'Modified by Lydia 2017/03/28 +判斷空白列印
               'Printer.Print ChangeNumber(txt1(7))
               If bolSpace = True And Val(Trim(txt1(7))) = 0 Then
                   Printer.Print String(12, "　") & "元整"
               Else
                   Printer.Print ChangeNumber(txt1(7))
               End If
               'end 2017/03/28
               Printer.FontBold = False
           End If
           If tmpI = 21 Then
               Printer.FontBold = True
               Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 11) - 50
               Printer.CurrentY = iY
               'Modified by Lydia 2017/03/28 +判斷空白列印
               'Printer.Print ChangeNumber(txt1(8))
               If bolSpace = True And Val(Trim(txt1(8))) = 0 Then
                   Printer.Print String(12, "　") & "元整"
               Else
                   Printer.Print ChangeNumber(txt1(8))
               End If
               'end 2017/03/28
               Printer.FontBold = False
           End If
           iY = iY + ((Printer.TextHeight(iStr(tmpI)) / 3) * 4)
           '畫線
           Select Case tmpI
           Case 4
                Printer.Line (1000 + (Printer.TextWidth("　") * 14), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 5
                Printer.Line (1000 + (Printer.TextWidth("　") * 4), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 7
                Printer.Line (1000 + (Printer.TextWidth("　") * 11.5), iY - 50)-(1000 + (Printer.TextWidth("　") * 39.5), iY - 50)
           Case 8, 9, 10
                Printer.Line (1000 + (Printer.TextWidth("　") * 11.5), iY - 50)-(1000 + (Printer.TextWidth("　") * 40.5), iY - 50)
           Case 15
                Printer.Line (1000 + (Printer.TextWidth("　") * 29), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 38, 40, 42, 43
                Printer.Line (1000 + (Printer.TextWidth("　") * 13), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case 39, 41
                Printer.Line (1000 + (Printer.TextWidth("　") * 13), iY - 50)-(1000 + (Printer.TextWidth("　") * 23), iY - 50)
                Printer.Line (1000 + (Printer.TextWidth("　") * 27), iY - 50)-(1000 + (Printer.TextWidth("　") * 40), iY - 50)
           Case Else
           End Select
       End If
   Next tmpI
   '畫格子
   Printer.Line (1000, tBoxTop)-(1000 + (Printer.TextWidth("　") * 5), tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 8)), , B
   Printer.Line (1000, tBoxTop)-(1000 + (Printer.TextWidth("　") * 2.5), tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 8)), , B
   Printer.Line (1000, tBoxTop)-(1000 + (Printer.TextWidth("　") * 5), tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 4)), , B
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 0.5)
   Printer.Print "會"
   If opt1(0).Value = True Then
       Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 3.25)
       Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 1.5)
       Printer.Print "Ｖ"
   End If
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 2.5)
   Printer.Print "稿"
   
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 4.5)
   Printer.Print "不"
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 5.5)
   Printer.Print "會"
   If opt1(1).Value = True Then
       Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 3.25)
       Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 5.5)
       Printer.Print "Ｖ"
   End If
   Printer.CurrentX = 1000 + (Printer.TextWidth("　") * 0.75)
   Printer.CurrentY = tBoxTop + (((Printer.TextHeight("　") / 3) * 4) * 6.5)
   Printer.Print "稿"
   Printer.FontSize = 16
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(iStr(UBound(iStr)))) / 2
   Printer.CurrentY = iY + 50
   Printer.Print iStr(UBound(iStr))
   'add by nickc 2007/05/04
   'edit by nickc 2007/07/12 試著解決第二頁的格子線會不見的問題
   'If iCount = Val(txtPCnt) Then
       Printer.EndDoc
   'Else
   '    Printer.NewPage
   'End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

'Modified by Lydia 2022/01/22 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
'Add By Sindy 98/02/11
Dim intLen As Integer
   
   If KeyAscii <> 8 Then
      intLen = GetTextLength(txt1(Index))
      intLen = intLen + GetTextLength(Chr(KeyAscii))
      '2014/5/13 modify by sonia
      'If intLen > txt1(Index).MaxLength Then KeyAscii = 0
      If CheckLengthIsOK(txt1(Index).Text & Chr(KeyAscii), txt1(Index).MaxLength) = False Then
         KeyAscii = 0
      End If
      'end 2014/5/13
   End If
   '98/02/11 End
   If Index = 7 Or Index = 8 Then
       If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 46 Then
           KeyAscii = 0
       End If
   End If
   '2009/11/13 ADD BY SONIA
   If Index = 16 Or Index = 17 Or Index = 18 Then
      If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
         KeyAscii = 0
      End If
   End If
   '2009/11/13 END
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   'Modified by Lydia 2018/04/13
   'txt1(Index).Text = Replace(Replace(txt1(Index).Text, Chr(10), ""), Chr(13), "")
   txt1(Index).Text = PUB_StringFilter(txt1(Index).Text)
   Cancel = False
   'modify by sonia 2020/7/23 商品/服務名稱的第一,二行提醒訊息改掉
   'If CheckLengthIsOK(txt1(Index).Text, txt1(Index).MaxLength) = False Then
   '    txt1(Index).SetFocus
   '    txt1_GotFocus Index
   '    Cancel = True
   'End If
   Select Case Index
      Case 3, 4   '輸入之資料過長，超過 個字（註：中文算兩個字)
         If CheckLengthIsOK(txt1(Index).Text, txt1(Index).MaxLength, False) = False Then
             ShowMsg "輸入之資料過長, 超過" & Format(txt1(Index).MaxLength) & "個字（註：中文算兩個字)，超過的文字請移至次行輸入！"
             txt1(Index).SetFocus
             txt1_GotFocus Index
             Cancel = True
         End If
      Case Else
         If CheckLengthIsOK(txt1(Index).Text, txt1(Index).MaxLength) = False Then
             txt1(Index).SetFocus
             txt1_GotFocus Index
             Cancel = True
         End If
   End Select
   'end 2020/7/23
End Sub

Private Sub txtPCnt_GotFocus()
   txtPCnt.SelStart = 0
   txtPCnt.SelLength = Len(txtPCnt)
End Sub

Private Sub txtPCnt_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 46 Then
       KeyAscii = 0
   End If
End Sub

'Add By Sindy 2010/4/29
Private Function SetCustTxt(strCUCode As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   SetCustTxt = False
   strCUCode = Left(strCUCode & "000000000", 9)
   'Modified by Morgan 2021/5/5
   'StrSQLa = "Select * From Customer,nation,potcustcont Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "' and CU10=na01(+) and pcc01(+)=substr(CU08, 1, 8) And pcc02(+)=substr(CU08, 9, 1) "
   StrSQLa = "Select * From Customer Where CU01='" & Mid(strCUCode, 1, 8) & "' And CU02='" & Mid(strCUCode, 9, 1) & "'"
   'end 2021/5/5
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      SetCustTxt = True
      '申請人中文
      Me.txt1(9).Text = "" & rsA("CU04").Value
      'ID No.
      Me.txt1(10).Text = "" & rsA("CU11").Value
      '申請地址
      Me.txt1(12).Text = "" & rsA("CU23").Value
'      '國籍
'      Me.txt1(8).Text = "" & rsA("NA03").Value
'      '聯絡人地址
'      If "" & rsA("CU08").Value <> "" Then
'         Me.txt1(9).Text = "" & rsA("pcc22").Value
'      Else
'         Me.txt1(9).Text = "" & rsA("CU31").Value
'      End If
      '電話1
      Me.txt1(13).Text = "" & rsA("CU16").Value
      '傳真1
      Me.txt1(14).Text = "" & rsA("CU18").Value
      '代表人1中文
      Me.txt1(11).Text = "" & rsA("CU07").Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Add by Amy 2016/08/19
Private Sub UpdReceiptCmp(ByVal stNowCustNo As String, ByVal stNowCmp As String)
    Dim strUpd As String
    
    Exit Sub 'Added by Lydia 2022/08/30 受任人下拉預設只剩下台一國際智慧財產事務所，所以不必再更新客戶檔的了。
    
    'Add by Amy 2016/12/30 +同業務區或為MCTF同組人員才可回寫收據公司別
    If ChkSameCuArea(stNowCustNo, strUserNum) = False Then Exit Sub
    
    'Modified by Lydia 2019/04/12 拿掉UpdateID,Date,Time(CU84,CU85,CU86)
    'strUpd = "Update Customer Set CU84='" & strUserNum & "',CU85=to_number(to_char(sysdate,'YYYYMMDD')),CU86=to_number(to_char(sysdate,'HH24MI')),CU162='" & stNowCmp & "' " & _
                    "Where CU01='" & Left(stNowCustNo, 8) & "' And CU02='" & Mid(stNowCustNo, 9, 1) & "' "
    strUpd = "Update Customer Set CU162='" & stNowCmp & "' " & _
                    "Where CU01='" & Left(stNowCustNo, 8) & "' And CU02='" & Mid(stNowCustNo, 9, 1) & "' "
    Pub_SeekTbLog strUpd
    'Modified by Lydia 2019/04/23 觸發Trigger
    'cnnConnection.Execute strUpd
    cnnConnection.Execute "begin user_data.user_enabled:=1; " & strUpd & " ; end; "
End Sub

'Added by Lydia 2017/04/13 列印:先轉PDF,列印後刪檔
Private Sub Print2PDF(ByVal bSpace As Boolean)
Dim strFileName As String
Dim strOldName As String 'Added by Lydia 2017/06/07

'Added by Lydia 2017/04/25 VB印表機實際列印的左邊界、右邊界
Set Printer = Printers(PUB_PrinterIndex(Combo1.Text))
d_Top = Format((Printer.Height - Printer.ScaleHeight) / 2, "0") '直印
d_Left = Format((Printer.Width - Printer.ScaleWidth) / 2, "0")
'end 2017/04/25

strDetail = "" 'Added by Lydia 2017/05/16
strOldName = App.Title 'Added by Lydia 2017/06/07

Screen.MousePointer = vbHourglass
    'Modified by Lydia 2022/01/22 先產生Word檔，後轉成PDF檔逐一列印
'    For iCount = 1 To Val(txtPCnt)
'        iPrintC = iCount
'        'Modified by Lydia 2017/06/06 改用App.Title變更印表機列印文件名稱(執行exe檔有效,VB跑無效)
'        'strFileName = strUserNum & "_T_" & IIf(bSpace = False, IIf(Trim(txt1(9)) <> "", Mid(Trim(txt1(9)), 1, 4), Mid(Trim(txt1(11)), 1, 4)), "空白") & iCount & ".pdf"
'        'If Dir(App.path & "\" & strFileName) <> "" Then
'        '   Kill App.path & "\" & strFileName
'        'End If
'        ''轉PDF
'        'frmPDF.Show
'        'frmPDF.StartProcess App.path, strFileName
'        'Call StrMenu(bSpace)
'        'frmPDF.EndtProcess
'        'Unload frmPDF
'        strFileName = strUserNum & "_T_" & IIf(bSpace = False, IIf(Trim(txt1(5)) <> "", Mid(Trim(txt1(5)), 1, 4), Mid(Trim(txt1(6)), 1, 4)), "空白") & iCount
'        App.Title = strFileName
'        Call StrMenu(bSpace)
'        'end 2017/06/07
'
'        'Added by Lydia 2017/05/16 用印記錄移到pdf建立
'        If iCount = 1 And strDetail <> "" Then
'           'If Dir(App.path & "\" & strFileName) <> "" Then 'Remove by Lydia 2020/03/16 因為不存檔案所以取消檔案檢查(自2017/06/08~2020/03/16無用印記錄)
'              If PUB_AddRecSeal("3", txtPCnt.Text, IIf(bSpace = True, "Y", ""), strDetail, Combo2.Text) Then
'              End If
'           'End If 'Remove by Lydia 2020/03/16
'        End If
'        'end 2017/05/16
'
'        'Remove by Lydia 2017/06/07
'        ''列印PDF
'        'PUB_PrintPDF App.path & "\" & strFileName, Me.Combo1
'        ''刪除PDF
'        'Kill App.path & "\" & strFileName
'    Next iCount
    Call runWordProc(bSpace)
    If m_TempPDF <> "" Then
        For iCount = 1 To Val(txtPCnt)
            iPrintC = iCount
            strFileName = strUserNum & "_T_" & m_TempFN & iCount
            PUB_PrintPDF App.path & "\" & strUserNum & "\" & m_TempPDF, Combo1.Text
            App.Title = strFileName
        Next iCount
    End If
'--------------先產生Word檔，後轉成PDF檔逐一列印
    App.Title = strOldName 'Added by Lydia 2017/06/07
    
End Sub

'Added by Lydia 2022/01/22 下載Word範本套印
Private Sub runWordProc(ByVal pSpace As Boolean)
Dim iStr(1 To 48) As String    '用印記錄(全文)
Dim strSealFile As String '公司章圖檔
Dim strSpaceAmt As String
Dim strName As String
Dim strText
Dim intA As Integer
Dim m_FileName As String, m_TempFileName As String
Dim m_DefPath As String
Dim oShape
Dim oWord

On Error GoTo ErrHand
   
   '上傳檔案
   'Modified by Lydia 2024/07/22 改用變數
   'intI = SaveImgByteFile("\\" & pub_getspecman("FTP_VOL_IP_LINUX") & "\PolyCOM\TaieNew\RptSample\M51-000300-0-03 智權部委任契約書_T.docx", "M51", "000300", "0", "03", "4", "1")
    
   m_DefPath = App.path & "\" & strUserNum
   'Added by Lydia 2022/01/25
   m_TempPDF = ""
   '變更Word印表機
   PUB_SetOsDefaultPrinter Combo1
   PUB_SetWordActivePrinter
   'end 2022/01/25
   
   '下載範本檔: M51-000300-0-03 智權部委任契約書_T.docx
   m_TempFN = Pub_RepFileName(IIf(pSpace = False, IIf(Trim(txt1(0)) <> "", Mid(Trim(txt1(0)), 1, 4), Mid(Trim(txt1(1)), 1, 4)), "空白")) 'Move by Lydia 2022/01/25 從m_TempFileName移過來
   'Modified by Lydia 2022/01/25 改成Word直接印，所以範本一開始就先命名好
   'm_FileName = "$$" & Me.Name & ".docx"
   m_FileName = "$$" & strUserNum & "_T_" & m_TempFN & ".docx"
   If Dir(m_DefPath & "\" & m_FileName) <> "" Then
      Kill m_DefPath & "\" & m_FileName
   End If
   If PUB_GetSampleFile(m_FileName, "M51-000300-0-03", , m_DefPath) = False Then
        Exit Sub
   End If
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   'Remove by Lydia 2022/01/25 不用改存PDF檔
   'm_TempFileName = "$$" & strUserNum & "_T_" & m_TempFN & ".pdf"
   'If Dir(m_DefPath & "\" & m_TempFileName) <> "" Then
   '   Kill m_DefPath & "\" & m_TempFileName
   'End If
   'end 2022/01/25
   '改成直接用範本檔
   'Q: AddToRecentFiles:=False還是會新增到最近開啟記錄
   g_WordAp.Documents.Open m_DefPath & "\" & m_FileName, False, False, False
  
   With g_WordAp
      .Selection.WholeStory
      .Selection.Copy
      For intA = 0 To 24
         strName = "PS" & Format(intA, "000")
         strText = ""
'-------第一條
         If intA = 0 Then
              '委辦內容或案件名稱
              strText = PUB_StrToStr(txt1(0), 56)
         ElseIf intA = 1 Then
              '委辦內容或案件名稱
              strText = PUB_StrToStr(txt1(1), 76)
         ElseIf intA = 2 Then
              '申請種類：
              strText = IIf(Chk1(0).Value = 1, "■", "□") & "商標  " & IIf(Chk1(1).Value = 1, "■", "□") & "團體標章  " & IIf(Chk1(2).Value = 1, "■", "□") & "證明標章"
         ElseIf intA = 3 Then
              '商品/服務類別
              strText = PUB_StrToStr(txt1(2) & " ", 56, True)
         ElseIf intA = 4 Then
               '商品/服務名稱：1
              strText = PUB_StrToStr(txt1(3), 62)
         ElseIf intA = 5 Then
               '商品/服務名稱：2
              strText = PUB_StrToStr(txt1(4), 62)
         ElseIf intA = 6 Then
               '商品/服務名稱：3
              strText = PUB_StrToStr(txt1(5), 62)
         ElseIf intA = 7 Then
               '案件性質：1
               'Modified by Lydia 2022/10/04 第一項多加一個全形空白
               strText = IIf(Chk2(0).Value = 1, "■", "□") & "查名　　　" & IIf(Chk2(1).Value = 1, "■", "□") & "申請　　" & IIf(Chk2(2).Value = 1, "■", "□") & "延展　　 " & IIf(Chk2(3).Value = 1, "■", "□") & "變更　　 " & IIf(Chk2(4).Value = 1, "■", "□") & "分割　　 " & IIf(Chk2(5).Value = 1, "■", "□") & "移轉　　" & IIf(Chk2(6).Value = 1, "■", "□") & "授權　　"
         ElseIf intA = 8 Then
               '案件性質：2
              'Modified by Lydia 2022/10/04 第一項多加一個全形空白
              strText = IIf(Chk2(7).Value = 1, "■", "□") & "補證　　　" & IIf(Chk2(8).Value = 1, "■", "□") & "再授權　" & IIf(Chk2(9).Value = 1, "■", "□") & "質權登記 " & IIf(Chk2(10).Value = 1, "■", "□") & "英文證明 " & IIf(Chk2(11).Value = 1, "■", "□") & "申請理由 " & IIf(Chk2(12).Value = 1, "■", "□") & "異議　　" & IIf(Chk2(13).Value = 1, "■", "□") & "評定　　"
         ElseIf intA = 9 Then
               '案件性質：3
               'Modified by Lydia 2022/08/30 不顯示訴願、第一審行政訴訟; 第一項多加一個全形空白
               'Memo by Lydia 2022/09/23 (還原)經協商後，專利、商標案件之訴願程序將由智慧所承辦
               'Modified by Lydia 2022/10/04 (debug) 只還原訴願，不還原第一審行政訴訟
               'strText = IIf(Chk2(14).Value = 1, "■", "□") & "廢止　　" & IIf(Chk2(15).Value = 1, "■", "□") & "訴願　　" & IIf(Chk2(16).Value = 1, "■", "□") & "答辯　　 " & IIf(Chk2(17).Value = 1, "■", "□") & "補充理由、補充答辯  " & IIf(Chk2(18).Value = 1, "■", "□") & "第一審行政訴訟　"
               strText = IIf(Chk2(14).Value = 1, "■", "□") & "廢止　　　" & IIf(Chk2(15).Value = 1, "■", "□") & "訴願　　" & IIf(Chk2(16).Value = 1, "■", "□") & "答辯　　 " & IIf(Chk2(17).Value = 1, "■", "□") & "補充理由、補充答辯  "
         ElseIf intA = 10 Then
               '案件性質：4
              'Modified by Lydia 2022/08/30 不顯示上訴審行政訴訟
              'Memo by Lydia 2022/09/23 (還原)經協商後，專利、商標案件之訴願程序將由智慧所承辦
               'Modified by Lydia 2022/10/04 (debug) 只還原訴願，不還原第一審行政訴訟
              'strText = IIf(Chk2(19).Value = 1, "■", "□") & "上訴審行政訴訟　　" & IIf(Chk2(20).Value = 1, "■", "□") & "補呈文件 " & IIf(Chk2(21).Value = 1, "■", "□") & "註冊費　 " & IIf(Chk2(22).Value = 1, "■", "□") & "其他：" & PUB_StrToStr(txt1(6), 20)
              strText = IIf(Chk2(20).Value = 1, "■", "□") & "補呈文件  " & IIf(Chk2(21).Value = 1, "■", "□") & "註冊費　" & IIf(Chk2(22).Value = 1, "■", "□") & "其他：" & PUB_StrToStr(txt1(6), 20)

'-------第二條 委辦費用
         ElseIf intA = 11 Then
                '酬金要區分項目描述
                strExc(1) = "": strExc(2) = ""
                If Val(Trim(txt1(7))) = 0 Then
                     If pSpace = True Then strExc(1) = "　　　　　　　　　　　　元整"
                Else
                     strExc(1) = ChangeNumber(txt1(7))
                End If
               strExc(3) = PUB_StrToStr("　　一、前酬金新台幣　|#PS012#|，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 82 - Len(strExc(1)) * 2 + 9)
               strExc(4) = "　　　　" & Replace("　　一、前酬金新台幣　|#PS012#|，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", PUB_StrToStr("　　一、前酬金新台幣　|#PS012#|，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 82 - Len(strExc(1)) * 2 + 9), "")
       
              If pSpace = True Or (Val(Trim(txt1(7))) > 0 And Val(Trim(txt1(8))) > 0) Then
                   strText = strExc(3) & vbCrLf & strExc(4) & vbCrLf & _
                                "　　二、後酬金新台幣　|#PS013#|，於本程序終結時，由甲方一次付清。"
              Else
                   If Val(Trim(txt1(7))) > 0 Then
                        strText = Replace(strExc(3) & vbCrLf & strExc(4), "一、", "　　") & "|#PS013#|"
                   ElseIf Val(Trim(txt1(8))) > 0 Then
                        strText = "|#PS012#|　　　　後酬金新台幣　|#PS013#|，於本程序終結時，由甲方一次付清。"
                   Else
                        strText = "　　|#PS012#||#PS013#|"
                   End If
              End If
         ElseIf intA = 12 Then
              '前酬金
              If Val(Trim(txt1(7))) = 0 Then
                   If pSpace = True Then strText = "　　　　　　　　　　　　元整"
              Else
                   strText = ChangeNumber(txt1(7))
              End If
         ElseIf intA = 13 Then
              '後酬金
              If Val(Trim(txt1(8))) = 0 Then
                   If pSpace = True Then strText = "　　　　　　　　　　　　元整"
              Else
                   strText = ChangeNumber(txt1(8))
              End If
'-------
         ElseIf intA = 14 Then
              '會稿
              strText = IIf(opt1(0).Value = True, "V", "")
         ElseIf intA = 15 Then
              '不會稿
              strText = IIf(opt1(1).Value = True, "V", "")
         ElseIf intA = 16 Then
              '委任人
              strText = PUB_StrToStr(txt1(9).Text, 54)
         ElseIf intA = 17 Then
              '委任人-ID NO. + 代表人
              strText = PUB_StrToStr(txt1(10).Text & " ", 20, True) & "　" & "代表人：" & PUB_StrToStr(txt1(11).Text, 26)
         ElseIf intA = 18 Then
              '委任人-地  址
              strText = PUB_StrToStr(txt1(12).Text, 54)
         ElseIf intA = 19 Then
              '委任人-電  話+傳　真
              strText = PUB_StrToStr(txt1(13).Text & " ", 20, True) & "　" & "傳  真：" & PUB_StrToStr(txt1(14).Text, 26)
         ElseIf intA = 20 Then
              '受任人
              strText = Combo2.Text
         ElseIf intA = 21 Then
              '經手人
              strText = PUB_StrToStr(txt1(15).Text, 54)
         ElseIf intA = 22 Then
              '受任人-地址
              strText = PUB_SetAddrTofrm210114(Combo2.Text)
         ElseIf intA = 23 Then
              strText = "        中    華    民    國 " & String((8 - LenB(StrConv((txt1(16)), vbFromUnicode))) / 2, " ") & txt1(16) & String((8 - LenB(StrConv((txt1(16)), vbFromUnicode))) / 2, " ") & "年" & String((8 - LenB(StrConv((txt1(17)), vbFromUnicode))) / 2, " ") & txt1(17) & String((8 - LenB(StrConv((txt1(17)), vbFromUnicode))) / 2, " ") & "月" & String((8 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & txt1(18) & String((8 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & "日"
         ElseIf intA = 24 Then
              strText = ""
         Else
         End If
         
         If Trim(strName) <> "" Then
            .Selection.Find.ClearFormatting
            .Selection.Find.Text = "|#" & strName & "#|"
            .Selection.Find.Replacement.Text = ""
            .Selection.Find.Forward = True
            .Selection.Find.Wrap = wdFindContinue
            .Selection.Find.Format = False
            .Selection.Find.MatchCase = False
            .Selection.Find.MatchWholeWord = False
            .Selection.Find.MatchWildcards = False
            .Selection.Find.MatchSoundsLike = False
            .Selection.Find.MatchAllWordForms = False
            .Selection.Find.MatchByte = True
            .Selection.Find.Execute
            .Selection.Delete
            '保留;因為先全部以細明體-ExtB,最後全選改字型;
            If (intA >= 0 And intA <= 6) Or intA = 9 Or intA = 11 Or intA = 12 Or intA = 15 Then
               '有Unicode字需要換字型
               .Selection.Font.Name = "細明體-ExtB"
            End If
            If intA = 12 Or intA = 13 Then
                '金額要粗體
                .Selection.Font.Bold = True
            End If
            If intA = 14 Or intA = 15 Then
                '會稿/不會稿勾選
                .Selection.Font.Size = 12
            End If
            If intA = 24 And bolAddSeal = True Then  '公司章: 放在受任人的儲存格
                strExc(9) = Mid(strCompSeal, InStr(strCompSeal, Combo2))
                If InStr(strExc(9), ",") > 0 Then
                    strExc(9) = Right(Mid(strExc(9), 1, InStr(strExc(9), ",") - 1), 2)
                Else
                    strExc(9) = Right(strExc(9), 2)
                End If
                If PUB_ReadDB2File(m_DefPath & "\$$" & Me.Name & "TempFile", Val(strExc(9))) Then
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=m_DefPath & "\$$" & Me.Name & "TempFile", LinkToFile:=False, SaveWithDocument:=True)
                    '--------設定圖片=文蓋圖(文字在前)
                        oShape.Fill.Visible = msoFalse
                        oShape.Fill.Solid
                        oShape.Fill.Transparency = 0#
                        oShape.Line.Weight = 0.75
                        oShape.Line.DashStyle = msoLineSolid
                        oShape.Line.Style = msoLineSingle
                        oShape.Line.Transparency = 0#
                        oShape.Line.Visible = msoFalse
                        oShape.LockAspectRatio = msoTrue
                        oShape.Rotation = 0#
                        oShape.PictureFormat.Brightness = 0.5
                        oShape.PictureFormat.Contrast = 0.5
                        oShape.PictureFormat.ColorType = msoPictureAutomatic
                        oShape.PictureFormat.CropLeft = 0#
                        oShape.PictureFormat.CropRight = 0#
                        oShape.PictureFormat.CropTop = 0#
                        oShape.PictureFormat.CropBottom = 0#

                        oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                        oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                        oShape.Left = .CentimetersToPoints(8.25)
                        oShape.Top = .CentimetersToPoints(0.3)
                        oShape.LockAnchor = False
                        oShape.LayoutInCell = True
                        oShape.WrapFormat.AllowOverlap = True
                        oShape.WrapFormat.Side = wdWrapBoth
                        oShape.WrapFormat.DistanceTop = .CentimetersToPoints(0)
                        oShape.WrapFormat.DistanceBottom = .CentimetersToPoints(0)
                        oShape.WrapFormat.DistanceLeft = .CentimetersToPoints(0.32)
                        oShape.WrapFormat.DistanceRight = .CentimetersToPoints(0.32)
                        oShape.WrapFormat.Type = 3
                        oShape.ZOrder 5 '文蓋圖(文字在前)
                        '---------------------------
                End If
          
            End If
            .Selection.Font.ColorIndex = wdBlack
            .Selection.TypeText strText
            '保留;因為先全部以細明體-ExtB,最後全選改字型;
            If (intA >= 0 And intA <= 6) Or intA = 9 Or intA = 11 Or intA = 12 Or intA = 15 Then
               '有Unicode字需要換字型=>還原
               .Selection.Font.Name = "標楷體"
            End If

            If intA = 12 Or intA = 13 Then
                '金額要粗體=>還原
                .Selection.Font.Bold = False
            End If
            If intA = 14 Or intA = 15 Then
                '會稿/不會稿勾選=>還原
                .Selection.Font.Size = 12
            End If
         End If
         
      Next intA
'      '因為先全部以細明體-ExtB,最後全選改字型;
      .Selection.WholeStory
      .Selection.Font.Name = "標楷體"
   End With
    
   '改存成PDF檔
   'Memo by Lydia 2022/01/25  因為受PDF redirect設定灰階列印影響，改成Word直接印
   intA = IIf(Val(txtPCnt) = 0, 1, Val(txtPCnt))
   For intI = 1 To intA
       g_WordAp.PrintOut Background:=False, Range:=4, Item:=0, Copies:=1, Pages:="1", Collate:=True
   Next intI
   
   '保留: 存檔
   'g_WordAp.ActiveDocument.Close wdSaveChanges
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing '避免快速開啟Word,程式出錯
   m_TempPDF = m_FileName 'Added by Lydia 2022/01/25
   
   'Mark by Lydia 2022/01/25 因為受PDF redirect設定灰階列印影響，改成Word直接印
   'If PUB_PrintWord2PDF(g_WordAp, m_DefPath, m_TempFileName, m_TempPDF) = False Then
   '    Exit Sub
   'End If
   'end 2022/01/19
   
If bolAddSeal = True Then  '用印記錄
   iStr(1) = "商標案件委任契約書"
   iStr(2) = "委任人(甲方)茲委任受任人(乙方)辦理國內商標案件，雙方同意條件如下："
   iStr(3) = "第一條　委辦範圍："
   iStr(4) = "    一、委辦內容或案件名稱：" & StrToStr(txt1(0) & String(52, " "), 26)
   iStr(5) = "　　　　" & StrToStr(txt1(1) & String(72, " "), 36)
   iStr(6) = "　　二、申請種類：" & IIf(Chk1(0).Value = 1, "■", "□") & "商標  " & IIf(Chk1(1).Value = 1, "■", "□") & "團體標章  " & IIf(Chk1(2).Value = 1, "■", "□") & "證明標章"
   iStr(7) = "　　　　商品/服務類別：" & StrToStr(txt1(2) & String(56, " "), 28) & "類"
   iStr(8) = "　　　　商品/服務名稱：" & StrToStr(txt1(3) & String(58, " "), 29)
   iStr(9) = "　　　　　　　　　　　 " & StrToStr(txt1(4) & String(58, " "), 29)
   iStr(10) = "　　　　　　　　　　　 " & StrToStr(txt1(5) & String(58, " "), 29)
   iStr(11) = "　　三、案件性質："
   iStr(12) = "　　　　" & IIf(Chk2(0).Value = 1, "■", "□") & "查名　　" & IIf(Chk2(1).Value = 1, "■", "□") & "申請　　" & IIf(Chk2(2).Value = 1, "■", "□") & "延展　　 " & IIf(Chk2(3).Value = 1, "■", "□") & "變更　　 " & IIf(Chk2(4).Value = 1, "■", "□") & "分割　　 " & IIf(Chk2(5).Value = 1, "■", "□") & "移轉　　" & IIf(Chk2(6).Value = 1, "■", "□") & "授權　　"
   iStr(13) = "　　　　" & IIf(Chk2(7).Value = 1, "■", "□") & "補證　　" & IIf(Chk2(8).Value = 1, "■", "□") & "再授權　" & IIf(Chk2(9).Value = 1, "■", "□") & "質權登記 " & IIf(Chk2(10).Value = 1, "■", "□") & "英文證明 " & IIf(Chk2(11).Value = 1, "■", "□") & "申請理由 " & IIf(Chk2(12).Value = 1, "■", "□") & "異議　　" & IIf(Chk2(13).Value = 1, "■", "□") & "評定　　"
   'Modified by Lydia 2022/08/30 不顯示訴願、第一審行政訴訟、上訴審行政訴訟
   'Memo by Lydia 2022/09/23 (還原)經協商後，專利、商標案件之訴願程序將由智慧所承辦
   'Modified by Lydia 2022/10/04 (debug) 只要還原訴願(2022/09/23 )
   'iStr(14) = "　　　　" & IIf(chk2(14).Value = 1, "■", "□") & "廢止　　" & IIf(chk2(15).Value = 1, "■", "□") & "訴願　　" & IIf(chk2(16).Value = 1, "■", "□") & "答辯　　 " & IIf(chk2(17).Value = 1, "■", "□") & "補充理由、補充答辯  " & IIf(chk2(18).Value = 1, "■", "□") & "第一審行政訴訟　"
   'iStr(15) = "　　　　" & IIf(chk2(19).Value = 1, "■", "□") & "上訴審行政訴訟　　" & IIf(chk2(20).Value = 1, "■", "□") & "補呈文件 " & IIf(chk2(21).Value = 1, "■", "□") & "註冊費　 " & IIf(chk2(22).Value = 1, "■", "□") & "其他：" & StrToStr(txt1(6).Text & String(20, " "), 10)
   iStr(14) = "　　　　" & IIf(Chk2(14).Value = 1, "■", "□") & "廢止　　 " & IIf(Chk2(15).Value = 1, "■", "□") & "訴願　　" & IIf(Chk2(16).Value = 1, "■", "□") & "答辯　　 " & IIf(Chk2(17).Value = 1, "■", "□") & "補充理由、補充答辯  "
   iStr(15) = "　　　　" & IIf(Chk2(20).Value = 1, "■", "□") & "補呈文件 " & IIf(Chk2(21).Value = 1, "■", "□") & "註冊費　 " & IIf(Chk2(22).Value = 1, "■", "□") & "其他：" & StrToStr(txt1(6).Text & String(20, " "), 10)
   'end 2022/10/04
   iStr(16) = "　　　　乙方根據甲方所提供資料，依前項約定之範圍，代撰必要書件向本程序主管機關提"
   iStr(17) = "　　　　出，並代為收受有關文件。"
   iStr(18) = "第二條　委辦費用："
   If Val(Trim(txt1(7))) = 0 And pSpace = False Then
       iStr(19) = ""
       iStr(20) = ""
   Else
       If Val(Trim(txt1(7))) = 0 Then
          strSpaceAmt = String(12, "　") & "元整"
       Else
          strSpaceAmt = ChangeNumber(txt1(7))
       End If
       iStr(19) = StrToStr("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40)
       iStr(20) = "　　　　" & Replace("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", StrToStr("　　一、前酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40), "")
       strExc(3) = StrToStr("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40)
       strExc(4) = "　　　　" & Replace("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", StrToStr("　　一、前酬金新台幣　" & strSpaceAmt & "，於本契約簽訂同時由甲方一次付清，並由乙方製據為憑，乙方於收到本費用後始有履行本契約之義務。", 40), "")
   End If
   '空白委任書要保留費用
   If Val(Trim(txt1(8))) = 0 And pSpace = False Then
       iStr(21) = ""
       iStr(22) = ""
       iStr(19) = Replace(iStr(21), "一、", "　　")
       iStr(20) = Replace(iStr(22), "一、", "　　")
       strExc(3) = Replace(strExc(3), "一、", "　　")
       strExc(4) = Replace(strExc(4), "一、", "　　")
   Else
       If Val(Trim(txt1(8))) = 0 Then
          strSpaceAmt = String(12, "　") & "元整"
       Else
          strSpaceAmt = ChangeNumber(txt1(8))
       End If
       iStr(21) = StrToStr("　　二、後酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本程序終結時，由甲方一次付清。", 40)
       iStr(22) = "　　　　" & Replace("　　二、後酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本程序終結時，由甲方一次付清。", StrToStr("　　二、後酬金新台幣　" & String(LenB(StrConv(strSpaceAmt, vbFromUnicode)), " ") & "，於本程序終結時，由甲方一次付清。", 40), "")
       strExc(5) = StrToStr("　　二、後酬金新台幣　" & strSpaceAmt & "，於本程序終結時，由甲方一次付清。", 40)
       strExc(6) = "　　　　" & Replace("　　二、後酬金新台幣　" & strSpaceAmt & "，於本程序終結時，由甲方一次付清。", StrToStr("　　二、後酬金新台幣　" & strSpaceAmt & "，於本程序終結時，由甲方一次付清。", 40), "")
       If iStr(19) = "" Then
         iStr(21) = Replace(iStr(21), "二、", "　　")
         iStr(22) = Replace(iStr(22), "二、、", "　　")
         strExc(5) = Replace(strExc(5), "二、", "　　")
         strExc(6) = Replace(strExc(6), "二、", "　　")
       End If
   End If
   iStr(23) = "第三條　乙方對於甲方所委辦之案件內容，於辦理中應嚴守秘密不得外洩，並不得發生足以"
   iStr(24) = "　　　　影響甲方權益之疏誤，否則應對甲方負損害賠償責任。但以不超過第二條所載前酬"
   iStr(25) = "　　　　金金額之三倍為限。"
   iStr(26) = "第四條　甲方確保所交付予乙方之資料均無虛偽情事，如因不實致生損害或法律責任時，概"
   iStr(27) = "　　　　由甲方負責，與乙方無關。"
   iStr(28) = "第五條　乙方於辦理過程中，應隨時將辦理經過如申請日期、案號及其他重要函件，儘速通"
   iStr(29) = "　　　　知或交付甲方。但甲方於簽約後變更連絡處所，未即時通知乙方，因而連絡不及致"
   iStr(30) = "　　　　延誤時限者，乙方不負責任。"
   iStr(31) = "第六條　凡經乙方正式通知甲方之任何事項，如甲方未依限答覆致延誤時限者，乙方不負責。"
   iStr(32) = "　　　　經乙方通知甲方繳費而未依限繳納者，亦同。"
   iStr(33) = "第七條　甲方如逕自撤回所委辦程序，或未經乙方同意終止契約時，所約定之費用，仍應全"
   iStr(34) = "　　　　數給付。"
   iStr(35) = "第八條　本契約一式二份，經甲方暨乙方之經手人簽字或蓋章後生效，但有增刪修改時，需"
   iStr(36) = "　　　　甲乙雙方於更動處蓋章始生效力，並由雙方各執乙份為憑。"
   iStr(37) = "           "
   iStr(38) = "　　　　　　甲方：委任人：" & StrToStr(txt1(9) & String(54, " "), 27)
   iStr(39) = "　　　　　　　　　ID NO.：" & StrToStr(txt1(10) & String(20, " "), 10) & "代表人：" & StrToStr(txt1(11) & String(26, " "), 13)
   iStr(40) = "　　　　　　      地  址：" & StrToStr(txt1(12) & String(54, " "), 27)
   iStr(41) = "　　　　　　　　　電  話：" & StrToStr(txt1(13) & String(20, " "), 10) & "傳  真：" & StrToStr(txt1(14) & String(26, " "), 13)
   iStr(42) = "　　　　　　乙方：受任人：" & Combo2.Text 'Add By Sindy 2011/3/23 台一國際專利商標事務所                               "
   iStr(43) = "　　　　　　　　　經手人：" & StrToStr(txt1(15) & String(54, " "), 27)
   iStr(44) = "　　　　　　　　　地　址：" & PUB_SetAddrTofrm210114(Combo2.Text)
   iStr(45) = "　　　　　　　　　電  話：(02)25061023(總機)   FAX:(02)25011666"
   iStr(46) = "　　　　　　　　　網  址：www.taie.com.tw"
   iStr(47) = "　　　　　　　　　E-mail：ipdept@taie.com.tw"
   iStr(48) = "  中    華    民    國 " & String((10 - LenB(StrConv((txt1(16)), vbFromUnicode))) / 2, " ") & txt1(16) & String((10 - LenB(StrConv((txt1(16)), vbFromUnicode))) / 2, " ") & "年" & String((10 - LenB(StrConv((txt1(17)), vbFromUnicode))) / 2, " ") & txt1(17) & String((10 - LenB(StrConv((txt1(17)), vbFromUnicode))) / 2, " ") & "月" & String((10 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & txt1(18) & String((10 - LenB(StrConv((txt1(18)), vbFromUnicode))) / 2, " ") & "日"
    strDetail = ""
    For intI = 1 To UBound(iStr)
       If Trim(iStr(intI)) <> "" Then
          If intI >= 19 And intI <= 22 Then
             Select Case intI
                 Case 19: strDetail = strDetail & vbCrLf & RTrim(strExc(3))
                 Case 20: strDetail = strDetail & vbCrLf & RTrim(strExc(4))
                 Case 21: strDetail = strDetail & vbCrLf & RTrim(strExc(5))
                 Case 22: strDetail = strDetail & vbCrLf & RTrim(strExc(6))
             End Select
          Else
             If intI <= 18 Or (intI >= 38 And intI <= 43) Or intI = 48 Then
                 If intI = 38 Then
                   strDetail = strDetail & vbCrLf & vbCrLf & IIf(opt1(0).Value = True, "會稿", "不會稿") & vbCrLf
                 End If
                strDetail = strDetail & vbCrLf & RTrim(iStr(intI))
             End If
          End If
       End If
    Next
    If PUB_AddRecSeal("3", txtPCnt.Text, IIf(pSpace = True, "Y", ""), strDetail, Combo2.Text) Then
    End If
End If
          
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Number & ":" & Err.Description, , "錯誤 "
   End If
   
End Sub

'Added by Lydia 2022/01/22 刪除暫存檔
Private Sub RunEndProc(ByVal bolSleep As Boolean)
   If bolSleep = True Then Sleep 3000
   PUB_KillTempFile (strUserNum & "\$$" & strUserNum & "*_T*.*")
   PUB_KillTempFile (strUserNum & "\$$" & Me.Name & "*.*")
    
End Sub
