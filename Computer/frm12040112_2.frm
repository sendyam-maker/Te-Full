VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040112_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "資料刪除記錄檔"
   ClientHeight    =   6192
   ClientLeft      =   540
   ClientTop       =   516
   ClientWidth     =   8640
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6192
   ScaleWidth      =   8640
   Begin VB.TextBox textDD28 
      Height          =   285
      Left            =   1860
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1023
      Width           =   1215
   End
   Begin VB.TextBox textDD23 
      Height          =   285
      Left            =   5970
      MaxLength       =   6
      TabIndex        =   19
      Top             =   4053
      Width           =   852
   End
   Begin VB.TextBox textDD15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   3447
      Width           =   2052
   End
   Begin VB.TextBox textDD10_2 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2242
      Width           =   2412
   End
   Begin VB.TextBox textDD07_2 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   2490
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   1932
      Width           =   2172
   End
   Begin VB.TextBox textDD01 
      Height          =   285
      Left            =   1860
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox textDD02 
      Height          =   285
      Left            =   2340
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox textDD03 
      Height          =   285
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox textDD04 
      Height          =   285
      Left            =   3300
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox textDD06 
      Height          =   285
      Left            =   1860
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1629
      Width           =   1215
   End
   Begin VB.TextBox textDD07 
      Height          =   285
      Left            =   1860
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1932
      Width           =   492
   End
   Begin VB.TextBox textDD08 
      Height          =   285
      Left            =   5970
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1932
      Width           =   2532
   End
   Begin VB.TextBox textDD09 
      Height          =   285
      Left            =   5970
      MaxLength       =   12
      TabIndex        =   10
      Top             =   2235
      Width           =   1572
   End
   Begin VB.TextBox textDD10 
      Height          =   285
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2235
      Width           =   252
   End
   Begin VB.TextBox textDD11 
      Height          =   285
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   13
      Top             =   3144
      Width           =   252
   End
   Begin VB.TextBox textDD12 
      Height          =   285
      Left            =   1860
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2538
      Width           =   1215
   End
   Begin VB.TextBox textDD13 
      Height          =   285
      Left            =   1860
      MaxLength       =   9
      TabIndex        =   12
      Top             =   2841
      Width           =   1215
   End
   Begin VB.TextBox textDD14 
      Height          =   285
      Left            =   5910
      MaxLength       =   9
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox textDD15 
      Height          =   285
      Left            =   1860
      MaxLength       =   4
      TabIndex        =   14
      Top             =   3447
      Width           =   612
   End
   Begin VB.TextBox textDD16 
      Height          =   285
      Left            =   1860
      MaxLength       =   7
      TabIndex        =   16
      Top             =   3750
      Width           =   1215
   End
   Begin VB.TextBox textDD17 
      Height          =   285
      Left            =   5970
      MaxLength       =   7
      TabIndex        =   17
      Top             =   3750
      Width           =   1215
   End
   Begin VB.TextBox textDD18 
      Height          =   285
      Left            =   5970
      MaxLength       =   7
      TabIndex        =   15
      Top             =   3447
      Width           =   1215
   End
   Begin VB.TextBox textDD19 
      Height          =   285
      Left            =   1860
      MaxLength       =   6
      TabIndex        =   18
      Top             =   4053
      Width           =   852
   End
   Begin VB.TextBox textDD20 
      Height          =   285
      Left            =   1860
      MaxLength       =   8
      TabIndex        =   20
      Top             =   4356
      Width           =   1215
   End
   Begin VB.TextBox textDD21 
      Height          =   285
      Left            =   5970
      MaxLength       =   8
      TabIndex        =   21
      Top             =   4356
      Width           =   1215
   End
   Begin VB.TextBox textDD22 
      Height          =   285
      Left            =   1860
      MaxLength       =   15
      TabIndex        =   22
      Top             =   4659
      Width           =   1812
   End
   Begin VB.TextBox textDD25 
      Height          =   285
      Left            =   1860
      MaxLength       =   7
      TabIndex        =   23
      Top             =   4962
      Width           =   1215
   End
   Begin VB.TextBox textDD26 
      Height          =   285
      Left            =   5970
      MaxLength       =   6
      TabIndex        =   24
      Top             =   4962
      Width           =   852
   End
   Begin VB.TextBox textDD27 
      Height          =   285
      Left            =   1860
      MaxLength       =   7
      TabIndex        =   25
      Top             =   5265
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1110
      Top             =   5670
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040112_2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040112_2.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040112_2.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040112_2.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040112_2.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040112_2.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040112_2.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040112_2.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040112_2.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040112_2.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040112_2.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   528
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   931
      ButtonWidth     =   1101
      ButtonHeight    =   889
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox textDD24 
      Height          =   495
      Left            =   1860
      TabIndex        =   26
      Top             =   5580
      Width           =   6615
      VariousPropertyBits=   -1466941413
      MaxLength       =   60
      Size            =   "11668;873"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD26_2 
      Height          =   285
      Left            =   6900
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4962
      Width           =   1305
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "2302;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD23_2 
      Height          =   285
      Left            =   6900
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   4053
      Width           =   1305
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "2302;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD19_2 
      Height          =   285
      Left            =   2820
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4053
      Width           =   1305
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "2302;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD13_2 
      Height          =   285
      Left            =   3200
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2841
      Width           =   5295
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "9340;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD12_2 
      Height          =   285
      Left            =   3200
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2538
      Width           =   5295
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "9340;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD05 
      Height          =   285
      Left            =   1860
      TabIndex        =   58
      Top             =   1326
      Width           =   6615
      VariousPropertyBits=   -1467989989
      MaxLength       =   160
      Size            =   "11668;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD06_2 
      Height          =   285
      Left            =   3150
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   1629
      Width           =   5295
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "9340;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "序號 :"
      Height          =   255
      Left            =   120
      TabIndex        =   56
      Top             =   1023
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   240
      Left            =   120
      TabIndex        =   52
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "案件中文名稱："
      Height          =   240
      Left            =   120
      TabIndex        =   51
      Top             =   1326
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "申請人："
      Height          =   240
      Left            =   120
      TabIndex        =   50
      Top             =   1629
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "申請國家："
      Height          =   240
      Left            =   120
      TabIndex        =   49
      Top             =   1932
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "申請案號："
      Height          =   240
      Left            =   4890
      TabIndex        =   48
      Top             =   1932
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "分所案號："
      Height          =   255
      Left            =   4890
      TabIndex        =   47
      Top             =   2250
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "專利/商標種類代號："
      Height          =   240
      Left            =   120
      TabIndex        =   46
      Top             =   2257
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "目前准駁："
      Height          =   240
      Left            =   120
      TabIndex        =   45
      Top             =   3144
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "FC代理人："
      Height          =   240
      Left            =   120
      TabIndex        =   44
      Top             =   2538
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "年費/延展代理人："
      Height          =   240
      Left            =   120
      TabIndex        =   43
      Top             =   2841
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "總收文號："
      Height          =   240
      Left            =   4860
      TabIndex        =   42
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label15 
      Caption         =   "案件性質："
      Height          =   240
      Left            =   120
      TabIndex        =   41
      Top             =   3447
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "本所期限："
      Height          =   240
      Left            =   120
      TabIndex        =   40
      Top             =   3750
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "法定期限："
      Height          =   240
      Left            =   4890
      TabIndex        =   39
      Top             =   3750
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "收文日："
      Height          =   240
      Left            =   4890
      TabIndex        =   38
      Top             =   3447
      Width           =   855
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   120
      TabIndex        =   37
      Top             =   4053
      Width           =   900
   End
   Begin VB.Label Label20 
      Caption         =   "費用："
      Height          =   225
      Left            =   120
      TabIndex        =   36
      Top             =   4356
      Width           =   615
   End
   Begin VB.Label Label21 
      Caption         =   "規費："
      Height          =   240
      Left            =   4890
      TabIndex        =   35
      Top             =   4356
      Width           =   615
   End
   Begin VB.Label Label22 
      Caption         =   "收據編號/請款編號："
      Height          =   240
      Left            =   120
      TabIndex        =   34
      Top             =   4659
      Width           =   1815
   End
   Begin VB.Label Label23 
      Caption         =   "失誤人員："
      Height          =   240
      Left            =   4890
      TabIndex        =   33
      Top             =   4053
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "刪除備註："
      Height          =   240
      Left            =   120
      TabIndex        =   32
      Top             =   5580
      Width           =   972
   End
   Begin VB.Label Label25 
      Caption         =   "原資料產生日期："
      Height          =   240
      Left            =   120
      TabIndex        =   31
      Top             =   4962
      Width           =   1575
   End
   Begin VB.Label Label26 
      Caption         =   "原資料產生人員："
      Height          =   255
      Left            =   4410
      TabIndex        =   30
      Top             =   4962
      Width           =   1455
   End
   Begin VB.Label Label27 
      Caption         =   "刪除日期："
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5265
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "准(1)/駁(2)"
      Height          =   255
      Left            =   2250
      TabIndex        =   28
      Top             =   3144
      Width           =   855
   End
End
Attribute VB_Name = "frm12040112_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/15 改成Form2.0 ; textDD05、textDD06_2、textDD12_2、textDD13_2、textDD19_2、textDD23_2、textDD24
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Const MAX_FIELD = 28

' 本所案號
Dim m_DD01 As String
Dim m_DD02 As String
Dim m_DD03 As String
Dim m_DD04 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList(MAX_FIELD) As FIELDITEM

' 變數宣告區
Dim m_EditMode As Integer
Dim m_SubMode As Integer

' 記錄所有資料的串列
Dim m_DataList() As String
' 記錄所有可瀏覽資料的總筆數
Dim m_DataListCount As Integer

' 目前正在作用的資料項目索引
Dim m_CurrDL As Integer

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

' Load Form
Private Sub Form_Load()
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040112_2", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040112_2", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040112_2", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040112_2", strFind, False)
   
   textDD01.BackColor = &H8000000F
   textDD02.BackColor = &H8000000F
   textDD03.BackColor = &H8000000F
   textDD04.BackColor = &H8000000F
   textDD06_2.BackColor = &H8000000F
   textDD07_2.BackColor = &H8000000F
   textDD10_2.BackColor = &H8000000F
   textDD12_2.BackColor = &H8000000F
   textDD13_2.BackColor = &H8000000F
   textDD15_2.BackColor = &H8000000F
   textDD19_2.BackColor = &H8000000F
   textDD23_2.BackColor = &H8000000F
   textDD26_2.BackColor = &H8000000F
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ClearDataList
   ClearFieldList
   'Add By Cheng 2002/07/18
   Set frm12040112_2 = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' 設定資料
Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_DD01 = Empty
      m_DD02 = Empty
      m_DD03 = Empty
      m_DD04 = Empty
      ClearDataList
   End If
   
   Select Case nType
      ' 本所案號
      Case 0: m_DD01 = strData
      Case 1: m_DD02 = strData
      Case 2: m_DD03 = strData
      Case 3: m_DD04 = strData
      Case 4: SetDataListItem strData
   End Select
End Sub

' 刪除資料串列
Private Sub ClearDataList()
   If m_DataListCount > 0 Then
      Erase m_DataList
   End If
   m_DataListCount = 0
End Sub

' 設定資料串列
Private Sub SetDataListItem(ByVal strData As String)
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   For nIndex = 0 To m_DataListCount - 1
      If m_DataList(nIndex) = strData Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_DataList(m_DataListCount + 1)
      m_DataList(m_DataListCount) = strData
      m_DataListCount = m_DataListCount + 1
   End If
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "DD" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 16, 17, 18, 20, 21, 25, 27, 28:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
End Sub

Private Sub ClearFieldList()
   Erase m_FieldList
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To MAX_FIELD - 1
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 更新欄位的內容
Private Sub UpdateFieldNewData()
   SetFieldNewData "DD01", m_DD01
   SetFieldNewData "DD02", m_DD02
   SetFieldNewData "DD03", m_DD03
   SetFieldNewData "DD04", m_DD04
   SetFieldNewData "DD05", textDD05
   If IsEmptyText(textDD06) = False Then
      SetFieldNewData "DD06", textDD06 & String(9 - Len(textDD06), "0")
   Else
      SetFieldNewData "DD06", textDD06
   End If
   SetFieldNewData "DD07", textDD07
   SetFieldNewData "DD08", textDD08
   SetFieldNewData "DD09", textDD09
   SetFieldNewData "DD10", textDD10
   SetFieldNewData "DD11", textDD11
   If IsEmptyText(textDD12) = False Then
      SetFieldNewData "DD12", textDD12 & String(9 - Len(textDD12), "0")
   Else
      SetFieldNewData "DD12", textDD12
   End If
   If IsEmptyText(textDD13) = False Then
      SetFieldNewData "DD13", textDD13 & String(9 - Len(textDD13), "0")
   Else
      SetFieldNewData "DD13", textDD13
   End If
   SetFieldNewData "DD14", textDD14
   SetFieldNewData "DD15", textDD15
   If IsEmptyText(textDD16) = False Then
      SetFieldNewData "DD16", DBDATE(textDD16)
   Else
      SetFieldNewData "DD16", textDD16
   End If
   If IsEmptyText(textDD17) = False Then
      SetFieldNewData "DD17", DBDATE(textDD17)
   Else
      SetFieldNewData "DD17", textDD17
   End If
   If IsEmptyText(textDD18) = False Then
      SetFieldNewData "DD18", DBDATE(textDD18)
   Else
      SetFieldNewData "DD18", textDD18
   End If
   SetFieldNewData "DD19", textDD19
   SetFieldNewData "DD20", textDD20
   SetFieldNewData "DD21", textDD21
   SetFieldNewData "DD22", textDD22
   SetFieldNewData "DD23", textDD23
   SetFieldNewData "DD24", textDD24
   If IsEmptyText(textDD25) = False Then
      SetFieldNewData "DD25", DBDATE(textDD25)
   Else
      SetFieldNewData "DD25", textDD25
   End If
   SetFieldNewData "DD26", textDD26
   If IsEmptyText(textDD27) = False Then
      SetFieldNewData "DD27", DBDATE(textDD27)
   Else
      SetFieldNewData "DD27", textDD27
   End If
   SetFieldNewData "DD28", textDD28
   
   If m_EditMode = 1 Then
      SetFieldNewData "DD28", CStr(GetMaxNumber())
   End If
End Sub

Private Function GetMaxNumber() As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   strSql = "SELECT MAX(DD28) FROM DATADELETERECORD "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields(0)) = False Then
         GetMaxNumber = rsTmp.Fields(0)
         GetMaxNumber = Val(GetMaxNumber) + 1
      End If
   End If
   If IsEmptyText(GetMaxNumber) = True Then
      GetMaxNumber = 1
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 讀取資料庫所有的資料
Public Sub QueryDB()
   textDD01 = m_DD01
   textDD02 = m_DD02
   textDD03 = m_DD03
   textDD04 = m_DD04
   EnableTextBox textDD28, False
   
   InitialField
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   textDD01 = m_DD01
   textDD02 = m_DD02
   textDD03 = m_DD03
   textDD04 = m_DD04
   textDD05 = Empty
   textDD06 = Empty
   textDD06_2 = Empty
   textDD07 = Empty
   textDD07_2 = Empty
   textDD08 = Empty
   textDD09 = Empty
   textDD10 = Empty
   textDD10_2 = Empty
   textDD11 = Empty
   textDD12 = Empty
   textDD12_2 = Empty
   textDD13 = Empty
   textDD13_2 = Empty
   textDD14 = Empty
   textDD15 = Empty
   textDD15_2 = Empty
   textDD16 = Empty
   textDD17 = Empty
   textDD18 = Empty
   textDD19 = Empty
   textDD19_2 = Empty
   textDD20 = Empty
   textDD21 = Empty
   textDD22 = Empty
   textDD23 = Empty
   textDD23_2 = Empty
   textDD24 = Empty
   textDD25 = Empty
   textDD26 = Empty
   textDD26_2 = Empty
   textDD27 = Empty
   textDD28 = Empty
   
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textDD05.Locked = bEnable
   textDD06.Locked = bEnable
   textDD07.Locked = bEnable
   textDD08.Locked = bEnable
   textDD09.Locked = bEnable
   textDD10.Locked = bEnable
   textDD11.Locked = bEnable
   textDD12.Locked = bEnable
   textDD13.Locked = bEnable
   textDD14.Locked = bEnable
   textDD15.Locked = bEnable
   textDD16.Locked = bEnable
   textDD17.Locked = bEnable
   textDD18.Locked = bEnable
   textDD19.Locked = bEnable
   textDD20.Locked = bEnable
   textDD21.Locked = bEnable
   textDD22.Locked = bEnable
   textDD23.Locked = bEnable
   textDD24.Locked = bEnable
   textDD25.Locked = bEnable
   textDD26.Locked = bEnable
   textDD27.Locked = bEnable
   'textDD28.Locked = bEnable
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textDD14.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   If m_CurrDL < 0 Or m_CurrDL >= m_DataListCount Then
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM DATADELETERECORD " & _
            "WHERE DD01 = '" & m_DD01 & "' AND " & _
                  "DD02 = '" & m_DD02 & "' AND " & _
                  "DD03 = '" & m_DD03 & "' AND " & _
                  "DD04 = '" & m_DD04 & "' AND " & _
                  "DD28 = " & m_DataList(m_CurrDL) & " "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If Not IsNull(rsTmp.Fields("DD05")) Then: textDD05 = rsTmp.Fields("DD05")
      If Not IsNull(rsTmp.Fields("DD06")) Then: textDD06 = rsTmp.Fields("DD06")
      If Not IsNull(rsTmp.Fields("DD07")) Then
         textDD07 = rsTmp.Fields("DD07")
         'm_DD07 = rsTmp.Fields("DD07")
      End If
      If Not IsNull(rsTmp.Fields("DD08")) Then: textDD08 = rsTmp.Fields("DD08")
      If Not IsNull(rsTmp.Fields("DD09")) Then: textDD09 = rsTmp.Fields("DD09")
      If Not IsNull(rsTmp.Fields("DD10")) Then: textDD10 = rsTmp.Fields("DD10")
      If Not IsNull(rsTmp.Fields("DD11")) Then: textDD11 = rsTmp.Fields("DD11")
      If Not IsNull(rsTmp.Fields("DD12")) Then: textDD12 = rsTmp.Fields("DD12"):
      If Not IsNull(rsTmp.Fields("DD13")) Then: textDD13 = rsTmp.Fields("DD13")
      If Not IsNull(rsTmp.Fields("DD14")) Then: textDD14 = rsTmp.Fields("DD14")
      If Not IsNull(rsTmp.Fields("DD15")) Then: textDD15 = rsTmp.Fields("DD15")
      If Not IsNull(rsTmp.Fields("DD16")) Then: textDD16 = TAIWANDATE(rsTmp.Fields("DD16"))
      If Not IsNull(rsTmp.Fields("DD17")) Then: textDD17 = TAIWANDATE(rsTmp.Fields("DD17"))
      If Not IsNull(rsTmp.Fields("DD18")) Then: textDD18 = TAIWANDATE(rsTmp.Fields("DD18"))
      If Not IsNull(rsTmp.Fields("DD19")) Then: textDD19 = rsTmp.Fields("DD19")
      If Not IsNull(rsTmp.Fields("DD20")) Then: textDD20 = rsTmp.Fields("DD20")
      If Not IsNull(rsTmp.Fields("DD21")) Then: textDD21 = rsTmp.Fields("DD21")
      If Not IsNull(rsTmp.Fields("DD22")) Then: textDD22 = rsTmp.Fields("DD22")
      If Not IsNull(rsTmp.Fields("DD23")) Then: textDD23 = rsTmp.Fields("DD23")
      If Not IsNull(rsTmp.Fields("DD24")) Then: textDD24 = rsTmp.Fields("DD24")
      If Not IsNull(rsTmp.Fields("DD25")) Then: textDD25 = TAIWANDATE(rsTmp.Fields("DD25"))
      If Not IsNull(rsTmp.Fields("DD26")) Then: textDD26 = rsTmp.Fields("DD26")
      If Not IsNull(rsTmp.Fields("DD27")) Then: textDD27 = TAIWANDATE(rsTmp.Fields("DD27"))
      If Not IsNull(rsTmp.Fields("DD28")) Then: textDD28 = rsTmp.Fields("DD28")
      ' 更新欄位的內容
      UpdateFieldOldData rsTmp
      
      textDD06_Validate False
      textDD07_Validate False
      textDD10_Validate False
      textDD12_Validate False
      textDD13_Validate False
      textDD15_Validate False
      textDD19_Validate False
      textDD23_Validate False
      textDD26_Validate False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
EXITSUB:
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strDD28 As String)
   Dim strTemp As String
   Dim nIndex As Integer
   
   If IsRecordExist(strDD28) = True Then
      For nIndex = 0 To m_DataListCount - 1
         If m_DataList(nIndex) = strDD28 Then
            m_CurrDL = nIndex
            Exit For
         End If
      Next nIndex
   Else
      m_CurrDL = 0
      strTemp = Empty
      For nIndex = 0 To m_DataListCount - 1
         If strDD28 > strTemp Then
            m_CurrDL = nIndex
            Exit For
         End If
         strTemp = m_DataList(nIndex)
      Next nIndex
   End If
   UpdateCtrlData
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   If m_DataListCount > 0 Then
      m_CurrDL = 0
   End If
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   If m_CurrDL = 0 Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   If m_CurrDL > 0 Then
      m_CurrDL = m_CurrDL - 1
   End If
   
   UpdateCtrlData
   
EXITSUB:
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   If m_CurrDL >= m_DataListCount - 1 Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   If m_CurrDL < (m_DataListCount - 1) Then
      m_CurrDL = m_CurrDL + 1
   End If
   
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   If m_DataListCount > 0 Then
      m_CurrDL = m_DataListCount - 1
   End If
   
   UpdateCtrlData
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         ' 90.07.13 modify by louis (依照權限設定其工具列的按紐狀態)
         'tlbar.Buttons(1).Enabled = True
         'tlbar.Buttons(2).Enabled = True
         'tlbar.Buttons(3).Enabled = True
         'tlbar.Buttons(4).Enabled = True
         'tlbar.Buttons(6).Enabled = True
         'tlbar.Buttons(7).Enabled = True
         'tlbar.Buttons(8).Enabled = True
         'tlbar.Buttons(9).Enabled = True
         'tlbar.Buttons(11).Enabled = False
         'tlbar.Buttons(12).Enabled = False
         'tlbar.Buttons(14).Enabled = True
         
         If m_bInsert Then
            tlbar.Buttons(1).Enabled = True
         Else
            tlbar.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            tlbar.Buttons(2).Enabled = True
         Else
            tlbar.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            tlbar.Buttons(3).Enabled = True
         Else
            tlbar.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(4).Enabled = True
         Else
            tlbar.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(6).Enabled = True
            tlbar.Buttons(7).Enabled = True
            tlbar.Buttons(8).Enabled = True
            tlbar.Buttons(9).Enabled = True
         Else
            tlbar.Buttons(6).Enabled = False
            tlbar.Buttons(7).Enabled = False
            tlbar.Buttons(8).Enabled = False
            tlbar.Buttons(9).Enabled = False
         End If
         tlbar.Buttons(11).Enabled = False
         tlbar.Buttons(12).Enabled = False
         tlbar.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         tlbar.Buttons(1).Enabled = False
         tlbar.Buttons(2).Enabled = False
         tlbar.Buttons(3).Enabled = False
         tlbar.Buttons(4).Enabled = False
         tlbar.Buttons(6).Enabled = False
         tlbar.Buttons(7).Enabled = False
         tlbar.Buttons(8).Enabled = False
         tlbar.Buttons(9).Enabled = False
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         tlbar.Buttons(14).Enabled = False
   End Select
   
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 案件名稱(中)
Private Sub textDD05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textDD05, 160) = False Then
      Cancel = True
      textDD05_GotFocus
   End If
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: textDD05.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textDD06_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textDD08_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textDD09_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 專利/商標種類
Private Sub textDD10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textDD10_2 = Empty
   If IsEmptyText(textDD10) = False Then
      Select Case m_DD01
         Case "T", "TF", "CFT", "FCT":
            If textDD07 < "010" Then
               textDD10_2 = GetTradeMarkName(textDD10, 0)
            Else
               textDD10_2 = GetTradeMarkName(textDD10, 1)
            End If
            If IsEmptyText(textDD10_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "專利/商標種類代號不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textDD10_GotFocus
            End If
         Case "P", "CFP", "FCP":
            If textDD07 < "010" Then
               textDD10_2 = GetPatentName(textDD10, 0)
            Else
               textDD10_2 = GetPatentName(textDD10, 1)
            End If
            If IsEmptyText(textDD10_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "專利/商標種類代號不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textDD10_GotFocus
            End If
      End Select
   End If
End Sub

' 申請國家
Private Sub textDD07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textDD07_2 = Empty
   If IsEmptyText(textDD07) = False Then
      textDD07_2 = GetNationName(textDD07, 0)
      If IsEmptyText(textDD07_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請國家代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD07_GotFocus
      End If
   End If
End Sub

' 目前准駁
Private Sub textDD11_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDD11) = False Then
      Select Case textDD11
         Case "", " ", "1", "2":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入 1 或 2 "
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDD11_GotFocus
      End Select
   End If
End Sub

Private Sub textDD12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textDD13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 年費延展代理人
Private Sub textDD13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDD13) = False Then
      Select Case Mid(textDD13, 1, 1)
         Case "X":
            textDD13_2 = GetCustomerName(textDD13, 0)
            If IsEmptyText(textDD13_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "年費延展代理人代號不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textDD13_GotFocus
            End If
         Case "Y":
            textDD13_2 = GetFAgentName(textDD13)
            If IsEmptyText(textDD13_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "年費延展代理人代號不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textDD13_GotFocus
            End If
      End Select
   End If
End Sub

Private Sub textDD14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 總收文號
Private Sub textDD14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textDD14) = False Then
      Select Case m_EditMode
         Case 1, 2, 4:
            Select Case Mid(textDD14, 1, 1)
               'Modified by Lydia 2016/12/26 + D類收文
               Case "A", "B", "C", "D":
               Case Else:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "收文號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textDD14_GotFocus
               GoTo EXITSUB
            End Select
         Case Else:
      End Select
      Select Case m_EditMode
         Case 1:
            If IsRecordExist(textDD14) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "該筆記錄已存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textDD14_GotFocus
               GoTo EXITSUB
            End If
         Case Else:
      End Select
   End If
EXITSUB:
End Sub

' 案件性質
Private Sub textDD15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textDD15_2 = Empty
   If IsEmptyText(textDD15) = False Then
      'If textDD07 < "010" Then
      '   textDD15_2 = GetCaseTypeName(m_DD01, textDD15, 0)
      'Else
      '   textDD15_2 = GetCaseTypeName(m_DD01, textDD15, 1)
      'End If
      '2005/5/13 MODIFY BY SONIA
      'textDD15_2 = GetCaseTypeName(m_DD01, textDD15, 0)
      'modify by sonia 2021/10/28 改前一畫面之申請國家暫存欄位，否則T-232571會顯示（無）
      'Select Case textDD07
      Select Case frm12040112_1.strDD07
         Case "000", ""
            textDD15_2 = GetCaseTypeName(m_DD01, textDD15, 0)
         Case Else
            textDD15_2 = GetCaseTypeName(m_DD01, textDD15, 1)
      End Select
      '2005/5/13 END
      If IsEmptyText(textDD15_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD15_GotFocus
      End If
   End If
End Sub

' 本所期限
Private Sub textDD16_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDD16) = False Then
      If CheckIsTaiwanDate(textDD16, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "本所期限日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD16_GotFocus
      End If
   End If
End Sub

' 法定期限
Private Sub textDD17_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDD17) = False Then
      If CheckIsTaiwanDate(textDD17, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "法定期限日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD17_GotFocus
      End If
   End If
End Sub

' 申請人欄位
Private Sub textDD06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   textDD06_2 = Empty
   If IsEmptyText(textDD06) = False Then
      textDD06_2 = GetCustomerName(textDD06, 0)
      If IsEmptyText(textDD06_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD06_GotFocus
      End If
   End If
End Sub

' FC代理人
Private Sub textDD12_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   textDD12_2 = Empty
   If IsEmptyText(textDD12) = False Then
      textDD12_2 = GetFAgentName(textDD12)
      If IsEmptyText(textDD12_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "FC代理人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD12_GotFocus
      End If
   End If
End Sub

' 收文日
Private Sub textDD18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDD18) = False Then
      If CheckIsTaiwanDate(textDD18, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD18_GotFocus
      End If
   End If
End Sub

Private Sub textDD19_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 智權人員代號
Private Sub textDD19_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textDD19_2 = Empty
   If IsEmptyText(textDD19) = False Then
      textDD19_2 = GetStaffName(textDD19, True)
      If IsEmptyText(textDD19_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "智權人員代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD19_GotFocus
      End If
   End If
End Sub

' 費用
Private Sub textDD20_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDD20) = False Then
      If IsNumeric(textDD20) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "費用只可輸入數值的資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD20_GotFocus
      End If
   End If
End Sub

' 規費
Private Sub textDD21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDD21) = False Then
      If IsNumeric(textDD21) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "規費只可輸入數值的資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD21_GotFocus
      End If
   End If
End Sub

Private Sub textDD22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textDD23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 失誤人員代號
Private Sub textDD23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textDD23_2 = Empty
   If IsEmptyText(textDD23) = False Then
      textDD23_2 = GetStaffName(textDD23, True)
      If IsEmptyText(textDD23_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "失誤人員代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD23_GotFocus
      End If
   End If
End Sub

' 刪除備註
Private Sub textDD24_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textDD24, 60) = False Then
      Cancel = True
      textDD24_GotFocus
   End If
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: textDD24.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 原資料產生日期
Private Sub textDD25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDD25) = False Then
      If CheckIsTaiwanDate(textDD25, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "原資料產生日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD25_GotFocus
      End If
   End If
End Sub

Private Sub textDD26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 原資料產生人員
Private Sub textDD26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textDD26_2 = Empty
   If IsEmptyText(textDD26) = False Then
      textDD26_2 = GetStaffName(textDD26, True)
      If IsEmptyText(textDD26_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "原資料產生人員代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD26_GotFocus
      End If
   End If
End Sub

' 刪除日期
Private Sub textDD27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDD27) = False Then
      If CheckIsTaiwanDate(textDD27, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "刪除日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD27_GotFocus
      End If
   End If
End Sub

' 序號
Private Sub textDD28_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(textDD28) = False Then
      If IsNumeric(textDD28) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "序號請輸入數值的資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDD28_GotFocus
      End If
   End If
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 90.07.13 modify by louis
      ' 新增
      'Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
      '   If m_EditMode = 0 Then
      '      OnAction KeyCode
      '      KeyCode = 0
      '   End If
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   m_SubMode = 0
   EnableTextBox textDD28, False
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            OnWork
            If m_DataListCount <= 0 Then
               GoTo EXITSUB
            Else
               UpdateToolbarState
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         EnableTextBox textDD28, True
         ClearField
         UpdateToolbarState
         SetInputEntry
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         OnWork
         UpdateToolbarState
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
         frm12040112_1.ClearRemark
         frm12040112_1.Show
   End Select
EXITSUB:
End Sub

' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: OnAction vbKeyF5
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

' 檢查資料庫中該筆記錄是否存在
Private Function IsDataBaseExist(ByVal strDD14) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   IsDataBaseExist = False
   strSql = "SELECT * FROM DataDeleteRecord " & _
            "WHERE DD01 = '" & m_DD01 & "' AND " & _
                  "DD02 = '" & m_DD02 & "' AND " & _
                  "DD03 = '" & m_DD03 & "' AND " & _
                  "DD04 = '" & m_DD04 & "' AND " & _
                  "DD14 = '" & strDD14 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsDataBaseExist = True
   Else
      IsDataBaseExist = False
   End If
   rsTmp.Close
EXITSUB:
   Set rsTmp = Nothing
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strDD28) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   IsRecordExist = False
   bFind = False
   For nIndex = 0 To m_DataListCount - 1
      If m_DataList(nIndex) = strDD28 Then
         bFind = True
      End If
   Next nIndex
   If bFind = False Then
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM DataDeleteRecord " & _
            "WHERE DD01 = '" & m_DD01 & "' AND " & _
                  "DD02 = '" & m_DD02 & "' AND " & _
                  "DD03 = '" & m_DD03 & "' AND " & _
                  "DD04 = '" & m_DD04 & "' AND " & _
                  "DD28 = '" & strDD28 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
EXITSUB:
   Set rsTmp = Nothing
End Function

' 新增記錄
Private Sub AddRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strDD28 As String

   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName = "DD28" Then
         strDD28 = m_FieldList(nIndex).fiNewData
         Exit For
      End If
   Next nIndex
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO DataDeleteRecord ("
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & m_FieldList(nIndex).fiNewData & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      ' 將新增的資料加入到串列中
      SetDataListItem strDD28
      ' 顯示該筆記錄
      ShowCurrRecord strDD28
      ' 通知前畫面有新增的記錄
      frm12040112_1.ModRecord strDD28
   End If
EXITSUB:
End Sub

' 修改記錄
Private Sub ModRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strDD28 As String
   
   strDD28 = m_DataList(m_CurrDL)
   
   strSql = "UPDATE DataDeleteRecord SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_FieldList(nIndex).fiName & " = '" & m_FieldList(nIndex).fiNewData & "'"
            End If
         Else
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   
   strSql = strSql & " " & _
            "WHERE DD01 = '" & m_DD01 & "' AND " & _
                  "DD02 = '" & m_DD02 & "' AND " & _
                  "DD03 = '" & m_DD03 & "' AND " & _
                  "DD04 = '" & m_DD04 & "' AND " & _
                  "DD28 = '" & strDD28 & "' "
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      ShowCurrRecord strDD28
      ' 通知前畫面有更新的記錄
      frm12040112_1.ModRecord strDD28
   End If
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strDD28 As String
   Dim strDataList() As String
   Dim nDataListCount As Integer
   Dim nIndex As Integer
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nPos As Integer
   
   strDD28 = m_DataList(m_CurrDL)
   nPos = m_CurrDL
   
   strSql = "DELETE FROM DataDeleteRecord " & _
            "WHERE DD01 = '" & m_DD01 & "' AND " & _
                        "DD02 = '" & m_DD02 & "' AND " & _
                        "DD03 = '" & m_DD03 & "' AND " & _
                        "DD04 = '" & m_DD04 & "' AND " & _
                        "DD28 = '" & strDD28 & "' "
   cnnConnection.Execute strSql
   
   ' 通知前畫面有刪除的記錄
   frm12040112_1.DelRecord strDD28

   ' 刪除記錄時, 除了要刪除資料庫中的記錄還要刪除在本程式模組中所記錄的資料
   ' 以供瀏覽資料時會更新資料串列的順序及正確性
   nDataListCount = 0
   For nIndex = 0 To m_DataListCount - 1
      If m_DataList(nIndex) <> strDD28 Then
         ReDim Preserve strDataList(nDataListCount + 1)
         strDataList(nDataListCount) = m_DataList(nIndex)
         nDataListCount = nDataListCount + 1
      End If
   Next nIndex
   ' 清除資料串列
   ClearDataList
   ' 將資料更新回到資料串列記錄中
   For nIndex = 0 To nDataListCount - 1
      SetDataListItem strDataList(nIndex)
   Next nIndex
   ' 刪除暫存串列
   If m_DataListCount <= 0 Then
      strTit = "資料顯示"
      strMsg = "該筆本所案號已無資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Unload Me
      frm12040112_1.ClearRemark
      frm12040112_1.Show
   Else
      'ShowCurrRecord strDD28
      If nPos <= m_DataListCount - 1 Then
         m_CurrDL = nPos
      Else
         m_CurrDL = m_DataListCount - 1
      End If
      UpdateCtrlData
   End If
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strDD28 As String
   Dim strFirst As String
   Dim bFirst As Boolean
   
   QueryRecord = False
   
   strDD28 = Empty
   strSql = "SELECT DD28 FROM DATADELETERECORD " & _
            "WHERE DD01 = '" & m_DD01 & "' AND " & _
                  "DD02 = '" & m_DD02 & "' AND " & _
                  "DD03 = '" & m_DD03 & "' AND " & _
                  "DD04 = '" & m_DD04 & "' "
   If IsEmptyText(textDD14) = False Then
      strSql = strSql & " AND DD14 = '" & textDD14 & "' "
   End If
   If IsEmptyText(textDD28) = False Then
      strSql = strSql & " AND DD28 = " & textDD28 & " "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      bFirst = True
      Do While rsTmp.EOF = False
         If IsNull(rsTmp.Fields("DD28")) = False Then
            strDD28 = rsTmp.Fields("DD28")
            SetDataListItem strDD28
            If bFirst = True Then
               strFirst = rsTmp.Fields("DD28")
               bFirst = False
            End If
         End If
         rsTmp.MoveNext
      Loop
   Else
      QueryRecord = False
      rsTmp.Close
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   QueryRecord = True
   ShowCurrRecord strFirst
   
   UpdateToolbarState
EXITSUB:
   Set rsTmp = Nothing
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Select Case m_EditMode
      Case 1:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            AddRecord
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
               
            ModRecord
         Else
            GoTo EXITSUB
         End If
      Case 3:
         DelRecord
         If m_DataListCount <= 0 Then
            GoTo EXITSUB
         End If
      Case 4:
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textDD14.SetFocus
      Case 2: textDD05.SetFocus
      Case 4: textDD14.SetFocus
   End Select
End Sub

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2:
         ' 失誤人員
         If IsEmptyText(textDD23) = True Then
            strTit = "檢核資料"
            strMsg = "失誤人員不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDD23.SetFocus
            GoTo EXITSUB
         End If
         ' 刪除備註
         If IsEmptyText(textDD24) = True Then
            strTit = "檢核資料"
            strMsg = "刪除備註不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDD24.SetFocus
            GoTo EXITSUB
         End If
         ' 原資料產生日期
         If IsEmptyText(textDD25) = True Then
            strTit = "檢核資料"
            strMsg = "原資料產生日期不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDD25.SetFocus
            GoTo EXITSUB
         End If
         ' 原資料產生人員
         If IsEmptyText(textDD26) = True Then
            strTit = "檢核資料"
            strMsg = "原資料產生人員不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDD26.SetFocus
            GoTo EXITSUB
         End If
         ' 刪除日期
         If IsEmptyText(textDD27) = True Then
            strTit = "檢核資料"
            strMsg = "刪除日期不可空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDD27.SetFocus
            GoTo EXITSUB
         End If
         ' 本所期限不可超過法定期限
         If IsEmptyText(textDD16) = False And IsEmptyText(textDD17) = False Then
            If Val(DBDATE(textDD16)) > Val(DBDATE(textDD17)) Then
               strTit = "檢核資料"
               strMsg = "本所期限不可超過法定期限"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textDD16.SetFocus
               GoTo EXITSUB
            End If
         End If
   End Select
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textDD05_GotFocus()
   InverseTextBox textDD05
   'edit by nickc 2007/07/11 切換輸入法改用API
   'textDD05.IMEMode = 1
   OpenIme
End Sub

Private Sub textDD06_GotFocus()
   InverseTextBox textDD06
End Sub

Private Sub textDD07_GotFocus()
   InverseTextBox textDD07
End Sub

Private Sub textDD08_GotFocus()
   InverseTextBox textDD08
End Sub

Private Sub textDD09_GotFocus()
   InverseTextBox textDD09
End Sub

Private Sub textDD10_GotFocus()
   InverseTextBox textDD10
End Sub

Private Sub textDD11_GotFocus()
   InverseTextBox textDD11
End Sub

Private Sub textDD12_GotFocus()
   InverseTextBox textDD12
End Sub

Private Sub textDD13_GotFocus()
   InverseTextBox textDD13
End Sub

Private Sub textDD14_GotFocus()
   InverseTextBox textDD14
End Sub

Private Sub textDD15_GotFocus()
   InverseTextBox textDD15
End Sub

Private Sub textDD16_GotFocus()
   InverseTextBox textDD16
End Sub

Private Sub textDD17_GotFocus()
   InverseTextBox textDD17
End Sub

Private Sub textDD18_GotFocus()
   InverseTextBox textDD18
End Sub

Private Sub textDD19_GotFocus()
   InverseTextBox textDD19
End Sub

Private Sub textDD20_GotFocus()
   InverseTextBox textDD20
End Sub

Private Sub textDD21_GotFocus()
   InverseTextBox textDD21
End Sub

Private Sub textDD22_GotFocus()
   InverseTextBox textDD22
End Sub

Private Sub textDD23_GotFocus()
   InverseTextBox textDD23
End Sub

Private Sub textDD24_GotFocus()
   InverseTextBox textDD24
   'edit by nickc 2007/07/11 切換輸入法改用API
   'textDD24.IMEMode = 1
   OpenIme
End Sub

Private Sub textDD25_GotFocus()
   InverseTextBox textDD25
End Sub

Private Sub textDD26_GotFocus()
   InverseTextBox textDD26
End Sub

Private Sub textDD27_GotFocus()
   InverseTextBox textDD27
End Sub

Private Sub textDD28_GotFocus()
   InverseTextBox textDD28
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textDD05.Enabled = True Then
   Cancel = False
   textDD05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD06.Enabled = True Then
   Cancel = False
   textDD06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD07.Enabled = True Then
   Cancel = False
   textDD07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD10.Enabled = True Then
   Cancel = False
   textDD10_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD11.Enabled = True Then
   Cancel = False
   textDD11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD12.Enabled = True Then
   Cancel = False
   textDD12_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD13.Enabled = True Then
   Cancel = False
   textDD13_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD14.Enabled = True Then
   Cancel = False
   textDD14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD15.Enabled = True Then
   Cancel = False
   textDD15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD16.Enabled = True Then
   Cancel = False
   textDD16_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD17.Enabled = True Then
   Cancel = False
   textDD17_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD18.Enabled = True Then
   Cancel = False
   textDD18_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD19.Enabled = True Then
   Cancel = False
   textDD19_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD20.Enabled = True Then
   Cancel = False
   textDD20_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD21.Enabled = True Then
   Cancel = False
   textDD21_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD23.Enabled = True Then
   Cancel = False
   textDD23_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD24.Enabled = True Then
   Cancel = False
   textDD24_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD25.Enabled = True Then
   Cancel = False
   textDD25_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD26.Enabled = True Then
   Cancel = False
   textDD26_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD27.Enabled = True Then
   Cancel = False
   textDD27_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDD28.Enabled = True Then
   Cancel = False
   textDD28_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Lydia 2021/10/15 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

TxtValidate = True
End Function

