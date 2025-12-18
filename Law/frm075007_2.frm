VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075007_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "下一程序資料維護"
   ClientHeight    =   5616
   ClientLeft      =   180
   ClientTop       =   960
   ClientWidth     =   9156
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5616
   ScaleWidth      =   9156
   Begin VB.TextBox textNP24 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      Height          =   288
      Left            =   8010
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1065
   End
   Begin VB.TextBox textNP23 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3465
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2985
      Width           =   1092
   End
   Begin VB.CommandButton cmdPrintReturnSheet 
      Caption         =   "列印案件回覆單(&R)"
      Height          =   345
      Left            =   7125
      TabIndex        =   41
      Top             =   1050
      Width           =   1755
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印接洽結案單(&P)"
      Height          =   345
      Left            =   7125
      TabIndex        =   40
      Top             =   660
      Width           =   1755
   End
   Begin VB.TextBox textNP22 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   5640
      MaxLength       =   9
      TabIndex        =   1
      Top             =   1080
      Width           =   1092
   End
   Begin VB.TextBox textNP12_2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      Height          =   288
      Left            =   6480
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3300
      Width           =   2532
   End
   Begin VB.TextBox textNP07_2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      Height          =   288
      Left            =   2025
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2700
      Width           =   2292
   End
   Begin VB.TextBox textCP27 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      Height          =   288
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox textCP05 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      Height          =   288
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   720
      Width           =   1092
   End
   Begin VB.TextBox textNP05 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   2625
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1080
      Width           =   372
   End
   Begin VB.TextBox textNP04 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   2385
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1080
      Width           =   252
   End
   Begin VB.TextBox textNP03 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   1665
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1080
      Width           =   732
   End
   Begin VB.TextBox textNP02 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   1185
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1080
      Width           =   492
   End
   Begin VB.TextBox textNP01 
      BackColor       =   &H00FFFFFF&
      Height          =   288
      Left            =   1185
      MaxLength       =   9
      TabIndex        =   0
      Top             =   720
      Width           =   1092
   End
   Begin VB.TextBox textNP06 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1350
      MaxLength       =   1
      TabIndex        =   10
      Top             =   3900
      Width           =   375
   End
   Begin VB.TextBox textNP10 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   3
      Top             =   2700
      Width           =   732
   End
   Begin VB.TextBox textNP08 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1350
      MaxLength       =   7
      TabIndex        =   4
      Top             =   2985
      Width           =   1092
   End
   Begin VB.TextBox textNP09 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6000
      MaxLength       =   7
      TabIndex        =   6
      Top             =   2985
      Width           =   1092
   End
   Begin VB.TextBox textNP07 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1350
      MaxLength       =   4
      TabIndex        =   2
      Top             =   2700
      Width           =   612
   End
   Begin VB.TextBox textNP11 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1350
      MaxLength       =   7
      TabIndex        =   7
      Top             =   3300
      Width           =   1092
   End
   Begin VB.TextBox textNP12 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   8
      Top             =   3300
      Width           =   375
   End
   Begin VB.TextBox textNP13 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1350
      MaxLength       =   50
      TabIndex        =   9
      Top             =   3600
      Width           =   7680
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8550
      Top             =   600
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
            Picture         =   "frm075007_2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075007_2.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075007_2.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075007_2.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075007_2.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075007_2.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075007_2.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075007_2.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075007_2.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075007_2.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm075007_2.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9156
      _ExtentX        =   16150
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
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
   Begin MSForms.TextBox textNP14 
      Height          =   285
      Left            =   1350
      TabIndex        =   50
      Top             =   4200
      Width           =   7680
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "13547;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNP15 
      Height          =   735
      Left            =   1350
      TabIndex        =   11
      Top             =   4500
      Width           =   7710
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13600;1296"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNP10_2 
      Height          =   285
      Left            =   6810
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2700
      Width           =   2175
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3831;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1185
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2130
      Width           =   2175
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3836;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1185
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1440
      Width           =   7725
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "13626;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1170
      TabIndex        =   46
      Top             =   1770
      Width           =   7725
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13626;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   285
      Left            =   150
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5280
      Width           =   8865
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "15637;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      Caption         =   "下一單據編號："
      Height          =   255
      Left            =   6690
      TabIndex        =   44
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label lblNP23 
      Caption         =   "約定期限："
      Height          =   255
      Left            =   2565
      TabIndex        =   42
      Top             =   3000
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   105
      X2              =   8925
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   105
      X2              =   8925
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Label Label12 
      Caption         =   "序       號 ："
      Height          =   252
      Left            =   4680
      TabIndex        =   39
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label9 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   105
      TabIndex        =   30
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "智權人員："
      Height          =   252
      Left            =   4680
      TabIndex        =   29
      Top             =   2712
      Width           =   1020
   End
   Begin VB.Label Label20 
      Caption         =   "本所期限："
      Height          =   255
      Left            =   105
      TabIndex        =   28
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label7 
      Caption         =   "收  文  日："
      Height          =   255
      Left            =   4680
      TabIndex        =   27
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號："
      Height          =   255
      Left            =   105
      TabIndex        =   26
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "承  辦  人："
      Height          =   255
      Left            =   105
      TabIndex        =   25
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "法定期限："
      Height          =   255
      Left            =   4680
      TabIndex        =   24
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "發  文  日："
      Height          =   252
      Left            =   4680
      TabIndex        =   23
      Top             =   2160
      Width           =   972
   End
   Begin VB.Label Label15 
      Caption         =   "是否續辦："
      Height          =   255
      Left            =   105
      TabIndex        =   22
      Top             =   3900
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "下一程序："
      Height          =   255
      Left            =   105
      TabIndex        =   21
      Top             =   2700
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   105
      TabIndex        =   20
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "申  請  人："
      Height          =   255
      Left            =   105
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "解除期限日期："
      Height          =   255
      Left            =   105
      TabIndex        =   18
      Top             =   3300
      Width           =   1335
   End
   Begin VB.Label Labe20 
      Caption         =   "解除期限原因："
      Height          =   252
      Left            =   4680
      TabIndex        =   17
      Top             =   3300
      Width           =   1332
   End
   Begin VB.Label Label5 
      Caption         =   "機關文號：  "
      Height          =   255
      Left            =   105
      TabIndex        =   16
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "相  關  人： "
      Height          =   255
      Left            =   105
      TabIndex        =   15
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "備        註："
      Height          =   255
      Left            =   105
      TabIndex        =   14
      Top             =   4530
      Width           =   1095
   End
   Begin VB.Label Label21 
      Caption         =   "(Y/N)"
      Height          =   255
      Left            =   1905
      TabIndex        =   13
      Top             =   3900
      Width           =   495
   End
End
Attribute VB_Name = "frm075007_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/12 改成Form2.0 ; cmbTM05、textTM23、textCP14、textNP10_2、textNP15
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

' 本所案號
Dim m_NP02 As String
Dim m_NP03 As String
Dim m_NP04 As String
Dim m_NP05 As String
' 申請國家
Dim m_Nation As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
'Modify by Morgan 2010/1/8
'Const MAX_FIELD = 22 '改用 TF_NP
'Dim m_FieldList(MAX_FIELD) As FIELDITEM
Dim m_FieldList() As FIELDITEM

' 變數宣告區
Dim m_EditMode As Integer 'Memo by Lydia 2020/07/13  1-新增, 2-修改
Dim m_SubMode As Integer

' 儲存單筆記錄的結構
Private Type DATAITEM
   diNP01 As String
   diNP07 As String
   diNP22 As String
End Type
' 記錄所有資料的串列
Dim m_DataList() As DATAITEM
' 記錄所有可瀏覽資料的總筆數
Dim m_DataListCount As Integer

' 目前正在作用的資料項目索引
Dim m_CurrDL As Integer
'
Dim m_AddData As Boolean

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_CP13 As String 'Add By Sindy 2014/9/11 智權人員
Dim m_MeTrackMode  As String 'Added by Lydia 2021/10/20 Form2.0 記錄鍵盤傳入順序
Dim bolUpdNp06 As Boolean 'Added by Lydia 2022/09/13 是否可以更新續辦欄位NP06

Private Sub cmdPrint_Click()
Dim strProgressNo As String
strProgressNo = Empty
'Add by Cheng 2001/12/10 (列印接洽結案單)
If IsEmptyText(textNP22) = False Then
    g_PrtForm001.PrintForm textNP22, Me.textNP02.Text, Me.textNP03.Text, Me.textNP04, Me.textNP05.Text
End If
End Sub

Private Sub cmdPrintReturnSheet_Click()
   Dim stOsPrinter As String
   If m_EditMode = 0 Then
      Screen.MousePointer = vbHourglass
      'Added by Morgan 2023/11/20 已改用Word列印,要切換控制台的印表機
      stOsPrinter = PUB_GetOsDefaultPrinter
      PUB_SetOsDefaultPrinter Printer.DeviceName
      'end 2023/11/20
      'Modify By Sindy 2022/4/18 + , , , , Me
      Call g_PrtForm001.PrintReturnSheet(textNP01, textNP07, DBDATE(textNP09), , , , , textNP02 & textNP03 & textNP04 & textNP05, , , , Me)
      PUB_SetOsDefaultPrinter stOsPrinter 'Added by Morgan 2023/11/20
      Screen.MousePointer = vbDefault
   Else
      MsgBox "編輯模式不可列印！", vbCritical
   End If
End Sub

Private Sub Form_Initialize()
   ReDim m_FieldList(TF_NP) As FIELDITEM
End Sub

' Load Form
Private Sub Form_Load()
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm075007_2", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm075007_2", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm075007_2", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm075007_2", strFind, False)
    
   'add by sonia 2013/11/29
   'FCP程序可查詢FMP案
   If Pub_StrUserSt03 = "F22" And (frm075007_1.textNP02 = "P" Or frm075007_1.textNP02 = "CFP") Then
      m_bInsert = False
      m_bUpdate = False
      m_bDelete = False
      m_bQuery = True
   End If
   '2013/11/29 end
   
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
    If FMP2open = True Then
       If PUB_FMPtoCheck(1, 1, Pub_strUserST05, frm075007_1.textNP02, frm075007_1.textNP03, IIf(frm075007_1.textNP04 = "", "0", frm075007_1.textNP04), IIf(frm075007_1.textNP05 = "", "00", frm075007_1.textNP05)) = True Then
        m_bInsert = True
        m_bUpdate = True
        m_bDelete = False
        m_bQuery = True
       End If
    End If
    
   textCP05.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   'textTM05.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP27.BackColor = &H8000000F
   textNP02.BackColor = &H8000000F
   textNP03.BackColor = &H8000000F
   textNP04.BackColor = &H8000000F
   textNP05.BackColor = &H8000000F
   textNP07_2.BackColor = &H8000000F
   textNP10_2.BackColor = &H8000000F
   textNP12_2.BackColor = &H8000000F
   textNP24.BackColor = &H8000000F 'Add By Sindy 2018/10/17
   textCUID.BackColor = &H8000000F
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
   '2010/5/4 ADD BY SONIA FCT,S才可印接洽結案單
   If frm075007_1.textNP02 = "FCT" Or frm075007_1.textNP02 = "S" Then
      cmdPrint.Enabled = True
      cmdPrint.Visible = True
   Else
      cmdPrint.Enabled = False
      cmdPrint.Visible = False
   End If
   '2010/5/4 END
   
   
   'Added by Morgan 2014/11/11
   'Modified by Morgan 2014/11/19 +P,PS
   'If frm075007_1.textNP02 = "FCP" Or frm075007_1.textNP02 = "FG" Then
   If frm075007_1.textNP02 = "FCP" Or frm075007_1.textNP02 = "FG" Or frm075007_1.textNP02 = "P" Or frm075007_1.textNP02 = "PS" Then
      textNP09.TabIndex = 4
   Else
      textNP09.TabIndex = 5
   End If
   'end 2014/11/11
   
   'Added by Lydia 2022/09/13 外商承辦F11判斷人員職稱等級決定是否鎖住「是否續辦」
   bolUpdNp06 = True
   If Pub_StrUserSt03 = "F11" Then
       strExc(0) = "select nvl(st20,'99') st20 from staff where st01='" & strUserNum & "' "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
           If "" & RsTemp.Fields("st20") > "52" Then
               bolUpdNp06 = False
           End If
       End If
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ClearDataList
   ClearFieldList
   'Add By Cheng 2002/07/18
   Set frm075007_2 = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' 設定資料
Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_NP02 = Empty
      m_NP03 = Empty
      m_NP04 = Empty
      m_NP05 = Empty
      ClearDataList
      m_AddData = False
   End If
   
   Select Case nType
      ' 本所案號
      Case 0: m_NP02 = strData
      Case 1: m_NP03 = strData
      Case 2: m_NP04 = strData
      Case 3: m_NP05 = strData
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
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim bFind As Boolean
   Dim nPos As Integer

   For nIndex = 0 To m_DataListCount - 1
      If m_DataList(nIndex).diNP22 = strData Then
         bFind = True
         nPos = nIndex
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_DataList(m_DataListCount + 1)
      m_DataList(m_DataListCount).diNP22 = strData
      nPos = m_DataListCount
      m_DataListCount = m_DataListCount + 1
   End If

   ' 取得總收文號及下一程序代號
   strSql = "SELECT NP01, NP07 FROM NEXTPROGRESS " & _
            "WHERE NP02 = '" & m_NP02 & "' AND " & _
                  "NP03 = '" & m_NP03 & "' AND " & _
                  "NP04 = '" & m_NP04 & "' AND " & _
                  "NP05 = '" & m_NP05 & "' AND " & _
                  "NP22 = " & strData & " "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("NP01")) = False Then
         m_DataList(nPos).diNP01 = rsTmp.Fields("NP01")
      End If
      If IsNull(rsTmp.Fields("NP07")) = False Then
         m_DataList(nPos).diNP07 = rsTmp.Fields("NP07")
      End If
   End If
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To TF_NP
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "NP" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 7, 8, 9, 11, 17, 18, 20, 21, 22:
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
   For nIndex = 0 To TF_NP - 1
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
   Dim strNP22 As String
   SetFieldNewData "NP01", textNP01
   SetFieldNewData "NP02", m_NP02
   SetFieldNewData "NP03", m_NP03
   SetFieldNewData "NP04", m_NP04
   SetFieldNewData "NP05", m_NP05
   SetFieldNewData "NP06", textNP06
   SetFieldNewData "NP07", textNP07
   If IsEmptyText(textNP08) = False Then
      SetFieldNewData "NP08", DBDATE(textNP08)
   Else
      SetFieldNewData "NP08", textNP08
   End If
   If IsEmptyText(textNP09) = False Then
      SetFieldNewData "NP09", DBDATE(textNP09)
   Else
      SetFieldNewData "NP09", textNP09
   End If
   SetFieldNewData "NP10", textNP10
   If IsEmptyText(textNP11) = False Then
      SetFieldNewData "NP11", DBDATE(textNP11)
   Else
      SetFieldNewData "NP11", textNP11
   End If
   SetFieldNewData "NP12", textNP12
   SetFieldNewData "NP13", textNP13
   SetFieldNewData "NP14", textNP14
   SetFieldNewData "NP15", textNP15
   If m_EditMode = 1 Then
      strNP22 = GetNextProgressNo()
      SetFieldNewData "NP22", strNP22
   End If
   
   SetFieldNewData "NP23", DBDATE(textNP23) 'Add by Morgan 2010/1/8
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To TF_NP - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 讀取資料庫所有的資料
Public Sub QueryDB()
   textNP02 = m_NP02
   textNP03 = m_NP03
   textNP04 = m_NP04
   textNP05 = m_NP05
   
   EnableTextBox textNP22, False
   
   InitialField
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   textNP01 = Empty
   textNP02 = m_NP02
   textNP03 = m_NP03
   textNP04 = m_NP04
   textNP05 = m_NP05
   textNP06 = Empty
   textNP06.Tag = textNP06 'Add by Amy 2022/06/20
   textNP07 = Empty
   textNP07_2 = Empty
   textNP08 = Empty
   textNP08.Tag = textNP08 'Added by Morgan 2014/11/19
   textNP09 = Empty
   textNP09.Tag = textNP09 'Added by Morgan 2014/11/19
   textNP10 = Empty
   textNP10_2 = Empty
   textNP11 = Empty
   textNP12 = Empty
   textNP24 = Empty 'Add By Sindy 2018/10/17
   textNP12_2 = Empty
   textNP13 = Empty
   textNP14 = Empty
   textNP15 = Empty
   textNP22 = Empty
   textNP23 = Empty 'Add by Morgan 2010/1/8
   
   For nIndex = 0 To TF_NP - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
   'Add by Morgan 2010/3/5
   lblNP23.Visible = False
   textNP23.Visible = False
   'Add By Sindy 2021/4/23 + 開放FCP,FG
   If m_NP02 = "FCP" Or m_NP02 = "FG" Then
      lblNP23.Visible = True
      textNP23.Visible = True
   End If
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textNP01.Locked = bEnable
   textNP06.Locked = bEnable
   textNP07.Locked = bEnable
   textNP08.Locked = bEnable
   textNP09.Locked = bEnable
   textNP10.Locked = bEnable
   textNP11.Locked = bEnable
   textNP12.Locked = bEnable
   textNP13.Locked = bEnable
   textNP14.Locked = bEnable
   textNP15.Locked = bEnable
   textNP23.Locked = bEnable 'Add by Morgan 2010/1/8
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textNP01.Locked = bEnable
End Sub

' 讀取商標基本檔
Private Function QueryTradeMark(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryTradeMark = False
   strSql = "SELECT * FROM TRADEMARK " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryTradeMark = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem "中 : " & rsTmp.Fields("TM05")
      End If
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem "英 : " & rsTmp.Fields("TM06")
      End If
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem "日 : " & rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_Nation = rsTmp.Fields("TM10")
      End If
   'Else
   '   textTM05 = "此筆資料不存在於商標基本檔"
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取服務業務基本檔
Private Function QueryServicePractice(ByVal strSP01 As String, ByVal strSP02 As String, ByVal strSP03 As String, ByVal strSP04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryServicePractice = False
   strSql = "SELECT * FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & strSP01 & "' AND " & _
                  "SP02 = '" & strSP02 & "' AND " & _
                  "SP03 = '" & strSP03 & "' AND " & _
                  "SP04 = '" & strSP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryServicePractice = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem "中 : " & rsTmp.Fields("SP05")
      End If
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem "英 : " & rsTmp.Fields("SP06")
      End If
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem "日 : " & rsTmp.Fields("SP07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_Nation = rsTmp.Fields("SP09")
      End If
   'Else
   '   textTM05 = "此筆資料不存在於服務業務基本檔"
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取專利基本檔
Private Function QueryPatent(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   QueryPatent = False
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & strPA01 & "' AND " & _
                  "PA02 = '" & strPA02 & "' AND " & _
                  "PA03 = '" & strPA03 & "' AND " & _
                  "PA04 = '" & strPA04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryPatent = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("PA05")) = False Then
         cmbTM05.AddItem "中 : " & rsTmp.Fields("PA05")
      End If
      If IsNull(rsTmp.Fields("PA06")) = False Then
         cmbTM05.AddItem "英 : " & rsTmp.Fields("PA06")
      End If
      If IsNull(rsTmp.Fields("PA07")) = False Then
         cmbTM05.AddItem "日 : " & rsTmp.Fields("PA07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("PA26")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("PA26"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("PA09")) = False Then
         m_Nation = rsTmp.Fields("PA09")
      End If
   'Else
   '   textTM05 = "此筆資料不存在於專利基本檔"
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取法務基本檔
Private Function QueryLawCase(ByVal strLC01 As String, ByVal strLC02 As String, ByVal strLC03 As String, ByVal strLC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryLawCase = False
   strSql = "SELECT * FROM LAWCASE " & _
            "WHERE LC01 = '" & strLC01 & "' AND " & _
                  "LC02 = '" & strLC02 & "' AND " & _
                  "LC03 = '" & strLC03 & "' AND " & _
                  "LC04 = '" & strLC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryLawCase = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("LC05")) = False Then
         cmbTM05.AddItem "中 : " & rsTmp.Fields("LC05")
      End If
      If IsNull(rsTmp.Fields("LC06")) = False Then
         cmbTM05.AddItem "英 : " & rsTmp.Fields("LC06")
      End If
      If IsNull(rsTmp.Fields("LC07")) = False Then
         cmbTM05.AddItem "日 : " & rsTmp.Fields("LC07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("LC11")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("LC11"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("LC15")) = False Then
         m_Nation = rsTmp.Fields("LC15")
      End If
   'Else
   '   textTM05 = "此筆資料不存在於法務基本檔"
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取顧問案件基本檔
Private Function QueryHireCase(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryHireCase = False
   strSql = "SELECT * FROM HIRECASE " & _
            "WHERE HC01 = '" & strHC01 & "' AND " & _
                  "HC02 = '" & strHC02 & "' AND " & _
                  "HC03 = '" & strHC03 & "' AND " & _
                  "HC04 = '" & strHC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryHireCase = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("HC06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("HC06")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("HC05")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("HC05"), 0)
      End If
   'Else
   '   textTM05 = "此筆資料不存在於顧問案件基本檔"
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function QueryCaseProgress(ByVal strCP09 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryCaseProgress = False
'Modify by Morgan 2005/3/16 只要總收文號就好
'   StrSql = "SELECT * FROM CASEPROGRESS " & _
'            "WHERE CP09 = '" & strCP09 & "' AND " & _
'                  "CP01 = '" & m_NP02 & "' AND " & _
'                  "CP02 = '" & m_NP03 & "' AND " & _
'                  "CP03 = '" & m_NP04 & "' AND " & _
'                  "CP04 = '" & m_NP05 & "' "
   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & strCP09 & "'"
'2005/3/16 end
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      QueryCaseProgress = True
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' 承辦人
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 發文日
      If IsNull(rsTmp.Fields("CP27")) = False Then
         textCP27 = TAIWANDATE(rsTmp.Fields("CP27"))
      End If
      'Add By Sindy 2014/9/11
      '智權人員
      m_CP13 = ""
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
      End If
      '2014/9/11 END
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
 
Private Sub QueryNextProgress()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   If m_DataListCount <= 0 Then
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM NEXTPROGRESS,CASEPROGRESS " & _
            "WHERE NP02 = '" & m_NP02 & "' AND " & _
                  "NP03 = '" & m_NP03 & "' AND " & _
                  "NP04 = '" & m_NP04 & "' AND " & _
                  "NP05 = '" & m_NP05 & "' AND " & _
                  "NP22 = " & m_DataList(m_CurrDL).diNP22 & " AND CP09(+)=NP01"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   textNP10.Tag = "" 'Add By Sindy 2014/9/11
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("NP01")) = False Then: textNP01 = rsTmp.Fields("NP01")
      If IsNull(rsTmp.Fields("NP06")) = False Then: textNP06 = rsTmp.Fields("NP06")
      textNP06.Tag = textNP06 'Add by Amy 2022/06/20
      If IsNull(rsTmp.Fields("NP07")) = False Then: textNP07 = rsTmp.Fields("NP07")
      If IsNull(rsTmp.Fields("NP08")) = False Then: textNP08 = TAIWANDATE(rsTmp.Fields("NP08"))
      If IsNull(rsTmp.Fields("NP09")) = False Then: textNP09 = TAIWANDATE(rsTmp.Fields("NP09"))
      textNP08.Tag = textNP08 'Added by Morgan 2014/11/19
      textNP09.Tag = textNP09 'Added by Morgan 2014/11/19
      'Modify By Sindy 2014/9/11 +textNP10.Tag = rsTmp.Fields("NP10")
      If IsNull(rsTmp.Fields("NP10")) = False Then: textNP10 = rsTmp.Fields("NP10"): textNP10.Tag = rsTmp.Fields("NP10")
      If IsNull(rsTmp.Fields("NP11")) = False Then: textNP11 = TAIWANDATE(rsTmp.Fields("NP11"))
      If IsNull(rsTmp.Fields("NP12")) = False Then: textNP12 = rsTmp.Fields("NP12")
      If IsNull(rsTmp.Fields("NP24")) = False Then: textNP24 = rsTmp.Fields("NP24") 'Add By Sindy 2018/10/17
      If IsNull(rsTmp.Fields("NP13")) = False Then: textNP13 = rsTmp.Fields("NP13")
      If IsNull(rsTmp.Fields("NP14")) = False Then: textNP14 = rsTmp.Fields("NP14")
      If IsNull(rsTmp.Fields("NP15")) = False Then: textNP15 = rsTmp.Fields("NP15")
      If IsNull(rsTmp.Fields("NP22")) = False Then: textNP22 = rsTmp.Fields("NP22")
      'Add by Morgan 2010/1/8
      textNP23 = TransDate("" & rsTmp.Fields("NP23"), 1)
      'Modify By Sindy 2021/4/23 + 開放FCP,FG
      If ((rsTmp.Fields("CP01") = "P" Or rsTmp.Fields("CP01") = "PS") _
         And Left(rsTmp.Fields("CP12"), 1) = "F") Or _
         rsTmp.Fields("CP01") = "FCP" Or _
         rsTmp.Fields("CP01") = "FG" Then
         lblNP23.Visible = True
         textNP23.Visible = True
      'Added by Lydia 2025/10/29 開放內專P案的約定期限
      ElseIf (rsTmp.Fields("CP01") = "P" Or rsTmp.Fields("CP01") = "PS") And strSrvDate(1) >= 內專本所約定期限啟用日 Then
         lblNP23.Visible = True
         textNP23.Visible = True
      'end 2025/10/29
      Else
         lblNP23.Visible = False
         textNP23.Visible = False
      End If
      'END 2010/1/8
      
      ' 更新CreateID及UpdateID
      UpdateCUID rsTmp
      ' 更新欄位的內容
      UpdateFieldOldData rsTmp
      
   End If
   rsTmp.Close
   Set rsTmp = Nothing
EXITSUB:
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim bQuery As Boolean
   
   'Add By Cheng 2002/07/17
   m_Nation = ""
   
   ' 依本所案號讀取基本檔案
   Select Case m_NP02
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         QueryTradeMark m_NP02, m_NP03, m_NP04, m_NP05
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         QueryPatent m_NP02, m_NP03, m_NP04, m_NP05
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/24 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         QueryLawCase m_NP02, m_NP03, m_NP04, m_NP05
      ' 讀取顧問案件基本檔
      Case "LA":
         QueryHireCase m_NP02, m_NP03, m_NP04, m_NP05
      ' 讀取服務業務基本檔
      Case Else:
         QueryServicePractice m_NP02, m_NP03, m_NP04, m_NP05
   End Select
   ' 讀取案件進度檔
   QueryNextProgress
   ' 讀取下一程序檔
   bQuery = QueryCaseProgress(textNP01)
   
   textNP07_Validate False
   textNP10_Validate False
   textNP12_Validate False
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
   If IsNull(rsSrcTmp.Fields("NP16")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NP16")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("NP16"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NP17")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NP17")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("NP17"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NP18")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NP18")) = False Then
         strTemp = rsSrcTmp.Fields("NP18")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NP19")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NP19")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("NP19"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NP20")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NP20")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("NP20"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NP21")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NP21")) = False Then
         strTemp = rsSrcTmp.Fields("NP21")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " : " & strCDate & " " & _
              " : " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " : " & strUDate & " " & _
              " : " & strUTime
              
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strNP22 As String)
   Dim strTemp As String
   Dim nIndex As Integer
   
   If IsRecordExist(strNP22) = True Then
      For nIndex = 0 To m_DataListCount - 1
         If m_DataList(nIndex).diNP22 = strNP22 Then
            m_CurrDL = nIndex
            Exit For
         End If
      Next nIndex
   Else
      m_CurrDL = 0
      strTemp = Empty
      For nIndex = 0 To m_DataListCount - 1
         If strNP22 > strTemp Then
            m_CurrDL = nIndex
            Exit For
         End If
         strTemp = m_DataList(nIndex).diNP22
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
         'Added by Lydia 2018/01/10 非編輯狀態可以列印
         cmdPrint.Enabled = True
         cmdPrintReturnSheet.Enabled = True
         'end 2018/01/10
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
         'Added by Lydia 2018/01/10 編輯狀態不可列印
         cmdPrint.Enabled = False
         cmdPrintReturnSheet.Enabled = False
         'end 2018/01/10
   End Select
   
End Sub


Private Sub textNP01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 總收文號
Private Sub textNP01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textNP01) = False Then
      Select Case m_EditMode
         Case 1, 2, 4:
            Select Case Mid(textNP01, 1, 1)
               'Modified by Lydia 2016/12/26 + D類收文
               Case "A", "B", "C", "D":
               Case Else:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "收文號不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textNP01_GotFocus
                  GoTo EXITSUB
            End Select
         Case Else:
      End Select
      Select Case m_EditMode
         Case 1:
            textCP05 = Empty
            textCP14 = Empty
            textCP27 = Empty
            If QueryCaseProgress(textNP01) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "該筆收文資料不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNP01_GotFocus
            End If
      End Select
   End If
EXITSUB:
End Sub

Private Sub textNP06_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否續辦
Private Sub textNP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textNP06) = False Then
      Select Case textNP06
         Case "Y", "N":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "是否續辦欄位只可輸入Y或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNP06_GotFocus
      End Select
   End If
End Sub

' 案件性質
Private Sub textNP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textNP07_2 = Empty
   If IsEmptyText(textNP07) = False Then
      If m_Nation < "010" Then
         textNP07_2 = GetCaseTypeName(m_NP02, textNP07, 0)
      Else
         textNP07_2 = GetCaseTypeName(m_NP02, textNP07, 1)
      End If
      If IsEmptyText(textNP07_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP07_GotFocus
      End If
      
      'Added by Morgan 2022/10/25
      If Cancel = False And m_EditMode = 1 And textNP02 = "FCP" And (textNP07 = "205" Or textNP07 = "107") Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請從機關來函輸入產生期限"
         nResponse = MsgBox(strMsg, vbExclamation, strTit)
         textNP07_GotFocus
      End If
      'end 2022/10/25
   End If
   
   'Added by Morgan 2012/8/20
   If Cancel = False Then
      If textNP02 = "CFP" Then
         'Modified by Morgan 2021/9/2 改以公用常數判斷
         'If textNP07 = "107" Then
         If (Len(textNP07) = 3 And InStr(CFPAppDatePtyList, textNP07) > 0) Then
         'end 2021/9/2
            lblNP23.Visible = True
            textNP23.Visible = True
         Else
            lblNP23.Visible = False
            textNP23.Visible = False
         End If
      End If
   End If
   
End Sub

' 本所期限
Private Sub textNP08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If m_EditMode = 1 Or m_EditMode = 2 Then  'Added by Lydia 2020/07/13 新增或修改
        If IsEmptyText(textNP08) = False Then
           If CheckIsTaiwanDate(textNP08, False) = False Then
              Cancel = True
              strTit = "檢核資料"
              strMsg = "本所期限日期格式不正確"
              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              textNP08_GotFocus
           'Add By Sindy 2012/8/20
           'modify by sonia 2019/8/7 +ACS
           'Modified by Lydia 2020/07/07 本所期限檢查：所有系統類別的本所期限都要控制是工作日
           'ElseIf ChkWorkDay(DBDATE(textNP08)) = False And (textNP02 = "P" Or textNP02 = "PS" Or textNP02 = "CFP" Or textNP02 = "CPS" Or textNP02 = "ACS") Then
           'Modified by Lydia 2022/05/27 已處理不受限制+ Trim(textNP06) = ""
           ElseIf ChkWorkDay(DBDATE(textNP08)) = False And Trim(textNP06) = "" Then
              'Added by Lydia 2022/05/27 系統類別FCT、CFT、CFC、S的案件改為「本所期限必須為工作天，將自動改為前一工作天! 」
              If InStr("FCT,CFT,CFC,S,", textNP02 & ",") > 0 Then
                 strTit = CompWorkDay(1, DBDATE(textNP08), 1)
                 textNP08 = TransDate(strTit, 1)
                 strTit = "檢核資料"
                 strMsg = "本所期限必須為工作天，將自動改為前一工作天! "
                 nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              Else
              'end 2022/05/27
                 Cancel = True
                 strTit = "檢核資料"
                 strMsg = "本所期限必須為工作天"
                 nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                 textNP08_GotFocus
              End If 'Added by Lydia 2022/05/27
           '2012/8/20 End
           End If
        End If
   End If 'Added by Lydia 2020/07/13
End Sub

' 法定期限
Private Sub textNP09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textNP09) = False Then
      If CheckIsTaiwanDate(textNP09, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "法定期限日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP09_GotFocus
      End If
   End If
   'Added by Morgan 2014/11/11
   If Cancel = False Then
      'Modified by Morgan 2014/11/18 +P,PS;另FCP改為年費預設-2個日曆天,實審預設-4個日曆天,其他不預設
      'Modify By Sindy 2021/8/9 + Or textNP02 = "FG"
      If m_Nation = "000" And (textNP02 = "FCP" Or textNP02 = "FG" Or textNP02 = "P" Or textNP02 = "PS") _
         And textNP09 <> "" Then
         
         If textNP08 = "" Then
            'Modify By Sindy 2021/8/9 + And (textNP07 = "605" Or textNP07 = "416")
            If textNP02 = "FCP" And (textNP07 = "605" Or textNP07 = "416") Then
               If textNP07 = "605" Then
                  'textNP08 = TransDate(CompDate(2, -2, textNP09), 1)
                  'Add by Sindy 2021/8/2 改計算天數及加約定期限
                  textNP08 = TransDate(PUB_GetFCPOurDeadline(textNP09, -2, , strExc(0)), 1) '均為法定期限前2個工作天
                  textNP23 = TransDate(strExc(0), 1)
               ElseIf textNP07 = "416" Then
                  'textNP08 = TransDate(CompDate(2, -4, textNP09), 1)
                  'Add by Sindy 2021/8/2 改計算天數及加約定期限
                  textNP08 = TransDate(PUB_GetFCPOurDeadline(textNP09, -4, , strExc(0)), 1) '均為法定期限前2個工作天
                  textNP23 = TransDate(strExc(0), 1)
               End If
            'Modify By Sindy 2021/8/9
            ElseIf textNP02 = "FCP" Or textNP02 = "FG" Then
               nResponse = MsgBox("是: 約定期限為本所期限前4工作天" & vbCrLf & vbCrLf & _
                                  "否: 約定期限為本所期限前2工作天", vbYesNo + vbDefaultButton1)
               If nResponse = vbYes Then
                  'Modify By Sindy 2025/2/12 傳入本所案號,總收文號
                  textNP08 = TransDate(PUB_GetFCPOurDeadline(textNP09, -4, , strExc(0), , , textNP02, textNP03, textNP04, textNP05, textNP01), 1)
                  textNP23 = TransDate(strExc(0), 1)
               Else
                  'Modify By Sindy 2025/2/12 傳入本所案號,總收文號
                  textNP08 = TransDate(PUB_GetFCPOurDeadline(textNP09, -2, , strExc(0), , , textNP02, textNP03, textNP04, textNP05, textNP01), 1)
                  textNP23 = TransDate(strExc(0), 1)
               End If
            '2021/8/9 END
            Else
               textNP08 = TransDate(PUB_GetOurDeadline(textNP09), 1)
            End If
            
            textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         End If
      End If
   End If
   'end 2014/11/11
End Sub

'Add By Sindy 2010/11/26
Private Sub textNP10_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 解除期限日期
Private Sub textNP11_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textNP11) = False Then
      If CheckIsTaiwanDate(textNP11, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "解除期限日期日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP11_GotFocus
      End If
   End If
End Sub

' 智權人員代號
Private Sub textNP10_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   textNP10_2 = Empty
   If IsEmptyText(textNP10) = False Then
      textNP10_2 = GetStaffName(textNP10, True)
      If IsEmptyText(textNP10_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "智權人員代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP10_GotFocus
      End If
   End If
End Sub

' 解除期限原因
Private Sub textNP12_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textNP12_2 = Empty
   If IsEmptyText(textNP12) = False Then
      strSql = "SELECT * FROM REASONOFRELIEF " & _
               "WHERE ROR01 = '" & textNP12 & "' "
      Set rsTmp = New ADODB.Recordset
      ' 讀取資料庫
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      ' 檢查讀取的資料筆數
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("ROR02")) = False Then
            textNP12_2 = rsTmp.Fields("ROR02")
         End If
      Else
         Cancel = True
         strTit = "檢核資料"
         strMsg = "解除期限原因代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP12_GotFocus
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
End Sub

' 機關文號
Private Sub textNP13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modify by Morgan 2011/1/3 機關文號欄位改長度(百年問題)改抓MaxLength屬性控制
   If CheckLengthIsOK(textNP13, textNP13.MaxLength) = False Then
      Cancel = True
      textNP13_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textNP13.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 相關人
Private Sub textNP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textNP14, 60) = False Then
      Cancel = True
      textNP14_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textNP14.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 備註
Private Sub textNP15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textNP15, 2000) = False Then
      Cancel = True
      textNP15_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textNP15.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   'Memo by Lydia 2021/10/20 原程式搬到Form_KeyUp

   Call PUB_SaveMeTrackMode(m_MeTrackMode, 0, KeyCode)  'Added by Lydia 2021/10/20 Form2.0 記錄鍵盤傳入順序
   
End Sub

'Added by Lydia 2021/10/20
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Call PUB_SaveMeTrackMode(m_MeTrackMode, 1, KeyCode)  'Added by Lydia 2021/10/20 Form2.0 記錄鍵盤傳入順序
    
'Memo by Lydia 2021/10/20 從Form_KeyDown搬來
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
      'Remove by Lydia 2021/11/22 取消以ENTER控制為換行的功能 (Form2.0修改之維護資料功能Toolbar之修改統一)
      'Case vbKeyReturn:
      '   If m_EditMode <> 0 Then
      '      OnAction vbKeyF9
      '   End If
      'end 2021/11/22
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
   
   EnableTextBox textNP22, False
   
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
         'Added by Lydia 2022/09/13 外商承辦F11判斷人員職稱等級決定是否鎖住「是否續辦」
         If bolUpdNp06 = False Then
             textNP06.Locked = True
         End If
         'end 2022/09/13
         UpdateToolbarState
         SetInputEntry
         Me.textNP06.Tag = Me.textNP06.Text 'Add by Amy 2022/06/20
         'Add By Cheng 2002/01/14
         Me.textNP08.Tag = Me.textNP08.Text
         Me.textNP09.Tag = Me.textNP09.Text
         Me.textNP23.Tag = Me.textNP23.Text 'Add by Morgan 2010/1/8
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
         EnableTextBox textNP22, True
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
         'Added by Lydia 2021/10/20 Form2.0 控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkMeTrackMode(m_MeTrackMode) = False Then
             Exit Sub
         End If
         'end 2021/10/20
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
'         UpdateFieldNewData
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
         frm075007_1.ClearRemark
         frm075007_1.Show
         'Modify By Cheng 2002/01/08
'         If m_AddData = True Then: frm075007_1.RefreshList
'         frm075007_1.RefreshList
        'Add By Cheng 2002/12/11
        frm075007_1.textNP03.SetFocus
        TextInverse frm075007_1.textNP03
   End Select
EXITSUB:
End Sub

Private Sub textNP23_GotFocus()
   TextInverse textNP23
   CloseIme
End Sub

'Add by Morgan 2010/1/8
Private Sub textNP23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(textNP23) = False Then
      If InStr(",997,998,995,996,999,411,1204,1503,", "," & textNP07 & ",") > 0 Then
         Cancel = True
         MsgBox "本案件性質不可輸入約定期限！"
         textNP23_GotFocus
      ElseIf CheckIsTaiwanDate(textNP23, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "本所期限日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP23_GotFocus
      End If
   End If
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
Private Function IsDataBaseExist(ByVal strNP22) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   IsDataBaseExist = False
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & m_NP02 & "' AND " & _
                  "NP03 = '" & m_NP03 & "' AND " & _
                  "NP04 = '" & m_NP04 & "' AND " & _
                  "NP05 = '" & m_NP05 & "' AND " & _
                  "NP22 = '" & strNP22 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
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
Private Function IsRecordExist(ByVal strNP22) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   IsRecordExist = False
   bFind = False
   For nIndex = 0 To m_DataListCount - 1
      If m_DataList(nIndex).diNP22 = strNP22 Then
         bFind = True
      End If
   Next nIndex
   If bFind = False Then
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & m_NP02 & "' AND " & _
                  "NP03 = '" & m_NP03 & "' AND " & _
                  "NP04 = '" & m_NP04 & "' AND " & _
                  "NP05 = '" & m_NP05 & "' AND " & _
                  "NP22 = '" & strNP22 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
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
'Private Sub AddRecord()
Private Function AddRecord() As Boolean
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strNP22 As String
   
   'add by nickc 2006/06/08
   AddRecord = False
   
   For nIndex = 0 To TF_NP - 1
      If m_FieldList(nIndex).fiName = "NP22" Then
         strNP22 = m_FieldList(nIndex).fiNewData
         Exit For
      End If
   Next nIndex

   bFirst = True
   bDifference = False
   strSql = "INSERT INTO NEXTPROGRESS ("
   For nIndex = 0 To TF_NP - 1
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
   For nIndex = 0 To TF_NP - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            ' 91.04.04 modify by louis (修改單引號)
            'strTmp = "'" & m_FieldList(nIndex).fiNewData & "'"
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
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
      'add by nickc 2006/03/16 紀錄分析語法
      On Error GoTo oErr
      cnnConnection.BeginTrans
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      'add by nickc 2006/06/07 加入 trans
      cnnConnection.CommitTrans
      AddRecord = True
      ' 將新增的資料加入到串列中
      SetDataListItem strNP22
      ' 顯示該筆記錄
      ShowCurrRecord strNP22
      ' 通知前畫面有新增的記錄
      frm075007_1.ModRecord strNP22
      ' 設定有新增資料
      m_AddData = True
   End If
EXITSUB:
'add by nickc 2006/06/07
Exit Function
oErr:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
End Function

' 修改記錄
'edit by nickc 2006/06/07
'Private Sub ModRecord()
Private Function ModRecord() As Boolean
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strNP22 As String
   
   'add by nickc 2006/06/08
   ModRecord = False
   
   strNP22 = m_DataList(m_CurrDL).diNP22
   '910910  nick tigger
   '***** start
   'strSQL = "UPDATE NEXTPROGRESS SET "
   'edit by nickc 2006/06/07 紀錄 log
   'StrSql = "begin user_data.user_enabled:=1; UPDATE NEXTPROGRESS SET "
   strSql = " UPDATE NEXTPROGRESS SET "
   '***** end
   bFirst = True
   bDifference = False
   For nIndex = 0 To TF_NP - 1
      strTmp = Empty
      '92.05.22 nick 跳過create & update 相關項目
      If nIndex < 15 Or nIndex > 20 Then
        If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
           If m_FieldList(nIndex).fiType = 0 Then
              If m_FieldList(nIndex).fiNewData = Empty Then
                 strTmp = m_FieldList(nIndex).fiName & " = NULL "
              Else
                 ' 91.04.04 modify by louis (修改單引號)
                 'strTmp = m_FieldList(nIndex).fiName & " = '" & m_FieldList(nIndex).fiNewData & "'"
                 strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
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
      End If
   Next nIndex
    '910910 nick tigger
   '***** start
   'strSQL = strSQL & " " & _
            "WHERE NP02 = '" & m_NP02 & "' AND " & _
                  "NP03 = '" & m_NP03 & "' AND " & _
                  "NP04 = '" & m_NP04 & "' AND " & _
                  "NP05 = '" & m_NP05 & "' AND " & _
                  "NP22 = " & strNP22 & " "
   'edit by nickc 2006/06/07 紀錄 log
   'StrSql = StrSql & " " & _
            "WHERE NP02 = '" & m_NP02 & "' AND " & _
                  "NP03 = '" & m_NP03 & "' AND " & _
                  "NP04 = '" & m_NP04 & "' AND " & _
                  "NP05 = '" & m_NP05 & "' AND " & _
                  "NP22 = " & strNP22 & ";end; "
   strSql = strSql & " " & _
            "WHERE NP02 = '" & m_NP02 & "' AND " & _
                  "NP03 = '" & m_NP03 & "' AND " & _
                  "NP04 = '" & m_NP04 & "' AND " & _
                  "NP05 = '" & m_NP05 & "' AND " & _
                  "NP22 = " & strNP22 & " "
    
    '***** end
'910910 nick tigger
'***** start
On Error GoTo ErrHand
'***** end
                  
   If bDifference = True Then
      '910910 nick tigger
      '**** start
      cnnConnection.BeginTrans
      '***** end
      'add by nickc 2006/03/16 紀錄分析語法
      Pub_SeekTbLog strSql
      'edit by nickc 2006/06/07  紀錄 log
      'cnnConnection.Execute StrSql
      cnnConnection.Execute "begin user_data.user_enabled:=1;" & strSql & "; end;"
      '910910 nick tigger
      '***** start
      'Add by Amy 2022/06/20 商標延展將續辦由N->null,則刪除 T102Inform 且未取消延展的資料(續辦可能N->Y)
      'Modify by Amy 2022/06/24 拿掉And ti06 is null,避免當天無法重做
      If ((textNP02 = "T" Or textNP02 = "TF") And (textNP07 = "102" Or textNP07 = "109" Or textNP07 = "716")) _
        And textNP06.Tag <> textNP06 And textNP06.Tag = "N" And textNP06 = MsgText(601) Then
         strExc(0) = "Select * From T102Inform Where ti02='" & textNP01 & "' And ti04='" & textNP22 & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "Delete T102Inform Where  ti02='" & textNP01 & "' And ti04='" & textNP22 & "' "
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2022/06/20
      'Added by Morgan 2012/8/20
      If textNP23.Tag <> textNP23 Then
         strExc(0) = "select cp64 from caseprogress where cp09='" & textNP01 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            intI = InStr("" & RsTemp(0), "約定期限：" & textNP23.Tag)
            If intI > 0 Then
               strExc(1) = Replace(RsTemp(0), "約定期限：" & textNP23.Tag, "約定期限：" & textNP23)
            Else
               strExc(1) = "約定期限：" & textNP23 & ";" & RsTemp(0)
            End If
            strSql = "update caseprogress set cp64='" & ChgSQL(strExc(1)) & "' where cp09='" & textNP01 & "'"
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2012/8/20
      cnnConnection.CommitTrans
      
      'Added by Lydia 2024/11/08 內專內商人員(ST03=P22,P12)在進度檔、下一程序修改本所期限和法定期限時，Email通知主管
      If InStr("P22,P12", Pub_StrUserSt03) > 0 And (Me.textNP08.Tag <> Me.textNP08.Text Or Me.textNP09.Tag <> Me.textNP09.Text) Then
         'Modified by Lydia 2025/05/19 在系統特殊設定區分要不要發通知; ex. 員工編號(N)=>不寄, 員工編號(Y)=>要寄
         'strExc(0) = "select s1.st01,s1.st02,s1.st93 from setspecman,staff s1 where ocode='期限修改郵件收受者' and instr(oman,s1.st01) > 0 and s1.st93='" & Pub_StrUserSt93 & "' "
         strExc(0) = "select s1.st01,s1.st02,s1.st93,substr(oman,instr(oman,s1.st01)+length(s1.st01),3) o1 from setspecman,staff s1 where ocode='期限修改郵件收受者' and instr(oman,s1.st01) > 0 and s1.st93='" & Pub_StrUserSt93 & "' "
         intI = 1
         strExc(1) = ""
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modified by Lydia 2025/07/04 debug: 本人操作才不用發Email; or "" & RsTemp.Fields("st01") <> strUserNum
            If "" & RsTemp.Fields("O1") <> "(N)" Or "" & RsTemp.Fields("st01") <> strUserNum Then    'Added by Lydia 2025/05/19 在系統特殊設定區分要不要發通知
               strExc(1) = "" & RsTemp.Fields("st01")
            End If
         Else
            strExc(1) = Pub_GetSpecMan("程式管理人員")
         End If
         If strExc(1) <> "" Then
            '主旨：P/CFP-XXXXXX「XXXXXXX」(案件性質)-->進度檔/下一程序之期限有更動！ [（更動之程序人員）]
            strExc(2) = m_NP02 & "-" & m_NP03 & IIf(m_NP04 <> "0", "-" & m_NP04, "") & IIf(m_NP05 <> "00", "-" & m_NP05, "") & "「" & Trim(textNP07_2) & "」-->下一程序之期限有更動！ [" & strUserName & "]"
            strExc(3) = ""
            If Me.textNP08.Tag <> Me.textNP08.Text Then
               strExc(3) = strExc(3) & "修改前本所期限：" & ChangeWStringToTDateString(DBDATE(textNP08.Tag)) & vbCrLf & _
                           "修改後本所期限：" & ChangeWStringToTDateString(DBDATE(textNP08.Text)) & vbCrLf & vbCrLf
            End If
            If Me.textNP09.Tag <> Me.textNP09.Text Then
               strExc(3) = strExc(3) & "修改前法定期限：" & ChangeWStringToTDateString(DBDATE(textNP09.Tag)) & vbCrLf & _
                           "修改後法定期限：" & ChangeWStringToTDateString(DBDATE(textNP09.Text)) & vbCrLf & vbCrLf
            End If
            PUB_SendMail strUserNum, strExc(1), "", strExc(2), strExc(3)
         End If
      End If
      'end 2024/11/08
      
   End If
      'add by nickc 2006/06/07
      ModRecord = True
      '***** end
      ShowCurrRecord strNP22
      ' 通知前畫面有更新的記錄
      frm075007_1.ModRecord strNP22
'910910 nick tigger
'***** start
   Exit Function
ErrHand:
    MsgBox (Err.Description)
    cnnConnection.RollbackTrans
'******* end
End Function

' 刪除記錄
'edit by nickc 2006/06/07
'Private Sub DelRecord()
Private Function DelRecord() As Boolean
   Dim strSql As String
   Dim strNP22 As String
   Dim strDataList() As String
   Dim nDataListCount As Integer
   Dim nIndex As Integer
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nPos As Integer

   'add by nickc 2006/06/08
   DelRecord = False
   
   nPos = m_CurrDL
   strNP22 = m_DataList(m_CurrDL).diNP22

   If OnDataDeleteRecord(2, strNP22, m_NP02 & m_NP03 & m_NP04 & m_NP05) <> 0 Then
      GoTo EXITSUB
   End If
   
   strSql = "DELETE FROM NEXTPROGRESS " & _
            "WHERE NP02 = '" & m_NP02 & "' AND " & _
                  "NP03 = '" & m_NP03 & "' AND " & _
                  "NP04 = '" & m_NP04 & "' AND " & _
                  "NP05 = '" & m_NP05 & "' AND " & _
                  "NP22 = " & strNP22 & " "
   'add by nickc 2006/03/16 紀錄分析語法
   On Error GoTo pErr
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   
   cnnConnection.Execute strSql
   'add by nickc 2006/06/07 加入 trans
   cnnConnection.CommitTrans
   DelRecord = True
   
   ' 通知前畫面有刪除的記錄
   frm075007_1.DelRecord strNP22

   ' 刪除記錄時, 除了要刪除資料庫中的記錄還要刪除在本程式模組中所記錄的資料
   ' 以供瀏覽資料時會更新資料串列的順序及正確性
   nDataListCount = 0
   For nIndex = 0 To m_DataListCount - 1
      If m_DataList(nIndex).diNP22 <> strNP22 Then
         ReDim Preserve strDataList(nDataListCount + 1)
         strDataList(nDataListCount) = m_DataList(nIndex).diNP22
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
      'strTit = "資料顯示"
      'strMsg = "該筆本所案號已無資料"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Unload Me
      frm075007_1.ClearRemark
      frm075007_1.Show
   Else
      'ShowCurrRecord strNP22
      If nPos <= m_DataListCount - 1 Then
         m_CurrDL = nPos
      Else
         m_CurrDL = m_DataListCount - 1
      End If
      UpdateCtrlData
   End If
EXITSUB:
'add by nickc 2006/06/07
Exit Function
pErr:
    cnnConnection.RollbackTrans
    MsgBox Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strNP22 As String
   Dim strFirst As String
   Dim bFirst As Boolean
   Dim nIndex As Integer
   Dim bFind As Boolean
   Dim nPos As Integer
   
   QueryRecord = False
   strSql = "SELECT * FROM NEXTPROGRESS " & _
            "WHERE NP02 = '" & m_NP02 & "' AND " & _
                  "NP03 = '" & m_NP03 & "' AND " & _
                  "NP04 = '" & m_NP04 & "' AND " & _
                  "NP05 = '" & m_NP05 & "' "
   If IsEmptyText(textNP01) = False Then
      strSql = strSql & " AND NP01 = '" & textNP01 & "' "
   End If
   If IsEmptyText(textNP07) = False Then
      strSql = strSql & " AND NP07 = " & textNP07 & " "
   End If
   If IsEmptyText(textNP22) = False Then
      strSql = strSql & " AND NP22 = " & textNP22 & " "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      bFirst = True
      Do While rsTmp.EOF = False
         If IsNull(rsTmp.Fields("NP22")) = False Then
            strNP22 = rsTmp.Fields("NP22")
            SetDataListItem strNP22
            If bFirst = True Then
               strFirst = rsTmp.Fields("NP22")
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
   
   QueryRecord = True
   ShowCurrRecord strFirst
   
   UpdateToolbarState
EXITSUB:
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Select Case m_EditMode
      Case 1:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            UpdateFieldNewData
            'edit by nickc 2006/06/07
            'AddRecord
            If AddRecord = False Then GoTo EXITSUB
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            UpdateFieldNewData
            'Add By Cheng 2002/01/14
            '若有改變本所期限或法定期限值, 則顯示是否確定修改的詢息
            If Me.textNP08.Tag <> Me.textNP08.Text Then
               If MsgBox("您確定要更改本所期限???", vbExclamation + vbOKCancel) = vbCancel Then
                  Me.textNP08.SetFocus
                  Exit Sub
               End If
            End If
            
            'Add by Morgan 2010/1/8
            If Me.textNP23.Tag <> Me.textNP23.Text Then
               If MsgBox("您確定要更改約定期限???", vbExclamation + vbOKCancel) = vbCancel Then
                  Me.textNP23.SetFocus
                  Exit Sub
               End If
            End If
            
            If Me.textNP09.Tag <> Me.textNP09.Text Then
               If MsgBox("您確定要更改法定期限???", vbExclamation + vbOKCancel) = vbCancel Then
                  Me.textNP09.SetFocus
                  Exit Sub
               End If
            End If
            'edit by nickc 2006/06/07
            'ModRecord
            If ModRecord = False Then GoTo EXITSUB
         Else
            GoTo EXITSUB
         End If
      Case 3:
         UpdateFieldNewData
         'edit by nickc 2006/06/07
         'DelRecord
         If DelRecord = False Then GoTo EXITSUB
         ' 若已無資料在串列中, 則離開
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
      Case 1: textNP01.SetFocus
      Case 2: textNP07.SetFocus
      Case 4: textNP01.SetFocus
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
         ' 總收文號不可空白
         If IsEmptyText(textNP01) = True Then
            strTit = "查詢資料"
            strMsg = "請輸入總收文號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNP01.SetFocus
            GoTo EXITSUB
         End If
         ' 下一程序代號不可空白
         If IsEmptyText(textNP07) = True Then
            strTit = "查詢資料"
            strMsg = "請輸入下一程序"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNP07.SetFocus
            GoTo EXITSUB
         End If
         ' 智權人員代號不可空白
         If IsEmptyText(textNP10) = True Then
            strTit = "查詢資料"
            strMsg = "請輸入智權人員"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNP10.SetFocus
            GoTo EXITSUB
         End If
         'Add By Sindy 2014/9/11
         'Modified by Lydia 2022/09/21 增加CFC案
         'Modified by Lydia 2022/09/29 改用共用變數判斷
         'If (m_NP02 = "CFT" Or m_NP02 = "CFC") And _
            (textNP07 = "305" Or textNP07 = "997" Or textNP07 = "998" Or textNP07 = "1711" Or textNP07 = "312") And _
            textNP10.Text <> textNP10.Tag Then
         If (m_NP02 = "CFT" Or m_NP02 = "CFC") And textNP10.Text <> textNP10.Tag And InStr(TMnp07NotIn, textNP07) > 0 Then
            'Modified by Lydia 2016/03/11 +案號
            'If GetNP69(textNP10, m_Nation, m_CP13) = False Then
            'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
            If GetNA69(textNP10, m_Nation, m_CP13, , m_NP02, m_NP03, m_NP04, m_NP05) = False Then
               If MsgBox("智權人員非該國的CFT承辦人，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                  textNP10.SetFocus
                  GoTo EXITSUB
               End If
            End If
         End If
         '2014/9/11 END
         ' 本所期限不可空白
         If IsEmptyText(textNP08) = True Then
            strTit = "查詢資料"
            strMsg = "請輸入本所期限"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNP08.SetFocus
            GoTo EXITSUB
         End If
         If Me.textNP09.Tag <> "" Then 'Add By Sindy 2015/8/31 +if 修改時,法定期限若原來沒值,就不必檢查不可空白 ex.P-111299
            ' 法定期限不可空白
            If IsEmptyText(textNP09) = True Then
               strTit = "查詢資料"
               strMsg = "請輸入法定期限"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNP09.SetFocus
               GoTo EXITSUB
            End If
         End If
         ' 本所期限不可超過法定期限
         If IsEmptyText(textNP08) = False And IsEmptyText(textNP09) = False Then
            If Val(DBDATE(textNP08)) > Val(DBDATE(textNP09)) Then
               strTit = "查詢資料"
               strMsg = "本所期限不可超過法定期限"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNP08.SetFocus
               GoTo EXITSUB
            End If
         End If
            
         'Add by Morgan 2010/1/8
         ' 約定期限不可超過本所期限
         If textNP23.Visible And IsEmptyText(textNP08) = False And IsEmptyText(textNP23) = False Then
            If Val(DBDATE(textNP23)) > Val(DBDATE(textNP08)) Then
               strTit = "查詢資料"
               strMsg = "約定期限不可晚於本所期限"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNP23.SetFocus
               GoTo EXITSUB
            End If
         End If
         
         'add by nickc 2007/06/13
         If textNP02 = "TF" And Trim(textNP07) = "102" And Trim(textNP09) <> "" Then
              Dim tmpstrNP08 As String
              tmpstrNP08 = DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(DBDATE(textNP09))))
              'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
              'If DBDATE(textNP08) <> tmpstrNP08 Then
              If DBDATE(textNP08) <> PUB_GetWorkDay1(tmpstrNP08, True) Then
                    strTit = "查詢資料"
                    'Modified by Lydia 2020/07/07
                    'strMsg = "TF 延展 本所期限 應為 法定期限 前一個月"
                    strMsg = "TF 延展 本所期限 應為 法定期限 前一個月(最近的工作天)"
                    nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                    textNP08.SetFocus
                    GoTo EXITSUB
              End If
         End If
         
         'Add By Sindy 2013/5/27
         If Trim(textNP06) = "" And (Val(Trim(textNP11)) > 0 Or Trim(textNP12) <> "") Then
            strTit = "查詢資料"
            strMsg = "是否續辦為空值時，不可有解除期限日期及原因!!!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            If Val(Trim(textNP11)) > 0 Then
               textNP11.SetFocus
            ElseIf Trim(textNP12) <> "" Then
               textNP12.SetFocus
            End If
            GoTo EXITSUB
         End If
         '2013/5/27 End
         
      If Trim(textNP06) = "" Then 'Added by Morgan 2022/4/21 要管制的期限才檢查，因為下面的程式只適用最新的期限，且若已無期限還會有錯 Ex:CFP-008623(AA1044289)延展費 N->Y (秀玲有確認)
         
         '910703 Sieg 411
         'edit by nickc 2006/07/12
         'Dim pA(1 To T_PA) As String
         Dim pa() As String
         ReDim pa(1 To TF_PA) As String
         
         Dim DATE1 As String, DATE2 As String, strTmp1 As String, strTmp As String
         Dim varTmp As Variant
         Dim i As Integer
         pa(1) = textNP02
         pa(2) = textNP03
         pa(3) = textNP04
         If pa(3) = "" Then pa(3) = "0"
         pa(4) = textNP05
         If pa(4) = "" Then pa(4) = "00"
         
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetSystemKind(textNP02, intI) Then
         If ClsPDGetSystemKind(textNP02, intI) Then
            If intI = 1 Then ClsPDReadPatentDatabase pa(), 0, False 'Added by Morgan 2022/12/22
         
            'Added by Morgan 2022/10/25
            '實體審查法定期限檢查
            If intI = 1 And textNP07 = 實體審查 Then
               strExc(1) = "": strExc(2) = ""
               strExc(0) = "select cp10,decode('" & pa(9) & "','000',cpm03,cpm04) cp10N from caseprogress,casepropertymap" & _
                  " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                  " and cp10 like '3%' and cpm01(+)=cp01 and cpm02(+)=cp10 order by cp05 desc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(1) = RsTemp("cp10")
                  strExc(2) = RsTemp("cp10n")
               End If
               'P或台灣案分割/改請提醒自行確認
               If (pa(1) = "P" Or pa(9) = "000") And strExc(1) <> "" Then
                  If MsgBox("本案為「" & strExc(2) & "」案，請自行確認法定期限是否正確！" & vbCrLf & vbCrLf & "是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                     textNP09.SetFocus
                     GoTo EXITSUB
                  End If
               Else
                  PUB_GetExamDate pa(1), pa(2), pa(3), pa(4), textNP01, , , strTmp1, , , , IIf(strExc(1) = "307", True, False)
                  If strTmp1 = "" Then
                     If MsgBox("系統無法推算法定期限，請自行確認是否正確！是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                        textNP09.SetFocus
                        GoTo EXITSUB
                     End If
                  Else
                     strTmp1 = TransDate(strTmp1, 1)
                     If textNP09.Text <> strTmp1 Then
                        If pa(9) = "000" Then
                           MsgBox "法定期限(" & textNP09 & ")與系統推算的日期(" & strTmp1 & ")「不同」！請再確認。", vbExclamation
                           textNP09.SetFocus
                           GoTo EXITSUB
                        Else
                           If MsgBox("法定期限(" & textNP09 & ")與系統推算的日期(" & strTmp1 & ")「不同」！" & vbCrLf & vbCrLf & "是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                              textNP09.SetFocus
                              GoTo EXITSUB
                           End If
                        End If
                     End If
                  End If
               End If
            End If
            'FCP補文件法定期限檢查
            If pa(1) = "FCP" And textNP07 = 補文件 Then
               strTmp1 = "": strTmp = ""
               'Modified by Morgan 2022/10/26 --Sharon
               'If InStr(textNP15, "優先權證明文件正本") > 0 Then
               If InStr(textNP15, "優先權證明") > 0 Then
               'end 2022/10/26
                  strExc(1) = PUB_GetFirstPriDate(pa)
                  If strExc(1) <> "" Then
                     If pa(8) = "3" Then
                        strTmp1 = CompDate(1, 10, strExc(1))
                     Else
                        strTmp1 = CompDate(1, 16, strExc(1))
                     End If
                  End If
                  If strTmp1 = "" Then
                     If MsgBox("系統無法推算法定期限，請自行確認是否正確！是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                        textNP09.SetFocus
                        GoTo EXITSUB
                     End If
                  Else
                     strTmp1 = TransDate(strTmp1, 1)
                     If textNP09.Text <> strTmp1 Then
                        MsgBox "法定期限(" & textNP09 & ")與系統推算的日期(" & strTmp1 & ")「不同」！請再確認。", vbExclamation
                        textNP09.SetFocus
                        GoTo EXITSUB
                     End If
                  End If
               ElseIf InStr(textNP15, "專利申請書") > 0 Then
                  strExc(0) = "select pa10 from patent where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "' and pa10>0"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     pa(10) = RsTemp("pa10")
                  End If
                  If pa(10) <> "" Then
                     strTmp = CompDate(1, 4, pa(10))
                     strTmp1 = CompDate(1, 6, pa(10))
                  End If
                  If strTmp1 = "" Then
                     If MsgBox("系統無法推算法定期限，請自行確認是否正確！是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                        textNP09.SetFocus
                        GoTo EXITSUB
                     End If
                  Else
                     strTmp = TransDate(strTmp, 1)
                     strTmp1 = TransDate(strTmp1, 1)
                     If textNP09.Text <> strTmp And textNP09.Text <> strTmp1 Then
                        MsgBox "法定期限(" & textNP09 & ")與系統推算的日期(" & strTmp & " or " & strTmp1 & " )「不同」！請再確認。", vbExclamation
                        textNP09.SetFocus
                        GoTo EXITSUB
                     End If
                  End If
               End If
               
            End If
            'end 2022/10/25
         
            '2008/11/7 modify by sonia
            'If intI = 1 And textNP07 = 年費  Then
            If intI = 1 And (textNP07 = 年費 Or textNP07 = 維持費 Or textNP07 = 延展費) Then
            '2008/11/7 end
               'edit by nickc 2007/02/07 不用 dll 了
               'If objPublicData.ReadPatentDatabase(pa(), 0) Then
               If ClsPDReadPatentDatabase(pa(), 0) Then
                  'Added by Morgan 2016/11/18
                  If Val(pa(25)) > 0 And DBDATE(textNP09) > DBDATE(pa(25)) Then
                     MsgBox "法定期限不可晚於專用期止日(" & pa(25) & ")。", vbExclamation
                     textNP09.SetFocus
                     GoTo EXITSUB
                  End If
                  'end 2016/11/18
               End If
               
               If GetMoneyDate(Val(pa(8)), pa(9), pa(), DATE1, strTmp1, DATE2) Then
                  strTmp = pa(72)
                  
                  'Added by Morgan 2022/4/21
                  If strTmp = strTmp1 Then
                     MsgBox "本案應已無需繳納" & textNP07_2 & "，請再確認！", vbExclamation
                     GoTo EXITSUB
                  End If
                  'end 2022/4/21
                  
                    'Modify By Cheng 2002/11/14
                    '若基本檔有繳費年度資料, 才要檢查下次繳費日法定期限的正確性
                    'Modified by Morgan 2016/11/18 +要有起算日
                    'If strTmp <> "" Then
                    If strTmp <> "" And DATE1 <> "" Then
                    'end 2016/11/18
                        Do While InStr(strTmp, ",,") > 0
                           strTmp = Replace(strTmp, ",,", ",")
                        Loop
                        varTmp = Split(strTmp, ",")
                        i = UBound(varTmp)
                        varTmp = Split(strTmp1, ",")
                        '2008/11/7 modify by sonia
                        'strTmp = Format(varTmp(I))
                        If textNP07 = 年費 Then
                           If pa(9) = "013" And pa(8) <> "1" Then    '香港U,D為延展費但NP掛年費
                              strTmp = Format(varTmp(i + 1))
                           Else
                              strTmp = Format(varTmp(i))
                           End If
                        'Added by Morgan 2015/5/21
                        '香港發明(標準專利)公布後滿5年至第二階段批准請求前要逐年繳維持費(期限和年費一樣用申請日算)
                        ElseIf pa(9) = "013" And pa(8) = "1" And textNP07 = 維持費 Then
                           strTmp = Format(varTmp(i))
                        'end 2015/5/21
                        Else
                           strTmp = Format(varTmp(i + 1))
                        End If
                        '2008/11/7 end
                        
                        'Modified by Morgan 2017/3/17
                        'Modified by Morgan 2019/12/11 +607
                        'If pa(1) = "CFP" And textNP07 = "605" Then
                        If pa(1) = "CFP" And (textNP07 = "605" Or textNP07 = "607") Then
                        'end 2019/12/11
                            strTmp1 = GetFeeNextDate(DATE1, Val(strTmp), pa(9), pa(8))
                            strTmp1 = TransDate(DBDATE(strTmp1), 1)
                        Else
                        'end 2017/3/17
                        
                           strTmp1 = CompDate(0, Val(strTmp), DATE1)
                           'Add by Morgan 2006/3/10 大陸不必減一天
                           'Modify by Morgan 2006/11/2 非台灣都不用減一天
                           'strTmp1 = TransDate(CompDate(2, -1, strTmp1), 1)
                           If pa(9) = "000" Then
                              strTmp1 = TransDate(CompDate(2, -1, strTmp1), 1)
                           Else
                              strTmp1 = TransDate(strTmp1, 1)
                           End If
                           
                        End If 'Added by Morgan 2017/3/17
                        
                        'Removed by Morgan 2017/3/17 已修法併入上面程式
                        ''Added by Morgan 2013/8/23
                        'If pa(9) = "017" And textNP07 = "605" Then
                        '   If Right(pa(21), 4) < Right(pa(10), 4) Then
                        '      strTmp1 = CompDate(0, 2, strTmp1)
                        '   Else
                        '      strTmp1 = CompDate(0, 1, strTmp1)
                        '   End If
                        '   strTmp1 = TransDate(strTmp1, 1)
                        'End If
                        ''end 2013/8/23
                        'end 2017/3/17
                        
                        '92.1.12 modify by sonia
                        'If textNP09.Text <> strTmp1 Then
                        'Modify by Morgan 2011/4/20 CFP 恢復要管制,若期限不一致時由電腦中心新增以減少錯誤
                        'If textNP09.Text <> strTmp1 And textNP02 <> "CFP" Then
                        If textNP09.Text <> strTmp1 Then
                        '92.1.12 end
                           'Modified by Morgan 2016/1/11 改確認後可存檔(因有新舊法或管制半年...等狀況)
                           'MsgBox "下次繳費日法定期限應為 " & strTmp1, vbCritical
                           If MsgBox("法定期限(" & textNP09 & ")與系統推算的日期(" & strTmp1 & ")不同！是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                              textNP09.SetFocus
                              GoTo EXITSUB
                           End If
                        End If
'Removed by Morgan 2012/10/19 取消才能掛香港維持費
'                    '2008/11/7 add by sonia
'                    Else
'                        varTmp = Split(strTmp1, ",")
'                        strTmp = Format(varTmp(0))
'                        'Modify by Morgan 2008/12/29
'                        'strTmp1 = CompDate(0, Val(strTmp), DATE1)
'                        If textNP07 = 年費 Then
'                           strTmp1 = CompDate(0, Val(strTmp) - 1, DATE1)
'                        Else
'                           strTmp1 = CompDate(0, Val(strTmp), DATE1)
'                        End If
'                        '非台灣都不用減一天
'                        If pa(9) = "000" Then
'                           strTmp1 = TransDate(CompDate(2, -1, strTmp1), 1)
'                        Else
'                           strTmp1 = TransDate(strTmp1, 1)
'                        End If
'                        If textNP09.Text <> strTmp1 Then
'                           MsgBox "下次繳費日法定期限應為 " & strTmp1, vbCritical
'                           textNP09.SetFocus
'                           GoTo EXITSUB
'                        End If
                    End If
               End If
            End If
         End If
         
      End If 'Added by Morgan 2022/4/21
      
      Case Else:
   End Select
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textNP01_GotFocus()
   InverseTextBox textNP01
End Sub

Private Sub textNP06_GotFocus()
   InverseTextBox textNP06
End Sub

Private Sub textNP07_GotFocus()
   InverseTextBox textNP07
End Sub

Private Sub textNP08_GotFocus()
   InverseTextBox textNP08
End Sub

Private Sub textNP09_GotFocus()
   InverseTextBox textNP09
End Sub

Private Sub textNP10_GotFocus()
   InverseTextBox textNP10
End Sub

Private Sub textNP11_GotFocus()
   InverseTextBox textNP11
End Sub

Private Sub textNP12_GotFocus()
   InverseTextBox textNP12
End Sub

Private Sub textNP13_GotFocus()
   InverseTextBox textNP13
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textNP13.IMEMode = 1
   OpenIme
End Sub

Private Sub textNP14_GotFocus()
   InverseTextBox textNP14
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textNP14.IMEMode = 1
   OpenIme
End Sub

Private Sub textNP15_GotFocus()
   InverseTextBox textNP15
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textNP15.IMEMode = 1
   OpenIme
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strCU13 As String  '2011/5/12 ADD BY SONIA

   TxtValidate = False
   If Me.textNP01.Enabled = True Then
      Cancel = False
      textNP01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textNP06.Enabled = True Then
      Cancel = False
      textNP06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      
      'Added by Morgan 2018/12/3
      '管制半年的年費期限不可上不續辦(相關收文號為不續辦或閉卷)，否則會影響後續的催年費 Ex:FCP-39324
      If (textNP02 = "P" Or textNP02 = "FCP") And textNP06 = "N" And textNP07 = "605" And Left(textNP01, 1) = "B" Then
         strExc(0) = "select cp10 from caseprogress where cp09='" & textNP01 & "' and cp10 in ('907','913')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "管制半年的年費期限不可上不續辦，若確定不管制請通知電腦中心刪除！", vbCritical
            Exit Function
         End If
      End If
      'end 2018/12/3
   End If
   
   If Me.textNP07.Enabled = True Then
      Cancel = False
      textNP07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textNP08.Enabled = True Then
      Cancel = False
      textNP08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      
      'Added by Morgan 2014/11/11
      'Modified by Morgan 2014/11/18 +P,PS;另FCP改為年費預設-2個日曆天,實審預設-4個日曆天,其他不預設
      If (textNP08.Tag <> textNP08 Or textNP09.Tag <> textNP09) And m_Nation = "000" And textNP09 <> "" And (textNP02 = "FCP" Or textNP02 = "P" Or textNP02 = "PS") Then
         If textNP02 = "FCP" Then
            If textNP07 = "605" Then
               'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
               'If textNP08 <> TransDate(CompDate(2, -2, textNP09), 1) Then
               '   If MsgBox("本所期限非法定期限的前2天，是否確定？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               If textNP08 <> TransDate(PUB_GetWorkDay1(CompDate(2, -2, textNP09), True), 1) Then
                  If MsgBox("本所期限非法定期限的前2天(最近的工作天)，是否確定？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               'end 2020/07/07
                     textNP08_GotFocus
                     textNP08.SetFocus
                     Exit Function
                  End If
               End If
            ElseIf textNP07 = "416" Then
               'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
               'If textNP08 <> TransDate(CompDate(2, -4, textNP09), 1) Then
               '   If MsgBox("本所期限非法定期限的前4天，是否確定？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               'Modify By Sindy 2021/8/2 外專本所期限均改為法定期限前2個工作天
               'If textNP08 <> TransDate(PUB_GetWorkDay1(CompDate(2, -4, textNP09), True), 1) Then
               If textNP08 <> TransDate(PUB_GetWorkDay1(CompDate(2, -2, textNP09), True), 1) Then
                  If MsgBox("本所期限非法定期限的前2天(最近的工作天)，是否確定？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               'end 2020/07/07
                     textNP08_GotFocus
                     textNP08.SetFocus
                     Exit Function
                  End If
               End If
            End If
         Else
            If textNP08 <> TransDate(PUB_GetOurDeadline(textNP09), 1) Then
               If MsgBox("本所期限非法定期限的前2個工作天，是否確定？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  textNP08_GotFocus
                  textNP08.SetFocus
                  Exit Function
               End If
            End If
         End If
      End If
      'end 2014/11/18
      'end 2014/11/11
   End If
   
   If Me.textNP09.Enabled = True Then
      Cancel = False
      textNP09_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textNP10.Enabled = True Then
      Cancel = False
      textNP10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      '2011/5/12 ADD BY SONIA 非程序管制案件性質, 依智權人員規則檢查是否相同
      strCU13 = ""
      '程序管制案件性質不檢查
      If m_NP02 = "L" Or m_NP02 = "FCL" Or m_NP02 = "CFL" Or m_NP02 = "LA" Or m_NP02 = "LIN" Then
         If textNP07 = "6001" Then GoTo TAG1
      ElseIf m_NP02 = "P" Or m_NP02 = "PS" Or m_NP02 = "CFP" Or m_NP02 = "CPS" Or m_NP02 = "FCP" Or m_NP02 = "FG" Then
         'Modified by Moragn 2012/5/22 +1209,1603
         If textNP07 = "997" Or textNP07 = "998" Or textNP07 = "994" Or textNP07 = "995" Or textNP07 = "996" Or textNP07 = "999" Or textNP07 = "411" Or textNP07 = "1204" Or textNP07 = "1503" Or textNP07 = "1209" Or textNP07 = "1603" Then
            GoTo TAG1
         End If
      '2015/5/28  MODIFY BY SONIA 商標加1711通知使用宣誓
      'Modified by Lydia 2016/10/20 TC案+994陸代申請書; 商標案+1701 註冊證
      ElseIf textNP07 = "994" Or textNP07 = "997" Or textNP07 = "998" Or textNP07 = "995" Or textNP07 = "996" Or textNP07 = "999" Or textNP07 = "305" Or textNP07 = "1403" Or textNP07 = "312" Or textNP07 = "1701" Or textNP07 = "1711" Then
         GoTo TAG1
      End If
      
      Select Case m_NP02
         Case "FCP", "FG"
             strCU13 = PUB_GetFCPSalesNo(m_NP02, m_NP03, m_NP04, m_NP05)
        Case "FCT"
            strCU13 = PUB_GetFCTSalesNo(m_NP02, m_NP03, m_NP04, m_NP05)
         Case Else
            strCU13 = PUB_GetAKindSalesNo(m_NP02, m_NP03, m_NP04, m_NP05)
      End Select
      If strCU13 <> "" And textNP10 <> strCU13 Then
         If MsgBox("智權人員代號與客戶管制智權人員 " & GetStaffName(strCU13) & " 不符，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
            Cancel = True
            textNP10.SetFocus
            Exit Function
         End If
      End If
      '2011/5/12 END
   End If
   
TAG1:
   If Me.textNP11.Enabled = True Then
      Cancel = False
      textNP11_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textNP12.Enabled = True Then
      Cancel = False
      textNP12_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textNP13.Enabled = True Then
      Cancel = False
      textNP13_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textNP14.Enabled = True Then
      Cancel = False
      textNP14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textNP15.Enabled = True Then
      Cancel = False
      textNP15_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add by Morgan 2010/1/8
   If textNP23.Visible And textNP23.Enabled Then
      Cancel = False
      textNP23_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Added by Lydia 2021/10/12 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If

   TxtValidate = True
End Function
