VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040102 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件國家收費表維護"
   ClientHeight    =   5940
   ClientLeft      =   270
   ClientTop       =   1040
   ClientWidth     =   9150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9150
   Begin VB.TextBox textCF32 
      Height          =   270
      Left            =   8760
      MaxLength       =   15
      TabIndex        =   24
      Top             =   4980
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.TextBox textCF31 
      Height          =   270
      Left            =   8760
      MaxLength       =   15
      TabIndex        =   23
      Top             =   4680
      Visible         =   0   'False
      Width           =   2292
   End
   Begin VB.TextBox textCF02 
      Height          =   270
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   1
      Top             =   930
      Width           =   612
   End
   Begin VB.TextBox textCF03 
      Height          =   270
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1230
      Width           =   612
   End
   Begin VB.TextBox textCF04 
      Height          =   270
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox textCF05 
      Height          =   270
      Left            =   5250
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox textCF06 
      Height          =   270
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1980
      Width           =   855
   End
   Begin VB.TextBox textCF07 
      Height          =   270
      Left            =   5250
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1980
      Width           =   855
   End
   Begin VB.TextBox textCF08 
      Height          =   270
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox textCF09 
      Height          =   270
      Left            =   5250
      MaxLength       =   12
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox textCF10 
      Height          =   270
      Left            =   6090
      MaxLength       =   20
      TabIndex        =   10
      Top             =   2580
      Width           =   2865
   End
   Begin VB.TextBox textCF11 
      Height          =   270
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   9
      Top             =   2570
      Width           =   855
   End
   Begin VB.TextBox textCF12 
      Height          =   270
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   11
      Top             =   2870
      Width           =   855
   End
   Begin VB.TextBox textCF13 
      Height          =   270
      Left            =   1920
      MaxLength       =   7
      TabIndex        =   13
      Top             =   3180
      Width           =   855
   End
   Begin VB.TextBox textCF14 
      Height          =   270
      Left            =   5250
      MaxLength       =   7
      TabIndex        =   14
      Top             =   3180
      Width           =   855
   End
   Begin VB.TextBox textCF15 
      Height          =   270
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   15
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox textCF22 
      Height          =   270
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   18
      Top             =   4090
      Width           =   855
   End
   Begin VB.TextBox textCF23 
      Height          =   270
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   20
      Top             =   4390
      Width           =   855
   End
   Begin VB.TextBox textCF24 
      Height          =   270
      Left            =   5250
      MaxLength       =   20
      TabIndex        =   19
      Top             =   4080
      Width           =   2292
   End
   Begin VB.TextBox textCF25 
      Height          =   270
      Left            =   5250
      MaxLength       =   2
      TabIndex        =   17
      Top             =   3780
      Width           =   855
   End
   Begin VB.TextBox textCF02_2 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   930
      Width           =   2292
   End
   Begin VB.TextBox textCF26 
      Height          =   270
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   22
      Top             =   4690
      Width           =   372
   End
   Begin VB.TextBox textCF03_2 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1230
      Width           =   2292
   End
   Begin VB.TextBox textCF15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3480
      Width           =   6045
   End
   Begin VB.TextBox textCF28 
      Height          =   270
      Left            =   5250
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2870
      Width           =   396
   End
   Begin VB.TextBox textCF27 
      Height          =   270
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   16
      Top             =   3780
      Width           =   300
   End
   Begin VB.TextBox textCF30 
      Height          =   270
      Left            =   5733
      MaxLength       =   1
      TabIndex        =   21
      Top             =   4390
      Width           =   372
   End
   Begin VB.TextBox textCF01 
      Height          =   270
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   0
      Top             =   630
      Width           =   612
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "擇一同時收文之案件性質:"
      Height          =   180
      Index           =   23
      Left            =   6690
      TabIndex        =   57
      Top             =   5030
      Visible         =   0   'False
      Width           =   2090
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "必要存在的案件性質 :"
      Height          =   180
      Index           =   22
      Left            =   7050
      TabIndex        =   56
      Top             =   4730
      Visible         =   0   'False
      Width           =   1750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統別 :"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   55
      Top             =   690
      Width           =   850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國家代碼 :"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   54
      Top             =   990
      Width           =   940
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質 :"
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   53
      Top             =   1290
      Width           =   940
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作天數(天) :"
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   52
      Top             =   1710
      Width           =   1120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "審查時間(天) :"
      Height          =   180
      Index           =   4
      Left            =   3810
      TabIndex        =   51
      Top             =   1680
      Width           =   1220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "費用(起) :"
      Height          =   180
      Index           =   5
      Left            =   240
      TabIndex        =   50
      Top             =   2010
      Width           =   970
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "規費 : "
      Height          =   180
      Index           =   7
      Left            =   240
      TabIndex        =   49
      Top             =   2310
      Width           =   850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "回音 : "
      Height          =   180
      Index           =   8
      Left            =   3810
      TabIndex        =   48
      Top             =   2280
      Width           =   860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主管機關(列印申請書使用) : "
      Height          =   180
      Index           =   9
      Left            =   3810
      TabIndex        =   47
      Top             =   2610
      Width           =   2210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "提申期限(天) : "
      Height          =   180
      Index           =   10
      Left            =   240
      TabIndex        =   46
      Top             =   2600
      Width           =   1280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "下次管制天數 :"
      Height          =   180
      Index           =   11
      Left            =   240
      TabIndex        =   45
      Top             =   2900
      Width           =   1180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "標準價(點數) :"
      Height          =   180
      Index           =   12
      Left            =   240
      TabIndex        =   44
      Top             =   3210
      Width           =   1210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "費用(迄) : "
      Height          =   180
      Index           =   6
      Left            =   3810
      TabIndex        =   43
      Top             =   1980
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "底價(點數) : "
      Height          =   180
      Index           =   13
      Left            =   3810
      TabIndex        =   42
      Top             =   3210
      Width           =   1220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "下一救濟程序 :"
      Height          =   180
      Index           =   14
      Left            =   240
      TabIndex        =   41
      Top             =   3510
      Width           =   1300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "延期天數 : "
      Height          =   180
      Index           =   15
      Left            =   240
      TabIndex        =   40
      Top             =   4120
      Width           =   860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人收達天數 :"
      Height          =   180
      Index           =   16
      Left            =   240
      TabIndex        =   39
      Top             =   4420
      Width           =   1360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主管機關文書 :"
      Height          =   180
      Index           =   17
      Left            =   3810
      TabIndex        =   38
      Top             =   4140
      Width           =   1290
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   240
      X2              =   8940
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   240
      X2              =   8940
      Y1              =   1570
      Y2              =   1570
   End
   Begin VB.Label Label4 
      Caption         =   "延期月數 :"
      Height          =   260
      Left            =   3810
      TabIndex        =   37
      Top             =   3810
      Width           =   980
   End
   Begin VB.Label Label2 
      Caption         =   "是否通知實體對象 :"
      Height          =   250
      Left            =   240
      TabIndex        =   36
      Top             =   4720
      Width           =   1570
   End
   Begin VB.Label Label3 
      Caption         =   "Y: 通知"
      Height          =   250
      Left            =   2520
      TabIndex        =   35
      Top             =   4690
      Width           =   970
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "延期起算日                        (1.當日 2.次日)"
      Height          =   180
      Index           =   18
      Left            =   240
      TabIndex        =   34
      Top             =   3830
      Width           =   3200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "下次管制月數 : "
      Height          =   180
      Index           =   19
      Left            =   3800
      TabIndex        =   33
      Top             =   2900
      Width           =   1230
   End
   Begin VB.Label Label5 
      Caption         =   "－＞（ 數字：發文預設催審期限                      0：發文不設催審期限       　     空白：不催審）"
      Height          =   590
      Left            =   6300
      TabIndex        =   32
      Top             =   1680
      Width           =   2690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否自動內部收文下一程序: "
      Height          =   180
      Index           =   20
      Left            =   3450
      TabIndex        =   31
      Top             =   4440
      Width           =   2250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Y: 是"
      Height          =   180
      Index           =   21
      Left            =   6180
      TabIndex        =   30
      Top             =   4440
      Width           =   390
   End
   Begin VB.Label Label6 
      Caption         =   "注意：設定FMP案，請使用FCP+申請國家(大陸、香港、澳門)"
      ForeColor       =   &H000000C0&
      Height          =   220
      Left            =   240
      TabIndex        =   26
      Top             =   5670
      Width           =   8650
   End
   Begin MSForms.TextBox textCUID 
      Height          =   290
      Left            =   240
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5340
      Width           =   8630
      VariousPropertyBits=   671105055
      Size            =   "15214;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm12040102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/17 改成Form2.0 ; textCUID
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit


'Modified by Lydia 2020/11/05 目前到CF30
'Const MAX_FIELD = 28
'Modified by Sindy 2024/11/13 目前到CF32
'Const MAX_FIELD = 30
Const MAX_FIELD = 32

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

' 辦識其為外商還是內商的程式
' 0 表內商
' 1 表外商
Dim m_SysKind As Integer

' 第一筆資料的本所案號
Dim m_FirstCF(3) As String
' 最後一筆資料的本所案號
Dim m_LastCF(3) As String
' 目前正在顯示的本所案號
Dim m_CurrCF(3) As String

Dim m_QuerySystem As String

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Const m_strNoRightMsg As String = "您無權限查詢或維護此系統類別+案件性質資料"


Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   'strSQL = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
   '         "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE) AND " & _
   '               "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
   '                       "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE)) AND " & _
   '               "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
   '                       "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE) AND " & _
   '                             "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
   '                                     "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE))) "
   strSql = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
            "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
                          "WHERE CF01 IN " & m_QuerySystem & ") AND " & _
                  "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
                          "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE WHERE CF01 IN " & m_QuerySystem & ")) AND " & _
                  "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
                          "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE WHERE CF01 IN " & m_QuerySystem & ") AND " & _
                                "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
                                        "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE WHERE CF01 IN " & m_QuerySystem & "))) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_FirstCF(0) = rsTmp.Fields("CF01")
      If IsNull(rsTmp.Fields("CF02")) = False Then: m_FirstCF(1) = rsTmp.Fields("CF02")
      If IsNull(rsTmp.Fields("CF03")) = False Then: m_FirstCF(2) = rsTmp.Fields("CF03")
   End If
   rsTmp.Close

   'strSQL = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
   '         "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE) AND " & _
   '               "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
   '                       "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE)) AND " & _
   '               "CF03 = (SELECT MAX(CF03) FROM CASEFEE " & _
   '                       "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE) AND " & _
   '                             "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
   '                                     "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE))) "
   strSql = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
            "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE WHERE CF01 IN " & m_QuerySystem & ") AND " & _
                  "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
                          "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE WHERE CF01 IN " & m_QuerySystem & ")) AND " & _
                  "CF03 = (SELECT MAX(CF03) FROM CASEFEE " & _
                          "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE WHERE CF01 IN " & m_QuerySystem & ") AND " & _
                                "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
                                        "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE WHERE CF01 IN " & m_QuerySystem & "))) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_LastCF(0) = rsTmp.Fields("CF01")
      If IsNull(rsTmp.Fields("CF02")) = False Then: m_LastCF(1) = rsTmp.Fields("CF02")
      If IsNull(rsTmp.Fields("CF03")) = False Then: m_LastCF(2) = rsTmp.Fields("CF03")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' Load Form
Private Sub Form_Load()
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040102", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040102", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040102", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040102", strFind, False)
   
   textCF02_2.BackColor = &H8000000F
   textCF03_2.BackColor = &H8000000F
   textCF15_2.BackColor = &H8000000F
   textCUID.BackColor = &H8000000F
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
   FilterSystem
   
   InitialField
   
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   
End Sub

Private Sub FilterSystem()
   Dim nIndex As Integer
   Dim nCount As Integer
   Dim strSys As String
   Dim strTemp As String
   m_QuerySystem = Empty
   
   strSys = GetUserSystemKind
   nCount = GetSubStringCount(strSys)
   For nIndex = 1 To nCount
      strTemp = GetSubString(strSys, nIndex)
      If IsEmptyText(m_QuerySystem) = False Then m_QuerySystem = m_QuerySystem & ","
      m_QuerySystem = m_QuerySystem & "'" & strTemp & "'"
NextRecord:
   Next nIndex
   
   m_QuerySystem = "(" & m_QuerySystem & ")"

End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CF" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 4, 5, 6, 7, 8, 11, 12, 13, 14, 17, 18, 20, 21, 22, 23, 25:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
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
   SetFieldNewData "CF01", textCF01
   SetFieldNewData "CF02", textCF02
   SetFieldNewData "CF03", textCF03
   SetFieldNewData "CF04", textCF04
   SetFieldNewData "CF05", textCF05
   SetFieldNewData "CF06", textCF06
   SetFieldNewData "CF07", textCF07
   SetFieldNewData "CF08", textCF08
   SetFieldNewData "CF09", textCF09
   SetFieldNewData "CF10", textCF10
   SetFieldNewData "CF11", textCF11
   SetFieldNewData "CF12", textCF12
   SetFieldNewData "CF13", textCF13
   SetFieldNewData "CF14", textCF14
   SetFieldNewData "CF15", textCF15
   SetFieldNewData "CF22", textCF22
   SetFieldNewData "CF23", textCF23
   SetFieldNewData "CF24", textCF24
   SetFieldNewData "CF25", textCF25
   SetFieldNewData "CF26", textCF26
   SetFieldNewData "CF27", textCF27
   SetFieldNewData "CF28", textCF28
   SetFieldNewData "CF30", textCF30 'Added  by Lydia 2020/11/05
   'Add by Sindy 2024/11/13
   SetFieldNewData "CF31", textCF31
   SetFieldNewData "CF32", textCF32
   '2024/11/13 END
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
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
Private Sub QueryDB()
   
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   'modify by sonia 2025/8/28
   'textCF01 = Empty
   'textCF02 = Empty
   'textCF02_2 = Empty
   If m_EditMode <> 4 Then
      textCF01 = Empty
      textCF02 = Empty
      textCF02_2 = Empty
   End If
   'end 2025/8/28
   textCF03 = Empty
   textCF03_2 = Empty
   textCF04 = Empty
   textCF05 = Empty
   textCF06 = Empty
   textCF07 = Empty
   textCF08 = Empty
   textCF09 = Empty
   textCF10 = Empty
   textCF11 = Empty
   textCF12 = Empty
   textCF13 = Empty
   textCF14 = Empty
   textCF15 = Empty
   textCF15_2 = Empty
   textCF22 = Empty
   textCF23 = Empty
   textCF24 = Empty
   textCF25 = Empty
   textCF26 = Empty
   textCF27 = Empty
   textCF28 = Empty
   textCF30 = Empty  'Added by Lydia 2020/11/05
   'Add by Sindy 2024/11/13
   textCF31 = Empty
   textCF32 = Empty
   '2024/11/13 END
   textCUID = Empty
   
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textCF01.Locked = bEnable
   textCF02.Locked = bEnable
   textCF03.Locked = bEnable
   textCF04.Locked = bEnable
   'Modify by Morgan 2009/8/21 控制只有電腦中心能改(其他可由發文或批次設定功能修改)
   'Modify by Morgan 2011/5/30 改控制專利處程序不可改,因為其他部門還是要能輸入
   'If Pub_StrUserSt03 <> "M51" Then
   If Pub_StrUserSt03 = "P12" Then
      textCF05.Locked = True
   Else
      textCF05.Locked = bEnable
   End If
   textCF06.Locked = bEnable
   textCF07.Locked = bEnable
   textCF08.Locked = bEnable
   textCF09.Locked = bEnable
   textCF10.Locked = bEnable
   textCF11.Locked = bEnable
   textCF12.Locked = bEnable
   textCF13.Locked = bEnable
   textCF14.Locked = bEnable
   textCF15.Locked = bEnable
   textCF22.Locked = bEnable
   textCF23.Locked = bEnable
   textCF24.Locked = bEnable
   textCF25.Locked = bEnable
   textCF26.Locked = bEnable
   textCF27.Locked = bEnable
   textCF28.Locked = bEnable
   textCF30.Locked = bEnable 'Added by Lydia 2020/11/05
   'Add by Sindy 2024/11/13
   textCF31.Locked = bEnable
   textCF32.Locked = bEnable
   '2024/11/13 END
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textCF01.Locked = bEnable
   textCF02.Locked = bEnable
   textCF03.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   'Modify By Cheng 2002/02/04
'   strSQL = "SELECT * FROM CASEFEE " & _
'            "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                  "CF02 = '" & m_CurrCF(1) & "' AND " & _
'                  "CF03 = '" & m_CurrCF(2) & "' "
   strSql = "SELECT * FROM CASEFEE " & _
            "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                     m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                  "CF02 = '" & m_CurrCF(1) & "' AND " & _
                  "CF03 = '" & m_CurrCF(2) & "' "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("CF01")) = False Then
         textCF01 = rsTmp.Fields("CF01")
         'add by sonia 2022/7/29
         If textCF01 = "ACS" Then
            Label1(3) = "工作天數(日曆天) :"
         Else
            Label1(3) = "工作天數(天) :"
         End If
         'end 2022/7/29
      End If
      If IsNull(rsTmp.Fields("CF02")) = False Then
         textCF02 = rsTmp.Fields("CF02")
      End If
      If IsNull(rsTmp.Fields("CF03")) = False Then
         textCF03 = rsTmp.Fields("CF03")
      End If
      If IsNull(rsTmp.Fields("CF04")) = False Then
         textCF04 = rsTmp.Fields("CF04")
      End If
      If IsNull(rsTmp.Fields("CF05")) = False Then
         textCF05 = rsTmp.Fields("CF05")
      End If
      If IsNull(rsTmp.Fields("CF06")) = False Then
         textCF06 = rsTmp.Fields("CF06")
      End If
      If IsNull(rsTmp.Fields("CF07")) = False Then
         textCF07 = rsTmp.Fields("CF07")
      End If
      If IsNull(rsTmp.Fields("CF08")) = False Then
         textCF08 = rsTmp.Fields("CF08")
      End If
      If IsNull(rsTmp.Fields("CF09")) = False Then
         textCF09 = rsTmp.Fields("CF09")
      End If
      If IsNull(rsTmp.Fields("CF10")) = False Then
         textCF10 = rsTmp.Fields("CF10")
      End If
      If IsNull(rsTmp.Fields("CF11")) = False Then
         textCF11 = rsTmp.Fields("CF11")
      End If
      If IsNull(rsTmp.Fields("CF12")) = False Then
         textCF12 = rsTmp.Fields("CF12")
      End If
      If IsNull(rsTmp.Fields("CF13")) = False Then
         textCF13 = rsTmp.Fields("CF13")
      End If
      If IsNull(rsTmp.Fields("CF14")) = False Then
         textCF14 = rsTmp.Fields("CF14")
      End If
      If IsNull(rsTmp.Fields("CF15")) = False Then
         textCF15 = rsTmp.Fields("CF15")
      End If
      If IsNull(rsTmp.Fields("CF22")) = False Then
         textCF22 = rsTmp.Fields("CF22")
      End If
      If IsNull(rsTmp.Fields("CF23")) = False Then
         textCF23 = rsTmp.Fields("CF23")
      End If
      If IsNull(rsTmp.Fields("CF24")) = False Then
         textCF24 = rsTmp.Fields("CF24")
      End If
      If IsNull(rsTmp.Fields("CF25")) = False Then
         textCF25 = rsTmp.Fields("CF25")
      End If
      If IsNull(rsTmp.Fields("CF26")) = False Then
         textCF26 = rsTmp.Fields("CF26")
      End If
      If IsNull(rsTmp.Fields("CF27")) = False Then
         textCF27 = rsTmp.Fields("CF27")
      End If
      If IsNull(rsTmp.Fields("CF28")) = False Then
         textCF28 = rsTmp.Fields("CF28")
      End If
      'Added by Lydia 2020/11/05
      If IsNull(rsTmp.Fields("CF30")) = False Then
         textCF30 = rsTmp.Fields("CF30")
      End If
      'end 2020/11/05
      'Add by Sindy 2024/11/13
      If IsNull(rsTmp.Fields("CF31")) = False Then
         textCF31 = rsTmp.Fields("CF31")
      End If
      If IsNull(rsTmp.Fields("CF32")) = False Then
         textCF32 = rsTmp.Fields("CF32")
      End If
      '2024/11/13 END
      
      ' 更新CUID
      UpdateCUID rsTmp
      
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
      
      textCF02_Validate False
      textCF03_Validate False
      textCF15_Validate False
      
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
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
   If IsNull(rsSrcTmp.Fields("CF16")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CF16")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("CF16"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CF17")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CF17")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("CF17"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CF18")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CF18")) = False Then
         strTemp = rsSrcTmp.Fields("CF18")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CF19")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CF19")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("CF19"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CF20")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CF20")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("CF20"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CF21")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CF21")) = False Then
         strTemp = rsSrcTmp.Fields("CF21")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strCF01 As String, ByVal strCF02 As String, ByVal strCF03 As String)
   Dim strSql As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   'Add By Cheng 2002/02/04
   '若檢查使用者無權限新增此系統類別
   If IsRightExist(strCF01, strCF03) = False Then
      strMsg = m_strNoRightMsg
      nResponse = MsgBox(strMsg, vbOKOnly)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   If IsRecordExist(strCF01, strCF02, strCF03) = True Then
      m_CurrCF(0) = strCF01
      m_CurrCF(1) = strCF02
      m_CurrCF(2) = strCF03
   Else
      
      'Modify By Cheng 2002/02/04
'      strSQL = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
'               "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                     "CF02 = '" & m_CurrCF(1) & "' AND " & _
'                     "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
'                             "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                   "CF02 = '" & m_CurrCF(1) & "' AND " & _
'                                   "CF03 > '" & m_CurrCF(2) & "')"
      strSql = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
               "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                     m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                     "CF02 = '" & m_CurrCF(1) & "' AND " & _
                     "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
                             "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                                    m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                                   "CF02 = '" & m_CurrCF(1) & "' AND " & _
                                   "CF03 > '" & m_CurrCF(2) & "')"
      
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrCF(0) = rsTmp.Fields("CF01")
         If IsNull(rsTmp.Fields("CF02")) = False Then: m_CurrCF(1) = rsTmp.Fields("CF02")
         If IsNull(rsTmp.Fields("CF03")) = False Then: m_CurrCF(2) = rsTmp.Fields("CF03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      'Modify By Cheng 2002/02/04
'      strSQL = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
'               "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                     "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
'                             "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                   "CF02 > '" & m_CurrCF(1) & "') AND " & _
'                     "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
'                             "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                   "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
'                                           "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                                 "CF02 > '" & m_CurrCF(1) & "')) "
      strSql = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
               "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                     m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                     "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
                             "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                                    m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                                   "CF02 > '" & m_CurrCF(1) & "') AND " & _
                     "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
                             "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                                    m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                                   "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
                                           "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                                                  m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                                                 "CF02 > '" & m_CurrCF(1) & "')) "
   
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrCF(0) = rsTmp.Fields("CF01")
         If IsNull(rsTmp.Fields("CF02")) = False Then: m_CurrCF(1) = rsTmp.Fields("CF02")
         If IsNull(rsTmp.Fields("CF03")) = False Then: m_CurrCF(2) = rsTmp.Fields("CF03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      'strSQL = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
      '         "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
      '                       "WHERE CF01 > '" & m_CurrCF(0) & "') AND " & _
      '               "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
      '                       "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
      '                                     "WHERE CF01 > '" & m_CurrCF(0) & "')) AND " & _
      '               "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
      '                       "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
      '                                     "WHERE CF01 > '" & m_CurrCF(0) & "') AND " & _
      '                                           "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
      '                                                   "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
      '                                                                 "WHERE CF01 > '" & m_CurrCF(0) & "'))) "
      strSql = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
            "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
                          "WHERE CF01 > '" & m_CurrCF(0) & "' AND " & _
                                "CF01 IN " & m_QuerySystem & ") AND " & _
                  "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
                          "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
                                        "WHERE CF01 > '" & m_CurrCF(0) & "' AND " & _
                                              "CF01 IN " & m_QuerySystem & ")) AND " & _
                  "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
                          "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
                                        "WHERE CF01 > '" & m_CurrCF(0) & "' AND " & _
                                              "CF01 IN " & m_QuerySystem & ") AND " & _
                                "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
                                        "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
                                                      "WHERE CF01 > '" & m_CurrCF(0) & "' AND " & _
                                                            "CF01 IN " & m_QuerySystem & "))) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrCF(0) = rsTmp.Fields("CF01")
         If IsNull(rsTmp.Fields("CF02")) = False Then: m_CurrCF(1) = rsTmp.Fields("CF02")
         If IsNull(rsTmp.Fields("CF03")) = False Then: m_CurrCF(2) = rsTmp.Fields("CF03")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrCF(0) = m_FirstCF(0)
   m_CurrCF(1) = m_FirstCF(1)
   m_CurrCF(2) = m_FirstCF(2)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrCF(0) = m_FirstCF(0) And m_CurrCF(1) = m_FirstCF(1) And m_CurrCF(2) = m_FirstCF(2) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   'Modify By Cheng 2002/02/04
'   strSQL = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
'            "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                  "CF02 = '" & m_CurrCF(1) & "' AND " & _
'                  "CF03 = (SELECT MAX(CF03) FROM CASEFEE " & _
'                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                "CF02 = '" & m_CurrCF(1) & "' AND " & _
'                                "CF03 < '" & m_CurrCF(2) & "')"
   strSql = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
            "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                  m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                  "CF02 = '" & m_CurrCF(1) & "' AND " & _
                  "CF03 = (SELECT MAX(CF03) FROM CASEFEE " & _
                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                                 m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                                "CF02 = '" & m_CurrCF(1) & "' AND " & _
                                "CF03 < '" & m_CurrCF(2) & "')"
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrCF(0) = rsTmp.Fields("CF01")
      If IsNull(rsTmp.Fields("CF02")) = False Then: m_CurrCF(1) = rsTmp.Fields("CF02")
      If IsNull(rsTmp.Fields("CF03")) = False Then: m_CurrCF(2) = rsTmp.Fields("CF03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
'Modify By Cheng 2002/02/04
'   strSQL = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
'            "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                  "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
'                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                "CF02 < '" & m_CurrCF(1) & "') AND " & _
'                  "CF03 = (SELECT MAX(CF03) FROM CASEFEE " & _
'                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
'                                        "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                        "CF02 < '" & m_CurrCF(1) & "')) "
   strSql = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
            "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                  m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                  "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                                 m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                                "CF02 < '" & m_CurrCF(1) & "') AND " & _
                  "CF03 = (SELECT MAX(CF03) FROM CASEFEE " & _
                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                                 m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                                "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
                                        "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
                                        "CF02 < '" & m_CurrCF(1) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrCF(0) = rsTmp.Fields("CF01")
      If IsNull(rsTmp.Fields("CF02")) = False Then: m_CurrCF(1) = rsTmp.Fields("CF02")
      If IsNull(rsTmp.Fields("CF03")) = False Then: m_CurrCF(2) = rsTmp.Fields("CF03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   'strSQL = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
   '         "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE " & _
   '                       "WHERE CF01 < '" & m_CurrCF(0) & "') AND " & _
   '               "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
   '                       "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE " & _
   '                                     "WHERE CF01 < '" & m_CurrCF(0) & "')) AND " & _
   '               "CF03 = (SELECT MAX(CF03) FROM CASEFEE " & _
   '                       "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE " & _
   '                                     "WHERE CF01 < '" & m_CurrCF(0) & "') AND " & _
   '                                           "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
   '                                                   "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE " & _
   '                                                                 "WHERE CF01 < '" & m_CurrCF(0) & "'))) "
   strSql = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
            "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE " & _
                          "WHERE CF01 < '" & m_CurrCF(0) & "' AND " & _
                                "CF01 IN " & m_QuerySystem & ") AND " & _
                  "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
                          "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE " & _
                                        "WHERE CF01 < '" & m_CurrCF(0) & "' AND " & _
                                              "CF01 IN " & m_QuerySystem & ")) AND " & _
                  "CF03 = (SELECT MAX(CF03) FROM CASEFEE " & _
                          "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE " & _
                                        "WHERE CF01 < '" & m_CurrCF(0) & "' AND " & _
                                              "CF01 IN " & m_QuerySystem & ") AND " & _
                                "CF02 = (SELECT MAX(CF02) FROM CASEFEE " & _
                                        "WHERE CF01 = (SELECT MAX(CF01) FROM CASEFEE " & _
                                                      "WHERE CF01 < '" & m_CurrCF(0) & "' AND " & _
                                                            "CF01 IN " & m_QuerySystem & "))) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrCF(0) = rsTmp.Fields("CF01")
      If IsNull(rsTmp.Fields("CF02")) = False Then: m_CurrCF(1) = rsTmp.Fields("CF02")
      If IsNull(rsTmp.Fields("CF03")) = False Then: m_CurrCF(2) = rsTmp.Fields("CF03")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrCF(0) = m_LastCF(0) And m_CurrCF(1) = m_LastCF(1) And m_CurrCF(2) = m_LastCF(2) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   'Modify By Cheng 2002/02/04
'   strSQL = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
'            "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                  "CF02 = '" & m_CurrCF(1) & "' AND " & _
'                  "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
'                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                "CF02 = '" & m_CurrCF(1) & "' AND " & _
'                                "CF03 > '" & m_CurrCF(2) & "')"
   strSql = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
            "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                  m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                  "CF02 = '" & m_CurrCF(1) & "' AND " & _
                  "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                                 m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                                "CF02 = '" & m_CurrCF(1) & "' AND " & _
                                "CF03 > '" & m_CurrCF(2) & "')"
                                
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrCF(0) = rsTmp.Fields("CF01")
      If IsNull(rsTmp.Fields("CF02")) = False Then: m_CurrCF(1) = rsTmp.Fields("CF02")
      If IsNull(rsTmp.Fields("CF03")) = False Then: m_CurrCF(2) = rsTmp.Fields("CF03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   'Modify By Cheng 2002/02/04
'   strSQL = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
'            "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                  "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
'                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                "CF02 > '" & m_CurrCF(1) & "') AND " & _
'                  "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
'                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
'                                        "WHERE CF01 = '" & m_CurrCF(0) & "' AND " & _
'                                              "CF02 > '" & m_CurrCF(1) & "')) "
   strSql = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
            "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                  m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                  "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                                 m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                                "CF02 > '" & m_CurrCF(1) & "') AND " & _
                  "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
                          "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                                 m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                                "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
                                        "WHERE CF01 = '" & m_CurrCF(0) & "' AND '" & _
                                             m_CurrCF(0) & "' IN " & m_QuerySystem & " And " & _
                                              "CF02 > '" & m_CurrCF(1) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrCF(0) = rsTmp.Fields("CF01")
      If IsNull(rsTmp.Fields("CF02")) = False Then: m_CurrCF(1) = rsTmp.Fields("CF02")
      If IsNull(rsTmp.Fields("CF03")) = False Then: m_CurrCF(2) = rsTmp.Fields("CF03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   'strSQL = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
   '         "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
   '                       "WHERE CF01 > '" & m_CurrCF(0) & "') AND " & _
   '               "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
   '                       "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
   '                                     "WHERE CF01 > '" & m_CurrCF(0) & "')) AND " & _
   '               "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
   '                       "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
   '                                     "WHERE CF01 > '" & m_CurrCF(0) & "') AND " & _
   '                                           "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
   '                                                   "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
   '                                                                 "WHERE CF01 > '" & m_CurrCF(0) & "'))) "
   strSql = "SELECT CF01,CF02,CF03 FROM CASEFEE " & _
            "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
                          "WHERE CF01 > '" & m_CurrCF(0) & "' AND " & _
                                "CF01 IN " & m_QuerySystem & ") AND " & _
                  "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
                          "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
                                        "WHERE CF01 > '" & m_CurrCF(0) & "' AND " & _
                                              "CF01 IN " & m_QuerySystem & ")) AND " & _
                  "CF03 = (SELECT MIN(CF03) FROM CASEFEE " & _
                          "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
                                        "WHERE CF01 > '" & m_CurrCF(0) & "' AND " & _
                                              "CF01 IN " & m_QuerySystem & ") AND " & _
                                "CF02 = (SELECT MIN(CF02) FROM CASEFEE " & _
                                        "WHERE CF01 = (SELECT MIN(CF01) FROM CASEFEE " & _
                                                      "WHERE CF01 > '" & m_CurrCF(0) & "' AND " & _
                                                            "CF01 IN " & m_QuerySystem & "))) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrCF(0) = rsTmp.Fields("CF01")
      If IsNull(rsTmp.Fields("CF02")) = False Then: m_CurrCF(1) = rsTmp.Fields("CF02")
      If IsNull(rsTmp.Fields("CF03")) = False Then: m_CurrCF(2) = rsTmp.Fields("CF03")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrCF(0) = m_LastCF(0)
   m_CurrCF(1) = m_LastCF(1)
   m_CurrCF(2) = m_LastCF(2)
   
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

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm12040102 = Nothing
End Sub

Private Sub textCF01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 系統類別
Private Sub textCF01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF01) = False Then
      Select Case m_EditMode
         Case 1, 4:
            If IsAlphabetic(textCF01) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "系統類別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCF01_GotFocus
               GoTo EXITSUB
            End If
            If IsUserHasRightOfSystem(strUserNum, textCF01) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "您沒有使用該系統類別的權限"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCF01_GotFocus
            End If
            If IsCorrectSysKind(textCF01) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "系統類別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCF01_GotFocus
            End If
      End Select
      'add by sonia 2022/7/29
      If textCF01 = "ACS" Then
         Label1(3) = "工作天數(日曆天) :"
      Else
         Label1(3) = "工作天數(天) :"
      End If
      'end 2022/7/29
   End If
EXITSUB:
End Sub

' 國家代碼
Private Sub textCF02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCF02_2 = Empty
   If IsEmptyText(textCF02) = False Then
      textCF02_2 = GetNationName(textCF02, 0)
      Select Case m_EditMode
         Case 1, 2, 4:
            If IsEmptyText(textCF02_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "國家代碼不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCF02_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

Private Sub textCF03_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(textCF03) = False Then
      If m_EditMode = 1 Then
         'Add By Cheng 2002/02/04
         '若檢查使用者無權限新增此系統類別
         If IsRightExist(textCF01, textCF03) = False Then
            strMsg = m_strNoRightMsg
            nResponse = MsgBox(strMsg, vbOKOnly)
            textCF03.SetFocus
            Exit Sub
         End If
         If IsRecordExist(textCF01, textCF02, textCF03) = True Then
            'Cancel = True
            strTit = "檢核資料"
            strMsg = "該筆資料已經存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCF03.SetFocus
         End If
      End If
   End If
End Sub

' 案件性質代號
Private Sub textCF03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCF03_2 = Empty
   If IsEmptyText(textCF03) = False Then
      If textCF02 > "010" Then
         '1:表取得大陸件性質名稱
         'Modify By Sindy 2022/3/30 FMP案件性質改用P系統別去讀取案件名稱
         If textCF01 = "FCP" Then
            textCF03_2 = GetCaseTypeName("P", textCF03, 1)
         Else
         '2022/3/30 END
            textCF03_2 = GetCaseTypeName(textCF01, textCF03, 1)
         End If
      Else
         '0:表取得國內案件性質名稱
         textCF03_2 = GetCaseTypeName(textCF01, textCF03, 0)
      End If
      Select Case m_EditMode
         Case 1, 2, 4:
            If IsEmptyText(textCF03_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "案件性質代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCF03_GotFocus
            End If
         Case Else:
      End Select
      'If m_EditMode = 1 Then
      '   If IsRecordExist(textCF01, textCF02, textCF03) = True Then
      '      Cancel = True
      '      strTit = "檢核資料"
      '      strMsg = "該筆資料已經存在"
      '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '      textCF03_GotFocus
      '   End If
      'End If
   End If
End Sub

' 工作天數
Private Sub textCF04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF04) = False Then
      If IsNumeric(textCF04) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "工作天數(天)請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF04_GotFocus
      End If
   End If
End Sub

' 審查時間
Private Sub textCF05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF05) = False Then
      If IsNumeric(textCF05) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "審查時間(天)請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF05_GotFocus
      End If
   End If
End Sub

' 費用(起)
Private Sub textCF06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF06) = False Then
      If IsNumeric(textCF06) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "費用(起)請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF06_GotFocus
      End If
   End If
End Sub

' 費用(迄)
Private Sub textCF07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF07) = False Then
      If IsNumeric(textCF07) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "費用(迄)請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF07_GotFocus
      End If
   End If
End Sub

' 規費
Private Sub textCF08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF08) = False Then
      If IsNumeric(textCF08) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "規費請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF08_GotFocus
      End If
   End If
End Sub

' 回音
Private Sub textCF09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCF09, 12) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "回音內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCF09_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCF09.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 主管機關
Private Sub textCF10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCF10, 30) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "主管機關內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCF10_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCF10.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 提申期限(天)
Private Sub textCF11_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF11) = False Then
      If IsNumeric(textCF11) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "提申期限(天)請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF11_GotFocus
      End If
   End If
End Sub

' 下次期限(天)
Private Sub textCF12_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF12) = False Then
      If IsNumeric(textCF12) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "下次期限(天)請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF12_GotFocus
      End If
   End If
End Sub

' 標準價(點數)
Private Sub textCF13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF13) = False Then
      If IsNumeric(textCF13) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "標準價(點數)請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF13_GotFocus
      End If
   End If
End Sub

' 底價(點數)
Private Sub textCF14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF14) = False Then
      If IsNumeric(textCF14) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "底價(點數)請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF14_GotFocus
      End If
   End If
End Sub

' 下一救濟程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textCF15_2 = Empty
   If IsEmptyText(textCF15) = False Then
      If textCF02 > "010" Then
         textCF15_2 = GetCaseTypeName(textCF01, textCF15, 1)
      Else
         textCF15_2 = GetCaseTypeName(textCF01, textCF15, 0)
      End If
      Select Case m_EditMode
         Case 1, 2:
            If IsEmptyText(textCF15_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "下一救濟程序代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCF15_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

' 延期天數
Private Sub textCF22_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF22) = False Then
      If IsNumeric(textCF22) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延期天數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF22_GotFocus
      End If
   End If
End Sub

' 代理人收達天數
Private Sub textCF23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF23) = False Then
      If IsNumeric(textCF23) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人收達天數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF23_GotFocus
      End If
   End If
End Sub

' 主管機關文書
Private Sub textCF24_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCF24, 20) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "主管機關文書內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCF24_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCF24.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 延期月數
Private Sub textCF25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF25) = False Then
      If IsNumeric(textCF25) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延期月數請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         TextInverse textCF25
      End If
   End If
End Sub

Private Sub textCF27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   Select Case textCF27.Text
      Case "1", "2", ""
      Case Else
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延期起算日只能輸入 1 或 2 !"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         TextInverse textCF27
   End Select
End Sub

Private Sub textCF28_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF28) = False Then
      If IsNumeric(textCF28) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "下次管制月份請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF28_GotFocus
      End If
   End If
End Sub

Private Sub textCF26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否通知實體對象
Private Sub textCF26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF26) = False Then
      Select Case textCF26
         Case "Y", "", " ":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "是否通知實體對象請輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCF26_GotFocus
      End Select
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
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
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
'edit by nickc 2006/11/13
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub
'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub
' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   m_SubMode = 0
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
            UpdateToolbarState
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
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
   End Select
End Sub

'Add By Sindy 2024/11/13
Private Sub textCF31_GotFocus()
   InverseTextBox textCF31
End Sub
Private Sub textCF31_Validate(Cancel As Boolean)
   Call ChkCmpCode(textCF31, Cancel)
   If Cancel = True Then
      textCF31_GotFocus
   End If
End Sub
Private Sub textCF32_GotFocus()
   InverseTextBox textCF32
End Sub
Private Sub textCF32_Validate(Cancel As Boolean)
   Call ChkCmpCode(textCF32, Cancel)
   If Cancel = True Then
      textCF32_GotFocus
   End If
End Sub
Private Sub ChkCmpCode(oCmp As Object, Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim ii As Integer, strWord As String
Dim arrTmp As Variant
   
   Cancel = False
   
   If oCmp.Text <> "" Then
      For ii = 1 To Len(oCmp)
         strWord = Mid(oCmp.Text, ii, 1)
         If IsNumeric(strWord) = True Then
         ElseIf strWord = "," Then
         Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "此欄位只能輸入 數字 或 逗號(,) !!!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Exit Sub
         End If
      Next ii
      arrTmp = Split(oCmp.Text, ",")
      For ii = LBound(arrTmp) To UBound(arrTmp)
         If textCF02 > "010" Then
            '1:表取得大陸件性質名稱
            'FMP案件性質改用P系統別去讀取案件名稱
            If textCF01 = "FCP" Then
               strExc(10) = GetCaseTypeName("P", arrTmp(ii), 1)
            Else
               strExc(10) = GetCaseTypeName(textCF01, arrTmp(ii), 1)
            End If
         Else
            '0:表取得國內案件性質名稱
            strExc(10) = GetCaseTypeName(textCF01, arrTmp(ii), 0)
         End If
         If IsEmptyText(strExc(10)) = True Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = arrTmp(ii) & "案件性質代號不存在，請確認！！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Exit Sub
         End If
      Next ii
   End If
End Sub
'2024/11/13 END

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

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strCF01 As String, ByVal strCF02 As String, ByVal strCF03 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   
   strSql = "SELECT * FROM CASEFEE " & _
            "WHERE CF01 = '" & strCF01 & "' AND " & _
                  "CF02 = '" & strCF02 & "' AND " & _
                  "CF03 = '" & strCF03 & "' "
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   'Modify By Cheng 2002/02/04
'   rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 檢查使用者是否有權限
Private Function IsRightExist(ByVal strCF01 As String, ByVal strCF03 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRightExist = False
   
   'Modify By Sindy 2022/3/30 FMP權限 ex:FCP系統別無936案件性質
   '外專不能修改, 但至少電腦中心可以在此作業代改
   If strCF01 = "FCP" And textCF02 <> "000" And InStr(m_QuerySystem, "'P'") > 0 Then
      strCF01 = "P"
   End If
   '2022/3/30 END
   
   strSql = "SELECT SG01,SG02,SG03 FROM Staff,Staff_Group " & _
            " WHERE ST11=SG01(+) AND SG02 IN " & m_QuerySystem & _
            " AND SG02='" & strCF01 & "' And ST01='" & strUserNum & "' " & _
            " And SG03='" & strCF03 & "'"
                  
   ' 讀取資料庫
   If rsTmp.State <> adStateClosed Then rsTmp.Close
   Set rsTmp = Nothing
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRightExist = True
   Else
      IsRightExist = False
   End If
   rsTmp.Close
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
   Dim strCF01 As String
   Dim strCF02 As String
   Dim strCF03 As String
   
   strCF01 = textCF01
   strCF02 = textCF02
   strCF03 = textCF03
   
   'Add By Cheng 2002/02/04
   '若檢查使用者無權限新增此系統類別
   If IsRightExist(strCF01, strCF03) = False Then
      strMsg = m_strNoRightMsg
      nResponse = MsgBox(strMsg, vbOKOnly)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strCF01, strCF02, strCF03) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO CASEFEE ("
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

   cnnConnection.Execute strSql
   
   If ((strCF01 & strCF02 & strCF03) < (m_FirstCF(0) & m_FirstCF(1) & m_FirstCF(2))) Or ((strCF01 & strCF02 & strCF03) > (m_LastCF(0) & m_LastCF(1) & m_LastCF(2))) Then
      RefreshRange
   End If
   
   ShowCurrRecord strCF01, strCF02, strCF03
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
   Dim strCF01 As String
   Dim strCF02 As String
   Dim strCF03 As String
   
   strCF01 = m_CurrCF(0)
   strCF02 = m_CurrCF(1)
   strCF03 = m_CurrCF(2)
   '910910  nick tigger
   '***** start
   'strSQL = "UPDATE CASEFEE SET "
   strSql = "begin user_data.user_enabled:=1; UPDATE CASEFEE SET "
   '***** end
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      '92.05.22 nick 跳過create & update 相關項目
      If nIndex < 15 Or nIndex > 20 Then
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
        End If
   Next nIndex
   '910910 nick tigger
   '***** start
   'strSQL = strSQL & " " & _
                  "WHERE CF01 = '" & strCF01 & "' AND " & _
                        "CF02 = '" & strCF02 & "' AND " & _
                        "CF03 = '" & strCF03 & "' "
    strSql = strSql & " " & _
                  "WHERE CF01 = '" & strCF01 & "' AND " & _
                        "CF02 = '" & strCF02 & "' AND " & _
                        "CF03 = '" & strCF03 & "'; end; "
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
      cnnConnection.Execute strSql
      '910910 nick tigger
      '***** start
      cnnConnection.CommitTrans
      '***** end
      ShowCurrRecord strCF01, strCF02, strCF03
   End If
'910910 nick tigger
'***** start
   Exit Sub
ErrHand:
    MsgBox (Err.Description)
    cnnConnection.RollbackTrans
'******* end
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strCF01 As String
   Dim strCF02 As String
   Dim strCF03 As String
   
   strCF01 = m_CurrCF(0)
   strCF02 = m_CurrCF(1)
   strCF03 = m_CurrCF(2)

   strSql = "DELETE FROM CASEFEE " & _
            "WHERE CF01 = '" & strCF01 & "' AND " & _
                  "CF02 = '" & strCF02 & "' AND " & _
                  "CF03 = '" & strCF03 & "' "

   cnnConnection.Execute strSql

   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strCF01 = m_LastCF(0) And strCF02 = m_LastCF(1) And strCF03 = m_LastCF(2)) Or (strCF01 = m_FirstCF(0) And strCF02 = m_FirstCF(1) And strCF03 = m_FirstCF(2)) Then
      RefreshRange
   End If
   ShowCurrRecord strCF01, strCF02, strCF03
   
EXITSUB:
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strMsg As String
Dim nResponse

   QueryRecord = False
   
   'Add By Cheng 2002/02/04
   '若檢查使用者無權限新增此系統類別
   If IsRightExist(textCF01, textCF03) = False Then
      strMsg = m_strNoRightMsg
      nResponse = MsgBox(strMsg, vbOKOnly)
      UpdateCtrlData
      Exit Function
   End If

   If IsRecordExist(textCF01, textCF02, textCF03) = True Then
      m_CurrCF(0) = textCF01
      m_CurrCF(1) = textCF02
      m_CurrCF(2) = textCF03
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   UpdateToolbarState
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
            
            AddRecord
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            ModRecord
         Else
            GoTo EXITSUB
         End If
      Case 3:
         DelRecord
         RefreshRange
      Case 4:
         If CheckDataValid() = True Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textCF01.SetFocus
      Case 2: textCF04.SetFocus
      Case 4: textCF03.SetFocus    'modify by sonia 2025/8/28 原停在textCF01
   End Select
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2, 4:
         ' 系統別不可空白
         If IsEmptyText(textCF01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入系統別"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCF01.SetFocus
            GoTo EXITSUB
         End If
         ' 國家代碼不可為空白
         If IsEmptyText(textCF02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入國家代碼"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCF02.SetFocus
            GoTo EXITSUB
         End If
         ' 案件性質不可為空白
         If IsEmptyText(textCF03) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入案件性質"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCF03.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
   
   Select Case m_EditMode
      Case 1, 2:
         If IsEmptyText(textCF06) = False And IsEmptyText(textCF07) = False Then
            If Val(textCF06) > Val(textCF07) Then
               strTit = "檢核資料"
               strMsg = "費用起迄範圍不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCF06.SetFocus
               GoTo EXITSUB
            End If
         End If
         'Add by Morgan 2010/2/5 P大陸改資料要檢查是否有FMP設定
         If textCF01 = "P" And textCF02 <> "000" Then
            strExc(0) = "select * from casefee where cf01='FCP' and cf02='" & textCF02 & "' and cf03='" & textCF03 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "本筆資料同時有FMP設定，若需一併調整請通知電腦中心人員！"
            End If
         End If
   End Select
      
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCF01_GotFocus()
   InverseTextBox textCF01
   CloseIme
End Sub

Private Sub textCF02_GotFocus()
   InverseTextBox textCF02
End Sub

Private Sub textCF03_GotFocus()
   InverseTextBox textCF03
End Sub

Private Sub textCF04_GotFocus()
   InverseTextBox textCF04
End Sub

Private Sub textCF05_GotFocus()
   InverseTextBox textCF05
End Sub

Private Sub textCF06_GotFocus()
   InverseTextBox textCF06
End Sub

Private Sub textCF07_GotFocus()
   InverseTextBox textCF07
End Sub

Private Sub textCF08_GotFocus()
   InverseTextBox textCF08
End Sub

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCF09.IMEMode = 1
   OpenIme
End Sub

Private Sub textCF10_GotFocus()
   InverseTextBox textCF10
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCF10.IMEMode = 1
   OpenIme
End Sub

Private Sub textCF11_GotFocus()
   InverseTextBox textCF11
End Sub

Private Sub textCF12_GotFocus()
   InverseTextBox textCF12
End Sub

Private Sub textCF13_GotFocus()
   InverseTextBox textCF13
End Sub

Private Sub textCF14_GotFocus()
   InverseTextBox textCF14
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

Private Sub textCF22_GotFocus()
   InverseTextBox textCF22
End Sub

Private Sub textCF23_GotFocus()
   InverseTextBox textCF23
End Sub

Private Sub textCF24_GotFocus()
   InverseTextBox textCF24
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCF24.IMEMode = 1
   OpenIme
End Sub

Private Sub textCF25_GotFocus()
   InverseTextBox textCF25
End Sub

Private Sub textCF26_GotFocus()
   InverseTextBox textCF26
End Sub

Private Sub textCF27_GotFocus()
   InverseTextBox textCF27
End Sub

Private Sub textCF28_GotFocus()
   InverseTextBox textCF28
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCF01.Enabled = True Then
   Cancel = False
   textCF01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF02.Enabled = True Then
   Cancel = False
   textCF02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF03.Enabled = True Then
   Cancel = False
   textCF03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF04.Enabled = True Then
   Cancel = False
   textCF04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF05.Enabled = True Then
   Cancel = False
   textCF05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF06.Enabled = True Then
   Cancel = False
   textCF06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF07.Enabled = True Then
   Cancel = False
   textCF07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF08.Enabled = True Then
   Cancel = False
   textCF08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF09.Enabled = True Then
   Cancel = False
   textCF09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF10.Enabled = True Then
   Cancel = False
   textCF10_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF11.Enabled = True Then
   Cancel = False
   textCF11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF12.Enabled = True Then
   Cancel = False
   textCF12_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF13.Enabled = True Then
   Cancel = False
   textCF13_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF14.Enabled = True Then
   Cancel = False
   textCF14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF15.Enabled = True Then
   Cancel = False
   textCF15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF22.Enabled = True Then
   Cancel = False
   textCF22_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF23.Enabled = True Then
   Cancel = False
   textCF23_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF24.Enabled = True Then
   Cancel = False
   textCF24_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF25.Enabled = True Then
   Cancel = False
   textCF25_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF26.Enabled = True Then
   Cancel = False
   textCF26_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF27.Enabled = True Then
   Cancel = False
   textCF27_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF28.Enabled = True Then
   Cancel = False
   textCF28_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Lydia 2020/11/05
If Me.textCF30.Enabled = True Then
   Cancel = False
   textCF30_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'end 2020/11/05

'Add by Sindy 2024/11/13
If Me.textCF31.Enabled = True Then
   Cancel = False
   textCF31_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textCF32.Enabled = True Then
   Cancel = False
   textCF32_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'2024/11/13 END

TxtValidate = True
End Function

'Added by Lydia 2020/11/05
Private Sub textCF30_GotFocus()
   InverseTextBox textCF30
End Sub

Private Sub textCF30_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCF30_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCF30) = False Then
      Select Case textCF30
         Case "Y", "", " ":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "是否自動內部收文下一程序請輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCF30_GotFocus
      End Select
   End If
End Sub
