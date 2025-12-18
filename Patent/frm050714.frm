VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050714 
   BorderStyle     =   1  '單線固定
   Caption         =   "作業失誤資料維護"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9150
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   30
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
            Picture         =   "frm050714.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050714.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050714.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050714.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050714.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050714.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050714.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050714.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050714.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050714.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050714.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
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
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5025
      Left            =   120
      TabIndex        =   23
      Top             =   690
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8864
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm050714.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblStaffName"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "textMD05"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textKEY02"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textMD01"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "textMD02"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textMD03"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textMD04"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textKEY02_2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textKEY04"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textKEY03"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textKEY01"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textMD06"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdSelCp09"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm050714.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(3)"
      Tab(1).Control(1)=   "txt1(4)"
      Tab(1).Control(2)=   "txt1(1)"
      Tab(1).Control(3)=   "txt1(0)"
      Tab(1).Control(4)=   "txt1(2)"
      Tab(1).Control(5)=   "textCP01"
      Tab(1).Control(6)=   "textCP03"
      Tab(1).Control(7)=   "textCP04"
      Tab(1).Control(8)=   "textCP02_2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdQuery"
      Tab(1).Control(10)=   "textCP02"
      Tab(1).Control(11)=   "grdList"
      Tab(1).Control(12)=   "Label1(7)"
      Tab(1).Control(13)=   "Label1(6)"
      Tab(1).Control(14)=   "Label1(5)"
      Tab(1).Control(15)=   "Label8"
      Tab(1).Control(16)=   "Label1(3)"
      Tab(1).Control(17)=   "lblNation"
      Tab(1).Control(18)=   "lblCustName"
      Tab(1).Control(19)=   "Label7"
      Tab(1).Control(20)=   "Label6"
      Tab(1).Control(21)=   "lblFilingNo"
      Tab(1).Control(22)=   "Label5"
      Tab(1).Control(23)=   "Label4"
      Tab(1).ControlCount=   24
      Begin VB.CommandButton cmdSelCp09 
         Caption         =   "選擇總收文號"
         Height          =   300
         Left            =   3600
         TabIndex        =   5
         Top             =   537
         Width           =   1440
      End
      Begin VB.TextBox textMD06 
         Height          =   270
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2220
         Width           =   975
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -69930
         MaxLength       =   3
         TabIndex        =   20
         Top             =   675
         Width           =   525
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   4
         Left            =   -69210
         MaxLength       =   3
         TabIndex        =   21
         Top             =   675
         Width           =   525
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72720
         MaxLength       =   7
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73920
         MaxLength       =   7
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -69930
         MaxLength       =   6
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox textKEY01 
         Height          =   264
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   1
         Top             =   900
         Width           =   492
      End
      Begin VB.TextBox textKEY03 
         Height          =   264
         Left            =   2880
         MaxLength       =   1
         TabIndex        =   3
         Top             =   900
         Width           =   252
      End
      Begin VB.TextBox textKEY04 
         Height          =   264
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   4
         Top             =   900
         Width           =   492
      End
      Begin VB.TextBox textKEY02_2 
         Height          =   264
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   900
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.TextBox textMD04 
         Height          =   270
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   9
         Top             =   1905
         Width           =   975
      End
      Begin VB.TextBox textCP01 
         Height          =   264
         Left            =   -73920
         MaxLength       =   3
         TabIndex        =   15
         Top             =   675
         Width           =   612
      End
      Begin VB.TextBox textCP03 
         Height          =   264
         Left            =   -72360
         MaxLength       =   1
         TabIndex        =   18
         Top             =   675
         Width           =   252
      End
      Begin VB.TextBox textCP04 
         Height          =   264
         Left            =   -72120
         MaxLength       =   2
         TabIndex        =   19
         Top             =   675
         Width           =   492
      End
      Begin VB.TextBox textCP02_2 
         Height          =   264
         Left            =   -72600
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   675
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.TextBox textMD03 
         Height          =   270
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1575
         Width           =   975
      End
      Begin VB.TextBox textMD02 
         Height          =   270
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   7
         Top             =   1245
         Width           =   975
      End
      Begin VB.TextBox textMD01 
         Height          =   270
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   0
         Top             =   552
         Width           =   1212
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Default         =   -1  'True
         Height          =   400
         Left            =   -67140
         TabIndex        =   22
         Top             =   360
         Width           =   912
      End
      Begin VB.TextBox textCP02 
         Height          =   264
         Left            =   -73320
         MaxLength       =   6
         TabIndex        =   16
         Top             =   675
         Width           =   972
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   3315
         Left            =   -74880
         TabIndex        =   30
         Top             =   1650
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5847
         _Version        =   393216
         ScrollTrack     =   -1  'True
         FillStyle       =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox textKEY02 
         Height          =   264
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   2
         Top             =   900
         Width           =   972
      End
      Begin MSForms.TextBox textMD05 
         Height          =   1380
         Left            =   1440
         TabIndex        =   11
         Top             =   2580
         Width           =   7125
         VariousPropertyBits=   -1467989989
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "12568;2434"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "處置金額 : "
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   45
         Top             =   2228
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "~"
         Height          =   255
         Index           =   7
         Left            =   -69360
         TabIndex        =   44
         Top             =   720
         Width           =   225
      End
      Begin VB.Label Label1 
         Caption         =   "失誤人員部門 :"
         Height          =   255
         Index           =   6
         Left            =   -71220
         TabIndex        =   43
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "~"
         Height          =   255
         Index           =   5
         Left            =   -72900
         TabIndex        =   42
         Top             =   390
         Width           =   225
      End
      Begin VB.Label Label8 
         Caption         =   "失誤日期 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   390
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "失誤人員 :"
         Height          =   255
         Index           =   3
         Left            =   -70860
         TabIndex        =   40
         Top             =   390
         Width           =   975
      End
      Begin VB.Label lblNation 
         Height          =   255
         Left            =   -73920
         TabIndex        =   39
         Top             =   1350
         Width           =   1665
      End
      Begin MSForms.Label lblCustName 
         Height          =   255
         Left            =   -73920
         TabIndex        =   38
         Top             =   1050
         Width           =   6615
         VariousPropertyBits=   27
         Size            =   "11668;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "申請國家 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "申請人 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   1050
         Width           =   855
      End
      Begin VB.Label lblFilingNo 
         Height          =   255
         Left            =   -69930
         TabIndex        =   35
         Top             =   1350
         Width           =   1665
      End
      Begin VB.Label Label5 
         Caption         =   "申請案號 :"
         Height          =   255
         Left            =   -70860
         TabIndex        =   34
         Top             =   1350
         Width           =   855
      End
      Begin MSForms.Label lblStaffName 
         Height          =   255
         Left            =   2520
         TabIndex        =   33
         Top             =   1575
         Width           =   2235
         VariousPropertyBits=   27
         Size            =   "3942;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         Caption         =   "本所案號 : "
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "備註 : "
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   31
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "失誤金額 : "
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   28
         Top             =   1905
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "失誤人員 :"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   27
         Top             =   1575
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "失誤日期 :"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1245
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "總收文號 :"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   552
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "本所案號 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   24
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm050714"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/12 改成Form2.0 (grdList,textMD05,lblStaffName,lblCustName)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit
'Modified by Lydia 2015/10/26 +處置金額
'Const MAX_FIELD = 5
Const MAX_FIELD = 6

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
' 第一筆資料的本所案號
Dim m_FirstKEY(1) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(1) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(1) As String
Dim m_CurrSel As Integer
' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_ST04 As String 'Added by Lydia 2016/09/02 人員在職狀態

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT MD01 FROM MissData " & _
            "WHERE MD01 = (SELECT MIN(MD01) FROM MissData ) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("MD01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("MD01")
   End If
   rsTmp.Close

   strSql = "SELECT MD01 FROM MissData " & _
            "WHERE MD01 = (SELECT MAX(MD01) FROM MissData ) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("MD01")) = False Then: m_LastKEY(0) = rsTmp.Fields("MD01")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Sub cmdQuery_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim s As Integer 'Add By Sindy 2010/01/08
   
   'Add By Sindy 2010/01/08
   If Len(textCP01.Text) = 0 And _
      Len(Trim(txt1(0).Text)) = 0 And _
      Len(Trim(txt1(1).Text)) = 0 And _
      Len(Trim(txt1(2).Text)) = 0 And _
      Len(Trim(txt1(3).Text)) = 0 And _
      Len(Trim(txt1(4).Text)) = 0 Then
      s = MsgBox("請至少輸入一項查詢條件！", , "輸入條件不足!")
      Me.textCP01.SetFocus
      Exit Sub
   End If
   
   'Modify By Sindy 2010/01/08
   'If IsEmptyText(textCP01) = True Or IsEmptyText(textCP02) = True Then
   If IsEmptyText(textCP01) <> True And IsEmptyText(textCP02) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入本所案號!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   If textCP01 = "TF" Then
      If IsEmptyText(textCP02_2) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入本所案號!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   
   ' 查詢資料
   If QueryMDFromCP() = False Then
      strTit = "查詢資料"
      strMsg = "無資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
EXITSUB:
End Sub

' Load Form
Private Sub Form_Load()
   SSTab1.Tab = 1

   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm050714", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm050714", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm050714", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm050714", strFind, False)
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
   textKEY01.BackColor = &H8000000F
   textKEY02.BackColor = &H8000000F
   textKEY02_2.BackColor = &H8000000F
   textKEY03.BackColor = &H8000000F
   textKEY04.BackColor = &H8000000F
   
   'Add By Sindy 2010/8/12
   If Pub_StrUserSt03 <> "M51" Then
      txt1(3).Text = Left(Trim(Pub_StrUserSt03), 1)
      txt1(4).Text = Left(Trim(Pub_StrUserSt03), 1) & "99"
   End If
   
   InitialField

   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "MD" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字
      Select Case nIndex
         Case 2, 4:
            m_FieldList(nIndex - 1).fiType = 1 '數字
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
    SetFieldNewData "MD01", textMD01
    If IsEmptyText(textMD02) = False Then
        SetFieldNewData "MD02", DBDATE(textMD02)
    Else
        SetFieldNewData "MD02", textMD02
    End If
    SetFieldNewData "MD03", textMD03
    SetFieldNewData "MD04", textMD04
    SetFieldNewData "MD05", textMD05
    SetFieldNewData "MD06", textMD06 'Added by Lydia 2015/10/26
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

' 清除欄位內的資料內容
Private Sub ClearField()
Dim nIndex As Integer
   
   textMD01 = Empty
   textMD02 = Empty
   textMD03 = Empty
   textMD04 = Empty
   textMD05 = Empty
   textMD06 = Empty 'Added by Lydia 2015/10/26
   textKEY01 = Empty
   textKEY02 = Empty
   textKEY02_2 = Empty
   textKEY03 = Empty
   textKEY04 = Empty
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textMD01.Locked = bEnable
   textMD02.Locked = bEnable
   textMD03.Locked = bEnable
   textMD04.Locked = bEnable
   textMD05.Locked = bEnable
   'Added by Lydia 2015/10/26
   textMD06.Locked = bEnable
   If m_EditMode = 1 Then
        cmdSelCp09.Enabled = True
        textKEY01.Enabled = True
        textKEY02.Enabled = True
        textKEY02_2.Enabled = True
        textKEY03.Enabled = True
        textKEY04.Enabled = True
        textKEY01.BackColor = &H80000005
        textKEY02.BackColor = &H80000005
        textKEY02_2.BackColor = &H80000005
        textKEY03.BackColor = &H80000005
        textKEY04.BackColor = &H80000005
   Else
        cmdSelCp09.Enabled = False
        textKEY01.Enabled = False
        textKEY02.Enabled = False
        textKEY02_2.Enabled = False
        textKEY03.Enabled = False
        textKEY04.Enabled = False
        textKEY01.BackColor = &H8000000F
        textKEY02.BackColor = &H8000000F
        textKEY02_2.BackColor = &H8000000F
        textKEY03.BackColor = &H8000000F
        textKEY04.BackColor = &H8000000F
   End If
   'end 2015/10/26

End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textMD01.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
    strSql = "SELECT * FROM MissData " & _
                "WHERE MD01 = '" & m_CurrKEY(0) & "' Order By MD01 "
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        ClearField
        If IsNull(rsTmp.Fields("MD01")) = False Then
            textMD01 = rsTmp.Fields("MD01")
        End If
        If IsNull(rsTmp.Fields("MD02")) = False Then
            textMD02 = ChangeWStringToTString(rsTmp.Fields("MD02"))
        End If
        If IsNull(rsTmp.Fields("MD03")) = False Then
            textMD03 = rsTmp.Fields("MD03")
        End If
        If IsNull(rsTmp.Fields("MD04")) = False Then
            textMD04 = rsTmp.Fields("MD04")
        End If
        If IsNull(rsTmp.Fields("MD05")) = False Then
           textMD05 = rsTmp.Fields("MD05")
        End If
        'Added by Lydia 2015/10/26
        If IsNull(rsTmp.Fields("MD06")) = False Then
           textMD06 = rsTmp.Fields("MD06")
        End If
        'end 2015/10/26
        ' 更新暫存區的資料
        UpdateFieldOldData rsTmp
    End If
    rsTmp.Close
   
   strSql = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & textMD01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If rsTmp.Fields("CP01") = "TF" Then
         textKEY02_2.Visible = True
         If IsNull(rsTmp.Fields("CP01")) = False Then: textKEY01 = rsTmp.Fields("CP01")
         If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02 = Mid(rsTmp.Fields("CP02"), 1, 5)
         If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02_2 = Mid(rsTmp.Fields("CP02"), 6, 1)
         If IsNull(rsTmp.Fields("CP03")) = False Then: textKEY03 = Mid(rsTmp.Fields("CP03"), 1, 5)
         If IsNull(rsTmp.Fields("CP04")) = False Then: textKEY04 = Mid(rsTmp.Fields("CP04"), 1, 5)
      Else
         textKEY02_2.Visible = False
         If IsNull(rsTmp.Fields("CP01")) = False Then: textKEY01 = rsTmp.Fields("CP01")
         If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02 = rsTmp.Fields("CP02")
         If IsNull(rsTmp.Fields("CP03")) = False Then: textKEY03 = Mid(rsTmp.Fields("CP03"), 1, 5)
         If IsNull(rsTmp.Fields("CP04")) = False Then: textKEY04 = Mid(rsTmp.Fields("CP04"), 1, 5)
      End If
   Else
      textKEY02_2.Visible = False
      textKEY01 = Empty
      textKEY02 = Empty
      textKEY02_2 = Empty
      textKEY03 = Empty
      textKEY04 = Empty
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
   Else
      strSql = "SELECT MD01 FROM MissData " & _
               "WHERE MD01 >= '" & m_CurrKEY(0) & "' Order By MD01 "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("MD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("MD01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   strSql = "SELECT MD01 FROM MissData " & _
            "WHERE MD01 < '" & m_CurrKEY(0) & "' Order By MD01 Desc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("MD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("MD01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
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
   
   If m_CurrKEY(0) = m_LastKEY(0) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   strSql = "SELECT MD01 FROM MissData " & _
            "WHERE MD01 > '" & m_CurrKEY(0) & "' Order By MD01 "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("MD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("MD01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
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
   Set frm050714 = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
   If SSTab1.Tab = 1 Then
      cmdQuery.Default = True
      textCP01.SetFocus
   Else
      cmdQuery.Default = False
   End If
End Sub



Private Sub textCP01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號的系統別
Private Sub textCP01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse

   Cancel = False
   If IsEmptyText(textCP01) = False Then
      ' 使用者沒有權限
      If IsUserHasRightOfSystem(strUserNum, textCP01) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "您沒有使用該系統類別的權限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP01_GotFocus
         GoTo EXITSUB
      End If
      
      Select Case textCP01
         Case "TF":
            textCP02_2.Visible = True
            textCP02_2.Locked = False
            textCP02_2.TabStop = True
            textCP02.MaxLength = 5
         Case Else:
            textCP02_2.Visible = False
            textCP02_2.Locked = True
            textCP02_2.TabStop = False
            textCP02.MaxLength = 6
      End Select
   Else
      textCP02_2.Visible = False
      textCP02_2.Locked = True
      textCP02_2.TabStop = False
      textCP02.MaxLength = 6
   End If
EXITSUB:
End Sub

Private Sub textMD01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 總收文號
Private Sub textMD01_Validate(Cancel As Boolean)
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textMD01) = False Then
      Select Case m_EditMode
         Case 1:
            If IsRecordExist(textMD01) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "該筆延期記錄資料已存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textMD01_GotFocus
               GoTo EXITSUB
            End If
            strSql = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS " & _
                     "WHERE CP09 = '" & textMD01 & "' "
            Set rsTmp = New ADODB.Recordset
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <= 0 Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "該筆收文記錄不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textMD01_GotFocus
            Else
               'Added by Lydia 2015/10/26 判斷輸入的案號和收文號是否為同一案
               If textKEY01 <> "" Or textKEY02 <> "" Then
                  If textKEY01 <> rsTmp.Fields("CP01") Or textKEY02 <> rsTmp.Fields("CP02") _
                     Or (textKEY03 <> "" And textKEY03 <> rsTmp.Fields("CP03")) Or (textKEY03 = "" And rsTmp.Fields("CP03") <> "0") _
                     Or (textKEY04 <> "" And textKEY04 <> rsTmp.Fields("CP04")) Or (textKEY04 = "" And rsTmp.Fields("CP04") <> "00") Then
                     Cancel = True
                     strTit = "檢核資料"
                     strMsg = "該筆收文記錄與本所案號不一致"
                     nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                     cmdSelCp09.SetFocus
                     GoTo EXITSUB
                  End If
               End If
               'end 2015/10/26
               
               If rsTmp.Fields("CP01") = "TF" Then
                  textKEY02_2.Visible = True
                  If IsNull(rsTmp.Fields("CP01")) = False Then: textKEY01 = rsTmp.Fields("CP01")
                  If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02 = Mid(rsTmp.Fields("CP02"), 1, 5)
                  If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02_2 = Mid(rsTmp.Fields("CP02"), 6, 1)
                  If IsNull(rsTmp.Fields("CP03")) = False Then: textKEY03 = Mid(rsTmp.Fields("CP03"), 1, 5)
                  If IsNull(rsTmp.Fields("CP04")) = False Then: textKEY04 = Mid(rsTmp.Fields("CP04"), 1, 5)
               Else
                  textKEY02_2.Visible = False
                  If IsNull(rsTmp.Fields("CP01")) = False Then: textKEY01 = rsTmp.Fields("CP01")
                  If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02 = rsTmp.Fields("CP02")
                  If IsNull(rsTmp.Fields("CP03")) = False Then: textKEY03 = Mid(rsTmp.Fields("CP03"), 1, 5)
                  If IsNull(rsTmp.Fields("CP04")) = False Then: textKEY04 = Mid(rsTmp.Fields("CP04"), 1, 5)
               End If
            End If
            rsTmp.Close
            Set rsTmp = Nothing
      End Select
   End If
EXITSUB:
End Sub

' 失誤日期
Private Sub textMD02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textMD02) = False Then
      If CheckIsTaiwanDate(textMD02, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "失誤日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textMD02_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Sub textMD03_Change()
    'Modified by Lydia 2016/09/02
    'Me.lblStaffName.Caption = GetStaffName(Me.textMD03.Text, False)
    Me.lblStaffName.Caption = GetStaffName(Me.textMD03.Text, True)
End Sub

'Add By Sindy 2010/11/26
Private Sub textMD03_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 原本所期限
Private Sub textMD03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textMD03) = False Then
        'Modified by Lydia 2016/09/02 離職人員改輸入訊息
        'Me.lblStaffName.Caption = GetStaffName(Me.textMD03.Text, False)
        'If Me.lblStaffName.Caption = "" Then
        '    Cancel = True
        '    strTit = "檢核資料"
        '    strMsg = "失誤人員代號不正確"
        '    nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        '    textMD03_GotFocus
        'End If
        Me.lblStaffName.Caption = "": m_ST04 = ""
        strSql = "select ST01,ST02,ST04 from staff where st01='" & Me.textMD03.Text & "' "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
           Me.lblStaffName.Caption = "" & RsTemp.Fields("ST02")
           m_ST04 = "" & RsTemp.Fields("ST04")
        Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "失誤人員代號不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textMD03_GotFocus
        End If
        'end 2016/09/02
   End If
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
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

Private Sub textMD05_GotFocus()
   TextInverse textMD05
End Sub

Private Sub textKEY01_GotFocus()
  TextInverse textKEY01
End Sub

Private Sub textKEY02_2_GotFocus()
  TextInverse textKEY02_2
End Sub

Private Sub textKEY02_GotFocus()
  TextInverse textKEY02
End Sub

Private Sub textKEY03_GotFocus()
  TextInverse textKEY03
End Sub

Private Sub textKEY04_GotFocus()
  TextInverse textKEY04
End Sub

' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Added by Lydia 2015/10/26 控制頁籤
   If Button.Index < 5 Then
      SSTab1.Tab = 0
      SSTab1.TabEnabled(1) = False
   Else
      SSTab1.TabEnabled(1) = True
   End If
   'end 2015/10/26
   
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
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM MissData " & _
            "WHERE MD01 = '" & strKEY01 & "' "
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
   Dim strMD01 As String
   Dim strMD02 As String
   
   strMD01 = textMD01
   ' 檢查記錄是否已存在
   If IsRecordExist(textMD01) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO MissData ("
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
   
   If ((strMD01) < (m_FirstKEY(0))) Or ((strMD01) > (m_LastKEY(0))) Then
      RefreshRange
   End If
   
   ShowCurrRecord strMD01
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
   Dim strMD01 As String
   Dim strMD02 As String
   
   strMD01 = m_CurrKEY(0)
   strSql = "UPDATE MissData SET "
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
                  "WHERE MD01 = '" & strMD01 & "' "
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      ShowCurrRecord strMD01
   End If
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strMD01 As String
   Dim strMD02 As String
   
   strMD01 = m_CurrKEY(0)

   strSql = "DELETE FROM MissData " & _
            "WHERE MD01 = '" & strMD01 & "' "
   cnnConnection.Execute strSql

   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strMD01 = m_LastKEY(0)) Or (strMD01 = m_FirstKEY(0)) Then
      RefreshRange
   End If
   'Modified by Lydia 2015/10/26
   'ShowCurrRecord strMD01
   ShowNextRecord
   SSTab1.TabEnabled(1) = True
   
EXITSUB:
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strMD01 As String
   Dim strMD02 As String
   
   strMD01 = textMD01
   
   QueryRecord = False

   If IsRecordExist(strMD01) = True Then
      m_CurrKEY(0) = strMD01
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
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            AddRecord
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
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
      Case 1: textMD01.SetFocus
      Case 2: textMD02.SetFocus
      Case 4: textMD01.SetFocus
   End Select
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2, 4:
         ' 總收文號不可空白
         If IsEmptyText(textMD01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入總收文號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textMD01.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
   
   Select Case m_EditMode
      Case 1, 2:
         ' 失誤日期不可為空白
         If IsEmptyText(textMD02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入失誤日期"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textMD02.SetFocus
            GoTo EXITSUB
         End If
         ' 失誤人員不可為空白
         If IsEmptyText(textMD03) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入失誤人員"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textMD03.SetFocus
            GoTo EXITSUB
         End If
         ' 失誤金額不可為空白
         If IsEmptyText(textMD04) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入失誤金額"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textMD04.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textMD01_GotFocus()
   InverseTextBox textMD01
End Sub

Private Sub textMD02_GotFocus()
   InverseTextBox textMD02
End Sub

Private Sub textMD03_GotFocus()
   InverseTextBox textMD03
End Sub

Private Sub textMD04_GotFocus()
   InverseTextBox textMD04
End Sub

Private Sub textCP01_GotFocus()
   InverseTextBox textCP01
End Sub

Private Sub textCP02_GotFocus()
   InverseTextBox textCP02
End Sub

Private Sub textCP02_2_GotFocus()
   InverseTextBox textCP02_2
End Sub

Private Sub textCP03_GotFocus()
   InverseTextBox textCP03
End Sub

Private Sub textCP04_GotFocus()
   InverseTextBox textCP04
End Sub

' 初始化列表
Public Sub InitialGridList()
'Added by Lydia 2015/10/26
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   arrGridHeadText = Array("", "收文日", "本所案號", "總收文號", "案件性質", "失誤日期", "失誤人員", "失誤金額", "處置金額", "備註")
   arrGridHeadWidth = Array(300, 860, 1100, 1000, 1100, 860, 860, 860, 860, 1920)
'end 2015/10/26

   
   GrdList.Clear
   'Modified by Morgan 2022/1/11
   'GrdList.Rows = 1
   GrdList.Rows = 2
   'end 2022/1/11
   'Modified by Lydia 2015/10/26
   'grdList.Cols = 8
   GrdList.Cols = UBound(arrGridHeadText) + 1
   
   GrdList.ColWidth(0) = 300
   GrdList.row = 0
   
'   grdList.col = 0
'   grdList.ColAlignment(0) = flexAlignCenterCenter
'   grdList.col = 1
'   grdList.Text = "收文日"
'   grdList.ColWidth(1) = 1000
'   grdList.ColAlignment(1) = flexAlignLeftCenter
'   grdList.col = 2
'   grdList.Text = "總收文號"
'   grdList.ColWidth(2) = 1000
'   grdList.ColAlignment(2) = flexAlignLeftCenter
'   grdList.col = 3
'   grdList.Text = "案件性質"
'   grdList.ColWidth(3) = 1000
'   grdList.ColAlignment(3) = flexAlignLeftCenter
'   grdList.col = 4
'   grdList.Text = "失誤日期"
'   grdList.ColWidth(4) = 1000
'   grdList.ColAlignment(4) = flexAlignLeftCenter
'   grdList.col = 5
'   grdList.Text = "失誤人員"
'   grdList.ColWidth(5) = 1000
'   grdList.ColAlignment(5) = flexAlignLeftCenter
'   grdList.col = 6
'   grdList.Text = "失誤金額"
'   grdList.ColWidth(6) = 1000
'   grdList.ColAlignment(6) = flexAlignRightCenter
'   grdList.col = 7
'   grdList.Text = "備註"
'   grdList.ColWidth(7) = 2000
'   grdList.ColAlignment(7) = flexAlignLeftCenter
   With GrdList
        For iRow = 0 To .Cols - 1
         '  .row = 0
           .col = iRow
           .Text = arrGridHeadText(iRow)
           .ColWidth(iRow) = arrGridHeadWidth(iRow)
           If iRow >= 7 And iRow < .Cols - 1 Then
              .ColAlignment(iRow) = flexAlignRightCenter
           Else
            '  .CellAlignment = flexAlignCenterCenter
              .ColAlignment(iRow) = flexAlignLeftCenter
           End If
        Next
   End With
End Sub

Private Sub grdList_Click()
   grdList_ShowSelection
End Sub

Private Sub grdList_SelChange()
   Dim nRow As Integer
   grdList_ShowSelection
   
   If GrdList.row > 0 And GrdList.row <= GrdList.Rows - 1 Then
      nRow = GrdList.row
      'Modify By Sindy 2010/01/13
      'ShowCurrRecord grdList.TextMatrix(nRow, 1)
      'Modified by Lydia 2015/10/26
      'ShowCurrRecord grdList.TextMatrix(nRow, 2)
      ShowCurrRecord GrdList.TextMatrix(nRow, 3)
   End If
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = GrdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = GrdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < GrdList.Rows Then
      GrdList.row = m_CurrSel
      GrdList.col = 1
      If GrdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To GrdList.Cols - 1
            GrdList.col = nCol
            If GrdList.CellBackColor <> &H80000005 Then: GrdList.CellBackColor = &H80000005
            If GrdList.CellForeColor <> &H80000008 Then: GrdList.CellForeColor = &H80000008
         Next nCol
      End If
      GrdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < GrdList.Rows Then
      GrdList.row = m_CurrSel
      GrdList.col = 1
      For nCol = 1 To GrdList.Cols - 1
         GrdList.col = nCol
         GrdList.CellBackColor = &H8000000D
         GrdList.CellForeColor = &H80000005
      Next nCol
      GrdList.col = 0
   End If
EXITSUB:
End Sub

Private Function QueryMDFromCP() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String, StrSQLa As String
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
   Dim nRow As Integer
   
   QueryMDFromCP = False
   
   ' 組成本所案號
   strCP01 = textCP01
   strCP02 = textCP02
   If strCP01 = "TF" Then: strCP02 = strCP02 & textCP02_2
   strCP03 = textCP03
   If IsEmptyText(strCP03) = True Then: strCP03 = "0"
   strCP04 = textCP04
   If IsEmptyText(strCP04) = True Then: strCP04 = "00"
   
   InitialGridList
   
   'Add By Sindy 2010/01/08
   StrSQLa = ""
   If strCP01 <> "" And strCP02 <> "" And strCP03 <> "" And strCP04 <> "" Then
      StrSQLa = StrSQLa & " And CP01 = '" & strCP01 & "' AND CP02 = '" & strCP02 & "' AND CP03 = '" & strCP03 & "' AND CP04 = '" & strCP04 & "' "
   End If
   If Trim(txt1(0)) <> "" Then
      StrSQLa = StrSQLa & " And MD02 >= " & ChangeTStringToWString(Trim(txt1(0)))
   End If
   If Trim(txt1(1)) <> "" Then
      StrSQLa = StrSQLa & " And MD02 <= " & ChangeTStringToWString(Trim(txt1(1)))
   End If
   If Trim(txt1(2)) <> "" Then
      StrSQLa = StrSQLa & " And MD03 = '" & Trim(txt1(2)) & "' "
   End If
   'Add By Sindy 2010/8/12
   If Trim(txt1(3)) <> "" Then
      StrSQLa = StrSQLa & " And ST03 >= '" & Trim(txt1(3)) & "' "
   End If
   If Trim(txt1(4)) <> "" Then
      StrSQLa = StrSQLa & " And ST03 <= '" & Trim(txt1(4)) & "' "
   End If
   '2010/8/12 End
   
'   strSql = "SELECT " & SQLDate("CP05") & " As CP05, MD01, Decode(PA09, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, PA11 As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Patent, CasePropertyMap, Nation " & _
'            "WHERE MD01 = CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And PA09=NA01 And CP01 = '" & strCP01 & "' AND CP02 = '" & strCP02 & "' AND CP03 = '" & strCP03 & "' AND CP04 = '" & strCP04 & "' "
'   strSql = strSql & " Union SELECT " & SQLDate("CP05") & " As CP05, MD01, Decode(TM10, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, TM12 As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Trademark, CasePropertyMap, Nation " & _
'            "WHERE MD01 = CP09 And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And TM10=NA01 And CP01 = '" & strCP01 & "' AND CP02 = '" & strCP02 & "' AND CP03 = '" & strCP03 & "' AND CP04 = '" & strCP04 & "' "
'   strSql = strSql & " Union SELECT " & SQLDate("CP05") & " As CP05, MD01, Decode(LC15, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, '' As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Lawcase, CasePropertyMap, Nation " & _
'            "WHERE MD01 = CP09 And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And LC15=NA01 And CP01 = '" & strCP01 & "' AND CP02 = '" & strCP02 & "' AND CP03 = '" & strCP03 & "' AND CP04 = '" & strCP04 & "' "
'   strSql = strSql & " Union SELECT " & SQLDate("CP05") & " As CP05, MD01, Decode('000', '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, '' As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Hirecase, CasePropertyMap, Nation " & _
'            "WHERE MD01 = CP09 And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And '000'=NA01 And CP01 = '" & strCP01 & "' AND CP02 = '" & strCP02 & "' AND CP03 = '" & strCP03 & "' AND CP04 = '" & strCP04 & "' "
'   strSql = strSql & " Union SELECT " & SQLDate("CP05") & " As CP05, MD01, Decode(SP09, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, SP11 As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Servicepractice, CasePropertyMap, Nation " & _
'            "WHERE MD01 = CP09 And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And SP09=NA01 And CP01 = '" & strCP01 & "' AND CP02 = '" & strCP02 & "' AND CP03 = '" & strCP03 & "' AND CP04 = '" & strCP04 & "' "
'   strSql = strSql & " Order By CP05, MD01 "
   'Modify By Sindy 2010/01/08
   'Modify By Sindy 2011/2/17 因用SQLDate排序或取MAX或MIN,修改百年蟲問題
'   strSql = "SELECT " & SQLDate("CP05") & " As CP05, MD01, Decode(PA09, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, PA11 As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Patent, CasePropertyMap, Nation " & _
'            "WHERE MD01 = CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And PA09=NA01" & StrSQLa
'   strSql = strSql & " Union SELECT " & SQLDate("CP05") & " As CP05, MD01, Decode(TM10, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, TM12 As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Trademark, CasePropertyMap, Nation " & _
'            "WHERE MD01 = CP09 And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And TM10=NA01" & StrSQLa
'   strSql = strSql & " Union SELECT " & SQLDate("CP05") & " As CP05, MD01, Decode(LC15, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, '' As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Lawcase, CasePropertyMap, Nation " & _
'            "WHERE MD01 = CP09 And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And LC15=NA01" & StrSQLa
'   strSql = strSql & " Union SELECT " & SQLDate("CP05") & " As CP05, MD01, Decode('000', '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, '' As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Hirecase, CasePropertyMap, Nation " & _
'            "WHERE MD01 = CP09 And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And '000'=NA01" & StrSQLa
'   strSql = strSql & " Union SELECT " & SQLDate("CP05") & " As CP05, MD01, Decode(SP09, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, SP11 As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Servicepractice, CasePropertyMap, Nation " & _
'            "WHERE MD01 = CP09 And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And SP09=NA01" & StrSQLa
'   strSql = strSql & " Order By CP05, MD01 "
   'Modified by Lydia 2015/10/26 +本所案號,處置金額MD06
   strExc(1) = ",Decode(CP03||CP04,'000',CP01||'-'||CP02,CP01||'-'||CP02||'-'||CP03||'-'||CP04) CaseNo"
   strSql = "SELECT '', substrb(' '||sqldatet(cp05),-9) As CP05" & strExc(1) & ", MD01, Decode(PA09, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, MD06, PA11 As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Patent, CasePropertyMap, Nation " & _
            "WHERE MD01 = CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And PA09=NA01" & StrSQLa
   strSql = strSql & " Union SELECT '', substrb(' '||sqldatet(cp05),-9) As CP05" & strExc(1) & ", MD01, Decode(TM10, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, MD06, TM12 As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Trademark, CasePropertyMap, Nation " & _
            "WHERE MD01 = CP09 And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And TM10=NA01" & StrSQLa
   strSql = strSql & " Union SELECT '', substrb(' '||sqldatet(cp05),-9) As CP05" & strExc(1) & ", MD01, Decode(LC15, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, MD06, '' As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Lawcase, CasePropertyMap, Nation " & _
            "WHERE MD01 = CP09 And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And LC15=NA01" & StrSQLa
   strSql = strSql & " Union SELECT '', substrb(' '||sqldatet(cp05),-9) As CP05" & strExc(1) & ", MD01, Decode('000', '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, MD06, '' As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Hirecase, CasePropertyMap, Nation " & _
            "WHERE MD01 = CP09 And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And '000'=NA01" & StrSQLa
   strSql = strSql & " Union SELECT '', substrb(' '||sqldatet(cp05),-9) As CP05" & strExc(1) & ", MD01, Decode(SP09, '020', CPM04, CPM03) As CPM03, " & SQLDate("MD02") & " AS MD02, ST02 As MD03, MD04, MD05, MD06, SP11 As FilingNo, NA03 FROM MissData, CASEPROGRESS, Staff, Servicepractice, CasePropertyMap, Nation " & _
            "WHERE MD01 = CP09 And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04 And CP01=CPM01 And CP10=CPM02 AND MD03=ST01 And SP09=NA01" & StrSQLa
   strSql = strSql & " Order By CP05, MD01 "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
        QueryMDFromCP = True
        'Add By Sindy 2010/01/08
        If strCP01 <> "" And strCP02 <> "" And strCP03 <> "" And strCP04 <> "" Then
        '2010/01/08 End
            Me.lblFilingNo.Caption = "" & rsTmp("FilingNo").Value
            Me.lblCustName.Caption = PUB_GetCustName(strCP01 & strCP02 & strCP03 & strCP04)
            Me.lblNation.Caption = "" & rsTmp("NA03").Value
        Else
            Me.lblFilingNo.Caption = ""
            Me.lblCustName.Caption = ""
            Me.lblNation.Caption = ""
        End If
        UpdateGridList rsTmp
   Else
        Me.lblFilingNo.Caption = ""
        Me.lblCustName.Caption = ""
        Me.lblNation.Caption = ""
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset)
   Dim nRow As Integer
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      'Modified by Morgan 2022/1/12
      'GrdList.Rows = GrdList.Rows + 1
      If GrdList.TextMatrix(1, 1) <> "" Then
         GrdList.Rows = GrdList.Rows + 1
      End If
      'end 2022/1/12
      nRow = GrdList.Rows - 1
      If IsNull(rsTmp.Fields("CP05")) = False Then
         GrdList.TextMatrix(nRow, 1) = rsTmp.Fields("CP05")
      End If
      'Added by Lydia 2015/10/26
      If IsNull(rsTmp.Fields("CaseNO")) = False Then
         GrdList.TextMatrix(nRow, 2) = rsTmp.Fields("CaseNo")
      End If
      If IsNull(rsTmp.Fields("MD01")) = False Then
        ' grdList.TextMatrix(nRow, 2) = rsTmp.Fields("MD01")
        GrdList.TextMatrix(nRow, 3) = rsTmp.Fields("MD01")
      End If
      If IsNull(rsTmp.Fields("CPM03")) = False Then
         'grdList.TextMatrix(nRow, 3) = rsTmp.Fields("CPM03")
         GrdList.TextMatrix(nRow, 4) = rsTmp.Fields("CPM03")
      End If
      If IsNull(rsTmp.Fields("MD02")) = False Then
         'grdList.TextMatrix(nRow, 4) = rsTmp.Fields("MD02")
         GrdList.TextMatrix(nRow, 5) = rsTmp.Fields("MD02")
      End If
      If IsNull(rsTmp.Fields("MD03")) = False Then
         'grdList.TextMatrix(nRow, 5) = rsTmp.Fields("MD03")
         GrdList.TextMatrix(nRow, 6) = rsTmp.Fields("MD03")
      End If
      If IsNull(rsTmp.Fields("MD04")) = False Then
         'grdList.TextMatrix(nRow, 6) = rsTmp.Fields("MD04")
         GrdList.TextMatrix(nRow, 7) = rsTmp.Fields("MD04")
      End If
      'Added by Lydia 2015/10/26
      If IsNull(rsTmp.Fields("MD06")) = False Then
         GrdList.TextMatrix(nRow, 8) = rsTmp.Fields("MD06")
      End If
      If IsNull(rsTmp.Fields("MD05")) = False Then
        ' grdList.TextMatrix(nRow, 7) = rsTmp.Fields("MD05")
         GrdList.TextMatrix(nRow, 9) = rsTmp.Fields("MD05")
      End If
      'end 2015/10/26
      rsTmp.MoveNext
   Loop
   
   GrdList.row = 1 'Added by Morgan 2022/1/11 要指定列否則第1筆會反白
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCP01.Enabled = True Then
   Cancel = False
   textCP01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textMD01.Enabled = True Then
   Cancel = False
   textMD01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textMD02.Enabled = True Then
   Cancel = False
   textMD02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textMD03.Enabled = True Then
   Cancel = False
   textMD03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   
   'Added by Lydia 2016/09/02
   If Me.textMD03 <> "" And m_ST04 = "2" Then
      If MsgBox("此人員已離職, 是否要重輸失誤人員？", vbYesNo) = vbYes Then
         textMD03.SetFocus
         textMD03_GotFocus
         Exit Function
      End If
   End If
End If

TxtValidate = True
End Function

'Add By Sindy 2010/01/08
Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

'Add By Sindy 2010/01/08
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 0, 1
      KeyAscii = Pub_NumAscii(KeyAscii)
   Case 2, 3, 4 'Modify By Sindy 2010/8/12
      KeyAscii = UpperCase(KeyAscii)
   Case Else
End Select
End Sub

'Add By Sindy 2010/01/08
Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
   Case 0 '失誤日起
         If Me.txt1(Index) <> "" Then
            If CheckIsTaiwanDate(Me.txt1(Index)) = False Then
               Me.txt1(Index).SetFocus
               Exit Sub
            End If
         End If
   Case 1 '失誤日迄
         If Me.txt1(Index) <> "" Then
            If CheckIsTaiwanDate(Me.txt1(Index)) = False Then
               Me.txt1(Index).SetFocus
               Exit Sub
            End If
            If Val(Me.txt1(0).Text) > Val(Me.txt1(1).Text) Then
               MsgBox "失誤日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.txt1(0).SetFocus
               Exit Sub
            End If
         End If
   'Add By Sindy 2010/8/12
   Case 4 '失誤人員部門迄
         If Me.txt1(Index) <> "" Then
            If Me.txt1(3).Text > Me.txt1(4).Text Then
               MsgBox "失誤人員部門範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.txt1(3).SetFocus
               Exit Sub
            End If
         End If
   '2010/8/12 End
   Case Else
End Select
End Sub
'Added by Lydia 2015/10/26 +處置金額
Private Sub textMD06_GotFocus()
   InverseTextBox textMD06
End Sub
Private Sub textKEY01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textKEY02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textKEY02_2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textKEY03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textKEY04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub cmdSelCp09_Click()
Dim rsRead As New ADODB.Recordset
Dim sqlB As String
 
If Trim(textKEY01) <> "" And Trim(textKEY02) <> "" Then
    Me.Tag = ""
    textKEY03.Text = IIf(textKEY03 = "", "0", textKEY03)
    textKEY04.Text = IIf(textKEY04 = "", "00", textKEY04)
    sqlB = "select '' V," & SQLDate("CP05") & " as 收文日,cp09 as 總收文號,decode(pa09,'000',cpm03,cpm04) as 案件性質,s1.st02 as 承辦人,s2.st02 as 智權人員," & SQLDate("CP27") & " as 發文日 " & _
           "from caseprogress,casepropertymap,staff s1,staff s2,patent " & _
           "where cp01='" & textKEY01 & "' and cp02='" & textKEY02 & "' and cp03='" & textKEY03 & "' and cp04='" & textKEY04 & "' and cp01=cpm01(+) and cp10=cpm02(+) and cp14=s1.st01(+) and cp13=s2.st01(+) " & _
           "and cp57 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) order by cp05 desc"
    intI = 0
    Set rsRead = ClsLawReadRstMsg(intI, sqlB)
    If intI = 1 Then
       Set frm880012.grdDataList.Recordset = rsRead
       Set frm880012.fmParent = Me
       frm880012.iTyp = "1"
       frm880012.Show vbModal
       If Me.Tag <> "" And m_EditMode = 1 Then
          textMD01.Text = Me.Tag
          textMD02.SetFocus
       End If
    End If
Else
   MsgBox "請先輸入本所案號！", vbExclamation, "警告！"
   If Me.textKEY01.Enabled = True Then Me.textKEY01.SetFocus
End If
End Sub

'end 2015/10/26
