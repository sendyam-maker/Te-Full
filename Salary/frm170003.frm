VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170003 
   BorderStyle     =   1  '單線固定
   Caption         =   "同仁其他給付資料"
   ClientHeight    =   5052
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8172
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5052
   ScaleWidth      =   8172
   Begin TabDlg.SSTab SSTab1 
      Height          =   4290
      Left            =   30
      TabIndex        =   6
      Top             =   720
      Width           =   8115
      _ExtentX        =   14309
      _ExtentY        =   7557
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm170003.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(17)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDsp(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblDsp(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(9)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textCUID"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtOD(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtOD(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtOD(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtOD(10)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtOD(11)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtNHI10"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtNet"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm170003.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtSum(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtSum(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "GRD1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdok"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txt1(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txt1(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt1(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt1(3)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label8"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label6"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label5"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label12"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label16"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   1
         Left            =   -70680
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3900
         Width           =   1005
      End
      Begin VB.TextBox txtSum 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   0
         Left            =   -72975
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3900
         Width           =   1005
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   270
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   25
         Text            =   "8888888"
         Top             =   2400
         Width           =   915
      End
      Begin VB.TextBox txtNHI10 
         Height          =   285
         Left            =   1485
         MaxLength       =   6
         TabIndex        =   1
         Text            =   "120000"
         Top             =   870
         Width           =   855
      End
      Begin VB.TextBox txtOD 
         Alignment       =   1  '靠右對齊
         Enabled         =   0   'False
         Height          =   270
         Index           =   11
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "666666"
         Top             =   2100
         Width           =   915
      End
      Begin VB.TextBox txtOD 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   270
         Index           =   10
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "1"
         Top             =   1800
         Width           =   315
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm170003.frx":0038
         Height          =   2925
         Left            =   -75000
         TabIndex        =   20
         Top             =   840
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   5165
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "給付日期|員工代號|姓名|公司別|給付金額|補充保費"
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
         _Band(0).Cols   =   6
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   400
         Left            =   -68280
         TabIndex        =   14
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72780
         MaxLength       =   7
         TabIndex        =   11
         Top             =   405
         Width           =   735
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73800
         MaxLength       =   7
         TabIndex        =   10
         Top             =   405
         Width           =   735
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -70680
         MaxLength       =   6
         TabIndex        =   12
         Top             =   405
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -69570
         MaxLength       =   6
         TabIndex        =   13
         Top             =   405
         Width           =   915
      End
      Begin VB.TextBox txtOD 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Index           =   3
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   3
         Text            =   "8888888"
         Top             =   1500
         Width           =   915
      End
      Begin VB.TextBox txtOD 
         Height          =   270
         Index           =   2
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "999999"
         Top             =   1185
         Width           =   915
      End
      Begin VB.TextBox txtOD 
         Height          =   285
         Index           =   1
         Left            =   1470
         MaxLength       =   7
         TabIndex        =   0
         Text            =   "1020101"
         Top             =   560
         Width           =   855
      End
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3870
         Width           =   7005
         VariousPropertyBits=   671105055
         Size            =   "7223;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "合計："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74820
         TabIndex        =   31
         Top             =   3915
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "補充保費："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71670
         TabIndex        =   30
         Top             =   3915
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "給付金額："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73965
         TabIndex        =   29
         Top             =   3915
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "給付淨額："
         Height          =   180
         Index           =   9
         Left            =   510
         TabIndex        =   26
         Top             =   2445
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "給付時間：                     (格式：HHMMSS)"
         Height          =   180
         Left            =   510
         TabIndex        =   24
         Top             =   922
         Width           =   3225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補充保費："
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   23
         Top             =   2145
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公  司  別： "
         Height          =   180
         Index           =   3
         Left            =   510
         TabIndex        =   22
         Top             =   1845
         Width           =   945
      End
      Begin MSForms.Label lblDsp 
         Height          =   285
         Index           =   2
         Left            =   1905
         TabIndex        =   21
         Top             =   1845
         Width           =   2280
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "4022;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblDsp 
         Height          =   285
         Index           =   1
         Left            =   2505
         TabIndex        =   18
         Top             =   1215
         Width           =   750
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1323;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "給付日期：                   －"
         Height          =   180
         Left            =   -74760
         TabIndex        =   16
         Top             =   450
         Width           =   1935
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "員工代號：                      －"
         Height          =   180
         Left            =   -71640
         TabIndex        =   15
         Top             =   450
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "給付金額："
         Height          =   180
         Index           =   17
         Left            =   510
         TabIndex        =   9
         Top             =   1545
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   8
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "給付日期："
         Height          =   180
         Left            =   510
         TabIndex        =   7
         Top             =   600
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
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
            Picture         =   "frm170003.frx":004D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170003.frx":0369
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170003.frx":0685
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170003.frx":0861
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170003.frx":0B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170003.frx":0E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170003.frx":11B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170003.frx":14D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170003.frx":17ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170003.frx":1B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170003.frx":1E25
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8175
      _ExtentX        =   14415
      _ExtentY        =   1080
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
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frm170003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/20 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2008/12/25 add by sonia
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Dim m_FieldList() As FIELDITEM
Dim TF_OD As Integer '欄位數
Dim oText As Object, oLabel As Object
Dim idx As Integer
Dim m_bConfirmCheck As Boolean
Dim m_bActived As Boolean

'Added by Morgan 2013/1/31
Dim stNHI() As String
Dim m_arrOD() As String


Private Sub cmdok_Click()
   If txt1(0) & txt1(1) & txt1(2) & txt1(3) <> "" Then
      If RunNick(txt1(0), txt1(1)) Then
         txt1(0).SetFocus
         Exit Sub
      End If
      If RunNick(txt1(2), txt1(3)) Then
         txt1(2).SetFocus
         Exit Sub
      End If
      GetData
   Else
      MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
      txt1(0).SetFocus
   End If
End Sub

Sub GetData()
Dim stCon As String
   
   txtSum(0) = ""
   txtSum(1) = ""
   stCon = ""
   If txt1(0) <> "" Then
      'Modified by Morgan 2013/1/24
      'stCon = stCon & " and od01>='" & Val(txt1(0)) + 191100 & "' "
      stCon = stCon & " and od01>='" & DBDATE(txt1(0)) & "' "
   End If
   If txt1(1) <> "" Then
      'Modified by Morgan 2013/1/24
      'stCon = stCon & " and od01<='" & Val(txt1(1)) + 191100 & "' "
      stCon = stCon & " and od01<='" & DBDATE(txt1(1)) & "' "
   End If
   If txt1(2) <> "" Then
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'stCon = stCon & " and replace(od02,'A','0')>='" & txt1(2) & "' "
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      'stCon = stCon & " and substr(od02,1,1)||replace(substr(od02,2),'A','0')>='" & txt1(2) & "' "
      stCon = stCon & " and substr(od02,1,2)||replace(substr(od02,3,1),'A','0')||substr(od02,4)>='" & txt1(2) & "' "
   End If
   If txt1(3) <> "" Then
      'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
      'stCon = stCon & " and replace(od02,'A','0')<='" & txt1(3) & "' "
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      'stCon = stCon & " and substr(od02,1,1)||replace(substr(od02,2),'A','0')<='" & txt1(3) & "' "
      stCon = stCon & " and substr(od02,1,2)||replace(substr(od02,3,1),'A','0')||substr(od02,4)<='" & txt1(3) & "' "
   End If
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'strExc(0) = "SELECT od01-191100 給付年月,od02 員工代號,ST02 姓名,od03 給付金額 FROM OtherPayData,staff " & _
               " where replace(od02,'A','0')=st01(+) " & stCon & " order by od01,od02"
   'Modified by Morgan 2013/1/29 給付年月改日期
   'Modified by Morgan 2013/2/4 +公司別,補充保費
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   strExc(0) = "SELECT od01-19110000 給付日期,od02 員工代號,ST02 姓名,od10 公司別,od03 給付金額,od11 補充保費 FROM OtherPayData,staff " & _
               " where substr(od02,1,2)||replace(substr(od02,3,1),'A','0')||substr(od02,4)=st01(+) " & stCon & " order by od01,od02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set GRD1.Recordset = RsTemp.Clone
      GRD1.FormatString = GRD1.FormatString
      'Added by Morgan 2013/2/4 數字靠右
      GRD1.ColAlignment(4) = 7
      GRD1.ColAlignment(5) = 7
      'end 2013/2/4
      
      'Added by Morgan 2013/10/17
      GRD1.ColWidth(2) = 1000
      GRD1.ColWidth(4) = 1000
      GRD1.ColWidth(5) = 1000
      For intI = 0 To GRD1.Cols - 1
         GRD1.ColAlignmentFixed(intI) = flexAlignCenterCenter
         If intI < 4 Then
            GRD1.ColAlignment(intI) = flexAlignLeftCenter
         Else
            GRD1.ColAlignment(intI) = flexAlignRightCenter
         End If
      Next
      If RsTemp.RecordCount > 0 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            txtSum(0) = Val(txtSum(0)) + Val("" & RsTemp("給付金額"))
            txtSum(1) = Val(txtSum(1)) + Val("" & RsTemp("補充保費"))
            RsTemp.MoveNext
         Loop
         txtSum(0) = Format(txtSum(0), "#,##0")
         txtSum(1) = Format(txtSum(1), "#,##0")
      End If
      'end 2013/10/17
   End If
End Sub

Private Sub Form_Activate()
   If m_bActived = False Then
      SetInputEntry
      m_bActived = True
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   
   InitialField
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      Form_KeyDown vbKeyF2, 0
   End If
   UpdateToolbarState
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170003 = Nothing
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   strExc(0) = "select * from OtherPayData where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      With RsTemp
      TF_OD = .Fields.Count
      ReDim m_FieldList(TF_OD) As FIELDITEM
      ReDim m_arrOD(TF_OD) As String 'Added by Morgan 2013/1/31
      
      For Each oText In txtOD
         idx = oText.Index
         m_FieldList(idx).fiName = "OD" & Format(idx, "00")
         'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
         'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
            m_FieldList(idx).fiType = 0
         'Else
         '   m_FieldList(idx).fiType = 1
         'End If
         'end 2017/06/29
      Next
      End With
   End If
   
   ReDim stNHI(TF_NHI) As String 'Added by Morgan 2013/1/3
End Sub
' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
Dim stKey01 As String
Dim stKey02 As String
Dim stKey03 As String
Dim adoRst As New ADODB.Recordset
   
   'Modified by Morgan 2013/1/24
   'stKey01 = Val(txtOD(1)) + 191100
   stKey01 = DBDATE(txtOD(1))
   'end 2013/1/24
   stKey02 = txtOD(2)
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT * FROM OtherPayData,nhi2nd" & _
            " WHERE od01 = " & stKey01 & " and od02= '" & stKey02 & "' and nhi01(+)=od02 and nhi02(+)=od01 and nhi03(+)='50' and nhi04(+)='3'"
      Case -2
         strExc(0) = "SELECT * FROM OtherPayData,nhi2nd where  nhi01(+)=od02 and nhi02(+)=od01 and nhi03(+)='50' and nhi04(+)='3' order by od01 ASC,od02 ASC"
      Case -1
         strExc(0) = "SELECT * FROM OtherPayData,nhi2nd" & _
            " WHERE od01||od02 <'" & stKey01 & stKey02 & "' and nhi01(+)=od02 and nhi02(+)=od01 and nhi03(+)='50' and nhi04(+)='3' order by od01 DESC,od02 DESC"
      Case 1
         strExc(0) = "SELECT * FROM OtherPayData,nhi2nd" & _
            " WHERE od01||od02 >'" & stKey01 & stKey02 & "' and nhi01(+)=od02 and nhi02(+)=od01 and nhi03(+)='50' and nhi04(+)='3' order by od01 ASC,od02 ASC"
      Case 2
         strExc(0) = "SELECT * FROM OtherPayData,nhi2nd where nhi01(+)=od02 and nhi02(+)=od01 and nhi03(+)='50' and nhi04(+)='3' order by od01 DESC,od02 DESC"
   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      UpdateCtrlData adoRst
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         MsgBox "查無資料！", vbInformation
         ClearField
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtOD(1).SetFocus
      txtOD_GotFocus 1
   End If
End Function

Private Sub GRD1_Click()
   Dim lCurRow As Long, i As Integer, j As Integer
   lCurRow = GRD1.row
   If lCurRow > 0 Then
      If GRD1.TextMatrix(lCurRow, 0) <> "" Then
         If GRD1.CellBackColor <> &HFFC0C0 Then
            GRD1.Visible = False
            For j = 1 To GRD1.Rows - 1
               GRD1.row = j
               If GRD1.CellBackColor <> QBColor(15) Then
                  For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = QBColor(15)
                  Next i
               End If
            Next j
            GRD1.row = lCurRow
            For i = 0 To GRD1.Cols - 1
                GRD1.col = i
                GRD1.CellBackColor = &HFFC0C0
            Next i
            GRD1.Visible = True
         End If
      End If
   End If
End Sub

Private Sub GRD1_DblClick()
Dim lCurRow As Long
   
   lCurRow = GRD1.row
   '呼叫查詢
   If lCurRow > 0 Then
      If GRD1.TextMatrix(lCurRow, 0) <> "" Then
         If TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
            If txtOD(1).Locked = False Then
               txtOD(1).Text = GRD1.TextMatrix(lCurRow, 0)
               txtOD(2).Text = GRD1.TextMatrix(lCurRow, 1)
               If TBar1.Buttons(11).Enabled = True Then
                  Call Tbar1_ButtonClick(TBar1.Buttons(11))
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 2 Then
      txt1(0).SetFocus
      TextInverse txt1(0)
   ElseIf SSTab1.Tab = 0 And PreviousTab = 2 Then
      GRD1_DblClick
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   CloseIme
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtOD_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtOD(Index)
End Sub

Private Sub ClearField()
   For Each oText In txtOD
      oText.Text = Empty
   Next
   For Each oLabel In lblDsp
      oLabel.Caption = Empty
   Next
   For intI = 1 To TF_OD
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""
   
   'Added by Morgan 2013/1/31
   txtNHI10 = ""
   Erase m_arrOD
   ReDim m_arrOD(TF_OD) As String
   'end 2013/1/31
   
   m_bConfirmCheck = False
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   With p_Rst
   If .RecordCount > 0 Then
      For Each oText In txtOD
         idx = oText.Index
         Select Case idx
         '年月轉民國年月
         'Modified by Morgan 2013/1/28 +補充保費代扣日期
         Case 1, 12
            'Modified by Morgan 2013/1/24 改為給付日期
            'm_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName) - 191100
            m_FieldList(idx).fiOldData = TransDate("" & .Fields(m_FieldList(idx).fiName), 1)
            'end 2013/1/24
         Case Else
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
         End Select
         m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
         oText.Text = m_FieldList(idx).fiOldData
         'oText.Tag = m_FieldList(idx).fiOldData
      Next
      
      If ClsPDGetStaffN(txtOD(2), strExc(1), , True) Then
         lblDsp(1) = strExc(1)
      End If
      
      'Added by Morgan 2013/1/24
      If txtOD(10) <> "" Then
         lblDsp(2) = CompNameQuery(txtOD(10))
      End If
      txtNHI10 = "" & .Fields("nhi10")
      'end 2013/1/24
                     
      CUID(1) = "" & .Fields("od04")
      CUID(2) = "" & .Fields("od05")
      CUID(3) = "" & .Fields("od06")
      CUID(4) = "" & .Fields("od07")
      CUID(5) = "" & .Fields("od08")
      CUID(6) = "" & .Fields("od09")
   End If
   End With
   UpdateCUID CUID, textCUID
   txtOD(1).Tag = txtOD(1)
   txtOD(2).Tag = txtOD(2)
   SetNet 'Added by Morgan 2013/2/22
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtOD
      oText.Locked = bLocked
   Next
   txtNHI10.Locked = bLocked 'Added by Morgan 2013/1/31
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

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
         
      Case vbKeyEscape:
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction KeyCode
            End If
         End If
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If

   End Select
End Sub

' 執行指令
Public Sub OnAction(ByVal KeyCode As Integer)
   Dim bCancel As Boolean
   
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         SSTab1.Tab = 0
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         'Added by Morgan 2013/1/31
         txtOD(1) = strSrvDate(2)
         txtNHI10 = ServerTime
         txtOD(2).SetFocus
         'end 2013/1/31

      Case vbKeyF3 ' 修改
         SSTab1.Tab = 0
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry

      Case vbKeyF5 ' 刪除
         SSTab1.Tab = 0
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         SSTab1.Tab = 0
         m_EditMode = 4
         SetCtrlReadOnly True
         ClearField
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
      Case vbKeyF9 ' 確定
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         SetInputEntry
         
      Case vbKeyF10 ' 取消
         bCancel = False
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  bCancel = True
               End If
            Case Else
               bCancel = True
         End Select
         If bCancel = True Then
            txtOD(1) = txtOD(1).Tag
            txtOD(2) = txtOD(2).Tag
            m_EditMode = 0
            SetInputEntry
            ShowRecord
            UpdateToolbarState
         End If
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate And txtOD(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtOD(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtOD(1) <> "" Then
            TBar1.Buttons(6).Enabled = True
            TBar1.Buttons(7).Enabled = True
            TBar1.Buttons(8).Enabled = True
            TBar1.Buttons(9).Enabled = True
         Else
            TBar1.Buttons(6).Enabled = False
            TBar1.Buttons(7).Enabled = False
            TBar1.Buttons(8).Enabled = False
            TBar1.Buttons(9).Enabled = False
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
      
      Case 1, 2, 3, 4 '維護
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1
         txtOD(1).Locked = False
         txtOD(2).Locked = False 'Added by Morgan 2013/1/25
         txtNHI10.Locked = False 'Added by Morgan 2013/1/31
         If Me.Visible = True Then
            txtOD(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 2
         txtOD(1).Locked = True
         txtOD(2).Locked = True 'Added by Morgan 2013/1/25
         txtNHI10.Locked = True 'Added by Morgan 2013/1/31
         If Me.Visible = True Then
            txtOD(3).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case 4
         txtOD(1).Locked = False
         txtOD(2).Locked = False 'Added by Morgan 2013/1/25
         If Me.Visible = True Then
            txtOD(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = False
      Case Else
         txtOD(1).Locked = True
         txtOD(2).Locked = False 'Added by Morgan 2013/1/25
         If Me.Visible = True Then
            txtOD(1).SetFocus
         End If
         SSTab1.TabEnabled(1) = True
   End Select
End Sub

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If AddRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 3: '刪除
         'Added by Morgan 2013/3/7
         If ChkHi2ndIsPaid(Left(DBDATE(txtOD(1)), 6)) = True Then
            MsgBox "給付日期月份的補充保費已繳納不可再有異動！", vbExclamation
            Exit Function
         End If
         'end 2013/3/7
      
         If DelRecord = True Then
            OnWork = True
            m_EditMode = 0
            ShowRecord 2
         End If
      
      Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtOD(1).SetFocus
               txtOD_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
   
   m_bConfirmCheck = True
   
   For Each oText In txtOD
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtOD_Validate idx, bCancel
         If bCancel = True Then
            txtOD(idx).SetFocus
            txtOD_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   '查詢
   If m_EditMode = 4 Then
      If txtOD(1) = "" Then
         'Modified by Morgan 2013/1/24
         'ShowMsg "請輸入給付年月 !"
         ShowMsg "請輸入給付日期 !"
         txtOD(1).SetFocus
         txtOD_GotFocus 1
         GoTo EscPoint
      End If
      If txtOD(2) = "" Then
         ShowMsg "請輸入員工代號 !"
         txtOD(2).SetFocus
         txtOD_GotFocus 2
         GoTo EscPoint
      End If
      
   '維護
   Else
      If txtOD(1) = "" And txtOD(1).Locked = False Then
         'Modified by Morgan 2013/1/24
         'ShowMsg "請輸入給付年月 !"
         ShowMsg "請輸入給付日期 !"
         txtOD(1).SetFocus
         txtOD_GotFocus 1
         GoTo EscPoint
      End If
      If txtOD(2) = "" And txtOD(2).Locked = False Then
         ShowMsg "請輸入員工代號 !"
         txtOD(2).SetFocus
         txtOD_GotFocus 2
         GoTo EscPoint
      End If
      If txtOD(3) = "" And txtOD(3).Locked = False Then
         ShowMsg "請輸入給付金額 !"
         txtOD(3).SetFocus
         txtOD_GotFocus 3
         GoTo EscPoint
      End If
      
      'Added by Morgan 2013/3/7
      If ChkHi2ndIsPaid(Left(DBDATE(txtOD(1)), 6)) = True Then
         MsgBox "給付日期月份的補充保費已繳納不可再有異動！", vbExclamation
         GoTo EscPoint
      End If
      'end 2013/3/7
   End If
   TxtValidate = True
   
EscPoint:
   m_bConfirmCheck = False
    
End Function

Private Function AddRecord() As Boolean
Dim stCols As String, stValues As String, stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Added by Morgan 2013/2/6
   If PUB_ChkNHi2nd(stNHI(1), stNHI(2), stNHI(10), True) = False Then
      GoTo EscPoint
   End If
   SetNHI06 '此處要重算以避免輸入過程中同時有新增較早資料而沒算到
   'end 2013/2/6
      
   '畫面有的欄位才更新
   stCols = "": stValues = ""
   For Each oText In txtOD
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            '獎金年月轉西元年月
            'If idx = 1 Then
            '   stValues = stValues & "," & CNULL(Val((m_FieldList(idx).fiNewData) + 191100), True)
            'Else
               stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
            'End If
         End If
      End If
   Next
   stCols = Mid(stCols, 2)
   stValues = Mid(stValues, 2)
   stSQL = "INSERT INTO OtherPayData (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
'   stSQL = "select max(od02) from OtherPayData where od01='" & txtOD(1) & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
'   If intI = 1 Then
'      txtOD(2) = RsTemp.Fields(0)
'   End If
   
   PUB_InsertNHI2nd stNHI 'Added by Morgan 2013/1/31
   
   cnnConnection.CommitTrans
   
   AddRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

EscPoint:

End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Added by Morgan 2013/2/6
   '不可有晚於該筆資料的補充保費
   If PUB_ChkNHi2nd(stNHI(1), stNHI(2), stNHI(10)) = False Then
      GoTo EscPoint
   End If
   'end 2013/2/6
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE OtherPayData SET "
   stSet = ""
   For Each oText In txtOD
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      'Modified by Morgan 2013/1/24
      'stSQL = stSQL & stSet & " where od01='" & Val(txtOD(1)) + 191100 & "' and od02='" & txtOD(2) & "'; end; "
      stSQL = stSQL & stSet & " where od01=" & DBDATE(txtOD(1)) & " and od02='" & txtOD(2) & "'; end; "
      
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
      
      PUB_InsertNHI2nd stNHI 'Added by Morgan 2013/1/31
   End If
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

EscPoint:
End Function

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub UpdateFieldNewData()
   For Each oText In txtOD
      idx = oText.Index
      Select Case idx
         Case 1, 12
            'Modified by Morgan 2013/1/24
            'm_FieldList(idx).fiNewData = Val(oText.Text) + 191100
            m_FieldList(idx).fiNewData = DBDATE(oText.Text)
         Case Else
            m_FieldList(idx).fiNewData = oText.Text
      End Select
   Next
End Sub

Private Sub txtOD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 2
      Case Else
         If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtOD_Validate(Index As Integer, Cancel As Boolean)
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Select Case Index
         Case 1
            If txtOD(Index) <> "" Then
            
               'Modified by Morgan 2013/1/25 改放日期
               'If Right(txtOD(Index), 2) > 12 Then
               '   ShowMsg "給付月份不可超過12 !"
               '   Cancel = True
               'End If
               If ChkDate(txtOD(Index)) = False Then
                  Cancel = True
               'Added by Morgan 2013/2/6
               ElseIf Val(txtOD(Index)) > Val(strSrvDate(2)) Then
                  MsgBox "給付日期不可晚於系統日！", vbExclamation
                  Cancel = True
               ElseIf Val(txtOD(Index)) \ 10000 < (Val(strSrvDate(2)) \ 10000 - 1) Then
                  MsgBox "給付日期不可早於去年度！", vbExclamation
                  Cancel = True
               'end 2013/2/6
               End If
               'end 2013/1/25
            End If
         Case 2
            If txtOD(Index) <> "" Then
               'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
               'If ChkStaffID(Replace(txtOD(Index), "A", "0")) = True Then
               If ChkStaffID(Left(txtOD(Index), 1) & Replace(Mid(txtOD(Index), 2), "A", "0")) = True Then
                  Cancel = True
               End If
               If Cancel = False And ClsPDGetStaffN(txtOD(Index), strExc(1), , True) = False Then
                  Cancel = True
               Else
                  lblDsp(1) = strExc(1)
                  
                  'Added by Morgan 2013/1/24
                  '不可為輸入翻譯人員編號,若日後有需求時補充保費計算中判斷投保公司部分也要一併修改
                  If Left(txtOD(Index), 1) = "F" Then
                     MsgBox "不可為輸入翻譯人員編號！", vbExclamation
                     Cancel = True
                  End If
               
                  '設定薪資公司別
                  strExc(0) = "select sd19,sd28 from salarydata where sd01='" & Left(txtOD(Index), 1) & Replace(Mid(txtOD(Index), 2), "A", "0") & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If Mid(txtOD(2), 2, 1) = "A" Then
                        txtOD(10) = "" & RsTemp(1)
                     Else
                        txtOD(10) = "" & RsTemp(0)
                     End If
                     If txtOD(10) <> "" Then
                        lblDsp(2) = CompNameQuery(txtOD(10))
                     End If
                  End If
                  'end 2013/1/24
                  
               End If
            End If
      End Select
      
      If Cancel = True Then TextInverse txtOD(Index)
      
      '若是按確定的檢查時略過, 檢查代號檔
      If Cancel = False And m_bConfirmCheck = False Then
         Select Case Index
         'Added by Morgan 2013/1/31
         Case 1, 2, 3
            If m_arrOD(Index) <> txtOD(Index) Then
               SetNHI06
            End If
         'end 2013/1/31
         End Select
      End If
      
      'Added by Morgan 2013/1/31
      If Cancel = False Then
         m_arrOD(Index) = txtOD(Index)
      End If
      'end 2013/1/31
   End If
End Sub

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Added by Morgan 2013/2/6
   '檢查不可有晚於該筆資料的補充保費
   SetNHI06
   If PUB_ChkNHi2nd(stNHI(1), stNHI(2), stNHI(10)) = False Then
      GoTo EscPoint
   End If
   'end 2013/2/6
   
   '刪除
   'Modified by Morgan 2013/1/24
   'stSQL = "delete from OtherPayData where od01='" & Val(txtOD(1)) + 191100 & "' and od02='" & txtOD(2) & "'"
   stSQL = "delete from OtherPayData where od01='" & DBDATE(txtOD(1)) & "' and od02='" & txtOD(2) & "'"
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   'Added by Morgan 2013/1/24
   '刪除補充保費
   strSql = "DELETE NHI2ND WHERE NHI01='" & txtOD(2) & "' AND NHI02=" & DBDATE(txtOD(1)) & " AND NHI03='50' AND NHI04='3'"
   cnnConnection.Execute strSql, intI
   'end 2013/1/24
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   txtOD(1).Tag = ""
   txtOD(2).Tag = ""
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
    
EscPoint:

End Function

'計算補充保費
Private Sub SetNHI06()
   
   If txtOD(1) = "" Or txtOD(2) = "" Or txtOD(3) = "" Or txtNHI10 = "" Then
      txtOD(11) = ""
   Else
      stNHI(1) = txtOD(2)
      stNHI(2) = DBDATE(txtOD(1))
      stNHI(3) = "50"
      stNHI(4) = "3"
      stNHI(5) = ""
      stNHI(6) = ""
      stNHI(7) = txtOD(3)
      stNHI(8) = ""
      stNHI(10) = txtNHI10
      stNHI(11) = txtOD(10) 'Added by Morgan 2013/2/26
      PUB_NHI2nd stNHI(1), stNHI(2), stNHI(3), stNHI(4), stNHI(7), stNHI(5), stNHI(6), stNHI(8), stNHI(10), stNHI(11), stNHI(13) 'Modified by Morgan 2013/3/12 +NHI13
      txtOD(11) = Val(stNHI(6))
   End If
   SetNet 'Added by Morgan 2013/2/22
End Sub

Private Sub txtNHI10_GotFocus()
   TextInverse txtNHI10
   CloseIme
End Sub

Private Sub txtNHI10_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtNHI10_Validate(Cancel As Boolean)
   If txtNHI10.Tag <> txtNHI10 Then
      SetNHI06
   End If
   txtNHI10.Tag = txtNHI10
End Sub


'Added by Morgan 2013/2/22
Private Sub SetNet()
   txtNet = Val(txtOD(3)) - Val(txtOD(11))
End Sub
