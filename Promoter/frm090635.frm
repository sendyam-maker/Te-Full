VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090635 
   BorderStyle     =   1  '單線固定
   Caption         =   "價目表權限維護"
   ClientHeight    =   5460
   ClientLeft      =   6090
   ClientTop       =   1550
   ClientWidth     =   9140
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9140
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
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
            Picture         =   "frm090635.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090635.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090635.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090635.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090635.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090635.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090635.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090635.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090635.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090635.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090635.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   60
      TabIndex        =   12
      Top             =   720
      Width           =   9015
      _ExtentX        =   15893
      _ExtentY        =   8273
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm090635.frx":20F4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lstUsers"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label23"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lstDept"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TextPLQ03_P"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TextPLQ03_D"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdAdd"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdRemove"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm090635.frx":2110
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(4)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label9"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "GRD1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdOK(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdOK(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Text1(0)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmdRemove 
         Caption         =   "移除 ->"
         Height          =   285
         Left            =   -67470
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1770
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "<- 新增"
         Height          =   285
         Left            =   -67470
         TabIndex        =   5
         Top             =   1470
         Width           =   735
      End
      Begin VB.TextBox TextPLQ03_D 
         Height          =   285
         Left            =   -72090
         TabIndex        =   23
         Text            =   "TextPLQ03_D"
         Top             =   3510
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.TextBox TextPLQ03_P 
         Height          =   285
         Left            =   -72090
         TabIndex        =   22
         Text            =   "TextPLQ03_P"
         Top             =   3870
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.ListBox lstDept 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2920
         ItemData        =   "frm090635.frx":212C
         Left            =   -73680
         List            =   "frm090635.frx":2133
         MultiSelect     =   1  '簡易多重選取
         Sorted          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1725
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   1095
         Left            =   -73740
         TabIndex        =   20
         Top             =   1050
         Width           =   2805
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   30
            TabIndex        =   1
            Text            =   "cboDept"
            Top             =   90
            Width           =   2535
         End
         Begin VB.CommandButton cmdAddDept 
            Caption         =   "<- 新增"
            Height          =   285
            Left            =   1800
            TabIndex        =   2
            Top             =   420
            Width           =   735
         End
         Begin VB.CommandButton cmdRemoveDept 
            Caption         =   "移除 ->"
            Height          =   285
            Left            =   1800
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   -73920
         MaxLength       =   2
         TabIndex        =   0
         Top             =   570
         Width           =   465
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1170
         MaxLength       =   2
         TabIndex        =   8
         Top             =   450
         Width           =   465
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "全部查詢"
         Height          =   400
         Index           =   1
         Left            =   7740
         TabIndex        =   10
         Top             =   810
         Width           =   975
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "類別查詢"
         Height          =   400
         Index           =   0
         Left            =   6690
         TabIndex        =   9
         Top             =   810
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm090635.frx":2140
         Height          =   3170
         Left            =   90
         TabIndex        =   24
         Top             =   1200
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   5592
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "系統類別|種類|可查詢部門或個人"
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
         _Band(0).Cols   =   3
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  '沒有框線
         Height          =   405
         Left            =   -69240
         TabIndex        =   18
         Top             =   1050
         Width           =   2805
         Begin MSForms.ComboBox cboUser 
            Height          =   300
            Left            =   30
            TabIndex        =   28
            Top             =   90
            Width           =   2550
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "4498;529"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Label Label9 
         Caption         =   "註：資料列點二下才會查詢明細資料。"
         ForeColor       =   &H000000C0&
         Height          =   250
         Left            =   90
         TabIndex        =   30
         Top             =   4410
         Width           =   4690
      End
      Begin MSForms.Label Label23 
         Height          =   195
         Left            =   -74430
         TabIndex        =   29
         Top             =   4410
         Width           =   7905
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13944;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   2940
         Left            =   -69210
         TabIndex        =   7
         Top             =   1440
         Width           =   1725
         VariousPropertyBits=   746586139
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "3043;5186"
         MatchEntry      =   0
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "（01.國內專利   02.大陸專利       03.香港澳門專利   04.CFP"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -73410
         TabIndex        =   27
         Top             =   390
         Width           =   5805
      End
      Begin VB.Label Label3 
         Caption         =   "　05.國內商標   06.大陸商標       07.馬德里商標       08.國內著作權   09.大陸著作權"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -73410
         TabIndex        =   26
         Top             =   630
         Width           =   6855
      End
      Begin VB.Label Label2 
         Caption         =   "　10.CFT            11.美國著作權   12.顧問及法務）"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -73410
         TabIndex        =   25
         Top             =   870
         Width           =   5805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "可查詢部門："
         Height          =   180
         Index           =   0
         Left            =   -74820
         TabIndex        =   21
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "可查詢個人："
         Height          =   180
         Index           =   3
         Left            =   -70290
         TabIndex        =   19
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "系統類別："
         Height          =   180
         Index           =   5
         Left            =   -74820
         TabIndex        =   17
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label7 
         Caption         =   "　10.CFT            11.美國著作權   12.顧問及法務）"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   1680
         TabIndex        =   16
         Top             =   870
         Width           =   5805
      End
      Begin VB.Label Label6 
         Caption         =   "　05.國內商標   06.大陸商標       07.馬德里商標       08.國內著作權   09.大陸著作權"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   1680
         TabIndex        =   15
         Top             =   630
         Width           =   6855
      End
      Begin VB.Label Label5 
         Caption         =   "（01.國內專利   02.大陸專利       03.香港澳門專利   04.CFP"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   1680
         TabIndex        =   14
         Top             =   390
         Width           =   5805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "系統類別："
         Height          =   180
         Index           =   4
         Left            =   270
         TabIndex        =   13
         Top             =   510
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9140
      _ExtentX        =   16122
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
End
Attribute VB_Name = "frm090635"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/1 Form2.0已修改
'Create By Sindy 2014/2/25
'Memo by Lydia 2019/08/08 表單名稱「價目表查詢權限維護」=>「價目表權限維護」
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
' 第一筆資料的系統類別
Dim m_FirstKEY(1) As String
' 最後一筆資料的系統類別
Dim m_LastKEY(1) As String
' 目前正在顯示的系統類別
Dim m_CurrKEY(1) As String
Dim rsA As New ADODB.Recordset
Dim tf_PLQ As Integer
Dim m_PLQ03_D As String, m_PLQ03_P As String


Private Sub cmdOK_Click(Index As Integer)
Dim rsTmp As New ADODB.Recordset
Dim i As Integer
Dim strText As String, strTemp As String
   
   Select Case Index
      Case 0 '類別查詢
         If Text1(0) = "" Then
            MsgBox "系統類別不可空白！", vbExclamation, "操作錯誤！"
            Text1(0).SetFocus
            Exit Sub
         End If
         strSql = "SELECT decode(plq01,'01','國內專利','02','大陸專利','03','香港澳門專利','04','CFP','05','國內商標','06','大陸商標','07','馬德里商標','08','國內著作權','09','大陸著作權','10','CFT','11','美國著作權','12','顧問及法務',plq01)" & _
                  ",decode(plq02,'D','部門','P','個人',''),plq03,plq01" & _
                  " FROM pricelistquery where plq01='" & Text1(0) & "' order by plq02"
         If rsTmp.State = 1 Then rsTmp.Close
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         Set GRD1.Recordset = rsTmp
         SetGrd
         If rsTmp.RecordCount = 0 Then
            GRD1.Rows = 2
            GRD1.row = 1
            GRD1.col = 0
            MsgBox "無此資料", vbOKOnly, "查詢資料"
         End If
      Case 1 '全部查詢
         strSql = "SELECT decode(plq01,'01','國內專利','02','大陸專利','03','香港澳門專利','04','CFP','05','國內商標','06','大陸商標','07','馬德里商標','08','國內著作權','09','大陸著作權','10','CFT','11','美國著作權','12','顧問及法務',plq01)" & _
                  ",decode(plq02,'D','部門','P','個人',''),plq03,plq01" & _
                  " FROM pricelistquery order by plq01,plq02"
         If rsTmp.State = 1 Then rsTmp.Close
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         Set GRD1.Recordset = rsTmp
         SetGrd
         If rsTmp.RecordCount = 0 Then
            GRD1.Rows = 2
            GRD1.row = 1
            GRD1.col = 0
            MsgBox "無此資料", vbOKOnly, "查詢資料"
         End If
   End Select
   
   For i = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(i, 1) <> "" And GRD1.TextMatrix(i, 2) <> "" Then
         strText = Trim(GRD1.TextMatrix(i, 2))
         If Left(strText, 1) = "," Then strText = Mid(strText, 2)
         If Right(strText, 1) = "," Then strText = Left(strText, Len(strText) - 1)
         strText = Replace(strText, ",", "','")
         If GRD1.TextMatrix(i, 1) = "部門" Then
            strExc(0) = "SELECT A0901,A0902 From ACC090" & _
                        " WHERE A0901 in('" & strText & "') AND A0902 is not null order by A0901"
         ElseIf GRD1.TextMatrix(i, 1) = "個人" Then
            strExc(0) = "SELECT st01,st02 FROM STAFF" & _
                        " WHERE st01 in('" & strText & "') AND st02 is not null order by st03,st01"
         End If
         intI = 1: strTemp = ""
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               strTemp = strTemp & "," & RsTemp.Fields(1)
               RsTemp.MoveNext
            Loop
            strTemp = Mid(strTemp, 2)
            GRD1.TextMatrix(i, 2) = strTemp
         End If
      End If
   Next i
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from pricelistquery where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_PLQ = rsA.Fields.Count
   SetGrd
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

Private Sub Form_Load()
   ReDim m_FieldList(tf_PLQ) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   
   ClearField
   Call SetComboData '組下拉式選單
   
'   InitialField
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
End Sub

Private Sub SetComboData()
Dim Rs As New ADODB.Recordset
   
   '部門
   Me.CboDept.Clear
   If Rs.State <> adStateClosed Then Rs.Close
   Rs.CursorLocation = adUseClient
   '除電腦中心及人事處外,其他人只能看到有在職員工的部門(王副總提需求江總同意)
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M21" Then
      Rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' and substr(A0901,1,1) not in('S','D') Order By A0901", _
               cnnConnection, adOpenStatic, adLockReadOnly
   Else
      Rs.Open "Select A0901,A0902 From ACC090 Where a0904 <> 'Y' and substr(A0901,1,1) not in('S','D') and a0901<>'P29' and a0901 in (select distinct st03 from staff where st04='1' and st01>'6' and substr(st01,1,1)<'G' and substr(st01,4,1)<>'9') Order By A0901", _
               cnnConnection, adOpenStatic, adLockReadOnly
   End If
   Me.CboDept.AddItem ""
   While Not Rs.EOF
      Me.CboDept.AddItem Trim(Rs.Fields(0).Value) & " " & Rs.Fields(1).Value
      Rs.MoveNext
   Wend
   
   '個人
   Me.cboUser.Clear
   If Rs.State <> adStateClosed Then Rs.Close
   Rs.CursorLocation = adUseClient
   Rs.Open "Select st01,st02 From staff Where st04='1' and substr(st15,1,1) not in('S','D') and st01>'63' and st01<'F' and substr(st01,4,1)<>'9' and st01 not in('96029','96030') order by st03,st01", _
            cnnConnection, adOpenStatic, adLockReadOnly
   Me.cboUser.AddItem ""
   While Not Rs.EOF
      Me.cboUser.AddItem Trim(Rs.Fields(0).Value) & " " & Rs.Fields(1).Value
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090635 = Nothing
End Sub

Private Sub GRD1_DblClick()
Dim tmpMouseRow
   
   tmpMouseRow = GRD1.row
   If tmpMouseRow <> 0 Then
      GRD1.row = tmpMouseRow
      GRD1.col = 0
      If GRD1.CellBackColor = &HFFC0C0 Then
         Text1(1).Text = GRD1.TextMatrix(tmpMouseRow, 3)
         QueryRecord
      End If
   End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow GRD1, x, y, nCol, nRow
   If nRow < 0 Then nRow = 0
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j

   tmpMouseRow = GRD1.row
   If tmpMouseRow <> 0 Then
      GRD1.row = tmpMouseRow
      GRD1.col = 0
      If GRD1.CellBackColor <> &HFFC0C0 Then
         GRD1.Visible = False
         For j = 1 To GRD1.Rows - 1
            GRD1.row = j
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               GRD1.CellBackColor = QBColor(15)
            Next i
         Next j
         GRD1.row = tmpMouseRow
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
         GRD1.Visible = True
      End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 1 Then
      If GRD1.Rows - 1 >= 1 Then
         If GRD1.TextMatrix(1, 0) <> "" Then '有查出資料時
            If Text1(0) <> "" Then
               Call cmdOK_Click(0) '類別查詢
            Else
               Call cmdOK_Click(1) '全部查詢
            End If
         End If
      End If
   End If
End Sub

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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If Me.Text1(1).Enabled = True Then
      Cancel = False
      Text1_Validate 1, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

'' 設定欄位的內容
'Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
'Dim nIndex As Integer
'
'   For nIndex = 0 To tf_PLQ - 1
'      If strName = m_FieldList(nIndex).fiName Then
'         If strData = "#==#" Then
'            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
'         Else
'            m_FieldList(nIndex).fiNewData = strData
'         End If
'         Exit For
'      End If
'   Next nIndex
'End Sub
'
'' 從記錄中更新欄位內容
'Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
'Dim nIndex As Integer
'Dim strTmp As String
'
'   For nIndex = 0 To tf_PLQ - 1
'      If m_FieldList(nIndex).fiName <> Empty Then
'         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
'            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
'            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
'         Else
'            m_FieldList(nIndex).fiOldData = Empty
'            m_FieldList(nIndex).fiNewData = Empty
'         End If
'      End If
'   Next nIndex
'EXITSUB:
'End Sub

' 新增記錄
Private Function AddRecord() As Boolean
Dim strPLQ01 As String, strPLQ02 As String, strPLQ03 As String
   
   AddRecord = False
   
   strPLQ01 = Text1(1)
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   If TextPLQ03_D <> "" Then
      strPLQ02 = "D"
      strPLQ03 = TextPLQ03_D
      If IsRecordExist(strPLQ01, strPLQ02) = False Then '檢查記錄是否已存在,否時,則新增資料
         strSql = "INSERT INTO pricelistquery(plq01,plq02,plq03)" & _
                  " VALUES(" & CNULL(strPLQ01) & "," & CNULL(strPLQ02) & "," & CNULL(strPLQ03) & ")"
         cnnConnection.Execute strSql
      End If
   End If
   If TextPLQ03_P <> "" Then
      strPLQ02 = "P"
      strPLQ03 = TextPLQ03_P
      If IsRecordExist(strPLQ01, strPLQ02) = False Then '檢查記錄是否已存在,否時,則新增資料
         strSql = "INSERT INTO pricelistquery(plq01,plq02,plq03)" & _
                  " VALUES(" & CNULL(strPLQ01) & "," & CNULL(strPLQ02) & "," & CNULL(strPLQ03) & ")"
         cnnConnection.Execute strSql
      End If
   End If
   
   If (strPLQ01 < m_FirstKEY(0)) Or (strPLQ01 > m_LastKEY(0)) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   ShowCurrRecord strPLQ01
   AddRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim strPLQ01 As String, strPLQ02 As String, strPLQ03 As String
       
   ModRecord = False
   
   strPLQ01 = m_CurrKEY(0)
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   If TextPLQ03_D <> "" Then
      strPLQ02 = "D"
      strPLQ03 = TextPLQ03_D
      If IsRecordExist(strPLQ01, strPLQ02) = False Then '檢查記錄是否已存在,否時,則新增資料
         strSql = "INSERT INTO pricelistquery(plq01,plq02,plq03)" & _
                  " VALUES(" & CNULL(strPLQ01) & "," & CNULL(strPLQ02) & "," & CNULL(strPLQ03) & ")"
         cnnConnection.Execute strSql
      Else
         If m_PLQ03_D <> strPLQ03 Then
            strSql = "update pricelistquery set plq03='" & strPLQ03 & "'" & _
                     " where plq01='" & strPLQ01 & "' and plq02='" & strPLQ02 & "'"
            cnnConnection.Execute strSql
         End If
      End If
   Else
      strPLQ02 = "D"
      If IsRecordExist(strPLQ01, strPLQ02) = True Then '檢查記錄是否已存在,有時,則刪除資料
         strSql = "delete from pricelistquery where plq01='" & strPLQ01 & "' and plq02='" & strPLQ02 & "'"
         cnnConnection.Execute strSql
      End If
   End If
   If TextPLQ03_P <> "" Then
      strPLQ02 = "P"
      strPLQ03 = TextPLQ03_P
      If IsRecordExist(strPLQ01, strPLQ02) = False Then '檢查記錄是否已存在,否時,則新增資料
         strSql = "INSERT INTO pricelistquery(plq01,plq02,plq03)" & _
                  " VALUES(" & CNULL(strPLQ01) & "," & CNULL(strPLQ02) & "," & CNULL(strPLQ03) & ")"
         cnnConnection.Execute strSql
      Else
         If m_PLQ03_P <> strPLQ03 Then
            strSql = "update pricelistquery set plq03='" & strPLQ03 & "'" & _
                     " where plq01='" & strPLQ01 & "' and plq02='" & strPLQ02 & "'"
            cnnConnection.Execute strSql
         End If
      End If
   Else
      strPLQ02 = "P"
      If IsRecordExist(strPLQ01, strPLQ02) = True Then '檢查記錄是否已存在,有時,則刪除資料
         strSql = "delete from pricelistquery where plq01='" & strPLQ01 & "' and plq02='" & strPLQ02 & "'"
         cnnConnection.Execute strSql
      End If
   End If
   
   cnnConnection.CommitTrans

   ShowCurrRecord strPLQ01
      
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strDel As String, strPLQ01 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   If Text1(1) <> m_CurrKEY(0) Then
      MsgBox "系統記錄的目前系統類別（" & m_CurrKEY(0) & "）與畫面上的系統類別不一致，請重新確認！"
      Exit Function
   End If
   
   strDel = ""
   If TextPLQ03_D <> "" And TextPLQ03_P <> "" Then
      strDel = InputBox("確定要刪除資料嗎？" & vbCrLf & _
                        "1.只刪除查詢部門的權限。" & vbCrLf & _
                        "2.只刪除查詢個人的權限。" & vbCrLf & _
                        "3.刪除二者（全部）。")
      If strDel = "" Then Exit Function
   Else
      If MsgBox("是否要刪除此筆資料？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
         Exit Function
      End If
   End If
   
   cnnConnection.BeginTrans
   
   strPLQ01 = m_CurrKEY(0)
   
   If strDel = "" Or strDel = "3" Then
      strSql = "DELETE FROM pricelistquery WHERE plq01='" & strPLQ01 & "'"
   ElseIf strDel = "1" Then
      strSql = "DELETE FROM pricelistquery WHERE plq01='" & strPLQ01 & "' and plq02='D'"
   ElseIf strDel = "2" Then
      strSql = "DELETE FROM pricelistquery WHERE plq01='" & strPLQ01 & "' and plq02='P'"
   End If
   cnnConnection.Execute strSql
   
   If strPLQ01 = m_LastKEY(0) Or strPLQ01 = m_FirstKEY(0) Then
      RefreshRange
   End If
   ShowCurrRecord strPLQ01
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strPLQ01 As String
   
   QueryRecord = False
   strPLQ01 = Text1(1)
   If IsRecordExist(strPLQ01) = True Then
      m_CurrKEY(0) = strPLQ01
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
      ClearField
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse

   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
'            UpdateFieldNewData
            If AddRecord = True Then
                RefreshRange
            Else
                Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '修改
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
'            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If Text1(1) <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            If Text1(1) = "" Then
               MsgBox "請輸入系統類別才可進行查詢動作！", vbInformation
               Text1(1).SetFocus
               GoTo EXITSUB
            End If
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: If Me.Visible = True Then Text1(1).SetFocus
      'Case 2: If Me.Visible = True Then cboDept.SetFocus
      Case 4: If Me.Visible = True Then Text1(1).SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, Optional strKEY02 As String = "") As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM pricelistquery WHERE plq01='" & strKEY01 & "'" & _
            IIf(strKEY02 <> "", " and plq02='" & strKEY02 & "'", "")
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
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

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
   Else
      strSql = "SELECT plq01 FROM pricelistquery WHERE plq01='" & m_CurrKEY(0) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("plq01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("plq01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT plq01 FROM pricelistquery" & _
               " WHERE plq01 = (SELECT min(plq01) FROM pricelistquery)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("plq01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("plq01")
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
   
   strSql = "SELECT plq01 FROM pricelistquery" & _
            "  WHERE plq01 = (SELECT max(plq01) FROM pricelistquery WHERE plq01<'" & m_CurrKEY(0) & "')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("plq01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("plq01")
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
   
   strSql = "SELECT plq01 FROM pricelistquery" & _
            " WHERE plq01 = (SELECT min(plq01) FROM pricelistquery WHERE plq01>'" & m_CurrKEY(0) & "')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("plq01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("plq01")
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

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
'         strTit = "詢問"
'         strMsg = "是否要刪除此筆資料?"
'         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
'         If nResponse = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
'         End If
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
'         UpdateFieldNewData
         If OnWork = True Then
            Me.SSTab1.TabEnabled(1) = True
            UpdateToolbarState
         Else
            Exit Sub
         End If
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  Me.SSTab1.TabEnabled(1) = True
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               Me.SSTab1.TabEnabled(1) = True
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT plq01 FROM pricelistquery " & _
            "WHERE plq01 = (SELECT Min(plq01) FROM pricelistquery) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("plq01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("plq01")
   End If
   rsTmp.Close
   
   strSql = "SELECT plq01 FROM pricelistquery " & _
            "WHERE plq01 = (SELECT Max(plq01) FROM pricelistquery) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("plq01")) = False Then: m_LastKEY(0) = rsTmp.Fields("plq01")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCName As String, strCDate As String, strCTime As String
Dim strUName As String, strUDate As String, strUTime As String
   
   strSql = "SELECT * FROM pricelistquery WHERE plq01='" & m_CurrKEY(0) & "' order by plq05 asc, plq08 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         If IsNull(rsTmp.Fields("plq01")) = False Then: Text1(1) = rsTmp.Fields("plq01")
         If rsTmp.Fields("plq02") = "D" Then '部門
            If IsNull(rsTmp.Fields("plq03")) = False Then: TextPLQ03_D = rsTmp.Fields("plq03"): m_PLQ03_D = rsTmp.Fields("plq03")
            SetList lstDept, TextPLQ03_D, "D"
         Else '個人
            If IsNull(rsTmp.Fields("plq03")) = False Then: TextPLQ03_P = rsTmp.Fields("plq03"): m_PLQ03_P = rsTmp.Fields("plq03")
            SetList lstUsers, TextPLQ03_P, "P"
         End If
         If strCName = "" Then
            If IsNull(rsTmp.Fields("plq04")) = False Then: strCName = rsTmp.Fields("plq04")
            If IsNull(rsTmp.Fields("plq05")) = False Then: strCDate = rsTmp.Fields("plq05")
            If IsNull(rsTmp.Fields("plq06")) = False Then: strCTime = rsTmp.Fields("plq06")
         Else
            If Val(rsTmp.Fields("plq05")) < Val(strCDate) Or _
               (Val(rsTmp.Fields("plq05")) = Val(strCDate) And Val(rsTmp.Fields("plq06")) < Val(strCTime)) Then
               If IsNull(rsTmp.Fields("plq04")) = False Then: strCName = rsTmp.Fields("plq04")
               If IsNull(rsTmp.Fields("plq05")) = False Then: strCDate = rsTmp.Fields("plq05")
               If IsNull(rsTmp.Fields("plq06")) = False Then: strCTime = rsTmp.Fields("plq06")
            End If
         End If
         If strUName = "" Then
            If IsNull(rsTmp.Fields("plq07")) = False Then: strUName = rsTmp.Fields("plq07")
            If IsNull(rsTmp.Fields("plq08")) = False Then: strUDate = rsTmp.Fields("plq08")
            If IsNull(rsTmp.Fields("plq09")) = False Then: strUTime = rsTmp.Fields("plq09")
         Else
            If Val("" & rsTmp.Fields("plq08")) > Val(strUDate) Or _
               (Val("" & rsTmp.Fields("plq08")) = Val(strUDate) And Val("" & rsTmp.Fields("plq09")) > Val(strUTime)) Then
               If IsNull(rsTmp.Fields("plq07")) = False Then: strUName = rsTmp.Fields("plq07")
               If IsNull(rsTmp.Fields("plq08")) = False Then: strUDate = rsTmp.Fields("plq08")
               If IsNull(rsTmp.Fields("plq09")) = False Then: strUTime = rsTmp.Fields("plq09")
            End If
         End If
         rsTmp.MoveNext
      Loop
      
      ' 更新CUID
      UpdateCUID strCName, strCDate, strCTime, strUName, strUDate, strUTime
'      ' 更新暫存區的資料
'      UpdateFieldOldData rsTmp
      SSTab1.Tab = 0
   End If
   
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(strCName As String, strCDate As String, strCTime As String, _
                       strUName As String, strUDate As String, strUTime As String)
   
   If strCName <> "" Then
      strCName = GetStaffName(strCName, True)
      strCDate = Format(TAIWANDATE(strCDate), "###/##/##")
      strCTime = Format(strCTime, "##:##:##")
   End If
   If strUName <> "" Then
      strUName = GetStaffName(strUName, True)
      strUDate = Format(TAIWANDATE(strUDate), "###/##/##")
      strUTime = Format(strUTime, "##:##:##")
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Sub SetList(oList As Object, p_stData As String, strKind As String)
   oList.Clear
   If p_stData <> "" Then
      If Left(p_stData, 1) = "," Then p_stData = Mid(p_stData, 2)
      If Right(p_stData, 1) = "," Then p_stData = Left(p_stData, Len(p_stData) - 1)
      p_stData = Replace(p_stData, ",", "','")
      If strKind = "D" Then '部門
         strExc(0) = "SELECT A0901,A0902 From ACC090" & _
                     " WHERE A0901 in('" & p_stData & "') AND A0902 is not null order by A0901"
      Else '個人
         strExc(0) = "SELECT st01,st02 FROM STAFF" & _
                     " WHERE st01 in('" & p_stData & "') AND st02 is not null order by st03,st01"
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            oList.AddItem RsTemp.Fields(0) & " " & RsTemp.Fields(1), 0
            RsTemp.MoveNext
         Loop
      End If
   End If
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
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
         ' 新增
      Case 1, 2, 3, 4:
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

Private Function CheckDataValid() As Boolean
   CheckDataValid = False
   
   If Text1(1).Text = "" Then
      MsgBox "系統類別不可以空白！", vbExclamation
      Text1(1).SetFocus
      Exit Function
   End If
   
   If TextPLQ03_D = "" And TextPLQ03_P = "" Then
      MsgBox "請設定權限！", vbExclamation
      CboDept.SetFocus
      Exit Function
   End If
   
   CheckDataValid = True
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   Text1(1).Enabled = Not bEnable
   If bEnable Then Text1(1).BackColor = &H8000000F Else Text1(1).BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   Text1(1).Enabled = Not bEnable
   If bEnable Then Text1(1).BackColor = &H8000000F Else Text1(1).BackColor = &H80000005
   CboDept.Enabled = Not bEnable
   cmdAddDept.Enabled = Not bEnable
   cmdRemoveDept.Enabled = Not bEnable
   cboUser.Enabled = Not bEnable
   cmdAdd.Enabled = Not bEnable
   cmdRemove.Enabled = Not bEnable
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   Text1(1) = Empty
   TextPLQ03_D = Empty: m_PLQ03_D = Empty
   TextPLQ03_P = Empty: m_PLQ03_P = Empty
   lstDept.Clear
   lstUsers.Clear
   CboDept.Text = Empty
   cboUser.Text = Empty
   Label23 = Empty
   
   SetGrd
   For nIndex = 0 To tf_PLQ - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

'Private Sub UpdateFieldNewData()
'Dim MyArr As Variant
'   '若新增資料
'   If m_EditMode = 1 Then
'      SetFieldNewData "PLQ01", Text1(1)
'      SetFieldNewData "PLQ02", "D"
'   End If
'   SetFieldNewData "PLQ03", TextPLQ03_D
'End Sub
'
'' 初始化欄位陣列
'Private Sub InitialField()
'Dim nIndex As Integer
'Dim strTmp As String
'   ' 初始化欄位陣列
'   For nIndex = 1 To tf_PLQ
'      strTmp = Format(nIndex, "00")
'      m_FieldList(nIndex - 1).fiName = "PLQ" & strTmp
'      m_FieldList(nIndex - 1).fiOldData = Empty
'      m_FieldList(nIndex - 1).fiNewData = Empty
'      m_FieldList(nIndex - 1).fiType = 0 '文字型態
'      Select Case nIndex
'         Case 2, 3, 4, 5, 6, 7:
'            m_FieldList(nIndex - 1).fiType = 1 '數值型態
'      End Select
'   Next nIndex
'End Sub

'帶預設資料
Private Sub InitialData()
   SetGrd
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("系統類別", "種類", "可查詢部門或個人", "plq01")
   arrGridHeadWidth = Array(1200, 500, 6800, 0)
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

Private Sub Text1_GotFocus(Index As Integer)
   InverseTextBox Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   If Trim(Text1(Index)) <> "" And Len(Trim(Text1(Index))) <> 2 Then
      Text1(Index) = Right("00" & Trim(Text1(Index)), 2)
   End If
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Text1(Index) <> "" Then
      If Val(Text1(Index)) > 12 Or Val(Text1(Index)) < 1 Then
         MsgBox "系統類別只可以輸入01 ~ 12！", vbInformation
         Call Text1_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      If m_EditMode = 1 Then
         If IsRecordExist(Text1(1)) = True And Text1(1).Enabled = True And Text1(1).Locked = False Then
            MsgBox "該系統類別已有資料，請修改！", vbInformation
            Call Text1_GotFocus(1)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

'新增開發人員
Private Sub cmdAdd_Click()
   If InStr(cboUser, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
      cboUser.SetFocus
      Exit Sub
   End If
   AddLstFrmCbo cboUser, lstUsers
   TextPLQ03_P = ComposeList(lstUsers)
   'cboUser.SetFocus
End Sub

'移除開發人員
Private Sub cmdRemove_Click()
   RemoveList lstUsers
   TextPLQ03_P = ComposeList(lstUsers)
End Sub

'新增部門
Private Sub cmdAddDept_Click()
   If InStr(CboDept, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請改用其他符號！", vbExclamation
      CboDept.SetFocus
      Exit Sub
   End If
   AddLstFrmCbo CboDept, lstDept
   TextPLQ03_D = ComposeList(lstDept)
   'cboDept.SetFocus
End Sub

'移除部門
Private Sub cmdRemoveDept_Click()
   RemoveList lstDept
   TextPLQ03_D = ComposeList(lstDept)
End Sub

Private Function ComposeList(oList As Object) As String
Dim varTemp As Variant
   
   strExc(1) = ""
   If oList.ListCount > 0 Then
      varTemp = Split(oList.List(0), " ")
      strExc(1) = varTemp(0)
      For intI = 1 To oList.ListCount - 1
         varTemp = Split(oList.List(intI), " ")
         strExc(1) = strExc(1) & "," & varTemp(0)
      Next
   End If
   ComposeList = strExc(1)
End Function

Private Sub AddLstFrmCbo(oCombo As Object, oList As Object)
   Dim idx As Integer, bFound As Boolean
   
   If oCombo <> "" Then
      For idx = 0 To oList.ListCount - 1
         If oList.List(idx) = oCombo Then
            MsgBox "資料已存在！"
            oCombo.SetFocus
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         oList.AddItem oCombo, 0
         oCombo = ""
      End If
   End If
End Sub

Private Sub RemoveList(oList As Object)
   Dim idx As Integer, ii As Integer
   If oList.ListCount > 0 Then
      ii = 0
      For idx = 0 To oList.ListCount - 1
         If oList.Selected(ii) = True Then
            oList.Selected(ii) = False 'Add By Sindy 2021/6/1
            oList.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Next
   End If
End Sub

Private Sub cboDept_GotFocus()
   If CboDept.Locked = False Then
      CloseIme
      'SendMessage cboDept.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

Private Sub CboDept_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cbouser_GotFocus()
   If cboUser.Locked = False Then
      CloseIme
      'SendMessage cboUser.hWnd, CB_SHOWDROPDOWN, 1, 0
   End If
End Sub

Private Sub cboUser_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
