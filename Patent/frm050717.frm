VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050717 
   BorderStyle     =   1  '單線固定
   Caption         =   "CF代理人報價附件維護"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8880
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8880
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   60
      TabIndex        =   16
      Top             =   660
      Width           =   8745
      _ExtentX        =   15431
      _ExtentY        =   8911
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm050717.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label23"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LabelCQ01"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(17)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "textCQ03"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CommonDialog1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Winsock1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textCQ04"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdOpenAtt"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdAddAtt"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdRemAtt"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lstAtt"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textCQ01"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCQ02"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textCQ11"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmd1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm050717.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "Line4"
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(3)=   "Line3"
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(5)=   "Line1"
      Tab(1).Control(6)=   "GRD1"
      Tab(1).Control(7)=   "txt1(0)"
      Tab(1).Control(8)=   "txt1(1)"
      Tab(1).Control(9)=   "txt1(2)"
      Tab(1).Control(10)=   "txt1(3)"
      Tab(1).Control(11)=   "cmdok"
      Tab(1).Control(12)=   "txt1(4)"
      Tab(1).Control(13)=   "txt1(5)"
      Tab(1).ControlCount=   14
      Begin VB.CommandButton cmd1 
         Caption         =   "搬檔"
         Height          =   375
         Left            =   7560
         TabIndex        =   29
         Top             =   4080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox textCQ11 
         Height          =   270
         Left            =   1320
         TabIndex        =   28
         Top             =   3960
         Visible         =   0   'False
         Width           =   6315
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   5
         Left            =   -72690
         MaxLength       =   4
         TabIndex        =   12
         Top             =   780
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   4
         Left            =   -73740
         MaxLength       =   4
         TabIndex        =   11
         Top             =   780
         Width           =   915
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Default         =   -1  'True
         Height          =   315
         Left            =   -68130
         TabIndex        =   13
         Top             =   750
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -69270
         MaxLength       =   7
         TabIndex        =   10
         Top             =   450
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -70320
         MaxLength       =   7
         TabIndex        =   9
         Top             =   450
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72480
         MaxLength       =   9
         TabIndex        =   8
         Top             =   450
         Width           =   1125
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73740
         MaxLength       =   9
         TabIndex        =   7
         Top             =   450
         Width           =   1125
      End
      Begin VB.TextBox textCQ02 
         Height          =   270
         Left            =   1305
         MaxLength       =   7
         TabIndex        =   1
         Top             =   750
         Width           =   1485
      End
      Begin VB.TextBox textCQ01 
         Height          =   270
         Left            =   1305
         MaxLength       =   9
         TabIndex        =   0
         Top             =   420
         Width           =   1485
      End
      Begin VB.ListBox lstAtt 
         Height          =   1848
         ItemData        =   "frm050717.frx":0038
         Left            =   1305
         List            =   "frm050717.frx":003F
         MultiSelect     =   2  '進階多重選取
         Sorted          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2130
         Width           =   6540
      End
      Begin VB.CommandButton cmdRemAtt 
         Caption         =   "-> 移除"
         Height          =   255
         Left            =   7860
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2730
         Width           =   735
      End
      Begin VB.CommandButton cmdAddAtt 
         Caption         =   "<- 新增"
         Height          =   255
         Left            =   7860
         TabIndex        =   5
         Top             =   2430
         Width           =   735
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   255
         Left            =   7860
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2130
         Width           =   735
      End
      Begin VB.TextBox textCQ04 
         Height          =   270
         Left            =   210
         TabIndex        =   17
         Top             =   2430
         Visible         =   0   'False
         Width           =   1035
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   210
         Top             =   2850
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   210
         Top             =   3300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm050717.frx":004B
         Height          =   3795
         Left            =   -74880
         TabIndex        =   14
         Top             =   1140
         Width           =   8490
         _ExtentX        =   14965
         _ExtentY        =   6689
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "代理人編號|名稱|日期|內容"
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
         _Band(0).Cols   =   4
      End
      Begin MSForms.TextBox textCQ03 
         Height          =   990
         Left            =   1305
         TabIndex        =   2
         Top             =   1080
         Width           =   6540
         VariousPropertyBits=   -1467989989
         MaxLength       =   4000
         ScrollBars      =   2
         Size            =   "11536;1746"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line1 
         X1              =   -73050
         X2              =   -72300
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "代理人國籍："
         Height          =   180
         Left            =   -74850
         TabIndex        =   27
         Top             =   810
         Width           =   1080
      End
      Begin VB.Line Line3 
         X1              =   -69690
         X2              =   -68940
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   -70920
         TabIndex        =   26
         Top             =   480
         Width           =   540
      End
      Begin VB.Line Line4 
         X1              =   -72780
         X2              =   -72180
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "代理人編號："
         Height          =   180
         Left            =   -74850
         TabIndex        =   25
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日　期："
         Height          =   180
         Index           =   17
         Left            =   525
         TabIndex        =   24
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人編號："
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   23
         Top             =   465
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "內　容："
         Height          =   180
         Index           =   2
         Left            =   525
         TabIndex        =   22
         Top             =   1140
         Width           =   720
      End
      Begin MSForms.Label LabelCQ01 
         Height          =   240
         Left            =   2820
         TabIndex        =   21
         Top             =   450
         Width           =   5325
         VariousPropertyBits=   27
         Size            =   "9393;423"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(民國年月日)"
         Height          =   180
         Index           =   0
         Left            =   2820
         TabIndex        =   20
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "附　件："
         Height          =   180
         Index           =   3
         Left            =   525
         TabIndex        =   19
         Top             =   2190
         Width           =   720
      End
      Begin MSForms.Label Label23 
         Height          =   225
         Left            =   150
         TabIndex        =   18
         Top             =   4290
         Width           =   8100
         VariousPropertyBits=   27
         Caption         =   "CREATE : 　　　  101/09/03  13:54:00          UPDATE : 　　　  101/09/04  09:21:44"
         Size            =   "14287;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
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
            Picture         =   "frm050717.frx":0060
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050717.frx":037C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050717.frx":0698
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050717.frx":0874
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050717.frx":0B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050717.frx":0EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050717.frx":11C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050717.frx":14E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050717.frx":1800
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050717.frx":1B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050717.frx":1E38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
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
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frm050717"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/4 改成Form2.0 (GRD1,textCQ03,LabelCQ01,Label23)
'Create By Sindy 2012/12/19
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
' 第一筆資料的Key值
Dim m_FirstKEY(2) As String
' 最後一筆資料的Key值
Dim m_LastKEY(2) As String
' 目前正在顯示的Key值
Dim m_CurrKEY(2) As String
Dim rsA As New ADODB.Recordset
Dim tf_CQ As Integer
Dim strText As String, arrKey As Variant
Dim m_CQ05 As String

Private Const cTableName As String = "CFQUOTATION" 'Added by Lydia 2017/08/09 指定FTP資料夾名稱

'Add By Sindy 2014/12/3
Private Sub cmdOK_Click()
   If txt1(0) & txt1(1) & txt1(2) & txt1(3) & txt1(4) & txt1(5) <> "" Then
      If txt1(0) <> "" Then txt1(0) = Left(txt1(0) & "00000000", 9)
      If txt1(1) <> "" Then txt1(1) = Left(txt1(1) & "00000000", 9)
      '代理人(起)
      If Len(Me.txt1(0).Text) > 0 Then
         If Left(Me.txt1(0).Text, 1) <> "Y" Then
            MsgBox "代理人代碼輸入錯誤！", vbExclamation, "輸入錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
      End If
      '代理人(迄)
      If Len(Me.txt1(1).Text) > 0 Then
         If Left(Me.txt1(1).Text, 1) <> "Y" Then
            MsgBox "代理人代碼輸入錯誤！", vbExclamation, "輸入錯誤！"
            txt1(1).SetFocus
            Exit Sub
         End If
      End If
      If RunNick(txt1(0), txt1(1)) Then
         txt1(0).SetFocus
         Exit Sub
      End If
      '日期
      If RunNick2(txt1(2), txt1(3)) Then
         txt1(2).SetFocus
         Exit Sub
      End If
      '國籍
      If RunNick(txt1(4), txt1(5)) Then
         txt1(4).SetFocus
         Exit Sub
      End If
      GetData
   Else
      MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
      txt1(0).SetFocus
   End If
End Sub

'Add By Sindy 2014/12/3
Sub GetData()
Dim rsTmp As New ADODB.Recordset
Dim strCol As String
   
   strCol = ""
   If txt1(0) <> "" Then
       strCol = strCol & " and CQ01>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       strCol = strCol & " and CQ01<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       strCol = strCol & " and CQ02>='" & DBDATE(txt1(2)) & "' "
   End If
   If txt1(3) <> "" Then
       strCol = strCol & " and CQ02<='" & DBDATE(txt1(3)) & "' "
   End If
   If txt1(4) <> "" Then
       strCol = strCol & " and FA10>='" & txt1(4) & "' "
   End If
   If txt1(5) <> "" Then
       strCol = strCol & " and FA10<='" & txt1(5) & "' "
   End If
   '抓取資料
   'Modified by Lydia 2017/08/09 +CQ11
   strSql = "SELECT CQ01,NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),sqldatet(CQ02),CQ03,FA10,CQ11" & _
            " FROM CFQuotation,Fagent" & _
            " WHERE substr(CQ01,1,8)=FA01(+) AND substr(CQ01,9)=FA02(+)" & strCol & _
            " ORDER BY CQ01,CQ02 asc"
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set grd1.Recordset = rsTmp
   SetGrd
End Sub

'Add By Sindy 2014/12/3
Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modified by Lydia 2017/08/09 +CQ11
   'arrGridHeadText = Array("代理人編號", "名稱", "日期", "內容")
   'arrGridHeadWidth = Array(1000, 3000, 850, 3000)
   arrGridHeadText = Array("代理人編號", "名稱", "日期", "內容", "CQ11")
   arrGridHeadWidth = Array(1000, 3000, 850, 3000, 0)
   'end 2017/08/09
   grd1.Visible = False
   grd1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1.Cols - 1
      grd1.row = 0
      grd1.col = iRow
      grd1.Text = arrGridHeadText(iRow)
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1.CellAlignment = flexAlignCenterCenter
   Next
   grd1.Visible = True
End Sub

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from CFQuotation where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_CQ = rsA.Fields.Count
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
   ReDim m_FieldList(tf_CQ) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   
   InitialField
   RefreshRange
   ShowFirstRecord
   SetCtrlReadOnly True
   UpdateToolbarState
   SSTab1.Tab = 0 'Add By Sindy 2014/12/3
   
   'Added by Lydia 2017/08/09
   If Pub_StrUserSt03 <> "M51" Then Cmd1.Visible = False
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   PUB_KillTempFile "$$*.*" 'Added by Lydia 2017/08/09 清除暫存檔
   
   Set frm050717 = Nothing
End Sub

'Add By Sindy 2014/12/3
Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow grd1, x, y, nCol, nRow
   grd1.col = nCol
   grd1.row = nRow
End Sub

'Add By Sindy 2014/12/3
Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j
   
   grd1.Visible = False
   tmpMouseRow = grd1.row
   If tmpMouseRow <> 0 Then
      grd1.row = tmpMouseRow
      grd1.col = 0
      If grd1.CellBackColor <> &HFFC0C0 Then
         For j = 1 To grd1.Rows - 1
            grd1.row = j
            For i = 0 To grd1.Cols - 1
               grd1.col = i
               grd1.CellBackColor = QBColor(15)
            Next i
         Next j
         grd1.row = tmpMouseRow
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = &HFFC0C0
         Next i
         textCQ01.Text = grd1.TextMatrix(tmpMouseRow, 0)
         textCQ02.Text = ChangeTDateStringToTString(grd1.TextMatrix(tmpMouseRow, 2))
         m_CurrKEY(0) = textCQ01
         m_CurrKEY(1) = DBDATE(textCQ02)
         QueryRecord
      End If
   End If
   grd1.Visible = True
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim tmpArr1 As Variant, tmpArr2 As Variant 'Added by Lydia 2017/08/09

   TxtValidate = False
   
   'Added by Morgan 2022/1/4 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2022/1/4
   
   If Me.textCQ01.Enabled = True Then
      Cancel = False
      textCQ01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCQ02.Enabled = True Then
      Cancel = False
      textCQ02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Added by Lydia 2017/08/09 檢查長度
   If CheckLengthIsOK(Me.textCQ04, 600, False) = False Then
       MsgBox "全部的附件檔名超過最大長度600字元！" & vbCrLf & "(1個中文=2個字元)", vbCritical
       Exit Function
   End If
   strExc(1) = "附件順序有誤，請全部移除後再新增附件"
   If (textCQ04 = "" And textCQ11 <> "") Or (textCQ04 <> "" And textCQ11 = "") Then
       ShowMsg strExc(1)
       Exit Function
   End If
    
   tmpArr1 = Empty: tmpArr2 = Empty
   tmpArr1 = Split(textCQ04, ",")
   tmpArr2 = Split(textCQ11, ",")
   If UBound(tmpArr1) <> UBound(tmpArr2) Then
       ShowMsg strExc(1)
       Exit Function
   End If
    
   '預估一個ftp路徑約55字
   If UBound(tmpArr2) > Format(600 / 55, "0") Then
      MsgBox "附件數量超過最大上限(" & Format(600 / 55, "0") & ")！", vbCritical
      Exit Function
   End If
   For intI = 0 To UBound(tmpArr1)
      If (Trim(tmpArr1(intI)) = "" And Trim(tmpArr2(intI)) <> "") Or (Trim(tmpArr1(intI)) <> "" And Trim(tmpArr2(intI)) = "") Then
         ShowMsg strExc(1)
         Exit Function
      End If
   Next intI
   'end 2017/08/09

   
   TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
   
   For nIndex = 0 To tf_CQ - 1
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

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String
   
   For nIndex = 0 To tf_CQ - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim iErr As Integer, sErrMsg As String
   
   AddRecord = False
   
   ' 檢查記錄是否已存在
   If IsRecordExist(textCQ01, textCQ02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO CFQuotation ("
   For nIndex = 0 To tf_CQ - 1
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
   For nIndex = 0 To tf_CQ - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         'Added by Lydia 2017/08/09 跳過FTP路徑
         If nIndex = 10 Then
            strSql = strSql & ",NULL "
         Else
         'end 2017/08/09
            If bFirst = True Then
               strSql = strSql & strTmp
               bFirst = False
            Else
               strSql = strSql & "," & strTmp
            End If
         End If 'end 2017/08/09
      End If
   Next nIndex
   strSql = strSql & ")"
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   'Modify By Sindy 2014/12/3
   If Pub_StrUserSt03 = "M51" Then
      If MsgBox("注意：此程式是直接對FTP（正式目錄）的附件做異動，確定還要繼續嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         GoTo ErrHand
      End If
   End If
   '2014/12/3 END
   
   'Added by Lydia 2017/08/09 判斷移檔日期
   If strSrvDate(1) >= CR_NewDate Then
      If UpdateAttFile(m_FieldList(0).fiNewData & m_FieldList(1).fiNewData, iErr, sErrMsg) = False Then
         GoTo ErrHand
      Else
         If textCQ11.Text <> textCQ11.Tag Then
            strSql = "UPDATE CFQUOTATION SET CQ11='" & textCQ11.Text & "' WHERE CQ01||CQ02='" & m_FieldList(0).fiNewData & m_FieldList(1).fiNewData & "' "
            cnnConnection.Execute strSql
         End If
      End If
   Else
   'end 2017/08/09
        '上傳附件檔
        If UploadAtt(m_FieldList(0).fiNewData, iErr, sErrMsg) = False Then
           GoTo ErrHand
        End If
   End If '2017/08/09
   
   cnnConnection.CommitTrans
   
   If ((textCQ01 & textCQ02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or _
      ((textCQ01 & textCQ02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   
   ShowCurrRecord textCQ01, textCQ02
   AddRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   If iErr > 0 Then
      MsgBox sErrMsg, vbExclamation, "新增失敗！"
   Else
      If Err.NUMBER <> 0 Then MsgBox Err.Description, vbExclamation, "新增失敗！"
   End If
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim iErr As Integer, sErrMsg As String
Dim bolRemove As Boolean, ii As Integer
Dim arrFile1
   
   ModRecord = False
   
   strSql = "begin user_data.user_enabled:=1; UPDATE CFQuotation SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_CQ - 1
      strTmp = Empty
      'If nIndex < 7 Or nIndex > 12 Then
         'Added by Lydia 2017/08/09 跳過FTP路徑
         If nIndex = 10 Then
         Else
            If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
               If m_FieldList(nIndex).fiType = 0 Then
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
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
         End If 'end 2017/08/09
      'End If
   Next nIndex
   
   strSql = strSql & " " & _
            "WHERE CQ01='" & m_CurrKEY(0) & "' and CQ02=" & m_CurrKEY(1) & "; end;"
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      
      'Modify By Sindy 2014/12/3
      If Pub_StrUserSt03 = "M51" Then
         If MsgBox("注意：此程式是直接對FTP（正式目錄）的附件做異動，確定還要繼續嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
            GoTo ErrHand
         End If
      End If
      '2014/12/3 END
      'Added by Lydia 2017/08/09 判斷移檔日期
      If strSrvDate(1) >= CR_NewDate Then
            If UpdateAttFile(m_FieldList(0).fiNewData & m_FieldList(1).fiNewData, iErr, sErrMsg) = False Then
               GoTo ErrHand
            Else
               If textCQ11.Text <> textCQ11.Tag Then
                  strSql = "UPDATE CFQUOTATION SET CQ11='" & textCQ11.Text & "' WHERE CQ01||CQ02='" & m_FieldList(0).fiNewData & m_FieldList(1).fiNewData & "' "
                  cnnConnection.Execute strSql
               End If
            End If
      Else
      'end 2017/08/09
            'Modify By Sindy 2014/12/4 必須先檢查有沒有要刪除的檔案,再檢查有沒有要新增的檔案
            '檔案有異動時，移掉的要刪除
            bolRemove = False
            If m_FieldList(3).fiNewData <> m_FieldList(3).fiOldData Then
               arrFile1 = Split(m_FieldList(3).fiOldData, ",")
               For ii = LBound(arrFile1) To UBound(arrFile1)
                  If InStr(m_FieldList(3).fiNewData & ",", arrFile1(ii) & ",") > 0 Then
                     arrFile1(ii) = ""
                  Else
                     bolRemove = True
                  End If
               Next
               If bolRemove = True Then
                  If RemoveAtt(m_FieldList(0).fiNewData, Join(arrFile1, ","), iErr, sErrMsg) = False Then
                     GoTo ErrHand
                  End If
               End If
            End If
            '上傳附件檔
            If UploadAtt(m_FieldList(0).fiNewData, iErr, sErrMsg) = False Then
               GoTo ErrHand
            End If
      End If 'end 2017/08/09
   End If
   
   cnnConnection.CommitTrans

   ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
      
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   If iErr > 0 Then
      MsgBox sErrMsg, vbExclamation, "修改失敗！"
   Else
      If Err.NUMBER <> 0 Then MsgBox Err.Description, vbExclamation, "修改失敗！"
   End If
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim iErr As Integer, sErrMsg As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   'Added by Lydia 2017/08/09 判斷移檔日期
   If textCQ11 <> "" And strSrvDate(1) >= CR_NewDate Then
      textCQ11.Text = ""
      If UpdateAttFile(m_FieldList(0).fiNewData & m_FieldList(1).fiNewData, iErr, sErrMsg) = False Then
         GoTo ErrHand
      End If
   End If
   'end 2017/08/09
   
   strSql = "DELETE FROM CFQuotation " & _
            "WHERE CQ01 = '" & m_CurrKEY(0) & "'  and CQ02=" & m_CurrKEY(1)
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   'Modify By Sindy 2014/12/3
   If Pub_StrUserSt03 = "M51" Then
      If MsgBox("注意：此程式是直接對FTP（正式目錄）的附件做異動，確定還要繼續嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         GoTo ErrHand
      End If
   End If
   '2014/12/3 END
   
   '刪除附件
   'Modifie by Lydia 2017/08/09 判斷移檔日期之前
   'If m_FieldList(3).fiOldData <> "" Then
   If m_FieldList(3).fiOldData <> "" And strSrvDate(1) < CR_NewDate Then
      If RemoveAtt(m_FieldList(0).fiNewData, m_FieldList(3).fiOldData, iErr, sErrMsg) = False Then
         GoTo ErrHand
      End If
   End If
   
   cnnConnection.CommitTrans
   
   If (m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1)) Or _
      (m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
   DelRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   If iErr > 0 Then
      MsgBox sErrMsg, vbExclamation, "刪除失敗！"
   Else
      If Err.NUMBER <> 0 Then MsgBox Err.Description, vbExclamation, "刪除失敗！"
   End If
End Function

' 查詢記錄
Public Function QueryRecord() As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   QueryRecord = False
   
   If IsRecordExist(textCQ01, textCQ02) = True Then
      m_CurrKEY(0) = textCQ01
      m_CurrKEY(1) = DBDATE(textCQ02)
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
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
            UpdateFieldNewData
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
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textCQ01 <> "" And textCQ02 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            If textCQ01 = "" Or textCQ02 = "" Then
               MsgBox "查詢條件必須全部輸齊，才可進行查詢動作！", vbInformation
            End If
            GoTo EXITSUB
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
      Case 1: If Me.Visible = True Then textCQ01.SetFocus
      Case 2: If Me.Visible = True Then textCQ03.SetFocus
      Case 4: If Me.Visible = True Then textCQ01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM CFQuotation " & _
             "WHERE CQ01='" & strKEY01 & "' and CQ02=" & DBDATE(strKEY02)
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
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "select * from CFQuotation where rownum <2"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount = 0 Then
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   rsTmp.Close
   
   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = DBDATE(strKEY02)
   Else
      strSql = "SELECT CQ01,CQ02 FROM CFQuotation " & _
                "WHERE CQ01='" & m_CurrKEY(0) & "' and CQ02=" & m_CurrKEY(1)
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CQ01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CQ01")
         If IsNull(rsTmp.Fields("CQ02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("CQ02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT min(CQ01||'-'||CQ02) FROM CFQuotation "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         strText = "" & rsTmp.Fields(0)
         If Trim(strText) > "" Then
            arrKey = Split(strText, "-")
            m_CurrKEY(0) = arrKey(0)
            m_CurrKEY(1) = arrKey(1)
         End If
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
   m_CurrKEY(1) = m_FirstKEY(1)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "select max(CQ01||'-'||CQ02) From CFQuotation " & _
             "where CQ01||CQ02<'" & m_CurrKEY(0) & m_CurrKEY(1) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      strText = "" & rsTmp.Fields(0)
      If Trim(strText) > "" Then
         arrKey = Split(strText, "-")
         m_CurrKEY(0) = arrKey(0)
         m_CurrKEY(1) = arrKey(1)
      End If
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "select min(CQ01||'-'||CQ02) From CFQuotation " & _
             "where CQ01||CQ02>'" & m_CurrKEY(0) & m_CurrKEY(1) & m_CurrKEY(2) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      strText = "" & rsTmp.Fields(0)
      If Trim(strText) > "" Then
         arrKey = Split(strText, "-")
         m_CurrKEY(0) = arrKey(0)
         m_CurrKEY(1) = arrKey(1)
      End If
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
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
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         If m_CQ05 <> strUserNum And Pub_StrUserSt03 <> "M51" Then
            MsgBox "無此筆資料的修改權限！", vbExclamation
            Exit Sub
         End If
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
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
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
         If OnWork = True Then
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

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "select * from CFQuotation where rownum <2"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount = 0 Then
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   rsTmp.Close
   
   strSql = "SELECT min(CQ01||'-'||CQ02) FROM CFQuotation "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      strText = "" & rsTmp.Fields(0)
      If Trim(strText) > "" Then
         arrKey = Split(strText, "-")
         m_FirstKEY(0) = arrKey(0)
         m_FirstKEY(1) = arrKey(1)
      End If
   End If
   rsTmp.Close
   
   strSql = "SELECT max(CQ01||'-'||CQ02) FROM CFQuotation "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      strText = "" & rsTmp.Fields(0)
      If Trim(strText) > "" Then
         arrKey = Split(strText, "-")
         m_LastKEY(0) = arrKey(0)
         m_LastKEY(1) = arrKey(1)
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer
   
   strSql = "SELECT * FROM CFQuotation " & _
            "WHERE CQ01='" & m_CurrKEY(0) & "' and CQ02=" & Val(DBDATE(m_CurrKEY(1)))
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ClearField
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CQ01")) = False Then: textCQ01 = rsTmp.Fields("CQ01")
      LabelCQ01 = GetPrjName1(textCQ01)
      If IsNull(rsTmp.Fields("CQ02")) = False Then: textCQ02 = TAIWANDATE(rsTmp.Fields("CQ02"))
      If IsNull(rsTmp.Fields("CQ03")) = False Then: textCQ03 = rsTmp.Fields("CQ03")
      If IsNull(rsTmp.Fields("CQ04")) = False Then: textCQ04 = rsTmp.Fields("CQ04")
      If IsNull(rsTmp.Fields("CQ05")) = False Then: m_CQ05 = rsTmp.Fields("CQ05")
      'Added by Lydia 2017/08/09
      If IsNull(rsTmp.Fields("CQ11")) = False Then: textCQ11 = "" & rsTmp.Fields("CQ11")
      textCQ11.Tag = textCQ11.Text
      'end 2017/08/09
      
      SetList lstAtt, textCQ04
      
      cmdOpenAtt.Enabled = True
      Call UpdateCUID(rsTmp)
      
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
   End If

   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub SetList(oList As ListBox, p_stList As String)
   Dim arrID
   oList.Clear
   If p_stList <> "" Then
      arrID = Split(p_stList, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
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
   
   If IsNull(rsSrcTmp.Fields("CQ05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ05")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("CQ05"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CQ06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ06")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("CQ06"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CQ07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ07")) = False Then
         strTemp = rsSrcTmp.Fields("CQ07")
         'Modified by Lydia 2017/08/17 "##:##" => "##:##:##"
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CQ08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ08")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("CQ08"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CQ09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ09")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("CQ09"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CQ10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CQ10")) = False Then
         strTemp = rsSrcTmp.Fields("CQ10")
         'Modified by Lydia 2017/08/17 "##:##" => "##:##:##"
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
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
Dim nResponse As Boolean
Dim strTmp  As String
Dim strTit As String
Dim strMsg As String
Dim tmpArr1 As Variant, tmpArr2 As Variant 'Added by Lydia 2017/08/09

   CheckDataValid = False
   
   If textCQ01.Text = "" Then
      MsgBox "代理人編號不可空白！", vbExclamation
      textCQ01.SetFocus
      Exit Function
   End If
   
   If textCQ02.Text = "" Then
      MsgBox "日期不可空白！", vbExclamation
      textCQ02.SetFocus
      Exit Function
   End If
   
   If textCQ03.Text = "" Then
      MsgBox "內容不可空白！", vbExclamation
      textCQ03.SetFocus
      Exit Function
   End If
   
   If textCQ04.Text = "" Then
      MsgBox "附件不可空白！", vbExclamation
      Exit Function
   End If

   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textCQ01.Locked = bEnable
   textCQ02.Locked = bEnable
   If bEnable Then textCQ01.BackColor = &H8000000F Else textCQ01.BackColor = &H80000005
   If bEnable Then textCQ02.BackColor = &H8000000F Else textCQ02.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textCQ01.Locked = bEnable
   textCQ02.Locked = bEnable
   If bEnable Then textCQ01.BackColor = &H8000000F Else textCQ01.BackColor = &H80000005
   If bEnable Then textCQ02.BackColor = &H8000000F Else textCQ02.BackColor = &H80000005
   textCQ03.Locked = bEnable
   textCQ04.Locked = bEnable
   textCQ11.Locked = bEnable 'Added by Lydia 2017/08/09
   
   If bEnable = False Then
      'cmdOpenAtt.Enabled = True
      cmdAddAtt.Enabled = True
      cmdRemAtt.Enabled = True
   Else
      'cmdOpenAtt.Enabled = False
      cmdAddAtt.Enabled = False
      cmdRemAtt.Enabled = False
   End If
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textCQ01 = Empty
   LabelCQ01 = Empty
   textCQ02 = Empty
   textCQ03 = Empty
   textCQ04 = Empty
   textCQ11 = Empty: textCQ11.Tag = textCQ11.Text  'Added by Lydia 2017/08/09
   Label23.Caption = Empty
   m_CQ05 = Empty
   
   lstAtt.Clear
   cmdOpenAtt.Enabled = False
   cmdAddAtt.Enabled = False
   cmdRemAtt.Enabled = False
   
   For nIndex = 0 To tf_CQ - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "CQ01", textCQ01
      SetFieldNewData "CQ02", DBDATE(textCQ02)
   End If
   SetFieldNewData "CQ03", textCQ03
   SetFieldNewData "CQ04", textCQ04
   SetFieldNewData "CQ11", textCQ11 'Added by Lydia 2017/08/09
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To tf_CQ
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CQ" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 2:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

Private Sub textCQ01_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textCQ01
      CloseIme
   End If
End Sub

Private Sub textCQ01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCQ01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   LabelCQ01 = Empty
   If IsEmptyText(textCQ01) = False Then
      textCQ01 = Left(textCQ01 & "000000000", 9)
      LabelCQ01 = GetPrjName1(textCQ01)
      Select Case m_EditMode
         Case 1, 4:
            If Left(textCQ01, 1) <> "Y" Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "必須輸入代理人編號"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCQ01_GotFocus
               GoTo EXITSUB
            End If
            If IsEmptyText(LabelCQ01) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "此代理人編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCQ01_GotFocus
            End If
      End Select
   End If
EXITSUB:
End Sub

Private Sub textCQ02_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textCQ02
      CloseIme
   End If
End Sub

Private Sub textCQ02_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub textCQ02_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textCQ02 <> "" Then
      If CheckIsTaiwanDate(textCQ02, False) = False Then
         Call textCQ02_GotFocus
         Cancel = True
         MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
         Exit Sub
      ElseIf ChkWork(ChangeTStringToWString(textCQ02)) = False Then
         Call textCQ02_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub textCQ03_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textCQ03
      OpenIme
   End If
End Sub

'開啟附件
Private Sub cmdOpenAtt_Click()
'Added by Lydia 2017/08/09
Dim tmpArr As Variant, ii As Integer
Dim stFileName As String
Dim hLocalFile As Long
'end 2017/08/09

   If lstAtt.Text = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      'Added by Lydia 2017/08/09 判斷移檔日期
      If strSrvDate(1) >= CR_NewDate And textCQ11 <> "" Then
         tmpArr = Empty
         tmpArr = Split(textCQ11.Text, ",")
         ii = lstAtt.ListIndex
         If ii > UBound(tmpArr) Then Exit Sub
         If Trim(tmpArr(ii)) <> "" Then
            strExc(1) = Trim(Mid(lstAtt.Text, 1, InStrRev(lstAtt.Text, "(") - 1))
            stFileName = App.path & "\$$" & strExc(1)
            If PUB_GetFtpFile(Trim(tmpArr(ii)), stFileName, cTableName) Then
                ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            End If
         End If
      Else
      'end 2017/08/09
         PUB_OpenFtpFile textCQ01, lstAtt.Text, Winsock1, "4"
      End If 'end 2017/08/09
      
   End If
End Sub

'可多選,顯示檔案大小
Private Sub cmdAddAtt_Click()
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s
   Dim strMid As String, strList As String 'Added by Lydia 2017/08/09
   
On Error GoTo ErrHnd
   
   stFileName = "*.*"
   strList = textCQ11.Text  'Added by Lydia 2017/08/09
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            For ii = 1 To UBound(sFile)
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               'Modified by Lydia 2017/08/09 存FTP檔名
               'AddListX lstAtt, stFileName & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)"
               strMid = PUB_GetNewFileNameSec(Mid(stFileName, InStrRev(stFileName, "\") + 1), , strList)
               AddListX lstAtt, PUB_StringFilter(stFileName) & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)", strMid
               'end 2017/08/09
            Next
         Else
            stFileName = .FileName
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            'Modified by Lydia 2017/08/09 存FTP檔名
            'AddListX lstAtt, stFileName & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)"
            strMid = PUB_GetNewFileNameSec(Mid(stFileName, InStrRev(stFileName, "\") + 1), , strList)
            AddListX lstAtt, PUB_StringFilter(stFileName) & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)", strMid
            'end 2017/08/09
         End If
         textCQ04 = ComposeAttList(lstAtt)
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.NUMBER <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

'Modified by Lydia 2017/08/09 +存FTP檔名 stFtpName
Private Function AddListX(oList As ListBox, stNewItem As String, stFtpName As String) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String
   If InStr(stNewItem, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請重新命名！", vbExclamation
      cmdAddAtt.SetFocus
      Exit Function
   End If
   If stNewItem <> "" Then
      For idx = 0 To oList.ListCount - 1
         stFileName = GetFileName(oList.List(idx))
         If GetFileName(stNewItem) = stFileName Then
            MsgBox "附件[" & stFileName & "]已存在！"
            AddListX = False
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         oList.AddItem stNewItem, 0
         AddListX = True
         'Added by Lydia 2017/08/09 存FTP檔名 (堆疊)
         textCQ11 = stFtpName & IIf(textCQ11 <> "", ",", "") & textCQ11
      End If
   End If
End Function

Private Function GetFileName(ByVal FullPath As String) As String
   Dim stItem As String, iPos As Integer
   stItem = FullPath
   iPos = InStr(stItem, "\")
   Do While iPos > 0
      stItem = Mid(stItem, iPos + 1)
      iPos = InStr(stItem, "\")
   Loop
   GetFileName = stItem
End Function

'附件
Private Function ComposeAttList(oList As ListBox) As String
   Dim iPos As Integer, stItem As String, stRtn As String, idx As Integer
   If oList.ListCount > 0 Then
      'Modify By Sindy 2014/12/4
      'stItem = oList.List(0)
      'stRtn = GetFileName(stItem)
      'For idx = 1 To oList.ListCount - 1
      For idx = 0 To oList.ListCount - 1
         stItem = oList.List(idx)
         'stRtn = stRtn & "," & GetFileName(stItem)
         stItem = GetFileName(stItem)
         If InStr(oList.List(idx), "\") > 0 Then
            If InStr(stItem, textCQ02) = 0 Then
               stItem = textCQ02 & "_" & Trim(stItem)
            End If
         End If
         stRtn = stRtn & "," & stItem
      Next
      If stRtn <> "" Then stRtn = Mid(stRtn, 2)
      '2014/12/4 END
   End If
   ComposeAttList = stRtn
End Function

Private Sub cmdRemAtt_Click()
   'Modify By Sindy 2023/8/23 Mark
'   If InStr(lstAtt, "\") = 0 And Pub_StrUserSt03 <> "M51" Then
'         MsgBox "已上傳檔案不可移除！"
'   Else
   '2023/8/23 END
   If RemoveList(lstAtt) = True Then
      textCQ04 = ComposeList(lstAtt)
      cmdAddAtt.SetFocus
   End If
End Sub

Private Function ComposeList(oList As ListBox, Optional p_iOpt As Integer = 0) As String
   Dim iPos As Integer, stItem As String
   strExc(1) = ""
   If oList.ListCount > 0 Then
      For intI = 0 To oList.ListCount - 1
         If p_iOpt = 0 Then
            iPos = InStr(oList.List(intI), Chr(1))
            If iPos > 0 Then
               stItem = Left(oList.List(intI), iPos - 1)
            Else
               stItem = oList.List(intI)
            End If
         Else
            stItem = Format(oList.ItemData(intI), "00")
         End If
         stItem = GetFileName(stItem)
         If intI = 0 Then
            strExc(1) = stItem
         Else
            strExc(1) = strExc(1) & "," & stItem
         End If
      Next
   End If
   ComposeList = strExc(1)
End Function

Private Function RemoveList(oList As ListBox) As Boolean
   Dim ii As Integer
   Dim tmpArr As Variant 'Added by Lydia 2017/08/09
   
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            RemoveList = True
            oList.RemoveItem ii
            'Added by Lydia 2017/08/09 移除FTP檔名(可複選)
            If textCQ11 <> "" Then
               '重整FTP檔名
               textCQ11 = Replace(textCQ11, ",,", ",")
               If Left(textCQ11, 1) = "," Then textCQ11 = Mid(textCQ11, 2)
               If Right(textCQ11, 1) = "," Then textCQ11 = Mid(textCQ11, 1, Len(textCQ11) - 1)
               tmpArr = Empty
               tmpArr = Split(textCQ11, ",")
               If Trim(tmpArr(ii)) <> "" Then textCQ11 = Replace(textCQ11, Trim(tmpArr(ii)), "")
            End If
            'end 2017/08/09
            
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
      
      'Added by Lydia 2017/08/09 重整FTP檔名
      textCQ11 = Replace(textCQ11, ",,", ",")
      If Left(textCQ11, 1) = "," Then textCQ11 = Mid(textCQ11, 2)
      If Right(textCQ11, 1) = "," Then textCQ11 = Mid(textCQ11, 1, Len(textCQ11) - 1)
      'end 2017/08/09
   End If
End Function

'上傳附件檔
Private Function UploadAtt(ByVal stKEY As String, Optional iErrNo As Integer, Optional stErrMsg As String) As Boolean
   Dim hOpen As Long
   Dim hConnection As Long
   Dim hDir As Long
   Dim bReturn As Boolean
   Dim dwInternetFlags As Integer
   Dim stDir As String
   Dim stRemoteFile As String
   Dim stLocalFile As String
   Dim stItem As String
   Dim idx As Integer
   Dim iPos As Integer
   Dim IsTimeOut As Boolean
   Dim SeekTimer
   Dim ACT_FTP_IP As String
   Dim arrIP
   Dim ii As Integer
   
   iErrNo = 0
   stErrMsg = ""
   
   stDir = CF代理人報價附件存放路徑
   If lstAtt.ListCount > 0 Then
      For idx = 0 To lstAtt.ListCount - 1
         stItem = lstAtt.List(idx)
         iPos = InStr(stItem, "\")
         If iPos > 0 Then
            If InStrRev(stItem, " (") > 0 Then
               stLocalFile = Left(stItem, InStrRev(stItem, " (") - 1)
            Else
               stLocalFile = stItem
            End If
            stRemoteFile = GetFileName(stLocalFile)
            
            'Add By Sindy 2014/12/4 檔名要加日期,因代理人進來的電子檔若檔名一樣時,存到Server上會被蓋掉
            If InStr(stRemoteFile, textCQ02) = 0 Then
               stRemoteFile = textCQ02 & "_" & stRemoteFile
            End If
            '2014/12/4 END
            
            If hOpen = 0 Then
               hOpen = InternetOpen("Taie FTP", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
               If hOpen = 0 Then
                  iErrNo = 1
                  stErrMsg = "網路錯誤！"
                  GoTo OutPort
               Else
                  IsTimeOut = True
                  If GOOD_FTP_IP <> "" Then
                     arrIP = Split(GOOD_FTP_IP & ";" & FTP_IP, ";")
                  Else
                     arrIP = Split(FTP_IP, ";")
                  End If
                  For ii = LBound(arrIP) To UBound(arrIP)
                     ACT_FTP_IP = arrIP(ii)
                     If ACT_FTP_IP <> "" Then
                        '偵測 FTPServer 是否存在
                        If Winsock1.State Then Winsock1.Close
                        Winsock1.Connect ACT_FTP_IP, 21
                        IsTimeOut = False
                        SeekTimer = Timer
                        Do While Winsock1.State = 6 And IsTimeOut = False
                           DoEvents
                           If Timer - SeekTimer > 1 Then
                              IsTimeOut = True
                           End If
                        Loop
                        If Winsock1.State Then Winsock1.Close
                        If IsTimeOut = False Then
                           Exit For
                        End If
                     End If
                  Next
                  
                  '若是超過時間
                  If IsTimeOut = True Then
                     iErrNo = 2
                     stErrMsg = "無法與FTP Server建立連線！"
                     GoTo OutPort
                  Else
                     GOOD_FTP_IP = ACT_FTP_IP
                  End If
               
                  hConnection = InternetConnect(hOpen, ACT_FTP_IP, FTP_Port, _
                     "pgmid", "pgmpwd", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
                  If hConnection = 0 Then
                     iErrNo = 3
                     stErrMsg = "無法與FTP Server建立連線！"
                     GoTo OutPort
                  ElseIf FtpSetCurrentDirectory(hConnection, stDir) = False Then
                     iErrNo = 4
                     stErrMsg = "切換至CF代理人報價目錄失敗！"
                     GoTo OutPort
                  '切換至CF代理人報價單號目錄
                  ElseIf FtpSetCurrentDirectory(hConnection, stKEY) = False Then
                     hDir = FtpCreateDirectory(hConnection, stKEY)
                     If hDir = 0 Then
                        iErrNo = 5
                        stErrMsg = "建立CF代理人報價單號目錄失敗！"
                        GoTo OutPort
                     ElseIf FtpSetCurrentDirectory(hConnection, stKEY) = False Then
                        iErrNo = 6
                        stErrMsg = "切換至CF代理人報價單號目錄失敗！"
                        GoTo OutPort
                     End If
                  End If
               End If
            End If

            dwInternetFlags = FTP_TRANSFER_TYPE_BINARY
            bReturn = FtpPutFile(hConnection, stLocalFile, stRemoteFile, dwInternetFlags, 0)
            If bReturn = False Then
               iErrNo = 7
               stErrMsg = "檔案（" & stRemoteFile & "）上傳失敗！"
               GoTo OutPort
            End If
         End If
      Next
   End If
   UploadAtt = True
   
OutPort:
   If hOpen <> 0 Then InternetCloseHandle (hOpen)
   If hConnection <> 0 Then InternetCloseHandle (hConnection)
   If Winsock1.State Then Winsock1.Close
End Function

'刪除附件檔
Private Function RemoveAtt(ByVal stKEY As String, stFiles As String, Optional iErrNo As Integer, Optional stErrMsg As String) As Boolean
   Dim hOpen As Long
   Dim hConnection As Long
   Dim bReturn As Boolean
   Dim stDir As String
   Dim IsTimeOut As Boolean
   Dim SeekTimer
   Dim ACT_FTP_IP As String
   Dim arrIP
   Dim ii As Integer, jj As Integer
   Dim arrFile
   Dim stRemoteFile As String
   
   iErrNo = 0
   stErrMsg = ""
   
   stDir = CF代理人報價附件存放路徑
   arrFile = Split(stFiles, ",")
   For jj = LBound(arrFile) To UBound(arrFile)
      If arrFile(jj) <> "" Then
         If hOpen = 0 Then
            hOpen = InternetOpen("Taie FTP", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
            If hOpen = 0 Then
               iErrNo = 1
               stErrMsg = "網路錯誤！"
               GoTo OutPort
            Else
               IsTimeOut = True
               If GOOD_FTP_IP <> "" Then
                  arrIP = Split(GOOD_FTP_IP & ";" & FTP_IP, ";")
               Else
                  arrIP = Split(FTP_IP, ";")
               End If
               For ii = LBound(arrIP) To UBound(arrIP)
                  ACT_FTP_IP = arrIP(ii)
                  If ACT_FTP_IP <> "" Then
                     '偵測 FTPServer 是否存在
                     If Winsock1.State Then Winsock1.Close
                     Winsock1.Connect ACT_FTP_IP, 21
                     IsTimeOut = False
                     SeekTimer = Timer
                     Do While Winsock1.State = 6 And IsTimeOut = False
                        DoEvents
                        If Timer - SeekTimer > 1 Then
                           IsTimeOut = True
                        End If
                     Loop
                     If Winsock1.State Then Winsock1.Close
                     If IsTimeOut = False Then
                        Exit For
                     End If
                  End If
               Next
               
               '若是超過時間
               If IsTimeOut = True Then
                  iErrNo = 2
                  stErrMsg = "無法與FTP Server建立連線！"
                  GoTo OutPort
               Else
                  GOOD_FTP_IP = ACT_FTP_IP
               End If
               
               hConnection = InternetConnect(hOpen, ACT_FTP_IP, FTP_Port, _
                  "pgmid", "pgmpwd", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
               If hConnection = 0 Then
                  iErrNo = 3
                  stErrMsg = "無法與FTP Server建立連線！"
                  GoTo OutPort
               ElseIf FtpSetCurrentDirectory(hConnection, stDir) = False Then
                  iErrNo = 4
                  stErrMsg = "切換至CF代理人報價目錄失敗！"
                  GoTo OutPort
               '切換至CF代理人報價單號目錄
               ElseIf FtpSetCurrentDirectory(hConnection, stKEY) = False Then
                  '無法切換時當作已刪除
                  'iErrNo = 6
                  'stErrMsg = "切換至CF代理人報價單號目錄失敗！"
                  'GoTo OutPort
                  Exit For
               End If
            End If
         End If
         If InStrRev(arrFile(jj), " (") > 0 Then
            stRemoteFile = Left(arrFile(jj), InStrRev(arrFile(jj), " (") - 1)
         Else
            stRemoteFile = arrFile(jj)
         End If
         '刪除檔案不控制成功與否
         bReturn = FtpDeleteFile(hConnection, stRemoteFile)
         If bReturn = False Then
            iErrNo = 8
            stErrMsg = "檔案（" & stRemoteFile & "）刪除失敗！"
            GoTo OutPort
         End If
      End If
   Next
   
   RemoveAtt = True
   
OutPort:
   If hOpen <> 0 Then InternetCloseHandle (hOpen)
   If hConnection <> 0 Then InternetCloseHandle (hConnection)
   If Winsock1.State Then Winsock1.Close
End Function

'Add By Sindy 2014/12/3
Private Sub txt1_Change(Index As Integer)
   Select Case Index
      Case 0
         If Len(Me.txt1(Index).Text) > 6 Then Exit Sub
         Me.txt1(1).Text = Me.txt1(Index).Text
   End Select
End Sub

'Add By Sindy 2014/12/3
Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

'Add By Sindy 2014/12/3
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = UpperCase(KeyAscii)
      Case 2, 3, 4, 5
         KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub

'Add By Sindy 2014/12/3
Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1, 4, 5
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 2, 3
         If CheckIsTaiwanDate(txt1(Index), False) = False And Trim(txt1(Index)) <> "" Then
            Call txt1_GotFocus(Index)
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub

'Added by Lydia 2017/08/17
Private Sub lstAtt_DblClick()
   If cmdOpenAtt.Enabled = True Then
      cmdOpenAtt.Value = True
   End If
End Sub

'Added by Lydia 2017/08/09 新增／刪除附件
Private Function UpdateAttFile(ByVal stKEY As String, Optional iErrNo As Integer, Optional stErrMsg As String) As Boolean
Dim arrTmp As Variant, arrOldTmp As Variant
Dim stFtpPath As String
Dim ii As Integer
Dim strMid As String
Dim stFileName As String

On Error GoTo OutPort
   
   iErrNo = 0
   stErrMsg = ""

   arrTmp = Empty: arrOldTmp = Empty
   arrTmp = Split(textCQ11.Text, ",")
   arrOldTmp = Split(textCQ11.Tag, ",")
   
   '先：刪除附件
   If textCQ11.Tag <> "" Then
    For ii = 0 To UBound(arrOldTmp)
       If Trim(arrOldTmp(ii)) <> "" And InStr(textCQ11.Text, Trim(arrOldTmp(ii))) = 0 Then
          If PUB_DelFtpFile2(stKEY, Trim(arrOldTmp(ii)), cTableName) = False Then
             GoTo OutPort
          End If
       End If
    Next ii
   End If
   
   '後：新增附件
   If textCQ11.Text <> "" Then
    For ii = 0 To UBound(arrTmp)
       If Trim(arrTmp(ii)) <> "" And InStr(textCQ11.Tag, Trim(arrTmp(ii))) = 0 Then
          stFileName = Trim(Mid(lstAtt.List(ii), 1, InStrRev(lstAtt.List(ii), "(") - 1))
          strExc(1) = m_FieldList(0).fiNewData & "_" & m_FieldList(1).fiNewData & "_"
          strExc(2) = IIf(InStr(Trim(arrTmp(ii)), strExc(1)) = 0, strExc(1), "") & Trim(arrTmp(ii))
          If PUB_PutFtpFile(stFileName, Left(stKEY, 9), strExc(2), stFtpPath, cTableName) = False Then
             GoTo OutPort
          Else
             strMid = strMid & IIf(strMid <> "", ",", "") & stFtpPath
          End If
       ElseIf Trim(arrTmp(ii)) <> "" Then
          strMid = strMid & IIf(strMid <> "", ",", "") & Trim(arrTmp(ii))
       End If
    Next ii
    textCQ11.Text = strMid
   End If
   
   UpdateAttFile = True
   
   Exit Function
   
OutPort:
   iErrNo = Err.NUMBER
   stErrMsg = Err.Description
   
End Function

'Added by Lydia 2017/08/09 搬檔
Private Sub Cmd1_Click()
Dim stSQL As String, intR As Integer
Dim rsQuery As ADODB.Recordset
Dim stOldDir As String, stNewDir As String, stNewPath As String
Dim oFileName As String, mFileName As String
Dim strGrp As String, strList As String, strNameList As String
Dim tmpArr As Variant
Dim strTmpExc As String
Dim stDownFile As String
Dim strLost As String, strLostId As String

   stOldDir = CF代理人報價附件存放路徑
   stNewDir = PUB_GetFtpTableDir(stNewDir) & cTableName
   stSQL = "select CQ01,CQ02,CQ04 from CFQUOTATION where NVL(CQ04,'N') <> 'N' AND NVL(CQ11,'N')='N' order by CQ01,CQ02 "
   intR = 0
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With rsQuery
         .MoveFirst
         MsgBox "開始工作，共" & .RecordCount & "筆記錄!"
         Do While Not .EOF
            '清除暫存檔
            PUB_KillTempFile "$$*.*"
                        
            If strGrp <> "" & .Fields("CQ01") & .Fields("CQ02") Then
               If strGrp <> "" Then
                  strTmpExc = strTmpExc & "UPDATE CFQUOTATION SET CQ11='" & strList & "' WHERE CQ01||CQ02='" & strGrp & "' ;"
               End If
               strList = "": strNameList = ""
               strGrp = "" & .Fields("CQ01") & .Fields("CQ02")
               tmpArr = Empty
               tmpArr = Split("" & .Fields("CQ04"), ",")
            End If
            
            For intR = 0 To UBound(tmpArr)
               If Trim(tmpArr(intR)) <> "" Then
                   '先下載檔案
                   stDownFile = ""
                   '因為有附件檔名有包含刮號,直接到模組處理舊檔名
                   strExc(1) = PUB_StringFilter(Trim(tmpArr(intR)))
                   If InStr(strExc(1), "(") > 0 And InStr(strExc(1), " (") = 0 Then
                      strExc(1) = Mid(strExc(1), 1, InStrRev(strExc(1), "(") - 1) & " " & Mid(strExc(1), InStrRev(strExc(1), "("))
                   End If
                   PUB_OpenFtpFile "" & .Fields("CQ01"), strExc(1), Winsock1, "4", False, stDownFile
                   
                   If stDownFile = "" Then
                       strLostId = strLostId & .Fields("CQ01") & "," & IIf(Len(strLostId) > 50, vbCrLf, "")
                       strLost = strLost & .Fields("CQ01") & "_" & Trim(tmpArr(intR)) & vbCrLf
                   Else
                        oFileName = Trim(tmpArr(intR))
                        oFileName = Trim(Mid(oFileName, 1, InStrRev(oFileName, "(") - 1))
                        '新-FTP檔名(非中文)
                        mFileName = PUB_GetNewFileNameSec(oFileName, "2", strNameList, .Fields("CQ01") & "_" & .Fields("CQ02"))
                        If PUB_PutFtpFile(stDownFile, "" & .Fields("CQ01"), mFileName, stNewPath, cTableName) = True Then
                           strList = strList & IIf(strList <> "", ",", "") & stNewPath
                        Else
                           MsgBox "Error !"
                           Exit Sub
                        End If
                   End If
               End If
            Next intR
            .MoveNext
         Loop
         
         '最後一筆
         strTmpExc = strTmpExc & "UPDATE CFQUOTATION SET CQ11='" & strList & "' WHERE CQ01||CQ02='" & strGrp & "' ;"
      End With
      
      '清除暫存檔
      PUB_KillTempFile "$$*.*"
        
      If strTmpExc <> "" Then
         tmpArr = Empty
         tmpArr = Split(strTmpExc, ";")
         cnnConnection.BeginTrans
           For intR = 0 To UBound(tmpArr)
              If Trim(tmpArr(intR)) <> "" Then
                 cnnConnection.Execute Trim(tmpArr(intR)), intI
              End If
           Next intR
         cnnConnection.CommitTrans
         MsgBox "工作結束!"
      End If
   End If
  
   If strLost <> "" Then
      PUB_SendMail "QPGMR", "A3034", "", CF代理人報價附件存放路徑 & "在NT2缺少檔案", "資料夾:" & strLostId & vbCrLf & vbCrLf & "檔案名稱:" & strLost
   End If
   
   Set rsQuery = Nothing
   Exit Sub
   
ErrHandle:
   cnnConnection.RollbackTrans
   
OutPort:
   Exit Sub
   
End Sub

