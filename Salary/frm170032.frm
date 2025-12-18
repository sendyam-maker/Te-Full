VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170032 
   BorderStyle     =   1  '單線固定
   Caption         =   "尾牙摸彩、年資、全勤獎金維護"
   ClientHeight    =   5976
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8040
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5976
   ScaleWidth      =   8040
   Begin TabDlg.SSTab SSTab1 
      Height          =   5232
      Left            =   24
      TabIndex        =   18
      Top             =   672
      Width           =   7968
      _ExtentX        =   14055
      _ExtentY        =   9229
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "資料輸入"
      TabPicture(0)   =   "frm170032.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "textCUID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "MSHFlexGrid2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtYEAR"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Combo1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtStaffNo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdCut"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Combo3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdBuild"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtZone"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdIns"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtRows"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtTot"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "獎項維護"
      TabPicture(1)   =   "frm170032.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHFlexGrid1"
      Tab(1).Control(1)=   "txtInput"
      Tab(1).Control(2)=   "Command2(4)"
      Tab(1).Control(3)=   "Command2(3)"
      Tab(1).Control(4)=   "Command2(2)"
      Tab(1).Control(5)=   "Command2(1)"
      Tab(1).Control(6)=   "Command2(0)"
      Tab(1).Control(7)=   "Combo2"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "多筆瀏覽"
      TabPicture(2)   =   "frm170032.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt1(0)"
      Tab(2).Control(1)=   "txt1(1)"
      Tab(2).Control(2)=   "txt1(2)"
      Tab(2).Control(3)=   "txt1(3)"
      Tab(2).Control(4)=   "cmdok"
      Tab(2).Control(5)=   "GRD1"
      Tab(2).Control(6)=   "Line5"
      Tab(2).Control(7)=   "Line4"
      Tab(2).Control(8)=   "Label15"
      Tab(2).Control(9)=   "Label16"
      Tab(2).ControlCount=   10
      Begin VB.TextBox txtTot 
         Alignment       =   2  '置中對齊
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         Height          =   192
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "0"
         Top             =   4560
         Width           =   900
      End
      Begin VB.TextBox txtRows 
         Alignment       =   2  '置中對齊
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         Height          =   192
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0"
         Top             =   3984
         Width           =   900
      End
      Begin VB.CommandButton cmdIns 
         Caption         =   "加入(Ins)"
         Height          =   285
         Left            =   6936
         TabIndex        =   32
         Top             =   984
         Width           =   924
      End
      Begin VB.TextBox txtZone 
         Height          =   285
         Left            =   3072
         MaxLength       =   1
         TabIndex        =   1
         Top             =   404
         Width           =   324
      End
      Begin VB.CommandButton cmdBuild 
         Caption         =   "整批建立"
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   696
         Width           =   1116
      End
      Begin VB.ComboBox Combo3 
         Height          =   276
         ItemData        =   "frm170032.frx":0054
         Left            =   288
         List            =   "frm170032.frx":0061
         Style           =   2  '單純下拉式
         TabIndex        =   0
         Top             =   384
         Width           =   1740
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         ItemData        =   "frm170032.frx":007B
         Left            =   -74712
         List            =   "frm170032.frx":008B
         Style           =   2  '單純下拉式
         TabIndex        =   30
         Top             =   408
         Width           =   1740
      End
      Begin VB.CommandButton cmdCut 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6936
         Picture         =   "frm170032.frx":00A5
         Style           =   1  '圖片外觀
         TabIndex        =   5
         ToolTipText     =   "取消"
         Top             =   1344
         Width           =   350
      End
      Begin VB.TextBox txtStaffNo 
         Height          =   285
         Left            =   4752
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1008
         Width           =   948
      End
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   888
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   1008
         Width           =   2796
      End
      Begin VB.CommandButton Command2 
         Caption         =   "新增"
         Height          =   285
         Index           =   0
         Left            =   -74760
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改"
         Height          =   285
         Index           =   1
         Left            =   -74028
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "刪除"
         Height          =   285
         Index           =   2
         Left            =   -73284
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "存檔"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72528
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取消"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -71784
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtInput 
         Appearance      =   0  '平面
         Height          =   375
         Left            =   -73584
         TabIndex        =   26
         Top             =   2868
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -74040
         MaxLength       =   6
         TabIndex        =   13
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72990
         MaxLength       =   6
         TabIndex        =   14
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -71100
         MaxLength       =   4
         TabIndex        =   15
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -70110
         MaxLength       =   4
         TabIndex        =   16
         Top             =   360
         Width           =   915
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   300
         Left            =   -68730
         TabIndex        =   17
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txtYEAR 
         Height          =   285
         Left            =   888
         MaxLength       =   4
         TabIndex        =   2
         Top             =   708
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   4524
         Left            =   -74928
         TabIndex        =   22
         Top             =   696
         Width           =   7752
         _ExtentX        =   13674
         _ExtentY        =   7980
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "年度|獎項|金額|員工號|姓名"
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4080
         Left            =   -74760
         TabIndex        =   27
         Top             =   1068
         Width           =   3744
         _ExtentX        =   6604
         _ExtentY        =   7197
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FormatString    =   "代碼|獎項名稱|金額"
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3480
         Left            =   240
         TabIndex        =   7
         Top             =   1344
         Width           =   6672
         _ExtentX        =   11769
         _ExtentY        =   6138
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         AllowBigSelection=   0   'False
         FormatString    =   "獎項|金額|員工號|姓名"
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "合計:"
         Height          =   180
         Left            =   6984
         TabIndex        =   36
         Top             =   4296
         Width           =   408
      End
      Begin MSForms.Label lblName 
         Height          =   180
         Left            =   5808
         TabIndex        =   35
         Top             =   1056
         Width           =   1044
         Caption         =   "XXX"
         Size            =   "1841;317"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "筆數:"
         Height          =   180
         Left            =   6984
         TabIndex        =   33
         Top             =   3672
         Width           =   408
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "所別：         (1:北 2:中 3:南 4:高)"
         Height          =   180
         Left            =   2520
         TabIndex        =   31
         Top             =   456
         Width           =   2484
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "得獎人："
         Height          =   180
         Left            =   4008
         TabIndex        =   29
         Top             =   1056
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "獎項："
         Height          =   180
         Left            =   264
         TabIndex        =   28
         Top             =   1056
         Width           =   540
      End
      Begin MSForms.TextBox textCUID 
         Height          =   300
         Left            =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4848
         Width           =   6420
         VariousPropertyBits=   671105055
         Size            =   "7223;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line5 
         X1              =   -70380
         X2              =   -69780
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line4 
         X1              =   -73344
         X2              =   -72654
         Y1              =   516
         Y2              =   516
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         Height          =   180
         Left            =   -74952
         TabIndex        =   24
         Top             =   396
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "年度："
         Height          =   180
         Left            =   -71820
         TabIndex        =   23
         Top             =   396
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   2196
         TabIndex        =   20
         Top             =   1128
         Width           =   48
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年度：                     (ex:112)"
         Height          =   180
         Left            =   264
         TabIndex        =   19
         Top             =   760
         Width           =   2124
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
            Picture         =   "frm170032.frx":070F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170032.frx":0A2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170032.frx":0D47
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170032.frx":0F23
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170032.frx":123F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170032.frx":155B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170032.frx":1877
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170032.frx":1B93
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170032.frx":1EAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170032.frx":21CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm170032.frx":24E7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   8040
      _ExtentX        =   14182
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
Attribute VB_Name = "frm170032"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2023/10/25
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Const clColorSel As Long = &HFFC0C0
Dim iLstRow1 As Integer '前次點選列數1
Dim iLstRow2 As Integer '前次點選列數2
Dim iLstRow3 As Integer '前次點選列數3
Dim iRow As Integer, iCol As Integer '本次點選列數,行數
Dim rsQuery As ADODB.Recordset
Dim m_MB03 As String, m_MB04 As String
Dim data() As String

Private Function TxtValidate() As Boolean
   If m_EditMode = 1 Then
      If txtYEAR = "" Then
         MsgBox "請輸入年度！", vbExclamation
         txtYEAR.SetFocus
         Exit Function
      ElseIf Val(txtYEAR) < 100 Or Val(txtYEAR) > 200 Then
         MsgBox "年度輸入錯誤！", vbCritical
         txtYEAR.SetFocus
         Exit Function
      End If
   End If
   TxtValidate = True
End Function
Private Sub cmdBuild_Click()
   Screen.MousePointer = vbHourglass
   BuildBatch
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCut_Click()
   GridDelRow2
End Sub

Private Sub cmdIns_Click()
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If Combo1 = "" Then
         MsgBox "請選擇獎項！", vbExclamation
         Combo1.SetFocus
         Exit Sub
      ElseIf lblName = "" Then
         MsgBox "員工號輸入錯誤！", vbCritical
         txtStaffNo.SetFocus
         txtStaffNo_GotFocus
         Exit Sub
      End If
      
      If Combo3.ListIndex = 0 Then
         If txtZone = "" Then
            MsgBox "請輸入所別！", vbCritical
            txtZone.SetFocus
            Exit Sub
         'Removed by Morgan 2024/1/29 取消(南高會合併舉辦)
         'ElseIf PUB_GetST06(txtStaffNo) <> txtZone Then
         '   MsgBox "員工所別錯誤！", vbCritical
         '   txtStaffNo.SetFocus
         '   txtStaffNo_GotFocus
         '   Exit Sub
         End If
      End If
      
      GridAddRow2
   End If
End Sub

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


Private Sub Combo2_Click()
   LoadGrid1
End Sub

Private Sub Combo3_Click()
   If Val(Combo3.Tag) <> Combo3.ListIndex Then
      txtYEAR = ""
      setCombo1
      If m_EditMode = 0 Then
         ShowRecord -2
      End If
      UpdateToolbarState
   End If
   Combo3.Tag = Combo3.ListIndex
End Sub

Private Sub SetEnable()
   '年資及全勤新增時可整批建立
   If m_EditMode = 1 And (Combo3.ListIndex = 1 Or Combo3.ListIndex = 2) Then
      cmdBuild.Enabled = True
   Else
      cmdBuild.Enabled = False
   End If
   
   If (m_EditMode = 1 Or m_EditMode = 4) And Combo3.ListIndex = 0 Then
      txtZone.Enabled = True
   Else
      txtZone.Enabled = False
   End If
End Sub

Private Sub Command2_Click(Index As Integer)
   CmdEnable Index
   Select Case Index
      Case 0 '新增
         GridAddRow
      Case 1 '修改
         If MSHFlexGrid1.TextMatrix(1, 0) = "" Then
            GridAddRow
         End If
      Case 2 '刪除
         GridDelRow
      Case 3 '存檔
         If SaveGrid1 = False Then
            Exit Sub
         End If
      Case 4 '取消
         LoadGrid1
   End Select
   
End Sub

Private Sub CmdEnable(pIdx As Integer)
   Select Case pIdx
   Case 0, 1, 2
      Command2(1).Enabled = False
      Command2(3).Enabled = True
      Command2(4).Enabled = True
      Combo2.Locked = True
      TabEnable False, SSTab1.Tab
      
   Case 3, 4
      Command2(1).Enabled = True
      Command2(3).Enabled = False
      Command2(4).Enabled = False
      txtInput.Visible = False
      Combo2.Locked = False
      TabEnable True
   End Select
End Sub

Private Sub TabEnable(pEnable As Boolean, Optional pActTab As Integer)
   Dim ii As Integer
   If pEnable Then
      For ii = 0 To SSTab1.Tabs - 1
         SSTab1.TabEnabled(ii) = True
      Next
      TBar1.Enabled = True
   Else
      For ii = 0 To SSTab1.Tabs - 1
         If ii <> pActTab Then
            SSTab1.TabEnabled(ii) = False
         End If
      Next
      TBar1.Enabled = False
   End If
End Sub

Private Sub Form_Activate()
   Static m_bActived As Boolean
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
   
   SetCombo
   setCombo1
   
   InitialField
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      'Form_KeyDown vbKeyF2, 0
   End If
   UpdateToolbarState
End Sub

'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   If SSTab1.Tab = 1 Then
      TBar1.Buttons(1).Enabled = False
      TBar1.Buttons(2).Enabled = False
      TBar1.Buttons(3).Enabled = False
      TBar1.Buttons(4).Enabled = False
      TBar1.Buttons(6).Enabled = False
      TBar1.Buttons(7).Enabled = False
      TBar1.Buttons(8).Enabled = False
      TBar1.Buttons(9).Enabled = False
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
   Else
      Select Case m_EditMode
         Case 0 ' 無任何動作
            If m_bInsert Then
               TBar1.Buttons(1).Enabled = True
            Else
               TBar1.Buttons(1).Enabled = False
            End If
            If m_bUpdate And txtYEAR <> "" Then
               TBar1.Buttons(2).Enabled = True
            Else
               TBar1.Buttons(2).Enabled = False
            End If
            If m_bDelete And txtYEAR <> "" Then
               TBar1.Buttons(3).Enabled = True
            Else
               TBar1.Buttons(3).Enabled = False
            End If
            If m_bQuery Then
               TBar1.Buttons(4).Enabled = True
            Else
               TBar1.Buttons(4).Enabled = False
            End If
            If m_bQuery And txtYEAR <> "" Then
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
            cmdCut.Enabled = False
            cmdIns.Enabled = False
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
            If m_EditMode = 1 Or m_EditMode = 2 Then
               cmdCut.Enabled = True
               cmdIns.Enabled = True
            Else
               cmdCut.Enabled = False
               cmdIns.Enabled = False
            End If
      End Select
      SetEnable
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsQuery = Nothing
   Set frm170032 = Nothing
End Sub

Private Sub GRD1_Click()
   With GRD1
   .row = .MouseRow
   .col = .MouseCol
   End With
   GridClick3
End Sub

Private Sub GRD1_DblClick()
   Dim lCurRow As Long
   
   If m_MB03 <> "" Then Exit Sub
   
   lCurRow = GRD1.row
   '呼叫查詢
   If lCurRow > 0 Then
      If GRD1.TextMatrix(lCurRow, 0) <> "" Then
         If TBar1.Buttons(4).Enabled = True Then
            Call Tbar1_ButtonClick(TBar1.Buttons(4))
            If txtYEAR.Locked = False Then
               SetComboListByItemData Val(GRD1.TextMatrix(lCurRow, 7)), Combo3
               txtZone = "" & GRD1.TextMatrix(lCurRow, 9)
               txtYEAR = GRD1.TextMatrix(lCurRow, 0)
               If TBar1.Buttons(11).Enabled = True Then
                  m_MB03 = GRD1.TextMatrix(lCurRow, 8)
                  m_MB04 = GRD1.TextMatrix(lCurRow, 4)
                  Call Tbar1_ButtonClick(TBar1.Buttons(11))
                  m_MB03 = "": m_MB04 = ""
               End If
            End If
         End If
         'SSTab1.Tab = 2
      End If
   End If
End Sub

Private Sub MSHFlexGrid1_Click()
   With MSHFlexGrid1
   .row = .MouseRow
   .col = .MouseCol
   End With
   GridClick
End Sub


Private Sub MSHFlexGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   With MSHFlexGrid2
   .row = .MouseRow
   .col = .MouseCol
   End With
   GridClick2
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Static flg1 As Byte, flg2 As Byte
      
   If SSTab1.Tab = 1 Then
      If flg1 = 0 Then
         LoadGrid1
         flg1 = 1
      End If
   ElseIf SSTab1.Tab = 0 Then
      If PreviousTab = 2 Then
         GRD1_DblClick
      End If
   ElseIf SSTab1.Tab = 2 Then
      If flg2 = 0 Then
         SetGrd True
         flg2 = 1
      End If
   End If
   UpdateToolbarState
End Sub

Private Sub GridAddRow()
   Dim iNo As Integer
   Dim ii As Integer
   
   With MSHFlexGrid1
   If .TextMatrix(1, 0) = "" Then
      iNo = 1
   Else
      iNo = Val(.TextMatrix(.Rows - 1, 0)) + 1
      .Rows = .Rows + 1
   End If
   .row = .Rows - 1
   .TextMatrix(.row, 0) = Format(iNo, "00")
   SetGridColor MSHFlexGrid1, iLstRow1
   iLstRow1 = .row
   .TopRow = .row
   .Refresh
   .col = 1
   End With
   GridClick
   
End Sub

Private Sub GridAddRow2()
   
On Error GoTo ErrHnd

   data = Split(Combo1, " ")
      
   strSql = "update rdatafactory set ROWSEQ=ROWSEQ where formname='" & Me.Name & "' and id='" & strUserNum & "' and r001='" & Format(Combo1.ItemData(Combo1.ListIndex), "00") & "' and r002='" & txtStaffNo & "'"
   cnnConnection.Execute strSql, intI
   If intI > 0 Then
      MsgBox "資料已存在！", vbExclamation
      
   Else
      strSql = "INSERT INTO rdatafactory (formname,id,seqno,r001,r002,r003,r004)" & _
         " select '" & Me.Name & "','" & strUserNum & "',sqlno,'" & Format(Combo1.ItemData(Combo1.ListIndex), "00") & "','" & txtStaffNo & "','" & Format(data(1)) & "','" & ChgSQL(data(0)) & "'" & _
         " from (select nvl(max(seqno),0)+1 sqlno from rdatafactory where formname='" & Me.Name & "' and id='" & strUserNum & "') x"
      
      cnnConnection.Execute strSql, intI
      
      Combo3.Enabled = False
      txtZone.Enabled = False
      
      ShowGrid2
      txtStaffNo = ""
      MSHFlexGrid2.TopRow = MSHFlexGrid2.row
   End If
   Exit Sub
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Sub

Private Sub ShowGrid2(Optional pClick As Boolean = False)
   SetGridHead2 True
   strExc(0) = "select r004,trim(to_char(r003,'999,990')) r003,r002,st02,a0922,r001" & _
      " from rdatafactory,staff,acc090new where formname='" & Me.Name & "' and id='" & strUserNum & "' and st01(+)=r002 and a0921(+)=st93"
   If Combo3.ListIndex = 2 Then
      strExc(0) = strExc(0) & " order by st06,st93,st01"
   Else
      strExc(0) = strExc(0) & " order by r001,r002"
   End If
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Set MSHFlexGrid2.Recordset = rsQuery
      If txtStaffNo <> "" Then
         rsQuery.Find "r001='" & Format(Combo1.ItemData(Combo1.ListIndex), "00") & "'"
         If Not rsQuery.EOF Then
            rsQuery.Find "r002='" & txtStaffNo & "'"
         End If
         If Not rsQuery.EOF Then
            MSHFlexGrid2.row = rsQuery.AbsolutePosition
         End If
         GridClick2
      Else
         If pClick Then GridClick2
      End If
      SetGridHead2
   End If
End Sub

Private Sub SetGridColor(pGrid As MSHFlexGrid, pLstRow As Integer)
   Dim ii As Integer
   Dim lColor As Long
   Dim iRow As Integer
   
   With pGrid
   If pLstRow <> .row Then
      iRow = .row
      For ii = 0 To .Cols - 1
         .col = ii
         .CellBackColor = clColorSel
      Next
      If pLstRow > 0 Then
         .row = pLstRow
         For ii = 0 To .Cols - 1
            .col = ii
            .CellBackColor = .BackColor
         Next
      End If
      .row = iRow
   End If
   .Refresh
   End With
End Sub

Private Sub SetBox(pGrid As MSHFlexGrid, pText As TextBox, Optional pValue As String = "")
   Dim ii As Integer
   Dim lngLeft As Long, lngTop As Long
   
   With pGrid
      If .row > 0 And .col > 0 Then
         pText.FontName = .CellFontName
         pText.FontSize = .CellFontSize
         pText.Alignment = .CellAlignment \ 5
         If pValue <> "" Then
            pText.Text = pValue
         Else
            pText.Text = Format(.TextMatrix(.row, .col))
         End If
         pText.Tag = pText.Text
         pText.Width = .ColWidth(.col)
         pText.Height = .RowHeight(.row)
         
         If .CellAlignment < 3 Then
            pText.Alignment = 0
         ElseIf .CellAlignment < 6 Then
            pText.Alignment = 2
         Else
            pText.Alignment = 1
         End If
         lngLeft = .Left + 25
         lngTop = .Top + .RowHeight(0) + 25
         For ii = 0 To .col - 1
            lngLeft = lngLeft + .ColWidth(ii)
         Next
         For ii = .TopRow To .row - 1
            lngTop = lngTop + .RowHeight(ii)
         Next
         pText.Left = lngLeft: pText.Top = lngTop
         pText.Visible = True
         If pText.Locked = False Then
            pText.SetFocus
            TextInverse pText
         End If
         iRow = .row: iCol = .col
      End If
   End With
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

Private Sub txt1_GotFocus(Index As Integer)
   CloseIme
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   If iCol <> 1 And Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   Else
   
      If KeyAscii = vbKeyReturn Then
         With MSHFlexGrid1
         If .col = 2 Then
            .TextMatrix(iRow, iCol) = Format(txtInput.Text, "#,##0")
         Else
            .TextMatrix(iRow, iCol) = txtInput.Text
         End If
         If iCol > 0 And iCol < 2 Then
            .col = iCol + 1
            SetBox MSHFlexGrid1, txtInput
         ElseIf iCol = 2 And .row < .Rows - 1 Then
            .row = .row + 1
            .col = 1
            GridClick
         Else
            txtInput.Visible = False
         End If
         End With
      ElseIf KeyAscii = vbKeyEscape Then
         txtInput = txtInput.Tag
         TextInverse txtInput
      End If
      
   End If
End Sub

Private Sub LoadGrid1()
   Dim stSQL As String, intQ As Integer, stAC01 As String
   
   '14 尾牙摸彩
   '15 年資
   '16 全勤
   If Combo2.ListIndex = -1 Then Combo2.ListIndex = Combo3.ListIndex
   stAC01 = 13 + Combo2.ItemData(Combo2.ListIndex)
   
   SetGridHead1 True
   iLstRow1 = 0
   If stAC01 >= "14" And stAC01 <= "16" Then
      stSQL = "select ac02,ac03,trim(to_char(ac10,'999,990')) ac10 from allcode where ac01='" & stAC01 & "' order by 1"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         Set MSHFlexGrid1.Recordset = rsQuery.Clone
         SetGridHead1
         MSHFlexGrid1_Click
      ElseIf intQ = 0 Then
         MsgBox "尚未建立資料！", vbExclamation
      End If
   End If
End Sub

Private Sub SetGridHead1(Optional bolReset As Boolean = False)
   With MSHFlexGrid1
   .Visible = False
   If bolReset Then
      .Clear
      .Rows = 2
      .Cols = 3
      .col = 2
   End If
   .TextMatrix(0, 0) = "代碼"
   .ColWidth(0) = 500
   .ColAlignmentFixed(0) = flexAlignCenterCenter
   .ColAlignment(0) = flexAlignCenterCenter
   .TextMatrix(0, 1) = "獎項名稱"
   .ColWidth(1) = 1500
   .ColAlignmentFixed(1) = flexAlignCenterCenter
   .ColAlignment(1) = flexAlignLeftCenter
   .TextMatrix(0, 2) = "金額"
   .ColWidth(2) = 1000
   .ColAlignmentFixed(2) = flexAlignCenterCenter
   .ColAlignment(2) = flexAlignRightCenter
   If .Rows = 1 Then .col = 2
   .Visible = True
   End With
End Sub

Private Sub GridDelRow()
   Dim iRow As Integer

   With MSHFlexGrid1
   If .row = 0 Then
      MsgBox "請點選要刪除的資料！", vbExclamation
   ElseIf .Rows = 2 Then
      'MsgBox "最後一筆資料不可刪除！", vbExclamation
      SetGridHead1 True
   Else
      .RemoveItem .row
      .row = 1
      iLstRow1 = 0
      .Refresh
      GridClick
   End If
   End With
End Sub

Private Sub GridDelRow2()
   Dim iTopRow As Integer, iRow As Integer

   With MSHFlexGrid2
   If .row = 0 Or iLstRow2 = 0 Then
      MsgBox "請點選要刪除的資料！", vbExclamation
   ElseIf .TextMatrix(.row, 0) <> "" Then
      iTopRow = .TopRow
      iRow = .row
      
      strSql = "delete rdatafactory where formname='" & Me.Name & "' and id='" & strUserNum & "' and r001='" & .TextMatrix(.row, 5) & "' and r002='" & .TextMatrix(.row, 2) & "'"
      cnnConnection.Execute strSql, intI
      If intI > 0 Then
         ShowGrid2
         If .TextMatrix(1, 0) = "" Then
            If m_EditMode = 1 Then
               Combo3.Enabled = True
               SetEnable
            End If
         Else
            MSHFlexGrid2.TopRow = iTopRow
         End If
      End If
   End If
   End With
End Sub

Private Function SaveGrid1() As Boolean
   Dim ii As Integer, jj As Integer, stSQL As String, bolCheck As Boolean
   Dim stAC01 As String
   
   stAC01 = 13 + Combo2.ItemData(Combo2.ListIndex)
   
   With MSHFlexGrid1
   bolCheck = True
   For ii = 1 To .Rows - 1
      If .TextMatrix(ii, 0) <> "" Then
         If .TextMatrix(ii, 1) = "" Then
            MsgBox "獎項名稱不可空白！", vbExclamation
            bolCheck = False
            jj = 1
         ElseIf Val(Format(.TextMatrix(ii, 2))) <= 0 Then
            MsgBox "獎金金額必須大於0！", vbExclamation
            bolCheck = False
            jj = 2
         End If
      End If
      If bolCheck = False Then
         .row = ii: .col = jj
         Exit For
      End If
   Next
   End With
   If bolCheck = False Then
      GridClick
      Exit Function
   End If
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd1
   cnnConnection.Execute "delete allcode where ac01='" & stAC01 & "'", intI
   If RsTemp.State <> adStateClosed Then RsTemp.Close
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open "select * from allcode where ac01='" & stAC01 & "'", cnnConnection, adOpenDynamic, adLockBatchOptimistic
   With MSHFlexGrid1
   For ii = 1 To .Rows - 1
      If .TextMatrix(ii, 0) <> "" Then
         RsTemp.AddNew
         RsTemp.Fields("AC01") = stAC01
         RsTemp.Fields("AC02") = .TextMatrix(ii, 0)
         RsTemp.Fields("AC03") = .TextMatrix(ii, 1)
         RsTemp.Fields("AC10") = Val(Format(.TextMatrix(ii, 2)))
      End If
   Next
   End With
   RsTemp.UpdateBatch
   cnnConnection.CommitTrans
   SaveGrid1 = True
   Exit Function
   
ErrHnd1:
   cnnConnection.RollbackTrans
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub txtInput_Validate(Cancel As Boolean)
   
   With MSHFlexGrid1
   If .col = 2 Then
      .TextMatrix(iRow, iCol) = Format(txtInput.Text, "#,##0")
   Else
      .TextMatrix(iRow, iCol) = txtInput.Text
   End If
   End With
   
   'Modified by Morgan 2023/11/22 改設定不可見位置,否則會自我觸發導致堆疊空間不足錯誤28
   'txtInput.Visible = False
   txtInput.Top = -1000
   'end 2023/11/22
End Sub

Private Sub GridClick()
   With MSHFlexGrid1
   If Command2(3).Enabled = True Then
      SetBox MSHFlexGrid1, txtInput
   End If
   If .TextMatrix(.row, 0) <> "" Then
      SetGridColor MSHFlexGrid1, iLstRow1
      iLstRow1 = .row
   End If
   End With
End Sub

Private Sub GridClick2()
   With MSHFlexGrid2
   If .row > 0 And .TextMatrix(.row, 0) <> "" Then
      SetGridColor MSHFlexGrid2, iLstRow2
      iLstRow2 = .row
   End If
   End With
End Sub

Private Sub GridClick3()
   With GRD1
   If .TextMatrix(.row, 0) <> "" Then
      SetGridColor GRD1, iLstRow3
      iLstRow3 = .row
   End If
   End With
End Sub

' 執行指令
Public Sub OnAction(ByVal KeyCode As Integer)
   Dim bCancel As Boolean
   
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         SSTab1.Tab = 0
         m_EditMode = 1
         InitialField
         UpdateToolbarState
         SetInputEntry
         Combo3.SetFocus
         
      Case vbKeyF3 ' 修改
         SSTab1.Tab = 0
         If CheckData() = True Then
            m_EditMode = 2
            InitialField
            UpdateToolbarState
            SetInputEntry
         End If

      Case vbKeyF5 ' 刪除
         SSTab1.Tab = 0
         If CheckData() = True Then
            strExc(0) = "是否要刪除"
            If txtZone = "1" Then
               strExc(0) = strExc(0) & "【北所】"
            ElseIf txtZone = "2" Then
               strExc(0) = strExc(0) & "【中所】"
            ElseIf txtZone = "3" Then
               strExc(0) = strExc(0) & "【南所】"
            ElseIf txtZone = "4" Then
               strExc(0) = strExc(0) & "【高所】"
            End If
            strExc(0) = strExc(0) & txtYEAR & "年度" & Trim(Combo3) & "資料?"
            If MsgBox(strExc(0), vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
               m_EditMode = 3
               If OnWork = True Then
                   UpdateToolbarState
               Else
                   Exit Sub
               End If
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         SSTab1.Tab = 0
         m_EditMode = 4
         InitialField
         UpdateToolbarState
         SetInputEntry
         Combo3.SetFocus
         
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
            If m_MB03 <> "" Then
               For intI = 1 To MSHFlexGrid2.Rows - 1
                  If MSHFlexGrid2.TextMatrix(intI, 5) = m_MB03 And MSHFlexGrid2.TextMatrix(intI, 2) = m_MB04 Then
                     MSHFlexGrid2.row = intI
                     MSHFlexGrid2.TopRow = MSHFlexGrid2.row
                     GridClick2
                  End If
               Next
            End If
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
            txtYEAR = txtYEAR.Tag
            txtZone = txtZone.Tag
            m_EditMode = 0
            SetInputEntry
            ShowRecord
            UpdateToolbarState
         End If
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

Private Sub SetGridHead2(Optional bolReset As Boolean = False)
   Dim ii As Integer, iShowCol As Integer
   With MSHFlexGrid2
   If bolReset Then
      .Clear
      .Rows = 2
      .Cols = 9
      iLstRow2 = 0
      txtRows = 0
   Else
      txtRows = .Rows - 1
      'Added by Morgan 2025/1/9
      txtTot = ""
      For ii = 1 To .Rows - 1
         txtTot = Val(txtTot) + Val(Format(.TextMatrix(ii, 1)))
      Next
      txtTot = Format(txtTot, "#,###")
      'end 2025/1/9
   End If
   .TextMatrix(0, 0) = "獎項"
   If Combo3.ListIndex = 2 Then
      .ColWidth(0) = 1600
   Else
      .ColWidth(0) = 2000
   End If
   .ColAlignmentFixed(0) = flexAlignCenterCenter
   .ColAlignment(0) = flexAlignCenterCenter
   .TextMatrix(0, 1) = "金額"
   .ColWidth(1) = 1100
   .ColAlignmentFixed(1) = flexAlignCenterCenter
   .ColAlignment(1) = flexAlignCenterCenter
   .TextMatrix(0, 2) = "員工號"
   .ColWidth(2) = 1000
   .ColAlignmentFixed(2) = flexAlignCenterCenter
   .ColAlignment(2) = flexAlignCenterCenter
   .TextMatrix(0, 3) = "姓名"
   .ColWidth(3) = 1400
   .ColAlignmentFixed(3) = flexAlignCenterCenter
   .ColAlignment(3) = flexAlignCenterCenter
   iShowCol = 3
   If Combo3.ListIndex = 2 Then
      .TextMatrix(0, 4) = "部門"
      .ColWidth(4) = 1200
      .ColAlignmentFixed(4) = flexAlignCenterCenter
      .ColAlignment(4) = flexAlignLeftCenter
      iShowCol = 4
   End If
   
   For ii = iShowCol + 1 To .Cols - 1
      .ColWidth(ii) = 0
   Next
   .MergeCol(0) = True
   .MergeCol(1) = True
   .MergeCells = flexMergeRestrictColumns
   End With
End Sub

Private Sub InitialField()
   Dim iRow As Integer, iTopRow As Integer
   Dim stAC01 As String
   
   stAC01 = 13 + Combo3.ItemData(Combo3.ListIndex)
   
   iTopRow = MSHFlexGrid2.TopRow
   iRow = iLstRow2
   txtStaffNo = ""
   lblName = ""
   SetGridHead2 True
   If m_EditMode = 1 Or m_EditMode = 2 Then
      setCombo1
      cnnConnection.Execute "delete from rdatafactory where FORMNAME='" & Me.Name & "' and ID=" & CNULL(strUserNum), intI
      If m_EditMode = 1 Then
         txtYEAR = ""
         txtZone = ""
         textCUID = ""
      Else
         'Modified by Morgan 2024/2/5 維護時帶新的獎項名稱
         strSql = "INSERT INTO rdatafactory (formname,id,seqno,r001,r002,r003,r004)" & _
            "select '" & Me.Name & "','" & strUserNum & "',rownum,mb03,mb04,mb05,ac03 from MiscBonus,staff,allcode where mb01='" & (Val(txtYEAR) + 1911) & "' and mb02='" & Format(Combo3.ItemData(Combo3.ListIndex), "00") & "'" & IIf(txtZone <> "", " and mb10='" & txtZone & "'", "") & " and st01(+)=mb04 and ac01(+)='" & stAC01 & "' and ac02(+)=mb03 order by mb03,mb04,mb05"
         cnnConnection.Execute strSql, intI
         ShowGrid2
         If iRow > 0 Then
            MSHFlexGrid2.TopRow = iTopRow
            MSHFlexGrid2.row = iRow
            GridClick2
         End If
      End If
   Else
      txtYEAR = ""
      txtZone = ""
      textCUID = ""
   End If
End Sub

Private Sub setCombo1()
   Dim stSQL As String, intQ As Integer
   Dim stAC01 As String
   
   stAC01 = 13 + Combo3.ItemData(Combo3.ListIndex)
   
   Combo1.Clear
   stSQL = "select ac02,ac03,ac10 from allcode where ac01='" & stAC01 & "' order by 1 desc"
   intQ = 1
   Set RsTemp = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With RsTemp
      Do While Not .EOF
         Combo1.AddItem .Fields("ac03") & " " & Format(.Fields("ac10"), "###,##0"), 0
         Combo1.ItemData(0) = .Fields("ac02")
         .MoveNext
      Loop
      End With
      If stAC01 = "16" Then Combo1.ListIndex = 0
   End If
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1
         Combo3.Enabled = True
         txtYEAR.Enabled = True
         If Me.Visible = True Then
            txtYEAR.SetFocus
            txtYEAR_GotFocus
         End If
         SSTab1.TabEnabled(1) = False
         SSTab1.TabEnabled(2) = False
      Case 2
         txtYEAR.Enabled = False
         Combo3.Enabled = False
         txtZone.Enabled = False
         SSTab1.TabEnabled(1) = False
         SSTab1.TabEnabled(2) = False
      Case 4
         Combo3.Enabled = True
         txtYEAR.Enabled = True
         If Me.Visible = True Then
            txtYEAR.SetFocus
            txtYEAR_GotFocus
         End If
         SSTab1.TabEnabled(1) = False
         SSTab1.TabEnabled(2) = False
      Case Else
         Combo3.Enabled = True
         txtYEAR.Enabled = True
         SSTab1.TabEnabled(1) = True
         SSTab1.TabEnabled(2) = True
   End Select
End Sub

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 1: '新增
         If TxtValidate = True Then
            If SaveGrid2 = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
         
      Case 2: '修改
         If SaveGrid2 = True Then
            OnWork = True
            m_EditMode = 0
            ShowRecord
         End If
         
      Case 3: '刪除
         If DelRecord = True Then
            OnWork = True
            m_EditMode = 0
            ShowRecord 2
         End If
      
      Case 4: '查詢
         If txtYEAR = "" Then
            MsgBox "請輸入查詢年度！", vbExclamation
            txtYEAR.SetFocus
            Exit Function
         End If
         If ShowRecord = True Then
            OnWork = True
            m_EditMode = 0
         Else
            txtYEAR.SetFocus
            txtYEAR_GotFocus
         End If
         
   End Select
End Function

Private Sub txtStaffNo_Change()
   lblName = ""
   If Len(txtStaffNo) = 5 Then
      If Left(txtStaffNo, 1) = "F" Then
         MsgBox "不可輸入外譯編號!!"
      ElseIf ChkStaffID(txtStaffNo) = False Then
         If ClsPDGetStaffN(txtStaffNo, strExc(1)) = True Then
            lblName = strExc(1)
         End If
      End If
   End If
End Sub

Private Sub txtStaffNo_GotFocus()
   TextInverse txtStaffNo
End Sub

Private Sub txtStaffNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtYEAR_GotFocus()
   TextInverse txtYEAR
End Sub

Private Sub txtYEAR_KeyPress(KeyAscii As Integer)
   If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   End If
End Sub

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
   Dim CUID(1 To 6) As String
   Dim stMB02 As String, stMB10 As String
   Dim stCon As String, stSort As String
   
   
   stMB02 = Format(Combo3.ItemData(Combo3.ListIndex), "00")
   
   If p_iWay = 0 Then
      If txtZone <> "" Then
         stCon = " and mb10='" & txtZone & "'"
      End If
   Else
      stCon = (Val(txtYEAR) + 1911) & txtZone
   End If
   
   If Combo3.ListIndex = 2 Then
      stSort = " order by st06,st93,mb03,mb04,mb05"
   Else
      stSort = " order by mb03,mb04,mb05"
   End If
   
   Select Case p_iWay
      Case 0
         strExc(0) = "select mb06,trim(to_char(mb05,'999,990')) mb05,mb04,st02,a0922,mb03,mb07,mb08,mb09,mb01,mb10 from MiscBonus,staff,acc090new where mb01='" & (Val(txtYEAR) + 1911) & "' and mb02='" & stMB02 & "'" & stCon & _
         " and st01(+)=mb04 and a0921(+)=st93"
      Case -2
         strExc(0) = "select mb06,trim(to_char(mb05,'999,990')) mb05,mb04,st02,a0922,mb03,mb07,mb08,mb09,mb01,mb10 from MiscBonus,staff,acc090new where mb02='" & stMB02 & "' and st01(+)=mb04 and a0921(+)=st93" & _
            " and mb01||mb10=(select min(mb01||mb10) from MiscBonus where mb02='" & stMB02 & "')"
      Case -1
         strExc(0) = "select mb06,trim(to_char(mb05,'999,990')) mb05,mb04,st02,a0922,mb03,mb07,mb08,mb09,mb01,mb10 from MiscBonus,staff,acc090new where mb02='" & stMB02 & "' and st01(+)=mb04 and a0921(+)=st93" & _
            " and mb01||mb10=(select max(mb01||mb10) from MiscBonus where mb01||mb10<'" & stCon & "' and mb02='" & stMB02 & "')"
      Case 1
         strExc(0) = "select mb06,trim(to_char(mb05,'999,990')) mb05,mb04,st02,a0922,mb03,mb07,mb08,mb09,mb01,mb10 from MiscBonus,staff,acc090new where mb02='" & stMB02 & "' and st01(+)=mb04 and a0921(+)=st93" & _
            " and mb01||mb10=(select min(mb01||mb10) from MiscBonus where mb01||mb10>'" & stCon & "' and mb02='" & stMB02 & "')"
      Case 2
         strExc(0) = "select mb06,trim(to_char(mb05,'999,990')) mb05,mb04,st02,a0922,mb03,mb07,mb08,mb09,mb01,mb10 from MiscBonus,staff,acc090new where mb02='" & stMB02 & "' and st01(+)=mb04 and a0921(+)=st93" & _
            " and mb01||mb10=(select max(mb01||mb10) from MiscBonus where mb02='" & stMB02 & "')"
   End Select
   strExc(0) = strExc(0) & stSort
   
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      CUID(1) = "" & rsQuery.Fields("mb07")
      CUID(2) = "" & rsQuery.Fields("mb08")
      CUID(3) = "" & rsQuery.Fields("mb09")
      txtYEAR = Val("" & rsQuery.Fields("mb01")) - 1911
      txtYEAR.Tag = txtYEAR
      txtZone = "" & rsQuery.Fields("mb10")
      txtZone.Tag = txtZone
      UpdateCUID CUID, textCUID
      
      SetGridHead2 True
      iLstRow2 = 0
      Set MSHFlexGrid2.Recordset = rsQuery
      SetGridHead2
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         SetGridHead2 True
         MsgBox "查無資料！", vbInformation
         InitialField
      End If
   End If
   
   Set rsQuery = Nothing
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Me.TBar1.Enabled = False Then Exit Sub
   
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

      Case vbKeyInsert
         If cmdIns.Enabled Then cmdIns.Value = True
                  
   End Select
End Sub

Private Function SaveGrid2() As Boolean
   Dim ii As Integer, stMB01 As String, stMB02 As String, stMB10 As String
   
On Error GoTo EscPoint

   stMB01 = Val(txtYEAR) + 1911
   stMB02 = Format(Combo3.ItemData(Combo3.ListIndex), "00")
   stMB10 = txtZone
   
   If m_EditMode = 1 Then
      strSql = "update MiscBonus set mb07=mb07 where mb01=" & stMB01 & " and mb02='" & stMB02 & "' and rownum<2" & IIf(stMB10 <> "", " and mb10='" & stMB10 & "'", "")
      cnnConnection.Execute strSql, intI
      If intI > 0 Then
         strExc(0) = ""
         If txtZone = "1" Then
            strExc(0) = strExc(0) & "【北所】"
         ElseIf txtZone = "2" Then
            strExc(0) = strExc(0) & "【中所】"
         ElseIf txtZone = "3" Then
            strExc(0) = strExc(0) & "【南所】"
         ElseIf txtZone = "4" Then
            strExc(0) = strExc(0) & "【高所】"
         End If
         strExc(0) = strExc(0) & txtYEAR & "年度" & Trim(Combo3) & "已存在，不可再新增！"
         
          MsgBox strExc(0), vbCritical
          Exit Function
      End If
   End If

   cnnConnection.BeginTrans
   
On Error GoTo ErrHand
   If m_EditMode = 2 Then
      If CheckData(True) = False Then GoTo ErrHand
      
      strSql = "delete MiscBonus where mb01=" & stMB01 & " and mb02='" & stMB02 & "'" & IIf(stMB10 <> "", " and mb10='" & stMB10 & "'", "")
      cnnConnection.Execute strSql, intI
   End If
   With MSHFlexGrid2
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) <> "" Then
            strSql = "insert into MiscBonus(mb01,mb02,mb03,mb04,mb05,mb06,mb07,mb08,mb09,mb10)"
            strSql = strSql & "values(" & stMB01 & ",'" & stMB02 & "','" & .TextMatrix(ii, 5) & "','" & .TextMatrix(ii, 2) & "'," & Format(.TextMatrix(ii, 1)) & ",'" & ChgSQL(.TextMatrix(ii, 0)) & "','" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & stMB10 & "')"
            cnnConnection.Execute strSql, intI
         End If
      Next
   End With
   cnnConnection.CommitTrans
   SaveGrid2 = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    
EscPoint:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Function DelRecord() As Boolean
   Dim stMB01 As String, stMB02 As String
   
   stMB01 = Val(txtYEAR) + 1911
   stMB02 = Format(Combo3.ItemData(Combo3.ListIndex), "00")
   
On Error GoTo EscPoint
   cnnConnection.BeginTrans
   
On Error GoTo ErrHand
   If CheckData(True) = False Then GoTo ErrHand
   
   strSql = "delete MiscBonus where mb01=" & stMB01 & " and mb02='" & stMB02 & "'" & IIf(txtZone <> "", " and mb10='" & txtZone & "'", "")
   cnnConnection.Execute strSql, intI
      
   cnnConnection.CommitTrans
   DelRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    
EscPoint:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function

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
      strCTime = Format(p_CUID(3), "##:##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Sub GetData()
   Dim stCon As String, stMB02SQL As String
   
   SetGrd True
   stCon = ""
   If txt1(0) <> "" Then
      stCon = stCon & " and mb04>='" & txt1(0) & "'"
   End If
   If txt1(1) <> "" Then
      stCon = stCon & " and mb04<='" & txt1(1) & "'"
   End If
   
   If txt1(2) <> "" Then
      stCon = stCon & " and mb01>=" & (Val(txt1(2)) + 1911)
   End If
   If txt1(3) <> "" Then
      stCon = stCon & " and mb01<=" & (Val(txt1(3)) + 1911)
   End If
   
   strExc(0) = "select mb01-1911 Yr,ac03,mb06,trim(to_char(mb05,'999,990')) Amt,mb04,st02,mb01,mb02,mb03,mb10,a0922" & _
      " from MiscBonus,staff,allcode,acc090new where st01(+)=mb04 and ac01(+)='17' and ac02(+)=mb02 and a0921(+)=st93" & stCon & _
      " order by mb01,mb02,mb10,mb03,decode(mb02,'03',st06||st93),mb04,mb05"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set GRD1.Recordset = RsTemp.Clone
      GRD1.FormatString = GRD1.FormatString
      SetGrd
   End If
End Sub
Private Sub SetGrd(Optional bolReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iCol As Integer
   
   '格式顯示中文,代號隱藏
   arrGridHeadText = Array("年度", "類別", "獎項", "金額", "員工號", "姓名", "", "", "", "", "部門")
   arrGridHeadWidth = Array(500, 1000, 1600, 800, 800, 1000, 0, 0, 0, 0, 1400)
   With GRD1
   .Visible = False
   If bolReset Then
      .Clear
      .Rows = 2
      iLstRow3 = -1
   End If
   .Cols = UBound(arrGridHeadText) + 1
   For iCol = 0 To .Cols - 1
      .row = 0
      .col = iCol
      .Text = arrGridHeadText(iCol)
      .ColWidth(iCol) = arrGridHeadWidth(iCol)
      .CellAlignment = flexAlignCenterCenter
      
      If iCol = 3 Then
         .ColAlignment(iCol) = flexAlignRightCenter
      Else
         .ColAlignment(iCol) = flexAlignCenterCenter
      End If
   Next
   .Visible = True
   End With
End Sub

Private Sub txtZone_GotFocus()
   TextInverse txtZone
End Sub

Private Sub txtZone_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("1") And KeyAscii <= Asc("4")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub SetCombo()
   Combo2.Clear
   Combo3.Clear
   strExc(0) = "select ac02,ac03 from allcode where ac01='17' order by ac02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         Combo2.AddItem .Fields("ac03")
         Combo2.ItemData(Combo2.ListCount - 1) = Val(.Fields("ac02"))
         
         Combo3.AddItem .Fields("ac03")
         Combo3.ItemData(Combo3.ListCount - 1) = Val(.Fields("ac02"))
         .MoveNext
      Loop
      End With
      Combo3.ListIndex = 0
      Combo2.ListIndex = 0
   End If
End Sub

Private Sub BuildBatch()
   Dim stMB03 As String, iSEQNO As Integer, iYear As Integer
   
   If TxtValidate = True Then
      If MSHFlexGrid2.TextMatrix(1, 0) <> "" Then
         If MsgBox("現有資料會先清除，是否確定要繼續？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
      End If
      cnnConnection.Execute "delete from rdatafactory where FORMNAME='" & Me.Name & "' and ID=" & CNULL(strUserNum), intI
           
      SetGridHead2 True
      '全勤
      If Combo3.ListIndex = 2 Then
         data = Split(Combo1, " ")
         stMB03 = Format(Combo1.ItemData(Combo1.ListIndex), "00")
         strExc(0) = PUB_GetFullAttendanceStaff(DBDATE(txtYEAR & "0101"), DBDATE(txtYEAR & "1231"))
         strSql = "INSERT INTO rdatafactory (formname,id,seqno,r001,r002,r003,r004)" & _
               " select '" & Me.Name & "','" & strUserNum & "',rownum,'" & stMB03 & "' mb03" & _
               ",st01 mb04,'" & Format(data(1)) & "' mb05,'" & ChgSQL(data(0)) & "' mb06 from (" & strExc(0) & ")"
         cnnConnection.Execute strSql, intI
         If intI > 0 Then
            Combo3.Enabled = False
            ShowGrid2
         End If
      '年資
      ElseIf Combo3.ListIndex = 1 Then
         iSEQNO = 0
         'Modified by Morgan 2024/5/16 排除第4碼是9的編號
         strExc(0) = "select st01,st02 from staff,SalaryData where ST04='1' and ST13>0 and substr(st01,4,1)<>'9' and st93 not in('R04') and SD01(+)=st01 and sd01 is not null and (sd02 not in('P','F') or sd02 is null) order by 1"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            Do While Not .EOF
               iYear = CalYear(RsTemp(0), DBDATE(txtYEAR & "1231"))
               If iYear >= 5 Then
                  If iYear Mod 5 = 0 Then
                     intI = iYear / 5
                     'Added by Morgan 2024/1/3
                     If intI > Combo1.ListCount Then
                        MsgBox iYear & "年年資獎項未建立，無法新增" & RsTemp("st01") & "(" & RsTemp("st02") & ")資料！", vbExclamation
                     Else
                     'end 2024/1/3
                        data = Split(Combo1.List(intI - 1), " ")
                        stMB03 = Format(intI, "00")
                        iSEQNO = iSEQNO + 1
                        strSql = "INSERT INTO rdatafactory (formname,id,seqno,r001,r002,r003,r004)" & _
                           " VALUES('" & Me.Name & "','" & strUserNum & "'," & iSEQNO & ",'" & stMB03 & "','" & RsTemp("st01") & "','" & Format(data(1)) & "','" & ChgSQL(data(0)) & "')"
                        cnnConnection.Execute strSql, intI
                     End If
                  End If
               End If
               .MoveNext
            Loop
            End With
         End If
         If iSEQNO > 0 Then
            Combo3.Enabled = False
            ShowGrid2
         End If
      End If
   End If
End Sub

Private Sub SetComboListByItemData(pItemData As Long, pCombo As ComboBox)
   Dim ii As Integer
   For ii = 0 To pCombo.ListCount - 1
      If pCombo.ItemData(ii) = pItemData Then
         pCombo.ListIndex = ii
      End If
   Next
End Sub

Private Function CheckData(Optional pSaveCheck As Boolean) As Boolean
   Dim stMB01 As String, stMB02 As String, stMB10 As String
   
   stMB01 = Val(txtYEAR) + 1911
   stMB02 = Format(Combo3.ItemData(Combo3.ListIndex), "00")
   stMB10 = txtZone
   
   strSql = "update MiscBonus set mb11=mb11 where mb01=" & stMB01 & " and mb02='" & stMB02 & "'" & IIf(stMB10 <> "", " and mb10='" & stMB10 & "'", "") & " and mb11>0" & IIf(pSaveCheck = False, " and rownum<2", "")
   cnnConnection.Execute strSql, intI
   If intI > 0 Then
      MsgBox "獎金已轉檔，不可再異動！", vbExclamation
      Exit Function
   End If
   CheckData = True
End Function
