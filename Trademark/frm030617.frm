VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030617 
   BorderStyle     =   1  '單線固定
   Caption         =   "公報特定公司不列印者"
   ClientHeight    =   5730
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8950
   Begin TabDlg.SSTab SSTab1 
      Height          =   5052
      Left            =   0
      TabIndex        =   6
      Top             =   648
      Width           =   8952
      _ExtentX        =   15787
      _ExtentY        =   8908
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "公報特定公司不列印者"
      TabPicture(0)   =   "frm030617.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtTBNP01"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grd1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "公報特殊字對照檔"
      TabPicture(1)   =   "frm030617.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(11)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(15)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtBS02"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtBS03"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "grd2(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Combo1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "電子報特殊名單"
      TabPicture(2)   =   "frm030617.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "txtTBNP01_M"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(23)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "LblCnt"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label4"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label1(3)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtTBNP09"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label5"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "grd2(2)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtTBNP10"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin VB.TextBox txtTBNP10 
         Height          =   270
         Left            =   6636
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1290
         Width           =   330
      End
      Begin VB.ComboBox Combo1 
         Height          =   260
         ItemData        =   "frm030617.frx":0054
         Left            =   -73740
         List            =   "frm030617.frx":005E
         Style           =   2  '單純下拉式
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "多筆勾選新增"
         ForeColor       =   &H00FF0000&
         Height          =   4395
         Left            =   -70650
         TabIndex        =   7
         Top             =   630
         Width           =   4545
         Begin VB.CommandButton Command1 
            Caption         =   "新增"
            Height          =   345
            Left            =   3330
            TabIndex        =   2
            Top             =   150
            Width           =   1005
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
            Bindings        =   "frm030617.frx":007A
            Height          =   3825
            Index           =   0
            Left            =   90
            TabIndex        =   1
            Top             =   510
            Width           =   4350
            _ExtentX        =   7691
            _ExtentY        =   6756
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            ScrollTrack     =   -1  'True
            AllowUserResizing=   3
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
            _Band(0).Cols   =   1
         End
         Begin VB.Label Label2 
            Caption         =   "已是不列印的特定公司顯示為淺綠色"
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2925
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Bindings        =   "frm030617.frx":008F
         Height          =   4395
         Left            =   -74910
         TabIndex        =   3
         Top             =   600
         Width           =   4215
         _ExtentX        =   7426
         _ExtentY        =   7743
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
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
         _Band(0).Cols   =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
         Height          =   3645
         Index           =   1
         Left            =   -74820
         TabIndex        =   12
         Top             =   1350
         Width           =   8550
         _ExtentX        =   15064
         _ExtentY        =   6438
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   4
         FixedCols       =   0
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "V|欄位    |轉檔前               | 轉檔後                "
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
         Bindings        =   "frm030617.frx":00A4
         Height          =   4392
         Index           =   2
         Left            =   288
         TabIndex        =   16
         Top             =   600
         Width           =   4896
         _ExtentX        =   8643
         _ExtentY        =   7743
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
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
         _Band(0).Cols   =   1
      End
      Begin VB.Label Label5 
         Caption         =   "欲新增E-Mail時，先使用”申請人查詢”檢查此Mail是否已存在系統中。不要重覆新增，以免擾民。"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   710
         Left            =   5440
         TabIndex        =   26
         Top             =   2760
         Width           =   2820
      End
      Begin MSForms.TextBox txtTBNP09 
         Height          =   300
         Left            =   6750
         TabIndex        =   25
         Top             =   320
         Width           =   820
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "1446;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "編號："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   6180
         TabIndex        =   24
         Top             =   380
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "(空白查全部)"
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   4790
         TabIndex        =   23
         Top             =   360
         Width           =   1130
      End
      Begin VB.Label LblCnt 
         Caption         =   "共   筆"
         Height          =   300
         Left            =   5472
         TabIndex        =   22
         Top             =   4572
         Width           =   1524
      End
      Begin VB.Label Label3 
         Caption         =   "注意：不要直接刪除。而是註記態樣；這樣才能防止發生又匯入資料又再處理一次的狀況。"
         ForeColor       =   &H000000FF&
         Height          =   680
         Left            =   5440
         TabIndex        =   21
         Top             =   2010
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄電子報：       (N:不寄 D:無效)"
         Height          =   180
         Index           =   23
         Left            =   5400
         TabIndex        =   20
         Top             =   1320
         Width           =   2860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   312
         TabIndex        =   18
         Top             =   384
         Width           =   672
      End
      Begin MSForms.TextBox txtTBNP01_M 
         Height          =   300
         Left            =   1030
         TabIndex        =   17
         Top             =   320
         Width           =   3700
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "6526;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtBS03 
         Height          =   330
         Left            =   -73740
         TabIndex        =   11
         Top             =   960
         Width           =   5535
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "9763;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtBS02 
         Height          =   300
         Left            =   -73740
         TabIndex        =   10
         Top             =   660
         Width           =   5535
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "9763;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTBNP01 
         Height          =   300
         Left            =   -73230
         TabIndex        =   0
         Top             =   330
         Width           =   4455
         VariousPropertyBits=   671107099
         MaxLength       =   80
         Size            =   "7858;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "欄位："
         Height          =   180
         Index           =   1
         Left            =   -74310
         TabIndex        =   15
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "轉檔後："
         Height          =   180
         Index           =   15
         Left            =   -74490
         TabIndex        =   14
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "轉檔前："
         Height          =   180
         Index           =   11
         Left            =   -74490
         TabIndex        =   13
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "不列印的特定公司："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   -74850
         TabIndex        =   4
         Top             =   390
         Width           =   1620
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030617.frx":00B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030617.frx":03D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030617.frx":06F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030617.frx":08CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030617.frx":0BE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030617.frx":0F05
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030617.frx":1221
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030617.frx":153D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030617.frx":1859
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030617.frx":1B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030617.frx":1E91
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8950
      _ExtentX        =   15787
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
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frm030617"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/24 改成Form2.0 (txtTBNP01,txtBS02,txtBS03)
'Memo By Sindy 2012/12/5 智權人員欄已修改
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 第一筆資料的Key
Dim m_FirstKEY(1) As String
' 最後一筆資料的Key
Dim m_LastKEY(1) As String
' 目前正在顯示的Key
Dim m_CurrKEY(1) As String
Dim i As Integer, j As Integer
Dim dblPrevRow As Double
Dim dblPrevRow2 As Double 'Add By Sindy 2017/5/8
Dim m_strFindKey01 As String 'Add By Sindy 2023/8/25
Public m_WorkType As String 'Add By Sindy 2023/9/1 M=電子報特殊名單維護


'Add By Sindy 2017/10/30
Private Sub Combo1_Click()
   If m_EditMode = 1 Then '新增
      'Modify By Sindy 2020/1/14 和秀玲確認過,讓使用者自行設定
'      If Trim(Left(Combo1.Text, 2)) = "2" And Pub_StrUserSt03 <> "M51" Then
'         MsgBox "請通知電腦中心新增代理人資料！", vbInformation
'         Combo1.ListIndex = 0
'         Exit Sub
'      End If
   End If
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   If m_EditMode = 1 And Combo1.ListIndex >= 0 And txtBS02 <> "" Then
      ' 檢查記錄是否已存在
      If IsRecordExist(Left(Combo1.Text, 1), txtBS02) = True Then
         MsgBox "該筆記錄已存在", vbInformation
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub Command1_Click() '新增
Dim bolSelect As Boolean
   
On Error GoTo ErrHand
   
   bolSelect = False
   Screen.MousePointer = vbHourglass
   'cnnConnection.BeginTrans
   For i = 1 To grd2(0).Rows - 1
      grd2(0).col = 0
      grd2(0).row = i
      If Trim(grd2(0).Text) = "V" Then
         bolSelect = True
         '新增
         'Modify By Sindy 2013/3/1 +TBNP08
         strSql = "insert into TMBulletinNp (TBNP01,TBNP08) values(" & CNULL(grd2(0).TextMatrix(i, 1)) & ",'T')"
         cnnConnection.Execute strSql
      End If
   Next i
   'cnnConnection.CommitTrans
   Screen.MousePointer = vbDefault
   If bolSelect = False Then
      MsgBox "請點選欲新增的資料！"
      Exit Sub
   Else
      ReadAllData
      ShowFirstRecord
   End If
   
   Exit Sub
   
ErrHand:
   'cnnConnection.RollbackTrans
   MsgBox "新增失敗！" & vbCrLf & Err.Description
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
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   SSTab1.TabVisible(1) = False 'Add By Sindy 2024/1/17 公報特殊字對照檔
   If m_WorkType = "M" Then
      Me.Caption = "電子報特殊名單維護"
      SSTab1.TabVisible(0) = False
      SSTab1.TabVisible(1) = False
      SSTab1.Tab = 2
   Else
      SSTab1.Tab = 0
      SSTab1.TabVisible(2) = False
   End If
   
   ClearField
   RefreshRange
   ReadAllData
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   'OnAction vbKeyF4
   OnAction vbKeyF10
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030617 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grd1, x, y, nCol, nRow
grd1.col = nCol
grd1.row = nRow
End Sub

Private Sub grd1_SelChange()
grd1.Visible = False
If grd1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      grd1.col = 0
      grd1.row = dblPrevRow
      For i = 0 To grd1.Cols - 1
         grd1.col = i
         grd1.CellBackColor = QBColor(15)
      Next i
   End If
   '目前資料列反白
   grd1.col = 0
   grd1.row = grd1.MouseRow
   dblPrevRow = grd1.row
   For i = 0 To grd1.Cols - 1
      grd1.col = i
      grd1.CellBackColor = &HFFC0C0
   Next i
   '查詢目前資料列
   ShowCurrRecord grd1.TextMatrix(grd1.row, 0)
   m_CurrKEY(0) = grd1.TextMatrix(grd1.row, 0)
   UpdateCtrlData
End If
grd1.Visible = True
End Sub

Private Sub grd2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grd2(Index), x, y, nCol, nRow
grd2(Index).col = nCol
grd2(Index).row = nRow
End Sub

Private Sub GRD2_SelChange(Index As Integer)
Dim tmpMouseRow As Long

'grd2(index).Visible = False
tmpMouseRow = grd2(Index).row
If tmpMouseRow <> 0 Then
   grd2(Index).col = 0
   grd2(Index).row = tmpMouseRow
   If grd2(Index).TextMatrix(tmpMouseRow, 1) <> "" Then
      If SSTab1.Tab = 0 Then
         If grd2(Index).Text = "V" Then
            grd2(Index).Text = ""
            For i = 0 To grd2(Index).Cols - 1
               grd2(Index).col = i
               grd2(Index).CellBackColor = QBColor(15)
            Next i
         Else
            grd2(Index).Text = "V"
            For i = 0 To grd2(Index).Cols - 1
               grd2(Index).col = i
               grd2(Index).CellBackColor = &HFFC0C0
            Next i
            ' 檢查記錄是否已存在
            If IsRecordExist(grd2(Index).TextMatrix(grd2(Index).row, 1), grd2(Index).TextMatrix(grd2(Index).row, 2)) = True Then
               MsgBox "該筆記錄已存在", vbInformation
               grd2(Index).TextMatrix(grd2(Index).row, 0) = ""
               For i = 0 To grd2(Index).Cols - 1
                  grd2(Index).col = i
                  grd2(Index).CellBackColor = &H80FF80
               Next i
            End If
         End If
      ElseIf SSTab1.Tab = 1 Or SSTab1.Tab = 2 Then
         '上一筆資料列清除反白
         If dblPrevRow2 > 0 Then
            grd2(Index).TextMatrix(dblPrevRow2, 0) = ""
            grd2(Index).col = 0
            grd2(Index).row = dblPrevRow2
            For i = 0 To grd2(Index).Cols - 1
               grd2(Index).col = i
               grd2(Index).CellBackColor = QBColor(15)
            Next i
         End If
         '目前資料列反白
         dblPrevRow2 = tmpMouseRow
         grd2(Index).TextMatrix(dblPrevRow2, 0) = "V"
         grd2(Index).col = 0
         grd2(Index).row = dblPrevRow2 'GRD2(Index).MouseRow
         For i = 0 To grd2(Index).Cols - 1
            grd2(Index).col = i
            grd2(Index).CellBackColor = &HFFC0C0
         Next i
         '查詢目前資料列
         If SSTab1.Tab = 2 Then
            ShowCurrRecord grd2(Index).TextMatrix(grd2(Index).row, 1), ""
         Else
            ShowCurrRecord Left(grd2(Index).TextMatrix(grd2(Index).row, 1), 1), grd2(Index).TextMatrix(grd2(Index).row, 2)
            m_CurrKEY(0) = Left(grd2(Index).TextMatrix(grd2(Index).row, 1), 1)
            m_CurrKEY(1) = grd2(Index).TextMatrix(grd2(Index).row, 2)
            txtBS03.Text = grd2(Index).TextMatrix(grd2(Index).row, 3)
         End If
         UpdateCtrlData
      End If
   End If
End If
'grd2(index).Visible = True
End Sub

'Add By Sindy 2017/5/8
Private Sub SSTab1_Click(PreviousTab As Integer)
   ClearField
   RefreshRange
   ReadAllData
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   'OnAction vbKeyF4
   OnAction vbKeyF10
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
Dim Cancel As Boolean
   
   TxtValidate = False
   'Added by Morgan 2021/12/24 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/24
   
   If SSTab1.Tab = 0 Then
      If Trim(txtTBNP01.Text) = "" Then
          MsgBox "特定公司不可以空白！", vbExclamation
          txtTBNP01.SetFocus
          Exit Function
      End If
      If m_EditMode = 1 Then
         ' 檢查記錄是否已存在
         If IsRecordExist(txtTBNP01) = True Then
            MsgBox "該筆記錄已存在", vbOKOnly, "更新資料"
            txtTBNP01.SetFocus
            Exit Function
         End If
      End If
      Cancel = False
      txtTBNP01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   
   'Add By Sindy 2023/8/21
   ElseIf SSTab1.Tab = 2 Then
      If Trim(txtTBNP01_M.Text) = "" Then
          MsgBox "E-Mail不可以空白！", vbExclamation
          txtTBNP01_M.SetFocus
          Exit Function
      End If
      If m_EditMode = 1 Then
         ' 檢查記錄是否已存在
         If IsRecordExist(txtTBNP01_M) = True Then
            MsgBox "該筆記錄已存在", vbOKOnly, "更新資料"
            txtTBNP01_M.SetFocus
            Exit Function
         End If
      End If
      Cancel = False
      txtTBNP01_M_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      '2023/8/21 END
      
   ElseIf SSTab1.Tab = 1 Then
      If Trim(Combo1.Text) = "" Then
          MsgBox "欄位不可以空白！", vbExclamation
          Exit Function
      End If
      If Trim(txtBS02.Text) = "" Then
          MsgBox "轉檔前不可以空白！", vbExclamation
          txtBS02.SetFocus
          Exit Function
      End If
      Cancel = False
      Combo1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      
      If Trim(txtBS03.Text) = "" Then
          MsgBox "轉檔後不可以空白！", vbExclamation
          txtBS03.SetFocus
          Exit Function
      End If
      If Trim(txtBS02.Text) = Trim(txtBS03.Text) Then
          MsgBox "轉檔前和轉檔後,資料不可一樣！", vbExclamation
          txtBS03.SetFocus
          Exit Function
      End If
      
      'Add By Sindy 2017/10/30
      '新增代理人時,檢查代理人資料是否存在
      If m_EditMode = 1 And Left(Combo1, 1) = "2" Then
         strExc(0) = "SELECT TA02,TA03 FROM TAGENT" & _
                     " WHERE TA03='" & txtBS03.Text & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            MsgBox "代理人(" & txtBS03.Text & ")資料不存在！", vbExclamation
            txtBS03.SetFocus
            Exit Function
         End If
      End If
      '2017/10/30 END
      
      Cancel = False
      txtBS02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      Cancel = False
      txtBS03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

' 更新資料
Private Function SaveData(strEditMode As Integer) As Boolean
Dim strKEY01 As String, strKEY02 As String, strBS03 As String
Dim bolReSave As Boolean
   
On Error GoTo ErrHand
   
   SaveData = False
   
   bolReSave = False
   
   If SSTab1.Tab = 0 Then
      strKEY01 = Trim(txtTBNP01)
   ElseIf SSTab1.Tab = 1 Then
      strKEY01 = Left(Combo1.Text, 1)
      strKEY02 = Trim(txtBS02)
      strBS03 = Trim(txtBS03)
   'Add By Sindy 2023/8/21
   ElseIf SSTab1.Tab = 2 Then
      strKEY01 = Trim(txtTBNP01_M)
      '2023/8/21 END
   End If
   
   'cnnConnection.BeginTrans
ReSave:
   If SSTab1.Tab = 0 Then
      '新增
      If strEditMode = 1 Then
         'Modify By Sindy 2013/3/1 +TBNP08
         strSql = "insert into TMBulletinNp (TBNP01,TBNP08) values(" & CNULL(strKEY01) & ",'T')"
      '修改
      ElseIf strEditMode = 2 Then
         'Modify By Sindy 2013/3/1 +TBNP08
         strSql = "update TMBulletinNp set " & _
                     "TBNP01=" & CNULL(strKEY01) & _
                  " where TBNP01=" & CNULL(strKEY01) & " and TBNP08='T'"
      End If
      cnnConnection.Execute strSql
      'cnnConnection.CommitTrans
      If (strKEY01 < m_FirstKEY(0)) Or (strKEY01 > m_LastKEY(0)) Then
         RefreshRange
      End If
      ShowCurrRecord strKEY01
      
   'Add By Sindy 2017/5/5
   ElseIf SSTab1.Tab = 1 Then
      '新增
      If strEditMode = 1 Then
         strSql = "insert into BulletinSpecWord (BS01,BS02,BS03) values(" & CNULL(strKEY01) & "," & CNULL(strKEY02) & "," & CNULL(strBS03) & ")"
      '修改
      ElseIf strEditMode = 2 Then
         strSql = "update BulletinSpecWord set " & _
                     "BS03=" & CNULL(strBS03) & _
                  " where BS01=" & CNULL(strKEY01) & _
                  " and BS02=" & CNULL(strKEY02)
      End If
      cnnConnection.Execute strSql
      'cnnConnection.CommitTrans
      If (strKEY01 & strKEY02 < m_FirstKEY(0) & m_FirstKEY(1)) Or (strKEY01 & strKEY02 > m_LastKEY(0) & m_LastKEY(1)) Then
         RefreshRange
      End If
      ShowCurrRecord strKEY01, strKEY02
   '2017/5/5 END
   
   'Add By Sindy 2023/8/21
   ElseIf SSTab1.Tab = 2 Then
      '新增
      If strEditMode = 1 Then
         strSql = "insert into TMBulletinNp (TBNP01,TBNP08,TBNP10) values(" & CNULL(strKEY01) & ",'M'," & CNULL(Me.txtTBNP10) & ")"
      '修改
      ElseIf strEditMode = 2 Then
         strSql = "update TMBulletinNp set " & _
                     "TBNP10=" & CNULL(Me.txtTBNP10) & _
                  " where TBNP01=" & CNULL(strKEY01) & " and TBNP08='M'"
      End If
      cnnConnection.Execute strSql
      'cnnConnection.CommitTrans
      If (strKEY01 < m_FirstKEY(0)) Or (strKEY01 > m_LastKEY(0)) Then
         RefreshRange
      End If
      ShowCurrRecord strKEY01
   End If
   
   SaveData = True
   Exit Function
   
ErrHand:
   'cnnConnection.RollbackTrans
   If Err.Number = -2147217900 And bolReSave = False Then '造字錯誤,必須最後加空白才可存入DB
      bolReSave = True
      If SSTab1.Tab = 0 Then
         strKEY01 = Trim(txtTBNP01) & " "
      ElseIf SSTab1.Tab = 1 Then
         strKEY01 = Left(Combo1.Text, 1)
         strKEY02 = Trim(txtBS02) & " "
         strBS03 = Trim(txtBS03) & " "
      'Add By Sindy 2023/8/21
      ElseIf SSTab1.Tab = 2 Then
         strKEY01 = Trim(txtTBNP01_M)
      End If
      GoTo ReSave
   End If
   MsgBox " 更新失敗！" & vbCrLf & Err.Description
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strKEY01 As String, strKEY02 As String

On Error GoTo ErrHand

   DelRecord = False
   'cnnConnection.BeginTrans
   
   If SSTab1.Tab = 0 Then
      strKEY01 = txtTBNP01
      'Modify By Sindy 2013/3/1 +TBNP08
      strSql = "DELETE FROM TMBulletinNp WHERE ltrim(rtrim(TBNP01))=ltrim(rtrim(" & CNULL(strKEY01) & ")) and TBNP08='T'"
      cnnConnection.Execute strSql
   
   'Add By Sindy 2017/5/5
   ElseIf SSTab1.Tab = 1 Then
      strKEY01 = Left(Combo1.Text, 1)
      strKEY02 = txtBS02
      strSql = "DELETE FROM BulletinSpecWord WHERE BS01='" & strKEY01 & "' and ltrim(rtrim(BS02))=ltrim(rtrim(" & CNULL(strKEY02) & ")) and ltrim(rtrim(BS03))=ltrim(rtrim(" & CNULL(txtBS03) & "))"
      cnnConnection.Execute strSql
      '2017/5/5 END
   End If
   
   'cnnConnection.CommitTrans

   DelRecord = True
   Exit Function

ErrHand:
   'cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strKEY01 As String
Dim strKEY02 As String

   QueryRecord = False
   
   If SSTab1.Tab = 0 Then
      strKEY01 = txtTBNP01
   ElseIf SSTab1.Tab = 1 Then
      strKEY01 = Left(Combo1.Text, 1)
      strKEY02 = txtBS02
   'Add By Sindy 2023/8/21
   ElseIf SSTab1.Tab = 2 Then
      strKEY01 = txtTBNP01_M
      strKEY02 = txtTBNP09 'Add By Sindy 2023/9/26
      'Add By Sindy 2023/8/25
      If strKEY01 = "" And strKEY02 = "" Then
         ReadAllData
         QueryRecord = True
         Exit Function
      End If
      '2023/8/25 END
   End If
   
   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
      QueryRecord = True
      UpdateCtrlData
'      ReadAllData
   Else
      QueryRecord = False
      'Add By Sindy 2017/11/16
      If SSTab1.Tab = 1 Then
         QueryRecord = QueryRecordGrd2
      End If
      '2017/11/16 END
   End If
   
   UpdateToolbarState
End Function

'Add By Sindy 2017/11/16 查詢公報特殊字對照檔
Private Function QueryRecordGrd2() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCon As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   strCon = ""
   If Left(Combo1.Text, 1) <> "" Then
      If strCon <> "" Then strCon = strCon & " and"
      strCon = strCon & " BS01='" & Left(Combo1.Text, 1) & "'"
   End If
   If txtBS02 <> "" Then
      If strCon <> "" Then strCon = strCon & " and"
      strCon = strCon & " BS02='" & txtBS02 & "'"
   End If
   
   strSql = "SELECT ' ',decode(BS01,'1','1申請人名稱','2','2代理人',BS01),BS02,BS03" & _
            " FROM BulletinSpecWord where" & strCon & _
            " order by BS01 asc,BS02 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      grd2(SSTab1.Tab).Rows = 2
      grd2(SSTab1.Tab).Clear
      QueryRecordGrd2 = True
      Set grd2(SSTab1.Tab).Recordset = rsTmp
      Call SetDataListWidth2(SSTab1.Tab)
      Call GetSelChage3(1)
   Else
      QueryRecordGrd2 = False
   End If
   rsTmp.Close
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse
   
   m_strFindKey01 = "" 'Add By Sindy 2023/8/25
   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         If SaveData(m_EditMode) = True Then
             RefreshRange
             ReadAllData
             SetKeyReadOnly True
         Else
             Exit Function
         End If
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         If SaveData(m_EditMode) = False Then Exit Function
         ReadAllData
         SetKeyReadOnly True
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
            ReadAllData
            SetKeyReadOnly True
         Else
            Exit Function
         End If
      Case 4: '查詢
         'Modify By Sindy 2017/11/16
         If SSTab1.Tab = 0 Then
            If txtTBNP01.Text = "" Then
                MsgBox "特定公司不可以空白！", vbExclamation
                txtTBNP01.SetFocus
                GoTo EXITSUB
            Else
               m_strFindKey01 = txtTBNP01.Text 'Add By Sindy 2023/8/25
            End If
         ElseIf SSTab1.Tab = 1 Then
            If Combo1.ListIndex = -1 Then
                MsgBox "欄位不可以空白！", vbExclamation
                Combo1.SetFocus
                GoTo EXITSUB
            End If
         'Add By Sindy 2023/8/21
         ElseIf SSTab1.Tab = 2 Then
            If txtTBNP01_M.Text = "" Then
               '空白代表查詢全部資料
'                MsgBox "E-Mail不可以空白！", vbExclamation
'                txtTBNP01_M.SetFocus
'                GoTo EXITSUB
            Else
               m_strFindKey01 = txtTBNP01_M.Text 'Add By Sindy 2023/8/25
            End If
            'txtTBNP01_M.Tag = txtTBNP01_M.Text
         End If
         '2017/11/16 END
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            'Add By Sindy 2023/8/25
            m_CurrKEY(0) = ""
            m_CurrKEY(1) = ""
            m_strFindKey01 = ""
            If SSTab1.Tab = 2 Then
               Call ReadAllData
            End If
            '2023/8/25 END
            UpdateCtrlData
         End If
         'txtTBNP01_M.Text = txtTBNP01_M.Tag 'Add By Sindy 2023/8/25
         SetKeyReadOnly True
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True

EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 0, 1, 4:
         If SSTab1.Tab = 0 Then
            If Me.txtTBNP01.Visible = True Then txtTBNP01.SetFocus
         ElseIf SSTab1.Tab = 1 Then
            If Me.Combo1.Visible = True Then Combo1.SetFocus
            If Me.txtBS02.Visible = True Then txtBS02.SetFocus
         'Add By Sindy 2023/8/21
         ElseIf SSTab1.Tab = 2 Then
            If Me.txtTBNP01_M.Visible = True Then txtTBNP01_M.SetFocus
         End If
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, Optional ByVal strKEY02 As String = "") As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   
   If SSTab1.Tab = 0 Then
      'Modify By Sindy 2013/3/1 +TBNP08
      strSql = "SELECT * FROM TMBulletinNp WHERE ltrim(rtrim(TBNP01))=ltrim(rtrim(" & CNULL(strKEY01 & " ") & ")) and TBNP08='T'"
   ElseIf SSTab1.Tab = 1 Then
      strSql = "SELECT * FROM BulletinSpecWord WHERE BS01='" & strKEY01 & "' and ltrim(rtrim(BS02))=ltrim(rtrim(" & CNULL(strKEY02 & " ") & "))"
   'Add By Sindy 2023/8/21
   ElseIf SSTab1.Tab = 2 Then  '電子報特殊名單
      If m_strFindKey01 <> "" And m_strFindKey01 = strKEY01 Then
         strSql = "SELECT * FROM TMBulletinNp WHERE instr(NLS_Upper(TBNP01),'" & UCase(ChgSQL(m_strFindKey01)) & "') > 0 and TBNP08='M'" & IIf(strKEY02 <> "", " and TBNP09=" & strKEY02, "")
      ElseIf strKEY01 <> "" Then
         strSql = "SELECT * FROM TMBulletinNp WHERE upper(ltrim(rtrim(TBNP01)))=upper(ltrim(rtrim(" & CNULL(strKEY01 & " ") & "))) and TBNP08='M'" & IIf(strKEY02 <> "", " and TBNP09=" & strKEY02, "")
      Else
         strSql = "SELECT * FROM TMBulletinNp WHERE TBNP08='M'" & IIf(strKEY02 <> "", " and TBNP09=" & strKEY02, "")
      End If
   End If
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
Private Sub ShowCurrRecord(ByVal strKEY01 As String, Optional ByVal strKEY02 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTBNP08 As String 'Add By Sindy 2023/8/21
   
   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      'Modify By Sindy 2023/8/21
      If SSTab1.Tab = 0 Or SSTab1.Tab = 2 Then
         If SSTab1.Tab = 0 Then
            strTBNP08 = "T"
         ElseIf SSTab1.Tab = 2 Then
            strTBNP08 = "M"
         End If
         '2023/8/21 END
         
         'Modify By Sindy 2013/3/1 +TBNP08
         strSql = "SELECT TBNP01 FROM TMBulletinNp WHERE TBNP01='" & m_CurrKEY(0) & "' and TBNP08='" & strTBNP08 & "'"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
            rsTmp.Close
            UpdateCtrlData
            GoTo EXITSUB
         End If
         rsTmp.Close
         'Modify By Sindy 2013/3/1 +TBNP08
         strSql = "SELECT TBNP01 FROM TMBulletinNp WHERE TBNP08='" & strTBNP08 & "' order by TBNP01 asc"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         Else
            ShowLastRecord
            GoTo EXITSUB
         End If
         rsTmp.Close
         
      'Add By Sindy 2017/5/5
      ElseIf SSTab1.Tab = 1 Then
         strSql = "SELECT BS01,BS02 FROM BulletinSpecWord WHERE BS01='" & m_CurrKEY(0) & "' and BS02='" & m_CurrKEY(1) & "'"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
            If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
            rsTmp.Close
            UpdateCtrlData
            GoTo EXITSUB
         End If
         rsTmp.Close
         strSql = "SELECT BS01,BS02 FROM BulletinSpecWord WHERE BS01='" & m_CurrKEY(0) & "' order by BS01 asc,BS02 asc"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
            If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
         Else
            ShowLastRecord
            GoTo EXITSUB
         End If
         rsTmp.Close
         strSql = "SELECT BS01,BS02 FROM BulletinSpecWord order by BS01 asc,BS02 asc"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
            If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
         Else
            ShowLastRecord
            GoTo EXITSUB
         End If
         rsTmp.Close
      '2017/5/5 END
      End If
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
   Dim strTBNP08 As String 'Add By Sindy 2023/8/21
   
   If SSTab1.Tab = 0 Or SSTab1.Tab = 2 Then
      If m_CurrKEY(0) = m_FirstKEY(0) Then
         ShowMsg MsgText(9008)
         GoTo EXITSUB
      End If
      'Modify By Sindy 2023/8/21
      If SSTab1.Tab = 0 Then
         strTBNP08 = "T"
      ElseIf SSTab1.Tab = 2 Then
         strTBNP08 = "M"
      End If
      '2023/8/21 END
      
      'Modify By Sindy 2013/3/1 +TBNP08
      strSql = "SELECT TBNP01 FROM TMBulletinNp WHERE TBNP01<'" & m_CurrKEY(0) & "' and TBNP08='" & strTBNP08 & "' order by TBNP01 desc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      'Modify By Sindy 2013/3/1 +TBNP08
      strSql = "SELECT TBNP01 FROM TMBulletinNp WHERE TBNP08='" & strTBNP08 & "' order by TBNP01 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      End If
      rsTmp.Close
      
   'Add By Sindy 2017/5/5
   ElseIf SSTab1.Tab = 1 Then
      If m_CurrKEY(0) & m_CurrKEY(1) = m_FirstKEY(0) & m_FirstKEY(1) Then
         ShowMsg MsgText(9008)
         GoTo EXITSUB
      End If
      strSql = "SELECT BS01,BS02 FROM BulletinSpecWord WHERE BS01||BS02<'" & m_CurrKEY(0) & m_CurrKEY(1) & "' order by BS01 desc,BS02 desc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT BS01,BS02 FROM BulletinSpecWord order by BS01 asc,BS02 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      End If
      rsTmp.Close
   '2017/5/5 END
   End If
   
   UpdateCtrlData

EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTBNP08 As String 'Add By Sindy 2023/8/21
   
   If SSTab1.Tab = 0 Or SSTab1.Tab = 2 Then
      If m_CurrKEY(0) = m_LastKEY(0) Then
         ShowMsg MsgText(9009)
         GoTo EXITSUB
      End If
      'Modify By Sindy 2023/8/21
      If SSTab1.Tab = 0 Then
         strTBNP08 = "T"
      ElseIf SSTab1.Tab = 2 Then
         strTBNP08 = "M"
      End If
      '2023/8/21 END
      
      'Modify By Sindy 2013/3/1 +TBNP08
      strSql = "SELECT TBNP01 FROM TMBulletinNp WHERE TBNP01>'" & m_CurrKEY(0) & "' and TBNP08='" & strTBNP08 & "' order by TBNP01 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      'Modify By Sindy 2013/3/1 +TBNP08
      strSql = "SELECT TBNP01 FROM TMBulletinNp WHERE TBNP08='" & strTBNP08 & "' order by TBNP01 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      End If
      rsTmp.Close
      
   ElseIf SSTab1.Tab = 1 Then
      If m_CurrKEY(0) & m_CurrKEY(1) = m_LastKEY(0) & m_LastKEY(1) Then
         ShowMsg MsgText(9009)
         GoTo EXITSUB
      End If
      strSql = "SELECT BS01,BS02 FROM BulletinSpecWord WHERE BS01||BS02>'" & m_CurrKEY(0) & m_CurrKEY(1) & "' order by BS01 asc,BS02 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT BS01,BS02 FROM BulletinSpecWord order by BS01 asc,BS02 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      End If
      rsTmp.Close
   End If
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
         SetKeyReadOnly False
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
         'Add By Sindy 2023/8/21
         If SSTab1.Tab = 2 Then
            MsgBox "電子報特殊名單不提供刪除，請針對狀況更新[是否寄電子報]！", vbExclamation
            Exit Sub
         End If
         '2023/8/21 END
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
         PUB_FilterFormText Me 'Add By Sindy 2017/7/31 防止儲存到跳行符號
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
                  SetKeyReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               SetKeyReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   'Modify By Sindy 2017/11/16
   If SSTab1.Tab = 0 Then
      txtTBNP01.Locked = bEnable
      If bEnable Then txtTBNP01.BackColor = &H8000000F Else txtTBNP01.BackColor = &H80000005
   ElseIf SSTab1.Tab = 1 Then
      Combo1.Locked = bEnable
      If bEnable Then Combo1.BackColor = &H8000000F Else Combo1.BackColor = &H80000005
      txtBS02.Locked = bEnable
      If bEnable Then txtBS02.BackColor = &H8000000F Else txtBS02.BackColor = &H80000005
   'Add By Sindy 2023/8/21
   ElseIf SSTab1.Tab = 2 Then
      txtTBNP01_M.Locked = bEnable
      If bEnable Then txtTBNP01_M.BackColor = &H8000000F Else txtTBNP01_M.BackColor = &H80000005
   End If
   '2017/11/16 END
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   If SSTab1.Tab = 0 Then
      txtTBNP01.Locked = bEnable
      If bEnable Then txtTBNP01.BackColor = &H8000000F Else txtTBNP01.BackColor = &H80000005
   ElseIf SSTab1.Tab = 1 Then
      If m_EditMode = 1 Then '新增
         Combo1.Locked = False
         Combo1.BackColor = &H80000005
         txtBS02.Locked = False
         txtBS02.BackColor = &H80000005
      Else
         Combo1.Locked = True
         Combo1.BackColor = &H8000000F
         txtBS02.Locked = True
         txtBS02.BackColor = &H8000000F
      End If
      txtBS03.Locked = bEnable
      If bEnable Then txtBS03.BackColor = &H8000000F Else txtBS03.BackColor = &H80000005
   'Add By Sindy 2023/8/21
   ElseIf SSTab1.Tab = 2 Then
      txtTBNP10.Locked = bEnable
      If bEnable Then txtTBNP10.BackColor = &H8000000F Else txtTBNP10.BackColor = &H80000005
      '2023/8/21 END
   End If
End Sub

Private Sub ClearField()
   If SSTab1.Tab = 0 Then
      txtTBNP01 = Empty
   ElseIf SSTab1.Tab = 1 Then
      If m_EditMode = 1 Then '新增
         Combo1.ListIndex = 0
      Else
         Combo1.ListIndex = -1
      End If
      txtBS02 = Empty
      txtBS03 = Empty
   'Add By Sindy 2023/8/21
   ElseIf SSTab1.Tab = 2 Then
      txtTBNP01_M = Empty
      Me.txtTBNP10 = Empty
      Me.txtTBNP09 = Empty
   End If
End Sub

'將資料庫中的資料更新到所有欄位中
Private Sub ReadAllData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

   Screen.MousePointer = vbHourglass
   Me.Enabled = False
      
   If SSTab1.Tab = 0 Then
      grd1.Rows = 2
      grd1.Clear
      'Modify By Sindy 2013/3/1 +TBNP08
      'Modify By Sindy 2022/5/4 +建檔人員,建檔時間
      strSql = "select TBNP01,st02,sqldatet(tbnp03)||' '||sqltime(tbnp04||'00') " & _
               "from TMBulletinNp,staff " & _
               "where TBNP08='T' and tbnp02=st01(+) " & _
               "order by TBNP01 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         Set grd1.Recordset = rsTmp
      End If
      rsTmp.Close
      SetDataListWidth
      GetSelChage
   End If
   
   grd2(SSTab1.Tab).Rows = 2
   grd2(SSTab1.Tab).Clear
   If SSTab1.Tab = 0 Then
      'Modify By Sindy 2013/3/1 +TBNP08
      strSql = "SELECT ' ',a.TBOR03,TBNP01 " & _
               "FROM TMBulletinNp,(SELECT TBOR03 FROM TMBulletinOwner group by TBOR03) a " & _
               "WHERE TBNP01(+)=a.TBOR03 and TBNP08(+)='T' " & _
               "order by a.TBOR03 asc "
   ElseIf SSTab1.Tab = 1 Then
      strSql = "SELECT ' ',decode(BS01,'1','1申請人名稱','2','2代理人',BS01),BS02,BS03 " & _
               "FROM BulletinSpecWord " & _
               "order by BS01 asc,BS02 asc"
   'Modify By Sindy 2023/8/21
   ElseIf SSTab1.Tab = 2 Then
      strSql = "SELECT ' ',TBNP01,TBNP10,TBNP09 " & _
               "FROM TMBulletinNp " & _
               "WHERE TBNP08='M' " & IIf(Trim(m_strFindKey01) <> "", "and instr(NLS_Upper(TBNP01),'" & UCase(ChgSQL(m_strFindKey01)) & "') > 0 ", "") & _
               "order by TBNP01 asc "
      If m_strFindKey01 = "" Then m_CurrKEY(0) = ""
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If SSTab1.Tab = 2 Then LblCnt.Caption = "共 " & rsTmp.RecordCount & " 筆" 'Add By Sindy 2023/8/25
   If rsTmp.RecordCount > 0 Then
      Set grd2(SSTab1.Tab).Recordset = rsTmp
      If rsTmp.RecordCount > 1 Then m_strFindKey01 = "" 'Modify By Sindy 2023/8/25
   End If
   rsTmp.Close
   Call SetDataListWidth2(SSTab1.Tab)
   If SSTab1.Tab = 0 Then
      GetSelChage2
   ElseIf SSTab1.Tab = 1 Then
      Call GetSelChage3(1)
   ElseIf SSTab1.Tab = 2 Then
      Call GetSelChage3(2)
   End If
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub GetSelChage()
grd1.Visible = False
If grd1.Rows - 1 > 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      grd1.col = 0
      grd1.row = dblPrevRow
      For i = 0 To grd1.Cols - 1
         grd1.col = i
         grd1.CellBackColor = QBColor(15)
      Next i
   End If
   '尋找目前資料列
   For j = 1 To grd1.Rows - 1
      If grd1.TextMatrix(j, 0) = m_CurrKEY(0) Then
         If m_strFindKey01 <> "" Then grd1.TopRow = j '移動捲軸 Add By Sindy 2023/8/25
         grd1.col = 0
         grd1.row = j
         dblPrevRow = grd1.row
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = &HFFC0C0
         Next i
         Exit For
      End If
   Next j
End If
grd1.Visible = True
End Sub

Private Sub GetSelChage2()
grd2(0).Visible = False
If grd2(0).Rows - 1 > 0 Then
   '已是不列印者顯示為淺綠色
   For j = 1 To grd2(0).Rows - 1
      If grd2(0).TextMatrix(j, 2) <> "" Then
         grd2(0).col = 0
         grd2(0).row = j
         For i = 0 To grd2(0).Cols - 1
            grd2(0).col = i
            grd2(0).CellBackColor = &H80FF80
         Next i
      End If
   Next j
End If
grd2(0).Visible = True
End Sub

'Add By Sindy 2017/5/8
Private Sub GetSelChage3(Index As Integer)
Dim strKey1 As String, StrKey2 As String

grd2(Index).Visible = False
If grd2(Index).Rows - 1 > 0 Then
   '上一筆資料列清除反白
   If dblPrevRow2 > 0 And dblPrevRow2 <= grd2(Index).Rows Then
      grd2(Index).col = 0
      grd2(Index).row = dblPrevRow2
      grd2(Index).Text = ""
      For i = 0 To grd2(Index).Cols - 1
         grd2(Index).col = i
         grd2(Index).CellBackColor = QBColor(15)
      Next i
   Else
      dblPrevRow2 = 0
   End If
   '尋找目前資料列
   If m_CurrKEY(0) <> "" Then
      For j = 1 To grd2(Index).Rows - 1
         If Index = 1 Then
            strKey1 = Left(grd2(Index).TextMatrix(j, 1), 1)
            StrKey2 = grd2(Index).TextMatrix(j, 2)
         ElseIf Index = 2 Then
            strKey1 = grd2(Index).TextMatrix(j, 1)
            StrKey2 = grd2(Index).TextMatrix(j, 3)
         Else
            strKey1 = grd2(Index).TextMatrix(j, 1)
            StrKey2 = ""
         End If
         If strKey1 = m_CurrKEY(0) And StrKey2 = m_CurrKEY(1) Then
            If m_strFindKey01 <> "" Then grd2(Index).TopRow = j '移動捲軸 Add By Sindy 2023/8/25
            grd2(Index).col = 0
            grd2(Index).row = j
            dblPrevRow2 = grd2(Index).row
            For i = 0 To grd2(Index).Cols - 1
               grd2(Index).col = i
               grd2(Index).CellBackColor = &HFFC0C0
            Next i
            Exit For
         End If
      Next j
   End If
End If
grd2(Index).Visible = True
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   ClearField
   
   If SSTab1.Tab = 0 Then
      m_CurrKEY(1) = "" 'Add By Sindy 2023/8/21
      'Modify By Sindy 2013/3/1 +TBNP08
      strSql = "SELECT TBNP01 " & _
               "FROM TMBulletinNp " & _
               "WHERE ltrim(rtrim(TBNP01))=ltrim(rtrim('" & m_CurrKEY(0) & " " & "')) and TBNP08='T'"
   ElseIf SSTab1.Tab = 1 Then
      strSql = "SELECT decode(BS01,'1','1申請人名稱','2','2代理人',BS01),BS02,BS03 " & _
               "FROM BulletinSpecWord " & _
               "WHERE BS01='" & m_CurrKEY(0) & "' and ltrim(rtrim(BS02))=ltrim(rtrim('" & m_CurrKEY(1) & " " & "'))"
   'Add By Sindy 2023/8/21
   ElseIf SSTab1.Tab = 2 Then
      If m_strFindKey01 <> "" And m_strFindKey01 = m_CurrKEY(0) Then
         strSql = "SELECT TBNP01,TBNP10,TBNP09 " & _
                  "FROM TMBulletinNp " & _
                  "WHERE instr(NLS_Upper(TBNP01),'" & UCase(ChgSQL(m_CurrKEY(0))) & "') > 0 and TBNP08='M'"
      ElseIf m_CurrKEY(0) <> "" Then
         strSql = "SELECT TBNP01,TBNP10,TBNP09 " & _
                  "FROM TMBulletinNp " & _
                  "WHERE ltrim(rtrim(TBNP01))=ltrim(rtrim('" & m_CurrKEY(0) & "')) and TBNP08='M'" & IIf(m_CurrKEY(1) <> "", " and TBNP09=" & m_CurrKEY(1), "")
      Else
         strSql = "SELECT TBNP01,TBNP10,TBNP09 " & _
                  "FROM TMBulletinNp " & _
                  "WHERE TBNP08='M'" & IIf(m_CurrKEY(1) <> "", " and TBNP09=" & m_CurrKEY(1), "")
      End If
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If SSTab1.Tab = 0 Then
         If IsNull(rsTmp.Fields("TBNP01")) = False Then: txtTBNP01 = rsTmp.Fields("TBNP01")
      ElseIf SSTab1.Tab = 1 Then
         If IsNull(rsTmp.Fields(0)) = False Then: Combo1.ListIndex = Val(Left(rsTmp.Fields(0), 1)) - 1
         If IsNull(rsTmp.Fields("BS02")) = False Then: txtBS02 = rsTmp.Fields("BS02")
         If IsNull(rsTmp.Fields("BS03")) = False Then: txtBS03 = rsTmp.Fields("BS03")
      'Add By Sindy 2023/8/21
      ElseIf SSTab1.Tab = 2 Then
         If IsNull(rsTmp.Fields("TBNP01")) = False Then: txtTBNP01_M = rsTmp.Fields("TBNP01")
         Me.txtTBNP10 = "" & rsTmp.Fields("TBNP10")
         Me.txtTBNP09 = "" & rsTmp.Fields("TBNP09")
         If m_CurrKEY(1) <> "" Then
            m_strFindKey01 = txtTBNP01_M
            m_CurrKEY(0) = txtTBNP01_M
            m_CurrKEY(1) = txtTBNP09
         ElseIf m_CurrKEY(0) <> "" Then
            m_CurrKEY(1) = txtTBNP09
         End If
         '2023/8/21 END
      End If
   End If
   rsTmp.Close
   If SSTab1.Tab = 0 Then
      GetSelChage
   ElseIf SSTab1.Tab = 1 Then
      Call GetSelChage3(1)
   ElseIf SSTab1.Tab = 2 Then
      'Add By Sindy 2023/8/25
      If m_strFindKey01 <> "" Then
         Call ReadAllData
      End If
      '2023/8/25 END
      Call GetSelChage3(2)
   End If

   Me.Enabled = True
   Screen.MousePointer = vbDefault

EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTBNP08 As String
   
   'Modify By Sindy 2023/8/21
   If SSTab1.Tab = 0 Or SSTab1.Tab = 2 Then
      If SSTab1.Tab = 0 Then
         strTBNP08 = "T"
      ElseIf SSTab1.Tab = 2 Then
         strTBNP08 = "M"
      End If
      '2023/8/21 END
      
      'Modify By Sindy 2013/3/1 +TBNP08
      strSql = "select TBNP01 from TMBulletinNp WHERE TBNP08='" & strTBNP08 & "' order by TBNP01 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then m_FirstKEY(0) = rsTmp.Fields(0)
      End If
      rsTmp.Close
      
      'Modify By Sindy 2013/3/1 +TBNP08
      strSql = "select TBNP01 from TMBulletinNp WHERE TBNP08='" & strTBNP08 & "' order by TBNP01 desc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then m_LastKEY(0) = rsTmp.Fields(0)
      End If
      rsTmp.Close
      
   'Add By Sindy 2017/5/5
   ElseIf SSTab1.Tab = 1 Then
      strSql = "select BS01,BS02 from BulletinSpecWord order by BS01 asc,BS02 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then m_FirstKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then m_FirstKEY(1) = rsTmp.Fields(1)
      End If
      rsTmp.Close
      
      strSql = "select BS01,BS02 from BulletinSpecWord order by BS01 desc,BS02 desc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then m_LastKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then m_LastKEY(1) = rsTmp.Fields(1)
      End If
      rsTmp.Close
      '2017/5/5 END
   End If
   
   Set rsTmp = Nothing
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

Private Sub SetDataListWidth()
grd1.row = 0
grd1.col = 0: grd1.Text = "不列印者"
grd1.ColWidth(0) = 3000
grd1.CellAlignment = flexAlignLeftCenter
'Add By Sindy 2022/5/4
grd1.col = 1: grd1.Text = "建檔人員"
grd1.ColWidth(1) = 800
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 2: grd1.Text = "建檔時間"
grd1.ColWidth(2) = 1500
grd1.CellAlignment = flexAlignLeftCenter
'2022/5/4 END
End Sub

Private Sub SetDataListWidth2(Index As Integer)
   If Index = 0 Or Index = 2 Then
      grd2(Index).row = 0
      grd2(Index).col = 0: grd2(Index).Text = "V"
      grd2(Index).ColWidth(0) = 200
      grd2(Index).CellAlignment = flexAlignLeftCenter
      If Index = 2 Then
         grd2(Index).col = 1: grd2(Index).Text = "E-Mail"
         grd2(Index).ColWidth(1) = 2200
         grd2(Index).CellAlignment = flexAlignLeftCenter
         grd2(Index).col = 2: grd2(Index).Text = "是否寄電子報"
         grd2(Index).ColWidth(2) = 1000
         grd2(Index).CellAlignment = flexAlignLeftCenter
         grd2(Index).col = 3: grd2(Index).Text = "編號"
         grd2(Index).ColWidth(3) = 800
         grd2(Index).CellAlignment = flexAlignLeftCenter
      Else
         grd2(Index).col = 1: grd2(Index).Text = "指定為不列印的特定公司"
         grd2(Index).ColWidth(1) = 3700
         grd2(Index).CellAlignment = flexAlignLeftCenter
         grd2(Index).col = 2: grd2(Index).Text = "TBNP01"
         grd2(Index).ColWidth(2) = 0
         grd2(Index).CellAlignment = flexAlignLeftCenter
      End If
   ElseIf Index = 1 Then
      grd2(Index).row = 0
      grd2(Index).col = 0: grd2(Index).Text = "V"
      grd2(Index).ColWidth(0) = 200
      grd2(Index).CellAlignment = flexAlignLeftCenter
      grd2(Index).col = 1: grd2(Index).Text = "欄位"
      grd2(Index).ColWidth(1) = 1200
      grd2(Index).CellAlignment = flexAlignLeftCenter
      grd2(Index).col = 2: grd2(Index).Text = "轉檔前"
      grd2(Index).ColWidth(2) = 3500
      grd2(Index).CellAlignment = flexAlignLeftCenter
      grd2(Index).col = 3: grd2(Index).Text = "轉檔後"
      grd2(Index).ColWidth(3) = 3500
      grd2(Index).CellAlignment = flexAlignLeftCenter
   End If
End Sub

'Add By Sindy 2017/5/4
Private Sub txtBS02_GotFocus()
   InverseTextBox txtBS02
   OpenIme
End Sub
Private Sub txtBS02_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtBS02_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If txtBS02.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtBS02, txtBS02.MaxLength) Then
      Cancel = True
   End If
   If m_EditMode = 1 Then
      ' 檢查記錄是否已存在
      If IsRecordExist(Left(Combo1.Text, 1), txtBS02) = True Then
         MsgBox "該筆記錄已存在", vbInformation
         Call txtBS02_GotFocus
         Cancel = True
         Exit Sub
      End If
      If Left(Combo1, 1) = "1" Then '申請人名稱
         If Len(Trim(txtBS02.Text)) <= 3 Then
            MsgBox "申請人名稱不可小於等於3碼！", vbExclamation
            Call txtBS02_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
      'Add By Sindy 2018/8/2
      If InStr(txtBS02.Text, "?") = 0 Then
         MsgBox "需有造字?存在!", vbInformation
         Call txtBS02_GotFocus
         Cancel = True
         Exit Sub
      End If
      '2018/8/2 END
   End If
End Sub
Private Sub txtBS03_GotFocus()
   InverseTextBox txtBS03
   OpenIme
End Sub
Private Sub txtBS03_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtBS03_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If txtBS03.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtBS03, txtBS03.MaxLength) Then
      Cancel = True
   End If
End Sub
'2017/5/4 END

Private Sub txtTBNP01_GotFocus()
   InverseTextBox txtTBNP01
End Sub

Private Sub txtTBNP01_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtTBNP01_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And txtTBNP01 <> "" Then
      If m_EditMode = 1 And txtTBNP01 <> "" Then
         ' 檢查記錄是否已存在
         If IsRecordExist(txtTBNP01) = True Then
            MsgBox "該筆記錄已存在", vbInformation
            Call txtTBNP01_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

'Add By Sindy 2023/8/21
Private Sub txtTBNP10_GotFocus()
   CloseIme
   TextInverse txtTBNP10
End Sub
Private Sub txtTBNP10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("D") Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub txtTBNP10_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If txtTBNP10.Text = "" Then Exit Sub
   If txtTBNP10.Text <> "N" And txtTBNP10.Text <> "D" Then
      ShowMsg "輸入錯誤 !"
      Cancel = True
   End If
End Sub
Private Sub txtTBNP01_M_GotFocus()
   InverseTextBox txtTBNP01_M
End Sub
Private Sub txtTBNP01_M_KeyPress(KeyAscii As MSForms.ReturnInteger)
   'KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtTBNP01_M_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And txtTBNP01_M <> "" Then
      If m_EditMode = 1 And txtTBNP01_M <> "" Then
         ' 檢查記錄是否已存在
         If IsRecordExist(txtTBNP01_M) = True Then
            MsgBox "該筆記錄已存在", vbInformation
            Call txtTBNP01_M_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub
'2023/8/21 END
