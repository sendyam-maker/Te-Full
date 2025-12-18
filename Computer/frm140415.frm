VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140415 
   BorderStyle     =   1  '單線固定
   Caption         =   "各項指示分類維護"
   ClientHeight    =   5484
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8292
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5484
   ScaleWidth      =   8292
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7695
      Top             =   450
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
            Picture         =   "frm140415.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140415.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140415.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140415.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140415.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140415.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140415.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140415.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140415.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140415.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140415.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8292
      _ExtentX        =   14626
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4710
      Left            =   90
      TabIndex        =   7
      Top             =   720
      Width           =   8115
      _ExtentX        =   14309
      _ExtentY        =   8319
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "資料維護"
      TabPicture(0)   =   "frm140415.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textCUID"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtField(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtField(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtField(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtField(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtField(11)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtField(12)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm140415.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MGrid1"
      Tab(1).ControlCount=   1
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   5
         Top             =   3570
         Width           =   600
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   11
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   4
         Top             =   2550
         Width           =   6240
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   3
         Top             =   2160
         Width           =   600
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Index           =   3
         Left            =   1320
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1410
         Width           =   6240
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1000
         Width           =   600
      End
      Begin VB.TextBox txtField 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   0
         Top             =   540
         Width           =   600
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
         Bindings        =   "frm140415.frx":212C
         Height          =   4125
         Left            =   -74910
         TabIndex        =   8
         Top             =   420
         Width           =   7905
         _ExtentX        =   13949
         _ExtentY        =   7281
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSForms.TextBox textCUID 
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   4320
         Width           =   7860
         VariousPropertyBits=   671105055
         Size            =   "13864;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(D：顯示民國年/月/日  N：讀取X/Y編號的名稱)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   1980
         TabIndex        =   18
         Top             =   3623
         Width           =   4650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "顯示格式："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   135
         TabIndex        =   17
         Top             =   3623
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "P.S 基本檔對應欄位的欄位名稱後面請加上;"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   7
         Left            =   1320
         TabIndex        =   16
         Top             =   3270
         Width           =   4245
      End
      Begin VB.Label Label1 
         Caption         =   "基  本  檔  對應欄位："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   6
         Left            =   135
         TabIndex        =   15
         Top             =   2610
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(P專利、T商標)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   1980
         TabIndex        =   14
         Top             =   2220
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "使用部門："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   135
         TabIndex        =   13
         Top             =   2220
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "P.S.分類的輸入和顯示在Table(AddressA4List.INTYPE)的控制"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   1980
         TabIndex        =   12
         Top             =   540
         Width           =   5940
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分類部門："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   585
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分類代號："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   1050
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "說　　明："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   1485
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frm140415"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/11/23 改成Form2.0 ; textCUID
'Created by Lydia 2016/11/09 各項指示分類維護
Option Explicit

Dim m_EditMode As Integer '0:瀏覽 1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim oText As TextBox
Dim strKind(1 To 3) As String 'Added by Lydia 2020/05/12 分類第1碼: 改用Table控制

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Screen.MousePointer = vbHourglass
   Select Case KeyCode
      Case vbKeyF2 '新增
         KeyCode = 0: Action 1
      Case vbKeyF3 '修改
         KeyCode = 0: Action 2
      Case vbKeyF4: '查詢
         KeyCode = 0: Action 4
      Case vbKeyF5 '刪除
         KeyCode = 0: Action 3
      Case vbKeyHome '第一筆
         KeyCode = 0: Action 6
      Case vbKeyPageUp '上一筆
         KeyCode = 0: Action 7
      Case vbKeyPageDown '下一筆
         KeyCode = 0: Action 8
      Case vbKeyEnd: '最後筆
         KeyCode = 0: Action 9
      Case vbKeyF9, vbKeyReturn '確定
         KeyCode = 0: Action 11
      Case vbKeyF10 '取消
         KeyCode = 0: Action 12
      Case vbKeyEscape '結束
         KeyCode = 0: Action 14
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
  
   MoveFormToCenter Me
   
   'Added by Lydia 2020/05/12 分類第1碼: 改用Table控制
   strKind(1) = PUB_GetInType("1")
   strKind(2) = PUB_GetInType("2")
   strKind(3) = PUB_GetInType("3")
   Label1(3).Caption = "(" & strKind(2) & ")"
   'end 2020/05/12
   
   Action 6 '預設第一筆
   Call SetGrid(True)
   UpdateToolbarState
   
   textCUID.BackColor = &H8000000F

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140415 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
   Action Button.Index
   Screen.MousePointer = vbDefault
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
         If m_bUpdate And txtField(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And txtField(1) <> "" Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtField(1) <> "" Then
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

Private Sub TxtLock()
   Select Case m_EditMode
   Case 0 '瀏覽
      For Each oText In txtField
         oText.Locked = True
      Next
      SSTab1.TabEnabled(1) = True
   Case Else
      If m_EditMode = 4 Or m_EditMode = 1 Then
         For Each oText In txtField
           oText.Locked = False
         Next
         txtField(1).SetFocus
      Else
         txtField(3).Locked = False
         txtField(3).SetFocus
         txtField(10).Locked = False  'Added by Lydia 2020/05/13 使用部門IT10
         'Added by Lydia 2020/05/14 基本檔對應欄位IT11,IT12
         txtField(11).Locked = False
         txtField(12).Locked = False
      End If
      
      SSTab1.TabEnabled(1) = False
   End Select
End Sub
Private Sub Action(Index As Integer)
   
   If TBar1.Buttons(Index).Enabled = False Then Exit Sub

On Error GoTo ErrHand

   SSTab1.Tab = 0
   Select Case Index
      Case 1 '按下新增
        m_EditMode = 1
        FormReset
        SSTab1.TabEnabled(1) = False
        UpdateCUID 0
      Case 2 '按下修改
        m_EditMode = 2
        SSTab1.TabEnabled(1) = False
      Case 3 '按下刪除
         If txtField(1).Text = "" Or txtField(2).Text = "" Then
             MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
             Exit Sub
         End If
         '檢查各項指示檔Instructions
         strExc(0) = "select count(*) from Instructions where ITS03='" & txtField(1) & txtField(2) & "' and NVL(ITS05,'Y')='Y' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Val(RsTemp(0)) > 0 Then
               If MsgBox("尚有" & RsTemp(0) & "筆客戶/代理人/案件各項指示資料用到此分類，確定要繼續刪除？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
            End If
         End If
         If DelMsg() = True Then
            If FormDelete() = False Then
               MsgBox "刪除失敗!", vbCritical
               Exit Sub
            '刪除後移到最末筆
            Else
               ShowRecord 3
               Call SetGrid(False)
            End If
         End If

      Case 4 '按下查詢
         FormReset
         m_EditMode = 4
      Case 6 '第一筆
         ShowRecord 0
      Case 7 '前一筆
         ShowRecord 1
      Case 8 '後一筆
         ShowRecord 2
      Case 9 '最後筆
         ShowRecord 3
      Case 11 '按下確定
         Select Case m_EditMode
            '新增,修改
            Case 1, 2
               If TxtValidate = False Then
                  Exit Sub
               Else
                 If txtField(1).Text <> txtField(1).Tag Or txtField(2).Text <> txtField(2).Tag Then
                    If RecIsExist = True Then Exit Sub
                 End If
                  If FormSave() = False Then
                     MsgBox "存檔失敗!", vbCritical
                     Exit Sub
                  Else
                     m_EditMode = 0
                     For Each oText In txtField
                         oText.Tag = oText.Text
                     Next
                     ReadData txtField(1), txtField(2)
                     Call SetGrid(False)
                  End If
               End If
               SSTab1.TabEnabled(1) = True
            '查詢
            Case 4
               If ReadData(txtField(1), txtField(2)) = False Then
                  MsgBox "無資料!", vbExclamation
                  Exit Sub
               Else
                  m_EditMode = 0
               End If
         End Select
      Case 12 '按下取消
         m_EditMode = 0
         SSTab1.TabEnabled(1) = True
         For Each oText In txtField
             oText.Text = oText.Tag
         Next
         If txtField(1) <> "" And txtField(2) <> "" Then
            If ReadData(txtField(1), txtField(2)) = False Then
               ShowRecord 3
            End If
         End If
      Case 14 '結束
         Unload Me
         Exit Sub
   End Select
   UpdateToolbarState
   TxtLock
   Exit Sub
   
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

' 顯示資料
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
 Dim stKEY As String

On Error GoTo ErrHand

   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case p_iWay
      Case 0 '第一筆
         strExc(0) = "SELECT nvl(min(IT01||IT02),0) FROM InstType "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 1 '前一筆
         strExc(0) = "SELECT nvl(max(IT01||IT02),0) FROM InstType where IT01||IT02 < " & CNULL(txtField(1) & FdFmt(txtField(2)))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 6
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 2 '後一筆
         strExc(0) = "SELECT nvl(min(IT01||IT02),0) FROM InstType where IT01||IT02>" & CNULL(txtField(1) & FdFmt(txtField(2)))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               DataErrorMessage 7
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 3 '最後筆
         strExc(0) = "SELECT nvl(max(IT01||IT02),0) FROM InstType "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
   End Select
   
   If stKEY <> "" Then
      ReadData Mid(stKEY, 1, 1), Mid(stKEY, 2, 2)
      ShowRecord = True
   End If
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Private Function ReadData(Optional ByVal pKey01 As String, Optional ByVal pKey02 As String) As Boolean
Dim stCon As String
Dim rsAD As New ADODB.Recordset
   If Trim(pKey01) <> "" Then stCon = stCon & "and IT01='" & pKey01 & "' "
   If Trim(pKey02) <> "" Then stCon = stCon & "and IT02='" & FdFmt(pKey02) & "' "

   FormReset

   strExc(0) = "select * from InstType where 1=1 " & stCon & " order by IT01,IT02"
  
   intI = 1
   Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsAD.MoveFirst
      With rsAD
         For Each oText In txtField
            oText.Text = "" & .Fields("IT" & Format(oText.Index, "00"))
            oText.Tag = oText.Text
         Next
      End With
      ReadData = True
   End If
   
   UpdateCUID 1, rsAD
End Function

Private Sub SetGrid(ByVal bolShow As Boolean)
Dim rsD As New ADODB.Recordset
Dim idR As Integer, intR As Integer
    
    'Modified by Lydia 2017/08/01 +C 通訊
    'Modified by Lydia 2020/04/20 +D 業拓; 並且代碼加註在分類部門之前
    'strExc(1) = "select decode(IT01,'A','通用','C','通訊','F','財務','P','專利','T','商標','L','法務',IT01) type,IT01,IT02,IT03 from InstType order by IT01,IT02"
    'Modified by Lydia 2020/05/11 原本「D 業拓=>D 個案」,+U承辦人
    'strExc(1) = "select IT01||decode(IT01,'A','通用','C','通訊','D','業拓','F','財務','P','專利','T','商標','L','法務',IT01) type,IT01,IT02,IT03 from InstType order by IT01,IT02"
    'Modified by Lydia 2020/05/12 改用變數
    'strExc(1) = "select IT01||decode(IT01,'A','通用','C','通訊','D','個案或特殊指示','E','業拓','F','財務','L','法務','P','專利','T','商標','U','承辦人',IT01) type,IT01,IT02,IT03 from InstType order by IT01,IT02"
    'Modified by Lydia 2020/05/13 +使用部門IT10
    'Modified by Lydia 2020/05/14 +基本檔對應欄位IT11,IT12
    strExc(1) = "select IT01||decode(IT01," & GetAddStr(strKind(3)) & ",IT01) type," & _
                     "IT01,IT02,IT03,IT10||decode(IT10,'P','專利','T','商標','') as IT10,IT11,IT12 from InstType order by IT01,IT02"
    intI = 0
    Set rsD = ClsLawReadRstMsg(intI, strExc(1))
    If intI = 1 Then
       Set MGrid1.Recordset = rsD
       'Modified by Lydia 2020/05/13 +使用部門IT10
       'Modified by Lydia 2020/05/14 +基本檔對應欄位IT11,IT12
       MGrid1.FormatString = "分類部門|IT01|分類代號|說明|使用部門|基本檔對應欄位|顯示格式"
       MGrid1.ColWidth(0) = 1100
       MGrid1.ColWidth(1) = 0
       MGrid1.ColWidth(2) = 1100
       MGrid1.ColWidth(3) = 2500
       MGrid1.ColWidth(4) = 1100 'Added by Lydia 2020/05/13 使用部門IT10
       'Added by Lydia 2020/05/14 基本檔對應欄位IT11,IT12
       MGrid1.ColWidth(5) = 1100
       MGrid1.ColWidth(6) = 1000
       'end 2020/05/14
       'Modified by Lydia 2017/06/28 靠左對齊
'       For idR = 4 To MGrid1.Cols - 1
'          MGrid1.ColWidth(idR) = 0
'       Next
       For intR = 0 To rsD.RecordCount
          For idR = 0 To MGrid1.Cols - 1
             MGrid1.row = intR
             MGrid1.col = idR
             MGrid1.CellAlignment = flexAlignLeftCenter
         Next idR
       Next intR
       'end 2017/06/28
       If bolShow = True Then
          SSTab1.Tab = 1
       Else
          SSTab1.Tab = 0
       End If
    End If
         
End Sub
' 更新 Create 及 Update 的人
Private Sub FormReset()
   Dim oText As TextBox
   
   For Each oText In txtField
      oText.Text = ""
      'Text.Tag = "" 'Remove by Lydia 2020/04/20 影響還原上一筆
   Next
   
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   TextInverse txtField(Index)
   If Index <> 3 Then
      CloseIme
   End If
End Sub

Private Sub txtField_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index <> 3 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
Dim iLen As Integer

   Cancel = False
   If m_EditMode = 0 Or m_EditMode = 4 Then Exit Sub
   If txtField(Index) = "" Then Exit Sub 'Added by Lydia 2020/04/20 欄位空白,不檢查輸入值
   
   Select Case Index
        Case 1 '分類部門
           If Trim(txtField(Index)) = "" Then
              MsgBox "分類部門不可空白!", vbCritical, "輸入錯誤"
              GoTo JumpSet
           'Modified by Lydia 2017/08/01 +C 通訊
           'Modified by Lydia 2020/04/20 +D 業拓
           'Modified by Lydia 2020/05/11 原本「D 業拓=>D 個案」,+U承辦人
           'ElseIf txtField(Index) <> "A" And txtField(Index) <> "C" And txtField(Index) <> "D" And txtField(Index) <> "F" And txtField(Index) <> "P" And txtField(Index) <> "T" And txtField(Index) <> "L" Then
           'Modified by Lydia 2020/05/12 改用變數
           'ElseIf InStr("A,C,D,E,F,L,P,T,U", txtField(Index)) = 0 Or Trim(txtField(Index)) = "," Then
           '    MsgBox "請輸入A、C、D、E、F、L、P、T、U!", vbCritical, "輸入錯誤"
           'end 2017/08/01
           ElseIf InStr(strKind(1), txtField(Index)) = 0 Or Trim(txtField(Index)) = "、" Then
               MsgBox "請輸入" & strKind(1) & " !", vbCritical, "輸入錯誤"
           'end 2020/05/12
               GoTo JumpSet
           End If
        Case 2 '分類代號
           If Trim(txtField(Index)) = "" Then
              MsgBox "分類代號不可空白!", vbCritical, "輸入錯誤"
              GoTo JumpSet
           'Modified by Lydia 2017/08/01
           'ElseIf Val(txtField(Index)) < 1 Or Val(txtField(Index)) > 99 Then
           '    MsgBox "請輸入01~99!", vbCritical, "輸入錯誤"
           ElseIf Val(txtField(Index)) < 0 Or Val(txtField(Index)) > 99 Then
               MsgBox "請輸入00~99!", vbCritical, "輸入錯誤"
           'end 2017/08/01
               GoTo JumpSet
           End If
        Case 3 '說明
           If txtField(Index) = "" Then
              MsgBox "請輸入說明!", vbCritical, "輸入錯誤"
              GoTo JumpSet
           Else
              txtField(Index).Text = PUB_StringFilter(txtField(Index).Text)  'Added by Lydia 2020/05/14 清除字串中的enter
           End If
        'Added by Lydia 2020/05/13
        Case 10 '使用部門IT10
           If txtField(Index) <> "" Then
              If txtField(Index) <> "P" And txtField(Index) <> "T" Then
                  MsgBox "請輸入P、T !", vbCritical, "輸入錯誤"
                  GoTo JumpSet
              End If
           End If
        'Added by Lydia 2020/05/14
        Case 11 '基本檔對應欄位IT11
           If txtField(Index) <> "" Then
              If Right(txtField(Index), 1) <> ";" Then
                  MsgBox "欄位名稱請加上;", vbCritical, "輸入錯誤"
                  GoTo JumpSet
              End If
              txtField(Index).Text = PUB_RepToOneSpace(PUB_StringFilter(txtField(Index).Text))  '清除字串中的enter & 清除連續空白
           End If
        Case 12 '基本檔對應欄位-顯示格式IT12
           If txtField(Index) <> "" Then
              If txtField(Index) <> "D" And txtField(Index) <> "N" Then
                  MsgBox "請輸入D、N !", vbCritical, "輸入錯誤"
                  GoTo JumpSet
              End If
           End If
   End Select
   
   If Not CheckLengthIsOK(txtField(Index), iLen) Then
      Cancel = True
   End If
   
   Exit Sub

JumpSet:
   txtField(Index).SetFocus
   txtField_GotFocus (Index)
   Cancel = True
End Sub

Private Function TxtValidate() As Boolean
Dim bCancel As Boolean
'Added by Lydia 2020/05/14
Dim tmpArr As Variant
Dim intP As Integer, strList As String

   'Added by Lydia 2020/04/20 檢查欄位是否空白
   If Trim(txtField(1)) = "" Then
       MsgBox "分類部門不可空白!", vbCritical, "輸入錯誤"
       txtField(1).SetFocus
       txtField_GotFocus 1
       Exit Function
   End If
   If Trim(txtField(2)) = "" Then
       MsgBox "分類代號不可空白!", vbCritical, "輸入錯誤"
       txtField(2).SetFocus
       txtField_GotFocus 2
       Exit Function
   End If
   If Trim(txtField(3)) = "" Then
       MsgBox "說明不可空白!", vbCritical, "輸入錯誤"
       txtField(3).SetFocus
       txtField_GotFocus 3
       Exit Function
   End If
   'end 2020/04/20
   
   txtField(2) = FdFmt(txtField(2))
   For Each oText In txtField
        txtField_Validate oText.Index, bCancel
        If bCancel = True Then
           txtField(oText.Index).SetFocus
           txtField_GotFocus (oText.Index)
           Exit Function
        End If
   Next
   
   'Added by Lydia 2020/05/14 基本檔對應欄位IT11
   If txtField(11).Text <> "" Then
        '檢查是否重複輸入
        tmpArr = Split(txtField(11).Text, ";")
        For intP = 0 To UBound(tmpArr)
            If Trim(tmpArr(intP)) <> "" Then
                If strList = "" Then
                    strList = strList & tmpArr(intP) & ";"
                Else
                    If InStr(strList, tmpArr(intP) & ";") > 0 Then
                         MsgBox "基本檔對應欄位重複輸入" & tmpArr(intP), vbCritical, "輸入錯誤"
                         txtField(11).SetFocus
                         txtField_GotFocus 11
                         Exit Function
                    Else
                        strList = strList & tmpArr(intP) & ";"
                    End If
                End If
                '檢查其他分類
                strExc(0) = "SELECT IT01,IT02,IT03,IT11 FROM INSTTYPE WHERE IT01||IT02<>'" & txtField(1) & txtField(2) & "' AND INSTR(IT11,'" & tmpArr(intP) & ";') > 0 "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                    If MsgBox(RsTemp.Fields("it01") & RsTemp.Fields("it02") & RsTemp.Fields("it03") & "的基本檔對應欄位有" & tmpArr(intP) & vbCrLf & "請確認是否繼續存檔？", vbYesNo + vbExclamation + vbDefaultButton2, "檢查其他分類") = vbNo Then
                        txtField(11).SetFocus
                        txtField_GotFocus 11
                        Exit Function
                    End If
                End If
            End If
        Next intP
   Else
        'Mark: 暫不開放
'        If InStr(txtField(3), "-") > 0 Then
'            MsgBox "說明有 - 為基本檔對應欄位，配合例外輸出的格式。" & vbCrLf & "基本檔對應欄位不可空白! ", vbCritical
'            txtField(11).SetFocus
'            txtField_GotFocus 11
'            Exit Function
'        End If
        '基本檔對應欄位-顯示格式IT12
        If txtField(12).Text <> "" Then
            MsgBox "基本檔對應欄位不可空白! ", vbCritical
            txtField(11).SetFocus
            txtField_GotFocus 11
            Exit Function
        End If
   End If
   'end 2020/05/14
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
       If m_EditMode = 1 Then '新增
          'Modified by Lydia 2020/05/13 +使用部門IT10
          'Modified by Lydia 2020/05/14 +基本檔對應欄位IT11,IT12
          strSql = "insert into InstType(IT01,IT02,IT03,IT04,IT05,IT06,IT10,IT11,IT12)" & _
             " Values ('" & txtField(1).Text & "','" & FdFmt(txtField(2).Text) & "','" & ChgSQL(txtField(3).Text) & "','" & strUserNum & "'," & strSrvDate(1) & "," & Mid(Format(ServerTime, "000000"), 1, 4) & _
             ", " & CNULL(txtField(10)) & ", " & CNULL(txtField(11)) & ", " & CNULL(txtField(12)) & ") "
       Else         '修改
          'Modified by Lydia 2020/05/13 +使用部門IT10
          'Modified by Lydia 2020/05/14 +基本檔對應欄位IT11,IT12
          strSql = "update InstType set IT03='" & ChgSQL(txtField(3).Text) & "', IT07='" & strUserNum & "', IT08=" & strSrvDate(1) & ", IT09=" & Mid(Format(ServerTime, "000000"), 1, 4) & _
                   ", IT10=" & CNULL(txtField(10)) & ", IT11=" & CNULL(txtField(11)) & ", IT12=" & CNULL(txtField(12)) & " where IT01='" & txtField(1).Text & "' and IT02='" & FdFmt(txtField(2).Text) & "' "
       End If
       cnnConnection.Execute strSql, intI

   cnnConnection.CommitTrans
   FormSave = True
   
   Exit Function
   
ErrHnd:
   If Err.Number > 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function

Private Sub MGrid1_DblClick()
   If MGrid1.row > 0 And MGrid1.TextMatrix(MGrid1.row, 0) <> "" Then
      If ReadData(MGrid1.TextMatrix(MGrid1.row, 1), MGrid1.TextMatrix(MGrid1.row, 2)) Then
         SSTab1.Tab = 0
      End If
   End If
End Sub

Private Function FormDelete() As Boolean
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
      strSql = "delete from InstType where IT01='" & txtField(1) & "' and IT02='" & txtField(2) & "' "
      cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   FormDelete = True
   Exit Function
   
ErrHnd:
   If Err.Number > 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function

'判斷是否存在相同分類的記錄
Private Function RecIsExist() As Boolean
   Dim iR As Integer
   Dim rsQa As ADODB.Recordset
   Dim strCon As String
   
If Trim(txtField(1)) <> "" Then
   strCon = strCon & "and IT01='" & txtField(1) & "' "
End If
If Trim(txtField(2)) <> "" Then
   strCon = strCon & "and IT02='" & FdFmt(txtField(2)) & "' "
End If

If Left(strCon, 3) = "and" Then strCon = Mid(strCon, 4, Len(strCon) - 4)

   strExc(1) = " select * from InstType where " & strCon
   iR = 1
   Set rsQa = ClsLawReadRstMsg(iR, strExc(1))
   If iR = 1 Then
      RecIsExist = True
      MsgBox "已存在同樣分類的記錄，請先查詢!!", vbCritical
   Else
      RecIsExist = False
   End If
   Set rsQa = Nothing
   
End Function

Private Function FdFmt(ByVal Str01 As String) As String
   FdFmt = Right("00" & Trim("" & Str01), 2)
End Function

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByVal actType As Integer, Optional ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If actType = 0 Then
      strCName = GetStaffName(strUserNum, True)
      strCDate = Format(strSrvDate(2), "###/##/##")
      strCTime = ""
   Else
        If IsNull(rsSrcTmp.Fields("IT04")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("IT04")) = False Then
              strCName = GetStaffName(rsSrcTmp.Fields("IT04"), True)
           End If
        End If
        If IsNull(rsSrcTmp.Fields("IT05")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("IT05")) = False Then
              strTemp = TAIWANDATE(rsSrcTmp.Fields("IT05"))
              strCDate = Format(strTemp, "###/##/##")
           End If
        End If
        If IsNull(rsSrcTmp.Fields("IT06")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("IT06")) = False Then
              strTemp = rsSrcTmp.Fields("IT06")
              strCTime = Format(strTemp, "00:00")
           End If
        End If
        If IsNull(rsSrcTmp.Fields("IT07")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("IT07")) = False Then
              strUName = GetStaffName(rsSrcTmp.Fields("IT07"), True)
           End If
        End If
        If IsNull(rsSrcTmp.Fields("IT08")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("IT08")) = False Then
              strTemp = TAIWANDATE(rsSrcTmp.Fields("IT08"))
              strUDate = Format(strTemp, "###/##/##")
           End If
        End If
        If IsNull(rsSrcTmp.Fields("IT09")) = False Then
           If IsEmptyText(rsSrcTmp.Fields("IT09")) = False Then
              strTemp = rsSrcTmp.Fields("IT09")
              strUTime = Format(strTemp, "00:00")
           End If
        End If
   End If
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & " " & vbTab & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

