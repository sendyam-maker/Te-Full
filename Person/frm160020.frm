VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160020 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工旅遊補助金資料"
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   324
      Index           =   5
      Left            =   7155
      Style           =   1  '圖片外觀
      TabIndex        =   25
      Top             =   180
      Visible         =   0   'False
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢下一筆(&N)"
      CausesValidation=   0   'False
      Height          =   324
      Index           =   6
      Left            =   5790
      Style           =   1  '圖片外觀
      TabIndex        =   24
      Top             =   180
      Visible         =   0   'False
      Width           =   1320
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5040
      Left            =   30
      TabIndex        =   15
      Top             =   690
      Width           =   8895
      _ExtentX        =   15681
      _ExtentY        =   8890
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm160020.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(1)=   "textSTD05"
      Tab(0).Control(2)=   "textSTD04"
      Tab(0).Control(3)=   "textSTD01"
      Tab(0).Control(4)=   "textSTD02"
      Tab(0).Control(5)=   "textSTD03"
      Tab(0).Control(6)=   "GRD2"
      Tab(0).Control(7)=   "LblPayDate"
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(9)=   "textSTD06"
      Tab(0).Control(10)=   "textSTD01_2"
      Tab(0).Control(11)=   "LblUserFee"
      Tab(0).Control(12)=   "LblFee"
      Tab(0).Control(13)=   "Line1"
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(15)=   "Label7"
      Tab(0).Control(16)=   "Label1(0)"
      Tab(0).Control(17)=   "Label3"
      Tab(0).Control(18)=   "Label1(17)"
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160020.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txt1(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txt1(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "GRD1"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt1(5)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txt1(4)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdOK(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin VB.Frame Frame1 
         Caption         =   "可申請近二年補助資料："
         ForeColor       =   &H00FF0000&
         Height          =   552
         Left            =   3096
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   3792
         Begin VB.Label LblUserFee2 
            Caption         =   "LblUserFee2"
            ForeColor       =   &H000000C0&
            Height          =   228
            Left            =   324
            TabIndex        =   35
            Top             =   252
            Width           =   2772
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "更新餘額"
         Height          =   405
         Left            =   -67470
         TabIndex        =   30
         Top             =   1830
         Width           =   1035
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "查詢"
         Height          =   324
         Index           =   0
         Left            =   6930
         TabIndex        =   29
         Top             =   324
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   4
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   11
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   5
         Left            =   2130
         MaxLength       =   7
         TabIndex        =   12
         Top             =   660
         Width           =   915
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   3690
         TabIndex        =   26
         Top             =   360
         Width           =   1935
         Begin VB.TextBox txt1 
            Height          =   270
            Index           =   2
            Left            =   570
            MaxLength       =   3
            TabIndex        =   9
            Top             =   0
            Width           =   495
         End
         Begin VB.TextBox txt1 
            Height          =   270
            Index           =   3
            Left            =   1200
            MaxLength       =   3
            TabIndex        =   10
            Top             =   0
            Width           =   495
         End
         Begin VB.Line Line5 
            X1              =   900
            X2              =   1500
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "部門："
            Height          =   180
            Left            =   0
            TabIndex        =   27
            Top             =   30
            Width           =   540
         End
      End
      Begin VB.TextBox textSTD05 
         Height          =   270
         Left            =   -70200
         MaxLength       =   7
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox textSTD04 
         Height          =   270
         Left            =   -71370
         MaxLength       =   7
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox textSTD01 
         Height          =   270
         Left            =   -73620
         MaxLength       =   6
         TabIndex        =   0
         Top             =   420
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160020.frx":0038
         Height          =   4035
         Left            =   30
         TabIndex        =   13
         Top             =   960
         Width           =   8790
         _ExtentX        =   15522
         _ExtentY        =   7108
         _Version        =   393216
         Cols            =   9
         FixedCols       =   2
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "員工編號|姓名|申請日|申請金額|補助年度|補助額度|補助金額|旅遊期間|備註"
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
         _Band(0).Cols   =   9
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   8
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   7
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox textSTD02 
         Height          =   270
         Left            =   -73620
         MaxLength       =   7
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox textSTD03 
         Height          =   270
         Left            =   -73620
         TabIndex        =   4
         Top             =   1050
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD2 
         Bindings        =   "frm160020.frx":004D
         Height          =   2445
         Left            =   -74940
         TabIndex        =   6
         Top             =   2310
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   4322
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "申請日期|補助年度|年資|補助額度|補助金額|旅遊期間|備註|剩餘額度"
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
         _Band(0).Cols   =   8
      End
      Begin MSForms.Label LblPayDate 
         Height          =   225
         Left            =   -68340
         TabIndex        =   33
         Top             =   450
         Width           =   2115
         ForeColor       =   255
         BackColor       =   12632256
         VariousPropertyBits=   27
         Caption         =   "付款日期："
         Size            =   "3731;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label23 
         Height          =   225
         Left            =   -74670
         TabIndex        =   32
         Top             =   4770
         Width           =   7395
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13044;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSTD06 
         Height          =   630
         Left            =   -73620
         TabIndex        =   5
         Top             =   1380
         Width           =   5895
         VariousPropertyBits=   -1466939365
         MaxLength       =   200
         ScrollBars      =   3
         Size            =   "10398;1111"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label textSTD01_2 
         Height          =   225
         Left            =   -72840
         TabIndex        =   31
         Top             =   450
         Width           =   1395
         BackColor       =   12632256
         VariousPropertyBits=   27
         Size            =   "2461;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line2 
         X1              =   1770
         X2              =   2460
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請日期："
         Height          =   180
         Left            =   150
         TabIndex        =   28
         Top             =   690
         Width           =   900
      End
      Begin VB.Label LblUserFee 
         Caption         =   "LblUserFee"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   -72870
         TabIndex        =   23
         Top             =   2100
         Width           =   2775
      End
      Begin VB.Label LblFee 
         AutoSize        =   -1  'True
         Caption         =   "可申請近二年補助資料："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   -74880
         TabIndex        =   22
         Top             =   2100
         Width           =   1980
      End
      Begin VB.Line Line1 
         X1              =   -70620
         X2              =   -69930
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "旅遊日期："
         Height          =   180
         Left            =   -72270
         TabIndex        =   21
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "備　　註："
         Height          =   180
         Left            =   -74520
         TabIndex        =   20
         Top             =   1410
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         Height          =   180
         Index           =   0
         Left            =   -74520
         TabIndex        =   19
         Top             =   465
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   150
         TabIndex        =   18
         Top             =   390
         Width           =   900
      End
      Begin VB.Line Line4 
         X1              =   1770
         X2              =   2460
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "申請日期："
         Height          =   180
         Left            =   -74520
         TabIndex        =   17
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請費用："
         Height          =   180
         Index           =   17
         Left            =   -74520
         TabIndex        =   16
         Top             =   1080
         Width           =   900
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
            Picture         =   "frm160020.frx":0062
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160020.frx":037E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160020.frx":069A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160020.frx":0876
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160020.frx":0B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160020.frx":0EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160020.frx":11CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160020.frx":14E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160020.frx":1802
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160020.frx":1B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160020.frx":1E3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   14
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
Attribute VB_Name = "frm160020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/13 Form2.0已修改
'Create by Sindy 2019/7/31
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
' 第一筆資料的本所案號
Dim m_FirstKEY(3) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(3) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(3) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_STD As Integer
Public UpForm As Form 'Add By Sindy 2019/9/9
Dim strFeeYear1 As String, strFeeYear2 As String
Dim strFeeYear0 As String 'Add By Sindy 2020/12/1


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      'Add By Sindy 2012/6/20
      Case 5 '結束
         Unload Me
         UpForm.Show
      Case 6 '查詢下一筆
         Unload Me
         UpForm.cmdState = 6
         UpForm.PubShowNextData
      '2012/6/20 End
      Case 0 '查詢
         'Modify By Sindy 2019/10/23 + & txt1(4) & txt1(5)
         If txt1(0) & txt1(1) & txt1(2) & txt1(3) & txt1(4) & txt1(5) <> "" Then
            If RunNick(txt1(0), txt1(1)) Then
                txt1(0).SetFocus
                Exit Sub
            End If
            If RunNick2(txt1(2), txt1(3)) Then
                txt1(2).SetFocus
                Exit Sub
            End If
            GetData
            'Add By Sindy 2023/7/25
            If Me.Frame1.Visible = True Then
               If txt1(0) = txt1(1) Then
                  Call frm160020.GetFeeMoney(txt1(0))
               Else
                  LblUserFee2.Caption = "僅供查詢個人資料"
               End If
            End If
            '2023/7/25 END
         Else
            MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
         End If
   End Select
End Sub

'Add By Sindy 2021/3/10
Private Sub Command1_Click()
Dim ii As Integer
Dim strYear As String
Dim strOverFee As String
Dim dblFee As Double
   
   If m_EditMode = 0 Then
      strYear = Trim(InputBox("請輸入欲更新的【補助年度】？" & vbCrLf & vbCrLf & "(註:沒輸入代表取消)", "更新餘額"))
      If strYear <> "" Then
         If Len(strYear) <> 4 Then
            strYear = Val(strYear) + 1911
         End If
         strOverFee = Trim(InputBox("請輸入欲更新的【餘額】？" & vbCrLf & vbCrLf & "(註:沒輸入代表取消)", "更新餘額"))
         If Trim(strOverFee) <> "" Then
            strSql = "update staff_travelFee set STF13=" & Val(strOverFee) & _
                     " where STF01='" & textSTD01 & "'" & _
                     " and STF03='" & strYear & "' and (STF13>0 or STF13 is null)"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
            Call GetTravelDetail(textSTD01)
            MsgBox "更新成功！"
         End If
      End If
      
   ElseIf m_EditMode = 2 Then '修改時
      dblFee = 0: strYear = ""
      For ii = 1 To GRD2.Rows - 1
         If Trim(GRD2.TextMatrix(ii, 0)) <> "" Then
            '檢查是否有同補助年度餘額
            If strYear = "" Or Val(strYear) <> Val(GRD2.TextMatrix(ii, 1)) Then
               dblFee = CDbl(GRD2.TextMatrix(ii, 3)) - CDbl(GRD2.TextMatrix(ii, 4))
               GRD2.TextMatrix(ii, 7) = Format(dblFee, "#,##0")
            Else
               If dblFee > 0 And Val(strYear) = Val(GRD2.TextMatrix(ii, 1)) Then
                  dblFee = dblFee - CDbl(GRD2.TextMatrix(ii, 4))
                  GRD2.TextMatrix(ii - 1, 7) = 0
                  GRD2.TextMatrix(ii, 7) = Format(dblFee, "#,##0")
               Else
                  dblFee = CDbl(GRD2.TextMatrix(ii, 3)) - CDbl(GRD2.TextMatrix(ii, 4))
                  GRD2.TextMatrix(ii, 7) = Format(dblFee, "#,##0")
               End If
            End If

'            If j > 1 And Val(GRD2.TextMatrix(j - 1, 1)) + 1911 = Val(GRD2.TextMatrix(j, 1)) + 1911 And _
'               Val(GRD2.TextMatrix(j - 1, 7)) > 0 Then
'               strOverFee = Val(Format(GRD2.TextMatrix(j - 1, 7), "##0"))
'            Else
'               strOverFee = Val(Format(GRD2.TextMatrix(j, 3), "##0"))
'            End If
'            strOverFee = strOverFee - Val(Format(GRD2.TextMatrix(j, 4), "##0"))
'            GRD2.TextMatrix(j, 7) = Format(strOverFee, "#,##0")
            '餘額已為0時,此年度的餘額應該都為0了
            If CDbl(GRD2.TextMatrix(ii, 7)) = 0 Then
               strSql = "update staff_travelFee set STF13=0" & _
                        " where STF01='" & textSTD01 & "'" & _
                        " and STF03='" & Val(GRD2.TextMatrix(ii, 1)) + 1911 & "'" & _
                        " and nvl(STF13,0)<>" & CDbl(GRD2.TextMatrix(ii, 7))
               cnnConnection.Execute strSql, intI
               'CDbl(grd2.TextMatrix(ii, 7))
            Else
               strSql = "update staff_travelFee set STF13=" & CDbl(GRD2.TextMatrix(ii, 7)) & _
                        " where STF01='" & textSTD01 & "'" & _
                        " and STF02=" & DBDATE(GRD2.TextMatrix(ii, 0)) & _
                        " and STF03='" & Val(GRD2.TextMatrix(ii, 1)) + 1911 & "'" & _
                        " and nvl(STF13,0)<>" & CDbl(GRD2.TextMatrix(ii, 7))
               cnnConnection.Execute strSql, intI
            End If
         End If
         strYear = GRD2.TextMatrix(ii, 1)
      Next ii
      Call GetTravelDetail(textSTD01)
      'MsgBox "更新成功！"
   End If
End Sub

Private Sub Form_Initialize()
Set rsA = New ADODB.Recordset
If rsA.State = 1 Then rsA.Close
rsA.CursorLocation = adUseClient
rsA.Open "select * from staff_TravelData where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
tf_STD = rsA.Fields.Count
SetGrd
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
'Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub

Private Sub Form_Load()
   ReDim m_FieldList(tf_STD) As FIELDITEM
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   LblUserFee.Caption = ""
   LblUserFee2.Caption = "" 'Add By Sindy 2023/7/25
   
   InitialField
   InitialData
   RefreshRange
   ShowLastRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
   
   '多看一年的申請資料
   strFeeYear0 = Val(Left(strSrvDate(1), 4)) - 2
   '可申請近二年補助資料
   strFeeYear1 = Val(Left(strSrvDate(1), 4)) - 1
   strFeeYear2 = Left(strSrvDate(1), 4)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160020 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow GRD1, X, Y, nCol, nRow
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j

   GRD1.Visible = False
   tmpMouseRow = GRD1.row
   GRD1.Visible = True
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
           
         textSTD01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         textSTD02.Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 2))
         textSTD03.Text = Val(GRD1.TextMatrix(tmpMouseRow, 4)) + 1911
         If textSTD01.Text = "" Then
            textSTD01.Text = GRD1.TextMatrix(tmpMouseRow, 10)
            textSTD02.Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 11))
         End If
         QueryRecord
         GRD1.Visible = True
      End If
   End If
End Sub

Private Sub grd2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow GRD2, X, Y, nCol, nRow
   GRD2.col = nCol
   GRD2.row = nRow
End Sub

Private Sub GRD2_SelChange()
Dim tmpMouseRow
Dim i, j

   GRD2.Visible = False
   tmpMouseRow = GRD2.row
   GRD2.Visible = True
   If tmpMouseRow <> 0 Then
      GRD2.row = tmpMouseRow
      GRD2.col = 0
      If GRD2.CellBackColor <> &HFFC0C0 Then
         GRD2.Visible = False
         For j = 1 To GRD2.Rows - 1
            GRD2.row = j
            For i = 0 To GRD2.Cols - 1
               GRD2.col = i
               GRD2.CellBackColor = QBColor(15)
            Next i
         Next j
         'Add By Sindy 2021/3/11
         If m_EditMode = 0 Then
         '2021/3/11 END
            If textSTD01 <> "" And Val(GRD2.TextMatrix(tmpMouseRow, 0)) > 0 Then
               textSTD01.Text = textSTD01 '員工編號
               textSTD02.Text = ChangeTDateStringToTString(GRD2.TextMatrix(tmpMouseRow, 0)) '申請日期
               Call QueryRecord(False)
            End If
         End If
         GRD2.row = tmpMouseRow
         For i = 0 To GRD2.Cols - 1
            GRD2.col = i
            GRD2.CellBackColor = &HFFC0C0
         Next i
         GRD2.Visible = True
      End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If Me.SSTab1.TabVisible(0) = True Then
      If PreviousTab = 0 Then
         cmdOK(0).SetFocus
         cmdOK(0).Default = True
      Else
         cmdOK(0).Default = False
      End If
   End If
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

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String

   If IsNull(rsSrcTmp.Fields("STD07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STD07")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("STD07"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STD08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STD08")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("STD08"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STD09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STD09")) = False Then
         strTemp = rsSrcTmp.Fields("STD09")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STD10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STD10")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("STD10"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STD11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STD11")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("STD11"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STD12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STD12")) = False Then
         strTemp = rsSrcTmp.Fields("STD12")
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

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If m_EditMode = 1 Then
      If CDbl(LblUserFee.Tag) = 0 Then
         MsgBox "無補助金可申請！", vbExclamation
         Exit Function
      Else
         '費用不可大於補助金
         If CDbl(Val(textSTD03.Text)) > CDbl(LblUserFee.Tag) Then
            MsgBox "費用(" & CDbl(Val(textSTD03.Text)) & ")不可大於補助金(" & CDbl(LblUserFee.Tag) & ")", vbExclamation
            Exit Function
         End If
      End If
   End If

   If textSTD01.Text = "" Then
       MsgBox "員工代號不可以空白！", vbExclamation
       textSTD01.SetFocus
       Exit Function
   End If
   '增加判斷員工代號+日期是否人員已離職
   If ChkStaffST04(textSTD01, True, textSTD02) = True Then
      textSTD01.SetFocus
      Exit Function
   End If
   
   If textSTD02.Text = "" Then
      MsgBox "申請日不可以空白！", vbExclamation
      textSTD02.SetFocus
      Exit Function
   End If
   
   'Add By Sindy 2021/3/10
   If m_EditMode = 1 Then
      strExc(0) = "SELECT count(*)" & _
                  " FROM staff_TravelData" & _
                  " WHERE STD01='" & textSTD01 & "' and substr(STD02,1,4)=" & Left(DBDATE(textSTD02), 4)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If RsTemp.Fields(0) >= 2 Then
         If MsgBox("今年已超過２次旅遊補助申請，是否要繼續？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
            Exit Function
         End If
      End If
   End If
   '2021/3/10 END
                        
   If CDbl(Val(textSTD03.Text)) = 0 Then
      MsgBox "申請費用不可以空白！", vbExclamation
      textSTD03.SetFocus
      Exit Function
   End If
   If textSTD04.Text = "" Then
      MsgBox "旅遊起始日不可以空白！", vbExclamation
      textSTD04.SetFocus
      Exit Function
   End If
   If textSTD05.Text = "" Then
      MsgBox "旅遊截止日不可以空白！", vbExclamation
      textSTD05.SetFocus
      Exit Function
   End If
   
   If Val(textSTD04) > Val(strSrvDate(2)) Then
      If MsgBox("旅遊起始日期大於系統日，是否要繼續？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
         textSTD04.SetFocus
         Exit Function
      End If
   End If
   If Val(textSTD05) > Val(strSrvDate(2)) Then
      If MsgBox("旅遊截止日期大於系統日，是否要繼續？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
         textSTD05.SetFocus
         Exit Function
      End If
   End If
   
   'Add by Sindy 2021/9/1 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/9/1 END
   
   TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To tf_STD - 1
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
   
   For nIndex = 0 To tf_STD - 1
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
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strSTD01 As String
Dim strSTD02 As String
Dim strSTF03 As String
Dim dblFee As Double
Dim ii As Integer
Dim strUpdSTF05 As String
Dim bolConn As Boolean
Dim bFirst As Boolean
Dim bDifference As Boolean
Dim nIndex As Integer
Dim strTmp As String
Dim strUpdSTF13 As String 'Add By Sindy 2021/3/11

   AddRecord = False

   strSTD01 = textSTD01
   strSTD02 = DBDATE(textSTD02)
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strSTD01, strSTD02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO staff_TravelData("
   For nIndex = 0 To tf_STD - 1
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
   For nIndex = 0 To tf_STD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
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
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans: bolConn = True
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   dblFee = CDbl(textSTD03.Text): strSTF03 = ""
   For ii = 1 To GRD2.Rows - 1
      'Modify By Sindy 2021/3/12 只檢查可申請的近二年補助資料
      If Val(GRD2.TextMatrix(ii, 1)) + 1911 > strFeeYear0 Then
      '2021/3/12 END
         If CDbl(GRD2.TextMatrix(ii, 7)) > 0 Then '剩餘額度
            strSTF03 = Val(GRD2.TextMatrix(ii, 1)) + 1911 '補助年度
            If dblFee > CDbl(GRD2.TextMatrix(ii, 7)) Then
               dblFee = dblFee - CDbl(GRD2.TextMatrix(ii, 7))
               strUpdSTF05 = CDbl(GRD2.TextMatrix(ii, 7))
               strUpdSTF13 = 0 'Add By Sindy 2021/3/11
            Else
               strUpdSTF05 = dblFee: dblFee = 0
               strUpdSTF13 = CDbl(GRD2.TextMatrix(ii, 7)) - CDbl(strUpdSTF05) 'Add By Sindy 2021/3/11
            End If
            
            'Add By Sindy 2021/3/11
            '當此年度餘額已為0,該年的其他筆資料餘額也應歸0了
            'Modify By Sindy 2021/12/28 下列會重新紀錄此年度餘額,此年度舊資料餘額歸0
'            If strUpdSTF13 = 0 Then
               strSql = "UPDATE staff_travelFee SET STF13=0" & _
                        " WHERE STF01=" & CNULL(textSTD01) & _
                        " AND STF03=" & CNULL(strSTF03)
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql, intI
'            End If
            '2021/3/11 END
            
            'Add By Sindy 2021/3/11 + STF13
            strSql = "INSERT INTO staff_travelFee(STF01,STF02,STF03,STF04,STF05,STF06,STF13)" & _
                     " VALUES (" & CNULL(textSTD01) & "," & DBDATE(textSTD02) & "," & CNULL(strSTF03) & "," & CDbl(GRD2.TextMatrix(ii, 3)) & _
                     "," & strUpdSTF05 & "," & GRD2.TextMatrix(ii, 2) & "," & strUpdSTF13 & ")"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
            
            If dblFee = 0 Then Exit For
         End If
      End If
   Next ii
   
   cnnConnection.CommitTrans: bolConn = False
   
   If ((strSTD01 & strSTD02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or _
      ((strSTD01 & strSTD02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   
   LblFee.Tag = ""
   ShowCurrRecord strSTD01, DBDATE(strSTD02)
   
   AddRecord = True
   Exit Function
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

' 修改記錄
Private Function ModRecord() As Boolean
   Dim strSTD01 As String
   Dim strSTD02 As String
   Dim strSTF03 As String
   Dim strConSql As String
   Dim ii As Integer
   Dim strUpdSTF05 As Double
   Dim dblFee As Double
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim nIndex As Integer
   Dim strTmp As String
   Dim bolConn As Boolean
   Dim dblCompFee As Double
   Dim strUpdSTF13 As String 'Add By Sindy 2021/3/11
   
   ModRecord = False

   strSTD01 = m_CurrKEY(0)
   strSTD02 = DBDATE(m_CurrKEY(1))
   
   If CDbl(Val(textSTD03.Tag)) <> CDbl(Val(textSTD03.Text)) Then
      '費用不可大於補助金
      If CDbl(Val(textSTD03.Text)) > (CDbl(textSTD03.Tag) + CDbl(LblUserFee.Tag)) Then
         MsgBox "費用(" & CDbl(Val(textSTD03.Text)) & ")不可大於補助金(" & (CDbl(textSTD03.Tag) + CDbl(LblUserFee.Tag)) & ")", vbExclamation
         textSTD03.SetFocus
         Exit Function
      End If
   End If
   
   strSql = "begin user_data.user_enabled:=1; UPDATE staff_TravelData SET "
   
   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_STD - 1
      strTmp = Empty
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
   Next nIndex
   
   strSql = strSql & " " & _
                  "WHERE STD01 = '" & strSTD01 & "' and STD02='" & strSTD02 & "'; end; "
On Error GoTo ErrHand
   cnnConnection.BeginTrans: bolConn = True
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      
      If CDbl(textSTD03.Tag) <> CDbl(textSTD03.Text) Then
         dblFee = CDbl(textSTD03.Text): strSTF03 = ""
         For ii = 1 To GRD2.Rows - 1
            'Modify By Sindy 2021/3/12 只檢查可申請的近二年補助資料
            If Val(GRD2.TextMatrix(ii, 1)) + 1911 > strFeeYear0 Then
            '2021/3/12 END
               '申請日要同一張
               If DBDATE(GRD2.TextMatrix(ii, 0)) = strSTD02 Or _
                  Trim(DBDATE(GRD2.TextMatrix(ii, 0))) = "" Then
                  
                  strSTF03 = Val(GRD2.TextMatrix(ii, 1)) + 1911
                  If Trim(GRD2.TextMatrix(ii, 0)) = "" Then
                     dblCompFee = CDbl(GRD2.TextMatrix(ii, 3)) '補助額度
                  Else
                     dblCompFee = CDbl(GRD2.TextMatrix(ii, 4)) '補助金額
                  End If
                  If dblFee = 0 And DBDATE(GRD2.TextMatrix(ii, 0)) = strSTD02 Then
                     '申請金額變少了的狀況下
                     strSql = "DELETE FROM staff_TravelFee" & _
                                 " WHERE STF01=" & CNULL(strSTD01) & _
                                       " and STF02=" & strSTD02 & _
                                       " and STF03=" & CNULL(strSTF03)
                     Pub_SeekTbLog strSql
                     cnnConnection.Execute strSql
                     
                  ElseIf dblFee <> dblCompFee Or _
                     Trim(DBDATE(GRD2.TextMatrix(ii, 0))) = "" Then
                     
                     If dblFee > dblCompFee Then
                        If dblFee > CDbl(GRD2.TextMatrix(ii, 3)) Then
                           dblFee = dblFee - CDbl(GRD2.TextMatrix(ii, 3))
                           strUpdSTF05 = CDbl(GRD2.TextMatrix(ii, 3))
                           strUpdSTF13 = 0 'Add By Sindy 2021/3/11
                        Else
                           strUpdSTF05 = dblFee: dblFee = 0
                           strUpdSTF13 = CDbl(GRD2.TextMatrix(ii, 3)) - strUpdSTF05 'Add By Sindy 2021/3/11
                        End If
                     Else
                        strUpdSTF05 = dblFee: dblFee = 0
                        strUpdSTF13 = CDbl(GRD2.TextMatrix(ii, 3)) - strUpdSTF05 'Add By Sindy 2021/3/11
                     End If
                     If Trim(GRD2.TextMatrix(ii, 0)) = "" Then
                        'Modify By Sindy 2021/3/11 + strUpdSTF13
                        strSql = "INSERT INTO staff_travelFee(STF01,STF02,STF03,STF04,STF05,STF06,STF13)" & _
                                 " VALUES (" & CNULL(textSTD01) & "," & DBDATE(textSTD02) & "," & CNULL(strSTF03) & "," & CDbl(GRD2.TextMatrix(ii, 3)) & _
                                 "," & strUpdSTF05 & "," & GRD2.TextMatrix(ii, 2) & "," & strUpdSTF13 & ")"
                     Else
                        'Modify By Sindy 2021/3/11 + strUpdSTF13
                        strSql = "UPDATE staff_TravelFee SET STF05=" & strUpdSTF05 & ",STF13=" & strUpdSTF13 & _
                                 " WHERE STF01=" & CNULL(strSTD01) & _
                                       " and STF02=" & strSTD02 & _
                                       " and STF03=" & CNULL(strSTF03)
                     End If
                     Pub_SeekTbLog strSql
                     cnnConnection.Execute strSql
                  ElseIf dblFee = dblCompFee Then
                     dblFee = 0
                  End If
                  'If dblFee = 0 Then Exit For
               End If
            End If
         Next ii
      End If
   End If
        
   cnnConnection.CommitTrans: bolConn = False
   
   LblFee.Tag = ""
   ShowCurrRecord strSTD01, DBDATE(strSTD02)
   ModRecord = True
   Exit Function
   
ErrHand:
   If bolConn = True Then cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strSTD01 As String
Dim strSTD02 As String
Dim rsTmp As New ADODB.Recordset
Dim dblSTF13 As Double 'Add By Sindy 2021/12/28
   
   DelRecord = False
   
   strSTD01 = m_CurrKEY(0)
   strSTD02 = DBDATE(m_CurrKEY(1))
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   'Add By Sindy 2021/10/22 補回餘額
   strSql = "select * FROM staff_TravelFee WHERE STF01=" & CNULL(strSTD01) & " and STF02=" & strSTD02 & " order by STF01 asc,STF02 asc,STF03 asc"
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         'Add By Sindy 2021/12/28 重新計算該補助年度的餘額
         strExc(0) = "SELECT stf04-sum(stf05) FROM staff_TravelFee" & _
                  " WHERE STF01=" & CNULL(strSTD01) & " and STF03=" & rsTmp.Fields("stf03") & _
                  " and STF02 in (select STF02 from staff_TravelFee WHERE STF01=" & CNULL(strSTD01) & " and STF03=" & rsTmp.Fields("stf03") & " and STF02<>'" & strSTD02 & "')" & _
                  " group by stf04"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If RsTemp.RecordCount > 0 Then
            dblSTF13 = RsTemp.Fields(0)
            If dblSTF13 <> rsTmp.Fields("stf05") Then
               'MsgBox "補回餘額有問題，請洽電腦中心。"
               PUB_SendMail strUserNum, "97038", "", strSTD01 & "補回餘額有問題(" & rsTmp.Fields("stf05") & "元)計算是(" & dblSTF13 & "元)，請洽電腦中心。", _
                           "重新計算該補助年度的餘額=" & vbCrLf & vbCrLf & strExc(0)
               'GoTo ErrHand
            End If
         End If
         '2021/12/28 END
         'rsTmp.Fields("stf05") => dblSTF13
         strSql = "UPDATE staff_TravelFee SET stf13=" & dblSTF13 & _
                  " WHERE STF01=" & CNULL(strSTD01) & " and STF03=" & rsTmp.Fields("stf03") & _
                  " and STF02=(select nvl(max(STF02),0) from staff_TravelFee WHERE STF01=" & CNULL(strSTD01) & " and STF03=" & rsTmp.Fields("stf03") & " and STF02<>'" & strSTD02 & "')"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql, intI
         If intI > 1 Then
            MsgBox "補回餘額有問題，請洽電腦中心。"
            PUB_SendMail strUserNum, "97038", "", strSTD01 & "補回餘額有問題，請洽電腦中心。", _
                        "補回餘額有問題，請洽電腦中心。" & vbCrLf & vbCrLf & strSql
            GoTo ErrHand
         End If
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   '2021/10/22 END
   
   strSql = "DELETE FROM staff_TravelFee " & _
            "WHERE STF01=" & CNULL(strSTD01) & " and STF02=" & strSTD02
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   strSql = "DELETE FROM staff_TravelData " & _
            "WHERE STD01=" & CNULL(strSTD01) & " and STD02=" & strSTD02
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
      
   If (strSTD01 = m_LastKEY(0) And strSTD02 = m_LastKEY(1)) Or _
      (strSTD01 = m_FirstKEY(0) And strSTD02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   
   LblFee.Tag = ""
   ShowCurrRecord strSTD01, DBDATE(strSTD02)
   
   DelRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord(Optional bolQueryFee As Boolean = True) As Boolean
Dim strSTD01 As String
Dim strSTD02 As String

   QueryRecord = False
   
   strSTD01 = textSTD01
   strSTD02 = DBDATE(textSTD02)
   If IsRecordExist(strSTD01, strSTD02) = True Then
      m_CurrKEY(0) = strSTD01
      m_CurrKEY(1) = strSTD02
      QueryRecord = True
      Call UpdateCtrlData(bolQueryFee)
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
         textSTD03 = ""
         'Add By Sindy 2022/4/11
         If textSTD01.Text = "" Then
            MsgBox "員工代號不可以空白！", vbExclamation
            textSTD01.SetFocus
            Exit Function
         End If
         If textSTD02 = "" Then
            MsgBox "申請日期不可以空白！", vbExclamation
            textSTD02.SetFocus
            Exit Function
         End If
         '2022/4/11 END
         If textSTD01 <> "" And textSTD02 <> "" Then
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
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: If Me.Visible = True Then textSTD01.SetFocus
      Case 2: If Me.Visible = True Then textSTD05.SetFocus
      Case 4: If Me.Visible = True Then textSTD01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   If strKEY01 = "" Or strKEY02 = "" Then Exit Function
   
   strSql = "SELECT * FROM staff_TravelData " & _
            "WHERE STD01='" & strKEY01 & "' and STD02=" & strKEY02
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

   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "SELECT STD01,STD02,STD03 FROM staff_TravelData " & _
               "WHERE STD01 = '" & m_CurrKEY(0) & "' and STD02=" & m_CurrKEY(1)
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("STD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STD01")
         If IsNull(rsTmp.Fields("STD02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STD02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT STD01,STD02,STD03 FROM staff_TravelData " & _
               "WHERE STD02=(SELECT MIN(STD02) FROM staff_TravelData where STD01=(select min(STD01) from staff_TravelData))" & _
               " and STD01=(SELECT MIN(STD01) FROM staff_TravelData) Order BY STD03 ASC"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("STD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STD01")
         If IsNull(rsTmp.Fields("STD02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STD02")
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
   
   If m_CurrKEY(0) = m_FirstKEY(0) And _
      m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
      
'   strSql = "SELECT STD01,STD02,STD03 FROM staff_TravelData " & _
'            "WHERE STD01 = '" & m_CurrKEY(0) & "' AND " & _
'                  "STD02 = (SELECT MAX(STD02) FROM staff_TravelData " & _
'                          "WHERE STD01 = '" & m_CurrKEY(0) & "' AND " & _
'                                "STD02 < '" & m_CurrKEY(1) & "' ) Order By STD03 DESC "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      If IsNull(rsTmp.Fields("STD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STD01")
'      If IsNull(rsTmp.Fields("STD02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STD02")
'      rsTmp.Close
'      UpdateCtrlData
'      GoTo EXITSUB
'   End If
'   rsTmp.Close
   
   strSql = "SELECT STD01,STD02,STD03 FROM staff_TravelData " & _
            "WHERE STD01 = (SELECT MAX(STD01) FROM staff_TravelData " & _
                           "WHERE STD01 < '" & m_CurrKEY(0) & "') AND " & _
                  "STD02 = (SELECT MAX(STD02) FROM staff_TravelData " & _
                           "WHERE STD01 = (SELECT MAX(STD01) FROM staff_TravelData " & _
                                          "WHERE STD01 < '" & m_CurrKEY(0) & "')) Order BY STD03 DESC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("STD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STD01")
      If IsNull(rsTmp.Fields("STD02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STD02")
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And _
      m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
'   strSql = "SELECT STD01,STD02,STD03 FROM staff_TravelData " & _
'            "WHERE STD01 = '" & m_CurrKEY(0) & "' AND " & _
'                  "STD02 = (SELECT MIN(STD02) FROM staff_TravelData " & _
'                          "WHERE STD01 = '" & m_CurrKEY(0) & "' AND " & _
'                                "STD02 > '" & m_CurrKEY(1) & "' ) Order by STD03 ASC "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      If IsNull(rsTmp.Fields("STD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STD01")
'      If IsNull(rsTmp.Fields("STD02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STD02")
'      rsTmp.Close
'      UpdateCtrlData
'      GoTo EXITSUB
'   End If
'   rsTmp.Close
   
   strSql = "SELECT STD01,STD02,STD03 FROM staff_TravelData " & _
            "WHERE STD01 = (SELECT MIN(STD01) FROM staff_TravelData " & _
                           "WHERE STD01 > '" & m_CurrKEY(0) & "') AND " & _
                  "STD02 = (SELECT max(STD02) FROM staff_TravelData " & _
                           "WHERE STD01 = (SELECT MIN(STD01) FROM staff_TravelData " & _
                                          "WHERE STD01 > '" & m_CurrKEY(0) & "')) Order BY STD03 ASC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("STD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STD01")
      If IsNull(rsTmp.Fields("STD02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("STD02")
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
   
   Command1.Visible = False 'Add By Sindy 2021/3/11 更新餘額按鈕
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         '檢查今年的補助金額是否已輸入
         strExc(0) = "select * from staff_travelMoney where stm01=" & Left(strSrvDate(1), 4)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            strTit = "詢問"
            strMsg = "今年的補助金額尚未輸入，是否要繼續？"
            nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
            If nResponse = vbNo Then Exit Sub
         End If
         m_EditMode = 1
         ClearField
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
         Call GetFeeMoney(m_CurrKEY(0)) '可申請近二年補助資料
      ' 修改
      Case vbKeyF3:
         'Add By Sindy 2021/12/28 付款日期
         If LblPayDate.Visible = True Then
            MsgBox "已付款，不可修改資料！", vbInformation
            Exit Sub
         End If
         '2021/12/28 END
         m_EditMode = 2
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
         Call GetFeeMoney(m_CurrKEY(0)) '可申請近二年補助資料
      ' 刪除
      Case vbKeyF5:
         'Add By Sindy 2021/12/28 付款日期
         If LblPayDate.Visible = True Then
            MsgBox "已付款，不可刪除資料！", vbInformation
            Exit Sub
         End If
         '2021/12/28 END
         'Add By Sindy 2022/1/5
         If SSTab1.Tab <> 0 Then
            MsgBox "請先切換到明細資料，查看後才能刪除！", vbInformation
            Exit Sub
         End If
         '2022/1/5 END
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
         UpdateFieldNewData
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
   End If
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT STD01,STD02,STD03 FROM staff_TravelData " & _
            "WHERE STD01 = (SELECT MIN(STD01) FROM staff_TravelData) AND " & _
                  "STD02 = (SELECT max(STD02) FROM staff_TravelData " & _
                           "WHERE STD01 = (SELECT MIN(STD01) FROM staff_TravelData)) Order BY STD03 ASC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("STD01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("STD01")
      If IsNull(rsTmp.Fields("STD02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("STD02")
   End If
   rsTmp.Close
   
   strSql = "SELECT STD01,STD02,STD03 FROM staff_TravelData " & _
            "WHERE STD01 = (SELECT MAX(STD01) FROM staff_TravelData) AND " & _
                  "STD02 = (SELECT MAX(STD02) FROM staff_TravelData " & _
                           "WHERE STD01 = (SELECT MAX(STD01) FROM staff_TravelData)) Order BY STD03 ASC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("STD01")) = False Then: m_LastKEY(0) = rsTmp.Fields("STD01")
      If IsNull(rsTmp.Fields("STD02")) = False Then: m_LastKEY(1) = rsTmp.Fields("STD02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(Optional bolQueryFee As Boolean = True)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   If m_CurrKEY(0) <> "" Then
      strSql = "SELECT *" & _
               " FROM staff_TravelData,staff" & _
               " WHERE STD01='" & m_CurrKEY(0) & "' and STD02=" & m_CurrKEY(1) & _
               " and STD01=ST01(+)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         ClearField
         If IsNull(rsTmp.Fields("STD01")) = False Then
            textSTD01 = rsTmp.Fields("STD01"): textSTD01_2 = rsTmp.Fields("ST02")
         End If
         If IsNull(rsTmp.Fields("STD02")) = False Then
            textSTD02 = TAIWANDATE(rsTmp.Fields("STD02"))
         End If
         If IsNull(rsTmp.Fields("STD03")) = False Then
            textSTD03 = Val(rsTmp.Fields("STD03"))
            textSTD03.Tag = textSTD03
         End If
         If IsNull(rsTmp.Fields("STD04")) = False Then
            textSTD04 = TAIWANDATE(rsTmp.Fields("STD04"))
         End If
         If IsNull(rsTmp.Fields("STD05")) = False Then
            textSTD05 = TAIWANDATE(rsTmp.Fields("STD05"))
         End If
         If IsNull(rsTmp.Fields("STD06")) = False Then
            textSTD06 = "" & rsTmp.Fields("STD06")
         End If
         
         'Add By Sindy 2021/12/28 付款日期
         LblPayDate.Visible = False
         If IsNull(rsTmp.Fields("STD13")) = False Then
            LblPayDate.Visible = True
            LblPayDate.Caption = "付款日期：" & ChangeWStringToTDateString(rsTmp.Fields("STD13"))
         End If
         '2021/12/28 END
         
         ' 更新CUID
         UpdateCUID rsTmp
         ' 更新暫存區的資料
         UpdateFieldOldData rsTmp
      End If
      rsTmp.Close
      
'      If m_EditMode = 1 Or m_EditMode = 2 Or m_EditMode = 3 Then
'      'If bolQueryFee = True Then
'         Call GetFeeMoney(m_CurrKEY(0)) '可申請近二年補助資料
'      Else
         Call GetTravelDetail(m_CurrKEY(0))
'      End If
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2021/3/10
'補助資料
Sub GetTravelDetail(strEmp As String)
   
On Error GoTo ErrHand
   
   If strEmp <> "" Then
      GRD2.Clear
      Call SetGrd2(True)
      LblFee.Visible = False: LblFee.Tag = ""
      LblUserFee.Visible = False
      strExc(0) = "SELECT sqldateT(STF02),STF03-1911,TO_NUMBER(STF06)," & _
                  "to_char(STF04,'999,999'),to_char(STF05,'999,999'),sqldateT(STD04)||'~'||sqldateT(STD05),STD06,to_char(STF13,'999,999')" & _
                  " FROM staff_TravelFee,staff_TravelData" & _
                  " WHERE STF01='" & strEmp & "'" & _
                  " AND STF01=STD01(+) AND STF02=STD02(+)" & _
                  " order by STF03 asc,STF02 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      Set GRD2.Recordset = RsTemp
      If RsTemp.RecordCount > 0 Then
         Command1.Visible = True 'Add By Sindy 2021/3/11 更新餘額按鈕
         Call SetGrd2(False)
      End If
   End If
   
Exit Sub

ErrHand:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'可申請近二年補助資料
'Modify By Sindy 2020/12/1 若為1月時,還可以看到前年的資料
Public Sub GetFeeMoney(strEmp As String, Optional bolChkErr As Boolean = False)
Dim jj As Integer, ii As Integer
Dim strYear As String, m_Year As String
Dim bolFind As Boolean
Dim dblFee As Double
Dim strST13 As String, strST04 As String, strSTF03_max As String
'Dim intStarCnt As Integer 'Add By Sindy 2020/12/1
   
On Error GoTo ErrHand
   
   'Add By Sindy 2021/1/5 有輸入員工編號後,申請日不可空白,下列的補助年度依申請日的年判斷
   If textSTD01.Text <> "" Then
      If textSTD02.Text = "" Then
         MsgBox "申請日不可以空白！", vbExclamation
         textSTD02.SetFocus
         Exit Sub
      End If
   End If
   '2021/1/5 END
   
   If strEmp <> "" Then
      If LblFee.Tag = "" Or LblFee.Tag <> strEmp Then
         LblFee.Visible = True
         LblUserFee.Visible = True
         LblUserFee.Caption = "": LblUserFee.Tag = 0
         LblUserFee2.Caption = "": LblUserFee2.Tag = 0 'Add By Sindy 2023/7/25
         'LblFee.Tag = strEmp
         GRD2.Clear
         Call SetGrd2(True)
         
         strSql = "select * from staff" & _
                  ",(select nvl(max(stf03),0) stf03_max from staff_travelfee where stf01='" & strEmp & "')" & _
                  " where ST01='" & strEmp & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strST13 = "" & RsTemp.Fields("st13")
            'Add By Sindy 2021/3/5
            strST04 = "" & RsTemp.Fields("st04")
            strSTF03_max = "" & RsTemp.Fields("stf03_max")
            '2021/3/5 END
         End If
         
'         'Add By Sindy 2020/12/1 若為1月時,還可以看到前年的資料
'         'Modify By Sindy 2021/3/5 改看3年資料
''         If Mid(strSrvDate(1), 5, 2) = "01" Then
'            strFeeYear0 = Val(Left(strSrvDate(1), 4)) - 2
'            intStarCnt = 0
''         Else
''            strFeeYear0 = ""
''            intStarCnt = 1
''         End If
'         '2020/12/1 END
'         strFeeYear1 = Val(Left(strSrvDate(1), 4)) - 1
'         strFeeYear2 = Left(strSrvDate(1), 4)
         'Add By Sindy 2021/1/5
         '異動資料時:
         '多看一年的申請資料
         If m_EditMode = 1 And textSTD02 <> "" Then 'Add By Sindy 2021/1/5 新增時,補助年度依申請日的年判斷
            strFeeYear0 = Val(Left(DBDATE(textSTD02), 4)) - 2
            '可申請近二年補助資料
            strFeeYear1 = Val(Left(DBDATE(textSTD02), 4)) - 1
            strFeeYear2 = Left(DBDATE(textSTD02), 4)
         Else
            strFeeYear0 = Val(Left(strSrvDate(1), 4)) - 2
            '可申請近二年補助資料
            strFeeYear1 = Val(Left(strSrvDate(1), 4)) - 1
            strFeeYear2 = Left(strSrvDate(1), 4)
         End If
         '2021/1/5 END
         strExc(0) = "SELECT sqldateT(STF02),STF03-1911,TO_NUMBER(STF06)," & _
                     "to_char(STF04,'999,999'),to_char(STF05,'999,999'),sqldateT(STD04)||'~'||sqldateT(STD05),STD06,to_char(stf13,'999,999')" & _
                     " FROM staff_TravelFee,staff_TravelData" & _
                     " WHERE STF01='" & strEmp & "' and STF03 in(" & IIf(strFeeYear0 <> "", strFeeYear0 & ",", "") & strFeeYear1 & "," & strFeeYear2 & ")" & _
                     " AND STF01=STD01(+) AND STF02=STD02(+)" & _
                     " order by STF03 asc,STF02 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         Set GRD2.Recordset = RsTemp
         If RsTemp.RecordCount > 0 Then LblFee.Tag = strEmp
         If bolChkErr = False Then
            For jj = 0 To 2 '補助二年
               If jj = 0 Then strYear = strFeeYear0 'Add By Sindy 2020/12/1
               If jj = 1 Then strYear = strFeeYear1
               If jj = 2 Then strYear = strFeeYear2
               bolFind = False
               For ii = 1 To GRD2.Rows - 1 '費用資料
                  If Val(GRD2.TextMatrix(ii, 1)) = Val(strYear) - 1911 Then
                     bolFind = True
                     Exit For
                  End If
               Next ii
               If bolFind = False Then
                  '年資
                  m_Year = Trim(CalYear(strEmp, Val(strYear) - 1 & "1231"))
                  'Modify By Sindy 2020/7/13 控管工作時間滿一年才能顯示,才能申請補助
                  'Val(strSrvDate(1)) >= Val(DBDATE(DateAdd("yyyy", 1, ChangeWStringToWDateString(strST13))))
                  If Trim(CalYear(strEmp, strSrvDate(1))) >= 1 And m_Year > 0 Then
                  '2020/7/13 END
                     'If strST13 <= Val(strYear) & "1231" Then
                        '讀取補助金額
                        strExc(0) = "SELECT to_char(STM03,'999,999'),STM03" & _
                                    " FROM staff_TravelMoney" & _
                                    " WHERE STM01='" & strYear & "' and STM02<=" & m_Year & _
                                    " order by STM02 DESC"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 0 Then
                           strExc(0) = "SELECT to_char(STM03,'999,999'),STM03" & _
                                       " FROM staff_TravelMoney" & _
                                       " WHERE STM01='" & strYear & "'" & _
                                       " order by STM02 asc"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        End If
                        If intI = 1 Then
                           LblFee.Tag = strEmp
                           GRD2.AddItem " "
                           GRD2.TextMatrix(ii, 1) = Val(strYear) - 1911
                           GRD2.TextMatrix(ii, 2) = m_Year '年資
                           GRD2.TextMatrix(ii, 3) = RsTemp.Fields(0)
                           GRD2.TextMatrix(ii, 7) = Format(RsTemp.Fields(1), "#,##0")
                        End If
                     'End If
                  End If
               End If
            Next jj
         End If
         '目前可使用的補助金
         dblFee = 0: strYear = ""
         For ii = 1 To GRD2.Rows - 1
            If GRD2.TextMatrix(ii, 1) <> "" Then '補助年度
               '剩餘額度
               If GRD2.TextMatrix(ii, 4) = "" Then
                  GRD2.TextMatrix(ii, 7) = GRD2.TextMatrix(ii, 3)
                  dblFee = 0
'               Else
'                  If strYear = "" Or Val(strYear) <> Val(GRD2.TextMatrix(ii, 1)) Then
'                     dblFee = CDbl(GRD2.TextMatrix(ii, 3)) - CDbl(GRD2.TextMatrix(ii, 4))
'                     GRD2.TextMatrix(ii, 7) = Format(dblFee, "#,##0")
'                  Else
'                     If dblFee > 0 And Val(strYear) = Val(GRD2.TextMatrix(ii, 1)) Then
'                        dblFee = dblFee - CDbl(GRD2.TextMatrix(ii, 4))
'                        GRD2.TextMatrix(ii - 1, 7) = 0
'                        GRD2.TextMatrix(ii, 7) = Format(dblFee, "#,##0")
'                     Else
'                        dblFee = CDbl(GRD2.TextMatrix(ii, 3)) - CDbl(GRD2.TextMatrix(ii, 4))
'                        GRD2.TextMatrix(ii, 7) = Format(dblFee, "#,##0")
'                     End If
'                  End If
               End If
               strYear = GRD2.TextMatrix(ii, 1)
            End If
         Next ii
         dblFee = 0
         For ii = 1 To GRD2.Rows - 1
            'Add By Sindy 2021/3/5 增加顯示前2年申請資料，但不列入餘額計算。
            If Val(GRD2.TextMatrix(ii, 1)) > Val(strFeeYear0) - 1911 Then
            '2021/3/5 END
               dblFee = dblFee + CDbl(Val(Format(GRD2.TextMatrix(ii, 7), "##0")))
            End If
         Next ii
         LblUserFee.Tag = Format(dblFee, "#,##0")
         LblUserFee2.Tag = LblUserFee.Tag 'Add By Sindy 2023/7/25
         LblUserFee.Caption = " (可使用額度：" & LblUserFee.Tag & ")"
         LblUserFee2.Caption = " (可使用額度：" & LblUserFee2.Tag & ")" 'Add By Sindy 2023/7/25
         If GRD2.Rows > 1 Then
            GRD2.row = 1
            GRD2.col = 0
         End If
         Call SetGrd2(False)
      End If
   End If
   
Exit Sub

ErrHand:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Public Sub GetData()
Dim rsTmp As New ADODB.Recordset
Dim ii As Integer
Dim strKey1 As String, StrKey2 As String
   
   strSql = ""
   If txt1(0) <> "" Then
       strSql = strSql & " and STD01>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       strSql = strSql & " and STD01<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
      'Modify By Sindy 2023/12/22
      If strSrvDate(1) >= 新部門啟用日 Then
         strSql = strSql & " and ST93>='" & txt1(2) & "' "
      Else
      '2023/12/22 END
         strSql = strSql & " and ST03>='" & txt1(2) & "' "
      End If
   End If
   If txt1(3) <> "" Then
      'Modify By Sindy 2023/12/22
      If strSrvDate(1) >= 新部門啟用日 Then
         strSql = strSql & " and ST93<='" & txt1(3) & "' "
      Else
      '2023/12/22 END
         strSql = strSql & " and ST03<='" & txt1(3) & "' "
      End If
   End If
   'Add By Sindy 2019/10/23
   If txt1(4) <> "" Then
       strSql = strSql & " and STD02>=" & DBDATE(txt1(4))
   End If
   If txt1(5) <> "" Then
       strSql = strSql & " and STD02<=" & DBDATE(txt1(5))
   End If
   '2019/10/23 END
   
   '抓取資料
   'Modify By Sindy 2020/3/16 + ,sqldateT(STD13)
   strSql = "SELECT STD01,st02,sqldateT(STD02),to_char(STD03,'999,999'),STF03-1911," & _
            "to_char(STF04,'999,999'),to_char(STF05,'999,999'),sqldateT(STD04)||'~'||sqldateT(STD05),STD06,sqldateT(STD13),STD01 as STD01_1,sqldateT(STD02) as STD02_1" & _
            " FROM staff_TravelFee,staff,staff_TravelData" & _
            " where STD01=st01(+) and STD01=STF01(+) and STD02=STF02(+)" & strSql & _
            " order by STD01 desc,STD02 desc,STF03 asc"
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   GRD1.FixedCols = 0 'Add By Sindy 2019/10/23
   Set GRD1.Recordset = rsTmp
   SetGrd
   GRD1.FixedCols = 2 'Add By Sindy 2019/10/23
   For ii = 1 To GRD1.Rows - 1
      If strKey1 <> "" And GRD1.TextMatrix(ii, 0) <> "" And _
         (strKey1 = GRD1.TextMatrix(ii, 0) And StrKey2 = GRD1.TextMatrix(ii, 2)) Then
         GRD1.TextMatrix(ii, 0) = ""
         GRD1.TextMatrix(ii, 1) = ""
         GRD1.TextMatrix(ii, 2) = ""
         GRD1.TextMatrix(ii, 3) = ""
         GRD1.TextMatrix(ii, 7) = ""
         GRD1.TextMatrix(ii, 8) = ""
      Else
         strKey1 = GRD1.TextMatrix(ii, 0)
         StrKey2 = GRD1.TextMatrix(ii, 2)
      End If
   Next ii
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

   CheckDataValid = False

   nResponse = False
   textSTD01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSTD02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSTD04_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSTD05_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSTD01.Locked = bEnable
   textSTD02.Locked = bEnable
   If bEnable Then textSTD01.BackColor = &H8000000F Else textSTD01.BackColor = &H80000005
   If bEnable Then textSTD02.BackColor = &H8000000F Else textSTD02.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textSTD01.Locked = bEnable
   textSTD02.Locked = bEnable
   If bEnable Then textSTD01.BackColor = &H8000000F Else textSTD01.BackColor = &H80000005
   If bEnable Then textSTD02.BackColor = &H8000000F Else textSTD02.BackColor = &H80000005
   If bEnable Then textSTD03.BackColor = &H8000000F Else textSTD03.BackColor = &H80000005
   If bEnable Then textSTD04.BackColor = &H8000000F Else textSTD04.BackColor = &H80000005
   If bEnable Then textSTD05.BackColor = &H8000000F Else textSTD05.BackColor = &H80000005
   If bEnable Then textSTD06.BackColor = &H8000000F Else textSTD06.BackColor = &H80000005
   textSTD03.Locked = bEnable
   textSTD04.Locked = bEnable
   textSTD05.Locked = bEnable
   textSTD06.Locked = bEnable
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textSTD01 = Empty
   textSTD01_2 = Empty
   textSTD02 = Empty
   textSTD03 = Empty: textSTD03.Tag = Empty
   textSTD04 = Empty
   textSTD05 = Empty
   textSTD06 = Empty
   Label23 = Empty
   SetGrd
   Call SetGrd2(True)
   
   For nIndex = 0 To tf_STD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
   Command1.Visible = False 'Add By Sindy 2021/3/11 更新餘額按鈕
   LblPayDate.Visible = False 'Add By Sindy 2021/12/28 付款日期
End Sub

Private Sub UpdateFieldNewData()
   Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "STD01", textSTD01
      SetFieldNewData "STD02", DBDATE(textSTD02)
   End If
   SetFieldNewData "STD03", textSTD03
   SetFieldNewData "STD04", DBDATE(textSTD04)
   SetFieldNewData "STD05", DBDATE(textSTD05)
   SetFieldNewData "STD06", textSTD06
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To tf_STD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "STD" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 2, 3, 4, 5:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
   SetGrd
End Sub

Private Sub textSTD01_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSTD01
      CloseIme
   End If
End Sub

Private Sub textSTD01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSTD01_Validate(Cancel As Boolean)
Dim rsTmp As New ADODB.Recordset

   If m_EditMode = 1 And textSTD01 <> "" Then
      textSTD01_2 = GetStaffName(textSTD01, True)
      If IsRecordExist(textSTD01, DBDATE(textSTD02)) = True And textSTD01.Enabled = True And textSTD01.Locked = False Then
         MsgBox "該員工當天已有申請資料，請修改！", vbInformation
'         Call GetFeeMoney(textSTD01) '可申請近二年補助資料
         Cancel = True
         Exit Sub
      End If
      If textSTD01_2 = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         Cancel = True
      Else
         If ChkStaffST04(textSTD01, False) = True Then
            MsgBox "此員工已離職！", vbInformation
            Cancel = True
         End If
      End If
      'Add By Sindy 2022/4/11
      'Modify By Sindy 2022/6/6 + And Val(textSTD02) = 0
      If Cancel = False And Val(textSTD02) = 0 Then
         textSTD02 = strSrvDate(2)
         Call textSTD02_Validate(False)
      End If
      '2022/4/11 END
'      Call GetFeeMoney(textSTD01, Cancel) '可申請近二年補助資料
'      LblFee.Tag = ""
   End If
End Sub

Private Sub textSTD02_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSTD02
   End If
End Sub

Private Sub textSTD02_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSTD02_Validate(Cancel As Boolean)
   If m_EditMode = 1 And textSTD02 <> "" Then
      If IsRecordExist(textSTD01, DBDATE(textSTD02)) = True And textSTD02.Enabled = True And textSTD02.Locked = False Then
         MsgBox "該員工當天已有申請資料，請修改！", vbInformation
         Call GetFeeMoney(textSTD01) '可申請近二年補助資料 'Add By Sindy 2021/1/5
         Cancel = True
         Exit Sub
      End If
      If CheckIsTaiwanDate(textSTD02, False) = False Then
         Cancel = True
         MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
         Exit Sub
      End If
      If Val(textSTD02) > Val(strSrvDate(2)) Then
         Cancel = True
         MsgBox "申請日不可大於系統日！", vbInformation, "輸入日期錯誤"
         Exit Sub
      End If
      
      Call GetFeeMoney(textSTD01, Cancel) '可申請近二年補助資料
      LblFee.Tag = ""
   End If
End Sub

Private Sub textSTD03_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSTD03
   End If
End Sub

Private Sub textSTD03_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSTD04_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSTD04
   End If
End Sub

Private Sub textSTD04_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSTD04_Validate(Cancel As Boolean)
   If m_EditMode = 1 And textSTD04 <> "" Then
      If CheckIsTaiwanDate(textSTD04, False) = False Then
         Cancel = True
         MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
         Exit Sub
      End If
      If textSTD04 <> "" And textSTD05 = "" Then
         textSTD05 = textSTD04
      End If
   End If
End Sub

Private Sub textSTD05_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSTD05
   End If
End Sub

Private Sub textSTD05_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSTD05_Validate(Cancel As Boolean)
   If m_EditMode = 1 And textSTD05 <> "" Then
      If CheckIsTaiwanDate(textSTD05, False) = False Then
         Cancel = True
         MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
         Exit Sub
      End If
      If RunNick(textSTD04, textSTD05) Then
         Call textSTD05_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   '                        0           1       2         3           4           5           6           7           8       9           10       11
   arrGridHeadText = Array("員工編號", "姓名", "申請日", "申請金額", "補助年度", "補助額度", "補助金額", "旅遊期間", "備註", "付款日期", "STD01_1", "STD02_1")
   arrGridHeadWidth = Array(800, 800, 800, 800, 850, 850, 800, 1700, 2000, 800, 0, 0)
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

Private Sub SetGrd2(bolDefault As Boolean)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   '                        0         1           2       3           4           5           6       7
   arrGridHeadText = Array("申請日", "補助年度", "年資", "補助額度", "補助金額", "旅遊期間", "備註", "剩餘額度")
   arrGridHeadWidth = Array(800, 900, 450, 900, 900, 1800, 2000, 900)
   GRD2.Visible = False
   GRD2.Cols = UBound(arrGridHeadText) + 1
   If bolDefault = True Then GRD2.Rows = 2
   For iRow = 0 To GRD2.Cols - 1
      GRD2.row = 0
      GRD2.col = iRow
      GRD2.Text = arrGridHeadText(iRow)
      GRD2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD2.CellAlignment = flexAlignCenterCenter
   Next
   GRD2.Visible = True
End Sub

Private Sub textSTD06_GotFocus()
   InverseTextBox textSTD06
   '切換輸入法改用API
   OpenIme
End Sub

Private Sub textSTD06_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textSTD06, textSTD06.MaxLength) = False Then
      Cancel = True
      textSTD06_GotFocus
   End If
   '切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
'      Case 2, 3
'         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0, 1 '員工編號
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
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
      Case 2, 3 '部門
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 4, 5 '申請日期
         If CheckIsTaiwanDate(txt1(Index), False) = False Then
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
         If Index = 5 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
