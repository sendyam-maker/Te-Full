VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160017 
   BorderStyle     =   1  '單線固定
   Caption         =   "可補休資料"
   ClientHeight    =   5070
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8190
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢下一筆(&N)"
      CausesValidation=   0   'False
      Height          =   324
      Index           =   6
      Left            =   4950
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   324
      Index           =   5
      Left            =   6320
      Style           =   1  '圖片外觀
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   756
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4380
      Left            =   30
      TabIndex        =   13
      Top             =   660
      Width           =   8120
      _ExtentX        =   14323
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm160017.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(2)=   "Label23"
      Tab(0).Control(3)=   "textSRR02_2"
      Tab(0).Control(4)=   "Label1(17)"
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(10)=   "textSRR04"
      Tab(0).Control(11)=   "textSRR02"
      Tab(0).Control(12)=   "textSRR05"
      Tab(0).Control(13)=   "textSRR01"
      Tab(0).Control(14)=   "textSRR03"
      Tab(0).Control(15)=   "txtB1008_14(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textSRR12"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160017.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Line1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt1(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt1(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt1(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txt1(3)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdOK(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "GRD1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txt1(5)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txt1(4)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtB1008_14(0)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      Begin VB.TextBox textSRR12 
         Height          =   270
         Left            =   -73560
         Locked          =   -1  'True
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   2700
         Width           =   500
      End
      Begin VB.TextBox txtB1008_14 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   -70200
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "@可補休：剩餘 3.5 天"
         Top             =   440
         Width           =   2760
      End
      Begin VB.TextBox txtB1008_14 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "@可補休：剩餘 3.5 天"
         Top             =   720
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   4
         Left            =   4200
         MaxLength       =   7
         TabIndex        =   9
         Top             =   660
         Width           =   890
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   5
         Left            =   5190
         MaxLength       =   7
         TabIndex        =   10
         Top             =   660
         Width           =   890
      End
      Begin VB.TextBox textSRR03 
         Height          =   270
         Left            =   -73560
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1050
         Width           =   500
      End
      Begin VB.TextBox textSRR01 
         Height          =   270
         Left            =   -73560
         MaxLength       =   7
         TabIndex        =   1
         Top             =   720
         Width           =   1010
      End
      Begin VB.TextBox textSRR05 
         Height          =   270
         Left            =   -73560
         MaxLength       =   7
         TabIndex        =   4
         Top             =   2370
         Width           =   1010
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160017.frx":0038
         Height          =   3380
         Left            =   30
         TabIndex        =   14
         Top             =   960
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   5962
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "員工編號|姓名|發生日期|時數|事由|補休到期日"
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
      Begin VB.CommandButton cmdOK 
         Caption         =   "查詢"
         Height          =   345
         Index           =   0
         Left            =   6390
         TabIndex        =   11
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   5190
         MaxLength       =   7
         TabIndex        =   8
         Top             =   360
         Width           =   890
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   4200
         MaxLength       =   7
         TabIndex        =   7
         Top             =   360
         Width           =   890
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   6
         Top             =   360
         Width           =   740
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   1050
         MaxLength       =   6
         TabIndex        =   5
         Top             =   360
         Width           =   740
      End
      Begin VB.TextBox textSRR02 
         Height          =   270
         Left            =   -73560
         MaxLength       =   6
         TabIndex        =   0
         Top             =   390
         Width           =   735
      End
      Begin VB.TextBox textSRR04 
         Height          =   930
         Left            =   -73560
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   3
         Top             =   1380
         Width           =   5900
      End
      Begin VB.Label Label8 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "目前："
         Height          =   180
         Left            =   -70800
         TabIndex        =   32
         Top             =   440
         Width           =   540
      End
      Begin VB.Label Label7 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "過期時數："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74540
         TabIndex        =   31
         Top             =   2760
         Width           =   950
      End
      Begin VB.Line Line1 
         X1              =   4920
         X2              =   5520
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "補休到期日："
         Height          =   180
         Left            =   3120
         TabIndex        =   25
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label5 
         Caption         =   "注意：輸入可補休時數時，要注意所內特殊工作時數的人員。"
         ForeColor       =   &H000000FF&
         Height          =   430
         Left            =   -74520
         TabIndex        =   24
         Top             =   3390
         Width           =   6520
      End
      Begin VB.Label Label4 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "時數："
         Height          =   180
         Left            =   -74130
         TabIndex        =   23
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "補休到期日："
         Height          =   180
         Left            =   -74670
         TabIndex        =   22
         Top             =   2400
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "事由："
         Height          =   180
         Index           =   17
         Left            =   -74130
         TabIndex        =   15
         Top             =   1410
         Width           =   540
      End
      Begin MSForms.Label textSRR02_2 
         Height          =   230
         Left            =   -72780
         TabIndex        =   21
         Top             =   420
         Width           =   1400
         BackColor       =   12632256
         VariousPropertyBits=   27
         Size            =   "2461;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label23 
         Height          =   200
         Left            =   -74760
         TabIndex        =   20
         Top             =   4050
         Width           =   7700
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13582;353"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "發生日期起："
         Height          =   180
         Left            =   3120
         TabIndex        =   19
         Top             =   390
         Width           =   1080
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
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line5 
         X1              =   4920
         X2              =   5520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "發生日期："
         Height          =   180
         Left            =   -74490
         TabIndex        =   17
         Top             =   780
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Index           =   0
         Left            =   -74490
         TabIndex        =   16
         Top             =   440
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
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
            Picture         =   "frm160017.frx":004D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160017.frx":0369
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160017.frx":0685
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160017.frx":0861
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160017.frx":0B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160017.frx":0E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160017.frx":11B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160017.frx":14D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160017.frx":17ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160017.frx":1B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160017.frx":1E25
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8190
      _ExtentX        =   14446
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
Attribute VB_Name = "frm160017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Sindy 2024/10/15
Option Explicit

Dim RcMain As New ADODB.Recordset, RsAdo As New ADODB.Recordset
' 變數宣告區
Dim m_EditMode As Integer
Dim m_SubMode As Integer
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
Dim m_FirstKEY(2) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(2) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(2) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_SRR As Integer
Public UpForm As Form


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 5 '結束
         Unload Me
         UpForm.Show
      Case 6 '查詢下一筆
         Unload Me
         UpForm.cmdState = 7
         UpForm.PubShowNextData
      Case 0 '查詢
         If txt1(0) & txt1(1) & txt1(2) & txt1(3) & txt1(4) & txt1(5) <> "" Then
             If RunNick(txt1(0), txt1(1)) Then
                 txt1(0).SetFocus
                 Exit Sub
             End If
             If RunNick2(txt1(2), txt1(3)) Then
                 txt1(2).SetFocus
                 Exit Sub
             End If
             If txt1(4) <> "" And txt1(5) <> "" Then
               If RunNick2(txt1(4), txt1(5)) Then
                  txt1(4).SetFocus
                  Exit Sub
               End If
             End If
             GetData
         Else
             MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
         End If
   End Select
End Sub

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from Staff_RepayRest where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_SRR = rsA.Fields.Count
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
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

ReDim m_FieldList(tf_SRR) As FIELDITEM

   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   textSRR01.BackColor = &H8000000F
   textSRR02.BackColor = &H8000000F

   MoveFormToCenter Me

   InitialField
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160017 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
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
            If SSTab1.TabVisible(0) = True Then
               textSRR01.Text = TAIWANDATE(GRD1.TextMatrix(tmpMouseRow, 2))
               textSRR02.Text = GRD1.TextMatrix(tmpMouseRow, 0)
               QueryRecord
            End If
            GRD1.Visible = True
       End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      If cmdok(0).Visible = True And cmdok(0).Enabled = True Then cmdok(0).SetFocus
      cmdok(0).Default = True
   Else
      cmdok(0).Default = False
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

   If IsNull(rsSrcTmp.Fields("SRR06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SRR06")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("SRR06"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SRR07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SRR07")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SRR07"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SRR08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SRR08")) = False Then
         strTemp = rsSrcTmp.Fields("SRR08")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SRR09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SRR09")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("SRR09"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SRR10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SRR10")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SRR10"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SRR11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SRR11")) = False Then
         strTemp = rsSrcTmp.Fields("SRR11")
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
   
   If textSRR02.Text = "" Then
      MsgBox "員工編號不可以空白！", vbExclamation
      textSRR02.SetFocus
      Exit Function
   End If
   If Me.textSRR02.Enabled = True Then
      Cancel = False
      textSRR02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textSRR01.Text = "" Then
      MsgBox "日期不可以空白！", vbExclamation
      textSRR01.SetFocus
      Exit Function
   End If
   If Me.textSRR01.Enabled = True Then
      Cancel = False
      textSRR01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   '增加判斷員工代號+日期是否人員已離職
   If ChkStaffST04(textSRR02, True, textSRR01) = True Then
      textSRR02.SetFocus
      Exit Function
   End If
   
   If textSRR03.Text = "" Then
      MsgBox "時數不可以空白！", vbExclamation
      textSRR03.SetFocus
      Exit Function
   End If
   If Me.textSRR03.Enabled = True Then
      Cancel = False
      textSRR03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2025/4/8
   If Me.textSRR12.Enabled = True Then
      Cancel = False
      textSRR12_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2025/4/8 END
   
   If textSRR04.Text = "" Then
      MsgBox "事由不可以空白！", vbExclamation
      textSRR04.SetFocus
      Exit Function
   End If
   If Me.textSRR04.Enabled = True Then
      Cancel = False
      textSRR04_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textSRR05.Text = "" Then
      MsgBox "補休到期日不可以空白！", vbExclamation
      textSRR05.SetFocus
      Exit Function
   'Modify By Sindy 2024/1/10 二個日期是會有相同的狀況發生
'   ElseIf Val(textSRR05.Text) <= Val(textSRR01.Text) Then
'      MsgBox "補休到期日必須大於發生日期！", vbExclamation
'      textSRR05.SetFocus
'      Exit Function
   End If
   If Me.textSRR05.Enabled = True Then
      Cancel = False
      textSRR05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer

   For nIndex = 0 To tf_SRR - 1
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

   For nIndex = 0 To tf_SRR - 1
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
Dim strSRR01 As String
Dim strSRR02 As String

   AddRecord = False

   strSRR02 = textSRR02
   strSRR01 = DBDATE(textSRR01)

   ' 檢查記錄是否已存在
   If IsRecordExist(strSRR01, strSRR02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If

   bFirst = True
   bDifference = False
   strSql = "INSERT INTO Staff_RepayRest ("
   For nIndex = 0 To tf_SRR - 1
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
   For nIndex = 0 To tf_SRR - 1
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
   cnnConnection.BeginTrans

   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql

   If ((strSRR01 & strSRR02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strSRR01 & strSRR02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans

   ShowCurrRecord DBDATE(strSRR01), strSRR02
   AddRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox " 新增失敗！" & vbCrLf & Err.Description

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
   Dim strSRR01 As String
   Dim strSRR02 As String

   ModRecord = False

   strSRR01 = m_CurrKEY(0)
   strSRR02 = m_CurrKEY(1)

   strSql = "begin user_data.user_enabled:=1; UPDATE Staff_RepayRest SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_SRR - 1
      strTmp = Empty
      'If nIndex < 3 Or nIndex > 8 Then
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
        'End If
   Next nIndex

   strSql = strSql & " " & _
                  "WHERE SRR01 = '" & strSRR01 & "' and SRR02='" & strSRR02 & "' ; end; "
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   cnnConnection.CommitTrans

   ShowCurrRecord DBDATE(strSRR01), strSRR02

   ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)

End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strSRR01 As String
Dim strSRR02 As String

   DelRecord = False

On Error GoTo ErrHand

   cnnConnection.BeginTrans

   strSRR01 = m_CurrKEY(0)
   strSRR02 = m_CurrKEY(1)

   strSql = "DELETE FROM Staff_RepayRest " & _
            "WHERE SRR01 = '" & strSRR01 & "'  and SRR02='" & strSRR02 & "' "

   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql

   If (strSRR01 = m_LastKEY(0) And strSRR02 = m_LastKEY(1)) Or (strSRR01 = m_FirstKEY(0) And strSRR02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord DBDATE(strSRR01), strSRR02
   DelRecord = True
   cnnConnection.CommitTrans

   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strSRR01 As String
Dim strSRR02 As String

   QueryRecord = False
   strSRR01 = DBDATE(textSRR01) '發生日期
   strSRR02 = textSRR02 '員工編號
   If IsRecordExist(strSRR01, strSRR02) = True Then
      m_CurrKEY(0) = strSRR01
      m_CurrKEY(1) = strSRR02
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
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         UpdateFieldNewData
         If AddRecord = True Then
             RefreshRange
         Else
             Exit Function
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         UpdateFieldNewData
         If ModRecord = False Then Exit Function
         
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textSRR01 <> "" And textSRR02 <> "" Then
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
      Case 1: If Me.Visible = True Then textSRR02.SetFocus
      Case 2: If Me.Visible = True Then textSRR03.SetFocus
      Case 4: If Me.Visible = True Then textSRR02.SetFocus
   End Select
End Sub
' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strSRR01 As String, ByVal strSRR02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

   IsRecordExist = False
   strSql = "SELECT * FROM Staff_RepayRest " & _
            "WHERE SRR01 = '" & strSRR01 & "'  and SRR02='" & strSRR02 & "'  "

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
Private Sub ShowCurrRecord(ByVal strSRR01 As String, ByVal strSRR02 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If IsRecordExist(strSRR01, strSRR02) = True Then
      m_CurrKEY(0) = strSRR01
      m_CurrKEY(1) = strSRR02
   Else
      strSql = "SELECT SRR01,SRR02 FROM Staff_RepayRest " & _
               "WHERE SRR01 = '" & m_CurrKEY(0) & "' and SRR02='" & m_CurrKEY(1) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SRR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SRR01")
         If IsNull(rsTmp.Fields("SRR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SRR02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close

      strSql = "SELECT SRR01,SRR02 FROM Staff_RepayRest " & _
               "WHERE SRR02 = (SELECT MIN(SRR02) FROM Staff_RepayRest where SRR01=(select min(SRR01) from Staff_RepayRest) ) and SRR01=(select min(SRR01) from Staff_RepayRest) "

      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SRR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SRR01")
         If IsNull(rsTmp.Fields("SRR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SRR02")
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

   strSql = "SELECT SRR01,SRR02 FROM Staff_RepayRest " & _
            "WHERE SRR01 = '" & m_CurrKEY(0) & "' AND " & _
                  "SRR02 = (SELECT MAX(SRR02) FROM Staff_RepayRest " & _
                          "WHERE SRR01 = '" & m_CurrKEY(0) & "' AND " & _
                                "SRR02 < '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SRR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SRR01")
      If IsNull(rsTmp.Fields("SRR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SRR02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close

   strSql = "SELECT SRR01,SRR02 FROM Staff_RepayRest " & _
            "WHERE SRR01 = (SELECT MAX(SRR01) FROM Staff_RepayRest " & _
                           "WHERE SRR01 < '" & m_CurrKEY(0) & "') AND " & _
                  "SRR02 = (SELECT MAX(SRR02) FROM Staff_RepayRest " & _
                           "WHERE SRR01 = (SELECT MAX(SRR01) FROM Staff_RepayRest " & _
                                          "WHERE SRR01 < '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SRR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SRR01")
      If IsNull(rsTmp.Fields("SRR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SRR02")
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

   strSql = "SELECT SRR01,SRR02 FROM Staff_RepayRest " & _
            "WHERE SRR01 = '" & m_CurrKEY(0) & "' AND " & _
                  "SRR02 = (SELECT MIN(SRR02) FROM Staff_RepayRest " & _
                          "WHERE SRR01 = '" & m_CurrKEY(0) & "' AND " & _
                                "SRR02 > '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SRR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SRR01")
      If IsNull(rsTmp.Fields("SRR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SRR02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close

   strSql = "SELECT SRR01,SRR02 FROM Staff_RepayRest " & _
            "WHERE SRR01 = (SELECT MIN(SRR01) FROM Staff_RepayRest " & _
                           "WHERE SRR01 > '" & m_CurrKEY(0) & "') AND " & _
                  "SRR02 = (SELECT MIN(SRR02) FROM Staff_RepayRest " & _
                           "WHERE SRR01 = (SELECT MIN(SRR01) FROM Staff_RepayRest " & _
                                          "WHERE SRR01 > '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SRR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SRR01")
      If IsNull(rsTmp.Fields("SRR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SRR02")
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

   m_SubMode = 0
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

   strSql = "SELECT SRR01,SRR02 FROM Staff_RepayRest " & _
            "WHERE SRR01 = (SELECT MIN(SRR01) FROM Staff_RepayRest) AND " & _
                  "SRR02 = (SELECT MIN(SRR02) FROM Staff_RepayRest " & _
                           "WHERE SRR01 = (SELECT MIN(SRR01) FROM Staff_RepayRest)) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SRR01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("SRR01")
      If IsNull(rsTmp.Fields("SRR02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("SRR02")
   End If
   rsTmp.Close

   strSql = "SELECT SRR01,SRR02 FROM Staff_RepayRest " & _
            "WHERE SRR01 = (SELECT MAX(SRR01) FROM Staff_RepayRest) AND " & _
                  "SRR02 = (SELECT MAX(SRR02) FROM Staff_RepayRest " & _
                           "WHERE SRR01 = (SELECT MAX(SRR01) FROM Staff_RepayRest)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SRR01")) = False Then: m_LastKEY(0) = rsTmp.Fields("SRR01")
      If IsNull(rsTmp.Fields("SRR02")) = False Then: m_LastKEY(1) = rsTmp.Fields("SRR02")
   End If
   rsTmp.Close

   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer

   strSql = "SELECT * FROM Staff_RepayRest " & _
            "WHERE SRR01='" & m_CurrKEY(0) & "' and SRR02 = '" & m_CurrKEY(1) & "'   "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("SRR01")) = False Then: textSRR01 = TAIWANDATE(rsTmp.Fields("SRR01"))
      If IsNull(rsTmp.Fields("SRR02")) = False Then
         textSRR02 = rsTmp.Fields("SRR02")
         textSRR02_2 = GetStaffName(textSRR02, True)
      End If
      If IsNull(rsTmp.Fields("SRR03")) = False Then: textSRR03 = rsTmp.Fields("SRR03")
      If IsNull(rsTmp.Fields("SRR04")) = False Then: textSRR04 = rsTmp.Fields("SRR04")
      If IsNull(rsTmp.Fields("SRR05")) = False Then: textSRR05 = TAIWANDATE(rsTmp.Fields("SRR05"))
      If IsNull(rsTmp.Fields("SRR12")) = False Then: textSRR12 = rsTmp.Fields("SRR12")
      Call Pub_GetSpecWorkHour(textSRR02, DBDATE(textSRR01))
      Me.txtB1008_14(1) = GetCurrFor14RestDay(m_CurrKEY(1))
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
   End If

   rsTmp.Close

EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset

   strSql = ""
   If txt1(0) <> "" Then
       strSql = strSql & " and SRR02>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       strSql = strSql & " and SRR02<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       strSql = strSql & " and SRR01>=" & DBDATE(txt1(2))
   End If
   If txt1(3) <> "" Then
       strSql = strSql & " and SRR01<=" & DBDATE(txt1(3))
   End If
   If txt1(4) <> "" Then
       strSql = strSql & " and SRR05>=" & DBDATE(txt1(4))
   End If
   If txt1(5) <> "" Then
       strSql = strSql & " and SRR05<=" & DBDATE(txt1(5))
   End If
   '抓取資料
   '員工編號|姓名|發生日期|時數|事由|補休到期日
   strSql = "SELECT SRR02,st02,sqldateT(SRR01),SRR03,SRR04,sqldateT(SRR05),SRR12" & _
            " FROM Staff_RepayRest,staff where SRR02=st01(+) " & strSql & _
            " order by SRR02,SRR05,SRR01 "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set GRD1.Recordset = rsTmp
   SetGrd
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

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSRR01.Locked = bEnable
   textSRR02.Locked = bEnable
   If bEnable Then textSRR01.BackColor = &H8000000F Else textSRR01.BackColor = &H80000005
   If bEnable Then textSRR02.BackColor = &H8000000F Else textSRR02.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer

   textSRR01.Locked = bEnable
   textSRR02.Locked = bEnable
   If bEnable Then textSRR01.BackColor = &H8000000F Else textSRR01.BackColor = &H80000005
   If bEnable Then textSRR02.BackColor = &H8000000F Else textSRR02.BackColor = &H80000005
   textSRR03.Locked = bEnable
   textSRR04.Locked = bEnable
   textSRR05.Locked = bEnable
   textSRR12.Locked = bEnable
End Sub

Private Sub ClearField()
Dim nIndex As Integer

   textSRR01 = Empty
   textSRR02_2 = Empty
   textSRR02 = Empty
   textSRR03 = Empty
   textSRR04 = Empty
   textSRR05 = Empty
   textSRR12 = Empty
   Label23 = Empty
   SetGrd
   For nIndex = 0 To tf_SRR - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   txtB1008_14(0).Text = Empty
   txtB1008_14(1).Text = Empty
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "SRR01", DBDATE(textSRR01)
      SetFieldNewData "SRR02", textSRR02
   End If
   SetFieldNewData "SRR03", textSRR03
   SetFieldNewData "SRR04", textSRR04
   SetFieldNewData "SRR05", DBDATE(textSRR05)
   SetFieldNewData "SRR12", textSRR12
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String

   ' 初始化欄位陣列
   For nIndex = 1 To tf_SRR
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "SRR" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 1:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
   SetGrd
End Sub

Private Sub textSRR01_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSRR01
   End If
End Sub

Private Sub textSRR01_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSRR01_Validate(Cancel As Boolean)
   If textSRR01 = "" Then Exit Sub
   If CheckIsTaiwanDate(textSRR01, False) = False Then
       Cancel = True
       textSRR01.SetFocus
       MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
       Exit Sub
   End If
   If m_EditMode = 1 Then
       If IsRecordExist(DBDATE(textSRR01), textSRR02) = True And textSRR01.Enabled = True And textSRR01.Locked = False Then
           MsgBox "該員工當天已有資料，請修改！", vbInformation
           textSRR01.SetFocus
           Cancel = True
           Exit Sub
       End If
   End If
End Sub

Private Sub textSRR02_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSRR02
       CloseIme
   End If
End Sub

Private Sub textSRR02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSRR02_Validate(Cancel As Boolean)
   If textSRR02 = "" Then Exit Sub
   textSRR02_2 = GetStaffName(textSRR02, True)
   If textSRR02_2 = "" Then
      MsgBox "員工編號錯誤！查無此員工！", vbInformation
      textSRR02.SetFocus
      Cancel = True
      Exit Sub
   End If
   If m_EditMode = 1 Then
      If IsRecordExist(DBDATE(textSRR01), textSRR02) = True And textSRR02.Enabled = True And textSRR02.Locked = False Then
         MsgBox "該員工當天已有資料，請修改！", vbInformation
         textSRR02.SetFocus
         Cancel = True
         Exit Sub
      End If
   End If
   If m_EditMode = 1 Or m_EditMode = 2 Then
      Call Pub_GetSpecWorkHour(textSRR02, DBDATE(textSRR01))
   End If
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("員工編號", "姓名", "發生日期", "時數", "事由", "補休到期日", "過期時數")
   'Add By Sindy 2025/5/26
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M21" Then
      arrGridHeadWidth = Array(900, 900, 900, 900, 2400, 900, 800)
   Else
   '2025/5/26 END
      arrGridHeadWidth = Array(900, 900, 900, 900, 3000, 1000, 0)
   End If
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

Private Sub textSRR03_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSRR03
   End If
End Sub

Private Sub textSRR03_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc(".") And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textSRR03_Validate(Cancel As Boolean)
   If textSRR03 <> "" Then
      If CheckLengthIsOK(textSRR03, textSRR03.MaxLength) = False Then
         textSRR03.SetFocus
         Cancel = True
         Exit Sub
      End If
      If Val(textSRR03) > PUB_intWkHour Then
         MsgBox "此人員一天最多只有工作 " & PUB_intWkHour & " 小時！"
         textSRR03.SetFocus
         Cancel = True
         Exit Sub
      End If
   End If
   CloseIme
End Sub

Private Sub textSRR04_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSRR04
   End If
End Sub

Private Sub textSRR04_Validate(Cancel As Boolean)
   If textSRR04 <> "" Then
      If CheckLengthIsOK(textSRR04, textSRR04.MaxLength) = False Then
          textSRR04.SetFocus
          Cancel = True
          Exit Sub
      End If
   End If
   CloseIme
End Sub

Private Sub textSRR05_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSRR05
   End If
End Sub

Private Sub textSRR05_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSRR05_Validate(Cancel As Boolean)
   If textSRR05 <> "" Then
       If CheckIsTaiwanDate(textSRR05, False) = False Then
           textSRR05.SetFocus
           Cancel = True
           MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
           Exit Sub
       End If
       If ChkWorkDay(DBDATE(textSRR05)) = False Then
           textSRR05.SetFocus
           Cancel = True
           MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
           Exit Sub
       End If
   End If
End Sub

'Add By Sindy 2025/4/8
Private Sub textSRR12_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSRR12
   End If
End Sub
Private Sub textSRR12_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc(".") And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub textSRR12_Validate(Cancel As Boolean)
   If textSRR12 <> "" Then
      If CheckLengthIsOK(textSRR12, textSRR12.MaxLength) = False Then
         textSRR12.SetFocus
         Cancel = True
         Exit Sub
      End If
      If Val(textSRR12) > Val(textSRR03) Then
         MsgBox "過期時數不可超過 " & textSRR03 & " 小時！"
         textSRR12.SetFocus
         Cancel = True
         Exit Sub
      End If
   End If
   CloseIme
End Sub
'2025/4/8 END

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
   Case 0, 1
           KeyAscii = UpperCase(KeyAscii)
   Case 2, 3, 4, 5
           KeyAscii = Pub_NumAscii(KeyAscii)
   Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0, 1
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
      Case 2, 3
              If CheckIsTaiwanDate(txt1(Index), False) = False Then
                  Cancel = True
                  MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
                  Exit Sub
              End If
              If Index = 3 Then
                  If RunNick2(txt1(Index - 1), txt1(Index)) Then
                      Cancel = True
                      Exit Sub
                  End If
              End If
      Case 4, 5
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
