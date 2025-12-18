VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140502 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶應收帳款收文檢查上限"
   ClientHeight    =   5070
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8170
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8170
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
            Picture         =   "frm140502.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140502.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140502.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140502.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140502.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140502.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140502.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140502.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140502.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140502.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140502.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8170
      _ExtentX        =   14411
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4350
      Left            =   60
      TabIndex        =   4
      Top             =   660
      Width           =   8115
      _ExtentX        =   14323
      _ExtentY        =   7673
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm140502.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblCU13Nm"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblCU13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label23"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LabelCRA01"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblCustName"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label7(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textCRA01"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textCRA02"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCustNO"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCU183"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm140502.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(1)=   "Line2"
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "GRD1"
      Tab(1).Control(5)=   "cmdok"
      Tab(1).Control(6)=   "txt1(2)"
      Tab(1).Control(7)=   "txt1(1)"
      Tab(1).Control(8)=   "txt1(0)"
      Tab(1).ControlCount=   9
      Begin VB.TextBox txtCU183 
         Height          =   270
         Left            =   2370
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1670
         Width           =   1485
      End
      Begin VB.TextBox txtCustNO 
         Height          =   270
         Left            =   1650
         MaxLength       =   8
         TabIndex        =   1
         Top             =   900
         Width           =   1005
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73950
         MaxLength       =   6
         TabIndex        =   9
         Top             =   360
         Width           =   885
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -70890
         MaxLength       =   8
         TabIndex        =   10
         Top             =   360
         Width           =   1000
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -69750
         MaxLength       =   8
         TabIndex        =   11
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Default         =   -1  'True
         Height          =   315
         Left            =   -68310
         TabIndex        =   12
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox textCRA02 
         Height          =   270
         Left            =   2370
         MaxLength       =   8
         TabIndex        =   3
         Top             =   2055
         Width           =   1485
      End
      Begin VB.TextBox textCRA01 
         Height          =   270
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   5
         Top             =   390
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm140502.frx":212C
         Height          =   3615
         Left            =   -74970
         TabIndex        =   16
         Top             =   690
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6368
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "客戶編號|客戶名稱|智權人員|應收帳款上限"
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
      Begin VB.Label Label7 
         Caption         =   "備註2：應收帳款收文額度上限基本為30萬。"
         ForeColor       =   &H000000C0&
         Height          =   200
         Index           =   1
         Left            =   510
         TabIndex        =   28
         Top             =   3480
         Width           =   6230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CRA01："
         Height          =   180
         Index           =   4
         Left            =   750
         TabIndex        =   27
         Top             =   450
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label7 
         Caption         =   "備註1：應收帳款上限分開管制為""應收帳款上限""和""集團應收帳款上限"""
         ForeColor       =   &H000000C0&
         Height          =   200
         Index           =   2
         Left            =   510
         TabIndex        =   26
         Top             =   3000
         Width           =   6230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "應收帳款上限："
         Height          =   180
         Index           =   3
         Left            =   750
         TabIndex        =   25
         Top             =   1715
         Width           =   1260
      End
      Begin MSForms.Label lblCustName 
         Height          =   300
         Left            =   2700
         TabIndex        =   24
         Top             =   915
         Width           =   4845
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74910
         TabIndex        =   23
         Top             =   390
         Width           =   900
      End
      Begin VB.Line Line4 
         X1              =   -73260
         X2              =   -72660
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   -71640
         TabIndex        =   22
         Top             =   390
         Width           =   540
      End
      Begin VB.Line Line3 
         X1              =   -70410
         X2              =   -69660
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label LabelCRA01 
         AutoSize        =   -1  'True
         Height          =   240
         Left            =   2670
         TabIndex        =   21
         Top             =   405
         Visible         =   0   'False
         Width           =   4845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "集團應收帳款上限："
         Height          =   180
         Index           =   2
         Left            =   750
         TabIndex        =   20
         Top             =   2100
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Index           =   0
         Left            =   750
         TabIndex        =   19
         Top             =   945
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Index           =   1
         Left            =   750
         TabIndex        =   18
         Top             =   1330
         Width           =   900
      End
      Begin MSForms.Label Label23 
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   4050
         Width           =   6615
         VariousPropertyBits=   27
         Caption         =   "Create ID:           Date         Time             Update ID:                Date                  Time"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LblCU13 
         Height          =   300
         Left            =   1650
         TabIndex        =   15
         Top             =   1305
         Width           =   735
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LblCU13Nm 
         Height          =   300
         Left            =   2520
         TabIndex        =   14
         Top             =   1305
         Width           =   2265
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "智權人員："
         Height          =   225
         Left            =   -74880
         TabIndex        =   13
         Top             =   390
         Width           =   945
      End
      Begin VB.Label Label5 
         Caption         =   "客戶編號："
         Height          =   225
         Left            =   -71820
         TabIndex        =   8
         Top             =   390
         Width           =   945
      End
      Begin VB.Line Line2 
         X1              =   -69990
         X2              =   -69600
         Y1              =   480
         Y2              =   480
      End
      Begin MSForms.Label Label6 
         Height          =   300
         Left            =   -73020
         TabIndex        =   7
         Top             =   390
         Width           =   1035
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "，只有母號的客戶才可設定""集團上限""。"
         ForeColor       =   &H000000C0&
         Height          =   200
         Index           =   0
         Left            =   1110
         TabIndex        =   6
         Top             =   3240
         Width           =   3650
      End
   End
End
Attribute VB_Name = "frm140502"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/17 Form2.0已修改
'Memo By Sindy 2012/12/10 智權人員欄已修改
'Create By Sindy 2012/12/10
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
Dim m_FirstKEY(1) As String
' 最後一筆資料的Key值
Dim m_LastKEY(1) As String
' 目前正在顯示的Key值
Dim m_CurrKEY(1) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_CRA As Integer
Dim strText As String, arrKey As Variant


Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   'Modified by Lydia 2020/01/30 共用語法
   'rsA.Open "select * from CustRecAmtLmt where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   strSql = GetSql
   rsA.Open strSql & " and rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   'end 2020/01/30
   tf_CRA = rsA.Fields.Count
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
   ReDim m_FieldList(tf_CRA) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   'Added by Lydia 2019/08/16 財務處只開放查詢功能
   'Mark by Lydia 2020/02/07 經呈報總經理核可，改由財務處瑞婷輸入，請惠予調整系統權限
   'If Pub_StrUserSt03 = "M31" Then
   '    m_bInsert = False
   '    m_bUpdate = False
   '    m_bDelete = False
   '    m_bQuery = True
   'End If
   ''end 2019/08/16
   
   MoveFormToCenter Me
   
   InitialField
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   
   Me.SSTab1.Tab = 0 'Added by Lydia 2015/10/14
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140502 = Nothing
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
   
   If IsNull(rsSrcTmp.Fields("CRA03")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CRA03")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("CRA03"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CRA04")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CRA04")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("CRA04"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CRA05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CRA05")) = False Then
         strTemp = rsSrcTmp.Fields("CRA05")
         'Modified by Lydia 2016/02/15
         'strCTime = Format(strTemp, "##:##")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CRA06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CRA06")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("CRA06"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CRA07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CRA07")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("CRA07"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("CRA08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("CRA08")) = False Then
         strTemp = rsSrcTmp.Fields("CRA08")
         'Modified by Lydia 2016/02/15
         'strUTime = Format(strTemp, "##:##")
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
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   'Modified by Lydia 2020/01/30
   'If Me.textCRA01.Enabled = True Then
   '   Cancel = False
   '   textCRA01_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   If Me.txtCustNO.Enabled = True Then
      Cancel = False
      txtCustNo_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer

   For nIndex = 0 To tf_CRA - 1
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

   For nIndex = 0 To tf_CRA - 1
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
   
   AddRecord = False
      
   ' 檢查記錄是否已存在
   'Mark by Lydia 2020/01/30 改在CheckDataValid
'   If IsRecordExist(textCRA01) = True Then
'      strTit = "新增資料"
'      strMsg = "該筆記錄已存在"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'     'Remove by Lydia 2015/10/14 不顯示已存在記錄
'      'UpdateCtrlData
'      Exit Function
'   End If
   'end 2020/01/30
   
   If Right(ChangeCustomerL(txtCustNO), 3) = "000" Then 'Added by Lydia 2020/01/30 判斷母號
        bFirst = True
        bDifference = False
        strSql = "INSERT INTO CustRecAmtLmt ("
        For nIndex = 0 To tf_CRA - 1
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
        For nIndex = 0 To tf_CRA - 1
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
   End If
   'end 2020/01/30
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Added by Lydia 2020/01/30
   If txtCU183.Tag <> txtCU183.Text Then  '更新個人"應收帳款上限CU183"
      strTmp = "update customer set cu183=" & CNULL(txtCU183.Text, True) & ", cu84='" & strUserNum & "', cu85=" & strSrvDate(1) & ", cu86=" & Left(Format(ServerTime, "000000"), 4) & _
                    " where cu01=" & CNULL(txtCustNO) & " and cu02='0' "
      Pub_SeekTbLog strTmp
      cnnConnection.Execute strTmp
   End If
   If strSql <> "" Then  '新增"集團應收帳款上限"
   'end 2020/01/30
        Pub_SeekTbLog strSql
        cnnConnection.Execute strSql
   End If  'end 2020/01/30
   
   'Modified by Lydia 2020/01/30
   'If (textCRA01 < m_FirstKEY(0)) Or (textCRA01 > m_LastKEY(0)) Then
   If (txtCustNO < m_FirstKEY(0)) Or (txtCustNO > m_LastKEY(0)) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   'Modified by Lydia 2020/01/30
   'ShowCurrRecord textCRA01
   ShowCurrRecord txtCustNO
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
       
   ModRecord = False
   
   strSql = "begin user_data.user_enabled:=1; UPDATE CustRecAmtLmt SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_CRA - 1
      strTmp = Empty
      'If nIndex < 7 Or nIndex > 12 Then
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
   
   'Modified by Lydia 2020/01/30 取前6碼
   strSql = strSql & " " & _
            "WHERE CRA01='" & Left(m_CurrKEY(0), 6) & "'; end;"
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   'Added by Lydia 2020/01/30
   If txtCU183.Tag <> txtCU183.Text Then  '更新個人"應收帳款上限CU183"
      strTmp = "update customer set cu183=" & CNULL(txtCU183.Text, True) & ", cu84='" & strUserNum & "', cu85=" & strSrvDate(1) & ", cu86=" & Left(Format(ServerTime, "000000"), 4) & _
                    " where cu01=" & CNULL(txtCustNO) & " and cu02='0' "
      Pub_SeekTbLog strTmp
      cnnConnection.Execute strTmp
   End If
   'end 2020/01/30
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   cnnConnection.CommitTrans
  
   'Modified by Lydia 2020/01/30
   'ShowCurrRecord m_CurrKEY(0)
   ShowCurrRecord txtCustNO
      
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans

   'Modified by Lydia 2020/01/30 取前6碼
   strSql = "DELETE FROM CustRecAmtLmt WHERE CRA01 = '" & Left(m_CurrKEY(0), 6) & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If (m_CurrKEY(0) = m_LastKEY(0)) Or (m_CurrKEY(0) = m_FirstKEY(0)) Then
      RefreshRange
   End If
   'ShowCurrRecord m_CurrKEY(0) 'Remove by Lydia 2020/01/30
   
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   QueryRecord = False
      
   'Modified by Lydia 2020/01/30
   'If IsRecordExist(textCRA01) = True Then
   '   m_CurrKEY(0) = textCRA01
   If IsRecordExist(txtCustNO) = True Then
      m_CurrKEY(0) = txtCustNO
   'end 2020/01/30
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
         'Added by Lydia 2020/01/30 針對"集團應收帳款上限",判斷母號
         If IsRecordExist(txtCustNO, strTit) = True Then
            If strTit = txtCustNO.Text Then
         'end 2020/01/30
                If DelRecord = True Then
                   RefreshRange
                   ClearField
                   ShowCurrRecord m_CurrKEY(0)
                Else
                   Exit Function
                End If
         'Added by Lydia 2020/01/30
            Else
               strMsg = "非母號的客戶只能修改應收帳款上限!"
               strTit = "刪除資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            End If
         Else '兩者皆不存在
            strMsg = "無此資料"
            strTit = "刪除資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End If
         'end 2020/01/30
      Case 4: '查詢
         'Modified by Lydia 2020/01/30
         'If textCRA01 <> "" Then
         If txtCustNO <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            'Modified by Lydia 2020/01/30
            'If textCRA01 = "" Then
            If txtCustNO = "" Then
               MsgBox "請輸入客戶編號！", vbInformation
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
      'Modified by Lydia 2020/01/30
'      Case 1: If Me.Visible = True Then textCRA01.SetFocus
'      Case 2: If Me.Visible = True Then textCRA02.SetFocus
'      Case 4: If Me.Visible = True Then textCRA01.SetFocus
      Case 1: If Me.Visible = True Then txtCustNO.SetFocus
      Case 2: If Me.Visible = True Then txtCU183.SetFocus
      Case 4: If Me.Visible = True Then txtCustNO.SetFocus
      'end 2020/01/30
   End Select
End Sub

' 檢查記錄是否已經存在
'Modified by Lydia 2020/01/03 + strCRA01
Private Function IsRecordExist(ByVal strKEY01 As String, Optional ByRef strCRA01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   'Modified by Lydia 2020/01/30 共用語法
   'strSql = "SELECT * FROM CustRecAmtLmt WHERE CRA01='" & strKEY01 & "'"
   strSql = GetSql
   strSql = strSql & "AND INSTR(CU01||','||CRA01||'00','" & Left(ChangeCustomerL(strKEY01), 8) & "') > 0 "
   strCRA01 = ""
   'end 2020/01/30
   
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      strCRA01 = "" & rsTmp.Fields("cra01") 'Added by Lydia 2020/01/30
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
   
   'Remove by Lydia 2020/01/30
'   strSql = "select * from CustRecAmtLmt where rownum <2"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount = 0 Then
'      rsTmp.Close
'      Set rsTmp = Nothing
'      Exit Sub
'   End If
'   rsTmp.Close
   'end 2020/01/30
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
   Else
      'Modified by Lydia 2020/01/30 共用語法
      'strSql = "SELECT CRA01 FROM CustRecAmtLmt WHERE CRA01='" & m_CurrKEY(0) & "'"
      strSql = GetSql
      strSql = strSql & "AND INSTR(CU01||','||CRA01||'00','" & Left(ChangeCustomerL(strKEY01), 8) & "') > 0 "
      strSql = strSql & ""
      'end 2020/01/30
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CRA01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CRA01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      'Modified by Lydia 2020/01/30 共用語法
      'strSql = "SELECT min(CRA01) FROM CustRecAmtLmt "
      strSql = GetSql
      strSql = "SELECT min(CNO) FROM (" & strSql & ") "
      'end 2020/01/30
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         m_CurrKEY(0) = "" & rsTmp.Fields(0)
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
   
   'Modified by Lydia 2020/01/30 共用語法
   'strSql = "select max(CRA01) From CustRecAmtLmt where CRA01<'" & m_CurrKEY(0) & "' "
   strSql = GetSql
   strSql = "select max(CNO) From (" & strSql & " and CU01<'" & m_CurrKEY(0) & "') "
   'end 2020/01/30
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_CurrKEY(0) = "" & rsTmp.Fields(0)
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
   
   'Modified by Lydia 2020/01/30 共用語法
   'strSql = "select min(CRA01) From CustRecAmtLmt where CRA01>'" & m_CurrKEY(0) & "'"
   strSql = GetSql
   strSql = "select min(CNO) From (" & strSql & " and CU01>'" & m_CurrKEY(0) & "') "
   'end 2020/01/30
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_CurrKEY(0) = "" & rsTmp.Fields(0)
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
   
   'Remove by Lydia 2020/01/30
'   strSql = "select * from CustRecAmtLmt where rownum <2"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount = 0 Then
'      rsTmp.Close
'      Set rsTmp = Nothing
'      Exit Sub
'   End If
'   rsTmp.Close
   'end 2020/01/30
   
   'Modified by Lydia 2020/01/30 共用語法
   'strSql = "SELECT min(CRA01) FROM CustRecAmtLmt "
   strSql = GetSql
   strSql = "SELECT min(CNO) FROM (" & strSql & ") "
   'end 2020/01/30
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_FirstKEY(0) = "" & rsTmp.Fields(0)
   End If
   rsTmp.Close
   
   'Modified by Lydia 2020/01/30 共用語法
   'strSql = "SELECT max(CRA01) FROM CustRecAmtLmt "
   strSql = GetSql
   strSql = "SELECT max(CNO) FROM (" & strSql & ") "
   'end 2020/01/30
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_LastKEY(0) = "" & rsTmp.Fields(0)
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer
   
   'Modified by Lydia 2020/01/30 共用語法
   'strSql = "SELECT * FROM CustRecAmtLmt WHERE CRA01='" & m_CurrKEY(0) & "'"
   strSql = GetSql("2")
   'strSql = strSql & "AND INSTR(CU01||','||CRA01||'00','" & Left(ChangeCustomerL(m_CurrKEY(0)), 8) & "') > 0 "
   strSql = strSql & "AND CU01='" & m_CurrKEY(0) & "' "
   strSql = strSql & " ORDER BY CNO"
   'end 2020/01/30
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ClearField
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CRA01")) = False Then: textCRA01 = rsTmp.Fields("CRA01")
      If IsNull(rsTmp.Fields("CRA02")) = False Then: textCRA02 = rsTmp.Fields("CRA02")
      
      ' 更新CUID
      UpdateCUID rsTmp
      
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
      
      LabelCRA01 = GetPrjPeople1(textCRA01 & "00", "1")
      
      Call QueryCustData 'Added by Lydia 2015/10/14
   End If

   rsTmp.Close
   
EXITSUB:
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

Private Function CheckDataValid() As Boolean
Dim nResponse As Boolean
Dim strTmp  As String
Dim strTit As String
Dim strMsg As String
   
   CheckDataValid = False
   
   'Modified by Lydia 2020/01/30 textCRA01=> txtCustNo
   If txtCustNO.Text = "" Then
       MsgBox "客戶編號不可空白！", vbExclamation
       txtCustNO.SetFocus
       Exit Function
   End If
   'end 2020/01/30
   
   'Modified by Lydia 2020/01/30 2020/01/30 判斷個人"應收帳款上限"
'   If Val(textCRA02.Text) = 0 Then
'       MsgBox "應收帳款上限不可空白！", vbExclamation
'       textCRA02.SetFocus
'       Exit Function
   If Val(txtCU183.Text) = 0 Then
       If MsgBox("應收帳款上限空白，是否繼續？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
           txtCU183.SetFocus
           Exit Function
       End If
   'end 2020/01/30
   End If
   
   'Added by Lydia 2020/01/30
   If Right(ChangeCustomerL(txtCustNO), 3) = "000" Then
        If m_EditMode = 1 Then '新增
             If Val(textCRA02) > 0 Then
                 If IsRecordExist(txtCustNO, strTmp) = True Then '檢查記錄是否已存在
                    If strTmp <> "" Then
                       strTit = "新增資料"
                       strMsg = "集團應收帳款上限記錄已存在"
                       nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                       txtCustNO.SetFocus
                       Exit Function
                    End If
                 End If
             End If
        End If
   End If
   If Val(txtCU183) > 0 And Val(textCRA02) > 0 And Val(txtCU183) > Val(textCRA02) Then
       MsgBox "應收帳款上限不可大於集團上限！", vbExclamation
       txtCU183.SetFocus
       Exit Function
   End If
   'end 2020/01/30
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textCRA01.Locked = bEnable
   If bEnable Then textCRA01.BackColor = &H8000000F Else textCRA01.BackColor = &H80000005
   'Added by Lydia 2020/01/30
   txtCustNO.Locked = bEnable
   If bEnable Then txtCustNO.BackColor = &H8000000F Else txtCustNO.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textCRA01.Locked = bEnable
   If bEnable Then textCRA01.BackColor = &H8000000F Else textCRA01.BackColor = &H80000005
   textCRA02.Locked = bEnable
   'Added by Lydia 2015/10/14 編輯資料時,關閉切換頁籤
   Me.SSTab1.TabEnabled(1) = bEnable
   If bEnable = False Then Me.SSTab1.Tab = 0
   'Added by Lydia 2020/01/30
   txtCustNO.Locked = bEnable
   If bEnable Then txtCustNO.BackColor = &H8000000F Else txtCustNO.BackColor = &H80000005
   txtCU183.Locked = bEnable
   If bEnable = False Then
        If Right(ChangeCustomerL(txtCustNO), 3) = "000" Then  '判斷非母號不可修改
           textCRA02.Enabled = True
        Else
           textCRA02.Enabled = False
        End If
   End If
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textCRA01 = Empty
   LabelCRA01 = Empty
   textCRA02 = Empty
   Label23 = Empty
   'Added by Lydia 2015/10/14
   LblCU13 = Empty
   LblCU13Nm = Empty
   'Added by Lydia 2020/01/30
   txtCustNO.Text = "": txtCustNO.Tag = ""
   lblCustName.Caption = ""
   txtCU183.Text = "": txtCU183.Tag = ""
   
   For nIndex = 0 To tf_CRA - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      If textCRA01 = "" Then textCRA01 = Left(txtCustNO, 6) 'Added by Lydia 2020/01/30 抓客戶代號前6碼
      SetFieldNewData "CRA01", textCRA01
   End If
   SetFieldNewData "CRA02", textCRA02
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To tf_CRA
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CRA" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 2:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
   
   'Added by Lydia 2015/10/14
   For nIndex = 0 To 2
      txt1(nIndex).Text = ""
   Next nIndex
   Label6.Caption = ""
   SetGrd
End Sub

Private Sub textCRA01_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textCRA01
   End If
End Sub

Private Sub textCRA01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCRA01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   LabelCRA01 = Empty
   If IsEmptyText(textCRA01) = False Then
      LabelCRA01 = GetPrjPeople1(Left(textCRA01 & "00000000", 8), "1")
      Select Case m_EditMode
         Case 1, 4:
            If Left(textCRA01, 1) <> "X" Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "必須輸入客戶編號"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCRA01_GotFocus
               GoTo EXITSUB
            End If
            If IsEmptyText(LabelCRA01) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "此客戶編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCRA01_GotFocus
            'Added by Lydia 2015/10/14
            Else
               QueryCustData
            End If
      End Select
   End If
EXITSUB:
End Sub

Private Sub textCRA02_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textCRA02
      CloseIme
   End If
End Sub

Private Sub textCRA02_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

'Added by Lydia 2015/10/14 +多筆查詢
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
  If txt1(Index).Text <> "" Then
    Select Case Index
        Case 0
            If ClsPDGetStaff(txt1(0).Text, strExc(1), strExc(2)) Then
               Label6.Caption = strExc(1)
            Else
               txt1(0).SetFocus
               Cancel = True
            End If
        Case 1, 2
            If Left(txt1(Index), 1) <> "X" Then
               MsgBox "客戶編號請輸入X編號!", vbCritical
               txt1(Index).SetFocus
               Cancel = True
            ElseIf txt1(1).Text <> "" And Index = 2 And txt1(1).Text > txt1(2).Text Then
               MsgBox "客戶編號起不可大於客戶編號止!", vbCritical
               txt1(Index).SetFocus
               Cancel = True
            End If
    End Select
  Else
    If Index = 0 Then Label6.Caption = ""
  End If
End Sub

Private Sub cmdOK_Click()
Dim Cancel As Boolean
Dim rsRead As New ADODB.Recordset

   For intI = 0 To 2
       txt1_Validate intI, Cancel
       If Cancel = True Then
          Exit Sub
       End If
   Next intI
   strExc(1) = "": strSql = ""
   If txt1(0).Text <> "" Then strExc(1) = strExc(1) & " AND CU13=" & CNULL(txt1(0).Text)
'Modified by Lydia 2020/01/30 共用語法
'   If txt1(1).Text <> "" Then strExc(1) = strExc(1) & " AND CRA01>=" & CNULL(txt1(1).Text)
'   If txt1(2).Text <> "" Then strExc(1) = strExc(1) & " AND CRA01<=" & CNULL(txt1(2).Text)
'
'   strSql = "SELECT CU01 CNO,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CNAME,CU13 ST01,ST02,CRA02 " & _
'            "FROM CUSTRECAMTLMT, CUSTOMER,STAFF WHERE CRA01=SUBSTR(CU01,1,6) AND CU02='0' AND CU13=ST01(+) " & strExc(1) & _
'            " ORDER BY 1"
   If txt1(1).Text <> "" Then strExc(1) = strExc(1) & " AND CU01||CU02>=" & CNULL(ChangeCustomerL(txt1(1).Text))
   If txt1(2).Text <> "" Then strExc(1) = strExc(1) & " AND CU01||CU02<=" & CNULL(ChangeCustomerL(txt1(2).Text))

   strSql = GetSql
   strSql = strSql & strExc(1) & " ORDER BY 1"
   'end 2020/01/30
   
   intI = 0
   Set rsRead = ClsLawReadRstMsg(intI, strSql)
   Set GRD1.Recordset = rsRead
   GRD1.FixedCols = 0
   SetGrd (IIf(rsRead.RecordCount = 0, 2, rsRead.RecordCount + 1))

End Sub
Private Sub SetGrd(Optional ByVal iR As Integer = 2)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer
   'Modified by Lydia 2020/01/30 共用語法
   'arrGridHeadText = Array("客戶編號", "客戶名稱", "ST01", "智權人員", "應收帳款上限")
   'arrGridHeadWidth = Array(1200, 2500, 0, 1000, 1500)
   arrGridHeadText = Array("客戶編號", "客戶名稱", "ST01", "智權人員", "應收帳款上限", "集團應收帳款上限", "CRA01")
   arrGridHeadWidth = Array(1000, 2500, 0, 1000, 1200, 1600, 0)
   'end 2020/01/30
   
   GRD1.Visible = False
   GRD1.Rows = iR
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iCol = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iCol
      GRD1.Text = arrGridHeadText(iCol)
      GRD1.ColWidth(iCol) = arrGridHeadWidth(iCol)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub
Private Sub GRD1_DblClick()
   'Modified by Lydia 2020/01/30
   'If InStr(GRD1.TextMatrix(GRD1.row, 0), textCRA01.Text) > 0 Then
   If InStr(GRD1.TextMatrix(GRD1.row, 0), txtCustNO.Text) > 0 Then
      Me.SSTab1.Tab = 0
   End If
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
         'Modified by Lydia 2020/01/30
         'textCRA01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         txtCustNO.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         QueryRecord
         GRD1.Visible = True
    End If
End If
End Sub

Private Sub QueryCustData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   'Modified by Lydia 2020/01/30 共用語法
   'strSql = "SELECT cu13,st02 FROM Customer,staff WHERE CU01='" & Left(Trim(IIf(textCRA01 <> m_CurrKEY(0), textCRA01.Text, m_CurrKEY(0))) & "00000000", 8) & "' and CU02='0' and cu13=st01(+)"
   strSql = GetSql("0")
   strSql = strSql & " and CU01='" & Left(Trim(IIf(txtCustNO <> "" And txtCustNO <> m_CurrKEY(0), txtCustNO.Text, m_CurrKEY(0))) & "00000000", 8) & "' "
   'end 2020/01/30
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      'Modified by Lydia 2020/01/30
      'If IsNull(rsTmp.Fields("cu13")) = False Then: LblCU13 = "" & rsTmp.Fields("cu13")
      If IsNull(rsTmp.Fields("st01")) = False Then: LblCU13 = "" & rsTmp.Fields("st01")
      If IsNull(rsTmp.Fields("st02")) = False Then: LblCU13Nm = "" & rsTmp.Fields("st02")
      'Added by Lydia 2020/01/30
      If IsNull(rsTmp.Fields("cra01")) = False Then
          textCRA01.Text = "" & rsTmp.Fields("cra01")
          textCRA02.Text = "" & rsTmp.Fields("cra02")
      End If
      txtCustNO = "" & rsTmp.Fields("cno")
      txtCustNO.Tag = txtCustNO.Text
      lblCustName = "" & rsTmp.Fields("cname")
      txtCU183.Text = "" & rsTmp.Fields("cu183")
      txtCU183.Tag = txtCU183.Tag
      'end 2020/01/30
   End If
   
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub
'end 2015/10/14

'Added by Lydia 2020/01/30
Private Sub txtCustNo_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If txtCustNO.Tag = txtCustNO.Text Then Exit Sub
   
   Cancel = False
   lblCustName = Empty
   If IsEmptyText(txtCustNO) = False Then
      If Len(txtCustNO) < 8 Then txtCustNO = Left(ChangeCustomerL(txtCustNO), 8)
      lblCustName = GetPrjPeople1(txtCustNO)
      Select Case m_EditMode
         Case 1, 4:
            If Left(txtCustNO, 1) <> "X" Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "必須輸入客戶編號"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               txtCustNo_GotFocus
               GoTo EXITSUB
            End If
            If IsEmptyText(lblCustName) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "此客戶編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               txtCustNo_GotFocus
            Else
               QueryCustData
               '新增時，母號可維護集團上限
               If m_EditMode = 1 Then
                   SetCtrlReadOnly False
               End If
            End If
      End Select
   End If
EXITSUB:
End Sub

Private Sub txtCustNo_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox txtCustNO
   End If
End Sub

Private Sub txtCustNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCU183_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox txtCU183
      CloseIme
   End If
End Sub

Private Sub txtCU183_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

'共用語法-應收帳款上限分開管制為個人"應收帳款上限"和"集團應收帳款上限"
Private Function GetSql(Optional ByVal pStatus As String = "1") As String
Dim tmpSQL As String
   
   '預設CRA01補足8碼
   tmpSQL = "SELECT CU01 CNO,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CNAME,CU13 ST01,ST02,CU183,CRA02,DECODE(CRA01,NULL,'',CRA01||'00') CRA01 "
   If pStatus = "2" Then tmpSQL = tmpSQL & ", CRA03,CRA04,CRA05,CRA06,CRA07,CRA08 "
   tmpSQL = tmpSQL & "FROM CUSTOMER,STAFF,CUSTRECAMTLMT WHERE CU02='0' AND CU13=ST01(+) AND SUBSTR(CU01,1,6)=CRA01(+) "
               
   If pStatus = "1" Then tmpSQL = tmpSQL & "AND (NVL(CU183,0) > 0 OR (SUBSTR(CU01,7,2)='00' AND NVL(CRA02,0)>0)) "  '抓有設定個人"應收帳款上限"或"集團應收帳款上限"
   GetSql = tmpSQL
   
End Function

