VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075007_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "下一程序資料維護"
   ClientHeight    =   5740
   ClientLeft      =   -2500
   ClientTop       =   2100
   ClientWidth     =   9330
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5740
   ScaleWidth      =   9330
   Begin VB.CommandButton CmdNext 
      Caption         =   "下一筆"
      Height          =   400
      Left            =   5520
      TabIndex        =   19
      Top             =   70
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.TextBox textNP02 
      Height          =   264
      Left            =   1380
      MaxLength       =   3
      TabIndex        =   0
      Top             =   516
      Width           =   732
   End
   Begin VB.TextBox textNP04 
      Height          =   264
      Left            =   3180
      MaxLength       =   1
      TabIndex        =   3
      Top             =   516
      Width           =   372
   End
   Begin VB.TextBox textNP05 
      Height          =   264
      Left            =   3540
      MaxLength       =   2
      TabIndex        =   4
      Top             =   516
      Width           =   732
   End
   Begin VB.TextBox textNP03_2 
      Height          =   264
      Left            =   2820
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   516
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8388
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7560
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6732
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1488
      Width           =   1932
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   5550
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1488
      Width           =   2052
   End
   Begin VB.TextBox textNP03 
      Height          =   264
      Left            =   2100
      MaxLength       =   6
      TabIndex        =   1
      Top             =   516
      Width           =   1092
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3720
      Left            =   120
      TabIndex        =   22
      Top             =   1920
      Width           =   9090
      _ExtentX        =   16051
      _ExtentY        =   6562
      _Version        =   393216
      Cols            =   15
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
      _Band(0).Cols   =   15
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1380
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1164
      Width           =   7725
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "13626;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1380
      TabIndex        =   20
      Top             =   840
      Width           =   7725
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13626;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "綠色：結案中"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7470
      TabIndex        =   18
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbldestroy 
      Caption         =   "北所銷卷"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8400
      TabIndex        =   17
      Top             =   1500
      Width           =   795
   End
   Begin VB.Label Label5 
      Caption         =   "紅色：與進度檔本所案號不同"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4560
      TabIndex        =   16
      Top             =   600
      Width           =   2565
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7650
      TabIndex        =   15
      Top             =   1500
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   9180
      Y1              =   1848
      Y2              =   1848
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   9180
      Y1              =   1872
      Y2              =   1872
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 ："
      Height          =   252
      Left            =   180
      TabIndex        =   14
      Top             =   516
      Width           =   1092
   End
   Begin VB.Label Label6 
      Caption         =   "申  請  人 ："
      Height          =   252
      Left            =   180
      TabIndex        =   13
      Top             =   1164
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 ："
      Height          =   252
      Left            =   180
      TabIndex        =   12
      Top             =   840
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家 ："
      Height          =   252
      Left            =   180
      TabIndex        =   11
      Top             =   1488
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "申請案號："
      Height          =   255
      Left            =   4500
      TabIndex        =   10
      Top             =   1485
      Width           =   975
   End
End
Attribute VB_Name = "frm075007_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/12 改成Form2.0 ; cmbTM05、textTM23、grdList改字型=新細明體-ExtB(MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim m_KeySel As Integer
Dim m_CurrSel As Integer

Dim m_NP02 As String
Dim m_NP03 As String
Dim m_NP04 As String
Dim m_NP05 As String

Dim m_Nation As String
'Added by Lydia 2019/03/21
Dim m_PrevForm As Form


'Added by Lydia 2019/03/21
Public Sub SetParent(ByRef pFM As Form)
   Set m_PrevForm = pFM
End Sub
'Added by Lydia 2019/03/21
Private Sub cmdNext_Click()
   '回到前畫面
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
        Me.Hide
        m_PrevForm.Show
        m_PrevForm.cmdState = 2
        Call m_PrevForm.PubShowNextData
   End If
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(textNP02) = True Then
      strTit = "檢核資料"
      strMsg = "請先輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If IsEmptyText(textNP03) = True Then
      strTit = "檢核資料"
      strMsg = "請先輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If textNP02 = "TF" And IsEmptyText(textNP03_2) = True Then
      strTit = "檢核資料"
      strMsg = "請先輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If IsDBExist() = False Then
      strTit = "檢核資料"
      strMsg = "基本檔資料不存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   cmdQuery.SetFocus
   DisplayNextForm
EXITSUB:
End Sub

Private Sub Form_Activate()
'Modify By Cheng 2002/11/04
''Add By Cheng 2002/01/09
'If Me.Visible And Me.textNP02.Text <> "" Then
If Me.Visible And Me.textNP02.Text <> "" And Me.textNP03.Text <> "" Then
   frm075007_1.RefreshList
End If
End Sub

Private Sub Form_Load()
   'textTM05.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   
   MoveFormToCenter Me
   'Add By Cheng 2002/04/25
   ClearField
    'Add By Cheng 2002/11/04
    Me.textNP02.Text = strSysKind
    'Add By Cheng 2002/12/11
   '92.4.25 modify by sonia
    'SendKeys "{Tab}"
   If strSysKind <> "" Then SendKeys "{Tab}"
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
   'Add By Sindy 2015/8/25
   Label5.Visible = False
   Label7.Visible = False
   '2015/8/25 END
   
   'Added by Lydia 2019/03/21 從電子公文稽核來
   If TypeName(m_PrevForm) = "frm060123" Then
       CmdNext.Visible = True
   End If
End Sub

'Modified by Lydia 2019/03/21
'Private Sub cmdQuery_Click()
Public Sub cmdQuery_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   '2005/7/4 ADD BY SONIA
   Label5.Visible = False
   Label7.Visible = False 'Add By Sindy 2015/8/25
   '2005/7/4 END
   m_NP02 = textNP02
   m_NP03 = textNP03
   If m_NP02 = "TF" Then: m_NP03 = m_NP03 & textNP03_2
   m_NP04 = textNP04
   If IsEmptyText(m_NP04) = True Then: m_NP04 = "0"
   m_NP05 = textNP05
   If IsEmptyText(m_NP05) = True Then: m_NP05 = "00"
   
   If IsEmptyText(textNP02) = True Or IsEmptyText(textNP03) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If textNP02 = "TF" Then
      If IsEmptyText(textNP03_2) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入本所案號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   ' 查詢資料
   'If QueryData() = False Then
   '   strTit = "查詢資料"
   '   strMsg = "無資料"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   'End If
   QueryData
   
EXITSUB:
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Added by Lydia 2019/03/21 回到前畫面
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
        m_PrevForm.Show
        m_PrevForm.cmdState = 0
        Call m_PrevForm.PubShowNextData
   End If
   'end 2019/03/21
   
   'Add By Cheng 2002/07/18
   Set frm075007_1 = Nothing
End Sub

Private Sub textNP02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub ClearField()
   cmbTM05.Clear
   textTM23 = Empty
   textTM10 = Empty
   textTM12 = Empty
   'Add By Cheng 2002/04/25
   Me.lblClose.Caption = Empty
   Me.lbldestroy.Caption = Empty  '2012/7/17 ADD BY SONIA
End Sub


' 本所案號的系統別
Private Sub textNP02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textNP02) = False Then
   
      'cancel by sonia 2013/11/29 考慮FCP程序可查詢FMP案改按查詢時檢查權限
      '' 使用者沒有權限
      'If IsUserHasRightOfSystem(strUserNum, textNP02) = False Then
      '   Cancel = True
      '   strTit = "檢核資料"
      '   strMsg = "您沒有使用該系統類別的權限"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textNP02_GotFocus
      '   GoTo EXITSUB
      'End If
      '2013/11/29 end
      
      Select Case textNP02
         Case "TF":
            textNP03_2.Visible = True
            textNP03_2.Locked = False
            textNP03_2.TabStop = True
            textNP03.MaxLength = 5
         Case Else:
            textNP03_2.Visible = False
            textNP03_2.Locked = True
            textNP03_2.TabStop = False
            textNP03.MaxLength = 6
      End Select
   Else
      textNP03_2.Visible = False
      textNP03_2.Locked = True
      textNP03_2.TabStop = False
      textNP03.MaxLength = 6
   End If
EXITSUB:
End Sub

' 讀取商標基本檔
Private Function QueryTradeMark(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryTradeMark = False
   strSql = "SELECT * FROM TRADEMARK " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryTradeMark = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem "中 : " & rsTmp.Fields("TM05")
      End If
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem "英 : " & rsTmp.Fields("TM06")
      End If
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem "日 : " & rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_Nation = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      'Add By Cheng 2002/04/25
      '顯示是否閉卷
      If rsTmp("TM29") = "Y" Then
         Me.lblClose.Caption = "已閉卷"
      End If
      '2012/7/17 ADD BY SONIA
      '顯示北所是否銷卷
      If IsNull(rsTmp("TM57")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      '2012/7/17 END
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取服務業務基本檔
Private Function QueryServicePractice(ByVal strSP01 As String, ByVal strSP02 As String, ByVal strSP03 As String, ByVal strSP04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryServicePractice = False
   strSql = "SELECT * FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & strSP01 & "' AND " & _
                  "SP02 = '" & strSP02 & "' AND " & _
                  "SP03 = '" & strSP03 & "' AND " & _
                  "SP04 = '" & strSP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryServicePractice = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem "中 : " & rsTmp.Fields("SP05")
      End If
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem "英 : " & rsTmp.Fields("SP06")
      End If
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem "日 : " & rsTmp.Fields("SP07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_Nation = rsTmp.Fields("SP09")
         textTM10 = GetNationName(rsTmp.Fields("SP09"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      'Add By Cheng 2002/04/25
      '顯示是否閉卷
      If rsTmp("SP15") = "Y" Then
         Me.lblClose.Caption = "已閉卷"
      End If
      '2012/7/17 ADD BY SONIA
      '顯示北所是否銷卷
      If IsNull(rsTmp("SP61")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      '2012/7/17 END
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取專利基本檔
Private Function QueryPatent(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryPatent = False
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & strPA01 & "' AND " & _
                  "PA02 = '" & strPA02 & "' AND " & _
                  "PA03 = '" & strPA03 & "' AND " & _
                  "PA04 = '" & strPA04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryPatent = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("PA05")) = False Then
         cmbTM05.AddItem "中 : " & rsTmp.Fields("PA05")
      End If
      If IsNull(rsTmp.Fields("PA06")) = False Then
         cmbTM05.AddItem "英 : " & rsTmp.Fields("PA06")
      End If
      If IsNull(rsTmp.Fields("PA07")) = False Then
         'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
         cmbTM05.AddItem "外 : " & rsTmp.Fields("PA07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("PA26")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("PA26"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("PA09")) = False Then
         m_Nation = rsTmp.Fields("PA09")
         textTM10 = GetNationName(rsTmp.Fields("PA09"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("PA11")) = False Then
         textTM12 = rsTmp.Fields("PA11")
      End If
      'Add By Cheng 2002/04/25
      '顯示是否閉卷
      If rsTmp("PA57") = "Y" Then
         Me.lblClose.Caption = "已閉卷"
      End If
      '2012/7/17 ADD BY SONIA
      '顯示北所是否銷卷
      If IsNull(rsTmp("PA108")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      '2012/7/17 END
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取法務基本檔
Private Function QueryLawCase(ByVal strLC01 As String, ByVal strLC02 As String, ByVal strLC03 As String, ByVal strLC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryLawCase = False
   strSql = "SELECT * FROM LAWCASE " & _
            "WHERE LC01 = '" & strLC01 & "' AND " & _
                  "LC02 = '" & strLC02 & "' AND " & _
                  "LC03 = '" & strLC03 & "' AND " & _
                  "LC04 = '" & strLC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryLawCase = True
      ' 案件名稱
      If IsNull(rsTmp.Fields("LC05")) = False Then
         cmbTM05.AddItem "中 : " & rsTmp.Fields("LC05")
      End If
      If IsNull(rsTmp.Fields("LC06")) = False Then
         cmbTM05.AddItem "英 : " & rsTmp.Fields("LC06")
      End If
      If IsNull(rsTmp.Fields("LC07")) = False Then
         cmbTM05.AddItem "日 : " & rsTmp.Fields("LC07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("LC11")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("LC11"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("LC15")) = False Then
         m_Nation = rsTmp.Fields("LC15")
         textTM10 = GetNationName(rsTmp.Fields("LC15"), 0)
      End If
      'Add By Cheng 2002/04/25
      '顯示是否閉卷
      If rsTmp("LC08") = "Y" Then
         Me.lblClose.Caption = "已閉卷"
      End If
      '2012/7/17 ADD BY SONIA
      '顯示北所是否銷卷
      If IsNull(rsTmp("LC34")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      '2012/7/17 END
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取顧問案件基本檔
Private Function QueryHireCase(ByVal strHC01 As String, ByVal strHC02 As String, ByVal strHC03 As String, ByVal strHC04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   QueryHireCase = False
   strSql = "SELECT * FROM HIRECASE " & _
            "WHERE HC01 = '" & strHC01 & "' AND " & _
                  "HC02 = '" & strHC02 & "' AND " & _
                  "HC03 = '" & strHC03 & "' AND " & _
                  "HC04 = '" & strHC04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryHireCase = True
     ' 案件名稱
      If IsNull(rsTmp.Fields("HC06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("HC06")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("HC05")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("HC05"), 0)
      End If
      'Add By Cheng 2002/04/25
      '顯示是否閉卷
      If rsTmp("HC09") = "Y" Then
         Me.lblClose.Caption = "已閉卷"
      End If
      '2012/7/17 ADD BY SONIA
      '顯示北所是否銷卷
      If IsNull(rsTmp("HC19")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      m_Nation = "000"
      textTM10 = GetNationName(m_Nation, 0)
      '2012/7/17 END
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取下一程序檔資料
Private Function QueryNextProgress(ByVal strNP02 As String, ByVal strNP03 As String, ByVal strNP04 As String, ByVal strNP05 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strCPSQL As String
   Dim rsCPTmp As ADODB.Recordset
   Dim strCP09 As String
   Dim nRow As Integer
   
   QueryNextProgress = False
   
   ' 設定查詢資料庫的SQL語法
   If m_Nation < "010" Then
      '2005/6/29 MODIFY BY SONIA 抓CP本所案號, 若與NP不同則顯示紅色提醒
      'strSQL = "SELECT NP01,NP06,NVL(C2.CPM03,NP07) AS NP07,NVL(NP08 - 19110000, NULL) AS NP08, NVL(NP09 - 19110000, NULL) AS NP09,NVL(S1.ST02,NP10) AS NP10,NP22,NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(S2.ST02,CP14) AS CP14,NVL(CP27 - 19110000, NULL) AS CP27,NP14,NP15 FROM NEXTPROGRESS, CASEPROGRESS C1, CASEPROPERTYMAP C2, STAFF S1, STAFF S2 " & _
      '         "WHERE NP02 = '" & strNP02 & "' AND " & _
      '               "NP03 = '" & strNP03 & "' AND " & _
      '               "NP04 = '" & strNP04 & "' AND " & _
      '               "NP05 = '" & strNP05 & "' AND " & _
      '               "NP01 = C1.CP09(+) AND " & _
      '               "NP10 = S1.ST01(+) AND " & _
      '               "CP14 = S2.ST01(+) AND " & _
      '               "NP02 = C2.CPM01(+) AND " & _
      '               "NP07 = C2.CPM02(+) " & _
      '         "ORDER BY CP05 DESC, NP01 DESC, NP08 DESC "
      'Modify By Sindy 2021/8/9 +, NVL(NP23 - 19110000, NULL) AS NP23
      strSql = "SELECT NP01,NP06,NVL(C2.CPM03,NP07) AS NP07,NVL(NP08 - 19110000, NULL) AS NP08, NVL(NP09 - 19110000, NULL) AS NP09" & _
               ",NVL(S1.ST02,NP10) AS NP10,NP22,NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(S2.ST02,CP14) AS CP14" & _
               ",NVL(CP27 - 19110000, NULL) AS CP27,NP14,NP15,CP01,CP02,CP03,CP04,NP24, NVL(NP23 - 19110000, NULL) AS NP23 " & _
               "FROM NEXTPROGRESS, CASEPROGRESS C1, CASEPROPERTYMAP C2, STAFF S1, STAFF S2 " & _
               "WHERE NP02 = '" & strNP02 & "' AND " & _
                     "NP03 = '" & strNP03 & "' AND " & _
                     "NP04 = '" & strNP04 & "' AND " & _
                     "NP05 = '" & strNP05 & "' AND " & _
                     "NP01 = C1.CP09(+) AND " & _
                     "NP10 = S1.ST01(+) AND " & _
                     "CP14 = S2.ST01(+) AND " & _
                     "NP02 = C2.CPM01(+) AND " & _
                     "NP07 = C2.CPM02(+) " & _
               "ORDER BY CP05 DESC, NP01 DESC, NP08 DESC "
      '2005/6/29 END
   Else
'2005/6/29 MODIFY BY SONIA 抓CP本所案號, 若與NP不同則顯示紅色提醒
'      strSQL = "SELECT NP01,NP06,NVL(C2.CPM04,NP07) AS NP07,NVL(NP08 - 19110000, NULL) AS NP08,NVL(NP09 - 19110000, NULL) AS NP09,NVL(S1.ST02,NP10) AS NP10,NP22,NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(S2.ST02,CP14) AS CP14,NVL(CP27 - 19110000, NULL) AS CP27,NP14,NP15 FROM NEXTPROGRESS, CASEPROGRESS C1, CASEPROPERTYMAP C2, STAFF S1, STAFF S2 " & _
'               "WHERE NP02 = '" & strNP02 & "' AND " & _
'                     "NP03 = '" & strNP03 & "' AND " & _
'                     "NP04 = '" & strNP04 & "' AND " & _
'                     "NP05 = '" & strNP05 & "' AND " & _
'                     "NP01 = C1.CP09(+) AND " & _
'                     "NP10 = S1.ST01(+) AND " & _
'                     "CP14 = S2.ST01(+) AND " & _
'                     "NP02 = C2.CPM01(+) AND " & _
'                     "NP07 = C2.CPM02(+) " & _
'               "ORDER BY CP05 DESC, NP01 DESC, NP08 DESC "
      'Modify By Sindy 2021/8/9 +, NVL(NP23 - 19110000, NULL) AS NP23
      strSql = "SELECT NP01,NP06,NVL(C2.CPM04,NP07) AS NP07,NVL(NP08 - 19110000, NULL) AS NP08,NVL(NP09 - 19110000, NULL) AS NP09" & _
               ",NVL(S1.ST02,NP10) AS NP10,NP22,NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(S2.ST02,CP14) AS CP14" & _
               ",NVL(CP27 - 19110000, NULL) AS CP27,NP14,NP15,CP01,CP02,CP03,CP04,NP24, NVL(NP23 - 19110000, NULL) AS NP23 " & _
               "FROM NEXTPROGRESS, CASEPROGRESS C1, CASEPROPERTYMAP C2, STAFF S1, STAFF S2 " & _
               "WHERE NP02 = '" & strNP02 & "' AND " & _
                     "NP03 = '" & strNP03 & "' AND " & _
                     "NP04 = '" & strNP04 & "' AND " & _
                     "NP05 = '" & strNP05 & "' AND " & _
                     "NP01 = C1.CP09(+) AND " & _
                     "NP10 = S1.ST01(+) AND " & _
                     "CP14 = S2.ST01(+) AND " & _
                     "NP02 = C2.CPM01(+) AND " & _
                     "NP07 = C2.CPM02(+) " & _
               "ORDER BY CP05 DESC, NP01 DESC, NP08 DESC "
'2005/6/29 END
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryNextProgress = True
      rsTmp.MoveFirst
      m_NP02 = strNP02
      m_NP03 = strNP03
      m_NP04 = strNP04
      m_NP05 = strNP05
      Do While rsTmp.EOF = False
         ' 新增一筆記錄
         grdList.Rows = grdList.Rows + 1
         nRow = grdList.Rows - 1
         ' 收文日
         If IsNull(rsTmp.Fields("CP05")) = False Then
            grdList.TextMatrix(nRow, 1) = rsTmp.Fields("CP05")
         End If
         ' 暫存總收文號
         strCP09 = Empty
         If IsNull(rsTmp.Fields("NP01")) = False Then
            strCP09 = rsTmp.Fields("NP01")
         End If
         ' 總收文號的欄位
         grdList.TextMatrix(nRow, 2) = strCP09
         
         ' 下一程序
         If IsNull(rsTmp.Fields("NP07")) = False Then
            grdList.TextMatrix(nRow, 3) = rsTmp.Fields("NP07")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("NP08")) = False Then
            grdList.TextMatrix(nRow, 4) = rsTmp.Fields("NP08")
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            grdList.TextMatrix(nRow, 5) = rsTmp.Fields("NP09")
         End If
         
         'Add By Sindy 2021/8/9
         ' 約定期限
         If IsNull(rsTmp.Fields("NP23")) = False Then
            grdList.TextMatrix(nRow, 6) = rsTmp.Fields("NP23")
         End If
         '2021/8/9 END
         
         ' 是否續辦欄位
         If IsNull(rsTmp.Fields("NP06")) = False Then
            grdList.TextMatrix(nRow, 7) = rsTmp.Fields("NP06")
         End If
         ' 智權人員
         If IsNull(rsTmp.Fields("NP10")) = False Then
            grdList.TextMatrix(nRow, 8) = rsTmp.Fields("NP10")
         End If
         ' 相關人
         If IsNull(rsTmp.Fields("NP14")) = False Then
            grdList.TextMatrix(nRow, 9) = rsTmp.Fields("NP14")
         End If
         ' 備註
         If IsNull(rsTmp.Fields("NP15")) = False Then
            grdList.TextMatrix(nRow, 10) = rsTmp.Fields("NP15")
         End If
         ' 承辦人
         If IsNull(rsTmp.Fields("CP14")) = False Then
            grdList.TextMatrix(nRow, 11) = rsTmp.Fields("CP14")
         End If
         ' 發文日
         If IsNull(rsTmp.Fields("CP27")) = False Then
            grdList.TextMatrix(nRow, 12) = rsTmp.Fields("CP27")
         End If
         ' 序號
         If IsNull(rsTmp.Fields("NP22")) = False Then
            grdList.TextMatrix(nRow, 13) = rsTmp.Fields("NP22")
         End If
         ' 相關案件性質  2011/10/5 ADD BY SONIA
         If IsNull(rsTmp.Fields("NP01")) = False Then
            grdList.TextMatrix(nRow, 3) = grdList.TextMatrix(nRow, 3) & PUB_GetNextCasePropertyName(grdList.TextMatrix(nRow, 2), grdList.TextMatrix(nRow, 13), "1")
         End If
         '2011/10/5 END
         If IsNull(rsTmp.Fields("CP09")) = False Then
            grdList.TextMatrix(nRow, 14) = 0
            '2005/6/29 ADD BY SONIA
            If rsTmp.Fields("CP01") <> strNP02 Then
               grdList.TextMatrix(nRow, 14) = 1
            End If
            If rsTmp.Fields("CP02") <> strNP03 Then
               grdList.TextMatrix(nRow, 14) = 1
            End If
            If rsTmp.Fields("CP03") <> strNP04 Then
               grdList.TextMatrix(nRow, 14) = 1
            End If
            If rsTmp.Fields("CP04") <> strNP05 Then
               grdList.TextMatrix(nRow, 14) = 1
            End If
         '2005/6/29 END
         Else
            grdList.TextMatrix(nRow, 14) = 1
         End If
         'Add By Sindy 2015/8/19
         ' NP24
         If IsNull(rsTmp.Fields("NP24")) = False Then
            grdList.TextMatrix(nRow, 15) = rsTmp.Fields("NP24")
         End If
         '2015/8/19 END
         ' 設定顯示的顏色
         Call SetColor(nRow) 'Modify By Sindy 2015/8/19
'         Select Case grdList.TextMatrix(nRow, 13)
'            Case "1":
'               For nCol = 1 To grdList.Cols - 1
'                  grdList.row = nRow
'                  grdList.col = nCol
'                  grdList.CellBackColor = &HFF&
'                  grdList.CellForeColor = &H80000008
'               Next nCol
'               '2005/7/4 ADD BY SONIA
'               Label5.Visible = True
'               '2005/7/4 END
'            Case Else:
'         End Select
         rsTmp.MoveNext
      Loop
      
      'Added by Lydia 2022/01/11 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
      If grdList.Rows > 1 Then
         grdList.FixedRows = 1
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add By Sindy 2015/8/19
Private Sub SetColor(nRow As Integer)
Dim nCol As Integer
   
   ' 設定顯示的顏色
   Select Case grdList.TextMatrix(nRow, 14)
      Case "1":
         For nCol = 1 To grdList.Cols - 1
            grdList.row = nRow
            grdList.col = nCol
            grdList.CellBackColor = &HFF& '紅色
            grdList.CellForeColor = &H80000008
         Next nCol
         Label5.Visible = True
      Case Else:
         'Add By Sindy 2015/8/19
         If Trim(grdList.TextMatrix(nRow, 7)) = "" And Len(Trim(grdList.TextMatrix(nRow, 15))) = 8 Then '長度8是結案單編號
            For nCol = 1 To grdList.Cols - 1
               grdList.row = nRow
               grdList.col = nCol
               grdList.CellBackColor = &H8000& '綠色
               grdList.CellForeColor = &H80000008
            Next nCol
            Label7.Visible = True
         End If
         '2015/8/19 END
   End Select
End Sub
'2015/8/19 END

' 讀取資料
Private Function QueryData() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strNP02 As String
   Dim strNP03 As String
   Dim strNP04 As String
   Dim strNP05 As String
   Dim bQuery As Boolean
   Dim bQueryNP As Boolean
   Dim StrSQLa As String            '2009/8/19 ADD BY SONIA
   Dim rsA As New ADODB.Recordset   '2009/8/19 ADD BY SONIA
   
   QueryData = False
   
   'textTM05 = Empty
   cmbTM05.Clear
   textTM10 = Empty
   textTM12 = Empty
   textTM23 = Empty
   
   ClearField
   InitialGridList
   
   ' 組成本所案號
   strNP02 = textNP02
   strNP03 = textNP03
   If strNP02 = "TF" Then: strNP03 = strNP03 & textNP03_2
   strNP04 = textNP04
   If IsEmptyText(strNP04) = True Then: strNP04 = "0"
   strNP05 = textNP05
   If IsEmptyText(strNP05) = True Then: strNP05 = "00"
   
   'Add By Cheng 2002/07/17
   m_Nation = ""
   bQuery = False
   
   'add by sonia 2013/11/29
   'FCP程序可查詢FMP案
   If CheckSR09(strUserNum, strNP02, "Y", , strNP02, strNP03, strNP04, strNP05) = False Then
      Set rsA = Nothing
      Exit Function
   End If
   '2013/11/29 end
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   'Mark by Lydia 2020/07/08 開放專利國外部程序可操作FMP非寰華案，但僅限進度檔及下一程序檔之維護 ! (不管P或FCP系統)
   'If FMP2open = True Then
   '   If PUB_FMPtoCheck(0, 1, Pub_strUserST05, strNP02, strNP03, strNP04, strNP05) = False Then
   '      Set rsA = Nothing
   '      Exit Function
   '   End If
   'End If
   'end 2020/07/08
   
   '2009/9/8 加註 by sonia T非台灣案非外商收文之案件不必寫程式控制,因為在系統類別外商人員即不可使用T案件
   '2009/8/19 add by sonia FCT無爭議程序之案件內商人員不可查詢(該案有內商承辦人者為FCT爭議案)
   If strNP02 = "FCT" And Mid(GetStaffDepartment(strUserNum), 1, 2) = "P2" Then
      'modify by sonia 2021/9/23 FCT-047943異議案林靖傑承辦,桂英無法操作
      'StrSQLa = "Select * From CASEPROGRESS,STAFF Where CP01='" & strNP02 & "' AND CP02='" & strNP03 & "' AND CP03='" & strNP04 & "' AND CP04='" & strNP05 & "' AND CP14=ST01(+) AND SUBSTR(ST03,1,2)='P2' "
      'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
      StrSQLa = "Select * From CASEPROGRESS,STAFF Where CP01='" & strNP02 & "' AND CP02='" & strNP03 & "' AND CP03='" & strNP04 & "' AND CP04='" & strNP05 & "' AND CP14=ST01(+) AND (SUBSTR(ST03,1,2)='P2' or (cp10 in (" & TMdebate & ") And Not (cp01 = 'FCT' And InStr(" & FCT_NotTMdebate & ", cp10) > 0)))"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic
      If rsA.RecordCount = 0 Then
         strMsg = "非FCT爭議案，您沒有使用該案號資料的權限"
         strTit = "查詢資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         Exit Function
      Else
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
   End If
   '2009/8/19 END
   
   ' 依本所案號讀取基本檔案
   Select Case strNP02
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         bQuery = QueryTradeMark(strNP02, strNP03, strNP04, strNP05)
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         bQuery = QueryPatent(strNP02, strNP03, strNP04, strNP05)
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/24 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         bQuery = QueryLawCase(strNP02, strNP03, strNP04, strNP05)
      ' 讀取顧問案件基本檔
      Case "LA":
         bQuery = QueryHireCase(strNP02, strNP03, strNP04, strNP05)
      ' 讀取服務業務基本檔
      Case Else:
         bQuery = QueryServicePractice(strNP02, strNP03, strNP04, strNP05)
   End Select
   
   ' 讀取下一程序檔
   bQueryNP = False
   bQueryNP = QueryNextProgress(strNP02, strNP03, strNP04, strNP05)
   
   If bQuery = False Then
      strTit = "查詢資料"
      strMsg = "該筆不存在於基本檔中"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      If bQueryNP = False Then
         strTit = "查詢資料"
         strMsg = "該筆案件無下一程序的資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   QueryData = bQueryNP
End Function

' 初始化列表
Public Sub InitialGridList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 16 '15

   grdList.ColWidth(0) = 300
   grdList.row = 0

   grdList.col = 0
   grdList.ColAlignment(0) = flexAlignCenterCenter
   grdList.col = 1
   grdList.Text = "收文日"
   grdList.ColWidth(1) = 800
   grdList.ColAlignment(1) = flexAlignCenterCenter
   grdList.col = 2
   grdList.Text = "總收文號"
   grdList.ColWidth(2) = 1000
   grdList.ColAlignment(2) = flexAlignCenterCenter
   grdList.col = 3
   grdList.Text = "下一程序"
   grdList.ColWidth(3) = 1250
   grdList.ColAlignment(3) = flexAlignLeftCenter
   grdList.col = 4
   grdList.Text = "本所期限"
   grdList.ColWidth(4) = 800
   grdList.ColAlignment(4) = flexAlignCenterCenter
   grdList.col = 5
   grdList.Text = "法定期限"
   grdList.ColWidth(5) = 800
   grdList.ColAlignment(5) = flexAlignCenterCenter
   
   'Add By Sindy 2021/8/9
   grdList.col = 6
   grdList.Text = "約定期限"
   If Left(Pub_StrUserSt03, 2) = "F2" Or Pub_StrUserSt03 = "M51" Then '外專部門加看約定期限
      grdList.ColWidth(6) = 800
   Else
      grdList.ColWidth(6) = 0
   End If
   grdList.ColAlignment(6) = flexAlignCenterCenter
   '2021/8/9 END
   
   grdList.col = 7
   grdList.Text = "續辦"
   grdList.ColWidth(7) = 450
   grdList.ColAlignment(7) = flexAlignCenterCenter
   grdList.col = 8
   grdList.Text = "智權人員"
   grdList.ColWidth(8) = 800
   grdList.ColAlignment(8) = flexAlignLeftCenter
   grdList.col = 9
   grdList.Text = "相關人"
   grdList.ColWidth(9) = 1250
   grdList.ColAlignment(9) = flexAlignLeftCenter
   grdList.col = 10
   grdList.Text = "備　註"
   grdList.ColWidth(10) = 3000
   grdList.ColAlignment(10) = flexAlignLeftCenter
   grdList.col = 11
   grdList.Text = "承辦人"
   'grdList.ColWidth(11) = 800
   grdList.ColWidth(11) = 0
   grdList.ColAlignment(11) = flexAlignLeftCenter
   grdList.col = 12
   grdList.Text = "發文日"
   'grdList.ColWidth(12) = 800
   grdList.ColWidth(12) = 0
   grdList.ColAlignment(12) = flexAlignCenterCenter
   grdList.col = 13
   grdList.Text = "序號"
   'grdList.ColWidth(13) = 1000
   grdList.ColWidth(13) = 0
   grdList.ColAlignment(13) = flexAlignLeftCenter
   grdList.col = 14
   grdList.Text = "案件進度檔是否存在"
   grdList.ColWidth(14) = 0
   grdList.ColAlignment(14) = flexAlignLeftCenter
   'Add By Sindy 2015/8/19
   grdList.col = 15
   grdList.Text = "NP24"
   grdList.ColWidth(15) = 0
   grdList.ColAlignment(15) = flexAlignLeftCenter
   '2015/8/19 END
End Sub

Private Sub grdList_Click()
   ' 案件性質必須為延期的才可以選取
   If grdList.row > 0 Then
      grdList.col = 0
      If grdList.Text = "V" Then
         grdList.Text = Empty
      Else
         grdList.Text = "V"
         cmdOK.SetFocus
      End If
   End If
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
   ' 案件性質必須為延期的才可以選取
   If KeyCode = vbKeySpace Then
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
            grdList.Text = "V"
         End If
      End If
   End If
End Sub

Private Sub grdList_SelChange()
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         'Select Case grdList.TextMatrix(grdList.row, 12)
         Select Case grdList.TextMatrix(grdList.row, 14)
            Case "1":
               For nCol = 1 To grdList.Cols - 1
                  grdList.col = nCol
                  If grdList.CellBackColor <> &HFF& Then: grdList.CellBackColor = &HFF& '紅色
                  If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
               Next nCol
            Case Else:
               'Add By Sindy 2015/8/19
               If Trim(grdList.TextMatrix(grdList.row, 7)) = "" And Trim(grdList.TextMatrix(grdList.row, 15)) <> "" Then
                  For nCol = 1 To grdList.Cols - 1
                     grdList.col = nCol
                     If grdList.CellBackColor <> &H8000& Then: grdList.CellBackColor = &H8000& '綠色
                     If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                  Next nCol
               Else
               '2015/8/19 END
                  For nCol = 1 To grdList.Cols - 1
                     grdList.col = nCol
                     If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
                     If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
                  Next nCol
               End If
         End Select
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

Private Sub DisplayNextForm()
   Dim nRow As Integer
   Dim strNP02 As String
   Dim strNP03 As String
   Dim strNP04 As String
   Dim strNP05 As String
   Dim bFind As Boolean
   bFind = False
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 0) = "V" Then
         bFind = True
         Exit For
      End If
   Next nRow
   If bFind = True Then
      frm075007_2.SetData 0, m_NP02, True
      frm075007_2.SetData 1, m_NP03, False
      frm075007_2.SetData 2, m_NP04, False
      frm075007_2.SetData 3, m_NP05, False
   Else
      ' 組成本所案號
      strNP02 = textNP02
      strNP03 = textNP03
      If strNP02 = "TF" Then: strNP03 = strNP03 & textNP03_2
      strNP04 = textNP04
      If IsEmptyText(strNP04) = True Then: strNP04 = "0"
      strNP05 = textNP05
      If IsEmptyText(strNP05) = True Then: strNP05 = "00"
      frm075007_2.SetData 0, strNP02, True
      frm075007_2.SetData 1, strNP03, False
      frm075007_2.SetData 2, strNP04, False
      frm075007_2.SetData 3, strNP05, False
   End If
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 0) = "V" Then
         frm075007_2.SetData 4, grdList.TextMatrix(nRow, 13), False
      End If
   Next nRow
   frm075007_2.Show
   frm075007_2.QueryDB
   Me.Hide
End Sub

' 刪除該筆記錄
Public Sub DelRecord(ByVal strData As String)
   Dim nRow As Integer
   Dim bFind As Boolean
   bFind = False
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 11) = strData Then
         bFind = True
         Exit For
      End If
   Next nRow

   If bFind = True Then
      If grdList.Rows <= 2 Then
         InitialGridList
      Else
         grdList.RemoveItem nRow
      End If
   End If

End Sub

' 更新該筆記錄
Public Sub ModRecord(ByVal strData As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim rsCPTmp As ADODB.Recordset
   Dim strCPSQL As String
   Dim nRow As Integer
   Dim nCol As Integer
   Dim bFind As Boolean
   Dim strNP22 As String
   Dim strCP09 As String
   
   strNP22 = strData
   
   bFind = False
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 13) = strData Then
         bFind = True
         Exit For
      End If
   Next nRow
   
   ' 設定查詢資料庫的SQL語法
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & m_NP02 & "' AND " & _
                  "NP03 = '" & m_NP03 & "' AND " & _
                  "NP04 = '" & m_NP04 & "' AND " & _
                  "NP05 = '" & m_NP05 & "' AND " & _
                  "NP22 = " & strNP22 & " "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 新增一筆
      If bFind = False Then
         grdList.Rows = grdList.Rows + 1
         nRow = grdList.Rows - 1
      End If
      ' 清空
      For nCol = 1 To grdList.Cols - 1
         grdList.TextMatrix(nRow, nCol) = Empty
      Next nCol
   
      ' 暫存總收文號
      strCP09 = Empty
      If IsNull(rsTmp.Fields("NP01")) = False Then
         strCP09 = rsTmp.Fields("NP01")
      End If
      
      ' 總收文號的欄位
      grdList.TextMatrix(nRow, 2) = strCP09
      ' 檢查案件進度檔
      If IsEmptyText(grdList.TextMatrix(nRow, 2)) = False Then
         strCPSQL = "SELECT * FROM CASEPROGRESS " & _
                    "WHERE CP01 = '" & m_NP02 & "' AND " & _
                          "CP02 = '" & m_NP03 & "' AND " & _
                          "CP03 = '" & m_NP04 & "' AND " & _
                          "CP04 = '" & m_NP05 & "' AND " & _
                          "CP09 = '" & strCP09 & "' "
         Set rsCPTmp = New ADODB.Recordset
         rsCPTmp.CursorLocation = adUseClient
         rsCPTmp.Open strCPSQL, cnnConnection, adOpenStatic, adLockReadOnly
         If rsCPTmp.RecordCount > 0 Then
            rsCPTmp.MoveFirst
            ' 收文日
            If IsNull(rsCPTmp.Fields("CP05")) = False Then
               grdList.TextMatrix(nRow, 1) = TAIWANDATE(rsCPTmp.Fields("CP05"))
            End If
            ' 承辦人
            If IsNull(rsCPTmp.Fields("CP14")) = False Then
               grdList.TextMatrix(nRow, 11) = GetStaffName(rsCPTmp.Fields("CP14"), True)
            End If
            ' 發文日
            If IsNull(rsCPTmp.Fields("CP27")) = False Then
               grdList.TextMatrix(nRow, 12) = TAIWANDATE(rsCPTmp.Fields("CP27"))
            End If
         Else
            ' 記錄該筆下一程序資料為案件進度檔不存在
            grdList.TextMatrix(nRow, 14) = 1
         End If
      Else
         ' 記錄該筆下一程序資料為案件進度檔不存在
         grdList.TextMatrix(nRow, 14) = 1
      End If
      
      ' 下一程序
      If IsNull(rsTmp.Fields("NP07")) = False Then
         ' 依國別取案件性質表的中文或大陸名稱
         If m_Nation < "010" Then
            grdList.TextMatrix(nRow, 3) = GetCaseTypeName(m_NP02, rsTmp.Fields("NP07"), 0)
         Else
            grdList.TextMatrix(nRow, 3) = GetCaseTypeName(m_NP02, rsTmp.Fields("NP07"), 1)
         End If
      End If
      ' 本所期限
      If IsNull(rsTmp.Fields("NP08")) = False Then
         grdList.TextMatrix(nRow, 4) = TAIWANDATE(rsTmp.Fields("NP08"))
      End If
      ' 法定期限
      If IsNull(rsTmp.Fields("NP09")) = False Then
         grdList.TextMatrix(nRow, 5) = TAIWANDATE(rsTmp.Fields("NP09"))
      End If
      
      'Add By Sindy 2021/8/9
      ' 約定期限
      If IsNull(rsTmp.Fields("NP23")) = False Then
         grdList.TextMatrix(nRow, 6) = TAIWANDATE(rsTmp.Fields("NP23"))
      End If
      '2021/8/9 END
      
      ' 是否續辦欄位
      If IsNull(rsTmp.Fields("NP06")) = False Then
         grdList.TextMatrix(nRow, 7) = rsTmp.Fields("NP06")
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("NP10")) = False Then
         grdList.TextMatrix(nRow, 8) = GetStaffName(rsTmp.Fields("NP10"), True)
      End If
      ' 相關人
      If IsNull(rsTmp.Fields("NP14")) = False Then
         grdList.TextMatrix(nRow, 9) = rsTmp.Fields("NP14")
      End If
      ' 備註
      If IsNull(rsTmp.Fields("NP15")) = False Then
         grdList.TextMatrix(nRow, 10) = rsTmp.Fields("NP15")
      End If
      ' 序號
      If IsNull(rsTmp.Fields("NP22")) = False Then
         grdList.TextMatrix(nRow, 13) = rsTmp.Fields("NP22")
      End If
      
      ' 設定該筆資料為選取的狀態
      grdList.TextMatrix(nRow, 0) = "V"
      
      ' 設定顯示的顏色
      Call SetColor(nRow) 'Modify By Sindy 2015/8/19
'      Select Case grdList.TextMatrix(nRow, 13)
'         Case "1":
'            For nCol = 1 To grdList.Cols - 1
'               grdList.row = nRow
'               grdList.col = nCol
'               grdList.CellBackColor = &HFF&
'               grdList.CellForeColor = &H80000008
'            Next nCol
'         Case Else:
'      End Select
      
      rsTmp.MoveNext
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub textNP02_GotFocus()
   InverseTextBox textNP02
   CloseIme
End Sub

Private Sub textNP03_2_GotFocus()
   InverseTextBox textNP03_2
End Sub

Private Sub textNP03_GotFocus()
   InverseTextBox textNP03
   CloseIme
End Sub

Private Sub textNP04_GotFocus()
   InverseTextBox textNP04
End Sub

Private Sub textNP04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textNP05_GotFocus()
   InverseTextBox textNP05
End Sub

Public Sub ClearRemark()
   Dim nRow As Integer
   For nRow = 0 To grdList.Rows - 1
      grdList.TextMatrix(nRow, 0) = Empty
   Next nRow
End Sub

Public Sub RefreshList()
   QueryData
End Sub

Private Function IsDBExist() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
  
   IsDBExist = False
   Select Case m_NP02
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         strSql = "SELECT * FROM TRADEMARK " & _
                  "WHERE TM01 = '" & m_NP02 & "' AND " & _
                        "TM02 = '" & m_NP03 & "' AND " & _
                        "TM03 = '" & m_NP04 & "' AND " & _
                        "TM04 = '" & m_NP05 & "' "
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         strSql = "SELECT * FROM PATENT " & _
                  "WHERE PA01 = '" & m_NP02 & "' AND " & _
                        "PA02 = '" & m_NP03 & "' AND " & _
                        "PA03 = '" & m_NP04 & "' AND " & _
                        "PA04 = '" & m_NP05 & "' "
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/24 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         strSql = "SELECT * FROM LAWCASE " & _
                  "WHERE LC01 = '" & m_NP02 & "' AND " & _
                        "LC02 = '" & m_NP03 & "' AND " & _
                        "LC03 = '" & m_NP04 & "' AND " & _
                        "LC04 = '" & m_NP05 & "' "
      ' 讀取顧問案件基本檔
      Case "LA":
         strSql = "SELECT * FROM HIRECASE " & _
                  "WHERE HC01 = '" & m_NP02 & "' AND " & _
                        "HC02 = '" & m_NP03 & "' AND " & _
                        "HC03 = '" & m_NP04 & "' AND " & _
                        "HC04 = '" & m_NP05 & "' "
      ' 讀取服務業務基本檔
      Case Else:
         strSql = "SELECT * FROM SERVICEPRACTICE " & _
                  "WHERE SP01 = '" & m_NP02 & "' AND " & _
                        "SP02 = '" & m_NP03 & "' AND " & _
                        "SP03 = '" & m_NP04 & "' AND " & _
                        "SP04 = '" & m_NP05 & "' "
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsDBExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

