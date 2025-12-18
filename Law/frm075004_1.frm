VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075004_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件進度檔資料維護"
   ClientHeight    =   5420
   ClientLeft      =   160
   ClientTop       =   1010
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5420
   ScaleWidth      =   9310
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3720
      Left            =   150
      TabIndex        =   24
      Top             =   1650
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
   Begin VB.CommandButton cmdQueryAll 
      Caption         =   "含來函查詢(&L)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6216
      TabIndex        =   6
      Top             =   10
      Width           =   1300
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   5568
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2172
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   1284
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1545
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "不含來函查詢(&F)"
      Height          =   400
      Left            =   4692
      TabIndex        =   5
      Top             =   10
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7536
      TabIndex        =   7
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8364
      TabIndex        =   8
      Top             =   10
      Width           =   800
   End
   Begin VB.TextBox textCP02_2 
      Height          =   300
      Left            =   2724
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   192
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textCP04 
      Height          =   300
      Left            =   3384
      MaxLength       =   2
      TabIndex        =   4
      Top             =   192
      Width           =   732
   End
   Begin VB.TextBox textCP03 
      Height          =   300
      Left            =   3084
      MaxLength       =   1
      TabIndex        =   3
      Top             =   192
      Width           =   372
   End
   Begin VB.TextBox textCP01 
      Height          =   300
      Left            =   1284
      MaxLength       =   3
      TabIndex        =   0
      Top             =   192
      Width           =   732
   End
   Begin VB.TextBox textCP02 
      Height          =   300
      Left            =   2004
      MaxLength       =   6
      TabIndex        =   1
      Top             =   192
      Width           =   1092
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1284
      TabIndex        =   23
      Top             =   528
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
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1284
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   864
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
   Begin VB.Label lblCaseMap2 
      AutoSize        =   -1  'True
      Caption         =   "lblCaseMap2"
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
      Height          =   180
      Left            =   3600
      TabIndex        =   21
      Top             =   1410
      Width           =   1770
   End
   Begin VB.Label lblCMboth 
      Caption         =   "lblCMboth"
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
      Height          =   180
      Left            =   7740
      TabIndex        =   20
      Top             =   1365
      Width           =   945
   End
   Begin VB.Label lblCaseMap 
      AutoSize        =   -1  'True
      Caption         =   "lblCaseMap"
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
      Height          =   180
      Left            =   3600
      TabIndex        =   18
      Top             =   1230
      Width           =   1005
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
      Left            =   8520
      TabIndex        =   17
      Top             =   1155
      Width           =   800
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
      Left            =   7740
      TabIndex        =   16
      Top             =   1155
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   150
      X2              =   9210
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label4 
      Caption         =   "申請案號："
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家 ："
      Height          =   252
      Left            =   192
      TabIndex        =   14
      Top             =   1200
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 ："
      Height          =   252
      Left            =   192
      TabIndex        =   13
      Top             =   528
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "申   請  人："
      Height          =   252
      Left            =   192
      TabIndex        =   12
      Top             =   864
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 ："
      Height          =   252
      Left            =   192
      TabIndex        =   11
      Top             =   192
      Width           =   1092
   End
   Begin VB.Label lblDivision 
      AutoSize        =   -1  'True
      Caption         =   "lblDivision"
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
      Height          =   180
      Left            =   2880
      TabIndex        =   19
      Top             =   1230
      Width           =   720
   End
End
Attribute VB_Name = "frm075004_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/08 改成Form2.0 ; grdList改字型=新細明體-ExtB、cmbTM05、textTM23
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim m_KeySel As Integer
Dim m_CurrSel As Integer

Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
' 國家
Dim m_Nation As String
'
Dim m_QueryAll As Boolean
Dim m_SP18 As String 'Added by Lydia 2016/04/12
Dim pYYMM As String 'Added by Lydia 2022/09/06 限制收文年月
   
Private Sub ClearField()
   cmbTM05.Clear
   textTM23 = Empty
   textTM10 = Empty
   textTM12 = Empty
   'Add By Cheng 2002/04/25
   Me.lblClose.Caption = Empty
   Me.lbldestroy.Caption = Empty  '2007/8/29 ADD BY SONIA
   'Added by Lydia 2015/11/03
   Me.lblCaseMap.Caption = Empty
   Me.lblCaseMap2.Caption = Empty 'Added by Lydia 2019/11/28
   Me.lblDivision.Caption = Empty
   'Added by Lydia 2016/06/14
   Me.lblCMboth.Caption = Empty
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(textCP01) = True Then
      strTit = "檢核資料"
      strMsg = "請先輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If IsEmptyText(textCP02) = True Then
      strTit = "檢核資料"
      strMsg = "請先輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If textCP01 = "TF" And IsEmptyText(textCP02_2) = True Then
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
   cmdQueryAll.SetFocus
   DisplayNextForm
EXITSUB:
End Sub

Private Sub cmdQueryAll_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   m_CP01 = Empty
   m_CP02 = Empty
   m_CP03 = Empty
   m_CP04 = Empty
   m_SP18 = Empty 'Added by Lydia 2016/04/12
   
   If IsEmptyText(textCP01) = True Or IsEmptyText(textCP02) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If textCP01 = "TF" Then
      If IsEmptyText(textCP02_2) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入本所案號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   ' 查詢資料
   m_QueryAll = True
   'If QueryData(True) = False Then
   '   strTit = "查詢資料"
   '   strMsg = "無資料"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   'End If
   QueryData True
EXITSUB:
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
    Me.textCP01.Text = strSysKind
   'Add By Cheng 2002/12/11
   '92.4.25 modify by sonia
   'SendKeys "{Tab}"
   If strSysKind <> "" Then SendKeys "{Tab}"
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub cmdQuery_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   m_CP01 = Empty
   m_CP02 = Empty
   m_CP03 = Empty
   m_CP04 = Empty
   m_SP18 = Empty 'Added by Lydia 2016/04/12
   
   If IsEmptyText(textCP01) = True Or IsEmptyText(textCP02) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If textCP01 = "TF" Then
      If IsEmptyText(textCP02_2) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入本所案號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   ' 查詢資料
   m_QueryAll = False
   'If QueryData(False) = False Then
   '   strTit = "查詢資料"
   '   strMsg = "無資料"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   'End If
   QueryData False
   
EXITSUB:
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm075004_1 = Nothing
End Sub

Private Sub textCP01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號的系統別
Private Sub textCP01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse

   Cancel = False
   If IsEmptyText(textCP01) = False Then
      
      'Removed by Morgan 2013/4/23 考慮FCP程序可維護FMP案改按查詢時檢查權限
      '' 使用者沒有權限
      'If IsUserHasRightOfSystem(strUserNum, textCP01) = False Then
      '   Cancel = True
      '   strTit = "檢核資料"
      '   strMsg = "您沒有使用該系統類別的權限"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCP01_GotFocus
      '   GoTo EXITSUB
      'End If
      'end 2013/4/23
      
      Select Case textCP01
         Case "TF":
            textCP02_2.Visible = True
            textCP02_2.Locked = False
            textCP02_2.TabStop = True
            textCP02.MaxLength = 5
         Case Else:
            textCP02_2.Visible = False
            textCP02_2.Locked = True
            textCP02_2.TabStop = False
            textCP02.MaxLength = 6
      End Select
   Else
      textCP02_2.Visible = False
      textCP02_2.Locked = True
      textCP02_2.TabStop = False
      textCP02.MaxLength = 6
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
      'Add By Cheng 2002/07/17
      m_Nation = ""
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
      '2007/8/29 ADD BY SONIA
      '顯示北所是否銷卷
      If IsNull(rsTmp("TM57")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      '2007/8/29 END
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
      'Add By Cheng 2002/07/17
      m_Nation = ""
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
      '2007/8/29 ADD BY SONIA
      '顯示北所是否銷卷
      If IsNull(rsTmp("SP61")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      '2007/8/29 END
      m_SP18 = "" & Trim(rsTmp("SP18"))
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
      'Add By Cheng 2002/07/17
      m_Nation = ""
      ' 申請國家
      If IsNull(rsTmp.Fields("PA09")) = False Then
         m_Nation = rsTmp.Fields("PA09")
         textTM10 = GetNationName(rsTmp.Fields("PA09"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("PA11")) = False Then
         textTM12 = rsTmp.Fields("PA11")
      End If
      '顯示是否閉卷
      If rsTmp("PA57") = "Y" Then
         Me.lblClose.Caption = "已閉卷"
      End If
      '2007/8/29 ADD BY SONIA
      '顯示北所是否銷卷
      If IsNull(rsTmp("PA108")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      '2007/8/29 END
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
      'Add By Cheng 2002/07/17
      m_Nation = ""
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
      '2007/8/29 ADD BY SONIA
      '顯示北所是否銷卷
      If IsNull(rsTmp("LC34")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      '2007/8/29 END
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
      '2007/8/29 ADD BY SONIA
      '顯示北所是否銷卷
      If IsNull(rsTmp("HC19")) = False Then
         Me.lbldestroy.Caption = "北所銷卷"
      End If
      '2007/8/29 END
      '92.9.26 ADD BY SONIA
      m_Nation = "000"
      textTM10 = GetNationName(m_Nation, 0)
      '92.9.26 END
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取案件進度檔資料
'Modify by Amy 2025/01/07 +m_CP09/intRow ,避免修改完回來後(Run ModRecord),程式有未改到故統一跑此函數
Private Function QueryCaseProgress(ByVal strCP01 As String, ByVal strCP02 As String, ByVal strCP03 As String, ByVal strCP04 As String, ByVal bAll As Boolean, _
  Optional ByVal m_CP09 As String = "", Optional ByVal intRow As Integer = 0) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strCP09 As String
   Dim nRow As Integer
   Dim nCol As Integer
   'Add By Cheng 2002/07/09
   Dim StrSQLa As String
   Dim strWhereSql As String 'Added by Lydia 2022/09/06
   
   QueryCaseProgress = False
   'Add By Cheng 2002/07/09
   StrSQLa = "DECODE(SK03,0,NVL(F1.FA04,DECODE(F1.FA05,NULL,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)),DECODE(F1.FA05,NULL,NVL(F1.FA04,F1.FA06),F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)) "
      
   'Added by Lydia 2022/09/06 限制收文年月
   If pYYMM <> "" Then
       strWhereSql = strWhereSql & " AND CP05>=" & Val(pYYMM & "01") + 19110000 & " AND CP05<=" & Val(pYYMM & "31") + 19110000
   End If
   'end 2022/09/06
   'Add by Amy 2025/01/07 +cp09 (避免修改完回來後,程式有改到統一跑此函數)
   If m_CP09 <> MsgText(601) Then
      strWhereSql = strWhereSql & " And CP09='" & m_CP09 & "' "
   End If
   
   ' 設定查詢資料庫的SQL語法
   'If bAll = True Then
   '   strSQL = "SELECT * FROM CaseProgress " & _
   '            "WHERE CP01 = '" & strCP01 & "' AND " & _
   '                  "CP02 = '" & strCP02 & "' AND " & _
   '                  "CP03 = '" & strCP03 & "' AND " & _
   '                  "CP04 = '" & strCP04 & "' " & _
   '            "ORDER BY CP05, CP09 "
   'Else
   '   strSQL = "SELECT * FROM CaseProgress " & _
   '            "WHERE CP01 = '" & strCP01 & "' AND " & _
   '                  "CP02 = '" & strCP02 & "' AND " & _
   '                  "CP03 = '" & strCP03 & "' AND " & _
   '                  "CP04 = '" & strCP04 & "' AND " & _
   '                  "SUBSTR(CP09,1,1) <> 'C' " & _
   '            "ORDER BY CP05, CP09 "
   'End If
   If bAll = True Then
        '若申請國家為台灣
      If m_Nation < "010" Then
         'Modify By Cheng 2002/07/09
'         strSQL = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10) AS CP10,NVL(S1.ST02,CP13) AS CP13,NVL(S2.ST02,CP14) AS CP14,CP24,NVL(CP27 - 19110000, NULL) AS CP27,CP43,NVL(F1.FA05,NVL(F1.FA04,NVL(F1.FA06,CP44))) AS CP44,NVL(CP57 - 19110000, NULL) AS CP57,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1 " & _
'                  "WHERE CP01 = '" & strcp01 & "' AND " & _
'                        "CP02 = '" & strcp02 & "' AND " & _
'                        "CP03 = '" & strcp03 & "' AND " & _
'                        "CP04 = '" & strcp04 & "' AND " & _
'                        "CP13 = S1.ST01(+) AND " & _
'                        "CP14 = S2.ST01(+) AND " & _
'                        "SUBSTR(CP44,1,8) = FA01(+) AND " & _
'                        "SUBSTR(CP44,9,1) = FA02(+) AND " & _
'                        "CP01 = CPM01(+) AND " & _
'                        "CP10 = CPM02(+) " & _
'                  "ORDER BY CP05 DESC, CP09 DESC"
        'Modify By Cheng 2003/01/27
        '加相關人及對造號數
'         strSQL = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10) AS CP10,NVL(S1.ST02,CP13) AS CP13,NVL(S2.ST02,CP14) AS CP14,CP24,NVL(CP27 - 19110000, NULL) AS CP27,CP43," & strSQLA & " AS CP44,NVL(CP57 - 19110000, NULL) AS CP57,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1,SYSTEMKIND "
         'Modified by Lydia 2015/02/11 專利權消滅(1604)的案件性質改顯示CP64
         'Modified by Lydia 2022/09/06 +strWhereSql
         strSql = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,decode(cp01||cp10,'P1604',decode(sign(instr(cp64,'專利權消滅')),1,substr(cp64,instr(cp64,'專利權消滅'),15),NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10)),'CFP1604',decode(sign(instr(cp64,'專利權消滅')),1,substr(cp64,instr(cp64,'專利權消滅'),15),NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10)),NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10)) AS CP10," & _
                  "NVL(S1.ST02,CP13) AS CP13,NVL(S2.ST02,CP14) AS CP14,CP24,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))),CP36,NVL(CP27 - 19110000, NULL) AS CP27,CP43," & StrSQLa & " AS CP44,NVL(CP57 - 19110000, NULL) AS CP57," & _
                  "NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 " & _
                    " FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1,SYSTEMKIND,CUSTOMER " & _
                    " WHERE CP01 = '" & strCP01 & "' AND " & _
                        "CP02 = '" & strCP02 & "' AND " & _
                        "CP03 = '" & strCP03 & "' AND " & _
                        "CP04 = '" & strCP04 & "' AND " & _
                        "CP13 = S1.ST01(+) AND " & _
                        "CP14 = S2.ST01(+) AND " & _
                        "SUBSTR(CP44,1,8) = FA01(+) AND " & _
                        "SUBSTR(CP44,9,1) = FA02(+) AND " & _
                        "CP01 = CPM01(+) AND " & _
                        "CP10 = CPM02(+) AND CP01=SK01(+) " & _
                        "AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) " & strWhereSql & _
                  "ORDER BY CP05 DESC, CP66 DESC, CP67 DESC, CP09 DESC"   '2009/2/20 MODIFY BY SONIA 排序原為收文日+收文號,改收文日+CREATE DATE+CREATE TIME+收文號
        '若申請國家非台灣
      Else
         'Modify By Cheng 2002/07/09
'         strSQL = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10) AS CP10,NVL(S1.ST02,CP13) AS CP13,NVL(S2.ST02,CP14) AS CP14,CP24,NVL(CP27 - 19110000, NULL) AS CP27,CP43,NVL(F1.FA05,NVL(F1.FA04,NVL(F1.FA06,CP44))) AS CP44,NVL(CP57 - 19110000, NULL) AS CP57,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1 " & _
'                  "WHERE CP01 = '" & strcp01 & "' AND " & _
'                        "CP02 = '" & strcp02 & "' AND " & _
'                        "CP03 = '" & strcp03 & "' AND " & _
'                        "CP04 = '" & strcp04 & "' AND " & _
'                        "CP13 = S1.ST01(+) AND " & _
'                        "CP14 = S2.ST01(+) AND " & _
'                        "SUBSTR(CP44,1,8) = FA01(+) AND " & _
'                        "SUBSTR(CP44,9,1) = FA02(+) AND " & _
'                        "CP01 = CPM01(+) AND " & _
'                        "CP10 = CPM02(+) " & _
'                  "ORDER BY CP05 DESC, CP09 DESC"
        'Modify By Cheng 2003/01/27
        '加相關人及對造號數
'         strSQL = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10) AS CP10,NVL(S1.ST02,CP13) AS CP13,NVL(S2.ST02,CP14) AS CP14,CP24,NVL(CP27 - 19110000, NULL) AS CP27,CP43," & strSQLA & " AS CP44,NVL(CP57 - 19110000, NULL) AS CP57,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1,SYSTEMKIND "
         'Modified by Lydia 2015/02/11 專利權消滅(1604)的案件性質改顯示CP64
         'modify by sonia 2015/10/30 非台灣案CP64之專利權消滅改消滅
         'strSql = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,decode(cp01||cp10,'P1604',decode(sign(instr(cp64,'專利權消滅')),1,substr(cp64,instr(cp64,'專利權消滅'),15),NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10)),'CFP1604',decode(sign(instr(cp64,'專利權消滅')),1,substr(cp64,instr(cp64,'專利權消滅'),15),NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10)),NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10)) AS CP10," & _
                  "NVL(S1.ST02,CP13) AS CP13,NVL(S2.ST02,CP14) AS CP14,CP24,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))),CP36,NVL(CP27 - 19110000, NULL) AS CP27,CP43," & StrSQLa & " AS CP44,NVL(CP57 - 19110000, NULL) AS CP57," & _
                  "NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 " & _
                    " FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1,SYSTEMKIND,CUSTOMER " & _
                    "WHERE CP01 = '" & strCP01 & "' AND CP02 = '" & strCP02 & "' AND CP03 = '" & strCP03 & "' AND CP04 = '" & strCP04 & "' AND " & _
                        "CP13 = S1.ST01(+) AND CP14 = S2.ST01(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1) = FA02(+) AND " & _
                        "CP01 = CPM01(+) AND CP10 = CPM02(+) AND CP01=SK01(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) " & _
                  "ORDER BY CP05 DESC, CP66 DESC, CP67 DESC, CP09 DESC"      '2009/2/20 MODIFY BY SONIA 排序原為收文日+收文號,改收文日+CREATE DATE+CREATE TIME+收文號
         'Modified by Lydia 2022/09/06 +strWhereSql
         strSql = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,decode(cp01||cp10,'P1604',decode(sign(instr(cp64,'專利權消滅')),1,substr(cp64,instr(cp64,'專利權消滅'),15),decode(sign(instr(cp64,'消滅')),1,substr(cp64,instr(cp64,'消滅'),12),NVL(DECODE('" & m_Nation & "','000',CPM03,CPM04),CP10))),'CFP1604',decode(sign(instr(cp64,'消滅')),1,substr(cp64,instr(cp64,'消滅'),12),NVL(DECODE('" & m_Nation & "','000',CPM03,CPM04),CP10)),NVL(DECODE('" & m_Nation & "','000',CPM03,CPM04),CP10)) AS CP10," & _
                  "NVL(S1.ST02,CP13) AS CP13,NVL(S2.ST02,CP14) AS CP14,CP24,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))),CP36,NVL(CP27 - 19110000, NULL) AS CP27,CP43," & StrSQLa & " AS CP44,NVL(CP57 - 19110000, NULL) AS CP57," & _
                  "NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 " & _
                    " FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1,SYSTEMKIND,CUSTOMER " & _
                    "WHERE CP01 = '" & strCP01 & "' AND CP02 = '" & strCP02 & "' AND CP03 = '" & strCP03 & "' AND CP04 = '" & strCP04 & "' AND " & _
                        "CP13 = S1.ST01(+) AND CP14 = S2.ST01(+) AND SUBSTR(CP44,1,8) = FA01(+) AND SUBSTR(CP44,9,1) = FA02(+) AND " & _
                        "CP01 = CPM01(+) AND CP10 = CPM02(+) AND CP01=SK01(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) " & strWhereSql & _
                  "ORDER BY CP05 DESC, CP66 DESC, CP67 DESC, CP09 DESC"      '2009/2/20 MODIFY BY SONIA 排序原為收文日+收文號,改收文日+CREATE DATE+CREATE TIME+收文號
         'end 2015/10/30
      End If
   Else
      If m_Nation < "010" Then
         'Modify By Cheng 2002/07/09
'         strSQL = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10) AS CP10,S1.ST02 AS CP13,S2.ST02 AS CP14,CP24,NVL(CP27 - 19110000, NULL) AS CP27,CP43,NVL(F1.FA05,NVL(F1.FA04,NVL(F1.FA06,CP44))) AS CP44,NVL(CP57 - 19110000, NULL) AS CP57,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1 " & _
'                  "WHERE CP01 = '" & strcp01 & "' AND " & _
'                        "CP02 = '" & strcp02 & "' AND " & _
'                        "CP03 = '" & strcp03 & "' AND " & _
'                        "CP04 = '" & strcp04 & "' AND " & _
'                        "CP13 = S1.ST01(+) AND " & _
'                        "CP14 = S2.ST01(+) AND " & _
'                        "SUBSTR(CP44,1,8) = FA01(+) AND " & _
'                        "SUBSTR(CP44,9,1) = FA02(+) AND " & _
'                        "CP01 = CPM01(+) AND " & _
'                        "CP10 = CPM02(+) AND " & _
'                        "CP09 < 'C' " & _
'                  "ORDER BY CP05 DESC, CP09 DESC"
        'Modify By Cheng 2003/01/27
        '加相關人及對造號數
'         strSQL = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10) AS CP10,S1.ST02 AS CP13,S2.ST02 AS CP14,CP24,NVL(CP27 - 19110000, NULL) AS CP27,CP43," & strSQLA & " AS CP44,NVL(CP57 - 19110000, NULL) AS CP57,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1,SYSTEMKIND "
        'Modified by Lydia 2015/02/11 專利權消滅(1604)的案件性質改顯示CP64
         'Modified by Lydia 2022/09/06 +strWhereSql
         strSql = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10) AS CP10,S1.ST02 AS CP13,S2.ST02 AS CP14,CP24,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))),CP36,NVL(CP27 - 19110000, NULL) AS CP27,CP43," & StrSQLa & " AS CP44,NVL(CP57 - 19110000, NULL) AS CP57,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 " & _
                    " FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1,SYSTEMKIND,CUSTOMER " & _
                    "WHERE CP01 = '" & strCP01 & "' AND " & _
                        "CP02 = '" & strCP02 & "' AND " & _
                        "CP03 = '" & strCP03 & "' AND " & _
                        "CP04 = '" & strCP04 & "' AND " & _
                        "CP13 = S1.ST01(+) AND " & _
                        "CP14 = S2.ST01(+) AND " & _
                        "SUBSTR(CP44,1,8) = FA01(+) AND " & _
                        "SUBSTR(CP44,9,1) = FA02(+) AND " & _
                        "CP01 = CPM01(+) AND " & _
                        "CP10 = CPM02(+) AND " & _
                        "CP09 < 'C' AND CP01=SK01(+) " & _
                        "AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) " & strWhereSql & _
                  "ORDER BY CP05 DESC, CP66 DESC, CP67 DESC, CP09 DESC"   '2009/2/20 MODIFY BY SONIA 排序原為收文日+收文號,改收文日+CREATE DATE+CREATE TIME+收文號
      Else
         'Modify By Cheng 2002/07/09
'         strSQL = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10) AS CP10,S1.ST02 AS CP13,S2.ST02 AS CP14,CP24,NVL(CP27 - 19110000, NULL) AS CP27,CP43,NVL(F1.FA05,NVL(F1.FA04,NVL(F1.FA06,CP44))) AS CP44,NVL(CP57 - 19110000, NULL) AS CP57,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1 " & _
'                  "WHERE CP01 = '" & strcp01 & "' AND " & _
'                        "CP02 = '" & strcp02 & "' AND " & _
'                        "CP03 = '" & strcp03 & "' AND " & _
'                        "CP04 = '" & strcp04 & "' AND " & _
'                        "CP13 = S1.ST01(+) AND " & _
'                        "CP14 = S2.ST01(+) AND " & _
'                        "SUBSTR(CP44,1,8) = FA01(+) AND " & _
'                        "SUBSTR(CP44,9,1) = FA02(+) AND " & _
'                        "CP01 = CPM01(+) AND " & _
'                        "CP10 = CPM02(+) AND " & _
'                        "CP09 < 'C' " & _
'                  "ORDER BY CP05 DESC, CP09 DESC"
        'Modify By Cheng 2003/01/27
        '加相關人及對造號數
'         strSQL = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10) AS CP10,S1.ST02 AS CP13,S2.ST02 AS CP14,CP24,NVL(CP27 - 19110000, NULL) AS CP27,CP43," & strSQLA & " AS CP44,NVL(CP57 - 19110000, NULL) AS CP57,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1,SYSTEMKIND "
         'Modified by Lydia 2022/09/06 +strWhereSql
         strSql = "SELECT NVL(CP05 - 19110000, NULL) AS CP05,CP09,NVL(DECODE('" & m_Nation & "','000',C1.CPM03,C1.CPM04),CP10) AS CP10,S1.ST02 AS CP13,S2.ST02 AS CP14,CP24,NVL(CP40,NVL(CP41,NVL(CP42,NVL(CP50,NVL(CP51,NVL(CP52,DECODE(CP56,CU01||CU02,CU04))))))),CP36,NVL(CP27 - 19110000, NULL) AS CP27,CP43," & StrSQLa & " AS CP44,NVL(CP57 - 19110000, NULL) AS CP57,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07 " & _
                    " FROM CASEPROGRESS, STAFF S1, STAFF S2, FAGENT F1, CASEPROPERTYMAP C1,SYSTEMKIND,CUSTOMER " & _
                    "WHERE CP01 = '" & strCP01 & "' AND " & _
                        "CP02 = '" & strCP02 & "' AND " & _
                        "CP03 = '" & strCP03 & "' AND " & _
                        "CP04 = '" & strCP04 & "' AND " & _
                        "CP13 = S1.ST01(+) AND " & _
                        "CP14 = S2.ST01(+) AND " & _
                        "SUBSTR(CP44,1,8) = FA01(+) AND " & _
                        "SUBSTR(CP44,9,1) = FA02(+) AND " & _
                        "CP01 = CPM01(+) AND " & _
                        "CP10 = CPM02(+) AND " & _
                        "CP09 < 'C' AND CP01=SK01(+) " & _
                        "AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) " & strWhereSql & _
                  "ORDER BY CP05 DESC, CP66 DESC, CP67 DESC, CP09 DESC"    '2009/2/20 MODIFY BY SONIA 排序原為收文日+收文號,改收文日+CREATE DATE+CREATE TIME+收文號
      End If
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   grdList.Visible = False 'Add by Morgan 2007/10/17
   If rsTmp.RecordCount > 0 Then
      QueryCaseProgress = True
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         'Modify by Amy 2025/01/07 +if (將ModRecord 合併於此)
         If m_CP09 <> MsgText(601) Then
            ' 若找不到則新增一筆
            If intRow = 0 Then
               grdList.Rows = grdList.Rows + 1
               nRow = grdList.Rows - 1
            Else
               nRow = intRow
            End If
            ' 清空
            For nCol = 1 To grdList.Cols - 1
               grdList.TextMatrix(nRow, nCol) = Empty
            Next nCol
         Else
            ' 新增一筆記錄
            grdList.Rows = grdList.Rows + 1
            nRow = grdList.Rows - 1
         End If
         ' 收文日
         If IsEmptyText(rsTmp.Fields("CP05")) = False Then
            grdList.TextMatrix(nRow, 1) = rsTmp.Fields("CP05")
         End If
         ' 總收文號的欄位
         If IsNull(rsTmp.Fields("CP09")) = False Then
            grdList.TextMatrix(nRow, 2) = rsTmp.Fields("CP09")
         End If
         ' 案件性質
         If IsNull(rsTmp.Fields("CP10")) = False Then
            grdList.TextMatrix(nRow, 3) = rsTmp.Fields("CP10")
         End If
        'Modify by Morgan 2007/10/17 判斷有相關總收文號才做
        'Me.grdList.TextMatrix(nRow, 3) = Me.grdList.TextMatrix(nRow, 3) & PUB_GetRelateCasePropertyName(Me.grdList.TextMatrix(nRow, 2), "1")
         ' 相關總收文號
         'If IsNull(rsTmp.Fields("CP43")) = False Then
            grdList.TextMatrix(nRow, 4) = "" & rsTmp.Fields("CP43")
            Me.grdList.TextMatrix(nRow, 3) = Me.grdList.TextMatrix(nRow, 3) & PUB_GetRelateCasePropertyName(Me.grdList.TextMatrix(nRow, 2), "1")
         'End If
         'end 2007/10/17
         
         ' 承辦人
         If IsNull(rsTmp.Fields("CP14")) = False Then
            grdList.TextMatrix(nRow, 5) = rsTmp.Fields("CP14")
         End If
         ' 智權人員
         If IsNull(rsTmp.Fields("CP13")) = False Then
            grdList.TextMatrix(nRow, 6) = rsTmp.Fields("CP13")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("CP06")) = False Then
            grdList.TextMatrix(nRow, 7) = rsTmp.Fields("CP06")
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("CP07")) = False Then
            grdList.TextMatrix(nRow, 8) = rsTmp.Fields("CP07")
         End If
         ' 發文日
         If IsNull(rsTmp.Fields("CP27")) = False Then
            grdList.TextMatrix(nRow, 9) = rsTmp.Fields("CP27")
         End If
         ' 結果
         If IsNull(rsTmp.Fields("CP24")) = False Then
            Select Case rsTmp.Fields("CP24")
            Case "1":
               grdList.TextMatrix(nRow, 10) = "准,勝"
            Case "2":
               grdList.TextMatrix(nRow, 10) = "駁,敗"
            'add by sonia 2024/8/2
            Case "3":
               grdList.TextMatrix(nRow, 10) = "部分勝"
            Case Else
               grdList.TextMatrix(nRow, 10) = "錯誤"
            'end 2024/8/2
         End Select
         End If
        'Add By Cheng 2003/02/18
        '相關人
        grdList.TextMatrix(nRow, 11) = "" & rsTmp.Fields(6).Value
        '對造號收
        grdList.TextMatrix(nRow, 12) = "" & rsTmp.Fields("CP36").Value
         ' 取消收文日
         If IsNull(rsTmp.Fields("CP57")) = False Then
            'Modify By Cheng 2003/02/18
            '移位置
'            grdList.TextMatrix(nRow, 11) = rsTmp.Fields("CP57")
            grdList.TextMatrix(nRow, 13) = rsTmp.Fields("CP57")
         End If
         ' 代理人
         If IsNull(rsTmp.Fields("CP44")) = False Then
            'Modify By Cheng 2003/02/18
            '移位置
'            grdList.TextMatrix(nRow, 12) = rsTmp.Fields("CP44")
            grdList.TextMatrix(nRow, 14) = rsTmp.Fields("CP44")
         End If
         rsTmp.MoveNext
      Loop
      'Modify by Amy 2025/01/07 +if (將ModRecord 合併於此)
      If m_CP09 <> MsgText(601) Then
         grdList.TextMatrix(nRow, 0) = "V"
      Else
         'Added by Lydia 2022/01/11 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
         If grdList.Rows > 1 Then
            grdList.FixedRows = 1
         End If
      End If
   End If
   grdList.Visible = True 'Add by Morgan 2007/10/17
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取資料
Private Function QueryData(ByVal bAll As Boolean) As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bQuery As Boolean
   Dim bQueryCP As Boolean
   Dim StrSQLa As String            '2009/8/19 ADD BY SONIA
   Dim rsA As New ADODB.Recordset   '2009/8/19 ADD BY SONIA
   
   bQuery = False
   
   ClearField
   InitialGridList
   
   ' 組成本所案號
   m_CP01 = textCP01
   m_CP02 = textCP02
   If m_CP01 = "TF" Then: m_CP02 = m_CP02 & textCP02_2
   m_CP03 = textCP03
   If IsEmptyText(m_CP03) = True Then: m_CP03 = "0"
   m_CP04 = textCP04
   If IsEmptyText(m_CP04) = True Then: m_CP04 = "00"
   
   'Added by Morgan 2013/4/23
   'FCP程序可維護FMP案
   If CheckSR09(strUserNum, m_CP01, "Y", , m_CP01, m_CP02, m_CP03, m_CP04) = False Then
      Set rsA = Nothing
      Exit Function
   End If
   'end 2013/4/23
   
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   'Mark by Lydia 2020/07/08 開放專利國外部程序可操作FMP非寰華案，但僅限進度檔及下一程序檔之維護 ! (不管P或FCP系統)
   'If FMP2open = True Then
   '   If PUB_FMPtoCheck(0, 1, Pub_strUserST05, m_CP01, m_CP02, m_CP03, m_CP04) = False Then
   '      Set rsA = Nothing
   '      Exit Function
   '   End If
   'End If
   'end 2020/07/08
   
   '2009/9/8 加註 by sonia T非台灣案非外商收文之案件不必寫程式控制,因為在系統類別外商人員即不可使用T案件
   '2009/8/19 add by sonia FCT無爭議程序之案件內商人員不可查詢(該案有內商承辦人者為FCT爭議案)
   If m_CP01 = "FCT" And Mid(GetStaffDepartment(strUserNum), 1, 2) = "P2" Then
      'modify by sonia 2021/9/23 FCT-047943異議案林靖傑承辦,桂英無法操作
      'StrSQLa = "Select * From CASEPROGRESS,STAFF Where CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "' AND CP03='" & m_CP03 & "' AND CP04='" & m_CP04 & "' AND CP14=ST01(+) AND SUBSTR(ST03,1,2)='P2' "
      'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
      StrSQLa = "Select * From CASEPROGRESS,STAFF Where CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "' AND CP03='" & m_CP03 & "' AND CP04='" & m_CP04 & "' AND CP14=ST01(+) AND (SUBSTR(ST03,1,2)='P2' or (cp10 in (" & TMdebate & ") And Not (cp01 = 'FCT' And InStr(" & FCT_NotTMdebate & ", cp10) > 0)))"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic
      If rsA.RecordCount = 0 Then
         'Add By Sindy 2023/3/7 FCT爭議案件之C類來函由內商程序發文
         rsA.Close
         If Pub_StrUserSt03 = "P22" And m_CP01 = "FCT" Then
            StrSQLa = "select * from staff_group,CASEPROGRESS where CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "' AND CP03='" & m_CP03 & "' AND CP04='" & m_CP04 & "' AND sg02='" & m_CP01 & "' and sg01='C1' and length(sg03)=4 and cp10=sg03"
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic
            If rsA.RecordCount = 0 Then
         '2023/3/7 END
               strMsg = "非FCT爭議案，您沒有使用該案號資料的權限"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               Exit Function
            End If
         End If
         '2023/3/7 END
         
'         strMsg = "非FCT爭議案，您沒有使用該案號資料的權限"
'         strTit = "查詢資料"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'         Exit Function
      Else
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
   End If
   '2009/8/19 END
   
   ' 依本所案號讀取基本檔案
   Select Case m_CP01
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         bQuery = QueryTradeMark(m_CP01, m_CP02, m_CP03, m_CP04)
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         bQuery = QueryPatent(m_CP01, m_CP02, m_CP03, m_CP04)
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/24 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         bQuery = QueryLawCase(m_CP01, m_CP02, m_CP03, m_CP04)
      ' 讀取顧問案件基本檔
      Case "LA":
         bQuery = QueryHireCase(m_CP01, m_CP02, m_CP03, m_CP04)
      ' 讀取服務業務基本檔
      Case Else:
         bQuery = QueryServicePractice(m_CP01, m_CP02, m_CP03, m_CP04)
   End Select
   
'Added by Lydia 2022/09/06 當以本所案號查詢時，若條件為TT-999999則點進度時再增加輸入收文年月對話框(預設當月，取消表示全部)以減少等待時間。
    pYYMM = ""
    'Modified by Lydia 2023/01/11 +LA-999999
    'If m_CP01 & m_CP02 = "TT999999" Then
    If InStr("TT999999,LA999999", m_CP01 & m_CP02) > 0 Then
JumpToReInput:
        pYYMM = InputBox("請輸入收文年月以減少等待時間，不輸入年月或按取消表示查詢全部資料。", "輸入收文年月", Left(strSrvDate(2), 5))
        If pYYMM <> "" Then
           If Val(Left(pYYMM, 5)) > Val(Left(strSrvDate(2), 5)) Then
               MsgBox "收文年月不可大於系統年月！", vbInformation
               GoTo JumpToReInput
           End If
           Me.Caption = "案件進度檔資料維護-收文年月：" & pYYMM
        End If
    End If
'end 2022/09/06
   
   
   ' 讀取案件進度檔
   bQueryCP = QueryCaseProgress(m_CP01, m_CP02, m_CP03, m_CP04, bAll)
   
   lblCaseMap.Caption = "": lblDivision.Caption = "" 'Added by Lydia 2015/11/03
   lblCaseMap2.Caption = "" 'Added by Lydia 2019/11/28
   lblCMboth.Caption = "" 'Added by Lydia 2016/06/14
   If bQuery = False Then
      strTit = "查詢資料"
      strMsg = "該筆不存在於基本檔中"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
        'Added by Lydia 2015/11/03　顯示一案兩請，擬制喪失新穎性案件，有分割案(母案顯示)
        If PUB_GetRefCaseChk(m_CP01, m_CP02, m_CP03, m_CP04, "CASEMAP", "3") = True Then
           lblCaseMap.Caption = "一案兩請"
        End If
        If PUB_GetRefCaseChk(m_CP01, m_CP02, m_CP03, m_CP04, "CASEMAP", "6") = True Then
           'Modified by Lydia 2019/11/28 P-123733有一案兩請和擬制喪失新穎性案件
           'lblCaseMap.Caption = "擬制喪失新穎性案件"
           lblCaseMap2.Caption = "擬制喪失新穎性案件"
        End If
        If PUB_GetRefCaseChk(m_CP01, m_CP02, m_CP03, m_CP04, "DIVISIONCASE", , "2") = True Then
           lblDivision.Caption = "有分割案"
        End If
        'end 2015/11/03
        'Added by Lydia 2016/06/14 +台灣大陸案件提示
        If (m_CP01 = "P" Or m_CP01 = "FCP") And m_Nation = 台灣國家代號 Then
           If PUB_GetRefCaseChk(m_CP01, m_CP02, m_CP03, m_CP04, "CASEMAP", "0", "A", 大陸國家代號) Then
              lblCMboth.Caption = "有大陸案"
           End If
        ElseIf m_CP01 = "P" And m_Nation = 大陸國家代號 Then
           If PUB_GetRefCaseChk(m_CP01, m_CP02, m_CP03, m_CP04, "CASEMAP", "0", "A", 台灣國家代號) Then
              lblCMboth.Caption = "有台灣案"
           End If
        End If
        'end 2016/06/14
      If bQueryCP = False Then
         strTit = "查詢資料"
         'Added by Lydia 2016/04/12 系統類別為 TS或S，且該案件備註欄(SP18)內含'轉入商標：'...字樣時，改顯示 "此案進度已轉入商標：XXX-XXXXX"，例 S-000307, S-000327, TS-001248
         If (m_CP01 = "TS" Or m_CP01 = "S") And InStr(m_SP18, "轉入商標") > 0 Then
            strMsg = "此案進度已" & Mid(m_SP18, InStr(m_SP18, "轉入商標"))
         Else
         'end 2016/04/12
            strMsg = "該筆案件無案件進度的資料"
         End If
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   QueryData = bQueryCP
End Function

' 初始化列表
Public Sub InitialGridList()
   
   grdList.Clear
   grdList.Rows = 1
   
    'Modify By Cheng 2003/01/28
'   grdList.Cols = 13
   grdList.Cols = 15

   grdList.ColWidth(0) = 300
   grdList.row = 0

   grdList.col = 0
   grdList.ColAlignment(0) = flexAlignCenterCenter
   grdList.col = 1
   grdList.Text = "收文日"
   grdList.ColWidth(1) = 800 '原800
   grdList.ColAlignment(1) = flexAlignCenterCenter
   grdList.col = 2
   grdList.Text = "總收文號"
   grdList.ColWidth(2) = 1000 '原1000
   grdList.ColAlignment(2) = flexAlignCenterCenter
   grdList.col = 3
   grdList.Text = "案件性質"
   grdList.ColWidth(3) = 1300
   grdList.ColAlignment(3) = flexAlignLeftCenter
   grdList.col = 4
   grdList.Text = "相關收文號"
   grdList.ColWidth(4) = 1050 '原1200
   grdList.ColAlignment(4) = flexAlignLeftCenter
   grdList.col = 5
   grdList.Text = "承辦人"
   grdList.ColWidth(5) = 700 '原800
   grdList.ColAlignment(5) = flexAlignLeftCenter
   grdList.col = 6
   grdList.Text = "智權人員"
   grdList.ColWidth(6) = 700 '原800
   grdList.ColAlignment(6) = flexAlignLeftCenter
   grdList.col = 7
   grdList.Text = "本所期限"
   grdList.ColWidth(7) = 800 '原800
   grdList.ColAlignment(7) = flexAlignCenterCenter
   grdList.col = 8
   grdList.Text = "法定期限"
   grdList.ColWidth(8) = 800 '原800
   grdList.ColAlignment(8) = flexAlignCenterCenter
   grdList.col = 9
   grdList.Text = "發文日"
   grdList.ColWidth(9) = 800 '原800
   grdList.ColAlignment(9) = flexAlignCenterCenter
   grdList.col = 10
   grdList.Text = "結果"
   grdList.ColWidth(10) = 600
   grdList.ColAlignment(10) = flexAlignCenterCenter
    'Add By Cheng 2003/01/27
    '加相關人
   grdList.col = 11
   grdList.Text = "相關人"
   grdList.ColWidth(11) = 1000
   grdList.ColAlignment(11) = flexAlignLeftCenter
    '加對造號數
   grdList.col = 12
   grdList.Text = "對造號數"
   grdList.ColWidth(12) = 1000
   grdList.ColAlignment(12) = flexAlignLeftCenter
   grdList.col = 13
   grdList.Text = "取消收文日"
   grdList.ColWidth(13) = 1000
   grdList.ColAlignment(13) = flexAlignCenterCenter
   grdList.col = 14
   grdList.Text = "代理人"
   grdList.ColWidth(14) = 1200
   grdList.ColAlignment(14) = flexAlignLeftCenter
   
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
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
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
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
   Dim bFind As Boolean
   bFind = False
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 0) = "V" Then
         bFind = True
         Exit For
      End If
   Next nRow
   If bFind = True Then
      frm075004_2.SetData 0, m_CP01, True
      frm075004_2.SetData 1, m_CP02, False
      frm075004_2.SetData 2, m_CP03, False
      frm075004_2.SetData 3, m_CP04, False
   Else
      ' 組成本所案號
      strCP01 = textCP01
      strCP02 = textCP02
      If strCP01 = "TF" Then: strCP02 = strCP02 & textCP02_2
      strCP03 = textCP03
      If IsEmptyText(strCP03) = True Then: strCP03 = "0"
      strCP04 = textCP04
      If IsEmptyText(strCP04) = True Then: strCP04 = "00"
      frm075004_2.SetData 0, strCP01, True
      frm075004_2.SetData 1, strCP02, False
      frm075004_2.SetData 2, strCP03, False
      frm075004_2.SetData 3, strCP04, False
   End If
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 0) = "V" Then
         frm075004_2.SetData 4, grdList.TextMatrix(nRow, 2), False
      End If
   Next nRow
   'frm075004_2.SetParent Me 'Added by Morgan 2018/10/9
   frm075004_2.Show
   frm075004_2.QueryDB
   Me.Hide
End Sub

' 刪除該筆記錄
Public Sub DelRecord(ByVal strData As String)
   Dim nRow As Integer
   Dim bFind As Boolean
   bFind = False
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 2) = strData Then
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
   Dim strCPSQL As String
   Dim nRow As Integer
   Dim nCol As Integer
   Dim bFind As Boolean
   
   ' 比對總收文號
   bFind = False
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 2) = strData Then
         bFind = True
         Exit For
      End If
   Next nRow
   
   'Add by Amy 2025/01/07 避免程式有未改到,統一抓QueryCaseProgress
   If bFind = False Then nRow = 0
   Call QueryCaseProgress(m_CP01, m_CP02, m_CP03, m_CP04, m_QueryAll, strData, nRow)
   
   'Mark by Amy 2025/01/07 以下程式不使用
'   ' 設定查詢資料庫的SQL語法
'   strSql = "SELECT * FROM CaseProgress " & _
'            "WHERE CP01 = '" & m_CP01 & "' AND " & _
'                  "CP02 = '" & m_CP02 & "' AND " & _
'                  "CP03 = '" & m_CP03 & "' AND " & _
'                  "CP04 = '" & m_CP04 & "' AND " & _
'                  "CP09 = '" & strData & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      ' 若找不到則新增一筆
'      If bFind = False Then
'         grdList.Rows = grdList.Rows + 1
'         nRow = grdList.Rows - 1
'      End If
'      ' 清空
'      For nCol = 1 To grdList.Cols - 1
'         grdList.TextMatrix(nRow, nCol) = Empty
'      Next nCol
'
'      ' 收文日
'      If IsEmptyText(rsTmp.Fields("CP05")) = False Then
'         If rsTmp.Fields("CP05") <> "0" Then
'            grdList.TextMatrix(nRow, 1) = ChangeWStringToTString(rsTmp.Fields("CP05"))
'         End If
'      End If
'      ' 總收文號的欄位
'      If IsNull(rsTmp.Fields("CP09")) = False Then
'         grdList.TextMatrix(nRow, 2) = rsTmp.Fields("CP09")
'      End If
'      ' 案件性質
'      If IsNull(rsTmp.Fields("CP10")) = False Then
'         If m_Nation < "010" Then
'            grdList.TextMatrix(nRow, 3) = GetCaseTypeName(m_CP01, rsTmp.Fields("CP10"), 0)
'         Else
'            grdList.TextMatrix(nRow, 3) = GetCaseTypeName(m_CP01, rsTmp.Fields("CP10"), 1)
'         End If
'      End If
'      ' 相關總收文號
'      If IsNull(rsTmp.Fields("CP43")) = False Then
'         grdList.TextMatrix(nRow, 4) = rsTmp.Fields("CP43")
'      End If
'      ' 承辦人
'      If IsNull(rsTmp.Fields("CP14")) = False Then
'         grdList.TextMatrix(nRow, 5) = GetStaffName(rsTmp.Fields("CP14"), True)
'      End If
'      ' 智權人員
'      If IsNull(rsTmp.Fields("CP13")) = False Then
'         grdList.TextMatrix(nRow, 6) = GetStaffName(rsTmp.Fields("CP13"), True)
'      End If
'      ' 本所期限
'      If IsNull(rsTmp.Fields("CP06")) = False Then
'         If rsTmp.Fields("CP06") <> "0" Then
'            grdList.TextMatrix(nRow, 7) = ChangeWStringToTString(rsTmp.Fields("CP06"))
'         End If
'      End If
'      ' 法定期限
'      If IsNull(rsTmp.Fields("CP07")) = False Then
'         If rsTmp.Fields("CP07") <> "0" Then
'            grdList.TextMatrix(nRow, 8) = ChangeWStringToTString(rsTmp.Fields("CP07"))
'         End If
'      End If
'      ' 發文日
'      If IsNull(rsTmp.Fields("CP27")) = False Then
'         If rsTmp.Fields("CP27") <> "0" Then
'            grdList.TextMatrix(nRow, 9) = ChangeWStringToTString(rsTmp.Fields("CP27"))
'         End If
'      End If
'      ' 結果
'      If IsNull(rsTmp.Fields("CP24")) = False Then
'         Select Case rsTmp.Fields("CP24")
'            Case "1":
'               grdList.TextMatrix(nRow, 10) = "准,勝"
'            Case "2":
'               grdList.TextMatrix(nRow, 10) = "駁,敗"
'            'add by sonia 2024/8/2
'            Case "3":
'               grdList.TextMatrix(nRow, 10) = "部分勝"
'            Case Else
'               grdList.TextMatrix(nRow, 10) = "錯誤"
'            'end 2024/8/2
'         End Select
'      End If
'      ' 取消收文日
'      If IsNull(rsTmp.Fields("CP57")) = False Then
'         If rsTmp.Fields("CP57") <> "0" Then
'            grdList.TextMatrix(nRow, 13) = ChangeWStringToTString(rsTmp.Fields("CP57"))
'         End If
'      End If
'      ' 代理人
'      If IsNull(rsTmp.Fields("CP44")) = False Then
'         '移位置
'         'grdList.TextMatrix(nRow, 12) = GetFAgentName(rsTmp.Fields("CP44"))
'         grdList.TextMatrix(nRow, 14) = GetFAgentName(rsTmp.Fields("CP44"))
'      End If
'      grdList.TextMatrix(nRow, 0) = "V"
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
End Sub

Private Sub textCP01_GotFocus()
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP01.IMEMode = 2
   CloseIme
   InverseTextBox textCP01
End Sub

Private Sub textCP02_GotFocus()
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP02.IMEMode = 2
   CloseIme
   InverseTextBox textCP02
End Sub

Private Sub textCP02_2_GotFocus()
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP02_2.IMEMode = 2
   CloseIme
   InverseTextBox textCP02_2
End Sub

Private Sub textCP03_GotFocus()
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06 切換輸入法改用API
   'extCP03.IMEMode = 2
   CloseIme
   InverseTextBox textCP03
End Sub

Private Sub textCP03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP04_GotFocus()
   'Add by Morgan 2006/3/15
   '加輸入法控制
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP04.IMEMode = 2
   CloseIme
   InverseTextBox textCP04
End Sub

Public Sub ClearRemark()
   Dim nRow As Integer
   For nRow = 0 To grdList.Rows - 1
      grdList.TextMatrix(nRow, 0) = Empty
   Next nRow
End Sub

Public Sub RefreshList()
   QueryData m_QueryAll
End Sub

Private Function IsDBExist() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   IsDBExist = False
   Select Case m_CP01
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         strSql = "SELECT * FROM TRADEMARK " & _
                  "WHERE TM01 = '" & m_CP01 & "' AND " & _
                        "TM02 = '" & m_CP02 & "' AND " & _
                        "TM03 = '" & m_CP03 & "' AND " & _
                        "TM04 = '" & m_CP04 & "' "
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         strSql = "SELECT * FROM PATENT " & _
                  "WHERE PA01 = '" & m_CP01 & "' AND " & _
                        "PA02 = '" & m_CP02 & "' AND " & _
                        "PA03 = '" & m_CP03 & "' AND " & _
                        "PA04 = '" & m_CP04 & "' "
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/24 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         strSql = "SELECT * FROM LAWCASE " & _
                  "WHERE LC01 = '" & m_CP01 & "' AND " & _
                        "LC02 = '" & m_CP02 & "' AND " & _
                        "LC03 = '" & m_CP03 & "' AND " & _
                        "LC04 = '" & m_CP04 & "' "
      ' 讀取顧問案件基本檔
      Case "LA":
         strSql = "SELECT * FROM HIRECASE " & _
                  "WHERE HC01 = '" & m_CP01 & "' AND " & _
                        "HC02 = '" & m_CP02 & "' AND " & _
                        "HC03 = '" & m_CP03 & "' AND " & _
                        "HC04 = '" & m_CP04 & "' "
      ' 讀取服務業務基本檔
      Case Else:
         strSql = "SELECT * FROM SERVICEPRACTICE " & _
                  "WHERE SP01 = '" & m_CP01 & "' AND " & _
                        "SP02 = '" & m_CP02 & "' AND " & _
                        "SP03 = '" & m_CP03 & "' AND " & _
                        "SP04 = '" & m_CP04 & "' "
   End Select
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsDBExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
