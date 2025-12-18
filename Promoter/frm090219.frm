VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090219 
   BorderStyle     =   1  '單線固定
   Caption         =   "電話回覆主管機關"
   ClientHeight    =   5130
   ClientLeft      =   900
   ClientTop       =   1050
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   4410
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   7710
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "新增電話回覆記錄(&O)"
      Height          =   400
      Index           =   0
      Left            =   5670
      TabIndex        =   11
      Top             =   120
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      Caption         =   "變更事項："
      Height          =   3060
      Left            =   90
      TabIndex        =   18
      Top             =   1320
      Width           =   8775
      Begin VB.CheckBox Check1 
         Caption         =   "變更商品"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   2220
         Width           =   1095
      End
      Begin MSForms.TextBox textTM67 
         Height          =   420
         Left            =   1410
         TabIndex        =   8
         Top             =   2520
         Width           =   7155
         VariousPropertyBits=   -1467989989
         MaxLength       =   200
         Size            =   "12621;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE45 
         Height          =   810
         Left            =   1410
         TabIndex        =   6
         Top             =   1380
         Width           =   7155
         VariousPropertyBits=   -1467989989
         MaxLength       =   200
         ScrollBars      =   2
         Size            =   "12621;1429"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41_1 
         Height          =   1005
         Left            =   1410
         TabIndex        =   5
         Top             =   270
         Width           =   7155
         VariousPropertyBits=   -1467989989
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "12621;1773"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE42 
         Height          =   330
         Left            =   1410
         TabIndex        =   14
         Top             =   600
         Width           =   7155
         VariousPropertyBits=   -1467989989
         MaxLength       =   180
         Size            =   "12621;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41 
         Height          =   330
         Left            =   1410
         TabIndex        =   13
         Top             =   270
         Width           =   7155
         VariousPropertyBits=   -1467989989
         MaxLength       =   160
         Size            =   "12621;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE43 
         Height          =   330
         Left            =   1410
         TabIndex        =   15
         Top             =   930
         Width           =   7155
         VariousPropertyBits=   -1467989989
         MaxLength       =   160
         Size            =   "12621;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label35 
         Caption         =   "案件名稱(英) :"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "案件名稱(日) :"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label Label38 
         Caption         =   "縮減商品 :"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label40 
         Caption         =   "放棄專用權 :"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2550
         Width           =   1155
      End
      Begin VB.Label Label34 
         Caption         =   "案件名稱(中) :"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label45 
         Caption         =   "案件名稱 :"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.TextBox textCP04 
      Height          =   300
      Left            =   2775
      MaxLength       =   2
      TabIndex        =   3
      Top             =   255
      Width           =   375
   End
   Begin VB.TextBox textCP03 
      Height          =   300
      Left            =   2535
      MaxLength       =   1
      TabIndex        =   2
      Top             =   255
      Width           =   255
   End
   Begin VB.TextBox textCP02 
      Height          =   300
      Left            =   1575
      MaxLength       =   6
      TabIndex        =   1
      Top             =   255
      Width           =   975
   End
   Begin VB.TextBox textCP01 
      Height          =   300
      Left            =   1095
      MaxLength       =   3
      TabIndex        =   0
      Top             =   255
      Width           =   495
   End
   Begin MSForms.Label Label16 
      Height          =   255
      Left            =   6090
      TabIndex        =   32
      Top             =   990
      Width           =   1485
      Size            =   "2619;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label13 
      Height          =   255
      Left            =   1110
      TabIndex        =   31
      Top             =   990
      Width           =   1485
      Size            =   "2619;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   7740
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "13652;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   585
      Left            =   1080
      TabIndex        =   9
      Top             =   4440
      Width           =   7755
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      Size            =   "13679;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      Caption         =   "承辦人："
      Height          =   210
      Left            =   5250
      TabIndex        =   26
      Top             =   990
      Width           =   795
   End
   Begin VB.Label Label24 
      Caption         =   "智權人員："
      Height          =   210
      Left            =   90
      TabIndex        =   25
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "業務區："
      Height          =   210
      Left            =   2670
      TabIndex        =   24
      Top             =   990
      Width           =   825
   End
   Begin VB.Label Label14 
      Height          =   270
      Left            =   3540
      TabIndex        =   23
      Top             =   990
      Width           =   1455
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
      Height          =   180
      Left            =   3210
      TabIndex        =   20
      Top             =   300
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "進度備註："
      Height          =   210
      Left            =   90
      TabIndex        =   19
      Top             =   4470
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   210
      Left            =   90
      TabIndex        =   17
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   210
      Left            =   90
      TabIndex        =   16
      Top             =   300
      Width           =   975
   End
End
Attribute VB_Name = "frm090219"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 改成Form2.0 ;textCP64、Combo1、textTM67、Label13、Label16、textCE45、textCE41、textCE42、textCE43、textCE41_1
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit

'本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String

Dim m_TM05 As String
Dim m_TM06 As String
Dim m_TM07 As String
Dim m_TM09 As String
Dim m_TM10 As String
Dim m_TM67 As String

'收文號
Dim m_CP09 As String

Dim m_CP12 As String
Dim m_CP13 As String
Dim m_CP64 As String

Public ChkTG As Boolean

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0 '新增電話回覆記錄
         '檢查輸入資料的有效性
         If CheckDataValidate = False Then Exit Sub
         If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      Case 1 '查詢
         QueryData
      Case 2 '結束
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   cmdOK(0).Enabled = False
   lblClose = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090219 = Nothing
End Sub

Private Sub QueryData()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strTemp As String
   
   onClear
   
   If textCP01 = "" Or textCP02 = "" Then
      MsgBox "請輸入案號!!!", vbExclamation + vbOKOnly
      If textCP01 = "" Then
         Me.textCP01.SetFocus
      ElseIf textCP02 = "" Then
         Me.textCP02.SetFocus
      End If
      Exit Sub
   End If
   
   Label16 = strUserName
   
   If textCP03 = "" Then textCP03 = "0"
   If textCP04 = "" Then textCP04 = "00"
   
   m_TM01 = Trim(textCP01)
   m_TM02 = Trim(textCP02)
   m_TM03 = Trim(textCP03)
   m_TM04 = Trim(textCP04)
   
   strSql = "SELECT * FROM TradeMark,Nation " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' AND " & _
                  "TM10 = NA01(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      m_TM10 = "" & rsTmp.Fields("TM10")
      ' 商標名稱(中)
      m_TM05 = "" & rsTmp.Fields("TM05")
      If IsNull(rsTmp.Fields("TM05")) = False Then
         Combo1.AddItem rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      m_TM06 = "" & rsTmp.Fields("TM06")
      If IsNull(rsTmp.Fields("TM06")) = False Then
         Combo1.AddItem rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      m_TM07 = "" & rsTmp.Fields("TM07")
      If IsNull(rsTmp.Fields("TM07")) = False Then
         Combo1.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If Combo1.ListCount > 0 Then
         Combo1.ListIndex = 0
      End If
      ' 商品類別
      m_TM09 = "" & rsTmp.Fields("TM09")
      ' 是否閉卷
      If IsNull(rsTmp.Fields("TM29")) = False Then
         If rsTmp.Fields("TM29") = "Y" Then
            lblClose = "已閉卷"
         End If
      End If
      ' 智權人員
      If m_TM01 = "FCP" Or m_TM01 = "FG" Then
         m_CP13 = PUB_GetFCPSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      ElseIf m_TM01 = "FCL" Or m_TM01 = "LIN" Then
         m_CP13 = PUB_GetFCLSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      ElseIf m_TM01 = "FCT" Or _
               (m_TM01 = "S" And Trim("" & rsTmp.Fields("TM10")) = "000") Then
         m_CP13 = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      Else
         m_CP13 = PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      End If
      Label13 = GetStaffName(m_CP13)
      ' 業務區
      m_CP12 = GetST15(m_CP13, strTemp)
      Label14 = strTemp
      ' 放棄專用權
      m_TM67 = "" & rsTmp.Fields("TM67")
      If IsNull(rsTmp.Fields("TM67")) = False Then
         textTM67 = rsTmp.Fields("TM67")
      End If
   Else
      MsgBox "查無此案號!!!", vbExclamation + vbOKOnly
      Me.textCP02.SetFocus
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   If lblClose = "已閉卷" Or m_TM10 <> "000" Then
      If lblClose = "已閉卷" Then
         MsgBox "案號已閉卷不可執行此作業!!!", vbExclamation + vbOKOnly
         Me.textCP01.SetFocus
      ElseIf m_TM10 <> "000" Then
         MsgBox "申請國家為非台灣不可執行此作業!!!", vbExclamation + vbOKOnly
         Me.textCP02.SetFocus
      End If
      textCP01.BorderStyle = 1
      textCP01.Locked = False
      textCP01.BackColor = &H80000005
      textCP02.BorderStyle = 1
      textCP02.Locked = False
      textCP02.BackColor = &H80000005
      textCP03.BorderStyle = 1
      textCP03.Locked = False
      textCP03.BackColor = &H80000005
      textCP04.BorderStyle = 1
      textCP04.Locked = False
      textCP04.BackColor = &H80000005
      cmdOK(0).Enabled = False
      cmdOK(0).Default = False
      cmdOK(1).Enabled = True
      cmdOK(1).Default = True
      Exit Sub
   Else
      textCP01.BorderStyle = 0
      textCP01.Locked = True
      textCP01.BackColor = &H8000000F
      textCP02.BorderStyle = 0
      textCP02.Locked = True
      textCP02.BackColor = &H8000000F
      textCP03.BorderStyle = 0
      textCP03.Locked = True
      textCP03.BackColor = &H8000000F
      textCP04.BorderStyle = 0
      textCP04.Locked = True
      textCP04.BackColor = &H8000000F
      cmdOK(0).Enabled = True
      cmdOK(0).Default = True
      cmdOK(1).Enabled = False
      cmdOK(1).Default = False
      textCP64 = "分機："
      m_CP64 = Trim(textCP64)
   End If
End Sub

Private Sub onClear()
   Combo1.Clear
   lblClose = ""
   Label13 = ""
   Label14 = ""
   Label16 = ""
   textCE41 = ""
   textCE41_1 = ""
   textCE42 = ""
   textCE43 = ""
   textCE45 = ""
   textCP64 = ""
   textTM67 = ""
   Check1.Value = 0
End Sub

Public Function OnSaveData() As Boolean
Dim strSql As String
Dim strTmp As String
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strCP64 As String
   
'On Error GoTo ErrorHandler
OnSaveData = True
   
   If Trim(textCE45) <> "" Or Check1.Value = 1 Then
      frm03010303_04.Hide
      Set frm03010303_04.UpForm = Me
      frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      frm03010303_04.AllClass = m_TM09
      frm03010303_04.cmdOK(2).Visible = True
      Me.Hide
      frm03010303_04.QueryData
      frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
   End If
   
   cnnConnection.BeginTrans
   
   ' 進度備註
   strCP64 = Trim(textCP64)
   If Trim(textCE41) <> "" Then strCP64 = strCP64 & "；原案件名稱(中)：" & m_TM05
   If Trim(textCE41_1) <> "" Then strCP64 = strCP64 & "；原案件名稱：" & m_TM05
   If Trim(textCE42) <> "" Then strCP64 = strCP64 & "；原案件名稱(英)：" & m_TM06
   If Trim(textCE43) <> "" Then strCP64 = strCP64 & "；原案件名稱(日)：" & m_TM07
   If Trim(textTM67) <> m_TM67 Then strCP64 = strCP64 & "；原放棄專用權：" & m_TM67
   If Check1.Value = 1 Then strCP64 = strCP64 & "；變更商品"
   
   ' 新增電話回覆記錄
   m_CP09 = AutoNo("B", 6)
   strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13," & _
                  "cp14,cp20,cp26,cp27,cp32,cp64,cp83) " & _
                  "values (" & CNULL(m_TM01) & ", " & _
                  CNULL(m_TM02) & ", " & _
                  CNULL(m_TM03) & ", " & _
                  CNULL(m_TM04) & ", " & _
                  CNULL(strSrvDate(1)) & ", " & _
                  CNULL(m_CP09) & ", " & _
                  "209, " & _
                  CNULL(m_CP12) & ", " & _
                  CNULL(m_CP13) & ", " & _
                  "'" & strUserNum & "', " & _
                  "'N', " & _
                  "'N', " & _
                  strSrvDate(1) & ", " & _
                  "'N', " & _
                  CNULL(ChgSQL(strCP64)) & ", " & _
                  "'" & strUserNum & "')"
   cnnConnection.Execute strSql
   
   ' 先刪除掉已存在的資料
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.Close
      strSql = "DELETE FROM ChangeEvent " & _
               "WHERE CE01 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   Else
      rsTmp.Close
   End If
   
   ' 新增資料到變更事項檔
   bDifference = False
   ' 執行SQL指令
   strSql = "INSERT INTO ChangeEvent (CE01"
   ' 欄位名稱
   If Trim(textCE41) <> "" Or Trim(textCE41_1) <> "" Then
      bDifference = True
      strSql = strSql & ",CE41"
   End If
   If Trim(textCE42) <> "" Then
      bDifference = True
      strSql = strSql & ",CE42"
   End If
   If Trim(textCE43) <> "" Then
      bDifference = True
      strSql = strSql & ",CE43"
   End If
   If Trim(textCE45) <> "" Then
      bDifference = True
      strSql = strSql & ",CE45"
   End If
   strSql = strSql & ") VALUES (" & CNULL(m_CP09)
   ' 欄位值
   If Trim(textCE41) <> "" Or Trim(textCE41_1) <> "" Then
      If Trim(textCE41) <> "" Then
         strSql = strSql & "," & CNULL(textCE41)
      ElseIf Trim(textCE41_1) <> "" Then
         strSql = strSql & "," & CNULL(textCE41_1)
      End If
   End If
   If Trim(textCE42) <> "" Then
      strSql = strSql & "," & CNULL(textCE42)
   End If
   If Trim(textCE43) <> "" Then
      strSql = strSql & "," & CNULL(textCE43)
   End If
   If Trim(textCE45) <> "" Then
      strSql = strSql & "," & CNULL(textCE45)
   End If
   strSql = strSql & ")"
   If bDifference = True Then
      cnnConnection.Execute strSql
   End If
   
   ' 更新商標基本檔
   bDifference = False
   bFirst = True
   ' 執行SQL指令
   strSql = "UPDATE Trademark SET "
   If Trim(textCE41) <> "" Then
      bDifference = True
      If bFirst = False Then strSql = strSql & ","
      strSql = strSql & "TM05='" & textCE41 & "'": bFirst = False
   End If
   If Trim(textCE41_1) <> "" Then
      bDifference = True
      If bFirst = False Then strSql = strSql & ","
      strSql = strSql & "TM05='" & textCE41_1 & "'": bFirst = False
   End If
   If Trim(textCE42) <> "" Then
      bDifference = True
      If bFirst = False Then strSql = strSql & ","
      strSql = strSql & "TM06='" & textCE42 & "'": bFirst = False
   End If
   If Trim(textCE43) <> "" Then
      bDifference = True
      If bFirst = False Then strSql = strSql & ","
      strSql = strSql & "TM07='" & textCE43 & "'": bFirst = False
   End If
   If Trim(textTM67) <> m_TM67 Then
      bDifference = True
      If bFirst = False Then strSql = strSql & ","
      strSql = strSql & "TM67='" & ChgSQL(textTM67) & "'": bFirst = False
   End If
   If bDifference = True Then
      strSql = strSql & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                                 "TM02 = '" & m_TM02 & "' AND " & _
                                 "TM03 = '" & m_TM03 & "' AND " & _
                                 "TM04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
   End If
   
   Set rsTmp = Nothing
   cnnConnection.CommitTrans
   
   textCP01 = "": textCP02 = "": textCP03 = "": textCP04 = ""
   onClear
   textCP01.BorderStyle = 1
   textCP01.Locked = False
   textCP01.BackColor = &H80000005
   textCP02.BorderStyle = 1
   textCP02.Locked = False
   textCP02.BackColor = &H80000005
   textCP03.BorderStyle = 1
   textCP03.Locked = False
   textCP03.BackColor = &H80000005
   textCP04.BorderStyle = 1
   textCP04.Locked = False
   textCP04.BackColor = &H80000005
   cmdOK(0).Enabled = False
   cmdOK(0).Default = False
   cmdOK(1).Enabled = True
   cmdOK(1).SetFocus
   cmdOK(1).Default = True
   Exit Function
   
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
    
End Function

Private Function CheckDataValidate() As Boolean
Dim Cancel As Boolean
    
    CheckDataValidate = False
    Cancel = False
    
    If textCE41_1 = "" And textCE41 = "" And textCE42 = "" And textCE43 = "" And _
       textCE45 = "" And textTM67 = m_TM67 And Check1.Value <> 1 And _
       (textCP64 = m_CP64 Or Trim(textCP64) = "") Then
       MsgBox "請輸入資料 !", vbCritical
       Exit Function
    End If
    If textCE41 <> "" Then
       textCE41_Validate Cancel
       If Cancel = True Then Exit Function
    End If
    If textCE41_1 <> "" Then
       textCE41_1_Validate Cancel
       If Cancel = True Then Exit Function
    End If
    If textCE42 <> "" Then
       textCE42_Validate Cancel
       If Cancel = True Then Exit Function
    End If
    If textCE43 <> "" Then
       textCE43_Validate Cancel
       If Cancel = True Then Exit Function
    End If
    If textCE45 <> "" Then
       textCE45_Validate Cancel
       If Cancel = True Then Exit Function
    End If
    If textTM67 <> "" Then
       textTM67_Validate Cancel
       If Cancel = True Then Exit Function
    End If
    
    'Added by Lydia 2021/12/23 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If

    CheckDataValidate = True
End Function

Private Sub textCE41_1_GotFocus()
TextInverse Me.textCE41_1
End Sub

Private Sub textCE41_1_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE41_1) Then Exit Sub
If CheckLengthIsOK(textCE41_1.Text, textCE41_1.MaxLength) = False Then
    MsgBox "案件名稱超過長度!!!", vbExclamation + vbOKOnly
    Me.textCE41_1.SetFocus
    textCE41_1_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub

Private Sub textCE41_GotFocus()
InverseTextBox textCE41
End Sub

Private Sub textCE41_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE41) Then Exit Sub
If CheckLengthIsOK(textCE41.Text, textCE41.MaxLength) = False Then
    MsgBox "案件名稱(中)超過長度!!!", vbExclamation + vbOKOnly
    Me.textCE41.SetFocus
    textCE41_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub

Private Sub textCE42_GotFocus()
InverseTextBox textCE42
End Sub

Private Sub textCE42_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE42) Then Exit Sub
If CheckLengthIsOK(textCE42.Text, textCE42.MaxLength) = False Then
    MsgBox "案件名稱(英)超過長度!!!", vbExclamation + vbOKOnly
    Me.textCE42.SetFocus
    textCE42_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub

Private Sub textCE43_GotFocus()
InverseTextBox textCE43
End Sub

Private Sub textCE43_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE43) Then Exit Sub
If CheckLengthIsOK(textCE43.Text, textCE43.MaxLength) = False Then
    MsgBox "案件名稱(日)超過長度!!!", vbExclamation + vbOKOnly
    Me.textCE43.SetFocus
    textCE43_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub

Private Sub textCE45_GotFocus()
InverseTextBox textCE45
End Sub

Private Sub textCE45_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textCE45) Then Exit Sub
If CheckLengthIsOK(textCE45.Text, textCE45.MaxLength) = False Then
    MsgBox "縮減商品超過長度!!!", vbExclamation + vbOKOnly
    Me.textCE45.SetFocus
    textCE45_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub

Private Sub textCP03_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM67_GotFocus()
InverseTextBox textTM67
End Sub

Private Sub textTM67_Validate(Cancel As Boolean)
Cancel = False
If IsEmpty(textTM67) Then Exit Sub
If CheckLengthIsOK(textTM67.Text, textTM67.MaxLength) = False Then
    MsgBox "放棄專用權超過長度!!!", vbExclamation + vbOKOnly
    Me.textTM67.SetFocus
    textTM67_GotFocus
    Cancel = True
    Exit Sub
End If
End Sub

Private Sub textCP01_GotFocus()
    TextInverse textCP01
    CloseIme
End Sub

Private Sub textCP01_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP01_Validate(Cancel As Boolean)
   Dim strTemp1
   Dim strTemp2
   Dim ii As Integer
   Dim jj As Integer
   Dim ss As Integer
   Dim m_Dept As String   '2006/12/11 ADD BY SONIA
   
   If textCP01.Text = "" Then Exit Sub
   
   If textCP01.Text <> "T" And textCP01.Text <> "FCT" Then
      MsgBox "系統別只可以輸入T或FCT!!!", vbExclamation + vbOKOnly
      Me.textCP01.SetFocus
      textCP01_GotFocus
      Cancel = True
      Exit Sub
   End If
   
   strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
   strTemp2 = Split(Replace(UCase(textCP01.Text), ",,", ""), ",")
   For ii = 0 To UBound(strTemp2)
       ss = 0
       For jj = 0 To UBound(strTemp1)
           If strTemp2(ii) = strTemp1(jj) Then
               ss = 1
               Exit For
           End If
       Next jj
       If ss = 0 Then
          '2006/12/11 ADD BY SONIA 開放FF案件之權限
          m_Dept = GetStaffDepartment(strUserNum)
          Select Case m_Dept
            'Modify by Morgan 2007/4/11 加F61
            'Modify by Morgan 2008/4/8 加F81
             Case "F21", "F23", "F61", "F81"  '開放F21,F23使用P,PS,CFP,CPS權限
                If textCP01.Text = "P" Or textCP01.Text = "PS" Or textCP01.Text = "CFP" Or textCP01.Text = "CPS" Then
                   Exit For
                End If
             Case "F10", "F11"    '開放F10,F11使用T權限
                If textCP01.Text = "T" Then
                   Exit For
                End If
          End Select
          '2006/12/11 END
           ss = MsgBox(strUserName & " 沒有 " & strTemp2(ii) & " 的權限!! ", , "USER 權限問題")
           textCP01.SetFocus
           textCP01_GotFocus
           Cancel = True
       End If
   Next ii
   
   Select Case Trim(textCP01)
      Case "T", "FCT", "CFT", "TF", "TS"
         Me.Label34.Visible = False
         Me.Label35.Visible = False
         Me.Label36.Visible = False
         Me.textCE41.Visible = False
         Me.textCE41.Enabled = False
         Me.textCE42.Visible = False
         Me.textCE42.Enabled = False
         Me.textCE43.Visible = False
         Me.textCE43.Enabled = False
         Me.Label45.Visible = True
         Me.textCE41_1.Visible = True
         Me.textCE41_1.Enabled = True
      Case Else
         Me.Label34.Visible = True
         Me.Label35.Visible = True
         Me.Label36.Visible = True
         Me.textCE41.Visible = True
         Me.textCE41.Enabled = True
         Me.textCE42.Visible = True
         Me.textCE42.Enabled = True
         Me.textCE43.Visible = True
         Me.textCE43.Enabled = True
         Me.Label45.Visible = False
         Me.textCE41_1.Visible = False
         Me.textCE41_1.Enabled = False
   End Select
End Sub

Private Sub textCP02_GotFocus()
    TextInverse textCP02
    CloseIme
End Sub

Private Sub textCP02_Validate(Cancel As Boolean)
   If textCP03 = "" Then textCP03 = "0"
End Sub

Private Sub textCP03_GotFocus()
    TextInverse textCP03
    CloseIme
End Sub

Private Sub textCP03_Validate(Cancel As Boolean)
   If textCP03 = "" Then textCP03 = "0"
   If textCP04 = "" Then textCP04 = "00"
End Sub

Private Sub textCP04_GotFocus()
    TextInverse textCP04
    CloseIme
End Sub

Private Sub textCP04_Validate(Cancel As Boolean)
   If textCP03 = "" Then textCP03 = "0"
   If textCP04 = "" Then textCP04 = "00"
End Sub
