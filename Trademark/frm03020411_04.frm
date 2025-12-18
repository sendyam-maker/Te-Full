VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020411_04 
   BorderStyle     =   1  '單線固定
   Caption         =   "修改承辦人"
   ClientHeight    =   3210
   ClientLeft      =   -3255
   ClientTop       =   4830
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   9150
   Begin VB.CommandButton CmdOpen 
      Cancel          =   -1  'True
      Caption         =   "卷宗區"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5880
      TabIndex        =   24
      Top             =   30
      Width           =   912
   End
   Begin VB.TextBox textCP07 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   23
      Top             =   2160
      Width           =   1125
   End
   Begin VB.TextBox textCP06 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   22
      Top             =   1830
      Width           =   1125
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   510
      Width           =   2532
   End
   Begin VB.TextBox textCP14 
      Height          =   285
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   0
      Top             =   2160
      Width           =   732
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   510
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1500
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1170
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1500
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   4
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4920
      TabIndex        =   2
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6840
      TabIndex        =   3
      Top             =   30
      Width           =   1212
   End
   Begin MSForms.TextBox textCP64 
      Height          =   525
      Left            =   1170
      TabIndex        =   1
      Top             =   2490
      Width           =   7725
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13626;926"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_2 
      Height          =   285
      Left            =   1980
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1785
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3149;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1170
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1170
      Width           =   2532
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4466;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1170
      TabIndex        =   25
      Top             =   840
      Width           =   7485
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13203;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "法定期限 :"
      Height          =   255
      Index           =   4
      Left            =   4710
      TabIndex        =   21
      Top             =   2175
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限 :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   1845
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2490
      Width           =   975
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
      Left            =   3750
      TabIndex        =   18
      Top             =   562
      Width           =   645
   End
   Begin VB.Label Label27 
      Caption         =   "申請案號 :"
      Height          =   255
      Left            =   4710
      TabIndex        =   17
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2175
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   855
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1185
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   255
      Index           =   2
      Left            =   4710
      TabIndex        =   11
      Top             =   1515
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   255
      Index           =   3
      Left            =   4710
      TabIndex        =   10
      Top             =   1185
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   1515
      Width           =   975
   End
End
Attribute VB_Name = "frm03020411_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/13 改成Form2.0 ; cmbTM05、textTM23、textCP14_2、textCP64
'Create by Lydia 2018/01/10 FCT-修改承辦人
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 來函收文日
Dim m_CP05 As String
' 收文號
Dim m_CP09 As String
' 原案件性質
Dim m_CP10 As String
'Added by Lydia 2018/02/05 原智權人員
Dim m_CP13 As String
'原承辦人
Dim m_CP14 As String
'原進度備註
Dim m_CP64 As String
' 國家代碼
Dim m_TM10 As String
Dim m_Kind As String 'Added by Lydia 2018/02/05 1.修改承辦人 2.修改智權人員
Dim m_CP43 As String 'Added by Lydia 2018/10/02 相關總收文號
' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm03020411_03.Show
End Sub

Private Sub cmdExit_Click()
   Me.Enabled = False
   Unload frm03020411_03
   Unload frm03020411_02
   Unload frm03020411_01
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid = True Then
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload Me
      Unload frm03020411_03
      Unload frm03020411_02
      frm03020411_01.Show
   End If
End Sub

Private Sub cmdOpen_Click()
    Screen.MousePointer = vbHourglass
    frm100101_L.m_strKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
    frm100101_L.SetParent Me
    If frm100101_L.QueryData = True Then
       frm100101_L.Show
       Me.Hide
    Else
       Unload frm100101_L
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP06.BackColor = &H8000000F
   textCP07.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
  
   MoveFormToCenter Me
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
   End Select
End Sub

' 讀取商標基本檔
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
      
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         If m_TM10 < "010" Then
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
         Else
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
         End If
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
      End If
      '閉卷提示
      If IsNull(rsTmp.Fields("tm29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取案件進度檔
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      '本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then
         textCP06 = TAIWANDATE(rsTmp.Fields("CP06"))
      End If
      '法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then
         textCP07 = TAIWANDATE(rsTmp.Fields("CP07"))
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      
       m_CP14 = "" & rsTmp.Fields("CP14")  '預設承辦人
      'Added by Lydia 2018/02/05 增加一個月內C類來函1701註冊證和1001核准並且承辦人為「程序人員」之進度
       m_CP13 = "" & rsTmp.Fields("CP13")  '預設智權人員
      If InStr("1001,1701", m_CP10) > 0 Then
            m_Kind = "2"
            Me.textCP14.Text = m_CP13
            Me.textCP14_2.Text = GetStaffName(m_CP13)
            Me.Caption = "修改智權人員"
            Me.Label24.Caption = "智權人員:"
      Else
            m_Kind = "1"
            Me.Caption = "修改承辦人"
            Me.Label24.Caption = "承辦人:"
      'end 2018/02/05
            Me.textCP14.Text = m_CP14
            Me.textCP14_2.Text = GetStaffName(m_CP14)
      End If  'end 2018/02/05
      
      m_CP43 = "" & rsTmp.Fields("CP43") 'Added by Lydia 2018/10/02
      
      '進度備註
      m_CP64 = "" & rsTmp.Fields("CP64")
      Me.textCP64.Text = m_CP64
   End If
   rsTmp.Close
   Set rsTmp = Nothing

End Sub

Public Sub QueryData()

   ' 本所案號
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   
   ' 讀取商標基本檔
   QueryTradeMark
   
   ' 讀取案件進度檔
   QueryCaseProgress

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm03020411_04 = Nothing
End Sub

Private Function OnSaveData() As Boolean
   Dim strSql As String
   Dim strST03 As String 'Added by Lydia 2018/10/02
   
   OnSaveData = True

On Error GoTo CheckingErr
cnnConnection.BeginTrans

      strST03 = PUB_GetST03(textCP14) 'Added by Lydia 2018/10/02
      
      If m_Kind = "1" Then  '修改進度的承辦人 'Added by Lydia 2018/02/05 +判斷
            'Modified by Lydia 2018/02/02 修改進度的承辦人CP14,同時修改智權人員為相同CP13,ex.FCT-040434
            strSql = "UPDATE CaseProgress SET CP12='" & strST03 & "', CP13 = '" & textCP14 & "' , CP14 = '" & textCP14 & "' ,CP64=" & CNULL(ChgSQL(textCP64)) & _
                     " WHERE CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strSql, intI
            If m_CP14 <> textCP14 Then
                '更新下一程序
                strSql = "Update Nextprogress SET NP10='" & textCP14 & "' Where NP01='" & m_CP09 & "' and np06 is null and np10='" & m_CP14 & "' "
                cnnConnection.Execute strSql, intI
            End If
            'Added by Lydia 2018/10/02 若修改外商發文(722)的承辦人，一併更新C類相關總收文號和下一程序的智權人員。(ex.FCT-42719)
            If m_CP10 = "722" And Left(m_CP43, 1) = "C" Then
                 strSql = "Update caseprogress set CP12='" & strST03 & "', CP13 = '" & textCP14 & "' where cp09='" & m_CP43 & "' "
                 cnnConnection.Execute strSql, intI
                 strSql = "Update Nextprogress SET NP10='" & textCP14 & "' Where NP01='" & m_CP43 & "' and np06 is null "
                 cnnConnection.Execute strSql, intI
            End If
            'end 2018/10/02
      'Added by Lydia 2018/02/05 增加一個月內C類來函1701註冊證和1001核准並且承辦人為「程序人員」之進度，更改案件進度之智權人員及下一程序之智權人員。
      ElseIf m_Kind = "2" Then   '修改進度的智權人員
            strSql = "UPDATE CaseProgress SET CP12='" & strST03 & "', CP13 = '" & textCP14 & "'  ,CP64=" & CNULL(ChgSQL(textCP64)) & _
                     " WHERE CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strSql, intI
            If m_CP13 <> textCP14 Then
                '更新下一程序
                strSql = "Update Nextprogress SET NP10='" & textCP14 & "' Where NP01='" & m_CP09 & "' and np06 is null and np10='" & m_CP13 & "' "
                cnnConnection.Execute strSql, intI
            End If
      End If 'end 2018/02/05
      
  cnnConnection.CommitTrans
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
    OnSaveData = False
End Function

Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14)
      If IsEmptyText(textCP14_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         'Modified by Lydia 2018/02/05 +判斷
         'strMsg = "承辦人代號不存在"
         If m_Kind = "1" Then
             strMsg = "承辦人代號不存在"
         ElseIf m_Kind = "2" Then
             strMsg = "智權人員代號不存在"
         End If
         'end 2018/02/05
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP14_GotFocus
         textCP14.SetFocus 'Added by Lydia 2018/10/02
      End If
   End If
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "進度備註資料內容長度太長"
      textCP64_GotFocus
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim Cancel As Boolean 'Added by Lydia 2018/10/02
   
   CheckDataValid = False
   
   ' 承辦人不可空白
   If IsEmptyText(textCP14) = True Then
      strTit = "資料檢核"
      'Modified by Lydia 2018/02/05 +判斷
      'strMsg = "請輸入承辦人"
      If m_Kind = "1" Then
          strMsg = "請輸入承辦人"
      ElseIf m_Kind = "2" Then
          strMsg = "請輸入智權人員"
      End If
      'end 2018/02/05
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP14.SetFocus
      GoTo EXITSUB
   End If
   'Added by Lydia 2018/10/02 檢查輸入是否正確
   Call textCP14_Validate(Cancel)
   If Cancel = True Then
       GoTo EXITSUB
   End If
   'end 2018/10/02
   
   'Added by Lydia 2021/09/13 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
   End If
    
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub
