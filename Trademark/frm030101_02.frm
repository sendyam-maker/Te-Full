VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030101_02 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文"
   ClientHeight    =   5160
   ClientLeft      =   6348
   ClientTop       =   1872
   ClientWidth     =   8124
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8124
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "繼續發文(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6864
      TabIndex        =   22
      Top             =   48
      Width           =   1212
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "E-Mail(S)"
      Height          =   400
      Left            =   5880
      TabIndex        =   21
      Top             =   48
      Width           =   912
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印聯絡單(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4440
      TabIndex        =   20
      Top             =   60
      Width           =   1392
   End
   Begin VB.TextBox textCP18 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4620
      Width           =   2532
   End
   Begin VB.TextBox textCP16 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4260
      Width           =   2532
   End
   Begin VB.TextBox textCP17 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3180
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3540
      Width           =   2532
   End
   Begin VB.TextBox textCP10_2 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1692
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2820
      Width           =   732
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2460
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2532
   End
   Begin MSForms.TextBox textTM07 
      Height          =   285
      Left            =   1440
      TabIndex        =   26
      Top             =   2100
      Width           =   6495
      VariousPropertyBits=   671105055
      Size            =   "11456;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM06 
      Height          =   285
      Left            =   1440
      TabIndex        =   25
      Top             =   1740
      Width           =   6495
      VariousPropertyBits=   671105055
      Size            =   "11456;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   285
      Left            =   1440
      TabIndex        =   24
      Top             =   1380
      Width           =   6495
      VariousPropertyBits=   671105055
      Size            =   "11456;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   1440
      TabIndex        =   23
      Top             =   3900
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "點數 :"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   4620
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "費用 :"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   4260
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   120
      TabIndex        =   14
      Top             =   3900
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   252
      Index           =   10
      Left            =   120
      TabIndex        =   13
      Top             =   3540
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "規費 :"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   3180
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   2820
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "此程序未收款, 是否向智權人員發出E-Mail"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   6732
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   252
      Index           =   8
      Left            =   120
      TabIndex        =   6
      Top             =   2460
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   1380
      Width           =   1332
   End
   Begin VB.Label Label4 
      Caption         =   "案件英文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   1740
      Width           =   1212
   End
   Begin VB.Label Label5 
      Caption         =   "案件日文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   2100
      Width           =   1452
   End
End
Attribute VB_Name = "frm030101_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Lydia 2021/08/10 改成Form2.0 ; textCP13、textTM05、textTM06、textTM07
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 收文號
Dim m_CP09 As String
' 本所期限
Dim m_CP06 As String
' 法定期限
Dim m_CP07 As String
' 申請國家
Dim m_TM10 As String

Private Sub cmdEmail_Click()
Dim dblAmt As Double, dblPFee As Double, dblTFee As Double   '2011/7/28 ADD BY SONIA
Dim m_StrTo As String, m_StrSub As String, m_StrCont As String 'Added by Lydia 2022/05/30 整理frm880005改用寄信模組

   GetBillData PUB_GetCustNo(m_TM01 & m_TM02 & m_TM03 & m_TM04), dblAmt, dblPFee, dblTFee  '2011/7/28 ADD BY SONIA 抓關係企業已發文應收金額
   'Added by Lydia 2016/10/07 傳預設收件人,並增加副本收件人選項
   If Trim(textCP13.Tag) <> "" Then
      'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
      'frm880005.txtEmail(0) = Trim(textCP13.Tag)
      'frm880005.bolCCList = True
      m_StrTo = Trim(textCP13.Tag)
   End If
   'end 2016/10/07
   
   'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
   'frm880005.txtEmail(1) = textTMKey + " 請儘快收款以便發文"
   ''2011/7/28 modify by sonia 加申請人,申請國家,含關係企業已發文未收金額
   'frm880005.txtEmail(2) = "智權人員姓名：" + textCP13 + vbCrLf & _
                           "本所案號：" + textTMKey + vbCrLf & _
                           "案件名稱(中)：" + textTM05 + vbCrLf & _
                           "申請人　：" + PUB_GetCustName(m_TM01 & m_TM02 & m_TM03 & m_TM04) + vbCrLf & _
                           "申請國家：" + textTM10 + vbCrLf & _
                           "收文日：" + textCP05 + vbTab + vbTab & _
                           "案件性質：" + textCP10_2 + vbCrLf & _
                           "費用：" + textCP16 + vbTab + vbTab & _
                           "規費：" + textCP17 + vbTab + vbTab & _
                           "點數：" + textCP18 + vbCrLf & _
                           "本所期限：" + ChangeWStringToTString(m_CP06) + vbTab + vbTab + vbTab & _
                           "法定期限：" + ChangeWStringToTString(m_CP07) + vbCrLf + vbCrLf & _
                           "含關係企業已發文未收金額：" + Format(dblAmt, "#,##0") + vbCrLf + vbCrLf & _
                           "此案件已可發文但尚未收款，請儘快收款以便發文。"
   'frm880005.Show vbModal
   ' 設定完EMail後回前畫面
   'If frm880005.bolLeave Then
   '   Unload Me
   '   frm030101_01.Show
   'End If
   m_StrSub = textTMKey + " 請儘快收款以便發文"
   m_StrCont = "智權人員姓名：" + textCP13 + vbCrLf & _
                           "本所案號：" + textTMKey + vbCrLf & _
                           "案件名稱(中)：" + textTM05 + vbCrLf & _
                           "申請人　：" + PUB_GetCustName(m_TM01 & m_TM02 & m_TM03 & m_TM04) + vbCrLf & _
                           "申請國家：" + textTM10 + vbCrLf & _
                           "收文日：" + textCP05 + vbTab + vbTab & _
                           "案件性質：" + textCP10_2 + vbCrLf & _
                           "費用：" + textCP16 + vbTab + vbTab & _
                           "規費：" + textCP17 + vbTab + vbTab & _
                           "點數：" + textCP18 + vbCrLf & _
                           "本所期限：" + ChangeWStringToTString(m_CP06) + vbTab + vbTab + vbTab & _
                           "法定期限：" + ChangeWStringToTString(m_CP07) + vbCrLf + vbCrLf & _
                           "含關係企業已發文未收金額：" + Format(dblAmt, "#,##0") + vbCrLf + vbCrLf & _
                           "此案件已可發文但尚未收款，請儘快收款以便發文。"
   PUB_SendMail strUserNum, m_StrTo, m_CP09, m_StrSub, m_StrCont
   frm030101_01.Show  '回前畫面
   'end 2022/05/30
End Sub

Private Sub cmdOK_Click()
   Unload Me
   ' 呼叫前一畫面的顯示下一個畫面的功能
   frm030101_01.DisplayNextForm
End Sub

Private Sub cmdPrint_Click()
   'edit by nickc 2007/02/06 不用 dll 了
   'Dim objPrintDllPublic As Object
   Dim intCaseKind As Integer
   'edit by nickc 2007/02/06 不用 dll 了
   'If objPublicData.GetSystemKind(m_TM01, intCaseKind) Then
   '   Set objPrintDllPublic = CreateObject("prjPrintDllPublic.clsPrintPublic")
   '   objPrintDllPublic.PrintEmail intCaseKind, intPWhere, m_CP09, strUserName
   '   Set objPrintDllPublic = Nothing
   If ClsPDGetSystemKind(m_TM01, intCaseKind) Then
      ClsPPPrintEmail intCaseKind, intPWhere, m_CP09, strUserName, , textTM10, PUB_GetCustName(m_TM01 & m_TM02 & m_TM03 & m_TM04)
   End If
   ' 印完回前一畫面
   Unload Me
   frm030101_01.Show
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM05.BackColor = &H8000000F
   textTM06.BackColor = &H8000000F
   textTM07.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP10_2.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP16.BackColor = &H8000000F
   textCP17.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
  
   MoveFormToCenter Me
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 收文號
      Case 0: m_CP09 = strData
   End Select
End Sub
' 由案件性質代碼取得案件性質名稱
Private Function GetCaseType(ByVal strKey1 As String, ByVal StrKey2 As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCaseType = Empty
   If IsEmptyText(strKey1) = False And IsEmptyText(StrKey2) = False Then
      strSql = "SELECT * FROM CasePropertyMap " & _
               "WHERE CPM01 = '" & strKey1 & "' AND " & _
                     "CPM02 = '" & StrKey2 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CPM03")) = False Then
            GetCaseType = rsTmp.Fields("CPM03")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function
' 由客戶代碼取得客戶名稱
Private Function GetCustomer(ByVal strData As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCustomer = Empty
   If IsEmptyText(strData) = False Then
      Set rsTmp = New ADODB.Recordset
      If Len(strData) > 8 Then
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                        "CU02 = '" & Mid(strData, 9, 1) & "'"
      Else
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CU04")) = False Then
            GetCustomer = rsTmp.Fields("CU04")
         ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
            GetCustomer = rsTmp.Fields("CU05")
         ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
            GetCustomer = rsTmp.Fields("CU06")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function
' 取得國家的名稱
Private Function GetNation(ByVal strNation As String, ByRef nNA14 As Integer) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetNation = Empty
   nNA14 = 0
   If IsEmptyText(strNation) = False Then
      strSql = "SELECT * FROM NATION " & _
               "WHERE NA01 = '" & strNation & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("NA03")) = False Then
            GetNation = rsTmp.Fields("NA03")
         End If
         ' 延展年度
         If IsNull(rsTmp.Fields("NA14")) = False Then
            nNA14 = rsTmp.Fields("NA14")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

' 取得員工姓名
'Remove by Lydia 2016/10/07
'Private Function GetStaffName(ByVal strKey As String) As String
'   Dim rsTmp As New ADODB.Recordset
'   Dim strSql As String
'
'   GetStaffName = Empty
'   strSql = "SELECT * FROM Staff " & _
'            "WHERE ST01 = '" & strKey & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      If IsNull(rsTmp.Fields("ST02")) = False Then
'         GetStaffName = rsTmp.Fields("ST02")
'      End If
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'End Function

' 讀取商標基本檔的相關欄位
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 取得案件進度檔及商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      'Add By Cheng 2002/07/18
      m_TM10 = Empty
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05 = rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         textTM06 = rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         textTM07 = rsTmp.Fields("TM07")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取服務業務基本檔的相關欄位
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 取得案件進度檔及商標基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      'Add By Cheng 2002/07/18
      m_TM10 = Empty
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textTM05 = rsTmp.Fields("SP05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textTM06 = rsTmp.Fields("SP06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textTM07 = rsTmp.Fields("SP07")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取案件進度檔
Private Sub QueryCaseProgress()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
   
   ' 取得案件進度檔的相關項目
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         textCP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10_2 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10_2 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = ChangeWStringToTString(rsTmp.Fields("CP05"))
      End If
      ' 本所期限
      'Add By Cheng 2002/07/18
      m_CP06 = Empty
      If IsNull(rsTmp.Fields("CP06")) = False Then
         m_CP06 = rsTmp.Fields("CP06")
      End If
      ' 法定期限
      'Add By Cheng 2002/07/18
      m_CP07 = Empty
      If IsNull(rsTmp.Fields("CP07")) = False Then
         m_CP07 = rsTmp.Fields("CP07")
      End If
      ' 規費
      If IsNull(rsTmp.Fields("CP17")) = False Then
         textCP17 = rsTmp.Fields("CP17")
      End If
      ' 費用
      If IsNull(rsTmp.Fields("CP16")) = False Then
         textCP16 = rsTmp.Fields("CP16")
      End If
      ' 點數
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      '2014/2/25 add by sonia 應扣除銷帳 CFP-026595
      StrSQLa = "Select NVL(SUM(A1U07),0) A1U07,NVL(SUM(A1U09),0) A1U09 From ACC1U0 Where A1U03 = '" & m_CP09 & "' GROUP BY A1U03"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         textCP16 = Val("" & rsTmp.Fields("CP16")) - Val(rsA.Fields(0).Value) - Val(rsA.Fields(1).Value)
         textCP17 = Val("" & rsTmp.Fields("CP17")) - Val(rsA.Fields(1).Value)
         textCP18 = (Val(textCP16) - Val(textCP17)) / 1000
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      '2014/2/25 end
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         'Memo by Lydia 2016/10/07 原本有自設模組GetStaffName,改用共用模組
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
         'Added by Lydia 2016/10/07
         If Trim(textCP13.Text) <> "" Then
            textCP13.Tag = rsTmp.Fields("CP13")
         Else '員工已離職
            textCP13.Tag = PUB_GetAKindSalesNo(rsTmp.Fields("CP01"), rsTmp.Fields("CP02"), rsTmp.Fields("CP03"), rsTmp.Fields("CP04"))
            textCP13.Text = GetStaffName(textCP13.Tag)
         End If
         'end 2016/10/07
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 取得案件進度檔及商標基本檔的相關項目
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 本所案號
      If IsNull(rsTmp.Fields("CP01")) = False Then: m_TM01 = rsTmp.Fields("CP01")
      If IsNull(rsTmp.Fields("CP02")) = False Then: m_TM02 = rsTmp.Fields("CP02")
      If IsNull(rsTmp.Fields("CP03")) = False Then: m_TM03 = rsTmp.Fields("CP03")
      If IsNull(rsTmp.Fields("CP04")) = False Then: m_TM04 = rsTmp.Fields("CP04")
   End If
   rsTmp.Close
   
   ' 顯示本所案號
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   
   ' 讀取商標基本檔或服務業務基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "CFT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   
   '91.12.22 ADD BY SONIA
   textTM10 = GetNationName(m_TM10, 0)
   '91.12.22 END
   
   ' 讀取案件進度檔
   QueryCaseProgress
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030101_02 = Nothing
End Sub
