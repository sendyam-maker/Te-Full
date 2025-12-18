VERSION 5.00
Begin VB.Form frm12040121 
   BorderStyle     =   1  '單線固定
   Caption         =   "來文資料稽核表"
   ClientHeight    =   2760
   ClientLeft      =   1335
   ClientTop       =   1965
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5955
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   4956
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4128
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox textSys 
      Height          =   855
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   2
      Top             =   1392
      Width           =   4215
   End
   Begin VB.TextBox textDate_2 
      Height          =   270
      Left            =   3000
      MaxLength       =   7
      TabIndex        =   1
      Top             =   888
      Width           =   972
   End
   Begin VB.TextBox textDate_1 
      Height          =   270
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   0
      Top             =   888
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "－"
      Height          =   132
      Left            =   2640
      TabIndex        =   7
      Top             =   936
      Width           =   252
   End
   Begin VB.Label Label2 
      Caption         =   "系統類別："
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   1368
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日："
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   888
      Width           =   1212
   End
End
Attribute VB_Name = "frm12040121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
'2007/9/14 整理 by sonia
Option Explicit

Const m_CharWidth = 120
Const m_CharHeight = 240
Const m_PaperSize = "A4"
' 宣告報表表頭的欄位其資料型態
Private Type REPORTFIELD
   Name As String
   Left As Long
   Width As Long
End Type
' 表頭欄位的內容
Dim m_Field(8) As REPORTFIELD
' 報表左方留白的寬度
Dim m_LeftMargin As Integer
' 報表上方留白的高度
Dim m_TopMargin As Integer
' 報表頁首的高度
Dim m_HeaderHeight As Integer
' 報表文件的寬度
Dim m_ReportWidth As Integer
' 報表文件中可容納的資料列數
Dim m_ReportDataRows As Integer
' 開始日期
Dim m_DateFrom As String
' 結束日期
Dim m_DateTo As String
' 系統類別
'Dim m_Sys As String
' 預設印表機
'Dim m_DefaultPrinter As String
'Add By Cheng 2002/09/10
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub Form_Load()
   Dim Prn As Printer
   Dim nIndex As Integer
   Dim nSel As Integer
   
   MoveFormToCenter Me
   
   textSys = GetUserSystemKind()
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040121 = Nothing
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bData As Boolean
   
   If CheckDataValid() = True Then
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      
      'm_Sys = textSys
      If IsEmptyText(textDate_1) = False Then
         m_DateFrom = DBDATE(textDate_1)
      Else
         m_DateFrom = textDate_1
      End If
      If IsEmptyText(textDate_2) = False Then
         m_DateTo = DBDATE(textDate_2)
      Else
         m_DateTo = textDate_2
      End If
      ' 設定欄位規格
      BuildField
      
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/21 清除查詢印表記錄檔欄位
      ' 產生報表
      bData = GenerateReport
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      If bData = True Then
         strTit = "輸出報表"
         strMsg = "列印結束"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

' 搜尋資料
Private Function GenerateReport() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strSql As String
   Dim strSubSQL As String
   Dim strMRSys As String
   Dim strCPSys As String
   Dim strTemp As String
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim bData As Boolean

   GenerateReport = False
   bData = False
   
   If Trim(textDate_1) <> "" Or Trim(textDate_2) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1 & textDate_1 & "-" & textDate_2 'Add By Sindy 2010/12/21
   End If
   If Trim(textSys) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label2 & textSys 'Add By Sindy 2010/12/21
   End If
   
   strMRSys = Empty
   strCPSys = Empty
   For nIndex = 1 To GetSubStringCount(textSys)
      strTemp = GetSubString(textSys, nIndex)
      If IsEmptyText(strTemp) = False Then
         If IsEmptyText(strMRSys) = False Then
            strMRSys = strMRSys & "OR "
         End If
         If IsEmptyText(strCPSys) = False Then
            strCPSys = strCPSys & "OR "
         End If
         strMRSys = strMRSys & "MR12 = '" & strTemp & "' "
         strCPSys = strCPSys & "CP01 = '" & strTemp & "' "
      End If
   Next nIndex
    
   'Memo by Lydia 2022/02/23 每日批次strMenu111參考「來函記錄檔有而案件進度檔」，若有變更請兩邊檢查一下
   ' 來函記錄檔有而案件進度檔沒有的資料
   'strSQL = strSQL & "SELECT NVL(CP01,MR12) AS KEY1, NVL(CP02,MR13) AS KEY2, NVL(CP03,MR14) AS KEY3, NVL(CP04,MR15) AS KEY4,CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM "
   
   strSql = "SELECT NVL(CP01,MR12) AS KEY1, NVL(CP02,MR13) AS KEY2, NVL(CP03,MR14) AS KEY3, NVL(CP04,MR15) AS KEY4,NVL(MR02,CP05) AS KEY5,1 AS KEY6,CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM " & _
               "(SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM MAILREC, CASEPROGRESS " & _
                "WHERE MR12 = CP01(+) AND " & _
                      "MR13 = CP02(+) AND " & _
                      "MR14 = CP03(+) AND " & _
                      "MR15 = CP04(+) AND " & _
                      "MR02 = CP05(+) AND " & _
                      "(" & strMRSys & ") "
   If IsEmptyText(m_DateFrom) = False Then: strSql = strSql & "AND MR02 >= " & m_DateFrom & " "
   If IsEmptyText(m_DateTo) = False Then: strSql = strSql & "AND MR02 <= " & m_DateTo & " "
   strSql = strSql & ") "
   strSql = strSql & "WHERE CP09 IS NULL "
   
   strSql = strSql & "UNION ALL "
   
   ' 案件進度檔有而來函記錄檔沒有的資料
'Modify By Cheng 2002/12/09
'   strSQL = strSQL & "SELECT NVL(CP01,MR12) AS KEY1, NVL(CP02,MR13) AS KEY2, NVL(CP03,MR14) AS KEY3, NVL(CP04,MR15) AS KEY4,NVL(CP05,MR02) AS KEY5,2 AS KEY6,CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM " & _
'               "(SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM CASEPROGRESS, MAILREC " & _
'                "WHERE CP01 = MR12(+) AND " & _
'                      "CP02 = MR13(+) AND " & _
'                      "CP03 = MR14(+) AND " & _
'                      "CP04 = MR15(+) AND " & _
'                      "CP05 = MR02(+) AND " & _
'                      "CP09>'C' AND " & _
'                      "(" & strCPSys & ") "
   '2007/9/14 MODIFY BY SONIA T智慧局註冊費通知函不檢查,因櫃台不輸
   'strSQL = strSQL & "SELECT NVL(CP01,MR12) AS KEY1, NVL(CP02,MR13) AS KEY2, NVL(CP03,MR14) AS KEY3, NVL(CP04,MR15) AS KEY4,NVL(CP05,MR02) AS KEY5,2 AS KEY6,CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM " & _
   '            "(SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM CASEPROGRESS, MAILREC, SYSTEMKIND " & _
   '             "WHERE CP01 = MR12(+) AND " & _
   '                   "CP02 = MR13(+) AND " & _
   '                   "CP03 = MR14(+) AND " & _
   '                   "CP04 = MR15(+) AND " & _
   '                   "CP05 = MR02(+) AND " & _
   '                   "CP01 = SK01(+) AND " & _
   '                   "((SK02='1' OR SK02='2' OR SK02='5' OR SK02='6') AND CP10<>'1101') AND " & _
   '                   "CP09>'C' AND " & _
   '                   "(" & strCPSys & ") "
   '2012/12/21 modify by sonia 加1720通知繳納註冊費
   'Modify By Sindy 2013/1/3 "CP09>'C' AND CP01||CP10 NOT IN ('T1715','T1716','T1717','T1720') AND " ==> "CP09>'C' AND CP10 NOT IN ('1715','1716','1717','1720','1721','1722','1723') AND "
   'Modify By Sindy 2015/3/4 +排除1725通知期限
   'moidfy by sonia 2017/2/6 只抓C類,故加CP09<'D'
   'Modified by Morgan 2017/5/25 排除電子公文
   strSql = strSql & "SELECT NVL(CP01,MR12) AS KEY1, NVL(CP02,MR13) AS KEY2, NVL(CP03,MR14) AS KEY3, NVL(CP04,MR15) AS KEY4,NVL(CP05,MR02) AS KEY5,2 AS KEY6,CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM " & _
               "(SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM CASEPROGRESS, MAILREC, SYSTEMKIND " & _
                "WHERE CP01 = MR12(+) AND " & _
                      "CP02 = MR13(+) AND " & _
                      "CP03 = MR14(+) AND " & _
                      "CP04 = MR15(+) AND " & _
                      "CP05 = MR02(+) AND " & _
                      "CP01 = SK01(+) AND " & _
                      "((SK02='1' OR SK02='2' OR SK02='5' OR SK02='6') AND CP10<>'1101') AND " & _
                      "CP09>'C' AND CP09<'D' AND CP10 NOT IN ('1715','1716','1717','1720','1721','1722','1723','1725') AND not exists(select * from edocument where ed11=cp09) and " & _
                      "(" & strCPSys & ") "
   If IsEmptyText(m_DateFrom) = False Then: strSql = strSql & "AND CP05 >= " & m_DateFrom & " "
   If IsEmptyText(m_DateTo) = False Then: strSql = strSql & "AND CP05 <= " & m_DateTo & " "
   strSql = strSql & ") "
   strSql = strSql & "WHERE MR01 IS NULL "
   
   strSql = strSql & "UNION ALL "
   
   ' 案件進度檔及來函記錄檔有但本所期限不符的資料
'   strSQL = strSQL & "SELECT NVL(CP01,MR12) AS KEY1, NVL(CP02,MR13) AS KEY2, NVL(CP03,MR14) AS KEY3, NVL(CP04,MR15) AS KEY4,NVL(CP05,MR02) AS KEY5,3 AS KEY6,CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM " & _
'               "(SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM CASEPROGRESS, MAILREC " & _
'                "WHERE CP01 = MR12 AND " & _
'                      "CP02 = MR13 AND " & _
'                      "CP03 = MR14 AND " & _
'                      "CP04 = MR15 AND " & _
'                      "CP05 = MR02 AND " & _
'                      "SUBSTR(CP09,1,1) = 'C' AND " & _
'                      "(" & strMRSys & ") AND " & _
'                      "(" & strCPSys & ") "
   'moidfy by sonia 2017/2/6 只抓C類,故加CP09<'D'
   strSql = strSql & "SELECT NVL(CP01,MR12) AS KEY1, NVL(CP02,MR13) AS KEY2, NVL(CP03,MR14) AS KEY3, NVL(CP04,MR15) AS KEY4,NVL(CP05,MR02) AS KEY5,3 AS KEY6,CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM " & _
               "(SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP64,MR01,MR02,MR09,MR12,MR13,MR14,MR15,MR16,MR17 FROM CASEPROGRESS, MAILREC, SYSTEMKIND " & _
                "WHERE CP01 = MR12 AND " & _
                      "CP02 = MR13 AND " & _
                      "CP03 = MR14 AND " & _
                      "CP04 = MR15 AND " & _
                      "CP05 = MR02 AND " & _
                      "CP01 = SK01(+) AND " & _
                      "((SK02='1' OR SK02='2' OR SK02='5' OR SK02='6') AND CP10<>'1101') AND " & _
                      "CP09>'C' AND CP09<'D' AND " & _
                      "(" & strMRSys & ") AND " & _
                      "(" & strCPSys & ") "
   If IsEmptyText(m_DateFrom) = False Then: strSql = strSql & "AND MR02 >= " & m_DateFrom & " "
   If IsEmptyText(m_DateTo) = False Then: strSql = strSql & "AND MR02 <= " & m_DateTo & " "
   If IsEmptyText(m_DateFrom) = False Then: strSql = strSql & "AND CP05 >= " & m_DateFrom & " "
   If IsEmptyText(m_DateTo) = False Then: strSql = strSql & "AND CP05 <= " & m_DateTo & " "
   strSql = strSql & ") "
'   'strSQL = strSQL & "WHERE CP06 <> MR16 OR CP07 <> MR17 ), "
    'Modify By Cheng 2002/12/09
    '取消法定期限不同之判斷
'   strSQL = strSQL & "WHERE NVL(CP06,0) <> NVL(MR16,0) OR " & _
'                           "NVL(CP07,0) <> NVL(MR17,0) " & _
'                     "ORDER BY KEY6,KEY5,KEY1,KEY2,KEY3,KEY4 "
   strSql = strSql & " WHERE NVL(CP06,0) <> NVL(MR16,0)  " & _
                     " ORDER BY KEY6,KEY5,KEY1,KEY2,KEY3,KEY4 "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/12/21
      If PrintData(rsTmp) = True Then
         bData = True
      End If
   End If
   rsTmp.Close
   
   If bData = False Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/21
      strTit = "搜尋資料"
      strMsg = "資料庫中沒有符合的資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   
   Set rsTmp = Nothing
   GenerateReport = bData

End Function

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Cheng 2002/12/09
   Dim blnCancel As Boolean
    
   CheckDataValid = False
   blnClkSure = False
   
   ' 來函收文日(起)
   If IsEmptyText(textDate_1) = True Then
      strTit = "查詢資料"
      strMsg = "請輸入來函收文日起日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textDate_1.SetFocus
      GoTo EXITSUB
   End If
   
   ' 來函收文日(迄)
   If IsEmptyText(textDate_2) = True Then
      strTit = "查詢資料"
      strMsg = "請輸入來函收文日迄日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textDate_2.SetFocus
      GoTo EXITSUB
   End If

   'Add By Cheng 2002/03/21
   If PUB_CheckKeyInDate(Me.textDate_1) = -1 Then
      Me.textDate_1.SetFocus
      textDate_1_GotFocus
      GoTo EXITSUB
   End If
   If PUB_CheckKeyInDate(Me.textDate_2) = -1 Then
      Me.textDate_2.SetFocus
      textDate_2_GotFocus
      GoTo EXITSUB
   End If

   ' 系統類別
   If IsEmptyText(textSys) = True Then
      strTit = "查詢資料"
      strMsg = "請輸入系統類別"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSys.SetFocus
      GoTo EXITSUB
   End If
    'Add By Cheng 2002/12/09
    blnCancel = False
    textSys_Validate blnCancel
    If blnCancel Then GoTo EXITSUB
   
   ' 起日不可超過迄日
   If IsEmptyText(textDate_1) = False And IsEmptyText(textDate_2) = False Then
      If Val(DBDATE(textDate_1)) > Val(DBDATE(textDate_2)) Then
         strTit = "查詢資料"
         strMsg = "來函收文日起日不可超過迄日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         blnClkSure = True
         textDate_1.SetFocus
         textDate_1_GotFocus
         GoTo EXITSUB
      End If
   End If
   CheckDataValid = True
EXITSUB:
End Function

' 來函收文日起日
Private Sub textDate_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDate_1) = False Then
      If CheckIsTaiwanDate(textDate_1, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "來函收文日起日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDate_1_GotFocus
      End If
   End If
End Sub

Private Sub textDate_2_LostFocus()
   'Add By Cheng 2002/09/10
   If blnClkSure = False Then
      If Me.textDate_1.Text <> "" And Me.textDate_2.Text <> "" Then
         If Val(Me.textDate_1.Text) > Val(Me.textDate_2.Text) Then
            MsgBox "來函收文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.textDate_1.SetFocus
            textDate_1_GotFocus
            Exit Sub
         End If
      End If
   Else
      blnClkSure = False
   End If
End Sub

' 來函收文日迄日
Private Sub textDate_2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDate_2) = False Then
      If CheckIsTaiwanDate(textDate_2, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "來函收文日迄日格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDate_2_GotFocus
      End If
   End If
End Sub

Private Sub textDate_1_GotFocus()
   InverseTextBox textDate_1
End Sub

Private Sub textDate_2_GotFocus()
   InverseTextBox textDate_2
End Sub

Private Sub textSys_GotFocus()
   InverseTextBox textSys
End Sub

' 列印分隔線
Public Sub PrintSplitLine(ByVal nRow As Integer)
   Dim nCount As Integer
   For nCount = 0 To m_ReportWidth - 1
      Printer.CurrentX = (m_LeftMargin + nCount) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "-"
   Next nCount
End Sub

' 列印分隔線
Public Sub PrintTerminateLine(ByVal nRow As Integer)
   Dim nCount As Integer
   For nCount = 0 To m_ReportWidth - 1
      Printer.CurrentX = (m_LeftMargin + nCount) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "="
   Next nCount
End Sub

' 設定報表欄位的左方位置及其名稱
Public Sub BuildField()
   Dim nIndex As Integer
   Dim nFieldWidth As Integer
   Dim nLeft As Integer
   
   Select Case m_PaperSize
      Case "REPORT"
         m_LeftMargin = 1
         m_TopMargin = 3
         m_ReportWidth = 154
         m_ReportDataRows = 45
         nFieldWidth = 9
      Case Else
         m_LeftMargin = 5
         m_TopMargin = 2
         m_ReportWidth = 130
         m_ReportDataRows = 30
         nFieldWidth = 7
   End Select
   
   nLeft = m_LeftMargin
   For nIndex = 0 To 7
      m_Field(nIndex).Left = nLeft
      Select Case nIndex
         Case 0:
            m_Field(nIndex).Name = "來函收文日"
            m_Field(nIndex).Width = 10
         Case 1:
            m_Field(nIndex).Name = "收件號"
            m_Field(nIndex).Width = 12
         Case 2:
            m_Field(nIndex).Name = "本所案號"
            m_Field(nIndex).Width = 16
         Case 3:
            m_Field(nIndex).Name = "案件名稱"
            m_Field(nIndex).Width = 24
         Case 4:
            m_Field(nIndex).Name = "申請人"
            m_Field(nIndex).Width = 20
         Case 5:
            m_Field(nIndex).Name = "本所期限"
            m_Field(nIndex).Width = 10
         Case 6:
            m_Field(nIndex).Name = "法定期限"
            m_Field(nIndex).Width = 10
         Case 7:
            m_Field(nIndex).Name = "備註"
            m_Field(nIndex).Width = 30
      End Select
      nLeft = nLeft + m_Field(nIndex).Width
   Next nIndex
End Sub

' 列印表頭
Private Sub PrintPageHeader(ByVal nPage As Integer)
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim nRow As Integer
   Dim nX As Long
   Dim nY As Long
   Dim nCenter As Long
   Dim strTemp As String

   ' 表頭
   nRow = 1
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.FontSize = 24
   Printer.Font.Underline = True
   nX = m_LeftMargin + m_ReportWidth / 2 - 14
   Printer.CurrentX = nX * m_CharWidth
   Printer.Print "來文資料稽核表"
   
   Printer.Font.Underline = False
   ' 下二列
   nRow = nRow + 2
   Printer.FontSize = 12
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "列印人 : " & strUserName
      
   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "製表日期 : " & Format(Date, "EE/MM/DD")
   ' 下一列
   nRow = nRow + 1
   ' 系統別
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "系統別 : " & textSys
   
   ' 下一列
   nRow = nRow + 1
   
   ' 來函收文日
   'nX = m_LeftMargin + m_ReportWidth / 2 - 16
   nX = m_LeftMargin
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "來函收文日 : "
   ' 印日期的起迄
   nX = nX + 12
   If IsEmptyText(m_DateFrom) = False Then
      Printer.CurrentX = nX * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print Format(TAIWANDATE(m_DateFrom), "###/##/##")
   End If
   nX = nX + 9
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "-"
   nX = nX + 2
   If IsEmptyText(m_DateTo) = False Then
      Printer.CurrentX = nX * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print Format(TAIWANDATE(m_DateTo), "###/##/##")
   End If
   
   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "頁"
   
   nX = m_LeftMargin + m_ReportWidth - 14
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "次 : " & nPage
   
   ' 列印分隔線
   nRow = nRow + 1
   PrintSplitLine nRow

   nRow = nRow + 1
   For nIndex = 0 To 7
      Select Case nIndex
         Case 0, 1, 2, 3, 5, 6:
            nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
            strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
            Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
            Printer.Print strTemp
         Case Else:
            strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
            Printer.CurrentX = m_Field(nIndex).Left * m_CharWidth
            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
            Printer.Print strTemp
      End Select
   Next nIndex
     
   ' 列印分隔線
   nRow = nRow + 1
   PrintSplitLine nRow
   
   m_HeaderHeight = nRow
End Sub

' 列印資料
Private Function PrintData(ByRef rsTmp As ADODB.Recordset) As Boolean
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(8) As String
   Dim strKey1 As String
   Dim StrKey2 As String
   Dim strKey3 As String
   Dim strKey4 As String
   Dim strSql As String
   Dim rsSubTmp As New ADODB.Recordset
   Dim nType As Integer
   Dim nIndex As Integer
   Dim nCenter As Long
   Dim nLeft As Long
   Dim nRight As Long
   Dim bPrintHeader As Boolean
   
   PrintData = False
   bPrintHeader = False
      
   ' 紙張大小, 方向
   Select Case m_PaperSize
      Case "A4":
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
      Case "REPORT":
         Printer.PaperSize = vbPRPSFanfoldUS
      Case Else:
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
   End Select
   
   Printer.Font.Name = "新細明體"
      
   ' 印表頭
   nPage = 1
   'PrintPageHeader nPage
   
   nRow = 1
   Do While rsTmp.EOF = False
      ' 若列數超過頁面的高度限制時則換頁
      If nRow > m_ReportDataRows Then
         Printer.NewPage
         nPage = nPage + 1
         PrintPageHeader nPage
         nRow = 1
      End If
      
      strKey1 = rsTmp.Fields("KEY1")
      StrKey2 = rsTmp.Fields("KEY2")
      strKey3 = rsTmp.Fields("KEY3")
      strKey4 = rsTmp.Fields("KEY4")
      Select Case rsTmp.Fields("KEY1")
         ' 讀取商標基本檔
         Case "T", "TF", "CFT", "FCT":
            strSql = "SELECT TM05,TM06,TM07,TM10 AS NA01,NVL(CU04,TM23) AS CU04,CU05,CU06 FROM TRADEMARK,CUSTOMER " & _
                     "WHERE TM01 = '" & strKey1 & "' AND " & _
                           "TM02 = '" & StrKey2 & "' AND " & _
                           "TM03 = '" & strKey3 & "' AND " & _
                           "TM04 = '" & strKey4 & "' AND " & _
                           "SUBSTR(TM23,1,8) = CU01(+) AND " & _
                           "SUBSTR(TM23,9,1) = CU02(+) "
         ' 讀取專利基本檔
         Case "P", "CFP", "FCP":
            strSql = "SELECT PA05,PA06,PA07,PA09 AS NA01,CU04,CU05,CU06 FROM PATENT,CUSTOMER " & _
                     "WHERE PA01 = '" & strKey1 & "' AND " & _
                           "PA02 = '" & StrKey2 & "' AND " & _
                           "PA03 = '" & strKey3 & "' AND " & _
                           "PA04 = '" & strKey4 & "' AND " & _
                           "SUBSTR(PA26,1,8) = CU01(+) AND " & _
                           "SUBSTR(PA26,9,1) = CU02(+) "
         ' 讀取法務基本檔
         'Modify By Sindy 2009/07/24 增加LIN系統類別
         'modify by sonia 2019/7/29 +ACS系統類別
         Case "L", "CFL", "FCL", "LIN", "ACS":
            strSql = "SELECT LC05,LC06,LC07,LC15 AS NA01,CU04,CU05,CU06 FROM LAWCASE,CUSTOMER " & _
                     "WHERE LC01 = '" & strKey1 & "' AND " & _
                           "LC02 = '" & StrKey2 & "' AND " & _
                           "LC03 = '" & strKey3 & "' AND " & _
                           "LC04 = '" & strKey4 & "' AND " & _
                           "SUBSTR(LC11,1,8) = CU01(+) AND " & _
                           "SUBSTR(LC11,9,1) = CU02(+) "
         ' 讀取顧問案件基本檔
         Case "LA":
            strSql = "SELECT HC06,'010' AS NA01,CU04,CU05,CU06 FROM HIRECASE,CUSTOMER " & _
                     "WHERE HC01 = '" & strKey1 & "' AND " & _
                           "HC02 = '" & StrKey2 & "' AND " & _
                           "HC03 = '" & strKey3 & "' AND " & _
                           "HC04 = '" & strKey4 & "' AND " & _
                           "SUBSTR(HC05,1,8) = CU01(+) AND " & _
                           "SUBSTR(HC05,9,1) = CU02(+) "
         ' 讀取服務業務基本檔
         Case Else:
            strSql = "SELECT SP05,SP06,SP07,SP09 AS NA01,CU04,CU05,CU06 FROM SERVICEPRACTICE,CUSTOMER " & _
                     "WHERE SP01 = '" & strKey1 & "' AND " & _
                           "SP02 = '" & StrKey2 & "' AND " & _
                           "SP03 = '" & strKey3 & "' AND " & _
                           "SP04 = '" & strKey4 & "' AND " & _
                           "SUBSTR(SP08,1,8) = CU01(+) AND " & _
                           "SUBSTR(SP08,9,1) = CU02(+) "
      End Select
      rsSubTmp.CursorLocation = adUseClient
      rsSubTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      'Add By Cheng 2003/02/21
      '若基本檔有資料
      If rsSubTmp.RecordCount > 0 Then
        rsSubTmp.MoveFirst
        
        ' 申請國家為台灣的才要
        If IsNull(rsSubTmp.Fields("NA01")) = False Then
           If rsSubTmp.Fields("NA01") > "010" Then
              GoTo NextRecord
           End If
        End If
      End If
      ' 列印表頭
      If bPrintHeader = False Then
         PrintPageHeader nPage
         bPrintHeader = True
      End If
      
      ' 清除欄位
      For nIndex = 0 To 7: fld(nIndex) = Empty: Next nIndex
      
      If IsNull(rsTmp.Fields("MR01")) = False Then
         If IsNull(rsTmp.Fields("CP09")) = False Then
            ' 來函記錄檔與案件進度檔均有
            nType = 0
         Else
            ' 來函記錄檔有而案件進度檔沒有
            nType = 1
         End If
      Else
         ' 案件進度檔有而來函記錄檔沒有
         nType = 2
      End If

      Select Case nType
         Case 0: ' 來函記錄檔與案件進度檔均有
            ' 來函收文日
            'If IsNull(rsTmp.Fields("MR02")) = False Then: fld(0) = rsTmp.Fields("MR02")
            If IsNull(rsTmp.Fields("MR02")) = False Then
               fld(0) = TAIWANYEAR(rsTmp.Fields("MR02")) & "/" & TAIWANMONTH(rsTmp.Fields("MR02")) & "/" & TAIWANDAY(rsTmp.Fields("MR02"))
            End If
            ' 收件號
            If IsNull(rsTmp.Fields("MR01")) = False Then: fld(1) = rsTmp.Fields("MR01")
            ' 本所案號
            If IsNull(rsTmp.Fields("MR12")) = False Then
               Select Case rsTmp.Fields("MR12")
                  Case "TF":
                     fld(2) = rsTmp.Fields("MR12") & "-" & Mid(rsTmp.Fields("MR13"), 1, 5) & "-" & Mid(rsTmp.Fields("MR13"), 6, 1) & "-" & rsTmp.Fields("MR14") & "-" & rsTmp.Fields("MR15")
                  Case Else:
                     fld(2) = rsTmp.Fields("MR12") & "-" & rsTmp.Fields("MR13") & "-" & rsTmp.Fields("MR14") & "-" & rsTmp.Fields("MR15")
               End Select
            End If
            ' 案件名稱
            Select Case rsTmp.Fields("MR12")
               ' 讀取商標基本檔
               Case "T", "TF", "CFT", "FCT":
                  If IsNull(rsSubTmp.Fields("TM05")) = False Then
                     fld(3) = rsSubTmp.Fields("TM05")
                  ElseIf IsNull(rsSubTmp.Fields("TM06")) = False Then
                     fld(3) = rsSubTmp.Fields("TM06")
                  ElseIf IsNull(rsSubTmp.Fields("TM07")) = False Then
                     fld(3) = rsSubTmp.Fields("TM07")
                  End If
               ' 讀取專利基本檔
               Case "P", "CFP", "FCP":
                  If IsNull(rsSubTmp.Fields("PA05")) = False Then
                     fld(3) = rsSubTmp.Fields("PA05")
                  ElseIf IsNull(rsSubTmp.Fields("PA06")) = False Then
                     fld(3) = rsSubTmp.Fields("PA06")
                  ElseIf IsNull(rsSubTmp.Fields("PA07")) = False Then
                     fld(3) = rsSubTmp.Fields("PA07")
                  End If
               ' 讀取法務基本檔
               'Modify By Sindy 2009/07/24 增加LIN系統類別
               'modify by sonia 2019/7/29 +ACS系統類別
               Case "L", "CFL", "FCL", "LIN", "ACS":
                  If IsNull(rsSubTmp.Fields("LC05")) = False Then
                     fld(3) = rsSubTmp.Fields("LC05")
                  ElseIf IsNull(rsSubTmp.Fields("LC06")) = False Then
                     fld(3) = rsSubTmp.Fields("LC06")
                  ElseIf IsNull(rsSubTmp.Fields("LC07")) = False Then
                     fld(3) = rsSubTmp.Fields("LC07")
                  End If
               ' 讀取顧問案件基本檔
               Case "LA":
                  If IsNull(rsSubTmp.Fields("HC06")) = False Then: fld(3) = rsTmp.Fields("HC06")
               ' 讀取服務業務基本檔
               Case Else:
                  If IsNull(rsSubTmp.Fields("SP05")) = False Then
                     fld(3) = rsSubTmp.Fields("SP05")
                  ElseIf IsNull(rsSubTmp.Fields("SP06")) = False Then
                     fld(3) = rsSubTmp.Fields("SP06")
                  ElseIf IsNull(rsSubTmp.Fields("SP07")) = False Then
                     fld(3) = rsSubTmp.Fields("SP07")
                  End If
            End Select
            ' 申請人
            If IsNull(rsSubTmp.Fields("CU04")) = False Then
               fld(4) = rsSubTmp.Fields("CU04")
            End If
            If IsEmptyText(fld(4)) = True And IsNull(rsSubTmp.Fields("CU05")) = False Then
               fld(4) = rsSubTmp.Fields("CU05")
            End If
            If IsEmptyText(fld(4)) = True And IsNull(rsSubTmp.Fields("CU06")) = False Then
               fld(4) = rsSubTmp.Fields("CU06")
            End If
            ' 本所期限
            If IsNull(rsTmp.Fields("MR16")) = False Then: fld(5) = rsTmp.Fields("MR16")
            ' 法定期限
            If IsNull(rsTmp.Fields("MR17")) = False Then: fld(6) = rsTmp.Fields("MR17")
            ' 備註
            If IsNull(rsTmp.Fields("MR09")) = False Then: fld(7) = Mid(rsTmp.Fields("MR09"), 1, 20)
            
            ' 輸出
            Printer.FontSize = 10
            For nIndex = 0 To 7
               Select Case nIndex
                  Case 1, 2, 3, 4, 7:
                     If nIndex = 3 Or nIndex = 4 Or nIndex = 7 Then: Printer.FontSize = 9
                     Printer.CurrentX = m_Field(nIndex).Left * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nIndex), m_Field(nIndex).Width)
                     If nIndex = 3 Or nIndex = 4 Or nIndex = 7 Then: Printer.FontSize = 10
                  Case Else:
                     nLeft = m_Field(nIndex).Left + (m_Field(nIndex).Width / 2) - (StrLength(fld(nIndex)) / 2)
                     Printer.CurrentX = nLeft * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nIndex), m_Field(nIndex).Width)
               End Select
            Next nIndex
            Printer.FontSize = 12
                        
            ' 新增一列
            nRow = nRow + 1
            
            ' 清除欄位
            For nIndex = 0 To 7: fld(nIndex) = Empty: Next nIndex
            
            ' 來函收文日
            'If IsNull(rsTmp.Fields("CP05")) = False Then: fld(0) = rsTmp.Fields("CP05")
            If IsNull(rsTmp.Fields("CP05")) = False Then
               fld(0) = TAIWANYEAR(rsTmp.Fields("CP05")) & "/" & TAIWANMONTH(rsTmp.Fields("CP05")) & "/" & TAIWANDAY(rsTmp.Fields("CP05"))
            End If
            Select Case rsTmp.Fields("CP01")
               Case "TF":
                  fld(2) = rsTmp.Fields("CP01") & "-" & Mid(rsTmp.Fields("CP02"), 1, 5) & "-" & Mid(rsTmp.Fields("CP02"), 6, 1) & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
               Case Else:
                  fld(2) = rsTmp.Fields("CP01") & "-" & rsTmp.Fields("CP02") & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
            End Select
            ' 案件名稱
            Select Case rsTmp.Fields("CP01")
               ' 讀取商標基本檔
               Case "T", "TF", "CFT", "FCT":
                  If IsNull(rsSubTmp.Fields("TM05")) = False Then
                     fld(3) = rsSubTmp.Fields("TM05")
                  ElseIf IsNull(rsSubTmp.Fields("TM06")) = False Then
                     fld(3) = rsSubTmp.Fields("TM06")
                  ElseIf IsNull(rsSubTmp.Fields("TM07")) = False Then
                     fld(3) = rsSubTmp.Fields("TM07")
                  End If
               ' 讀取專利基本檔
               Case "P", "CFP", "FCP":
                  If IsNull(rsSubTmp.Fields("PA05")) = False Then
                     fld(3) = rsSubTmp.Fields("PA05")
                  ElseIf IsNull(rsSubTmp.Fields("PA06")) = False Then
                     fld(3) = rsSubTmp.Fields("PA06")
                  ElseIf IsNull(rsSubTmp.Fields("PA07")) = False Then
                     fld(3) = rsSubTmp.Fields("PA07")
                  End If
               ' 讀取法務基本檔
               'Modify By Sindy 2009/07/24 增加LIN系統類別
               'modify by sonia 2019/7/29 +ACS系統類別
               Case "L", "CFL", "FCL", "LIN", "ACS":
                  If IsNull(rsSubTmp.Fields("LC05")) = False Then
                     fld(3) = rsSubTmp.Fields("LC05")
                  ElseIf IsNull(rsSubTmp.Fields("LC06")) = False Then
                     fld(3) = rsSubTmp.Fields("LC06")
                  ElseIf IsNull(rsSubTmp.Fields("LC07")) = False Then
                     fld(3) = rsSubTmp.Fields("LC07")
                  End If
               ' 讀取顧問案件基本檔
               Case "LA":
                  If IsNull(rsSubTmp.Fields("HC06")) = False Then: fld(3) = rsTmp.Fields("HC06")
               ' 讀取服務業務基本檔
               Case Else:
                  If IsNull(rsSubTmp.Fields("SP05")) = False Then
                     fld(3) = rsSubTmp.Fields("SP05")
                  ElseIf IsNull(rsSubTmp.Fields("SP06")) = False Then
                     fld(3) = rsSubTmp.Fields("SP06")
                  ElseIf IsNull(rsSubTmp.Fields("SP07")) = False Then
                     fld(3) = rsSubTmp.Fields("SP07")
                  End If
            End Select
            ' 申請人
            If IsNull(rsSubTmp.Fields("CU04")) = False Then
               fld(4) = rsSubTmp.Fields("CU04")
            End If
            If IsEmptyText(fld(4)) = True And IsNull(rsSubTmp.Fields("CU05")) = False Then
               fld(4) = rsSubTmp.Fields("CU05")
            End If
            If IsEmptyText(fld(4)) = True And IsNull(rsSubTmp.Fields("CU06")) = False Then
               fld(4) = rsSubTmp.Fields("CU06")
            End If
            ' 本所期限
            If IsNull(rsTmp.Fields("CP06")) = False Then: fld(5) = rsTmp.Fields("CP06")
            ' 法定期限
            If IsNull(rsTmp.Fields("CP07")) = False Then: fld(6) = rsTmp.Fields("CP07")
            ' 備註
            If IsNull(rsTmp.Fields("CP64")) = False Then: fld(7) = Mid(rsTmp.Fields("CP64"), 1, 20)
            
            ' 輸出
            Printer.FontSize = 10
            For nIndex = 0 To 7
               Select Case nIndex
                  Case 1, 2, 3, 4, 7:
                     If nIndex = 3 Or nIndex = 4 Or nIndex = 7 Then: Printer.FontSize = 9
                     Printer.CurrentX = m_Field(nIndex).Left * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nIndex), m_Field(nIndex).Width)
                     If nIndex = 3 Or nIndex = 4 Or nIndex = 7 Then: Printer.FontSize = 10
                  Case Else:
                     nLeft = m_Field(nIndex).Left + (m_Field(nIndex).Width / 2) - (StrLength(fld(nIndex)) / 2)
                     Printer.CurrentX = nLeft * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nIndex), m_Field(nIndex).Width)
               End Select
            Next nIndex
            Printer.FontSize = 12
            
         Case 1: ' 來函記錄檔有而案件進度檔沒有
            ' 來函收文日
            'If IsNull(rsTmp.Fields("MR02")) = False Then: fld(0) = rsTmp.Fields("MR02")
            If IsNull(rsTmp.Fields("MR02")) = False Then
               fld(0) = TAIWANYEAR(rsTmp.Fields("MR02")) & "/" & TAIWANMONTH(rsTmp.Fields("MR02")) & "/" & TAIWANDAY(rsTmp.Fields("MR02"))
            End If
            ' 收件號
            If IsNull(rsTmp.Fields("MR01")) = False Then: fld(1) = rsTmp.Fields("MR01")
            ' 本所案號
            If IsNull(rsTmp.Fields("MR12")) = False Then
               Select Case rsTmp.Fields("MR12")
                  Case "TF":
                     fld(2) = rsTmp.Fields("MR12") & "-" & Mid(rsTmp.Fields("MR13"), 1, 5) & "-" & Mid(rsTmp.Fields("MR13"), 6, 1) & "-" & rsTmp.Fields("MR14") & "-" & rsTmp.Fields("MR15")
                  Case Else:
                     fld(2) = rsTmp.Fields("MR12") & "-" & rsTmp.Fields("MR13") & "-" & rsTmp.Fields("MR14") & "-" & rsTmp.Fields("MR15")
               End Select
            End If
            'Add By Cheng 2003/02/21
            fld(3) = "": fld(4) = ""
            If rsSubTmp.RecordCount > 0 Then
                ' 案件名稱
                Select Case rsTmp.Fields("MR12")
                   ' 讀取商標基本檔
                   Case "T", "TF", "CFT", "FCT":
                      If IsNull(rsSubTmp.Fields("TM05")) = False Then
                         fld(3) = rsSubTmp.Fields("TM05")
                      ElseIf IsNull(rsSubTmp.Fields("TM06")) = False Then
                         fld(3) = rsSubTmp.Fields("TM06")
                      ElseIf IsNull(rsSubTmp.Fields("TM07")) = False Then
                         fld(3) = rsSubTmp.Fields("TM07")
                      End If
                   ' 讀取專利基本檔
                   Case "P", "CFP", "FCP":
                      If IsNull(rsSubTmp.Fields("PA05")) = False Then
                         fld(3) = rsSubTmp.Fields("PA05")
                      ElseIf IsNull(rsSubTmp.Fields("PA06")) = False Then
                         fld(3) = rsSubTmp.Fields("PA06")
                      ElseIf IsNull(rsSubTmp.Fields("PA07")) = False Then
                         fld(3) = rsSubTmp.Fields("PA07")
                      End If
                   ' 讀取法務基本檔
                   'Modify By Sindy 2009/07/24 增加LIN系統類別
                   'modify by sonia 2019/7/29 +ACS系統類別
                   Case "L", "CFL", "FCL", "LIN", "ACS":
                      If IsNull(rsSubTmp.Fields("LC05")) = False Then
                         fld(3) = rsSubTmp.Fields("LC05")
                      ElseIf IsNull(rsSubTmp.Fields("LC06")) = False Then
                         fld(3) = rsSubTmp.Fields("LC06")
                      ElseIf IsNull(rsSubTmp.Fields("LC07")) = False Then
                         fld(3) = rsSubTmp.Fields("LC07")
                      End If
                   ' 讀取顧問案件基本檔
                   Case "LA":
                      If IsNull(rsSubTmp.Fields("HC06")) = False Then: fld(3) = rsTmp.Fields("HC06")
                   ' 讀取服務業務基本檔
                   Case Else:
                      If IsNull(rsSubTmp.Fields("SP05")) = False Then
                         fld(3) = rsSubTmp.Fields("SP05")
                      ElseIf IsNull(rsSubTmp.Fields("SP06")) = False Then
                         fld(3) = rsSubTmp.Fields("SP06")
                      ElseIf IsNull(rsSubTmp.Fields("SP07")) = False Then
                         fld(3) = rsSubTmp.Fields("SP07")
                      End If
                End Select
                ' 申請人
                If IsNull(rsSubTmp.Fields("CU04")) = False Then
                   fld(4) = rsSubTmp.Fields("CU04")
                End If
                If IsEmptyText(fld(4)) = True And IsNull(rsSubTmp.Fields("CU05")) = False Then
                   fld(4) = rsSubTmp.Fields("CU05")
                End If
                If IsEmptyText(fld(4)) = True And IsNull(rsSubTmp.Fields("CU06")) = False Then
                   fld(4) = rsSubTmp.Fields("CU06")
                End If
            End If
            ' 本所期限
            If IsNull(rsTmp.Fields("MR16")) = False Then: fld(5) = rsTmp.Fields("MR16")
            ' 法定期限
            If IsNull(rsTmp.Fields("MR17")) = False Then: fld(6) = rsTmp.Fields("MR17")
            ' 備註
            If IsNull(rsTmp.Fields("MR09")) = False Then: fld(7) = Mid(rsTmp.Fields("MR09"), 1, 20)
            
            ' 輸出
            Printer.FontSize = 10
            For nIndex = 0 To 7
               Select Case nIndex
                  Case 1, 2, 3, 4, 7:
                     If nIndex = 3 Or nIndex = 4 Or nIndex = 7 Then: Printer.FontSize = 9
                     Printer.CurrentX = m_Field(nIndex).Left * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nIndex), m_Field(nIndex).Width)
                     If nIndex = 3 Or nIndex = 4 Or nIndex = 7 Then: Printer.FontSize = 10
                  Case Else:
                     nLeft = m_Field(nIndex).Left + (m_Field(nIndex).Width / 2) - (StrLength(fld(nIndex)) / 2)
                     Printer.CurrentX = nLeft * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nIndex), m_Field(nIndex).Width)
               End Select
            Next nIndex
            Printer.FontSize = 12
         
         Case 2: ' 案件進度檔有而來函記錄檔沒有
            ' 來函收文日
            'If IsNull(rsTmp.Fields("CP05")) = False Then: fld(0) = rsTmp.Fields("CP05")
            If IsNull(rsTmp.Fields("CP05")) = False Then
               fld(0) = TAIWANYEAR(rsTmp.Fields("CP05")) & "/" & TAIWANMONTH(rsTmp.Fields("CP05")) & "/" & TAIWANDAY(rsTmp.Fields("CP05"))
            End If
            Select Case rsTmp.Fields("CP01")
               Case "TF":
                  fld(2) = rsTmp.Fields("CP01") & "-" & Mid(rsTmp.Fields("CP02"), 1, 5) & "-" & Mid(rsTmp.Fields("CP02"), 6, 1) & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
               Case Else:
                  fld(2) = rsTmp.Fields("CP01") & "-" & rsTmp.Fields("CP02") & "-" & rsTmp.Fields("CP03") & "-" & rsTmp.Fields("CP04")
            End Select
            ' 案件名稱
            Select Case rsTmp.Fields("CP01")
               ' 讀取商標基本檔
               Case "T", "TF", "CFT", "FCT":
                  If IsNull(rsSubTmp.Fields("TM05")) = False Then
                     fld(3) = rsSubTmp.Fields("TM05")
                  ElseIf IsNull(rsSubTmp.Fields("TM06")) = False Then
                     fld(3) = rsSubTmp.Fields("TM06")
                  ElseIf IsNull(rsSubTmp.Fields("TM07")) = False Then
                     fld(3) = rsSubTmp.Fields("TM07")
                  End If
               ' 讀取專利基本檔
               Case "P", "CFP", "FCP":
                  If IsNull(rsSubTmp.Fields("PA05")) = False Then
                     fld(3) = rsSubTmp.Fields("PA05")
                  ElseIf IsNull(rsSubTmp.Fields("PA06")) = False Then
                     fld(3) = rsSubTmp.Fields("PA06")
                  ElseIf IsNull(rsSubTmp.Fields("PA07")) = False Then
                     fld(3) = rsSubTmp.Fields("PA07")
                  End If
               ' 讀取法務基本檔
               'Modify By Sindy 2009/07/24 增加LIN系統類別
               'modify by sonia 2019/7/29 +ACS系統類別
               Case "L", "CFL", "FCL", "LIN", "ACS":
                  If IsNull(rsSubTmp.Fields("LC05")) = False Then
                     fld(3) = rsSubTmp.Fields("LC05")
                  ElseIf IsNull(rsSubTmp.Fields("LC06")) = False Then
                     fld(3) = rsSubTmp.Fields("LC06")
                  ElseIf IsNull(rsSubTmp.Fields("LC07")) = False Then
                     fld(3) = rsSubTmp.Fields("LC07")
                  End If
               ' 讀取顧問案件基本檔
               Case "LA":
                  If IsNull(rsSubTmp.Fields("HC06")) = False Then: fld(3) = rsTmp.Fields("HC06")
               ' 讀取服務業務基本檔
               Case Else:
                  If IsNull(rsSubTmp.Fields("SP05")) = False Then
                     fld(3) = rsSubTmp.Fields("SP05")
                  ElseIf IsNull(rsSubTmp.Fields("SP06")) = False Then
                     fld(3) = rsSubTmp.Fields("SP06")
                  ElseIf IsNull(rsSubTmp.Fields("SP07")) = False Then
                     fld(3) = rsSubTmp.Fields("SP07")
                  End If
            End Select
            ' 申請人
            If IsNull(rsSubTmp.Fields("CU04")) = False Then
               fld(4) = rsSubTmp.Fields("CU04")
            End If
            If IsEmptyText(fld(4)) = True And IsNull(rsSubTmp.Fields("CU05")) = False Then
               fld(4) = rsSubTmp.Fields("CU05")
            End If
            If IsEmptyText(fld(4)) = True And IsNull(rsSubTmp.Fields("CU06")) = False Then
               fld(4) = rsSubTmp.Fields("CU06")
            End If
            ' 本所期限
            If IsNull(rsTmp.Fields("CP06")) = False Then: fld(5) = rsTmp.Fields("CP06")
            ' 法定期限
            If IsNull(rsTmp.Fields("CP07")) = False Then: fld(6) = rsTmp.Fields("CP07")
            ' 備註
            If IsNull(rsTmp.Fields("CP64")) = False Then: fld(7) = Mid(rsTmp.Fields("CP64"), 1, 20)
            
            ' 輸出
            Printer.FontSize = 10
            For nIndex = 0 To 7
               Select Case nIndex
                  Case 1, 2, 3, 4, 7:
                     If nIndex = 3 Or nIndex = 4 Or nIndex = 7 Then: Printer.FontSize = 9
                     Printer.CurrentX = m_Field(nIndex).Left * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nIndex), m_Field(nIndex).Width)
                     If nIndex = 3 Or nIndex = 4 Or nIndex = 7 Then: Printer.FontSize = 10
                  Case Else:
                     nLeft = m_Field(nIndex).Left + (m_Field(nIndex).Width / 2) - (StrLength(fld(nIndex)) / 2)
                     Printer.CurrentX = nLeft * m_CharWidth
                     Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
                     Printer.Print LeftStr(fld(nIndex), m_Field(nIndex).Width)
               End Select
            Next nIndex
            Printer.FontSize = 12
         Case Else:
      End Select
      
      ' 列數加一
      nRow = nRow + 1
      
NextRecord:
      rsSubTmp.Close
      rsTmp.MoveNext
   Loop
   
   If bPrintHeader = True Then
      PrintTerminateLine m_HeaderHeight + nRow
      Printer.EndDoc
      PrintData = True
   End If

End Function

Private Sub textSys_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Public Function LeftStr(ByVal strData As String, ByVal nLen As Integer) As String
   LeftStr = StrConv(MidB(StrConv(strData, vbFromUnicode), 1, nLen), vbUnicode)
End Function

' 系統別
Private Sub textSys_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim strTemp As String
   Dim nIndex As Integer
   Dim nSubIndex As Integer
   Dim nCount As Integer
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textSys) = False Then
      nCount = GetSubStringCount(textSys)
      For nIndex = 1 To nCount
         strTemp = GetSubString(textSys, nIndex)
         If IsUserHasRightOfSystem(strUserNum, strTemp) = False Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "系統別<" & strTemp & ">不正確或是您沒有使用該系統別的權限"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSys_GotFocus
            GoTo EXITSUB
         End If
      Next nIndex

      For nIndex = 1 To nCount
         For nSubIndex = nIndex + 1 To nCount
            If GetSubString(textSys, nIndex) = GetSubString(textSys, nSubIndex) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "系統別<" & strTemp & ">資料重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSys_GotFocus
               GoTo EXITSUB
            End If
         Next nSubIndex
      Next nIndex
   End If
   
EXITSUB:
End Sub
