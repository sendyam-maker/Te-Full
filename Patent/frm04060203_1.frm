VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060203_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸專利公報查詢列印"
   ClientHeight    =   2940
   ClientLeft      =   510
   ClientTop       =   4620
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4920
   Begin VB.TextBox text01 
      Height          =   300
      Left            =   1740
      MaxLength       =   3
      TabIndex        =   0
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox text03_01 
      Height          =   300
      Left            =   1740
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1440
      Width           =   1212
   End
   Begin VB.TextBox text03_02 
      Height          =   300
      Left            =   3420
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1440
      Width           =   1212
   End
   Begin VB.TextBox text04 
      Height          =   300
      Left            =   1740
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1800
      Width           =   372
   End
   Begin VB.TextBox text05 
      Height          =   300
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   4
      Top             =   2160
      Width           =   1212
   End
   Begin VB.TextBox text06 
      Height          =   300
      Left            =   1740
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2520
      Width           =   372
   End
   Begin VB.CommandButton buttonExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4008
      TabIndex        =   7
      Top             =   96
      Width           =   800
   End
   Begin VB.CommandButton buttonOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3180
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin MSForms.TextBox text02 
      Height          =   300
      Left            =   1740
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2895
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5106;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   3060
      X2              =   3300
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "代理事務所："
      Height          =   252
      Left            =   300
      TabIndex        =   15
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "事務所名稱："
      Height          =   252
      Left            =   300
      TabIndex        =   14
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "公告日："
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   300
      TabIndex        =   13
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label Label4 
      Caption         =   "查詢方式："
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   300
      TabIndex        =   12
      Top             =   1800
      Width           =   1092
   End
   Begin VB.Label Label5 
      Caption         =   "起始公告號："
      Height          =   252
      Left            =   300
      TabIndex        =   11
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label Label6 
      Caption         =   "是否含明細："
      Height          =   252
      Left            =   300
      TabIndex        =   10
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label7 
      Caption         =   "(1:螢幕查詢  2:報表列印)"
      Height          =   255
      Left            =   2220
      TabIndex        =   9
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "(空白:含  N:不含)"
      Height          =   255
      Left            =   2220
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "frm04060203_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/18 改成Form2.0 ; text02 ; Printer列印未改
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Const m_CharWidth = 120
Const m_CharHeight = 240

Private Sub buttonExit_Click()
   Unload Me
End Sub
Private Sub buttonOK_Click()
   Dim bListDetail As Boolean
   
   If CheckDataValid = False Then
      GoTo EXITSUB
   End If
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
   bListDetail = True
   If text06 = "N" Or text06 = "n" Then
      bListDetail = False
      pub_QL05 = pub_QL05 & ";" & Label6 & "N:不含" 'Add By Sindy 2010/12/2
   Else
      pub_QL05 = pub_QL05 & ";" & Label6 & "空白:含" 'Add By Sindy 2010/12/2
   End If
   
   If text04 = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label4 & "1:螢幕查詢" 'Add By Sindy 2010/12/2
      ' 設定滑鼠游標成等待狀態
      Screen.MousePointer = vbHourglass
      frm04060203_2.SetData text01, text05, text03_01, text03_02, bListDetail
      frm04060203_2.UpdateCtrlData
      ' 設定滑鼠游標成預設
      Screen.MousePointer = vbDefault
      frm04060203_2.Show
   Else
      pub_QL05 = pub_QL05 & ";" & Label4 & "2:報表列印" 'Add By Sindy 2010/12/2
      PrintReport
   End If
EXITSUB:
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   UpdateState
End Sub

Public Sub UpdateState()
   text02.BackColor = &H8000000F
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm04060203_1 = Nothing
End Sub

Private Sub text01_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Dim rsTmp As ADODB.Recordset
   Cancel = False
   text02 = Empty
   If IsEmptyText(text01) = False Then
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT * FROM CAgent WHERE FNM01 = '" & text01 & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         text02 = rsTmp.Fields("FNM02")
      Else
         strMsg = "無此代理人資料"
         strTit = "錯誤"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text01_GotFocus
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Sub

Private Sub text03_01_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(text03_01) = False Then
      If CheckIsTaiwanDate(text03_01, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的公告日 !"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   Else
      strMsg = "公告日必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Cancel = True
   End If
   If Cancel Then TextInverse text03_01
End Sub

Private Sub text03_02_LostFocus()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   If IsEmpty(text03_02) = False Then
      If CheckIsTaiwanDate(text03_02, False) = False Then
         strMsg = "請輸入正確的公告日 !"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         text03_02.SetFocus
         TextInverse text03_02
      Else
         If Not ChkRange(text03_01, text03_02, "公告日") Then
         
         End If
      End If
   Else
      strMsg = "公告日必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text03_02.SetFocus
   End If
End Sub
' 轉換成大寫
Private Sub text04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 檢查查詢方式
Private Sub text04_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(text04) = False Then
      Select Case text04
         Case "1", "2":
         Case Else
            Cancel = True
            strMsg = "請輸入 1 或 2 !"
            strTit = "檢核輸入"
            nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
            text04_GotFocus
      End Select
   End If
End Sub

Private Sub text05_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 轉換成大寫
Private Sub text06_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 檢查是否包含明細
Private Sub text06_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmpty(text06) = False Then
      Select Case text06.Text
         Case " ", "N":
         Case Else
            Cancel = True
            strMsg = "請輸入空白或N"
            strTit = "檢核輸入"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            text06_GotFocus
      End Select
   End If
End Sub

Public Function CheckDataValid() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   CheckDataValid = True
   
   If IsEmptyText(text03_02) = True Then
      CheckDataValid = False
      strMsg = "公告日必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   If IsEmptyText(text03_01) = False And IsEmptyText(text03_02) = False Then
      If Val(text03_01) > Val(text03_02) Then
         CheckDataValid = False
         strMsg = "公告日範圍不正確"
         strTit = "檢核輸入"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   
   Select Case text04
      Case "1", "2":
      Case Else
         CheckDataValid = False
         strMsg = "請選擇查詢方式"
         strTit = "檢核輸入"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
   End Select
   
EXITSUB:
End Function

Public Sub PrintReport()
   Dim strSql As String
   Dim strSubSQL As String
   Dim rsData As New ADODB.Recordset
   Dim strMsg, strTit As String
   Dim nResponse
   Dim strCurr1, strCurr2 As String
   Dim nPage As Integer
   Dim Fld1, Fld2, Fld3, Fld4, Fld5 As String
   Dim nRow As Integer
   Dim nCount As Integer
   Dim arrayAmount(3) As Integer
   Dim nIndex As Integer
   
   ' 設定滑鼠游標成等待狀態
   Screen.MousePointer = vbHourglass
   
   'strSQL = "SELECT * FROM CPBulletin "
   strSql = "SELECT CPB01,CPB02,CPB03,CPB04,CPB05,CPB06,CPB07,CPB08,PA01,PA02,PA03,PA04,FNM02 FROM CPBulletin, PATENT, CAGENT "
   strSubSQL = Empty
   If IsEmpty(text01) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & "AND "
      strSubSQL = strSubSQL & "CPB06 = '" & text01 & "' "
      pub_QL05 = pub_QL05 & ";" & Label1 & text01 & text02 'Add By Sindy 2010/12/2
   End If
   If IsEmpty(text03_01) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & "AND "
      strSubSQL = strSubSQL & "CPB03 >= " & ChangeTStringToWString(text03_01) & " "
   End If
   If IsEmpty(text03_02) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & "AND "
      strSubSQL = strSubSQL & "CPB03 <= " & ChangeTStringToWString(text03_02) & " "
   End If
   If IsEmpty(text03_01) = False Or IsEmpty(text03_02) = False Then
      pub_QL05 = pub_QL05 & ";" & Label3 & text03_01 & "-" & text03_02 'Add By Sindy 2010/12/2
   End If
   If IsEmpty(text05) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & "AND "
      strSubSQL = strSubSQL & "CPB02 >= " & text05 & " "
      pub_QL05 = pub_QL05 & ";" & Label5 & text05 'Add By Sindy 2010/12/2
   End If
   If strSubSQL <> Empty Then
      strSql = strSql & " WHERE " & strSubSQL & " AND " & _
                                 "CPB01 = PA11(+) AND " & _
                                 "CPB06 = FNM01(+) "
   Else
      strSql = strSql & " WHERE CPB01 = PA11(+) AND " & _
                               "CPB06 = FNM01(+) "
   End If
   strSql = strSql & "ORDER BY CPB04, CPB05, CPB02 ASC"
   
   rsData.CursorLocation = adUseClient
   rsData.Open strSql, cnnConnection, adOpenDynamic
   
   If rsData.RecordCount <= 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
      strMsg = "無資料"
      strTit = "列印資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      rsData.Close
      GoTo EXITSUB
   Else
      InsertQueryLog (rsData.RecordCount) 'Add By Sindy 2010/12/2
   End If
   
PrintData:
   nPage = 1
   rsData.MoveFirst
   strCurr1 = rsData.Fields("CPB04")
   strCurr2 = rsData.Fields("CPB05")
   
   Printer.Orientation = vbPRORLandscape
   ' 印第一頁的表頭
   PrintPageHeader strCurr1, strCurr2, nPage
   
   ' 清除合計值
   For nIndex = 0 To 2
      arrayAmount(nIndex) = 0
   Next nIndex
   
   nRow = 1
   While rsData.EOF <> True
      Fld1 = rsData.Fields("CPB01")
      Fld2 = CStr(rsData.Fields("CPB02"))
      Fld3 = Empty
      Fld4 = Empty
      
      ' 若卷號或期數不同則換頁
      If rsData.Fields("CPB04") <> strCurr1 Or rsData.Fields("CPB05") <> strCurr2 Then
         strCurr1 = rsData.Fields("CPB04")
         strCurr2 = rsData.Fields("CPB05")
         ' 印合計值
         PrintPageTailer arrayAmount(0), arrayAmount(1), arrayAmount(2)
         ' 換頁
         Printer.NewPage
         ' 清除合計值
         For nIndex = 0 To 2
            arrayAmount(nIndex) = 0
         Next nIndex
         ' 頁碼更新
         nPage = nPage + 1
         ' 印表頭
         PrintPageHeader strCurr1, strCurr2, nPage
         nRow = 1
      End If
      
      ' 計算 發明, 新型, 設計的總數
      Select Case Mid(Fld1, 3, 1)
         Case "1":
            arrayAmount(0) = arrayAmount(0) + 1
         Case "2":
            arrayAmount(1) = arrayAmount(1) + 1
         Case "3":
            arrayAmount(2) = arrayAmount(2) + 1
      End Select
      
      ' 若列數超過 27 列則換頁
      If nRow > 27 Then
         Printer.NewPage
         nPage = nPage + 1
         PrintPageHeader strCurr1, strCurr2, nPage
         nRow = 1
      End If
      
      ' 代理事務所
      Fld3 = rsData.Fields("FNM02")
      
      ' 本所案號
      Fld4 = rsData.Fields("PA01") & "-" & rsData.Fields("PA02") & "-" & rsData.Fields("PA03") & "-" & rsData.Fields("PA04")
      
      Printer.CurrentX = 12 * m_CharWidth
      Printer.CurrentY = (nRow + 13) * m_CharHeight
      Printer.Print Fld1
      
      Printer.CurrentX = 32 * m_CharWidth
      Printer.CurrentY = (nRow + 13) * m_CharHeight
      Printer.Print Fld2
   
      Printer.CurrentX = 52 * m_CharWidth
      Printer.CurrentY = (nRow + 13) * m_CharHeight
      Printer.Print Fld3
   
      Printer.CurrentX = 82 * m_CharWidth
      Printer.CurrentY = (nRow + 13) * m_CharHeight
      Printer.Print Fld4
   
      nRow = nRow + 1
      rsData.MoveNext
   Wend
   
   ' 印合計值
   PrintPageTailer arrayAmount(0), arrayAmount(1), arrayAmount(2)
   ' 結束列印
   Printer.EndDoc
   
EXITSUB:
   ' 設定滑鼠游標成預設
   Screen.MousePointer = vbDefault
   Set rsData = Nothing
End Sub

Public Sub PrintPageHeader(ByVal Str1 As String, ByVal Str2 As String, ByVal nPage As Integer)
   Dim i As Integer
   Dim strDate1 As String
   Dim strDate2 As String
   
   strDate1 = text03_01
   strDate2 = text03_02
   If IsEmpty(strDate1) = True Then
      strDate1 = "        "
   Else
      strDate1 = ChangeTStringToTDateString(strDate1)
   End If
   If IsEmpty(strDate2) = True Then
      strDate2 = "        "
   Else
      strDate2 = ChangeTStringToTDateString(strDate2)
   End If
   
   ' 表頭
   Printer.CurrentX = 54 * m_CharWidth
   Printer.CurrentY = 5 * m_CharHeight
   Printer.FontSize = 24
   Printer.Font.Underline = True
   Printer.Print "大陸專利公報列印"
   
   Printer.CurrentX = 58 * m_CharWidth
   Printer.CurrentY = 8 * m_CharHeight
   Printer.FontSize = 12
   Printer.Font.Underline = False
   Printer.Print "公告日 : " & strDate1 & " - " & strDate2
   
   Printer.CurrentX = 12 * m_CharWidth
   Printer.CurrentY = 9 * m_CharHeight
   Printer.Print "列印人 : " & strUserName
   
   Printer.CurrentX = 101 * m_CharWidth
   Printer.CurrentY = 9 * m_CharHeight
   Printer.Print "製表日期 : " & Format(ChangeWStringToWDateString(GetTodayDate), "EE/MM/DD")
   
   Printer.CurrentX = 12 * m_CharWidth
   Printer.CurrentY = 10 * m_CharHeight
   Printer.Print Str1 & " 卷 " & Str2; " 號"
   ' 頁
   Printer.CurrentX = 101 * m_CharWidth
   Printer.CurrentY = 10 * m_CharHeight
   Printer.Print "頁"
   ' 次
   Printer.CurrentX = 107 * m_CharWidth
   Printer.CurrentY = 10 * m_CharHeight
   Printer.Print "次 : " & nPage
      
   For i = 0 To 107
      Printer.CurrentX = (i + 12) * m_CharWidth
      Printer.CurrentY = 11 * m_CharHeight
      Printer.Print "-"
   Next i
   
   Printer.CurrentX = 12 * m_CharWidth
   Printer.CurrentY = 12 * m_CharHeight
   Printer.Print "申請案號"
   
   Printer.CurrentX = 32 * m_CharWidth
   Printer.CurrentY = 12 * m_CharHeight
   Printer.Print "公告號"
   
   Printer.CurrentX = 52 * m_CharWidth
   Printer.CurrentY = 12 * m_CharHeight
   Printer.Print "代理事務所"
   
   Printer.CurrentX = 82 * m_CharWidth
   Printer.CurrentY = 12 * m_CharHeight
   Printer.Print "本所案號"
   
   For i = 0 To 107
      Printer.CurrentX = (i + 12) * m_CharWidth
      Printer.CurrentY = 13 * m_CharHeight
      Printer.Print "-"
   Next i
   
End Sub

Public Sub PrintPageTailer(ByVal value1 As Integer, ByVal value2 As Integer, ByVal value3 As Integer)
   Dim nIndex As Integer
   
   For nIndex = 0 To 107
      Printer.CurrentX = (nIndex + 12) * m_CharWidth
      Printer.CurrentY = 41 * m_CharHeight
      Printer.Print "-"
   Next nIndex
   
   Printer.CurrentX = 12 * m_CharWidth
   Printer.CurrentY = 42 * m_CharHeight
   Printer.Print "發明合計 : " & value1 & "            " & _
                 "新型合計 : " & value2 & "            " & _
                 "設計合計 : " & value3 & "            " & _
                 "合計 : " & value1 + value2 + value3
   
End Sub

' 判斷資料是否為空的
Public Function IsEmpty(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   IsEmpty = False
   
   If Len(strData) <= 0 Then
      IsEmpty = True
   Else
      IsEmpty = True
      For nIndex = 1 To Len(strData)
         If Mid(strData, nIndex, 1) <> " " Then
            IsEmpty = False
            Exit For
         End If
      Next nIndex
   End If
End Function

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As TextBox)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Private Sub text01_GotFocus()
   InverseAll text01
End Sub

Private Sub text03_01_GotFocus()
   InverseAll text03_01
End Sub

Private Sub text03_02_GotFocus()
   InverseAll text03_02
End Sub

Private Sub text04_GotFocus()
   InverseAll text04
End Sub

Private Sub text05_GotFocus()
   InverseAll text05
End Sub

Private Sub text06_GotFocus()
   InverseAll text06
End Sub

Public Sub SetInputCPB06()
   text01.SetFocus
End Sub
