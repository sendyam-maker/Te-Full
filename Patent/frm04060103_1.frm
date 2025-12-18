VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060103_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利公報查詢列印"
   ClientHeight    =   3495
   ClientLeft      =   150
   ClientTop       =   1950
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5280
   Begin VB.CheckBox Check1 
      Caption         =   "是否只顯示本所案件"
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   1995
   End
   Begin VB.CommandButton buttonOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3480
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton buttonExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4305
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox text06 
      Height          =   300
      Left            =   1620
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2520
      Width           =   492
   End
   Begin VB.TextBox text05 
      Height          =   300
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   4
      Top             =   2154
      Width           =   972
   End
   Begin VB.TextBox text04 
      Height          =   300
      Left            =   1620
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1788
      Width           =   492
   End
   Begin VB.TextBox text03_02 
      Height          =   300
      Left            =   3060
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1422
      Width           =   972
   End
   Begin VB.TextBox text03_01 
      Height          =   300
      Left            =   1620
      MaxLength       =   8
      TabIndex        =   1
      Top             =   1422
      Width           =   972
   End
   Begin VB.TextBox text01_01 
      Height          =   300
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   0
      Top             =   690
      Width           =   732
   End
   Begin MSForms.TextBox text01_02 
      Height          =   285
      Left            =   2430
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   690
      Width           =   1935
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "3408;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox text02 
      Height          =   300
      Left            =   1620
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1056
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
   Begin VB.Label Label6 
      Caption         =   "是否含明細："
      Height          =   252
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label8 
      Caption         =   "(空白:含  N:不含)"
      Height          =   252
      Left            =   2250
      TabIndex        =   15
      Top             =   2520
      Width           =   1812
   End
   Begin VB.Line Line1 
      X1              =   2700
      X2              =   2940
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label7 
      Caption         =   "(1:螢幕查詢  2:報表列印)"
      Height          =   255
      Left            =   2250
      TabIndex        =   14
      Top             =   1788
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "起始公告號："
      Height          =   252
      Left            =   240
      TabIndex        =   13
      Top             =   2154
      Width           =   1092
   End
   Begin VB.Label Label4 
      Caption         =   "查詢方式："
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   1788
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "公告日："
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   240
      TabIndex        =   11
      Top             =   1422
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "事務所名稱："
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   1056
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   690
      Width           =   1092
   End
End
Attribute VB_Name = "frm04060103_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/18 改成Form2.0 ; text01_02、text02 ; Printer列印未改
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Const m_CharWidth = 120
Const m_CharHeight = 240
Dim bListDetail As Boolean
Dim nRow As Integer
Dim nRowTot As Double 'Add By Sindy 2011/12/16


Private Sub Form_Load()
   MoveFormToCenter Me
   UpdateState
End Sub

Private Sub buttonExit_Click()
   Unload Me
End Sub
' 清除畫面中的欄位
Private Sub ClearFields()
   text01_01 = Empty
   text01_02 = Empty
   text02 = Empty
   text03_01 = Empty
   text03_02 = Empty
   text04 = Empty
   text05 = Empty
   text06 = Empty
   Check1.Value = 0 'Add By Sindy 2012/1/10
End Sub

Private Sub buttonOK_Click()
   
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
      frm04060103_2.SetData text01_01, text05, text03_01, text03_02, bListDetail
      frm04060103_2.UpdateCtrlData
      frm04060103_2.Show
      frm04060103_1.Hide
   Else
      pub_QL05 = pub_QL05 & ";" & Label4 & "2:報表列印" 'Add By Sindy 2010/12/2
      PrintReport
   End If
   ClearFields
EXITSUB:
End Sub

Public Sub UpdateState()
   text01_02.BackColor = &H8000000F
   text02.BackColor = &H8000000F
End Sub

Private Function GetNation(ByVal strNation As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM NATION " & _
            "WHERE NA01 = '" & strNation & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly 'edit by nickc 2007/02/06 , adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      GetNation = rsTmp.Fields("NA03")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm04060103_1 = Nothing
End Sub

Private Sub text01_01_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Dim rsTmp As ADODB.Recordset
   
   text01_02 = Empty
   text02 = Empty
   
   Cancel = False
   If IsEmpty(text01_01) = False Then
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT * FROM TAGENT WHERE TA01 = 'P' AND TA02 = '" & text01_01 & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         text01_02 = rsTmp.Fields("TA03")
         text02 = rsTmp.Fields("TA04")
      Else
         Cancel = True
         strMsg = "無此代理人資料"
         strTit = "錯誤"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Sub
'Remove by Morgan 2006/9/19 確定時再檢查就好
'Private Sub text03_01_Validate(Cancel As Boolean)
'   Dim strMsg As String
'   Dim strTit As String
'   Dim nResponse
'   Cancel = False
'   If IsEmpty(text03_01) = False Then
'      If CheckIsTaiwanDate(text03_01, False) = False Then
'         Cancel = True
'         strMsg = "請輸入正確的公告日 !"
'         strTit = "資料檢核"
'         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
'      End If
'   Else
'      Cancel = True
'      strMsg = "公告日必須輸入"
'      strTit = "檢核輸入"
'      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
'   End If
'   If Cancel Then TextInverse text03_01
'End Sub
'
'Private Sub text03_02_LostFocus()
'   Dim strMsg As String
'   Dim strTit As String
'   Dim nResponse
'   If IsEmpty(text03_02) = False Then
'      If CheckIsTaiwanDate(text03_02, False) = False Then
'         strMsg = "請輸入正確的公告日 !"
'         strTit = "資料檢核"
'         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
'         text03_02.SetFocus
'         TextInverse text03_02
'      Else
'         If Not ChkRange(text03_01, text03_02, "公告日") Then
'
'         End If
'      End If
'   Else
'      strMsg = "公告日必須輸入"
'      strTit = "檢核輸入"
'      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
'      text03_02.SetFocus
'   End If
'End Sub

Private Sub text04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

' 檢查查詢方式
Private Sub text04_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmpty(text04) = False Then
      Select Case text04
         Case "1", "2":
         Case Else
            strMsg = "請輸入1或2 !"
            strTit = "檢核輸入"
            nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
            Cancel = True
            TextInverse text04
      End Select
   End If
End Sub

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
      Select Case text06
         Case "", " ", "N":
         Case Else
            strMsg = "請輸入空白或N !"
            strTit = "檢核輸入"
            nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
            Cancel = True
            TextInverse text06
      End Select
   End If
End Sub

Public Function CheckDataValid() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   CheckDataValid = True

   If IsEmpty(text03_02) = True Then
      CheckDataValid = False
      strMsg = "公告日必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      GoTo EXITSUB
   ElseIf CheckIsTaiwanDate(text03_02, False) = False Then
      CheckDataValid = False
      strMsg = "請輸入正確的公告日"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      GoTo EXITSUB
   End If

   ' 檢查公告日起日是否小於止日
   If IsEmpty(text03_01) = False And IsEmpty(text03_02) = False Then
      If Val(text03_01) > Val(text03_02) Then
         CheckDataValid = False
         strMsg = "公告日起日必須小於止日"
         strTit = "檢核輸入"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         text03_01.SetFocus
         TextInverse text03_01
         GoTo EXITSUB
      End If
   End If
   
   If IsEmpty(text04) = True Then
      CheckDataValid = False
        'Modify By Cheng 2002/11/22
'      strMsg = "請輸入查詢方式D或R"
      strMsg = "請輸入查詢方式 1 或 2 "
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      GoTo EXITSUB
   End If

EXITSUB:
End Function

Public Sub PrintReport()
   Dim strSql As String
   Dim strSubSQL As String
   Dim rsData As New ADODB.Recordset
   Dim rsTmp As New ADODB.Recordset
   Dim strMsg, strTit As String
   Dim nResponse
   Dim strCurr1, strCurr2 As String
   Dim nPage As Integer
   Dim Fld1, Fld2, Fld3, Fld4, Fld5 As String
   Dim nCount As Integer
   Dim arrayAmount(3) As Integer
   Dim arrayTAmount(3) As Integer
   Dim nIndex As Integer
      
   nRow = 0
   nRowTot = 0 'Add By Sindy 2011/12/16
   strSubSQL = Empty
   strSql = "SELECT * FROM TPBulletin "
   
   If IsEmpty(text03_01) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & " AND "
      strSubSQL = strSubSQL & "TPB03 >= " & ChangeTStringToWString(text03_01) & " "
   End If
   If IsEmpty(text03_02) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & " AND "
      strSubSQL = strSubSQL & "TPB03 <= " & ChangeTStringToWString(text03_02) & " "
   End If
   If IsEmpty(text03_01) = False Or IsEmpty(text03_02) = False Then
      pub_QL05 = pub_QL05 & ";" & Label3 & text03_01 & "-" & text03_02 'Add By Sindy 2010/12/2
   End If
   If IsEmpty(text01_01) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & " AND "
      'Modify by Morgan 2006/9/19
      'strSubSQL = strSubSQL & "TPB07 >= '" & text01_01 & "' "
      strSubSQL = strSubSQL & "TPB07 = '" & text01_01 & "' "
      pub_QL05 = pub_QL05 & ";" & Label1 & text01_01 & text01_02 'Add By Sindy 2010/12/2
   End If
   If IsEmpty(text05) = False Then
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & " AND "
      strSubSQL = strSubSQL & "TPB02 >= '" & text05 & "' "
      pub_QL05 = pub_QL05 & ";" & Label5 & text05 'Add By Sindy 2010/12/2
   End If
   'Add By Sindy 2011/12/16
   If Me.Check1 = 1 Then
      pub_QL05 = pub_QL05 & ";只顯示本所案件"
   End If
   '2011/12/16 End
   If strSubSQL <> Empty Then
      strSql = strSql & "WHERE " & strSubSQL
   End If
   strSql = strSql & "ORDER BY TPB04, TPB05, TPB02 ASC"
   
   rsData.CursorLocation = adUseClient
   rsData.Open strSql, cnnConnection, adOpenDynamic
   
   If rsData.RecordCount <= 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
      strMsg = "無資料"
      strTit = "列印資料"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      rsData.Close
      GoTo EXITSUB
'   Else
'      InsertQueryLog (rsData.RecordCount) 'Add By Sindy 2010/12/2
   End If
   
PrintData:
   nPage = 1
   rsData.MoveFirst
   strCurr1 = rsData.Fields("TPB04")
   strCurr2 = rsData.Fields("TPB05")
   
   Printer.Orientation = vbPRORLandscape
   ' 印第一頁的表頭
   PrintPageHeader strCurr1, strCurr2, nPage
   
   ' 清除合計值
   For nIndex = 0 To 2
      arrayAmount(nIndex) = 0
   Next nIndex
   
   nRow = 1
   While rsData.EOF <> True
      Fld1 = rsData.Fields("TPB01")
      Fld2 = CStr(rsData.Fields("TPB02"))
      Fld3 = Empty
      Fld4 = Empty
      Fld5 = Empty
      
      ' 若卷號或期數不同則換頁
      If rsData.Fields("TPB04") <> strCurr1 Or rsData.Fields("TPB05") <> strCurr2 Then
         strCurr1 = rsData.Fields("TPB04")
         strCurr2 = rsData.Fields("TPB05")
         If bListDetail = True Then
            ' 印合計值
            PrintPageTailer arrayAmount(0), arrayAmount(1), arrayAmount(2)
            ' 換頁
            Printer.NewPage
            ' 頁碼更新
            nPage = nPage + 1
            ' 印表頭
            PrintPageHeader strCurr1, strCurr2, nPage
            nRow = 1
         Else
            PrintPageTailer arrayAmount(0), arrayAmount(1), arrayAmount(2), strCurr1 & "卷" & strCurr2 & "期"
         End If
         
         ' 清除合計值
         For nIndex = 0 To 2
            arrayAmount(nIndex) = 0
         Next nIndex
         
      End If
      
      ' 若列數超過 27 列則換頁
      If nRow > 27 Then
         Printer.NewPage
         nPage = nPage + 1
         PrintPageHeader strCurr1, strCurr2, nPage
         nRow = 1
      End If
      
      If bListDetail = True Then
      
         ' 申請人國籍
         If IsNull(rsData.Fields("TPB06")) = False Then
            Fld3 = GetNation(rsData.Fields("TPB06"))
         End If
         
         ' 代理人
         If IsNull(rsData.Fields("TPB07")) = False Then
            strSql = "SELECT * FROM TAGENT WHERE TA02 = '" & rsData.Fields("TPB07") & "'"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenDynamic
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               Fld4 = rsTmp.Fields("TA03")
            End If
            rsTmp.Close
         End If
         
         ' 本所案號
         strSql = "SELECT * FROM Patent WHERE PA11 = '" & rsData.Fields("TPB01") & "'"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenDynamic
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Fld5 = rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02") & "-" & rsTmp.Fields("PA03") & "-" & rsTmp.Fields("PA04")
         End If
         rsTmp.Close
         
         If Me.Check1 = 1 And Fld5 = "" Then GoTo ReadNext 'Add By Sindy 2011/12/16
         
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
      
         Printer.CurrentX = 102 * m_CharWidth
         Printer.CurrentY = (nRow + 13) * m_CharHeight
         Printer.Print Fld5
         
         nRow = nRow + 1
      End If
      
      ' 計算 發明, 新型, 設計的總數
      'Modify by Morgan 2010/12/28 申請案號改碼數
      'Select Case Mid(Fld1, 3, 1)
      Select Case Mid(Fld1, 4, 1)
         Case "1":
            arrayAmount(0) = arrayAmount(0) + 1
            arrayTAmount(0) = arrayTAmount(0) + 1
         Case "2":
            arrayAmount(1) = arrayAmount(1) + 1
            arrayTAmount(1) = arrayTAmount(1) + 1
         Case "3":
            arrayAmount(2) = arrayAmount(2) + 1
            arrayTAmount(2) = arrayTAmount(2) + 1
      End Select
      
ReadNext: 'Add By Sindy 2011/12/16
      rsData.MoveNext
   Wend
      
   If bListDetail = True Then
      ' 印合計值
      PrintPageTailer arrayAmount(0), arrayAmount(1), arrayAmount(2)
   Else
      PrintPageTailer arrayAmount(0), arrayAmount(1), arrayAmount(2), strCurr1 & "卷" & strCurr2 & "期"
   End If
   'Add by Morgan 2007/3/30 加印總計
   PrintPageTailer arrayTAmount(0), arrayTAmount(1), arrayTAmount(2), , True
   ' 結束列印
   Printer.EndDoc
   
   InsertQueryLog (nRowTot) 'Add By Sindy 2010/12/2
   
EXITSUB:
   Set rsTmp = Nothing
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
      strDate2 = Empty
   Else
      strDate2 = ChangeTStringToTDateString(strDate2)
   End If
   
   ' 表頭
   Printer.CurrentX = 58 * m_CharWidth
   Printer.CurrentY = 5 * m_CharHeight
   Printer.FontSize = 24
   Printer.Font.Underline = True
   Printer.Print "專利公報列印"
   
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
   
   'Modify by Morgan 2007/3/30 改判斷要印明細時才印卷期
   If bListDetail = True Then
      Printer.CurrentX = 12 * m_CharWidth
      Printer.CurrentY = 10 * m_CharHeight
      Printer.Print Str1 & " 卷 " & Str2 & " 期"
   End If
   'end 2007/3/30
   
   Printer.CurrentX = 101 * m_CharWidth
   Printer.CurrentY = 10 * m_CharHeight
   Printer.Print "頁"
   
   Printer.CurrentX = 107 * m_CharWidth
   Printer.CurrentY = 10 * m_CharHeight
   Printer.Print "次 : " & nPage
      
   For i = 0 To 107
      Printer.CurrentX = (i + 12) * m_CharWidth
      Printer.CurrentY = 11 * m_CharHeight
      Printer.Print "-"
   Next i

   'Modify by Morgan 2007/3/30 改判斷要印明細時才印欄位名稱
   If bListDetail = True Then
      
      Printer.CurrentX = 12 * m_CharWidth
      Printer.CurrentY = 12 * m_CharHeight
      Printer.Print "申請案號"
      
      Printer.CurrentX = 32 * m_CharWidth
      Printer.CurrentY = 12 * m_CharHeight
      Printer.Print "公告號"
      
      Printer.CurrentX = 52 * m_CharWidth
      Printer.CurrentY = 12 * m_CharHeight
      Printer.Print "申請人國籍"
      
      Printer.CurrentX = 82 * m_CharWidth
      Printer.CurrentY = 12 * m_CharHeight
      Printer.Print "代理人"
      
      Printer.CurrentX = 102 * m_CharWidth
      Printer.CurrentY = 12 * m_CharHeight
      Printer.Print "本所案號"
         
      For i = 0 To 107
         Printer.CurrentX = (i + 12) * m_CharWidth
         Printer.CurrentY = 13 * m_CharHeight
         Printer.Print "-"
      Next i
      
   End If
   'end 2007/3/30
   
End Sub

Public Sub PrintPageTailer(ByVal value1 As Integer, ByVal value2 As Integer, ByVal value3 As Integer, Optional ByVal stAdd As String, Optional ByVal bolRF As Boolean = False)
      
   If bListDetail = True And bolRF = False Then
      nRow = 30
   End If
   
   If stAdd = "" Then
      Printer.CurrentX = 12 * m_CharWidth
      Printer.CurrentY = (nRow + 11) * m_CharHeight
      Printer.Print String(150, "-") '108
      nRow = nRow + 1
   End If
   If bolRF = True Then
      Printer.CurrentX = 12 * m_CharWidth
      Printer.CurrentY = (nRow + 11) * m_CharHeight
      Printer.Print stAdd
      Printer.CurrentX = 32 * m_CharWidth
      Printer.CurrentY = (nRow + 11) * m_CharHeight
      Printer.Print "發明總計 : " & value1
      Printer.CurrentX = 52 * m_CharWidth
      Printer.CurrentY = (nRow + 11) * m_CharHeight
      Printer.Print "新型總計 : " & value2
      Printer.CurrentX = 72 * m_CharWidth
      Printer.CurrentY = (nRow + 11) * m_CharHeight
      Printer.Print "設計總計 : " & value3
      Printer.CurrentX = 92 * m_CharWidth
      Printer.CurrentY = (nRow + 11) * m_CharHeight
      Printer.Print "總計 : " & value1 + value2 + value3
      
      nRowTot = (value1 + value2 + value3) 'Add By Sindy 2011/12/16
   Else
      Printer.CurrentX = 12 * m_CharWidth
      Printer.CurrentY = (nRow + 11) * m_CharHeight
      Printer.Print stAdd
      Printer.CurrentX = 32 * m_CharWidth
      Printer.CurrentY = (nRow + 11) * m_CharHeight
      Printer.Print "發明合計 : " & value1
      Printer.CurrentX = 52 * m_CharWidth
      Printer.CurrentY = (nRow + 11) * m_CharHeight
      Printer.Print "新型合計 : " & value2
      Printer.CurrentX = 72 * m_CharWidth
      Printer.CurrentY = (nRow + 11) * m_CharHeight
      Printer.Print "設計合計 : " & value3
      Printer.CurrentX = 92 * m_CharWidth
      Printer.CurrentY = (nRow + 11) * m_CharHeight
      Printer.Print "合計 : " & value1 + value2 + value3
   End If
   nRow = nRow + 1
End Sub

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

Public Sub SetInputEntry()
   text01_01.SetFocus
End Sub

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As TextBox)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Private Sub text01_01_GotFocus()
   InverseAll text01_01
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


