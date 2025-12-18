VERSION 5.00
Begin VB.Form frm04060302 
   BorderStyle     =   1  '單線固定
   Caption         =   "公開市場佔有率統計表列印"
   ClientHeight    =   2610
   ClientLeft      =   420
   ClientTop       =   1905
   ClientWidth     =   5430
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5430
   Begin VB.CheckBox Check1 
      Caption         =   "亞洲包含大陸"
      Height          =   225
      Left            =   1560
      TabIndex        =   4
      Top             =   1815
      Width           =   3270
   End
   Begin VB.CommandButton buttonExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4524
      TabIndex        =   8
      Top             =   36
      Width           =   800
   End
   Begin VB.CommandButton buttonOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3696
      TabIndex        =   7
      Top             =   36
      Width           =   800
   End
   Begin VB.TextBox text04_02 
      Height          =   264
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox text04_01 
      Height          =   264
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2160
      Width           =   972
   End
   Begin VB.TextBox text03 
      Height          =   264
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1440
      Width           =   612
   End
   Begin VB.TextBox text02 
      Height          =   264
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1080
      Width           =   612
   End
   Begin VB.TextBox text01_02 
      Height          =   264
      Left            =   3720
      MaxLength       =   7
      TabIndex        =   1
      Top             =   720
      Width           =   1452
   End
   Begin VB.TextBox text01_01 
      Height          =   264
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   0
      Top             =   720
      Width           =   1452
   End
   Begin VB.Label Label12 
      Caption         =   "表3選擇："
      Height          =   255
      Left            =   255
      TabIndex        =   16
      Top             =   1785
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   3840
      X2              =   4080
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label6 
      Caption         =   "申請人國籍 :"
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "表3,4選擇："
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "(1:國內  2:國外  空白:全部)"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "表1選擇："
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   1092
   End
   Begin VB.Label Label7 
      Caption         =   "(1:表1  2:表2  3:表3  4:表4  空白:全部)"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000004&
      Caption         =   "報表選擇："
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "公開日："
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1092
   End
   Begin VB.Line Line1 
      X1              =   3240
      X2              =   3480
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frm04060302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Const m_CharWidth = 120
Const m_CharHeight = 240
'Const m_PaperSize = "REPORT"
Const m_PaperSize = "A4"

' 宣告報表表頭的欄位其資料型態
Private Type REPORTFIELD
   Name As String
   DataCode As String
   DataName As String
   Left As Long
   Width As Long
End Type
' 表頭欄位的內容
Dim m_Field(17) As REPORTFIELD
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

' 宣告代理人項目的資料型態
Private Type AGENTITEM
   AgentCode As String
   AgentName As String
   Count As Long
   Type1 As Long
   Type2 As Long
   Type3 As Long
End Type
' 宣告地區項目的資料型態
Private Type ZONEITEM
   ' 事務所代號
   ZoneCode As String
   ZoneName As String
   Count As Long
   TaieCount(3) As Long
   NoAgentItem As AGENTITEM
   AgentList() As AGENTITEM
   AgentCount As Long
End Type
' 定義地區陣列
Dim m_ZoneList() As ZONEITEM
' 定義代理人陣列
Dim m_AgentList() As AGENTITEM
' 地區串列中的資料筆數
Dim m_ZoneCount As Long
' 代理人串列的資料筆數
Dim m_AgentCount As Long
' 儲存原預設印表機的字串
Dim m_DefaultPrinter As String
'Add By Cheng 2003/05/19
Dim m_dblMaterialCnt As Double '實體審查件數
Dim m_dblTotCnt As Double '所有件數

' 使用者案下OK的按紐
Private Sub buttonOK_Click()
   Dim Prn As Printer
   
   If CheckDataValid = True Then
      '搜尋 Printer
      Screen.MousePointer = vbHourglass
        'Modify By Cheng 2002/11/26
        '使用預設印表機
'      For Each Prn In Printers
'         If Prn.DeviceName = cmbPrinter.Text Then
'            Set Printer = Prn
'            Exit For
'         End If
'      Next
      Select Case text02
         Case "1": Print_RP1
         Case "2": Print_RP2
         Case "3": Print_RP3
         Case "4": Print_RP4
         ' 報表選擇為空白時, 只列印表一, 表二及表三
         Case " ", "":
            Print_RP1
            Print_RP2
            Print_RP3
            'Print_RP4
      End Select
   
      ' 清除欄位
      ClearField
      ' 設定第一個欄位為輸入的Focus
      text01_01.SetFocus
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub ClearField()
   text01_01 = Empty
   text01_02 = Empty
   text02 = Empty
   text03 = Empty
   text04_01 = Empty
   text04_02 = Empty
End Sub

' 檢核輸入資料是否正確
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = True
   
   'Add By Sindy 2014/8/29
   strSql = "select * from tagent where ta01='P' and (ta04 is null or ta04='')"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      CheckDataValid = False
      strTit = "資料有誤"
      strMsg = "公報代理人檔，尚有事務所名稱為空白，怕影響統計結果" & vbCrLf & "請先補足資料再執行此作業！"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      GoTo EXITSUB
   End If
   '2014/8/29 END
   
   If IsEmpty(text01_02) = True Then
      CheckDataValid = False
      strTit = "資料輸入不正確"
      strMsg = "請輸入正確的公開日"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      GoTo EXITSUB
   ElseIf CheckIsTaiwanDate(text01_02, False) = False Then
      CheckDataValid = False
      strTit = "資料輸入不正確"
      strMsg = "請輸入正確的公開日"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      GoTo EXITSUB
   End If
   
   If IsEmpty(text01_01) = False Then
      If CheckIsTaiwanDate(text01_01) = False Then
         CheckDataValid = False
         strTit = "資料輸入不正確"
         strMsg = "請輸入正確的公開日"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         GoTo EXITSUB
      ElseIf Val(text01_01) > Val(text01_02) Then
         CheckDataValid = False
         strTit = "資料輸入不正確"
         strMsg = "請輸入正確的公開日範圍"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         text01_01.SetFocus
         TextInverse text01_01
         GoTo EXITSUB
      End If
   End If
   
   If IsEmpty(text04_01) = False And IsEmpty(text04_02) = False Then
      If Val(text04_01) > Val(text04_02) Then
         CheckDataValid = False
         strTit = "資料輸入不正確"
         strMsg = "請輸入正確的申請人國籍範圍"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         GoTo EXITSUB
      End If
   End If
   
   Select Case text02
      Case "1":
         Select Case text03
            Case " ", "":
            Case "1":
            Case "2":
            Case "3":
            Case Else
               CheckDataValid = False
               strTit = "資料輸入不正確"
               strMsg = "請輸入表1選擇正確的地區"
               nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         End Select
      Case "2":
      Case "3":
      Case "4":
      Case " ", "":
         Select Case text03
            Case "1":
            Case "2":
            Case " ", "":
            Case Else
               CheckDataValid = False
               strTit = "資料輸入不正確"
               strMsg = "請輸入表1選擇正確的地區"
               nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
               GoTo EXITSUB
         End Select
      Case Else
         CheckDataValid = False
         strTit = "資料輸入不正確"
         strMsg = "請輸入報表種類"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
   End Select
EXITSUB:
End Function

Private Sub Form_Load()
   Dim Prn As Printer
   Dim nIndex As Integer
   Dim nSel As Integer
      
   m_DefaultPrinter = Printer.DeviceName
   MoveFormToCenter Me
   
   nSel = 0
   nIndex = 0
   'For Each Prn In Printers
   '   cmbPrinter.AddItem Prn.DeviceName
   '   If Prn.DeviceName = strDeviceName Then
   '      nSel = nIndex
   '   End If
   '   nIndex = nIndex + 1
   'Next
   'cmbPrinter.ListIndex = nSel
    'Modify By Cheng 2002/11/26
    '使用預設
'   For Each Prn In Printers
'      If Prn.DeviceName <> m_DefaultPrinter Then
'         cmbPrinter.AddItem Prn.DeviceName
'      End If
'   Next
'   If cmbPrinter.ListCount > 0 Then: cmbPrinter.ListIndex = 0
End Sub

Private Sub buttonExit_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim Prn As Printer
    'Modify By Cheng 2002/11/26
    '使用預設印表機
'   '搜尋 Printer
'   For Each Prn In Printers
'      If Prn.DeviceName = m_DefaultPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
   'Add By Cheng 2002/07/18
   Set frm04060302 = Nothing
End Sub

Private Sub text01_01_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   If IsEmpty(text01_01) = False Then
      If CheckIsTaiwanDate(text01_01, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的公開日 !"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   Else
      Cancel = True
      strMsg = "公開日必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
   End If
   If Cancel Then TextInverse text01_01
End Sub

Private Sub text01_02_LostFocus()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   If IsEmpty(text01_02) = False Then
      If CheckIsTaiwanDate(text01_02, False) = False Then
         strMsg = "請輸入正確的公開日 !"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         text01_02.SetFocus
         TextInverse text01_02
      Else
         If Not ChkRange(text01_01, text01_02, "公開日") Then
         
         End If
      End If
   Else
      strMsg = "公開日必須輸入"
      strTit = "檢核輸入"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      text01_02.SetFocus
   End If
End Sub

Private Sub text02_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   Select Case text02.Text
      Case " ", "": Cancel = False
      Case "1": Cancel = False
      Case "2": Cancel = False
      Case "3": Cancel = False
      Case "4": Cancel = False
      Case " ": Cancel = False
      Case Else
         Cancel = True
         strMsg = "請輸入正確的選擇"
         strTit = "報表選擇"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         text02_GotFocus
   End Select
End Sub

Private Sub text03_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   
   If IsEmpty(text03) = False Then
      Select Case text03
         Case " ", "1", "2":
         Case Else
            Cancel = True
            strMsg = "請輸入正確的選擇"
            strTit = "表一選擇"
            nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
            text03_GotFocus
      End Select
   End If
End Sub

' 由國籍代碼取得國籍的名稱
Public Function GetNationName(ByVal strNation As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetNationName = Empty
   strSql = "SELECT * FROM NATION WHERE NA01 = '" & strNation & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("NA03")) = False Then
         GetNationName = rsTmp.Fields("NA03")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
' 由代理人代碼取得事務所的名稱
Public Function GetAgentCompany(ByVal strAgent As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetAgentCompany = Empty
   strSql = "SELECT * FROM TAGENT WHERE TA02 = '" & strAgent & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("TA04")) = False Then
         GetAgentCompany = rsTmp.Fields("TA04")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 列印報表表一
Private Sub Print_RP1()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
   If Len(text01_01) <> 0 Or Len(text01_02) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label3 & text01_01 & "-" & text01_02 'Add By Sindy 2010/12/2
   End If
   pub_QL05 = pub_QL05 & ";" & Label4 & "1:表1" 'Add By Sindy 2010/12/2
   If text03 = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label1 & "1:國內" 'Add By Sindy 2010/12/2
   ElseIf text03 = "2" Then
      pub_QL05 = pub_QL05 & ";" & Label1 & "2:國外" 'Add By Sindy 2010/12/2
   Else
      pub_QL05 = pub_QL05 & ";" & Label1 & "空白:全部" 'Add By Sindy 2010/12/2
   End If
   
   If GetDBData_RP(1) = True Then
      InsertQueryLog ("") 'Add By Sindy 2010/12/2
      BuildField_RP (1)
      Generate_RP1
      Clear
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
      strMsg = "無資料"
      strTit = "錯誤"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
   End If
End Sub

' 列印報表表二
Private Sub Print_RP2()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
   If Len(text01_01) <> 0 Or Len(text01_02) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label3 & text01_01 & "-" & text01_02 'Add By Sindy 2010/12/2
   End If
   pub_QL05 = pub_QL05 & ";" & Label4 & "2:表2" 'Add By Sindy 2010/12/2
   
   If GetDBData_RP(2) = True Then
      InsertQueryLog ("") 'Add By Sindy 2010/12/2
      BuildField_RP (2)
      Generate_RP2
      Clear
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
      strMsg = "無資料"
      strTit = "錯誤"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
   End If
End Sub

' 列印報表表三
Private Sub Print_RP3()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
   If Len(text01_01) <> 0 Or Len(text01_02) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label3 & text01_01 & "-" & text01_02 'Add By Sindy 2010/12/2
   End If
   pub_QL05 = pub_QL05 & ";" & Label4 & "3:表3" 'Add By Sindy 2010/12/2
   If Check1.Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label12 & Check1.Caption 'Add By Sindy 2010/12/2
   End If
   If Len(text04_01) <> 0 Or Len(text04_02) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label6 & text04_01 & "-" & text04_02 'Add By Sindy 2010/12/2
   End If
   
   If GetDBData_RP(3) = True Then
      InsertQueryLog ("") 'Add By Sindy 2010/12/2
      BuildField_RP (3)
      Generate_RP3
      Clear
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
      strMsg = "無資料"
      strTit = "錯誤"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
   End If
End Sub

' 列印報表表四
Private Sub Print_RP4()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
   If Len(text01_01) <> 0 Or Len(text01_02) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label3 & text01_01 & "-" & text01_02 'Add By Sindy 2010/12/2
   End If
   pub_QL05 = pub_QL05 & ";" & Label4 & "4:表4" 'Add By Sindy 2010/12/2
   If Len(text04_01) <> 0 Or Len(text04_02) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label6 & text04_01 & "-" & text04_02 'Add By Sindy 2010/12/2
   End If
   
   If GetDBData_RP(4) = True Then
      InsertQueryLog ("") 'Add By Sindy 2010/12/2
      BuildField_RP (4)
      Generate_RP4
      Clear
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/2
      strMsg = "無資料"
      strTit = "錯誤"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
   End If
End Sub

' 設定報表欄位的左方位置及其名稱
Public Sub BuildField_RP(ByVal nReport As Integer)
   Dim nIndex As Integer
   Dim nFieldWidth As Integer
   
   Select Case m_PaperSize
      Case "REPORT"
         m_LeftMargin = 1
         m_TopMargin = 5
         m_ReportWidth = 154
         m_ReportDataRows = 45
         nFieldWidth = 9
      Case Else
         m_LeftMargin = 10
         m_TopMargin = 5
         m_ReportWidth = 120
         m_ReportDataRows = 27
         nFieldWidth = 7
   End Select
   
   For nIndex = 0 To 16
      m_Field(nIndex).Width = nFieldWidth - 1
      m_Field(nIndex).Left = m_LeftMargin + (nIndex * nFieldWidth)
      Select Case nIndex
         Case 0:
            m_Field(nIndex).Name = "排名"
            Select Case nReport
               Case 1, 4: m_Field(nIndex).DataName = "事務所"
               Case 2: m_Field(nIndex).DataName = "地區"
            End Select
         Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14:
            m_Field(nIndex).Name = CStr(nIndex)
            Select Case nReport
               Case 1, 4:
                  If nIndex <= m_AgentCount Then
                     m_Field(nIndex).DataCode = m_AgentList(nIndex - 1).AgentCode
                     m_Field(nIndex).DataName = m_AgentList(nIndex - 1).AgentName
                  End If
               Case 2:
                  If nIndex <= m_ZoneCount Then
                     m_Field(nIndex).DataCode = m_ZoneList(nIndex - 1).ZoneCode
                     m_Field(nIndex).DataName = m_ZoneList(nIndex - 1).ZoneName
                  End If
            End Select
         Case 15:
            m_Field(nIndex).Name = "總計"
         Case 16:
            m_Field(nIndex).Name = "百分比"
      End Select
   Next nIndex
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

' 列印表一的表頭
Public Sub PrintPageHeader_RP(ByVal nReport As Integer, ByVal nPage As Integer, ByVal strZone As String)
   Dim nCount As Integer
   Dim strDate1 As String
   Dim strDate2 As String
   Dim nIndex As Integer
   Dim nRow As Integer
   Dim nX As Long
   Dim ny As Long
   Dim nCenter As Long
   Dim strTemp As String
      
   strDate1 = text01_01
   strDate2 = text01_02
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
   nRow = 0
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.FontSize = 24
   Printer.Font.Underline = True
   Select Case nReport
      Case 1:
         nX = m_LeftMargin + m_ReportWidth / 2 - 20
         Printer.CurrentX = nX * m_CharWidth
         Printer.Print "專利市場統計表(表一)"
      Case 2:
         nX = m_LeftMargin + m_ReportWidth / 2 - 28
         Printer.CurrentX = nX * m_CharWidth
         Printer.Print "台一各區市場佔有率排名(表二)"
      Case 3:
         nX = m_LeftMargin + m_ReportWidth / 2 - 24
         Printer.CurrentX = nX * m_CharWidth
         Printer.Print "各區市場佔有率排名(表三)"
      Case 4:
         nX = m_LeftMargin + m_ReportWidth / 2 - 20
         Printer.CurrentX = nX * m_CharWidth
         Printer.Print "代理人排名統計(表四)"
   End Select
   
   nRow = 3
   nX = m_LeftMargin + m_ReportWidth / 2 - 14
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.FontSize = 12
   Printer.Font.Underline = False
   Printer.Print "公開日 : " & strDate1 & " - " & strDate2
   
   nRow = nRow + 1
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "列印人 : " & strUserName
   
   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "製表日期 : " & Format(ChangeWStringToWDateString(GetTodayDate), "EE/MM/DD")
      
   ' 地
   nRow = nRow + 1
   If (nReport = 1 And IsEmpty(text03) = False) Or nReport = 4 Then
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "地"
      ' 區
      Printer.CurrentX = (m_LeftMargin + 4) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "區 : "
      ' 地區數值
      Printer.CurrentX = (m_LeftMargin + 10) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Select Case nReport
         Case 1, 2: Printer.Print strZone
         Case 4:
            If text04_01.Text = text04_02.Text Then
               Printer.Print text04_01
            Else
               Printer.Print text04_01 & "---" & text04_02
            End If
      End Select
   End If
   
   ' 頁
   nX = m_LeftMargin + m_ReportWidth - 20
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "頁"
   ' 次
   nX = m_LeftMargin + m_ReportWidth - 14
   Printer.CurrentX = nX * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "次 : " & nPage
      
   nRow = nRow + 1
   ' 列印分隔線
   PrintSplitLine nRow
   
   Select Case nReport
   Case 1, 4:
      nRow = nRow + 1
      For nIndex = 0 To 16
         nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
         strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
         Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print strTemp
      Next nIndex
      nRow = nRow + 1
      Printer.FontSize = 8
      For nIndex = 0 To 16
         nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
         strTemp = LeftStr(m_Field(nIndex).DataName, 12)
         Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print strTemp
      Next nIndex
      Printer.FontSize = 12
      
      nRow = nRow + 1
      For nX = 0 To 16
         For ny = m_Field(nX).Left To m_Field(nX).Left + (m_Field(nX).Width - 1)
            Printer.CurrentX = ny * m_CharWidth
            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
            Printer.Print "-"
         Next ny
      Next nX
      
      ' 列印分隔線
      'PrintSplitLine nRow
   Case 2:
      nRow = nRow + 1
      For nIndex = 0 To 15
         nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
         strTemp = LeftStr(m_Field(nIndex).Name, m_Field(nIndex).Width)
         Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print strTemp
      Next nIndex
      nRow = nRow + 1
      Printer.FontSize = 8
      For nIndex = 0 To 15
         nCenter = ((m_Field(nIndex).Left * m_CharWidth) + (m_Field(nIndex).Left + m_Field(nIndex).Width) * m_CharWidth) / 2
         strTemp = LeftStr(m_Field(nIndex).DataName, 12)
         Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print strTemp
      Next nIndex
      Printer.FontSize = 12
      
      nRow = nRow + 1
      For nX = 0 To 15
         For ny = m_Field(nX).Left To m_Field(nX).Left + m_Field(nX).Width - 1
            Printer.CurrentX = ny * m_CharWidth
            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
            Printer.Print "-"
         Next ny
      Next nX
      ' 列印分隔線
      'PrintSplitLine nRow
   End Select
   
   m_HeaderHeight = nRow
End Sub
' 列印表一的內容
Public Sub Generate_RP1()
   Dim strZone As String
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nAgentCount As Integer
   Dim bFindAgent As Boolean
   Dim nType As Integer
   Dim nAmount As Double
   Dim nTaieAmount As Long
   Dim nNoAgentAmount As Long
   Dim nTotalAmount As Long
   Dim nZoneCount As Long
   Dim nCount As Long
   Dim fValue As Single
   Dim nFinalAmount As Long
   Dim nNoAgent(3) As Long
   Dim nX As Integer
   Dim nRight As Long
   Dim strTemp As String
    Dim nTmpRow As Integer
   
   ' 紙張大小
   Select Case m_PaperSize
      Case "A4":
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
      Case "REPORT":
         Printer.PaperSize = vbPRPSFanfoldUS
      Case Else
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
   End Select
   ' 紙張方向
   
   Select Case text03
      Case "1":
         strZone = "國內"
      Case "2":
         strZone = "國外"
      Case " ":
         strZone = "全部"
   End Select
   
   ' 印第一頁的表頭
   nPage = 1
   PrintPageHeader_RP 1, nPage, strZone
   
   ' 依序列印表一的第一部份
   nRow = 1
   For nZoneCount = 0 To m_ZoneCount - 1
      ' 超過最多筆數為
        'Modify By Cheng 2003/05/19
'      If nRow > (m_ReportDataRows - 6) Then
      If nRow > (m_ReportDataRows - 4) Then
         ' 換頁
         Printer.NewPage
         nRow = 1
         nPage = nPage + 1
         ' 印頁首
         PrintPageHeader_RP 1, nPage, strZone
      End If
      
      ' 依序產生 發明, 新型, 設計, 小計的資料
      For nType = 1 To 4
            'Modify By Cheng 2003/05/19
            '公開只有發明資料
            If nType = 4 Then
                ' 第一個欄位的內容
                Select Case nType
                   Case 1: fld(0) = "發明"
                   Case 2: fld(0) = "新型"
                   Case 3: fld(0) = "設計"
                   Case 4: fld(0) = m_ZoneList(nZoneCount).ZoneName
                End Select
                ' 清除欄位的內容
                For nAgentCount = 1 To 16
                   fld(nAgentCount) = Empty
                Next nAgentCount
                ' 依序計算出各欄位的內容
                For nAgentCount = 0 To Min(13, m_AgentCount - 1)
                   bFindAgent = False
                   nAmount = 0
                   Select Case nType
                      Case 1:
                         'nAmount = GetZoneAgentAmount(m_ZoneList(nZoneCount), m_AgentList(nAgentCount).AgentCode, 1, bFindAgent)
                         nAmount = GetZoneAgentAmount(m_ZoneList(nZoneCount), m_AgentList(nAgentCount).AgentName, 1, bFindAgent)
                      Case 2:
                         'nAmount = GetZoneAgentAmount(m_ZoneList(nZoneCount), m_AgentList(nAgentCount).AgentCode, 2, bFindAgent)
                         nAmount = GetZoneAgentAmount(m_ZoneList(nZoneCount), m_AgentList(nAgentCount).AgentName, 2, bFindAgent)
                      Case 3:
                         'nAmount = GetZoneAgentAmount(m_ZoneList(nZoneCount), m_AgentList(nAgentCount).AgentCode, 3, bFindAgent)
                         nAmount = GetZoneAgentAmount(m_ZoneList(nZoneCount), m_AgentList(nAgentCount).AgentName, 3, bFindAgent)
                      Case 4:
                         'nAmount = GetZoneAgentAmount(m_ZoneList(nZoneCount), m_AgentList(nAgentCount).AgentCode, 0, bFindAgent)
                         nAmount = GetZoneAgentAmount(m_ZoneList(nZoneCount), m_AgentList(nAgentCount).AgentName, 0, bFindAgent)
                   End Select
                         
                   ' 將資料放入欄位中
                   fld(nAgentCount + 1) = CStr(nAmount)
                Next nAgentCount
                
                ' 台一的件數
                nTaieAmount = 0
                Select Case nType
                   Case 1: nTaieAmount = GetZoneTaieAmount(m_ZoneList(nZoneCount), 1)
                   Case 2: nTaieAmount = GetZoneTaieAmount(m_ZoneList(nZoneCount), 2)
                   Case 3: nTaieAmount = GetZoneTaieAmount(m_ZoneList(nZoneCount), 3)
                   Case 4: nTaieAmount = GetZoneTaieAmount(m_ZoneList(nZoneCount), 0)
                End Select
                
                ' 總件數
                Select Case nType
                   Case 1: nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount), 1, 0)
                   Case 2: nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount), 2, 0)
                   Case 3: nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount), 3, 0)
                   Case 4: nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount), 0, 0)
                End Select
                
                ' 總計欄位
                fld(15) = CStr(nTotalAmount)
                ' 百分比欄位
                If nTotalAmount > 0 Then
                   fValue = (nTaieAmount * 100) / nTotalAmount
                   fld(16) = Format(fValue, "##0.00")
                Else
                   fld(16) = Format(0, "##0.00")
                End If
                ' 將資料列印到印表機
                For nAgentCount = 0 To 16
                   Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
                   If nAgentCount > 0 Then
                      nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
                      strTemp = LeftStr(fld(nAgentCount), m_Field(nAgentCount).Width)
                      '911031 nick 將資料往後移
                      'Printer.CurrentX = nRight - Printer.TextWidth(strTemp)
                      Printer.CurrentX = nRight - Printer.TextWidth(strTemp) + 200
                   End If
                   Select Case nType
                      Case 1, 2, 3: Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow + (nType - 1)) * m_CharHeight
'                      Case 4: Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow + nType) * m_CharHeight
                      Case 4
                        Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nRow + nType - 4) * m_CharHeight
                   End Select
                   Printer.Print fld(nAgentCount)
                Next nAgentCount
                ' 列印分隔線
'                If nType = 3 Or nType = 4 Then
                If nType = 4 Then
'                   Dim nTmpRow As Integer
                   Select Case nType
                      Case 3: nTmpRow = m_HeaderHeight + nRow + nType
'                      Case 4: nTmpRow = m_HeaderHeight + nRow + nType + 1
                      Case 4: nTmpRow = m_HeaderHeight + nRow + nType - 3
                   End Select
                   ' 列印分隔線
                   PrintSplitLine nTmpRow
                End If
            End If
      Next nType
        'Modify By Cheng 2003/05/19
'      nRow = nRow + 6
      nRow = nRow + 6 - 4
   Next nZoneCount
   
   ' 列印表一的第二部份
   Printer.NewPage
   nPage = nPage + 1
   ' 印頁首
   PrintPageHeader_RP 1, nPage, strZone
   
   nNoAgentAmount = GetAllZoneAmount(0, 2)
   nFinalAmount = GetAllZoneAmount(0, 0)
   
   For nType = 1 To 6
        'Modify By Cheng 2003/05/19
        '公只有發明資料
        If nType <> 1 And nType <> 2 And nType <> 3 Then
            ' 第一個欄位的內容
            Select Case nType
               Case 1: fld(0) = "發明"
               Case 2: fld(0) = "新型"
               Case 3: fld(0) = "設計"
               Case 4: fld(0) = "總計"
               Case 5: fld(0) = "佔代理%"
               Case 6: fld(0) = "佔專利%"
            End Select
            ' 清除欄位的內容
            For nAgentCount = 1 To 16
               fld(nAgentCount) = Empty
            Next nAgentCount
            
            ' 台一的件數
            nTaieAmount = 0
            Select Case nType
               Case 1: nTaieAmount = GetAllZoneTaieAmount(1)
               Case 2: nTaieAmount = GetAllZoneTaieAmount(2)
               Case 3: nTaieAmount = GetAllZoneTaieAmount(3)
               Case 4: nTaieAmount = GetAllZoneTaieAmount(0)
            End Select
            
            ' 取得總件數
            nTotalAmount = 0
            Select Case nType
               Case 1: nTotalAmount = GetAllZoneAmount(1, 0)
               Case 2: nTotalAmount = GetAllZoneAmount(2, 0)
               Case 3: nTotalAmount = GetAllZoneAmount(3, 0)
               Case 4: nTotalAmount = GetAllZoneAmount(0, 0)
            End Select
            
            ' 欄位15, 16內容
            Select Case nType
               Case 1, 2, 3, 4:
                  fld(15) = CStr(nTotalAmount)
                  If nTotalAmount > 0 Then
                     fValue = (nTaieAmount * 100) / nTotalAmount
                     fld(16) = Format(fValue, "##0.00")
                  Else
                     fld(16) = Format(0, "##0.00")
                  End If
            End Select
            
            ' 列印欄位的內容
            For nAgentCount = 0 To Min(13, m_AgentCount - 1)
               nAmount = 0
               ' 設定欄位內的值
               Select Case nType
                  ' 發明
                  Case 1:
                     'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 1)
                     nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 1)
                     fld(nAgentCount + 1) = CStr(nAmount)
                  ' 新型
                  Case 2:
                     'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 2)
                     nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 2)
                     fld(nAgentCount + 1) = CStr(nAmount)
                  ' 設計
                  Case 3:
                     'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 3)
                     nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 3)
                     fld(nAgentCount + 1) = CStr(nAmount)
                  ' 總計
                  Case 4:
                     'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 0)
                     nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 0)
                     fld(nAgentCount + 1) = CStr(nAmount)
                  '佔代理
                  Case 5:
                     'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 0)
                     nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 0)
                     fValue = (nAmount * 100) / nNoAgentAmount
                     fld(nAgentCount + 1) = Format(fValue, "##0.00")
                  '佔專利
                  Case 6:
                     'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 0)
                     nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 0)
                     fValue = (nAmount * 100) / nFinalAmount
                     fld(nAgentCount + 1) = Format(fValue, "##0.00")
               End Select
            Next nAgentCount
            
            ' 將資料列印到印表機
            For nAgentCount = 0 To 16
               Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
               If nAgentCount > 0 Then
                  nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
                  strTemp = LeftStr(fld(nAgentCount), m_Field(nAgentCount).Width)
                  '911031 nick 將資料往後移
                  'Printer.CurrentX = nRight - Printer.TextWidth(strTemp)
                  Printer.CurrentX = nRight - Printer.TextWidth(strTemp) + 200
               End If
               Select Case nType
                  Case 1, 2, 3: Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nType) * m_CharHeight
                  Case 4: Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 5 - 4) * m_CharHeight
                  Case 5: Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 7 - 4) * m_CharHeight
                  Case 6: Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 9 - 4) * m_CharHeight
               End Select
               Printer.Print fld(nAgentCount)
            Next nAgentCount
            ' 列印分隔線
            If nType >= 3 Then
               Dim ny As Integer
               Select Case nType
                  Case 3: ny = (m_HeaderHeight + 4)
                  Case 4: ny = (m_HeaderHeight + 6)
                  Case 5: ny = (m_HeaderHeight + 8)
                  Case 6: ny = (m_HeaderHeight + 10)
               End Select
               ' 列印分隔線
               PrintSplitLine ny - 4
            End If
        End If
   Next nType
   
   nNoAgent(0) = GetAllZoneAmount(1, 1)
   nNoAgent(1) = GetAllZoneAmount(2, 1)
   nNoAgent(2) = GetAllZoneAmount(3, 1)
   
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 11 - 4) * m_CharHeight
    'Modify By Cheng 2003/05/19
'   Printer.Print "無代理人申請專利   " & _
'                 "發明 : " & nNoAgent(0) & "          " & _
'                 "新型 : " & nNoAgent(1) & "          " & _
'                 "設計 : " & nNoAgent(2)
   Printer.Print "無代理人申請專利" & _
                 " : " & nNoAgent(0)
   
   nTotalAmount = 0
   For nAgentCount = 0 To m_AgentCount - 1
      nTotalAmount = nTotalAmount + m_AgentList(nAgentCount).Type1 + m_AgentList(nAgentCount).Type2 + m_AgentList(nAgentCount).Type3
   Next nAgentCount
   fValue = (nNoAgent(0) + nNoAgent(1) + nNoAgent(2)) * 100 / nFinalAmount
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 12 - 4) * m_CharHeight
   Printer.Print "佔專利市場% : " & Format(fValue, "##0.00") & " %"
   'Add By Cheng 2003/05/19
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 13 - 4) * m_CharHeight
   Printer.Print "有實審案件 : " & m_dblMaterialCnt & " 件"
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 14 - 4) * m_CharHeight
   Printer.Print "有實審案件佔專利市場% : " & Format((m_dblMaterialCnt / m_dblTotCnt) * 100, "##0.00") & " %"
       
   Printer.EndDoc
   
End Sub
' 列印表二的內容
Public Sub Generate_RP2()
   Dim strZone
   Dim fld(17) As String
   Dim nCount As Integer
   Dim nType As Integer
   Dim nAmount As Double
   Dim nRow As Integer
   Dim nPage As Integer
   Dim nZoneCount As Long
   Dim nAgentCount As Long
   Dim nTaieAmount As Long
   Dim nTotalAmount As Long
   Dim fValue As Long
   Dim nRight As Long
   Dim strTemp As String
      
   ' 紙張大小
   Select Case m_PaperSize
      Case "A4":
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
      Case "REPORT":
         Printer.PaperSize = vbPRPSFanfoldUS
      Case Else
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
   End Select
   strZone = "國內"
   
   ' 印第一頁的表頭
   nPage = 1
   PrintPageHeader_RP 2, nPage, strZone
   For nType = 1 To 3
      ' 清除欄位內容
      For nCount = 0 To 16
         fld(nCount) = Empty
      Next nCount
      ' 設定第一個欄位的內容
      Select Case nType
         Case 1:
            fld(0) = "台一合計"
            ' 依地區計算該地區台一的總數
            For nZoneCount = 0 To Min(13, m_ZoneCount - 1)
               nTaieAmount = GetZoneTaieAmount(m_ZoneList(nZoneCount), 0)
               fld(nZoneCount + 1) = CStr(nTaieAmount)
            Next nZoneCount
            nAmount = GetAllZoneTaieAmount(0)
            fld(15) = CStr(nAmount)
         Case 2: fld(0) = "地區合計"
            ' 依地區計算該地區的總數
            For nZoneCount = 0 To Min(13, m_ZoneCount - 1)
               nAmount = GetZoneAmount(m_ZoneList(nZoneCount), 0, 0)
               fld(nZoneCount + 1) = CStr(nAmount)
            Next nZoneCount
            nAmount = GetAllZoneAmount(0, 0)
            fld(15) = CStr(nAmount)
         Case 3: fld(0) = "佔地區%"
            For nZoneCount = 0 To Min(13, m_ZoneCount - 1)
               nTaieAmount = GetZoneTaieAmount(m_ZoneList(nZoneCount), 0)
               nAmount = GetZoneAmount(m_ZoneList(nZoneCount), 0, 0)
               If nAmount > 0 Then
                  fValue = (nTaieAmount * 100) / nAmount
                  fld(nZoneCount + 1) = Format(fValue, "##0.00")
               Else
                  fld(nZoneCount + 1) = Format(0, "##0.00")
               End If
            Next nZoneCount
            
            nTaieAmount = GetAllZoneTaieAmount(0)
            nAmount = GetAllZoneAmount(0, 0)
            If nAmount > 0 Then
               fValue = (nTaieAmount * 100) / nAmount
               fld(15) = Format(fValue, "##0.00")
            Else
               fld(15) = Format(0, "##0.00")
            End If
      End Select
      
      ' 列印資料
      For nZoneCount = 0 To 15
         Printer.CurrentX = m_Field(nZoneCount).Left * m_CharWidth
         If nZoneCount > 0 Then
            nRight = (m_Field(nZoneCount).Left + m_Field(nZoneCount).Width - 2) * m_CharWidth
            strTemp = LeftStr(fld(nZoneCount), m_Field(nZoneCount).Width)
            '911031 nick 將資料往後移
            'Printer.CurrentX = nRight - Printer.TextWidth(strTemp)
            Printer.CurrentX = nRight - Printer.TextWidth(strTemp) + 200
         End If
         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + (nType * 2) - 1) * m_CharHeight
         Printer.Print fld(nZoneCount)
      Next nZoneCount
            
      ' 列印分隔線
      nRow = m_HeaderHeight + nType * 2
      PrintSplitLine nRow
   Next nType
   
   Printer.EndDoc
   
End Sub

' 列印表二的內容
Public Sub Generate_RP3()
   Dim strZone
   Dim fldTemp As String
   Dim nCount As Integer
   Dim nType As Integer
   Dim nAmount As Double
   Dim nRow As Integer
   Dim nPage As Integer
   Dim nZoneCount As Long
   Dim nAgentCount As Long
   Dim nTaieAmount As Long
   Dim nTotalAmount As Long
   Dim fValue As Long
   Dim nRight As Long
   Dim nCenter As Long
   Dim strTemp As String
   Dim nX As Integer
   Dim ny As Integer
   
   ' 紙張大小
   Select Case m_PaperSize
      Case "A4":
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
      Case "REPORT":
         Printer.PaperSize = vbPRPSFanfoldUS
      Case Else
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
   End Select
   strZone = "國內"
   
   nPage = 1
   PrintPageHeader_RP 3, nPage, strZone
   
   nRow = 1
   For nZoneCount = 0 To m_ZoneCount - 1
      If nRow > m_ReportDataRows - 5 Then
         Printer.NewPage
         nPage = nPage + 1
         PrintPageHeader_RP 3, nPage, strZone
         nRow = 1
      End If
      
      ' 計算總件數
      nTotalAmount = GetZoneAmount(m_ZoneList(nZoneCount), 0, 0)
      
      ' 列印地區名稱及代碼
      nRow = nRow + 1
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
      'edit by nickc 2006/04/27 加入洲
      'Printer.Print "<" & m_ZoneList(nZoneCount).ZoneCode & ">" & m_ZoneList(nZoneCount).ZoneName
      If Mid(m_ZoneList(nZoneCount).ZoneCode, 3) = "ZZZ" Then
         Printer.Print m_ZoneList(nZoneCount).ZoneName
      Else
         Printer.Print "<" & Mid(m_ZoneList(nZoneCount).ZoneCode, 3) & ">" & m_ZoneList(nZoneCount).ZoneName
      End If
      ' 列印分隔線
      nRow = nRow + 1
      PrintSplitLine m_HeaderHeight + nRow
      ' 列印排名
      nRow = nRow + 1
      For nCount = 0 To 15
         nCenter = ((m_Field(nCount).Left * m_CharWidth) + (m_Field(nCount).Left + m_Field(nCount).Width) * m_CharWidth) / 2
         strTemp = m_Field(nCount).Name
         Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
         Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
         Printer.Print strTemp
      Next nCount
      ' 列印事務所名稱
      nRow = nRow + 1
      Printer.FontSize = 8
      strTemp = "事務所"
      nCenter = ((m_Field(0).Left * m_CharWidth) + (m_Field(0).Left + m_Field(0).Width) * m_CharWidth) / 2
      Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
      Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
      Printer.Print strTemp
      ' 依序列印各事務所的名稱
      If m_ZoneList(nZoneCount).Count - m_ZoneList(nZoneCount).NoAgentItem.Count > 0 Then
         For nAgentCount = 0 To Min(13, m_ZoneList(nZoneCount).AgentCount - 1)
            nCenter = ((m_Field(nAgentCount + 1).Left * m_CharWidth) + (m_Field(nAgentCount + 1).Left + m_Field(nAgentCount + 1).Width) * m_CharWidth) / 2
            strTemp = LeftStr(m_ZoneList(nZoneCount).AgentList(nAgentCount).AgentName, 12)
            Printer.CurrentX = nCenter - Printer.TextWidth(strTemp) / 2
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print strTemp
         Next nAgentCount
      End If
      Printer.FontSize = 12
      ' 列印分隔線
      nRow = nRow + 1
      For nX = 0 To 15
         For ny = m_Field(nX).Left To m_Field(nX).Left + m_Field(nX).Width - 1
            Printer.CurrentX = ny * m_CharWidth
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print "-"
         Next ny
      Next nX
      
      ' 列印件數
      nRow = nRow + 1
      Printer.CurrentX = m_Field(0).Left * m_CharWidth
      Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
      Printer.Print "件數"
      ' 依序列印各事務所的件數
      If m_ZoneList(nZoneCount).Count - m_ZoneList(nZoneCount).NoAgentItem.Count > 0 Then
         For nAgentCount = 0 To Min(13, m_ZoneList(nZoneCount).AgentCount - 1)
            strTemp = m_ZoneList(nZoneCount).AgentList(nAgentCount).Count
            nRight = (m_Field(nAgentCount + 1).Left + m_Field(nAgentCount + 1).Width - 2) * m_CharWidth
            Printer.CurrentX = nRight - Printer.TextWidth(strTemp)
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print strTemp
         Next nAgentCount
      End If
      ' 列印總件數
      nRight = (m_Field(15).Left + m_Field(15).Width - 2) * m_CharWidth
      strTemp = nTotalAmount
      Printer.CurrentX = nRight - Printer.TextWidth(strTemp)
      Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
      Printer.Print strTemp
      ' 列印百分比
      nRow = nRow + 1
      Printer.CurrentX = m_Field(0).Left * m_CharWidth
      Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
      Printer.Print "百分比"
      ' 依序列印各事務所的百分比
      If m_ZoneList(nZoneCount).Count - m_ZoneList(nZoneCount).NoAgentItem.Count > 0 And nTotalAmount > 0 Then
         For nAgentCount = 0 To Min(13, m_ZoneList(nZoneCount).AgentCount - 1)
            nRight = (m_Field(nAgentCount + 1).Left + m_Field(nAgentCount + 1).Width - 2) * m_CharWidth
            fValue = (m_ZoneList(nZoneCount).AgentList(nAgentCount).Count * 100) / nTotalAmount
            strTemp = Format(fValue, "##0.00")
            Printer.CurrentX = nRight - Printer.TextWidth(strTemp)
            Printer.CurrentY = (m_TopMargin + m_HeaderHeight + nRow) * m_CharHeight
            Printer.Print strTemp
         Next nAgentCount
      End If
      ' 列印分隔線
      nRow = nRow + 1
      PrintSplitLine m_HeaderHeight + nRow
      ' 空白
      nRow = nRow + 1
   Next nZoneCount
   
   Printer.EndDoc
   
End Sub

' 列印表四的內容
Public Sub Generate_RP4()
   Dim strZone As String
   Dim nRow As Integer
   Dim nPage As Integer
   Dim fld(17) As String
   Dim nAgentCount As Integer
   Dim bFindAgent As Boolean
   Dim nType As Integer
   Dim nAmount As Double
   Dim nTaieAmount As Long
   Dim nNoAgentAmount As Long
   Dim nTotalAmount As Long
   Dim nZoneCount As Long
   Dim nCount As Integer
   Dim fValue As Single
   Dim nFinalAmount As Long
   Dim nNoAgent(3) As Long
   Dim nX As Integer
   Dim nRight As Long
   Dim strTemp As String
   
   ' 紙張大小
   Select Case m_PaperSize
      Case "A4":
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
      Case "REPORT":
         Printer.PaperSize = vbPRPSFanfoldUS
      Case Else
         Printer.PaperSize = vbPRPSA4
         Printer.Orientation = vbPRORLandscape
   End Select
   
   Select Case text03
      Case "1":
         strZone = "國內"
      Case "2":
         strZone = "國外"
      Case " ":
         strZone = "全部"
   End Select
      
   ' 印第一頁的表頭
   nPage = 1
   PrintPageHeader_RP 4, nPage, strZone
   nRow = 1
   
   nNoAgentAmount = GetAllZoneAmount(0, 2)
   nFinalAmount = GetAllZoneAmount(0, 0)
   
   For nType = 1 To 6
      ' 第一個欄位的內容
      Select Case nType
         Case 1: fld(0) = "發明"
         Case 2: fld(0) = "新型"
         Case 3: fld(0) = "設計"
         Case 4: fld(0) = "總計"
         Case 5: fld(0) = "佔代理%"
         Case 6: fld(0) = "佔專利%"
      End Select
      ' 清除欄位的內容
      For nAgentCount = 1 To 16
         fld(nAgentCount) = Empty
      Next nAgentCount
      
      ' 台一的件數
      nTaieAmount = 0
      Select Case nType
         Case 1: nTaieAmount = GetAllZoneTaieAmount(1)
         Case 2: nTaieAmount = GetAllZoneTaieAmount(2)
         Case 3: nTaieAmount = GetAllZoneTaieAmount(3)
         Case 4: nTaieAmount = GetAllZoneTaieAmount(0)
      End Select
      
      ' 取得總件數
      nTotalAmount = 0
      Select Case nType
         Case 1: nTotalAmount = GetAllZoneAmount(1, 0)
         Case 2: nTotalAmount = GetAllZoneAmount(2, 0)
         Case 3: nTotalAmount = GetAllZoneAmount(3, 0)
         Case 4: nTotalAmount = GetAllZoneAmount(0, 0)
      End Select
      
      ' 欄位15, 16內容
      Select Case nType
         Case 1, 2, 3, 4:
            fld(15) = CStr(nTotalAmount)
            If nTotalAmount > 0 Then
               fValue = (nTaieAmount * 100) / nTotalAmount
               fld(16) = Format(fValue, "##0.00")
            Else
               fld(16) = Format(0, "##0.00")
            End If
      End Select
      
      ' 列印欄位的內容
      For nAgentCount = 0 To Min(13, m_AgentCount - 1)
         nAmount = 0
         ' 設定欄位內的值
         Select Case nType
            ' 發明
            Case 1:
               'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 1)
               nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 1)
               fld(nAgentCount + 1) = CStr(nAmount)
            ' 新型
            Case 2:
               'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 2)
               nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 2)
               fld(nAgentCount + 1) = CStr(nAmount)
            ' 設計
            Case 3:
               'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 3)
               nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 3)
               fld(nAgentCount + 1) = CStr(nAmount)
            ' 總計
            Case 4:
               'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 0)
               nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 0)
               fld(nAgentCount + 1) = CStr(nAmount)
            '佔代理
            Case 5:
               'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 0)
               nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 0)
               fValue = (nAmount * 100) / nNoAgentAmount
               fld(nAgentCount + 1) = Format(fValue, "##0.00")
            '佔專利
            Case 6:
               'nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentCode, 0)
               nAmount = GetAllZoneAgentAmount(m_AgentList(nAgentCount).AgentName, 0)
               fValue = (nAmount * 100) / nFinalAmount
               fld(nAgentCount + 1) = Format(fValue, "##0.00")
         End Select
      Next nAgentCount
      
      ' 將資料列印到印表機
      For nAgentCount = 0 To 16
         Printer.CurrentX = m_Field(nAgentCount).Left * m_CharWidth
         If nAgentCount > 0 Then
            nRight = (m_Field(nAgentCount).Left + m_Field(nAgentCount).Width - 2) * m_CharWidth
            strTemp = fld(nAgentCount)
            Printer.CurrentX = nRight - Printer.TextWidth(strTemp)
         End If
         Select Case nType
            Case 1, 2, 3: Printer.CurrentY = (m_HeaderHeight + m_TopMargin + nType) * m_CharHeight
            Case 4: Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 5) * m_CharHeight
            Case 5: Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 7) * m_CharHeight
            Case 6: Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 9) * m_CharHeight
         End Select
         Printer.Print fld(nAgentCount)
      Next nAgentCount
      ' 列印分隔線
      If nType >= 3 Then
         Dim ny As Integer
         Select Case nType
            Case 3: ny = (m_HeaderHeight + 4)
            Case 4: ny = (m_HeaderHeight + 6)
            Case 5: ny = (m_HeaderHeight + 8)
            Case 6: ny = (m_HeaderHeight + 10)
         End Select
         ' 列印分隔線
         PrintSplitLine ny
      End If
   Next nType
   
   nNoAgent(0) = GetAllZoneAmount(1, 1)
   nNoAgent(1) = GetAllZoneAmount(2, 1)
   nNoAgent(2) = GetAllZoneAmount(3, 1)
   
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 11) * m_CharHeight
   Printer.Print "無代理人申請專利   " & _
                 "發明 : " & nNoAgent(0) & "          " & _
                 "新型 : " & nNoAgent(1) & "          " & _
                 "設計 : " & nNoAgent(2)
   
   nTotalAmount = 0
   For nAgentCount = 0 To m_AgentCount - 1
      nTotalAmount = nTotalAmount + m_AgentList(nAgentCount).Type1 + m_AgentList(nAgentCount).Type2 + m_AgentList(nAgentCount).Type3
   Next nAgentCount
   fValue = (nNoAgent(0) + nNoAgent(1) + nNoAgent(2)) * 100 / nFinalAmount
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_HeaderHeight + m_TopMargin + 12) * m_CharHeight
   Printer.Print "佔專利市場% : " & Format(fValue, "##0.00") & " % "
   
   Printer.EndDoc
   
End Sub

' 清除資料
Public Sub Clear()
   Dim nZoneCount As Integer
   Dim nAgentCount As Integer
   For nZoneCount = 0 To m_ZoneCount - 1
      If m_ZoneList(nZoneCount).Count - m_ZoneList(nZoneCount).NoAgentItem.Count > 0 Then
         Erase m_ZoneList(nZoneCount).AgentList
      End If
   Next nZoneCount
   
   If m_ZoneCount > 0 Then
      Erase m_ZoneList
   End If
   m_ZoneCount = 0
   
   For nAgentCount = 0 To 16
      m_Field(nAgentCount).Name = Empty
      m_Field(nAgentCount).DataCode = Empty
      m_Field(nAgentCount).DataName = Empty
      m_Field(nAgentCount).Left = 0
      m_Field(nAgentCount).Width = 0
   Next nAgentCount
   
   If m_AgentCount > 0 Then
      Erase m_AgentList
   End If
   m_AgentCount = 0
   
End Sub

'取得所有地區該事務所的資料
' Input : strAgent ==> 代理人的代碼
'         nType ==> 取得資訊的種類
'            0 : 總數量
'            1 : 表取得所有事務所的發明合計數量
'            2 : 表取得所有事務所的新型合計數量
'            3 : 表取得所有事務所的設計合計數量
'edit by nickc 2007/01/25 解決溢位
'Public Function GetAllZoneAgentAmount(ByRef strAgent As String, ByVal nType As Integer) As Integer
Public Function GetAllZoneAgentAmount(ByRef strAgent As String, ByVal nType As Integer) As Double
   Dim nAmount As Double
   Dim nAgentCount As Integer
   Dim nZoneCount As Integer
   Dim bFind As Boolean
   
   nAmount = 0
   For nZoneCount = 0 To m_ZoneCount - 1
      If m_ZoneList(nZoneCount).Count - m_ZoneList(nZoneCount).NoAgentItem.Count > 0 Then
         For nAgentCount = 0 To m_ZoneList(nZoneCount).AgentCount - 1
            'If m_ZoneList(nZoneCount).AgentList(nAgentCount).AgentCode = strAgent Then
            If m_ZoneList(nZoneCount).AgentList(nAgentCount).AgentName = strAgent Then
               Select Case nType
                  Case 0: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Count
                  Case 1: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type1
                  Case 2: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type2
                  Case 3: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type3
               End Select
            End If
         Next nAgentCount
      End If
   Next nZoneCount
   
   GetAllZoneAgentAmount = nAmount
End Function

'取得該地區該事務所的資料
' Input : ZoneData ==> 地區結構
'         strAgent ==> 代理人的代碼
'         nType ==> 取得資訊的種類
'            0 : 總數量
'            1 : 表取得所有事務所的發明合計數量
'            2 : 表取得所有事務所的新型合計數量
'            3 : 表取得所有事務所的設計合計數量
'         bFindAgent ==> 是否有找到該代理人的資訊
'edit by nickc 2007/01/25 解決溢位
'Private Function GetZoneAgentAmount(ByRef ZoneData As ZONEITEM, ByRef strAgent As String, ByVal nType As Integer, ByRef bFindAgent As Boolean) As Integer
Private Function GetZoneAgentAmount(ByRef ZoneData As ZONEITEM, ByRef strAgent As String, ByVal nType As Integer, ByRef bFindAgent As Boolean) As Double
   Dim nAmount As Double
   Dim nAgentCount As Integer
   Dim bFind As Boolean
   
   nAmount = 0
   bFind = False
   If ZoneData.Count - ZoneData.NoAgentItem.Count > 0 Then
      For nAgentCount = 0 To ZoneData.AgentCount - 1
         'If ZoneData.AgentList(nAgentCount).AgentCode = strAgent Then
         If ZoneData.AgentList(nAgentCount).AgentName = strAgent Then
            Select Case nType
               Case 0: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Count
               Case 1: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Type1
               Case 2: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Type2
               Case 3: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Type3
            End Select
         End If
      Next nAgentCount
   End If
   
   GetZoneAgentAmount = nAmount
End Function

' 取得所有地區台一代理資訊的數量
' Input : ZoneData ==> 地區結構
'         nType : 取得資訊的種類
'            0 : 總數量
'            1 : 表取得台一事務所的發明合計數量
'            2 : 表取得台一事務所的新型合計數量
'            3 : 表取得台一事務所的設計合計數量
'edit by nickc 2007/01/25 解決溢位
'Public Function GetAllZoneTaieAmount(ByVal nType) As Integer
Public Function GetAllZoneTaieAmount(ByVal nType) As Double
   Dim nAmount As Double
   Dim nAgentCount As Integer
   Dim nZoneCount As Integer
   
   nAmount = 0
   For nZoneCount = 0 To m_ZoneCount - 1
      If m_ZoneList(nZoneCount).Count - m_ZoneList(nZoneCount).NoAgentItem.Count > 0 Then
         For nAgentCount = 0 To m_ZoneList(nZoneCount).AgentCount - 1
            'If m_ZoneList(nZoneCount).AgentList(nAgentCount).AgentCode = "001" Then
            'If Val(m_ZoneList(nZoneCount).AgentList(nAgentCount).AgentCode = "001") = 1 Then
            If Val(m_ZoneList(nZoneCount).AgentList(nAgentCount).AgentCode) = 1 Then
               Select Case nType
                  Case 0: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Count
                  Case 1: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type1
                  Case 2: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type2
                  Case 3: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type3
               End Select
               Exit For
            End If
         Next nAgentCount
      End If
   Next nZoneCount
   GetAllZoneTaieAmount = nAmount
End Function

' 取得該地區台一代理資訊的數量
' Input : ZoneData ==> 地區結構
' nType : 取得資訊的種類
'         0 : 總數量
'         1 : 表取得台一事務所的發明合計數量
'         2 : 表取得台一事務所的新型合計數量
'         3 : 表取得台一事務所的設計合計數量
'edit by nickc 2007/01/25 解決溢位
'Private Function GetZoneTaieAmount(ByRef ZoneData As ZONEITEM, ByVal nType As Integer) As Integer
Private Function GetZoneTaieAmount(ByRef ZoneData As ZONEITEM, ByVal nType As Integer) As Double
   Dim nAmount As Double
   Dim nAgentCount As Integer
   
   nAmount = 0
   If ZoneData.Count - ZoneData.NoAgentItem.Count > 0 Then
      For nAgentCount = 0 To ZoneData.AgentCount - 1
         'If Val(ZoneData.AgentList(nAgentCount).AgentCode = "001") = 1 Then
         If Val(ZoneData.AgentList(nAgentCount).AgentCode) = 1 Then
            Select Case nType
               Case 0: nAmount = ZoneData.AgentList(nAgentCount).Count
               Case 1: nAmount = ZoneData.AgentList(nAgentCount).Type1
               Case 2: nAmount = ZoneData.AgentList(nAgentCount).Type2
               Case 3: nAmount = ZoneData.AgentList(nAgentCount).Type3
            End Select
            Exit For
         End If
      Next nAgentCount
   End If
   GetZoneTaieAmount = nAmount
End Function

' 取得所有地區代理資訊的數量
' Input : nType : 取得資訊的種類
'         0 : 總數量
'         1 : 表取得所有事務所的發明合計數量
'         2 : 表取得所有事務所的新型合計數量
'         3 : 表取得所有事務所的設計合計數量
'         nAgent : 關於代理的選項
'         0 : 全部不管有無代理事務所
'         1 : 單純無代理事務所
'         2 : 有代理事務所

Public Function GetAllZoneAmount(ByVal nType, ByVal nAgent) As Double
   Dim nAmount As Double
   Dim nAgentCount As Long
   Dim nZoneCount As Long
   
   nAmount = 0
   For nZoneCount = 0 To m_ZoneCount - 1
      Select Case nAgent
         Case 1:
            Select Case nType
               Case 1: nAmount = nAmount + m_ZoneList(nZoneCount).NoAgentItem.Type1
               Case 2: nAmount = nAmount + m_ZoneList(nZoneCount).NoAgentItem.Type2
               Case 3: nAmount = nAmount + m_ZoneList(nZoneCount).NoAgentItem.Type3
               Case 0: nAmount = nAmount + m_ZoneList(nZoneCount).NoAgentItem.Count
            End Select
         Case 2:
            If m_ZoneList(nZoneCount).Count - m_ZoneList(nZoneCount).NoAgentItem.Count > 0 Then
               For nAgentCount = 0 To m_ZoneList(nZoneCount).AgentCount - 1
                  Select Case nType
                     Case 1: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type1
                     Case 2: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type2
                     Case 3: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type3
                     Case 0: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Count
                  End Select
               Next nAgentCount
            End If
         Case 0:
            If m_ZoneList(nZoneCount).Count - m_ZoneList(nZoneCount).NoAgentItem.Count > 0 Then
               For nAgentCount = 0 To m_ZoneList(nZoneCount).AgentCount - 1
                  Select Case nType
                     Case 1: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type1
                     Case 2: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type2
                     Case 3: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Type3
                     Case 0: nAmount = nAmount + m_ZoneList(nZoneCount).AgentList(nAgentCount).Count
                  End Select
               Next nAgentCount
            End If
            Select Case nType
               Case 1: nAmount = nAmount + m_ZoneList(nZoneCount).NoAgentItem.Type1
               Case 2: nAmount = nAmount + m_ZoneList(nZoneCount).NoAgentItem.Type2
               Case 3: nAmount = nAmount + m_ZoneList(nZoneCount).NoAgentItem.Type3
               Case 0: nAmount = nAmount + m_ZoneList(nZoneCount).NoAgentItem.Count
            End Select
      End Select
   Next nZoneCount
   
   GetAllZoneAmount = nAmount
End Function

' 取得該地區代理資訊的數量
' Input : ZoneData ==> 地區結構
' nType : 取得資訊的種類
'         0 : 總數量
'         1 : 表取得所有事務所的發明合計數量
'         2 : 表取得所有事務所的新型合計數量
'         3 : 表取得所有事務所的設計合計數量
' nAgent : 關於代理的選項
'         0 : 全部不管有無代理事務所
'         1 : 單純無代理事務所
'         2 : 有代理事務所
'edit by nickc 2007/01/25 解決溢位
'Private Function GetZoneAmount(ByRef ZoneData As ZONEITEM, ByVal nType As Integer, ByVal nAgent As Integer) As Integer
Private Function GetZoneAmount(ByRef ZoneData As ZONEITEM, ByVal nType As Integer, ByVal nAgent As Integer) As Double
   Dim nAmount As Double
   Dim nAgentCount As Long
   
   nAmount = 0
   Select Case nAgent
      Case 1:
         Select Case nType
            Case 1: nAmount = ZoneData.NoAgentItem.Type1
            Case 2: nAmount = ZoneData.NoAgentItem.Type2
            Case 3: nAmount = ZoneData.NoAgentItem.Type3
            Case 0: nAmount = ZoneData.NoAgentItem.Count
         End Select
      Case 2:
         If ZoneData.Count - ZoneData.NoAgentItem.Count > 0 Then
            For nAgentCount = 0 To ZoneData.AgentCount - 1
               Select Case nType
                  Case 1: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Type1
                  Case 2: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Type2
                  Case 3: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Type3
                  Case 0: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Count
               End Select
            Next nAgentCount
         End If
      Case 0:
         If ZoneData.Count - ZoneData.NoAgentItem.Count > 0 Then
            For nAgentCount = 0 To ZoneData.AgentCount - 1
               Select Case nType
                  Case 1: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Type1
                  Case 2: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Type2
                  Case 3: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Type3
                  Case 0: nAmount = nAmount + ZoneData.AgentList(nAgentCount).Count
               End Select
            Next nAgentCount
         End If
         Select Case nType
            Case 1: nAmount = nAmount + ZoneData.NoAgentItem.Type1
            Case 2: nAmount = nAmount + ZoneData.NoAgentItem.Type2
            Case 3: nAmount = nAmount + ZoneData.NoAgentItem.Type3
            Case 0: nAmount = nAmount + ZoneData.NoAgentItem.Count
         End Select
   End Select
   
   GetZoneAmount = nAmount
End Function

' 從資料庫中取得所有的資料
Private Function GetDBData_RP(ByVal nReport As Integer) As Boolean
   Dim rsMain As New ADODB.Recordset
   Dim strSql As String
   Dim strSubSQL As String
   Dim strZone As String
   Dim strAgent As String
   Dim nZoneIndex As Integer
   Dim nAgentIndex As Integer
   Dim nCount As Integer
   Dim bFindZone As Boolean
   Dim bFindAgent As Boolean
   Dim nType As Long
   Dim nSortX As Integer
   Dim nSortY As Integer
   Dim agentTemp As AGENTITEM
   Dim ZoneTemp As ZONEITEM
   Dim bFromNation As Boolean
   Dim bToNation As Boolean
   Dim strCompany As String
   'add by nickc 2006/04/27
   Dim StrSQL4 As String
   Dim strSubSQL4 As String
   
    m_dblMaterialCnt = 0
    m_dblTotCnt = 0
   GetDBData_RP = True
   strSql = Empty
   strSubSQL = Empty
   'add by nickc 2006/04/27
   StrSQL4 = Empty
   strSubSQL4 = Empty
   ' 產生SQL查詢語法
   Select Case nReport
      Case 1:
         Select Case text03
            Case "1":
               strSql = "SELECT * FROM TPGAZETTE "
               strSubSQL = "TPG06 <= '010' "
            Case "2":
               strSql = "SELECT * FROM TPGAZETTE "
               strSubSQL = "TPG06 > '010' "
            Case " ", "":
               strSql = "SELECT * FROM TPGAZETTE "
         End Select
      'edit by nickc 2006/04/267
      'Case 2, 3, 4:
      Case 2, 4
         strSql = "SELECT * FROM TPGAZETTE "
      'add by nickc 206/04/27
      Case 3
         strSql = "SELECT TPG01,TPG02,TPG03,TPG04,TPG05,decode(tpg06,'015','C4','016','C4','039','C4'" & IIf(Check1.Value = 1, ",'020','C0'", "") & ",substr(na02,1,2))||TPG06 TPg06,TPG07,TPG08,TPG09,TPG10,TPG11,tpg12,tpg13,tpg14 FROM TPGAZETTE,nation "
         StrSQL4 = "SELECT TPG01,TPG02,TPG03,TPG04,TPG05,decode(tpg06,'015','C4ZZZ','016','C4ZZZ','039','C4ZZZ'" & IIf(Check1.Value = 1, ",'020','C0ZZZ'", "") & ",substr(na02,1,2)||decode(substr(na02,1,1),'C','ZZZ','ZZB')) TPG06,TPG07,TPG08,TPG09,tpg10,Tpg11,tpg12,tpg13,tpg14 FROM TPGAZETTE,nation "
   End Select
   
   If IsEmpty(text01_01) = False Then
      If strSubSQL <> Empty Then
         strSubSQL = strSubSQL & "AND "
         'add by nickc 2006/04/27
         strSubSQL4 = strSubSQL4 & "AND "
      End If
      strSubSQL = strSubSQL & " TPG03 >= " & ChangeTStringToWString(text01_01) & " "
      'add by nickc 2006/04/27
      strSubSQL4 = strSubSQL4 & " TPG03 >= " & ChangeTStringToWString(text01_01) & " "
   End If
   If IsEmpty(text01_02) = False Then
      If strSubSQL <> Empty Then
         strSubSQL = strSubSQL & "AND "
         'add by nickc 2006/04/27
         strSubSQL4 = strSubSQL4 & "AND "
      End If
      strSubSQL = strSubSQL & "TPG03 <= " & ChangeTStringToWString(text01_02) & " "
      'add by nickc 2006/04/27
      strSubSQL4 = strSubSQL4 & "TPG03 <= " & ChangeTStringToWString(text01_02) & " "
   End If
   If nReport = 3 Or nReport = 4 Then
      If IsEmpty(text04_01) = False Then
         If strSubSQL <> Empty Then
            strSubSQL = strSubSQL & "AND "
            'add by nickc 2006/04/27
            strSubSQL4 = strSubSQL4 & "AND "
         End If
         strSubSQL = strSubSQL & "TPG06 >= '" & text04_01 & "' "
         'add by nickc 2006/04/27
         strSubSQL4 = strSubSQL4 & "TPG06 >= '" & text04_01 & "' "
      End If
      If IsEmpty(text04_02) = False Then
         If strSubSQL <> Empty Then
            strSubSQL = strSubSQL & "AND "
            'add by nickc 2006/04/27
            strSubSQL4 = strSubSQL4 & "AND "
         End If
         strSubSQL = strSubSQL & "TPG06 <= '" & text04_02 & "' "
         'add by nickc 2006/04/27
         strSubSQL4 = strSubSQL4 & "TPG06 <= '" & text04_02 & "' "
      End If
   End If
   
   'add by nickc 2006/04/26
   If nReport = 3 Then
         If strSubSQL <> Empty Then
            strSubSQL = strSubSQL & "AND "
            strSubSQL4 = strSubSQL4 & "AND "
         End If
         strSubSQL = strSubSQL & " tpg06=na01(+) "
         strSubSQL4 = strSubSQL4 & " tpg06=na01(+) " & IIf(Check1.Value = 1, " and na02>'B' ", " and na02>'C' ") & " order by Tpg06 "
   End If
   
   If strSubSQL <> Empty Then
      strSql = strSql & "WHERE " & strSubSQL
      'add by nickc 2006/04/27
      StrSQL4 = StrSQL4 & "WHERE " & strSubSQL4
   End If
   
   ' 取得資料庫的資料
   rsMain.CursorLocation = adUseClient
   'edit by nickc 2006/04/27
   If nReport = 3 Then
      rsMain.Open strSql & " union " & StrSQL4, cnnConnection, adOpenStatic, adLockReadOnly
   Else
      rsMain.Open strSql, cnnConnection, adOpenDynamic
   End If
   ' 無資料則離開
   If rsMain.RecordCount <= 0 Then
      GetDBData_RP = False
      GoTo EXITSUB
   End If
   ' 設定初始值
   m_ZoneCount = 0
   m_AgentCount = 0
   
   rsMain.MoveFirst
   ' 依序從資料記錄中取出欄位的內容
   Do While Not rsMain.EOF
      strZone = Empty
      If IsNull(rsMain.Fields("TPG06")) = False Then
         strZone = rsMain.Fields("TPG06")
      End If
      
      ' 代理人代碼
      If IsNull(rsMain.Fields("TPG07")) = False Then
         strAgent = rsMain.Fields("TPG07")
      Else
         strAgent = Empty
      End If
      
      ' 事務所名稱
      If IsNull(rsMain.Fields("TPG08")) = False Then
         strCompany = rsMain.Fields("TPG08")
      Else
         strCompany = Empty
      End If
      
      ' 當產生表二時, 凡是地區大於010的均歸類於國外
      If nReport = 2 Then
         If strZone > "010" Then
            strZone = "999"
         End If
      End If
      
      ' 檢查申請案號的種類是屬於 發明, 新型還是設計
      nType = 0
      'Modify by Morgan 2010/12/28 申請案號改碼數
      'Select Case Mid(rsMain.Fields("TPG01"), 3, 1)
      Select Case Mid(rsMain.Fields("TPG01"), 4, 1)
         Case "1": nType = 1
         Case "2": nType = 2
         Case "3": nType = 3
      End Select
      
      ' 地區是否存在的旗標
      bFindZone = False
      ' 搜尋地區串列
      For nZoneIndex = 0 To m_ZoneCount - 1
         If m_ZoneList(nZoneIndex).ZoneCode = strZone Then
            bFindZone = True
            ' 地區數量累計
            m_ZoneList(nZoneIndex).Count = m_ZoneList(nZoneIndex).Count + 1
            ' 檢查是否無代理人並累計入該地區無代理人結構的數量中
            'If strAgent = Empty Then
            If strCompany = Empty Then
               m_ZoneList(nZoneIndex).NoAgentItem.Count = m_ZoneList(nZoneIndex).NoAgentItem.Count + 1
               Select Case nType
                  Case 1: m_ZoneList(nZoneIndex).NoAgentItem.Type1 = m_ZoneList(nZoneIndex).NoAgentItem.Type1 + 1
                  Case 2: m_ZoneList(nZoneIndex).NoAgentItem.Type2 = m_ZoneList(nZoneIndex).NoAgentItem.Type2 + 1
                  Case 3: m_ZoneList(nZoneIndex).NoAgentItem.Type3 = m_ZoneList(nZoneIndex).NoAgentItem.Type3 + 1
               End Select
            End If
            ' 累計代理人為台一
            If Val(strAgent) = 1 Then
               Select Case nType
                  Case 1: m_ZoneList(nZoneIndex).TaieCount(nType - 1) = m_ZoneList(nZoneIndex).TaieCount(0) + 1
                  Case 2: m_ZoneList(nZoneIndex).TaieCount(nType - 1) = m_ZoneList(nZoneIndex).TaieCount(1) + 1
                  Case 3: m_ZoneList(nZoneIndex).TaieCount(nType - 1) = m_ZoneList(nZoneIndex).TaieCount(2) + 1
               End Select
            End If
               
            ' 代理人是否存在的旗標
            'If strAgent <> Empty Then
            If strCompany <> Empty Then
               bFindAgent = False
               If m_ZoneList(nZoneIndex).Count - m_ZoneList(nZoneIndex).NoAgentItem.Count > 0 Then
                  ' 搜尋代理人串列
                  For nAgentIndex = 0 To m_ZoneList(nZoneIndex).AgentCount - 1
                     'If m_ZoneList(nZoneIndex).AgentList(nAgentIndex).AgentCode = strAgent Then
                     If m_ZoneList(nZoneIndex).AgentList(nAgentIndex).AgentName = strCompany Then
                        bFindAgent = True
                        m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Count = m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Count + 1
                        ' 發明或新型或設計
                        Select Case nType
                           Case 1: m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type1 = m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type1 + 1
                           Case 2: m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type2 = m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type2 + 1
                           Case 3: m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type3 = m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type3 + 1
                        End Select
                        Exit For
                     End If
                  Next nAgentIndex
               End If
               ' 無代理人資料時需產生一個代理人的結構並放入串列中
               If bFindAgent = False Then
                  ' 取得該地區中代理人串列的數目
                  nAgentIndex = 0
                  If m_ZoneList(nZoneIndex).Count - m_ZoneList(nZoneIndex).NoAgentItem.Count > 0 Then
                     nAgentIndex = m_ZoneList(nZoneIndex).AgentCount
                  End If
                  ReDim Preserve m_ZoneList(nZoneIndex).AgentList(nAgentIndex + 1)
                  m_ZoneList(nZoneIndex).AgentList(nAgentIndex).AgentCode = strAgent
                  'm_ZoneList(nZoneIndex).AgentList(nAgentIndex).AgentName = GetAgentCompany(strAgent)
                  m_ZoneList(nZoneIndex).AgentList(nAgentIndex).AgentName = strCompany
                  m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Count = 1
                  m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type1 = 0
                  m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type2 = 0
                  m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type3 = 0
                  m_ZoneList(nZoneIndex).AgentCount = m_ZoneList(nZoneIndex).AgentCount + 1
                  ' 發明或新型或設計
                  Select Case nType
                     Case 1: m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type1 = m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type1 + 1
                     Case 2: m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type2 = m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type2 + 1
                     Case 3: m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type3 = m_ZoneList(nZoneIndex).AgentList(nAgentIndex).Type3 + 1
                  End Select
               End If
            End If
            ' 離開迴圈
            Exit For
         End If
      Next nZoneIndex
      
      ' 無此地區的項目時需新增一個地區的結構並放入串列中
      If bFindZone = False Then
         If m_ZoneCount = 0 Then
            nZoneIndex = 0
         Else
            nZoneIndex = m_ZoneCount
         End If
         ' 擴大地區串列
         ReDim Preserve m_ZoneList(nZoneIndex + 1)
         m_ZoneCount = m_ZoneCount + 1
         m_ZoneList(nZoneIndex).ZoneCode = strZone
         If (strZone = "999") Then
            m_ZoneList(nZoneIndex).ZoneName = "國外"
         Else
            'edit by nickc 2006/04/27
            If nReport = 3 Then
                If Mid(strZone, 3) = "ZZZ" Then
                    m_ZoneList(nZoneIndex).ZoneName = "小計"
                Else
                    m_ZoneList(nZoneIndex).ZoneName = GetNationName(Mid(strZone, 3))
                End If
            Else
                m_ZoneList(nZoneIndex).ZoneName = GetNationName(strZone)
            End If
         End If
         ' 初始化資料
         m_ZoneList(nZoneIndex).Count = 0
         m_ZoneList(nZoneIndex).NoAgentItem.Count = 0
         m_ZoneList(nZoneIndex).NoAgentItem.Type1 = 0
         m_ZoneList(nZoneIndex).NoAgentItem.Type2 = 0
         m_ZoneList(nZoneIndex).NoAgentItem.Type3 = 0
         m_ZoneList(nZoneIndex).AgentCount = 0
         For nCount = 0 To 2
            m_ZoneList(nZoneIndex).TaieCount(nCount) = 0
         Next nCount
         ' 累計該地區的數量
         m_ZoneList(nZoneIndex).Count = m_ZoneList(nZoneIndex).Count + 1
         ' 檢查是否無代理人
         'If strAgent = Empty Then
         If strCompany = Empty Then
            m_ZoneList(nZoneIndex).NoAgentItem.Count = m_ZoneList(nZoneIndex).NoAgentItem.Count + 1
            Select Case nType
               Case 1: m_ZoneList(nZoneIndex).NoAgentItem.Type1 = m_ZoneList(nZoneIndex).NoAgentItem.Type1 + 1
               Case 2: m_ZoneList(nZoneIndex).NoAgentItem.Type2 = m_ZoneList(nZoneIndex).NoAgentItem.Type2 + 1
               Case 3: m_ZoneList(nZoneIndex).NoAgentItem.Type3 = m_ZoneList(nZoneIndex).NoAgentItem.Type3 + 1
            End Select
         End If
         ' 累計代理人為台一
         'If strAgent = "001" Then
         If Val(strAgent) = 1 Then
            Select Case nType
               Case 1, 2, 3: m_ZoneList(nZoneIndex).TaieCount(nType - 1) = m_ZoneList(nZoneIndex).TaieCount(nType - 1) + 1
            End Select
         End If
         ' 擴大該地區的代理人串列
         'If strAgent <> Empty Then
         If strCompany <> Empty Then
            ReDim Preserve m_ZoneList(nZoneIndex).AgentList(1)
            m_ZoneList(nZoneIndex).AgentList(0).AgentCode = strAgent
            m_ZoneList(nZoneIndex).AgentList(0).AgentName = strCompany
            m_ZoneList(nZoneIndex).AgentList(0).Count = 1
            m_ZoneList(nZoneIndex).AgentList(0).Type1 = 0
            m_ZoneList(nZoneIndex).AgentList(0).Type2 = 0
            m_ZoneList(nZoneIndex).AgentList(0).Type3 = 0
            m_ZoneList(nZoneIndex).AgentCount = 1
         
            ' 發明或新型或設計
            Select Case nType
               Case 1: m_ZoneList(nZoneIndex).AgentList(0).Type1 = m_ZoneList(nZoneIndex).AgentList(0).Type1 + 1
               Case 2: m_ZoneList(nZoneIndex).AgentList(0).Type2 = m_ZoneList(nZoneIndex).AgentList(0).Type2 + 1
               Case 3: m_ZoneList(nZoneIndex).AgentList(0).Type3 = m_ZoneList(nZoneIndex).AgentList(0).Type3 + 1
            End Select
         End If
      End If
      
      'If strAgent <> Empty Then
      If strCompany <> Empty Then
         ' 搜尋代理人串列
         bFindAgent = False
         For nAgentIndex = 0 To m_AgentCount - 1
            'If m_AgentList(nAgentIndex).AgentCode = strAgent Then
            If m_AgentList(nAgentIndex).AgentName = strCompany Then
               bFindAgent = True
               m_AgentList(nAgentIndex).Count = m_AgentList(nAgentIndex).Count + 1
               ' 發明或新型或設計
               Select Case nType
                  Case 1: m_AgentList(nAgentIndex).Type1 = m_AgentList(nAgentIndex).Type1 + 1
                  Case 2: m_AgentList(nAgentIndex).Type2 = m_AgentList(nAgentIndex).Type2 + 1
                  Case 3: m_AgentList(nAgentIndex).Type3 = m_AgentList(nAgentIndex).Type3 + 1
               End Select
               Exit For
            End If
         Next nAgentIndex
         If bFindAgent = False Then
            If m_AgentCount = 0 Then
               nAgentIndex = 0
            Else
               nAgentIndex = m_AgentCount
            End If
            ReDim Preserve m_AgentList(nAgentIndex + 1)
            m_AgentCount = m_AgentCount + 1
            m_AgentList(nAgentIndex).AgentCode = strAgent
            'm_AgentList(nAgentIndex).AgentName = GetAgentCompany(strAgent)
            m_AgentList(nAgentIndex).AgentName = strCompany
            m_AgentList(nAgentIndex).Count = 1
            m_AgentList(nAgentIndex).Type1 = 0
            m_AgentList(nAgentIndex).Type2 = 0
            m_AgentList(nAgentIndex).Type3 = 0
            ' 發明或新型或設計
            Select Case nType
               Case 1: m_AgentList(nAgentIndex).Type1 = m_AgentList(nAgentIndex).Type1 + 1
               Case 2: m_AgentList(nAgentIndex).Type2 = m_AgentList(nAgentIndex).Type2 + 1
               Case 3: m_AgentList(nAgentIndex).Type3 = m_AgentList(nAgentIndex).Type3 + 1
            End Select
         End If
      End If
        'Add By Cheng 2003/05/19
        '計算實體審查件數
        If "" & rsMain.Fields("TPG09") = "Y" Then m_dblMaterialCnt = m_dblMaterialCnt + 1
        '計算總件數
        m_dblTotCnt = m_dblTotCnt + 1
      ' 移到下一筆記錄
      rsMain.MoveNext
   Loop
   
   ' 對代理人串列依數量的多寡由大到小排序
   For nSortX = 0 To m_AgentCount - 1
      For nSortY = nSortX To m_AgentCount - 1
         If m_AgentList(nSortX).Count < m_AgentList(nSortY).Count Then
            agentTemp = m_AgentList(nSortX)
            m_AgentList(nSortX) = m_AgentList(nSortY)
            m_AgentList(nSortY) = agentTemp
         End If
      Next nSortY
   Next nSortX
   
   ' 地區排序 (依地區別)
   For nSortX = 0 To m_ZoneCount - 1
      For nSortY = nSortX To m_ZoneCount - 1
         Select Case nReport
            Case 1, 3, 4:
               If m_ZoneList(nSortX).ZoneCode > m_ZoneList(nSortY).ZoneCode Then
                  ZoneTemp = m_ZoneList(nSortX)
                  m_ZoneList(nSortX) = m_ZoneList(nSortY)
                  m_ZoneList(nSortY) = ZoneTemp
               End If
            Case 2:
               If GetZoneTaieAmount(m_ZoneList(nSortX), 0) < GetZoneTaieAmount(m_ZoneList(nSortY), 0) Then
                  ZoneTemp = m_ZoneList(nSortX)
                  m_ZoneList(nSortX) = m_ZoneList(nSortY)
                  m_ZoneList(nSortY) = ZoneTemp
               End If
         End Select
      Next nSortY
   Next nSortX
   
   ' 對每個地區的代理人串列做排序
   For nZoneIndex = 0 To m_ZoneCount - 1
      If m_ZoneList(nZoneIndex).Count - m_ZoneList(nZoneIndex).NoAgentItem.Count > 0 Then
         For nSortX = 0 To m_ZoneList(nZoneIndex).AgentCount - 1
            For nSortY = nSortX To m_ZoneList(nZoneIndex).AgentCount - 1
               If m_ZoneList(nZoneIndex).AgentList(nSortX).Count < m_ZoneList(nZoneIndex).AgentList(nSortY).Count Then
                  agentTemp = m_ZoneList(nZoneIndex).AgentList(nSortX)
                  m_ZoneList(nZoneIndex).AgentList(nSortX) = m_ZoneList(nZoneIndex).AgentList(nSortY)
                  m_ZoneList(nZoneIndex).AgentList(nSortY) = agentTemp
               ElseIf m_ZoneList(nZoneIndex).AgentList(nSortX).Count = m_ZoneList(nZoneIndex).AgentList(nSortY).Count Then
                  If m_ZoneList(nZoneIndex).AgentList(nSortX).AgentName > m_ZoneList(nZoneIndex).AgentList(nSortY).AgentName Then
                     agentTemp = m_ZoneList(nZoneIndex).AgentList(nSortX)
                     m_ZoneList(nZoneIndex).AgentList(nSortX) = m_ZoneList(nZoneIndex).AgentList(nSortY)
                     m_ZoneList(nZoneIndex).AgentList(nSortY) = agentTemp
                  End If
               End If
            Next nSortY
         Next nSortX
      End If
   Next nZoneIndex
   
EXITSUB:
   rsMain.Close
   Set rsMain = Nothing
End Function

Public Function Min(ByVal nValue1 As Integer, ByVal nValue2 As Integer) As Integer
   If nValue2 < nValue1 Then
      Min = nValue2
   Else
      Min = nValue1
   End If
End Function

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

Public Function Length(ByVal strData As String) As Integer
   Length = LenB(StrConv(strData, vbFromUnicode))
End Function

Public Function LeftStr(ByVal strData As String, ByVal nLen As Integer) As String
   LeftStr = StrConv(MidB(StrConv(strData, vbFromUnicode), 1, nLen), vbUnicode)
End Function

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As TextBox)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Private Sub text01_01_GotFocus()
   InverseAll text01_01
End Sub

Private Sub text01_02_GotFocus()
   InverseAll text01_02
End Sub

Private Sub text02_GotFocus()
   InverseAll text02
End Sub

Private Sub text03_GotFocus()
   InverseAll text03
End Sub

Private Sub text04_01_GotFocus()
   InverseAll text04_01
End Sub

Private Sub text04_02_GotFocus()
   InverseAll text04_02
End Sub

