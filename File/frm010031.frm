VERSION 5.00
Begin VB.Form frm010031 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文清單"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4920
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   4
      Top             =   1830
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   1
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   2280
      MaxLength       =   7
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3885
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2940
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發文日期："
      Height          =   180
      Left            =   1350
      TabIndex        =   6
      Top             =   870
      Width           =   900
   End
End
Attribute VB_Name = "frm010031"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
'Create by Sindy 2009/03/30
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_str2 As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 12) As Integer
Dim strTemp(1 To 12) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmt As Double, dblAmt2 As Double


Private Sub cmdok_Click(Index As Integer)
Dim strMsgText As String
   
   Select Case Index
   Case 0
           If Trim(txt1(0)) = "" Then
               MsgBox "發文日期不可空白！", vbInformation, "操作錯誤！"
               txt1(0).SetFocus
               Exit Sub
           Else
               'Add By Sindy 2009/05/04
               '檢查CP27 is null並且CP124>0,顯示警示視窗
               If GetNoSend2(strMsgText) = True Then
                  MsgBox strMsgText & "有錯誤之發文資料, 請通知電腦中心！", vbExclamation + vbOKOnly
                  Exit Sub
               End If
               '2009/05/04 End
               '檢查是否有資料尚待發文
               If GetNoSend = True Then
                  MsgBox "尚有資料待發文！", vbExclamation + vbOKOnly
                  Exit Sub
               End If
           End If
           
           Screen.MousePointer = vbHourglass
           m_StrSQL = ""
           If txt1(0) <> "" Then
               m_StrSQL = m_StrSQL & " AND CP124=" & ChangeTStringToWString(txt1(0)) & " "
           End If
           StrMenu1
           Screen.MousePointer = vbDefault
   Case 1
           Unload Me
   Case Else
   End Select
End Sub


Sub StrMenu1()
Dim intSendCnt As Integer
Dim m_j As Integer, intRow As Integer

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印

m_str = "SELECT * FROM ( "
m_str = m_str & "SELECT substr(CP28,3,6),CP01||'-'||CP02||'-'||CP03||'-'||CP04,substr(PA05,1,6),substr(CPM03,1,6),TO_CHAR(NVL(CP84,0),'9G999G999'),PA11, " & _
                              "substr(A0902,1,2), " & _
                              "CP130,CP123 " & _
                  "FROM Patent,CaseProgress,CasePropertyMap,Customer,Staff S1,ACC090 " & _
                "WHERE CP01 = PA01 And cp02 = pa02 And cp03 = pa03 And cp04 = pa04 " & _
                     "AND substr(PA26,1,8)=CU01(+) " & _
                     "AND substr(PA26,9,1)=CU02(+) " & _
                     "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                     "AND CP83=S1.ST01(+) " & _
                     "AND S1.ST03=A0901(+) " & m_StrSQL & " union all "
m_str = m_str & "SELECT substr(CP28,3,6),CP01||'-'||CP02||'-'||CP03||'-'||CP04,substr(TM05,1,6),substr(CPM03,1,6),TO_CHAR(NVL(CP84,0),'9G999G999'),NVL(TM15,TM12), " & _
                              "substr(A0902,1,2), " & _
                              "CP130,CP123 " & _
                  "FROM TradeMark,CaseProgress,CasePropertyMap,Customer,Staff S1,ACC090 " & _
                "WHERE CP01 = tm01 And cp02 = tm02 And cp03 = tm03 And cp04 = tm04 " & _
                     "AND substr(TM23,1,8)=CU01(+) " & _
                     "AND substr(TM23,9,1)=CU02(+) " & _
                     "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                     "AND CP83=S1.ST01(+) " & _
                     "AND S1.ST03=A0901(+) " & m_StrSQL & " union all "
m_str = m_str & "SELECT substr(CP28,3,6),CP01||'-'||CP02||'-'||CP03||'-'||CP04,substr(SP05,1,6),substr(CPM03,1,6),TO_CHAR(NVL(CP84,0),'9G999G999'),SP11, " & _
                              "substr(A0902,1,2), " & _
                              "CP130,CP123 " & _
                  "FROM ServicePractice,CaseProgress,CasePropertyMap,Customer,Staff S1,ACC090 " & _
                "WHERE CP01 = SP01 And cp02 = SP02 And cp03 = SP03 And cp04 = SP04 " & _
                     "AND substr(SP08,1,8)=CU01(+) " & _
                     "AND substr(SP08,9,1)=CU02(+) " & _
                     "AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                     "AND CP83=S1.ST01(+) " & _
                     "AND S1.ST03=A0901(+) " & m_StrSQL
m_str = m_str & ") Order By 1,2 ASC "

If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        .MoveFirst
        
        iLine = 1
        strType = ""
        intSendCnt = 0
        
        Do While Not .EOF
            For m_i = 1 To 8
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(0))
            strTemp(2) = CheckStr(m_rs.Fields(1))
            strTemp(3) = CheckStr(m_rs.Fields(2))
            strTemp(4) = CheckStr(m_rs.Fields(3))
            strTemp(5) = CheckStr(m_rs.Fields(4))
            strTemp(6) = CheckStr(m_rs.Fields(5))
            strTemp(7) = CheckStr(m_rs.Fields(6))
            strTemp(8) = CheckStr(m_rs.Fields(7))
            If Trim(CheckStr(m_rs.Fields(8))) = "Y" Then '是否算發文室件數
               intSendCnt = intSendCnt + 1
            End If
            
            If iLine > 51 Or iLine = 1 Then
               'If .AbsolutePosition <> .RecordCount Then
                   If strType <> "" Then Printer.NewPage
                   iLine = 1
                   PrintTitle '列印表頭
               'End If
            End If
            PrintDetail
            
            strType = CheckStr(m_rs.Fields(0))
            .MoveNext
        Loop
        '合計
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iLine * 300
        Printer.Print String(160, "-")
        
        iLine = iLine + 1
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iLine * 300
        Printer.Print "合　計：" & .RecordCount & " 筆"
        Printer.CurrentX = PLeft(7)
        Printer.CurrentY = iLine * 300
        Printer.Print "算件數：" & intSendCnt & " 筆"
        iLine = iLine + 3
        
        '最後加印取消發文的資料
        m_str = "select CP01||'-'||CP02||'-'||CP03||'-'||CP04,CPM03,CP09,CP131 " & _
                     "From caseprogress, casepropertymap " & _
                     "Where cp132=" & ChangeTStringToWString(txt1(0)) & " " & _
                     "and cp01=cpm01(+) " & _
                     "and cp10=cpm02(+) "
        If m_rs.State = 1 Then m_rs.Close
        m_rs.CursorLocation = adUseClient
        m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
        intRow = 0
        If Not m_rs.EOF And Not m_rs.BOF Then
            With m_rs
               .MoveFirst
               Do While Not .EOF
                  intRow = intRow + 1
                  If iLine > 51 Or iLine = 1 Then
                     'If .AbsolutePosition <> .RecordCount Then
                         Printer.NewPage
                         iLine = 1
                         PrintTitle '列印表頭
                     'End If
                  End If
                  For m_j = 0 To 3
                     If m_j = 0 Then
                        Printer.CurrentX = PLeft(1)
                        Printer.CurrentY = iLine * 300
                        Printer.Print "取消 " '& intRow
                     End If
                     Printer.CurrentX = PLeft(m_j + 2)
                     Printer.CurrentY = iLine * 300
                     Printer.Print m_rs.Fields(m_j)
                  Next m_j
                  iLine = iLine + 1
                  .MoveNext
               Loop
            End With
        End If
    End With
Else
    ShowNoData
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft()
PLeft(1) = 1300
PLeft(2) = 2200
PLeft(3) = 3700
PLeft(4) = 5000
PLeft(5) = 6800 '規費
PLeft(6) = 7000
PLeft(7) = 7900
PLeft(8) = 8400
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(txt1(0) & " 發文室發文清單") / 2)
Printer.CurrentY = iLine * 300
Printer.Print txt1(0) & " 發文室發文清單"

Printer.Font.Size = 10
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "列印人：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 4
Printer.CurrentX = PLeft(1)
Printer.CurrentY = (iLine + 1) * 300
Printer.Print "發文字號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = (iLine + 1) * 300
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = (iLine + 1) * 300
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = (iLine + 1) * 300
Printer.Print "案件性質"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("發文規費")
Printer.CurrentY = (iLine + 1) * 300
Printer.Print "發文規費"

Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "申請號/"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = (iLine + 1) * 300
Printer.Print "審定號"

Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "發文"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = (iLine + 1) * 300
Printer.Print "部門"

Printer.CurrentX = PLeft(8)
Printer.CurrentY = (iLine + 1) * 300
Printer.Print "發文對象"

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(160, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 8
   If m_j = 5 Then
      Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
   Else
      Printer.CurrentX = PLeft(m_j)
   End If
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSql Then
         SeekPrint = i
      End If
   Next i
   
   '發文日期
   txt1(0).Text = strSrvDate(2)
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm010031 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    InverseTextBox txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
         'KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index).Text <> "" Then
            If ChkDate(txt1(Index)) = False Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub

'檢查是否有文尚待發文
Public Function GetNoSend() As Boolean
Dim strSql As String
   CheckOC3
   GetNoSend = False
   '2012/10/22 MODIFY BY SONIA 加CP118 IS NULL 條件
   strSql = "SELECT * FROM CaseProgress,Staff S1 " & _
                   "WHERE CP27=" & ChangeTStringToWString(txt1(0)) & _
                        " AND CP123 is not null AND CP28 is null AND CP118 IS NULL " & _
                        " AND CP83=S1.ST01(+) " & _
                        " AND (S1.ST03 like 'F1%' or S1.ST03 like 'F2%' or S1.ST03 like 'P1%' or S1.ST03 like 'P2%') "
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
      GetNoSend = True
   End If
   CheckOC3
End Function

'Add By Sindy 2009/05/04
'檢查CP27 is null並且CP124>0,顯示警示視窗
Public Function GetNoSend2(strMsgText As String) As Boolean
Dim strSql As String
CheckOC3
GetNoSend2 = False
strMsgText = ""
strSql = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,CP09,CP10 FROM CaseProgress WHERE CP27 is null AND CP124 > 0 "
AdoRecordSet3.CursorLocation = adUseClient
AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
   GetNoSend2 = True
   With AdoRecordSet3
      AdoRecordSet3.MoveFirst
      Do While Not AdoRecordSet3.EOF
         strMsgText = strMsgText & "案號：" & AdoRecordSet3.Fields(0) & " 文號：" & AdoRecordSet3.Fields(1) & " 案件性質：" & AdoRecordSet3.Fields(2) & vbCrLf
         AdoRecordSet3.MoveNext
      Loop
   End With
End If
CheckOC3
End Function
