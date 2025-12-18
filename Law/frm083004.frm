VERSION 5.00
Begin VB.Form frm083004 
   BorderStyle     =   1  '單線固定
   Caption         =   "催審表"
   ClientHeight    =   2100
   ClientLeft      =   2670
   ClientTop       =   2070
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4545
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2652
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1416
      TabIndex        =   0
      Top             =   792
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1416
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1152
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2736
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1152
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   336
      TabIndex        =   6
      Top             =   792
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "催審期限："
      Height          =   180
      Index           =   1
      Left            =   336
      TabIndex        =   5
      Top             =   1152
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2376
      X2              =   2616
      Y1              =   1272
      Y2              =   1272
   End
End
Attribute VB_Name = "frm083004"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim SDay As String, EDay As String, PLeft(0 To 10) As Integer
Dim m_print As Integer

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   m_print = 0
   If Text1(0) = "" Then
      Text1(0).SetFocus
      MsgBox "系統類別不得為空值 !", vbCritical
      Exit Sub
   End If
   If ChkRange(Text1(1), Text1(2), "催審期限") = False Then Exit Sub
   'Add By Cheng 2002/03/22
   If PUB_CheckKeyInDate(Me.Text1(1)) = -1 Then
      Me.Text1(1).SetFocus
      Text1_GotFocus 1
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text1(2)) = -1 Then
      Me.Text1(2).SetFocus
      Text1_GotFocus 2
      Exit Sub
   End If
   
   Screen.MousePointer = 11
   GetPrintLeft
   PrintCase
   Screen.MousePointer = 0
   If m_print = 0 Then
      MsgBox "列印結束!", vbInformation
   End If
End Sub

Private Sub PrintCase()
 Dim i As Integer, St As String, Page As Integer, iPrint As Integer
 Dim TmpArea As String
On Error GoTo ErrHand
   strExc(0) = ""
   If Me.Tag = 0 Then
      strExc(0) = "SELECT decode(LENGTH(CP27),NULL,NULL,SUBSTR(CP27,1,4)-1911||'/'||SUBSTR(CP27,5,2)||'/'||SUBSTR(CP27,7,2))," & _
         "NP02||'-'||NP03||DECODE(NP04,'0','','-'||NP04)||DECODE(NP05,'00','','-'||NP05)," & _
         "HC06,CP35,DECODE(CP01||CP10,CPM01||CPM02,CPM03)," & _
         "DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))),DECODE(CP13,S2.ST01,S2.ST02)," & _
         "DECODE(CP14,S1.ST01,S1.ST02),NP09,CP27,NP02,NP03,NP04,NP05," & _
         "decode(LENGTH(NP08),NULL,NULL,SUBSTR(NP08,1,4)-1911||'/'||SUBSTR(NP08,5,2)||'/'||SUBSTR(NP08,7,2)) AS NP08N,NP22 " & _
         "FROM STAFF S1,STAFF S2,CASEPROGRESS," & _
         "HIRECASE,CUSTOMER,CASEPROPERTYMAP,NEXTPROGRESS WHERE " & _
         "CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND (CP01=CPM01(+) AND CP10=CPM02(+)) AND " & _
         "NP01=CP09 AND NP02=HC01 AND NP03=HC02 AND NP04=HC03 AND NP05=HC04" & _
         " and (NP07='6001' AND NP06 IS NULL) AND (SUBSTR(HC05,1,8)=CU01(+) AND HC09 IS NULL AND " & _
         "SUBSTR(HC05,9,1)=CU02(+)) " & strGetcdnSQL & " UNION "
   End If
   strExc(0) = strExc(0) & "SELECT DECODE(LENGTH(CP27),NULL,NULL,SUBSTR(CP27,1,4)-1911||'/'||SUBSTR(CP27,5,2)||'/'||SUBSTR(CP27,7,2))," & _
      "NP02||'-'||NP03||DECODE(NP04,'0','','-'||NP04)||DECODE(NP05,'00','','-'||NP05)," & _
      "NVL(LC05, NVL(LC06,LC07)),CP35,DECODE(CP01||CP10,CPM01||CPM02,CPM03)," & _
      "DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06))),DECODE(CP13,S2.ST01,S2.ST02)," & _
      "DECODE(CP14,S1.ST01,S1.ST02),NP09,CP27,NP02,NP03,NP04,NP05," & _
      "decode(LENGTH(NP08),NULL,NULL,SUBSTR(NP08,1,4)-1911||'/'||SUBSTR(NP08,5,2)||'/'||SUBSTR(NP08,7,2)) AS NP08N,NP22 " & _
      "FROM STAFF S1,STAFF S2,CASEPROGRESS," & _
      "LAWCASE,CUSTOMER,CASEPROPERTYMAP,NEXTPROGRESS WHERE " & _
      "CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND (CP01=CPM01(+) AND CP10=CPM02(+)) AND " & _
      "NP01=CP09 AND NP02=LC01 AND NP03=LC02 AND NP04=LC03 AND NP05=LC04 AND " & _
      "(NP07='6001' AND NP06 IS NULL) AND (SUBSTR(LC11,1,8)=CU01(+) AND LC08 IS NULL AND " & _
      "SUBSTR(LC11,9,1)=CU02(+)) " & strGetcdnSQL & " ORDER BY NP08N,NP05,NP02,NP03,NP04"
   If RsTemp.State = adStateOpen Then RsTemp.Close
   RsTemp.Open strExc(0), cnnConnection
   If RsTemp.EOF And RsTemp.BOF Then
      MsgBox "資料庫內無資料 !", vbInformation
      m_print = 1
      Exit Sub
   End If
   i = 1
   Page = 1
   CaseTitle TmpArea, 1
   iPrint = 2700
   With RsTemp
   Do While Not .EOF
      Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
      If .Fields(0) <> "//" Then
         St = .Fields(0)
      Else
         St = ""
      End If
      Printer.Print St  '發文日
      '催審期限
      Printer.CurrentX = PLeft(8):      Printer.CurrentY = iPrint
      If IsNull(.Fields("NP08N")) = False Then Printer.Print .Fields("NP08N")

      Printer.CurrentX = PLeft(1):      Printer.CurrentY = iPrint
      Printer.Print .Fields(1) '本所案號
      '案件名稱
      Printer.CurrentX = PLeft(2):      Printer.CurrentY = iPrint
      If IsNull(.Fields(2)) = False Then Printer.Print Format(Left(.Fields(2), 8), "!@@@@@@@@")
      '法院案號
      Printer.CurrentX = PLeft(3):      Printer.CurrentY = iPrint
      If IsNull(.Fields(3)) = False Then Printer.Print Format(Left(.Fields(3), 8), "!@@@@@@@@")
      '案件性質
      Printer.CurrentX = PLeft(4):      Printer.CurrentY = iPrint
      If IsNull(.Fields(4)) = False Then Printer.Print Format(Left(.Fields(4), 6), "!@@@@@@")
      
      Printer.CurrentX = PLeft(5):      Printer.CurrentY = iPrint
      If IsNull(.Fields(5)) = False Then Printer.Print Format(Left(.Fields(5), 6), "!@@@@@@")
      Printer.CurrentX = PLeft(6):      Printer.CurrentY = iPrint
      If IsNull(.Fields(6)) = False Then Printer.Print Format(.Fields(6), "!@@@@")
      Printer.CurrentX = PLeft(7):      Printer.CurrentY = iPrint
      If IsNull(.Fields(7)) = False Then Printer.Print Format(.Fields(7), "!@@@@")
      iPrint = iPrint + 300
      .MoveNext
      If Not .EOF Then
         If (i Mod 27 = 0) Then
            Printer.NewPage
            Page = Page + 1
            CaseTitle St, Page
            iPrint = 2700
            i = 0
         End If
         i = i + 1
         
      End If
   Loop
   End With
   Printer.EndDoc
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub GetPrintLeft()
   PLeft(0) = 1600   '發文日
   PLeft(1) = 2700   '本所案號
   PLeft(2) = 4100   '案件名稱
   PLeft(3) = 6100   '法院案號
   PLeft(4) = 8100   '案件性質
   PLeft(5) = 9700   '當事人
   PLeft(6) = 12100  '智權人員
   PLeft(7) = 13600  '承辦人
   PLeft(8) = 500    '催審期限
End Sub

Private Sub CaseTitle(ByVal Area As String, ByVal Page As String)
 Dim i As Integer, St As String
   i = 500
   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6500:         Printer.CurrentY = i
   Printer.Print "催  審  表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 6200:         Printer.CurrentY = i + 500
   Printer.Print "催審期限 : " & ChangeTStringToTDateString(Text1(1)) & _
      " - " & ChangeTStringToTDateString(Text1(2))
   Printer.Font.Bold = False
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 13000:            Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
   Printer.CurrentX = 13000:            Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
   Printer.CurrentX = 500:              Printer.CurrentY = i + 1400
   Printer.Print String(205, "-")
   Printer.CurrentX = PLeft(8):         Printer.CurrentY = i + 1700
   Printer.Print "催審期限"
   Printer.CurrentX = PLeft(0):         Printer.CurrentY = i + 1700
   Printer.Print "發文日"
   Printer.CurrentX = PLeft(1):         Printer.CurrentY = i + 1700
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(2):         Printer.CurrentY = i + 1700
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(3):         Printer.CurrentY = i + 1700
   Printer.Print "法院案號"
   Printer.CurrentX = PLeft(4):         Printer.CurrentY = i + 1700
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(5):         Printer.CurrentY = i + 1700
   Printer.Print "當事人"
   Printer.CurrentX = PLeft(6):         Printer.CurrentY = i + 1700
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(7):         Printer.CurrentY = i + 1700
   Printer.Print "承辦人"
   Printer.CurrentX = 500:          Printer.CurrentY = i + 2000
   Printer.Print String(205, "-")
End Sub

Private Sub Form_Activate()
  Text1(0).SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Text1(0).Text = GetSystemKindByNick
End Sub

Private Function strGetcdnSQL() As String
 Dim i As Integer, strcpdate As String
 Dim strNP02 As String
 Dim strTemp As Variant
 Dim strSql As String
 
 strSql = ""
 'strNP02 = Replace(Text1(0).Text, ",", "','")
 If Me.Tag = 0 Then
    strTemp = Split(UCase(Text1(0).Text), ",")
    For i = 0 To UBound(strTemp)
        If strTemp(i) = "L" Or strTemp(i) = "LA" Then
           strSql = strSql & strTemp(i)
           strSql = strSql & "','"
        End If
        
    Next i
 ElseIf Me.Tag = 1 Then
    strTemp = Split(UCase(Text1(0).Text), ",")
    For i = 0 To UBound(strTemp)
        'Modify By Sindy 2009/07/24 增加LIN系統類別
        If strTemp(i) = "CFL" Or strTemp(i) = "FCL" Or strTemp(i) = "LIN" Then
           strSql = strSql & strTemp(i)
           strSql = strSql & "','"
        End If
        
    Next i
 End If
   strExc(1) = " AND NP02 in('" & strSql & "')"
   If Text1(1).Text = "" And Text1(2).Text <> "" Then
      strExc(1) = strExc(1) + " and NP09<='" + ChangeTStringToWString(Text1(2)) + "'"
   ElseIf Text1(1).Text <> "" And Text1(2).Text <> "" Then
      strExc(1) = strExc(1) + " and (NP09 BETWEEN '" + ChangeTStringToWString(Text1(1)) + "' AND '" + ChangeTStringToWString(Text1(2)) + "')"
   End If
   strGetcdnSQL = strExc(1)
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm083004 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 1, 2
        If PUB_CheckKeyInDate(Text1(Index)) = -1 Then
           Text1(Index).SetFocus
           Text1_GotFocus (Index)
           Exit Sub
        End If
          If Index = 2 Then
            If RunNick(Text1(Index - 1), Text1(Index)) Then
               Text1(Index - 1).SetFocus
               Text1_GotFocus (Index - 1)
               Exit Sub
            End If
          End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTempName As String, i As Integer, t As Integer
 Dim strTemp1 As Variant
 Dim strTemp2 As Variant
 Dim j As Integer
 Dim s As Integer
 
   If Text1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0
        ' If ChkSysName(Text1(Index)) = False Then Cancel = True
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(Text1(Index).Text), ",,", ""), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            Cancel = True
            Exit Sub
        End If
     Next i
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub
