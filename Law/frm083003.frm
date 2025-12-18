VERSION 5.00
Begin VB.Form frm083003 
   BorderStyle     =   1  '單線固定
   Caption         =   "收文未發文明細表"
   ClientHeight    =   3312
   ClientLeft      =   2028
   ClientTop       =   1656
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3312
   ScaleWidth      =   4800
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   7
      Left            =   1224
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2568
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   6
      Left            =   1224
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2208
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   1224
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1848
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   2544
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1488
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1224
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1488
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2544
      TabIndex        =   2
      Top             =   1128
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1224
      TabIndex        =   1
      Top             =   1128
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1224
      TabIndex        =   0
      Top             =   768
      Width           =   3135
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2904
      TabIndex        =   15
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3732
      TabIndex        =   17
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   2
      Left            =   2556
      TabIndex        =   18
      Top             =   2616
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   1
      Left            =   2556
      TabIndex        =   16
      Top             =   2256
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   2556
      TabIndex        =   14
      Top             =   1896
      Width           =   1800
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2184
      X2              =   2424
      Y1              =   1608
      Y2              =   1608
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2184
      X2              =   2424
      Y1              =   1248
      Y2              =   1248
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   5
      Left            =   270
      TabIndex        =   13
      Top             =   2565
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業  務  區："
      Height          =   180
      Index           =   4
      Left            =   264
      TabIndex        =   12
      Top             =   2208
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承  辦  人："
      Height          =   180
      Index           =   3
      Left            =   264
      TabIndex        =   11
      Top             =   1848
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   2
      Left            =   264
      TabIndex        =   10
      Top             =   1488
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文天數："
      Height          =   180
      Index           =   1
      Left            =   264
      TabIndex        =   9
      Top             =   1128
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   264
      TabIndex        =   8
      Top             =   768
      Width           =   900
   End
End
Attribute VB_Name = "frm083003"
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
'Add By Cheng 2002/09/09
Dim strTempName As String

   m_print = 0
   If Text1(0) = "" Then
      Text1(0).SetFocus
      MsgBox "系統類別不得為空值 !", vbCritical
      Exit Sub
   End If
   'Add By Cheng 2002/09/09
   If Me.Text1(1).Text <> "" And Me.Text1(2).Text <> "" Then
      If Val(Me.Text1(1).Text) > Val(Me.Text1(2).Text) Then
         MsgBox "收文天數範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.Text1(1).SetFocus
         Text1_GotFocus 1
         Exit Sub
      End If
   End If
   If Me.Text1(3).Text <> "" And Me.Text1(4).Text <> "" Then
      If Me.Text1(3).Text > Me.Text1(4).Text Then
         MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.Text1(3).SetFocus
         Text1_GotFocus 3
         Exit Sub
      End If
   End If
   '檢查承辦人
   If Me.Text1(5).Text <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetStaff(Text1(5), strTempName) Then
      If ClsPDGetStaffN(Text1(5), strTempName) Then
         Label2(0) = strTempName
      Else
         Label2(0) = ""
         Me.Text1(5).SetFocus
         Text1_GotFocus 5
         Exit Sub
      End If
   End If
   '檢查業務區
   If Me.Text1(6).Text <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objLawDll.GetStaffDeptName(Text1(6), strTempName) Then
      If ClsPDGetStaffDeptName(Text1(6), strTempName) Then
         Label2(1) = strTempName
      Else
         Label2(1) = ""
         Me.Text1(6).SetFocus
         Text1_GotFocus 6
         Exit Sub
      End If
   End If
   '檢查智權人員
   If Me.Text1(7).Text <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetStaff(Text1(7), strTempName) Then
      If ClsPDGetStaffN(Text1(7), strTempName) Then
         Label2(2) = strTempName
      Else
         Label2(2) = ""
         Me.Text1(7).SetFocus
         Text1_GotFocus 7
         Exit Sub
      End If
   End If
   
   If Text1(1) <> "" Then SDay = Format(DateAdd("d", -Val(Text1(1)), Date), "YYYYMMDD")
   If Text1(2) <> "" Then EDay = Format(DateAdd("d", -Val(Text1(2)), Date), "YYYYMMDD")
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
      'MOdify By Cheng 2002/03/26
      '取消CP09的控制
'      strExc(0) = strExc(0) & "SELECT DECODE(CP12,A0901,A0902),DECODE(CP13,S1.ST01,S1.ST02)," & _
'         "DECODE(CP05,NULL,NULL,SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2))," & _
'         "DECODE(CP06,NULL,NULL,SUBSTR(CP06,1,4)-1911||'/'||SUBSTR(CP06,5,2)||'/'||SUBSTR(CP06,7,2))," & _
'         "DECODE(CP07,NULL,NULL,SUBSTR(CP07,1,4)-1911||'/'||SUBSTR(CP07,5,2)||'/'||SUBSTR(CP07,7,2))," & _
'         "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04)," & _
'         "HC06,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
'         "DECODE(CP01||CP10,CPM01||CPM02,CPM03),DECODE(CP14,S2.ST01,S2.ST02)," & _
'         "CP64,CP12,CP13,CP05,CP01,CP02,CP03,CP04,CP09 " & _
'         "FROM STAFF S1,STAFF S2,CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,CUSTOMER,ACC090 WHERE " & _
'         "CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (SUBSTR(HC05,1,8)=CU01(+) AND " & _
'         "SUBSTR(HC05,9,1)=CU02(+)) AND (CP01=CPM01(+) AND CP10=CPM02(+)) AND CP09<'C' AND " & _
'         "CP27 IS NULL AND CP57 IS NULL AND CP12=A0901(+) AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND HC09 IS NULL" & _
'         strGetcdnSQL & " UNION "
      strExc(0) = strExc(0) & "SELECT DECODE(CP12,A0901,A0902),DECODE(CP13,S1.ST01,S1.ST02)," & _
         "DECODE(CP05,NULL,NULL,SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2))," & _
         "DECODE(CP06,NULL,NULL,SUBSTR(CP06,1,4)-1911||'/'||SUBSTR(CP06,5,2)||'/'||SUBSTR(CP06,7,2))," & _
         "DECODE(CP07,NULL,NULL,SUBSTR(CP07,1,4)-1911||'/'||SUBSTR(CP07,5,2)||'/'||SUBSTR(CP07,7,2))," & _
         "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04)," & _
         "HC06,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
         "DECODE(CP01||CP10,CPM01||CPM02,CPM03),DECODE(CP14,S2.ST01,S2.ST02)," & _
         "CP64,CP12,CP13,CP05,CP01,CP02,CP03,CP04,CP09 " & _
         "FROM STAFF S1,STAFF S2,CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,CUSTOMER,ACC090 WHERE " & _
         "CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (SUBSTR(HC05,1,8)=CU01(+) AND " & _
         "SUBSTR(HC05,9,1)=CU02(+)) AND (CP01=CPM01(+) AND CP10=CPM02(+)) AND " & _
         "CP27 IS NULL AND CP57 IS NULL AND CP12=A0901(+) AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND HC09 IS NULL" & _
         strGetcdnSQL & " UNION "
   End If
   'MOdify By Cheng 2002/03/26
   '取消CP09的控制
'   strExc(0) = strExc(0) & "SELECT DECODE(CP12,A0901,A0902),DECODE(CP13,S1.ST01,S1.ST02)," & _
'      "DECODE(CP05,NULL,NULL,SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2))," & _
'      "DECODE(CP06,NULL,NULL,SUBSTR(CP06,1,4)-1911||'/'||SUBSTR(CP06,5,2)||'/'||SUBSTR(CP06,7,2))," & _
'      "DECODE(CP07,NULL,NULL,SUBSTR(CP07,1,4)-1911||'/'||SUBSTR(CP07,5,2)||'/'||SUBSTR(CP07,7,2))," & _
'      "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04)," & _
'      "NVL(LC05, NVL(LC06, lC07)),DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
'      "DECODE(CP01||CP10,CPM01||CPM02,CPM03),DECODE(CP14,S2.ST01,S2.ST02)," & _
'      "CP64,CP12,CP13,CP05,CP01,CP02,CP03,CP04,CP09 " & _
'      "FROM STAFF S1,STAFF S2,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,CUSTOMER,ACC090 WHERE " & _
'      "CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (SUBSTR(LC11,1,8)=CU01(+) AND " & _
'      "SUBSTR(LC11,9,1)=CU02(+)) AND (CP01=CPM01(+) AND CP10=CPM02(+)) AND CP09<'C' AND " & _
'      "CP27 IS NULL AND CP57 IS NULL AND CP12=A0901(+) AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND LC08 IS NULL" & _
'      strGetcdnSQL & " ORDER BY CP12,CP13,CP05,CP01,CP02,CP03,CP04"
   strExc(0) = strExc(0) & "SELECT DECODE(CP12,A0901,A0902),DECODE(CP13,S1.ST01,S1.ST02)," & _
      "DECODE(CP05,NULL,NULL,SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2))," & _
      "DECODE(CP06,NULL,NULL,SUBSTR(CP06,1,4)-1911||'/'||SUBSTR(CP06,5,2)||'/'||SUBSTR(CP06,7,2))," & _
      "DECODE(CP07,NULL,NULL,SUBSTR(CP07,1,4)-1911||'/'||SUBSTR(CP07,5,2)||'/'||SUBSTR(CP07,7,2))," & _
      "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04)," & _
      "NVL(LC05, NVL(LC06, lC07)),DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
      "DECODE(CP01||CP10,CPM01||CPM02,CPM03),DECODE(CP14,S2.ST01,S2.ST02)," & _
      "CP64,CP12,CP13,CP05,CP01,CP02,CP03,CP04,CP09 " & _
      "FROM STAFF S1,STAFF S2,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,CUSTOMER,ACC090 WHERE " & _
      "CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND (SUBSTR(LC11,1,8)=CU01(+) AND " & _
      "SUBSTR(LC11,9,1)=CU02(+)) AND (CP01=CPM01(+) AND CP10=CPM02(+)) AND " & _
      "CP27 IS NULL AND CP57 IS NULL AND CP12=A0901(+) AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND LC08 IS NULL" & _
      strGetcdnSQL & " ORDER BY CP12,CP13,CP05,CP01,CP02,CP03,CP04"
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
      If IsNull(.Fields(0)) = False Then Printer.Print Format(.Fields(0), "!@@@@@")
      Printer.CurrentX = PLeft(1):      Printer.CurrentY = iPrint
      If IsNull(.Fields(1)) = False Then Printer.Print Format(.Fields(1).Value, "!@@@@")
      Printer.CurrentX = PLeft(2):      Printer.CurrentY = iPrint
      Printer.Print .Fields(2)
      Printer.CurrentX = PLeft(3):      Printer.CurrentY = iPrint
      If .Fields(3) <> "//" Then
         If DBDATE(.Fields(3)) > Val(GetTaiwanTodayDate + 19110000) Then
            St = .Fields(3)
         ElseIf .Fields(3) = ChangeTStringToTDateString(GetTaiwanTodayDate) Then
            St = "V" & .Fields(3)
         ElseIf DBDATE(.Fields(3)) < Val(GetTaiwanTodayDate + 19110000) Then
            St = "*" & .Fields(3)
         End If
      Else
         St = ""
      End If
      Printer.Print St
      Printer.CurrentX = PLeft(4):      Printer.CurrentY = iPrint
      If .Fields(4) <> "//" Then
         St = .Fields(4)
      Else
         St = ""
      End If
      Printer.Print St
      Printer.CurrentX = PLeft(5):      Printer.CurrentY = iPrint
      Printer.Print .Fields(5)
      Printer.CurrentX = PLeft(6):      Printer.CurrentY = iPrint
      Printer.Print Format(Left(.Fields(6), 8), "!@@@@@@@@")
      Printer.CurrentX = PLeft(7):      Printer.CurrentY = iPrint
      If IsNull(.Fields(7)) = False Then Printer.Print Format(Left(.Fields(7), 6), "!@@@@@@")
      Printer.CurrentX = PLeft(8):      Printer.CurrentY = iPrint
      If IsNull(.Fields(8)) = False Then Printer.Print Format(Left(.Fields(8), 6), "!@@@@@@")
      Printer.CurrentX = PLeft(9):      Printer.CurrentY = iPrint
      If IsNull(.Fields(9)) = False Then Printer.Print Format(.Fields(9), "!@@@@")
      Printer.CurrentX = PLeft(10):     Printer.CurrentY = iPrint
      If IsNull(.Fields(10)) = False Then Printer.Print .Fields(10)
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
   PLeft(0) = 500:     PLeft(1) = 1750
   PLeft(2) = 2700:    PLeft(3) = 3800
   PLeft(4) = 5000:    PLeft(5) = 6200
   PLeft(6) = 7700:    PLeft(7) = 10100
   PLeft(8) = 12000:   PLeft(9) = 13100
   PLeft(10) = 14200
End Sub

Private Sub CaseTitle(ByVal Area As String, ByVal Page As String)
 Dim i As Integer
  
   i = 500
   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000:         Printer.CurrentY = i
   Printer.Print "收文未發文明細表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 6200:         Printer.CurrentY = i + 500
   Printer.Print "收文日期 : " & Left(EDay, 4) - 1911 & "/" & Mid(EDay, 5, 2) & _
      "/" & Right(EDay, 2) & " - " & Left(SDay, 4) - 1911 & "/" & Mid(SDay, 5, 2) & _
      "/" & Right(SDay, 2)
   Printer.Font.Bold = False
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 13000:            Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
   Printer.CurrentX = 13000:            Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
   Printer.CurrentX = 500:            Printer.CurrentY = i + 1100
   Printer.Print "收文天數 : " & Text1(1).Text & "-" & Text1(2).Text
   Printer.CurrentX = 500:              Printer.CurrentY = i + 1400
   Printer.Print String(205, "-")
   Printer.CurrentX = PLeft(0):         Printer.CurrentY = i + 1700
   Printer.Print "業務區"
   Printer.CurrentX = PLeft(1):         Printer.CurrentY = i + 1700
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(2):         Printer.CurrentY = i + 1700
   Printer.Print "收文日"
   Printer.CurrentX = PLeft(3):         Printer.CurrentY = i + 1700
   Printer.Print "本所期限"
   Printer.CurrentX = PLeft(4):         Printer.CurrentY = i + 1700
   Printer.Print "法定期限"
   Printer.CurrentX = PLeft(5):         Printer.CurrentY = i + 1700
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(6):         Printer.CurrentY = i + 1700
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(7):         Printer.CurrentY = i + 1700
   Printer.Print "當事人"
   Printer.CurrentX = PLeft(8):         Printer.CurrentY = i + 1700
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(9):         Printer.CurrentY = i + 1700
   Printer.Print "承辦人"
   Printer.CurrentX = PLeft(10):        Printer.CurrentY = i + 1700
   Printer.Print "進度備註"
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
Dim strCP01 As String
Dim strTemp As Variant
 
   If Me.Tag = 0 Then
      strTemp = Split(UCase(Text1(0).Text), ",")
      For i = 0 To UBound(strTemp)
          If strTemp(i) = "L" Or strTemp(i) = "LA" Then
             strCP01 = strCP01 & strTemp(i)
             strCP01 = strCP01 & "','"
          End If
      Next i
   ElseIf Me.Tag = 1 Then
      strTemp = Split(UCase(Text1(0).Text), ",")
      For i = 0 To UBound(strTemp)
          'Modify By Sindy 2009/07/24 增加LIN系統類別
          If strTemp(i) = "CFL" Or strTemp(i) = "FCL" Or strTemp(i) = "LIN" Then
             strCP01 = strCP01 & strTemp(i)
             strCP01 = strCP01 & "','"
          End If
      Next i
   'add by sonia 2019/8/6 +ACS系統類別
   ElseIf Me.Tag = 2 Then
      strTemp = Split(UCase(Text1(0).Text), ",")
      For i = 0 To UBound(strTemp)
          If strTemp(i) = "ACS" Then
             strCP01 = strCP01 & strTemp(i)
             strCP01 = strCP01 & "','"
          End If
      Next i
   'end 2019/8/6
   End If
      
   If SDay = "" And EDay <> "" Then
      strExc(1) = " AND CP01 IN('" & strCP01 & "') AND CP05>='" & EDay & "'"
   ElseIf EDay <> "" And SDay <> "" Then
      strExc(1) = " AND CP01 IN('" & strCP01 & "') AND (CP05 BETWEEN '" + _
         EDay + "' AND '" + SDay + "')"
   End If
   If Text1(5).Text <> "" Then strExc(1) = strExc(1) + " and CP14='" + Text1(5).Text + "'"
   If Text1(6).Text <> "" Then strExc(1) = strExc(1) + " and CP12='" + Text1(6).Text + "'"
   If Text1(7).Text <> "" Then strExc(1) = strExc(1) + " and CP13='" + Text1(7).Text + "'"
   If Text1(3).Text = "" And Text1(4).Text <> "" Then
      strExc(1) = strExc(1) + " and CP10 <='" + Text1(4) + "'"
   ElseIf Text1(3).Text <> "" And Text1(4).Text <> "" Then
      strExc(1) = strExc(1) + " and (CP10 BETWEEN '" + Text1(3) + "' AND '" + Text1(4) + "')"
   End If
   strGetcdnSQL = strExc(1)
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm083003 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   'If Index <> 7 Then KeyAscii = UpperCase(KeyAscii)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 2, 4
            If RunNick(Text1(Index - 1), Text1(Index)) Then
               Text1(Index - 1).SetFocus
            End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTempName As String, i As Integer, t As Integer
 Dim strTemp1 As Variant
 Dim strTemp2 As Variant
 Dim s As Integer
 Dim j As Integer
 Dim nPos As Integer
 Dim strProperty As String
 
   If Text1(Index) = "" Then
      If Index = 5 Or Index = 6 Or Index = 7 Then
         Label2(Index - 5).Caption = ""
      End If
      Exit Sub
   End If
   Select Case Index
      Case 0
         'If ChkSysName(Text1(Index)) = False Then Cancel = True
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(Text1(Index)), ",,", ""), ",")
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
         
      Case 6
         'edit by nickc 2007/02/07 不用 dll 了
         'If objLawDll.GetStaffDeptName(Text1(Index), strTempName) Then
         If ClsPDGetStaffDeptName(Text1(Index), strTempName) Then
            Label2(1) = strTempName
         Else
            Cancel = True
            Label2(1) = ""
         End If
      Case 5, 7
         If Index = 5 Then
            i = 0
         ElseIf Index = 7 Then
            i = 2
         End If
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(Text1(Index), strTempName) Then
         If ClsPDGetStaffN(Text1(Index), strTempName) Then
            Label2(i) = strTempName
         Else
            Label2(i) = ""
            Cancel = True
         End If
      End Select
      If Cancel Then TextInverse Text1(Index)
End Sub
