VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm083001 
   BorderStyle     =   1  '單線固定
   Caption         =   "期限管制表"
   ClientHeight    =   3888
   ClientLeft      =   2400
   ClientTop       =   1488
   ClientWidth     =   5052
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3888
   ScaleWidth      =   5052
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   4344
      Top             =   2388
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   1
      Left            =   1248
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1092
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   3
      Left            =   1248
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1452
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   2
      Left            =   2688
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1092
      Width           =   1092
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3168
      TabIndex        =   11
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3996
      TabIndex        =   12
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   0
      Left            =   1248
      TabIndex        =   0
      Top             =   732
      Width           =   3255
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   4
      Left            =   1248
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1812
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   5
      Left            =   1248
      MaxLength       =   6
      TabIndex        =   5
      Top             =   2172
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   7
      Left            =   1248
      MaxLength       =   9
      TabIndex        =   7
      Top             =   2892
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   8
      Left            =   2688
      MaxLength       =   9
      TabIndex        =   8
      Top             =   2892
      Width           =   1092
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00FFFFFF&
      Height          =   264
      Index           =   6
      Left            =   1248
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2532
      Width           =   495
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   9
      Left            =   1248
      MaxLength       =   9
      TabIndex        =   9
      Top             =   3252
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Height          =   264
      Index           =   10
      Left            =   2688
      MaxLength       =   9
      TabIndex        =   10
      Top             =   3252
      Width           =   1092
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(1.智權人員  2.承辦人 3.全部)"
      Height          =   180
      Index           =   8
      Left            =   1965
      TabIndex        =   24
      Top             =   2535
      Width           =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代  理  人："
      Height          =   180
      Index           =   7
      Left            =   288
      TabIndex        =   23
      Top             =   3252
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申  請  人："
      Height          =   180
      Index           =   6
      Left            =   288
      TabIndex        =   22
      Top             =   2892
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "管制對象："
      Height          =   180
      Index           =   5
      Left            =   288
      TabIndex        =   21
      Top             =   2532
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承  辦  人："
      Height          =   180
      Index           =   4
      Left            =   288
      TabIndex        =   20
      Top             =   2172
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   3
      Left            =   285
      TabIndex        =   19
      Top             =   1815
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業  務  區："
      Height          =   180
      Index           =   2
      Left            =   288
      TabIndex        =   18
      Top             =   1452
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Index           =   1
      Left            =   288
      TabIndex        =   17
      Top             =   732
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   288
      TabIndex        =   16
      Top             =   1092
      Width           =   900
   End
   Begin VB.Label lbe 
      Height          =   252
      Index           =   4
      Left            =   2448
      TabIndex        =   15
      Top             =   1812
      Width           =   1212
   End
   Begin VB.Label lbe 
      Height          =   252
      Index           =   5
      Left            =   2448
      TabIndex        =   14
      Top             =   2172
      Width           =   1092
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2448
      X2              =   2568
      Y1              =   1212
      Y2              =   1212
   End
   Begin VB.Label lbe 
      Height          =   252
      Index           =   3
      Left            =   2448
      TabIndex        =   13
      Top             =   1452
      Width           =   1692
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2448
      X2              =   2568
      Y1              =   3012
      Y2              =   3012
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2448
      X2              =   2568
      Y1              =   3372
      Y2              =   3372
   End
End
Attribute VB_Name = "frm083001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim PLeft(0 To 10) As Integer
Dim m_print As Integer
Dim m_PrintKind As Integer
'Add By Cheng 2002/09/09
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
'Add By Cheng 2002/09/09
Dim strTempName As String
   
   'Add By Cheng 2002/09/09
   blnClkSure = False
   
   If CheckCmd Then Exit Sub
   'Add By Cheng 2002/03/22
   If PUB_CheckKeyInDate(Me.Text(1)) = -1 Then
      Me.Text(1).SetFocus
      Text_GotFocus 1
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text(2)) = -1 Then
      Me.Text(2).SetFocus
      Text_GotFocus 2
      Exit Sub
   End If
   'Add By Cheng 2002/09/09
   If Me.Text(1).Text <> "" And Me.Text(2).Text <> "" Then
      If Val(Me.Text(1).Text) > Val(Me.Text(2).Text) Then
         MsgBox "本所期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.Text(1).SetFocus
         Text_GotFocus 1
         Exit Sub
      End If
   End If
   '檢查業務區
   If Me.Text(3).Text <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objLawDll.GetStaffDeptName(Me.Text(3).Text, strTempName) Then
      If ClsPDGetStaffDeptName(Me.Text(3).Text, strTempName) Then
         Me.lbe(3).Caption = strTempName
      Else
         lbe(3) = ""
         Me.Text(3).SetFocus
         Text_GotFocus 3
         Exit Sub
      End If
   End If
   '檢查智權人員
   If Me.Text(4).Text <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetStaff(Text(4), strTempName) Then
      If ClsPDGetStaffN(Text(4), strTempName) Then
         lbe(4) = strTempName
      Else
         lbe(4) = ""
         Me.Text(4).SetFocus
         Text_GotFocus 4
         Exit Sub
      End If
   End If
   '檢查承辦人員
   If Me.Text(5).Text <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetStaff(Text(5), strTempName) Then
      If ClsPDGetStaffN(Text(5), strTempName) Then
         lbe(5) = strTempName
      Else
         lbe(5) = ""
         Me.Text(5).SetFocus
         Text_GotFocus 5
         Exit Sub
      End If
   End If
   If Me.Text(7).Text <> "" And Me.Text(8).Text <> "" Then
      If Me.Text(7).Text > Me.Text(8).Text Then
         MsgBox "申請人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.Text(7).SetFocus
         Text_GotFocus 7
         Exit Sub
      End If
   End If
   If Me.Text(9).Text <> "" And Me.Text(10).Text <> "" Then
      If Me.Text(9).Text > Me.Text(10).Text Then
         MsgBox "代理人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         blnClkSure = True
         Me.Text(9).SetFocus
         Text_GotFocus 9
         Exit Sub
      End If
   End If
   
   Screen.MousePointer = 11
   m_print = 0
   Select Case Text(6)
      Case 1
         GetPrintLeft 1
         PrintSales
      Case 2
         GetPrintLeft 2
         PrintEnginer
      Case 3
         GetPrintLeft 1
         PrintSales
         GetPrintLeft 2
         PrintEnginer
   End Select
   Screen.MousePointer = 0
   If m_print <> 0 Then
      MsgBox "列印結束!", vbInformation
   Else
      MsgBox "資料庫內無資料 !", vbInformation
   End If
End Sub

Private Sub PrintEnginer()
Dim i As Integer, St As String, Page As Integer, iPrint As Integer
Dim TmpArea As String

On Error GoTo ErrHand
   
   m_PrintKind = 2
   strExc(0) = ""
   If Me.Tag = 0 Then
      strExc(0) = "SELECT DECODE(CP14,S1.ST01,S1.ST02)," & _
         "SUBSTR(CP06,1,4)-1911||'/'||SUBSTR(CP06,5,2)||'/'||SUBSTR(CP06,7,2)," & _
         "DECODE(LENGTH(CP07),NULL,NULL,SUBSTR(CP07,1,4)-1911||'/'||SUBSTR(CP07,5,2)||'/'||SUBSTR(CP07,7,2))," & _
         "SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," & _
         "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04)," & _
         "HC06,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
         "DECODE(CP01||CP10,CPM01||CPM02,CPM03),DECODE(CP13,S2.ST01,S2.ST02)," & _
         "CP64,CP14,CP06,CP01,CP02,CP03,CP04,CP09 FROM " & _
         "CASEPROGRESS,STAFF S1,STAFF S2,HIRECASE,CASEPROPERTYMAP,CUSTOMER WHERE " & _
         "CP27 IS NULL AND CP57 IS NULL AND CP14=S1.ST01(+) AND " & _
         "SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) AND " & _
         "(CP01=CPM01(+) AND CP10=CPM02(+)) AND CP13=S2.ST01(+) AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND HC09 IS NULL" & _
         strGetcdnSQL(0) & " UNION "
   End If
   strExc(0) = strExc(0) & "SELECT DECODE(CP14,S1.ST01,S1.ST02)," & _
      "SUBSTR(CP06,1,4)-1911||'/'||SUBSTR(CP06,5,2)||'/'||SUBSTR(CP06,7,2)," & _
      "DECODE(LENGTH(CP07),NULL,NULL,SUBSTR(CP07,1,4)-1911||'/'||SUBSTR(CP07,5,2)||'/'||SUBSTR(CP07,7,2))," & _
      "SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," & _
      "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04)," & _
      "NVL(LC05, NVL(LC06, lC07)),DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
      "DECODE(CP01||CP10,CPM01||CPM02,CPM03),DECODE(CP13,S2.ST01,S2.ST02)," & _
      "CP64,CP14,CP06,CP01,CP02,CP03,CP04,CP09 FROM " & _
      "CASEPROGRESS,STAFF S1,STAFF S2,LAWCASE,CASEPROPERTYMAP,CUSTOMER WHERE " & _
      "CP27 IS NULL AND CP57 IS NULL AND CP14=S1.ST01(+) AND " & _
      "SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND " & _
      "(CP01=CPM01(+) AND CP10=CPM02(+)) AND CP13=S2.ST01(+) AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and LC08 IS NULL" & _
      strGetcdnSQL(1)
      
   If Text(10) = "" Then
      strExc(0) = strExc(0) & " ORDER BY CP14,CP06,CP01,CP02,CP03,CP04"
     
   Else
      strExc(0) = strExc(0) & " union all select DECODE(CP14,S1.ST01,S1.ST02)," & _
      "SUBSTR(CP06,1,4)-1911||'/'||SUBSTR(CP06,5,2)||'/'||SUBSTR(CP06,7,2)," & _
      "SUBSTR(CP07,1,4)-1911||'/'||SUBSTR(CP07,5,2)||'/'||SUBSTR(CP07,7,2)," & _
      "SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," & _
      "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04)," & _
      "NVL(LC05, NVL(LC06, lC07)),DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
      "DECODE(CP01||CP10,CPM01||CPM02,CPM03),DECODE(CP13,S2.ST01,S2.ST02)," & _
      "CP64,CP14,CP06,CP01,CP02,CP03,CP04,CP09 FROM " & _
      "CASEPROGRESS,STAFF S1,STAFF S2,LAWCASE,CASEPROPERTYMAP,CUSTOMER WHERE " & _
      "CP27 IS NULL AND CP57 IS NULL AND CP14=S1.ST01(+) AND " & _
      "SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND " & _
      "(CP01=CPM01(+) AND CP10=CPM02(+)) AND CP13=S2.ST01(+) AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and LC08 IS NULL" & _
      strGetcdnSQL1(1) & " ORDER BY CP14,CP06,CP01,CP02,CP03,CP04"
   End If
      
   If RsTemp.State = adStateOpen Then RsTemp.Close
   RsTemp.Open strExc(0), cnnConnection
   If RsTemp.EOF And RsTemp.BOF Then
      Exit Sub
   Else
      m_print = 1
   End If
   i = 1
   Page = 1
   If IsNull(RsTemp.Fields(0).Value) = False Then
      TmpArea = RsTemp.Fields(0).Value
   Else
      TmpArea = ""
   End If
   EngineTitle TmpArea, 1
   iPrint = 2700
   With RsTemp
   Do While Not .EOF
'      Printer.CurrentX = 500:      Printer.CurrentY = iPrint
'      Printer.Print Format(.Fields(0), "!@@@@")
      Printer.CurrentX = PLeft(1):    Printer.CurrentY = iPrint
      If .Fields(1) = ChangeTStringToTDateString(GetTaiwanTodayDate) Then
         St = "V" & .Fields(1)
      ElseIf DBDATE(.Fields(1)) < Val(GetTaiwanTodayDate + 19110000) Then
         St = "*" & .Fields(1)
      ElseIf DBDATE(.Fields(1)) > Val(GetTaiwanTodayDate + 19110000) Then
         St = .Fields(1)
      End If
      Printer.Print St
      Printer.CurrentX = PLeft(2):    Printer.CurrentY = iPrint
      Printer.Print IIf(IsNull(.Fields(2)), "", .Fields(2))
      Printer.CurrentX = PLeft(3):    Printer.CurrentY = iPrint
      Printer.Print .Fields(3)
      Printer.CurrentX = PLeft(4):    Printer.CurrentY = iPrint
      Printer.Print .Fields(4)
      Printer.CurrentX = PLeft(5):    Printer.CurrentY = iPrint
      Printer.Print Format(Mid(.Fields(5), 1, 9), "!@@@@@@@@@@")
      Printer.CurrentX = PLeft(6):   Printer.CurrentY = iPrint
      Printer.Print Format(Mid(.Fields(6), 1, 7), "!@@@@@@@@")
      Printer.CurrentX = PLeft(7):   Printer.CurrentY = iPrint
      Printer.Print Format(.Fields(7), "!@@@@@@")
      Printer.CurrentX = PLeft(8):   Printer.CurrentY = iPrint
      Printer.Print Format(.Fields(8), "!@@@@")
      Printer.CurrentX = PLeft(9):   Printer.CurrentY = iPrint
      If IsNull(.Fields(9)) = False Then Printer.Print .Fields(9)
      If IsNull(.Fields(0).Value) = False Then
         TmpArea = .Fields(0).Value
      Else
         TmpArea = ""
      End If
      iPrint = iPrint + 300
      .MoveNext
      If Not .EOF Then
         If IsNull(.Fields(0).Value) = False Then
            St = .Fields(0).Value
         Else
            St = ""
         End If
         If (i Mod 30 = 0) Or (TmpArea <> St) Then
            Printer.NewPage
            Page = Page + 1
            EngineTitle St, Page
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

Private Sub EngineTitle(ByVal Area As String, ByVal Page As String)
 Dim i As Integer
   i = 500

   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000:         Printer.CurrentY = i
   Printer.Print "承辦人期限管制表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 6200:         Printer.CurrentY = i + 500
   Printer.Print "本所期限 : " & ChangeTStringToTDateString(Text(1)) & _
      " - " & ChangeTStringToTDateString(Text(2))
   Printer.Font.Bold = False
   Printer.CurrentX = 500:          Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 13000:        Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
   Printer.CurrentX = 500:          Printer.CurrentY = i + 1100
   Printer.Print "承辦人 : " & Area
   Printer.CurrentX = 13000:        Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
   Printer.CurrentX = 500:          Printer.CurrentY = i + 1400
   Printer.Print String(200, "-")
   Printer.CurrentX = PLeft(1):          Printer.CurrentY = i + 1700
   Printer.Print "本所期限"
   Printer.CurrentX = PLeft(2):          Printer.CurrentY = i + 1700
   Printer.Print "法定期限"
   Printer.CurrentX = PLeft(3):          Printer.CurrentY = i + 1700
   Printer.Print "收文日"
   Printer.CurrentX = PLeft(4):          Printer.CurrentY = i + 1700
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(5):          Printer.CurrentY = i + 1700
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(6):          Printer.CurrentY = i + 1700
   Printer.Print "當事人"
   Printer.CurrentX = PLeft(7):          Printer.CurrentY = i + 1700
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(8):          Printer.CurrentY = i + 1700
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(9):          Printer.CurrentY = i + 1700
   Printer.Print "備註"
   Printer.CurrentX = 500:          Printer.CurrentY = i + 2000
   Printer.Print String(200, "-")
End Sub

Private Sub GetPrintLeft(ByVal Sty As Integer)
   Erase PLeft
   If Sty = 1 Then
'500,1900,3500,5100,6600,9000,10400,12000,13300
      PLeft(0) = 500:     PLeft(1) = 1900
      PLeft(2) = 3200:    PLeft(3) = 4500
      PLeft(4) = 6000:    PLeft(5) = 8300
      PLeft(6) = 10500:    PLeft(7) = 11800
      PLeft(8) = 13000
   ElseIf Sty = 2 Then
'1900,3500,5000,6400,8000,10400,12000,13300,14600
      PLeft(1) = 700
      PLeft(2) = 2000:    PLeft(3) = 3500
      PLeft(4) = 4800:    PLeft(5) = 6100
      PLeft(6) = 8500:    PLeft(7) = 10700
      PLeft(8) = 12300:   PLeft(9) = 13600
   End If
End Sub

Private Sub PrintSales()
 Dim i As Integer, St As String, Page As Integer, iPrint As Integer
 Dim TmpArea As String
On Error GoTo ErrHand
   m_PrintKind = 1
   strExc(0) = ""
   If Me.Tag = 0 Then
      'Added by Lydia 2023/12/26
      If strSrvDate(1) >= 新部門啟用日 Then
         strExc(0) = strExc(0) & "SELECT DECODE(NP10,S1.ST01,S1.ST02)," & _
            "SUBSTR(NP08,1,4)-1911||'/'||SUBSTR(NP08,5,2)||'/'||SUBSTR(NP08,7,2)," & _
            "SUBSTR(NP09,1,4)-1911||'/'||SUBSTR(NP09,5,2)||'/'||SUBSTR(NP09,7,2)," & _
            "NP02||'-'||NP03||DECODE(NP04,'0','','-'||NP04)||DECODE(NP05,'00','','-'||NP05)," & _
            "HC06,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
            "DECODE(NP02||NP07,CPM01||CPM02,CPM03),DECODE(CP14,S2.ST01,S2.ST02)," & _
            "NP15,NVL(A0922,A0902) AS A0902,NP01,CP27,S3.ST03,NP10,NP08,NP02,NP03,NP04,NP05,NP22 " & _
            "FROM STAFF S1,STAFF S2,STAFF S3,CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,CUSTOMER,NEXTPROGRESS,ACC090,ACC090NEW WHERE " & _
            "NP10=S1.ST01(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) AND " & _
            "(NP02=CPM01(+) AND NP07=CPM02(+)) AND NP06 IS NULL AND S3.ST03=A0901(+) AND NP10=S3.ST01(+) AND S3.ST93=A0921(+) AND " & _
            "CP14=S2.ST01(+) AND NP01=CP09 AND HC01=NP02 AND HC02=NP03 AND HC03=NP04 AND HC04=NP05 AND HC09 IS NULL " & strGetcdnSQL(0) & _
            " UNION "
      Else
      'end 2023/12/26
         strExc(0) = strExc(0) & "SELECT DECODE(NP10,S1.ST01,S1.ST02)," & _
            "SUBSTR(NP08,1,4)-1911||'/'||SUBSTR(NP08,5,2)||'/'||SUBSTR(NP08,7,2)," & _
            "SUBSTR(NP09,1,4)-1911||'/'||SUBSTR(NP09,5,2)||'/'||SUBSTR(NP09,7,2)," & _
            "NP02||'-'||NP03||DECODE(NP04,'0','','-'||NP04)||DECODE(NP05,'00','','-'||NP05)," & _
            "HC06,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
            "DECODE(NP02||NP07,CPM01||CPM02,CPM03),DECODE(CP14,S2.ST01,S2.ST02)," & _
            "NP15,A0902,NP01,CP27,S3.ST03,NP10,NP08,NP02,NP03,NP04,NP05,NP22 " & _
            "FROM STAFF S1,STAFF S2,STAFF S3,CASEPROGRESS,HIRECASE,CASEPROPERTYMAP,CUSTOMER,NEXTPROGRESS,ACC090 WHERE " & _
            "NP10=S1.ST01(+) AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) AND " & _
            "(NP02=CPM01(+) AND NP07=CPM02(+)) AND NP06 IS NULL AND S3.ST03=A0901(+) AND NP10=S3.ST01(+) AND " & _
            "CP14=S2.ST01(+) AND NP01=CP09 AND HC01=NP02 AND HC02=NP03 AND HC03=NP04 AND HC04=NP05 AND HC09 IS NULL " & strGetcdnSQL(0) & _
            " UNION "
      End If
   End If
   
   'Added by Lydia 2023/12/26
   If strSrvDate(1) >= 新部門啟用日 Then
      strExc(0) = strExc(0) & "SELECT DECODE(NP10,S1.ST01,S1.ST02)," & _
         "SUBSTR(NP08,1,4)-1911||'/'||SUBSTR(NP08,5,2)||'/'||SUBSTR(NP08,7,2)," & _
         "SUBSTR(NP09,1,4)-1911||'/'||SUBSTR(NP09,5,2)||'/'||SUBSTR(NP09,7,2)," & _
         "NP02||'-'||NP03||DECODE(NP04,'0','','-'||NP04)||DECODE(NP05,'00','','-'||NP05)," & _
         "NVL(LC05, NVL(LC06, lC07)),DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
         "DECODE(NP02||NP07,CPM01||CPM02,CPM03),DECODE(CP14,S2.ST01,S2.ST02),NP15,NVL(A0922,A0902) AS A0902,NP01,CP27,S3.ST03,NP10,NP08,NP02,NP03,NP04,NP05,NP22 " & _
         "FROM STAFF S1,STAFF S2,STAFF S3,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,CUSTOMER,NEXTPROGRESS,ACC090,ACC090NEW WHERE " & _
         "NP10=S1.ST01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND " & _
         "(NP02=CPM01(+) AND NP07=CPM02(+)) AND NP06 IS NULL AND S3.ST03=A0901(+) AND NP10=S3.ST01(+) AND S3.ST93=A0921(+) AND " & _
         "CP14=S2.ST01(+) AND NP01=CP09 AND LC01=NP02 AND LC02=NP03 AND LC03=NP04 AND LC04=NP05 AND LC08 IS NULL " & strGetcdnSQL(1)
   Else
   'end 2023/12/26
      strExc(0) = strExc(0) & "SELECT DECODE(NP10,S1.ST01,S1.ST02)," & _
         "SUBSTR(NP08,1,4)-1911||'/'||SUBSTR(NP08,5,2)||'/'||SUBSTR(NP08,7,2)," & _
         "SUBSTR(NP09,1,4)-1911||'/'||SUBSTR(NP09,5,2)||'/'||SUBSTR(NP09,7,2)," & _
         "NP02||'-'||NP03||DECODE(NP04,'0','','-'||NP04)||DECODE(NP05,'00','','-'||NP05)," & _
         "NVL(LC05, NVL(LC06, lC07)),DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
         "DECODE(NP02||NP07,CPM01||CPM02,CPM03),DECODE(CP14,S2.ST01,S2.ST02),NP15,A0902,NP01,CP27,S3.ST03,NP10,NP08,NP02,NP03,NP04,NP05,NP22 " & _
         "FROM STAFF S1,STAFF S2,STAFF S3,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,CUSTOMER,NEXTPROGRESS,ACC090 WHERE " & _
         "NP10=S1.ST01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND " & _
         "(NP02=CPM01(+) AND NP07=CPM02(+)) AND NP06 IS NULL AND S3.ST03=A0901(+) AND NP10=S3.ST01(+) AND " & _
         "CP14=S2.ST01(+) AND NP01=CP09 AND LC01=NP02 AND LC02=NP03 AND LC03=NP04 AND LC04=NP05 AND LC08 IS NULL " & strGetcdnSQL(1)
   End If
      
   If Text(10) = "" Then
      strExc(0) = strExc(0) & " ORDER BY ST03,NP10,NP08,NP02,NP03,NP04,NP05"
   Else
      'Added by Lydia 2023/12/26
      If strSrvDate(1) >= 新部門啟用日 Then
         strExc(0) = strExc(0) & " union all select DECODE(NP10,S1.ST01,S1.ST02)," & _
         "SUBSTR(NP08,1,4)-1911||'/'||SUBSTR(NP08,5,2)||'/'||SUBSTR(NP08,7,2)," & _
         "SUBSTR(NP09,1,4)-1911||'/'||SUBSTR(NP09,5,2)||'/'||SUBSTR(NP09,7,2)," & _
         "NP02||'-'||NP03||DECODE(NP04,'0','','-'||NP04)||DECODE(NP05,'00','','-'||NP05)," & _
         "NVL(LC05, NVL(LC06, lC07)),DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
         "DECODE(NP02||NP07,CPM01||CPM02,CPM03),DECODE(CP14,S2.ST01,S2.ST02),NP15,NVL(A0922,A0902) AS A0902,NP01,CP27,S3.ST03,NP10,NP08,NP02,NP03,NP04,NP05,NP22 " & _
         "FROM STAFF S1,STAFF S2,STAFF S3,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,CUSTOMER,NEXTPROGRESS,ACC090,ACC090NEW WHERE " & _
         "NP10=S1.ST01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND " & _
         "(NP02=CPM01(+) AND NP07=CPM02(+)) AND NP06 IS NULL AND S3.ST03=A0901(+) AND NP10=S3.ST01(+) AND S3.ST93=A0921(+) AND " & _
         "CP14=S2.ST01(+) AND NP01=CP09 AND LC01=NP02 AND LC02=NP03 AND LC03=NP04 AND LC04=NP05 AND LC08 IS NULL " & strGetcdnSQL1(1) & _
         " ORDER BY ST03,NP10,NP08,NP02,NP03,NP04,NP05"
      Else
      'end 2023/12/26
         strExc(0) = strExc(0) & " union all select DECODE(NP10,S1.ST01,S1.ST02)," & _
         "SUBSTR(NP08,1,4)-1911||'/'||SUBSTR(NP08,5,2)||'/'||SUBSTR(NP08,7,2)," & _
         "SUBSTR(NP09,1,4)-1911||'/'||SUBSTR(NP09,5,2)||'/'||SUBSTR(NP09,7,2)," & _
         "NP02||'-'||NP03||DECODE(NP04,'0','','-'||NP04)||DECODE(NP05,'00','','-'||NP05)," & _
         "NVL(LC05, NVL(LC06, lC07)),DECODE(LC11,CU01||CU02,NVL(CU04,NVL(CU05,CU06)))," & _
         "DECODE(NP02||NP07,CPM01||CPM02,CPM03),DECODE(CP14,S2.ST01,S2.ST02),NP15,A0902,NP01,CP27,S3.ST03,NP10,NP08,NP02,NP03,NP04,NP05,NP22 " & _
         "FROM STAFF S1,STAFF S2,STAFF S3,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,CUSTOMER,NEXTPROGRESS,ACC090 WHERE " & _
         "NP10=S1.ST01(+) AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND " & _
         "(NP02=CPM01(+) AND NP07=CPM02(+)) AND NP06 IS NULL AND S3.ST03=A0901(+) AND NP10=S3.ST01(+) AND " & _
         "CP14=S2.ST01(+) AND NP01=CP09 AND LC01=NP02 AND LC02=NP03 AND LC03=NP04 AND LC04=NP05 AND LC08 IS NULL " & strGetcdnSQL1(1) & _
         " ORDER BY ST03,NP10,NP08,NP02,NP03,NP04,NP05"
      End If
   End If
   
   If RsTemp.State = adStateOpen Then RsTemp.Close
   RsTemp.Open strExc(0), cnnConnection
   If RsTemp.EOF And RsTemp.BOF Then
      Exit Sub
   Else
      m_print = 1
   End If
   i = 1
   Page = 1
 '  Dialog1.ShowPrinter
   If IsNull(RsTemp.Fields(9).Value) = False Then
      TmpArea = RsTemp.Fields(9).Value
   Else
      TmpArea = ""
   End If
   SalesTitle TmpArea, 1
   iPrint = 2700
   With RsTemp
   Do While Not .EOF
      Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
      Printer.Print Format(.Fields(0), "!@@@@")
      Printer.CurrentX = PLeft(1):     Printer.CurrentY = iPrint
      If DBDATE(.Fields(1)) > Val(GetTaiwanTodayDate + 19110000) Then
         St = .Fields(1)
      ElseIf .Fields(1) = ChangeTStringToTDateString(GetTaiwanTodayDate) Then
         St = "V" & .Fields(1)
      ElseIf DBDATE(.Fields(1)) < Val(GetTaiwanTodayDate + 19110000) Then
         St = "*" & .Fields(1)
      End If
      If Left(.Fields(10), 1) = "C" Then
         If IsNull(.Fields(11)) = True Or .Fields(11) = "" Then
            St = Replace(St, "V", "")
            St = Replace(St, "*", "")
            St = "#" & St
         End If
      End If
      Printer.Print St
      Printer.CurrentX = PLeft(2):     Printer.CurrentY = iPrint
      Printer.Print .Fields(2)
      Printer.CurrentX = PLeft(3):     Printer.CurrentY = iPrint
      Printer.Print .Fields(3)
      Printer.CurrentX = PLeft(4):     Printer.CurrentY = iPrint
      Printer.Print Format(Mid(.Fields(4), 1, 8), "!@@@@@@@@")
      Printer.CurrentX = PLeft(5):     Printer.CurrentY = iPrint
      Printer.Print Format(Mid(.Fields(5), 1, 8), "!@@@@@@@@")
      Printer.CurrentX = PLeft(6):     Printer.CurrentY = iPrint
      Printer.Print Format(.Fields(6), "!@@@@@@")
      Printer.CurrentX = PLeft(7):    Printer.CurrentY = iPrint
      Printer.Print Format(.Fields(7), "!@@@@")
      Printer.CurrentX = PLeft(8):    Printer.CurrentY = iPrint
      If IsNull(.Fields(8)) = False Then Printer.Print .Fields(8)
      If IsNull(.Fields(9).Value) = False Then
         TmpArea = .Fields(9).Value
      Else
         TmpArea = ""
      End If
      iPrint = iPrint + 300
      .MoveNext
      If Not .EOF Then
         If IsNull(.Fields(9).Value) = False Then
            St = .Fields(9).Value
         Else
            St = ""
         End If
         If (i Mod 27 = 0) Or (TmpArea <> St) Then
            Printer.NewPage
            Page = Page + 1
            SalesTitle St, Page
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

Private Sub SalesTitle(ByVal Area As String, ByVal Page As String)
 Dim i As Integer
   i = 500
   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000:         Printer.CurrentY = i
   Printer.Print "智權人員期限管制表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 6200:         Printer.CurrentY = i + 500
   Printer.Print "本所期限 : " & ChangeTStringToTDateString(Text(1)) & _
      " - " & ChangeTStringToTDateString(Text(2))
   Printer.Font.Bold = False
   Printer.CurrentX = 500:          Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 13000:        Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
   Printer.CurrentX = 500:          Printer.CurrentY = i + 1100
   Printer.Print "業務區 : " & Area
   Printer.CurrentX = 13000:        Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
   Printer.CurrentX = 500:          Printer.CurrentY = i + 1400
   Printer.Print String(200, "-")
   Printer.CurrentX = PLeft(0):          Printer.CurrentY = i + 1700
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(1):          Printer.CurrentY = i + 1700
   Printer.Print "本所期限"
   Printer.CurrentX = PLeft(2):          Printer.CurrentY = i + 1700
   Printer.Print "法定期限"
   Printer.CurrentX = PLeft(3):          Printer.CurrentY = i + 1700
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(4):          Printer.CurrentY = i + 1700
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(5):          Printer.CurrentY = i + 1700
   Printer.Print "當事人"
   Printer.CurrentX = PLeft(6):          Printer.CurrentY = i + 1700
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(7):          Printer.CurrentY = i + 1700
   Printer.Print "承辦人"
   Printer.CurrentX = PLeft(8):          Printer.CurrentY = i + 1700
   Printer.Print "備註"
   Printer.CurrentX = 500:          Printer.CurrentY = i + 2000
   Printer.Print String(200, "-")
End Sub

Private Sub Form_Activate()
  Text(1).SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Text(0).Text = GetSystemKindByNick
   m_print = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm083001 = Nothing
End Sub

Private Sub Text_Change(Index As Integer)
   Text(Index).BackColor = &H80000005
   Select Case Index
      Case 3, 4, 5
         If Text(Index) = "" Then lbe(Index) = ""
   End Select
End Sub

Private Sub Text_GotFocus(Index As Integer)
   TextInverse Text(Index)
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 6
         If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 0, 3, 4, 5, 7, 8, 9, 10
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
   Select Case Index
      Case 2
            'Add/Modify By Cheng 2002/09/09
            If blnClkSure = False Then
               If RunNick(Text(Index - 1), Text(Index)) Then
                  Text(Index - 1).SetFocus
               End If
            Else
               blnClkSure = False
            End If
      Case 8
            If Len(Text(Index - 1)) <> 0 Then
               If Left(Text(Index - 1), 6) <> Left(Text(Index), 6) Then
                   MsgBox "申請人前 6 碼必須相同", , "USER 輸入錯誤"
                   Text(Index - 1).SetFocus
                   Exit Sub
               End If
            End If
            'Add/Modify By Cheng 2002/09/09
            If blnClkSure = False Then
               If RunNick(Text(Index - 1), Text(Index)) Then
                  Text(Index - 1).SetFocus
               End If
            Else
               blnClkSure = False
            End If
      Case 10
            If Len(Text(Index - 1)) <> 0 Then
               If Left(Text(Index - 1), 6) <> Left(Text(Index), 6) Then
                   MsgBox "代理人前 6 碼必須相同", , "USER 輸入錯誤"
                   Text(Index - 1).SetFocus
                   Exit Sub
               End If
            End If
            'Add/Modify By Cheng 2002/09/09
            If blnClkSure = False Then
               If RunNick(Text(Index - 1), Text(Index)) Then
                  Text(Index - 1).SetFocus
               End If
            Else
               blnClkSure = False
            End If
   End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim strTempName As String, i As Integer, t As Integer
Dim strTit As String
Dim strMsg As String
Dim strTemp1 As Variant
Dim strTemp2 As Variant
Dim j As Integer
Dim s As Integer
 
   If Text(Index) = "" Then Exit Sub
   Select Case Index
      Case 0
         'If ChkSysName(Text(Index)) = False Then Cancel = True
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(Text(Index)), ",,", ""), ",")
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
      Case 2
         If CheckIsTaiwanDate(Text(Index)) Then
         Else
            Cancel = True
         End If
      Case 3
         'edit by nickc 2007/02/07 不用 dll 了
         'If objLawDll.GetStaffDeptName(Text(Index), strTempName) Then
         If ClsPDGetStaffDeptName(Text(Index), strTempName) Then
            lbe(Index) = strTempName
         Else
            Cancel = True
            lbe(Index) = ""
         End If
      Case 4, 5
         Text(Index) = UCase(Text(Index))
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(Text(Index), strTempName) Then
         If ClsPDGetStaffN(Text(Index), strTempName) Then
            lbe(Index) = strTempName
         Else
            lbe(Index) = ""
            Cancel = True
         End If
      Case 6
         If Text(Index) <> "1" And Text(Index) <> "2" And Text(Index) <> "3" Then
            Beep
            Cancel = True
         End If
      Case 7, 8
'         If objPublicData.GetCustomer(Text(Index), strTempName) = False Then
'            Cancel = True
'         Else
'            If Index = 8 Then
'               If Left(Text(7), 6) <> Left(Text(8), 6) Then
'                  MsgBox "申請人起迄號前六碼必須相同 !", vbCritical
'                  Cancel = True
'               End If
'            End If
'         End If
      Case 9, 10
'         If objPublicData.GetAgent(Text(Index), strTempName) = False Then
'            Cancel = True
'         Else
'            If Index = 10 Then
'               If Left(Text(9), 6) <> Left(Text(10), 6) Then
'                  MsgBox "代理人起迄號前六碼必須相同 !", vbCritical
'                  Cancel = True
'               End If
'            End If
'         End If
      End Select
      If Cancel Then TextInverse Text(Index)
End Sub

Private Function CheckCmd() As Boolean
Dim i As Integer, blnIsEnbled As Boolean
   
   For i = 0 To 6
      Select Case i
         Case 0, 1, 2, 6
            If Text(i) = "" Then
               blnIsEnbled = False
               Text(i).BackColor = &HFFC0C0
               Text(i).SetFocus
               MsgBox "請輸入資料 !", vbCritical
               CheckCmd = True
               Exit Function
            Else
               CheckCmd = False
            End If
      End Select
   Next
End Function

Private Function strGetcdnSQL(ByVal Situ As Integer) As String
Dim i As Integer
Dim strNP02 As String
Dim strTemp As Variant
 
   If Me.Tag = 0 Then
      strTemp = Split(UCase(Text(0).Text), ",")
      For i = 0 To UBound(strTemp)
          If strTemp(i) = "L" Or strTemp(i) = "LA" Then
             strNP02 = strNP02 & strTemp(i)
             strNP02 = strNP02 & "','"
          End If
      Next i
   ElseIf Me.Tag = 1 Then
      strTemp = Split(UCase(Text(0).Text), ",")
      For i = 0 To UBound(strTemp)
          'Modify By Sindy 2009/07/24 增加LIN系統類別
          If strTemp(i) = "CFL" Or strTemp(i) = "FCL" Or strTemp(i) = "LIN" Then
             strNP02 = strNP02 & strTemp(i)
             strNP02 = strNP02 & "','"
          End If
      Next i
   'add by sonia 2019/8/6 +ACS系統類別
   ElseIf Me.Tag = 2 Then
      strTemp = Split(UCase(Text(0).Text), ",")
      For i = 0 To UBound(strTemp)
          If strTemp(i) = "ACS" Then
             strNP02 = strNP02 & strTemp(i)
             strNP02 = strNP02 & "','"
          End If
      Next i
   'end 2019/8/6
   End If

   If m_PrintKind = 1 Then
      strExc(1) = " AND NP02 in('" & strNP02 & "') AND (NP08 BETWEEN '" + _
         ChangeTStringToWString(Text(1)) + "' AND '" + ChangeTStringToWString(Text(2)) + "')"
      If Text(3).Text <> "" Then strExc(1) = " and CP12='" + Text(3).Text + "'"
      If Text(4).Text <> "" Then strExc(1) = strExc(1) + " and NP10='" + Text(4).Text + "'"
      If Text(5) <> "" Then strExc(1) = strExc(1) + " and CP14='" + Text(5) + "'"
      If Situ = 0 Then
         If Text(7).Text = "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and HC05<='" + Text(8) + String(9 - Len(Text(8)), 0) + "'"
         ElseIf Text(7).Text <> "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and (HC05 BETWEEN '" + Text(7) + String(9 - Len(Text(7)), "0") + _
               "' AND '" + Text(8) + String(9 - Len(Text(8)), "0") + "')"
         End If
      Else
         If Text(7).Text = "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and LC11<='" + Text(8) + String(9 - Len(Text(8)), 0) + "'"
         ElseIf Text(7).Text <> "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and (LC11 BETWEEN '" + Text(7) + String(9 - Len(Text(7)), "0") + _
               "' AND '" + Text(8) + String(9 - Len(Text(8)), "0") + "')"
         End If
      End If
      If Text(9) = "" And Text(10).Text <> "" Then
         strExc(1) = strExc(1) + " and CP44<='" + Text(10) + String(9 - Len(Text(10)), 0) + "'"
      ElseIf Text(9) <> "" And Text(10).Text <> "" Then
         strExc(1) = strExc(1) + " and (CP44 BETWEEN '" + Text(9) + String(9 - Len(Text(9)), "0") + _
            "' AND '" + Text(10) + String(9 - Len(Text(10)), "0") + "')"
      End If
   ElseIf m_PrintKind = 2 Then
      strExc(1) = " AND CP01 in('" & strNP02 & "') AND (CP06 BETWEEN '" + _
         ChangeTStringToWString(Text(1)) + "' AND '" + ChangeTStringToWString(Text(2)) + "')"
      If Text(3).Text <> "" Then strExc(1) = " and CP12='" + Text(3).Text + "'"
      If Text(4).Text <> "" Then strExc(1) = strExc(1) + " and CP13='" + Text(4).Text + "'"
      If Text(5) <> "" Then strExc(1) = strExc(1) + " and CP14='" + Text(5) + "'"
      If Situ = 0 Then
         If Text(7).Text = "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and HC05<='" + Text(8) + String(9 - Len(Text(8)), 0) + "'"
         ElseIf Text(7).Text <> "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and (HC05 BETWEEN '" + Text(7) + String(9 - Len(Text(7)), "0") + _
               "' AND '" + Text(8) + String(9 - Len(Text(8)), "0") + "')"
         End If
      Else
         If Text(7).Text = "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and LC11<='" + Text(8) + String(9 - Len(Text(8)), 0) + "'"
         ElseIf Text(7).Text <> "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and (LC11 BETWEEN '" + Text(7) + String(9 - Len(Text(7)), "0") + _
               "' AND '" + Text(8) + String(9 - Len(Text(8)), "0") + "')"
         End If
      End If
      If Text(9) = "" And Text(10).Text <> "" Then
         strExc(1) = strExc(1) + " and CP44<='" + Text(10) + String(9 - Len(Text(10)), 0) + "'"
      ElseIf Text(9) <> "" And Text(10).Text <> "" Then
         strExc(1) = strExc(1) + " and (CP44 BETWEEN '" + Text(9) + String(9 - Len(Text(9)), "0") + "' AND '" + Text(10) + String(9 - Len(Text(10)), "0") + "')"
      End If
   End If
   strGetcdnSQL = strExc(1)
End Function

Private Function strGetcdnSQL1(ByVal Situ As Integer) As String
Dim i As Integer
Dim strNP02 As String
Dim strTemp As Variant
 
   If Me.Tag = 0 Then
      strTemp = Split(UCase(Text(0).Text), ",")
      For i = 0 To UBound(strTemp)
          If strTemp(i) = "L" Or strTemp(i) = "LA" Then
             strNP02 = strNP02 & strTemp(i)
             strNP02 = strNP02 & "','"
          End If
      Next i
   ElseIf Me.Tag = 1 Then
      strTemp = Split(UCase(Text(0).Text), ",")
      For i = 0 To UBound(strTemp)
          'Modify By Sindy 2009/07/24 增加LIN系統類別
          'add by sonia 2019/7/31 +ACS系統類別
          If strTemp(i) = "CFL" Or strTemp(i) = "FCL" Or strTemp(i) = "LIN" Then
             strNP02 = strNP02 & strTemp(i)
             strNP02 = strNP02 & "','"
          End If
      Next i
   'add by sonia 2019/8/6 +ACS系統類別
   ElseIf Me.Tag = 2 Then
      strTemp = Split(UCase(Text(0).Text), ",")
      For i = 0 To UBound(strTemp)
          If strTemp(i) = "ACS" Then
             strNP02 = strNP02 & strTemp(i)
             strNP02 = strNP02 & "','"
          End If
      Next i
   'end 2019/8/6
   End If

   If m_PrintKind = 1 Then
      strExc(1) = " AND NP02 in('" & strNP02 & "') AND (NP08 BETWEEN '" + _
         ChangeTStringToWString(Text(1)) + "' AND '" + ChangeTStringToWString(Text(2)) + "')"
      If Text(3).Text <> "" Then strExc(1) = " and CP12='" + Text(3).Text + "'"
      If Text(4).Text <> "" Then strExc(1) = strExc(1) + " and NP10='" + Text(4).Text + "'"
      If Text(5) <> "" Then strExc(1) = strExc(1) + " and CP14='" + Text(5) + "'"
      If Situ = 0 Then
         If Text(7).Text = "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and HC05<='" + Text(8) + String(9 - Len(Text(8)), 0) + "'"
         ElseIf Text(7).Text <> "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and (HC05 BETWEEN '" + Text(7) + String(9 - Len(Text(7)), "0") + _
               "' AND '" + Text(8) + String(9 - Len(Text(8)), "0") + "')"
         End If
      Else
         If Text(7).Text = "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and LC11<='" + Text(8) + String(9 - Len(Text(8)), 0) + "'"
         ElseIf Text(7).Text <> "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and (LC11 BETWEEN '" + Text(7) + String(9 - Len(Text(7)), "0") + _
               "' AND '" + Text(8) + String(9 - Len(Text(8)), "0") + "')"
         End If
      End If
      If Text(9) = "" And Text(10).Text <> "" Then
         strExc(1) = strExc(1) + " and LC22<='" + Text(10) + String(9 - Len(Text(10)), 0) + "'"
      ElseIf Text(9) <> "" And Text(10).Text <> "" Then
         strExc(1) = strExc(1) + " and (LC22 BETWEEN '" + Text(9) + String(9 - Len(Text(9)), "0") + _
            "' AND '" + Text(10) + String(9 - Len(Text(10)), "0") + "')"
      End If
   ElseIf m_PrintKind = 2 Then
      strExc(1) = " AND CP01 in('" & strNP02 & "') AND (CP06 BETWEEN '" + _
         ChangeTStringToWString(Text(1)) + "' AND '" + ChangeTStringToWString(Text(2)) + "')"
      If Text(3).Text <> "" Then strExc(1) = " and CP12='" + Text(3).Text + "'"
      If Text(4).Text <> "" Then strExc(1) = strExc(1) + " and CP13='" + Text(4).Text + "'"
      If Text(5) <> "" Then strExc(1) = strExc(1) + " and CP14='" + Text(5) + "'"
      If Situ = 0 Then
         If Text(7).Text = "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and HC05<='" + Text(8) + String(9 - Len(Text(8)), 0) + "'"
         ElseIf Text(7).Text <> "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and (HC05 BETWEEN '" + Text(7) + String(9 - Len(Text(7)), "0") + _
               "' AND '" + Text(8) + String(9 - Len(Text(8)), "0") + "')"
         End If
      Else
         If Text(7).Text = "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and LC11<='" + Text(8) + String(9 - Len(Text(8)), 0) + "'"
         ElseIf Text(7).Text <> "" And Text(8).Text <> "" Then
            strExc(1) = strExc(1) + " and (LC11 BETWEEN '" + Text(7) + String(9 - Len(Text(7)), "0") + _
               "' AND '" + Text(8) + String(9 - Len(Text(8)), "0") + "')"
         End If
      End If
      If Text(9) = "" And Text(10).Text <> "" Then
         strExc(1) = strExc(1) + " and LC22<='" + Text(10) + String(9 - Len(Text(10)), 0) + "'"
      ElseIf Text(9) <> "" And Text(10).Text <> "" Then
         strExc(1) = strExc(1) + " and (LC22 BETWEEN '" + Text(9) + String(9 - Len(Text(9)), "0") + "' AND '" + Text(10) + String(9 - Len(Text(10)), "0") + "')"
      End If
   End If
   strGetcdnSQL1 = strExc(1)
End Function
