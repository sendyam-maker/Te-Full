VERSION 5.00
Begin VB.Form frm12040139 
   BorderStyle     =   1  '單線固定
   Caption         =   "閉卷清單"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5910
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   2850
      MaxLength       =   7
      TabIndex        =   2
      Top             =   930
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   1350
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1260
      Width           =   405
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1350
      MaxLength       =   7
      TabIndex        =   1
      Top             =   930
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1350
      TabIndex        =   0
      Top             =   600
      Width           =   3045
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4020
      TabIndex        =   4
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4845
      TabIndex        =   5
      Top             =   60
      Width           =   800
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "－"
      Height          =   180
      Index           =   3
      Left            =   2520
      TabIndex        =   9
      Top             =   960
      Width           =   180
   End
   Begin VB.Label lbl 
      Caption         =   "排序條件 :                 (1.本所案號 2.閉卷日期 3.閉卷原因 4.智權人員)"
      Height          =   255
      Index           =   2
      Left            =   390
      TabIndex        =   8
      Top             =   1290
      Width           =   5385
   End
   Begin VB.Label lbl 
      Caption         =   "閉卷日期 : "
      Height          =   255
      Index           =   1
      Left            =   390
      TabIndex        =   7
      Top             =   960
      Width           =   945
   End
   Begin VB.Label lbl 
      Caption         =   "系統類別 :                                                                           (ALL：全部)"
      Height          =   255
      Index           =   0
      Left            =   390
      TabIndex        =   6
      Top             =   630
      Width           =   5355
   End
End
Attribute VB_Name = "frm12040139"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim PLeft(0 To 7) As Integer

Private Sub cmdok_Click(Index As Integer)
Dim ii As Integer

Select Case Index
Case 0 '確定
   For ii = 0 To Me.Text1.Count - 1
      If Me.Text1(ii).Enabled = True Then
         If CheckKeyIn(ii) = False Then
            Me.Text1(ii).SetFocus
            Text1_GotFocus ii
            Exit Sub
         End If
         'Add By Cheng 2002/06/10
         If ii = 2 Then
            If Val(Me.Text1(1).Text) > Val(Me.Text1(2).Text) Then
               MsgBox "閉卷日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.Text1(1).SetFocus
               TextInverse Me.Text1(1)
               Exit Sub
            End If
         End If
      End If
   Next ii
   Process
Case 1 '結束
   Unload Me
End Select
End Sub

Private Sub Process()
Dim strSql As String
Dim strSQL1 As String
Dim strSQL2 As String
Dim StrSQL3 As String
Dim StrSQL4 As String
Dim strSQL5 As String
Dim strSQL_1 As String
Dim rs As New ADODB.Recordset

Screen.MousePointer = vbHourglass
strSql = "": strSQL1 = "": strSQL2 = "": StrSQL3 = "": StrSQL4 = "": strSQL5 = "": strSQL_1 = ""

strSQL1 = strSQL1 & " AND PA01 IN (" & SQLGrpStr(IIf(Me.Text1(0).Text <> "ALL", Me.Text1(0).Text, GetAllSysKind(Me.Text1(0))), 1) & ") "
strSQL2 = strSQL2 & " AND TM01 IN (" & SQLGrpStr(IIf(Me.Text1(0).Text <> "ALL", Me.Text1(0).Text, GetAllSysKind(Me.Text1(0))), 2) & ") "
StrSQL3 = StrSQL3 & " AND LC01 IN (" & SQLGrpStr(IIf(Me.Text1(0).Text <> "ALL", Me.Text1(0).Text, GetAllSysKind(Me.Text1(0))), 3) & ") "
StrSQL4 = StrSQL4 & " AND HC01 IN (" & SQLGrpStr(IIf(Me.Text1(0).Text <> "ALL", Me.Text1(0).Text, GetAllSysKind(Me.Text1(0))), 4) & ") "
strSQL5 = strSQL5 & " AND SP01 IN (" & SQLGrpStr(IIf(Me.Text1(0).Text <> "ALL", Me.Text1(0).Text, GetAllSysKind(Me.Text1(0))), 5) & ") "

strSQL1 = strSQL1 & " AND PA58>=" & Val(Me.Text1(1).Text) + 19110000 & " AND PA58<=" & Val(Me.Text1(2).Text) + 19110000 & " "
strSQL2 = strSQL2 & " AND TM30>=" & Val(Me.Text1(1).Text) + 19110000 & " AND TM30<=" & Val(Me.Text1(2).Text) + 19110000 & " "
StrSQL3 = StrSQL3 & " AND LC09>=" & Val(Me.Text1(1).Text) + 19110000 & " AND LC09<=" & Val(Me.Text1(2).Text) + 19110000 & " "
StrSQL4 = StrSQL4 & " AND HC10>=" & Val(Me.Text1(1).Text) + 19110000 & " AND HC10<=" & Val(Me.Text1(2).Text) + 19110000 & " "
strSQL5 = strSQL5 & " AND SP16>=" & Val(Me.Text1(1).Text) + 19110000 & " AND SP16<=" & Val(Me.Text1(2).Text) + 19110000 & " "
'專利
strSQL_1 = "SELECT CP01,CP02,CP03,CP04,MAX(CP05) CP05,CP10 FROM CASEPROGRESS,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP10='913' " & strSQL1 & " GROUP BY CP01,CP02,CP03,CP04,CP10"
strSql = strSql & " SELECT PA01||'-'||PA02||'-'||PA03||'-'||PA04, NVL(PA05,NVL(PA06,PA07)), PA58, NVL(ROR02,' '), NVL(ST02,' '), NVL(C1.CP13,' '), NVL(C1.CP12,' '), PA01, NVL(ROR01,' ') " & _
                  " FROM PATENT, ReasonOfRelief, CaseProgress C1, ( " & strSQL_1 & " ) C2, Staff WHERE C2.CP01=C1.CP01 AND C2.CP02=C1.CP02 AND C2.CP03=C1.CP03 AND C2.CP04=C1.CP04 AND C2.CP05=C1.CP05(+) AND C2.CP10=C1.CP10 AND C1.CP01=PA01(+) AND C1.CP02=PA02(+) AND C1.CP03=PA03(+) AND C1.CP04=PA04(+) AND PA59=ROR01(+) AND C1.CP13=ST01(+) "
'商標
strSQL_1 = "SELECT CP01,CP02,CP03,CP04,MAX(CP05) CP05,CP10 FROM CASEPROGRESS,TRADEMARK WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP10='704' " & strSQL2 & " GROUP BY CP01,CP02,CP03,CP04,CP10"
strSql = strSql & " UNION SELECT TM01||'-'||TM02||'-'||TM03||'-'||TM04, NVL(TM05,NVL(TM06,TM07)), TM30, NVL(ROR02,' '), NVL(ST02,' '), NVL(C1.CP13,' '), NVL(C1.CP12,' '), TM01, NVL(ROR01,' ') " & _
                  " FROM TRADEMARK, ReasonOfRelief, CaseProgress C1, ( " & strSQL_1 & " ) C2, Staff WHERE C2.CP01=C1.CP01 AND C2.CP02=C1.CP02 AND C2.CP03=C1.CP03 AND C2.CP04=C1.CP04 AND C2.CP05=C1.CP05(+) AND C2.CP10=C1.CP10(+) AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) AND TM31=ROR01(+) AND C1.CP13=ST01(+) "
'法務
strSQL_1 = "SELECT CP01,CP02,CP03,CP04,MAX(CP05) CP05,CP10 FROM CASEPROGRESS,LAWCASE WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP10='999' " & StrSQL3 & " GROUP BY CP01,CP02,CP03,CP04,CP10"
strSql = strSql & " UNION SELECT LC01||'-'||LC02||'-'||LC03||'-'||LC04, NVL(LC05,NVL(LC06,LC07)), LC09, NVL(ROR02,' '), NVL(ST02,' '), NVL(C1.CP13,' '), NVL(C1.CP12,' '), LC01, NVL(ROR01,' ') " & _
                  " FROM LAWCASE, ReasonOfRelief, CaseProgress C1, ( " & strSQL_1 & " ) C2, Staff WHERE C2.CP01=C1.CP01 AND C2.CP02=C1.CP02 AND C2.CP03=C1.CP03 AND C2.CP04=C1.CP04 AND C2.CP05=C1.CP05(+) AND C2.CP10=C1.CP10(+) AND C1.CP01=LC01(+) AND C1.CP02=LC02(+) AND C1.CP03=LC03(+) AND C1.CP04=LC04(+) AND LC10=ROR01(+) AND C1.CP13=ST01(+) "
'顧問
strSQL_1 = "SELECT CP01,CP02,CP03,CP04,MAX(CP05) CP05,CP10 FROM CASEPROGRESS,HIRECASE WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP10='999' " & StrSQL4 & " GROUP BY CP01,CP02,CP03,CP04,CP10"
strSql = strSql & " UNION SELECT HC01||'-'||HC02||'-'||HC03||'-'||HC04, NVL(HC05,NVL(HC06,HC07)), HC10, NVL(ROR02,' '), NVL(ST02,' '), NVL(C1.CP13,' '), NVL(C1.CP12,' '), HC01, NVL(ROR01,' ') " & _
                  " FROM HIRECASE, ReasonOfRelief, CaseProgress C1, ( " & strSQL_1 & " ) C2, Staff WHERE C2.CP01=C1.CP01 AND C2.CP02=C1.CP02 AND C2.CP03=C1.CP03 AND C2.CP04=C1.CP04 AND C2.CP05=C1.CP05(+) AND C2.CP10=C1.CP10(+) AND C1.CP01=HC01(+) AND C1.CP02=HC02(+) AND C1.CP03=HC03(+) AND C1.CP04=HC04(+) AND HC11=ROR01(+) AND C1.CP13=ST01(+) "
'服務(專)
strSQL_1 = "SELECT CP01,CP02,CP03,CP04,MAX(CP05) CP05,CP10 FROM CASEPROGRESS,SERVICEPRACTICE,SYSTEMKIND WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP04=SK01(+) AND CP10='913' AND SK02=5 " & strSQL5 & " GROUP BY CP01,CP02,CP03,CP04,CP10"
strSql = strSql & " UNION SELECT SP01||'-'||SP02||'-'||SP03||'-'||SP04, NVL(SP05,NVL(SP06,SP07)), SP16, NVL(ROR02,' '), NVL(ST02,' '), NVL(C1.CP13,' '), NVL(C1.CP12,' '), SP01, NVL(ROR01,' ') " & _
                  " FROM SERVICEPRACTICE, ReasonOfRelief, CaseProgress C1, ( " & strSQL_1 & " ) C2, Staff WHERE C2.CP01=C1.CP01 AND C2.CP02=C1.CP02 AND C2.CP03=C1.CP03 AND C2.CP04=C1.CP04 AND C2.CP05=C1.CP05(+) AND C2.CP10=C1.CP10(+) AND C1.CP01=SP01(+) AND C1.CP02=SP02(+) AND C1.CP03=SP03(+) AND C1.CP04=SP04(+) AND SP17=ROR01(+) AND C1.CP13=ST01(+) "
'服務(商)
strSQL_1 = "SELECT CP01,CP02,CP03,CP04,MAX(CP05) CP05,CP10 FROM CASEPROGRESS,SERVICEPRACTICE,SYSTEMKIND WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP04=SK01(+) AND CP10='704' AND SK02=6 " & strSQL5 & " GROUP BY CP01,CP02,CP03,CP04,CP10"
strSql = strSql & " UNION SELECT SP01||'-'||SP02||'-'||SP03||'-'||SP04, NVL(SP05,NVL(SP06,SP07)), SP16, NVL(ROR02,' '), NVL(ST02,' '), NVL(C1.CP13,' '), NVL(C1.CP12,' '), SP01, NVL(ROR01,' ') " & _
                  " FROM SERVICEPRACTICE, ReasonOfRelief, CaseProgress C1, ( " & strSQL_1 & " ) C2, Staff WHERE C2.CP01=C1.CP01 AND C2.CP02=C1.CP02 AND C2.CP03=C1.CP03 AND C2.CP04=C1.CP04 AND C2.CP05=C1.CP05(+) AND C2.CP10=C1.CP10(+) AND C1.CP01=SP01(+) AND C1.CP02=SP02(+) AND C1.CP03=SP03(+) AND C1.CP04=SP04(+) AND SP17=ROR01(+) AND C1.CP13=ST01(+) "
'服務(法)
strSQL_1 = "SELECT CP01,CP02,CP03,CP04,MAX(CP05) CP05,CP10 FROM CASEPROGRESS,SERVICEPRACTICE,SYSTEMKIND WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP04=SK01(+) AND CP10='999' AND SK02=7 " & strSQL5 & " GROUP BY CP01,CP02,CP03,CP04,CP10"
strSql = strSql & " UNION SELECT SP01||'-'||SP02||'-'||SP03||'-'||SP04, NVL(SP05,NVL(SP06,SP07)), SP16, NVL(ROR02,' '), NVL(ST02,' '), NVL(C1.CP13,' '), NVL(C1.CP12,' '), SP01, NVL(ROR01,' ') " & _
                  " FROM SERVICEPRACTICE, ReasonOfRelief, CaseProgress C1, ( " & strSQL_1 & " ) C2, Staff WHERE C2.CP01=C1.CP01 AND C2.CP02=C1.CP02 AND C2.CP03=C1.CP03 AND C2.CP04=C1.CP04 AND C2.CP05=C1.CP05(+) AND C2.CP10=C1.CP10(+) AND C1.CP01=SP01(+) AND C1.CP02=SP02(+) AND C1.CP03=SP03(+) AND C1.CP04=SP04(+) AND SP17=ROR01(+) AND C1.CP13=ST01(+) "
'服務(顧)
strSQL_1 = "SELECT CP01,CP02,CP03,CP04,MAX(CP05) CP05,CP10 FROM CASEPROGRESS,SERVICEPRACTICE,SYSTEMKIND WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP04=SK01(+) AND CP10='999' AND SK02=8 " & strSQL5 & " GROUP BY CP01,CP02,CP03,CP04,CP10"
strSql = strSql & " UNION SELECT SP01||'-'||SP02||'-'||SP03||'-'||SP04, NVL(SP05,NVL(SP06,SP07)), SP16, NVL(ROR02,' '), NVL(ST02,' '), NVL(C1.CP13,' '), NVL(C1.CP12,' '), SP01, NVL(ROR01,' ') " & _
                  " FROM SERVICEPRACTICE, ReasonOfRelief, CaseProgress C1, ( " & strSQL_1 & " ) C2, Staff WHERE C2.CP01=C1.CP01 AND C2.CP02=C1.CP02 AND C2.CP03=C1.CP03 AND C2.CP04=C1.CP04 AND C2.CP05=C1.CP05(+) AND C2.CP10=C1.CP10(+) AND C1.CP01=SP01(+) AND C1.CP02=SP02(+) AND C1.CP03=SP03(+) AND C1.CP04=SP04(+) AND SP17=ROR01(+) AND C1.CP13=ST01(+) "
If Me.Text1(3).Text = "1" Then
   strSql = strSql & " ORDER BY 8,1 "
ElseIf Me.Text1(3).Text = "2" Then
   strSql = strSql & " ORDER BY 8,3,1 "
ElseIf Me.Text1(3).Text = "3" Then
   strSql = strSql & " ORDER BY 8,9,1 "
Else
   strSql = strSql & " ORDER BY 8,7,6,1 "
End If
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing
rs.CursorLocation = adUseClient
rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
   PrintData rs
   ShowPrintOk
Else
   ShowNoData
End If
If rs.State <> adStateClosed Then rs.Close
Set rs = Nothing

Screen.MousePointer = vbDefault
End Sub
Private Sub PrintData(rsP As ADODB.Recordset)
Dim intPage As Integer
Dim strSys As String
Dim iPrint As Integer
Dim j As Integer
Dim i As Integer
Dim intCnt As Integer

GetPrintLeft
Printer.Orientation = vbPRORPortrait
intCnt = 0
intPage = intPage + 1
strSys = "" & rsP.Fields(7).Value
PrintTitle intPage, strSys
rsP.MoveFirst
iPrint = 2700
Do While Not rsP.EOF
   For j = 0 To 4
      Printer.CurrentX = PLeft(j)
      Printer.CurrentY = iPrint
      If j = 1 Then
         Printer.Print Left("" & rsP.Fields(j).Value, 20)
      ElseIf j = 2 Then
         Printer.Print ChangeTStringToTDateString(Val(rsP.Fields(j).Value - 19110000))
      ElseIf j = 3 Then
         Printer.Print Left("" & rsP.Fields(j).Value, 7)
      Else
         Printer.Print "" & rsP.Fields(j).Value
      End If
   Next j
   iPrint = iPrint + 300
   i = i + 1
   intCnt = intCnt + 1
   rsP.MoveNext
   If rsP.EOF Then Exit Do
   If "" & rsP.Fields(7).Value <> strSys Then
      PrintEnd intCnt
      strSys = "" & rsP.Fields(7).Value
      Printer.NewPage
      intPage = intPage + 1
      PrintTitle intPage, strSys
      iPrint = 2700
      i = 0
      intCnt = 0
   ElseIf i > 40 Then
      Printer.NewPage
      intPage = intPage + 1
      PrintTitle intPage, strSys
      iPrint = 2700
      i = 0
   End If
Loop
PrintEnd intCnt
Printer.EndDoc

End Sub

Private Sub GetPrintLeft()
   PLeft(0) = 200
   PLeft(1) = 2250
   PLeft(2) = 7200
   PLeft(3) = 8250
   PLeft(4) = 10200
End Sub

Private Sub PrintTitle(ByVal Page As String, SYS As String)
'Page : 頁數
'Sys  : 系統類別
Dim i As Integer
 
i = 500
Printer.FontName = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 4750
Printer.CurrentY = i
Printer.Print "閉卷清單"

Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 800
Printer.Print "列印人　 : " & strUserName
Printer.CurrentX = 4000
Printer.CurrentY = i + 800
Printer.Print "閉卷日期 : " & ChangeTStringToTDateString(Me.Text1(1).Text) & "－" & ChangeTStringToTDateString(Me.Text1(2).Text)
Printer.CurrentX = 7000 + 1500
Printer.CurrentY = i + 800
Printer.Print "列印日期 : " & ChangeTStringToTDateString("" & (Val(ServerDate) - 19110000))

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1100
Printer.Print "系統類別 : " & SYS

Printer.CurrentX = 4000
Printer.CurrentY = i + 1100
Printer.Print "排序條件 : " & IIf(Me.Text1(3).Text = "1", "本所案號", IIf(Me.Text1(3).Text = "2", "閉卷日期", IIf(Me.Text1(3).Text = "3", "閉卷原因", "智權人員")))
Printer.CurrentX = 7000 + 1500
Printer.CurrentY = i + 1100
Printer.Print "頁　　次 : " & Page
Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1400
Printer.Print String(250, "-")

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1400
Printer.Print String(250, "-")
 
Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1700
Printer.Print "本所案號"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = i + 1700
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = i + 1700
Printer.Print "閉卷日期"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = i + 1700
Printer.Print "閉卷原因"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = i + 1700
Printer.Print "智權人員"

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 2000
Printer.Print String(250, "-")

End Sub

Private Sub PrintEnd(Cnt As Integer)
'Cnt : 筆數
Printer.CurrentX = PLeft(0)
Printer.CurrentY = 2700 + 41 * 300
Printer.Print String(250, "-")

Printer.CurrentX = PLeft(0)
Printer.CurrentY = 2700 + 42 * 300
Printer.Print "共 " & Format(Cnt, "#,##0") & " 筆"

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Me.Text1(0).Text = Systemkind_g
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040139 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Index = 3 Then
      If (KeyAscii < 49 Or KeyAscii > 52) And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Select Case Index
Case 2 '閉卷迄日
   If Val(Me.Text1(Index).Text) < Val(Me.Text1(1).Text) Then
      MsgBox "閉卷日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
      Me.Text1(1).SetFocus
      TextInverse Me.Text1(1)
   End If
End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Cancel = False
If Len(Me.Text1(Index).Text) <= 0 Then Exit Sub
If CheckKeyIn(Index) = False Then
   Cancel = True
   Me.Text1(Index).SetFocus
   Text1_GotFocus Index
   Exit Sub
End If
End Sub

Private Function CheckKeyIn(Index As Integer) As Boolean
CheckKeyIn = False
Select Case Index
Case 0 '系統類別
   If Len(Me.Text1(Index).Text) <= 0 Then
      MsgBox "請輸入系統類別!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
Case 1 '閉卷日期起
   If Len(Me.Text1(Index).Text) <= 0 Then
      MsgBox "請輸入閉卷起日!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
   If PUB_CheckKeyInDate(Me.Text1(Index)) = -1 Then
      Exit Function
   End If
   If Val(Me.Text1(Index).Text) + 19110000 > ServerDate Then
      MsgBox "閉卷日期不可大於系統日期!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
Case 2 '閉卷日期迄
   If Len(Me.Text1(Index).Text) <= 0 Then
      MsgBox "請輸入閉卷迄日!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
   If PUB_CheckKeyInDate(Me.Text1(Index)) = -1 Then
      Exit Function
   End If
   If Val(Me.Text1(Index).Text) + 19110000 > ServerDate Then
      MsgBox "閉卷日期不可大於系統日期!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
   'Modify By Cheng 2002/06/10
   '改在Lost_Focus時檢查
'   If Val(Me.Text1(Index).Text) < Val(Me.Text1(1).Text) Then
'      MsgBox "閉卷迄日不可小於閉卷起日!!!", vbExclamation + vbOKOnly
'      Exit Function
'   End If
Case 3 '排序條件
   If Len(Me.Text1(Index).Text) <= 0 Then
      MsgBox "請輸入排序條件!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
End Select
CheckKeyIn = True
End Function
