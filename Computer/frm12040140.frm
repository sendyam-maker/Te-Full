VERSION 5.00
Begin VB.Form frm12040140 
   BorderStyle     =   1  '單線固定
   Caption         =   "分所案號檢核表"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5910
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   1350
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1620
      Width           =   405
   End
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
      TabIndex        =   5
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
      TabIndex        =   6
      Top             =   60
      Width           =   800
   End
   Begin VB.Label lbl 
      Caption         =   "排序條件 :                 (1.收文日 2.智權人員 3.本所案號)"
      Height          =   255
      Index           =   4
      Left            =   390
      TabIndex        =   11
      Top             =   1650
      Width           =   5385
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "－"
      Height          =   180
      Index           =   3
      Left            =   2520
      TabIndex        =   10
      Top             =   960
      Width           =   180
   End
   Begin VB.Label lbl 
      Caption         =   "報表內容 :                 (1.檢核表 2.無分所資料清單)"
      Height          =   255
      Index           =   2
      Left            =   390
      TabIndex        =   9
      Top             =   1290
      Width           =   5385
   End
   Begin VB.Label lbl 
      Caption         =   "收文日期 : "
      Height          =   255
      Index           =   1
      Left            =   390
      TabIndex        =   8
      Top             =   960
      Width           =   945
   End
   Begin VB.Label lbl 
      Caption         =   "系統類別 :                                                                           (ALL：全部)"
      Height          =   255
      Index           =   0
      Left            =   390
      TabIndex        =   7
      Top             =   630
      Width           =   5355
   End
End
Attribute VB_Name = "frm12040140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim PLeft(0 To 7) As Integer
'Add By Cheng 2002/06/10
Dim m_blnValidate As Boolean

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
            Text1_LostFocus ii
            If m_blnValidate = False Then
               Me.Text1(1).SetFocus
               Text1_GotFocus 1
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
Dim rs As New ADODB.Recordset

Screen.MousePointer = vbHourglass
strSql = "": strSQL1 = "": strSQL2 = "": StrSQL3 = "": StrSQL4 = "": strSQL5 = ""

strSQL1 = strSQL1 & " AND CP05>=" & Val(Me.Text1(1).Text) + 19110000 & " AND CP05<=" & Val(Me.Text1(2).Text) + 19110000 & " AND 'Y'=CP31 "
strSQL2 = strSQL1
StrSQL3 = strSQL1
StrSQL4 = strSQL1
strSQL5 = strSQL1

strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(IIf(Me.Text1(0).Text <> "ALL", Me.Text1(0).Text, GetAllSysKind(Me.Text1(0))), 1) & ") "
strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(IIf(Me.Text1(0).Text <> "ALL", Me.Text1(0).Text, GetAllSysKind(Me.Text1(0))), 2) & ") "
StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(IIf(Me.Text1(0).Text <> "ALL", Me.Text1(0).Text, GetAllSysKind(Me.Text1(0))), 3) & ") "
StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(IIf(Me.Text1(0).Text <> "ALL", Me.Text1(0).Text, GetAllSysKind(Me.Text1(0))), 4) & ") "
strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr(IIf(Me.Text1(0).Text <> "ALL", Me.Text1(0).Text, GetAllSysKind(Me.Text1(0))), 5) & ") "

'專利
strSql = strSql & " SELECT PA01||'-'||PA02||'-'||PA03||'-'||PA04, NVL(PA47,' '), NVL(PA05,NVL(PA06,PA07)), NVL(NA03,NA04), CP05, NVL(ST02,' '), NVL(CP13,' '), DECODE(SUBSTR(CP12,1,2),'S1',SUBSTR(CP12,1,2),'S2',SUBSTR(CP12,1,2),'S3',SUBSTR(CP12,1,2),'S4',SUBSTR(CP12,1,2),'ZZ') " & _
                  " FROM PATENT, NATION, CaseProgress, Staff WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA09=NA01(+) AND CP13=ST01(+) " & strSQL1 & IIf(Me.Text1(3).Text = "2", " AND PA47 IS NULL ", " ")
'商標
strSql = strSql & " UNION SELECT TM01||'-'||TM02||'-'||TM03||'-'||TM04, NVL(TM34,' '), NVL(TM05,NVL(TM06,TM07)), NVL(NA03,NA04), CP05, NVL(ST02,' '), NVL(CP13,' '), DECODE(SUBSTR(CP12,1,2),'S1',SUBSTR(CP12,1,2),'S2',SUBSTR(CP12,1,2),'S3',SUBSTR(CP12,1,2),'S4',SUBSTR(CP12,1,2),'ZZ') " & _
                  " FROM TRADEMARK, NATION, CaseProgress, Staff WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND TM10=NA01(+) AND CP13=ST01(+) " & strSQL2 & IIf(Me.Text1(3).Text = "2", " AND TM34 IS NULL ", " ")
'法務
strSql = strSql & " UNION SELECT LC01||'-'||LC02||'-'||LC03||'-'||LC04, NVL(LC16,' '), NVL(LC05,NVL(LC06,LC07)), NVL(NA03,NA04), CP05, NVL(ST02,' '), NVL(CP13,' '), DECODE(SUBSTR(CP12,1,2),'S1',SUBSTR(CP12,1,2),'S2',SUBSTR(CP12,1,2),'S3',SUBSTR(CP12,1,2),'S4',SUBSTR(CP12,1,2),'ZZ') " & _
                  " FROM LAWCASE, NATION, CaseProgress, Staff WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND LC15=NA01(+) AND CP13=ST01(+) " & StrSQL3 & IIf(Me.Text1(3).Text = "2", " AND LC16 IS NULL ", " ")
'顧問
strSql = strSql & " UNION SELECT HC01||'-'||HC02||'-'||HC03||'-'||HC04, NVL(HC07,' '), NVL(HC06,' '), NVL(NA03,NA04), CP05, NVL(ST02,' '), NVL(CP13,' '), DECODE(SUBSTR(CP12,1,2),'S1',SUBSTR(CP12,1,2),'S2',SUBSTR(CP12,1,2),'S3',SUBSTR(CP12,1,2),'S4',SUBSTR(CP12,1,2),'ZZ') " & _
                  " FROM HIRECASE, NATION, CaseProgress, Staff WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND '000'=NA01(+) AND CP13=ST01(+) " & StrSQL4 & IIf(Me.Text1(3).Text = "2", " AND HC07 IS NULL ", " ")
'服務
strSql = strSql & " UNION SELECT SP01||'-'||SP02||'-'||SP03||'-'||SP04, NVL(SP28,' '), NVL(SP05,NVL(SP06,SP07)), NVL(NA03,NA04), CP05, NVL(ST02,' '), NVL(CP13,' '), DECODE(SUBSTR(CP12,1,2),'S1',SUBSTR(CP12,1,2),'S2',SUBSTR(CP12,1,2),'S3',SUBSTR(CP12,1,2),'S4',SUBSTR(CP12,1,2),'ZZ') " & _
                  " FROM SERVICEPRACTICE, NATION, CaseProgress, Staff WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND SP09=NA01(+) AND CP13=ST01(+) " & strSQL5 & IIf(Me.Text1(3).Text = "2", " AND SP28 IS NULL ", " ")

If Me.Text1(4).Text = "1" Then
   strSql = strSql & " ORDER BY 8,5,1 "
ElseIf Me.Text1(4).Text = "2" Then
   strSql = strSql & " ORDER BY 8,7,1 "
Else
   strSql = strSql & " ORDER BY 8,1 "
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
Dim strSaleZone As String
Dim iPrint As Integer
Dim j As Integer
Dim i As Integer
Dim intCnt As Integer

GetPrintLeft
Printer.Orientation = vbPRORPortrait
intCnt = 0
intPage = intPage + 1
strSaleZone = Left("" & rsP.Fields(7).Value, 2)
PrintTitle intPage
rsP.MoveFirst
iPrint = 3000
Do While Not rsP.EOF
   For j = 0 To 5
      Printer.CurrentX = PLeft(j)
      Printer.CurrentY = iPrint
      If j = 2 Then
         Printer.Print Left("" & rsP.Fields(j).Value, 15)
      ElseIf j = 3 Then
         Printer.Print Left("" & rsP.Fields(j).Value, 4)
      ElseIf j = 4 Then
         Printer.Print ChangeTStringToTDateString(Val(rsP.Fields(j).Value - 19110000))
      Else
         Printer.Print "" & rsP.Fields(j).Value
      End If
   Next j
   iPrint = iPrint + 300
   i = i + 1
   intCnt = intCnt + 1
   rsP.MoveNext
   If rsP.EOF Then Exit Do
   If Left("" & rsP.Fields(7).Value, 2) <> strSaleZone Then
      strSaleZone = Left("" & rsP.Fields(7).Value, 2)
      Printer.NewPage
      intPage = intPage + 1
      PrintTitle intPage
      iPrint = 3000
      i = 0
      intCnt = 0
   ElseIf i > 40 Then
      Printer.NewPage
      intPage = intPage + 1
      PrintTitle intPage
      iPrint = 3000
      i = 0
   End If
Loop
Printer.EndDoc

End Sub

Private Sub GetPrintLeft()
   PLeft(0) = 200
   PLeft(1) = 2200
   PLeft(2) = 3825
   PLeft(3) = 7700
   PLeft(4) = 8950
   PLeft(5) = 10200
End Sub

Private Sub PrintTitle(ByVal Page As String)
'Page : 頁數
Dim i As Integer
 
i = 500
Printer.FontName = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 4100
Printer.CurrentY = i
Printer.Print "分所案號檢核表"

Printer.Font.Underline = False
Printer.Font.Size = 12
Printer.Font.Bold = False

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 800
Printer.Print "列印人　 : " & strUserName
Printer.CurrentX = 3700
Printer.CurrentY = i + 800
Printer.Print "收文日期 : " & ChangeTStringToTDateString(Me.Text1(1).Text) & "－" & ChangeTStringToTDateString(Me.Text1(2).Text)
Printer.CurrentX = 7000 + 1500
Printer.CurrentY = i + 800
Printer.Print "列印日期 : " & ChangeTStringToTDateString("" & (Val(ServerDate) - 19110000))

Printer.CurrentX = 3700
Printer.CurrentY = i + 1100
Printer.Print "排序條件 : " & IIf(Me.Text1(4).Text = "1", "收文日", IIf(Me.Text1(4).Text = "2", "智權人員", "本所案號"))
Printer.CurrentX = 7000 + 1500
Printer.CurrentY = i + 1100
Printer.Print "頁　　次 : " & Page

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1400
Printer.Print "系統類別 : " & Me.Text1(0).Text
Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 1700
Printer.Print String(250, "-")
 
Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 2000
Printer.Print "本所案號"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = i + 2000
Printer.Print "分所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = i + 2000
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = i + 2000
Printer.Print "申請國家"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = i + 2000
Printer.Print "收文日"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = i + 2000
Printer.Print "智權人員"

Printer.CurrentX = PLeft(0)
Printer.CurrentY = i + 2300
Printer.Print String(250, "-")

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Me.Text1(0).Text = Systemkind_g
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040140 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Index = 3 Then
      If (KeyAscii < 49 Or KeyAscii > 50) And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   ElseIf Index = 4 Then
      If (KeyAscii < 49 Or KeyAscii > 51) And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
m_blnValidate = True
Select Case Index
Case 2 '收文迄日
If Val(Me.Text1(Index).Text) < Val(Me.Text1(1).Text) Then
   MsgBox "收文日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
   Me.Text1(1).SetFocus
   TextInverse Me.Text1(1)
   m_blnValidate = False
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
Case 1 '收文日期起
   If Len(Me.Text1(Index).Text) <= 0 Then
      MsgBox "請輸入收文起日!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
   If PUB_CheckKeyInDate(Me.Text1(Index)) = -1 Then
      Exit Function
   End If
   If Val(Me.Text1(Index).Text) + 19110000 > ServerDate Then
      MsgBox "收文日期不可大於系統日期!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
Case 2 '收文日期迄
   If Len(Me.Text1(Index).Text) <= 0 Then
      MsgBox "請輸入收文迄日!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
   If PUB_CheckKeyInDate(Me.Text1(Index)) = -1 Then
      Exit Function
   End If
   If Val(Me.Text1(Index).Text) + 19110000 > ServerDate Then
      MsgBox "收文日期不可大於系統日期!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
   'Modify By Cheng 2002/06/10
   '改在Lost_Focus時檢查
'   If Val(Me.Text1(Index).Text) < Val(Me.Text1(1).Text) Then
'      MsgBox "收文迄日不可小於收文起日!!!", vbExclamation + vbOKOnly
'      Exit Function
'   End If
Case 3 '報告內容
   If Len(Me.Text1(Index).Text) <= 0 Then
      MsgBox "請輸入報告內容!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
Case 4 '排序條件
   If Len(Me.Text1(Index).Text) <= 0 Then
      MsgBox "請輸入排序條件!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
End Select
CheckKeyIn = True
End Function
