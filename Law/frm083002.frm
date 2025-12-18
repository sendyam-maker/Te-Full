VERSION 5.00
Begin VB.Form frm083002 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人案件收達 / 提申管制表"
   ClientHeight    =   3195
   ClientLeft      =   2895
   ClientTop       =   1710
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4560
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   8
      Left            =   1344
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   7
      Left            =   1344
      MaxLength       =   9
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   6
      Left            =   2664
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   1344
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   2670
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1344
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2664
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1344
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1344
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3528
      TabIndex        =   10
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2700
      TabIndex        =   9
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   276
      Left            =   2544
      TabIndex        =   17
      Top             =   2160
      Width           =   1608
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2184
      X2              =   2544
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2184
      X2              =   2544
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2184
      X2              =   2544
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "列  印  別：           (1.未收達 2.未提申)"
      Height          =   180
      Index           =   5
      Left            =   384
      TabIndex        =   16
      Top             =   2520
      Width           =   2916
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代  理  人："
      Height          =   180
      Index           =   4
      Left            =   384
      TabIndex        =   15
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   384
      TabIndex        =   14
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發  文  日："
      Height          =   180
      Index           =   2
      Left            =   384
      TabIndex        =   13
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   1
      Left            =   384
      TabIndex        =   12
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   384
      TabIndex        =   11
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "frm083002"
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
   If ChkRange(Text1(3), Text1(4), "發文日") = False Then Exit Sub
   'Add By Cheng 2002/03/22
   If PUB_CheckKeyInDate(Me.Text1(3)) = -1 Then
      Me.Text1(3).SetFocus
      Text1_GotFocus 3
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Text1(4)) = -1 Then
      Me.Text1(4).SetFocus
      Text1_GotFocus 4
      Exit Sub
   End If
   
   If Text1(8) = "" Then
      Text1(8).SetFocus
      MsgBox "列印別不得為空值 !", vbCritical
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
   'Modify By Cheng 2002/03/26
   '加CP09<'C'的控制
'   strExc(0) = "SELECT DECODE(CP14,ST01,ST02)," & _
'      "SUBSTR(CP27,1,4)-1911||'/'||SUBSTR(CP27,5,2)||'/'||SUBSTR(CP27,7,2)," & _
'      "SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," & _
'      "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04)," & _
'      "NVL(LC05, NVL(LC06, lC07)),DECODE(LC15,NA01,NA03)," & _
'      "DECODE(CP01||CP10,CPM01||CPM02,CPM03),NVL(FA05,NVL(FA04,FA06))," & _
'      "SUBSTR(CP46,1,4)-1911||'/'||SUBSTR(CP46,5,2)||'/'||SUBSTR(CP46,7,2) " & _
'      "FROM STAFF,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP," & _
'      "FAGENT,NATION WHERE CP14=ST01(+) AND (CP01=CPM01(+) AND CP10=CPM02(+)) " & _
'      "AND LC15=NA01(+) AND substr(CP44,1,8)=FA01(+) AND substr(CP44,9,1)=FA02(+) AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 " & _
'      strGetcdnSQL & " ORDER BY CP14,CP27,CP01,CP02,CP03,CP04"
   strExc(0) = "SELECT DECODE(CP14,ST01,ST02)," & _
      "SUBSTR(CP27,1,4)-1911||'/'||SUBSTR(CP27,5,2)||'/'||SUBSTR(CP27,7,2)," & _
      "SUBSTR(CP05,1,4)-1911||'/'||SUBSTR(CP05,5,2)||'/'||SUBSTR(CP05,7,2)," & _
      "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04)," & _
      "NVL(LC05, NVL(LC06, lC07)),DECODE(LC15,NA01,NA03)," & _
      "DECODE(CP01||CP10,CPM01||CPM02,CPM03),NVL(FA05,NVL(FA04,FA06))," & _
      "SUBSTR(CP46,1,4)-1911||'/'||SUBSTR(CP46,5,2)||'/'||SUBSTR(CP46,7,2) " & _
      "FROM STAFF,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP," & _
      "FAGENT,NATION WHERE CP14=ST01(+) AND (CP01=CPM01(+) AND CP10=CPM02(+)) " & _
      "AND LC15=NA01(+) AND substr(CP44,1,8)=FA01(+) AND substr(CP44,9,1)=FA02(+) AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND CP09<'C' " & _
      strGetcdnSQL & " ORDER BY CP14,CP27,CP01,CP02,CP03,CP04"
   If RsTemp.State = adStateOpen Then RsTemp.Close
   RsTemp.Open strExc(0), cnnConnection
   If RsTemp.EOF And RsTemp.BOF Then
      m_print = 1
      MsgBox "資料庫內無資料 !", vbInformation
      Exit Sub
   End If
   i = 1
   Page = 1
 '  Dialog1.ShowPrinter
   CaseTitle TmpArea, 1
   iPrint = 2700
   With RsTemp
   Do While Not .EOF
      Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
      If IsNull(.Fields(0)) = False Then Printer.Print Format(.Fields(0), "!@@@@")
      Printer.CurrentX = PLeft(1):      Printer.CurrentY = iPrint
      If .Fields(1) <> "//" Then
         St = .Fields(1)
      Else
         St = ""
      End If
      Printer.Print St
      Printer.CurrentX = PLeft(2):      Printer.CurrentY = iPrint
      If .Fields(2) <> "//" Then
         St = .Fields(2)
      Else
         St = ""
      End If
      Printer.Print St
      Printer.CurrentX = PLeft(3):      Printer.CurrentY = iPrint
      Printer.Print .Fields(3)
      Printer.CurrentX = PLeft(4):      Printer.CurrentY = iPrint
      If IsNull(.Fields(4)) = False Then Printer.Print Format(Left(.Fields(4), 8), "!@@@@@@@@")
      Printer.CurrentX = PLeft(5):      Printer.CurrentY = iPrint
      If IsNull(.Fields(5)) = False Then Printer.Print Format(.Fields(5), "!@@@@@@")
      Printer.CurrentX = PLeft(6):      Printer.CurrentY = iPrint
      If IsNull(.Fields(6)) = False Then Printer.Print Format(.Fields(6), "!@@@@@@")
      Printer.CurrentX = PLeft(7):      Printer.CurrentY = iPrint
      If IsNull(.Fields(7)) = False Then Printer.Print Format(.Fields(7), "!@@@@@@@@@@@@@@@@@@@@")
      If Text1(8).Text = "2" Then
         Printer.CurrentX = PLeft(8):      Printer.CurrentY = iPrint
         If Text1(8).Text = "2" Then
            If .Fields(8) <> "//" Then
               St = .Fields(8)
            Else
               St = ""
            End If
            Printer.Print St
         End If
      End If
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
   PLeft(0) = 500:    PLeft(1) = 1600
   PLeft(2) = 2700:   PLeft(3) = 3800
   PLeft(4) = 5200:   PLeft(5) = 7500
   PLeft(6) = 9100:   PLeft(7) = 11600
   PLeft(8) = 14700
End Sub

Private Sub CaseTitle(ByVal Area As String, ByVal Page As String)
 Dim i As Integer, St As String
   i = 500
   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000:         Printer.CurrentY = i
   If Text1(8).Text = "1" Then
      St = "代理人案件收達管制表"
   Else
      St = "代理人案件提申管制表"
   End If
   Printer.Print St
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 6200:         Printer.CurrentY = i + 500
   Printer.Print "發文日 : " & ChangeTStringToTDateString(Text1(3)) & _
      " - " & ChangeTStringToTDateString(Text1(4))
   Printer.Font.Bold = False
   Printer.CurrentX = 500:              Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 13000:            Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
   Printer.CurrentX = 13000:            Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
   Printer.CurrentX = 500:              Printer.CurrentY = i + 1400
   Printer.Print String(205, "-")
   Printer.CurrentX = PLeft(0):         Printer.CurrentY = i + 1700
   Printer.Print "承辦人"
   Printer.CurrentX = PLeft(1):         Printer.CurrentY = i + 1700
   Printer.Print "發文日"
   Printer.CurrentX = PLeft(2):         Printer.CurrentY = i + 1700
   Printer.Print "收文日"
   Printer.CurrentX = PLeft(3):         Printer.CurrentY = i + 1700
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(4):         Printer.CurrentY = i + 1700
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(5):         Printer.CurrentY = i + 1700
   Printer.Print "申請國家"
   Printer.CurrentX = PLeft(6):         Printer.CurrentY = i + 1700
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(7):         Printer.CurrentY = i + 1700
   Printer.Print "代理人"
  If Text1(8).Text = "2" Then
   Printer.CurrentX = PLeft(8):         Printer.CurrentY = i + 1700
   Printer.Print "收達日"
  End If
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
 Dim strTemp As Variant
 Dim strSql As String
  
 strTemp = Split(UCase(Text1(0).Text), ",")
 For i = 0 To UBound(strTemp)
     'Modify By Sindy 2009/07/24 增加LIN系統類別
     If strTemp(i) = "CFL" Or strTemp(i) = "FCL" Or strTemp(i) = "LIN" Then
        strSql = strSql & strTemp(i)
        strSql = strSql & "','"
     End If
 Next i
 
 
   strExc(1) = " AND CP01 IN('" & strSql & "')"
   If Text1(8).Text = "1" Then
      strExc(1) = strExc(1) & " AND CP46 IS NULL AND CP47 IS NULL"
   ElseIf Text1(8).Text = "2" Then
      strExc(1) = strExc(1) & " AND CP47 IS NULL"
   End If
   If Text1(1).Text = "" And Text1(2).Text <> "" Then
      strExc(1) = strExc(1) + " and CP10<='" + Text1(1) + "'"
   ElseIf Text1(1).Text <> "" And Text1(2).Text <> "" Then
      strExc(1) = strExc(1) + " and (CP10 BETWEEN '" + Text1(1) + "' AND '" + Text1(2) + "')"
   End If
   If Text1(3).Text = "" And Text1(4).Text <> "" Then
      strExc(1) = strExc(1) + " and CP27<='" + ChangeTStringToWString(Text1(4)) + "'"
   ElseIf Text1(3).Text <> "" And Text1(4).Text <> "" Then
      strExc(1) = strExc(1) + " and (CP27 BETWEEN '" + ChangeTStringToWString(Text1(3)) + _
         "' AND '" + ChangeTStringToWString(Text1(4)) + "')"
   End If
   If Text1(5).Text = "" And Text1(6).Text <> "" Then
      strExc(1) = strExc(1) + " and LC15<='" + Text1(5) + "'"
   ElseIf Text1(5).Text <> "" And Text1(6).Text <> "" Then
      strExc(1) = strExc(1) + " and (LC15 BETWEEN '" + Text1(5) + "' AND '" + Text1(6) + "')"
   End If
   If Text1(7).Text <> "" Then strExc(1) = strExc(1) + " and CP44='" + GetNewFagent(Text1(7).Text) + "'"
   strGetcdnSQL = strExc(1)
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm083002 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 8
         If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case Else
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTempName As String, i As Integer, t As Integer
 Dim j As Integer
 Dim strTemp1 As Variant
 Dim strTemp2 As Variant
 Dim s As Integer
 
   If Text1(Index) = "" Then
      If Index = 7 Then Label2.Caption = ""
      Exit Sub
   End If
   Select Case Index
      Case 0
'         If ChkSysName(Text1(Index)) = False Then Cancel = True
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


      Case 1, 2
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCaseProperty(Text1(0).Text, Text1(Index), strTempName, False) = False Then
         If ClsPDGetCaseProperty(Text1(0).Text, Text1(Index), strTempName, False) = False Then
            Cancel = True
         End If
      Case 3, 4
         If CheckIsTaiwanDate(Text1(Index)) Then
            If Text1(3) <> "" And Text1(4) <> "" Then
               If Val(Text1(4)) - Val(Text1(3)) < 0 Then
                  MsgBox "前一日期大於後一日期", vbCritical
                  Cancel = True
               End If
            End If
         Else
            Cancel = True
         End If
      Case 5, 6
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetNation(Text1(Index), strTempName) = False Then
         If ClsPDGetNation(Text1(Index), strTempName) = False Then
            Cancel = True
         End If
      Case 7
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetAgent(Text1(Index), strTempName) Then
         If ClsPDGetAgent(Text1(Index), strTempName) Then
            Label2.Caption = strTempName
         Else
            Label2.Caption = ""
            Cancel = True
         End If
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub
