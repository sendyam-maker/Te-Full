VERSION 5.00
Begin VB.Form frm083005 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員收文明細表"
   ClientHeight    =   3735
   ClientLeft      =   2700
   ClientTop       =   780
   ClientWidth     =   4395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4395
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   10
      Left            =   2424
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2964
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   9
      Left            =   1344
      MaxLength       =   9
      TabIndex        =   9
      Top             =   2964
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   8
      Left            =   2424
      MaxLength       =   9
      TabIndex        =   8
      Top             =   2604
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   7
      Left            =   1344
      MaxLength       =   9
      TabIndex        =   7
      Top             =   2604
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   6
      Left            =   2424
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2244
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   1344
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2244
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   2424
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1884
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1344
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1884
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1344
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1524
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1344
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1164
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1344
      TabIndex        =   0
      Top             =   804
      Width           =   2535
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3312
      TabIndex        =   12
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2484
      TabIndex        =   11
      Top             =   120
      Width           =   800
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   3
      X1              =   2184
      X2              =   2304
      Y1              =   3084
      Y2              =   3084
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   2
      X1              =   2184
      X2              =   2304
      Y1              =   2724
      Y2              =   2724
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   252
      Index           =   1
      Left            =   2424
      TabIndex        =   21
      Top             =   1524
      Width           =   1464
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   276
      Index           =   0
      Left            =   2424
      TabIndex        =   20
      Top             =   1164
      Width           =   1464
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   2184
      X2              =   2304
      Y1              =   2364
      Y2              =   2364
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代  理  人："
      Height          =   180
      Index           =   6
      Left            =   384
      TabIndex        =   19
      Top             =   2964
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申  請  人："
      Height          =   180
      Index           =   5
      Left            =   384
      TabIndex        =   18
      Top             =   2604
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   4
      Left            =   384
      TabIndex        =   17
      Top             =   2244
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業  務  區："
      Height          =   180
      Index           =   1
      Left            =   384
      TabIndex        =   16
      Top             =   1164
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   384
      TabIndex        =   15
      Top             =   804
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   2184
      X2              =   2304
      Y1              =   2004
      Y2              =   2004
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文日期："
      Height          =   180
      Index           =   3
      Left            =   384
      TabIndex        =   14
      Top             =   1884
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   2
      Left            =   384
      TabIndex        =   13
      Top             =   1524
      Width           =   900
   End
End
Attribute VB_Name = "frm083005"
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
   If ChkRange(Text1(3), Text1(4), "收文日期") = False Then Exit Sub
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
   
   If Text1(5) <> "" And Text1(6) <> "" Then
      If ChkRange(Text1(5), Text1(6), "案件性質") = False Then Exit Sub
   End If
   If Text1(7) <> "" And Text1(8) <> "" Then
      If ChkRange(Text1(7), Text1(8), "申請人編號") = False Then Exit Sub
   End If
   If Text1(9) <> "" And Text1(10) <> "" Then
      If ChkRange(Text1(9), Text1(10), "代理人編號") = False Then Exit Sub
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

   strExc(0) = "SELECT A0902,ST02," & SQLDate("CP05") & "," & _
      "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04)," & _
      "NVL(LC05,NVL(LC06, lC07)),CPM03,NA03," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP13,CP12,CP05,CP01,CP02,CP03,CP04 " & _
      "FROM STAFF,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,CUSTOMER,ACC090,NATION WHERE " & _
      "CP13=ST01(+) AND CP12=A0901(+) AND (CP01=CPM01(+) AND CP10=CPM02(+)) AND " & _
      "CU10=NA01(+) AND CP57 IS NULL AND (SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+)) AND " & _
      "CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 " & strGetcdnSQL
   If Text1(10).Text <> "" Then
      strExc(0) = strExc(0) & " UNION "
      strExc(0) = strExc(0) & "SELECT A0902,ST02," & SQLDate("CP05") & "," & _
      "CP01||'-'||CP02||DECODE(CP03,'0','','-'||CP03)||DECODE(CP04,'00','','-'||CP04)," & _
      "NVL(LC05,NVL(LC06, lC07)),CPM03,NA03," & SQLDate("CP06") & "," & SQLDate("CP27") & ",CP13,CP12,CP05,CP01,CP02,CP03,CP04 " & _
      "FROM STAFF,CASEPROGRESS,LAWCASE,CASEPROPERTYMAP,CUSTOMER,ACC090,NATION WHERE " & _
      "CP13=ST01(+) AND CP12=A0901(+) AND (CP01=CPM01(+) AND CP10=CPM02(+)) AND " & _
      "CU10=NA01(+) AND CP57 IS NULL AND (SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+)) AND " & _
      "CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 " & strGetcdnSQL1
   
   End If
       strExc(0) = strExc(0) & " ORDER BY CP12,CP13,CP05,CP01,CP02,CP03,CP04"
   If RsTemp.State = adStateOpen Then RsTemp.Close
   RsTemp.Open strExc(0), cnnConnection
   If RsTemp.EOF And RsTemp.BOF Then
      MsgBox "資料庫內無資料 !", vbInformation
      m_print = 1
      Exit Sub
   End If
   i = 1
   Page = 1
 '  Dialog1.ShowPrinter
   If IsNull(RsTemp.Fields(0).Value) = False Then
      TmpArea = RsTemp.Fields(0).Value
   Else
      TmpArea = ""
   End If
   CaseTitle TmpArea, 1
   iPrint = 2700
   With RsTemp
   Do While Not .EOF
      Printer.CurrentX = PLeft(0):      Printer.CurrentY = iPrint
      Printer.Print Format(.Fields(1), "!@@@@")
      Printer.CurrentX = PLeft(1):    Printer.CurrentY = iPrint
      Printer.Print .Fields(2)
      Printer.CurrentX = PLeft(2):    Printer.CurrentY = iPrint
      Printer.Print .Fields(3)
      Printer.CurrentX = PLeft(3):    Printer.CurrentY = iPrint
      Printer.Print Format(.Fields(4), "!@@@@@@@@@@@@")
      Printer.CurrentX = PLeft(4):    Printer.CurrentY = iPrint
      Printer.Print Format(.Fields(5), "!@@@@@@")
      Printer.CurrentX = PLeft(5):    Printer.CurrentY = iPrint
      Printer.Print Format(.Fields(6), "!@@@@@@")
      Printer.CurrentX = PLeft(6):    Printer.CurrentY = iPrint
      If .Fields(7) <> "//" Then
         St = .Fields(7)
      Else
         St = ""
      End If
      Printer.Print St
      Printer.CurrentX = PLeft(7):    Printer.CurrentY = iPrint
      If .Fields(8) <> "//" Then
         St = .Fields(8)
      Else
         St = ""
      End If
      Printer.Print St
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
         If (i Mod 27 = 0) Or (TmpArea <> St) Then
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

Private Sub CaseTitle(ByVal Area As String, ByVal Page As String)
 Dim i As Integer
   i = 500
   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000:         Printer.CurrentY = i
   Printer.Print "智權人員收文明細表"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = 6200:         Printer.CurrentY = i + 500
   Printer.Print "收文日 : " & ChangeTStringToTDateString(Text1(3)) & _
      " - " & ChangeTStringToTDateString(Text1(4))
   Printer.Font.Bold = False
   Printer.CurrentX = 500:               Printer.CurrentY = i + 800
   Printer.Print "列印人 : " & strUserName
   Printer.CurrentX = 13000:             Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
   Printer.CurrentX = 500:               Printer.CurrentY = i + 1100
   Printer.Print "業務區 : " & Area
   Printer.CurrentX = 13000:             Printer.CurrentY = i + 1100
   Printer.Print "頁次 : " & Page
   Printer.CurrentX = 500:               Printer.CurrentY = i + 1400
   Printer.Print String(200, "-")
   Printer.CurrentX = PLeft(0):          Printer.CurrentY = i + 1700
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(1):          Printer.CurrentY = i + 1700
   Printer.Print "收文日"
   Printer.CurrentX = PLeft(2):          Printer.CurrentY = i + 1700
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(3):          Printer.CurrentY = i + 1700
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(4):          Printer.CurrentY = i + 1700
   Printer.Print "案件性質"
   Printer.CurrentX = PLeft(5):          Printer.CurrentY = i + 1700
   Printer.Print "客戶國籍"
   Printer.CurrentX = PLeft(6):          Printer.CurrentY = i + 1700
   Printer.Print "本所期限"
   Printer.CurrentX = PLeft(7):          Printer.CurrentY = i + 1700
   Printer.Print "發文日"
   Printer.CurrentX = 500:          Printer.CurrentY = i + 2000
   Printer.Print String(200, "-")
End Sub

Private Sub GetPrintLeft()
   Erase PLeft
   PLeft(0) = 500:     PLeft(1) = 1400
   PLeft(2) = 2500:    PLeft(3) = 4300
   PLeft(4) = 7500:    PLeft(5) = 9800
   PLeft(6) = 11300:    PLeft(7) = 13000
End Sub

Private Sub Form_Activate()
   Text1(0).SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Text1(0).Text = GetSystemKindByNick
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 7, 8, 9, 10
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTempName As String, i As Integer, t As Integer
 Dim s As Integer
 Dim j As Integer
 Dim strTemp1 As Variant
 Dim strTemp2 As Variant
 
   If Text1(Index) = "" Then
      If Index = 1 Or Index = 2 Then
         Label2(Index - 1).Caption = ""
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

      Case 1
         'edit by nickc 2007/02/07 不用 dll 了
         'If objLawDll.GetStaffDeptName(Text1(Index), strTempName) Then
         If ClsPDGetStaffDeptName(Text1(Index), strTempName) Then
            Label2(0) = strTempName
         Else
            Cancel = True
            Label2(0) = ""
         End If
      Case 2
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(Text1(Index), strTempName) Then
         If ClsPDGetStaffN(Text1(Index), strTempName) Then
            Label2(1) = strTempName
         Else
            Label2(1) = ""
            Cancel = True
         End If
      Case 3, 4
            If CheckIsTaiwanDate(Text1(Index)) = False Then Cancel = True
      Case 5, 6
          If Text1(5).Text <> "" And Text1(6).Text <> "" Then
             If CInt(Text1(5).Text) > CInt(Text1(6).Text) Then
                MsgBox "案件性質代碼起不可大於迄!", vbExclamation, "智權人員收文明細表"
                Cancel = True
             End If
          End If
               
      Case 7, 8
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCustomer(Text1(Index), strTempName) = False Then
         If ClsPDGetCustomer(Text1(Index), strTempName) = False Then
            Cancel = True
         Else
            If Index = 8 Then
               If Left(Text1(7), 6) <> Left(Text1(8), 6) Then
                  MsgBox "申請人起迄號前六碼必須相同 !", vbCritical
                  Cancel = True
               End If
            End If
         End If
      Case 9, 10
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetAgent(Text1(Index), strTempName) = False Then
         If ClsPDGetAgent(Text1(Index), strTempName) = False Then
            Cancel = True
         Else
            If Index = 10 Then
               If Left(Text1(9), 6) <> Left(Text1(10), 6) Then
                  MsgBox "代理人起迄號前六碼必須相同 !", vbCritical
                  Cancel = True
               End If
            End If
         End If
      End Select
      If Cancel Then TextInverse Text1(Index)
End Sub

Private Function strGetcdnSQL() As String
 Dim i As Integer
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
                                     
                      
   strExc(1) = " AND CP01 in('" & strSql & "')"
   If Text1(3) = "" And Text1(4) <> "" Then
      strExc(1) = strExc(1) & " AND CP05 <='" & ChangeTStringToWString(Text1(4)) + "'"
   ElseIf Text1(3) <> "" And Text1(4) <> "" Then
      strExc(1) = strExc(1) & " AND CP05 BETWEEN '" & _
         ChangeTStringToWString(Text1(3)) + "' AND '" + ChangeTStringToWString(Text1(4)) + "'"
   End If
   If Text1(1) <> "" Then strExc(1) = strExc(1) & " and CP12='" + Text1(1) + "'"
   If Text1(2) <> "" Then strExc(1) = strExc(1) & " and CP13='" + Text1(2) + "'"
   If Text1(5) = "" And Text1(6) <> "" Then
      strExc(1) = strExc(1) + " AND CP10<='" & Text1(5) & "'"
   ElseIf Text1(5) <> "" And Text1(6) <> "" Then
      strExc(1) = strExc(1) + " AND (CP10 BETWEEN '" & Text1(5) & "' AND '" & Text1(6) & "')"
   End If
   If Text1(7) = "" And Text1(8) <> "" Then
      strExc(1) = strExc(1) + " AND LC11<='" & GetNewFagent(Text1(7)) & "'"
   ElseIf Text1(7) <> "" And Text1(8) <> "" Then
      strExc(1) = strExc(1) + " AND (LC11 BETWEEN '" & GetNewFagent(Text1(7)) & "' AND '" & GetNewFagent(Text1(8)) & "')"
   End If
   If Text1(9) = "" And Text1(10) <> "" Then
      strExc(1) = strExc(1) + " AND LC22<='" & GetNewFagent(Text1(9)) & "'"
   ElseIf Text1(9) <> "" And Text1(10) <> "" Then
      strExc(1) = strExc(1) + " AND (LC22 BETWEEN '" & GetNewFagent(Text1(9)) & "' AND '" & GetNewFagent(Text1(10)) & "')"
   End If
   strGetcdnSQL = strExc(1) & " AND CP57 IS NULL AND CP09<'C'"
End Function
Private Function strGetcdnSQL1() As String
 Dim i As Integer
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
                                     
                      
   strExc(1) = " AND CP01 in('" & strSql & "')"
   If Text1(3) = "" And Text1(4) <> "" Then
      strExc(1) = strExc(1) & " AND CP05 <='" & ChangeTStringToWString(Text1(4)) + "'"
   ElseIf Text1(3) <> "" And Text1(4) <> "" Then
      strExc(1) = strExc(1) & " AND CP05 BETWEEN '" & _
         ChangeTStringToWString(Text1(3)) + "' AND '" + ChangeTStringToWString(Text1(4)) + "'"
   End If
   If Text1(1) <> "" Then strExc(1) = strExc(1) & " and CP12='" + Text1(1) + "'"
   If Text1(2) <> "" Then strExc(1) = strExc(1) & " and CP13='" + Text1(2) + "'"
   If Text1(5) = "" And Text1(6) <> "" Then
      strExc(1) = strExc(1) + " AND CP10<='" & Text1(5) & "'"
   ElseIf Text1(5) <> "" And Text1(6) <> "" Then
      strExc(1) = strExc(1) + " AND (CP10 BETWEEN '" & Text1(5) & "' AND '" & Text1(6) & "')"
   End If
   If Text1(7) = "" And Text1(8) <> "" Then
      strExc(1) = strExc(1) + " AND LC11<='" & GetNewFagent(Text1(7)) & "'"
   ElseIf Text1(7) <> "" And Text1(8) <> "" Then
      strExc(1) = strExc(1) + " AND (LC11 BETWEEN '" & GetNewFagent(Text1(7)) & "' AND '" & GetNewFagent(Text1(8)) & "')"
   End If
   If Text1(9) = "" And Text1(10) <> "" Then
      strExc(1) = strExc(1) + " AND CP44<='" & GetNewFagent(Text1(9)) & "'"
   ElseIf Text1(9) <> "" And Text1(10) <> "" Then
      strExc(1) = strExc(1) + " AND (CP44 BETWEEN '" & GetNewFagent(Text1(9)) & "' AND '" & GetNewFagent(Text1(10)) & "')"
   End If
   strGetcdnSQL1 = strExc(1) & " AND CP57 IS NULL AND CP09<'C'"
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm083005 = Nothing
End Sub

