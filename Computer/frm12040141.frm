VERSION 5.00
Begin VB.Form frm12040141 
   BorderStyle     =   1  '單線固定
   Caption         =   "新客戶清單"
   ClientHeight    =   3960
   ClientLeft      =   2955
   ClientTop       =   1620
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4935
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1110
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "1"
      Top             =   2460
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1110
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "1"
      Top             =   810
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1590
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "5"
      Top             =   1470
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   10
      Left            =   1110
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "1"
      Top             =   2130
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   9
      Left            =   1110
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3996
      TabIndex        =   9
      Top             =   75
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3168
      TabIndex        =   8
      Top             =   75
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   6
      Left            =   2670
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1140
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   1110
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1140
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1110
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "2"
      Top             =   1470
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "輸出方式                    (1:E-MAIL 2:報表)"
      Height          =   180
      Index           =   5
      Left            =   270
      TabIndex        =   18
      Top             =   2490
      Width           =   3630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "       2. 每月依全所, 依智權人員跳頁"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   300
      TabIndex        =   17
      Top             =   3630
      Width           =   2745
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS : 1. 每週只印分所, 依所別跳頁(已取消)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   2
      Left            =   300
      TabIndex        =   16
      Top             =   3330
      Width           =   3225
   End
   Begin VB.Label Label1 
      Caption         =   "列印別                        (1: 每週 2: 每月)"
      Height          =   180
      Index           =   1
      Left            =   270
      TabIndex        =   15
      Top             =   840
      Width           =   3630
   End
   Begin VB.Label Label1 
      Caption         =   "國內外                        (1: 國內 2: 國外)"
      Height          =   180
      Index           =   8
      Left            =   270
      TabIndex        =   14
      Top             =   2160
      Width           =   3630
   End
   Begin VB.Label lblName 
      Height          =   180
      Left            =   2400
      TabIndex        =   13
      Top             =   1860
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員"
      Height          =   180
      Index           =   6
      Left            =   270
      TabIndex        =   12
      Top             =   1845
      Width           =   720
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2430
      X2              =   2550
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "開發日期"
      Height          =   180
      Index           =   3
      Left            =   270
      TabIndex        =   11
      Top             =   1140
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所　　別         －        (1: 北所 2: 中所 3: 南所 4: 高所 5: 其他)"
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   10
      Top             =   1485
      Width           =   4665
   End
End
Attribute VB_Name = "frm12040141"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim intWhere As Integer, strReceiveNo As String, PLeft(0 To 10) As Integer
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
Dim TempFileName As String
Dim ff As Integer
Dim A01 As String, A02 As String, A03 As String, A04 As String, A05 As String

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0 '確定
    '檢查輸入的資料是否齊全完整
    If CheckDataValid() = False Then
        GoTo EXITSUB
    End If
    strSql = ""
    Screen.MousePointer = vbHourglass
    Process
    Screen.MousePointer = vbDefault
Case 1 '結束
    Unload Me
End Select
EXITSUB:
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    Text1(0) = "3"    '92.7.14 ADD BY SONIA
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm12040141 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    Select Case Index
    Case 0, 1 '所別
        If (KeyAscii > 53 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
    Case 2, 10, 3 '國內外
        If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
    End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 1, 6
         If blnClkSure = False Then
            If Text1(Index - 1) <> "" Then
               If RunNick(Text1(Index - 1), Text1(Index)) Then
                 Text1(Index - 1).SetFocus
               End If
            End If
         Else
            blnClkSure = False
         End If
      Case 2
         If Text1(Index) = "1" Then
            '92.7.14 ADD BY SONIA
            'Text1(0) = "2"
            Text1(0) = "3"
            '92.7.14 END
         Else
            Text1(0) = "1"
            '2013/5/16 ADD BY SONIA strSrvDate(1)
            Text1(5) = TransDate(CompDate(1, -1, (Left(strSrvDate(1), 6) & "01")), 1)   '預設上月1日
            Text1(6) = TransDate(CompDate(2, -1, (Left(strSrvDate(1), 6) & "01")), 1) '預設上月最後一日
            '2013/5/16 END

         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTmp As String
   Select Case Index
      Case 2, 10, 3
         'Add By Sindy 2012/10/4
         If Text1(Index) <> "1" And Text1(Index) <> "2" Then
            Cancel = True
            MsgBox "只可輸入1或2 !!!"
         End If
         '2012/10/4 End
      Case 5, 6
         If Text1(Index) <> "" Then
            Cancel = Not ChkDate(Text1(Index).Text)
         End If
      Case 9
         lblName.Caption = ""
         If Text1(Index) <> "" Then
            'edit by nickc 2007/02/09 不用 dll 了
            'If Not objPublicData.GetStaff(Text1(Index), strExc(0)) Then
            If Not ClsPDGetStaff(Text1(Index), strExc(0)) Then
               Cancel = True
            Else
               lblName.Caption = strExc(0)
            End If
         End If
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub

Private Sub Process()
Dim i As Integer, j As Integer, Page As Integer, iPrint As Integer
Dim strTmp As String, strTmp1 As String
Dim strNo As String '所別
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   On Error GoTo ErrHand
   '開發日期區間
   If Text1(5).Text <> "" And Text1(6).Text <> "" Then
       strSql = strSql & " AND CU14 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2)
   ElseIf Text1(5).Text = "" And Text1(6).Text <> "" Then
       strSql = strSql & " AND CU14 <=" & TransDate(Text1(6).Text, 2)
   ElseIf Text1(5).Text <> "" And Text1(6).Text = "" Then
       strSql = strSql & " AND CU14 >=" & TransDate(Text1(5).Text, 2) & _
                   " AND CU14 <=" & ServerDate & " "
   End If
   '所別
   If Me.Text1(0).Text <> "" Then
       strSql = strSql & " And ST06>='" & Me.Text1(0).Text & "' "
   End If
   If Me.Text1(1).Text <> "" Then
       strSql = strSql & " And ST06<='" & Me.Text1(1).Text & "' "
   End If
   '智權人員
   If Len(Me.Text1(9).Text) > 0 Then
       strSql = strSql + " AND CU13='" & Me.Text1(9).Text & "' "
   End If
   '國內外
   If Me.Text1(10).Text = "1" Then
       strSql = strSql + " AND (ST15<'F' OR ST15>'F99') "
   ElseIf Me.Text1(10).Text = "2" Then
       strSql = strSql + " AND ST15>='F' AND ST15<='F99' "
   Else
      '無動作
   End If
   strExc(0) = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,CU30,SUBSTR(NVL(CU31,CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,ST06,DECODE(ST06,'1','北所','2','中所','3','南所','4','高所','其他') FROM CUSTOMER,NATION,STAFF Where  CU10=NA01(+) AND CU13=ST01(+) " & strSql
   If Me.Text1(2) = "2" Then
      strExc(0) = strExc(0) + " AND CU12>='S' AND CU12<='S99' ORDER BY CU12,CU13,CU01,CU02 "
   Else
      'Modify By Sindy 2012/10/4
      If Me.Text1(3) = "1" Then 'E-Mail
         strExc(0) = strExc(0) + " ORDER BY CU12,CU13,CU01,CU02 "
      Else
      '2012/10/4 End
         strExc(0) = strExc(0) + " ORDER BY ST06,CU01,CU02 "
      End If
   End If
   intI = 0
   'edit by nickc 2007/02/09 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modify By Sindy 2012/10/4
      If Me.Text1(3) = "1" Then 'E-Mail
         PrintCase3
      Else
      '2012/10/4 End
         If Me.Text1(2) = "1" Then
            PrintCase1
         Else
            PrintCase2
         End If
      End If
   End If
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub PrintCase1()
Dim i As Integer, j As Integer, Page As Integer, iPrint As Integer
Dim strTmp As String, strTmp1 As String
Dim strNo As String '所別
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   GetPrintLeft
   Page = 1
   CaseTitle Page, "" & RsTemp.Fields(16).Value
   iPrint = 2700 + 300 + 300
   strNo = RsTemp.Fields(16).Value
   i = 0
   With RsTemp
      i = 0
      Do While Not .EOF
         For j = 0 To 4
            Printer.CurrentX = PLeft(j)
            Printer.CurrentY = iPrint
            If j = 0 Then
               '2007/9/3 MODIFY BY SONIA
               'If Not IsNull(RsTemp.Fields(8).Value) Or Not IsNull(RsTemp.Fields(9).Value) Then
               If Not IsNull(RsTemp.Fields(9).Value) Then
                  Printer.CurrentX = PLeft(j) - Printer.TextWidth("＊")
                  '2007/9/3 MODIFY BY SONIA
                  'Printer.Print "＊" & Left(.Fields(j) & "000", 9)
                  If Not IsNull(RsTemp.Fields(8).Value) Then
                     Printer.Print "＊" & Left(.Fields(j) & "000", 9) & " N"
                  Else
                     Printer.Print "＊" & Left(.Fields(j) & "000", 9)
                  End If
                  '2007/9/3 END
               Else
                  '2007/9/3 MODIFY BY SONIA
                  'Printer.Print "" & Left(.Fields(j) & "000", 9)
                  If Not IsNull(RsTemp.Fields(8).Value) Then
                     Printer.Print " " & Left(.Fields(j) & "000", 9) & " N"
                  Else
                     Printer.Print "" & Left(.Fields(j) & "000", 9)
                  End If
                  '2007/9/3 END
               End If
            'Add By Cheng 2003/02/20
            '若列印負責人
            ElseIf j = 2 Then
               Printer.Print Left("" & .Fields(j), 7)
            Else
               Printer.Print "" & .Fields(j)
            End If
         Next j
         
         iPrint = iPrint + 300
         
         For j = 5 To 7
            Printer.CurrentX = PLeft(j)
            Printer.CurrentY = iPrint
            Printer.Print "" & .Fields(j)
         Next j
         
         iPrint = iPrint + 300
                     
         Printer.CurrentX = PLeft(8)
         Printer.CurrentY = iPrint
         Printer.Print "" & .Fields(13)
         Printer.CurrentX = PLeft(9)
         Printer.CurrentY = iPrint
         Printer.Print "" & .Fields(11)
         Printer.CurrentX = PLeft(10)
         Printer.CurrentY = iPrint
         Printer.Print "" & .Fields(14)
         
         iPrint = iPrint + 300
         i = i + 1
         
         If i > 9 Or "" & RsTemp.Fields(16).Value <> strNo Then
            
             Printer.CurrentX = PLeft(0)
             Printer.CurrentY = iPrint
             Printer.Print String(250, "-")
             
             iPrint = iPrint + 300
            
            strNo = "" & RsTemp.Fields(16).Value
            '2007/9/3 ADD BY SONIA
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
            '2007/9/3 END
            Printer.NewPage
            Page = Page + 1
            CaseTitle Page, "" & RsTemp.Fields(16).Value
             iPrint = 2700 + 300 + 300
            i = 0
         End If
         
         '若為子公司
         If Right("" & .Fields(0).Value, 3) <> "000" Then
             StrSQLa = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,CU30,SUBSTR(NVL(CU31,CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,ST06,DECODE(ST06,'1','北所','2','中所','3','南所','4','高所','其他') FROM CUSTOMER,NATION,STAFF Where  CU10=NA01(+) AND CU13=ST01(+) AND CU01='" & Left(.Fields(0).Value, 6) & "00" & "' AND CU02='0' "
             rsA.CursorLocation = adUseClient
             rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
             If rsA.RecordCount > 0 Then
                 Printer.CurrentX = PLeft(0)
                 Printer.CurrentY = iPrint
                 Printer.Print "母公司資料："
                 iPrint = iPrint + 300
                                     
                 For j = 0 To 4
                    Printer.CurrentX = PLeft(j)
                    Printer.CurrentY = iPrint
                    If j = 0 Then
                       '2007/9/3 MODIFY BY SONIA
                       'If Not IsNull(rsA.Fields(8).Value) Or Not IsNull(rsA.Fields(9).Value) Then
                       If Not IsNull(rsA.Fields(9).Value) Then
                          Printer.CurrentX = PLeft(j) - Printer.TextWidth("＊")
                          '2007/9/3 MODIFY BY SONIA
                          'Printer.Print "＊" & Left(rsA.Fields(j) & "000", 9)
                          If Not IsNull(rsA.Fields(8).Value) Then
                             Printer.Print "＊" & Left(rsA.Fields(j) & "000", 9) & " N"
                          Else
                             Printer.Print "＊" & Left(rsA.Fields(j) & "000", 9)
                          End If
                          '2007/9/3 END
                       Else
                          '2007/9/3 MODIFY BY SONIA
                          'Printer.Print "" & Left(rsA.Fields(j) & "000", 9)
                          If Not IsNull(rsA.Fields(8).Value) Then
                             Printer.Print " " & Left(rsA.Fields(j) & "000", 9) & " N"
                          Else
                             Printer.Print "" & Left(rsA.Fields(j) & "000", 9)
                          End If
                          '2007/9/3 END
                       End If
                     'Add By Cheng 2003/02/20
                     '若列印負責人
                     ElseIf j = 2 Then
                        Printer.Print Left("" & rsA.Fields(j), 7)
                    Else
                       Printer.Print "" & rsA.Fields(j)
                    End If
                 Next j
                 
                 iPrint = iPrint + 300
                 
                 For j = 5 To 7
                    Printer.CurrentX = PLeft(j)
                    Printer.CurrentY = iPrint
                    Printer.Print "" & rsA.Fields(j)
                 Next j
                 
                 iPrint = iPrint + 300
                             
                 Printer.CurrentX = PLeft(8)
                 Printer.CurrentY = iPrint
                 Printer.Print "" & rsA.Fields(13)
                 Printer.CurrentX = PLeft(9)
                 Printer.CurrentY = iPrint
                 Printer.Print "" & rsA.Fields(11)
                 Printer.CurrentX = PLeft(10)
                 Printer.CurrentY = iPrint
                 Printer.Print "" & rsA.Fields(14)
                 
                 iPrint = iPrint + 300
                 i = i + 1
             End If
             If rsA.State <> adStateClosed Then rsA.Close
             Set rsA = Nothing
         End If
         
         If i <> 0 Then
             Printer.CurrentX = PLeft(0)
             Printer.CurrentY = iPrint
             Printer.Print String(250, "-")
             iPrint = iPrint + 300
         End If
                     
         .MoveNext
         If RsTemp.EOF Then Exit Do
         If i > 9 Or "" & RsTemp.Fields(16).Value <> strNo Then
            strNo = "" & RsTemp.Fields(16).Value
            '2007/9/3 ADD BY SONIA
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
            '2007/9/3 END
            Printer.NewPage
            Page = Page + 1
            CaseTitle Page, "" & RsTemp.Fields(16).Value
             iPrint = 2700 + 300 + 300
            i = 0
         End If
      Loop
   End With
   '2007/9/3 ADD BY SONIA
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
   '2007/9/3 END
   Printer.EndDoc
   ShowPrintOk
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub PrintCase2()
Dim i As Integer, j As Integer, Page As Integer, iPrint As Integer
Dim strTmp As String, strTmp1 As String
Dim strNo As String '所別
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   GetPrintLeft
   Page = 1
   CaseTitle Page, "" & RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
   iPrint = 2700 + 300 + 300
   strNo = RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
   i = 0
   With RsTemp
      i = 0
      Do While Not .EOF
         For j = 0 To 4
            Printer.CurrentX = PLeft(j)
            Printer.CurrentY = iPrint
            If j = 0 Then
               '2007/9/3 MODIFY BY SONIA
               'If Not IsNull(RsTemp.Fields(8).Value) Or Not IsNull(RsTemp.Fields(9).Value) Then
               If Not IsNull(RsTemp.Fields(9).Value) Then
                  Printer.CurrentX = PLeft(j) - Printer.TextWidth("＊")
                  '2007/9/3 MODIFY BY SONIA
                  'Printer.Print "＊" & Left(.Fields(j) & "000", 9)
                  If Not IsNull(RsTemp.Fields(8).Value) Then
                     Printer.Print "＊" & Left(.Fields(j) & "000", 9) & " N"
                  Else
                     Printer.Print "＊" & Left(.Fields(j) & "000", 9)
                  End If
                  '2007/9/3 END
               Else
                  '2007/9/3 MODIFY BY SONIA
                  'Printer.Print "" & Left(.Fields(j) & "000", 9)
                  If Not IsNull(RsTemp.Fields(8).Value) Then
                     Printer.Print " " & Left(.Fields(j) & "000", 9) & " N"
                  Else
                     Printer.Print "" & Left(.Fields(j) & "000", 9)
                  End If
                  '2007/9/3 END
               End If
            'Add By Cheng 2003/02/20
            '若列印負責人
            ElseIf j = 2 Then
               Printer.Print Left("" & .Fields(j), 7)
            Else
               Printer.Print "" & .Fields(j)
            End If
         Next j
         
         iPrint = iPrint + 300
         
         For j = 5 To 7
            Printer.CurrentX = PLeft(j)
            Printer.CurrentY = iPrint
            Printer.Print "" & .Fields(j)
         Next j
         
         iPrint = iPrint + 300
                     
         Printer.CurrentX = PLeft(8)
         Printer.CurrentY = iPrint
         Printer.Print "" & .Fields(13)
         Printer.CurrentX = PLeft(9)
         Printer.CurrentY = iPrint
         Printer.Print "" & .Fields(11)
         Printer.CurrentX = PLeft(10)
         Printer.CurrentY = iPrint
         Printer.Print "" & .Fields(14)
         
         iPrint = iPrint + 300
         i = i + 1
         
         If i > 9 Or RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value <> strNo Then
            
             Printer.CurrentX = PLeft(0)
             Printer.CurrentY = iPrint
             Printer.Print String(250, "-")
             
             iPrint = iPrint + 300
            
            strNo = RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
            '2007/9/3 ADD BY SONIA
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
            '2007/9/3 END
            Printer.NewPage
            Page = Page + 1
            CaseTitle Page, "" & RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
             iPrint = 2700 + 300 + 300
            i = 0
         End If
         
         '若為子公司
         If Right("" & .Fields(0).Value, 3) <> "000" Then
             StrSQLa = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,CU30,SUBSTR(NVL(CU31,CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,ST06,DECODE(ST06,'1','北所','2','中所','3','南所','4','高所','其他') FROM CUSTOMER,NATION,STAFF Where  CU10=NA01(+) AND CU13=ST01(+) AND CU01='" & Left(.Fields(0).Value, 6) & "00" & "' AND CU02='0' "
             rsA.CursorLocation = adUseClient
             rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
             If rsA.RecordCount > 0 Then
                 Printer.CurrentX = PLeft(0)
                 Printer.CurrentY = iPrint
                 Printer.Print "母公司資料："
                 iPrint = iPrint + 300
                                     
                 For j = 0 To 4
                    Printer.CurrentX = PLeft(j)
                    Printer.CurrentY = iPrint
                    If j = 0 Then
                       '2007/9/3 MODIFY BY SONIA
                       'If Not IsNull(rsA.Fields(8).Value) Or Not IsNull(rsA.Fields(9).Value) Then
                       If Not IsNull(rsA.Fields(9).Value) Then
                          Printer.CurrentX = PLeft(j) - Printer.TextWidth("＊")
                          '2007/9/3 MODIFY BY SONIA
                          'Printer.Print "＊" & Left(rsA.Fields(j) & "000", 9)
                          If Not IsNull(rsA.Fields(8).Value) Then
                             Printer.Print "＊" & Left(rsA.Fields(j) & "000", 9) & " N"
                          Else
                             Printer.Print "＊" & Left(rsA.Fields(j) & "000", 9)
                          End If
                          '2007/9/3 END
                       Else
                          '2007/9/3 MODIFY BY SONIA
                          'Printer.Print "" & Left(rsA.Fields(j) & "000", 9)
                          If Not IsNull(rsA.Fields(8).Value) Then
                             Printer.Print " " & Left(rsA.Fields(j) & "000", 9) & " N"
                          Else
                             Printer.Print "" & Left(rsA.Fields(j) & "000", 9)
                          End If
                          '2007/9/3 END
                       End If
                     'Add By Cheng 2003/02/20
                     '若列印負責人
                     ElseIf j = 2 Then
                        Printer.Print Left("" & rsA.Fields(j), 7)
                    Else
                       Printer.Print "" & rsA.Fields(j)
                    End If
                 Next j
                 
                 iPrint = iPrint + 300
                 
                 For j = 5 To 7
                    Printer.CurrentX = PLeft(j)
                    Printer.CurrentY = iPrint
                    Printer.Print "" & rsA.Fields(j)
                 Next j
                 
                 iPrint = iPrint + 300
                             
                 Printer.CurrentX = PLeft(8)
                 Printer.CurrentY = iPrint
                 Printer.Print "" & rsA.Fields(13)
                 Printer.CurrentX = PLeft(9)
                 Printer.CurrentY = iPrint
                 Printer.Print "" & rsA.Fields(11)
                 Printer.CurrentX = PLeft(10)
                 Printer.CurrentY = iPrint
                 Printer.Print "" & rsA.Fields(14)
                 
                 iPrint = iPrint + 300
                 i = i + 1
             End If
             If rsA.State <> adStateClosed Then rsA.Close
             Set rsA = Nothing
         End If
         
         If i <> 0 Then
             Printer.CurrentX = PLeft(0)
             Printer.CurrentY = iPrint
             Printer.Print String(250, "-")
             iPrint = iPrint + 300
         End If
                     
         .MoveNext
         If RsTemp.EOF Then Exit Do
         If i > 9 Or RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value <> strNo Then
            strNo = RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
            '2007/9/3 ADD BY SONIA
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
            '2007/9/3 END
            Printer.NewPage
            Page = Page + 1
            CaseTitle Page, "" & RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
             iPrint = 2700 + 300 + 300
            i = 0
         End If
      Loop
   End With
   '2007/9/3 ADD BY SONIA
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
   '2007/9/3 END
   Printer.EndDoc
   ShowPrintOk
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub GetPrintLeft()
    '第一列
    PLeft(0) = 200
    PLeft(1) = 1500
    PLeft(2) = 4200 + 3000 - 500 - 500
    PLeft(3) = 5200 + 3000 - 500
    PLeft(4) = 6200 + 3000 + 500 - 500
    '第二列
    PLeft(5) = 200
    PLeft(6) = 1500
    PLeft(7) = 6200 + 3000 + 500 - 500
    '第三列
    PLeft(8) = 1500
    PLeft(9) = 5200 + 3000 - 500
    PLeft(10) = 6200 + 3000 + 500 - 500
End Sub

Private Sub CaseTitle(ByVal Page As String, ByVal strSNo As String)
'Page : 頁數
'strSNo : 業務區別+智權人員 或 所別
 Dim i As Integer
   
   i = 500
   If Page = 1 Then Printer.Orientation = vbPRORPortrait
   Printer.FontName = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000
   Printer.CurrentY = i
   Printer.Print "新　客　戶　清　單"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 800 - 300
   Printer.Print "列印人　 : " & strUserName
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 800
   If Me.Text1(2) = "1" Then
      Printer.Print "所別　　 : " & strSNo
   Else
      Printer.Print "智權人員 : " & strSNo
   End If
   Printer.CurrentX = 7000 + 1500
   Printer.CurrentY = i + 800
   Printer.Print "列印日期 : " & ChangeTStringToTDateString("" & (Val(ServerDate) - 19110000))
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 1100
   Printer.Print "開發日期 : " & ChangeTStringToTDateString(Me.Text1(5).Text) & " - " & ChangeTStringToTDateString(Me.Text1(6).Text)
   Printer.CurrentX = 7000 + 1500
   Printer.CurrentY = i + 1100
   Printer.Print "頁　　次 : " & Page
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 1400
   Printer.Print String(250, "-")
   
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = i + 1700
   Printer.Print "編號"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = i + 1700
   Printer.Print "公司名稱"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = i + 1700
   Printer.Print "負責人"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = i + 1700
   Printer.Print "國籍"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = i + 1700
   Printer.Print "電話"
   
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = i + 1700 + 300
   Printer.Print "郵遞區號"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = i + 1700 + 300
   Printer.Print "聯絡地址"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = i + 1700 + 300
   Printer.Print "傳真"
   
    Printer.CurrentX = PLeft(8)
    Printer.CurrentY = i + 1700 + 300 + 300
    Printer.Print "中文地址"
    'Add By Cheng 2003/02/20
    '加標題--智權人員
    Printer.CurrentX = PLeft(9)
    Printer.CurrentY = i + 1700 + 300 + 300
    Printer.Print "智權人員"
    Printer.CurrentX = PLeft(10)
    Printer.CurrentY = i + 1700 + 300 + 300
    Printer.Print "統一編號"
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = i + 2000 + 300 + 300
   Printer.Print String(250, "-")
   Printer.Font.Size = 10
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   blnClkSure = False
   
   '檢查開發日期
   If Len(Me.Text1(5).Text) <= 0 Then
      MsgBox "請輸入開發起日!!!", vbExclamation + vbOKOnly
      Me.Text1(5).SetFocus
      Text1_GotFocus 5
      GoTo EXITSUB
   End If
   If Len(Me.Text1(6).Text) <= 0 Then
      MsgBox "請輸入開發迄日!!!", vbExclamation + vbOKOnly
      Me.Text1(6).SetFocus
      Text1_GotFocus 6
      GoTo EXITSUB
   End If
   If PUB_CheckKeyInDate(Me.Text1(5)) = -1 Then
      Me.Text1(5).SetFocus
      Text1_GotFocus 5
      GoTo EXITSUB
   End If
   If PUB_CheckKeyInDate(Me.Text1(6)) = -1 Then
      Me.Text1(6).SetFocus
      Text1_GotFocus 6
      GoTo EXITSUB
   End If
   If Val("0" & Me.Text1(5).Text) > Val("0" & Me.Text1(6).Text) Then
      MsgBox "開發日期輸入範圍錯誤!!!", vbExclamation + vbOKOnly
      blnClkSure = True
      Me.Text1(5).SetFocus
      Text1_GotFocus 5
      GoTo EXITSUB
   End If
   
   '所別
   If Val(Me.Text1(0).Text) > Val(Me.Text1(1).Text) Then
      MsgBox "所別範圍輸入錯誤!!!", vbExclamation + vbOKOnly
      blnClkSure = True
      Me.Text1(0).SetFocus
      Text1_GotFocus 0
      GoTo EXITSUB
   End If
   
   '檢查智權人員
   lblName.Caption = ""
   If Text1(9) <> "" Then
      'edit by nickc 2007/02/09 不用 dll 了
      'If Not objPublicData.GetStaff(Text1(9), strExc(0)) Then
      If Not ClsPDGetStaff(Text1(9), strExc(0)) Then
         Me.Text1(9).SetFocus
         Text1_GotFocus 9
         GoTo EXITSUB
      Else
         lblName.Caption = strExc(0)
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

'Add By Sindy 2012/10/4
'寄E-Mail時
Private Sub PrintCase3()
Dim iRow As Integer, j As Integer
Dim strNo As String '所別
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strTo As String
Dim s As Integer
   
   GetPrintLeft_Txt
   CaseTitle_Txt "" & RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
   iRow = 0
   strTo = RsTemp.Fields(10).Value
   strNo = RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
   With RsTemp
      Do While Not .EOF
         iRow = iRow + 1
         For j = 0 To 4
            If j = 0 Then
               If Not IsNull(RsTemp.Fields(9).Value) Then
                  If Not IsNull(RsTemp.Fields(8).Value) Then
                     A01 = convForm("＊" & Left(.Fields(j) & "000", 9) & " N", PLeft(0))
                  Else
                     A01 = convForm("＊" & Left(.Fields(j) & "000", 9), PLeft(0))
                  End If
               Else
                  If Not IsNull(RsTemp.Fields(8).Value) Then
                     A01 = convForm(" " & Left(.Fields(j) & "000", 9) & " N", PLeft(0))
                  Else
                     A01 = convForm("" & Left(.Fields(j) & "000", 9), PLeft(0))
                  End If
               End If
            ElseIf j = 1 Then
               A02 = convForm("" & .Fields(j), PLeft(1))
            '若列印負責人
            ElseIf j = 2 Then
               A03 = convForm(Left("" & .Fields(j), 7), PLeft(2))
            ElseIf j = 3 Then
               A04 = convForm("" & .Fields(j), PLeft(3))
            ElseIf j = 4 Then
               A05 = convForm("" & .Fields(j), PLeft(4))
            End If
         Next j
         Print #ff, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
         
         For j = 5 To 7
            If j = 5 Then
               A01 = convForm("" & .Fields(j), PLeft(0))
            ElseIf j = 6 Then
               A02 = convForm("" & .Fields(j), PLeft(1))
               A03 = convForm(" ", PLeft(2))
               A04 = convForm(" ", PLeft(3))
            ElseIf j = 7 Then
               A05 = convForm("" & .Fields(j), PLeft(4))
            End If
         Next j
         Print #ff, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
         
         A01 = convForm(" ", PLeft(0))
         A02 = convForm("" & .Fields(13), PLeft(1))
         A03 = convForm(" ", PLeft(2))
         A04 = convForm("" & .Fields(11), PLeft(3))
         A05 = convForm("" & .Fields(14), PLeft(4))
         Print #ff, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
         
'         If RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value <> strNo Then
'            Call ShowLine
'            Print #ff, "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
'            Close ff
'            Call GoToSendMail(strTo)
'            strTo = RsTemp.Fields(10).Value
'            strNo = RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
'            CaseTitle_Txt "" & RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
'         End If
         
         '若為子公司
         If Right("" & .Fields(0).Value, 3) <> "000" Then
            StrSQLa = "SELECT CU01||CU02,SUBSTR(CU04,1,30),SUBSTR(CU07,1,10),SUBSTR(NA03,1,14),CU16,CU30,SUBSTR(NVL(CU31,CU23),1,65),CU18,CU32,CU80,CU13,ST02,CU12,CU23,CU11,ST06,DECODE(ST06,'1','北所','2','中所','3','南所','4','高所','其他') FROM CUSTOMER,NATION,STAFF Where  CU10=NA01(+) AND CU13=ST01(+) AND CU01='" & Left(.Fields(0).Value, 6) & "00" & "' AND CU02='0' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               Print #ff, "母公司資料："
               For j = 0 To 4
                  If j = 0 Then
                     If Not IsNull(rsA.Fields(9).Value) Then
                        If Not IsNull(rsA.Fields(8).Value) Then
                           A01 = convForm("＊" & Left(rsA.Fields(j) & "000", 9) & " N", PLeft(0))
                        Else
                           A01 = convForm("＊" & Left(rsA.Fields(j) & "000", 9), PLeft(0))
                        End If
                     Else
                        If Not IsNull(rsA.Fields(8).Value) Then
                           A01 = convForm(" " & Left(rsA.Fields(j) & "000", 9) & " N", PLeft(0))
                        Else
                           A01 = convForm("" & Left(rsA.Fields(j) & "000", 9), PLeft(0))
                        End If
                     End If
                  ElseIf j = 1 Then
                     A02 = convForm("" & rsA.Fields(j), PLeft(1))
                  '若列印負責人
                  ElseIf j = 2 Then
                     A03 = convForm(Left("" & rsA.Fields(j), 7), PLeft(2))
                  ElseIf j = 3 Then
                     A04 = convForm("" & rsA.Fields(j), PLeft(3))
                  ElseIf j = 4 Then
                     A05 = convForm("" & rsA.Fields(j), PLeft(4))
                  End If
               Next j
               Print #ff, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
               
               For j = 5 To 7
                  If j = 5 Then
                     A01 = convForm("" & rsA.Fields(j), PLeft(0))
                  ElseIf j = 6 Then
                     A02 = convForm("" & rsA.Fields(j), PLeft(1))
                     A03 = convForm(" ", PLeft(2))
                     A04 = convForm(" ", PLeft(3))
                  ElseIf j = 7 Then
                     A05 = convForm("" & rsA.Fields(j), PLeft(4))
                  End If
               Next j
               Print #ff, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
               
               A01 = convForm(" ", PLeft(0))
               A02 = convForm("" & rsA.Fields(13), PLeft(1))
               A03 = convForm(" ", PLeft(2))
               A04 = convForm("" & rsA.Fields(11), PLeft(3))
               A05 = convForm("" & rsA.Fields(14), PLeft(4))
               Print #ff, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         End If
         Call ShowLine
         
         .MoveNext
         If RsTemp.EOF Then Exit Do
         If RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value <> strNo Then
            Print #ff, "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
            Print #ff, ""
            Print #ff, "共 " & iRow & " 筆"
            Close ff
            Call GoToSendMail(strTo)
            iRow = 0
            strTo = RsTemp.Fields(10).Value
            strNo = RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
            CaseTitle_Txt "" & RsTemp.Fields(12).Value & "　" & RsTemp.Fields(11).Value
         End If
      Loop
   End With
   Print #ff, "PS : 編號與公司名稱之間, 若有 N 表示不寄台一雜誌"
   Print #ff, ""
   Print #ff, "共 " & iRow & " 筆"
   Close ff
   Call GoToSendMail(strTo)
   s = MsgBox("寄信完成!!", , "E-Mail")
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

'Add By Sindy 2012/10/4
Private Sub CaseTitle_Txt(ByVal strSNo As String)
'strSNo : 業務區別+智權人員 或 所別
   TempFileName = ""
   ff = FreeFile
   TempFileName = "新客戶清單"
   If ff > 0 Then Close #ff
   ff = FreeFile
   Open App.path & "\" & TempFileName & ".txt" For Output As ff
   Print #ff, "                                      新　客　戶　清　單                                      "
   Print #ff, "列印人　 : " & strUserName
   Print #ff, "智權人員 : " & strSNo & "　　　　　　　　　　　　　　　　　　　　          列印日期 : " & ChangeTStringToTDateString("" & (Val(ServerDate) - 19110000))
   Print #ff, "開發日期 : " & ChangeTStringToTDateString(Me.Text1(5).Text) & " - " & ChangeTStringToTDateString(Me.Text1(6).Text)
   Call ShowLine
   
   A01 = convForm("編號", PLeft(0))
   A02 = convForm("公司名稱", PLeft(1))
   A03 = convForm("負責人", PLeft(2))
   A04 = convForm("國籍", PLeft(3))
   A05 = convForm("電話", PLeft(4))
   Print #ff, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
   
   A01 = convForm("郵遞區號", PLeft(0))
   A02 = convForm("聯絡地址", PLeft(1))
   A03 = convForm(" ", PLeft(2))
   A04 = convForm(" ", PLeft(3))
   A05 = convForm("傳真", PLeft(4))
   Print #ff, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
   
   A01 = convForm(" ", PLeft(0))
   A02 = convForm("中文地址", PLeft(1))
   A03 = convForm(" ", PLeft(2))
   A04 = convForm("智權人員", PLeft(3))
   A05 = convForm("統一編號", PLeft(4))
   Print #ff, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05
   Call ShowLine
End Sub

Private Sub ShowLine()
   Print #ff, "-----------------------------------------------------------------------------------------------"
End Sub

'Add by Sindy 2012/10/4
'限定字串長度
'Remove by Lydia 2018/08/24 與basQuery重複
'Private Function convForm(ByVal p_InStr As String, ByVal p_Num As Integer, Optional ByVal p_Char As String = " ") As String
'   convForm = StrConv(LeftB(StrConv(p_InStr & String(p_Num, p_Char), vbFromUnicode), p_Num), vbUnicode)
'End Function

'Add by Sindy 2012/10/4
Private Sub GetPrintLeft_Txt()
   PLeft(0) = 12
   PLeft(1) = 42
   PLeft(2) = 12
   PLeft(3) = 12
   PLeft(4) = 13
End Sub

'Add By Sindy 2012/10/4
Private Sub GoToSendMail(ByVal strTo As String)
Dim rsA As New ADODB.Recordset
Dim strTemp As String
   
   '若智權人員離職則改發部門主管(ACC090之A0908)
   strSql = "select st01,st02,st04,st15,a0901,a0908 from staff,acc090 where st01='" & strTo & "' and a0901(+)=st15 "
   Set rsA = New ADODB.Recordset
   rsA.CursorLocation = adUseClient
   rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If rsA.Fields("st04") = "2" Then
         If IsNull(rsA.Fields("a0908")) Then
            strTemp = rsA.Fields("st02")
         Else
            strTo = rsA.Fields("a0908")
         End If
      End If
   Else
      strTemp = strTo
      strTo = ""
   End If
   rsA.Close
   Set rsA = Nothing
   
   If TempFileName <> "" Then
      If strTo = "" Then
         MsgBox strTemp & "已離職且無部門主管，信件無法寄出!!!"
      Else
         'modify by sonia 2016/4/1 收受者請假不彈訊息
         PUB_SendMail strUserNum, strTo, "", ChangeTStringToTDateString(Me.Text1(5).Text) & " - " & ChangeTStringToTDateString(Me.Text1(6).Text) & TempFileName, "Dear Sirs," & vbCrLf & "          " & TempFileName & " 如附件！" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心", , App.path & "\" & TempFileName & ".txt", , , , , , , , , False
      End If
   End If
End Sub
