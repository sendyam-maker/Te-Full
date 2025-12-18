VERSION 5.00
Begin VB.Form frm073011 
   BorderStyle     =   1  '單線固定
   Caption         =   "顧問客戶資料表"
   ClientHeight    =   2460
   ClientLeft      =   1635
   ClientTop       =   2100
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4320
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   210
      TabIndex        =   9
      Top             =   1590
      Width           =   4000
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   4
         Top             =   168
         Width           =   3200
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   10
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   2730
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1140
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1290
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3420
      TabIndex        =   6
      Top             =   75
      Width           =   760
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2736
      MaxLength       =   7
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1296
      MaxLength       =   7
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2580
      TabIndex        =   5
      Top             =   75
      Width           =   800
   End
   Begin VB.Line Line2 
      X1              =   2370
      X2              =   2610
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "業務區 ："
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   8
      Top             =   1140
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   2370
      X2              =   2610
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "到期日期 ："
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frm073011"
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
'Add By Cheng 2002/09/09
Dim blnClkSure As Boolean
'Add By Cheng 2003/03/31
Dim iPrint As Integer
Dim SeekPrint As Integer
Dim SeekPrintL As Integer

Private Sub cmdBack_Click()
    'Add By Cheng 2003/04/18
    '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim i As Integer
    'Add By Cheng 2002/09/09
    blnClkSure = False
    m_print = 0
    'Add By Cheng 2003/04/18
    If Me.Text1(0).Text = "" Then
        MsgBox "請輸入到期起日!!!", vbExclamation + vbOKOnly
        Me.Text1(0).SetFocus
        Text1_GotFocus 0
        Exit Sub
    End If
    If Me.Text1(1).Text = "" Then
        MsgBox "請輸入到期迄日!!!", vbExclamation + vbOKOnly
        Me.Text1(1).SetFocus
        Text1_GotFocus 1
        Exit Sub
    End If
    'Add By Cheng 2002/03/22
    If PUB_CheckKeyInDate(Me.Text1(0)) = -1 Then
       Me.Text1(0).SetFocus
       Text1_GotFocus 0
       Exit Sub
    End If
    If PUB_CheckKeyInDate(Me.Text1(1)) = -1 Then
       Me.Text1(1).SetFocus
       Text1_GotFocus 1
       Exit Sub
    End If
    If ChkRange(Text1(0), Text1(1), "到期日期") = False Then
        blnClkSure = True
        Exit Sub
    End If
    'Add By Cheng 2003/04/18
    If Me.Text1(2).Text <> "" And Me.Text1(3).Text <> "" Then
        If Me.Text1(2).Text > Me.Text1(3).Text Then
            blnClkSure = True
            MsgBox "業務區範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.Text1(2).SetFocus
            Text1_GotFocus 2
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = Me.Combo1.Text Then
            Set Printer = Printers(i)
            Exit For
        End If
    Next i
    PrintCase
    Screen.MousePointer = vbDefault
End Sub

Private Sub PrintCase()
Dim Page As Integer
Dim TmpArea As String
Dim strSaleZone As String '業務區
Dim strSales As String '智權人員
Dim arrDate

On Error GoTo ErrHand
    'Modify By Cheng 2003/04/28
    '依區別, 智權人員, 客戶排序
'    strExc(0) = "SELECT DECODE(CP12,A0901,A0902),DECODE(CP13,ST01,ST02),MIN((SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2))||'-'||(SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2))),HC05,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))), CU79,CP12,CP13 " & _
'                        "FROM HIRECASE,CASEPROGRESS,STAFF,CUSTOMER,ACC090 WHERE " & _
'                        "HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) AND CP13=ST01(+) AND " & _
'                        "CP12=A0901(+) AND (SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+))" & _
'                        strGetcdnSQL & _
'                        " GROUP BY DECODE(CP12,A0901,A0902),DECODE(CP13,ST01,ST02),HC05,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))),CU79,CP12,CP13 " & _
'                        " ORDER BY CP12,CP13 "
'edit by nick 2004/10/29
'    strExc(0) = "SELECT DECODE(CP12,A0901,A0902),DECODE(CP13,ST01,ST02),MIN((SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2))||'-'||(SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2))),HC05,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))), CU79,CP12,CP13 " & _
'                        "FROM HIRECASE,CASEPROGRESS,STAFF,CUSTOMER,ACC090 WHERE " & _
'                        "HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) AND CP13=ST01(+) AND " & _
'                        "CP12=A0901(+) AND (SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+))" & _
'                        strGetcdnSQL & _
'                        " GROUP BY DECODE(CP12,A0901,A0902),DECODE(CP13,ST01,ST02),HC05,DECODE(HC05,CU01||CU02,NVL(CU04,NVL(CU05,CU06))),CU79,CP12,CP13 " & _
'                        " ORDER BY CP12, CP13, HC05 "
    strExc(0) = "SELECT DECODE(CP12,A0901,A0902),DECODE(CP13,ST01,ST02),Max((SUBSTR(CP53,1,4)-1911||'/'||SUBSTR(CP53,5,2)||'/'||SUBSTR(CP53,7,2))||'-'||(SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2))),HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))),C1.CU79,CP12,CP13,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))),C2.CU79 " & _
                        "FROM HIRECASE,CASEPROGRESS,STAFF,CUSTOMER C1,CUSTOMER C2,ACC090 WHERE " & _
                        "HC01=CP01(+) AND HC02=CP02(+) AND HC03=CP03(+) AND HC04=CP04(+) AND CP13=ST01(+) AND " & _
                        "CP12=A0901(+) AND (SUBSTR(HC05,1,8)=C1.CU01(+) AND SUBSTR(HC05,9,1)=C1.CU02(+)) AND (SUBSTR(HC24,1,8)=C2.CU01(+) AND SUBSTR(HC24,9,1)=C2.CU02(+)) and cp57 is null " & _
                        strGetcdnSQL & _
                        " GROUP BY DECODE(CP12,A0901,A0902),DECODE(CP13,ST01,ST02),HC05,DECODE(HC05,C1.CU01||C1.CU02,NVL(C1.CU04,NVL(C1.CU05,C1.CU06))),C1.CU79,CP12,CP13,HC24,DECODE(HC24,C2.CU01||C2.CU02,NVL(C2.CU04,NVL(C2.CU05,C2.CU06))),C2.CU79 " & _
                        " ORDER BY CP12, CP13, HC05 "
    If RsTemp.State = adStateOpen Then RsTemp.Close
    'edit by nick 2004/10/29
    'rsTemp.Open strExc(0), cnnConnection
    RsTemp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
    If RsTemp.EOF And RsTemp.BOF Then
        MsgBox "資料庫內無資料 !", vbInformation
        m_print = 1
        Exit Sub
    End If
    strSaleZone = "" & RsTemp.Fields(0).Value
    strSales = ""
    Page = 1
    CaseTitle strSaleZone, Page
    With RsTemp
        Do While Not .EOF
            '若業務區不同
            If strSaleZone <> "" & .Fields(0).Value Then
                strSaleZone = "" & .Fields(0).Value
                Printer.NewPage
                Page = Page + 1
                CaseTitle strSaleZone, Page
            End If
            If iPrint > 14000 Then
               Printer.NewPage
               Page = Page + 1
               CaseTitle strSaleZone, Page
            End If
            'Modify By Sindy 2011/2/11 當事人1若為X65299時, 則當事人資料改抓當事人2
            If Left(Trim("" & .Fields(3).Value), 6) = "X65299" Then
               '客戶編號
               Printer.CurrentX = PLeft(0) + 125:   Printer.CurrentY = iPrint - 100
               Printer.Print "" & .Fields(8).Value
               '客戶名稱
               Printer.CurrentX = PLeft(1) + 125:    Printer.CurrentY = iPrint - 100
               Printer.Print "" & .Fields(9).Value
               '備註
               Printer.CurrentX = PLeft(5) + 125:    Printer.CurrentY = iPrint - 100
               '2011/8/8 modify by sonia
               'Printer.Print "" & .Fields(10).Value
               Printer.Print "(與謝律師合作)；" & .Fields(10).Value
               '2011/8/8 end
            '2011/2/11 End
            Else
               '客戶編號
               Printer.CurrentX = PLeft(0) + 125:   Printer.CurrentY = iPrint - 100
               Printer.Print "" & .Fields(3).Value
               '客戶名稱
               Printer.CurrentX = PLeft(1) + 125:    Printer.CurrentY = iPrint - 100
               Printer.Print "" & .Fields(4).Value
               '備註
               Printer.CurrentX = PLeft(5) + 125:    Printer.CurrentY = iPrint - 100
               Printer.Print "" & .Fields(5).Value
            End If
            arrDate = Split("" & .Fields(2).Value, "-")
            '開始期間
            Printer.CurrentX = PLeft(2) + (PLeft(3) - PLeft(2) - Printer.TextWidth("" & arrDate(0))) / 2: Printer.CurrentY = iPrint - 100
            Printer.Print "" & arrDate(0)
            '到期期間
            Printer.CurrentX = PLeft(3) + (PLeft(4) - PLeft(3) - Printer.TextWidth("" & arrDate(1))) / 2: Printer.CurrentY = iPrint - 100
            Printer.Print "" & arrDate(1)
            '智權人員
            Printer.CurrentX = PLeft(4) + (PLeft(5) - PLeft(4) - Printer.TextWidth("" & .Fields(1).Value)) / 2: Printer.CurrentY = iPrint - 100
            If strSales <> "" & .Fields(1).Value Then
                Printer.Print "" & .Fields(1).Value
                strSales = "" & .Fields(1).Value
            End If
            iPrint = iPrint + 300
            Printer.Line (PLeft(0), iPrint)-(PLeft(6), iPrint)
            PrintVerLine
            iPrint = iPrint + 300
            .MoveNext
        Loop
    End With
    Printer.EndDoc
    ShowPrintOk
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub CaseTitle(ByVal Area As String, ByVal Page As String)
Dim i As Integer
    GetPrintLeft
    i = 500
    'Modified by Lydia 2018/04/30
    'Printer.PaperSize = vbPRPSFanfoldUS
    Printer.PaperSize = PUB_GetPaperSize(15) '美國標準
    Printer.Font.Size = 22
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    Printer.CurrentX = PLeft(0) + (PLeft(6) - PLeft(0) - Printer.TextWidth("顧問客戶資料表")) / 2:        Printer.CurrentY = i
    Printer.Print "顧問客戶資料表"
    Printer.Font.Underline = False
    Printer.Font.Size = 12
    Printer.Font.Bold = False
    Printer.CurrentX = PLeft(0):               Printer.CurrentY = i + 500
    Printer.Print "列印人 : " & strUserName
    Printer.CurrentX = PLeft(0) + (PLeft(6) - PLeft(0) - Printer.TextWidth("到期日期 : " & ChangeTStringToTDateString(Text1(0)) & _
      " - " & ChangeTStringToTDateString(Text1(1)))) / 2: Printer.CurrentY = i + 500
    Printer.Print "到期日期 : " & ChangeTStringToTDateString(Text1(0)) & _
      " - " & ChangeTStringToTDateString(Text1(1))
    Printer.CurrentX = 16000:             Printer.CurrentY = i + 500
    Printer.Print "列印日期 : " & ChangeTStringToTDateString(GetTaiwanTodayDate)
    Printer.CurrentX = PLeft(0):               Printer.CurrentY = i + 800
    Printer.Print "業務區：" & Area
    Printer.CurrentX = 16000:             Printer.CurrentY = i + 800
    Printer.Print "頁次 : " & Page
    iPrint = i + 1100
    Printer.Line (PLeft(0), iPrint)-(PLeft(6), iPrint)
    iPrint = iPrint + 300
    Printer.CurrentX = PLeft(0) + (PLeft(1) - PLeft(0) - Printer.TextWidth("客戶編號")) / 2:       Printer.CurrentY = iPrint - 100
    Printer.Print "客戶編號"
    Printer.CurrentX = PLeft(1) + (PLeft(2) - PLeft(1) - Printer.TextWidth("客　　戶　　名　　稱")) / 2:          Printer.CurrentY = iPrint - 100
    Printer.Print "客　　戶　　名　　稱"
    Printer.CurrentX = PLeft(2) + (PLeft(3) - PLeft(2) - Printer.TextWidth("開始期限")) / 2:       Printer.CurrentY = iPrint - 100
    Printer.Print "開始期限"
    Printer.CurrentX = PLeft(3) + (PLeft(4) - PLeft(3) - Printer.TextWidth("到期期限")) / 2:       Printer.CurrentY = iPrint - 100
    Printer.Print "到期期限"
    Printer.CurrentX = PLeft(4) + (PLeft(5) - PLeft(4) - Printer.TextWidth("智權人員")) / 2:       Printer.CurrentY = iPrint - 100
    Printer.Print "智權人員"
    Printer.CurrentX = PLeft(5) + (PLeft(6) - PLeft(5) - Printer.TextWidth("備　　　　　　　　註")) / 2:       Printer.CurrentY = iPrint - 100
    Printer.Print "備　　　　　　　　註"
    iPrint = iPrint + 300
    Printer.Line (PLeft(0), iPrint)-(PLeft(6), iPrint)
    PrintVerLine
    iPrint = iPrint + 300
End Sub

Private Sub GetPrintLeft()
   Erase PLeft
   PLeft(0) = 0:     PLeft(1) = 1500
   PLeft(2) = 6500:    PLeft(3) = 8000
   PLeft(4) = 9500:    PLeft(5) = 11000
   PLeft(6) = 18500
End Sub

Private Sub Form_Load()
       
    MoveFormToCenter Me
    '*****************
    '印表設定
    '*****************
    SeekPrintL = Printer.Orientation
    PUB_SetPrinter Me.Name, Combo1, , , SeekPrint 'Modified by Morgan 2017/11/9 設定印表機改呼叫公用函數,原程式移除

End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    'Modify By Cheng 2003/03/31
    KeyAscii = UpperCase(KeyAscii)
    Select Case Index
    Case 0, 1 '到期日期
        If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
    End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Select Case Index
    Case 1 '到期日期
       'Add By Cheng 2002/09/09
       If blnClkSure = False Then
          If Me.Text1(0).Text <> "" And Me.Text1(1).Text <> "" Then
             If Val(Me.Text1(0).Text) > Val(Me.Text1(1).Text) Then
                MsgBox "到期日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                Me.Text1(0).SetFocus
                Text1_GotFocus 0
                Exit Sub
             End If
          End If
       Else
          blnClkSure = False
       End If
    Case 3 '業務區
        If blnClkSure = False Then
            If Me.Text1(2).Text <> "" And Me.Text1(3).Text <> "" Then
                If Me.Text1(2).Text > Me.Text1(3).Text Then
                    MsgBox "業務區範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                    Me.Text1(2).SetFocus
                    Text1_GotFocus 2
                    Exit Sub
                End If
            End If
        Else
            blnClkSure = False
        End If
    End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
            'Modify By Cheng 2003/04/18
            If Me.Text1(Index).Text <> "" Then
                If CheckIsTaiwanDate(Text1(Index)) = False Then Cancel = True
            End If
      End Select
      If Cancel Then TextInverse Text1(Index)
End Sub

Private Function strGetcdnSQL() As String
Dim i As Integer
    strExc(1) = ""
    If Me.Text1(0).Text <> "" Then
        strExc(1) = strExc(1) & " And CP54>=" & ChangeTStringToWString(Text1(0))
    End If
    If Me.Text1(1).Text <> "" Then
        strExc(1) = strExc(1) & " And CP54<=" & ChangeTStringToWString(Text1(1))
    End If
    If Me.Text1(2).Text <> "" Then
        strExc(1) = strExc(1) & " And CP12>='" & Me.Text1(2).Text & "' "
    End If
    If Me.Text1(3).Text <> "" Then
        strExc(1) = strExc(1) & " And CP12<='" & Me.Text1(3).Text & "' "
    End If
    strExc(1) = strExc(1) & " AND CP10='0' AND CP57 IS NULL"
    strGetcdnSQL = strExc(1)
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set Printer = Printers(SeekPrint)
    Printer.Orientation = SeekPrintL
    Set frm073011 = Nothing
End Sub

Private Sub PrintVerLine()
    Printer.Line (PLeft(0), iPrint - 600)-(PLeft(0), iPrint)
    Printer.Line (PLeft(1), iPrint - 600)-(PLeft(1), iPrint)
    Printer.Line (PLeft(2), iPrint - 600)-(PLeft(2), iPrint)
    Printer.Line (PLeft(3), iPrint - 600)-(PLeft(3), iPrint)
    Printer.Line (PLeft(4), iPrint - 600)-(PLeft(4), iPrint)
    Printer.Line (PLeft(5), iPrint - 600)-(PLeft(5), iPrint)
    Printer.Line (PLeft(6), iPrint - 600)-(PLeft(6), iPrint)
End Sub

