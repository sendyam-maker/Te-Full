VERSION 5.00
Begin VB.Form frm170213 
   BorderStyle     =   1  '單線固定
   Caption         =   "互助會得標金額明細"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4740
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3660
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Top             =   60
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   8
      Top             =   2340
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   9
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   1830
      MaxLength       =   5
      TabIndex        =   0
      Top             =   750
      Width           =   675
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   1830
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1110
      Width           =   435
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   2
      Left            =   2400
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1110
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS:計算完薪資才有資料"
      Height          =   180
      Left            =   900
      TabIndex        =   7
      Top             =   1710
      Width           =   1845
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "互助會號："
      Height          =   180
      Left            =   900
      TabIndex        =   6
      Top             =   1140
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   2700
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "得標月份："
      Height          =   180
      Index           =   0
      Left            =   900
      TabIndex        =   5
      Top             =   810
      Width           =   900
   End
End
Attribute VB_Name = "frm170213"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by SINDY 2009/01/05
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
'Dim PaperX As Double
'Dim paperY As Double
Dim iPgae As Integer, iLine As Integer
Dim strType As String


Private Sub cmdok_Click(Index As Integer)
Dim strYM As String
Select Case Index
Case 0
        If txt1(0) = "" And txt1(1) = "" And txt1(2) = "" Then
            MsgBox "至少輸入一項列印條件！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        If txt1(0) <> "" Then
            If Len(txt1(0)) <= 3 Then
                MsgBox "得標月份輸入錯誤！", vbInformation, "操作錯誤！"
                txt1(0).SetFocus
                Exit Sub
            End If
            If ChkDate(txt1(0) & "01") = False Then
                txt1(0).SetFocus
                Exit Sub
            End If
        End If
        If txt1(1) <> "" Or txt1(2) <> "" Then
            If RunNick(txt1(1), txt1(2)) Then
               txt1(1).SetFocus
               Exit Sub
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            strYM = Left(ChangeTStringToWString(txt1(0) & "01"), 6)
            m_StrSQL = m_StrSQL & " and substr(a.cm04,1,6)='" & strYM & "' "
        End If
        If txt1(1) <> "" Then
            m_StrSQL = m_StrSQL & " and a.cm01 >= '" & txt1(1) & "' "
        End If
        If txt1(2) <> "" Then
            m_StrSQL = m_StrSQL & " and a.cm01 <= '" & txt1(2) & "' "
        End If
        StrMenu1
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
End Select
End Sub


Sub StrMenu1()
Dim dblAmt As Double
Dim dblCnt As Double
Dim dblMainAmt As Double, strSql As String

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

'modify by sonia 2015/6/8 同一會號有二個以上60000時,會款金額會重覆計算 2015/5,會號315
'm_str = "select decode(wfa03,'60000','1','2') as sort1,a.cm01,a.cm04,a.cm03,a.cm05,b.cm02,c.ST02,wfa04,d.ST02 " & _
              "from CooperationMember a,WfAmount,CooperationMember b,Staff c,Staff d " & _
             "where a.cm04(+)=wfa01 and wfa05='2' and wfa02=a.cm01(+) " & _
               "and b.cm01(+)=wfa02 " & _
               "and b.cm03(+)=wfa03 " & _
               "and c.ST01(+)=wfa03 and d.ST01(+)=a.cm03 " & m_StrSQL & _
               "order by a.cm01,sort1,b.cm02 "
m_str = "select decode(wfa03,'60000','1','2') as sort1,a.cm01,a.cm04,a.cm03,a.cm05,b.cm02,c.ST02,decode(wfa03,'60000',decode(nvl(b.cm05,0),0,co02,co02+b.cm05),wfa04) as wfa04,d.ST02 " & _
              "from CooperationMember a,WfAmount,CooperationMember b,Staff c,Staff d,Cooperation " & _
             "where a.cm04(+)=wfa01 and wfa05='2' and wfa02=a.cm01(+) and a.cm01=co01(+) " & _
               "and b.cm01(+)=wfa02 " & _
               "and b.cm03(+)=wfa03 " & _
               "and c.ST01(+)=wfa03 and d.ST01(+)=a.cm03 " & m_StrSQL & _
               "order by a.cm01,sort1,b.cm02 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        iLine = 1
        'PrintTitle '列印表頭
        strType = "" '切頁條件
        dblAmt = 0
        dblCnt = 0
        Do While Not m_rs.EOF
            
            For m_i = 1 To 8
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(1)) '會號
            strTemp(2) = CheckStr(m_rs.Fields(2)) '得標日
            strTemp(3) = CheckStr(m_rs.Fields(3)) '得標人-代號
            strTemp(4) = CheckStr(m_rs.Fields(4)) '得標金額
            strTemp(5) = CheckStr(m_rs.Fields(5)) '編號
            strTemp(6) = CheckStr(m_rs.Fields(6)) '扣款員工-姓名
            strTemp(7) = CheckStr(m_rs.Fields(7)) '會款
            strTemp(8) = CheckStr(m_rs.Fields(8)) '得標人-姓名
            
            If iLine > 50 Or iLine = 1 Or _
                  (strType <> strTemp(1)) Then
                  
                If (strType <> "" And strType <> strTemp(1)) Then
                   Printer.CurrentX = 500
                   Printer.CurrentY = iLine * 300
                   Printer.Print String(140, "-")
                   
                   iLine = iLine + 1
                   Printer.CurrentX = PLeft(2)
                   Printer.CurrentY = iLine * 300
                   Printer.Print "合　計：" & dblCnt & "人"
                   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmt, "##,##0"))
                   Printer.CurrentY = iLine * 300
                   Printer.Print Format(dblAmt, "##,##0")
                   
                   dblAmt = 0
                   dblCnt = 0
                End If
                
                'If .AbsolutePosition <> .RecordCount Then
                    If strType <> "" Then Printer.NewPage
                    iLine = 1
                    PrintTitle '列印表頭
                'End If
            End If
            
            'Add By Sindy 2010/6/4
            If strType <> strTemp(1) Then
               '取得會號之投標金額
               dblMainAmt = 0
               strSql = "SELECT * FROM Cooperation WHERE CO01='" & strTemp(1) & "' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  dblMainAmt = Val(RsTemp("CO02"))
               End If
               '新增一筆會首資料
               '編號
               Printer.CurrentX = PLeft(1)
               Printer.CurrentY = iLine * 300
               Printer.Print "00"
               '姓　名
               Printer.CurrentX = PLeft(2)
               Printer.CurrentY = iLine * 300
               Printer.Print "台一"
               '會　款
               Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(7), "##,##0"))
               Printer.CurrentY = iLine * 300
               Printer.Print Format(dblMainAmt, "##,##0")
               dblAmt = dblAmt + dblMainAmt
               dblCnt = dblCnt + 1
               iLine = iLine + 1
            End If
            '2010/6/4 End
            
            PrintDetail '列印表中
            
            strType = strTemp(1) '依會號跳頁
            dblAmt = dblAmt + strTemp(7)
            dblCnt = dblCnt + 1
            m_rs.MoveNext
        Loop
        
        '列印表尾
        Printer.CurrentX = 500
        Printer.CurrentY = iLine * 300
        Printer.Print String(140, "-")
        
        iLine = iLine + 1
        Printer.CurrentX = PLeft(2)
        Printer.CurrentY = iLine * 300
        Printer.Print "合　計：" & dblCnt & "人"
        Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmt, "##,##0"))
        Printer.CurrentY = iLine * 300
        Printer.Print Format(dblAmt, "##,##0")
    End With
Else
    MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintTitle()
GetPleft

'PaperX = 12000
'paperY = 7500

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("互助會得標金額明細") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "互助會得標金額明細"

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "列印人：" & strUserName
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("得標月份：" & Left(strTemp(2), 4) - 1911 & "  年  " & Mid(strTemp(2), 5, 2) & "  月") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "得標月份：" & Left(strTemp(2), 4) - 1911 & "  年  " & Mid(strTemp(2), 5, 2) & "  月"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "會　號：" & strTemp(1)
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("得標月份：" & Left(strTemp(2), 4) - 1911 & "  年  " & Mid(strTemp(2), 5, 2) & "  月") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "得  標  人：" & strTemp(3) & " " & strTemp(8)
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "得標金額：" & strTemp(4)

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "編號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"
Printer.CurrentX = PLeft(3) - Printer.TextWidth("會　款")
Printer.CurrentY = iLine * 300
Printer.Print "會　款"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 3500
PLeft(2) = 5500
PLeft(3) = 8500
End Sub

Sub PrintDetail()
   '編號
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(5)
   '姓　名
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(6)
   '會　款
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(7), "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "##,##0")
   
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
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170213 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "01") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub


