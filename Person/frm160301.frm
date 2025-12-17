VERSION 5.00
Begin VB.Form frm160301 
   BorderStyle     =   1  '單線固定
   Caption         =   "出缺勤年統計及全勤名單"
   ClientHeight    =   3240
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5460
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   2490
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1260
      Width           =   555
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   1860
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1260
      Width           =   555
   End
   Begin VB.CheckBox Check3 
      Caption         =   "出缺勤年統計表(依假別)"
      Height          =   195
      Left            =   1860
      TabIndex        =   7
      Top             =   2100
      Width           =   2955
   End
   Begin VB.CheckBox Check2 
      Caption         =   "全勤名單"
      Height          =   195
      Left            =   1860
      TabIndex        =   6
      Top             =   1860
      Width           =   2955
   End
   Begin VB.CheckBox Check1 
      Caption         =   "出缺勤年統計表(依1~12月)"
      Height          =   195
      Left            =   1860
      TabIndex        =   5
      Top             =   1620
      Width           =   2955
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3480
      TabIndex        =   8
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4425
      TabIndex        =   9
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1860
      MaxLength       =   6
      TabIndex        =   1
      Top             =   930
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2640
      MaxLength       =   6
      TabIndex        =   2
      Top             =   930
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1860
      MaxLength       =   3
      TabIndex        =   0
      Top             =   600
      Width           =   525
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   10
      Top             =   2580
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   11
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   12
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2790
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   900
      TabIndex        =   16
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "報表種類："
      Height          =   180
      Index           =   2
      Left            =   930
      TabIndex        =   15
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   930
      TabIndex        =   14
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "年　　度："
      Height          =   180
      Left            =   930
      TabIndex        =   13
      Top             =   630
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2490
      X2              =   2730
      Y1              =   1050
      Y2              =   1050
   End
End
Attribute VB_Name = "frm160301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'Create by SINDY 2009/01/08
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 25) As Integer
Dim strTemp(1 To 6) As String
Dim strTemp01(1 To 25) As String, strTemp02(1 To 25) As String, strTemp03(1 To 25) As String, strTemp04(1 To 25) As String
Dim strTemp05(1 To 25) As String, strTemp06(1 To 25) As String, strTemp07(1 To 25) As String, strTemp08(1 To 25) As String
Dim strTemp09(1 To 25) As String, strTemp10(1 To 25) As String, strTemp11(1 To 25) As String, strTemp12(1 To 25) As String
'Dim PaperX As Double
'Dim paperY As Double
Dim iPgae As Integer, iLine As Integer
Dim LongPrintCurCnt As Long
Dim strYear As String
Dim dblAmt As Double


Private Sub cmdok_Click(Index As Integer)
Select Case Index
   Case 0
      If txt1(0) = "" Then
         MsgBox "年度不可以空白！", vbInformation, "操作錯誤！"
         txt1(0).SetFocus
         Exit Sub
      End If
      If Len(txt1(0)) <= 1 Then
         MsgBox "年度輸入錯誤！", vbInformation, "操作錯誤！"
         txt1(0).SetFocus
         Exit Sub
      End If
      'Modify By Sindy 2016/1/5 + And Check3.Value = 0
      If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 Then
         MsgBox "報表種類至少勾選一項！", vbInformation, "操作錯誤！"
         Check1.SetFocus
         Exit Sub
      End If
      
      strYear = Val(txt1(0)) + 1911
      
      Screen.MousePointer = vbHourglass
      m_StrSQL = ""
      If txt1(1) <> "" Then
         m_StrSQL = m_StrSQL & " and st01>='" & txt1(1) & "' "
      End If
      If txt1(2) <> "" Then
         m_StrSQL = m_StrSQL & " and st01<='" & txt1(2) & "' "
      End If
      'Add By Sindy 2016/1/5
      If txt1(3) <> "" Then
         'Modify By Sindy 2023/12/27 部門調整改抓ST93
         m_StrSQL = m_StrSQL & " AND ST93>='" & txt1(3) & "'"
      End If
      If txt1(4) <> "" Then
         'Modify By Sindy 2023/12/27 部門調整改抓ST93
         m_StrSQL = m_StrSQL & " AND ST93<='" & txt1(4) & "'"
      End If
      '2016/1/5 END
      '出缺勤年統計表
      If Check1.Value = 1 Then
         If StrMenu1 = False Then Check1.Value = 0
      End If
      '全勤名單
      If Check2.Value = 1 Then
         If StrMenu2 = False Then Check2.Value = 0
      End If
      'Add By Sindy 2016/1/5
      '出缺勤年統計表(假別)
      If Check3.Value = 1 Then
         If StrMenu3 = False Then Check3.Value = 0
      End If
      '2016/1/5 END
      
      'Modify By Sindy 2016/1/5 + Or Check3.Value = 1
      If Check1.Value = 1 Or Check2.Value = 1 Or Check3.Value = 1 Then
         ShowPrintOk
      End If
      
      Screen.MousePointer = vbDefault
   Case 1
      Unload Me
End Select
End Sub

'出缺勤年統計表
Function StrMenu1() As Boolean
Dim strSql As String
'Dim dblHour(18) As Double, dblCnt(18) As Double
'Dim dblHour(22) As Double, dblCnt(22) As Double...移到basPerson共用變數區
Dim i As Integer, j As Integer
Dim strSDate As String, strEDate As String

StrMenu1 = True

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

'63001.董事長 67004.副董事長 L01.律師 不印出缺勤統計表
'm_str = "SELECT ST01,ST02,a1.a0901||' '||a1.a0902,ST40 "
'Modify By Sindy 2016/10/11 L01.法務處律師只有3個人不打卡,其他都要(and ST03<>'L01'取消,改寫法)
'Modify By Sindy 2023/12/27 部門調整改抓ST93
m_str = "SELECT ST01,ST02,a1.a0921||' '||a1.a0922 " & _
               "FROM Staff,Acc090NEW a1,SalaryData " & _
               "WHERE ST04='1' and ST01=SD01 and (sd02 not in('P','F') or sd02 is null) " & _
               "and ST93=a1.a0921(+) " & _
               "and ST01 not in('63001','67004','" & Replace(Pub_GetSpecMan("不用打卡的律師"), ";", "','") & "') " & m_StrSQL & _
               "ORDER BY ST93,ST01 ASC "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
LongPrintCurCnt = 0
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        Do While Not m_rs.EOF
            LongPrintCurCnt = LongPrintCurCnt + 1
            
            For m_i = 1 To 6
               strTemp(m_i) = ""
            Next m_i
            For m_i = 1 To 25 '23 '22 '20
                strTemp01(m_i) = ""
                strTemp02(m_i) = ""
                strTemp03(m_i) = ""
                strTemp04(m_i) = ""
                strTemp05(m_i) = ""
                strTemp06(m_i) = ""
                strTemp07(m_i) = ""
                strTemp08(m_i) = ""
                strTemp09(m_i) = ""
                strTemp10(m_i) = ""
                strTemp11(m_i) = ""
                strTemp12(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields("ST01"))
            strTemp(2) = CheckStr(m_rs.Fields("ST02"))
            strTemp(3) = CheckStr(m_rs.Fields(2))
'            strTemp(4) = CheckStr(m_rs.Fields("ST40"))
            strSql = " and ST01='" & Trim(strTemp(1)) & "' "
            
            '1~12月
            For i = 1 To 12
               strSDate = strYear & Format(i, "00") & "01"
               strEDate = strYear & Format(i, "00") & "31"
               If PUB_GetAbsenceHour(strSql, strSDate, strEDate, dblHour(), dblCnt()) = True Then
                  '10.產假=產假+17.扣年終產假
                  '11.流產假=流產假+18.扣年終流產假
                  dblHour(10) = dblHour(10) + dblHour(17)
                  dblHour(11) = dblHour(11) + dblHour(18)
                  '1~18個假別
                  'For j = 1 To 18
                  'Modify By Sindy 2012/1/4 +19.陪產假
                  'Modify By Sindy 2012/1/4 +20.生理假 21.產檢假 22.家庭照顧假
                  For j = 1 To 25 '23 '22 '20
                     If i = 1 Then strTemp01(j) = dblHour(j)
                     If i = 2 Then strTemp02(j) = dblHour(j)
                     If i = 3 Then strTemp03(j) = dblHour(j)
                     If i = 4 Then strTemp04(j) = dblHour(j)
                     If i = 5 Then strTemp05(j) = dblHour(j)
                     If i = 6 Then strTemp06(j) = dblHour(j)
                     If i = 7 Then strTemp07(j) = dblHour(j)
                     If i = 8 Then strTemp08(j) = dblHour(j)
                     If i = 9 Then strTemp09(j) = dblHour(j)
                     If i = 10 Then strTemp10(j) = dblHour(j)
                     If i = 11 Then strTemp11(j) = dblHour(j)
                     If i = 12 Then strTemp12(j) = dblHour(j)
                  Next j
               End If
            Next i
            
            PrintTitle '列印表頭
            PrintDetail '列印表中、表尾
            
            '每二筆才換新頁
'            If LongPrintCurCnt Mod 2 = 0 Then
               Printer.NewPage
'            End If
            
            m_rs.MoveNext
        Loop
    End With
Else
   StrMenu1 = False
   MsgBox "出缺勤年統計表，無符合列印的資料！", vbExclamation + vbOKOnly
   Exit Function
End If
Printer.EndDoc
'ShowPrintOk
End Function

Sub PrintTitle()
GetPleft

'PaperX = 12000
'paperY = 7500

'列印行數
'If LongPrintCurCnt Mod 2 <> 0 Then
   iLine = 1 '新頁重頭列印
   
   Printer.Font.Size = 16
   Printer.Font.Underline = False
   Printer.FontBold = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("出　缺　勤　年　統　計　表") / 2)
   Printer.CurrentY = iLine * 250
   Printer.Print "出　缺　勤　年　統　計　表"
   
   Printer.Font.Size = 12
   Printer.FontBold = False
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 250
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 1
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(txt1(0) & "　年度") / 2)
   Printer.CurrentY = iLine * 250
   Printer.Print txt1(0) & "　年度"
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   Printer.CurrentY = iLine * 250
   Printer.Print "頁　　次：" & Printer.Page
'Else
'   iLine = iLine + 1 '接續列印
'End If

Printer.Font.Size = 12
Printer.FontBold = False
iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 250
Printer.Print "編號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 250
Printer.Print "姓　名"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 250
Printer.Print "項　目"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("1月")
Printer.CurrentY = iLine * 250
Printer.Print "1月"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("2月")
Printer.CurrentY = iLine * 250
Printer.Print "2月"
Printer.CurrentX = PLeft(6) - Printer.TextWidth("3月")
Printer.CurrentY = iLine * 250
Printer.Print "3月"
Printer.CurrentX = PLeft(7) - Printer.TextWidth("4月")
Printer.CurrentY = iLine * 250
Printer.Print "4月"
Printer.CurrentX = PLeft(8) - Printer.TextWidth("5月")
Printer.CurrentY = iLine * 250
Printer.Print "5月"
Printer.CurrentX = PLeft(9) - Printer.TextWidth("6月")
Printer.CurrentY = iLine * 250
Printer.Print "6月"
Printer.CurrentX = PLeft(10) - Printer.TextWidth("7月")
Printer.CurrentY = iLine * 250
Printer.Print "7月"
Printer.CurrentX = PLeft(11) - Printer.TextWidth("8月")
Printer.CurrentY = iLine * 250
Printer.Print "8月"
Printer.CurrentX = PLeft(12) - Printer.TextWidth("9月")
Printer.CurrentY = iLine * 250
Printer.Print "9月"
Printer.CurrentX = PLeft(13) - Printer.TextWidth("10月")
Printer.CurrentY = iLine * 250
Printer.Print "10月"
Printer.CurrentX = PLeft(14) - Printer.TextWidth("11月")
Printer.CurrentY = iLine * 250
Printer.Print "11月"
Printer.CurrentX = PLeft(15) - Printer.TextWidth("12月")
Printer.CurrentY = iLine * 250
Printer.Print "12月"
Printer.CurrentX = PLeft(16) - Printer.TextWidth("總計")
Printer.CurrentY = iLine * 250
Printer.Print "總計"

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 250
Printer.Print String(210, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 300
PLeft(2) = 1000
PLeft(3) = 2200
PLeft(4) = 4000
PLeft(5) = 5000
PLeft(6) = 6000
PLeft(7) = 7000
PLeft(8) = 8000
PLeft(9) = 9000
PLeft(10) = 10000
PLeft(11) = 11000
PLeft(12) = 12000
PLeft(13) = 13000
PLeft(14) = 14000
PLeft(15) = 15000
PLeft(16) = 16000
End Sub

Sub PrintDetail()
Dim m_i As Integer, m_j As Integer, Item As Integer

Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 250
Printer.Print strTemp(1)
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 250
Printer.Print strTemp(2)
'For m_j = 1 To 16
For m_j = 1 To 22 '21 '20 '17
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 250
   If m_j = 1 Then
      Printer.Print "忘打卡(次)": Item = 1
   ElseIf m_j = 2 Then
      Printer.Print "遲　到(次)": Item = 2
   ElseIf m_j = 3 Then
      Printer.Print "曠　職(分)": Item = 3
   ElseIf m_j = 4 Then
      Printer.Print "出　差": Item = 4
   ElseIf m_j = 5 Then
      Printer.Print "事　假": Item = 5
   'Add By Sindy 2014/12/10 22.家庭照顧假
   ElseIf m_j = 6 Then
      Printer.Print "家庭照顧假": Item = 22
   ElseIf m_j = 7 Then
      Printer.Print "病　假": Item = 6
   'Add By Sindy 2014/12/10 20.生理假
   ElseIf m_j = 8 Then
      Printer.Print "生理假": Item = 20
   'Add By Sindy 2015/1/5 23.健檢假
   ElseIf m_j = 9 Then
      GoTo goStep
      Printer.Print "健檢假": Item = 23
   ElseIf m_j = 10 Then
      Printer.Print "公　假": Item = 7
   ElseIf m_j = 11 Then
      Printer.Print "特別假": Item = 8
   ElseIf m_j = 12 Then
      Printer.Print "婚　假": Item = 9
   'Add By Sindy 2014/12/10 21.產檢假
   ElseIf m_j = 13 Then
      Printer.Print "產檢假": Item = 21
   ElseIf m_j = 14 Then
      Printer.Print "產　假": Item = 10
   ElseIf m_j = 15 Then
      Printer.Print "流產假": Item = 11
   'Add By Sindy 2012/1/4 +19.陪產假
   ElseIf m_j = 16 Then
      Printer.Print "陪產假": Item = 19
   ElseIf m_j = 17 Then
      Printer.Print "喪　假": Item = 12
   ElseIf m_j = 18 Then
      Printer.Print "公傷假": Item = 13
   ElseIf m_j = 19 Then
      Printer.Print "補　休": Item = 14
   ElseIf m_j = 20 Then
      Printer.Print "其　他": Item = 15
   'Add By Sindy 2025/11/17 +25.天災不給薪
   ElseIf m_j = 21 Then
      Printer.Print "天災不給薪": Item = 25
   '2025/11/17 END
   ElseIf m_j = 22 Then
      Printer.Print "加　班": Item = 16
   End If
   
   dblAmt = 0
   'Modify By Sindy 2025/11/17
   If Item = 1 Or Item = 2 Or Item = 3 Then
      For m_i = 1 To 12
         If m_i = 1 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp01(Item), "##0"))
         If m_i = 2 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp02(Item), "##0"))
         If m_i = 3 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp03(Item), "##0"))
         If m_i = 4 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp04(Item), "##0"))
         If m_i = 5 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp05(Item), "##0"))
         If m_i = 6 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp06(Item), "##0"))
         If m_i = 7 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp07(Item), "##0"))
         If m_i = 8 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp08(Item), "##0"))
         If m_i = 9 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp09(Item), "##0"))
         If m_i = 10 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp10(Item), "##0"))
         If m_i = 11 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp11(Item), "##0"))
         If m_i = 12 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp12(Item), "##0"))
               
         Printer.CurrentY = iLine * 250
         
         If m_i = 1 Then Printer.Print Format(strTemp01(Item), "##0")
         If m_i = 2 Then Printer.Print Format(strTemp02(Item), "##0")
         If m_i = 3 Then Printer.Print Format(strTemp03(Item), "##0")
         If m_i = 4 Then Printer.Print Format(strTemp04(Item), "##0")
         If m_i = 5 Then Printer.Print Format(strTemp05(Item), "##0")
         If m_i = 6 Then Printer.Print Format(strTemp06(Item), "##0")
         If m_i = 7 Then Printer.Print Format(strTemp07(Item), "##0")
         If m_i = 8 Then Printer.Print Format(strTemp08(Item), "##0")
         If m_i = 9 Then Printer.Print Format(strTemp09(Item), "##0")
         If m_i = 10 Then Printer.Print Format(strTemp10(Item), "##0")
         If m_i = 11 Then Printer.Print Format(strTemp11(Item), "##0")
         If m_i = 12 Then Printer.Print Format(strTemp12(Item), "##0")
      Next m_i
   Else
   '2025/11/17 END
      For m_i = 1 To 12
         If m_i = 1 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp01(Item), "##0.0"))
         If m_i = 2 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp02(Item), "##0.0"))
         If m_i = 3 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp03(Item), "##0.0"))
         If m_i = 4 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp04(Item), "##0.0"))
         If m_i = 5 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp05(Item), "##0.0"))
         If m_i = 6 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp06(Item), "##0.0"))
         If m_i = 7 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp07(Item), "##0.0"))
         If m_i = 8 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp08(Item), "##0.0"))
         If m_i = 9 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp09(Item), "##0.0"))
         If m_i = 10 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp10(Item), "##0.0"))
         If m_i = 11 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp11(Item), "##0.0"))
         If m_i = 12 Then Printer.CurrentX = PLeft(m_i + 3) - Printer.TextWidth(Format(strTemp12(Item), "##0.0"))
               
         Printer.CurrentY = iLine * 250
         
         If m_i = 1 Then Printer.Print Format(strTemp01(Item), "##0.0")
         If m_i = 2 Then Printer.Print Format(strTemp02(Item), "##0.0")
         If m_i = 3 Then Printer.Print Format(strTemp03(Item), "##0.0")
         If m_i = 4 Then Printer.Print Format(strTemp04(Item), "##0.0")
         If m_i = 5 Then Printer.Print Format(strTemp05(Item), "##0.0")
         If m_i = 6 Then Printer.Print Format(strTemp06(Item), "##0.0")
         If m_i = 7 Then Printer.Print Format(strTemp07(Item), "##0.0")
         If m_i = 8 Then Printer.Print Format(strTemp08(Item), "##0.0")
         If m_i = 9 Then Printer.Print Format(strTemp09(Item), "##0.0")
         If m_i = 10 Then Printer.Print Format(strTemp10(Item), "##0.0")
         If m_i = 11 Then Printer.Print Format(strTemp11(Item), "##0.0")
         If m_i = 12 Then Printer.Print Format(strTemp12(Item), "##0.0")
      Next m_i
   End If
   
   dblAmt = Val(strTemp01(Item)) + Val(strTemp02(Item)) + Val(strTemp03(Item)) + Val(strTemp04(Item)) + Val(strTemp05(Item)) + Val(strTemp06(Item)) + Val(strTemp07(Item)) + Val(strTemp08(Item)) + Val(strTemp09(Item)) + Val(strTemp10(Item)) + Val(strTemp11(Item)) + Val(strTemp12(Item))
   'Modify By Sindy 2025/11/17
   If Item = 1 Or Item = 2 Or Item = 3 Then
      Printer.CurrentX = PLeft(16) - Printer.TextWidth(Format(dblAmt, "##0"))
      Printer.CurrentY = iLine * 250
      Printer.Print Format(dblAmt, "##0")
   Else
   '2025/11/17 END
      Printer.CurrentX = PLeft(16) - Printer.TextWidth(Format(dblAmt, "##0.0"))
      Printer.CurrentY = iLine * 250
      Printer.Print Format(dblAmt, "##0.0")
   End If
   iLine = iLine + 1
goStep:
Next m_j
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
   Set frm160301 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      'Modify By Sindy 2016/1/5 + 3,4
      Case 1, 2, 3, 4
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "0101") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case 1, 2
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 1 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 2 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      'Add By Sindy 2016/1/5
      Case 3, 4
         If Index = 3 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 4 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      '2016/1/5 END
      Case Else
   End Select
End Sub

'全勤名單
Function StrMenu2() As Boolean
Dim strSql As String, strType As String
Dim i As Integer, intRow As Integer
Dim strSDate As String, strEDate As String

StrMenu2 = True

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

strSDate = strYear & "0101"
strEDate = strYear & "1231"
'Modify By Sindy 2015/2/16 73029廖宗岳,99029伊恩不列作全勤名單,及員工編號第4碼不可為9
'Modify By Sindy 2022/7/18 + 15.名譽所長
'Modified by Morgan 2023/11/7 SQL語法改共用(薪資系統也要用)
'm_str = "select a0902,ST01,ST02 " & _
          "From staff,acc090 " & _
         "where ST04='1' " & _
           "and st03=a0901 " & _
           "and substr(ST01,1,1) in ('6','7','8','9','A','B','C','D','E') " & _
           "and ST01 not in(select distinct SA01 from Staff_Assist where SA02 between " & strSDate & " and " & strEDate & " and (SA04>0 or SA05>0 or SA06>0)) " & _
           "and ST01 not in(select distinct SA01 from Staff_Absence where SA02 between " & strSDate & " and " & strEDate & " and SA06 in ('05','06')) " & _
           "and ST01 not in('67001','68007','86026','68091','68092','94099','97099','99998','60000','99997','99099','68099','99999','96029','96030','68096','73029','99029') " & _
           "and substr(ST01,4,1)<>'9' " & _
           "and ST01 not in(SELECT distinct ST01 From Staff,Staff_Change WHERE ST01=SC01 AND ST04='1' AND (ST13>=" & strSDate & " OR (SC02 between " & strSDate & " AND " & strEDate & " AND SC03='02'))) " & _
           "and ST01 not in(select st01 from staff,allcode where ac01(+)='01' and st20=ac02(+) and (ac02 in('01','02','21','22','11','12','15') or st01 in('94015','79037'))) " & _
           "and ST01 not in(select ST01 from staff where substr(st03,1,1)='R' and st04='1') " & _
           m_StrSQL & _
        "order by st06,st03,st01 "
'Modify By Sindy 2023/12/27 部門調整改抓ST93
m_str = "select a0922,ST01,ST02 " & _
          "From (" & PUB_GetFullAttendanceStaff(strSDate, strEDate) & ") X,acc090NEW " & _
         "where st93=a0921(+) " & m_StrSQL & _
        "order by st06,st93,st01 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
LongPrintCurCnt = 0
If Not m_rs.EOF And Not m_rs.BOF Then
   With m_rs
      m_rs.MoveFirst
      iLine = 1
      strType = ""
      'Do While Not m_rs.EOF
      For intRow = 1 To m_rs.RecordCount
         If (intRow Mod 2) <> 0 Then '筆數為單數時
            For i = 1 To 6
               strTemp(i) = ""
            Next i
            strTemp(1) = CheckStr(m_rs.Fields(0))
            strTemp(2) = CheckStr(m_rs.Fields(1))
            strTemp(3) = CheckStr(m_rs.Fields(2))
         End If
         '筆數為偶數或最後一筆資料
         If (intRow Mod 2) = 0 Or intRow = m_rs.RecordCount Then
            If (intRow Mod 2) = 0 Then
               strTemp(4) = CheckStr(m_rs.Fields(0))
               strTemp(5) = CheckStr(m_rs.Fields(1))
               strTemp(6) = CheckStr(m_rs.Fields(2))
            End If
            
            If iLine > 54 Or iLine = 1 Then
               If strType <> "" Then Printer.NewPage
               iLine = 1
               PrintTitle2 '列印表頭
            End If
            PrintDetail2
            
            strType = CheckStr(m_rs.Fields("ST01"))
         End If
         m_rs.MoveNext
      'Loop
      Next intRow
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iLine * 300
      Printer.Print "共計　" & m_rs.RecordCount & "　筆"
   End With
Else
   StrMenu2 = False
   MsgBox "全勤名單，無符合列印的資料！", vbExclamation + vbOKOnly
   Exit Function
End If
Printer.EndDoc
'ShowPrintOk
End Function

Sub GetPleft2()
PLeft(1) = 500
PLeft(2) = 2500
PLeft(3) = 4000
PLeft(4) = 6000
PLeft(5) = 8000
PLeft(6) = 9500
End Sub

Sub PrintTitle2()
GetPleft2

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(Val(txt1(0)) & "年度全勤名單") / 2)
Printer.CurrentY = 300
Printer.Print Val(txt1(0)) & "年度全勤名單"

Printer.Font.Size = 12
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 4
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "部　門"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "員工代碼"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "部　門"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "員工代碼"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail2()
Dim m_j As Integer
   For m_j = 1 To 6
      Printer.CurrentX = PLeft(m_j)
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine = iLine + 1
End Sub

'Add By Sindy 2016/1/5
'出缺勤年統計表(依假別)
Function StrMenu3() As Boolean
Dim strSql As String
Dim i As Integer, j As Integer
Dim strSDate As String, strEDate As String

StrMenu3 = True

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

'63001.董事長 67004.副董事長 L01.律師 不印出缺勤統計表
'Modify By Sindy 2016/10/11 L01.法務處律師只有3個人不打卡,其他都要(and ST03<>'L01'取消,改寫法)
'Modify By Sindy 2023/12/27 部門調整改抓ST93
m_str = "SELECT ST01,ST02,a1.a0921||' '||a1.a0922 " & _
               "FROM Staff,Acc090NEW a1,SalaryData " & _
               "WHERE ST04='1' and ST01=SD01 and (sd02 not in('P','F') or sd02 is null) " & _
               "and ST93=a1.a0921(+) " & _
               "and ST01 not in('63001','67004','" & Replace(Pub_GetSpecMan("不用打卡的律師"), ";", "','") & "') " & m_StrSQL & _
               "ORDER BY ST93,ST01 ASC "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        PrintTitle3 '列印表頭
        Do While Not m_rs.EOF
            If iLine > 36 Then
               Printer.NewPage
               PrintTitle3 '列印表頭
            End If
            
            For m_i = 1 To 6
               strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields("ST01"))
            strTemp(2) = CheckStr(m_rs.Fields("ST02"))
            strTemp(3) = CheckStr(m_rs.Fields(2))
            strSql = " and ST01='" & Trim(strTemp(1)) & "' "
            
            strSDate = strYear & "0101"
            strEDate = strYear & "1231"
            If PUB_GetAbsenceHour(strSql, strSDate, strEDate, dblHour(), dblCnt()) = True Then
               '10.產假=產假+17.扣年終產假
               '11.流產假=流產假+18.扣年終流產假
               dblHour(10) = dblHour(10) + dblHour(17)
               dblHour(11) = dblHour(11) + dblHour(18)
            End If
            
            PrintDetail3 '列印表中、表尾
            
            m_rs.MoveNext
        Loop
    End With
Else
   StrMenu3 = False
   MsgBox "出缺勤年統計表(假別)，無符合列印的資料！", vbExclamation + vbOKOnly
   Exit Function
End If
Printer.EndDoc
'ShowPrintOk
End Function

Sub PrintTitle3()
GetPleft3
iLine = 1
'PaperX = 12000
'paperY = 7500

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = True

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("出　缺　勤　年　統　計　表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "出　缺　勤　年　統　計　表"

Printer.Font.Size = 12
Printer.FontBold = False
iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(txt1(0) & "　年度") / 2)
Printer.CurrentY = iLine * 300
Printer.Print txt1(0) & "　年度"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "編號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "姓名"
Printer.CurrentX = PLeft(3) - Printer.TextWidth("忘打卡")
Printer.CurrentY = iLine * 300
Printer.Print "忘打卡"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("遲到")
Printer.CurrentY = iLine * 300
Printer.Print "遲到"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("曠職")
Printer.CurrentY = iLine * 300
Printer.Print "曠職"
Printer.CurrentX = PLeft(6) - Printer.TextWidth("出差")
Printer.CurrentY = iLine * 300
Printer.Print "出差"
Printer.CurrentX = PLeft(7) - Printer.TextWidth("事假")
Printer.CurrentY = iLine * 300
Printer.Print "事假"
Printer.CurrentX = PLeft(8) - Printer.TextWidth("照顧假")
Printer.CurrentY = iLine * 300
Printer.Print "照顧假"
Printer.CurrentX = PLeft(9) - Printer.TextWidth("病假")
Printer.CurrentY = iLine * 300
Printer.Print "病假"
Printer.CurrentX = PLeft(10) - Printer.TextWidth("生理假")
Printer.CurrentY = iLine * 300
Printer.Print "生理假"
Printer.CurrentX = PLeft(11) - Printer.TextWidth("公假")
Printer.CurrentY = iLine * 300
Printer.Print "公假"
Printer.CurrentX = PLeft(12) - Printer.TextWidth("特別假")
Printer.CurrentY = iLine * 300
Printer.Print "特別假"
Printer.CurrentX = PLeft(13) - Printer.TextWidth("婚假")
Printer.CurrentY = iLine * 300
Printer.Print "婚假"
Printer.CurrentX = PLeft(14) - Printer.TextWidth("產檢假")
Printer.CurrentY = iLine * 300
Printer.Print "產檢假"
Printer.CurrentX = PLeft(15) - Printer.TextWidth("產假")
Printer.CurrentY = iLine * 300
Printer.Print "產假"
Printer.CurrentX = PLeft(16) - Printer.TextWidth("流產假")
Printer.CurrentY = iLine * 300
Printer.Print "流產假"
Printer.CurrentX = PLeft(17) - Printer.TextWidth("陪產假")
Printer.CurrentY = iLine * 300
Printer.Print "陪產假"
Printer.CurrentX = PLeft(18) - Printer.TextWidth("喪假")
Printer.CurrentY = iLine * 300
Printer.Print "喪假"
Printer.CurrentX = PLeft(19) - Printer.TextWidth("公傷假")
Printer.CurrentY = iLine * 300
Printer.Print "公傷假"
Printer.CurrentX = PLeft(20) - Printer.TextWidth("補休")
Printer.CurrentY = iLine * 300
Printer.Print "補休"
Printer.CurrentX = PLeft(21) - Printer.TextWidth("其他")
Printer.CurrentY = iLine * 300
Printer.Print "其他"
Printer.CurrentX = PLeft(22) - Printer.TextWidth("加班")
Printer.CurrentY = iLine * 300
Printer.Print "加班"
Printer.CurrentX = PLeft(23) - Printer.TextWidth("總計")
Printer.CurrentY = iLine * 300
Printer.Print "總計"

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(220, "-")

iLine = iLine + 1
End Sub

Sub GetPleft3()
PLeft(1) = 300
PLeft(2) = 900
PLeft(3) = 1300 + 900
PLeft(4) = 2050 + 750
PLeft(5) = 2800 + 750
PLeft(6) = 3550 + 750
PLeft(7) = 4300 + 750
PLeft(8) = 5050 + 850
PLeft(9) = 5800 + 750
PLeft(10) = 6550 + 850
PLeft(11) = 7300 + 750
PLeft(12) = 8050 + 850
PLeft(13) = 8800 + 750
PLeft(14) = 9550 + 850
PLeft(15) = 10300 + 750
PLeft(16) = 11050 + 850
PLeft(17) = 11800 + 850
PLeft(18) = 12550 + 750
PLeft(19) = 13300 + 850
PLeft(20) = 14050 + 750
PLeft(21) = 14800 + 750
PLeft(22) = 15550 + 750
'PLeft(23) = 16300
End Sub

Sub PrintDetail3()
Dim m_j As Integer, Item As Integer
   
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   '假別List:
   '01   忘打卡
   '02   遲到
   '03   曠職
   '04   出差
   '05   事假
   '22   家庭照顧假
   '06   病假
   '20   生理假
   '07   公假
   '08   特別假
   '09   婚假
   '21   產檢假
   '10   產假
   '11   流產假
   '19   陪產假
   '12   喪假
   '13   公傷假
   '14   補休
   '15   其他
   '16   加班
   For m_j = 3 To 22
      If m_j = 3 Then Item = 1
      If m_j = 4 Then Item = 2
      If m_j = 5 Then Item = 3
      If m_j = 6 Then Item = 4
      If m_j = 7 Then Item = 5
      If m_j = 8 Then Item = 22
      If m_j = 9 Then Item = 6
      If m_j = 10 Then Item = 20
      If m_j = 11 Then Item = 7
      If m_j = 12 Then Item = 8
      If m_j = 13 Then Item = 9
      If m_j = 14 Then Item = 21
      If m_j = 15 Then Item = 10
      If m_j = 16 Then Item = 11
      If m_j = 17 Then Item = 19
      If m_j = 18 Then Item = 12
      If m_j = 19 Then Item = 13
      If m_j = 20 Then Item = 14
      If m_j = 21 Then Item = 15
      If m_j = 22 Then Item = 16
      
      'Modify By Sindy 2025/11/17
      If m_j = 3 Or m_j = 4 Or m_j = 5 Then
         Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(Format(dblHour(Item), "##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblHour(Item), "##0")
      Else
      '2025/11/17 END
         Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(Format(dblHour(Item), "##0.0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblHour(Item), "##0.0")
      End If
   Next m_j
   iLine = iLine + 1
End Sub
