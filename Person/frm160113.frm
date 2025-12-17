VERSION 5.00
Begin VB.Form frm160113 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人出缺勤明細表"
   ClientHeight    =   3250
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   4700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3250
   ScaleWidth      =   4700
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   9
      Top             =   2640
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   10
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   11
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2370
      MaxLength       =   3
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1380
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   2370
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1380
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1260
      Width           =   615
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3585
      TabIndex        =   8
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblNote 
      Caption         =   "注意事項"
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   288
      TabIndex        =   14
      Top             =   2064
      Width           =   4332
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "                "
      Height          =   180
      Left            =   2064
      TabIndex        =   13
      Top             =   1308
      Width           =   768
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Top             =   1305
      Width           =   900
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "部門代號：                 －"
      Height          =   180
      Left            =   450
      TabIndex        =   6
      Top             =   900
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "出缺勤日期：                 －"
      Height          =   180
      Left            =   300
      TabIndex        =   5
      Top             =   1710
      Width           =   2025
   End
End
Attribute VB_Name = "frm160113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'Create by SINDY 2008/12/23
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_str2 As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 20) As String
'Dim PaperX As Double
'Dim paperY As Double
Dim iPgae As Integer, iLine As Integer
Dim LongPrintCurCnt As Long
Dim StrMenu2Cnt As Long ' Add By Sindy 98/03/06


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
'        If txt1(4) = "" Or txt1(5) = "" Then
'            MsgBox "出缺勤日期不可以空白！", vbInformation, "操作錯誤！"
'            txt1(4).SetFocus
'            Exit Sub
'        End If
        'Modified by Morgan 2023/12/27 取消員工號迄,日期改必輸且不可跨新部門啟用日
        'If txt1(0) = "" And txt1(1) = "" And _
        '    txt1(2) = "" And txt1(3) = "" And _
        '    txt1(4) = "" And txt1(5) = "" Then
        '    MsgBox "部門代號或員工代號或出缺勤日期至少輸入一項！", vbInformation, "操作錯誤！"
        '    txt1(0).SetFocus
        '    Exit Sub
        'End If
        If txt1(4) = "" Or txt1(5) = "" Then
            MsgBox "日期起迄不可空白！", vbInformation, "操作錯誤！"
            If txt1(4) = "" Then
               txt1(4).SetFocus
            Else
               txt1(5).SetFocus
            End If
            Exit Sub
        ElseIf DBDATE(txt1(4)) < 新部門啟用日 And DBDATE(txt1(5)) >= 新部門啟用日 Then
            MsgBox "日期起迄不可跨新部啟用日(" & (新部門啟用日 - 19110000) & ")，請重新輸入！", vbInformation, "操作錯誤！"
            txt1(4).SetFocus
            Exit Sub
        End If
        'end 2023/12/27
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        'Modified by Morgan 2023/12/27 以日期區間判斷新舊部門
        If txt1(0) <> "" Then
            If DBDATE(txt1(4)) >= 新部門啟用日 Then
               m_StrSQL = m_StrSQL & " and st93>='" & txt1(0) & "' "
            Else
               m_StrSQL = m_StrSQL & " and st03>='" & txt1(0) & "' "
            End If
        End If
        If txt1(1) <> "" Then
            If DBDATE(txt1(4)) >= 新部門啟用日 Then
               m_StrSQL = m_StrSQL & " and st93<='" & txt1(1) & "' "
            Else
               m_StrSQL = m_StrSQL & " and st03<='" & txt1(1) & "' "
            End If
        End If
        'end 2023/12/27
        
        'Modified by Morgan 2023/12/27
        'If txt1(2) <> "" Then
        '    m_StrSQL = m_StrSQL & " and st01>='" & txt1(2) & "' "
        'End If
        'If txt1(3) <> "" Then
        '    m_StrSQL = m_StrSQL & " and st01<='" & txt1(3) & "' "
        'End If
        If txt1(2) <> "" Then
            m_StrSQL = m_StrSQL & " and st01='" & txt1(2) & "' "
        End If
        'end 2023/12/27
        
        StrMenu1
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
End Select
End Sub

Sub StrMenu1()
Dim strSaSQL As String, strSbSQL As String, strSa2SQL As String

Set Printer = Printers(Combo1.ListIndex)

'XP自定紙張需手動設定並將印表機預設為該紙張
'9 x
'If pub_OS = "1" Then
'   Printer.Height = 2880
'   Printer.Width = 13000
'Else
'   Printer.PaperSize = PUB_GetPaperSize(3)
'End If
'Printer.Font = "@新細明體"
'Printer.FontSize = 12

Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印
'Printer.PaperSize = 39 '中一刀
Printer.PaperSize = 9  'PDF

'm_str = "SELECT ST01,ST02,a1.a0901||' '||a1.a0902,ST40 "
'Modified by Moran 2023/12/26 +新部門
'Modify By Sindy 2024/10/7 取消ST04='1'的限制條件: ST04='1' and 取消
If DBDATE(txt1(4)) >= 新部門啟用日 Then
   m_str = "SELECT ST01,ST02,st93||' '||a0922 " & _
               "FROM Staff,Acc090new,SalaryData " & _
               "WHERE ST01=SD01 and sd02 not in('P','F') " & _
               "and a0921(+)=st93" & m_StrSQL & _
               " ORDER BY st93,ST01 ASC "
Else
   m_str = "SELECT ST01,ST02,a1.a0901||' '||a1.a0902 " & _
               "FROM Staff,Acc090 a1,SalaryData " & _
               "WHERE ST01=SD01 and sd02 not in('P','F') " & _
               "and ST03=a1.a0901(+)" & m_StrSQL & _
               " ORDER BY ST03,ST01 ASC "
End If
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
LongPrintCurCnt = 0
StrMenu2Cnt = 0 ' Add By Sindy 98/03/06
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        Do While Not m_rs.EOF
            
            For m_i = 1 To 20
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields("ST01"))
            strTemp(2) = CheckStr(m_rs.Fields("ST02"))
            strTemp(3) = CheckStr(m_rs.Fields(2))
            'strTemp(4) = CheckStr(m_rs.Fields("ST40"))
            
            '組Where SQL
            strSaSQL = " and SA01='" & strTemp(1) & "' "
            strSbSQL = " and SB01='" & strTemp(1) & "' "
            strSa2SQL = " and SA01='" & strTemp(1) & "' "
            If txt1(4) <> "" Then
               strSaSQL = strSaSQL & " and SA02>='" & ChangeTStringToWString(txt1(4)) & "' "
               strSbSQL = strSbSQL & " and SB02>='" & ChangeTStringToWString(txt1(4)) & "' "
               strSa2SQL = strSa2SQL & " and SA02>='" & ChangeTStringToWString(txt1(4)) & "' "
            End If
            If txt1(5) <> "" Then
               strSaSQL = strSaSQL & " and SA04<='" & ChangeTStringToWString(txt1(5)) & "' "
               strSbSQL = strSbSQL & " and SB04<='" & ChangeTStringToWString(txt1(5)) & "' "
               strSa2SQL = strSa2SQL & " and SA02<='" & ChangeTStringToWString(txt1(5)) & "' "
            End If
                        
            '明細表
            Call StrMenu2(strSaSQL, strSbSQL, strSa2SQL)
            
            m_rs.MoveNext
        Loop
    End With
Else
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If
' Add By Sindy 98/03/06
If StrMenu2Cnt = 0 Then
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If
' 98/03/06 End
Printer.EndDoc
ShowPrintOk
End Sub

'明細表
Sub StrMenu2(strSaSQL As String, strSbSQL As String, strSa2SQL As String)
Dim dblTotDay As Double, dblTotHour As Double

m_str2 = "select * from ( " & _
"select sb01 as T1,'04 出差' as T2, " & _
"sqldateT(sb02)||' '||substr(ltrim(to_char('0000'||to_char(sb03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(sb03),'0000')),3,2)||' -- '|| " & _
"sqldateT(sb04)||' '||substr(ltrim(to_char('0000'||to_char(sb05),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(sb05),'0000')),3,2) as T3, " & _
"SB06 as T4,SB07 as T5,SB10 as T6,sb02 as T7,sb03 as T8 From staff_busi_trip Where 1=1 " & strSbSQL & " union all " & _
"select sa01 as T1,ac02||' '||ac03 as T2, " & _
"sqldateT(sa02)||' '||substr(ltrim(to_char('0000'||to_char(sa03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(sa03),'0000')),3,2)||' -- '|| " & _
"sqldateT(sa04)||' '||substr(ltrim(to_char('0000'||to_char(sa05),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(sa05),'0000')),3,2) as T3, " & _
"Sa07 as T4,Sa08 as T5,sa09 as T6,sa02 as T7,sa03 as T8 From staff_Absence,allcode Where 1=1 and (ac01='04' and sa06=ac02(+)) " & strSaSQL & " union all " & _
"select sa01 as T1,'01 忘打卡' as T2,sqldateT(sa02) as T3,0 as T4,sa03 as T5,'' as T6,sa02 as T7,0 as T8 From Staff_Assist Where sa03>'0' " & strSa2SQL & " union all " & _
"select sa01 as T1,'02 遲到' as T2,sqldateT(sa02) as T3,0 as T4,sa04 as T5,'' as T6,sa02 as T7,0 as T8 From Staff_Assist Where sa04>'0' " & strSa2SQL & " union all " & _
"select sa01 as T1,'03 曠職' as T2,sqldateT(sa02) as T3,sa05 as T4,sa06 as T5,'' as T6,sa02 as T7,0 as T8 From Staff_Assist Where (sa05>'0' or sa06>'0') " & strSa2SQL & _
") order by T7,T8 "
If m_rs2.State = 1 Then m_rs2.Close
m_rs2.CursorLocation = adUseClient
m_rs2.Open m_str2, cnnConnection, adOpenStatic, adLockReadOnly
dblTotDay = 0
dblTotHour = 0
If Not m_rs2.EOF And Not m_rs2.BOF Then
    With m_rs2
        m_rs2.MoveFirst
        
        If LongPrintCurCnt > 0 Then
            Printer.NewPage
        End If
        iLine = 1
        PrintTitle '列印表頭
        StrMenu2Cnt = StrMenu2Cnt + 1 ' Add By Sindy 98/03/06
        Do While Not m_rs2.EOF
            LongPrintCurCnt = LongPrintCurCnt + 1
            
            strTemp(5) = CheckStr(m_rs2.Fields("T1"))
            strTemp(6) = CheckStr(m_rs2.Fields("T2"))
            strTemp(7) = CheckStr(m_rs2.Fields("T3"))
            strTemp(8) = CheckStr(m_rs2.Fields("T4"))
            strTemp(9) = CheckStr(m_rs2.Fields("T5"))
            strTemp(10) = CheckStr(m_rs2.Fields("T6"))
            strTemp(11) = CheckStr(m_rs2.Fields("T7"))
            strTemp(12) = CheckStr(m_rs2.Fields("T8"))
            
            '假別01.忘打卡和02.遲到是以次數為單位
            If Left(Trim(strTemp(6)), 2) <> "01" And Left(Trim(strTemp(6)), 2) <> "02" Then
               '累計時數
               dblTotHour = dblTotHour + (Val(strTemp(8)) * 8) + strTemp(9)
            End If
            
            PrintDetail '列印表中
            
            If iLine >= 50 Then
                If .AbsolutePosition <> .RecordCount Then
                    Printer.NewPage
                    iLine = 1
                    PrintTitle '列印表頭
                End If
            End If
            m_rs2.MoveNext
        Loop
        '列印表尾
        If dblTotHour > 0 Then
            dblTotDay = (dblTotHour * 10) \ (8 * 10)
            dblTotHour = dblTotHour - (dblTotDay * 8)
            
            Printer.CurrentX = 500
            Printer.CurrentY = iLine * 300
            Printer.Print String(140, "-")
            
            iLine = iLine + 1
            Printer.CurrentX = 5000
            Printer.CurrentY = iLine * 300
            Printer.Print "小計："
            '小計-日
            Printer.CurrentX = PLeft(8) - Printer.TextWidth(dblTotDay)
            Printer.CurrentY = iLine * 300
            Printer.Print dblTotDay & "日"
            '小計-時
            Printer.CurrentX = PLeft(9) - Printer.TextWidth(dblTotHour)
            Printer.CurrentY = iLine * 300
            Printer.Print dblTotHour & "時"
        End If
    End With
'Else
'   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
'   Exit Sub
End If
End Sub

Sub PrintTitle()
GetPleft

'PaperX = 12000
'paperY = 7500

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("個人出缺勤明細表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "個人出缺勤明細表"

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "部　　門：" & strTemp(3)
If txt1(4) <> "" And txt1(5) <> "" Then
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(ChangeTStringToTDateString(txt1(4)) & " -- " & ChangeTStringToTDateString(txt1(5))) / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print ChangeTStringToTDateString(txt1(4)) & " -- " & ChangeTStringToTDateString(txt1(5))
End If
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "員工姓名：" & strTemp(1) & "　" & strTemp(2)

iLine = iLine + 2
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "假　別"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "起　迄　時　間"
Printer.CurrentX = PLeft(3) - Printer.TextWidth("日")
Printer.CurrentY = iLine * 300
Printer.Print "日"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("時(次)")
Printer.CurrentY = iLine * 300
Printer.Print "時(次)"
'Printer.CurrentX = PLeft(5)
'Printer.CurrentY = iLine * 300
'Printer.Print "職務代理人"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
'明細抬頭
PLeft(1) = 500
PLeft(2) = 2500
PLeft(3) = 7000
PLeft(4) = 8500
PLeft(5) = 9500
'明細內文
PLeft(6) = 500
PLeft(7) = 2500
PLeft(8) = 7000
PLeft(9) = 8500
PLeft(10) = 9500
End Sub

Sub PrintDetail()
   '假別
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iLine * 300
   If Left(Trim(strTemp(6)), 2) = "01" Or Left(Trim(strTemp(6)), 2) = "02" Then
      Printer.Print strTemp(6) & "(次)"
   Else
      Printer.Print strTemp(6)
   End If
   '起迄時間
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(7)
   '日
   Printer.CurrentX = PLeft(8) - Printer.TextWidth(strTemp(8))
   Printer.CurrentY = iLine * 300
   If Left(Trim(strTemp(6)), 2) <> "01" And Left(Trim(strTemp(6)), 2) <> "02" Then
      Printer.Print strTemp(8)
   End If
   '時(次)
   Printer.CurrentX = PLeft(9) - Printer.TextWidth(strTemp(9))
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(9)
'   '職務代理人
'   Printer.CurrentX = PLeft(10)
'   Printer.CurrentY = iLine * 300
'   Printer.Print strTemp(10)
   
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
   
'   InitialData

'Added by Morgan 2023/12/27
If strSrvDate(1) >= 新部門啟用日 Then
   lblNote = "1.出缺勤日期區間不可跨新部門啟用日(" & (新部門啟用日 - 19110000) & ")。" & vbCrLf & "2.日期區間為啟用日之前請輸入舊部門代碼，之後則輸入新部門代碼。"
Else
   lblNote = ""
End If
'end 2023/12/27
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160113 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case 4, 5
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         
      Case 2, 3
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            'Added by Morgan 2023/12/27
            ElseIf ClsPDGetStaffN(txt1(Index), strExc(1)) = True Then
               lblName = strExc(1)
            'end 2023/12/27
            End If
         End If
         'Removed by Morgan 2023/12/27
         'If Index = 2 Then
         '   If txt1(Index) <> "" And txt1(Index + 1) = "" Then
         '      txt1(Index + 1) = txt1(Index)
         '   End If
         'ElseIf Index = 3 Then
         '   If RunNick(txt1(Index - 1), txt1(Index)) Then
         '      Call txt1_GotFocus(Index)
         '      Cancel = True
         '      Exit Sub
         '   End If
         'End If
         'end 2023/12/27
         
      Case 4, 5
         If CheckIsTaiwanDate(txt1(Index), False) = False And Trim(txt1(Index)) <> "" Then
            Call txt1_GotFocus(Index)
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
         If Index = 4 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 5 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub

