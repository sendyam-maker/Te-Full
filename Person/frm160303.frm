VERSION 5.00
Begin VB.Form frm160303 
   BorderStyle     =   1  '單線固定
   Caption         =   "年終考績考核"
   ClientHeight    =   3620
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3620
   ScaleWidth      =   4990
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   1290
      MaxLength       =   72
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frm160303.frx":0000
      Top             =   1950
      Width           =   3405
   End
   Begin VB.CheckBox Check1 
      Caption         =   "不須算考勤加減分"
      Height          =   225
      Left            =   1110
      TabIndex        =   5
      Top             =   1650
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   10
      Top             =   3000
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   7
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
      Index           =   4
      Left            =   2100
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1260
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   1290
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1260
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1950
      MaxLength       =   3
      TabIndex        =   2
      Top             =   900
      Width           =   555
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1290
      MaxLength       =   3
      TabIndex        =   1
      Top             =   900
      Width           =   555
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4005
      TabIndex        =   9
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3060
      TabIndex        =   8
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1290
      MaxLength       =   3
      TabIndex        =   0
      Top             =   540
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "備註：請自行修改預設的呈閱日期"
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   16
      Top             =   2460
      Width           =   3060
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "呈閱日期："
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   15
      Top             =   1980
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2010
      X2              =   2250
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Line Line1 
      X1              =   1860
      X2              =   2250
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   360
      TabIndex        =   13
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "年度："
      Height          =   180
      Left            =   720
      TabIndex        =   12
      Top             =   570
      Width           =   540
   End
End
Attribute VB_Name = "frm160303"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strStarDate As String, strEndDate As String
'每行往下移多少
Const CustMove = 580
'起始位置
Const StartMove = 700


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
        If Trim(txt1(0)) = "" Then
            MsgBox "年度不可空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        If Trim(txt1(1)) & Trim(txt1(2)) <> "" Then
            If Trim(txt1(1)) = "" Then txt1(1).SetFocus: MsgBox "請輸入起始部門代號！", vbInformation, "操作錯誤！": Exit Sub
            If Trim(txt1(2)) = "" Then txt1(2).SetFocus: MsgBox "請輸入終止部門代號！", vbInformation, "操作錯誤！": Exit Sub
        End If
        If Trim(txt1(3)) & Trim(txt1(4)) <> "" Then
            If Trim(txt1(3)) = "" Then txt1(3).SetFocus: MsgBox "請輸入起始員工編號！", vbInformation, "操作錯誤！": Exit Sub
            If Trim(txt1(4)) = "" Then txt1(4).SetFocus: MsgBox "請輸入終止員工編號！", vbInformation, "操作錯誤！": Exit Sub
        End If
        If Trim(txt1(5)) = "" Then
            MsgBox "呈閱日期不可空白！", vbInformation, "操作錯誤！"
            txt1(5).SetFocus
            Exit Sub
        End If
        
        strStarDate = Val(txt1(0)) + 1911 & "0101"
        strEndDate = Val(txt1(0)) + 1911 & "1231"
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(1) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
            m_StrSQL = m_StrSQL & " and st93 >='" & txt1(1) & "' "
        End If
        If txt1(2) <> "" Then
            'Modify By Sindy 2023/12/27 部門調整改抓ST93
            m_StrSQL = m_StrSQL & " and st93 <='" & txt1(2) & "' "
        End If
        If txt1(3) <> "" Then
            m_StrSQL = m_StrSQL & " and st01 >='" & txt1(3) & "' "
        End If
        If txt1(4) <> "" Then
            m_StrSQL = m_StrSQL & " and st01 <='" & txt1(4) & "' "
        End If
        StrMenu1
         'Add By Sindy 2018/1/3
         PUB_SaveLastDate Me.Name, "txt1(5)", txt1(5)
         '2018/1/3 END
        Screen.MousePointer = vbDefault
      Case 1
        Unload Me
   End Select
End Sub

Sub StrMenu1()
Dim intCurrRow As Integer

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   'Printer.PaperSize = 9  'PDF
   intCurrRow = 0
   '2011/12/21 ADD BY SONIA
   'Modify By Sindy 2022/7/18 + 15.名譽所長
   m_StrSQL = m_StrSQL & " AND (ST20 IS NULL OR ST20 NOT IN ('01','02','11','12','13','21','22','31','15')) AND ST03<>'R08' "
   '2011/12/21 END
   
   'modify by sonia 2016/1/6 104/7/27劉經理通知修改辦法
   '1.每年10/2(含)以後到職者不得參加年終考績,不印年終考績考核表
   '2.留職停薪者仍要參加年終考績故要印年終考績考核表
   'm_str = "select st03,a0902,st01,st02,st13,st20,st21,a1.ac03 as T1,a2.ac03 as T2 " & _
            "From staff, SalaryData, acc090,allcode a1,allcode a2 " & _
            "where ST04='1' and ST01=SD01 and (sd02 not in('P','F') or sd02 is null) " & _
            "and st03=a0901(+) " & _
            "and a1.ac01(+)='01' and st20=a1.ac02(+) " & _
            "and a2.ac01(+)='02' and st21=a2.ac02(+) " & m_StrSQL & _
            "Order By st03,st01 "
   
   '在職且到職日<當年10/2
   'Modify By Sindy 2023/12/27 部門調整改抓ST93
   'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
   m_str = "select st93,a0922,st01,st02,st13,st20,st21,a1.ac03 as T1,a2.ac03 as T2 " & _
            "From staff, SalaryData, acc090NEW,allcode a1,allcode a2 " & _
            "where ST04='1' and ST01=SD01 and (sd02 not in('P','F') or sd02 is null) and not(substr(st01,5,1)>='A') " & _
            "and st93=a0921(+) and nvl(st13,0)<" & Val(txt1(0)) + 1911 & "1002 " & _
            "and a1.ac01(+)='01' and st20=a1.ac02(+) and a2.ac01(+)='02' and st21=a2.ac02(+) " & m_StrSQL
   '再抓當時留職停薪人員
   'Modify By Sindy 2023/12/27 部門調整改抓ST93
   'Modify By Sindy 2024/4/30 + and not(substr(st01,5,1)>='A') 排除 B309A=宗家澔
   m_str = m_str & "union select st93,a0922,st01,st02,st13,st20,st21,a1.ac03 as T1,a2.ac03 as T2 " & _
            "From staff, SalaryData, acc090NEW,allcode a1,allcode a2, " & _
            "(select sc01 from staff_change where (sc01,sc02) in (select sc01,max(sc02) from staff_change group by sc01) and sc03='04') " & _
            "where ST04<>'1' and ST01=SD01 and (sd02 not in('P','F') or sd02 is null) and not(substr(st01,5,1)>='A') " & _
            "and st93=a0921(+) and nvl(st13,0)<" & Val(txt1(0)) + 1911 & "1002 " & _
            "and a1.ac01(+)='01' and st20=a1.ac02(+) and a2.ac01(+)='02' and st21=a2.ac02(+) " & m_StrSQL & " and st01=sc01 "
   '排序
   m_str = m_str & " Order By st93,st01 "
   'end 2016/1/6
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       With m_rs
           m_rs.MoveFirst
           
           Do While Not m_rs.EOF
               intCurrRow = intCurrRow + 1
               
               For m_i = 1 To 9
                   strTemp(m_i) = ""
               Next m_i
               
               strTemp(1) = CheckStr(m_rs.Fields("st01")) & "　" & CheckStr(m_rs.Fields("st02")) '姓名
               'Modify By Sindy 2023/12/27 部門調整改抓ST93
               strTemp(2) = CheckStr(m_rs.Fields("a0922")) '部門
               strTemp(3) = CheckStr(m_rs.Fields("T1")) '職稱
               strTemp(4) = CheckStr(m_rs.Fields("T2")) '職位
               '前三年考績
               'MODIFY BY SONIA 2015/12/25 考績檔YM02加入*不參加考核
               strSql = "select YM01,decode(YM02,'1','優','2','甲','3','乙','4','丙','*','不參加考核',YM02)" & _
                              " From YearMerit" & _
                              " Where YM01>=" & (Val(txt1(0)) + 1911) - 3 & " And YM01<=" & (Val(txt1(0)) + 1911) - 1 & _
                              " and YM03='" & m_rs.Fields("st01") & "'" & _
                              " order by YM01 asc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  With RsTemp
                     .MoveFirst
                     Do While Not .EOF
                        If strTemp(5) <> "" Then strTemp(5) = strTemp(5) & "、"
                        strTemp(5) = strTemp(5) & "(" & Val(.Fields(0)) - 1911 & ")" & .Fields(1)
                        .MoveNext
                     Loop
                  End With
               End If
               If Check1.Value = vbUnchecked Then
                  '(考勤類別 : 0.全部 1.遲到 2.曠職 3.全勤)
                  strTemp(6) = GetAssistAbsenceGrade(CheckStr(m_rs.Fields("st01")), strStarDate, strEndDate, 1) '遲到
                  strTemp(7) = GetAssistAbsenceGrade(CheckStr(m_rs.Fields("st01")), strStarDate, strEndDate, 2) '曠職
                  strTemp(8) = GetAssistAbsenceGrade(CheckStr(m_rs.Fields("st01")), strStarDate, strEndDate, 3) '全勤
                  strTemp(9) = GetRewardGrade(CheckStr(m_rs.Fields("st01")), strStarDate, strEndDate) '獎懲
               End If
               
               If intCurrRow > 1 Then Printer.NewPage
               iLine = 1
               '表格種類：
               '年終考績考核表-智權部人員
               'Modify By Sindy 2023/12/27 部門調整改抓ST93
               If m_rs.Fields("st93") >= "S10" And m_rs.Fields("st93") <= "S99" Then
                  Call PrintTitle("2")  '表頭
                  Call PrintDetail2 '表中
               '年終考績考核表-一般人員
               Else
                  Call PrintTitle("1")  '表頭
                  Call PrintDetail '表中
               End If
               
               m_rs.MoveNext
           Loop
       End With
   Else
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   Printer.EndDoc
   ShowPrintOk

End Sub

Sub PrintTitle(strRptType As String)
   Printer.Font.Size = 20
   Printer.Font.Underline = True
   Printer.FontBold = True
   Printer.FontName = "標楷體"
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(txt1(0) & " 年 度 年 終 考 績 考 核 表") / 2)
   Printer.CurrentY = iLine * CustMove
   Printer.Print txt1(0) & " 年 度 年 終 考 績 考 核 表"
   
   Printer.Font.Size = 14
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   'iLine = iLine + 1
   'Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   'Printer.CurrentY = iLine * CustMove
   'Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   'iLine = iLine + 1
   'Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(txt1(0) & "年度") / 2)
   'Printer.CurrentY = iLine * 450
   'Printer.Print txt1(0) & "年度"
   'Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
   'Printer.CurrentY = iLine * 450
   'Printer.Print "頁　　次：" & Printer.Page
   
   iLine = iLine + 1
   Printer.CurrentX = 1000
   Printer.CurrentY = iLine * CustMove
   Printer.Print "姓　名：" & strTemp(1)
   Printer.CurrentX = 6500
   Printer.CurrentY = iLine * CustMove
   Printer.Print "部　門：" & strTemp(2)
   iLine = iLine + 1
   Printer.CurrentX = 1000
   Printer.CurrentY = iLine * CustMove
   Printer.Print "職　稱：" & strTemp(3)
   Printer.CurrentX = 6500
   Printer.CurrentY = iLine * CustMove
   Printer.Print "職　位：" & strTemp(4)
   iLine = iLine + 1
   Printer.CurrentX = 1000
   Printer.CurrentY = iLine * CustMove
   Printer.Print "前三年考績：" & strTemp(5)
   Printer.CurrentX = 8000
   Printer.CurrentY = iLine * CustMove
   If strRptType = "1" Then
      Printer.Print "適用人員：一般人員"
   Else
      Printer.Print "適用人員：智權部人員"
   End If
   
   iLine = iLine + 1
End Sub

''年終考績考核表-一般人員
Sub PrintDetail()
Dim i As Integer
   
   Printer.Font.Size = 14
   Printer.DrawWidth = 4
   
   For i = 1 To 14
      If i = 1 Then
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 14) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(2500, CustMove * (iLine + 14) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(3500, CustMove * (iLine + 14) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(8000, CustMove * (iLine + 14) + CustMove), , B
         Printer.Line (8000, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
      End If
      If i = 2 Then
         Printer.Line (8000, CustMove * iLine)-(8812, CustMove * (iLine + 13) + CustMove), , B
         Printer.Line (8000, CustMove * iLine)-(9624, CustMove * (iLine + 13) + CustMove), , B
         Printer.Line (8000, CustMove * iLine)-(10436, CustMove * (iLine + 13) + CustMove), , B
      End If
      If i = 1 Then
         Printer.CurrentX = 8500
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "各 級 主 管 考 核"
      End If
      If i = 2 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "考 核 項 目"
         Printer.CurrentX = 2600
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "分 數"
         Printer.CurrentX = 4900
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "考 核 內 容"
      End If
      If i = 3 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "工 作 績 效"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "50"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove
         Printer.Print "１當年度工作目標達成率"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 300
         Printer.Print "２當年度工作目標成長率"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 600
         Printer.Print "３工作效率"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 900
         Printer.Print "４工作品質"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 1200
         Printer.Print "５工作正確性"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 2) + CustMove), , B
      End If
      If i = 6 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "工 作 態 度"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "10"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove
         Printer.Print "１對工作的興趣與熱忱"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 300
         Printer.Print "２對所承辦工作能悉心研究並提供"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 600
         Printer.Print "　意見"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 900
         Printer.Print "３能遵循上級所訂政策並執行"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
      End If
      If i = 8 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "協 調 合 作"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "10"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove
         Printer.Print "１服從指揮，遇事能適當回應，並"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 300
         Printer.Print "　適時向主管反應"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 600
         Printer.Print "２與同事相處融洽，與別人配合均"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 900
         Printer.Print "　能和睦協調"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
      End If
      If i = 10 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "品 德 操 守"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "10"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove
         Printer.Print "１處事剛正不阿，不接受饋贈及不"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 300
         Printer.Print "　正當招待"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 600
         Printer.Print "２潔身自愛，以本所及客戶的利益"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 900
         Printer.Print "　為考量"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
      End If
      If i = 12 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "學 識 才 能"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "10"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove
         Printer.Print "１具備足夠專業知識，並能掌握工"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 300
         Printer.Print "　作要點"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 600
         Printer.Print "２積極進取，隨時充實自己，力求"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 900
         Printer.Print "　成長"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
      End If
      If i = 14 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "創 意 能 力"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 120
         Printer.Print "10"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove
         Printer.Print "１以創新及有效率的方式解決問題"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 300
         Printer.Print "２提出改進方案以提高效率，簡化"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove) + 600
         Printer.Print "　流程"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
      End If
      iLine = iLine + 1
   Next i
   
   iLine = iLine + 1
   Printer.Line (StartMove, CustMove * iLine)-(3500, CustMove * (iLine + 1) + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
   Printer.CurrentX = 800
   Printer.CurrentY = iLine * CustMove + 120
   Printer.Print "考　　　勤"
   Printer.CurrentX = 3620
   Printer.CurrentY = iLine * CustMove + 100
   Printer.Print "遲　　到　" & strTemp(6)
   Printer.CurrentX = 3620
   Printer.CurrentY = (iLine * CustMove + 100) + 300
   Printer.Print "曠　　職　" & strTemp(7)
   Printer.CurrentX = 3620
   Printer.CurrentY = (iLine * CustMove + 100) + 600
   Printer.Print "全　　勤　" & strTemp(8)
   iLine = iLine + 2
   Printer.Line (StartMove, CustMove * iLine)-(2500, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
   Printer.CurrentX = 800
   Printer.CurrentY = iLine * CustMove + 120
   Printer.Print "獎　　　懲"
   Printer.CurrentX = 3620
   Printer.CurrentY = iLine * CustMove + 100
   Printer.Print "　　　　　" & strTemp(9)
   
   iLine = iLine + 1
   Printer.Line (StartMove, CustMove * iLine)-(8000, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(8812, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(9624, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(10436, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
   Printer.CurrentX = 800
   Printer.CurrentY = iLine * CustMove + 120
   Printer.Print "總　　　分"
   
   iLine = iLine + 1
   Printer.Line (StartMove, CustMove * iLine)-(2500, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(3500, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(7000, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(8000, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
   Printer.CurrentX = 800
   Printer.CurrentY = iLine * CustMove + 120
   Printer.Print "等　　　級"
   Printer.CurrentX = 2600
   Printer.CurrentY = iLine * CustMove + 120
   Printer.Print "初 核"
   Printer.CurrentX = 7100
   Printer.CurrentY = iLine * CustMove + 120
   Printer.Print "核 定"
   
   iLine = iLine + 1
'   Printer.Line (StartMove, CustMove * iLine)-(2500, CustMove * (iLine + 1) + CustMove), , B
'   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(2500, CustMove * (iLine + 1) + CustMove - 200), , B
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove - 200), , B
   Printer.CurrentX = 800
   Printer.CurrentY = iLine * CustMove + 120
   Printer.Print "主 管 評 語"
   
   iLine = iLine + 2
   Printer.CurrentX = StartMove
   'Printer.CurrentY = iLine * CustMove + 120
   Printer.CurrentY = iLine * CustMove + 120 - 200
   Printer.Print "※本表請按下列日期呈交各級主管：" & Trim(txt1(5).Text)
'   iLine = iLine + 1
'   Printer.CurrentX = StartMove
'   Printer.CurrentY = iLine * CustMove
'   Printer.Print Trim(txt1(5).Text)
End Sub

'年終考績考核表-智權部人員
Sub PrintDetail2()
Dim i As Integer
   
   Printer.Font.Size = 14
   Printer.DrawWidth = 4
   
   For i = 1 To 14
      If i = 1 Then
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 14) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(2500, CustMove * (iLine + 14) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(3500, CustMove * (iLine + 14) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(8000, CustMove * (iLine + 14) + CustMove), , B
         Printer.Line (8000, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
      End If
      If i = 2 Then
         Printer.Line (8000, CustMove * iLine)-(8812, CustMove * (iLine + 13) + CustMove), , B
         Printer.Line (8000, CustMove * iLine)-(9624, CustMove * (iLine + 13) + CustMove), , B
         Printer.Line (8000, CustMove * iLine)-(10436, CustMove * (iLine + 13) + CustMove), , B
      End If
      If i = 1 Then
         Printer.CurrentX = 8500
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "各 級 主 管 考 核"
      End If
      If i = 2 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "考 核 項 目"
         Printer.CurrentX = 2600
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "分 數"
         Printer.CurrentX = 4900
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "考 核 內 容"
      End If
      If i = 3 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "業績達成率"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "50"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "１當年度業績目標達成率"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 150) + 300
         Printer.Print "２各類業務目標達成率"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 150) + 600
         Printer.Print "３業績目標成長率"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 150) + 900
         Printer.Print "４對業務拓展的積極度"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 2) + CustMove), , B
      End If
      If i = 6 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "客戶滿意度"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "10"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "１客戶服務品質"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 150) + 300
         Printer.Print "２客戶對該員的回應狀況"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
      End If
      If i = 8 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "工 作 態 度"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "10"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove + 50
         Printer.Print "１平時出勤狀況"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 50) + 300
         Printer.Print "２主動積極開發客戶"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 50) + 600
         Printer.Print "３能遵循上級所訂政策並執行"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
      End If
      If i = 10 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "協 調 合 作"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "10"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove + 10
         Printer.Print "１服從指揮，遇事能適當回應，並"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 10) + 300
         Printer.Print "　適時向主管反應"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 10) + 600
         Printer.Print "２與同事相處融洽，與別人配合均"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 10) + 900
         Printer.Print "　能和睦協調"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
      End If
      If i = 12 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "品 德 操 守"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "10"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove + 10
         Printer.Print "１處事剛正不阿，不接受饋贈及不"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 10) + 300
         Printer.Print "　正當招待"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 10) + 600
         Printer.Print "２潔身自愛，以本所及客戶的利益"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 10) + 900
         Printer.Print "　為考量"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
      End If
      If i = 14 Then
         Printer.CurrentX = 800
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "學 識 才 能"
         Printer.CurrentX = 2850
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "10"
         Printer.CurrentX = 3620
         Printer.CurrentY = iLine * CustMove + 10
         Printer.Print "１具備足夠專業知識，並能掌握工"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 10) + 300
         Printer.Print "　作要點"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 10) + 600
         Printer.Print "２積極進取，隨時充實自己，力求"
         Printer.CurrentX = 3620
         Printer.CurrentY = (iLine * CustMove + 10) + 900
         Printer.Print "　成長"
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
      End If
      iLine = iLine + 1
   Next i
   
   iLine = iLine + 1
   Printer.Line (StartMove, CustMove * iLine)-(3500, CustMove * (iLine + 1) + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
   Printer.CurrentX = 800
   Printer.CurrentY = iLine * CustMove + 150
   Printer.Print "考　　　勤"
   Printer.CurrentX = 3620
   Printer.CurrentY = iLine * CustMove + 100
   Printer.Print "遲　　到　" & strTemp(6)
   Printer.CurrentX = 3620
   Printer.CurrentY = (iLine * CustMove + 100) + 300
   Printer.Print "曠　　職　" & strTemp(7)
   Printer.CurrentX = 3620
   Printer.CurrentY = (iLine * CustMove + 100) + 600
   Printer.Print "全　　勤　" & strTemp(8)
   iLine = iLine + 2
   Printer.Line (StartMove, CustMove * iLine)-(2500, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
   Printer.CurrentX = 800
   Printer.CurrentY = iLine * CustMove + 150
   Printer.Print "獎　　　懲"
   Printer.CurrentX = 3620
   Printer.CurrentY = iLine * CustMove + 100
   Printer.Print "　　　　　" & strTemp(9)
   
   iLine = iLine + 1
   Printer.Line (StartMove, CustMove * iLine)-(8000, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(8812, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(9624, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(10436, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
   Printer.CurrentX = 800
   Printer.CurrentY = iLine * CustMove + 150
   Printer.Print "總　　　分"
   
   iLine = iLine + 1
   Printer.Line (StartMove, CustMove * iLine)-(2500, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(3500, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(7000, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(8000, CustMove * iLine + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
   Printer.CurrentX = 800
   Printer.CurrentY = iLine * CustMove + 150
   Printer.Print "等　　　級"
   Printer.CurrentX = 2600
   Printer.CurrentY = iLine * CustMove + 150
   Printer.Print "初 核"
   Printer.CurrentX = 7100
   Printer.CurrentY = iLine * CustMove + 150
   Printer.Print "核 定"
   
   iLine = iLine + 1
'   Printer.Line (StartMove, CustMove * iLine)-(2500, CustMove * (iLine + 1) + CustMove), , B
'   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
   Printer.Line (StartMove, CustMove * iLine)-(2500, CustMove * (iLine + 1) + CustMove - 200), , B
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove - 200), , B
   Printer.CurrentX = 800
   Printer.CurrentY = iLine * CustMove + 150
   Printer.Print "主 管 評 語"
   
   iLine = iLine + 2
   Printer.CurrentX = StartMove
'   Printer.CurrentY = iLine * CustMove + 150
   Printer.CurrentY = iLine * CustMove + 150 - 200
   Printer.Print "※本表請按下列日期呈交各級主管：" & Trim(txt1(5).Text)
'   iLine = iLine + 1
'   Printer.CurrentX = StartMove
'   Printer.CurrentY = iLine * CustMove
'   Printer.Print Trim(txt1(5).Text)
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
   
   txt1(5) = PUB_GetLastDate(Me.Name, "txt1(5)") 'Add By Sindy 2018/1/2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160303 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 1, 2, 3, 4, 5
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
      Case 3, 4
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
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
      Case Else
   End Select
End Sub

'年終考績考核表-舊格式
Sub PrintDetail3()
Dim i As Integer
   
   Printer.Font.Size = 12
   Printer.DrawWidth = 4
   '列框
   For i = 1 To 5
      If i = 1 Then
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 4) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(1300, CustMove * (iLine + 4) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(3250, CustMove * (iLine + 4) + CustMove), , B
         Printer.Line (3250, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
      End If
      If i = 2 Then
         Printer.Line (3250, CustMove * iLine)-(5250, CustMove * (iLine + 3) + CustMove), , B
         Printer.Line (3250, CustMove * iLine)-(7250, CustMove * (iLine + 3) + CustMove), , B
         Printer.Line (3250, CustMove * iLine)-(9250, CustMove * (iLine + 3) + CustMove), , B
      End If
      If i = 3 Or i = 4 Then
         Printer.Line (1300, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
      End If
      Printer.CurrentX = 650
      Printer.CurrentY = iLine * CustMove + 150
      If i = 1 Then
         Printer.Print "工"
         Printer.CurrentX = 5000
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "各　　　級　　　主　　　管　　　考　　　核"
      End If
      If i = 2 Then
         Printer.Print "作"
         Printer.CurrentX = 1500
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "考核　　 項目"
      End If
      If i = 3 Then
         Printer.Print "評"
         Printer.CurrentX = 1500
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "工作表現　50%"
      End If
      If i = 4 Then
         Printer.Print "分"
         Printer.CurrentX = 1500
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "協調同仁　25%"
      End If
      If i = 5 Then
         Printer.Print "100%"
         Printer.CurrentX = 1500
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "配合主管　25%"
      End If
      iLine = iLine + 1
   Next i
   For i = 1 To 3
      If i = 1 Then
         Printer.Line (StartMove, CustMove * iLine)-(1300, CustMove * (iLine + 2) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(3250, CustMove * (iLine + 2) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 2) + CustMove), , B
      End If
      Printer.CurrentX = 650
      Printer.CurrentY = iLine * CustMove + 150
      If i = 1 Then
         Printer.Print "考"
         Printer.CurrentX = 1500
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "　遲　　　到　"
         Printer.CurrentX = 4000 - Printer.TextWidth(strTemp(5))
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print strTemp(5)
      End If
      If i = 2 Then
         Printer.CurrentX = 1500
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "　曠　　　職　"
         Printer.CurrentX = 4000 - Printer.TextWidth(strTemp(6))
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print strTemp(6)
      End If
      If i = 3 Then
         Printer.Print "勤"
         Printer.CurrentX = 1500
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "　全　　　勤　"
         Printer.CurrentX = 4000 - Printer.TextWidth(strTemp(7))
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print strTemp(7)
      End If
      iLine = iLine + 1
   Next i
   For i = 1 To 2
      Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
      If i = 2 Then
         Printer.Line (StartMove, CustMove * iLine)-(3250, CustMove * iLine + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(5250, CustMove * iLine + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(7250, CustMove * iLine + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(9250, CustMove * iLine + CustMove), , B
      End If
      Printer.CurrentX = 650
      Printer.CurrentY = iLine * CustMove + 150
      If i = 1 Then
         Printer.Print "獎"
         Printer.CurrentX = 2750 - 50
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "懲"
         Printer.CurrentX = 4000 - Printer.TextWidth(strTemp(8))
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print strTemp(8)
      End If
      If i = 2 Then
         Printer.Print "總"
         Printer.CurrentX = 2750 - 50
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "分"
      End If
      iLine = iLine + 1
   Next i
   For i = 1 To 2
      If i = 1 Then
         Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(1300, CustMove * (iLine + 1) + CustMove), , B
         Printer.Line (StartMove, CustMove * iLine)-(3250, CustMove * (iLine + 1) + CustMove), , B
      ElseIf i = 2 Then
         Printer.Line (1300, CustMove * iLine)-(11250, CustMove * iLine + CustMove), , B
      End If
      Printer.CurrentX = 650
      Printer.CurrentY = iLine * CustMove + 150
      If i = 1 Then
         Printer.Print "等"
         Printer.CurrentX = 1500
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "　初　　　核　"
      End If
      If i = 2 Then
         Printer.Print "級"
         Printer.CurrentX = 1500
         Printer.CurrentY = iLine * CustMove + 150
         Printer.Print "　核　　　定　"
      End If
      iLine = iLine + 1
   Next i
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 1) + CustMove), , B
   Printer.CurrentX = 4500
   Printer.CurrentY = iLine * CustMove + 150
   Printer.Print "主　　　管　　　評　　　語"
   iLine = iLine + 1
   Printer.CurrentX = 3500
   Printer.CurrentY = iLine * CustMove + 150
   Printer.Print "(　參酌　1. 敬業精神　2. 求知意願　3. 特殊專長　)"
   iLine = iLine + 1
   Printer.Line (StartMove, CustMove * iLine)-(11250, CustMove * (iLine + 6) + CustMove), , B
   iLine = (iLine + 6) + 1
   Printer.CurrentX = StartMove
   Printer.CurrentY = iLine * CustMove + 150
   Printer.Print "※本表請按下列日期呈交各級主管作業："
   iLine = iLine + 1
   Printer.CurrentX = StartMove
   Printer.CurrentY = iLine * CustMove
   Printer.Print Trim(txt1(5).Text)
End Sub
