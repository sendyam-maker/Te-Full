VERSION 5.00
Begin VB.Form frm170217 
   BorderStyle     =   1  '單線固定
   Caption         =   "年終獎金發放明細"
   ClientHeight    =   3120
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4752
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4752
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   3
      Left            =   2130
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "1"
      Top             =   1530
      Width           =   285
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   9
      Top             =   2460
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   4
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
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2610
      TabIndex        =   5
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3630
      TabIndex        =   6
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   2130
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   435
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   2130
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1200
      Width           =   765
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   2
      Left            =   3030
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "PS：紙張為A4橫印"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   1470
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "報表格式：        (1.依部門  2.依公司)"
      Height          =   180
      Left            =   1170
      TabIndex        =   11
      Top             =   1590
      Width           =   2820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "獎金年度："
      Height          =   180
      Left            =   1170
      TabIndex        =   8
      Top             =   870
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   1170
      TabIndex        =   7
      Top             =   1230
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   3300
      Y1              =   1290
      Y2              =   1290
   End
End
Attribute VB_Name = "frm170217"
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
Dim PLeft(1 To 15) As Integer
Dim strTemp(1 To 15) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmt1 As Double, dblAmt2 As Double, dblAmt3 As Double, dblAmt4 As Double, dblAmt5 As Double, dblAmt6 As Double, dblAmt7 As Double, dblAmt8 As Double, dblAmt9 As Double, dblAmt10 As Double
Dim dblTotAmt1 As Double, dblTotAmt2 As Double, dblTotAmt3 As Double, dblTotAmt4 As Double, dblTotAmt5 As Double, dblTotAmt6 As Double, dblTotAmt7 As Double, dblTotAmt8 As Double, dblTotAmt9 As Double, dblTotAmt10 As Double

Private Sub cmdok_Click(Index As Integer)
Dim strYM As String
Select Case Index
Case 0
        If txt1(0) = "" Then
            MsgBox "獎金年度不可空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        If txt1(3) = "" Then
            MsgBox "報表格式不可空白！", vbInformation, "操作錯誤！"
            txt1(3).SetFocus
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            strYM = Left(ChangeTStringToWString(txt1(0) & "0101"), 4)
            m_StrSQL = m_StrSQL & " yb01='" & strYM & "' "
        End If
        If txt1(1) <> "" Then
            'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
            'm_StrSQL = m_StrSQL & " and replace(yb02,'A','0') >= '" & txt1(1) & "' "
            'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
            m_StrSQL = m_StrSQL & " and substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4) >= '" & txt1(1) & "' "
        End If
        If txt1(2) <> "" Then
            'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
            'm_StrSQL = m_StrSQL & " and replace(yb02,'A','0') <= '" & txt1(2) & "' "
            'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
            m_StrSQL = m_StrSQL & " and substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4) <= '" & txt1(2) & "' "
        End If
        If txt1(3) = "1" Then '依部門
            StrMenu2
        ElseIf txt1(3) = "2" Then '依公司
            StrMenu1
        End If
        Printer.Font.Size = 12    'add by sonia 2018/1/15 還原字大小
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
End Select
End Sub

Sub StrMenu1()

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF
'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
'2013/1/21 MODIFY BY SONIA 加入YB25代扣補充保費
'modify by sonia 2018/1/15 +YB26
'modify by sonia 2018/1/30 婧瑄說應領不可扣除借支
'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
m_str = "select T.yb24,a0802,T.yb02,ST02,T4,T5,T6,T8,T15,T5+T6+T26+T8-T15,T16,T17,T18,T5+T6+T26+T8-T15-T17-T18-T16,T26 " & _
              "from staff,acc080, " & _
            "(select yb01,yb24,yb02,yb03,substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4) as T2,nvl(yb04,0) as T4,nvl(yb05,0) as T5,nvl(yb06,0) as T6,nvl(yb08,0) as T8,nvl(yb15,0) as T15,nvl(yb16,0) as T16,nvl(yb17,0) as T17,nvl(yb25,0) as T18,nvl(yb26,0) as T26 " & _
             "From yearbonus " & _
             "where " & m_StrSQL & ") T " & _
             "where T2=st01(+) " & _
             "and T.yb24=a0801(+) " & _
             "order by yb24,yb03,yb02 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        '預設值
        iLine = 1
        strType = "" '切頁條件
        dblAmt1 = 0: dblAmt2 = 0: dblAmt3 = 0: dblAmt4 = 0: dblAmt5 = 0
        dblAmt6 = 0: dblAmt7 = 0: dblAmt8 = 0: dblAmt9 = 0: dblAmt10 = 0
        dblTotAmt1 = 0: dblTotAmt2 = 0: dblTotAmt3 = 0: dblTotAmt4 = 0: dblTotAmt5 = 0
        dblTotAmt6 = 0: dblTotAmt7 = 0: dblTotAmt8 = 0: dblTotAmt9 = 0: dblTotAmt10 = 0
        
        Do While Not m_rs.EOF
            
            For m_i = 1 To 15
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(0)) '公司別
            strTemp(2) = CheckStr(m_rs.Fields(1))
            strTemp(3) = CheckStr(m_rs.Fields(2)) '編號
            strTemp(4) = CheckStr(m_rs.Fields(3))
            strTemp(5) = CheckStr(m_rs.Fields(4))
            strTemp(6) = CheckStr(m_rs.Fields(5))
            strTemp(7) = CheckStr(m_rs.Fields(6))
            strTemp(8) = CheckStr(m_rs.Fields(7))
            strTemp(9) = CheckStr(m_rs.Fields(8))
            strTemp(10) = CheckStr(m_rs.Fields(9))
            strTemp(11) = CheckStr(m_rs.Fields(10))
            strTemp(12) = CheckStr(m_rs.Fields(11))
            strTemp(13) = CheckStr(m_rs.Fields(12))
            strTemp(14) = CheckStr(m_rs.Fields(13))
            strTemp(15) = CheckStr(m_rs.Fields(14))    'add by sonia 2018/1/15
            
            If iLine > 34 Or iLine = 1 Or _
                  (strType <> strTemp(1)) Then
                  
                If (strType <> "" And strType <> strTemp(1)) Then
                   PrintEnd '合計
                End If
                
                'If .AbsolutePosition <> .RecordCount Then
                    If strType <> "" Then Printer.NewPage
                    iLine = 1
                    PrintTitle '列印表頭
                'End If
            End If
                        
            PrintDetail '列印表中
            
            strType = strTemp(1) '依公司別跳頁
            dblAmt1 = dblAmt1 + strTemp(6)
            dblAmt2 = dblAmt2 + strTemp(7)
            dblAmt3 = dblAmt3 + strTemp(8)
            dblAmt4 = dblAmt4 + strTemp(9)
            dblAmt5 = dblAmt5 + strTemp(10)
            dblAmt6 = dblAmt6 + strTemp(11)
            dblAmt7 = dblAmt7 + strTemp(12)
            dblAmt8 = dblAmt8 + strTemp(13)
            dblAmt9 = dblAmt9 + strTemp(14)
            dblAmt10 = dblAmt10 + strTemp(15)        'add by sonia 2018/1/15
            dblTotAmt1 = dblTotAmt1 + strTemp(6)
            dblTotAmt2 = dblTotAmt2 + strTemp(7)
            dblTotAmt3 = dblTotAmt3 + strTemp(8)
            dblTotAmt4 = dblTotAmt4 + strTemp(9)
            dblTotAmt5 = dblTotAmt5 + strTemp(10)
            dblTotAmt6 = dblTotAmt6 + strTemp(11)
            dblTotAmt7 = dblTotAmt7 + strTemp(12)
            dblTotAmt8 = dblTotAmt8 + strTemp(13)
            dblTotAmt9 = dblTotAmt9 + strTemp(14)
            dblTotAmt10 = dblTotAmt10 + strTemp(15)  'add by sonia 2018/1/15
            m_rs.MoveNext
        Loop
        PrintEnd '合計
         
         '總計
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(230, "-")
         
         iLine = iLine + 1
         Printer.CurrentX = PLeft(3) - Printer.TextWidth("總　計：")
         Printer.CurrentY = iLine * 300
         Printer.Print "總　計："
         Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblTotAmt1, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt1, "##,###,##0")
         Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblTotAmt2, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt2, "##,###,##0")
         'add by sonia 2018/1/15
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblTotAmt10, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt10, "##,###,##0")
         'end 2018/1/15
         Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(dblTotAmt3, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt3, "##,###,##0")
         Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(dblTotAmt4, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt4, "##,###,##0")
         Printer.CurrentX = PLeft(9) - Printer.TextWidth(Format(dblTotAmt5, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt5, "##,###,##0");
         Printer.CurrentX = PLeft(10) - Printer.TextWidth(Format(dblTotAmt6, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt6, "##,###,##0")
         Printer.CurrentX = PLeft(11) - Printer.TextWidth(Format(dblTotAmt7, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt7, "##,###,##0")
         Printer.CurrentX = PLeft(12) - Printer.TextWidth(Format(dblTotAmt8, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt8, "##,###,##0")
         Printer.CurrentX = PLeft(13) - Printer.TextWidth(Format(dblTotAmt9, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt9, "##,###,##0")
    End With
Else
    MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd()
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(230, "-")
   
   iLine = iLine + 1
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("合　計：")
   Printer.CurrentY = iLine * 300
   Printer.Print "合　計："
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmt1, "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt1, "##,###,##0")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmt2, "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt2, "##,###,##0")
   'add by sonia 2018/1/15
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblAmt10, "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt10, "##,###,##0")
   'end 2018/1/15
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(dblAmt3, "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt3, "##,###,##0")
   Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(dblAmt4, "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt4, "##,###,##0")
   Printer.CurrentX = PLeft(9) - Printer.TextWidth(Format(dblAmt5, "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt5, "##,###,##0");
   Printer.CurrentX = PLeft(10) - Printer.TextWidth(Format(dblAmt6, "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt6, "##,###,##0")
   Printer.CurrentX = PLeft(11) - Printer.TextWidth(Format(dblAmt7, "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt7, "##,###,##0")
   Printer.CurrentX = PLeft(12) - Printer.TextWidth(Format(dblAmt8, "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt8, "##,###,##0")
   Printer.CurrentX = PLeft(13) - Printer.TextWidth(Format(dblAmt9, "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmt9, "##,###,##0")
   
   dblAmt1 = 0
   dblAmt2 = 0
   dblAmt3 = 0
   dblAmt4 = 0
   dblAmt5 = 0
   dblAmt6 = 0
   dblAmt7 = 0
   dblAmt8 = 0
   dblAmt9 = 0
   dblAmt10 = 0
   iLine = iLine + 1
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 11
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("台一關係企業　" & txt1(0) & "年　年終獎金發放明細表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "台一關係企業　" & txt1(0) & "年　年終獎金發放明細表"

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "列印人：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

If txt1(3) = "2" Then
   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "公司別：" & strTemp(1) & "　" & strTemp(2)
End If

'2013/1/21 MODIFY BY SONIA 表頭改二行
iLine = iLine + 2
Printer.CurrentX = PLeft(5) - Printer.TextWidth("特殊功績")
Printer.CurrentY = iLine * 300
Printer.Print "特殊功績"
Printer.CurrentX = PLeft(12) - Printer.TextWidth("補充保費")
Printer.CurrentY = iLine * 300
Printer.Print "代　　扣"

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "編號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"
Printer.CurrentX = PLeft(3) - Printer.TextWidth("平均月薪")
Printer.CurrentY = iLine * 300
Printer.Print "平均月薪"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("年終獎金")
Printer.CurrentY = iLine * 300
Printer.Print "年終獎金"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("特殊功績")
Printer.CurrentY = iLine * 300
Printer.Print "獎　　金"
'add by sonia 2018/1/15
Printer.CurrentX = PLeft(6) - Printer.TextWidth("紅　　利")
Printer.CurrentY = iLine * 300
Printer.Print "紅　　利"
'end 2018/1/15
Printer.CurrentX = PLeft(7) - Printer.TextWidth("未休假代金")
Printer.CurrentY = iLine * 300
Printer.Print "未休假代金"
Printer.CurrentX = PLeft(8) - Printer.TextWidth("缺勤扣除額")
Printer.CurrentY = iLine * 300
Printer.Print "缺勤扣除額"
Printer.CurrentX = PLeft(9) - Printer.TextWidth("應發金額")
Printer.CurrentY = iLine * 300
Printer.Print "應發金額"
Printer.CurrentX = PLeft(10) - Printer.TextWidth("借支扣除額")
Printer.CurrentY = iLine * 300
Printer.Print "借支扣除額"
Printer.CurrentX = PLeft(11) - Printer.TextWidth("代扣稅額")
Printer.CurrentY = iLine * 300
Printer.Print "代扣稅額"
'2013/1/21 ADD BY SONIA
Printer.CurrentX = PLeft(12) - Printer.TextWidth("補充保費")
Printer.CurrentY = iLine * 300
Printer.Print "補充保費"
'2013/1/21 END
Printer.CurrentX = PLeft(13) - Printer.TextWidth("實發金額")
Printer.CurrentY = iLine * 300
Printer.Print "實發金額"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(230, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 500
   PLeft(2) = 1200
   PLeft(3) = 3100
   PLeft(4) = 4400
   PLeft(5) = 5700
   PLeft(6) = 7000
   PLeft(7) = 8300
   PLeft(8) = 9700
   PLeft(9) = 11000
   PLeft(10) = 12300
   PLeft(11) = 13600
   PLeft(12) = 14900
   PLeft(13) = 16200
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   'Modified by Morgan 2023/12/20
   'Printer.Print strTemp(4)
   PUB_PrintUnicodeText strTemp(4), Printer.CurrentX, Printer.CurrentY, 0
   'end 2023/12/20
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(5), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(5), "##,###,##0")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(strTemp(6), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "##,###,##0")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(7), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "##,###,##0")
   'add by sonia 2018/1/15
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(15), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(15), "##,###,##0")
   'end 2018/1/15
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(strTemp(8), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(8), "##,###,##0")
   Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(strTemp(9), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(9), "##,###,##0")
   Printer.CurrentX = PLeft(9) - Printer.TextWidth(Format(strTemp(10), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(10), "##,###,##0")
   Printer.CurrentX = PLeft(10) - Printer.TextWidth(Format(strTemp(11), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(11), "##,###,##0")
   Printer.CurrentX = PLeft(11) - Printer.TextWidth(Format(strTemp(12), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(12), "##,###,##0")
   Printer.CurrentX = PLeft(12) - Printer.TextWidth(Format(strTemp(13), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(13), "##,###,##0")
   Printer.CurrentX = PLeft(13) - Printer.TextWidth(Format(strTemp(14), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(14), "##,###,##0")
   
   iLine = iLine + 1
End Sub

Sub StrMenu2()

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

'2009/1/14 modify by sonia 依部門列印時台一投資分開統計
'm_str = "select ST03,ST01,ST02,yb04,T5,T6,T8,T15+T16,T5+T6+T8-(T15+T16),T17,T5+T6+T8-(T15+T16)-T17 " & _
             "from staff,yearbonus, " & _
             "(select replace(yb02,'A','0') as T2,sum(nvl(yb05,0)) as T5,sum(nvl(yb06,0)) as T6,sum(nvl(yb08,0)) as T8,sum(nvl(yb15,0)) as T15,sum(nvl(yb16,0)) as T16,sum(nvl(yb17,0)) as T17 " & _
             "From yearbonus " & _
             "where " & m_StrSQL & _
             "group by replace(yb02,'A','0')) T " & _
             "where T2=st01(+) and T2=yb02 and yb04>0 " & _
             "order by ST03,ST01 "
'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
'2013/1/21 MODIFY BY SONIA 加YB25代扣補充保費
'modify by sonia 2018/1/15 +YB26
'modify by sonia 2018/1/30 婧瑄說應領不可扣除借支
'Modified by Morgan 2023/12/20 ST03-->YB03 部門改抓年終獎金的
'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
m_str = "select YB03,ST01,ST02,yb04,T5,T6,T8,T15,T5+T6+T26+T8-T15,T16,T17,T18,T5+T6+T26+T8-T15-T17-T18-T16,comp,T26 from (" & _
        "select YB03,ST01,ST02,nvl(yb04,0) yb04,substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4) as T2,sum(nvl(yb05,0)) as T5,sum(nvl(yb06,0)) as T6,sum(nvl(yb08,0)) as T8,sum(nvl(yb15,0)) as T15,sum(nvl(yb16,0)) as T16,sum(nvl(yb17,0)) as T17,decode(YB03,'R04','2','1') comp,sum(nvl(yb25,0)) as T18,sum(nvl(yb26,0)) as T26 " & _
        "From staff,yearbonus where " & m_StrSQL & " and substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4)=st01(+) " & _
        "group by YB03,ST01,ST02,yb04,substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4)) " & _
        "order by comp,YB03,ST01"
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        '預設值
        iLine = 1
        strType = "" '切頁條件
        dblAmt1 = 0: dblAmt2 = 0: dblAmt3 = 0: dblAmt4 = 0: dblAmt5 = 0
        dblAmt6 = 0: dblAmt7 = 0: dblAmt8 = 0: dblAmt9 = 0: dblAmt10 = 0
        dblTotAmt1 = 0: dblTotAmt2 = 0: dblTotAmt3 = 0: dblTotAmt4 = 0: dblTotAmt5 = 0
        dblTotAmt6 = 0: dblTotAmt7 = 0: dblTotAmt8 = 0: dblTotAmt9 = 0: dblTotAmt10 = 0
        
        Do While Not m_rs.EOF
            
            For m_i = 1 To 15
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(0)) '部門別
            strTemp(2) = CheckStr(m_rs.Fields(1)) '編號
            strTemp(3) = CheckStr(m_rs.Fields(2)) '姓名
            strTemp(4) = CheckStr(m_rs.Fields(3))
            strTemp(5) = CheckStr(m_rs.Fields(4))
            strTemp(6) = CheckStr(m_rs.Fields(5))
            strTemp(7) = CheckStr(m_rs.Fields(6))
            strTemp(8) = CheckStr(m_rs.Fields(7))
            strTemp(9) = CheckStr(m_rs.Fields(8))
            strTemp(10) = CheckStr(m_rs.Fields(9))
            strTemp(11) = CheckStr(m_rs.Fields(10))
            strTemp(12) = CheckStr(m_rs.Fields(11))
            strTemp(13) = CheckStr(m_rs.Fields(12))
            strTemp(14) = CheckStr(m_rs.Fields(14))   'add by sonia 2018/1/15
            
'            If (strType <> "" And strType <> strTemp(1)) Then
'               PrintEnd '合計
'            End If
            
            '2009/1/15 modify by sonia
            'If iLine > 34 Or iLine = 1 Then
            If iLine > 34 Or iLine = 1 Or _
                  (strType <> "" & m_rs.Fields("comp")) Then
                  
                If (strType <> "" And strType <> "" & m_rs.Fields("comp")) Then
                   PrintEnd '合計
                End If
                
            '2009/1/15 end
                'If .AbsolutePosition <> .RecordCount Then
                    If strType <> "" Then Printer.NewPage
                    iLine = 1
                    PrintTitle '列印表頭
                'End If
            End If
            
            PrintDetail2 '列印表中
            
            strType = "" & m_rs.Fields("comp") '依comp別跳頁
            dblAmt1 = dblAmt1 + strTemp(5)
            dblAmt2 = dblAmt2 + strTemp(6)
            dblAmt3 = dblAmt3 + strTemp(7)
            dblAmt4 = dblAmt4 + strTemp(8)
            dblAmt5 = dblAmt5 + strTemp(9)
            dblAmt6 = dblAmt6 + strTemp(10)
            dblAmt7 = dblAmt7 + strTemp(11)
            dblAmt8 = dblAmt8 + strTemp(12)
            dblAmt9 = dblAmt9 + strTemp(13)
            dblAmt10 = dblAmt10 + strTemp(14)          'add by sonia 2018/1/15
            dblTotAmt1 = dblTotAmt1 + strTemp(5)
            dblTotAmt2 = dblTotAmt2 + strTemp(6)
            dblTotAmt3 = dblTotAmt3 + strTemp(7)
            dblTotAmt4 = dblTotAmt4 + strTemp(8)
            dblTotAmt5 = dblTotAmt5 + strTemp(9)
            dblTotAmt6 = dblTotAmt6 + strTemp(10)
            dblTotAmt7 = dblTotAmt7 + strTemp(11)
            dblTotAmt8 = dblTotAmt8 + strTemp(12)
            dblTotAmt9 = dblTotAmt9 + strTemp(13)
            dblTotAmt10 = dblTotAmt10 + strTemp(14)    'add by sonia 2018/1/15
            m_rs.MoveNext
        Loop
        PrintEnd '合計
         
         '總計
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(230, "-")
         
         iLine = iLine + 1
         Printer.CurrentX = PLeft(3) - Printer.TextWidth("總　計：")
         Printer.CurrentY = iLine * 300
         Printer.Print "總　計："
         Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblTotAmt1, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt1, "##,###,##0")
         Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblTotAmt2, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt2, "##,###,##0")
         'add by sonia 2018/1/15
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblTotAmt10, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt10, "##,###,##0")
         'end 2018/1/15
         Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(dblTotAmt3, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt3, "##,###,##0")
         Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(dblTotAmt4, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt4, "##,###,##0")
         Printer.CurrentX = PLeft(9) - Printer.TextWidth(Format(dblTotAmt5, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt5, "##,###,##0");
         Printer.CurrentX = PLeft(10) - Printer.TextWidth(Format(dblTotAmt6, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt6, "##,###,##0")
         Printer.CurrentX = PLeft(11) - Printer.TextWidth(Format(dblTotAmt7, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt7, "##,###,##0")
         Printer.CurrentX = PLeft(12) - Printer.TextWidth(Format(dblTotAmt8, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt8, "##,###,##0")
         Printer.CurrentX = PLeft(13) - Printer.TextWidth(Format(dblTotAmt9, "##,###,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblTotAmt9, "##,###,##0")
    End With
Else
    MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintDetail2()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   'Modified by Morgan 2023/12/20
   'Printer.Print strTemp(3)
   PUB_PrintUnicodeText strTemp(3), Printer.CurrentX, Printer.CurrentY, 0
   'end 2023/12/20
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(4), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(4), "##,###,##0")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(strTemp(5), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(5), "##,###,##0")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(6), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "##,###,##0")
   'add by sonia 2018/1/15
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(14), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(14), "##,###,##0")
   'end 2018/1/15
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(strTemp(7), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "##,###,##0")
   Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(strTemp(8), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(8), "##,###,##0")
   Printer.CurrentX = PLeft(9) - Printer.TextWidth(Format(strTemp(9), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(9), "##,###,##0")
   Printer.CurrentX = PLeft(10) - Printer.TextWidth(Format(strTemp(10), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(10), "##,###,##0")
   Printer.CurrentX = PLeft(11) - Printer.TextWidth(Format(strTemp(11), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(11), "##,###,##0")
   Printer.CurrentX = PLeft(12) - Printer.TextWidth(Format(strTemp(12), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(12), "##,###,##0")
   Printer.CurrentX = PLeft(13) - Printer.TextWidth(Format(strTemp(13), "##,###,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(13), "##,###,##0")
   
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
   Set frm170217 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 3
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 1, 2
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
         ' 判斷員工代號須為 6~9 或 F 開頭
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
      Case Else
   End Select
End Sub
