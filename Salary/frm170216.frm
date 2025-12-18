VERSION 5.00
Begin VB.Form frm170216 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工年終獎金明細"
   ClientHeight    =   3210
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   4700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4700
   Begin VB.CheckBox Check4 
      Caption         =   "EMail給自己"
      Height          =   255
      Left            =   552
      TabIndex        =   12
      Top             =   2304
      Width           =   1356
   End
   Begin VB.CheckBox Check3 
      Caption         =   "台一投資及離職同仁(以EMail發送薪資單)"
      Height          =   255
      Left            =   552
      TabIndex        =   11
      Top             =   1944
      Width           =   3540
   End
   Begin VB.CheckBox Check1 
      Caption         =   "只印要印薪資單者"
      Height          =   255
      Left            =   552
      TabIndex        =   10
      Top             =   1584
      Width           =   2175
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   2
      Left            =   2376
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   1476
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1200
      Width           =   765
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   1476
      MaxLength       =   3
      TabIndex        =   0
      Top             =   840
      Width           =   435
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3660
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   2580
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   7
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Line Line2 
      X1              =   1992
      X2              =   2652
      Y1              =   1296
      Y2              =   1296
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   516
      TabIndex        =   9
      Top             =   1236
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "獎金年度："
      Height          =   180
      Left            =   516
      TabIndex        =   8
      Top             =   876
      Width           =   900
   End
End
Attribute VB_Name = "frm170216"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by SINDY 2009/01/06
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 25) As Integer
Dim strTemp(1 To 30) As String
'Dim PaperX As Double
'Dim paperY As Double
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim strYM As String
Dim m_YearDay As Long       '年度總天數
Dim m_AttachPath As String 'Added by Morgan 2024/2/2

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        If txt1(0) = "" Then
            MsgBox "獎金年度不可空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        If txt1(0) <> "" Then
            If ChkDate(txt1(0) & "0101") = False Then
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
         
        'add by sonia 2016/1/11
        'Modified by Morgan 2024/2/2 +
        If txt1(1) = "" And txt1(2) = "" And Check1.Value = 0 And Check3.Value = 0 Then
           'If MsgBox("未勾選【只印要印薪資單者】，是否確定要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
           If MsgBox("未勾選【" & Check1.Caption & " 】或【" & Check3.Caption & "】，是否確定要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
              Exit Sub
           End If
        End If
        'end 2016/1/11
       
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
        'add by sonia 2016/1/11
        If Check1.Value = 1 Then
           m_StrSQL = m_StrSQL & " and sd50='Y' "
        End If
        'end 2016/1/11
         
        StrMenu1
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
End Select
End Sub

Sub StrMenu1()
   Dim stSubject As String, strCompNo As String, strPdfName As String, strMsg As String, strMsgOK As String, strMsgErr As String, stCon As String 'Added by Morgan 2024/2/2
   
   'Added by Morgan 2024/2/2
   If Check3.Value = vbChecked Then
      stCon = stCon & " and (yb03='R04' or (st04='2' and st51<yb19))"
   End If
   'end 2024/2/2
   
   '2010/1/18 ADD BY SONIA
   '取得計算年度之總天數
   If PUB_GetMonthDays((Val(txt1(0)) + 1911), 2) = 28 Then
      m_YearDay = 365
   Else
      m_YearDay = 366
   End If
   '2010/1/18 END
   
   '2009/1/16 modify by sonia 加先以部門排序
   'm_str = "select T2,ST02,a0902,nvl(YB04,0),TT5,TT6,TT8,(TT5+TT6+TT8),TT15,TT16,(TT15+TT16),(TT5+TT6+TT8)-(TT15+TT16),TT17,(TT5+TT6+TT8)-(TT15+TT16)-TT17,nvl(YB07,0),nvl(YB09,0),nvl(YB10,0),nvl(YB12,0),nvl(YB13,0),nvl(YB11,0),nvl(YB14,0) " & _
               "from staff,acc090,yearbonus, " & _
               "(select T2,sum(T5) as TT5,sum(T6) as TT6,sum(T8) as TT8,sum(T15) as TT15,sum(T16) as TT16,sum(T17) as TT17 " & _
               "From " & _
               "(select replace(yb02,'A','0') as T2,nvl(yb05,0) as T5,nvl(yb06,0) as T6,nvl(yb07,0) as T7,nvl(yb08,0) as T8,nvl(yb15,0) as T15,nvl(yb16,0) as T16,nvl(yb17,0) as T17 " & _
               "From yearbonus " & _
               "where " & m_StrSQL & ") T " & _
               "group by T2) Y " & _
               "where T2=ST01(+) " & _
               "and ST03=a0901(+) " & _
               "and (yb01='" & strYM & "' and yb02=T2) " & _
               "order by T2 "
   '2010/1/18 modify by sonia 加讀每人當年工作天
   'm_str = "select T2,ST02,a0902,nvl(YB04,0),TT5,TT6,TT8,(TT5+TT6+TT8),TT15,TT16,(TT15+TT16),(TT5+TT6+TT8)-(TT15+TT16),TT17,(TT5+TT6+TT8)-(TT15+TT16)-TT17,nvl(YB07,0),nvl(YB09,0),nvl(YB10,0),nvl(YB12,0),nvl(YB13,0),nvl(YB11,0),nvl(YB14,0) " & _
               "from staff,acc090,yearbonus, " & _
               "(select T1,T2,sum(T5) as TT5,sum(T6) as TT6,sum(T8) as TT8,sum(T15) as TT15,sum(T16) as TT16,sum(T17) as TT17 " & _
               "From " & _
               "(select yb03 T1,replace(yb02,'A','0') as T2,nvl(yb05,0) as T5,nvl(yb06,0) as T6,nvl(yb07,0) as T7,nvl(yb08,0) as T8,nvl(yb15,0) as T15,nvl(yb16,0) as T16,nvl(yb17,0) as T17 " & _
               "From yearbonus " & _
               "where " & m_StrSQL & ") T " & _
               "group by T1,T2) Y " & _
               "where T2=ST01(+) " & _
               "and T1=a0901(+) " & _
               "and (yb01='" & strYM & "' and yb02=T2) " & _
               "order by T1,T2 "
   '2010/2/3 MODIFY BY SONIA 加入所別排序
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'modify by sonia 2016/1/11 加入只印要印薪資單者sd50條件
   'm_str = "select T2,ST02,a0902,nvl(YB04,0),TT5,TT6,TT8,(TT5+TT6+TT8),TT15,TT16,(TT15+TT16),(TT5+TT6+TT8)-(TT15+TT16),TT17,(TT5+TT6+TT8)-(TT15+TT16)-TT17-TT18,nvl(YB07,0),nvl(YB09,0),nvl(YB10,0),nvl(YB12,0),nvl(YB13,0),nvl(YB11,0),nvl(YB14,0),Z.workday,st06,TT18 " & _
               "from staff,acc090,yearbonus, " & _
               "(select sm01,sum(sm27) workday from salarymonth where substr(sm02,1,4)='" & strYM & "' group by sm01) Z, " & _
               "(select T1,T2,sum(T5) as TT5,sum(T6) as TT6,sum(T8) as TT8,sum(T15) as TT15,sum(T16) as TT16,sum(T17) as TT17,sum(T18) as TT18 " & _
               "From " & _
               "(select yb03 T1,substr(YB02,1,1)||replace(substr(YB02,2),'A','0') as T2,nvl(yb05,0) as T5,nvl(yb06,0) as T6,nvl(yb07,0) as T7,nvl(yb08,0) as T8,nvl(yb15,0) as T15,nvl(yb16,0) as T16,nvl(yb17,0) as T17,nvl(yb25,0) as T18 " & _
               "From yearbonus " & _
               "where " & m_StrSQL & ") T " & _
               "group by T1,T2) Y " & _
               "where T2=ST01(+) and T1=a0901(+) " & _
               "and (yb01='" & strYM & "' and yb02=T2) and t2=z.sm01(+) " & _
               "order by st06,T1,T2 "
   'modify by sonia 2016/2/26 +YB19以抓補充保費費率
   'modify by sonia 2018/1/12 +YB26紅利
   'modify by sonia 2018/1/30 婧瑄說應領不可扣除借支TT16
   'modified by Morgan 2023/12/20 + acc090new 新部門啟用日的前年度要開始抓新部門名稱(a0922)，因為發放(扣繳)是隔年--秀玲
   'Modified by Morgan 2024/2/2 +yb24,st01,st02
   'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
   m_str = "select T2,ST02,decode(sign(yb01-" & Left(新部門啟用日, 4) & "+1),-1,a0902,a0922) a0902,nvl(YB04,0),TT5,TT6,TT8,(TT5+TT6+TT19+TT8),TT15,TT16,(TT15),(TT5+TT6+TT19+TT8)-(TT15),TT17,(TT5+TT6+TT19+TT8)-(TT15+TT16)-TT17-TT18,nvl(YB07,0),nvl(YB09,0),nvl(YB10,0),nvl(YB12,0),nvl(YB13,0),nvl(YB11,0),nvl(YB14,0),Z.workday,st06,TT18,YB19,TT19 " & _
               ",yb24,st01,st02 from staff,acc090,acc090new,yearbonus, " & _
               "(select sm01,sum(sm27) workday from salarymonth where substr(sm02,1,4)='" & strYM & "' group by sm01) Z, " & _
               "(select T1,T2,sum(T5) as TT5,sum(T6) as TT6,sum(T8) as TT8,sum(T15) as TT15,sum(T16) as TT16,sum(T17) as TT17,sum(T18) as TT18,sum(T19) as TT19 " & _
               "From " & _
               "(select yb03 T1,substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4) as T2,nvl(yb05,0) as T5,nvl(yb06,0) as T6,nvl(yb07,0) as T7,nvl(yb08,0) as T8,nvl(yb15,0) as T15,nvl(yb16,0) as T16,nvl(yb17,0) as T17,nvl(yb25,0) as T18,nvl(yb26,0) as T19 " & _
               "From yearbonus,salarydata " & _
               "where " & m_StrSQL & " and substr(YB02,1,2)||replace(substr(YB02,3,1),'A','0')||substr(YB02,4)=sd01(+)) T " & _
               "group by T1,T2) Y " & _
               "where T2=ST01(+) and T1=a0901(+) and a0921(+)=T1 " & _
               "and (yb01='" & strYM & "' and yb02=T2) and t2=z.sm01(+) " & stCon & _
               "order by st06,T1,T2 "
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       With m_rs
           m_rs.MoveFirst
           
           '預設值
           iLine = 1
           strType = "" '切頁條件
           
            'Added by Morgan 2024/2/2
            If Check3.Value = vbUnchecked And Check4.Value = vbUnchecked Then
               Set Printer = Printers(Combo1.ListIndex)
               Printer.EndDoc
               Printer.Orientation = 1 '1.直印 2.橫印
               'Modified by Morgan 2024/2/2 改A4
               'Printer.PaperSize = PUB_GetPaperSize(3)  '中一刀
               Printer.PaperSize = 9
            End If 'Added by Morgan 2024/2/2
            
            
            
      
           Do While Not m_rs.EOF
               'Added by Morgan 2024/2/2
               strCompNo = "" & .Fields("yb24")
               If Check3.Value = vbChecked Or Check4.Value = vbChecked Then
                  stSubject = Val(txt1(0)) & "年年終獎金明細(" & .Fields("st01") & ")"
                  strPdfName = .Fields("st01") & "_" & strYM & ".pdf"
                  frmPDF.Show
                  frmPDF.StartProcess m_AttachPath, strPdfName
                  Printer.Orientation = 1 '1.直印 2.橫印
                  Printer.PaperSize = 9
               End If
               'end 2024/2/2
               
               For m_i = 1 To 30
                   strTemp(m_i) = ""
               Next m_i
   '1. T2,
   '2. ST02,
   '3. a0902,
   '4. nvl(YB04,0),
   '5. TT5,
   '6. TT6,
   '7. TT19,
   '8. TT8,
   '9. (TT5+TT6+TT8),
   '10. TT15,
   '11.TT16,
   '12.(TT15),
   '13.(TT5+TT6+TT8)-(TT15),
   '14.TT17,
   '15.(TT5+TT6+TT8)-(TT15+TT16)-TT17-TT18,
   '16.nvl(YB07,0),
   '17.nvl(YB09,0),
   '18.nvl(YB10,0),
   '19.nvl(YB12,0),
   '20.nvl(YB13,0),
   '21.nvl(YB11,0),
   '22.nvl(YB14,0)
   '23.TT18,
   '28.YB19 create date
   
               strTemp(1) = CheckStr(m_rs.Fields(0)) '編號
               strTemp(2) = CheckStr(m_rs.Fields(1)) '姓名
               strTemp(3) = CheckStr(m_rs.Fields(2)) '部門
               strTemp(4) = CheckStr(m_rs.Fields(3)) '平均月薪
               strTemp(5) = CheckStr(m_rs.Fields(4))
               strTemp(6) = CheckStr(m_rs.Fields(5))
               strTemp(7) = CheckStr("" & m_rs.Fields("TT19"))
               strTemp(8) = CheckStr(m_rs.Fields(6))
               strTemp(9) = CheckStr(m_rs.Fields(7))
               strTemp(10) = CheckStr(m_rs.Fields(8))
               strTemp(11) = CheckStr(m_rs.Fields(9))
               strTemp(12) = CheckStr(m_rs.Fields(10))
               strTemp(13) = CheckStr(m_rs.Fields(11))
               strTemp(14) = CheckStr(m_rs.Fields(12))
               strTemp(15) = CheckStr(m_rs.Fields(13))
               strTemp(16) = CheckStr(m_rs.Fields(14)) '未休特別假代金
               strTemp(17) = CheckStr(m_rs.Fields(15)) '各假別
               strTemp(18) = CheckStr(m_rs.Fields(16))
               strTemp(19) = CheckStr(m_rs.Fields(17))
               strTemp(20) = CheckStr(m_rs.Fields(18))
               strTemp(21) = CheckStr(m_rs.Fields(19))
               strTemp(22) = CheckStr(m_rs.Fields(20))
               strTemp(23) = GetYearBonusMonth(txt1(0), strTemp(1)) '取得年終獎金基準月數
               '取得考績及核發獎金基數
               If GetYearMerit(txt1(0), strTemp(1), strTemp(24), strTemp(25)) = True Then
               Else
                  strTemp(24) = ""
                  strTemp(25) = ""
               End If
               strTemp(26) = CheckStr(m_rs.Fields(21))  '2010/1/18 ADD BY SONIA取得當年工作天數
               strTemp(27) = CheckStr(m_rs.Fields(23))  '2013/1/22 add by sonia 補充保費
               'add by sonia 2016/2/24 抓補充保費費率
               strTemp(28) = PUB_GetNhiRate(Val("" & m_rs.Fields("yb19")))
               'end 2016/2/24
               
               PrintTitle strCompNo '列印表頭
               PrintDetail '列印表中、表尾
               
               'Added by Morgan 2024/2/2
               If Check3.Value = vbChecked Or Check4.Value = vbChecked Then
                  Printer.EndDoc
                  frmPDF.EndtProcess
                  Unload frmPDF
                  If PUB_SalarySendMail(stSubject, .Fields("st01"), m_AttachPath & "\" & strPdfName, strMsg, Check4.Value) = True Then
                     strMsgOK = strMsgOK & .Fields("st01") & .Fields("st02") & IIf(strMsg <> "", "(" & strMsg & ")", "") & vbCrLf
                  Else
                     strMsgErr = strMsgErr & .Fields("st01") & .Fields("st02") & ":" & strMsg & vbCrLf
                  End If
               End If
               'end 2024/2/2
         
               m_rs.MoveNext
           Loop
       End With
       
      'Modified by Morgan 2024/2/2
      'Printer.EndDoc
      If Check3.Value = vbChecked Or Check4.Value = vbChecked Then
         MsgBoxU "EMail完成，清單如下：" & vbCrLf & "成功:" & vbCrLf & strMsgOK & IIf(strMsgErr <> "", vbCrLf & "失敗：" & strMsgErr, ""), vbExclamation
      Else
         Printer.EndDoc
         MsgBox "列印結束 !"
      End If
       
   Else
       MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
       Exit Sub
   End If
   
   ShowPrintOk
End Sub
'Modified by Morgan 2024/2/2 +pComp
Private Sub PrintTitle(Optional pCompNo As String)
   GetPleft
   
   iLine = 1 '新頁重頭列印
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = 3750
   Printer.CurrentY = iLine * 250
   'Added by Morgan 2024/2/2
   If pCompNo <> "" Then
      Printer.Print CompNameQuery(pCompNo) & "　" & txt1(0) & "年　年終獎金明細"
   Else
   'end 2024/2/2
      Printer.Print "台一關係企業　" & txt1(0) & "年　年終獎金明細"
   End If
   
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 250
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   
   iLine = iLine + 2
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 250
   Printer.Print "員工編號：" & strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 250
   'Modified by Morgan 2024/2/2
   'Printer.Print "姓　　名：" & strTemp(2)
   PUB_PrintUnicodeText "姓　　名：" & strTemp(2), Printer.CurrentX, Printer.CurrentY, 0
   'end 2024/2/2
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 250
   Printer.Print "部　　門：" & strTemp(3)
   
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 250
   Printer.Print "基準月數：" & strTemp(23)
   
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 250
   Printer.Print "考　　績：" & strTemp(24)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 250
   Printer.Print "核發獎金基數：" & strTemp(25)
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 250
   Printer.Print "平均基準月薪：" & Format(strTemp(4), "##,##0")
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 500
   PLeft(2) = 4000
   PLeft(3) = 8000
   PLeft(4) = 1500
   PLeft(5) = 7500
   PLeft(6) = 10000
End Sub

Sub PrintDetail()
Dim i As Integer, strText As String
Dim dblDay As Double
Dim dblHour As Double
Dim m_taxrate As String   '2010/12/30 add by sonia 非固定之薪資所得扣繳稅率
   
   Call Pub_GetSpecWorkHour(strTemp(1), Val(txt1(0)) + 19111231)   'add by sonia 2018/2/1
   
   For i = 1 To 4
      iLine = iLine + 1
      Printer.CurrentX = PLeft(4)
      Printer.CurrentY = iLine * 250
      If i = 1 Then
         Printer.Print "年終獎金"
      ElseIf i = 2 Then
         '2015/1/27 modify by sonia 無特殊功績獎金則不印該欄
         'Printer.Print "特殊功績獎金"
         If strTemp(6) > 0 Then
            Printer.Print "特殊功績獎金"
         Else
            iLine = iLine - 1
         End If
         '2015/1/27 end
      'add by sonia 2018/1/12
      ElseIf i = 3 Then
         '無紅利則不印該欄
         If strTemp(7) > 0 Then
            Printer.Print "紅　利"
         Else
            iLine = iLine - 1
         End If
      '2018/1/12 end
      ElseIf i = 4 Then
         'modify by sonia 2018/2/1 每日8小時改用上班特殊時數PUB_intWkHour
         'dblDay = (strTemp(16) * 10) \ (8 * 10)
         'dblHour = Round(strTemp(16) - (dblDay * 8), 1)
         dblDay = (strTemp(16) * 10) \ (PUB_intWkHour * 10)
         dblHour = Round(strTemp(16) - (dblDay * PUB_intWkHour), 1)
         'end 2018/2/1
         Printer.Print "未休特別假代金"
         Printer.CurrentX = 3600 - Printer.TextWidth(dblDay)
         Printer.CurrentY = iLine * 250
         Printer.Print dblDay
         Printer.CurrentX = 4000
         Printer.CurrentY = iLine * 250
         Printer.Print "日"
         Printer.CurrentX = 4600 - Printer.TextWidth(dblHour)
         Printer.CurrentY = iLine * 250
         Printer.Print dblHour
         Printer.CurrentX = 5000
         Printer.CurrentY = iLine * 250
         Printer.Print "時"
      End If
      'strTemp(6)-(8)
      Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(4 + i), "##,##0"))
      Printer.CurrentY = iLine * 250
      '2010/1/18 MODIFY BY SONIA 工作天不滿一年者列印計算公式
      'Printer.Print Format(strTemp(4 + i), "##,##0")
      If i = 1 And strTemp(26) <> m_YearDay Then
         Printer.Print Format(strTemp(4 + i), "##,##0") & " (" & Format(strTemp(4), "##,##0") & " * " & strTemp(23) & " * " & strTemp(25) & " * " & strTemp(26) & " / " & m_YearDay & ")"
      '2015/1/27 modify by sonia 無特殊功績獎金則不印該欄
      'Else
      '   Printer.Print Format(strTemp(4 + i), "##,##0")
      ElseIf i = 2 Then
         If strTemp(6) > 0 Then Printer.Print Format(strTemp(4 + i), "##,##0")
      'add by sonia 2018/1/12 +YB26紅利,但無值則不印
      ElseIf i = 3 Then
         If strTemp(7) > 0 Then Printer.Print Format(strTemp(4 + i), "##,##0")
      'end 2018/1/12
      Else
         Printer.Print Format(strTemp(4 + i), "##,##0")
      '2015/1/27 end
      End If
      '2010/1/18 END
      If i = 4 Then
         'strTemp(9)
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(9), "##,##0"))
         Printer.CurrentY = iLine * 250
         Printer.Print Format(strTemp(9), "##,##0")
         iLine = iLine + 1
         Printer.CurrentX = 6000
         Printer.CurrentY = iLine * 250
         Printer.Print String(58, "-")
      End If
   Next i
   
   For i = 1 To 6
      'strTemp(17)-(22)
      'modify by sonia 2018/2/1 每日8小時改用上班特殊時數PUB_intWkHour
      'dblDay = (strTemp(16 + i) * 10) \ (8 * 10)
      'dblHour = Round(strTemp(16 + i) - (dblDay * 8), 1)
      dblDay = (strTemp(16 + i) * 10) \ (PUB_intWkHour * 10)
      dblHour = Round(strTemp(16 + i) - (dblDay * PUB_intWkHour), 1)
      'end  2018/2/1
      
      iLine = iLine + 1
      If i = 1 Then
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 250
         Printer.Print "扣　除"
      End If
      Printer.CurrentX = PLeft(4)
      Printer.CurrentY = iLine * 250
      If i = 1 Then
         strText = "病　假"
      ElseIf i = 2 Then
         strText = "事　假"
      ElseIf i = 3 Then
         strText = "產　假"
      ElseIf i = 4 Then
         strText = "流產假"
      ElseIf i = 5 Then
         strText = "曠　職"
      ElseIf i = 6 Then
         strText = "公傷假"
      End If
      Printer.Print strText
      Printer.CurrentX = 3600 - Printer.TextWidth(dblDay)
      Printer.CurrentY = iLine * 250
      Printer.Print dblDay
      Printer.CurrentX = 4000
      Printer.CurrentY = iLine * 250
      Printer.Print "日"
      Printer.CurrentX = 4600 - Printer.TextWidth(dblHour)
      Printer.CurrentY = iLine * 250
      Printer.Print dblHour
      Printer.CurrentX = 5000
      Printer.CurrentY = iLine * 250
      Printer.Print "時"
   Next i
   
'modify by sonia 2018/1/30 婧瑄說借支扣款改放最下面
'   For i = 1 To 2
'      iLine = iLine + 1
'      Printer.CurrentX = PLeft(4)
'      Printer.CurrentY = iLine * 250
'      If i = 1 Then
'         Printer.Print "缺勤扣款"
'      ElseIf i = 2 Then
'         Printer.Print "借支扣款"
'      End If
'      'strTemp(10)-(11)
'      Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(9 + i), "##,##0"))
'      Printer.CurrentY = iLine * 250
'      Printer.Print Format(strTemp(9 + i), "##,##0")
'      If i = 2 Then
'         'strTemp(12)
'         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(12), "##,##0"))
'         Printer.CurrentY = iLine * 250
'         Printer.Print Format(strTemp(12), "##,##0")
'         iLine = iLine + 1
'         Printer.CurrentX = 6000
'         Printer.CurrentY = iLine * 250
'         Printer.Print String(58, "-")
'      End If
'   Next i
'
'   For i = 1 To 4
'      iLine = iLine + 1
'      Printer.CurrentX = PLeft(4)
'      Printer.CurrentY = iLine * 250
'      If i = 1 Then
'         Printer.Print "應領金額"
'      ElseIf i = 2 Then
'         '2010/12/30 modify by sonia 非固定之薪資所得扣繳稅率改抓 翻譯所得oc01='01'的稅率
'         'Printer.Print "代扣稅額＝(年終獎金＋特殊功績獎金－缺勤扣款) X 6%"
'         m_taxrate = 0
'         strExc(0) = "select oc04 from OtherSalaryCode where oc01='01'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            m_taxrate = "" & RsTemp.Fields(0)
'         End If
'         Printer.CurrentX = PLeft(1)
'         'modify by sonia 2018/1/12 特殊功績獎金及紅利二欄若沒值就不印
'         'Printer.Print "代扣稅額＝(年終獎金＋特殊功績獎金－缺勤扣款) X " & m_taxrate & "%"
'         If strTemp(6) + strTemp(7) = 0 Then
'            Printer.Print "代扣稅額＝(年終獎金－缺勤扣款) X " & m_taxrate & "%"
'         ElseIf strTemp(6) > 0 And strTemp(7) = 0 Then
'            Printer.Print "代扣稅額＝(年終獎金＋特殊功績獎金－缺勤扣款) X " & m_taxrate & "%"
'         ElseIf strTemp(6) = 0 And strTemp(7) > 0 Then
'            Printer.Print "代扣稅額＝(年終獎金＋紅利－缺勤扣款) X " & m_taxrate & "%"
'         Else
'            Printer.Print "代扣稅額＝(年終獎金＋特殊功績獎金＋紅利－缺勤扣款) X " & m_taxrate & "%"
'         End If
'         'end 2018/1/12
'
'         '2010/12/30 end
'      '2013/1/22 add by sonia 代扣補充保費
'      ElseIf i = 3 Then
'         Printer.CurrentX = PLeft(1)
'         'modify by sonia 2016/2/24 抓補充保費費率
'         'Printer.Print "代扣補充保費＝[(年終獎金＋特殊功績獎金－缺勤扣款)－(４倍投保金額)] X 2%"
'         'modify by sonia 2018/1/12 特殊功績獎金及紅利二欄若沒值就不印
'         'Printer.Print "代扣補充保費＝[(年終獎金＋特殊功績獎金－缺勤扣款)－(４倍投保金額)] X " & strTemp(27) & "%"
'         If strTemp(6) + strTemp(7) = 0 Then
'            Printer.Print "代扣補充保費＝[(年終獎金－缺勤扣款)－(４倍投保金額)] X " & strTemp(28) & "%"
'         ElseIf strTemp(6) > 0 And strTemp(7) = 0 Then
'            Printer.Print "代扣補充保費＝[(年終獎金＋特殊功績獎金－缺勤扣款)－(４倍投保金額)] X " & strTemp(28) & "%"
'         ElseIf strTemp(6) = 0 And strTemp(7) > 0 Then
'            Printer.Print "代扣補充保費＝[(年終獎金＋紅利－缺勤扣款)－(４倍投保金額)] X " & strTemp(28) & "%"
'         Else
'            Printer.Print "代扣補充保費＝[(年終獎金＋特殊功績獎金＋紅利－缺勤扣款)"
'            iLine = iLine + 1
'            Printer.CurrentX = PLeft(1)
'            Printer.CurrentY = iLine * 250
'            Printer.Print "　　　　　　－(４倍投保金額)] X " & strTemp(28) & "%"
'        End If
'         'end 2018/1/12
'      '2013/1/22 end
'      ElseIf i = 4 Then
'         Printer.Print "實領金額"
'      End If
'      'strTemp(13)-(16),2013/1/22 加strTemp(27)
'      If i <= 2 Then
'         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(12 + i), "##,##0"))
'         Printer.CurrentY = iLine * 250
'         Printer.Print Format(strTemp(12 + i), "##,##0")
'      '2013/1/22 add by sonia
'      ElseIf i = 3 Then
'         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(27), "##,##0"))
'         Printer.CurrentY = iLine * 250
'         Printer.Print Format(strTemp(27), "##,##0")
'      ElseIf i = 4 Then
'         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(12 + i - 1), "##,##0"))
'         Printer.CurrentY = iLine * 250
'         Printer.Print Format(strTemp(12 + i - 1), "##,##0")
'      End If
'      '2013/1/22 end
'      If i = 1 Or i = 3 Then
'         iLine = iLine + 1
'         Printer.CurrentX = 8500
'         Printer.CurrentY = iLine * 250
'         Printer.Print String(25, "-")
'      ElseIf i = 4 Then
'         iLine = iLine + 1
'         Printer.CurrentX = 8500
'         Printer.CurrentY = iLine * 250
'         Printer.Print String(15, "=")
'      End If
'   Next i
   iLine = iLine + 1
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 250
   Printer.Print "缺勤扣款"
   'strTemp(10)
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(10), "##,##0"))
   Printer.CurrentY = iLine * 250
   Printer.Print Format(strTemp(10), "##,##0")
   'strTemp(12)
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(12), "##,##0"))
   Printer.CurrentY = iLine * 250
   Printer.Print Format(strTemp(12), "##,##0")
   iLine = iLine + 1
   Printer.CurrentX = 6000
   Printer.CurrentY = iLine * 250
   Printer.Print String(58, "-")
   
   For i = 1 To 5
      iLine = iLine + 1
      Printer.CurrentX = PLeft(4)
      Printer.CurrentY = iLine * 250
      If i = 1 Then
         Printer.Print "應領金額"
      ElseIf i = 2 Then
         '2010/12/30 modify by sonia 非固定之薪資所得扣繳稅率改抓 翻譯所得oc01='01'的稅率
         'Printer.Print "代扣稅額＝(年終獎金＋特殊功績獎金－缺勤扣款) X 6%"
         m_taxrate = 0
         strExc(0) = "select oc04 from OtherSalaryCode where oc01='01'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_taxrate = "" & RsTemp.Fields(0)
         End If
         Printer.CurrentX = PLeft(1)
         'modify by sonia 2018/1/12 特殊功績獎金及紅利二欄若沒值就不印
         'Printer.Print "代扣稅額＝(年終獎金＋特殊功績獎金－缺勤扣款) X " & m_taxrate & "%"
         If strTemp(6) + strTemp(7) = 0 Then
            Printer.Print "代扣稅額＝(年終獎金－缺勤扣款) X " & m_taxrate & "%"
         ElseIf strTemp(6) > 0 And strTemp(7) = 0 Then
            Printer.Print "代扣稅額＝(年終獎金＋特殊功績獎金－缺勤扣款) X " & m_taxrate & "%"
         ElseIf strTemp(6) = 0 And strTemp(7) > 0 Then
            Printer.Print "代扣稅額＝(年終獎金＋紅利－缺勤扣款) X " & m_taxrate & "%"
         Else
            Printer.Print "代扣稅額＝(年終獎金＋特殊功績獎金＋紅利－缺勤扣款) X " & m_taxrate & "%"
         End If
         'end 2018/1/12
         
         '2010/12/30 end
      '2013/1/22 add by sonia 代扣補充保費
      ElseIf i = 3 Then
         Printer.CurrentX = PLeft(1)
         'modify by sonia 2016/2/24 抓補充保費費率
         'Printer.Print "代扣補充保費＝[(年終獎金＋特殊功績獎金－缺勤扣款)－(４倍投保金額)] X 2%"
         'modify by sonia 2018/1/12 特殊功績獎金及紅利二欄若沒值就不印
         'Printer.Print "代扣補充保費＝[(年終獎金＋特殊功績獎金－缺勤扣款)－(４倍投保金額)] X " & strTemp(27) & "%"
         If strTemp(6) + strTemp(7) = 0 Then
            Printer.Print "代扣補充保費＝[(年終獎金－缺勤扣款)－(４倍投保金額)] X " & strTemp(28) & "%"
         ElseIf strTemp(6) > 0 And strTemp(7) = 0 Then
            Printer.Print "代扣補充保費＝[(年終獎金＋特殊功績獎金－缺勤扣款)－(４倍投保金額)] X " & strTemp(28) & "%"
         ElseIf strTemp(6) = 0 And strTemp(7) > 0 Then
            Printer.Print "代扣補充保費＝[(年終獎金＋紅利－缺勤扣款)－(４倍投保金額)] X " & strTemp(28) & "%"
         Else
            Printer.Print "代扣補充保費＝[(年終獎金＋特殊功績獎金＋紅利－缺勤扣款)"
            iLine = iLine + 1
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iLine * 250
            Printer.Print "　　　　　　－(４倍投保金額)] X " & strTemp(28) & "%"
        End If
         'end 2018/1/12
      '2013/1/22 end
      ElseIf i = 4 Then
         Printer.Print "借支扣款"
      ElseIf i = 5 Then
         Printer.Print "實領金額"
      End If
      'strTemp(13)-(16),2013/1/22 加strTemp(27)
      If i <= 2 Then
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(12 + i), "##,##0"))
         Printer.CurrentY = iLine * 250
         Printer.Print Format(strTemp(12 + i), "##,##0")
      '2013/1/22 add by sonia
      ElseIf i = 3 Then
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(27), "##,##0"))
         Printer.CurrentY = iLine * 250
         Printer.Print Format(strTemp(27), "##,##0")
      ElseIf i = 4 Then
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(11), "##,##0"))
         Printer.CurrentY = iLine * 250
         Printer.Print Format(strTemp(11), "##,##0")
      ElseIf i = 5 Then
         Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(15), "##,##0"))
         Printer.CurrentY = iLine * 250
         Printer.Print Format(strTemp(15), "##,##0")
      End If
      '2013/1/22 end
      If i = 1 Or i = 4 Then
         iLine = iLine + 1
         Printer.CurrentX = 8500
         Printer.CurrentY = iLine * 250
         Printer.Print String(25, "-")
      ElseIf i = 5 Then
         iLine = iLine + 1
         Printer.CurrentX = 8500
         Printer.CurrentY = iLine * 250
         Printer.Print String(15, "=")
      End If
   Next i
   
   iLine = iLine + 1
   Printer.NewPage
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
   
   
   'Added by Morgan 2024/2/2
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   Else
      PUB_KillAttach m_AttachPath
   End If
   'end 2024/2/2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_KillAttach m_AttachPath 'Added by Morgan 2024/2/2
   Set frm170216 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
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
            'add by sonia 2016/1/11
            Else
               Check1.Value = 0
            'end 2016/1/11
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

