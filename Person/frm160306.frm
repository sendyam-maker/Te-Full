VERSION 5.00
Begin VB.Form frm160306 
   BorderStyle     =   1  '單線固定
   Caption         =   "特別假名單"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5760
   Begin VB.CheckBox ChkErr 
      Caption         =   "加註【修改通知】字樣在主旨上"
      Height          =   285
      Left            =   90
      TabIndex        =   19
      Top             =   1140
      Width           =   2955
   End
   Begin VB.TextBox txtNote 
      Height          =   915
      Left            =   1470
      TabIndex        =   3
      Top             =   2310
      Width           =   4185
   End
   Begin VB.OptionButton Option1 
      Caption         =   "列印個人紙本通知單"
      Height          =   255
      Index           =   1
      Left            =   3540
      TabIndex        =   8
      Top             =   1050
      Value           =   -1  'True
      Width           =   1965
   End
   Begin VB.OptionButton Option1 
      Caption         =   "列印原清單"
      Height          =   255
      Index           =   0
      Left            =   3540
      TabIndex        =   7
      Top             =   690
      Width           =   1965
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2190
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1950
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   2970
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1950
      Width           =   705
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "電子化EMail通知(&E)"
      Height          =   435
      Index           =   2
      Left            =   1800
      TabIndex        =   5
      Top             =   90
      Width           =   1785
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   1
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   4
      Top             =   690
      Width           =   315
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3675
      TabIndex        =   6
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4680
      TabIndex        =   9
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   2190
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1590
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   330
      TabIndex        =   11
      Top             =   3750
      Width           =   4875
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
         TabIndex        =   12
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "通知內容加註："
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   18
      Top             =   2370
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   1260
      TabIndex        =   17
      Top             =   1980
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2910
      X2              =   3150
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "註：若處理全公司的特別假通知時，不要輸入員工代號"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   2
      Left            =   330
      TabIndex        =   16
      Top             =   3300
      Width           =   4320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "　　因為在寄發Mail時會有其影響"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   3
      Left            =   330
      TabIndex        =   15
      Top             =   3510
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否只印分所人員名單：        (Y:只印分所)"
      Height          =   180
      Left            =   75
      TabIndex        =   14
      Top             =   735
      Width           =   3405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "新年度："
      Height          =   180
      Left            =   1440
      TabIndex        =   13
      Top             =   1620
      Width           =   720
   End
End
Attribute VB_Name = "frm160306"
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
Dim strTemp(1 To 40) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim intRow As Integer
Dim strItem As String


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If txt1(0) = "" Then
            MsgBox "新年度不可以空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         If txt1(1) = "" Then  '2009/12/21 ADD BY SONIA 加只印分所條件
            m_StrSQL = ""
            If txt1(3) <> "" Or txt1(4) <> "" Then
               If txt1(3) <> "" Then
                   m_StrSQL = m_StrSQL & " and ST01>='" & txt1(3) & "' "
               End If
               If txt1(4) <> "" Then
                   m_StrSQL = m_StrSQL & " and ST01<='" & txt1(4) & "' "
               End If
               strItem = "1"
               Call StrMenu(False)
            Else
               '不含台一投資和台一開發
               'm_StrSQL = " and sd19<>'A' "
               'Modify By Sindy 2011/12/26
               m_StrSQL = " and substr(st03,1,1)<>'R' "
               strItem = "1"
               If StrMenu(False) = True Then
                  '為台一投資
                  'm_StrSQL = " and sd19='A' "
                  'Modify By Sindy 2011/12/26
                  m_StrSQL = " and st03='R04' "
                  strItem = "2"
                  Call StrMenu(False)
                  'Add By Sindy 2011/12/26
                  '為台一開發
                  m_StrSQL = " and st03='R08' "
                  strItem = "3"
                  Call StrMenu(False)
               End If
               '2011/12/26 End
            End If
            
         '2009/12/21 ADD BY SONIA 加只印分所條件
         Else
            m_StrSQL = ""
            '不含台一投資 及 北所
            'm_StrSQL = " and sd19<>'A' AND ST06<>'1' "
            'Modify By Sindy 2011/12/26
            m_StrSQL = " and substr(st03,1,1)<>'R' AND ST06<>'1' "
            If txt1(3) <> "" Then
                m_StrSQL = m_StrSQL & " and ST01>='" & txt1(3) & "' "
            End If
            If txt1(4) <> "" Then
                m_StrSQL = m_StrSQL & " and ST01<='" & txt1(4) & "' "
            End If
            strItem = "1"
            Call StrMenu(False)
         End If
         '2009/12/21 END
         Screen.MousePointer = vbDefault
      Case 1
           Unload Me
      Case 2 '電子化EMail通知
         Screen.MousePointer = vbHourglass
         m_StrSQL = ""
         If txt1(3) <> "" Then
             m_StrSQL = m_StrSQL & " and ST01>='" & txt1(3) & "' "
         End If
         If txt1(4) <> "" Then
             m_StrSQL = m_StrSQL & " and ST01<='" & txt1(4) & "' "
         End If
         'Add By Sindy 2019/7/16
         If ChkErr.Value = 1 Then
            If MsgBox("確定是要發【修改通知】信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
         '2019/7/16 END
         strItem = "1"
         Call StrMenu(True)
         Screen.MousePointer = vbDefault
   End Select
End Sub

'明細表
'Modify By Sindy 2019/6/26 bolEPrint:電子化人員E-Mail通知
Function StrMenu(bolEPrint As Boolean) As Boolean
Dim int_i As Integer
Dim m_ST06 As String    '2009/12/21 ADD BY SONIA
Dim dblTotDay As Double
Dim BolChkEnd04 As Boolean 'Add By Sindy 2014/12/29 最後一筆是否為04留職停薪
Dim strDay As String 'Add By Sindy 2017/1/6
'Add By Sindy 2019/6/26
Dim strNote As String
Dim strBackTaieDate As String
Dim intECnt As Integer, intPCnt As Integer
Dim bolConn As Boolean
   
On Error GoTo ErrHnd
   
   StrMenu = True
   intECnt = 0: intPCnt = 0
   
   'Modify By Sindy 2017/1/5
   '計算期滿：to_char(add_months(to_date(st13,'YYYYMMDD'),6),'YYYYMMDD') as ST13_6,to_char(add_months(to_date(st13,'YYYYMMDD'),12),'YYYYMMDD') as ST13_12
   '+,ST13,ST40
   'Add By Sindy 2019/7/4
   If Option1(1).Value = True Or bolEPrint = True Then '列印個人紙本通知單
      m_StrSQL = m_StrSQL & " and st04='1' " '在職的才寄
      m_StrSQL = m_StrSQL & " and ST01 not in('63001','67004') " '排除人員
   End If
   '2019/7/4 END
   If txt1(1) = "" Then    '2009/12/21 ADD BY SONIA 加印分所依所別排序及跳頁
      'Modify By Sindy 2023/2/4 and substr(st01,4,1)<>'9'
      m_str = "SELECT ST01,ST02,YV03,YV04,sqldatet(ST13),a0802,to_char(add_months(to_date(st13,'YYYYMMDD'),6),'YYYYMMDD') as ST13_6,to_char(add_months(to_date(st13,'YYYYMMDD'),12),'YYYYMMDD') as ST13_12,ST13,ST40,YV12,YV13,ST14 " & _
               "FROM YearVacation,staff,acc080,SalaryData " & _
               "WHERE yv02=st01(+) and yv01='" & Mid(Trim(DBDATE(txt1(0) & "0101")), 1, 4) & "' " & _
               "and yv02=sd01(+) and sd19=a0801(+) and substr(st01,4,1)<>'9' " & m_StrSQL & _
               "Order BY yv03 DESC,yv04 DESC,yv02 ASC "
   '2009/12/21 ADD BY SONIA 加印分所依所別排序及跳頁
   Else
      'Modify By Sindy 2023/2/4 and substr(st01,4,1)<>'9'
      m_str = "SELECT ST01,ST02,YV03,YV04,sqldatet(ST13),ST06,to_char(add_months(to_date(st13,'YYYYMMDD'),6),'YYYYMMDD') as ST13_6,to_char(add_months(to_date(st13,'YYYYMMDD'),12),'YYYYMMDD') as ST13_12,ST13,ST40,YV12,YV13,ST14 " & _
               "FROM YearVacation,staff,SalaryData " & _
               "WHERE yv02=st01(+) and yv01='" & Mid(Trim(DBDATE(txt1(0) & "0101")), 1, 4) & "' " & _
               "and yv02=sd01(+) and substr(st01,4,1)<>'9' " & m_StrSQL & _
               "Order BY ST06,yv03 DESC,yv04 DESC,yv02 ASC "
   End If
   '2009/12/21 END
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       With m_rs
           m_rs.MoveFirst
           'Add By Sindy 2019/6/26
           If Option1(1).Value = True Or bolEPrint = True Then '列印個人紙本通知單
               cnnConnection.BeginTrans: bolConn = True
               strSql = "delete from mailcache where mc01='" & strUserNum & "'"
               cnnConnection.Execute strSql, intI
               Do While Not m_rs.EOF
                  strBackTaieDate = Pub_BackTaieToDate(m_rs.Fields("ST01"), txt1(0), strNote)
                  strTemp(1) = m_rs.Fields("ST02") & "同仁，您好：" & vbCrLf & vbCrLf & _
                     "　　　通知您 " & txt1(0) & " 年度特別休假日數如下:" & vbCrLf & vbCrLf & _
                     "　　　到職日為 " & Val(Left(m_rs.Fields("ST13"), 4)) - 1911 & " 年 " & Mid(m_rs.Fields("ST13"), 5, 2) & " 月 " & Mid(m_rs.Fields("ST13"), 7, 2) & " 日，年資 " & m_rs.Fields("YV03") & " 年" & vbCrLf & vbCrLf
                  If Val(strBackTaieDate) > 0 Then
                     strTemp(1) = strTemp(1) & "　　　" & strNote & _
                        "　　　留職停薪後特休的起算日為 " & Val(Left(strBackTaieDate, 4)) - 1911 & " 年 " & Mid(strBackTaieDate, 5, 2) & " 月 " & Mid(strBackTaieDate, 7, 2) & " 日" & vbCrLf & vbCrLf
                  End If
                  strTemp(1) = strTemp(1) & _
                     "　　　特別休假日數 " & m_rs.Fields("YV04") & " 日" & vbCrLf & vbCrLf
                  If ("" & m_rs.Fields("YV12") = "A" Or "" & m_rs.Fields("YV12") = "B") And _
                     "" & m_rs.Fields("YV13") <> "" And Val(m_rs.Fields("YV04")) < 30 Then
                     strTemp(1) = strTemp(1) & "　　　" & m_rs.Fields("YV13") & vbCrLf & vbCrLf
                  End If
                  
                  'Add By Sindy 2019/7/16
                  If Trim(txtNote) <> "" Then
                     strTemp(1) = strTemp(1) & vbCrLf & "　　　備註：" & Trim(txtNote) & vbCrLf & vbCrLf
                  End If
                  '2019/7/16 END
                  
                  strTemp(1) = strTemp(1) & vbCrLf & "　　　　　　人事處啟" & vbCrLf
                  
                  '電子化人員直接收Mail通知
                  'Modify By Sindy 2023/1/10 And Mid(m_rs.Fields("ST01"), 4, 1) <> "9"
                  If bolEPrint = True And "" & m_rs.Fields("ST14") <> "99997" And Mid(m_rs.Fields("ST01"), 4, 1) <> "9" Then
                     '記錄要發通知確認的E-Mail人員
                     'Add By Sindy 2019/7/16 + 發【修改通知】信件
                     strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08) values(" & _
                              "'" & strUserNum & "','" & m_rs.Fields("ST01") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                              ",'" & IIf(ChkErr.Value = 1, "【修改通知】", "") & txt1(0) & "年度特別休假通知函','" & strTemp(1) & "')"
                     cnnConnection.Execute strSql, intI
                     
                     intECnt = intECnt + 1
                  Else
                     intPCnt = intPCnt + 1
                     If intPCnt = 1 Then
                        Set Printer = Printers(Combo1.ListIndex)
                        Printer.EndDoc
                        Printer.Orientation = 1 '1.直印 2.橫印
                        Printer.PaperSize = 9  'PDF
                     End If
                     
                     Printer.Font.Size = 16
                     If intPCnt Mod 2 = 1 Then
                        If intPCnt > 1 Then Printer.NewPage
                        Printer.CurrentY = 1000
                     Else
                        Printer.CurrentY = 8000
                     End If
                     Printer.CurrentX = 1000
                     Printer.Print strTemp(1)
                  End If
                  
                  m_rs.MoveNext
               Loop
               cnnConnection.CommitTrans: bolConn = False
               
           Else '列印原清單
           '2019/6/26 END
              Set Printer = Printers(Combo1.ListIndex)
              Printer.EndDoc
              Printer.Orientation = 1 '1.直印 2.橫印
              Printer.PaperSize = 9  'PDF
              
              Printer.Font.Size = 12
              iLine = 1 '新的一頁開始
              strType = ""
              intRow = 0
              dblTotDay = 0
              
              Do While Not m_rs.EOF
                  intPCnt = intPCnt + 1
                  For m_i = 1 To 6
                      strTemp(m_i) = ""
                  Next m_i
                  strTemp(1) = CheckStr(.Fields(0)) '員工代號
   '               If Val(CheckStr(.Fields(2))) >= 1 Then
   '                  strTemp(1) = CheckStr(.Fields(0)) '員工代號
   '               Else
   '                  strTemp(1) = CheckStr(.Fields(0)) & "  *"
   '               End If
                  strTemp(2) = CheckStr(.Fields(1)) '姓名
                  strTemp(6) = CheckStr(.Fields(5)) '公司別或所別
                  
                  '為固定任職期間寫死的條件判斷
                  If strTemp(1) = "63001" Or strTemp(1) = "63002" Or _
                     strTemp(1) = "64001" Or strTemp(1) = "65001" Or _
                     strTemp(1) = "68001" Or strTemp(1) = "72010" Then
                     If strTemp(1) = "63001" Then strTemp(3) = "63/02/22 -- 65/09/30"
                     If strTemp(1) = "63002" Then strTemp(3) = "63/02/22 -- 65/09/30"
                     If strTemp(1) = "64001" Then strTemp(3) = "64/03/01 -- 65/09/30"
                     If strTemp(1) = "65001" Then strTemp(3) = "65/05/01 -- 65/09/30"
                     If strTemp(1) = "68001" Then strTemp(3) = "65/04/01 -- 65/09/30"
                     If strTemp(1) = "72010" Then strTemp(3) = "72/04/20 -- 77/06/30"
   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
   '                  If intRow >= 15 Or iLine = 1 Then
   '                       If .AbsolutePosition <> .RecordCount Then
   '                           If strType <> "" Then Printer.NewPage: intRow = 0
   '                           iLine = 1
   '                           Call PrintTitle(strItem)
   '                       End If
   '                  End If
                     PrintDetail
                     'strType = strTemp(6)   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
                  End If
                  
                  strTemp(4) = PUB_ChangeNianZi(Val(CheckStr(.Fields(2))))
                  'Add By Sindy 2016/12/19
                  If CheckStr(.Fields(3)) = 0 Then
                     strTemp(5) = ""
   '                  If .Fields("st01") = "A5017" Then
   '                     MsgBox "TEST"
   '                  End If
                     'Add By Sindy 2017/1/6
                     '當年滿半年者
                     If .Fields("st13_6") <> "" And Left(.Fields("st13_6"), 4) = Val(txt1(0)) + 1911 Then
                        strTemp(5) = ChangeWStringToTDateString(.Fields("st13_6")) & "起 3 天"
                        dblTotDay = dblTotDay + 3 '特休假總日數
                     End If
                     '當年滿1年者
                     If .Fields("st13_12") <> "" And Left(.Fields("st13_12"), 4) = Val(txt1(0)) + 1911 Then
                        strDay = CountRestDay(.Fields("st13"), 0, Val(txt1(0)))
                        strTemp(5) = IIf(strTemp(5) <> "", strTemp(5) & "&", "") & ChangeWStringToTDateString(.Fields("st13_12")) & "起 " & strDay & " 天"
                        dblTotDay = dblTotDay + Val(strDay) '特休假總日數
                     End If
                     '2017/1/6 END
                  Else
                  '2016/12/19 END
                     strTemp(5) = CheckStr(.Fields(3)) & " 天" '特休假天數
                     dblTotDay = dblTotDay + CheckStr(.Fields(3)) '特休假總日數 'Modify By Sindy 2017/1/6
                  End If
                  'dblTotDay = dblTotDay + CheckStr(.Fields(3)) '特休假總日數 'Modify By Sindy 2017/1/6 Mark
                  strTemp(3) = CheckStr(.Fields(4)) '到職日
                  
                  'Add By Sindy 2014/12/29 檢查最後一筆是否為04留職停薪
                  BolChkEnd04 = False
                  strSql = "select * from staff_change where sc01='" & strTemp(1) & "' " & _
                           "and sc02=(select max(sc02) from staff_change where sc01='" & strTemp(1) & "') " & _
                           "and sc03='04'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     BolChkEnd04 = True
                  End If
                  '2014/12/29 END
                  
                  '任職時間
                  m_str2 = "select sqldatet(sc02) as 日期,sc03 " & _
                           "from staff_change " & _
                           "where sc03 in ('02','03','04','08','09','10') and sc01='" & strTemp(1) & "' " & _
                           "order by sc02"
                  If m_rs2.State = 1 Then m_rs2.Close
                  m_rs2.CursorLocation = adUseClient
                  m_rs2.Open m_str2, cnnConnection, adOpenStatic, adLockReadOnly
                  If Not m_rs2.EOF And Not m_rs2.BOF Then
                      m_rs2.MoveFirst
                      int_i = 0
                      Do While Not m_rs2.EOF
                          int_i = int_i + 1
                          '有到職日並且異動資料只有一筆
                          If strTemp(3) <> "" And m_rs2.RecordCount = 1 Then
                              If Val(ChangeTDateStringToTString(strTemp(3))) <= Val(ChangeWStringToTString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "0101"))), "yyyy") & "1231")) Then
                                 If CheckStr(m_rs2.Fields(1)) = "04" Then
                                    strTemp(3) = strTemp(3) & " -- " & ChangeTStringToTDateString(ChangeWDateStringToTString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0))))))))
                                 Else
                                    strTemp(3) = strTemp(3) & " -- " & ChangeWStringToTDateString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "0101"))), "yyyy") & "1231")
                                 End If
                              Else
                                  strTemp(3) = strTemp(3) & " -- "
                              End If
   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
   '                           If intRow >= 15 Or iLine = 1 Then
   '                                If .AbsolutePosition <> .RecordCount Then
   '                                    If strType <> "" Then Printer.NewPage: intRow = 0
   '                                    iLine = 1
   '                                    Call PrintTitle(strItem)
   '                                End If
   '                           End If
                              PrintDetail
                              'strType = strTemp(6)   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
                              
                          '有到職日並且非異動資料的最後一筆
                          ElseIf strTemp(3) <> "" And m_rs2.RecordCount > 1 And m_rs2.AbsolutePosition <> m_rs2.RecordCount Then
                              If m_rs2.AbsolutePosition <> 1 Then
                                  strTemp(1) = ""
                                  strTemp(2) = ""
                              End If
                              strTemp(3) = strTemp(3) & " -- " & ChangeTStringToTDateString(ChangeWDateStringToTString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0))))))))
                              strTemp(4) = ""
                              strTemp(5) = ""
                              If int_i Mod 2 <> 0 Then
   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
   '                                 If intRow >= 15 Or iLine = 1 Then
   '                                      If .AbsolutePosition <> .RecordCount Then
   '                                          If strType <> "" Then Printer.NewPage: intRow = 0
   '                                          iLine = 1
   '                                          Call PrintTitle(strItem)
   '                                      End If
   '                                 End If
                                    PrintDetail
                                    'strType = strTemp(6)   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
                              End If
                              strTemp(3) = PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0)))
                          Else
                              strTemp(1) = ""
                              strTemp(2) = ""
                              strTemp(3) = strTemp(3) & " -- " & ChangeTStringToTDateString(ChangeWDateStringToTString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0))))))))
                              If int_i Mod 2 <> 0 Then
   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
   '                                 If iLine = 1 Then
   '                                      If .AbsolutePosition <> .RecordCount Then
   '                                          If strType <> "" Then Printer.NewPage
   '                                          iLine = 1
   '                                          Call PrintTitle(strItem)
   '                                      End If
   '                                 End If
                                    
                                    'Add By Sindy 2014/12/29
                                    If BolChkEnd04 = True Then
                                       strTemp(4) = PUB_ChangeNianZi(Val(CheckStr(.Fields(2))))
                                       'Add By Sindy 2016/12/19
                                       If CheckStr(.Fields(3)) = 0 Then
                                          strTemp(5) = ""
                                       Else
                                       '2016/12/19 END
                                          strTemp(5) = CheckStr(.Fields(3)) & " 天"
                                       End If
                                    End If
                                    '2014/12/29 END
                                    
                                    PrintDetail
                                    
                                    'Add By Sindy 2014/12/29
                                    If BolChkEnd04 = True Then
                                       Exit Do
                                    End If
                                    '2014/12/29 END
                                    
                                    'strType = strTemp(6)   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
                              End If
                              strTemp(3) = PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0)))
                              
                              If Val(ChangeTDateStringToTString(strTemp(3))) <= Val(ChangeWStringToTString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "0101"))), "yyyy") & "1231")) Then
                                 If CheckStr(m_rs2.Fields(1)) = "04" Then
                                    strTemp(3) = strTemp(3) & " -- " & ChangeTStringToTDateString(ChangeWDateStringToTString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0))))))))
                                 Else
                                    strTemp(3) = strTemp(3) & " -- " & ChangeWStringToTDateString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "0101"))), "yyyy") & "1231")
                                 End If
                              Else
                                  strTemp(3) = strTemp(3) & " -- "
                              End If
                              strTemp(4) = PUB_ChangeNianZi(Val(CheckStr(.Fields(2))))
                              'Add By Sindy 2016/12/19
                              If CheckStr(.Fields(3)) = 0 Then
                                 strTemp(5) = ""
                              Else
                              '2016/12/19 END
                                 strTemp(5) = CheckStr(.Fields(3)) & " 天"
                              End If
   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
   '                           If intRow >= 15 Or iLine = 1 Then
   '                                If .AbsolutePosition <> .RecordCount Then
   '                                    If strType <> "" Then Printer.NewPage: intRow = 0
   '                                    iLine = 1
   '                                    Call PrintTitle(strItem)
   '                                End If
   '                           End If
                              PrintDetail
                              'strType = strTemp(6)   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
                          End If
                          m_rs2.MoveNext
                      Loop
                  '沒異動資料
                  Else
                      '為固定任職期間寫死的條件判斷
                      If strTemp(1) = "63001" Or strTemp(1) = "63002" Or _
                        strTemp(1) = "64001" Or strTemp(1) = "65001" Or _
                        strTemp(1) = "68001" Or strTemp(1) = "72010" Then
                        If strTemp(1) = "72010" Then strTemp(3) = "77/07/01"
                        strTemp(1) = "": strTemp(2) = ""
                      End If
                      
                      If Val(ChangeTDateStringToTString(strTemp(3))) <= Val(ChangeWStringToTString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "0101"))), "yyyy") & "1231")) Then
                          strTemp(3) = strTemp(3) & " -- " & ChangeWStringToTDateString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "0101"))), "yyyy") & "1231")
                      Else
                          strTemp(3) = strTemp(3) & " -- "
                      End If
   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
   '                   If intRow >= 15 Or iLine = 1 Then
   '                       If .AbsolutePosition <> .RecordCount Then
   '                           If strType <> "" Then Printer.NewPage: intRow = 0
   '                           iLine = 1
   '                           Call PrintTitle(strItem)
   '                       End If
   '                   End If
                      PrintDetail
                      'strType = strTemp(6)   '2009/12/21 CANCEL BY SONIA 移至PrintDetail
                  End If
                  
      '            Printer.CurrentX = 500
      '            Printer.CurrentY = iLine * 300
      '            Printer.Print String(140, "-")
                  Printer.Line (500, iLine * 300)-(11000, iLine * 300), , B
                  iLine = iLine + 1
                  intRow = intRow + 1
                  m_rs.MoveNext
              Loop
              If txt1(1) = "" Then '2009/12/21分所不印合計
                  '合計
                  Printer.CurrentX = PLeft(4) - Printer.TextWidth("合　計：")
                  Printer.CurrentY = iLine * 300
                  Printer.Print "合　計："
                  Printer.CurrentX = PLeft(5) - Printer.TextWidth(dblTotDay & "天")
                  Printer.CurrentY = iLine * 300
                  Printer.Print dblTotDay & "天"
              End If
            End If
       End With
   Else
      StrMenu = False
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   'Modify By Sindy 2019/6/26
   If intECnt > 0 Then
      Screen.MousePointer = vbHourglass
      Call PUB_SendMailCache(, True)
      Screen.MousePointer = vbDefault
   End If
   If intPCnt > 0 Then
      Printer.EndDoc
      'If strItem = 2 Then
      'If txt1(1) <> "" Or strItem = 3 Then
      ShowPrintOk
      'End If
   End If
   '2019/6/26 END
   
   Exit Function

ErrHnd:
   If bolConn = True Then cnnConnection.RollbackTrans: bolConn = False
   
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function

Sub PrintTitle()

GetPleft

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

iLine = iLine + 2
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(txt1(0) & "年度特別假核給表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print txt1(0) & "年度特別假核給表"

iLine = iLine + 2
Printer.Font.Size = 12
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.Font.Size = 14
'2009/12/21 MODIFY BY SONIA 加印分所
'If strItem = 2 Then
If strItem = 2 Or strItem = 3 Or txt1(1) <> "" Then
'2009/12/21 END
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   '2009/12/21 MODIFY BY SONIA 加印分所不分公司別
   'Printer.Print "公司別：" & strTemp(6)
   If txt1(1) = "" Then
      Printer.Print "公司別：" & strTemp(6)
   Else
      Select Case strTemp(6)
         Case "2"
            Printer.Print "所別：中所"
         Case "3"
            Printer.Print "所別：南所"
         Case "4"
            Printer.Print "所別：高所"
         Case "5"
            Printer.Print "所別：其他"
      End Select
   End If
   '2009/12/21 END
End If

Printer.Font.Size = 12
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

Printer.Font.Size = 14
'iLine = iLine + 2
'Printer.CurrentX = 1500
'Printer.CurrentY = iLine * 300
'Printer.Print "備註：打＊者須於入所滿一年方可申請休特別假。"

iLine = iLine + 3
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "員　編"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "任　職　時　間"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("年　資")
Printer.CurrentY = iLine * 300
Printer.Print "年　資"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("特別假天數")
Printer.CurrentY = iLine * 300
Printer.Print "特別假天數"

iLine = iLine + 1
'Printer.CurrentX = 500
'Printer.CurrentY = iLine * 300
'Printer.Print String(140, "-")
Printer.Line (500, iLine * 300)-(11000, iLine * 300), , B

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 800
PLeft(2) = 2000
PLeft(3) = 4000
PLeft(4) = 7500
PLeft(5) = 9500
End Sub

Sub PrintDetail()
Dim i As Integer, k As Integer
Dim varTmp As Variant
   
   '2009/12/21 ADD BY SONIA
   If intRow >= 15 Or iLine = 1 Or (txt1(1) <> "" And strType <> strTemp(6)) Then
'      If strType <> "" Then
         If strType <> "" Then Printer.NewPage
         intRow = 0
         iLine = 1
         Call PrintTitle
'      End If
   End If
   '2009/12/21 END
   
   varTmp = Split(strTemp(5), "&")
   For k = 0 To IIf(UBound(varTmp) < 0, 0, UBound(varTmp))
      If UBound(varTmp) >= 0 Then strTemp(5) = varTmp(k)
      For i = 1 To 5
         If i = 4 Or (i = 5 And InStr(strTemp(5), "起") = 0) Then
            Printer.CurrentX = PLeft(i) - Printer.TextWidth(strTemp(i))
         Else
            If i = 5 Then
               Printer.CurrentX = 8000
            Else
               Printer.CurrentX = PLeft(i)
            End If
         End If
         Printer.CurrentY = iLine * 300
         Printer.Print strTemp(i)
      Next i
      If k < UBound(varTmp) Then
         For i = 1 To 4
            strTemp(i) = ""
         Next i
         iLine = iLine + 1
      End If
   Next k
'   If Mid(strTemp(1), 1, 5) = "98029" Then
'      Debug.Print
'   End If
   iLine = iLine + 1
   strType = strTemp(6)    '2009/12/21 ADD BY SONIA
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
   Set frm160306 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      '2009/12/21 ADD BY SONIA
      Case 1
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
            Beep
         End If
      '2009/12/21 END
      'Add By Sindy 2019/6/26
      Case 3, 4
         KeyAscii = UpperCase(KeyAscii)
      '2019/6/26 END
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(0).Text = "" Then Exit Sub
         If CheckIsTaiwanDate(txt1(0).Text & "0101", False) = False Then
             Cancel = True
             MsgBox "請輸入民國年度！", vbInformation, "輸入新年度錯誤"
             Exit Sub
         End If
      'Add By Sindy 2019/6/26
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
      '2019/6/26 END
      Case Else
   End Select
End Sub
