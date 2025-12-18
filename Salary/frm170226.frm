VERSION 5.00
Begin VB.Form frm170226 
   BorderStyle     =   1  '單線固定
   Caption         =   "薪資扣繳表－人事用"
   ClientHeight    =   3130
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   4690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3130
   ScaleWidth      =   4690
   Begin VB.CheckBox chk1 
      Caption         =   "含加班費"
      Height          =   255
      Left            =   870
      TabIndex        =   4
      Top             =   1710
      Width           =   1365
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2745
      MaxLength       =   5
      TabIndex        =   1
      Top             =   900
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   1830
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   0
      TabIndex        =   10
      Top             =   2400
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   5
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
      Height          =   255
      Index           =   4
      Left            =   2730
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1350
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1830
      MaxLength       =   5
      TabIndex        =   0
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3570
      TabIndex        =   7
      Top             =   180
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2490
      TabIndex        =   6
      Top             =   180
      Width           =   975
   End
   Begin VB.Line Line3 
      X1              =   2325
      X2              =   2945
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line2 
      X1              =   2340
      X2              =   3000
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   870
      TabIndex        =   9
      Top             =   1380
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "扣繳年月："
      Height          =   180
      Left            =   870
      TabIndex        =   8
      Top             =   930
      Width           =   900
   End
End
Attribute VB_Name = "frm170226"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/1/20 ADD BY SONIA 不管年終獎金及每月獎金
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL1 As String     '每月薪資資料
Dim m_StrSQL2 As String     '年終獎金資料 add by sonia 2016/7/13
Dim m_StrSQL3 As String     '每月獎金資料 add by sonia 2016/7/13
Dim m_i As Integer
Dim m_month As Integer      '列印中之月份(因為無該月資料也要空行)
Dim PLeft(1 To 20) As Integer
Dim strTemp(1 To 20) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblAmtS As Double, dblTAmtT As Double
Dim dblTAmtT1 As Double, dblTAmtT2 As Double, dblTAmtT3 As Double, dblTAmtT4 As Double, dblTAmtT5 As Double, dblTAmtT6 As Double, dblTAmtT7 As Double
Dim dblTAmtT8 As Double 'Add By Sindy 2024/8/8


Private Sub cmdok_Click(Index As Integer)
Dim strYM As String

   Select Case Index
      Case 0
         If txt1(1) = "" And txt1(2) = "" Then
            MsgBox "扣繳年月不可空白！", vbInformation, "操作錯誤！"
            txt1(1).SetFocus
            Exit Sub
         End If
         If RunNick(Val(txt1(1)), Val(txt1(2))) Then
            txt1(1).SetFocus
            Exit Sub
         End If
         '2009/5/13 add by sonia
         If Len(txt1(1)) = 5 Then
            If Left(Trim(txt1(1)), 3) <> Left(Trim(txt1(2)), 3) Then
               MsgBox "起迄扣繳年度不同！", vbInformation, "操作錯誤！"
               txt1(1).SetFocus
               Exit Sub
            End If
         Else
            If Left(Trim(txt1(1)), 2) <> Left(Trim(txt1(2)), 2) Then
               MsgBox "起迄扣繳年度不同！", vbInformation, "操作錯誤！"
               txt1(1).SetFocus
               Exit Sub
            End If
         End If
         '2009/5/13 end
         '2009/5/13 cancel by sonia
         'If txt1(3) = "" And txt1(4) = "" Then
         '   MsgBox "員工編號不可空白！", vbInformation, "操作錯誤！"
         '   txt1(3).SetFocus
         '   Exit Sub
         'End If
         '2009/5/13 end
         If RunNick(txt1(3), txt1(4)) Then
            txt1(3).SetFocus
            Exit Sub
         End If
           
         Screen.MousePointer = vbHourglass
         StrMenu
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Sub StrMenu()

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   
   m_StrSQL1 = "": m_StrSQL2 = "": m_StrSQL3 = ""
   If txt1(1) <> "" Then '扣繳年月起
      m_StrSQL1 = m_StrSQL1 & " and SM02 >=" & Val(txt1(1)) + 191100
      'add by sonia 2016/7/13
      m_StrSQL2 = m_StrSQL2 & " yb01=" & Left(Val(txt1(1)) + 191100 - 100, 4)
      m_StrSQL3 = m_StrSQL3 & " substr(MB01,1,4)='" & Left(Val(txt1(1)) + 191100, 4) & "' "
      'end 2016/7/13
   End If
   If txt1(2) <> "" Then '扣繳年月起
      m_StrSQL1 = m_StrSQL1 & " and SM02 <=" & Val(txt1(2)) + 191100
      m_StrSQL3 = m_StrSQL3 & " and 14+TRUNC((SUBSTR(MB01,5,2)-1)/4) <=" & Val(Right(txt1(2), 2))  'add by sonia 2016/8/31
   End If
   If txt1(3) <> "" Then '員工編號起
      '2009/5/13 modify by sonia 辜說輸條件時只抓該編號,不要抓第二家
      'm_StrSQL1 = m_StrSQL1 & " and replace(SM01,'A','0') >='" & Trim(txt1(3)) & "' "
      m_StrSQL1 = m_StrSQL1 & " and SM01>='" & Trim(txt1(3)) & "' "
      '2009/5/13 end
      'add by sonia 2016/7/13
      m_StrSQL2 = m_StrSQL2 & " and YB02>='" & Trim(txt1(3)) & "' "
      m_StrSQL3 = m_StrSQL3 & " and MB02>='" & Trim(txt1(3)) & "' "
      'end 2016/7/13
   End If
   If txt1(4) <> "" Then '員工編號迄
      '2009/5/13 modify by sonia 辜說輸條件時只抓該編號,不要抓第二家
      'm_StrSQL1 = m_StrSQL1 & " and replace(SM01,'A','0') <='" & Trim(txt1(4)) & "' "
      m_StrSQL1 = m_StrSQL1 & " and SM01<='" & Trim(txt1(4)) & "' "
      '2009/5/13 end
      'add by sonia 2016/7/13
      m_StrSQL2 = m_StrSQL2 & " and YB02<='" & Trim(txt1(4)) & "' "
      m_StrSQL3 = m_StrSQL3 & " and MB02<='" & Trim(txt1(4)) & "' "
      'end 2016/7/13
   End If
   
   '每月薪資資料
   '2009/5/13 modify by sonia 辜說應稅薪資不含超時加班費
   'm_str = "SELECT 公司別,員工編號,月份,應稅薪資,其他給付,所得稅,加班費,午餐津貼,ST02 姓名,ST26 身分證字號,ST34 戶籍地址,DECODE(SD03,'Y','有','無') 配偶,NVL(SD07,0) 扶養人數,A0802 公司名稱 FROM STAFF,SALARYDATA,ACC080,(" & _
           "SELECT SM37 公司別,SM01 員工編號,TO_CHAR(SUBSTR(SM02,5,2),'99') 月份," & _
           "NVL(SM04,0)+NVL(SM05,0)+NVL(SM28,0)-NVL(SM21,0) 應稅薪資," & _
           "NVL(OD03,0) 其他給付,NVL(SM24,0) 所得稅,NVL(SM12,0) 加班費,NVL(SM07,0) 午餐津貼 " & _
           "From SALARYMONTH, OtherPayData WHERE SM02=OD01(+) AND SM01=OD02(+) " & m_StrSQL1 & ")" & _
           "WHERE 公司別=A0801(+) AND REPLACE(員工編號,'A','0')=ST01(+) AND REPLACE(員工編號,'A','0')=SD01(+) ORDER BY REPLACE(員工編號,'A','0'),公司別,月份"
   'Modify by Morgan 2010/12/2 修正員工編號第一碼可以是英文問題
   'modify by sonia 2015/12/24 扣繳表之加班費欄應扣除超時加班費SM28
   'modify by sonia 2016/7/13 辜說要開放可跑年終13月及三節14,15,16月的資料
   'm_str = "SELECT 公司別,員工編號,月份,應稅薪資,其他給付,所得稅,加班費,午餐津貼,ST02 姓名,ST26 身分證字號,ST34 戶籍地址,DECODE(SD03,'Y','有','無') 配偶,NVL(SD07,0) 扶養人數,A0802 公司名稱 FROM STAFF,SALARYDATA,ACC080,(" & _
           "SELECT SM37 公司別,SM01 員工編號,TO_CHAR(SUBSTR(SM02,5,2),'99') 月份," & _
           "NVL(SM04,0)+NVL(SM05,0)-NVL(SM21,0) 應稅薪資," & _
           "NVL(OD03,0) 其他給付,NVL(SM24,0) 所得稅,NVL(SM12,0)-NVL(SM28,0) 加班費,NVL(SM07,0) 午餐津貼 " & _
           "From SALARYMONTH, OtherPayData WHERE SM02=substr(OD01(+),1,6) AND SM01=OD02(+) " & m_StrSQL1 & ")" & _
           "WHERE 公司別=A0801(+) AND substr(員工編號,1,1)||replace(substr(員工編號,2),'A','0')=ST01(+) AND substr(員工編號,1,1)||replace(substr(員工編號,2),'A','0')=SD01(+) ORDER BY substr(員工編號,1,1)||replace(substr(員工編號,2),'A','0'),公司別,月份"
   'modify by sonia 2018/7/19 辜說同仁其他給付OtherPayData都不抓,但格式保留
   'm_str = "SELECT 公司別,員工編號,月份,應稅薪資,其他給付,所得稅,加班費,午餐津貼,ST02 姓名,ST26 身分證字號,ST34 戶籍地址,DECODE(SD03,'Y','有','無') 配偶,NVL(SD07,0) 扶養人數,A0802 公司名稱 FROM STAFF,SALARYDATA,ACC080,(" & _
           "SELECT SM37 公司別,SM01 員工編號,TO_CHAR(SUBSTR(SM02,5,2),'99') 月份," & _
           "NVL(SM04,0)+NVL(SM05,0)-NVL(SM21,0) 應稅薪資," & _
           "NVL(OD03,0) 其他給付,NVL(SM24,0) 所得稅,NVL(SM12,0)-NVL(SM28,0) 加班費,NVL(SM07,0) 午餐津貼 " & _
           "From SALARYMONTH, OtherPayData WHERE SM02=substr(OD01(+),1,6) AND SM01=OD02(+) " & m_StrSQL1 & " "
   'Modify By Sindy 2020/6/25 + 證照津貼
   'Modify By Sindy 2024/8/8 +,NVL(SM14,0)+NVL(SM15,0) 勞健保費
   m_str = "SELECT 公司別,員工編號,月份,應稅薪資,其他給付,所得稅,加班費,午餐津貼,ST02 姓名,ST26 身分證字號,ST34 戶籍地址,DECODE(SD03,'Y','有','無') 配偶,NVL(SD07,0) 扶養人數,A0802 公司名稱,勞健保費 FROM STAFF,SALARYDATA,ACC080,(" & _
           "SELECT SM37 公司別,SM01 員工編號,TO_CHAR(SUBSTR(SM02,5,2),'99') 月份," & _
           "NVL(SM04,0)+NVL(SM05,0)+NVL(SM45,0)-NVL(SM21,0) 應稅薪資," & _
           "0 其他給付,NVL(SM24,0) 所得稅,NVL(SM12,0)-NVL(SM28,0) 加班費,NVL(SM07,0) 午餐津貼,NVL(SM14,0)+NVL(SM15,0) 勞健保費 " & _
           "From SALARYMONTH WHERE SM02=SM02 " & m_StrSQL1 & " "
   
   '年終獎金資料(前一年)
   If Val(Right(txt1(1), 2)) >= 13 Or (Val(Right(txt1(1), 2)) < 13 And Val(Right(txt1(2), 2)) >= 13) Then
      'modify by sonia 2017/9/12 區分年終獎金及特殊功績獎金,且二者改列其他給付欄
      'm_str = m_str & " UNION " & _
              "SELECT YB24 公司別,YB02 員工編號,TO_CHAR('13','99') 月份,NVL(YB05,0)+NVL(YB06,0)-NVL(YB15,0) 年終獎金,0,NVL(YB17,0) 所得稅,0,0 " & _
              "FROM YEARBONUS WHERE " & m_StrSQL2
      'Modify By Sindy 2024/8/8 +,0
      m_str = m_str & " UNION " & _
              "SELECT YB24 公司別,YB02 員工編號,TO_CHAR('13','99') 月份,0,NVL(YB05,0)-NVL(YB15,0) 年終獎金,NVL(YB17,0) 所得稅,0,0,0 " & _
              "FROM YEARBONUS WHERE " & m_StrSQL2
      'modify by sonia 2018/1/11 +YB26
      'Modify By Sindy 2024/8/8 +,0
      m_str = m_str & " UNION " & _
              "SELECT YB24 公司別,YB02 員工編號,TO_CHAR('131','999') 月份,0,NVL(YB06,0)+NVL(YB26,0) 特殊功績獎金,0 所得稅,0,0,0 " & _
              "FROM YEARBONUS WHERE " & m_StrSQL2
      'end 2017/9/12
   End If
   '每月獎金資料, 但第14,15,16月獎金要合併
   If Val(Right(txt1(1), 2)) >= 14 Or (Val(Right(txt1(1), 2)) < 14 And Val(Right(txt1(2), 2)) >= 14) Then
      'modify by sonia 2017/9/12 改列其他給付欄
      'Modified by Morgan 2024/5/10 修正員工編號第五碼可以是英文問題
      'Modify By Sindy 2024/8/8 +,0
      m_str = m_str & " UNION " & _
              "SELECT comp 公司別,MB02 員工編號,MB01A 月份," & _
              "0,NVL(sum(total),0) 獎金總額,NVL(sum(tax),0) 扣繳稅額,0,0,0 from (" & _
              "SELECT MB11 comp, TO_CHAR(14+TRUNC((SUBSTR(MB01,5,2)-1)/4),'99') MB01A,MB02,NVL(MB03,0) total,NVL(MB04,0) tax,MB01 " & _
              "FROM MonthBonus,salarydata WHERE " & m_StrSQL3 & " AND substr(mb02,3,1)='0' AND MB02=Sd01(+) UNION " & _
              "SELECT MB11 comp, TO_CHAR(14+TRUNC((SUBSTR(MB01,5,2)-1)/4),'99') MB01A,MB02,NVL(MB03,0) total,NVL(MB04,0) tax,MB01 " & _
              "FROM MonthBonus,salarydata WHERE " & m_StrSQL3 & " AND substr(mb02,3,1)='A' AND substr(MB02,1,2)||replace(substr(MB02,3,1),'A','0')||substr(MB02,4)=Sd01(+)) " & _
              " GROUP BY comp,MB02,MB01A"
   End If
   m_str = m_str & ") WHERE 公司別=A0801(+) AND substr(員工編號,1,2)||replace(substr(員工編號,3,1),'A','0')||substr(員工編號,4)=ST01(+) AND substr(員工編號,1,2)||replace(substr(員工編號,3,1),'A','0')||substr(員工編號,4)=SD01(+) ORDER BY substr(員工編號,1,2)||replace(substr(員工編號,3,1),'A','0')||substr(員工編號,4),公司別,月份"
   'end 2016/7/13
   
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         '預設值
         iLine = 1
         strType = "" '切頁條件
         dblTAmtT1 = 0
         dblTAmtT2 = 0
         dblTAmtT3 = 0
         dblTAmtT4 = 0
         dblTAmtT5 = 0
         dblTAmtT6 = 0
         dblTAmtT7 = 0
         dblTAmtT8 = 0 'Add By Sindy 2024/8/8
         
         Do While Not m_rs.EOF
             
            For m_i = 1 To 15 '14
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(0))   '公司別
            strTemp(2) = CheckStr(m_rs.Fields(1))   '員工編號
            strTemp(3) = CheckStr(m_rs.Fields(2))   '月份
            strTemp(4) = CheckStr(m_rs.Fields(3))   '應稅薪資
            '2009/5/13 modify by sonia 辜說不列印此欄
            'modify by sonia 2017/9/12 所有獎金改回其他給付
            strTemp(5) = CheckStr(m_rs.Fields(4))   '其他給付
            'strTemp(5) = 0                          '其他給付
            'end 2017/9/12
            strTemp(6) = CheckStr(m_rs.Fields(5))   '所得稅
            'Modify By Sindy 2012/5/11 增加判斷是否要含加班費
            If chk1.Value = 0 Then
               '不含加班費
               strTemp(7) = 0
            Else
            '2012/5/11 End
               '含加班費
               strTemp(7) = CheckStr(m_rs.Fields(6))   '加班費
            End If
            strTemp(8) = CheckStr(m_rs.Fields(7))   '午餐津貼
            strTemp(9) = CheckStr(m_rs.Fields(8))   '姓名
            strTemp(10) = CheckStr(m_rs.Fields(9))  '身份證字號
            strTemp(11) = CheckStr(m_rs.Fields(10)) '戶籍地址
            strTemp(12) = CheckStr(m_rs.Fields(11)) '配偶
            strTemp(13) = CheckStr(m_rs.Fields(12)) '扶養人數
            strTemp(14) = CheckStr(m_rs.Fields(13)) '公司名稱
            strTemp(15) = CheckStr(m_rs.Fields(14)) '勞健保費 Add By Sindy 2024/8/8
            
            If iLine > 50 Or iLine = 1 Or (strType <> strTemp(1) & strTemp(2)) Then
               If strType <> "" And strType <> strTemp(1) & strTemp(2) Then
                  PrintEnd '個人合計
               End If
               
               If iLine <> 1 Then Printer.NewPage
               iLine = 1
               PrintTitle '列印表頭
            End If
            
            dblAmtS = Val(strTemp(4)) + Val(strTemp(5))                               '給付合計
            'Modify By Sindy 2024/8/8 增加 - Val(strTemp(16))
            dblTAmtT = Val(dblAmtS) - Val(strTemp(6)) + Val(strTemp(7)) + Val(strTemp(8)) - Val(strTemp(15)) '給付實額
            
            'modify by sonia 2017/9/12
            'If strTemp(3) > m_month Then  '若該月無資料也要印空行
            '   PrintEmpty m_month, Val(strTemp(3)) - 1 '列印空行
            'End If
            If Val(Left(strTemp(3), 2)) > m_month Then '若該月無資料也要印空行
               PrintEmpty m_month, Val(Left(strTemp(3), 2)) - 1 '列印空行
            End If
            'end 2017/9/12
            
            PrintDetail '列印表中
            
            strType = strTemp(1) & strTemp(2) '依公司別,員工編號跳頁
            
            '個人合計 ********************************
            dblTAmtT1 = dblTAmtT1 + Val(strTemp(4))
            dblTAmtT2 = dblTAmtT2 + Val(strTemp(5))
            dblTAmtT3 = dblTAmtT3 + Val(dblAmtS)
            dblTAmtT4 = dblTAmtT4 + Val(strTemp(6))
            dblTAmtT5 = dblTAmtT5 + Val(strTemp(7))
            dblTAmtT6 = dblTAmtT6 + Val(strTemp(8))
            dblTAmtT7 = dblTAmtT7 + Val(dblTAmtT)
            dblTAmtT8 = dblTAmtT8 + Val(strTemp(15)) 'Add By Sindy 2024/8/8 勞健保費
            
            m_rs.MoveNext
         Loop
         
         '列印表尾
         PrintEnd '個人合計
           
      End With
   Else
      MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
      Exit Sub
   End If
   
   Printer.EndDoc
   ShowPrintOk

End Sub

Sub PrintEnd()
   
   If m_month < 12 Then  '若該月無資料也要印空行
      PrintEmpty m_month, 12 '列印空行
   End If
            
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "合計"
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(dblTAmtT1, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblTAmtT1, "##,###,###")
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblTAmtT2, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblTAmtT2, "##,###,###")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblTAmtT3, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblTAmtT3, "##,###,###")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblTAmtT4, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblTAmtT4, "##,###,###")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(dblTAmtT5, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblTAmtT5, "##,###,###")
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(dblTAmtT6, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblTAmtT6, "##,###,###")
   'Add By Sindy 2024/8/8 +勞健保費
   Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(dblTAmtT8, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblTAmtT8, "##,###,###")
   '2024/8/8 END
   Printer.CurrentX = PLeft(9) - Printer.TextWidth(Format(dblTAmtT7, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblTAmtT7, "##,###,###")

   dblTAmtT1 = 0
   dblTAmtT2 = 0
   dblTAmtT3 = 0
   dblTAmtT4 = 0
   dblTAmtT5 = 0
   dblTAmtT6 = 0
   dblTAmtT7 = 0
   dblTAmtT8 = 0 'Add By Sindy 2024/8/8
End Sub

Sub PrintTitle()

   GetPleft
   
   Printer.Font.Size = 14
   Printer.Font.Underline = True
   Printer.FontBold = True
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTemp(14)) / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(14)
   
   iLine = iLine + 1
   'modify by sonia 2016/7/13 劉經理要求改名稱
   'Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("薪　資　扣　繳　表") / 2)
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("薪　　　資　　　表") / 2)
   Printer.CurrentY = iLine * 300
   'modify by sonia 2016/7/13 劉經理要求改名稱
   'Printer.Print "薪　資　扣　繳　表"
   Printer.Print "薪　　　資　　　表"
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   '2009/5/13 modify by sonia
   'Printer.Print "扣繳年度：" & Val(txt1(1))
   'modify by sonia 2016/7/13 '扣繳年度'改'年度'
   If Len(txt1(1)) = 5 Then
      Printer.Print "年度：" & Left(Trim(txt1(1)), 3)
   Else
      Printer.Print "年度：" & Left(Trim(txt1(1)), 2)
   End If
   '2009/5/13 end
   Printer.CurrentX = 8000
   Printer.CurrentY = iLine * 300
   Printer.Print "編號　　　字第　　　　　　號"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　　名：" & strTemp(9)
   Printer.CurrentX = 8500
   Printer.CurrentY = iLine * 300
   Printer.Print "身份證字號：" & strTemp(10)
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "地址：" & strTemp(11)
   Printer.CurrentX = 8500
   Printer.CurrentY = iLine * 300
   Printer.Print "配偶：" & strTemp(12) & "　　扶養人數：" & strTemp(13)

   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "所得"
   Printer.CurrentX = 1200
   Printer.CurrentY = iLine * 300
   Printer.Print "< 給　　　付　　　明　　　細 >　　代　扣"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "月份"
   Printer.CurrentX = PLeft(2) - Printer.TextWidth("　薪　資")
   Printer.CurrentY = iLine * 300
   Printer.Print "　薪　資"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("其他給付")
   Printer.CurrentY = iLine * 300
   Printer.Print "其他給付"
   Printer.CurrentX = PLeft(4) - Printer.TextWidth("合　　計")
   Printer.CurrentY = iLine * 300
   Printer.Print "合　　計"
   Printer.CurrentX = PLeft(5) - Printer.TextWidth("所得稅")
   Printer.CurrentY = iLine * 300
   Printer.Print "所得稅"
   Printer.CurrentX = PLeft(6) - Printer.TextWidth("加班費")
   Printer.CurrentY = iLine * 300
   Printer.Print "加班費"
   Printer.CurrentX = PLeft(7) - Printer.TextWidth("午餐津貼")
   Printer.CurrentY = iLine * 300
   Printer.Print "午餐津貼"
   'Add By Sindy 2024/8/8 +勞健保費
   Printer.CurrentX = PLeft(8) - Printer.TextWidth("勞健保費")
   Printer.CurrentY = iLine * 300
   Printer.Print "勞健保費"
   '2024/8/8 END
   Printer.CurrentX = PLeft(9) - Printer.TextWidth("給付實額")
   Printer.CurrentY = iLine * 300
   Printer.Print "給付實額"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(145, "-")
   
   iLine = iLine + 1
   
   m_month = 1
End Sub

Sub GetPleft()
   'Modify By Sindy 2024/8/8
   PLeft(1) = 700
   PLeft(2) = 2000 '薪資
   PLeft(3) = 3100 '3300 '其他給付
   PLeft(4) = 4400 '4600 '合計
   PLeft(5) = 5500 '5900 '所得稅
   PLeft(6) = 6500 '7200 '加班費
   PLeft(7) = 7500 '8500 '午餐津貼
   PLeft(8) = 8500 '勞健保費
   PLeft(9) = 9800 '9800 '給付實額
End Sub

Sub PrintDetail()
   
   m_month = m_month + 1
         
   'modify by sonia 2016/7/13
   'If strTemp(3) = 1 Then iLine = iLine + 1
   'Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(strTemp(3), "##"))
   'Printer.CurrentY = iLine * 300
   'Printer.Print Format(strTemp(3), "##")
   Select Case strTemp(3)
      Case 13
         Printer.CurrentX = PLeft(1) - Printer.TextWidth("獎金")
         Printer.CurrentY = iLine * 300
         Printer.Print "年終"
         iLine = iLine + 1
         Printer.CurrentX = PLeft(1) - Printer.TextWidth("獎金")
         Printer.CurrentY = iLine * 300
         Printer.Print "獎金"
      'add by sonia 2017/9/12
      Case 131
'modify by sonia 2018/7/19 特殊功績獎金改紅利
'         Printer.CurrentX = PLeft(1) - Printer.TextWidth("獎金")
'         Printer.CurrentY = iLine * 300
'         Printer.Print "特殊功"
'         iLine = iLine + 1
'         Printer.CurrentX = PLeft(1) - Printer.TextWidth("獎金")
'         Printer.CurrentY = iLine * 300
'         Printer.Print "績獎金"
         Printer.CurrentX = PLeft(1) - Printer.TextWidth("紅利")
         Printer.CurrentY = iLine * 300
         Printer.Print "紅利"
      'end 2017/9/12
      Case 14
         Printer.CurrentX = PLeft(1) - Printer.TextWidth("獎金")
         Printer.CurrentY = iLine * 300
         Printer.Print "端節"
         iLine = iLine + 1
         Printer.CurrentX = PLeft(1) - Printer.TextWidth("獎金")
         Printer.CurrentY = iLine * 300
         Printer.Print "獎金"
      Case 15
         Printer.CurrentX = PLeft(1) - Printer.TextWidth("獎金")
         Printer.CurrentY = iLine * 300
         Printer.Print "秋節"
         iLine = iLine + 1
         Printer.CurrentX = PLeft(1) - Printer.TextWidth("獎金")
         Printer.CurrentY = iLine * 300
         Printer.Print "獎金"
      Case 16
         Printer.CurrentX = PLeft(1) - Printer.TextWidth("獎金")
         Printer.CurrentY = iLine * 300
         Printer.Print "春節"
         iLine = iLine + 1
         Printer.CurrentX = PLeft(1) - Printer.TextWidth("獎金")
         Printer.CurrentY = iLine * 300
         Printer.Print "獎金"
      Case Else
         If strTemp(3) = 1 Then iLine = iLine + 1
         Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(strTemp(3), "##"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(strTemp(3), "##")
   End Select
   'end 2016/7/13
   
   Printer.CurrentX = PLeft(2) - Printer.TextWidth(Format(strTemp(4), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(4), "##,###,###")
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(5), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(5), "##,###,###")
   Printer.CurrentX = PLeft(4) - Printer.TextWidth(Format(dblAmtS, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblAmtS, "##,###,###")
   Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(strTemp(6), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(6), "##,###,###")
   Printer.CurrentX = PLeft(6) - Printer.TextWidth(Format(strTemp(7), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(7), "##,###,###")
   Printer.CurrentX = PLeft(7) - Printer.TextWidth(Format(strTemp(8), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(8), "##,###,###")
   'Add By Sindy 2024/8/8 +勞健保費
   Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(strTemp(15), "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(15), "##,###,###")
   '2024/8/8 END
   Printer.CurrentX = PLeft(9) - Printer.TextWidth(Format(dblTAmtT, "##,###,###"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(dblTAmtT, "##,###,###")
   
   iLine = iLine + 2
End Sub

Sub PrintEmpty(ByRef start_month As Integer, ByRef end_month As Integer)
   
   For m_i = start_month To end_month
      
      If m_i = 1 Then iLine = iLine + 1
      
      Printer.CurrentX = PLeft(1) - Printer.TextWidth(Format(m_i, "##"))
      Printer.CurrentY = iLine * 300
      Printer.Print Format(m_i, "##")
      
      iLine = iLine + 2
      m_month = m_month + 1

   Next m_i

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
   Set frm170226 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 1, 2
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 3, 4
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1, 2
         'modify by sonia 2016/7/13 辜說要開放可跑年終13月及三節14,15,16月的資料
         'If txt1(Index) <> "" Then
         '   If ChkDate(txt1(Index) & "01") = False Then
         '      Cancel = True
         '   End If
         'End If
         If Right(txt1(Index), 2) <> "" Then
            If Right(txt1(Index), 2) > 16 Then
               MsgBox "月份輸入錯誤, 不可超過 16！", vbInformation, "操作錯誤！"
               Cancel = True
            End If
         End If
         'end 2016/7/13
         If Index = 1 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 2 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
            End If
         End If
      Case 3, 4
         ' 判斷員工代號須為 6~9 或 F 開頭
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Cancel = True
            End If
         End If
         If Index = 3 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 4 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
            End If
         End If
      Case Else
   End Select
   
   If Cancel = True Then TextInverse txt1(Index)
      
End Sub
