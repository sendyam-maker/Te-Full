VERSION 5.00
Begin VB.Form frm12040144 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外客戶/代理人地址條列印"
   ClientHeight    =   4860
   ClientLeft      =   2952
   ClientTop       =   1620
   ClientWidth     =   5304
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5304
   Begin VB.CheckBox Check1 
      Caption         =   "特定對象"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   3696
      TabIndex        =   24
      Top             =   1320
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2016
      MaxLength       =   1
      TabIndex        =   10
      Top             =   4104
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2016
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "N"
      Top             =   3744
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1056
      MaxLength       =   100
      TabIndex        =   4
      Top             =   2064
      Width           =   1950
   End
   Begin VB.CheckBox Check1 
      Caption         =   "代理人"
      Height          =   255
      Index           =   1
      Left            =   2232
      TabIndex        =   1
      Top             =   1320
      Width           =   1035
   End
   Begin VB.CheckBox Check1 
      Caption         =   "客戶"
      Height          =   255
      Index           =   0
      Left            =   1092
      TabIndex        =   0
      Top             =   1320
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   12
      Left            =   2016
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "N"
      Top             =   3432
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   11
      Left            =   1056
      MaxLength       =   3
      TabIndex        =   5
      Top             =   2424
      Width           =   465
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   276
      Left            =   1080
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   4500
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3996
      TabIndex        =   13
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3168
      TabIndex        =   12
      Top             =   45
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   6
      Left            =   2616
      MaxLength       =   7
      TabIndex        =   7
      Top             =   3060
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   1056
      MaxLength       =   7
      TabIndex        =   6
      Top             =   3060
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   2616
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1056
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "3.先用PDF reDirect先跑過一次全部清單取得流水號"
      ForeColor       =   &H000000FF&
      Height          =   204
      Left            =   240
      TabIndex        =   26
      Top             =   1056
      Width           =   4140
   End
   Begin VB.Label Label1 
      Caption         =   "特定對象請先用語法抓出資料寫入FAGENT_SPECIAL 1.必須為有地址的資料；預設只抓最新日期的記錄      2.畫面條件只過濾國籍和除外國籍條件"
      ForeColor       =   &H00FF0000&
      Height          =   540
      Index           =   4
      Left            =   240
      TabIndex        =   25
      Top             =   480
      Width           =   4140
   End
   Begin VB.Label Label5 
      Caption         =   "是否只寄電子報對象:             (Y: 只寄電子報對象)"
      Height          =   252
      Left            =   216
      TabIndex        =   23
      Top             =   4116
      Width           =   4008
   End
   Begin VB.Label Label4 
      Caption         =   "是否含寄電子報對象:             (N: 不含)"
      Height          =   252
      Left            =   216
      TabIndex        =   22
      Top             =   3756
      Width           =   3012
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "除外國籍:                                               (可複選，請以 , 區隔)"
      Height          =   180
      Index           =   1
      Left            =   192
      TabIndex        =   21
      Top             =   2100
      Width           =   4572
   End
   Begin VB.Label Label3 
      Caption         =   "是否含不寄雜誌對象:             (N: 不含)"
      Height          =   252
      Left            =   216
      TabIndex        =   20
      Top             =   3444
      Width           =   3012
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(A:代理人律師事務所 B:公司直接委辦 C:其他)"
      Height          =   180
      Index           =   7
      Left            =   216
      TabIndex        =   19
      Top             =   2772
      Width           =   3588
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "性質:                     (可複選，請以 , 區隔) 對象為代理人才適用"
      Height          =   180
      Index           =   20
      Left            =   216
      TabIndex        =   18
      Top             =   2460
      Width           =   4716
   End
   Begin VB.Label Label2 
      Caption         =   "印表機:"
      Height          =   252
      Left            =   240
      TabIndex        =   17
      Top             =   4500
      Width           =   612
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2376
      X2              =   2496
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2376
      X2              =   2496
      Y1              =   1812
      Y2              =   1812
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "開發日期:"
      Height          =   180
      Index           =   3
      Left            =   216
      TabIndex        =   16
      Top             =   3060
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "國籍:"
      Height          =   180
      Index           =   2
      Left            =   216
      TabIndex        =   15
      Top             =   1692
      Width           =   408
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "列印對象:"
      Height          =   180
      Index           =   0
      Left            =   216
      TabIndex        =   14
      Top             =   1332
      Width           =   768
   End
End
Attribute VB_Name = "frm12040144"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim intWhere As Integer, strReceiveNo As String, PLeft(0 To 7) As Integer
Dim m_DefaultPrinter As String '' 預設印表機
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
Dim m_dbl_LeftMargin As Double
Dim m_dbl_TopMargin As Double
Dim m_PageNo As Double

Private Sub cmdok_Click(Index As Integer)
Dim rsTemp1 As New ADODB.Recordset
Dim nPageNo As Integer
Dim Prn As Printer
Dim i As Integer
Dim varTmp As Variant
Dim strCU As String
Dim strPCU As String 'Added by Lydia 2018/11/01
Dim intCounter As Integer, intStart As Integer 'Added by Lydia 2023/11/2

    Select Case Index
    Case 0 '確定
        Screen.MousePointer = vbHourglass
        '檢查輸入的資料是否齊全完整
        If CheckDataValid() = False Then
            GoTo EXITSUB
        End If
        '設定使用者所選擇的印表機成預設印表機
        For Each Prn In Printers
            If Prn.DeviceName = cmbPrinter.Text Then
                Set Printer = Prn
                Exit For
            End If
        Next
        
         'add by sonia 2015/9/10 特定對象請先用語法抓出資料寫入FAGENT_SPECIAL(僅存編號8碼,必須有地址的資料)
         '此處只考慮國籍和除外國籍條件
         If Check1(2).Value = vbChecked Then
            strSql = ""
            strCU = "" 'Added by Morgan 2017/11/14
            strPCU = "" 'Added by Lydia 2018/11/01
            '國籍區間
            If Text1(3).Text <> "" And Text1(4).Text <> "" Then
                strSql = strSql & " AND (SUBSTR(FA10,1,3) >= '" & Text1(3).Text & "' AND SUBSTR(FA10,1,3) <= '" & Text1(4).Text & "') "
                strCU = strCU & " AND (SUBSTR(CU10,1,3) >= '" & Text1(3).Text & "' AND SUBSTR(CU10,1,3) <= '" & Text1(4).Text & "') " 'Added by Morgan 2017/11/14
                strPCU = strPCU & " AND (SUBSTR(PCU09,1,3) >= '" & Text1(3).Text & "' AND SUBSTR(PCU09,1,3) <= '" & Text1(4).Text & "') " 'Added by Lydia 2018/11/01
            ElseIf Text1(3).Text = "" And Text1(4).Text <> "" Then
                strSql = strSql & " AND SUBSTR(FA10,1,3) <='" & Text1(4).Text & "' "
                strCU = strCU & " AND SUBSTR(CU10,1,3) <= '" & Text1(4).Text & "' " 'Added by Morgan 2017/11/14
                strPCU = strPCU & "' AND SUBSTR(PCU09,1,3) <= '" & Text1(4).Text & "' "  'Added by Lydia 2018/11/01
            ElseIf Text1(3).Text <> "" And Text1(4).Text = "" Then
                strSql = strSql & " AND SUBSTR(FA10,1,3) >='" & Text1(3).Text & "' "
                strCU = strCU & " AND SUBSTR(CU10,1,3) >='" & Text1(3).Text & "' " 'Added by Morgan 2017/11/14
                strPCU = strPCU & " AND SUBSTR(PCU09,1,3) >= '" & Text1(3).Text & "'  " 'Added by Lydia 2018/11/01
            End If
            If Text1(0).Text <> "" Then
               varTmp = Split(Text1(0).Text, ",")
               strExc(1) = ""
               For i = 0 To UBound(varTmp)
                   strExc(1) = strExc(1) & "'" & varTmp(i) & "',"
               Next
               If Right(strExc(1), 1) = "," Then strExc(1) = Left(strExc(1), Len(strExc(1)) - 1)
               strSql = strSql & " AND SUBSTR(FA10,1,3) NOT IN (" & strExc(1) & ")"
               strCU = strCU & " AND SUBSTR(CU10,1,3) NOT IN (" & strExc(1) & ")" 'Added by Morgan 2017/11/14
               strPCU = strPCU & " AND SUBSTR(PCU09,1,3) NOT IN (" & strExc(1) & ")" 'Added by Lydia 2018/11/01
            End If
            'Added by Lydia 2018/10/25 預設抓最新日期
            intI = 1
            strExc(0) = "select max(fs01) mdate from fagent_special "
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                If Val("" & RsTemp.Fields("mdate")) > 0 Then
                    'Added by Lydia 2019/09/27 可以指定日期
                    'strSql = strSql & " and fs01=" & RsTemp.Fields("mdate")
                    'strCU = strCU & " and fs01=" & RsTemp.Fields("mdate")
                    'strPCU = strPCU & " and fs01=" & RsTemp.Fields("mdate") 'Added by Lydia 2018/11/01
                    strExc(2) = UCase(InputBox("若要改變日期，請在下方修改日期：", "日期條件", RsTemp.Fields("mdate")))
                    If strExc(2) = "" Then
                        MsgBox "請輸入日期!", vbExclamation
                        GoTo EXITSUB
                    End If
                    If ChkDate(strExc(2)) = True Then
                        strSql = strSql & " and fs01=" & strExc(2)
                        strCU = strCU & " and fs01=" & strExc(2)
                        strPCU = strPCU & " and fs01=" & strExc(2)
                        'Memo by Lydia 2019/09/27 因為有另外補充的Y編號,並且分區列印
                        'strSql = strSql & " and fs01>=" & strExc(2) & " and fs04='1' "
                        'strCU = strCU & " and fs01>=" & strExc(2) & " and fs04='1' "
                        'strPCU = strPCU & " and fs01>=" & strExc(2) & " and fs04='1' "
                        'Memo by Lydia 2022/10/12 手動分區列印
                        'strSql = strSql & " and fs04='3' "
                        'strCU = strCU & " and fs04='3' "
                        'strPCU = strPCU & " and fs04='3' "
                        'Added by Lydia 2023/11/02 取得列印洲別和最大流水號
                        strExc(5) = UCase(InputBox("若要指定洲別，請在下方輸大1~3：" & vbCrLf & "因總務郵務需求，依往例以地區分1-港澳大、2-亞洲大洋洲、3-歐洲美洲非洲，或者是0=全部。", "列印洲別條件", ""))
                        If strExc(5) <> "0" And strExc(5) <> "1" And strExc(5) <> "2" And strExc(5) <> "3" Then
                           MsgBox "請輸入指定洲別0~3!", vbExclamation
                           GoTo EXITSUB
                        Else
                           strExc(0) = "select decode(max(fs07),null,1,max(fs07)+1) as mno from fagent_special where fs01=" & strExc(2)
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              intCounter = Val("" & RsTemp.Fields("mno"))
                           End If
                        End If
                        If strExc(5) <> "0" Then
                           strSql = strSql & " and fs04=" & CNULL(strExc(5)) & " "
                           strCU = strCU & " and fs04=" & CNULL(strExc(5)) & " "
                           strPCU = strPCU & " and fs04=" & CNULL(strExc(5)) & " "
                        End If
                        'intStart = 598  'Memo by Lydia 2023/11/02 在印2024竹曆遇見卡紙現象
                        If intStart > 0 Then '指定列印流水號
                           strSql = strSql & " and fs07>=" & intStart & " "
                           strCU = strCU & " and fs07>=" & intStart & " "
                           strPCU = strPCU & " and fs07>=" & intStart & " "
                        End If
                        'end 2023/11/02
                    Else
                        GoTo EXITSUB
                    End If
                    'end 2019/09/27
                End If
            End If
            
            'end 2018/10/25
            
            'Modified by Morgan 2017/11/14 改呼叫地址條列印畫面
            'strExc(0) = "SELECT FA08,Decode(FA05,Null,Decode(FA04,Null,substr(FA06,1,20),substr(FA04,1,20)),FA05),Decode(FA05,Null,Decode(FA04,Null,substr(FA06,21,20),substr(FA04,21,20)),FA63),Decode(FA05,Null,'',FA64),Decode(FA05,Null,'',FA65)," & _
                        " Decode(FA18,Null,Nvl(FA17,FA23),FA18), Decode(FA18,Null,'',FA19), Decode(FA18,Null,'',FA20), Decode(FA18,Null,'',FA21), Decode(FA18,Null,'',FA22)," & _
                        " Decode(FA18,Null,'',FA70), FA01 As CodeNo, substr(FA10,1,3) As Nation, FA05||' '||FA63||' '||FA64||' '||FA65 As EngName, FA04 As ChnName, FA06 As JpnName,FS02,FS03 FROM FAGENT_SPECIAL,FAGENT Where FS01='00' and NO=FA01(+) AND FA02='0' " & strSql
            'strExc(0) = strExc(0) & " Order By 13, 14, 15, 16 "
            'Modified by Lydia 2018/10/25 FAGENT_SPECIAL傳入代號為9碼
            'strExc(0) = "select NO,FS01,FS02,FS03,substr(FA10,1,3) As FA10 from FAGENT_SPECIAL,FAGENT where substr(NO,1,1)='Y' and NO=FA01(+) AND FA02='0' " & strSql
            'strExc(0) = strExc(0) & " union all select NO,FS01,FS02,FS03,substr(CU10,1,3) As FA10 from FAGENT_SPECIAL,customer where substr(NO,1,1)='X' and NO=CU01(+) AND CU02='0' " & strCU
            'strExc(0) = strExc(0) & " Order By FA10,NO,FS01"
            'end 2017/11/14
            'Modified by Lydia 2022/10/13 +nvl(fs06,1) as fs06
            'Modified by Lydia 2023/11/2 +fsno, fs01,nvl(fs07,0) as fs07
            strExc(0) = "select fsno as no,fs02,fs03,fs04,substr(fa10,1,3) as fa10,nvl(fs06,1) as fs06,fsno,fs01,nvl(fs07,0) as fs07 from fagent_special,fagent where substr(fsno,1,1)='Y' and substr(fsno,1,8)=fa01(+) and fa02='0' " & strSql
            strExc(0) = strExc(0) & " union all select fsno as no,fs02,fs03,fs04,substr(cu10,1,3) as fa10,nvl(fs06,1) as fs06,fsno,fs01,nvl(fs07,0) as fs07 from fagent_special,customer where substr(fsno,1,1)='X' and substr(fsno,1,8)=cu01(+) and cu02='0' " & strCU
            'Added by Lydia 2018/11/01 增加抓國外潛在客戶(聯絡人)
            'Modified by Lydia 2023/11/2 +fsno, fs01,nvl(fs07,0) as fs07
            strExc(0) = strExc(0) & " union all select pcu01||pcu02 as no,decode(pcu36,'1',pcc05,pcc03) fs02,fs03,fs04,substr(pcu09,1,3) as fa10,nvl(fs06,1) as fs06,fsno,fs01,nvl(fs07,0) as fs07 from fagent_special,potcustomer,potcustcont where substr(fsno,1,1)='R' and substr(fsno,1,8)=pcu01(+) and pcu02='0' " & strPCU & _
                             " and pcu01=pcc01(+) and fsno=pcc01||'-'||pcc02 "
            'Added by Lydia 2019/09/27 增加抓國外潛在客戶(9碼，沒有聯絡人)
            'Modified by Lydia 2020/11/17 +and length(fsno)=9
            'Modified by Lydia 2022/10/13 +and pcu01 is not null
            'Modified by Lydia 2023/11/2 +fsno, fs01,nvl(fs07,0) as fs07
            strExc(0) = strExc(0) & " union all select pcu01||pcu02 as no,fs02,fs03,fs04,substr(pcu09,1,3) as fa10,nvl(fs06,1) as fs06,fsno ,fs01 ,nvl(fs07,0) as fs07 from fagent_special,potcustomer where substr(fsno,1,1)='R' and substr(fsno,1,8)=pcu01(+) and substr(fsno,9,1)=pcu02(+) and length(fsno)=9  and pcu01 is not null " & strPCU
            'Memo by Lydia 2022/10/13 國內潛在客戶請改用A4地址條列印
            'strExc(0) = strExc(0) & " union all select 'A'||poc01||poc02 as no,nvl(pcc05,pcc03) fs02,fs03,fs04,'000' as fa10,nvl(fs06,1) as fs06 from fagent_special,potcustomer1,potcustcont where substr(fsno,1,1)='R' and substr(fsno,1,8)=poc01(+) and poc02='0' " & strPCU & " and poc01=pcc01(+) and fsno=pcc01||'-'||pcc02"
            'strExc(0) = strExc(0) & " union all select 'A'||poc01||poc02 as no,fs02,fs03,fs04,'000' as fa10,nvl(fs06,1) as fs06 from fagent_special,potcustomer1 where substr(fsno,1,1)='R' and substr(fsno,1,8)=poc01(+) and substr(fsno,9,1)=poc02(+) and length(fsno)=9 and poc01 is not null " & strPCU
            'end 2022/10/13
            'Modified by Lydia 2023/11/02 若已有列印流水號,改成以流水號優先
            'strExc(0) = strExc(0) & " order by fs04 asc, fa10 asc, 1 asc "
            strExc(0) = strExc(0) & " order by fs07 asc, fs04 asc, fa10 asc, no asc "
            'end 2018/10/25
            intI = 1
            Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Modified by Morgan 2017/11/14 改呼叫地址條列印畫面
                'PrintCaseBatch rsTemp1
               pub_blnBatchPrintAddress = False
               '載入表單
               Load frm083014
               With frm083014
               .Show
               .Combo1.Text = cmbPrinter.Text '設定印表機
               'Modified by Lydia 2022/10/13 改成變數
               '.Text1(3).Text = "1" '列印份數
               .Text1(3).Text = Val("" & rsTemp1.Fields("fs06")) '列印份數
               .Text1(5).Text = "Y" '是否含不寄雜誌的對象
               
               rsTemp1.MoveFirst
               Do While Not rsTemp1.EOF
                  'Added by Lydia 2025/10/08
                  .p_FScon1 = ""
                  .p_FScon2 = ""
                  .p_FSconDept = ""
                  '最後還是補上英文地址，先保留程式；指定客戶定稿為日文---避免影響過大，在檢查資料後才設定
                  'If InStr("R1927300,R1927500,R1927600,R1928300,R1928500,Y5614600", Mid("" & rsTemp1.Fields("NO"), 1, 8)) > 0 Then
                  '   .p_SpecLan = "3"
                  'Else
                  '   .p_SpecLan = ""
                  'End If
                  '-------------------
                  '避免Form重抓
                  If InStr("X,Y,R", Left(rsTemp1("NO"), 1)) > 0 And Not IsNull(rsTemp1("FS02")) Then
                     .p_FScon1 = "" & rsTemp1("FS02")
                     .p_FScon2 = ""
                     .p_FSconDept = "" & rsTemp1("FS03")
                  End If
                  'end 2025/10/08
                  .m_InputNo = "" & rsTemp1("NO")
                  If Left(rsTemp1("NO"), 1) = "Y" Then
                     .opt1(1).Value = True
                     .Text1(1) = rsTemp1("NO")
                     .Text1_LostFocus 1
                  'Added by Lydia 2018/11/01 潛在客戶
                  ElseIf Left(rsTemp1("NO"), 1) = "R" Then
                     .opt1(3).Value = True
                     .Text1(16) = rsTemp1("NO")
                     .Text1_LostFocus 16
                  'end 2018/11/01
                  Else
                     .opt1(0).Value = True
                     .Text1(0) = "" & rsTemp1("NO")
                     .Text1_LostFocus 0
                  End If
                  'Modified by Lydia 2022/10/13
                  'If Not IsNull(rsTemp1("FS02")) Then
                  If InStr("X,Y,R", Left(rsTemp1("NO"), 1)) > 0 And Not IsNull(rsTemp1("FS02")) Then
                     'Modified by Lydia 2022/05/02
                     '.Text1(13) = rsTemp1("FS02")
                     '.Text1(14) = ""
                     '.Text1(15) = "" & rsTemp1("FS03")
                     .textFM2(0) = "" & rsTemp1("FS02")
                     .textFM2(1) = ""
                     .textFM2(2) = "" & rsTemp1("FS03")
                     'end 2022/05/01
                  End If
                  .Text1(3).Text = Val("" & rsTemp1.Fields("fs06")) 'Added by Lydia 2022/10/13 列印份數
                  .cmdPrint_Click '執行列印'
                  'Memo by Lydia 2018/10/25 測試用
                  'If rsTemp1.AbsolutePosition = 5 Then Exit Do
                  'Debug.Print rsTemp1.AbsolutePosition & ":" & rsTemp1("NO")
                  'Added by Lydia 2023/11/02 更新列印流水號
                  If Val("" & rsTemp1.Fields("fs07")) = 0 Then
                     'Modified by Lydia 2025/11/13 +區分判斷聯絡人fs02
                     strExc(5) = "Update fagent_special set fs07=" & intCounter & " where fs01='" & rsTemp1.Fields("fs01") & "' and fsno='" & rsTemp1.Fields("fsno") & "' " & _
                                 IIf("" & rsTemp1.Fields("fs02") <> "", " and fs02='" & rsTemp1.Fields("fs02") & "' ", "")
                     intCounter = intCounter + 1
                     cnnConnection.Execute strExc(5)
                  End If
                  'end 2023/11/02
                  rsTemp1.MoveNext
               Loop
               Printer.EndDoc
               End With
               Unload frm083014
               
                'end 2017/11/14
                MsgBox "列印結束，國籍：" & Me.Text1(3).Text & "－" & Me.Text1(4).Text & " 共 " & Format(rsTemp1.RecordCount, "#,##0") & " 筆 !!!", vbInformation
            Else
               MsgBox "無符合條件之資料可列印 !", vbInformation
            End If
            
            GoTo EXITSUB
         End If
         'end 2015/9/10
          
        strExc(0) = ""
        strExc(1) = ""
        '客戶
        If Me.Check1(0).Value = vbChecked Then
            strSql = ""
            '國籍區間
            If Text1(3).Text <> "" And Text1(4).Text <> "" Then
                strSql = strSql & " AND (SUBSTR(CU10,1,3) >= '" & Text1(3).Text & "' AND SUBSTR(CU10,1,3) <= '" & Text1(4).Text & "') "
            ElseIf Text1(3).Text = "" And Text1(4).Text <> "" Then
                strSql = strSql & " AND SUBSTR(CU10,1,3) <='" & Text1(4).Text & "' "
            ElseIf Text1(3).Text <> "" And Text1(4).Text = "" Then
                strSql = strSql & " AND SUBSTR(CU10,1,3) >='" & Text1(3).Text & "' "
            End If
            '92.8.4 ADD BY SONIA
            If Text1(0).Text <> "" Then
               varTmp = Split(Text1(0).Text, ",")
               strExc(0) = ""
               For i = 0 To UBound(varTmp)
                   strExc(0) = strExc(0) & "'" & varTmp(i) & "',"
               Next
               If Right(strExc(0), 1) = "," Then strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
               strSql = strSql & " AND SUBSTR(CU10,1,3) NOT IN (" & strExc(0) & ")"
            End If
            '92.8.4 END
'edit by nickc 2005/12/02 改成由代理人
'            '性質
'            If Text1(11).Text <> "" Then
'                varTmp = Split(Text1(11).Text, ",")
'                strExc(0) = ""
'                For i = 0 To UBound(varTmp)
'                    strExc(0) = strExc(0) & "'" & Format(varTmp(i)) & "',"
'                Next
'                If Right(strExc(0), 1) = "," Then strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
'                strSQL = strSQL & " AND CU101 IN (" & strExc(0) & ")"
'            End If
'2009/6/12因為尚未整理資料故先還原 BY SONIA
            If Text1(11).Text <> "" Then
                varTmp = Split(Text1(11).Text, ",")
                strExc(0) = ""
                For i = 0 To UBound(varTmp)
                    strExc(0) = strExc(0) & "'" & Format(varTmp(i)) & "',"
                Next
                If Right(strExc(0), 1) = "," Then strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
                strSql = strSql & " AND CU101 IN (" & strExc(0) & ")"
            End If
'2009/6/12 END
            '開發日期區間
            If Text1(5).Text <> "" And Text1(6).Text <> "" Then
                strSql = strSql & " AND CU14 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2)
            ElseIf Text1(5).Text = "" And Text1(6).Text <> "" Then
                strSql = strSql & " AND CU14 <=" & TransDate(Text1(6).Text, 2)
            ElseIf Text1(5).Text <> "" And Text1(6).Text = "" Then
                strSql = strSql & " AND CU14 >=" & TransDate(Text1(5).Text, 2) & _
                                  " AND CU14 <=" & ServerDate & " "
            End If
            '2010/10/25 add by sonia 是否只寄電子報對象
            If Text1(2) = "Y" Then
               strSql = strSql & " AND CU32 IS NULL AND CU20||CU116||CU117||CU118 IS NOT NULL AND UPPER(CU20||CU116||CU117||CU118)<>'NO' AND CU132 IS NULL AND (CU79 IS NULL OR INSTR(CU79,'90 fail')=0) "
            Else
            '2010/10/25 end
               '是否含不寄雜誌對象
               If Text1(12) = "N" Then
                   strSql = strSql & " AND CU32 IS NULL "
               End If
               '2009/3/4 add by sonia 是否含寄電子報對象
               If Text1(1) = "N" Then
                   strSql = strSql & " AND (CU20||CU116||CU117||CU118 IS NULL OR UPPER(CU20||CU116||CU117||CU118)='NO' OR CU132='N' OR INSTR(CU79,'90 fail')>0) "
               End If
            End If    '2010/10/25 add by sonia
            '2009/3/4 END
            
            '要有地址資料
            strSql = strSql & " And (CU24 Is Not Null Or CU23 Is Not Null Or CU29 Is Not Null ) "
'            strExc(0) = "SELECT CU01 As CodeNo, substr(CU10,1,3) As Nation, CU05||' '||CU88||' '||CU89||' '||CU90 As EngName, CU04 As ChnName, CU06 As JpnName FROM CUSTOMER Where CU01=CU01 " & strSQL & " AND CU02='0' AND CU03 IS NULL "
            'edit by nick 2004/07/02 先檢查 cu104 ，再檢查 cu04
            'strExc(0) = "SELECT CU59,Decode(CU05,Null,Decode(CU04,Null,substr(CU06,1,20),substr(CU04,1,20)),CU05),Decode(CU05,Null,Decode(CU04,Null,substr(CU06,21,20),substr(CU04,21,20)),CU88),Decode(CU05,Null,'',CU89),Decode(CU05,Null,'',CU90),DECODE(CU24,Null,Nvl(CU23,CU29),CU24)," & _
                            "DECODE(CU24,Null,'',CU25),DECODE(CU24,Null,'',CU26)," & _
                            "DECODE(CU24,Null,'',CU27),DECODE(CU24,Null,'',CU28)," & _
                            "DECODE(CU24,Null,'',CU102,''), CU01 As CodeNo, substr(CU10,1,3) As Nation, CU05||' '||CU88||' '||CU89||' '||CU90 As EngName, CU04 As ChnName, CU06 As JpnName FROM CUSTOMER Where CU01=CU01 " & strSQL & " AND CU02='0' AND CU03 IS NULL "
            strExc(0) = "SELECT CU59,Decode(CU05,Null,Decode(decode(cu104,null,CU04,cu104),Null,substr(CU06,1,20),substr(decode(cu104,null,CU04,cu104),1,20)),CU05),Decode(CU05,Null,Decode(decode(cu104,null,CU04,cu104),Null,substr(CU06,21,20),substr(decode(cu104,null,CU04,cu104),21,20)),CU88),Decode(CU05,Null,'',CU89),Decode(CU05,Null,'',CU90),DECODE(CU24,Null,Nvl(CU23,CU29),CU24)," & _
                            "DECODE(CU24,Null,'',CU25),DECODE(CU24,Null,'',CU26)," & _
                            "DECODE(CU24,Null,'',CU27),DECODE(CU24,Null,'',CU28)," & _
                            "DECODE(CU24,Null,'',CU102,''), CU01 As CodeNo, substr(CU10,1,3) As Nation, CU05||' '||CU88||' '||CU89||' '||CU90 As EngName, decode(cu104,null,CU04,cu104) As ChnName, CU06 As JpnName FROM CUSTOMER Where CU01=CU01 " & strSql & " AND CU02='0' AND CU03 IS NULL "
        End If
        '代理人
        If Me.Check1(1).Value = vbChecked Then
            strSql = ""
            '國籍區間
            If Text1(3).Text <> "" And Text1(4).Text <> "" Then
                strSql = strSql & " AND (SUBSTR(FA10,1,3) >= '" & Text1(3).Text & "' AND SUBSTR(FA10,1,3) <= '" & Text1(4).Text & "') "
            ElseIf Text1(3).Text = "" And Text1(4).Text <> "" Then
                strSql = strSql & " AND SUBSTR(FA10,1,3) <='" & Text1(4).Text & "' "
            ElseIf Text1(3).Text <> "" And Text1(4).Text = "" Then
                strSql = strSql & " AND SUBSTR(FA10,1,3) >='" & Text1(3).Text & "' "
            End If
            '92.8.4 ADD BY SONIA
            If Text1(0).Text <> "" Then
               varTmp = Split(Text1(0).Text, ",")
               strExc(1) = ""
               For i = 0 To UBound(varTmp)
                   strExc(1) = strExc(1) & "'" & varTmp(i) & "',"
               Next
               If Right(strExc(1), 1) = "," Then strExc(1) = Left(strExc(1), Len(strExc(1)) - 1)
               strSql = strSql & " AND SUBSTR(FA10,1,3) NOT IN (" & strExc(1) & ")"
            End If
            '92.8.4 END
            '開發日期區間
            If Text1(5).Text <> "" And Text1(6).Text <> "" Then
                strSql = strSql & " AND FA11 BETWEEN " & TransDate(Text1(5).Text, 2) & " AND " & TransDate(Text1(6).Text, 2)
            ElseIf Text1(5).Text = "" And Text1(6).Text <> "" Then
                strSql = strSql & " AND FA11 <=" & TransDate(Text1(6).Text, 2)
            ElseIf Text1(5).Text <> "" And Text1(6).Text = "" Then
                strSql = strSql & " AND FA11 >=" & TransDate(Text1(5).Text, 2) & _
                                " AND FA11 <=" & ServerDate & " "
            End If
'2009/6/12因為尚未整理資料故先還原 BY SONIA
'            'add by nickc 2005/12/02 性質
'            If Text1(11).Text <> "" Then
'                varTmp = Split(Text1(11).Text, ",")
'                strExc(1) = ""
'                For i = 0 To UBound(varTmp)
'                    strExc(1) = strExc(1) & "'" & Format(varTmp(i)) & "',"
'                Next
'                If Right(strExc(1), 1) = "," Then strExc(1) = Left(strExc(1), Len(strExc(1)) - 1)
'                strSQL = strSQL & " AND FA76 IN (" & strExc(1) & ")"
'            End If
'2009/6/12 END
            
            '2010/10/25 add by sonia 是否只寄電子報對象
            If Text1(2) = "Y" Then
               strSql = strSql & " AND FA24 IS NULL AND FA16||FA80||FA81||FA82 IS NOT NULL AND UPPER(FA16||FA80||FA81||FA82)<>'NO' AND FA97 IS NULL AND (FA29 IS NULL OR INSTR(FA29,'90 fail')=0) "
            Else
            '2010/10/25 end
               '是否含不寄雜誌對象
               If Text1(12) = "N" Then
                   strSql = strSql & " AND FA24 IS NULL "
               End If
               '2009/3/4 add by sonia 是否含寄電子報對象
               If Text1(1) = "N" Then
                   strSql = strSql & " AND (FA16||FA80||FA81||FA82 IS NULL OR UPPER(FA16||FA80||FA81||FA82)='NO' OR FA97='N' OR INSTR(FA29,'90 fail')>0) "
               End If
            End If    '2010/10/25 add by sonia
            '2009/3/4 END
            
            '要有地址資料
            strSql = strSql & " And (FA18 Is Not Null Or FA17 Is Not Null Or FA23 Is Not Null ) "
'            strExc(1) = "SELECT FA01 As CodeNo, substr(FA10,1,3) As Nation, FA05||' '||FA63||' '||FA64||' '||FA65 As EngName, FA04 As ChnName, FA06 As JpnName FROM FAGENT Where FA01=FA01 " & strSQL & " AND FA02='0' "
            '93.9.1 MODIFY BY SONIA
            'strExc(1) = "SELECT FA08,Decode(FA05,Null,Decode(FA04,Null,substr(FA06,1,20),substr(FA04,1,20)),FA05),Decode(FA05,Null,Decode(FA04,Null,substr(FA06,21,20),substr(FA04,21,20)),FA63),Decode(FA05,Null,'',FA64),Decode(FA05,Null,'',FA65)," & _
            '                " Decode(FA18,Null,Nvl(FA17,FA23),FA18), Decode(FA18,Null,'',FA19), Decode(FA18,Null,'',FA20), Decode(FA18,Null,'',FA21), Decode(FA18,Null,'',FA22)," & _
            '                " Decode(FA18,Null,'',FA70), FA01 As CodeNo, substr(FA10,1,3) As Nation, FA05||' '||FA63||' '||FA64||' '||FA65 As EngName, FA04 As ChnName, FA06 As JpnName FROM FAGENT Where FA01=FA01 " & strSQL & " AND FA02='0' "
            '2008/5/19 modify by sonia Fagent不必考慮CU101並取消FA01=FA01
            'strExc(1) = "SELECT FA08,Decode(FA05,Null,Decode(FA04,Null,substr(FA06,1,20),substr(FA04,1,20)),FA05),Decode(FA05,Null,Decode(FA04,Null,substr(FA06,21,20),substr(FA04,21,20)),FA63),Decode(FA05,Null,'',FA64),Decode(FA05,Null,'',FA65)," & _
            '                " Decode(FA18,Null,Nvl(FA17,FA23),FA18), Decode(FA18,Null,'',FA19), Decode(FA18,Null,'',FA20), Decode(FA18,Null,'',FA21), Decode(FA18,Null,'',FA22)," & _
            '                " Decode(FA18,Null,'',FA70), FA01 As CodeNo, substr(FA10,1,3) As Nation, FA05||' '||FA63||' '||FA64||' '||FA65 As EngName, FA04 As ChnName, FA06 As JpnName FROM FAGENT,CUSTOMER Where FA01=FA01 " & strSQL & " AND FA02='0' AND FA03=CU01(+) AND (CU101='A' OR CU101 IS NULL) "
            strExc(1) = "SELECT FA08,Decode(FA05,Null,Decode(FA04,Null,substr(FA06,1,20),substr(FA04,1,20)),FA05),Decode(FA05,Null,Decode(FA04,Null,substr(FA06,21,20),substr(FA04,21,20)),FA63),Decode(FA05,Null,'',FA64),Decode(FA05,Null,'',FA65)," & _
                            " Decode(FA18,Null,Nvl(FA17,FA23),FA18), Decode(FA18,Null,'',FA19), Decode(FA18,Null,'',FA20), Decode(FA18,Null,'',FA21), Decode(FA18,Null,'',FA22)," & _
                            " Decode(FA18,Null,'',FA70), FA01 As CodeNo, substr(FA10,1,3) As Nation, FA05||' '||FA63||' '||FA64||' '||FA65 As EngName, FA04 As ChnName, FA06 As JpnName FROM FAGENT Where FA02='0' " & strSql
            '93.9.1 END
        End If
        If strExc(0) <> "" Then
            If strExc(1) <> "" Then
                strExc(0) = strExc(0) & " Union "
            End If
        End If
        If strExc(1) <> "" Then
            strExc(0) = strExc(0) & strExc(1)
        End If
        strExc(0) = strExc(0) & " Order By 13, 14, 15, 16 "
        intI = 1
        'edit by nickc 2007/02/09 不用 dll 了
        'Set rsTemp1 = objLawDll.ReadRstMsg(intI, strExc(0))
        Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            PrintCaseBatch rsTemp1
            MsgBox "列印結束，國籍：" & Me.Text1(3).Text & "－" & Me.Text1(4).Text & " 共 " & Format(rsTemp1.RecordCount, "#,##0") & " 筆 !!!", vbInformation
        Else
           MsgBox "無符合條件之資料可列印 !", vbInformation
        End If
        Screen.MousePointer = vbDefault
    Case 1 '結束
        Unload Me
    End Select
Exit Sub
EXITSUB:
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Dim Prn As Printer
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' 將印表機設為原先的預設印表機
   PUB_RestorePrinter m_DefaultPrinter
   '若印表機變動, 則更新列印設定
   If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   
   Set frm12040144 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 7, 8
         If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 11
         'edit by nickc 2005/12/02
         'If (KeyAscii > 68 Or KeyAscii < 65) And KeyAscii <> 8 And KeyAscii <> 44 Then
         If (KeyAscii > 66 Or KeyAscii < 65) And KeyAscii <> 8 And KeyAscii <> 44 Then
            KeyAscii = 0
            Beep
         End If
      Case 1, 12
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      '2010/10/25 ADD BY SONIA
      Case 2
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            Beep
            KeyAscii = 0
         End If
      '2010/10/25 end
   End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 2, 4, 6
         'Modify By Cheng 2002/09/10
         If blnClkSure = False Then
            If Text1(Index - 1) <> "" Then
               If RunNick(Text1(Index - 1), Text1(Index)) Then
                 Text1(Index - 1).SetFocus
               End If
            End If
         Else
            blnClkSure = False
         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTmp As String
   Select Case Index
      Case 5, 6
         If Text1(Index) <> "" Then
            Cancel = Not ChkDate(Text1(Index).Text)
         End If
      'Modify By Cheng 2002/05/13
'      Case 0, 7, 8
      Case 7, 8, 10
         If Text1(Index) = "" Then
            MsgBox "請輸入資料 !", vbCritical
            Cancel = True
         Else
            ' 90.07.12 modify by louis
            If Index = 7 Then
               RefreshPrinterList
            End If
         End If
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub

' 更新可供選擇的印表機列表
Private Sub RefreshPrinterList()
   Dim Prn As Printer
   
   cmbPrinter.Clear
   Select Case Text1(7)
      ' 地址條
      Case 1:
         For Each Prn In Printers
            If Prn.DeviceName <> m_DefaultPrinter Then
               cmbPrinter.AddItem Prn.DeviceName
            End If
         Next
         If cmbPrinter.ListCount > 0 Then
            cmbPrinter.ListIndex = 0
         End If
         cmbPrinter.Enabled = True
      ' 名冊
      Case 2:
         cmbPrinter.AddItem m_DefaultPrinter
         If cmbPrinter.ListCount > 0 Then
            cmbPrinter.ListIndex = 0
         End If
         cmbPrinter.Enabled = False
      ' 其它
      Case Else:
         For Each Prn In Printers
            cmbPrinter.AddItem Prn.DeviceName
         Next
         cmbPrinter.Enabled = True
   End Select
End Sub

Private Function GetNation(ByVal strTmp As String) As Boolean
   GetNation = False
   strExc(0) = "SELECT COUNT(*) FROM NATION WHERE SUBSTR(NA01,1,3)='" & strTmp & "'"
   intI = 1
   'edit by nickc 2007/02/09 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.Fields(0) = 1 Then GetNation = True
   End If
End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   blnClkSure = False
    '列印對象
    'Modified by Lydia 2018/10/25 +特定對象
    'If Me.Check1(0).Value = vbUnchecked And Me.Check1(1).Value = vbUnchecked Then
    If Me.Check1(0).Value = vbUnchecked And Me.Check1(1).Value = vbUnchecked And Me.Check1(2).Value = vbUnchecked Then
        MsgBox "請選擇列印對象!!!", vbExclamation + vbOKOnly
        GoTo EXITSUB
    End If
   '開發日期
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
    '國籍範圍
    If Me.Text1(3).Text > Me.Text1(4).Text Then
        MsgBox "國籍輸入範圍錯誤!!!", vbExclamation + vbOKOnly
        blnClkSure = True
        Me.Text1(3).SetFocus
        Text1_GotFocus 3
        GoTo EXITSUB
    End If
    CheckDataValid = True
EXITSUB:
End Function

'批次列印地址條
Private Sub PrintCaseBatch(rsA As ADODB.Recordset)
Dim StrSQLa As String
Dim i As Integer
Dim St As String
Dim Page As Integer
Dim iPrint As Integer
Dim IntF As Integer
Dim PriType As Integer
Dim j As Integer
Dim Prn As Printer
Dim nRow As Integer
Dim intBgnRow As Integer '起始列數
Dim strTempText  As String
Dim blnNumber  As Boolean
Dim jj  As Integer
Dim ii As Integer
Dim kk As Integer
Dim intNoPos As Integer

On Error GoTo ErrorHandler

    iPrint = 1
    '移至第一筆資料
    rsA.MoveFirst
    While Not rsA.EOF
        '設定偏移值
        m_dbl_LeftMargin = CDbl(0) * 567
        m_dbl_TopMargin = CDbl(0) * 567
        PriType = 2
        Page = 2
        intI = 0
        Select Case Page
        Case 1
            IntF = 7
        Case 2
            IntF = 11
        Case 3
            IntF = 6
        Case 4
            IntF = 5
        End Select
        
         '左邊界
         Dim iCurrentX As Integer, iHeight As Integer
         iCurrentX = 0 + m_dbl_LeftMargin
         
         'Modify by Morgan 2009/4/17 列高不夠部分字母的下半段會被截(Ex:g,y)
         'Printer.Height = 2880
         'Printer.Width = 10000
         'iHeight = 230
         '95
         If pub_OS = "1" Then
            Printer.Height = 2880
            Printer.Width = 10000
         'NT
         Else
            Printer.PaperSize = PUB_GetPaperSize(2)
         End If
         iHeight = 270
         'end 2009/4/17
         
         Printer.Font.Size = 12
         '設定列印字型
         Printer.Font.Name = "Times New Roman"
         
        Select Case Page
        Case 2
            nRow = 0
            For i = 0 To IntF - 1
                Printer.CurrentX = iCurrentX
                '若地址只有一欄有資料
                If i = 5 And "" & rsA.Fields(5).Value <> "" And "" & rsA.Fields(6).Value = "" Then
                    strTempText = ""
                    blnNumber = False
                    jj = 0
                    For ii = 1 To Len("" & rsA.Fields(5).Value)
                        jj = jj + 1
                        If Mid("" & rsA.Fields(5).Value, ii, 1) = "０" Or Mid("" & rsA.Fields(5).Value, ii, 1) = "１" Or _
                            Mid("" & rsA.Fields(5).Value, ii, 1) = "２" Or Mid("" & rsA.Fields(5).Value, ii, 1) = "３" Or _
                            Mid("" & rsA.Fields(5).Value, ii, 1) = "４" Or Mid("" & rsA.Fields(5).Value, ii, 1) = "５" Or _
                            Mid("" & rsA.Fields(5).Value, ii, 1) = "６" Or Mid("" & rsA.Fields(5).Value, ii, 1) = "７" Or _
                            Mid("" & rsA.Fields(5).Value, ii, 1) = "８" Or Mid("" & rsA.Fields(5).Value, ii, 1) = "９" Then
                            If blnNumber = False Then
                                blnNumber = True
                                kk = ii
                            End If
                            strTempText = strTempText & Mid("" & rsA.Fields(5).Value, ii, 1)
                            If jj <= 18 Then
                                GoTo NextII
                            Else
                                Printer.CurrentX = iCurrentX
                                Printer.CurrentY = nRow * iHeight + m_dbl_TopMargin
                                Printer.Print Left(strTempText, Len(strTempText) - (ii - (kk - 1)))
                                nRow = nRow + 1
                                jj = 0
                                strTempText = ""
                                ii = ii - (ii - (kk - 1))
                            End If
                        Else
                            blnNumber = False
                            strTempText = strTempText & Mid("" & rsA.Fields(5).Value, ii, 1)
                        End If
                        If jj = 18 Then
                            Printer.CurrentX = iCurrentX
                            Printer.CurrentY = nRow * iHeight + m_dbl_TopMargin
                            Printer.Print strTempText
                            nRow = nRow + 1
                            jj = 0
                            strTempText = ""
                        End If
NextII:
                    Next ii
                    If strTempText <> "" Then
                        Printer.CurrentX = iCurrentX
                        Printer.CurrentY = nRow * iHeight + m_dbl_TopMargin
                        Printer.Print strTempText
                        nRow = nRow + 1
                    End If
                '其他
                Else
                    If IsNull(rsA.Fields(i)) = False Then
                        If IsEmptyText(rsA.Fields(i)) = False Then
                            ' 語文為英文時不空行
                            If PriType = 2 Then
                                Printer.CurrentY = nRow * iHeight + m_dbl_TopMargin
                            Else
                                Printer.CurrentY = i * iHeight + m_dbl_TopMargin
                            End If
                            nRow = nRow + 1
                        End If
                    End If
                    
                    If IsNull(rsA.Fields(i)) = False Then
                        If IsEmptyText(rsA.Fields(i)) = False Then
                            Printer.Print rsA.Fields(i)
                        End If
                    End If
                End If
                
                If i = 10 Then
                    Printer.CurrentX = 3200 + iCurrentX
                    If (nRow - 1) = i Then
                        Printer.CurrentY = (nRow - 1) * iHeight + m_dbl_TopMargin
                    Else
                        Printer.CurrentY = nRow * iHeight + m_dbl_TopMargin
                    End If
                    Printer.Print Format(iPrint, "000000")
                End If
            Next i
            iPrint = iPrint + 1
        End Select
        rsA.MoveNext
        If rsA.EOF = False Then
            Printer.NewPage
        Else
            Printer.EndDoc
        End If
    Wend
Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub


