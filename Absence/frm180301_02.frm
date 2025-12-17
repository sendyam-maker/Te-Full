VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180301_02 
   BorderStyle     =   1  '單線固定
   Caption         =   "出缺勤查詢－個人統計"
   ClientHeight    =   5750
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8960
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   360
      Index           =   1
      Left            =   6660
      TabIndex        =   2
      Top             =   60
      Width           =   1365
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Index           =   2
      Left            =   8055
      TabIndex        =   0
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm180301_02.frx":0000
      Height          =   5235
      Left            =   60
      TabIndex        =   1
      Top             =   450
      Width           =   8835
      _ExtentX        =   15593
      _ExtentY        =   9243
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm180301_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2023/12/19 修改抓新部門程式
'Memo By Sindy 2021/5/28 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/8/5
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Public cmdState As Integer '紀錄作用按鍵
Dim dblPrevRow As Double


'查詢明細資料
Public Sub PubShowNextData()
   Select Case cmdState
      Case 1 '回前畫面
         Unload Me
         frm180301.Show
      Case 2 '結束
         Unload Me
         Unload frm180301
      Case Else
   End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
Dim strConABS As String, strConSA As String, strConSO As String, strConSB As String
Dim strYear As String
Dim dblTotHour As Double
   
   QueryData = False
   
   m_blnColOrderAsc = True
   GRD1.Clear
   SetGrd
   
   dblTotHour = 0
   strCon = "": strConABS = "": strConSA = "": strConSO = "": strConSB = "":
   '表單日期
   If frm180301.txtDate(0) <> "" And frm180301.txtDate(1) <> "" Then
'      If Left(frm180301.txtDate(0), 3) = Left(frm180301.txtDate(1), 3) Then
'         strYear = Left(frm180301.txtDate(0), 3)
'      Else
'         strYear = Left(strSrvDate(2), 3) '系統年
'      End If
      strYear = Left(frm180301.txtDate(1), 3)
'      strConABS = strConABS & " and (B1004 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & " or " & _
'                              "B1006 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & ") "
'      strConSA = strConSA & " and (SA02 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & " or " & _
'                              "SA04 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & ") "
'      strConSO = strConSO & " and (So02 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & ") "
'      strConSB = strConSB & " and (SB02 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & " or " & _
'                              "SB04 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & ") "
      'Modify By Sindy 2012/8/6
      strConABS = strConABS & " and (" & DBDATE(frm180301.txtDate(0)) & " between B1004 and B1006 or " & _
                              DBDATE(frm180301.txtDate(1)) & " between B1004 and B1006 or " & _
                              "B1004 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & " or " & _
                              "B1006 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & ") "
      strConSA = strConSA & " and (" & DBDATE(frm180301.txtDate(0)) & " between SA02 and SA04 or " & _
                              DBDATE(frm180301.txtDate(1)) & " between SA02 and SA04 or " & _
                              "SA02 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & " or " & _
                              "SA04 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & ") "
      strConSO = strConSO & " and (So02 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & ") "
      strConSB = strConSB & " and (" & DBDATE(frm180301.txtDate(0)) & " between SB02 and SB04 or " & _
                              DBDATE(frm180301.txtDate(1)) & " between SB02 and SB04 or " & _
                              "SB02 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & " or " & _
                              "SB04 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & ") "
   End If
   '部門別
'   If frm180301.txtDept(0) <> "" Then
'      strCon = strCon & " and a1.A0901>=" & CNULL(frm180301.txtDept(0))
'   End If
'   If frm180301.txtDept(1) <> "" Then
'      strCon = strCon & " and a1.A0901<=" & CNULL(frm180301.txtDept(1))
'   End If
   'Modify By Sindy 2021/12/21
   If frm180301.cboDept(0) <> "" Then
      'Modify By Sindy 2023/12/19
      If strSrvDate(1) >= 新部門啟用日 Then
         strCon = strCon & " and a1.A0921>=" & CNULL(Left(Trim(frm180301.cboDept(0)), 3))
      Else
      '2023/12/19 END
         strCon = strCon & " and a1.A0901>=" & CNULL(Left(Trim(frm180301.cboDept(0)), 3))
      End If
   End If
   If frm180301.cboDept(1) <> "" Then
      'Modify By Sindy 2023/12/19
      If strSrvDate(1) >= 新部門啟用日 Then
         strCon = strCon & " and a1.A0921<=" & CNULL(Left(Trim(frm180301.cboDept(1)), 3))
      Else
      '2023/12/19 END
         strCon = strCon & " and a1.A0901<=" & CNULL(Left(Trim(frm180301.cboDept(1)), 3))
      End If
   End If
   '所屬簽核的人員
   If frm180301.m_IsAbsBossST03 <> "" Then
      If frm180301.m_strEmp <> "" Then
         strCon = strCon & " and s1.ST01 in(" & frm180301.m_strEmp & ")"
      End If
   End If
   '2021/12/21 END
   
   '部門別權限限制
   If frm180301.m_IsAbsBossST03 <> "" Then
'      '部門別權限裡沒有自己所屬的部門別時,若輸入自己的員工代號做查詢時,部門別的限制要增加所屬部門
'      If frm180301.txtB1003(0) = strUserNum And frm180301.txtB1003(1) = strUserNum Then
'         strCon = strCon & " and a1.A0901 in(" & frm180301.m_IsAbsBossST03 & ",'" & Pub_StrUserSt03 & "')"
'      Else
         'Modify By Sindy 2023/12/19
         If strSrvDate(1) >= 新部門啟用日 Then
            strCon = strCon & " and a1.A0921 in(" & frm180301.m_IsAbsBossST03 & ")"
         Else
         '2023/12/19 END
            strCon = strCon & " and a1.A0901 in(" & frm180301.m_IsAbsBossST03 & ")"
         End If
'      End If
   End If
   '員工代號
   If frm180301.txtB1003(0) <> "" Then
      strCon = strCon & " and s1.ST01>=" & CNULL(frm180301.txtB1003(0))
   End If
   If frm180301.txtB1003(1) <> "" Then
      strCon = strCon & " and s1.ST01<=" & CNULL(frm180301.txtB1003(1))
   End If
   '表單類別
   If frm180301.CboB1002 <> "" Then
      strConABS = strConABS & " and B1002=" & CNULL(Left(frm180301.CboB1002, 2))
   End If
   '假別
   If frm180301.CboB1008 <> "" Then
      strConABS = strConABS & " and B1008=" & CNULL(Left(frm180301.CboB1008, 2))
      strConSA = strConSA & " and SA06=" & CNULL(Left(frm180301.CboB1008, 2))
   End If
   '所別
   If frm180301.txtST06(0) <> "" Then
      strCon = strCon & " and s1.ST06>=" & CNULL(frm180301.txtST06(0))
   End If
   If frm180301.txtST06(1) <> "" Then
      strCon = strCon & " and s1.ST06<=" & CNULL(frm180301.txtST06(1))
   End If
   
   Screen.MousePointer = vbHourglass
   '出缺勤電子簽核主檔(人事處未簽收的表單),員工請假資料,員工加班資料,員工出差資料
   'Modify By Sindy 2023/12/19
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "Select s1.st06 ST06,a1.a0921 A0901,nvl(a1.A0922,'(舊)'||a3.A0902) AA,s1.st01 BB,s1.st02 CC,decode(B1002,'02','22','03','23',ac02) TT,decode(B1002,'02','加班','03','出差',ac03) DD,sum(nvl(B1009,0)) EE,decode(B1002,'02',sum(nvl(B1012,B1013)),sum(B1010)) FF " & _
               "From ABS010,Staff s1,ACC090NEW a1,allcode,Staff s2,ACC090NEW a2,ACC090 a3 " & _
               "Where B1003=s1.ST01(+) and s1.ST93=a1.A0921(+) and s1.ST03=a3.A0901(+) and B1017=a2.A0921(+) and ac01(+)='04' and B1008=ac02(+) and B1017=s2.ST01(+) and (B1019 is null) " & strCon & strConABS & _
               " group by s1.st06,a1.a0921,nvl(a1.A0922,'(舊)'||a3.A0902),s1.st01,s1.st02,B1002,ac02,ac03"
      If frm180301.CboB1002 = "" Or Left(frm180301.CboB1002, 2) = "01" Then
         strSql = strSql & " union all " & _
                  "Select s1.st06,a1.a0921,nvl(a1.A0922,'(舊)'||a3.A0902),s1.st01,s1.st02,ac02,ac03,sum(sa07),sum(sa08) " & _
                  "From Staff_Absence,Staff s1,ACC090NEW a1,allcode,ACC090 a3 " & _
                  "Where SA01=s1.ST01(+) and s1.ST93=a1.A0921(+) and s1.ST03=a3.A0901(+) and ac01(+)='04' and SA06=ac02(+) and (SA09 is null or (SA09 in(select B1001 from abs010 where B1002='01' and B1003=SA01 and B1019 is not null))) " & strCon & strConSA & _
                  " group by s1.st06,a1.a0921,nvl(a1.A0922,'(舊)'||a3.A0902),s1.st01,s1.st02,ac02,ac03"
      End If
      If frm180301.CboB1008 = "" Then
         If frm180301.CboB1002 = "" Or Left(frm180301.CboB1002, 2) = "02" Then
            strSql = strSql & " union all " & _
                     "Select s1.st06,a1.a0921,nvl(a1.A0922,'(舊)'||a3.A0902),s1.st01,s1.st02,'22','加班',0,sum(nvl(so05,so06)) " & _
                     "From Staff_Overtime,Staff s1,ACC090NEW a1,ACC090 a3 " & _
                     "Where SO01=s1.ST01(+) and s1.ST93=a1.A0921(+) and s1.ST03=a3.A0901(+) and (SO13 is null or (SO13 in(select B1001 from abs010 where B1002='02' and B1003=SO01 and B1019 is not null))) " & strCon & strConSO & _
                     " group by s1.st06,a1.a0921,nvl(a1.A0922,'(舊)'||a3.A0902),s1.st01,s1.st02"
         End If
         If frm180301.CboB1002 = "" Or Left(frm180301.CboB1002, 2) = "03" Then
            strSql = strSql & " union all " & _
                     "Select s1.st06,a1.a0921,nvl(a1.A0922,'(舊)'||a3.A0902),s1.st01,s1.st02,'23','出差',sum(sb06),sum(sb07) " & _
                     "From Staff_Busi_Trip,Staff s1,ACC090NEW a1,ACC090 a3 " & _
                     "Where SB01=s1.ST01(+) and s1.ST93=a1.A0921(+) and s1.ST03=a3.A0901(+) and (SB10 is null or (SB10 in(select B1001 from abs010 where B1002='03' and B1003=SB01 and B1019 is not null))) " & strCon & strConSB & _
                     " group by s1.st06,a1.a0921,nvl(a1.A0922,'(舊)'||a3.A0902),s1.st01,s1.st02"
         End If
      End If
   '2023/12/19 END
   Else
      strSql = "Select s1.st06 ST06,a1.a0901 A0901,a1.a0902 AA,s1.st01 BB,s1.st02 CC,decode(B1002,'02','22','03','23',ac02) TT,decode(B1002,'02','加班','03','出差',ac03) DD,sum(nvl(B1009,0)) EE,decode(B1002,'02',sum(nvl(B1012,B1013)),sum(B1010)) FF " & _
               "From ABS010,Staff s1,ACC090 a1,allcode,Staff s2,ACC090 a2 " & _
               "Where B1003=s1.ST01(+) and s1.ST03=a1.A0901(+) and B1017=a2.A0901(+) and ac01(+)='04' and B1008=ac02(+) and B1017=s2.ST01(+) and (B1019 is null) " & strCon & strConABS & _
               " group by s1.st06,a1.a0901,a1.a0902,s1.st01,s1.st02,B1002,ac02,ac03"
      If frm180301.CboB1002 = "" Or Left(frm180301.CboB1002, 2) = "01" Then
         strSql = strSql & " union all " & _
                  "Select s1.st06,a1.a0901,a1.a0902,s1.st01,s1.st02,ac02,ac03,sum(sa07),sum(sa08) " & _
                  "From Staff_Absence,Staff s1,ACC090 a1,allcode " & _
                  "Where SA01=s1.ST01(+) and s1.ST03=a1.A0901(+) and ac01(+)='04' and SA06=ac02(+) and (SA09 is null or (SA09 in(select B1001 from abs010 where B1002='01' and B1003=SA01 and B1019 is not null))) " & strCon & strConSA & _
                  " group by s1.st06,a1.a0901,a1.a0902,s1.st01,s1.st02,ac02,ac03"
      End If
      If frm180301.CboB1008 = "" Then
         If frm180301.CboB1002 = "" Or Left(frm180301.CboB1002, 2) = "02" Then
            strSql = strSql & " union all " & _
                     "Select s1.st06,a1.a0901,a1.a0902,s1.st01,s1.st02,'22','加班',0,sum(nvl(so05,so06)) " & _
                     "From Staff_Overtime,Staff s1,ACC090 a1 " & _
                     "Where SO01=s1.ST01(+) and s1.ST03=a1.A0901(+) and (SO13 is null or (SO13 in(select B1001 from abs010 where B1002='02' and B1003=SO01 and B1019 is not null))) " & strCon & strConSO & _
                     " group by s1.st06,a1.a0901,a1.a0902,s1.st01,s1.st02"
         End If
         If frm180301.CboB1002 = "" Or Left(frm180301.CboB1002, 2) = "03" Then
            strSql = strSql & " union all " & _
                     "Select s1.st06,a1.a0901,a1.a0902,s1.st01,s1.st02,'23','出差',sum(sb06),sum(sb07) " & _
                     "From Staff_Busi_Trip,Staff s1,ACC090 a1 " & _
                     "Where SB01=s1.ST01(+) and s1.ST03=a1.A0901(+) and (SB10 is null or (SB10 in(select B1001 from abs010 where B1002='03' and B1003=SB01 and B1019 is not null))) " & strCon & strConSB & _
                     " group by s1.st06,a1.a0901,a1.a0902,s1.st01,s1.st02"
         End If
      End If
   End If
   rsTmp.CursorLocation = adUseClient
   'Add By Sindy 2012/6/13 +,"" 未休天數
   'Modify By Sindy 2017/8/16 + 總時數
   rsTmp.Open "select ST06,A0901,AA 部門別,BB 員工代號,CC 姓名,TT,DD 假別,sum(EE) 日,sum(FF) 時,sum(FF) 總時數,'' 未休天數 from (" & strSql & ") group by ST06,A0901,AA,BB,CC,TT,DD order by ST06,A0901,BB,TT", cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryData = True
      Set GRD1.Recordset = rsTmp
      '逐筆計算日,時
      For i = 1 To GRD1.Rows - 1
         ' ”\”整除運算子不可使用於有小數位之數值, 因此先將數值*10做運算
         '99029伊恩一天只上5個小時
         Dim dblDay As Double, dblHour As Double
         If IsNull(GRD1.TextMatrix(i, 7)) Or GRD1.TextMatrix(i, 7) = "" Then GRD1.TextMatrix(i, 7) = 0
         If IsNull(GRD1.TextMatrix(i, 8)) Or GRD1.TextMatrix(i, 8) = "" Then GRD1.TextMatrix(i, 8) = 0
         'Modify By Sindy 2012/7/9 上班時數為特殊者
         Call Pub_GetSpecWorkHour(GRD1.TextMatrix(i, 3), DBDATE(frm180301.txtDate(0)))
'         If GRD1.TextMatrix(i, 3) = "99029" Then
'            If GRD1.TextMatrix(i, 8) >= 5 Then
'               dblDay = (GRD1.TextMatrix(i, 8) * 10) \ (5 * 10)
'               dblHour = Round(GRD1.TextMatrix(i, 8) - (dblDay * 5), 1)
'            End If
         If PUB_bSpecY <> True Then '非過渡期
            If Val(GRD1.TextMatrix(i, 8)) >= Val(PUB_intWkHour) Then
               dblDay = (GRD1.TextMatrix(i, 8) * 10) \ (PUB_intWkHour * 10)
               dblHour = Round(GRD1.TextMatrix(i, 8) - (dblDay * PUB_intWkHour), 1)
               GRD1.TextMatrix(i, 7) = GRD1.TextMatrix(i, 7) + dblDay
               GRD1.TextMatrix(i, 8) = dblHour
            End If
         End If
         'Add By Sindy 2012/6/13 若為特別假則計算未休天數
         If Trim(GRD1.TextMatrix(i, 6)) = "特別假" And strYear >= Left(strSrvDate(2), 3) Then
            GRD1.TextMatrix(i, 10) = GetCurrSpecRestDay(Trim(GRD1.TextMatrix(i, 3)), 2, strYear)
         'Add By Sindy 2024/12/10 若為補休則計算未休天數
         ElseIf Trim(GRD1.TextMatrix(i, 6)) = "補休" And strYear >= Left(strSrvDate(2), 3) Then
            GRD1.TextMatrix(i, 10) = GetCurrFor14RestDay(Trim(GRD1.TextMatrix(i, 3)), 2, frm180301.txtDate(0))
         End If
         '2012/6/13 End
         dblTotHour = dblTotHour + CDbl(Val(GRD1.TextMatrix(i, 9))) 'sum(總時數)
      Next i
      'Add By Sindy 2017/8/16 加班才顯示總時數的總合計
      If Left(frm180301.CboB1002, 2) = "02" Then
         GRD1.AddItem ""
         GRD1.TextMatrix(i, 2) = "總合計"
         GRD1.TextMatrix(i, 9) = dblTotHour
      End If
      '2017/8/16 END
   Else
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Screen.MousePointer = vbDefault
      Unload Me
      frm180301.Show
      Exit Function
   End If
   GRD1.TextMatrix(0, 10) = strYear & "年特/補休未休天數"
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
'   If rsTmp.RecordCount > 0 Then
'      'grd1.Text = "V"
'      For i = 0 To grd1.Cols - 1
'         grd1.col = i
'         grd1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Set rsTmp = Nothing
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180301_02 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Add By Sindy 2012/6/13 +特休未休天數
   'Modify By Sindy 2025/1/14 改為 特/補休未休天數
   arrGridHeadText = Array("ST06", "A0901", "部門別", "員工代號", "姓名", "TT", "假別", "日", "時", "總時數", "特/補休未休天數")
   'Modify By Sindy 2017/8/16 加班才顯示總時數
   If Left(frm180301.CboB1002, 2) = "02" Then
      arrGridHeadWidth = Array(0, 0, 1000, 1000, 1000, 0, 1000, 800, 800, 800, 1700)
   Else
   '2017/8/16 END
      arrGridHeadWidth = Array(0, 0, 1000, 1000, 1000, 0, 1000, 800, 800, 0, 1700)
   End If
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 2
      GRD1.row = dblPrevRow
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   dblPrevRow = GRD1.row
   For i = 0 To GRD1.Cols - 1
      GRD1.col = i
      GRD1.CellBackColor = &HFFC0C0
   Next i
End If
GRD1.Visible = True
End Sub

'Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim nCol As Long, nRow As Long
'   getGrdColRow GRD1, x, y, nCol, nRow
'   If nCol < 0 Or nRow < 0 Then Exit Sub
'   GRD1.col = nCol
'   GRD1.row = nRow
'   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
'      If Me.GRD1.Text = "表單編號" Or Me.GRD1.Text = "天數" Or Me.GRD1.Text = "時數" Then
'         If m_blnColOrderAsc = True Then
'            Me.GRD1.Sort = 3  '數值昇冪
'            m_blnColOrderAsc = False
'         Else
'            Me.GRD1.Sort = 4 '數值降冪
'            m_blnColOrderAsc = True
'         End If
'      Else
'         If m_blnColOrderAsc = True Then
'            Me.GRD1.Sort = 5 '字串昇冪
'            m_blnColOrderAsc = False
'         Else
'            Me.GRD1.Sort = 6 '字串降冪
'            m_blnColOrderAsc = True
'         End If
'      End If
'   End If
'End Sub
