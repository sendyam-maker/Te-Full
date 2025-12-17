VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180301_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "出缺勤查詢－清單"
   ClientHeight    =   5900
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5900
   ScaleWidth      =   8960
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   6
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4620
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   360
      Index           =   1
      Left            =   6640
      TabIndex        =   1
      Top             =   60
      Width           =   1365
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "明細資料"
      Height          =   360
      Index           =   0
      Left            =   5645
      TabIndex        =   0
      Top             =   60
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Index           =   2
      Left            =   8055
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm180301_01.frx":0000
      Height          =   4785
      Left            =   60
      TabIndex        =   3
      Top             =   450
      Width           =   8835
      _ExtentX        =   15593
      _ExtentY        =   8431
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
Attribute VB_Name = "frm180301_01"
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
'Added by Sindy 2021/12/20
Dim strPrinter As String
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim m_intColumn As Integer
'2021/12/20 END


'查詢明細資料
Public Sub PubShowNextData()
   Select Case cmdState
      Case 0 '明細資料
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then
               GRD1.col = 0
               GRD1.Text = ""
               For j = 0 To GRD1.Cols - 1
                  GRD1.col = j
                  GRD1.CellBackColor = QBColor(15)
               Next j
               GRD1.col = 12 '表單編號
               If Not IsNull(GRD1.Text) Then
                  Screen.MousePointer = vbHourglass
                  Me.Hide
                  'Add By Sindy 2013/6/26
                  If GRD1.TextMatrix(i, 13) = "5" Then
                     Call frm180301_04.SetParent(Me)
                     frm180301_04.m_OG01 = Pub_RplStr(GRD1.TextMatrix(i, 15))
                     frm180301_04.QueryData
                     frm180301_04.Show
                  Else
                  '2013/6/26 END
                     Call frm180301_03.SetParent(Me)
                     If GRD1.TextMatrix(i, 13) = "1" Then '出缺勤
                        frm180301_03.txtB1001 = Pub_RplStr(GRD1.Text)
                        frm180301_03.QueryData
                     Else
                        frm180301_03.txtB1003 = Pub_RplStr(GRD1.TextMatrix(i, 2))
                        frm180301_03.m_SA02 = Pub_RplStr(GRD1.TextMatrix(i, 14))
                        frm180301_03.m_SA03 = Pub_RplStr(GRD1.TextMatrix(i, 15))
                        If GRD1.TextMatrix(i, 13) = "2" Then '請假
                           frm180301_03.QueryData_2
                        ElseIf GRD1.TextMatrix(i, 13) = "3" Then '加班
                           frm180301_03.QueryData_3
                        ElseIf GRD1.TextMatrix(i, 13) = "4" Then '出差
                           frm180301_03.QueryData_4
                        End If
                     End If
                     frm180301_03.Show
                  End If
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
            End If
         Next i
         Me.Enabled = True
      Case 1 '回前畫面
         Unload Me
         frm180301.Show
      Case 2 '結束
         Unload Me
         Unload frm180301
      Case Else
   End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
Dim strConABS As String, strConSA As String, strConSO As String, strConSB As String
Dim strConOut As String 'Add By Sindy 2013/6/26
   
   QueryData = False
   
   m_blnColOrderAsc = True
   GRD1.Clear
   SetGrd
   
   strCon = "": strConABS = "": strConSA = "": strConSO = "": strConSB = "": strConOut = ""
   '表單日期
   If frm180301.txtDate(0) <> "" And frm180301.txtDate(1) <> "" Then
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
      'Add By Sindy 2013/6/26
      strConOut = strConOut & " and (og02 between " & DBDATE(frm180301.txtDate(0)) & " and " & DBDATE(frm180301.txtDate(1)) & ") "
      '2013/6/26 END
   End If
   '部門別
'   If frm180301.txtDept(0) <> "" Then
'      strCon = strCon & " and a1.A0901>=" & CNULL(frm180301.txtDept(0))
'   End If
'   If frm180301.txtDept(1) <> "" Then
'      strCon = strCon & " and a1.A0901<=" & CNULL(frm180301.txtDept(1))
'   End If
   'Modify By Sindy 2021/12/21
   If frm180301.CboDept(0) <> "" Then
      'Modify By Sindy 2023/12/19 新部門
      strCon = strCon & " and a1.A0921>=" & CNULL(Left(Trim(frm180301.CboDept(0)), 3))
   End If
   If frm180301.CboDept(1) <> "" Then
      'Modify By Sindy 2023/12/19 新部門
      strCon = strCon & " and a1.A0921<=" & CNULL(Left(Trim(frm180301.CboDept(1)), 3))
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
         'Modify By Sindy 2023/12/19 新部門
         strCon = strCon & " and a1.A0921 in(" & frm180301.m_IsAbsBossST03 & ")"
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
   '出缺勤電子簽核主檔(人事處未簽收或註銷的表單),員工請假資料,員工加班資料,員工出差資料
   'Modify By Sindy 2023/12/19 +新部門
   strSql = "Select ' ' as V,nvl(a1.A0922,'(舊)'||a3.A0902) 部門別,s1.ST01 員工代號,s1.ST02 姓名," & B1002CName & " 表單類別,AC03 假別,sqldateT(B1004)||' '||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),3,2) 起始日期時間,sqldateT(decode(B1002,'02',B1004,B1006))||' '||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),3,2) 結束日期時間,B1009 天數,decode(B1002,'02',nvl(B1012,B1013),B1010)||'' 時數,decode(B1018,'" & 註銷 & "',' ',nvl(s2.ST02,a2.A0922)) 待處理人員," & B1018CName & " 表單狀態,B1001 表單編號,'1' TableID,B1004,B1005,s1.ST06 ST06,nvl(s1.ST93,s1.ST03) ST93,s1.ST01 ST01 " & _
            "From ABS010,Staff s1,ACC090NEW a1,allcode,Staff s2,ACC090NEW a2,ACC090 a3 " & _
            "Where B1003=s1.ST01(+) and s1.ST93=a1.A0921(+) and s1.ST03=a3.A0901(+) and B1017=a2.A0921(+) and ac01(+)='04' and B1008=ac02(+) and B1017=s2.ST01(+) and (B1019 is null or B1018='" & 註銷 & "') " & strCon & strConABS
   If frm180301.CboB1002 = "" Or Left(frm180301.CboB1002, 2) = "01" Then
      strSql = strSql & " union " & _
               "Select ' ' as V,nvl(a1.A0922,'(舊)'||a3.A0902) 部門別,s1.ST01 員工代號,s1.ST02 姓名,'請假' 表單類別,AC03 假別,sqldateT(SA02)||' '||substr(ltrim(to_char('0000'||to_char(SA03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SA03),'0000')),3,2) 起始日期時間,sqldateT(SA04)||' '||substr(ltrim(to_char('0000'||to_char(SA05),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SA05),'0000')),3,2) 結束日期時間,SA07 天數,SA08||'' 時數,'' 待處理人員,'' 表單狀態,SA09 表單編號,'2' TableID,SA02,SA03,s1.ST06 ST06,nvl(s1.ST93,s1.ST03) ST93,s1.ST01 ST01 " & _
               "From Staff_Absence,Staff s1,ACC090NEW a1,allcode,ACC090 a3 " & _
               "Where SA01=s1.ST01(+) and s1.ST93=a1.A0921(+) and s1.ST03=a3.A0901(+) and ac01(+)='04' and SA06=ac02(+) and (SA09 is null or (SA09 in(select B1001 from abs010 where B1002='01' and B1003=SA01 and B1019 is not null))) " & strCon & strConSA
   End If
   If frm180301.CboB1008 = "" Then
      If frm180301.CboB1002 = "" Or Left(frm180301.CboB1002, 2) = "02" Then
         strSql = strSql & " union " & _
                  "Select ' ' as V,nvl(a1.A0922,'(舊)'||a3.A0902) 部門別,s1.ST01 員工代號,s1.ST02 姓名,'加班' 表單類別,'' 假別,sqldateT(SO02)||' '||substr(ltrim(to_char('0000'||to_char(SO03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SO03),'0000')),3,2) 起始日期時間,sqldateT(SO02)||' '||substr(ltrim(to_char('0000'||to_char(SO04),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SO04),'0000')),3,2) 結束日期時間,0 天數,nvl(SO05,SO06)||'' 時數,'' 待處理人員,'' 表單狀態,SO13 表單編號,'3' TableID,So02,So03,s1.ST06 ST06,nvl(s1.ST93,s1.ST03) ST93,s1.ST01 ST01 " & _
                  "From Staff_Overtime,Staff s1,ACC090NEW a1,ACC090 a3 " & _
                  "Where SO01=s1.ST01(+) and s1.ST93=a1.A0921(+) and s1.ST03=a3.A0901(+) and (SO13 is null or (SO13 in(select B1001 from abs010 where B1002='02' and B1003=SO01 and B1019 is not null))) " & strCon & strConSO
      End If
      If frm180301.CboB1002 = "" Or Left(frm180301.CboB1002, 2) = "03" Then
         strSql = strSql & " union " & _
                  "Select ' ' as V,nvl(a1.A0922,'(舊)'||a3.A0902) 部門別,s1.ST01 員工代號,s1.ST02 姓名,'出差' 表單類別,'' 假別,sqldateT(SB02)||' '||substr(ltrim(to_char('0000'||to_char(SB03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SB03),'0000')),3,2) 起始日期時間,sqldateT(SB04)||' '||substr(ltrim(to_char('0000'||to_char(SB05),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SB05),'0000')),3,2) 結束日期時間,SB06 天數,SB07||'' 時數,'' 待處理人員,'' 表單狀態,SB10 表單編號,'4' TableID,SB02,SB03,s1.ST06 ST06,nvl(s1.ST93,s1.ST03) ST93,s1.ST01 ST01 " & _
                  "From Staff_Busi_Trip,Staff s1,ACC090NEW a1,ACC090 a3 " & _
                  "Where SB01=s1.ST01(+) and s1.ST93=a1.A0921(+) and s1.ST03=a3.A0901(+) and (SB10 is null or (SB10 in(select B1001 from abs010 where B1002='03' and B1003=SB01 and B1019 is not null))) " & strCon & strConSB
      End If
      'Add By Sindy 2013/6/26
      If frm180301.CboB1002 = "" Or Left(frm180301.CboB1002, 2) = "04" Then
         strSql = strSql & " union " & _
                  "Select ' ' as V,nvl(a1.A0922,'(舊)'||a3.A0902) 部門別,s1.ST01 員工代號,s1.ST02 姓名,'外出' 表單類別,'' 假別,sqldateT(og02)||' '||og19 起始日期時間,sqldateT(og02)||' '||og20 結束日期時間,0 天數,og05||'' 時數,'' 待處理人員,'' 表單狀態,'' 表單編號,'5' TableID,og02,to_number(og01),s1.ST06 ST06,nvl(s1.ST93,s1.ST03) ST93,s1.ST01 ST01 " & _
                  "From outgoing,Staff s1,ACC090NEW a1,ACC090 a3 " & _
                  "Where og03=s1.ST01(+) and s1.ST93=a1.A0921(+) and s1.ST03=a3.A0901(+) " & strCon & strConOut
      End If
      '2013/6/26 END
   End If
   strSql = strSql & " order by ST06,ST93,ST01,7 "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryData = True
      Set GRD1.Recordset = rsTmp
      'Add By Sindy 2013/6/26
      For i = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(i, 4) = "外出" Then
            '天數和時數放空白
            GRD1.TextMatrix(i, 8) = ""
            GRD1.TextMatrix(i, 9) = ""
         End If
      Next i
      '2013/6/26 END
   Else
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Screen.MousePointer = vbDefault
      Unload Me
      frm180301.Show
      Exit Function
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
'   If rsTmp.RecordCount > 0 Then
'      grd1.Text = "V"
'      For i = 0 To grd1.Cols - 1
'         grd1.col = i
'         grd1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Set rsTmp = Nothing
   
   Screen.MousePointer = vbDefault
End Function

Private Sub Form_Load()
   If frm180301.cmdOK(1).Tag = "" Then
      MoveFormToCenter Me
   End If
   
   'Add By Sindy 2021/12/20
   If GetStaffDepartment(strUserNum) = "M51" Or _
      GetStaffDepartment(strUserNum) = "M21" Then
      cmdPrint.Visible = True
      Frame1.Visible = True
      PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True
   Else
      cmdPrint.Visible = False
      Frame1.Visible = False
   End If
   '2021/12/20 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180301_01 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   '                        0    1         2           3       4           5       6               7               8       9       10            11          12          13         14      15      16      17      18
   arrGridHeadText = Array("V", "部門別", "員工代號", "姓名", "表單類別", "假別", "起始日期時間", "結束日期時間", "天數", "時數", "待處理人員", "表單狀態", "表單編號", "TableID", "SA02", "SA03", "ST06", "ST93", "ST01")
   arrGridHeadWidth = Array(200, 900, 550, 650, 500, 600, 1250, 1250, 400, 400, 650, 650, 800, 0, 0, 0, 0, 0, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
'      If iRow = 8 Or iRow = 9 Then '時數
'         GRD1.CellAlignment = flexAlignLeftCenter 'flexAlignRightCenter
'      Else
         GRD1.CellAlignment = flexAlignCenterCenter
'      End If
   Next
   GRD1.Visible = True
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   If GRD1.TextMatrix(GRD1.MouseRow, 1) <> "" Then
      If GRD1.Text = "V" Then
         GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
      Else
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
End If
GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   'Add By Sindy 2012/4/16 '部門別置換為使用部門別代碼做排序
   If nCol = 1 Then nCol = 17
   '2012/4/16 End
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      'Modify By Sindy 2012/4/16 + Me.GRD1.Text = "部門別" Or
      If Me.GRD1.Text = "部門別" Or Me.GRD1.Text = "表單編號" Or Me.GRD1.Text = "天數" Or Me.GRD1.Text = "時數" Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

'Add By Sindy 2021/12/21
Private Sub cmdPrint_Click()
Screen.MousePointer = vbHourglass

PUB_SetOsDefaultPrinter Combo1

If GRD1.Rows > 1 And GRD1.TextMatrix(GRD1.Rows - 1, 1) <> "" Then
   PrintExcel
Else
   MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
   Screen.MousePointer = vbDefault
   Exit Sub
End If

PUB_SetOsDefaultPrinter strPrinter
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub PrintExcel()
Dim ii As Integer
Dim strTempFile As String
   
   '預設A4紙張/橫式/比例 80%/水平置中/邊界左右都改0-瑞婷
   Set xlsAnnuity = New Excel.Application
   'xlsAnnuity.Visible = True
   xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   wksAnnuity.Cells.NumberFormatLocal = "@" '文字 Add By Sindy 2024/9/5
   xlsAnnuity.ActiveWindow.Zoom = 75 '畫面比例100%太大了,調整為75%
   '把Excel的警告訊息關掉
   xlsAnnuity.DisplayAlerts = False
   
   wksAnnuity.PageSetup.PaperSize = 9 'A4
   wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
   'wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity.PageSetup.LeftMargin = 0 '邊界
   wksAnnuity.PageSetup.RightMargin = 0
   wksAnnuity.PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.4)
   wksAnnuity.PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.5)
   wksAnnuity.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
   
'   xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
'   xlsAnnuity.Workbooks.add
'   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   wksAnnuity.Activate
   
   '設定各欄位長度
   wksAnnuity.Columns("A:A").ColumnWidth = 6 'Add By Sindy 2024/9/5
   wksAnnuity.Columns("B:B").ColumnWidth = 10
   wksAnnuity.Columns("C:C").ColumnWidth = 9
   wksAnnuity.Columns("D:D").ColumnWidth = 9
   wksAnnuity.Columns("E:E").ColumnWidth = 9
   wksAnnuity.Columns("F:F").ColumnWidth = 9
   wksAnnuity.Columns("G:G").ColumnWidth = 15
   wksAnnuity.Columns("H:H").ColumnWidth = 15
   wksAnnuity.Columns("I:I").ColumnWidth = 8
   wksAnnuity.Columns("J:J").ColumnWidth = 8
   wksAnnuity.Columns("K:K").ColumnWidth = 11
   wksAnnuity.Columns("L:L").ColumnWidth = 10
   wksAnnuity.Columns("M:M").ColumnWidth = 8
   
   '標題
   m_intColumn = 1
   xlsAnnuity.Range("D" & m_intColumn).Value = "出缺勤明細"
   xlsAnnuity.Range("A" & m_intColumn & ":" & "M" & m_intColumn).Select
   With xlsAnnuity.Selection
      .HorizontalAlignment = xlCenter '置中
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = True
   End With
   With xlsAnnuity.Selection.Font
      .Bold = True '粗體
      .Name = "新細明體"
      .Size = 16
   End With
   m_intColumn = m_intColumn + 1
   xlsAnnuity.Range("A" & m_intColumn).Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   m_intColumn = m_intColumn + 2
   xlsAnnuity.Range("A" & m_intColumn).Value = "所別" 'Add By Sindy 2024/9/5
   xlsAnnuity.Range("B" & m_intColumn).Value = "部門別"
   'xlsAnnuity.Range("B" & m_intColumn).HorizontalAlignment = xlCenter
   xlsAnnuity.Range("C" & m_intColumn).Value = "員工代號"
   'xlsAnnuity.Range("C" & m_intColumn).HorizontalAlignment = xlCenter
   xlsAnnuity.Range("D" & m_intColumn).Value = "姓名"
   xlsAnnuity.Range("E" & m_intColumn).Value = "表單類別"
   xlsAnnuity.Range("F" & m_intColumn).Value = "假別"
   xlsAnnuity.Range("G" & m_intColumn).Value = "起始日期時間"
   xlsAnnuity.Range("H" & m_intColumn).Value = "結束日期時間"
   xlsAnnuity.Range("I" & m_intColumn).Value = "天數"
   xlsAnnuity.Range("J" & m_intColumn).Value = "時數"
   xlsAnnuity.Range("K" & m_intColumn).Value = "待處理人員"
   xlsAnnuity.Range("L" & m_intColumn).Value = "表單狀態"
   xlsAnnuity.Range("M" & m_intColumn).Value = "表單編號"
   xlsAnnuity.Range("A" & m_intColumn & ":" & "M" & m_intColumn).Select
'   With xlsAnnuity.Selection
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlCenter
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .IndentLevel = 0
'       .ShrinkToFit = False
'       .ReadingOrder = xlContext
'       .MergeCells = False
'   End With
'   xlsAnnuity.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'   xlsAnnuity.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'   xlsAnnuity.Selection.Borders(xlEdgeLeft).LineStyle = xlNone
   With xlsAnnuity.Selection.Borders(xlEdgeTop)
       .LineStyle = xlContinuous
       .ColorIndex = xlAutomatic
       .tintandshade = 0
       .Weight = xlThin
   End With
   With xlsAnnuity.Selection.Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .ColorIndex = xlAutomatic
       .tintandshade = 0
       .Weight = xlThin
   End With
   
   '列印明細
   With GRD1
   For ii = 1 To .Rows - 1
      m_intColumn = m_intColumn + 1
      'Add By Sindy 2024/9/5
      If .TextMatrix(ii, 16) = "1" Then
         xlsAnnuity.Range("A" & m_intColumn).Value = "北所"
      ElseIf .TextMatrix(ii, 16) = "2" Then
         xlsAnnuity.Range("A" & m_intColumn).Value = "中所"
      ElseIf .TextMatrix(ii, 16) = "3" Then
         xlsAnnuity.Range("A" & m_intColumn).Value = "南所"
      ElseIf .TextMatrix(ii, 16) = "4" Then
         xlsAnnuity.Range("A" & m_intColumn).Value = "高所"
      End If
      '2024/9/5 END
      xlsAnnuity.Range("B" & m_intColumn).Value = .TextMatrix(ii, 1)
      xlsAnnuity.Range("C" & m_intColumn).Value = .TextMatrix(ii, 2)
      xlsAnnuity.Range("D" & m_intColumn).Value = .TextMatrix(ii, 3)
      xlsAnnuity.Range("E" & m_intColumn).Value = .TextMatrix(ii, 4)
      xlsAnnuity.Range("F" & m_intColumn).Value = .TextMatrix(ii, 5)
      xlsAnnuity.Range("G" & m_intColumn).Value = .TextMatrix(ii, 6)
      xlsAnnuity.Range("H" & m_intColumn).Value = .TextMatrix(ii, 7)
      xlsAnnuity.Range("I" & m_intColumn).Value = .TextMatrix(ii, 8)
      xlsAnnuity.Range("J" & m_intColumn).Value = .TextMatrix(ii, 9)
      xlsAnnuity.Range("K" & m_intColumn).Value = .TextMatrix(ii, 10)
      xlsAnnuity.Range("L" & m_intColumn).Value = .TextMatrix(ii, 11)
      xlsAnnuity.Range("M" & m_intColumn).Value = .TextMatrix(ii, 12)
   Next ii
   m_intColumn = m_intColumn + 2
   xlsAnnuity.Range("A" & m_intColumn).Value = "共 " & .Rows - 1 & " 筆"
   End With
   xlsAnnuity.ActiveSheet.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
   'Modify By Sindy 2022/1/19 列印標題
   xlsAnnuity.ActiveSheet.PageSetup.PrintTitleRows = "$1:$4"
   
   strTempFile = App.path & "\$$demo.pdf"
   If Val(xlsAnnuity.Version) < 12 Then
      xlsAnnuity.Workbooks(1).SaveAs FileName:=strTempFile, FileFormat:=-4143
   Else
      xlsAnnuity.Workbooks(1).SaveAs FileName:=strTempFile, FileFormat:=56
   End If
   xlsAnnuity.Workbooks(1).PrintOut
   
'   'xlTypePDF   0  PDF：可攜式文件格式檔案 (.pdf)
'   'Quality:=xlQualityStandard 0  標準品質
'   '各參數解釋:
'   'Type 必要  XlFixedFormatTyp  匯出目標的檔案格式類型。
'   'FileName   選用  Variant  要儲存之檔案的檔案名稱。 可以包含完整路徑，否則 Microsoft Excel 會將檔案儲存在目前的資料夾中。
'   'Quality 選用  Variant  選用 XlFixedFormatQuality。 這會指定已發佈檔案的品質。
'   'IncludeDocProperties 選用  Variant  若要包含檔案屬性，則為 True 。否則 為 False。
'   'IgnorePrintAreas  選用  Variant  True 是表示忽略所有發佈時設定的列印範圍;否則 為 False。
'   'From 選用  Variant  要發佈的起始頁碼。 如果省略此引數，將從頭開始列印。
'   'To   選用  Variant  要發佈的最後一頁頁碼。 如果省略此引數，將發佈至最後一頁。
'   'OpenAfterPublish 選用  Variant  True 是表示在發佈後在檢視器中顯示檔案;否則 為 False。
'   'FixedFormatExtClassPtr 選用  Variant  FixedFormatExt 類別的指標。
'   xlsAnnuity.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strTempFile, Quality:=0, _
'   IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
'   'ShellExecute 0, "open", strTempFile, vbNullString, vbNullString, 1
   
   xlsAnnuity.Workbooks.Close 'SaveChanges:=False
   xlsAnnuity.Quit
   Set xlsAnnuity = Nothing
End Sub
