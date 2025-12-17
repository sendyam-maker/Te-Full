VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160501 
   BorderStyle     =   1  '單線固定
   Caption         =   "加班單異常查詢"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   8955
   Tag             =   "加班資料"
   Begin VB.CommandButton cmdOK 
      Caption         =   "消除異常"
      Height          =   360
      Left            =   5300
      TabIndex        =   18
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   780
      Left            =   2940
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "frm160501.frx":0000
      Top             =   4950
      Width           =   5520
   End
   Begin VB.CheckBox Check1 
      Caption         =   "查詢全部加班單"
      Height          =   255
      Left            =   555
      TabIndex        =   6
      Top             =   1050
      Width           =   1635
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   360
      Left            =   4230
      TabIndex        =   7
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7920
      TabIndex        =   9
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "當日打卡明細"
      Height          =   360
      Left            =   6370
      TabIndex        =   8
      Top             =   90
      Width           =   1455
   End
   Begin VB.TextBox txtST01 
      Height          =   300
      Index           =   1
      Left            =   2640
      MaxLength       =   6
      TabIndex        =   3
      Top             =   360
      Width           =   1005
   End
   Begin VB.TextBox txtDept 
      Height          =   300
      Index           =   1
      Left            =   2130
      MaxLength       =   3
      TabIndex        =   5
      Top             =   690
      Width           =   495
   End
   Begin VB.TextBox txtDept 
      Height          =   300
      Index           =   0
      Left            =   1500
      MaxLength       =   3
      TabIndex        =   4
      Top             =   690
      Width           =   495
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Index           =   1
      Left            =   2640
      MaxLength       =   7
      TabIndex        =   1
      Top             =   30
      Width           =   1005
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Index           =   0
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   0
      Top             =   30
      Width           =   1005
   End
   Begin VB.TextBox txtST01 
      Height          =   300
      Index           =   0
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   2
      Top             =   360
      Width           =   1005
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3555
      Left            =   60
      TabIndex        =   10
      Top             =   1350
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   6279
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|加班日期|部門|員工姓名|加班起始時間|加班迄止時間|上班打卡|下班打卡|應下班時段"
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
      _Band(0).Cols   =   9
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "上/下班打卡若有異常，其欄位會以淺藍色顯示"
      ForeColor       =   &H00C000C0&
      Height          =   180
      Left            =   5220
      TabIndex        =   17
      Top             =   1140
      Width           =   3645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "顏色說明：整列紅色代表加班單異常"
      ForeColor       =   &H00C000C0&
      Height          =   180
      Left            =   4350
      TabIndex        =   16
      Top             =   930
      Width           =   2880
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "注意：請勾選下面資料列，再進行明細查詢"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   4350
      TabIndex        =   14
      Top             =   720
      Width           =   3420
   End
   Begin VB.Line Line3 
      X1              =   2430.357
      X2              =   2670.491
      Y1              =   509.557
      Y2              =   509.557
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "部  門  別："
      Height          =   180
      Left            =   555
      TabIndex        =   13
      Top             =   750
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   1920.072
      X2              =   2400.34
      Y1              =   840.269
      Y2              =   840.269
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  '內實線
      X1              =   2460.374
      X2              =   2670.491
      Y1              =   149.87
      Y2              =   149.87
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      Height          =   180
      Left            =   555
      TabIndex        =   12
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Left            =   555
      TabIndex        =   11
      Top             =   420
      Width           =   900
   End
End
Attribute VB_Name = "frm160501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/14 Form2.0已修改
'Created by Sindy 2013/8/7
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Public bolClose As Boolean
Dim i As Integer, j As Integer


'打卡明細
Private Sub cmdDetail_Click()
'Dim bolSelisV As Boolean
'
'   bolSelisV = False
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.TextMatrix(i, 0) = "V" Then
'         bolSelisV = True
         GRD1.Text = ""
         For j = 0 To GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = QBColor(15)
         Next j
         bolClose = False
         Call frm180303_1.SetParent(Me)
         frm180303_1.m_B1401 = GRD1.TextMatrix(i, 9)
         frm180303_1.m_B1402 = GRD1.TextMatrix(i, 1)
         If frm180303_1.QueryData = True Then
            frm180303_1.Show vbModal '強制回應表單
         Else
            Unload frm180303_1
         End If
         If bolClose = True Then
            Exit Sub
         End If
      End If
   Next i
'   If bolSelisV = False Then
'      MsgBox "請勾選欲查詢的資料！"
'   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

'消除異常
Private Sub cmdok_Click()
Dim bolSelisV As Boolean
   
   bolSelisV = False
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.TextMatrix(i, 0) = "V" Then
         bolSelisV = True
         GRD1.Text = ""
         For j = 0 To GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = QBColor(15)
         Next j
         If SaveDB(i) = False Then
            Exit Sub
         End If
      End If
   Next i
   If bolSelisV = False Then
      MsgBox "請勾選欲消除異常的資料！"
   Else
      Call cmdQuery_Click
   End If
End Sub

Private Function SaveDB(intRow As Integer) As Boolean
   
On Error GoTo ErrHand
   
   SaveDB = False
   cnnConnection.BeginTrans
   strSql = "UPDATE staff_overtime SET so14='Y'" & _
            " WHERE so01='" & GRD1.TextMatrix(intRow, 9) & "' and so02=" & DBDATE(GRD1.TextMatrix(intRow, 1))
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   
   SaveDB = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "資料存檔失敗！" & vbCrLf & Err.Description
End Function

Private Sub cmdQuery_Click()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCon As String, strConSCD As String
   
   m_blnColOrderAsc = True
   
   GRD1.Clear
   SetGrd
   
   strCon = "": strConSCD = ""
   If Val(txtDate(0)) = 0 Or Val(txtDate(1)) = 0 Then
      MsgBox "請輸入起迄日期！", vbExclamation, "操作錯誤！"
      If Val(txtDate(0)) = 0 Then txtDate(0).SetFocus
      If Val(txtDate(1)) = 0 Then txtDate(1).SetFocus
      Exit Sub
   End If
   
   '員工代號
   If txtST01(0) <> "" And txtST01(1) <> "" Then
      strCon = strCon & " and ST01>='" & txtST01(0) & "' and ST01<='" & txtST01(1) & "'"
      strConSCD = strConSCD & " and scd01>='" & txtST01(0) & "' and scd01<='" & txtST01(1) & "'"
   End If
   '部門別
   If txtDept(0) <> "" And txtDept(1) <> "" Then
      'Modify By Sindy 2023/12/27 部門調整改抓ST93
      strCon = strCon & " and ST93>='" & txtDept(0) & "' and ST93<='" & txtDept(1) & "'"
   End If
   
   '只查詢異常
   If Check1.Value = 0 Then
      'strCon = strCon & " and so03<=1210 and so04>=1330" '排除中午加班者
      strCon = strCon & " and so14 is null" '排除已消除異常者
      strCon = strCon & " and (b1401 is not null" & _
                              " or (so03<=substr(max_pr02,1,length(max_pr02)-2) and so04>substr(max_pr02,1,length(max_pr02)-2))" & _
                              " or ((so03*100)<decode(sign(V4.min_pr02-80000),-1,170000,decode(sign(V4.min_pr02-83000),-1,173000,180000)) and V4.min_pr02 is not null)" & _
                             ")"
   End If
   
   Screen.MousePointer = vbHourglass
   'Modify By Sindy 2023/12/27 部門調整改抓ST93
   strSql = "select distinct ' ' as V,sqldatet(so02) as 加班日期,nvl(A0922,'(舊)'||A0902) as 部門,st02 as 姓名," & _
            "substr(sqltime(so03||'00'),1,5) as 起始時間,substr(sqltime(so04||'00'),1,5) as 迄止時間," & _
            "sqltime6(V4.min_pr02) as 上班打卡,sqltime6(V4.max_pr02) as 下班打卡," & _
            "decode(V4.min_pr02,null,'',decode(sign(V4.min_pr02-80000),-1,'17:00',decode(sign(V4.min_pr02-83000),-1,'17:30','18:00'))) as 應下班時段," & _
            "st01,nvl(A0922,A0902) A0901,b1401,' ' as Error,V4.min_pr02 as min_pr02,so02,so03," & _
            "decode(V1.AErr,1,'A','')||decode(V1.PErr,1,'P','') as 打卡異常狀況,so14 as 核銷" & _
            " from staff_overtime,staff," & _
            "(select b1401,b1402,sum(decode(b1403,'A',1,0)) as AErr,sum(decode(b1403,'P',1,0)) as PErr from abs014" & _
            " where b1402 between " & DBDATE(txtDate(0)) & " and " & DBDATE(txtDate(1)) & _
            " group by b1401,b1402) V1," & _
            "(select scd01,pr01,nvl(min(pr02),0) as min_pr02,nvl(max(pr02),0) as max_pr02" & _
            " from pollrecord,staffcarddata where pr03=scd02(+)" & _
            " and pr01>=" & DBDATE(txtDate(0)) & " and pr01<=" & DBDATE(txtDate(1)) & strConSCD & " group by scd01,pr01) V4,ACC090,ACC090NEW" & _
            " where so01=st01(+)" & _
            " and so01=b1401(+) and so02=b1402(+)" & _
            " and so01=V4.scd01(+) and so02=V4.pr01(+)" & _
            " and ST03=A0901(+) and ST93=A0921(+)" & _
            " and so02 between " & DBDATE(txtDate(0)) & " and " & DBDATE(txtDate(1)) & strCon & _
            " order by so02,A0901,st01,so03"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      For i = 1 To GRD1.Rows - 1
'         '排除中午加班者
'         If Val(Replace(GRD1.TextMatrix(i, 4), ":", "")) >= 1210 And Val(Replace(GRD1.TextMatrix(i, 5), ":", "")) <= 1330 Then
'            GoTo ReadNext
'         End If
         '排除已消除異常者
         If Trim(GRD1.TextMatrix(i, 17)) <> "" Then
            GoTo ReadNext
         End If
         
         'Add By Sindy 2013/8/26
         '排除非工作天者
         If ChkWorkDay(DBDATE(GRD1.TextMatrix(i, 1)), GRD1.TextMatrix(i, 9), True) = False Then
            If Check1.Value = 0 Then
               GRD1.RowHeight(i) = 0
            End If
            GoTo ReadNext
         End If
         '2013/8/26 END
         
         '上/下班有異常則異常
         If Trim(GRD1.TextMatrix(i, 11)) <> "" Then
            GRD1.TextMatrix(i, 12) = "有異常"
            Call SetColColor(i)
            GoTo ReadNext
         End If
         '加班起始時間是在上班時段內，但加班迄止時間晚於下班打卡時間
         If (Val(Replace(GRD1.TextMatrix(i, 4), ":", "")) * 100) <= Val(Replace(GRD1.TextMatrix(i, 7), ":", "")) And _
            (Val(Replace(GRD1.TextMatrix(i, 5), ":", "")) * 100) > Val(Replace(GRD1.TextMatrix(i, 7), ":", "")) Then
            GRD1.TextMatrix(i, 12) = "有異常"
            Call SetColColor(i)
            GoTo ReadNext
         End If
         '加班起始時間必須在下班時段之後
         If Val(Replace(GRD1.TextMatrix(i, 4), ":", "")) < Val(Replace(GRD1.TextMatrix(i, 8), ":", "")) Then
            GRD1.TextMatrix(i, 12) = "有異常"
            Call SetColColor(i)
            GoTo ReadNext
         End If
ReadNext:
      Next i
   Else
      ShowNoData
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
'   If rsTmp.RecordCount > 0 Then
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD1.Visible = True
   
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
   Screen.MousePointer = vbDefault
End Sub

'異常資料以紅色標註
Private Function SetColColor(intRow As Integer) As Boolean
   SetColColor = False
   GRD1.row = intRow
   '加班單異常顯示
   If Trim(GRD1.TextMatrix(intRow, 12)) = "有異常" Then
      SetColColor = True
      For j = 0 To GRD1.Cols - 1
         GRD1.col = j
         GRD1.CellBackColor = &H8080FF
      Next j
   End If
   '上班打卡異常顯示
   If InStr(Trim(GRD1.TextMatrix(intRow, 16)), "A") > 0 Then
      GRD1.col = 6
      GRD1.CellBackColor = &HFFFF80
   End If
   '下班打卡異常顯示
   If InStr(Trim(GRD1.TextMatrix(intRow, 16)), "P") > 0 Then
      GRD1.col = 7
      GRD1.CellBackColor = &HFFFF80
   End If
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   '前一個月的第一天
   'txtDate(0) = Left(ChangeWStringToTString(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##")))), 5) & "01"
'   '取得前一日的工作天
'   For i = 1 To 12
'      txtDate(0) = Val(DBDATE(DateAdd("d", Val("-" & i), ChangeWStringToWDateString(strSrvDate(1))))) - 19110000
'      If ChkWorkDay(DBDATE(DateAdd("d", Val("-" & i), ChangeWStringToWDateString(strSrvDate(1))))) = True Then
'         Exit For
'      End If
'   Next i
   '當月1日
   txtDate(0) = Left(ChangeWStringToTString(strSrvDate(1)), 5) & "01"
   txtDate(1) = strSrvDate(2)
'   txtDept(0) = Pub_StrUserSt03
'   txtDept(1) = Pub_StrUserSt03
'   txtST01(0) = strUserNum
'   txtST01(1) = strUserNum
   
   SetGrd
   'Call cmdQuery_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160501 = Nothing
End Sub

' 初始化列表
Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("V", "加班日期", "部門", "員工姓名", "加班起始時間", "加班迄止時間", _
                           "上班打卡", "下班打卡", "應下班時段", "st01", "A0901", "b1401", "Error", _
                           "min_pr02", "so02", "so03", "打卡異常狀況", "核銷")
   arrGridHeadWidth = Array(200, 900, 1200, 900, 1200, 1200, _
                           900, 900, 1000, 0, 0, 0, 0, _
                           0, 0, 0, 0, 500)
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
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   If GRD1.TextMatrix(GRD1.MouseRow, 1) <> "" Then
      If GRD1.Text = "V" Then
         GRD1.Text = ""
         If SetColColor(GRD1.MouseRow) = False Then
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               GRD1.CellBackColor = QBColor(15)
            Next i
         End If
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
   If nCol = 2 Then nCol = 10 '部門別置換為使用部門別代碼做排序
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If Me.GRD1.Text = "部門別" Then
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

Private Sub txtST01_GotFocus(Index As Integer)
   InverseTextBox txtST01(Index)
End Sub

Private Sub txtST01_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'2013/7/26 ADD BY SONIA
Private Sub txtST01_LostFocus(Index As Integer)
   Select Case Index
      Case 0
         txtST01(1) = txtST01(0)
   End Select
End Sub
'2013/7/26 END

Private Sub txtST01_Validate(Index As Integer, Cancel As Boolean)
   If txtST01(Index).Text <> "" Then
      If ChkStaffID(txtST01(Index)) Then
         Call txtST01_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
   End If
   If Index = 0 Then
      If txtST01(Index) <> "" And txtST01(Index + 1) = "" Then
         txtST01(Index + 1) = txtST01(Index)
      End If
      If txtST01(Index) > txtST01(Index + 1) Then
         txtST01(Index + 1) = txtST01(Index)
      End If
   ElseIf Index = 1 Then
      If txtST01(Index) <> "" And txtST01(Index - 1) = "" Then
         txtST01(Index - 1) = txtST01(Index)
      End If
      If txtST01(Index - 1) <> "" And txtST01(Index) <> "" Then
         If RunNick(txtST01(Index - 1), txtST01(Index)) Then
            Call txtST01_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtDept_GotFocus(Index As Integer)
   InverseTextBox txtDept(Index)
End Sub

Private Sub txtDept_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDept_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      If txtDept(Index) <> "" And txtDept(Index + 1) = "" Then
         txtDept(Index + 1) = txtDept(Index)
      End If
      If txtDept(Index) > txtDept(Index + 1) Then
         txtDept(Index + 1) = txtDept(Index)
      End If
   ElseIf Index = 1 Then
      If txtDept(Index) <> "" And txtDept(Index - 1) = "" Then
         txtDept(Index - 1) = txtDept(Index)
      End If
      If txtDept(Index - 1) <> "" And txtDept(Index) <> "" Then
         If RunNick(txtDept(Index - 1), txtDept(Index)) Then
            Call txtDept_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   InverseTextBox txtDate(Index)
End Sub

Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
   If txtDate(Index).Text <> "" Then
      If ChkDate(txtDate(Index)) = False Then
         Call txtDate_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
   End If
   If Index = 0 Then
      If txtDate(Index) <> "" And txtDate(Index + 1) = "" Then
         txtDate(Index + 1) = txtDate(Index)
      End If
      If Val(txtDate(Index)) > Val(txtDate(Index + 1)) Then
         txtDate(Index + 1) = txtDate(Index)
      End If
   ElseIf Index = 1 Then
      If txtDate(Index) <> "" And txtDate(Index - 1) = "" Then
         txtDate(Index - 1) = txtDate(Index)
      End If
      If txtDate(Index - 1) <> "" And txtDate(Index) <> "" Then
         If RunNick2(txtDate(Index - 1), txtDate(Index)) Then
            Call txtDate_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub
