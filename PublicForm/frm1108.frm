VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm1108 
   BorderStyle     =   1  '單線固定
   Caption         =   "部門別送件清單列印"
   ClientHeight    =   3912
   ClientLeft      =   456
   ClientTop       =   996
   ClientWidth     =   4776
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3912
   ScaleWidth      =   4776
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1215
      TabIndex        =   25
      Top             =   510
      Width           =   3405
   End
   Begin VB.Frame Frame1 
      Caption         =   "將指定案號改為下午送件"
      Height          =   735
      Left            =   45
      TabIndex        =   22
      Top             =   3090
      Width           =   4695
      Begin VB.TextBox txtCP 
         Height          =   285
         Index           =   1
         Left            =   1170
         MaxLength       =   3
         TabIndex        =   5
         Top             =   270
         Width           =   465
      End
      Begin VB.TextBox txtCP 
         Height          =   285
         Index           =   2
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   6
         Top             =   270
         Width           =   765
      End
      Begin VB.TextBox txtCP 
         Height          =   285
         Index           =   3
         Left            =   2550
         MaxLength       =   1
         TabIndex        =   7
         Top             =   270
         Width           =   225
      End
      Begin VB.TextBox txtCP 
         Height          =   285
         Index           =   4
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   8
         Top             =   270
         Width           =   345
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "存檔(&S)"
         Height          =   345
         Left            =   3330
         TabIndex        =   9
         Top             =   240
         Width           =   1230
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1470
         X2              =   3180
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "本所案號："
         Height          =   255
         Index           =   8
         Left            =   135
         TabIndex        =   23
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4095
      Top             =   1560
   End
   Begin VB.TextBox txtListTime 
      Enabled         =   0   'False
      Height          =   264
      Left            =   1215
      MaxLength       =   6
      TabIndex        =   4
      Top             =   2280
      Width           =   1230
   End
   Begin VB.ComboBox cboListType 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frm1108.frx":0000
      Left            =   1215
      List            =   "frm1108.frx":000A
      TabIndex        =   2
      Top             =   1530
      Width           =   2625
   End
   Begin VB.ComboBox cboListTime 
      Height          =   300
      ItemData        =   "frm1108.frx":001A
      Left            =   1215
      List            =   "frm1108.frx":001C
      TabIndex        =   3
      Top             =   1890
      Width           =   1230
   End
   Begin VB.TextBox txtPrinter 
      Enabled         =   0   'False
      Height          =   264
      Left            =   1215
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1200
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   3735
      TabIndex        =   11
      Top             =   30
      Width           =   912
   End
   Begin VB.TextBox txtCP27 
      Height          =   264
      Left            =   1215
      MaxLength       =   7
      TabIndex        =   0
      Top             =   870
      Width           =   1185
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2745
      TabIndex        =   10
      Top             =   30
      Width           =   912
   End
   Begin VB.Label Label2 
      Caption         =   "非智慧局清單一案印一張！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   180
      TabIndex        =   24
      Top             =   60
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   135
      X2              =   4635
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   135
      X2              =   4635
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblServerTime 
      Height          =   180
      Left            =   1260
      TabIndex        =   21
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "系統時間："
      Height          =   180
      Index           =   7
      Left            =   210
      TabIndex        =   20
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "( 格式：HHMMSS )"
      Height          =   180
      Index           =   4
      Left            =   2610
      TabIndex        =   19
      Top             =   2325
      Width           =   1605
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "分段時間："
      Height          =   180
      Index           =   6
      Left            =   210
      TabIndex        =   18
      Top             =   2325
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Index           =   5
      Left            =   390
      TabIndex        =   17
      Top             =   570
      Width           =   720
   End
   Begin MSForms.Label lblPrinter 
      Height          =   255
      Left            =   2475
      TabIndex        =   16
      Top             =   1215
      Width           =   1245
      VariousPropertyBits=   27
      Size            =   "2196;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "部門別："
      Height          =   180
      Index           =   2
      Left            =   390
      TabIndex        =   15
      Top             =   1590
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "送件時段："
      Height          =   180
      Index           =   3
      Left            =   210
      TabIndex        =   14
      Top             =   1950
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "列印人員："
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   13
      Top             =   1245
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "發文日期："
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   12
      Top             =   915
      Width           =   900
   End
End
Attribute VB_Name = "frm1108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/12 改成Form2.0 (lblPrinter,Printer列印未改)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'sonia 2010/8/19 日期欄已修改
'整理 by Morgan 2005/7/25
'Modify by Morgan 2008/3/14 簡化改下午送件功能
Option Explicit

Dim m_CP(1 To 4) As String
Dim m_OriPrinterName As String '原印表機名稱
Dim m_DBTime As Long '系統時間
Dim m_Dept As String '部門別
Dim m_intCnt As Integer 'Add By Sindy 2010/12/21 查詢出幾筆資料
Dim strPrinter As String
Dim prnPrint As Printer

Private Sub cboListTime_Click()
   GetTime
End Sub

Private Sub cmdExit_Click(Index As Integer)
   Unload Me
   Set frm1108 = Nothing
End Sub

Private Function CheckCaseNo() As Boolean
   Select Case m_Dept
      Case "P1"
         If Trim(txtCP(1)) <> "P" Then
            MsgBox "系統代碼輸入錯誤！", vbCritical
            txtCP(1).SetFocus
            txtCP_GotFocus 1
            Exit Function
         End If
         
      Case "P2"
         If Trim(txtCP(1)) <> "T" And Trim(txtCP(1)) <> "FCT" Then
            MsgBox "系統代碼輸入錯誤！", vbCritical
            txtCP(1).SetFocus
            txtCP_GotFocus 1
            Exit Function
         End If
      
      Case "F1"
         If Trim(txtCP(1)) <> "FCT" Then
            MsgBox "系統代碼輸入錯誤！", vbCritical
            txtCP(1).SetFocus
            txtCP_GotFocus 1
            Exit Function
         End If
            
      Case "F2"
         If Trim(txtCP(1)) <> "FCP" Then
            MsgBox "系統代碼輸入錯誤！", vbCritical
            txtCP(1).SetFocus
            txtCP_GotFocus 1
            Exit Function
         End If
   End Select
         
   If Trim(txtCP(2).Text) = "" Then
      MsgBox "案號不可空白！", vbExclamation
      txtCP(2).SetFocus
      txtCP_GotFocus 2
      Exit Function
   End If
         
   m_CP(1) = Trim(txtCP(1).Text)
   m_CP(2) = Right("000000" & Trim(txtCP(2).Text), 6)
   m_CP(3) = Left(Trim(txtCP(3).Text) & "0", 1)
   m_CP(4) = Left(Trim(txtCP(4).Text) & "00", 2)
   '檢查案號是否存在
   strExc(0) = "SELECT cp82 from caseprogress" & _
      " WHERE cp01='" & m_CP(1) & "' AND cp02='" & m_CP(2) & "' AND cp03='" & m_CP(3) & "' AND cp04='" & m_CP(4) & "'" & _
      " and cp27=" & strSrvDate(1) & " and cp82<=120000"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If MsgBox("是否確定將 " & m_CP(1) & "-" & m_CP(2) & "-" & m_CP(3) & "-" & m_CP(4) & " 改為下午送件？", vbYesNo + vbDefaultButton2) = vbYes Then
         CheckCaseNo = True
      End If
   Else
      MsgBox "案號 " & m_CP(1) & "-" & m_CP(2) & "-" & m_CP(3) & "-" & m_CP(4) & " 當日上午並無發文資料！", vbExclamation
      txtCP(2).SetFocus
      txtCP_GotFocus 2
   End If
End Function

Private Sub cmdMove_Click()
   If CheckCaseNo = True Then
      If UpdateCaseTime() = True Then
         txtCP(2).Text = "": txtCP(3).Text = "": txtCP(4).Text = ""
         txtCP_GotFocus 2
         txtCP(2).SetFocus
      End If
   End If
End Sub

'更改案件送件時間
Private Function UpdateCaseTime() As Boolean
   '檢查是否已開支票
   If isChecked() = True Then
      MsgBox "財務處已開立支票,清單內容不可變更!", vbExclamation
      Exit Function
   End If
   
   Dim strCon As String, i As Integer
   
On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
      
   strSql = "UPDATE CASEPROGRESS SET CP82=130000" & _
      " WHERE CP01='" & m_CP(1) & "' AND CP02='" & m_CP(2) & "' AND CP03='" & m_CP(3) & "' AND CP04='" & m_CP(4) & "'" & _
      " AND CP27=" & strSrvDate(1) & " AND CP09<'C'"
      
   cnnConnection.Execute strSql
      
   strSql = "DELETE APPLISTDETAIL" & _
      " WHERE ALD01=" & strSrvDate(1) & _
      " AND ALD05='" & m_CP(1) & "' AND ALD06='" & m_CP(2) & "' AND ALD07='" & m_CP(3) & "' AND ALD08='" & m_CP(4) & "'"
         
   cnnConnection.Execute strSql
      
   cnnConnection.CommitTrans
   
   UpdateCaseTime = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
End Function
'Add by Morgan 2006/1/3
'檢查當日清單是否已有支票號
'Modify by Morgan 2008/9/24
'檢查當日清單是否有無支票號的
Private Function isChecked(Optional p_bAll As Boolean = True) As Boolean
   
   isChecked = True
   
   strExc(0) = "SELECT AL06 FROM APPLIST" & _
      " WHERE AL01=" & TransDate(txtCP27, 2) & " AND AL02='" & m_Dept & "'"
   
   If p_bAll = False Then
      strExc(0) = strExc(0) & " and AL03='" & cboListTime.ListIndex & "'"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         If IsNull(.Fields(0)) Then
            isChecked = False
            Exit Function
         End If
         .MoveNext
      Loop
      End With
   Else
      isChecked = False
   End If
End Function

Private Function Process() As Boolean
   Dim strDate As String, arrList
   Dim strConCP As String, strInsSQL As String, iEffect As Integer
  
   strDate = DBDATE(txtCP27)
   strConCP = " and cp27=" & strDate
   '上午
   If cboListTime.ListIndex = 0 Then
      strConCP = strConCP & " and cp82<" & Format(txtListTime.Text)
   '下午
   Else
      strConCP = strConCP & " and cp82>=" & Format(txtListTime.Text)
   End If
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   
   '刪除舊資料
   strSql = " BEGIN" & _
      " DELETE FROM APPLISTDETAIL" & _
      " WHERE ALD01=" & strDate & " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "';" & _
      " DELETE FROM APPLIST" & _
      " WHERE AL01=" & strDate & " AND AL02='" & m_Dept & "' AND AL03='" & cboListTime.ListIndex & "' AND AL06 IS NULL;" & _
      " END;"
      
   cnnConnection.Execute strSql, iEffect
   
   '新增清單明細
   strSql = GetSql
   cnnConnection.Execute strSql, iEffect
   
'Removed by Morgan 2016/10/3 無此需求,取消--黃志佑
'   '判斷單一專利種類是否超過15個案件
'   If m_Dept = "P1" Or m_Dept = "F2" Then
'      strSql = "SELECT CP10 FROM APPLISTDETAIL,CASEPROGRESS A" & _
'         " WHERE ALD01=" & strDate & " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "'" & _
'         " AND CP01=ALD05 AND CP02=ALD06 AND CP03=ALD07 AND CP04=ALD08" & _
'         " AND CP10 IN('101','102','103')" & strConCP & _
'         " GROUP BY CP10 HAVING COUNT(*)>15"
'
'      intI = 1
'      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         With adoRecordset
'         Do While Not .EOF
'            Select Case .Fields(0)
'               Case "101"
'                  strSql = "UPDATE APPLISTDETAIL SET ALD04='4'" & _
'                     " WHERE ALD01=" & strDate & " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "'" & _
'                     " AND EXISTS(SELECT * FROM CASEPROGRESS A WHERE CP01=ALD05 AND CP02=ALD06 AND CP03=ALD07 AND CP04=ALD08 AND CP10='101'" & strConCP & ")"
'               Case "102"
'                  strSql = "UPDATE APPLISTDETAIL SET ALD04='5'" & _
'                     " WHERE ALD01=" & strDate & " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "'" & _
'                     " AND EXISTS(SELECT * FROM CASEPROGRESS A WHERE CP01=ALD05 AND CP02=ALD06 AND CP03=ALD07 AND CP04=ALD08 AND CP10='102'" & strConCP & ")"
'               Case "103"
'                  strSql = "UPDATE APPLISTDETAIL SET ALD04='6'" & _
'                     " WHERE ALD01=" & strDate & " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "'" & _
'                     " AND EXISTS(SELECT * FROM CASEPROGRESS A WHERE CP01=ALD05 AND CP02=ALD06 AND CP03=ALD07 AND CP04=ALD08 AND CP10='103'" & strConCP & ")"
'            End Select
'            cnnConnection.Execute strSql, iEffect
'            .MoveNext
'         Loop
'         End With
'      End If
'   End If
'end 2016/10/3
   
   '新增清單主檔
   strSql = " INSERT INTO APPLIST(AL01,AL02,AL03,AL04,AL05,AL07,AL08,AL09)" & _
      " SELECT ALD01,ALD02,ALD03,ALD04," & txtListTime & ",'" & strUserNum & "',TO_CHAR(SYSDATE,'YYYYMMDD'),TO_CHAR(SYSDATE,'HH24MI')" & _
      " FROM APPLISTDETAIL,APPLIST WHERE ALD01=" & strDate & " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "'" & _
      " AND AL01(+)=ALD01 AND AL02(+)=ALD02 AND AL03(+)=ALD03 AND AL04(+)=ALD04 AND AL06 IS NULL GROUP BY ALD01,ALD02,ALD03,ALD04"
      
   cnnConnection.Execute strSql, iEffect
   
   'Modify by Morgan 2008/9/24 考慮只有部分部分清單重跑
   'If iEffect = 0 Then
   '若無明細時新增一筆
   strExc(0) = "select * from applist where al01=" & strDate & " and al02='" & m_Dept & "' and al03='" & cboListTime.ListIndex & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
   'end 2008/9/24
      strSql = " INSERT INTO APPLIST(AL01,AL02,AL03,AL04,AL05,AL07,AL08,AL09)" & _
         " VALUES(" & strDate & ",'" & m_Dept & "','" & cboListTime.ListIndex & "','2'," & txtListTime & ",'" & strUserNum & "',TO_CHAR(SYSDATE,'YYYYMMDD'),TO_CHAR(SYSDATE,'HH24MI'))"
      cnnConnection.Execute strSql, iEffect
   End If
   
   cnnConnection.CommitTrans
   Process = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Function Process2() As Boolean
    Dim strDate As String, arrList
   Dim strConCP As String, strInsSQL As String, iEffect As Integer
  
   strDate = DBDATE(txtCP27)
   strConCP = " and cp27=" & strDate
   '上午
   If cboListTime.ListIndex = 0 Then
      strConCP = strConCP & " and cp82<" & Format(txtListTime.Text)
   '下午
   Else
      strConCP = strConCP & " and cp82>=" & Format(txtListTime.Text)
   End If
   
On Error GoTo ErrHnd

   cnnConnection.BeginTrans
   
   '刪除舊資料
   strSql = " DELETE FROM RecAppList" & _
      " WHERE RAL01=" & strDate & " AND RAL02='" & m_Dept & "' AND RAL03='" & cboListTime.ListIndex & "'"
   cnnConnection.Execute strSql, iEffect
   
   'Add by Morgan 2009/4/30
   '新增發文對象清單
   'Modified by Morgan 2023/8/31 排除TC案,因目前為承辦繳費且時間不定(有規費清單目前也沒有印)
   strSql = "insert into RecAppList(RAL01,RAL02,RAL03,RAL04,RAL05,RAL06,RAL07)" & _
      " select ald01,ald02,ald03,ald04,cp09 as RAL05" & _
      ",decode(cp130,null,'經濟部智慧財產局',decode(instr(cp130,','),0,cp130,substr(cp130,1,instr(cp130,',')-1))) as RAL06,cp84 as RAL07" & _
      " From applist, applistdetail, caseprogress" & _
      " where al01=" & strDate & " and al02='" & m_Dept & "' and al03='" & cboListTime.ListIndex & "'" & _
      " and ald01(+)=al01 and ald02(+)=al02 and ald03(+)=al03 and ald04(+)=al04" & _
      " and cp01(+)=ald05 and cp02(+)=ald06 and cp03(+)=ald07 and cp04(+)=ald08" & _
      " and cp27(+)=ald01 and cp84>0" & _
      " UNION SELECT " & strDate & " as RAL01,'" & m_Dept & "' as RAL02" & _
      ",'" & cboListTime.ListIndex & "' as RAL03,'9' as RAL04,CP09 as RAL05" & _
      ",decode(cp130,null,'經濟部智慧財產局',decode(instr(cp130,','),0,cp130,substr(cp130,1,instr(cp130,',')-1))) as RAL06,cp84 as RAL07" & _
      " FROM CASEPROGRESS,STAFF WHERE cp01<>'TC' and cp09<'C' AND NVL(CP84,0)=0 and cp123='Y'" & strConCP & _
      " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "'"
      
   cnnConnection.Execute strSql, iEffect
   
   '新增其他發文對象
   If iEffect > 0 Then
      strExc(0) = "select A.*,CP130 from RecAppList A,caseprogress where RAL01=" & strDate & " and RAL02='" & m_Dept & "' and RAL03='" & cboListTime.ListIndex & "' and cp09(+)=RAL05 and instr(CP130,',')>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            arrList = Split("" & .Fields("CP130"), ",")
            For intI = 1 To UBound(arrList)
               strSql = "insert into RecAppList(RAL01,RAL02,RAL03,RAL04,RAL05,RAL06)" & _
                  " VALUES(" & .Fields("RAL01") & ",'" & .Fields("RAL02") & "'" & _
                  ",'" & .Fields("RAL03") & "','9'" & _
                  ",'" & .Fields("RAL05") & "','" & arrList(intI) & "')"
               cnnConnection.Execute strSql, iEffect
            Next
            .MoveNext
         Loop
         End With
      End If
   End If
   'end 2009/4/30
   
   cnnConnection.CommitTrans
   Process2 = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Sub cmdPrint_Click()
   Dim strDesc As String
   
   If Trim(txtPrinter) = "" Then
      MsgBox "請輸入列印人員!!"
      txtPrinter.SetFocus
      Exit Sub
   End If
   
   If Trim(txtCP27) = "" Then
      MsgBox "請輸入發文日期!!"
      If txtCP27.Enabled Then txtCP27.SetFocus
      Exit Sub
   End If
   
   If Trim(txtListTime) = "" Then
      MsgBox "請輸入分段時間!!"
      If txtListTime.Enabled Then txtListTime.SetFocus
      Exit Sub
   End If

   Screen.MousePointer = vbHourglass
   If cboListType.ListIndex >= 0 Then
      '該時段未開支票才會重新產生清單資料
      If isChecked(False) = False Then
         If Process() = False Then
            MsgBox "無法產生清單資料！", vbCritical
            Exit Sub
         End If
      Else
         If MsgBox("財務處支票皆已開立，將不會重新建立有規費的清單明細！是否確定要列印？" & vbCrLf & vbCrLf & "( ※若案件有更動，請與財務處確認　貴部門該時段支票已作廢後再行列印！ )", vbOKCancel + vbQuestion + vbDefaultButton2, "重印確認") = vbCancel Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      End If
      
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/21 清除查詢印表記錄檔欄位
      If Trim(txtCP27) <> "" Then
         pub_QL05 = pub_QL05 & ";" & Label1(0) & txtCP27 'Add By Sindy 2010/12/21
      End If
      If Trim(txtPrinter) <> "" Then
         pub_QL05 = pub_QL05 & ";" & Label1(1) & txtPrinter & lblPrinter 'Add By Sindy 2010/12/21
      End If
      If Trim(cboListType.Text) <> "" Then
         pub_QL05 = pub_QL05 & ";" & Label1(2) & cboListType.Text 'Add By Sindy 2010/12/21
      End If
      If Trim(cboListTime.Text) <> "" Then
         pub_QL05 = pub_QL05 & ";" & Label1(3) & cboListTime.Text 'Add By Sindy 2010/12/21
      End If
      If Trim(txtListTime) <> "" Then
         pub_QL05 = pub_QL05 & ";" & Label1(6) & txtListTime 'Add By Sindy 2010/12/21
      End If
      m_intCnt = 0
      
      'Modify by Morgan 2009/5/6
      'If DoPrint1 = True Then
      'Modify by Morgan 2010/4/7 無規費清單無支票問題一律重新產生資料
      If Process2 Then
         PUB_RestorePrinter Combo1 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
         
         If DoPrint = True Then
            strDesc = "有【有規費】案件，清單列印完畢!!"
         Else
            strDesc = "無【有規費】案件!!"
         End If
      
         If DoPrint2 = True Then
            strDesc = strDesc & vbCrLf & vbCrLf & "有【無規費】案件，清單列印完畢!!"
         Else
            strDesc = strDesc & vbCrLf & vbCrLf & "無【無規費】案件!!"
         End If
         
         PUB_RestorePrinter strPrinter 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
      End If
      
      InsertQueryLog (m_intCnt) 'Add By Sindy 2010/12/21
      MsgBox strDesc, vbInformation
   Else
      MsgBox "該部門無清單可列印！", vbInformation
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub PrintHead(p_iCaseType As Integer, p_iNameType As Integer, Optional p_Office As String)

      Dim strTitle As String
      Dim lngY As Long
      
      Select Case p_iCaseType
         Case 1, 4, 5, 6 '新案
            strTitle = cboListType.Text & " " & ChangeTStringToTDateString(txtCP27.Text) & " " & cboListTime & "新申請案送件清單"
            Select Case p_iCaseType
               Case 4
                  strTitle = strTitle & "(發明)"
               Case 5
                  strTitle = strTitle & "(新型)"
               Case 6
                  strTitle = strTitle & "(設計)"
            End Select
         Case 2 '一般
            strTitle = cboListType.Text & " " & ChangeTStringToTDateString(txtCP27.Text) & " " & cboListTime & "送件清單"
            
         Case 3 '快速
            strTitle = cboListType.Text & " " & ChangeTStringToTDateString(txtCP27.Text) & " " & cboListTime & "快速收文櫃檯送件清單"
      
         Case 7 '非智慧局
            If p_Office <> "" Then
               strTitle = cboListType.Text & " " & ChangeTStringToTDateString(txtCP27.Text) & " " & cboListTime & "現金/" & p_Office & "送件清單"
            Else
               strTitle = cboListType.Text & " " & ChangeTStringToTDateString(txtCP27.Text) & " " & cboListTime & "現金/非智慧局送件清單"
            End If
            
         Case 8 '電子送件
            strTitle = cboListType.Text & " " & ChangeTStringToTDateString(txtCP27.Text) & " " & cboListTime & "電子送件清單"
      End Select
      '是否出名
      If p_iNameType = 1 Then strTitle = strTitle & "(不出名)"
      
      Printer.Print
      Printer.FontSize = 16
      Printer.CurrentX = 5000
      Printer.Print strTitle
      Printer.FontSize = 12
      Printer.CurrentX = 13000
      Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
      Printer.CurrentX = 0
      lngY = Printer.CurrentY
      Printer.Print "區分時段時間：" & Format(txtListTime.Text, "##:##:##")
      Printer.CurrentY = lngY
      Printer.CurrentX = 13000
      Printer.Print "列印時間：" & lblServerTime
      Printer.CurrentX = 13000
      Printer.Print "列印人員：" & lblPrinter
      
      Printer.Print
      If m_Dept = "P2" Or m_Dept = "F1" Then
         Printer.Print "本所案號        規費      規費小計 申請案號   案件性質　　 申請人               案件名稱                                 商品類別      "
         Printer.Print "--------------- --------- -------- ---------- ------------ -------------------- ---------------------------------------- --------------"
     'Added by Lydia 2016/04/28 外專+eE化,特殊請款
      ElseIf m_Dept = "F2" Then
         Printer.Print "本所案號        ｅ 特殊 規費      規費小計 申請案號   案件性質　　 申請人               案件名稱                                "
         Printer.Print "--------------- -- ---- --------- -------- ---------- ------------ -------------------- ----------------------------------------"
         
      Else
         Printer.Print "本所案號        規費      規費小計 申請案號   案件性質　　 申請人               案件名稱                                "
         Printer.Print "--------------- --------- -------- ---------- ------------ -------------------- ----------------------------------------"
      End If
End Sub

Private Sub PrintHead1(Optional strType As String, Optional strName As String)

      Dim strTitle As String
      Dim lngY As Long
      If strName <> "" Then
            strTitle = cboListType.Text & " " & ChangeTStringToTDateString(txtCP27.Text) & " " & cboListTime & "無規費送件清單(" & strName & ")"
      Else
         If strType = "1" Then
            strTitle = cboListType.Text & " " & ChangeTStringToTDateString(txtCP27.Text) & " " & cboListTime & "無規費送件清單(智慧局)"
         Else
            strTitle = cboListType.Text & " " & ChangeTStringToTDateString(txtCP27.Text) & " " & cboListTime & "無規費送件清單(非智慧局)"
         End If
      End If
      
      Printer.Print
      Printer.FontSize = 16
      Printer.CurrentX = 5000
      Printer.Print strTitle
      Printer.FontSize = 12
      Printer.CurrentX = 13000
      Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
      Printer.CurrentX = 0
      lngY = Printer.CurrentY
      Printer.Print "區分時段時間：" & Format(txtListTime.Text, "##:##:##")
      Printer.CurrentY = lngY
      Printer.CurrentX = 13000
      Printer.Print "列印時間：" & lblServerTime
      Printer.CurrentX = 13000
      Printer.Print "列印人員：" & lblPrinter
      
      Printer.Print
      Printer.Print "本所案號        申請案號     案件性質　　 案件備註"
      Printer.Print "--------------- ------------ ------------ ---------------------------------------------------------------------------------------------"
End Sub

Private Sub PrintTail(iPage As Integer, Optional p_lngTot As Long, Optional p_iRecs As Integer, Optional p_iCaseCnt As Integer, Optional p_PS As String)
   Dim stData As String
   
   If p_lngTot > 0 Then
      '內商不印案件數
      If m_Dept = "P2" Then
         stData = "總計" & Space(4) & Right(Space(8) & "規費 " & Format(p_lngTot, "###,###,###") & " 元", 16) & Space(4) & Right(Space(2) & "明細 " & Format(p_iRecs) & " 筆", 9) & Space(4)
      Else
         stData = "總計" & Space(4) & Right(Space(8) & "規費 " & Format(p_lngTot, "###,###,###") & " 元", 16) & Space(4) & Right(Space(2) & "明細 " & Format(p_iRecs) & " 筆", 9) & Space(4) & Right(Space(2) & "案號 " & Format(p_iCaseCnt) & " 筆", 9)
      End If
   End If
               
   Printer.FontSize = 12

   If m_Dept = "P2" Or m_Dept = "F1" Then
      Printer.Print String(135, "-")
   'Added by Lydia 2016/04/28 +外專F2
   ElseIf m_Dept = "F2" Then
      Printer.Print String(128, "-")
   'end 2016/04/28
   Else
      Printer.Print String(120, "-")
   End If
   Printer.Print stData
   
   'Added by Morgan 2021/4/27
   If p_PS <> "" Then
      Printer.Print
      Printer.Print "PS:" & p_PS
   End If
   'end 2021/4/27
   
   Printer.CurrentY = 10700
   Printer.CurrentX = 7000
   Printer.Print "第 " & Format(iPage) & " 頁"
   
End Sub

Private Sub PrintTail1(iPage As Integer, Optional p_iRecs As Integer)
   Dim stData As String
   
   Printer.FontSize = 12
   Printer.Print String(135, "-")
   If p_iRecs > 0 Then
      Printer.Print "共 " & p_iRecs & " 筆"
   End If
   
   Printer.CurrentY = 10700
   Printer.CurrentX = 7000
   Printer.Print "第 " & Format(iPage) & " 頁"
   
End Sub

Private Function GetSql() As String

   Dim strInsSQL As String, strCon As String, stVTable As String
   Dim strNewAppList As String
  
   strCon = " and A.cp27=" & TransDate(txtCP27, 2)
   '上午
   If cboListTime.ListIndex = 0 Then
      strCon = strCon & " and A.cp82<" & Format(txtListTime.Text)
   '下午
   Else
      strCon = strCon & " and A.cp82>=" & Format(txtListTime.Text)
   End If
   
   Select Case m_Dept
      'Memo by Morgan 2009/6/19
      '若同一案件同一時段送兩個不同櫃檯案件時需手動修改資料(原控制同一案件申請書會合併故應為單一櫃檯送件)
      '將原種類金額減少並新增另一種類的送件資料 Ex.P-89881,實審與申請有先權證明同時送件(儘量分不同時段送件以避免此問題)
      
      '專利分上下午,出名否,新舊案 8 種清單
      Case "P1", "F2"
         'Modify by Morgan 2005/9/29 加改請301~303
         'Modified by Morgan 2013/1/11 +125,308
         strNewAppList = "101,102,103,104,105,117,125,301,302,303,305,306,307,308"
         
         '條件:發文規費>0,申請國家=000,發文人員部門前二碼
         '快速收文櫃檯案件性質(申請優先權證明405,補發證書604)單獨列印 --> P.S.單獨送件才要
         '取Min控制若有收新案則1,有收一般則2,單獨收快速才3
         '快速發文案件也要考慮沒有規費的程序
         'Modify by Morgan 2005/8/16 補發證書先不要
         'stVTable = " SELECT CP01, CP02, CP03, CP04, MIN(DECODE(INSTR('101,102,103,104,105,117,305,306,307',CP10),0,DECODE(INSTR('405,604',CP10),0,'2','3'),'1')) ALD04,SUM(CP84) ALD09,MAX(CP22) ALD10"
         'Modify by Morgan 2007/2/9 補發證書又要--玲玲
         'stVTable = " SELECT CP01, CP02, CP03, CP04, MIN(DECODE(INSTR('101,102,103,104,105,117,301,302,303,305,306,307',CP10),0,DECODE(INSTR('405',CP10),0,'2','3'),'1')) ALD04,SUM(CP84) ALD09,MAX(CP22) ALD10"
         'Modify by Morgan 2008/1/29 FCP都不要快速送件 --靜芳
         'stVTable = " SELECT CP01, CP02, CP03, CP04, MIN(DECODE(INSTR('101,102,103,104,105,117,301,302,303,305,306,307',CP10),0,DECODE(INSTR('405,604',CP10),0,'2','3'),'1')) ALD04,SUM(CP84) ALD09,MAX(CP22) ALD10"
         'Modify by Morgan 2008/3/17 FCP 不管進度檔是否出名都算出名(第三人繳年費)
         'stVTable = " SELECT CP01, CP02, CP03, CP04, MIN(DECODE(INSTR('101,102,103,104,105,117,301,302,303,305,306,307',CP10),0,DECODE(INSTR('P405,P604',CP01||CP10),0,'2','3'),'1')) ALD04,SUM(CP84) ALD09,MAX(CP22) ALD10"
         'Modified by Morgan 2012/12/28 +FCP405,FCP604也快速
         stVTable = " SELECT CP01, CP02, CP03, CP04, MIN(DECODE(INSTR('" & strNewAppList & "',CP10),0,DECODE(INSTR('P405,P604,FCP405,FCP604',CP01||CP10),0,'2','3'),'1')) ALD04,SUM(CP84) ALD09,MAX(DECODE(CP01,'FCP',NULL,CP22)) ALD10"
         'End 2007/2/9
         stVTable = stVTable & " FROM CASEPROGRESS A,PATENT,STAFF"
         stVTable = stVTable & " where cp01='" & IIf(m_Dept = "P1", "P", "FCP") & "'" & strCon
         stVTable = stVTable & " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='000'"
         stVTable = stVTable & " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "'"
         'Add by Morgan 2008/7/10 加判斷電子送件
         'Mofieied by Morgan 2015/12/17 +判斷有經發文室的否則是否出名會錯誤(Ex.電話聯絡單發文沒有出名人會導致該案視為不出名)
         stVTable = stVTable & " and cp118 is null and cp123 is not null"
                  
         
         'Add by Morgan 2007/8/28 加判斷主管機關
         'Modify by Morgan 2009/4/28 改判斷CP130的第一個主管機關
         'stVTable = stVTable & " and NOT EXISTS(SELECT * FROM CASEPROGRESS X,CASEPROGRESS B,NEXTPROGRESS,CASEFEE" & _
            " WHERE X.CP09=A.CP09 AND B.CP09(+)=DECODE(X.CP10,'404',X.CP43) AND NP01(+)=DECODE(X.CP10,'404',X.CP43)" & _
            " AND CF01=PA01 AND CF02=PA09 AND CF03=DECODE(X.CP10,'404',DECODE(SUBSTR(X.CP09,1,1),'A',NP07,B.CP10),X.CP10)" & _
            " AND CF10<>'經濟部智慧財產局' )"
         stVTable = stVTable & " and (cp130 is null or decode(instr(cp130,','),0,cp130,substr(cp130,1,instr(cp130,',')-1))='經濟部智慧財產局')"
         'end 2009/4/28
         'end 2007/8/28
         
         stVTable = stVTable & " GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
         
         'Add by Morgan 2007/8/28 加判斷主管機關,非智慧局=7 (FCP的面詢費用由工程師去的時候繳所以主管機關有故意加空白以便清單單獨出;P案不用,因為該部門工程師認為先繳費審查委員比較會准面詢--敏惠)
         stVTable = stVTable & " UNION SELECT CP01, CP02, CP03, CP04,'7' ALD04,SUM(CP84) ALD09,MAX(CP22) ALD10"
         stVTable = stVTable & " FROM CASEPROGRESS A,PATENT,STAFF"
         stVTable = stVTable & " where cp01='" & IIf(m_Dept = "P1", "P", "FCP") & "'" & strCon
         stVTable = stVTable & " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='000'"
         stVTable = stVTable & " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "'"
         stVTable = stVTable & " and cp118 is null and cp123 is not null"
         
         'Modify by Morgan 2009/4/28
         'Modify by Morgan 2009/4/28 改判斷CP130的第一個主管機關
         'stVTable = stVTable & " and EXISTS(SELECT * FROM CASEPROGRESS X,CASEPROGRESS B,NEXTPROGRESS,CASEFEE" & _
            " WHERE X.CP09=A.CP09 AND B.CP09(+)=DECODE(X.CP10,'404',X.CP43) AND NP01(+)=DECODE(X.CP10,'404',X.CP43)" & _
            " AND CF01=PA01 AND CF02=PA09 AND CF03=DECODE(X.CP10,'404',DECODE(SUBSTR(X.CP09,1,1),'A',NP07,B.CP10),X.CP10)" & _
            " AND CF10<>'經濟部智慧財產局' )"
         stVTable = stVTable & " and decode(instr(cp130,','),0,cp130,substr(cp130,1,instr(cp130,',')-1))<>'經濟部智慧財產局'"
         'end 2009/4/28
         stVTable = stVTable & " GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
         'end 2007/8/28
         
'Remove by Morgan 2011/12/5 已改單獨功能無須再產生資料(Table也不同)
'         'Add by Morgan 2008/7/10 加電子送件=8
'         stVTable = stVTable & " UNION SELECT CP01, CP02, CP03, CP04,'8' ALD04,SUM(CP84) ALD09,MAX(CP22) ALD10"
'         stVTable = stVTable & " FROM CASEPROGRESS A,PATENT,STAFF"
'         stVTable = stVTable & " where cp01='" & IIf(m_Dept = "P1", "P", "FCP") & "'" & strCon
'         stVTable = stVTable & " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='000'"
'         stVTable = stVTable & " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "'"
'         stVTable = stVTable & " and cp118='Y'"
'         stVTable = stVTable & " GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
'         'end 2008/7/10
         
         If m_Dept = "F2" Then
            'Modified by Morgan 2025/1/22 因為有非智慧局案件，增加判斷主管機關
            stVTable = stVTable & " UNION ALL"
            stVTable = stVTable & " select CP01, CP02, CP03, CP04, '2' ALD04,SUM(CP84) ALD09,MAX(CP22) ALD10"
            stVTable = stVTable & " FROM CASEPROGRESS A, servicepractice, staff" & _
               " where cp01='FG'" & strCon & _
               " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp09='000'" & _
               " and (cp130 is null or decode(instr(cp130,','),0,cp130,substr(cp130,1,instr(cp130,',')-1))='經濟部智慧財產局')" & _
               " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "'" & _
               " GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
               
            
            stVTable = stVTable & " UNION ALL"
            stVTable = stVTable & " select CP01, CP02, CP03, CP04, '7' ALD04,SUM(CP84) ALD09,MAX(CP22) ALD10"
            stVTable = stVTable & " FROM CASEPROGRESS A, servicepractice, staff" & _
               " where cp01='FG'" & strCon & _
               " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp09='000'" & _
               " and (cp130 is null or decode(instr(cp130,','),0,cp130,substr(cp130,1,instr(cp130,',')-1))<>'經濟部智慧財產局')" & _
               " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "'" & _
               " GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
               
         End If
         
         strInsSQL = "INSERT INTO APPLISTDETAIL(ALD01,ALD02,ALD03,ALD04,ALD05,ALD06,ALD07,ALD08,ALD09,ALD10)" & _
            " SELECT " & TransDate(txtCP27, 2) & ",'" & m_Dept & "','" & cboListTime.ListIndex & "',ALD04, CP01,CP02,CP03,CP04,ALD09,ALD10" & _
            " FROM  (" & stVTable & ") X"
   
      '商標 不分新舊案
      Case "P2", "F1"
         strInsSQL = "INSERT INTO APPLISTDETAIL(ALD01,ALD02,ALD03,ALD04,ALD05,ALD06,ALD07,ALD08,ALD09,ALD10)" & _
            " SELECT " & TransDate(txtCP27, 2) & ",'" & m_Dept & "','" & cboListTime.ListIndex & "','2', CP01,CP02,CP03,CP04,SUM(CP84),MAX(CP22)" & _
            " FROM CASEPROGRESS A, trademark, staff" & _
            " where cp01 " & IIf(m_Dept = "P2", "IN ('T','FCT')", "='FCT'") & " and cp84>0" & strCon & _
            " and TM01(+)=cp01 and TM02(+)=cp02 and TM03(+)=cp03 and TM04(+)=cp04 and TM10='000'" & _
            " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "'"
         'Add by Morgan 2008/7/10 加判斷電子送件
         'Mofieied by Morgan 2015/12/17 +判斷有經發文室的否則是否出名會錯誤(Ex.電話聯絡單發文沒有出名人會導致該案視為不出名)
         strInsSQL = strInsSQL & " and cp118 is null and cp123 is not null"
         
         'Add by Morgan 2007/8/28 加判斷主管機關
         'Modify by Morgan 2009/4/28 改判斷CP130的第一個主管機關
         'strInsSQL = strInsSQL & " and NOT EXISTS(SELECT * FROM CASEPROGRESS X,CASEPROGRESS B,NEXTPROGRESS,CASEFEE" & _
            " WHERE X.CP09=A.CP09 AND B.CP09(+)=DECODE(X.CP10,'303',X.CP43) AND NP01(+)=DECODE(X.CP10,'303',X.CP43)" & _
            " AND CF01=TM01 AND CF02=TM10 AND CF03=DECODE(X.CP10,'303',DECODE(SUBSTR(X.CP09,1,1),'A',NP07,B.CP10),X.CP10)" & _
            " AND CF10<>'經濟部智慧財產局')"
         strInsSQL = strInsSQL & " and (cp130 is null or decode(instr(cp130,','),0,cp130,substr(cp130,1,instr(cp130,',')-1))='經濟部智慧財產局')"
         'end 2009/4/29
         'end 2007/8/28
         strInsSQL = strInsSQL & " GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
         
         strInsSQL = strInsSQL & " UNION SELECT " & TransDate(txtCP27, 2) & ",'" & m_Dept & "','" & cboListTime.ListIndex & "','7', CP01,CP02,CP03,CP04,SUM(CP84),MAX(CP22)" & _
            " FROM CASEPROGRESS A, trademark, staff" & _
            " where cp01 " & IIf(m_Dept = "P2", "IN ('T','FCT')", "='FCT'") & " and cp84>0" & strCon & _
            " and TM01(+)=cp01 and TM02(+)=cp02 and TM03(+)=cp03 and TM04(+)=cp04 and TM10='000'" & _
            " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "' and cp118 is null and cp123 is not null"
         'Add by Morgan 2007/8/28 加判斷主管機關
         'Modify by Morgan 2009/4/28 改判斷CP130的第一個主管機關
         'strInsSQL = strInsSQL & " and EXISTS(SELECT * FROM CASEPROGRESS X,CASEPROGRESS B,NEXTPROGRESS,CASEFEE" & _
            " WHERE X.CP09=A.CP09 AND B.CP09(+)=DECODE(X.CP10,'303',X.CP43) AND NP01(+)=DECODE(X.CP10,'303',X.CP43)" & _
            " AND CF01=TM01 AND CF02=TM10 AND CF03=DECODE(X.CP10,'303',DECODE(SUBSTR(X.CP09,1,1),'A',NP07,B.CP10),X.CP10)" & _
            " AND CF10<>'經濟部智慧財產局' )"
         strInsSQL = strInsSQL & " and decode(instr(cp130,','),0,cp130,substr(cp130,1,instr(cp130,',')-1))<>'經濟部智慧財產局'"
         
         strInsSQL = strInsSQL & " GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
         'end 2007/8/28
         
'Remove by Morgan 2011/12/5 已改單獨功能無須再產生資料(Table也不同)
'         'Add by Morgan 2008/7/10 加電子送件=8
'         strInsSQL = strInsSQL & " UNION SELECT " & TransDate(txtCP27, 2) & ",'" & m_Dept & "','" & cboListTime.ListIndex & "','8', CP01,CP02,CP03,CP04,SUM(CP84),MAX(CP22)" & _
'            " FROM CASEPROGRESS A, trademark, staff" & _
'            " where cp01 " & IIf(m_Dept = "P2", "IN ('T','FCT')", "='FCT'") & " and cp84>0" & strCon & _
'            " and TM01(+)=cp01 and TM02(+)=cp02 and TM03(+)=cp03 and TM04(+)=cp04 and TM10='000'" & _
'            " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "' and cp118='Y'"
'         strInsSQL = strInsSQL & " GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
'         'end 2008/7/10
            
         'Add by Morgan 2006/8/4 CFT的英文證明也要印
         If m_Dept = "F1" Then
            strInsSQL = strInsSQL & " UNION ALL" & _
               " SELECT " & TransDate(txtCP27, 2) & ",'" & m_Dept & "','" & cboListTime.ListIndex & "','2', CP01,CP02,CP03,CP04,SUM(CP84),MAX(CP22)" & _
               " FROM CASEPROGRESS A, staff" & _
               " where cp01='CFT' and cp10='304' and cp30 is not null and cp84>0" & strCon & _
               " and st01(+)=cp83 and substrb(st03,1,2)='" & m_Dept & "' and cp118 is null and cp123 is not null GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
         End If
         
      'Add by Morgan 2008/9/23
      '目前僅中所專利新案
      Case Else
         stVTable = " SELECT CP01, CP02, CP03, CP04, '1' ALD04,SUM(CP84) ALD09,MAX(CP22) ALD10"
         stVTable = stVTable & " FROM CASEPROGRESS A,PATENT,STAFF"
         'Mofieied by Morgan 2015/12/17 +判斷有經發文室的否則是否出名會錯誤(Ex.電話聯絡單發文沒有出名人會導致該案視為不出名)
         stVTable = stVTable & " where cp118 is null and cp123 is not null" & strCon
         stVTable = stVTable & " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='000'"
         stVTable = stVTable & " and st01(+)=cp83 and st06='" & m_Dept & "'"
         stVTable = stVTable & " GROUP BY CP01,CP02,CP03,CP04 HAVING SUM(CP84)>0"
         
         strInsSQL = "INSERT INTO APPLISTDETAIL(ALD01,ALD02,ALD03,ALD04,ALD05,ALD06,ALD07,ALD08,ALD09,ALD10)" & _
            " SELECT " & TransDate(txtCP27, 2) & ",'" & m_Dept & "','" & cboListTime.ListIndex & "',ALD04, CP01,CP02,CP03,CP04,ALD09,ALD10" & _
            " FROM  (" & stVTable & ") X"
   End Select
   
   GetSql = strInsSQL
      
End Function
'Add by Morgan 2009/3/18 無規費清單
Private Function DoPrint1() As Boolean

   Dim strConAL As String, strConCP As String
   Dim nCopys As Integer '份數
   Dim iPage As Integer '頁次
   Dim strTmp As String, iRec As Integer, iRecs As Integer
   Dim iCopys As Integer
   Dim stKey As String
   
   strConAL = " and al01=" & DBDATE(txtCP27)
   strConCP = ""
   '上午
   If cboListTime.ListIndex = 0 Then
      strConAL = strConAL & " and al03='0'"
      strConCP = strConCP & " and cp82<al05"
   '下午
   Else
      strConAL = strConAL & " and al03='1'"
      strConCP = strConCP & " and cp82>=al05"
   End If
   
   strSql = ""
   
   Select Case m_Dept
      '專利
      Case "P1", "F2"
         nCopys = 1
         strSql = "select RPAD(cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04),15,' ') 本所案號" & _
            ",RPAD(CPM03,12,' ') 案件性質,cp64 案件備註,PA01 CF1,PA09 CF2,CP10 CF3" & _
            " from (SELECT AL01,AL02,AL03,AL05 FROM applist WHERE al02='" & m_Dept & "'" & strConAL & " and rownum<2) X" & _
            ",caseprogress,staff,casepropertymap,patent" & _
            " where cp27(+)=al01 and cp09<'C' AND NVL(CP84,0)=0" & _
            " and cp01 in ('P','FCP') and cp123='Y'" & strConCP & _
            " and st01(+)=cp83 and substr(st03,1,2)='" & m_Dept & "'" & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='000'"

         strSql = strSql & " UNION ALL select RPAD(cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04),15,' ') 本所案號" & _
            ",RPAD(CPM03,12,' ') 案件性質,cp64 案件備註,SP01 CF1,SP09 CF2,CP10 CF3" & _
            " from (SELECT AL01,AL02,AL03,AL05 FROM applist WHERE al02='" & m_Dept & "'" & strConAL & " and rownum<2) X" & _
            ",caseprogress,staff,casepropertymap,servicepractice" & _
            " where cp27(+)=al01 and cp09<'C' AND NVL(CP84,0)=0" & _
            " and cp01 in ('PS', 'FG') and cp123='Y'" & strConCP & _
            " and st01(+)=cp83 and substr(st03,1,2)='" & m_Dept & "'" & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
            " and Sp01(+)=cp01 and Sp02(+)=cp02 and Sp03(+)=cp03 and Sp04(+)=cp04 and Sp09='000'"
            
      '商標
      Case "P2", "F1"
         nCopys = 1

         strSql = "select RPAD(cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04),15,' ') 本所案號" & _
            ",RPAD(CPM03,12,' ') 案件性質,cp64 案件備註,TM01 CF1,TM10 CF2,CP10 CF3" & _
            " from (SELECT AL01,AL02,AL03,AL05 FROM applist WHERE al02='" & m_Dept & "'" & strConAL & " and rownum<2) X" & _
            ",caseprogress,staff,casepropertymap,trademark" & _
            " where cp27(+)=al01 and cp09<'C' AND NVL(CP84,0)=0" & _
            " and cp01 in ('T','FCT') and cp123='Y'" & strConCP & _
            " and st01(+)=cp83 and substr(st03,1,2)='" & m_Dept & "'" & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
            " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm10='000'"

         strSql = strSql & " UNION ALL select RPAD(cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04),15,' ') 本所案號" & _
            ",RPAD(CPM03,12,' ') 案件性質,cp64 案件備註,SP01 CF1,SP09 CF2,CP10 CF3" & _
            " from (SELECT AL01,AL02,AL03,AL05 FROM applist WHERE al02='" & m_Dept & "'" & strConAL & " and rownum<2) X" & _
            ",caseprogress,staff,casepropertymap,servicepractice" & _
            " where cp27(+)=al01 and cp09<'C' AND NVL(CP84,0)=0" & _
            " and (cp01='S' or SUBSTR(CP01,1)='T') and cp123='Y'" & strConCP & _
            " and st01(+)=cp83 and substr(st03,1,2)='" & m_Dept & "'" & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
            " and SP01(+)=cp01 and SP02(+)=cp02 and SP03(+)=cp03 and SP04(+)=cp04 and SP09='000'"
            
   End Select

   If strSql = "" Then
      Exit Function
   Else
      strSql = "SELECT X.*,DECODE(CF10,'經濟部智慧財產局',1,2) SRT FROM (" & strSql & ") X,CASEFEE WHERE CF01(+)=CF1 AND CF02(+)=CF2 AND CF03(+)=CF3 ORDER BY SRT,1,2"
   End If

On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Printer.Orientation = 2
         Printer.Font = "細明體"
         For iCopys = 1 To nCopys
            .MoveFirst
            stKey = "" & .Fields("SRT")
            If iCopys > 1 Then Printer.NewPage
            iPage = 1: iRec = 0: iRecs = 0
            PrintHead1 stKey
            Do While Not .EOF
               If stKey <> "" & .Fields("SRT") Then
                  PrintTail1 iPage, iRecs
                  iRecs = 0
                  iRec = 0
                  iPage = 1
                  stKey = "" & .Fields("SRT")
                  Printer.NewPage
                  PrintHead1 stKey
               End If
               iRecs = iRecs + 1
               iRec = iRec + 1
               If iRec > 26 Then
                  PrintTail1 iPage
                  Printer.NewPage
                  iPage = iPage + 1
                  PrintHead1 stKey
                  iRec = 0
               End If
               strTmp = .Fields(0) & Space(1) & .Fields(1) & Space(1) & .Fields(2)
               Printer.CurrentY = Printer.CurrentY + 60
               Printer.Print strTmp
               .MoveNext
            Loop
            PrintTail1 iPage, iRecs
         Next
         Printer.EndDoc
         DoPrint1 = True
      End If
   End With

ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function
'Modify by Morgan 2011/2/24 電子送件不印
'Modified by Morgan 2012/8/6 規費扣除已銷帳
Private Function DoPrint() As Boolean

   Dim strTmp As String, iRec As Integer, strTitle As String, iRecs As Integer, i As Integer
   Dim lngTot As Long, strCon As String, strCon2 As String, stVTable As String, stLastNo As String, stLastC13 As String
   Dim iCaseNo As Integer '案件筆數
   Dim iPage As Integer '頁次
   Dim nCopys As Integer '份數
   Dim iCopys As Integer
   Dim iCaseType As Integer '案件類別 1.新案 2.一般 3.快速
   Dim iNameType As Integer '出名類別 0.出名 1.不出名
   Dim strOffice As String '發文機關
   Dim stVTable2 As String, stVTable3 As String
   Dim m_bolEmail As Boolean, m_bolPlusPaper As Boolean 'Added by Lydia 2016/04/28
   Dim strPS As String, arrNo(4) As String 'Added by Morgan 2021/4/27
   
   strCon = " and A.cp27=" & TransDate(txtCP27, 2)
   '上午
   If cboListTime.ListIndex = 0 Then
      strCon = strCon & " and A.cp82<" & Format(txtListTime.Text)
      'Added by Morgan 2019/4/26
      If m_Dept = "P1" Or m_Dept = "F2" Then
         strCon2 = " and c.cp82(+)<" & Format(txtListTime.Text)
      End If
      'end 2019/4/26
   '下午
   Else
      strCon = strCon & " and A.cp82>=" & Format(txtListTime.Text)
      'Added by Morgan 2019/4/26
      If m_Dept = "P1" Or m_Dept = "F2" Then
          strCon2 = " and c.cp82(+)>=" & Format(txtListTime.Text)
      End If
      'end 2019/4/26
   End If
   
   strSql = ""
   'Modified by Morgan 2016/11/14 +判斷有經發文室(要與GetSql語法一致)
   '上下午,出名否不同清單
   Select Case m_Dept
      
      '專利 清單種類:新案,一般,快速;新案單一專利種類超過*清單單獨印
      '排序:是否出名(C09),清單種類(C10),主管機關(C12),發文規費加總(C03),本所案號(C01),發文規費(C02),案件性質(C05)
      Case "P1", "F2"
         nCopys = 2
         
         stVTable = " SELECT ALD05 X01, ALD06 X02, ALD07 X03, ALD08 X04, PA05, PA11, PA26, ALD09 X05, ALD04 X06, ALD10,ALD01,ALD02,ALD03,ALD04 " & _
            " FROM APPLISTDETAIL A,PATENT" & _
            " WHERE ALD01=" & TransDate(txtCP27, 2) & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04<>'8'" & _
            " AND PA01(+)=ALD05 AND PA02(+)=ALD06 AND PA03(+)=ALD07 AND PA04(+)=ALD08"
         
         stVTable2 = "select cp09 T2C1,sum(a1u09) T2C2" & _
            " from applistdetail,caseprogress A,acc1u0" & _
            " WHERE ALD01=" & TransDate(txtCP27, 2) & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04<>'8'" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & " and A.cp84>0 and A.cp118 is null and A.cp123 is not null" & strCon & _
            " AND A1U03(+)=CP09 GROUP BY CP09"
         
         stVTable3 = "select CP43 T3C1,sum(a1u09) T3C2" & _
            " from applistdetail,caseprogress A,acc1u0" & _
            " WHERE ALD01=" & TransDate(txtCP27, 2) & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04<>'8'" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & " and A.cp84>0 and A.cp118 is null and A.cp123 is not null" & strCon & _
            " AND CP10='404' AND CP09>'B' AND CP43 IS NOT NULL AND A1U03(+)=CP43 GROUP BY CP43"

         '控制快速收文櫃檯案件性質單獨列印 申請優先權證明405,補發證書604
         'Modify by Morgan 2009/6/19 加以發文對象清單資料過濾收文號(快速送件相同案號有可能人工改為不同清單列印,Ex.P-89881)
         'Modified by Morgan 2012/3/29 延期改用cp30判斷是否已收文
         'Modified by Lydia 2016/04/28 +領證及繳年費(601)分成另一區塊(C13)
         'Modified by Morgan 2019/4/26 領證規費不可與其他性質小計--黃志佑
         strSql = "SELECT LPAD(A.CP01||'-'||A.CP02||'-'||A.CP03||'-'||A.CP04,15,' ') C01,DECODE(A.CP10,'601','1','0') C13,A.CP148 C14, A.CP84 C02" & _
            ",DECODE(A.CP10,'601',A.CP84,X05-NVL(C.CP84,0)) C03,LPAD(NVL(PA11,' '),10,' ') C04, RPAD(CPM03,12,' ') C05, RPAD(NVL(CU04,' '),20,' ') C06" & _
            ",RPAD(PA05,40,' ') C07, RPAD(' ',10,' ') C08, DECODE(ALD10,'N',1,0) C09, X06 C10" & _
            ",DECODE( NVL(DECODE(A.CP10,'404',DECODE(SUBSTR(A.CP09,1,1),'A',A.CP17-NVL(T2C2,0),B.CP17-NVL(T3C2,0)),A.CP17-NVL(T2C2,0)),0)-A.CP84,0,' ','*') C11" & _
            ",decode(A.cp130,null,'經濟部智慧財產局',decode(instr(A.cp130,','),0,A.cp130,substr(A.cp130,1,instr(A.cp130,',')-1))) C12 " & _
            " FROM  (" & stVTable & ") X, CASEPROGRESS A, CUSTOMER" & _
            ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP,(" & stVTable2 & ") T2,(" & stVTable3 & ") T3,CASEPROGRESS C" & _
            " where A.CP01=X01 AND A.CP02=X02 AND A.CP03=X03 AND A.CP04=X04" & " and A.cp84>0 and A.cp118 is null and A.cp123 is not null" & strCon & _
            " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
            " AND B.CP09(+)=DECODE(A.CP10,'404',A.CP43,NULL) AND NP01(+)=DECODE(A.CP10,'404',A.CP43,NULL) AND NP22(+)=DECODE(A.CP10,'404',A.CP30,NULL)" & _
            " AND CPM01=A.CP01 AND CPM02=DECODE(A.CP10,'404',NVL(NP07,B.CP10),A.CP10)" & _
            " AND EXISTS( SELECT * FROM RecAppList WHERE RAL01=ALD01 AND RAL02=ALD02 AND RAL03=ALD03 AND RAL04=ALD04 AND RAL05=A.CP09 AND RAL07>0 )" & _
            " AND T2C1(+)=A.CP09 AND T3C1(+)=B.CP09" & _
            " AND C.CP01(+)=A.CP01 AND C.CP02(+)=A.CP02 AND C.CP03(+)=A.CP03 AND C.CP04(+)=A.CP04 and C.CP10(+)='601' AND C.CP27(+)=A.CP27" & strCon2
            
      '商標 清單種類:一般,非智慧局案件
      '排序:是否出名(C09),發文規費加總(C03),本所案號(C01),發文規費(C02),案件性質(C05)
      Case "P2", "F1"
      
         nCopys = 2
         
         stVTable = " SELECT ALD05 X01, ALD06 X02, ALD07 X03, ALD08 X04, TM05, TM09, TM12, TM15, TM23, ALD09 X05, ALD04 X06,ALD10 " & _
            " FROM APPLISTDETAIL,TRADEMARK A" & _
            " WHERE ALD01=" & TransDate(txtCP27, 2) & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04<>'8'" & _
            " AND A.TM01(+)=ALD05 AND A.TM02(+)=ALD06 AND A.TM03(+)=ALD07 AND A.TM04(+)=ALD08"
         
         stVTable2 = "select cp09 T2C1,sum(a1u09) T2C2" & _
            " from applistdetail,caseprogress A,acc1u0" & _
            " WHERE ALD01=" & TransDate(txtCP27, 2) & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04<>'8'" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & " and A.cp84>0 and A.cp118 is null and A.cp123 is not null" & strCon & _
            " AND A1U03(+)=CP09 GROUP BY CP09"
         
         stVTable3 = "select CP43 T3C1,sum(a1u09) T3C2" & _
            " from applistdetail,caseprogress A,acc1u0" & _
            " WHERE ALD01=" & TransDate(txtCP27, 2) & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04<>'8'" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & " and A.cp84>0 and A.cp118 is null and A.cp123 is not null" & strCon & _
            " AND CP10='303' AND CP09>'B' AND CP43 IS NOT NULL AND A1U03(+)=CP43 GROUP BY CP43"

         'Modify by Morgan 2006/8/4 加控制CFT的申請號抓CP30(國內案的審定號)
         '若性質為延期303時:A類收文用CP43抓NP07，規費用延期的CP17比較；B類收文用CP43抓相關收文號的CP10，規費用相關收文號的CP17比較。
         'Modified by Lydia 2016/04/28 +C13,C14
         strSql = "SELECT LPAD(A.CP01||'-'||A.CP02||'-'||A.CP03||'-'||A.CP04,15,' ') C01,'0' C13,' ' C14 , A.CP84 C02, X05 C03" & _
            ", LPAD(NVL(DECODE(A.CP01,'CFT',A.CP30,NVL(TM15,TM12)),' '),10,' ') C04, RPAD(CPM03,12,' ') C05, RPAD(CU04,20,' ') C06" & _
            ", RPAD(TM05,40,' ') C07, RPAD(TM09,10,' ') C08, DECODE(ALD10,'N',1,0) C09,X06 C10" & _
            ", DECODE( NVL(DECODE(A.CP10,'303',DECODE(SUBSTR(A.CP09,1,1),'A',A.CP17-NVL(T2C2,0),B.CP17-NVL(T3C2,0)),A.CP17-NVL(T2C2,0)),0)-A.CP84,0,' ','*') C11" & _
            ",decode(A.cp130,null,'經濟部智慧財產局',decode(instr(A.cp130,','),0,A.cp130,substr(A.cp130,1,instr(A.cp130,',')-1))) C12 " & _
            " FROM  (" & stVTable & ") X, CASEPROGRESS A,  CUSTOMER" & _
            ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP,(" & stVTable2 & ") T2,(" & stVTable3 & ") T3" & _
            " where A.CP01=X01 AND A.CP02=X02 AND A.CP03=X03 AND A.CP04=X04" & " and A.cp84>0 and A.cp118 is null and A.cp123 is not null" & strCon & _
            " and cu01(+)=substr(TM23,1,8) and cu02(+)=substr(TM23,9,1)" & _
            " AND B.CP09(+)=DECODE(A.CP10,'303',A.CP43,NULL) AND NP01(+)=DECODE(A.CP10,'303',A.CP43,NULL) AND NP22(+)=DECODE(A.CP10,'303',A.CP30,NULL)" & _
            " AND CPM01=A.CP01 AND CPM02=DECODE(A.CP10,'303',NVL(NP07,B.CP10),A.CP10) " & _
            " AND T2C1(+)=A.CP09 AND T3C1(+)=B.CP09"

      'Add by Morgan 2008/9/23
      Case Else
         nCopys = 1
         
         stVTable = " SELECT ALD05 X01, ALD06 X02, ALD07 X03, ALD08 X04, PA05, PA11, PA26, ALD09 X05, ALD04 X06, ALD10 " & _
            " FROM APPLISTDETAIL A,PATENT" & _
            " WHERE ALD01=" & TransDate(txtCP27, 2) & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04<>'8'" & _
            " AND PA01(+)=ALD05 AND PA02(+)=ALD06 AND PA03(+)=ALD07 AND PA04(+)=ALD08"
            
         stVTable2 = "select cp09 T2C1,sum(a1u09) T2C2" & _
            " from applistdetail,caseprogress A,acc1u0" & _
            " WHERE ALD01=" & TransDate(txtCP27, 2) & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04<>'8'" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & " and A.cp84>0 and A.cp118 is null and A.cp123 is not null" & strCon & _
            " AND A1U03(+)=CP09 GROUP BY CP09"
         
         stVTable3 = "select CP43 T3C1,sum(a1u09) T3C2" & _
            " from applistdetail,caseprogress A,acc1u0" & _
            " WHERE ALD01=" & TransDate(txtCP27, 2) & _
            " AND ALD02='" & m_Dept & "' AND ALD03='" & cboListTime.ListIndex & "' AND ALD04<>'8'" & _
            " AND CP01(+)=ALD05 AND CP02(+)=ALD06 AND CP03(+)=ALD07 AND CP04(+)=ALD08" & " and A.cp84>0 and A.cp118 is null and A.cp123 is not null" & strCon & _
            " AND CP10='404' AND CP09>'B' AND CP43 IS NOT NULL AND A1U03(+)=CP43 GROUP BY CP43"
         'Modified by Lydia 2016/04/28 +C13,C14
         strSql = "SELECT LPAD(A.CP01||'-'||A.CP02||'-'||A.CP03||'-'||A.CP04,15,' ') C01,'0' C13,' ' C14, A.CP84 C02, X05 C03" & _
            ",LPAD(NVL(PA11,' '),10,' ') C04, RPAD(CPM03,12,' ') C05, RPAD(NVL(CU04,' '),20,' ') C06" & _
            ",RPAD(PA05,40,' ') C07, RPAD(' ',10,' ') C08, DECODE(ALD10,'N',1,0) C09, X06 C10" & _
            ",DECODE( NVL(DECODE(A.CP10,'404',DECODE(SUBSTR(A.CP09,1,1),'A',A.CP17-NVL(T2C2,0),B.CP17-NVL(T3C2,0)),A.CP17-NVL(T2C2,0)),0)-A.CP84,0,' ','*') C11" & _
            ",decode(A.cp130,null,'經濟部智慧財產局',decode(instr(A.cp130,','),0,A.cp130,substr(A.cp130,1,instr(A.cp130,',')-1))) C12 " & _
            " FROM  (" & stVTable & ") X, CASEPROGRESS A, CUSTOMER" & _
            ",CASEPROGRESS B, NEXTPROGRESS,CASEPROPERTYMAP,(" & stVTable2 & ") T2,(" & stVTable3 & ") T3" & _
            " where A.CP01=X01 AND A.CP02=X02 AND A.CP03=X03 AND A.CP04=X04" & " and A.cp84>0 and A.cp118 is null and A.cp123 is not null" & strCon & _
            " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
            " AND B.CP09(+)=DECODE(A.CP10,'404',A.CP43,NULL) AND NP01(+)=DECODE(A.CP10,'404',A.CP43,NULL) AND NP22(+)=DECODE(A.CP10,'404',A.CP30,NULL)" & _
            " AND CPM01=A.CP01 AND CPM02=DECODE(A.CP10,'404',NVL(NP07,B.CP10),A.CP10)" & _
            " AND T2C1(+)=A.CP09 AND T3C1(+)=B.CP09"
      
   End Select
   '排序: 出名否(C09=ALD10), 案件類別(C10=ALD04), 主管機關名稱(C12), 規費(C03=ALD09) DESC,本所案號(C01),發文規費(C02=CP84) DESC,案件性質(C05)
   'Modified by Lydia 2016/04/28 領證及繳年費(601)分成另一區塊(C13)
   'strSql = strSql & " order by 9, 10,12, 3 DESC, 1, 2 DESC, 5"
   If m_Dept = "F2" Then
      strSql = strSql & " order by C09, C10, C13, C12, C03 DESC, C01, C02 DESC, C05"
   Else
      strSql = strSql & " order by C09, C10, C12, C03 DESC, C01, C02 DESC, C05"
   End If
On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         m_intCnt = m_intCnt + .RecordCount 'Add By Sindy 2010/12/21
         Printer.Orientation = 2
         Printer.Font = "細明體"
         For iCopys = 1 To nCopys
            strPS = "" 'Added by Morgan 2021/4/27
            .MoveFirst
            If iCopys > 1 Then Printer.NewPage
         
            iPage = 1: iRec = 0: lngTot = 0: iCaseNo = 0: iRecs = 0: stLastNo = ""
            iCaseType = .Fields("C10"): iNameType = .Fields("C09")
            strOffice = .Fields("C12")
            PrintHead iCaseType, iNameType, strOffice
            
            Do While Not .EOF
            
               '案件類別或出名類別不同時跳頁
               'Modify by Morgan 2007/8/28 非智慧局案件一案一頁
               'If .Fields("C09") <> iNameType Or .Fields("C10") <> iCaseType Then
               If .Fields("C09") <> iNameType Or .Fields("C10") <> iCaseType Or (iCaseType = 7 And iRec > 0) Then
                  PrintTail iPage, lngTot, iRecs, iCaseNo
                  Printer.NewPage
                  iPage = 1: iRec = 0: lngTot = 0: iCaseNo = 0: iRecs = 0
                  iCaseType = .Fields("C10"): iNameType = .Fields("C09")
                  strOffice = .Fields("C12")
                  PrintHead iCaseType, iNameType, strOffice
               End If
               iRec = iRec + 1: iRecs = iRecs + 1
               If iRec > 26 Then
                  PrintTail iPage
                  Printer.NewPage
                  iPage = iPage + 1
                  PrintHead iCaseType, iNameType, strOffice
                  iRec = 0
               End If
               strTmp = ""
               
               'Modified by Lydia 2016/04/28 處理明細列
               'For i = 0 To 7
               For i = 0 To 9
                  '規費
                  'If i = 1 Then
                  If i = 3 Then
                     strTmp = strTmp & Right(Space(8) & Format(Val("" & .Fields(i)), "###,###"), 8) & .Fields("C11") & Space(1)
                     lngTot = lngTot + Val("" & .Fields(i))
                     
                  '規費小計
                  'ElseIf i = 2 Then
                  ElseIf i = 4 Then
                     '內商不印小計
                     'If .Fields(0) <> stLastNo Then
                     'Added by Morgan 2019/4/26
                     If (m_Dept = "P1" Or m_Dept = "F2") Then
                        '領證規費不與其他性質小計(小計=規費)
                        If .Fields("C13") = "1" Then
                           strTmp = strTmp & Right(Space(8) & Format(Val("" & .Fields(i)), "###,###"), 8) & Space(1)
                        ElseIf .Fields(0) <> stLastNo Or stLastC13 = "1" Then
                           strTmp = strTmp & Right(Space(8) & Format(Val("" & .Fields(i)), "###,###"), 8) & Space(1)
                        Else
                           strTmp = strTmp & Space(9)
                        End If
                     'end 2019/4/26
                     ElseIf m_Dept <> "P2" And .Fields(0) <> stLastNo Then
                        strTmp = strTmp & Right(Space(8) & Format(Val("" & .Fields(i)), "###,###"), 8) & Space(1)
                     Else
                        strTmp = strTmp & Space(9)
                     End If
                  '案件性質
                  'ElseIf i = 4 Then
                  ElseIf i = 6 Then
                     strTmp = strTmp & .Fields(i) & Space(1)
                  'Added by Lydia 2016/04/28
                  ElseIf i = 1 Or i = 2 Then
                     If m_Dept = "F2" Then
                        If i = 1 Then
                           m_bolEmail = PUB_GetEMailFlag(Trim(Replace(.Fields("C01"), "-", "")), True, , m_bolPlusPaper)
                           strExc(2) = IIf(m_bolEmail, IIf(m_bolPlusPaper, "Ｅ", "ｅ"), "  ")
                           strTmp = strTmp & strExc(2) & Space(1)
                        Else
                           strTmp = strTmp & Mid(.Fields(i) & Space(5), 1, 5)
                        End If
                     End If
                  'end 2016/02/28
                  Else
                     '內商同案號也要印
                     If m_Dept = "P2" Or .Fields(0) <> stLastNo Then
                        strTmp = strTmp & .Fields(i) & Space(1)
                     Else
                        strTmp = strTmp & Space(Len(.Fields(i)) + 1)
                     End If
                  End If
               Next
               'end 2016/04/28
               Printer.CurrentY = Printer.CurrentY + 60
               Printer.Print strTmp
               If .Fields(0) <> stLastNo Then iCaseNo = iCaseNo + 1
               stLastNo = .Fields(0)
               stLastC13 = .Fields("C13") 'Added by Morgan 2019/4/26
               
               'Added by Morgan 2021/4/27
               If Left(Trim(.Fields("C01")), 3) = "CFT" Then
                  Erase arrNo
                  ChgCaseNo Replace(Trim(.Fields("C01")), "-", ""), arrNo
                  If PUB_GetReceiptComp(arrNo(1), arrNo(2), arrNo(3), arrNo(4)) = "J" Then
                     strPS = strPS & .Fields("C01") & " "
                  End If
               End If
               'end 2021/4/27
               .MoveNext
            Loop
            
            If strPS <> "" Then strPS = strPS & "為智權公司出名" 'Added by Morgan 2021/4/27
            
            PrintTail iPage, lngTot, iRecs, iCaseNo, strPS
         Next
         
         Printer.EndDoc
         DoPrint = True
      End If
   End With


ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   CheckOC

End Function

Private Sub Form_Load()
 
   MoveFormToCenter Me
   
   '更新系統時間
   m_DBTime = ServerTime
   time = Format(m_DBTime, "##:##:##")
   lblServerTime.Caption = Format(m_DBTime, "##:##:##")
   Timer1.Interval = 1000
   
   '印表機
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
      
   '發文日期
   txtCP27.Text = strSrvDate(2)
   
   '列印人員
   txtPrinter.Text = strUserNum
   lblPrinter.Caption = GetStaffName(strUserNum)
   
   '清單種類
   cboListType.Clear
   cboListType.AddItem "內專"
   cboListType.AddItem "內商"
   cboListType.AddItem "外商"
   cboListType.AddItem "外專"
   'Add by Morgan 2008/9/23
   cboListType.AddItem "中所"
   cboListType.AddItem "南所"
   cboListType.AddItem "高所"
   
   '送件時段
   cboListTime.Clear
   cboListTime.AddItem "上午"
   cboListTime.AddItem "下午"
   '12點前預設上午
   If m_DBTime < 120000 Then
      cboListTime.ListIndex = 0
   Else
      cboListTime.ListIndex = 1
   End If
   
   SetRef
   
   If Pub_StrUserSt03 = "M51" Then
      txtPrinter.Enabled = True
   End If
   
End Sub

Private Sub SetRef()
   Dim stOffice As String
   
   stOffice = PUB_GetST06(txtPrinter)
   
   '部門別
   If stOffice = "1" Then
      m_Dept = Left(GetStaffDepartment(txtPrinter), 2)
   Else
      m_Dept = stOffice
   End If
   
   Select Case m_Dept
      Case "P1": cboListType.ListIndex = 0
      Case "P2": cboListType.ListIndex = 1
      Case "F1": cboListType.ListIndex = 2
      Case "F2": cboListType.ListIndex = 3
      Case "2": cboListType.ListIndex = 4
      Case "3": cboListType.ListIndex = 5
      Case "4": cboListType.ListIndex = 6
      Case Else: cboListType.ListIndex = -1
   End Select
   GetTime
End Sub

Private Sub Form_Unload(Cancel As Integer)
       '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    MenuEnabled
   Set frm1108 = Nothing
End Sub

Private Sub Timer1_Timer()
   m_DBTime = Format(Now, "HHMMSS")
   lblServerTime.Caption = Format(m_DBTime, "##:##:##")
End Sub

Private Sub txtCP_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCP27_GotFocus()
   TextInverse txtCP27
End Sub

Private Sub txtCP27_KeyPress(KeyAscii As Integer)
   '只能輸數字
   If Not (KeyAscii = vbKeyBack Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCP27_Validate(Cancel As Boolean)
   If ChkDate(txtCP27.Text) = False Then
      Cancel = True
   Else
      '非當日不可改送件時間
      If txtCP27.Text <> strSrvDate(2) Then
         txtCP(1).Text = ""
         txtCP(2).Text = ""
         txtCP(3).Text = ""
         txtCP(4).Text = ""
         Frame1.Enabled = False
      Else
         Frame1.Enabled = True
      End If
      GetTime
   End If
End Sub

Private Sub txtListTime_GotFocus()
   TextInverse txtListTime
End Sub

Private Sub txtListTime_KeyPress(KeyAscii As Integer)
   '只能輸數字
   If Not (KeyAscii = vbKeyBack Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9"))) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCP_GotFocus(Index As Integer)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtCP(Index).IMEMode = 2
   CloseIme
   TextInverse txtCP(Index)
End Sub

Private Function GetTime() As Boolean

On Error GoTo ErrHnd
   
   If cboListType.ListIndex >= 0 Then
      txtListTime.Enabled = True
      strSql = "select AL05,AL06 from APPLIST" & _
         " where AL01=" & TransDate(txtCP27, 2) & " and AL02='" & m_Dept & "' ORDER BY AL06"
      
      CheckOC
      With adoRecordset
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
         '已列印
         If .RecordCount > 0 Then
            txtListTime.Text = "" & .Fields("AL05")
            ' 若財務處當日已產生傳票分錄則分段時間不可再改
            If Not IsNull(.Fields(1)) Then
               txtListTime.Enabled = False
            End If
         '未列印
         Else
            '當日 10:30以前預設系統時間 否則預設10:30
            If Val(txtCP27.Text) = Val(strSrvDate(2)) Then
               txtListTime.Text = IIf(m_DBTime >= 103000, 103000, m_DBTime)
            '非當日 預設10:30
            Else
               '不必預設
               'txtListTime.Text = 103000
            End If
         End If
      End With
      'Add by Morgan 2005/8/31
      '下午都不可改分隔時間以免發生上下午不同
      If cboListTime.ListIndex = 1 Then
         txtListTime.Enabled = False
      End If
   End If
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description
   
End Function

Private Sub txtListTime_Validate(Cancel As Boolean)
   If Val(txtListTime.Text) > 120000 Then
      MsgBox "區分時段時間不可過12點！", vbCritical
      Cancel = True
   End If
   'Added by Lydia 2025/10/16 人工輸入1030沒有秒，造成無法產生applistdetail
   If Len(Trim(txtListTime.Text)) <> 6 Then
      MsgBox "請輸入時分秒(HHMMSS)！" & vbCrLf & "例如:103000", vbCritical
      Cancel = True
   End If
   'end 2025/10/16
End Sub

Private Sub txtPrinter_GotFocus()
   TextInverse txtPrinter
End Sub

'Add By Sindy 2010/11/29
Private Sub txtPrinter_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPrinter_Validate(Cancel As Boolean)
   Dim strTempName As String
   If txtPrinter <> "" Then
      If ClsPDGetStaff(txtPrinter, strTempName) = True Then
         lblPrinter = strTempName
         SetRef
      Else
         Cancel = True
      End If
   End If
End Sub

'Add by Morgan 2009/4/30 無規費清單
Private Function DoPrint2() As Boolean
   Dim strConRAL As String
   Dim nCopys As Integer '份數
   Dim iPage As Integer '頁次
   Dim strTmp As String, iRec As Integer, iRecs As Integer
   Dim iCopys As Integer
   Dim stKey As String
   
   strConRAL = " and RAL01=" & DBDATE(txtCP27) & " AND RAL02='" & m_Dept & "'"
   '上午
   If cboListTime.ListIndex = 0 Then
      strConRAL = strConRAL & " AND RAL03='0'"
   '下午
   Else
      strConRAL = strConRAL & " AND RAL03='1'"
   End If
   
   nCopys = 1
   
   'Modified by Morgan 2019/8/27 服務業務可能沒有申請號 RPAD(SP11,12,' ')-->RPAD(NVL(SP11,' '),12,' ')
   strSql = "SELECT * FROM (" & _
      " select RPAD(cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04),15,' ') 本所案號" & _
      ",RPAD(pa11,12,' ') 申請案號,RPAD(CPM03,12,' ') 案件性質,cp64 案件備註,PA01 CF1,PA09 CF2,CP10 CF3,RAL06" & _
      " From RecAppList, caseprogress, casepropertymap, patent" & _
      " WHERE RAL04='9'" & strConRAL & _
      " and cp09(+)=RAL05 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='000'" & _
      " Union All" & _
      " select RPAD(cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04),15,' ') 本所案號" & _
      ",RPAD(TM12,12,' ') 申請案號,RPAD(CPM03,12,' ') 案件性質,cp64 案件備註,TM01 CF1,TM10 CF2,CP10 CF3,RAL06" & _
      " From RecAppList, caseprogress, casepropertymap, trademark" & _
      " WHERE RAL04='9'" & strConRAL & _
      " and cp09(+)=RAL05 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm10='000'" & _
      " Union All" & _
      " select RPAD(cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04),15,' ') 本所案號" & _
      ",RPAD(NVL(SP11,' '),12,' ') 申請案號,RPAD(CPM03,12,' ') 案件性質,cp64 案件備註,SP01 CF1,SP09 CF2,CP10 CF3,RAL06" & _
      " From RecAppList, caseprogress, casepropertymap, servicepractice" & _
      " where RAL04='9'" & strConRAL & _
      " and cp09(+)=RAL05 and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and Sp01(+)=cp01 and Sp02(+)=cp02 and Sp03(+)=cp03 and Sp04(+)=cp04 and Sp09='000'" & _
      ") X ORDER BY RAL06,1,2"
      
      
On Error GoTo ErrHnd

   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With adoRecordset
      m_intCnt = m_intCnt + .RecordCount 'Add By Sindy 2010/12/21
      Printer.Orientation = 2
      Printer.Font = "細明體"
      For iCopys = 1 To nCopys
         .MoveFirst
         stKey = "" & .Fields("RAL06")
         If iCopys > 1 Then Printer.NewPage
         iPage = 1: iRec = 0: iRecs = 0
         PrintHead1 , stKey
         Do While Not .EOF
            If stKey <> "" & .Fields("RAL06") Then
               PrintTail1 iPage, iRecs
               iRecs = 0
               iRec = 0
               iPage = 1
               stKey = "" & .Fields("RAL06")
               Printer.NewPage
               PrintHead1 , stKey
            End If
            iRecs = iRecs + 1
            iRec = iRec + 1
            If iRec > 26 Then
               PrintTail1 iPage
               Printer.NewPage
               iPage = iPage + 1
               PrintHead1 , stKey
               iRec = 0
            End If
            strTmp = .Fields(0) & Space(1) & .Fields(1) & Space(1) & .Fields(2)
            Printer.CurrentY = Printer.CurrentY + 60
            Printer.Print strTmp
            .MoveNext
         Loop
         PrintTail1 iPage, iRecs
      Next
      Printer.EndDoc
      DoPrint2 = True
      End With
   End If

ErrHnd:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function
