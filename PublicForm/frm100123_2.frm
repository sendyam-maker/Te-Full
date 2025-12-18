VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100123_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "管制備註作業"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7830
   Begin VB.CommandButton CmdPDF 
      Caption         =   "PDF"
      Height          =   375
      Left            =   3750
      TabIndex        =   25
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "下一筆"
      Height          =   375
      Left            =   4740
      TabIndex        =   24
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "存檔"
      Default         =   -1  'True
      Height          =   375
      Left            =   5730
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "取消"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   120
      TabIndex        =   23
      Top             =   5100
      Width           =   6735
      VariousPropertyBits=   671107099
      Size            =   "11880;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   90
      X2              =   7590
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Label Label1 
      Caption         =   "管制備註內容："
      Height          =   465
      Index           =   8
      Left            =   90
      TabIndex        =   22
      Top             =   3660
      Width           =   855
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   1380
      Index           =   1
      Left            =   1050
      TabIndex        =   0
      Top             =   3600
      Width           =   6285
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "11086;2434"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   960
      Index           =   0
      Left            =   1050
      TabIndex        =   4
      Top             =   2325
      Width           =   6285
      VariousPropertyBits=   -1467989985
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "11086;1693"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "備　註："
      Height          =   225
      Index           =   7
      Left            =   90
      TabIndex        =   21
      Top             =   2400
      Width           =   945
   End
   Begin MSForms.Label lblFM2 
      Height          =   285
      Index           =   0
      Left            =   1950
      TabIndex        =   20
      Top             =   1245
      Width           =   5715
      VariousPropertyBits=   27
      Caption         =   "lblFM2(0)"
      Size            =   "10081;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "申請人1："
      Height          =   225
      Index           =   17
      Left            =   90
      TabIndex        =   19
      Top             =   1245
      Width           =   945
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(6)"
      Height          =   285
      Index           =   6
      Left            =   4140
      TabIndex        =   18
      Top             =   1950
      Width           =   1005
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(5)"
      Height          =   285
      Index           =   5
      Left            =   1050
      TabIndex        =   17
      Top             =   1950
      Width           =   1005
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(4)"
      Height          =   285
      Index           =   4
      Left            =   4140
      TabIndex        =   16
      Top             =   1590
      Width           =   1005
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(3)"
      Height          =   285
      Index           =   3
      Left            =   1050
      TabIndex        =   15
      Top             =   1590
      Width           =   1005
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(2)"
      Height          =   285
      Index           =   2
      Left            =   1050
      TabIndex        =   14
      Top             =   1245
      Width           =   825
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(1)"
      Height          =   285
      Index           =   1
      Left            =   4110
      TabIndex        =   13
      Top             =   540
      Width           =   1425
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(0)"
      Height          =   285
      Index           =   0
      Left            =   1050
      TabIndex        =   12
      Top             =   540
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "序號："
      Height          =   225
      Index           =   6
      Left            =   3180
      TabIndex        =   11
      Top             =   1950
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   225
      Index           =   5
      Left            =   90
      TabIndex        =   10
      Top             =   1950
      Width           =   945
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Top             =   885
      Width           =   6615
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11668;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   225
      Index           =   4
      Left            =   3150
      TabIndex        =   9
      Top             =   540
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限："
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1590
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   225
      Index           =   2
      Left            =   90
      TabIndex        =   7
      Top             =   885
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "法定期限："
      Height          =   225
      Index           =   1
      Left            =   3180
      TabIndex        =   6
      Top             =   1590
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   5
      Top             =   540
      Width           =   945
   End
End
Attribute VB_Name = "frm100123_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2021/05/07 管制備註維護作業
'Memo by Lydia 2021/05/07 Form2.0已修改 lblFM2(index)、Combo1 、txtFM2(index)、textCUID
Option Explicit
Dim m_PrevForm As Form  '前一畫面
Dim m_PKey01 As String, m_PKey02 As String  '(目前處理)收文號CP09+CP66 / 下一程序NP02+NP22
Dim m_CMR03 As String '(目前處理)記錄日期
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String '(目前處理)本所案號
Dim m_CP10 As String '(目前處理)案件性質
Dim strTmpQ As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim oObj
Dim m_ARow As Variant, m_APK01 As Variant, m_APK02 As Variant '分別記錄勾選列數、收文號、序號：可選多筆連續輸入管制備註，存檔後自動帶下一筆；取消則回前畫面
Dim m_Now As Integer '目前處理第X筆
Dim m_Max As Integer
Dim bolToPrint  As Boolean 'Added by Lydia 2021/06/15 存檔後列印

Public Sub SetParent(ByVal pFrm As Form, ByVal pKeyRow As String, ByVal pKey01 As String, ByVal pKey02 As String)
    Set m_PrevForm = pFrm
    m_ARow = Empty:      m_APK01 = Empty: m_APK02 = Empty
    m_ARow = Split(pKeyRow, ",")
    m_APK01 = Split(pKey01, ",")
    m_APK02 = Split(pKey02, ",")
    m_Max = UBound(m_ARow)
    m_Now = 0
    m_PKey01 = m_APK01(m_Now)
    m_PKey02 = m_APK02(m_Now)
End Sub

Private Sub cmdOK_Click()
    
    If Trim(txtFM2(1)) = "" Then
       MsgBox "管制備註內容不可空白！", vbCritical, "檢核資料"
       txtFM2(1).SetFocus
       txtFM2_GotFocus 1
       Exit Sub
    End If
    
    cmdok.Enabled = False
    If FormSave = True Then
        '前一畫面Grid中之管制備註欄內容同步更新
        Call frm100123.UpdateCMR04(Val("" & m_ARow(m_Now)), True, txtFM2(1).Text)
        
        'Added by Lydia 2021/06/15 先存檔後列印
        If bolToPrint = True Then
        Else
        'end 2021/06/15
            '預設跳下一筆或回前畫面
            If m_Max = 0 Then
                cmdok.Enabled = True
                Call cmdExit_Click
                Exit Sub
            Else
                If m_Now + 1 > m_Max Then
                    cmdok.Enabled = True
                    Call cmdExit_Click
                    Exit Sub
                Else
                    Call cmdNext_Click
                End If
            End If
        End If  'Added by Lydia 2021/06/15
    End If
    cmdok.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub doQuery()
Dim intCaseKind As Integer

    strTmpQ = "SELECT CP01,CP02,CP03,CP04,CP01||'-'||CP02||DECODE(CP03,'0',NULL,'-'||CP03)||DECODE(CP04,'00',NULL,'-'||CP04) CASENO,CP09 AS PKEY01, CP66 AS PKEY02, CP10 AS CASEPTY,CP06,CP07,CP64 AS CASEMEMO,CTLREMARK.* " & _
                     "FROM CASEPROGRESS,CTLREMARK WHERE CP09='" & m_PKey01 & "' AND CP66=" & m_PKey02 & " AND CP09=CMR01(+) AND CP66=CMR02(+) " & _
                     "Union All " & _
                     "SELECT NP02,NP03,NP04,NP05,NP02||'-'||NP03||DECODE(NP04,'0',NULL,'-'||NP04)||DECODE(NP05,'00',NULL,'-'||NP05) CASENO,NP01 AS PKEY01, NP22 AS PKEY02, NP07 AS CASEPTY,NP08,NP09,NP15 AS CASEMEMO,CTLREMARK.* " & _
                     "FROM NEXTPROGRESS,CTLREMARK WHERE NP01='" & m_PKey01 & "' AND NP22=" & m_PKey02 & " AND NP01=CMR01(+) AND NP22=CMR02(+)"
    intQ = 0
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 0 Then
        cmdok.Enabled = False
        Exit Sub
    End If
    
    rsQuery.MoveFirst
    '本所案號
    lblData(0) = "" & rsQuery.Fields("caseno")
    m_CP01 = "" & rsQuery.Fields("cp01")
    m_CP02 = "" & rsQuery.Fields("cp02")
    m_CP03 = "" & rsQuery.Fields("cp03")
    m_CP04 = "" & rsQuery.Fields("cp04")
    m_CP10 = "" & rsQuery.Fields("casepty")
    '本所期限
    lblData(3) = ChangeWStringToTDateString("" & rsQuery.Fields("cp06"))
    '法定期限
    lblData(4) = ChangeWStringToTDateString("" & rsQuery.Fields("cp07"))
    'CP09 / NP02
    lblData(5) = "" & rsQuery.Fields("pkey01")
    'CP66 / NP22
    lblData(6) = "" & rsQuery.Fields("pkey02")
    '備註：CP64 / NP15
    txtFM2(0) = "" & rsQuery.Fields("casememo")
    '管制備註：CP64 / NP15
    txtFM2(1) = "" & rsQuery.Fields("cmr04")
    txtFM2(0).Tag = txtFM2(0).Text
    txtFM2(1).Tag = txtFM2(1).Text
    '記錄日期
    m_CMR03 = "" & rsQuery.Fields("cmr03")
    If m_CMR03 <> "" Then
        textCUID = "Create Date: " & ChangeWStringToTDateString(m_CMR03) & String(10, " ") & "Update: " & GetStaffName("" & rsQuery.Fields("cmr05")) & "  " & ChangeWStringToTDateString("" & rsQuery.Fields("cmr06")) & "  " & Format("" & rsQuery.Fields("cmr07"), "00:00:00")
    End If
    
    Call frm100123.UpdateCMR04(m_ARow(m_Now), False, "") '已顯示資料=>取消勾選
    
    '案件基本資料
    strTmpQ = ""
    If ClsPDGetSystemKind(m_CP01, intCaseKind) Then
        Select Case intCaseKind
            Case 專利
              strTmpQ = "select pa05 as cname1,pa06 as cname2,pa07 as cname3 ,cu01||cu02 as cuno," & _
                 "NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cuname,pa09 as na01 " & _
                 "from patent,customer where pa01=" & CNULL(m_CP01) & " and " & _
                 "pa02=" & CNULL(m_CP02) & " and pa03=" & CNULL(m_CP03) & " and " & _
                 "pa04=" & CNULL(m_CP04) & " and substr(pa26,1,8)=cu01(+) and " & _
                 "substr(pa26,9,1)=cu02(+) "
           Case 商標
              strTmpQ = "select tm05 as cname1,tm06 as cname2,tm07 as cname3,cu01||cu02 as cuno," & _
                 "NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cuname,tm10 as na01 " & _
                 "from trademark,customer where tm01=" & CNULL(m_CP01) & " and " & _
                 "tm02=" & CNULL(m_CP02) & " and tm03=" & CNULL(m_CP03) & " and " & _
                 "tm04=" & CNULL(m_CP04) & " AND substr(tm23,1,8)=cu01(+) and " & _
                 "substr(tm23,9,1)=cu02(+)"
           Case 法務
              strTmpQ = "select lc05 as cname1,lc06 as cname2 ,lc07 as cname3,cu01||cu02 as cuno," & _
                 "NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cuname,lc15 as na01 " & _
                 "from lawcase,customer where lc01=" & CNULL(m_CP01) & " and " & _
                 "lc02=" & CNULL(m_CP02) & " and lc03=" & CNULL(m_CP03) & " and " & _
                 "lc04=" & CNULL(m_CP04) & " and substr(lc11,1,8)=cu01(+) and " & _
                 "substr(lc11,9,1)=cu02(+)"
           Case 顧問
              strTmpQ = "select hc06 as cname1,'' as cname2,'' as cname3,cu01||cu02 as cuno," & _
                 "NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cuname,'000' as na01 " & _
                 "from hirecase,customer where hc01=" & CNULL(m_CP01) & " and hc02=" & CNULL(m_CP02) & _
                 " and hc03=" & CNULL(m_CP03) & " and hc04=" & CNULL(m_CP04) & _
                 " and substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+)"
           Case Else
              strTmpQ = "select sp05 as cname1,sp06 as cname2,sp07 as cname3,cu01||cu02 as cuno," & _
                 "NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cuname,sp09 as na01 " & _
                 "from servicepractice,customer where sp01=" & CNULL(m_CP01) & _
                 " and sp02=" & CNULL(m_CP02) & " and sp03=" & CNULL(m_CP03) & " and sp04=" & _
                 CNULL(m_CP04) & " and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) "
        End Select
        If strTmpQ <> "" Then
            intQ = 0
            Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
            If intQ = 1 Then
                intQ = 0
                Combo1.AddItem "中：" & rsQuery.Fields("cname1"), 0
                If rsQuery.Fields("cname1") <> "" Then intQ = 1
                Combo1.AddItem "英：" & rsQuery.Fields("cname2"), 1
                If rsQuery.Fields("cname2") <> "" Then intQ = 2
                Combo1.AddItem "日：" & rsQuery.Fields("cname3"), 2
                If rsQuery.Fields("cname3") <> "" Then intQ = 3
                Combo1.ListIndex = intQ - 1
                
                lblData(2) = "" & rsQuery.Fields("cuno")
                lblFM2(0) = "" & rsQuery.Fields("cuname")
                
                If ClsPDGetCaseProperty(m_CP01, m_CP10, strExc(1), IIf("" & rsQuery.Fields("na01") <> "000", True, False)) Then
                    lblData(1) = strExc(1)
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()

    MoveFormToCenter Me
    
    textCUID.BackColor = &H8000000F
    
    Call ClearForm
    Call doQuery
    
    If m_Max = 0 Then
        CmdNext.Visible = False
    Else
        CmdNext.Visible = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
   If cmdok.Visible = True And cmdok.Enabled = True Then
       If Trim(txtFM2(0) & txtFM2(1)) <> Trim(txtFM2(0).Tag & txtFM2(1).Tag) Then
          If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
             Cancel = 1
          End If
       End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If TypeName(m_PrevForm) <> "Nothing" Then
       m_PrevForm.Show
   End If
   
   Set frm100123_2 = Nothing
End Sub

Private Sub ClearForm()

    For Each oObj In txtFM2
        oObj.Text = "0"
        oObj.Tag = "0"
    Next
    For Each oObj In lblData
       oObj.Caption = ""
    Next
    For Each oObj In lblFM2
       oObj.Caption = ""
    Next
    Combo1.Clear
    
    textCUID = ""
    m_CMR03 = ""
    m_CP01 = ""
    m_CP02 = ""
    m_CP03 = ""
    m_CP04 = ""
    bolToPrint = False 'Added by Lydia 2021/06/15
End Sub

Private Function FormSave() As Boolean

On Error GoTo ErrHandle
      
   strSql = ""
   If m_CMR03 = "" Then
        strSql = "insert into CtlRemark (cmr01,cmr02,cmr03,cmr04,cmr05,cmr06,cmr07) values (" & _
                    CNULL(m_PKey01) & "," & CNULL(m_PKey02) & "," & strSrvDate(1) & "," & CNULL(ChgSQL(txtFM2(1))) & "," & CNULL(strUserNum) & "," & strSrvDate(1) & "," & Format(ServerTime, "000000") & " )"
   Else
        strSql = "update ctlremark set cmr04=" & CNULL(ChgSQL(txtFM2(1))) & ", cmr05=" & CNULL(strUserNum) & ", cmr06=" & strSrvDate(1) & ", cmr07=" & Format(ServerTime, "000000") & _
                    " where cmr01=" & CNULL(m_PKey01) & " and cmr02=" & CNULL(m_PKey02) & " and cmr03=" & CNULL(m_CMR03)
   End If
   If strSql <> "" Then
        cnnConnection.BeginTrans
           cnnConnection.Execute strSql
        cnnConnection.CommitTrans
        txtFM2(0).Tag = txtFM2(0).Text
        txtFM2(1).Tag = txtFM2(1).Text
        FormSave = True
   End If
   
ErrHandle:
   If Err.Number <> 0 Then
        If strSql <> "" Then cnnConnection.RollbackTrans
        MsgBox "存檔失敗：" & vbCrLf & Err.Description, vbCritical
   End If
   
End Function

Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub cmdNext_Click()
    If cmdok.Visible = True And cmdok.Enabled = True Then
       If Trim(txtFM2(0) & txtFM2(1)) <> Trim(txtFM2(0).Tag & txtFM2(1).Tag) Then
          If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
             Exit Sub
          End If
       End If
    End If
   
    m_Now = m_Now + 1
    If m_Now > m_Max Then
        MsgBox "已經是最後一筆！", vbInformation
        Call cmdExit_Click
        Exit Sub
    Else
        m_PKey01 = m_APK01(m_Now)
        m_PKey02 = m_APK02(m_Now)
        Call ClearForm
        Call doQuery
        txtFM2(1).SetFocus
        Call txtFM2_GotFocus(1)
    End If
End Sub

'Added by Lydia 2021/06/15 列印管制備註PDF
Private Sub CmdPDF_Click()
   If cmdok.Visible = True And cmdok.Enabled = True Then
       If Trim(txtFM2(0) & txtFM2(1)) <> Trim(txtFM2(0).Tag & txtFM2(1).Tag) Then
          If MsgBox("資料已修改，是否存檔？", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
             Exit Sub
          Else
             bolToPrint = True
             Call cmdOK_Click
          End If
       End If
   End If
   
   
   '檔案存放路徑：個人桌面\案號.民國年月日＋時分.REPLY.pdf
   cmdPDF.Enabled = False
   Call PrintCtlRemarkPDF(m_CP01, m_CP02, m_CP03, m_CP04, m_PKey01, m_PKey02)
   bolToPrint = False
   cmdPDF.Enabled = True
End Sub

'Added by Lydia  2021/06/15 期限資料查詢/管制備註：列印管制備註PDF  ;
Private Sub PrintCtlRemarkPDF(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pKey01 As String, ByVal pKey02 As String)
Dim stSQL As String, intR As Integer
Dim stTmp1 As String
Dim rsQuery As ADODB.Recordset
Dim xlsReport
Dim wksReport
Dim strTempFile As String, stFileTime As String, stFullPath As String
Dim xRow As Integer
Dim tmpArr As Variant

On Error GoTo ErrHandle
  
   stSQL = "SELECT CMR04 FROM CTLREMARK WHERE CMR01='" & pKey01 & "' AND CMR02='" & pKey02 & "' "
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      stFileTime = Format(ServerTime, "000000")
      If Dir(App.path & "\" & strUserNum, vbDirectory) = "" Then
          MkDir App.path & "\" & strUserNum
      End If
      strTempFile = App.path & "\" & strUserNum & "\$" & strSrvDate(2) & stFileTime & "CMR.xls"
      tmpArr = Split("" & rsQuery.Fields("cmr04"), vbCrLf)
      
      Set xlsReport = CreateObject("Excel.Application")
      xlsReport.SheetsInNewWorkbook = 3 '工作表份數
      xlsReport.Workbooks.add
      Set wksReport = xlsReport.Worksheets(1)
      wksReport.Cells.NumberFormatLocal = "@"
      xRow = xRow + 1
      wksReport.Range("A" & xRow) = "期限資料查詢/管制備註" & String(10, " ") & "產生日期時間：" & ChangeTStringToTDateString(strSrvDate(2)) & " " & Mid(Format(stFileTime, "00:00:00"), 1, 5)
      wksReport.Range("A:A").ColumnWidth = 90
      wksReport.Range("1:1").RowHeight = 30
      wksReport.Range("A" & xRow).Font.Size = 16
      wksReport.Range("A" & xRow).Font.Bold = True
      wksReport.Range("A" & xRow).Font.Name = "標楷體"
      
      xRow = xRow + 1
      wksReport.Range(xRow & ":" & xRow).RowHeight = 30
      stSQL = GetStaffName(strUserNum, True, stTmp1)
      wksReport.Range("A" & xRow).Value = "列印人員：" & stTmp1 & " " & strUserName & String(2, " ")
      wksReport.Range("A" & xRow).Font.Name = "標楷體"
      wksReport.Range("A" & xRow).Font.Size = 16
      wksReport.Range("A" & xRow).Font.Bold = True
      wksReport.Range("A" & xRow).HorizontalAlignment = xlRight
      
      xRow = xRow + 1
      For intR = 0 To UBound(tmpArr)
         xRow = xRow + 1
         'P.S. 要能夠自動換列，不能設定在合併的儲存格，不能固定列高
         wksReport.Range("A" & xRow).Value = "" & tmpArr(intR)
         wksReport.Range("A" & xRow).WrapText = True                '設儲存格為自動換列
      Next intR
      wksReport.Range("A3:A" & xRow).Font.Size = 14
      wksReport.Range("A3:A" & xRow).Select
     
      stFullPath = PUB_Getdesktop & "\" & PUB_CaseNo2FileName(pCP01, pCP02, pCP03, pCP04) & "." & strSrvDate(2) & stFileTime & ".REPLY.PDF"
      
      '列印設定
      wksReport.PageSetup.PaperSize = 9 'A4
      wksReport.PageSetup.Orientation = wdOrientLandscape '直印
      wksReport.PageSetup.CenterHorizontally = True '垂直置中
      '列印邊界
      wksReport.PageSetup.LeftMargin = xlsReport.CentimetersToPoints(1)
      wksReport.PageSetup.RightMargin = xlsReport.CentimetersToPoints(1)
      wksReport.PageSetup.TopMargin = xlsReport.CentimetersToPoints(1.5)
      wksReport.PageSetup.BottomMargin = xlsReport.CentimetersToPoints(1.5)
      wksReport.PageSetup.HeaderMargin = xlsReport.CentimetersToPoints(1.5)
      wksReport.PageSetup.FooterMargin = xlsReport.CentimetersToPoints(1.5)
    
      '判斷版本2007
      If Val(xlsReport.Version) < 12 Then
           xlsReport.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=stFullPath, Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
           xlsReport.Workbooks(1).SaveAs FileName:=strTempFile, FileFormat:=56
      '版本2007以上
      Else
           xlsReport.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=stFullPath, Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
           xlsReport.Workbooks(1).SaveAs FileName:=strTempFile, FileFormat:=56
      End If
      xlsReport.Workbooks.Close
      xlsReport.Quit
      Kill strTempFile
      MsgBox "檔案產生完成！（檔案位置：" & stFullPath & "）"
   End If
ErrHandle:
   If Err.Number <> 0 Then
         MsgBox Err.Description, vbCritical
   End If

   Set rsQuery = Nothing
   Set wksReport = Nothing
   Set xlsReport = Nothing
End Sub

