VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100123_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "管制備註-查詢"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   8715
   Begin VB.CommandButton CmdCaseQuery 
      Caption         =   "全案管制備註"
      Height          =   375
      Left            =   6270
      TabIndex        =   20
      Top             =   60
      Width           =   1305
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "下一筆"
      Height          =   375
      Left            =   5190
      TabIndex        =   19
      Top             =   60
      Width           =   1065
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2655
      Left            =   60
      TabIndex        =   18
      Top             =   1380
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4683
      _Version        =   393216
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
   Begin VB.CommandButton CmdExit 
      Caption         =   "回前畫面"
      Height          =   375
      Left            =   7590
      TabIndex        =   1
      Top             =   60
      Width           =   1065
   End
   Begin MSForms.Label lblFM2 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   17
      Top             =   945
      Width           =   6645
      VariousPropertyBits=   27
      Caption         =   "lblFM2(0)"
      Size            =   "11721;503"
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
      TabIndex        =   16
      Top             =   945
      Width           =   945
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(6)"
      Height          =   285
      Index           =   6
      Left            =   8760
      TabIndex        =   15
      Top             =   1740
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(5)"
      Height          =   285
      Index           =   5
      Left            =   4110
      TabIndex        =   14
      Top             =   270
      Width           =   1005
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(4)"
      Height          =   285
      Index           =   4
      Left            =   9390
      TabIndex        =   13
      Top             =   810
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(3)"
      Height          =   285
      Index           =   3
      Left            =   8730
      TabIndex        =   12
      Top             =   1020
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(2)"
      Height          =   285
      Index           =   2
      Left            =   1050
      TabIndex        =   11
      Top             =   945
      Width           =   825
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(1)"
      Height          =   285
      Index           =   1
      Left            =   9420
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(0)"
      Height          =   285
      Index           =   0
      Left            =   1050
      TabIndex        =   9
      Top             =   240
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "序號："
      Height          =   225
      Index           =   6
      Left            =   8550
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   225
      Index           =   5
      Left            =   3150
      TabIndex        =   7
      Top             =   270
      Width           =   945
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1050
      TabIndex        =   0
      Top             =   585
      Width           =   7575
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13361;529"
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
      Left            =   8460
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "本所期限："
      Height          =   225
      Index           =   3
      Left            =   8520
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   225
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   585
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "法定期限："
      Height          =   225
      Index           =   1
      Left            =   8430
      TabIndex        =   3
      Top             =   810
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "frm100123_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2021/05/17 管制備註-查詢
'Memo by Lydia 2021/05/17 Form2.0已修改 lblFM2(index)、Combo1
Option Explicit
Dim m_PrevForm As Form  '前一畫面
Dim m_PKey01 As String, m_PKey02 As String  '收文號CP09+CP66 / 下一程序NP02+NP22
Dim m_CMR03 As String '記錄日期
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String '本所案號
Dim m_Na01 As String '申請國家
Dim strTmpQ As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim oObj
Dim m_ARow As Variant, m_APK01 As Variant, m_APK02 As Variant '分別記錄勾選列數、收文號、序號：可選多筆，當顯示資料後將前一畫面的勾選項取消。
Dim m_Now As Integer '目前處理第X筆
Dim m_Max As Integer

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

Private Sub CmdCaseQuery_Click()
    Call doQuery(True)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub doQuery(Optional ByVal bolCaseQuery As Boolean = False)
Dim strConCP As String, strConNP As String

    '案件基本資料
    strTmpQ = "select cp01,cp02,cp03,cp04,sk02 from caseprogress,systemkind where cp09='" & m_PKey01 & "' and cp01=sk01(+) "
    intQ = 0
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 1 Then
       m_CP01 = "" & rsQuery.Fields("cp01")
       m_CP02 = "" & rsQuery.Fields("cp02")
       m_CP03 = "" & rsQuery.Fields("cp03")
       m_CP04 = "" & rsQuery.Fields("cp04")
       Select Case Val("" & rsQuery.Fields("sk02"))
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
                '本所案號
                lblData(0) = m_CP01 & "-" & m_CP02 & IIf(m_CP03 <> "0", "-" & m_CP03, "") & IIf(m_CP04 <> "00", "-" & m_CP04, "")
                '收文號：CP09/NP01
                lblData(5) = "" & m_PKey01
                '序號：CP66/NP22
                lblData(6) = "" & m_PKey02
            End If
        End If
        
        If TypeName(m_PrevForm) <> "Nothing" And bolCaseQuery = False Then
            Call m_PrevForm.UpdateShowFlag(Val("" & m_ARow(m_Now)))
        End If
        
        '備註、管制備註
        If bolCaseQuery = True Then '全案
           strConCP = " AND CP01='" & m_CP01 & "' AND CP02='" & m_CP02 & "' AND CP03='" & m_CP03 & "' AND CP04='" & m_CP04 & "' AND CP159=0 "
           strConNP = " AND NP02='" & m_CP01 & "' AND NP03='" & m_CP02 & "' AND NP04='" & m_CP03 & "' AND NP05='" & m_CP04 & "' "
        Else
           strConCP = " AND CP09='" & m_PKey01 & "' " & IIf(m_PKey02 <> "A", " AND CP66=" & m_PKey02, "")
           strConNP = " AND NP01='" & m_PKey01 & "' " & IIf(m_PKey02 <> "A", " AND NP22=" & m_PKey02, "")
        End If
        '只抓有管制備註NVL(CMR04,'N') <> 'N'
        strTmpQ = "SELECT CP01||'-'||CP02||DECODE(CP03,'0',NULL,'-'||CP03)||DECODE(CP04,'00',NULL,'-'||CP04) CASENO,CP09 AS PKEY01, '進度' AS PKEY02, CPM03 AS CASEPTY,SUBSTR(SQLDATET(CP06),1,10) CP06,SUBSTR(SQLDATET(CP07),1,10) CP07, " & _
                         "SUBSTR(CP64,1,500) AS CASEMEMO,SUBSTR(CMR04,1,500) AS CMR04,CP01,CP02,CP03,CP04,CP10,CP66 " & _
                         "FROM CASEPROGRESS,CTLREMARK,CASEPROPERTYMAP WHERE CP09=CMR01(+) AND CP66=CMR02(+) " & _
                         "AND CP01=CPM01(+) AND CP10=CPM02(+) AND NVL(CMR04,'N') <> 'N' " & strConCP
        strTmpQ = strTmpQ & " Union All " & _
                         "SELECT NP02||'-'||NP03||DECODE(NP04,'0',NULL,'-'||NP04)||DECODE(NP05,'00',NULL,'-'||NP05) CASENO,NP01 AS PKEY01, '下一程序' AS PKEY02, CPM04 AS CASEPTY,SUBSTR(SQLDATET(NP08),1,10) NP08,SUBSTR(SQLDATET(NP09),1,10) NP09, " & _
                         "SUBSTR(NP15,1,500) AS CASEMEMO,SUBSTR(CMR04,1,500) AS CMR04, NP02,NP03,NP04,NP05,NP07,NP22 " & _
                         "FROM NEXTPROGRESS,CTLREMARK,CASEPROPERTYMAP WHERE NP01='" & m_PKey01 & "' AND NP01=CMR01(+) AND NP22=CMR02(+) " & _
                         "AND NP02=CPM01(+) AND NP07=CPM02(+) AND NVL(CMR04,'N') <> 'N' " & _
                         strConNP & strNpSqlOfNoSalesDuty
        Call SetGrd(True) '清空
        intQ = 1
        Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
        If intQ = 1 Then
            MSHFlexGrid1.FixedCols = 0
            Set MSHFlexGrid1.Recordset = rsQuery
            Call SetGrd
            rsQuery.MoveFirst
            '本所期限
            lblData(3) = ChangeWStringToTDateString("" & rsQuery.Fields("cp06"))
            '法定期限
            lblData(4) = ChangeWStringToTDateString("" & rsQuery.Fields("cp07"))
        Else
            If bolCaseQuery = False Then
                MsgBox m_CP01 & "-" & m_CP02 & IIf(m_CP03 <> "0", "-" & m_CP03, "") & IIf(m_CP04 <> "00", "-" & m_CP04, "") & "(" & m_PKey01 & ") 查無管制備註", vbInformation, "管制備註-查詢"
                If m_Now < m_Max Then '查無管制備註，自動跳下一筆
                    Call cmdNext_Click
                End If
            End If
        End If

    End If

End Sub

Private Sub cmdNext_Click()
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
    End If
End Sub

Private Sub Form_Load()

    MoveFormToCenter Me
    
    Call ClearForm
    Call doQuery
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If TypeName(m_PrevForm) <> "Nothing" Then
       m_PrevForm.Show
   End If
   
   Set frm100123_3 = Nothing
End Sub
 
Private Sub ClearForm()

    For Each oObj In lblData
       oObj.Caption = ""
    Next
    For Each oObj In lblFM2
       oObj.Caption = ""
    Next
    Combo1.Clear
    
    m_CMR03 = ""
    m_CP01 = ""
    m_CP02 = ""
    m_CP03 = ""
    m_CP04 = ""
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer, iR As Integer
   Dim lngColor As Long
   Dim strTmp1 As String


   arrGridHeadText = Array("本所案號", "總收文號", "來源", "案件性質", "本所期限", "法定期限", "備註", "管制備註", "CP01", "CP02", "CP03", "CP04", "CP10", "CP66")
   If m_PKey02 <> "A" Then '有傳入序號: 以期限管制日查詢frm100106_2 、frm100106_3
        arrGridHeadWidth = Array(1000, 960, 0, 1000, 860, 860, 1200, 3500, 0, 0, 0, 0, 0, 0)
   Else   '沒有傳入序號: 案件資料及案件進度查詢frm100101_2
        arrGridHeadWidth = Array(1000, 960, 800, 1000, 860, 860, 1200, 3500, 0, 0, 0, 0, 0, 0)
   End If
   
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         MSHFlexGrid1.Clear
         MSHFlexGrid1.Rows = 2
   End If
       
   For iRow = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.row = 0
       MSHFlexGrid1.col = iRow
       MSHFlexGrid1.Text = arrGridHeadText(iRow)
       MSHFlexGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
   Next

   MSHFlexGrid1.Visible = True
   
End Sub
