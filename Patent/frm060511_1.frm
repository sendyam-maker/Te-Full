VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060511_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專案件清單Excel：查詢歷史記錄"
   ClientHeight    =   5016
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8832
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5016
   ScaleWidth      =   8832
   Begin VB.Frame Frame1 
      Height          =   1428
      Left            =   72
      TabIndex        =   9
      Top             =   48
      Width           =   7476
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   5256
         MaxLength       =   6
         TabIndex        =   4
         Top             =   144
         Width           =   800
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1416
         MaxLength       =   9
         TabIndex        =   3
         Top             =   1080
         Width           =   1100
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1416
         MaxLength       =   9
         TabIndex        =   2
         Top             =   768
         Width           =   1100
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1416
         MaxLength       =   9
         TabIndex        =   1
         Top             =   456
         Width           =   1100
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1416
         MaxLength       =   6
         TabIndex        =   0
         Top             =   144
         Width           =   800
      End
      Begin VB.CommandButton cmdRe 
         Caption         =   "重新查詢"
         Height          =   350
         Left            =   6192
         TabIndex        =   5
         Top             =   120
         Width           =   1044
      End
      Begin VB.Label Label1 
         Caption         =   "記錄年度(西元年)："
         Height          =   228
         Index           =   4
         Left            =   3624
         TabIndex        =   18
         Top             =   172
         Width           =   1596
      End
      Begin MSForms.Label lblFM2 
         Height          =   300
         Index           =   3
         Left            =   2568
         TabIndex        =   17
         Top             =   1068
         Width           =   4800
         BackColor       =   16777215
         Size            =   "8467;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   300
         Index           =   2
         Left            =   2568
         TabIndex        =   16
         Top             =   756
         Width           =   4800
         BackColor       =   16777215
         Size            =   "8467;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   300
         Index           =   1
         Left            =   2568
         TabIndex        =   15
         Top             =   456
         Width           =   4800
         BackColor       =   16777215
         Size            =   "8467;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblFM2 
         Height          =   300
         Index           =   0
         Left            =   2280
         TabIndex        =   14
         Top             =   120
         Width           =   1212
         BackColor       =   16777215
         Size            =   "2138;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "固定請款對象："
         Height          =   228
         Index           =   3
         Left            =   96
         TabIndex        =   13
         Top             =   1104
         Width           =   1284
      End
      Begin VB.Label Label1 
         Caption         =   "客戶編號："
         Height          =   228
         Index           =   2
         Left            =   96
         TabIndex        =   12
         Top             =   792
         Width           =   1116
      End
      Begin VB.Label Label1 
         Caption         =   "代理人編號："
         Height          =   228
         Index           =   1
         Left            =   96
         TabIndex        =   11
         Top             =   492
         Width           =   1116
      End
      Begin VB.Label Label1 
         Caption         =   "需求人員："
         Height          =   228
         Index           =   0
         Left            =   96
         TabIndex        =   10
         Top             =   168
         Width           =   1116
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3252
      Left            =   72
      TabIndex        =   8
      Top             =   1680
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   5736
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "回前畫面(&U)"
      Height          =   444
      Left            =   7728
      TabIndex        =   6
      Top             =   96
      Width           =   1044
   End
   Begin VB.CommandButton cmdImport 
      BackColor       =   &H00FFFFC0&
      Caption         =   "載入記錄"
      Height          =   444
      Left            =   7728
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   888
      Width           =   1044
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFC0&
      BorderWidth     =   2
      X1              =   72
      X2              =   8602
      Y1              =   1584
      Y2              =   1584
   End
End
Attribute VB_Name = "frm060511_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/01/18
Option Explicit
Dim intLastRow As Integer
Dim mPrevForm As Form
Private Const cFixed As Integer = 5 '固定欄位
Dim strKeyNo As String
Dim strPUser As String, strFaNo As String, strCUNo As String, strCuNo2 As String  '從前一畫面傳入

Public Sub SetParent(ByVal fm As Form, Optional ByVal pUserNo As String, Optional ByVal pFaNo As String, Optional ByVal pCuNo As String, Optional ByVal pCuNo2 As String)
   Set mPrevForm = fm
   strPUser = pUserNo '需求人員
   strFaNo = pFaNo '代理人編號以;區隔
   strCUNo = pCuNo '客戶編號以;區隔
   strCuNo2 = pCuNo2 '固定請款對象
End Sub

Private Sub cmdExit_Click()
   '先不詢問
   'If strKeyNo <> "" Then
    '  If MsgBox("是否匯入勾選記錄：" & strKeyNo, vbExclamation + vbYesNo + vbDefaultButton1) = vbNo Then
    '     strKeyNo = ""
    '  End If
   'End If
   
   Unload Me
End Sub

Private Sub cmdImPort_Click()
   strKeyNo = ""
   PubShowNextData
   If strKeyNo <> "" Then
      Call cmdExit_Click
   End If
End Sub

Private Sub cmdRe_Click()
   If QueryData = False Then
      MsgBox "查無資料！" & IIf(Trim(Text1(0) & Text1(1) & Text1(2) & Text1(3)) <> "可以清除畫面輸入條件再重新查詢。", vbCrLf, ""), vbInformation
   End If
End Sub

Private Sub Form_Load()
Dim oObj As Control

   MoveFormToCenter Me
      
   For Each oObj In Text1
      oObj.Text = ""
      oObj.Tag = ""
   Next
   For Each oObj In lblFM2
      oObj.Caption = ""
   Next
   
   Call QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)

   If TypeName(mPrevForm) <> "Nothing" Then
      mPrevForm.Show
      If strKeyNo <> "" Then
         Call mPrevForm.QueryData(strKeyNo)
      End If
   End If
   
   Set frm060511_1 = Nothing
End Sub

Private Function QueryData() As Boolean
Dim rsQuery As New ADODB.Recordset
Dim intQ As Integer, strQuery As String
Dim strCon As String
    
   '需求人員
   If Text1(0) <> "" Then
      strCon = strCon & " AND FER05=" & CNULL(Text1(0))
   ElseIf strPUser <> "" Then
      strCon = strCon & " AND FER05=" & CNULL(strPUser)
   End If
   '代理人編號
   If Text1(1) <> "" Then
      strCon = strCon & " AND INSTR(FER06," & CNULL(Text1(1)) & ") > 0"
   ElseIf strFaNo <> "" Then
      strCon = strCon & " AND INSTR(FER06," & CNULL(strFaNo) & ") > 0"
   End If
   '客戶編號
   If Text1(2) <> "" Then
      strCon = strCon & " AND INSTR(FER07," & CNULL(Text1(2)) & ") > 0"
   ElseIf strCUNo <> "" Then
      strCon = strCon & " AND INSTR(FER07," & CNULL(strCUNo) & ") > 0"
   End If
   '固定請款對象
   If Text1(3) <> "" Then
      strCon = strCon & " AND INSTR(FER10," & CNULL(Text1(3)) & ") > 0"
   ElseIf strCuNo2 <> "" Then
      strCon = strCon & " AND INSTR(FER10," & CNULL(strCuNo2) & ") > 0"
   End If
   
   '沒有條件先預設3個月
   If strCon = "" Then
      strCon = " AND FER03>=" & CompDate(1, -3, strSrvDate(1))
   Else
      If Text1(4) <> "" Then
         If Trim(Text1(0) & Text1(1) & Text1(2) & Text1(3)) = "" Then
            '記錄年度(最優先)
            strCon = " AND SUBSTR(FER03,1,4)=" & CNULL(Text1(4))
         Else
            strCon = strCon & " AND SUBSTR(FER03,1,4)=" & CNULL(Text1(4))
         End If
      End If
   End If
   
   Call SetGrd(True)
   strQuery = " select '' as v, fer01,st02 as fer05n,sqldatet(fer03) as fer03t,fer35 ,getfagentnamelist(rtrim(substr(fer06,1,20)))||decode(rtrim(substr(fer06,21,30)),null,null,';(略)') as fer06n" & _
            " ,getcustomernamelist(rtrim(substr(fer07,1,20)))||decode(rtrim(substr(fer07,21,30)),null,null,';(略)') as fer07n" & _
            " ,decode(substr(fer10,1,1),'Y',getfagentnamelist(fer10),'X',getcustomernamelist(fer10),null) as fer10n" & _
            " ,(decode(fer12,'Y','案件清單,','')||decode(fer13,'Y','未發文,','')||" & _
            " decode(fer14,'Y','未請款'||decode(fer15||fer16,null,',',sqldatew(fer15)||'~'||sqldatew(fer16)||','),null)||" & _
            " decode(fer17,'Y','未收文'||decode(fer18||fer19,null,',',sqldatew(fer18)||'~'||sqldatew(fer19)||','),null)||" & _
            " decode(fer21,'Y','未付款'||decode(fer36||fer37,null,',',sqldatew(fer36)||'~'||sqldatew(fer37)||','),null)||" & _
            " decode(fer22,'Y','行事曆,','')||" & _
            " decode(fer32,'1','提申日期'||fer33||'~'||fer34||',','2','公告日期：'||fer33||'~'||fer34,'')" & _
            " ) as smemo,getferlist(fer26) as fer26n" & _
            " from fcpelistrec, staff where fer05=st01(+)" & strCon
   strQuery = strQuery & " order by fer03 desc, fer04 desc "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      QueryData = True
      MSHFlexGrid1.FixedCols = 0
      Set MSHFlexGrid1.Recordset = rsQuery
      Call SetGrd
      MSHFlexGrid1.FixedCols = cFixed
   End If
   
   Set rsQuery = Nothing
End Function

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTempA As String

   Select Case Index
      Case 0  '需求人員
         If Text1(Index).Text <> Text1(Index).Tag Then
            lblFM2(Index).Caption = ""
            strTempA = GetStaffName(Text1(Index).Text, True)
            If strTempA <> "" Then
               lblFM2(Index).Caption = strTempA
            Else
               MsgBox "請輸入正確的員工編號!"
               GoTo EXITSUB
            End If
         End If
      Case 1, 2, 3
         If Text1(Index).Text <> Text1(Index).Tag Then
            lblFM2(Index).Caption = ""
            Text1(Index) = ChangeCustomerL(Text1(Index))
            If Index = 1 Then   '代理人
               strTempA = GetFAgentName(Text1(Index))
            ElseIf Index = 2 Then   '申請人
               strTempA = GetCustomerName(Text1(Index), "1")
            Else '3
               If Left(Text1(Index), 1) = "Y" Then
                  strTempA = GetFAgentName(Text1(Index))
               ElseIf Left(Text1(Index), 1) = "X" Then
                  strTempA = GetCustomerName(Text1(Index), "1")
               Else
                  strTempA = ""
               End If
            End If
            If strTempA <> "" Then
               lblFM2(Index).Caption = strTempA
            Else
               MsgBox "資料庫無資料 !", vbInformation
               GoTo EXITSUB
            End If
         End If
       Case 4 '西元年檢查
         If Text1(Index) <> "" Then
            If Val(Text1(Index) & "0101") > strSrvDate(1) Then
               MsgBox "輸入西元年不可大於系統日！", vbExclamation
               GoTo EXITSUB
            ElseIf CheckIsDate(Text1(Index) & "0101") = False Then
               GoTo EXITSUB
            End If
         End If
   End Select
   Exit Sub
   
EXITSUB:
   Cancel = True
   Text1(Index).SetFocus
   Text1_GotFocus Index
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
  
                           '1   |2        |3          |4          |5        |6         |7             |8          |9          |10
   arrGridHeadText = Array("v", "流水號", "需求人員", "建立日期", "備　　註", "代理人", "申請人", "固定請款對象", "清單選項", "顯示ITEM")
   arrGridHeadWidth = Array(200, 0, 900, 900, 1500, 1500, 1500, 1500, 2500, 2500)
   
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
      If iRow >= 7 Then
         MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
      End If
   Next
   For intI = 1 To MSHFlexGrid1.Rows - 1
      MSHFlexGrid1.row = intI
      For iRow = 0 To MSHFlexGrid1.Cols - 1
         MSHFlexGrid1.col = iRow
         MSHFlexGrid1.CellBackColor = &HFFFFFF
      Next iRow
   Next intI
   MSHFlexGrid1.Visible = True
End Sub

Private Sub MSHFlexGrid1_Click()

   GridClick MSHFlexGrid1, intLastRow, 0, , , "V"
End Sub

Public Sub PubShowNextData()
Dim intX As Integer
Dim Str01 As String

On Error GoTo ErrHand01
    
    For intX = 1 To MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.row = intX
       MSHFlexGrid1.col = 0
       If Trim(MSHFlexGrid1.Text) = "V" Then
           MSHFlexGrid1.Text = ""
           MSHFlexGrid1.col = 0
           MSHFlexGrid1.CellBackColor = MSHFlexGrid1.BackColor
           Str01 = Trim(MSHFlexGrid1.TextMatrix(intX, 1)) '流水號
           Exit For
       End If
    Next intX
    strKeyNo = Str01
    
    Me.Enabled = True

    Exit Sub

ErrHand01:
    If Err.Number <> 0 Then
         MsgBox Err.Description
         Resume Next
    End If

End Sub
