VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm075013 
   BorderStyle     =   1  '單線固定
   Caption         =   "出庭費確認維護"
   ClientHeight    =   5484
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9432
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5484
   ScaleWidth      =   9432
   Begin VB.CheckBox Check1 
      Caption         =   "含不領取案件"
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   1848
      TabIndex        =   9
      Top             =   504
      Value           =   1  '核取
      Width           =   1380
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認維護(&E)"
      Height          =   400
      Index           =   0
      Left            =   5064
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   48
      Width           =   1140
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8568
      TabIndex        =   7
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4248
      TabIndex        =   6
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   1
      Left            =   6228
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   48
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   400
      Index           =   2
      Left            =   7752
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   48
      Width           =   795
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
      Height          =   4500
      Left            =   72
      TabIndex        =   3
      Top             =   864
      Width           =   9228
      _ExtentX        =   16277
      _ExtentY        =   7938
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtUsernum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1032
      MaxLength       =   6
      TabIndex        =   0
      Top             =   130
      Width           =   744
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   144
      TabIndex        =   2
      Top             =   182
      Width           =   900
   End
   Begin MSForms.Label lblUserName 
      Height          =   288
      Left            =   1848
      TabIndex        =   1
      Top             =   132
      Width           =   1116
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "1968;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm075013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/09/30 (113/11/01上線)
Option Explicit
Dim intLastRow As Integer '記錄MGrid1勾選最後一筆
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Public cmdState As Integer

Dim intQ As Integer, strQ1 As String
Dim rsQD As New ADODB.Recordset
Const cntFixed As Integer = 4
Dim colCp09 As Integer, colCaseNo As Integer, colLOS15 As String, colChkCP60 As String

Private Sub Check1_Click()
   Call doQuery(True)
End Sub

Private Sub cmdExit_Click()
   fnCloseAllFrm100  '關閉共同查詢的畫面
   Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   If cmdState = 0 Then
      If PUB_CheckFormExist("frm075013_1") Then
         MsgBox "請先關閉〔出庭費確認維護明細〕畫面！"
         Exit Sub
      End If
   End If
   
   PubShowNextData
End Sub

Public Sub PubShowNextData(Optional bolRefresh As Boolean = False)
Dim intA As Integer, StrTag As String, intB As Integer
Dim Str01 As String, strKeyNo As String

On Error GoTo ErrorHandler
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   For intA = 1 To MGrid1.Rows - 1
      MGrid1.col = 0
      MGrid1.row = intA
      If Trim(MGrid1.Text) = "V" Then
         bolRefresh = False
         MGrid1.col = 0
         MGrid1.Text = ""
         For intB = 0 To MGrid1.Cols - 1
            MGrid1.col = intB
            MGrid1.CellBackColor = &H80000005
         Next
                  
         StrTag = MGrid1.TextMatrix(intA, colCaseNo)
         '增加對執行”基本資料”和”案件進度查詢”的限閱案件控制
         If cmdState = 1 Or cmdState = 2 Then
            If PUB_ChkCufaByCaseNo(strUserNum, Me.Name, Replace(StrTag, "-", ""), "1") = False Then
               Exit For
            End If
         End If
         Str01 = SystemNumber(StrTag, 1)
         If cmdState = 1 Or cmdState = 2 Then
            If fnSaveParentForm(Me) = False Then
               Exit For
            End If
         End If
         strKeyNo = MGrid1.TextMatrix(intA, colCp09)
         
         Me.Show
         Select Case cmdState
            Case 0 '確認維護
               If Len(strKeyNo) = 9 Then
                  If "" & MGrid1.TextMatrix(intA, colChkCP60) <> "Y" Then
                     MsgBox "案件狀態為已請款，才可以確認領取或不領取！", vbInformation
                     Exit For
                  End If
                  Call frm075013_1.SetParent(Me, strKeyNo, txtUsernum)
                  frm075013_1.Show
                  Me.Hide
               End If
            Case 1 '案件基本資料
               Select Case Pub_RplStr(Str01)
                  Case "LA" '顧問
                      Screen.MousePointer = vbHourglass
                      frm100101_6.Show
                      frm100101_6.Tag = StrTag
                      frm100101_6.StrMenu
                      Screen.MousePointer = vbDefault
                  Case Else '法務
                      Screen.MousePointer = vbHourglass
                      frm100101_5.Show
                      frm100101_5.Tag = StrTag
                      frm100101_5.StrMenu
                      Screen.MousePointer = vbDefault
               End Select
               
            Case 2 '案件進度
               frm100101_2.Show
               frm100101_2.Tag = StrTag
               frm100101_2.StrMenu
         End Select
         Exit For
      End If
   Next intA
   
   If bolRefresh = True Then
      cmdQuery_Click
   End If
   
ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "(" & Err.Number & ")" & Err.Description
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuery_Click()
   Call doQuery(True)
End Sub

Private Sub Form_Load()
   
   MoveFormToCenter Me
   txtUsernum = strUserNum
   If Pub_StrUserSt03 = "M51" Then
      txtUsernum.Enabled = True
   Else
      txtUsernum.Enabled = False
   End If
   
   Call doQuery(False)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsQD = Nothing
   Set frm075013 = Nothing
End Sub

Private Sub MGrid1_Click()

   If "" & MGrid1.TextMatrix(MGrid1.row, 1) <> "" Then
      GridClick MGrid1, intLastRow, 0, 0, , "V"
   End If
   
End Sub

Private Sub MGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow MGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MGrid1.col = nCol
   MGrid1.row = nRow
   If Me.MGrid1.row < 1 And Me.MGrid1.Text <> "V" Then
      If InStr("出庭費", Me.MGrid1.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.MGrid1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.MGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub txtUsernum_GotFocus()
   TextInverse txtUsernum
End Sub

Private Sub txtUsernum_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUsernum_Change()
   If Len(txtUsernum) >= 5 Then
      lblUserName = GetStaffName(txtUsernum, True)
   Else
      lblUserName = ""
   End If
End Sub

Private Sub SetGrd1(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
    
   '收文號不顯示
   arrGridHeadText = Array("V", "律所案號", "智慧所案號", "案件名稱", "發文日期", "收文號", "案件性質", "出庭費", "已請款", "不領取日期", "EMAIL通知記錄", "確認記錄", "ST01", "ST02", "CP162", "LOS02", "LOS01", "CP10")
   arrGridHeadWidth = Array(300, 1300, 1300, 1600, 1000, 0, 1300, 900, 720, 1200, 1500, 1200, 0, 0, 0, 0, 0, 0)
        
   MGrid1.Visible = False
   MGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
      MGrid1.Clear
      MGrid1.Rows = 2
   End If
       
   For iRow = 0 To MGrid1.Cols - 1
      MGrid1.row = 0
      MGrid1.col = iRow
      MGrid1.Text = arrGridHeadText(iRow)
      MGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MGrid1.CellAlignment = flexAlignCenterCenter
   Next

   For intI = 1 To MGrid1.Rows - 1
      MGrid1.row = intI
      For iRow = 0 To MGrid1.Cols - 1
         MGrid1.col = iRow
         MGrid1.CellBackColor = &H80000005
         '靠右
         If InStr("07,", Format(iRow, "00")) > 0 Then
            MGrid1.CellAlignment = flexAlignRightCenter
         End If
         '置中
         If InStr("08,", Format(iRow, "00")) > 0 Then
            MGrid1.CellAlignment = flexAlignCenterCenter
         End If
      Next iRow
   Next intI
   
   If colCp09 = 0 Then
      colCp09 = PUB_MGridGetId("收文號", MGrid1)
      colCaseNo = PUB_MGridGetId("律所案號", MGrid1)
      colChkCP60 = PUB_MGridGetId("已請款", MGrid1)
      colLOS15 = PUB_MGridGetId("CP162", MGrid1)
   End If
   
   If pReset = False And MGrid1.Rows > 1 Then
      With MGrid1
         For iRow = 1 To .Rows - 1
            If "" & .TextMatrix(iRow, colLOS15) <> "" Then
               '檢查案源是否可以輸入出庭律師
               If Pub_ChkLosToCL("" & .TextMatrix(iRow, colCp09), False, strExc(1)) = False Then
                  .RowHeight(iRow) = 0 '隱藏
               End If
            End If
         Next iRow
      End With
   End If
   
   MGrid1.Visible = True
End Sub

Public Sub doQuery(ByVal bolMsg As Boolean)
   
   Call SetGrd1(True) '清空
   
   '條件:尚未確認(CL04), 已通知主管後不可以再進行確認counting(cl07), 7/26 判斷財務處已發放(CL06)則不顯示案件資料。
   'Memo by Lydia 2024/09/30 拿掉and instr('," & CaseLawerPtyList & ",',','||cpm12||',') > 0
   'Modified by Lydia 2025/04/07 +CL09財務確認律師不領取出庭費日期
   strQ1 = "select '' as v, lc01||'-'||lc02||'-'||lc03||'-'||lc04 as caseno,decode(c2.cp01,null,null,'TT',null,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04) as pcase," & _
           " nvl(lc05,nvl(lc06,lc07)) as casename,sqldatet(c1.cp27) as cp27t,c1.cp09,decode(lc15,'000',cpm03,cpm04) as cp10n, cl03," & _
           " decode(c1.cp60,null,decode(c2.cp60,null,null,'Y'),'Y') as chkcp60 ," & _
           " sqldatet(cl05) cl05t,cl07,cl08,st01,st02,c1.cp162,los02,los01,c1.cp10" & _
           " from caselawer,caseprogress c1,lawcase,casepropertymap,staff,lawofficesource, caseprogress c2" & _
           " where cl04 is null and counting(cl07) < 4 and cl03 > 0 and cl06||cl09 is null and cl01=c1.cp09(+) and c1.cp158 > 0 and c1.cp159=0 and c1.cp01=lc01(+) and c1.cp02=lc02(+) and c1.cp03=lc03(+) and c1.cp04=lc04(+)" & _
           " and c1.cp01=cpm01(+) and c1.cp10=cpm02(+) and cl02=st01(+) " & _
           " and c1.cp162=los15(+) and los01=c2.cp09(+)"

   If txtUsernum <> "" Then
      strQ1 = strQ1 & " and cl02=" & CNULL(txtUsernum)
   End If
   If Check1.Value = 0 Then
      strQ1 = strQ1 & " and cl05 is null"
   End If
   strQ1 = strQ1 & " order by cl07, c1.cp27"
   
   If bolMsg = True Then
     intQ = 0
   Else
     intQ = 1
   End If
   Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      MGrid1.FixedCols = 0
      Set MGrid1.Recordset = rsQD
      Call SetGrd1
      MGrid1.FixedCols = cntFixed
   End If
End Sub
