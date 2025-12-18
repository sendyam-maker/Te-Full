VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050110 
   BorderStyle     =   1  '單線固定
   Caption         =   "外翻人員給案維護"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8952
   Begin VB.CommandButton cmdSearch 
      Caption         =   "未收達"
      Height          =   348
      Index           =   3
      Left            =   5376
      TabIndex        =   21
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "達完稿期限"
      Height          =   348
      Index           =   2
      Left            =   6240
      TabIndex        =   20
      Top             =   1152
      Width           =   1092
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "未完稿"
      Height          =   348
      Index           =   1
      Left            =   5376
      TabIndex        =   19
      Top             =   1152
      Width           =   800
   End
   Begin VB.CommandButton cmdJPList 
      Caption         =   "日本部案件清單"
      Height          =   348
      Left            =   7392
      TabIndex        =   18
      Top             =   1152
      Width           =   1452
   End
   Begin VB.TextBox txtEP09 
      Height          =   276
      Index           =   2
      Left            =   2328
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1200
      Width           =   1020
   End
   Begin VB.TextBox txtEP09 
      Height          =   276
      Index           =   1
      Left            =   1104
      MaxLength       =   7
      TabIndex        =   7
      Top             =   1200
      Width           =   1020
   End
   Begin VB.TextBox txtTF26 
      Height          =   276
      Index           =   2
      Left            =   2328
      MaxLength       =   7
      TabIndex        =   6
      Top             =   864
      Width           =   1020
   End
   Begin VB.TextBox txtTF26 
      Height          =   276
      Index           =   1
      Left            =   1104
      MaxLength       =   7
      TabIndex        =   5
      Top             =   864
      Width           =   1020
   End
   Begin VB.TextBox txtCP14 
      Height          =   276
      Left            =   1104
      MaxLength       =   5
      TabIndex        =   4
      Top             =   504
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "輸入(&E)"
      Height          =   348
      Left            =   7164
      TabIndex        =   10
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   348
      Left            =   8004
      TabIndex        =   11
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   348
      Index           =   0
      Left            =   4488
      TabIndex        =   9
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox txtCaseNo 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   1104
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "CFP"
      Top             =   168
      Width           =   495
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   2
      Left            =   1584
      MaxLength       =   6
      TabIndex        =   1
      Top             =   168
      Width           =   855
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   3
      Left            =   2424
      MaxLength       =   1
      TabIndex        =   2
      Top             =   168
      Width           =   255
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   4
      Left            =   2664
      MaxLength       =   2
      TabIndex        =   3
      Top             =   168
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4056
      Left            =   36
      TabIndex        =   12
      Top             =   1572
      Width           =   8832
      _ExtentX        =   15579
      _ExtentY        =   7154
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblCP14N 
      AutoSize        =   -1  'True
      Caption         =   "外翻名稱"
      Height          =   180
      Left            =   2184
      TabIndex        =   17
      Top             =   552
      Width           =   1176
   End
   Begin VB.Line Line2 
      X1              =   2136
      X2              =   2376
      Y1              =   1344
      Y2              =   1344
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "完稿日:"
      Height          =   180
      Index           =   4
      Left            =   144
      TabIndex        =   16
      Top             =   1248
      Width           =   588
   End
   Begin VB.Line Line1 
      X1              =   2136
      X2              =   2376
      Y1              =   1008
      Y2              =   1008
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "完稿期限:"
      Height          =   180
      Index           =   3
      Left            =   144
      TabIndex        =   15
      Top             =   912
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "外翻編號:"
      Height          =   180
      Index           =   1
      Left            =   144
      TabIndex        =   14
      Top             =   552
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   144
      TabIndex        =   13
      Top             =   216
      Width           =   768
   End
End
Attribute VB_Name = "frm050110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2025/11/4
Option Explicit

Dim intLastRow As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Private Sub SetGridHead()
    Dim ii As Integer, jj As Integer
    FixGrid MSHFlexGrid1
    With MSHFlexGrid1
      .Visible = False
      .row = 0
      ii = 0
      .col = 0: .ColWidth(.col) = 200: .Text = "v"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 1100: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 900: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 1200: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 800: .Text = "給案日"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 950: .Text = "中文字數"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 950: .Text = "原文字數"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 800: .Text = "完稿期限"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 800: .Text = "本所期限"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 800: .Text = "收達日"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 800: .Text = "完稿日"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 800: .Text = "會稿日"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 800: .Text = "發文日"
      .ColAlignment(.col) = flexAlignCenterCenter
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 750: .Text = "分配點數"
      .ColAlignment(.col) = flexAlignRightCenter
      .CellAlignment = flexAlignCenterCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 1000: .Text = "備註"
      .CellAlignment = flexAlignCenterCenter
      For jj = ii + 1 To .Cols - 1
       .ColWidth(jj) = 0
      Next
      .Visible = True
    End With
End Sub

Private Sub ClearGrid()
    Dim rstGrid As New ADODB.Recordset, stSQL As String
    
    stSQL = "SELECT 0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20 FROM DUAL WHERE ROWNUM<1"
    rstGrid.CursorLocation = adUseClient
    rstGrid.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
    Set MSHFlexGrid1.Recordset = rstGrid
    SetGridHead
    Set rstGrid = Nothing
End Sub

Private Sub cmdExit_Click()
    blnIsFormBack = False
    Unload Me
End Sub

Private Sub cmdJPList_Click()
   Dim stDateTo As String, stXLSX As String, hLocalFile As Long
   stDateTo = Left(strSrvDate(1), 6) - 191100
   stDateTo = InputBox("請輸入發文月份：" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "※匯出上月26日到本月25日發文案件資料成Excel檔", , stDateTo)
   If stDateTo <> "" Then
      Screen.MousePointer = vbHourglass
      If PUB_JPTranCaseExport(stDateTo, stXLSX, , True) Then
         ShellExecute hLocalFile, "open", stXLSX, vbNullString, vbNullString, 1
      End If
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub cmdOK_Click()

    Dim ii As Integer
    
    If MSHFlexGrid1.Rows < 2 Then Exit Sub
      
    With MSHFlexGrid1
        .Visible = False
        For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 0) = "v" Then Exit For
        Next ii
        .Visible = True
        If ii = .Rows Then
            MsgBox "請點選欲輸入資料！"
        Else
            frm050110_1.Show
            strExc(1) = PUB_MGridGetValue(ii, "cp09", MSHFlexGrid1)
            frm050110_1.ReadAllData strExc(1)
            Me.Hide
        End If
        
    End With
   
End Sub

'更新維護後的資料
Public Sub UpdatRecord()
   Dim ii As Integer
   With MSHFlexGrid1
   .Visible = False
   For ii = 1 To .Rows - 1
      If .TextMatrix(ii, 0) = "v" Then
         strExc(1) = PUB_MGridGetValue(ii, "cp09", MSHFlexGrid1)
         strExc(0) = "select  SQLDATET(TF26) 完稿期限, decode(tf27,'5',TF23) 中文字數, decode(tf27,'5',null,TF23) 原文字數,SQLDATET(TF32) 收達日, TF04 分配點數,tf36 備註" & _
            " from transfee where tf01='" & strExc(1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            PUB_MGridSetValue ii, "完稿期限", "" & RsTemp(0), MSHFlexGrid1
            PUB_MGridSetValue ii, "中文字數", "" & RsTemp(1), MSHFlexGrid1
            PUB_MGridSetValue ii, "原文字數", "" & RsTemp(2), MSHFlexGrid1
            PUB_MGridSetValue ii, "收達日", "" & RsTemp(3), MSHFlexGrid1
            PUB_MGridSetValue ii, "分配點數", "" & RsTemp(4), MSHFlexGrid1
            PUB_MGridSetValue ii, "備註", "" & RsTemp(5), MSHFlexGrid1
         End If
         Exit For
      End If
   Next ii
   .Visible = True
   End With
End Sub

Private Sub SetGrid(idx As Integer)

On Error GoTo flgErr

   Dim rstGrid As New ADODB.Recordset
   Dim stSQL As String
   Dim arrCaseNo(1 To 4) As String
   Dim stCon As String
   
   ClearGrid
   
   stCon = ""
   arrCaseNo(1) = txtCaseNo(1)
   
   If idx = 0 Then
      If txtCaseNo(2) <> "" Then
         arrCaseNo(2) = Right("000000" & txtCaseNo(2), 6)
         arrCaseNo(3) = Right("0" & txtCaseNo(3), 1)
         arrCaseNo(4) = Right("00" & txtCaseNo(4), 2)
         stCon = stCon & " and cp02='" & arrCaseNo(2) & "' and cp03='" & arrCaseNo(3) & "' and cp04='" & arrCaseNo(4) & "'"
      End If
      '承辦人
      If txtCP14 <> "" Then
         stCon = stCon & " and cp14='" & txtCP14 & "'"
      End If
      '完稿期限
      If txtTF26(1) <> "" Then
         stCon = stCon & " and TF26>=" & DBDATE(txtTF26(1)) & ""
      End If
      If txtTF26(2) <> "" Then
         stCon = stCon & " and TF26<=" & DBDATE(txtTF26(2)) & ""
      End If
      '完稿日
      If txtEP09(1) <> "" Then
         stCon = stCon & " and EP09>=" & DBDATE(txtEP09(1)) & ""
      End If
      If txtEP09(2) <> "" Then
         stCon = stCon & " and EP09<=" & DBDATE(txtEP09(2)) & ""
      End If
      
      If stCon = "" Then
         MsgBox "輸入條件不可全部空白...", vbCritical
         txtCaseNo(2).SetFocus
         GoTo flgErr
      End If
      
      '依條件決定主表
      If txtCaseNo(2) <> "" Or txtCP14 <> "" Then
         stCon = stCon & " AND EP02(+)=CP09 AND TF01(+)=CP09"
      ElseIf txtTF26(1) <> "" Then
         stCon = stCon & " AND CP09(+)=TF01 AND EP02(+)=CP09"
      ElseIf txtEP09(1) <> "" Then
         stCon = stCon & " AND CP09(+)=EP02 AND TF01(+)=CP09"
      Else
         stCon = stCon & " AND EP02(+)=CP09 AND TF01(+)=CP09"
      End If
      
   '未完稿
   Else
      stCon = stCon & " AND EP02(+)=CP09 AND TF01(+)=CP09"
      
      stCon = stCon & " and cp158=0 and cp159=0 and ep09 is null"
      '達完稿期限
      If idx = 2 Then
         stCon = stCon & " and tf26<=to_char(sysdate,'yyyymmdd')"
      '未收達
      ElseIf idx = 3 Then
         stCon = stCon & " and tf32 is null"
      End If
   End If
   
   stCon = stCon & " and pa09 in ('011','231')" '目前限定日德案件，若要增加其他語種，翻譯費設定及相關程式也要改
   stCon = stCon & " and cp10 in (" & CaseMapOut & ",927)" '只抓新案及其他翻譯--郭
   
   stSQL = "SELECT '' V" & _
      ", CP01||'-'||CP02||DECODE(CP03||CP04,'000','',CP03||CP04) 本所案號" & _
      ", NVL(CPM03,CP10) 案件性質" & _
      ", CP14||' '||S1.ST02 承辦人" & _
      ", SQLDATET(CP157) 給案日" & _
      ", decode(tf27,'5',TF23) 中文字數" & _
      ", decode(tf27,'5',null,TF23) 原文字數" & _
      ", SQLDATET(TF26) 完稿期限" & _
      ", SQLDATET(CP06) 本所期限" & _
      ", SQLDATET(TF32) 收達日" & _
      ", SQLDATET(EP09) 完稿日, SQLDATET(EP07) 會稿日, SQLDATET(CP27) 發文日" & _
      ", TF04 分配點數" & _
      ", TF36 備註" & _
      ", NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ", CP09" & _
      " FROM CASEPROGRESS, ENGINEERPROGRESS, CASEPROPERTYMAP, STAFF S1,TRANSFEE,Staff_IdMap,STAFF S2, PATENT" & _
      " WHERE cp01='" & arrCaseNo(1) & "' and cp12 not like 'F%' and cp14 like 'F%'" & stCon & _
      " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND S1.ST01(+)=CP14 and sim02(+)=cp14 and S2.st01(+)=sim01" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04"
      
   stSQL = stSQL & " ORDER BY cp157,2"
   
   intI = 1
   Set rstGrid = ClsLawReadRstMsg(intI, stSQL)
   
   If intI = 1 Then
      txtCaseNo(1) = arrCaseNo(1)
      txtCaseNo(2) = arrCaseNo(2)
      txtCaseNo(3) = arrCaseNo(3)
      txtCaseNo(4) = arrCaseNo(4)
      
   Else
      ShowNoData
   End If
   
   Set MSHFlexGrid1.Recordset = rstGrid
   SetGridHead
    
flgErr:
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Set rstGrid = Nothing
   
End Sub

Private Sub cmdSearch_Click(index As Integer)

    SetGrid index
'    If Me.MSHFlexGrid1.Rows = 2 And Me.Visible = True Then
'        MSHFlexGrid1.row = 1
'        GridClick MSHFlexGrid1, intLastRow, 0
'        cmdok_Click
'   End If

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   ClearGrid
   lblCP14N = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm050110 = Nothing
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nCol As Long, nRow As Long
   getGrdColRow MSHFlexGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MSHFlexGrid1.col = nCol
   MSHFlexGrid1.row = nRow
   If MSHFlexGrid1.row < 1 And Me.MSHFlexGrid1.Text <> "v" Then
      If MSHFlexGrid1.Text = "點數" Then
         If m_blnColOrderAsc = True Then
            MSHFlexGrid1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            MSHFlexGrid1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            MSHFlexGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            MSHFlexGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

Private Sub MSHFlexGrid1_SelChange()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK.SetFocus
End Sub

Private Sub txtCaseNo_GotFocus(index As Integer)
   TextInverse txtCaseNo(index)
   CloseIme
End Sub

Private Sub txtCaseNo_KeyPress(index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCP14_Change()
   lblCP14N = ""
   If Len(txtCP14) = 5 Then
      If Left(txtCP14, 1) = "F" Then
         If ClsPDGetStaffN(txtCP14.Text, strExc(1)) = True Then
            lblCP14N = strExc(1)
         End If
      End If
   End If
End Sub

Private Sub txtCP14_GotFocus()
   TextInverse txtCP14
   CloseIme
End Sub

Private Sub txtCP14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCP14_Validate(Cancel As Boolean)
   If txtCP14 <> "" Then
      If Left(txtCP14, 1) <> "F" Then
         MsgBox "必須為F編號！", vbCritical
         Cancel = True
      End If
   End If
End Sub

Private Sub txtEP09_GotFocus(index As Integer)
   TextInverse txtEP09(index)
   CloseIme
End Sub

Private Sub txtTF26_GotFocus(index As Integer)
   TextInverse txtTF26(index)
   CloseIme
End Sub
