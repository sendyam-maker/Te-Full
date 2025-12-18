VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210106 
   BorderStyle     =   1  '單線固定
   Caption         =   "簽收資料查詢"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9435
   Begin VB.CommandButton cmdFind 
      Caption         =   "搜尋(&Q)"
      Height          =   315
      Left            =   4725
      TabIndex        =   7
      Top             =   1410
      Width           =   765
   End
   Begin VB.TextBox txtCustNo 
      Height          =   300
      Index           =   1
      Left            =   2565
      MaxLength       =   9
      TabIndex        =   5
      Top             =   1070
      Width           =   975
   End
   Begin VB.TextBox txtCustNo 
      Height          =   300
      Index           =   0
      Left            =   1260
      MaxLength       =   9
      TabIndex        =   4
      Top             =   1070
      Width           =   975
   End
   Begin VB.CheckBox chkUnRec 
      Caption         =   "永遠含已收款未繳收據資料"
      Height          =   225
      Left            =   225
      TabIndex        =   8
      Top             =   1770
      Value           =   1  '核取
      Width           =   3660
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "簽收確認(&C)"
      Height          =   400
      Left            =   7605
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   1590
      Width           =   1695
   End
   Begin VB.TextBox txtOffice 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   0
      Top             =   53
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   300
      Index           =   0
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   2
      Top             =   731
      Width           =   960
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   300
      Index           =   1
      Left            =   2565
      MaxLength       =   7
      TabIndex        =   3
      Top             =   731
      Width           =   960
   End
   Begin VB.TextBox txtSales 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   1
      Top             =   392
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7620
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8520
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   90
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3765
      Left            =   135
      TabIndex        =   12
      Top             =   2100
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   6641
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   2280
      TabIndex        =   19
      Top             =   392
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCuName 
      Height          =   330
      Left            =   1530
      TabIndex        =   6
      Top             =   1380
      Width           =   3120
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "5503;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCuNam 
      Caption         =   "客戶中文名稱："
      Height          =   180
      Left            =   225
      TabIndex        =   18
      Top             =   1410
      Width           =   1275
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2370
      X2              =   2490
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2370
      X2              =   2490
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "客戶代號"
      Height          =   180
      Left            =   225
      TabIndex        =   17
      Top             =   1070
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "所別"
      Height          =   180
      Left            =   225
      TabIndex        =   16
      Top             =   53
      Width           =   900
   End
   Begin VB.Label lblZone 
      Caption         =   "（1北所 2中所 3南所 4高所）"
      Height          =   180
      Left            =   2295
      TabIndex        =   15
      Top             =   53
      Width           =   4590
   End
   Begin VB.Label Label2 
      Caption         =   "繳款日期"
      Height          =   180
      Left            =   225
      TabIndex        =   14
      Top             =   731
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "智權人員"
      Height          =   180
      Left            =   225
      TabIndex        =   13
      Top             =   392
      Width           =   900
   End
End
Attribute VB_Name = "frm210106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/03 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、txtCuName、lblSalesName
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit
'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Public m_strCustCode As String
Public m_blnOneRec As Boolean
Dim stIdList As String 'Added by Lydia 2019/08/08 使用者清單(取代Me.tag)
Dim m_PrevForm As Form 'Added by Lydia 2021/07/27 記錄呼叫的表單; 智權-調整財務系統(20200909) 'Memo by Lydia 2021/08/27 上線

'Added by Lydia 2021/07/27 記錄前一畫面的表單名稱
Public Sub SetParent(ByVal pForm As Form)
    Set m_PrevForm = pForm
End Sub

'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim ii As Integer
   With grdDataList
      .Visible = False
      If p_bolHeaderOnly = False Then
         .Clear
         .Rows = 3: .Cols = 16: .FixedRows = 2: .FixedCols = 1
      End If
      .row = 0
      .col = 0: .ColWidth(.col) = 400: .Text = "V"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignmentFixed(.col) = flexAlignCenterCenter
      .col = 1: .ColWidth(.col) = 1100: .Text = "繳款日期"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignCenterCenter
      .col = 2: .ColWidth(.col) = 1100: .Text = "繳收據日"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignCenterCenter
      .col = 3: .ColWidth(.col) = 1400: .Text = "客戶名稱"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      .col = 4: .ColWidth(.col) = 1200: .Text = "扣繳金額"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      .col = 5: .ColWidth(.col) = 1200: .Text = "票額"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      .col = 6: .ColWidth(.col) = 1200: .Text = "現金"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      .col = 7: .ColWidth(.col) = 1200: .Text = "銀存"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      .col = 8: .ColWidth(.col) = 1200: .Text = "暫收"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      .col = 9: .ColWidth(.col) = 1200: .Text = "其他"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      .col = 10: .ColWidth(.col) = 3500: .Text = "收據號碼"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      .col = 11: .ColWidth(.col) = 2450: .Text = "備註"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      .col = 12: .ColWidth(.col) = 1150: .Text = "簽收單號"
      .CellAlignment = flexAlignLeftCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      .col = 13: .ColWidth(.col) = 0
      .col = 14: .ColWidth(.col) = 0
      .col = 15: .ColWidth(.col) = 0
      
      .row = 1
      .col = 2: .Text = "合計："
      .CellAlignment = flexAlignRightCenter
      For ii = 3 To 9
         .col = ii: .Text = ""
         .CellAlignment = flexAlignRightCenter
      Next
      For ii = 0 To .Cols - 1
         .col = ii
         .CellBackColor = &H90EE90
      Next
      '不看合計
      .RowHeight(1) = 0
      .Refresh
      .Visible = True
   End With
End Sub

Private Function doQuery() As Boolean

   Dim stCon As String, stUnion As String
   
   stCon = ""
   '所別
   If txtOffice <> "" Then
      stCon = " AND EXISTS(SELECT * FROM STAFF WHERE ST01=A2303 AND ST06='" & txtOffice & "')"
   End If
   '智權人員
   If txtSales <> "" Then
      stCon = stCon & " AND A2303='" & txtSales & "'"
   End If
   
   stUnion = ""
   If chkUnRec.Value = vbChecked Then
      stUnion = " UNION SELECT A2302,A2308,CU04,A2307,A2306,A2317,A2318,A2319,A2320,A2309,A2310,A2301,A2303,A2321,DECODE(A2308,NULL,' ',A2302) S1" & _
      " FROM ACC230, CUSTOMER WHERE CU01(+)=SUBSTR(A2304,1,8) AND CU02(+)=SUBSTR(A2304,9,1) AND A2308 IS NULL" & stCon
      'Added by Morgan 2014/10/23 排除手動更新簽收確認日期者(未輸繳款直接輸收款)
      stUnion = stUnion & " and (a2321 is null or to_char(a2321,'hh24miss')<>'000000') "
   End If
   
   '點數結算日
   If txtCloseDate(0) <> "" Then
      stCon = stCon & " AND A2302 >= " & txtCloseDate(0)
   End If
   If txtCloseDate(1) <> "" Then
      stCon = stCon & " AND A2302 <= " & txtCloseDate(1)
   End If
   
   If txtCustNo(0) <> "" Then
      stCon = stCon & " AND A2304>='" & txtCustNo(0) & "'"
   End If
   
   If txtCustNo(1) <> "" Then
      stCon = stCon & " AND A2304<='" & txtCustNo(1) & "'"
   End If

   
On Error GoTo ErrHnd
   
   strSql = "SELECT A2302,A2308,CU04,A2307,A2306,A2317,A2318,A2319,A2320,A2309,A2310,A2301,A2303,A2321,DECODE(A2308,NULL,' ',A2302) S1" & _
      " FROM ACC230, CUSTOMER WHERE CU01(+)=SUBSTR(A2304,1,8) AND CU02(+)=SUBSTR(A2304,9,1)" & stCon & stUnion & _
      " ORDER BY S1,A2301"
      
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         Set grdDataList.Recordset = AdoRecordSet3.Clone
         Call SetDataListWidth(True)
         Call Calculate
      Else
         MsgBox "無符合資料！", vbInformation
      End If
   End With
   
   doQuery = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function
Private Sub Calculate()
   Dim ii As Integer, dblSum(4 To 10) As Double, dblSum2 As Double, jj As Integer
   With grdDataList
      .Visible = False
      For ii = 2 To .Rows - 1
         If .TextMatrix(ii, 14) <> "" Then
            .TextMatrix(ii, 0) = "V"
         End If
         .TextMatrix(ii, 1) = Format(.TextMatrix(ii, 1), "###/##/##")
         .TextMatrix(ii, 2) = Format(.TextMatrix(ii, 2), "###/##/##")
         For jj = 4 To 10
            dblSum(jj) = dblSum(jj) + Val(.TextMatrix(ii, jj))
            .TextMatrix(ii, jj) = Format(.TextMatrix(ii, jj), "###,###,###.00")
         Next
      Next ii
      For jj = 4 To 10
         .TextMatrix(1, jj) = Format(dblSum(jj), "###,###,###.00")
      Next
      .Visible = True
   End With
End Sub
Private Sub cmdExit_Click()
 Unload Me
End Sub
Private Function SaveData(ByRef p_iRec As Integer) As Boolean

   Dim ii As Integer, iCnt As Integer, jj As Integer
   
On Error GoTo ErrHnd
   p_iRec = 0
   cnnConnection.BeginTrans
   With grdDataList
      For ii = 2 To .Rows - 1
         If .TextMatrix(ii, 0) = "V" And .TextMatrix(ii, 14) = "" Then
            p_iRec = p_iRec + 1
            .row = ii: .col = 0: .CellForeColor = vbBlack
            .TextMatrix(ii, 14) = "V"
            strSql = "UPDATE ACC230 SET A2321=SYSDATE WHERE A2301='" & .TextMatrix(ii, 12) & "' AND A2321 IS NULL"
            cnnConnection.Execute strSql
            For jj = 1 To .Cols - 1
               .col = jj
               .CellBackColor = &H80000018
            Next
         End If
      Next
   End With
   cnnConnection.CommitTrans
   SaveData = True
   
ErrHnd:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function
Private Sub CmdSave_Click()
   Dim iRec As Integer
   If MsgBox("確定簽收資料無誤？", vbYesNo + vbDefaultButton2) = vbYes Then
      If SaveData(iRec) = True Then
         If iRec = 0 Then
            MsgBox "無確認資料可更新！", vbInformation
         Else
            MsgBox "共確認 " & iRec & " 筆簽收資料！", vbInformation
         End If
      End If
   End If
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   If ConstrainCheck = True Then
      Call SetDataListWidth
      Call doQuery
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

   Dim stST05  As String, stST03 As String
   
   MoveFormToCenter Me
   Call SetDataListWidth
   
   stST05 = PUB_GetST05(strUserNum)
   stST03 = PUB_GetST03(strUserNum)
   txtOffice = pub_strUserOffice
   '財務人員只可查詢
   'MODIFY BY SONIA 2015/6/1 中所加CM
   If stST03 = "M31" Or stST05 = "CM" Or stST05 = "C1" Or stST05 = "NM" Or stST05 = "KM" Then
      cmdSave.Enabled = False
      txtSales.Enabled = True
      '財務處可看全所
      If stST03 = "M31" Then
         txtOffice.Enabled = True
      End If
   Else
      stIdList = strUserNum
      'Remove by Lydia 2017/01/26 改到main直接呼叫智權人員登入
      ''密碼空白的才彈視窗改Me.Tag值
      'If strPassWord = "" Then
      '   frm210106_1.setCaller Me
      '   frm210106_1.Show vbModal
      'End If
     
      'Memo by Lydia 2019/08/08 stIdList使用者清單(取代Me.tag)
      txtSales = stIdList
      '小真可確認總經理的簽收
      If stIdList = "65001" Then
         txtSales = "68001"
         txtSales.Enabled = True
         stIdList = stIdList & ",68001"
      'add by sonia 2014/6/9 美珍可確認林總經理的簽收
      ElseIf stIdList = "77027" Then
         txtSales = "94007"
         txtSales.Enabled = True
         stIdList = stIdList & ",94007"
      'end 2014/6/9
'cancel by sonia 2019/4/12阿蓮調職取消此權限
'      '阿蓮可確認杜副總的簽收
'      ElseIf stIdlist = "74028" Then
'         txtSales = "68006"
'         txtSales.Enabled = True
'         stIdlist = stIdlist & ",68006"
'end 2019/4/12
      'Added by Lydia 2019/08/08 創新業務部各組人員可以看該組所有人資料
      ElseIf stST03 <> "W00" And Left(stST03, 1) = "W" Then
         txtSales.Enabled = True
         stIdList = Replace(PUB_GetSalesList(strUserNum), "'", "")
      End If
   End If
   txtCloseDate(0) = strSrvDate(2)
   txtCloseDate(1) = strSrvDate(2)
   
   cmdSave.Visible = False 'Added by Morgan 2014/3/5 取消確認功能,改到繳款資料輸入作業
   
   Call Pub_AddPersonRec("frm210106") 'Added by Lydia 2019/06/27
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Added by Lydia 2021/07/27 回前一畫面
   If TypeName(m_PrevForm) <> "Nothing" Then
       m_PrevForm.Show
   End If
   'end 2021/07/27
   MenuEnabled
   Set frm210106 = Nothing
End Sub

Private Sub grdDataList_SelChange()
   Dim ii As Integer
   With grdDataList
      If .MouseRow > 1 And .MouseRow < .Rows Then
         .Visible = False
         .row = grdDataList.MouseRow
         '控制未確認過才可點選
         If .TextMatrix(.row, 14) = "" Then
            .col = 0
            If .Text = "V" Then
               .Text = ""
               For ii = 1 To .Cols - 1
                  .col = ii
                  .CellBackColor = &H80000018
               Next
            Else
               .CellForeColor = vbRed
               .Text = "V"
               For ii = 1 To .Cols - 1
                  .col = ii
                  .CellBackColor = &HFFC0C0
               Next
            End If
         End If
      End If
      .Visible = True
   End With
End Sub
'Remove by Morgan 2005/4/21 改成共同查詢模式
'Private Sub grdDataList_DblClick()
'   With grdDataList
'      If .Row > 1 And .Row < .Rows Then
'         If .TextMatrix(.Row, 14) = "" Then
'            If .TextMatrix(.Row, 0) = "V" Then
'               .TextMatrix(.Row, 0) = ""
'            Else
'               .col = 0
'               .CellForeColor = vbRed
'               .TextMatrix(.Row, 0) = "V"
'            End If
'         End If
'      End If
'   End With
'End Sub

Private Sub txtCustNo_GotFocus(Index As Integer)
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtCustNo(Index).IMEMode = 2
   CloseIme
   TextInverse txtCustNo(Index)
   If Index = 1 And Len(txtCustNo(0)) = 9 Then
      txtCustNo(Index) = Left(txtCustNo(0), 6) & "ZZZ"
      txtCustNo(Index).SelStart = 6
      txtCustNo(Index).SelLength = 3
   End If
End Sub

Private Sub txtCustNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCustNo_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      If Len(txtCustNo(0)) = 6 Then
         'Modify By Sindy 2014/8/11 999=>ZZZ
         'txtCustNo(1) = txtCustNo(0) & "999"
         txtCustNo(1) = txtCustNo(0) & "ZZZ"
         txtCustNo(0) = txtCustNo(0) & "000"
      ElseIf Len(txtCustNo(0)) = 9 Then
         txtCustNo(1) = txtCustNo(0)
      End If
   Else
      If txtCustNo(1) <> "" Then
         If txtCustNo(0) = "" Then
            MsgBox "請先輸入起始客戶代碼！", vbExclamation
            Cancel = True
         ElseIf Len(txtCustNo(1)) <> 9 Then
            MsgBox "客戶代碼需輸入九碼！", vbExclamation
            Cancel = True
         ElseIf Left(txtCustNo(1), 6) <> Left(txtCustNo(0), 6) Then
            MsgBox "客戶代碼前六碼需相同！", vbCritical
            Cancel = True
         ElseIf txtCustNo(1) < txtCustNo(0) Then
            MsgBox "客戶代碼區間輸入錯誤！", vbCritical
            Cancel = True
         End If
      End If
   End If
End Sub

Private Sub txtOffice_GotFocus()
   TextInverse txtOffice
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtOffice.IMEMode = 2
   CloseIme
End Sub

Private Sub txtOffice_KeyPress(KeyAscii As Integer)
   If (KeyAscii < Asc("1") Or KeyAscii > Asc("4")) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtSales_Change()
   If Len(txtSales) > 4 Then
      lblSalesName = StaffQuery(txtSales)
   Else
      lblSalesName = ""
   End If
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtSales.IMEMode = 2
   CloseIme
End Sub
Private Sub txtCloseDate_GotFocus(Index As Integer)
   TextInverse txtCloseDate(Index)
   'edit by nickc 2007/06/11  切換輸入法改用API'
   'txtCloseDate(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtCloseDate_Validate(Index As Integer, Cancel As Boolean)
   If txtCloseDate(Index) <> "" Then
      If ChkDate(txtCloseDate(Index)) = False Then
         Cancel = True
         txtCloseDate(Index).SetFocus
         txtCloseDate_GotFocus Index
      End If
   End If
End Sub

Private Function ConstrainCheck() As Boolean
   Dim bolCancel As Boolean
   'Add by Morgan 2008/5/28
   If txtSales.Enabled = True Then
      If stIdList <> "" Then
         If txtSales = "" Then
            MsgBox "智權人員不可空白！"
            txtSales.SetFocus
            Exit Function
         ElseIf InStr(stIdList, txtSales) = 0 Then
            MsgBox "智權人員輸入錯誤或無權限！"
            txtSales.SetFocus
            txtSales_GotFocus
            Exit Function
         End If
      End If
   End If
   
   If txtCloseDate(0) = "" Then
      MsgBox "請輸入繳款日期起日！", vbExclamation
      txtCloseDate(0).SetFocus
      txtCloseDate_GotFocus (0)
      Exit Function
   Else
      Call txtCloseDate_Validate(0, bolCancel)
      If bolCancel = True Then
         Exit Function
      End If
   End If
   
   If txtCloseDate(1) = "" Then
      MsgBox "請輸入繳款日期迄日！", vbExclamation
      txtCloseDate(1).SetFocus
      txtCloseDate_GotFocus (1)
      Exit Function
   Else
      Call txtCloseDate_Validate(1, bolCancel)
      If bolCancel = True Then
         Exit Function
      End If
   End If
   
   'Add by Morgan 2005/5/19
   Call txtCustNo_Validate(1, bolCancel)
   If bolCancel = True Then
      Exit Function
   End If
   
   ConstrainCheck = True
End Function

'Add By Sindy 2010/11/26
Private Sub txtSales_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Private Sub txtCuName_GotFocus()
   'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
   cmdSearch.Default = False: cmdFind.Default = True
   
   TextInverse Me.txtCuName
   OpenIme
End Sub

'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
Private Sub txtCuName_LostFocus()
   cmdSearch.Default = True: cmdFind.Default = False
End Sub
'Added by Lydia 2015/06/18 +客戶中文名稱查詢
Private Sub cmdFind_Click()
   If Me.txtCuName.Text = "" Then
      MsgBox "請輸入客戶中文名稱的關鍵字!!!", vbExclamation + vbOKOnly
      Me.txtCuName.SetFocus
      txtCuName_GotFocus
      Exit Sub
   End If
   frm090801_1.m_strCustChnName = Me.txtCuName.Text
   frm090801_1.lblName.Caption = Me.txtCuName.Text
   m_blnOneRec = False
   m_strCustCode = ""
   If frm090801_1.StrMenu = True Then
      If frm090801_1.m_blnOneRec = False Then
         frm090801_1.Show vbModal
      End If
      m_blnOneRec = frm090801_1.m_blnOneRec
      m_strCustCode = frm090801_1.m_strCustCode
      Unload frm090801_1
   Else
      Unload frm090801_1
   End If
   If m_blnOneRec = True And m_strCustCode <> "" Then
      Me.txtCustNo(0).Text = m_strCustCode
      Me.txtCustNo(1).Text = IIf(Right(m_strCustCode, 3) = "000", Mid(m_strCustCode, 1, 6) & "ZZZ", IIf(Right(m_strCustCode, 1) = "0", Mid(m_strCustCode, 1, 8) & "Z", m_strCustCode))
      Me.txtCuName.Text = GetCustomerName(m_strCustCode)
   End If
   'Added by Lydia 2015/06/25 輸入客戶名稱按Enter自動執行搜尋後,執行查詢
   If Me.txtCustNo(0).Text <> "" And Me.txtCustNo(1).Text <> "" Then
      Call cmdSearch_Click
   End If
End Sub
