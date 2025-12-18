VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm110103_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "閉卷"
   ClientHeight    =   5676
   ClientLeft      =   96
   ClientTop       =   996
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5676
   ScaleWidth      =   9312
   Begin VB.CheckBox Check2 
      Caption         =   "有年費期限"
      Height          =   300
      Left            =   2040
      TabIndex        =   10
      Top             =   144
      Width           =   1668
   End
   Begin VB.CheckBox Check1 
      Caption         =   "不含已閉卷"
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   144
      Width           =   1668
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   8388
      TabIndex        =   7
      Top             =   70
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6432
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   7260
      TabIndex        =   5
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "所有案件(&L)"
      Height          =   400
      Index           =   3
      Left            =   5304
      TabIndex        =   4
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "全部選取(&A)"
      Height          =   400
      Left            =   4176
      TabIndex        =   3
      Top             =   70
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4692
      Left            =   72
      TabIndex        =   8
      Top             =   960
      Width           =   9156
      _ExtentX        =   16150
      _ExtentY        =   8276
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
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
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblNo 
      Height          =   252
      Left            =   984
      TabIndex        =   2
      Top             =   576
      Width           =   852
   End
   Begin VB.Label lblName 
      Height          =   252
      Left            =   1848
      TabIndex        =   1
      Top             =   576
      Width           =   7188
   End
   Begin VB.Label lblTitle 
      Caption         =   "申請人："
      Height          =   250
      Left            =   96
      TabIndex        =   0
      Top             =   576
      Width           =   972
   End
End
Attribute VB_Name = "frm110103_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/10 改成Form2.0(grdDataList)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'sonia 2010/8/19 日期欄已修改
Option Explicit
'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'intChoose被選擇之收文號總數
Dim strCaseCodeC() As String, intNowCaseCodeC As Integer
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
'edit by nickc 2007/02/05 不用 dll 了
'Dim obj011 As Object
Dim strSql As String
Public intChoose As Integer
'Added by Morgan 2025/7/3
Dim RsNick As New ADODB.Recordset
Dim strLastSortType As String, strLastSortCol As String
Dim bolRechoose As Boolean
'end 2025/7/3


Public Sub ReChoose(ByRef intNowCaseCode As Integer, ByRef strCaseCode() As String)
intNowCaseCodeC = intNowCaseCode
strCaseCodeC() = strCaseCode()
bolRechoose = True 'Added by Morgan 2025/7/3
End Sub

'Added by Morgan 2025/7/3
Private Sub Check1_Click()
   intChoose = 0
   GetCloseCaseData
End Sub
'Added by Morgan 2025/7/3
Private Sub Check2_Click()
   intChoose = 0
   GetCloseCaseData
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim i As Integer

Select Case Index
             Case 0 '確定
                        frm110103_3.intWhereComeFrom = 2
                        Set frm110103_3.mPrev01 = Me 'Add By Sindy 2015/2/13
                        frm110103_3.Show
                        Me.Hide
             Case 1, 2 '回前畫面, 結束
                        If Index = 2 Then
                           intLeaveKind = 0
                        Else
                           intLeaveKind = 1
                        End If
                        Unload Me
             Case 3 '所有案件
                       frm110103_4.lblTitle = lblTitle
                       frm110103_4.lblName = lblName
                       frm110103_4.LblNo = LblNo
                       frm110103_4.Show
                       Me.Hide
End Select
End Sub
Private Sub cmdSelectAll_Click()
Dim i As Integer

For i = 1 To grdDataList.Rows - 1
       grdDataList.TextMatrix(i, 0) = "ˇ"
Next
intChoose = grdDataList.Rows - 1
cmdok(0).Enabled = True
End Sub
Private Sub GetCloseCaseData()
Dim varSaveCursor, i As Integer, j   As Integer, k As Integer, L As Integer
' Dim RsNick As New ADODB.Recordset 'Removed by Morgan 2025/7/3 改全域

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
If RsNick.State = 1 Then
   RsNick.Close
End If
With RsNick
   .CursorLocation = adUseClient
   strSql = ReadCloseCaseRst(GetNewFagent(LblNo), frm110103_1.txtCaseField(0))
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      Set grdDataList.Recordset = RsNick
   End If
End With
CheckOC
If grdDataList.Recordset Is Nothing Then
   Screen.MousePointer = varSaveCursor
   Unload Me
   Exit Sub
End If
SetDataListVision grdDataList, True, True
intLastRow = 0
If grdDataList.Rows > 1 Then
   ShowBar grdDataList, intLastRow, 5
   'Modified by Morgan 2025/7/3
   'If intNowCaseCodeC < intChoose Then
   If intNowCaseCodeC < intChoose And bolRechoose Then
      bolRechoose = False
   'end 2025/7/3
      j = intNowCaseCodeC
      For i = 1 To grdDataList.Rows - 1
             For L = 0 To 3
                    If grdDataList.TextMatrix(i, 6 + L) <> strCaseCodeC(L, j) Then
                       Exit For
                    End If
             Next
             If L = 4 Then
                grdDataList.TextMatrix(i, 0) = "ˇ"
                k = k + 1
                j = j + 1
                If j = intChoose Then Exit For
             End If
      Next
   End If
   intChoose = k
   cmdok(3).Enabled = True
   cmdSelectAll.Enabled = True
Else
   intChoose = 0
   cmdok(3).Enabled = False
   cmdSelectAll.Enabled = False
End If
grdDataList.col = 0
grdDataList.ColWidth(0) = 200
grdDataList.col = 1
grdDataList.ColWidth(1) = 1500
grdDataList.col = 2
grdDataList.ColWidth(2) = 500
grdDataList.col = 3
grdDataList.ColWidth(3) = 3200
grdDataList.col = 4
grdDataList.ColWidth(4) = 1000
grdDataList.col = 5
grdDataList.ColWidth(5) = 1500
grdDataList.col = 6
grdDataList.ColWidth(6) = 0
grdDataList.col = 7
grdDataList.ColWidth(7) = 0
grdDataList.col = 8
grdDataList.ColWidth(8) = 0
grdDataList.col = 9
grdDataList.ColWidth(9) = 0
'Added by Lydia 2016/10/12 pa08 as E
grdDataList.col = 10
grdDataList.ColWidth(10) = 0

If intChoose = 0 Then
   cmdok(0).Enabled = False
Else
   cmdok(0).Enabled = True
End If
Screen.MousePointer = varSaveCursor
End Sub
Private Sub Form_Activate()
GetCloseCaseData
End Sub
Private Sub Form_Load()
MoveFormToCenter Me
SetDataListWidth
'Added by Lydia 2016/01/04 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub
Private Sub SetDataListWidth()
Dim varGridWidth() As Variant

varGridWidth = Array(300, 400, 550, 150, 150, 250, 3000, 1000, 2000)
SetGridDataListWidth grdDataList, varGridWidth()
blnOKtoShow = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
intChoose = 0
If intLeaveKind = 1 Then
   frm110103_1.Show
   frm110103_1.Cleartxt
ElseIf intLeaveKind = 0 Then
  Unload frm110103_1
End If
   'Add By Cheng 2002/07/18
   Set frm110103_2 = Nothing
End Sub
Private Sub grdDataList_GotFocus()
GridGotFocus grdDataList
End Sub
Private Sub grdDataList_LostFocus()
GridLostFocus grdDataList
End Sub
Private Sub grdDataList_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then grdDataList_SelChange
End Sub

'Added by Morgan 2025/7/3
Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim strSortCol As String, strSortType As String, strSortSQL As String
   With grdDataList
   If .MouseRow < 1 And .Rows > 1 Then
      If strLastSortType = "ASC" Then
         strSortType = "DESC"
      Else
         strSortType = "ASC"
      End If
      
      strSortCol = RsNick.Fields(.MouseCol).Name
      If strSortCol = strLastSortCol Then
         strSortSQL = strSortCol & " " & strSortType
      Else
         strSortSQL = strSortCol & " ASC"
         strSortType = "ASC"
      End If
         
      If strSortCol <> "本所案號" Then
         strSortSQL = strSortSQL & ",本所案號 ASC"
      End If
      RsNick.Sort = strSortSQL
      Set grdDataList.Recordset = RsNick
      
      strLastSortCol = strSortCol
      strLastSortType = strSortType
   End If
   End With
End Sub
'end 2025/7/3

Private Sub grdDataList_SelChange()
If grdDataList.TextMatrix(grdDataList.row, 0) = "ˇ" Then
   grdDataList.TextMatrix(grdDataList.row, 0) = ""
   intChoose = intChoose - 1
Else
   grdDataList.TextMatrix(grdDataList.row, 0) = "ˇ"
   intChoose = intChoose + 1
End If
If intChoose = 0 Then cmdok(0).Enabled = False Else cmdok(0).Enabled = True
End Sub
Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      ShowBar grdDataList, intLastRow, 8
      blnOKtoShow = True
   End If
End If
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Boolean
Select Case intIndex
             Case 0, 1
'                        If txtDate(intIndex) = "" Then
'                           CheckKeyIn = True
'                        Else
'                           If intPWhere <> 國外_CF Then
'                              If CheckIsTaiwanDate(txtDate(intIndex)) Then
'                                 CheckKeyIn = True
'                              End If
'                           Else
'                              If CheckIsDate(txtDate(intIndex)) Then
'                                 CheckKeyIn = True
'                              End If
'                           End If
'                        End If
'                        If CheckKeyIn = False Then Exit Function
'                        If txtDate(0) <> "" And txtDate(1) = "" And intIndex = 1 Then
'                           ShowMsg MsgText(9169)
'                           CheckKeyIn = False
'                        ElseIf txtDate(1) <> "" And Val(txtDate(0)) > Val(txtDate(1)) Then
'                           ShowMsg MsgText(9170)
'                           CheckKeyIn = False
'                        End If
End Select
End Function

'讀取閉卷申請人或代理人資料
Private Function ReadCloseCaseRst(ByRef strNo As String, ByRef strSystemKind As String) As String
Dim strSql As String, rsRecordset As New ADODB.Recordset
'edit by nickc 2007/02/06 不用 dll 了
'Dim objPublicData As Object, strSQLS(4) As String, strSystem(4) As String, i As Integer
Dim strSQLS(4) As String, strSystem(4) As String, i As Integer
   
On Error GoTo ErrHand
'edit by nickc 2007/02/06 不用 dll 了
'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
'If objPublicData.AnalysisSystemKindString(strSystemKind, strSystem()) = False Then GoTo err1
If ClsPDAnalysisSystemKindString(strSystemKind, strSystem()) = False Then GoTo err1
strNo = ChangeCustomerL(strNo)
'If Left(strNo, 1) = 代理人編號 Then
'   strNo = Left(strNo, 8)
'End If
If strSystem(0) <> "" Then
   'StrSQL = "select '' ˇ,s01 本所案號,s02 本所案號,s03 本所案號,s04 本所案號,s05 本所案號,s06 案件名稱,s07 申請國家,s08 申請案號,s09,s10,s11,s12 from ("
   'Modified by Lydia 2016/01/13  開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   'strSQLS(0) = "select '' ˇ,pa01||'-'||PA02||DECODE(PA03,'0',DECODE(PA04,'00','',NULL,'','-'),NULL,'','-')||REPLACE(PA03,'0',' ')||DECODE(PA03,'0',DECODE(PA04,'00','',NULL,'','-'),NULL,'','-')||replace(PA04,'00',' ') as 本所案號,pa57 as 閉卷,NVL(pa05,nvl(pa06,pa07)) AS 案件名稱,na03 AS 申請國家,pa11 AS 申請案號,pa01 AS A,pa02 AS B,pa03  AS C,pa04 AS D from patent,nation where pa09=na01(+) "
   'Modified by Lydia 2016/10/12 +pa08
   strSQLS(0) = "select '' ˇ,pa01||'-'||PA02||DECODE(PA03,'0',DECODE(PA04,'00','',NULL,'','-'),NULL,'','-')||REPLACE(PA03,'0',' ')||DECODE(PA03,'0',DECODE(PA04,'00','',NULL,'','-'),NULL,'','-')||replace(PA04,'00',' ') as 本所案號,pa57 as 閉卷,NVL(pa05,nvl(pa06,pa07)) AS 案件名稱,na03 AS 申請國家,pa11 AS 申請案號,pa01 AS A,pa02 AS B,pa03  AS C,pa04 AS D,pa08 AS E from patent f0,nation where pa09=na01(+) "
   If Left(strNo, 1) = 客戶編號 Then
      strSQLS(0) = strSQLS(0) + "and (pa26=" + CNULL(strNo) + " or pa27=" + CNULL(strNo) + " or pa28=" + CNULL(strNo) + " or pa29=" + CNULL(strNo) + " or pa30=" + CNULL(strNo) + ")"
   Else
      strSQLS(0) = strSQLS(0) + "and pa75=" + CNULL(strNo)
   End If
   strSQLS(0) = strSQLS(0) + " and pa01 in (" + strSystem(0) + ") "
   'Added by Lydia 2016/01/13
   If FMP2open Then
      'Modified by Morgan 2021/5/4
      'strSQLS(0) = strSQLS(0) + FMP2openSQL
      'strSQLS(0) = Replace(strSQLS(0), "f0.CP", "f0.PA")
      If InStr(strSystem(0), "'P'") > 0 Then
         strSQLS(0) = strSQLS(0) & " and PA01<>'P' union " & strSQLS(0) & " and PA01='P' " & FMP2openSQL
         strSQLS(0) = Replace(strSQLS(0), "f0.CP", "f0.PA")
      End If
   End If
   '2016/01/13 end
End If
If strSystem(1) <> "" Then
   'Modified by Lydia 2016/10/12 +' ' AS E
   strSQLS(1) = "select '' ˇ,tm01||'-'||decode(tm01," + CNULL(馬德里案) + ",substr(tm02,1,5),tm02)||'-'||decode(tm01," + CNULL(馬德里案) + ",replace(substr(tm02,6,1),'0',' '),tm02)||DECODE(TM03,'0',DECODE(TM04,'00','',NULL,'','-'),NULL,'','-')||replace(tm03,'0',' ')||DECODE(TM03,'0',DECODE(TM04,'00','',NULL,'','-'),NULL,'','-')||replace(tm04,'00',' ') as 本所案號,tm29 as 閉卷,nvl(tm05,nvl(tm06,tm07)) AS 案件名稱,na03 AS 申請國家,tm12 AS 申請案號,tm01 AS A,tm02 AS B,tm03 AS C,tm04 AS D,' ' AS E from trademark,nation where tm10=na01(+) "
   If Left(strNo, 1) = 客戶編號 Then
      'Modify By Sindy 2011/2/18 增加tm78,tm79,tm80,tm81
      'strSQLS(1) = strSQLS(1) + "and tm23=" + CNULL(strNo)
      strSQLS(1) = strSQLS(1) + "and (tm23=" + CNULL(strNo) + " or tm78=" + CNULL(strNo) + " or tm79=" + CNULL(strNo) + " or tm80=" + CNULL(strNo) + " or tm81=" + CNULL(strNo) + ")"
   Else
      strSQLS(1) = strSQLS(1) + "and tm44=" + CNULL(strNo)
   End If
   strSQLS(1) = strSQLS(1) + " and tm01 in (" + strSystem(1) + ")"
End If
If strSystem(2) <> "" Then
   'Modified by Lydia 2016/10/12 + ' ' AS E
   strSQLS(2) = "select '' ˇ,lc01||'-'||LC02||DECODE(LC03,'0',DECODE(LC04,'00','',NULL,'','-'),NULL,'','-')||REPLACE(lc03,'0',' ')||DECODE(LC03,'0',DECODE(LC04,'00','',NULL,'','-'),NULL,'','-')||replace(lc04,'00',' ') as 本所案號,lc08 as 閉卷,nvl(lc05,nvl(lc06,lc07)) AS 案件名稱,na03 AS 申請國家,'' AS 申請案號,lc01 AS A,lc02 AS B,lc03 AS C,lc04 AS D,' ' AS E from lawcase,nation where lc15=na01(+) "
   If Left(strNo, 1) = 客戶編號 Then
      'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46
      'strSQLS(2) = strSQLS(2) + "and lc11=" + CNULL(strNo)
      strSQLS(2) = strSQLS(2) + "and (lc11=" + CNULL(strNo) + " or lc43=" + CNULL(strNo) + " or lc44=" + CNULL(strNo) + " or lc45=" + CNULL(strNo) + " or lc46=" + CNULL(strNo) + ")"
   Else
      strSQLS(2) = strSQLS(2) + "and lc22=" + CNULL(strNo)
   End If
   strSQLS(2) = strSQLS(2) + " and lc01 in (" + strSystem(2) + ")"
End If
If strSystem(3) <> "" Then
   'Modified by Lydia 2016/10/12 + ' ' AS E
   strSQLS(3) = "select '' ˇ,hc01||'-'||hc02||DECODE(HC03,'0',DECODE(HC04,'00','',NULL,'','-'),NULL,'','-')||replace(hc03,'0',' ')||DECODE(HC03,'0',DECODE(HC04,'00','',NULL,'','-'),NULL,'','-')||replace(hc04,'00',' ') as 本所案號,hc09 as 閉卷,hc06 AS 案件名稱,'' AS 申請國家,'' AS 申請案號,hc01 AS A,hc02 AS B,hc03 AS C,hc04 AS D,' ' AS E from hirecase where hc01 in (" + strSystem(3) + ") "
   'Add By Sindy 2011/2/18 增加HC24,HC25,HC26,HC27
   If Left(strNo, 1) = 客戶編號 Then
      strSQLS(3) = strSQLS(3) + "and (hc05=" + CNULL(strNo) + " or hc24=" + CNULL(strNo) + " or hc25=" + CNULL(strNo) + " or hc26=" + CNULL(strNo) + " or hc27=" + CNULL(strNo) + ")"
   End If
   '2011/2/18 End
End If
If strSystem(4) <> "" Then
   'Modified by Lydia 2016/10/12 + ' ' AS E
   strSQLS(4) = "select '' ˇ,sp01||'-'||sp02||DECODE(SP03,'0',DECODE(SP04,'00','',NULL,'','-'),NULL,'','-')||replace(sp03,'0',' ')||DECODE(SP03,'0',DECODE(SP04,'00','',NULL,'','-'),NULL,'','-')||replace(sp04,'00',' ') as 本所案號,sp15 as 閉卷,nvl(sp05,nvl(sp06,sp07)) AS 案件名稱,na03 AS 申請國家,sp11 AS 申請案號,sp01 AS A,sp02 AS B,sp03 AS C,sp04 AS D,' ' AS E from servicepractice,nation where sp09=na01(+) "
   If Left(strNo, 1) = 客戶編號 Then
      'Modify By Sindy 2011/2/18 增加sp65,sp66
      strSQLS(4) = strSQLS(4) + "and (sp08=" + CNULL(strNo) + " or sp58=" + CNULL(strNo) + " or sp59=" + CNULL(strNo) + " or sp65=" + CNULL(strNo) + " or sp66=" + CNULL(strNo) + ")"
   Else
      strSQLS(4) = strSQLS(4) + "and sp26=" + CNULL(strNo)
   End If
   strSQLS(4) = strSQLS(4) + " and sp01 in (" + strSystem(4) + ")"
End If
'StrSQL = "select '' ˇ,s01 本所案號,s02 本所案號,s03 本所案號,s04 本所案號,s05 本所案號,s06 案件名稱,s07 申請國家,s08 申請案號,s09,s10,s11,s12 from ("
For i = 0 To 4
       If strSQLS(i) <> "" Then
          strSql = strSql + strSQLS(i) & " UNION "
       End If
Next
strSql = Left(strSql, Len(strSql) - 7)

'Added by Morgan 2025/7/3
If Check2.Value = vbChecked Then
   strSql = "select * from (" & strSql & ") where 閉卷 is null and exists(select * from nextprogress where np02=A and np03=B and np04=C and np05=D and np07 in ('605','606','607') and np06 is null)"
ElseIf Check1.Value = vbChecked Then
   strSql = "select * from (" & strSql & ") where  閉卷 is null"
End If
'end 2025/7/3
strSql = strSql + " order by A,B,C,D"
'Set ReadCloseCaseRst = objPublicData.ReadRst(StrSQL)
ReadCloseCaseRst = strSql
err1:
'edit by nickc 2007/02/06 不用 dll 了
'Set objPublicData = Nothing
Exit Function
ErrHand:
'edit by nickc 2007/02/06 不用 dll 了
'Set objPublicData = Nothing
End Function

