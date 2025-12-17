Attribute VB_Name = "basPrint"
Option Explicit
'FCT中要交由T執行之項目
Global Const conFCTtoT = "異議,評定,撤銷,答辯"
Global strLastCommandText As String
Global str報表對象(1) As String
'Connection宣告
Global cnnConnection As ADODB.Connection
'顯示錯誤訊息視窗
Public Sub ErrorLog()
frm990003.Show vbModal
End Sub
'取得來函記錄檔之系統類別
Public Function GetCkindSys(ByRef strDate As String, ByRef strSys() As String) As Boolean
Dim strSQL As String, i As Integer, rsRecordset As New ADODB.Recordset

On Error GoTo err
strSQL = "select distinct mr12 from mailrec where mr02=" + strDate
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSQL, cnnConnection
If rsRecordset.RecordCount > 0 Then
   Do While Not rsRecordset.EOF
         ReDim Preserve strSys(i) As String
         strSys(i) = rsRecordset.Fields(0)
         rsRecordset.MoveNext
         i = i + 1
   Loop
   GetCkindSys = True
Else
   MsgBox "Mailrec檔案無資料!!", vbCritical
End If
rsRecordset.Close
Exit Function
err:
MsgBox "讀取Mailrec檔案失敗!!", vbCritical
ErrorLog
End Function
'將From移至畫面之中心
Public Sub MoveFormToCenter(ByRef frmTemp As Form)
Dim intX  As Integer, intY As Integer

intX = (Screen.Width - frmTemp.Width) / 2
intY = (Screen.Height - frmTemp.Height) / 2
frmTemp.Move intX, intY
End Sub

'轉換字串以塞入SQL語法
Public Function CNULL(ByRef strNULL As String) As String
If strNULL = "" Then
   CNULL = "NULL"
Else
   CNULL = "'" + strNULL + "'"
End If
End Function
Public Function GetRdset(ByRef strSQL As String) As ADODB.Recordset
Dim rsRecordset As New ADODB.Recordset
On Error GoTo err
    rsRecordset.CursorLocation = adUseClient
    rsRecordset.Open strSQL, cnnConnection, adOpenDynamic
    Set GetRdset = rsRecordset

Exit Function
err:
ErrorLog
End Function

'讀取分案資料,intCaseKind系統分類
Public Function GetSQL(ByRef intCaseKind As Integer, ByRef intWhere As Integer, ByRef strReceiveCode As String) As String
 Dim strDateLine As String, bolChk As Boolean
   If intWhere <> 國外_CF Then
      bolChk = True
   Else
      bolChk = False
   End If
   If intWhere <> 國外_CF Then strDateLine = "-1911"
   Select Case intCaseKind
      Case 專利
         GetSQL = "select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他','') a01,st02 a02," & ChgPatent("", 1) & _
            " a03,pa05 a04,pa06 a05,pa07 a06,decode(pa09," + CNULL(大陸國家代號) + ",cpm04,cpm03) a07," & _
            SQLDate("CP05", bolChk) & " a08,cp16 a09,cp17 a10,cp18 a11," & SQLDate("CP06", bolChk) & " a12," & _
            SQLDate("CP07", bolChk) & " a13 from caseprogress,patent,casepropertymap,staff where cp09=" + CNULL(strReceiveCode) & _
            " and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp01=cpm01 and cp10=cpm02 and cp13=st01(+)"
      Case 商標
         GetSQL = "select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他','') a01,st02 a02," & _
            "tm01||'- '||decode(tm01," + CNULL(馬德里案) + ",substr(tm02,1,5),tm02)||decode(tm01," + CNULL(馬德里案) + ",decode(substr(tm02,6,1),'0','','- '||substr(tm02,6,1)),'')||decode(tm03,'0','','- '||tm03)||decode(tm04,'00','','- '||tm04) a03," & _
            "tm05 a04,tm06 a05,tm07 a06,decode(tm10," + CNULL(大陸國家代號) + ",cpm04,cpm03) a07," & _
            SQLDate("CP05", bolChk) & " a08,cp16 a09,cp17 a10,cp18 a11," & SQLDate("CP06", bolChk) & " a12," & SQLDate("CP07", bolChk) & _
            " a13 from caseprogress,trademark,casepropertymap,staff where cp09=" + CNULL(strReceiveCode) & " and cp01=tm01 and " & _
            "cp02=tm02 and cp03=tm03 and cp04=tm04 and cp01=cpm01 and cp10=cpm02 and cp13=st01(+)"
      Case Else
         GetSQL = "select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他','') a01,st02 a02," & ChgService("", 1) & _
            " a03,sp05 a04,sp06 a05,sp07 a06,decode(sp09," + CNULL(大陸國家代號) + ",cpm04,cpm03) a07," & _
            SQLDate("CP05", bolChk) & " a08,cp16 a09,cp17 a10,cp18 a11," & SQLDate("CP06", bolChk) & " a12," & _
            SQLDate("CP07", bolChk) & " a13 from caseprogress,servicepractice,casepropertymap,staff where cp09=" + CNULL(strReceiveCode) & _
            " and cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and cp01=cpm01 and cp10=cpm02 and cp13=st01(+)"
   End Select
End Function

