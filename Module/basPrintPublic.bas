Attribute VB_Name = "basPrintPublic"
Option Explicit
'Connection宣告
Global cnnConnection As adodb.Connection
'顯示錯誤訊息視窗
Public Sub ErrorLog()
frm990003.Show vbModal
End Sub
'讀取分案資料,intCaseKind系統分類
Public Function GetSQL(ByRef intCaseKind As Integer, ByRef intWhere As Integer, ByRef strReceiveCode As String) As String
Dim strDateLine As String

If intWhere <> 國外_CF Then strDateLine = "-1911"
Select Case intCaseKind
             Case 專利
                        GetSQL = "select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他','') a01,st02 a02,pa01||'- '||pa02||decode(pa03,'0','','- '||pa03)||decode(pa04,'00','','- '||pa04) a03,pa05 a04,pa06 a05,pa07 a06,decode(pa09," + CNULL(大陸國家代號) + ",cpm04,cpm03) a07,substr(cp05,1,4)" + strDateLine + "||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) a08,cp16 a09,cp17 a10,cp18 a11,decode(cp06,null,'',substr(cp06,1,4)" + strDateLine + "||'/'||substr(cp06,5,2)||'/'||substr(cp06,7,2)) a12,decode(cp07,null,'',substr(cp07,1,4)" + strDateLine + "||'/'||substr(cp07,5,2)||'/'||substr(cp07,7,2)) a13  from caseprogress,patent,casepropertymap,staff where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp01=cpm01 and cp10=cpm02 and cp13=st01(+) and cp09=" + CNULL(strReceiveCode)
             Case 商標
                        GetSQL = "select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他','') a01,st02 a02,tm01||'- '||decode(tm01," + CNULL(馬德里案) + ",substr(tm02,1,5),tm02)||decode(tm01," + CNULL(馬德里案) + ",decode(substr(tm02,6,1),'0','','- '||substr(tm02,6,1)),'')||decode(tm03,'0','','- '||tm03)||decode(tm04,'00','','- '||tm04) a03,tm05 a04,tm06 a05,tm07 a06,decode(tm10," + CNULL(大陸國家代號) + ",cpm04,cpm03) a07,substr(cp05,1,4)" + strDateLine + "||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) a08,cp16 a09,cp17 a10,cp18 a11,decode(cp06,null,'',substr(cp06,1,4)" + strDateLine + "||'/'||substr(cp06,5,2)||'/'||substr(cp06,7,2)) a12,decode(cp07,null,'',substr(cp07,1,4)" + strDateLine + "||'/'||substr(cp07,5,2)||'/'||substr(cp07,7,2)) a13 from caseprogress,trademark,casepropertymap,staff where cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and cp01=cpm01 and cp10=cpm02 and cp13=st01(+) and cp09=" + CNULL(strReceiveCode)
             Case Else
                        GetSQL = "select decode(st06,'1','北所','2','中所','3','南所','4','高所','5','其他','') a01,st02 a02,sp01||'- '||sp02||decode(sp03,'0','','- '||sp03)||decode(sp04,'00','','- '||sp04) a03,sp05 a04,sp06 a05,sp07 a06,decode(sp09," + CNULL(大陸國家代號) + ",cpm04,cpm03) a07,substr(cp05,1,4)" + strDateLine + "||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) a08,cp16 a09,cp17 a10,cp18 a11,decode(cp06,null,'',substr(cp06,1,4)" + strDateLine + "||'/'||substr(cp06,5,2)||'/'||substr(cp06,7,2)) a12,decode(cp07,null,'',substr(cp07,1,4)" + strDateLine + "||'/'||substr(cp07,5,2)||'/'||substr(cp07,7,2)) a13 from caseprogress,servicepractice,casepropertymap,staff where cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and cp01=cpm01 and cp10=cpm02 and cp13=st01(+) and cp09=" + CNULL(strReceiveCode)
End Select
End Function
'轉換字串以塞入SQL語法
Public Function CNULL(ByRef strNULL As String) As String
If strNULL = "" Then
   CNULL = "NULL"
Else
   CNULL = "'" + strNULL + "'"
End If
End Function
