VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100106_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "一案二申請"
   ClientHeight    =   3930
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   7000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7000
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Height          =   315
      Left            =   5790
      TabIndex        =   1
      Top             =   30
      Width           =   1155
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3465
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   6121
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   8
   End
End
Attribute VB_Name = "frm100106_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/05/24 Form2.0已修改: grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/13 日期欄已修改
'add by nick 新增功能，秀一案二申請
Option Explicit
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件

Private Sub SetDataListWidth()
'Modified by Lydia 2019/11/01
Dim intField As Integer
'grdDataList.Cols = 4
intField = 9
grdDataList.Cols = intField
'end 2019/11/01

grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "已發文案號"
grdDataList.ColWidth(0) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "發文日"
grdDataList.ColWidth(1) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "未發文案號"
grdDataList.ColWidth(2) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "承辦人"
grdDataList.ColWidth(3) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
'Added by Lydia 2019/11/01 隱藏欄位：申請人1~5, FC代理人
For intI = 4 To intField - 1
     grdDataList.col = intI
     grdDataList.ColWidth(intI) = 0
Next intI
'end 2019/11/01
End Sub

Private Sub cmdok_Click()
Me.Hide
End Sub

Private Sub Form_Load()
bolToEndByNick = False
MoveFormToCenter Me
SetDataListWidth
End Sub

Sub Process()
Dim strSql As String
Dim strSQL1 As String
Dim strSQL2 As String
Dim StrSQL3 As String 'Add by Morgan 2010/3/11
Dim dblRow As Double 'Add By Sindy 2025/9/3

strSql = ""
'edit by nick 2004/10/05
'strSQL1 = "c1.cp27 is null and c2.cp01 in ('P','CFP') "
'strSQL2 = "c2.cp27 is null and c1.cp01 in ('P','CFP') "

'Remove by Morgan 2010/3/11 條件有問題,先取消(另有改語法故都不要)
'strSQL1 = "c1.cp27 is null and c2.cp01 in ('P','CFP') and c2.cp57 is not null "
'strSQL2 = "c2.cp27 is null and c1.cp01 in ('P','CFP') and c1.cp57 is not null "
'end 2010/3/11

'Added by Lydia 2019/11/01
m_AllSys = IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(, frm100106_1.txt5(0).Text))
intCufaCnt = 0
'end 2019/11/01

ClearQueryLog ("frm100106_1") 'Add By Sindy 2010/11/3 清除查詢印表記錄檔欄位
If frm100106_1.opt1(0).Value = True Then '本所期限
    If Len(Trim(frm100106_1.txt1(0).Text)) <> 0 Then
        strSQL1 = strSQL1 & " and c2.cp27 >=" & Val(ChangeTStringToWString(frm100106_1.txt1(0))) & " "
        strSQL2 = strSQL2 & " and c1.cp27 >=" & Val(ChangeTStringToWString(frm100106_1.txt1(0))) & " "
        StrSQL3 = StrSQL3 & " and cp27 >=" & Val(ChangeTStringToWString(frm100106_1.txt1(0))) & " "
    End If
    If Len(Trim(frm100106_1.txt1(1).Text)) <> 0 Then
        strSQL1 = strSQL1 & " and c2.cp27 <=" & Val(ChangeTStringToWString(frm100106_1.txt1(1))) & " "
        strSQL2 = strSQL2 & " and c1.cp27 <=" & Val(ChangeTStringToWString(frm100106_1.txt1(1))) & " "
        StrSQL3 = StrSQL3 & " and cp27 <=" & Val(ChangeTStringToWString(frm100106_1.txt1(1))) & " "
    End If
    If Len(Trim(frm100106_1.txt1(0).Text)) <> 0 Or Len(Trim(frm100106_1.txt1(1).Text)) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(0).Caption & frm100106_1.txt1(0) & "-" & frm100106_1.txt1(1) 'Add By Sindy 2010/11/3
    End If
Else
      If frm100106_1.opt1(1).Value = True Then '法定期限
            If Len(Trim(frm100106_1.txt2(0).Text)) <> 0 Then
                strSQL1 = strSQL1 & " and c2.cp27 >=" & Val(ChangeTStringToWString(frm100106_1.txt2(0))) & " "
                strSQL2 = strSQL2 & " and c1.cp27 >=" & Val(ChangeTStringToWString(frm100106_1.txt2(0))) & " "
                StrSQL3 = StrSQL3 & " and cp27 >=" & Val(ChangeTStringToWString(frm100106_1.txt2(0))) & " "
            End If
            If Len(Trim(frm100106_1.txt2(1).Text)) <> 0 Then
                strSQL1 = strSQL1 & " and c2.cp27 <=" & Val(ChangeTStringToWString(frm100106_1.txt2(1))) & " "
                strSQL2 = strSQL2 & " and c1.cp27 <=" & Val(ChangeTStringToWString(frm100106_1.txt2(1))) & " "
                StrSQL3 = StrSQL3 & " and cp27 <=" & Val(ChangeTStringToWString(frm100106_1.txt2(1))) & " "
            End If
            If Len(Trim(frm100106_1.txt2(0).Text)) <> 0 Or Len(Trim(frm100106_1.txt2(1).Text)) <> 0 Then
               pub_QL05 = pub_QL05 & ";" & frm100106_1.opt1(1).Caption & frm100106_1.txt2(0) & "-" & frm100106_1.txt2(1) 'Add By Sindy 2010/11/3
            End If
      End If
End If
'edit by nick 2004/07/14 不管無論如何，都是抓申請案，也就是，收文日最小，及收文號最小
'strSQL = "select distinct decode(c1.cp27,null,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04)," & _
               " decode(c1.cp27,null," & SQLDate("c2.cp27") & "," & SQLDate("c1.cp27") & ")," & _
               " decode(c1.cp27,null,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04)," & _
               " decode(c1.cp27,null,s2.st02,s1.st02) " & _
               " from caseprogress C1,casemap,caseprogress C2,staff s1,staff s2 " & _
               " Where cm01 = C1.CP01 And cm02 = C1.cp02 And cm03 = C1.cp03 And cm04 = C1.cp04 " & _
               " and cm05=C2.cp01 and cm06=C2.cp02 and cm07=C2.cp03 and cm08=C2.cp04 and cm10='3' and c1.cp14=s1.st01(+) and c2.cp14=s2.st01(+) " & _
               " and ((" & strSQL1 & ") or (" & strSQL2 & " ) ) "
'Modify by Morgan 2010/3/11 案件性質只要抓 101,102 就好(原語法太慢)
'strSql = "select distinct decode(c1.cp27,null,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04)," & _
               " decode(c1.cp27,null," & SQLDate("c2.cp27") & "," & SQLDate("c1.cp27") & ")," & _
               " decode(c1.cp27,null,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04)," & _
               " decode(c1.cp27,null,s1.st02,s2.st02) " & _
               " from caseprogress C1,caseprogress C2,staff s1,staff s2,(" & _
               " select cm01,cm02,cm03,cm04,min(Lcp05) lcp05,min(Lcp09) lcp09,cm05,cm06,cm07,cm08,min(rcp05) rcp05,min(rcp09) rcp09 from (" & _
               " select cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08,c3.cp05 lcp05,c3.cp09 lcp09, " & _
               " c4.cp05 rcp05,c4.cp09 rcp09 " & _
               " from casemap,caseprogress c3,caseprogress c4 where cm10='3' " & _
               " and c3.cp01=cm01 and c3.cp02=cm02 and c3.cp03=cm03 and c3.cp04=cm04 " & _
               " and c4.cp01=cm05 and c4.cp02=cm06 and c4.cp03=cm07 and c4.cp04=cm08 " & _
               " ) group by cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08 " & _
               ") newCaseMap " & _
               " Where newCaseMap.lcp09=c1.cp09  " & _
               " and newCaseMap.rcp09=c2.cp09 and c1.cp14=s1.st01(+) and c2.cp14=s2.st01(+) " & _
               " and ((" & strSQL1 & ") or (" & strSQL2 & " ) ) "
               
'Modified by Morgan 2012/6/21 原語法少抓新型未發文的情形
'strSql = "select distinct decode(c1.cp27,null,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04)," & _
               " decode(c1.cp27,null," & SQLDate("c2.cp27") & "," & SQLDate("c1.cp27") & ")," & _
               " decode(c1.cp27,null,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04)," & _
               " decode(c1.cp27,null,s1.st02,s2.st02) " & _
               " from (select cp09 X1,cm05,cm06,cm07,cm08 from caseprogress,casemap" & _
               " Where cp01||'' in ('P','CFP') and cp10||'' in ('101','102') " & StrSQL3 & _
               " and cm01(+)=cp01 and cm02(+)=cp02 and cm03(+)=cp03 and cm04(+)=cp04 and cm10='3'" & _
               " union select cp09 X1,cm01,cm02,cm03,cm04 from caseprogress c1,casemap m1" & _
               " Where cp01||'' in ('P','CFP') and cp10||'' in ('101','102') " & StrSQL3 & _
               " and cm05(+)=cp01 and cm06(+)=cp02 and cm07(+)=cp03 and cm08(+)=cp04 and cm10='3'" & _
               " ) X,caseprogress c1,caseprogress C2,staff s1,staff s2" & _
               " where c1.cp09(+)=X1 and c2.cp01(+)=cm05 and c2.cp02(+)=cm06 and c2.cp03(+)=cm07 and c2.cp04(+)=cm08" & _
               " and c2.cp10 in ('101','102') and c2.cp27 is null and c2.cp57 is null" & _
               " and s1.st01(+)=c1.cp14 and s2.st01(+)=c2.cp14"
'end 2010/3/11
'Modified by Lydia 2019/11/01 利益衝突案件：於c004後面，增加申請人1~5,FC代理人
'strSql = "select distinct decode(c1.cp27,null,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04)," & _
               " decode(c1.cp27,null," & SQLDate("c2.cp27") & "," & SQLDate("c1.cp27") & ")," & _
               " decode(c1.cp27,null,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04)," & _
               " decode(c1.cp27,null,s1.st02,s2.st02) " & _
               " from (select c1.cp09 X1,c2.cp09 X2 from caseprogress c1,casemap,caseprogress c2" & _
               " Where c1.cp01||'' in ('P','CFP') and c1.cp10||'' in ('101','102') " & strSQL2 & _
               " and cm01(+)=c1.cp01 and cm02(+)=c1.cp02 and cm03(+)=c1.cp03 and cm04(+)=c1.cp04 and cm10='3'" & _
               " and c2.cp01(+)=cm05 and c2.cp02(+)=cm06 and c2.cp03(+)=cm07 and c2.cp04(+)=cm08" & _
               " and c2.cp10 in ('101','102') and c2.cp27 is null and c2.cp57 is null" & _
               " union select c1.cp09 X1,c2.cp09 X2 from caseprogress c1,casemap,caseprogress c2" & _
               " Where c1.cp01||'' in ('P','CFP') and c1.cp10||'' in ('101','102') " & strSQL2 & _
               " and cm05(+)=c1.cp01 and cm06(+)=c1.cp02 and cm07(+)=c1.cp03 and cm08(+)=c1.cp04 and cm10='3'" & _
               " and c2.cp01(+)=cm01 and c2.cp02(+)=cm02 and c2.cp03(+)=cm03 and c2.cp04(+)=cm04" & _
               " and c2.cp10 in ('101','102') and c2.cp27 is null and c2.cp57 is null" & _
               " ) X,caseprogress c1,caseprogress C2,staff s1,staff s2" & _
               " where c1.cp09(+)=X1 and c2.cp09(+)=X2" & _
               " and s1.st01(+)=c1.cp14 and s2.st01(+)=c2.cp14"
'end 2012/6/21
strSql = "select distinct decode(c1.cp27,null,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04) as c001," & _
               " decode(c1.cp27,null," & SQLDate("c2.cp27") & "," & SQLDate("c1.cp27") & ") as c002," & _
               " decode(c1.cp27,null,c1.cp01||'-'||c1.cp02||'-'||c1.cp03||'-'||c1.cp04,c2.cp01||'-'||c2.cp02||'-'||c2.cp03||'-'||c2.cp04) as c003," & _
               " decode(c1.cp27,null,s1.st02,s2.st02) as c004"
strSql = strSql & ", decode(c1.cp27,null,p2.pa26,p1.pa26) as c005"
strSql = strSql & ", decode(c1.cp27,null,p2.pa27,p1.pa27) as c006"
strSql = strSql & ", decode(c1.cp27,null,p2.pa28,p1.pa28) as c007"
strSql = strSql & ", decode(c1.cp27,null,p2.pa29,p1.pa29) as c008"
strSql = strSql & ", decode(c1.cp27,null,p2.pa30,p1.pa30) as c009"
strSql = strSql & ", decode(c1.cp27,null,p2.pa75,p1.pa75) as c010"
strSql = strSql & " from (select c1.cp09 X1,c2.cp09 X2 from caseprogress c1,casemap,caseprogress c2" & _
               " Where c1.cp01||'' in ('P','CFP') and c1.cp10||'' in ('101','102') " & strSQL2 & _
               " and cm01(+)=c1.cp01 and cm02(+)=c1.cp02 and cm03(+)=c1.cp03 and cm04(+)=c1.cp04 and cm10='3'" & _
               " and c2.cp01(+)=cm05 and c2.cp02(+)=cm06 and c2.cp03(+)=cm07 and c2.cp04(+)=cm08" & _
               " and c2.cp10 in ('101','102') and c2.cp27 is null and c2.cp57 is null" & _
               " union select c1.cp09 X1,c2.cp09 X2 from caseprogress c1,casemap,caseprogress c2" & _
               " Where c1.cp01||'' in ('P','CFP') and c1.cp10||'' in ('101','102') " & strSQL2 & _
               " and cm05(+)=c1.cp01 and cm06(+)=c1.cp02 and cm07(+)=c1.cp03 and cm08(+)=c1.cp04 and cm10='3'" & _
               " and c2.cp01(+)=cm01 and c2.cp02(+)=cm02 and c2.cp03(+)=cm03 and c2.cp04(+)=cm04" & _
               " and c2.cp10 in ('101','102') and c2.cp27 is null and c2.cp57 is null" & _
               " ) X,caseprogress c1,caseprogress C2,staff s1,staff s2,patent p1, patent p2 "
strSql = strSql & " where c1.cp09(+)=X1 and c2.cp09(+)=X2" & _
               " and s1.st01(+)=c1.cp14 and s2.st01(+)=c2.cp14" & _
               " and c1.cp01=p1.pa01(+) and c1.cp02=p1.pa02(+) and c1.cp03=p1.pa03(+) and c1.cp04=p1.pa04(+)" & _
               " and c2.cp01=p2.pa01(+) and c2.cp02=p2.pa02(+) and c2.cp03=p2.pa03(+) and c2.cp04=p2.pa04(+)"
'end 2019/11/01

    CheckOC
    adoRecordset.CursorLocation = adUseClient
    'Modified by Lydia 2019/11/01 改變型態
    'adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    adoRecordset.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
    If adoRecordset.RecordCount = 0 Then
       InsertQueryLog (0) 'Add By Sindy 2010/11/3
JumpToNoData:   'Added by Lydia 2019/11/01
         Me.Hide
         Exit Sub
    Else
        dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/9/3
         
        'Added by Lydia 2019/11/01 逐案號判斷
        If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            adoRecordset.MoveFirst
            Do While adoRecordset.EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & adoRecordset.Fields(0), "" & adoRecordset.Fields("c005") & "," & adoRecordset.Fields("c006") & "," & adoRecordset.Fields("c007") & "," & adoRecordset.Fields("c008") & "," & adoRecordset.Fields("c009"), "" & adoRecordset.Fields("c010")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    adoRecordset.Delete
                End If
                adoRecordset.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow) 'Add By Sindy 2010/11/3
            If adoRecordset.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
        Else
            InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/11/3
        End If
       'end 2019/11/01
    End If
Set grdDataList.Recordset = adoRecordset
SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100106_6 = Nothing
End Sub

