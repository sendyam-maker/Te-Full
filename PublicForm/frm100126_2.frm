VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100126_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶重新委任案件查詢及列印"
   ClientHeight    =   6470
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6470
   ScaleWidth      =   8010
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   5955
      Left            =   60
      TabIndex        =   3
      Top             =   510
      Width           =   7875
      _ExtentX        =   13899
      _ExtentY        =   10513
      _Version        =   393216
      Rows            =   3
      Cols            =   1
      FixedRows       =   2
      FixedCols       =   0
      ScrollTrack     =   -1  'True
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
      _Band(0).Cols   =   1
   End
   Begin VB.ComboBox cbo1 
      Height          =   300
      Left            =   660
      Style           =   2  '單純下拉式
      TabIndex        =   0
      Top             =   600
      Width           =   1245
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5940
      TabIndex        =   1
      Top             =   30
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   7110
      TabIndex        =   2
      Top             =   30
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "發文日為11/11/11者為不必重新委任案件"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   165
      Left            =   1950
      TabIndex        =   5
      Top             =   660
      Width           =   5985
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶："
      Height          =   180
      Left            =   90
      TabIndex        =   4
      Top             =   660
      Width           =   540
   End
End
Attribute VB_Name = "frm100126_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Public cmdState As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件
Dim SeColPA As String

Sub SetGrd()
GRD1.Visible = False
   Dim arrGridHeadText, arrGridHeadWidth, arrGridHeadText1
   Dim iRow As Integer
   If frm100126_1.txt1(6) = "1" Then
        'Modified by Lydia 2019/11/01 +申請人1~5(cust01~cust05),FC代理人;
        'arrGridHeadText = Array("本所案號", "申請案號", "案件名稱", "申請日", "重新委任發文日", "", "")
        'arrGridHeadWidth = Array(1300, 1000, 2800, 1000, 1300, 0, 0)
        arrGridHeadText = Array("本所案號", "申請案號", "案件名稱", "申請日", "重新委任發文日", "FSort", "CaseNo" _
                                            , "CUST01", "CUST02", "CUST03", "CUST04", "CUST05", "FCNO")
        arrGridHeadWidth = Array(1300, 1000, 2800, 1000, 1300, 0, 0 _
                                            , 0, 0, 0, 0, 0, 0)
        'end 2019/11/01
        For iRow = 0 To GRD1.Cols - 1
            GRD1.row = 0
            GRD1.col = iRow
            GRD1.Text = arrGridHeadText(iRow)
            GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
            GRD1.CellAlignment = flexAlignCenterCenter
        Next
        If GRD1.Rows > 2 Then
            For iRow = 2 To GRD1.Rows - 1
               GRD1.row = iRow
               GRD1.col = 0
               GRD1.CellAlignment = flexAlignLeftCenter
               GRD1.col = 1
               GRD1.CellAlignment = flexAlignLeftCenter
               GRD1.col = 2
               GRD1.CellAlignment = flexAlignLeftCenter
               GRD1.col = 3
               GRD1.CellAlignment = flexAlignLeftCenter
               GRD1.col = 4
               GRD1.CellAlignment = flexAlignLeftCenter
            Next
            If GRD1.TextMatrix(GRD1.Rows - 1, 5) = "29" Then GRD1.Rows = GRD1.Rows - 1
            If GRD1.TextMatrix(GRD1.Rows - 1, 5) = "19" Then GRD1.Rows = GRD1.Rows - 1
        End If
   Else
        'Modified by Lydia 2019/11/01 +申請人1~5(cust01~cust05),FC代理人;
        'arrGridHeadText = Array("業務區", "智權人員", "客戶編號", "客戶名稱", "已重新", "待處理", "不必重新", "")
        'arrGridHeadText1 = Array("業務區", "智權人員", "客戶編號", "客戶名稱", "委任件數", "件數", "委任件數", "")
        'arrGridHeadWidth = Array(750, 750, 900, 2000, 1000, 1000, 1000, 0)
        arrGridHeadText = Array("業務區", "智權人員", "客戶編號", "客戶名稱", "已重新", "待處理", "不必重新", "CU13" _
                                        , "CUST01", "CUST02", "CUST03", "CUST04", "CUST05", "FCNO")
        arrGridHeadText1 = Array("業務區", "智權人員", "客戶編號", "客戶名稱", "委任件數", "件數", "委任件數", "CU13" _
                                        , "CUST01", "CUST02", "CUST03", "CUST04", "CUST05", "FCNO")
        arrGridHeadWidth = Array(750, 750, 900, 2000, 1000, 1000, 1000, 0 _
                                        , 0, 0, 0, 0, 0, 0)
        'end 2019/11/01
        GRD1.Cols = UBound(arrGridHeadText) + 1
        For iRow = 0 To GRD1.Cols - 1
            GRD1.row = 0
            GRD1.col = iRow
            GRD1.Text = arrGridHeadText(iRow)
            GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
            GRD1.CellAlignment = flexAlignCenterCenter
            GRD1.row = 1
            GRD1.col = iRow
            GRD1.Text = arrGridHeadText1(iRow)
            GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
            GRD1.CellAlignment = flexAlignCenterCenter
        Next
        GRD1.MergeCells = flexMergeRestrictRows
        GRD1.MergeRow(0) = True
        GRD1.MergeRow(1) = True
        GRD1.MergeCol(0) = True
        GRD1.MergeCol(1) = True
        GRD1.MergeCol(2) = True
        GRD1.MergeCol(3) = True
        GRD1.MergeCol(4) = True
        GRD1.MergeCol(5) = True
        If GRD1.Rows > 2 Then
            For iRow = 2 To GRD1.Rows - 1
               GRD1.row = iRow
               GRD1.col = 0
               GRD1.CellAlignment = flexAlignLeftCenter
               GRD1.col = 1
               GRD1.CellAlignment = flexAlignLeftCenter
               GRD1.col = 2
               GRD1.CellAlignment = flexAlignLeftCenter
               GRD1.col = 3
               GRD1.CellAlignment = flexAlignLeftCenter
               GRD1.col = 4
               GRD1.CellAlignment = flexAlignRightCenter
               GRD1.col = 5
               GRD1.CellAlignment = flexAlignRightCenter
               GRD1.col = 6
               GRD1.CellAlignment = flexAlignRightCenter
            Next
        End If
    End If
GRD1.Visible = True
End Sub

Private Sub cbo1_Change()
Dim rsTmp As New ADODB.Recordset
Dim strSql  As String
If rsTmp.State = 1 Then rsTmp.Close
strSql = "select NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) from customer where cu01||cu02='" & Cbo1.Text & "' "
rsTmp.CursorLocation = adUseClient
rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If rsTmp.RecordCount > 0 Then
    Label2.Caption = CheckStr(rsTmp.Fields(0))
Else
    Label2.Caption = ""
End If
rsTmp.Close
End Sub

Private Sub cbo1_Click()
Screen.MousePointer = vbHourglass
GRD1.MousePointer = flexArrowHourGlass
doQuery
GRD1.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdok_Click(Index As Integer)
cmdState = Index
PubShowNextData
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim tmpArr As Variant
MoveFormToCenter Me
If frm100126_1.txt1(6) = "1" Then
    pub_QL05 = pub_QL05 & ";" & frm100126_1.Label5 & "1.明細" 'Add By Sindy 2010/11/16
    Cbo1.Clear
    tmpArr = Split(frm100126_1.StrAllCU, ",")
    For i = 0 To UBound(tmpArr)
        Cbo1.AddItem tmpArr(i), i
    Next i
    Cbo1.Text = Cbo1.List(0)
    GRD1.Top = 960
    GRD1.Height = 5505
    InsertQueryLog (UBound(tmpArr) + 1) 'Add By Sindy 2010/11/16
Else
    pub_QL05 = pub_QL05 & ";" & frm100126_1.Label5 & "2.統計" 'Add By Sindy 2010/11/16
    Cbo1.Visible = False
    Label2.Visible = False
    GRD1.Top = 510
    GRD1.Height = 5955
    doQuery
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm100126_2 = Nothing
End Sub

Sub doQuery()
Dim rsTmp As New ADODB.Recordset
Dim strSql  As String
Dim intJump As Integer 'Added by Lydia 2019/11/01

If rsTmp.State = 1 Then rsTmp.Close
strSql = "select NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) from customer where cu01||cu02='" & Cbo1.Text & "' "
rsTmp.CursorLocation = adUseClient
rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If rsTmp.RecordCount > 0 Then
    Label2.Caption = CheckStr(rsTmp.Fields(0))
Else
    Label2.Caption = ""
End If
rsTmp.Close

If rsTmp.State = 1 Then rsTmp.Close
strSql = ""

 'Added by Lydia 2019/11/01利益衝突案件：於後面增加欄位
 SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
 intCufaCnt = 0
 m_AllSys = "FCP,P"
 'end 2019/11/01
 
If frm100126_1.txt1(6) = "1" Then  '明細
    If frm100126_1.Str100126SQL1 <> "" Then
        strSql = strSql & frm100126_1.Str100126SQL1 & " and pa26='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL1 & " and pa27='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL1 & " and pa28='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL1 & " and pa29='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL1 & " and pa30='" & Cbo1.Text & "' "
    End If
    If frm100126_1.Str100126SQL2 <> "" Then
        If strSql <> "" Then
            'Modified by Lydia 2019/11/01 +欄位
            'strSql = strSql & " union select '','','','','',19,'Z' from dual "
            strSql = strSql & " union select '','','','','',19,'Z','' as cust01,'' as cust02,'' as cust03,'' as cust04,'' as cust05,'' as fcno from dual "
            intJump = intJump + 1
            'end 2019/11/01
            strSql = strSql & " union "
        End If
        strSql = strSql & frm100126_1.Str100126SQL2 & " and pa26='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL2 & " and pa27='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL2 & " and pa28='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL2 & " and pa29='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL2 & " and pa30='" & Cbo1.Text & "' "
    End If
    If frm100126_1.Str100126SQL3 <> "" Then
        If strSql <> "" Then
            'Modified by Lydia 2019/11/01 +欄位
            'strSql = strSql & " union select '','','','','',29,'Z' from dual "
            strSql = strSql & " union select '','','','','',29,'Z','' as cust01,'' as cust02,'' as cust03,'' as cust04,'' as cust05,'' as fcno from dual "
            intJump = intJump + 1
            'end 2019/11/01
            strSql = strSql & " union "
        End If
        strSql = strSql & frm100126_1.Str100126SQL3 & " and pa26='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL3 & " and pa27='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL3 & " and pa28='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL3 & " and pa29='" & Cbo1.Text & "' "
        strSql = strSql & " union " & frm100126_1.Str100126SQL3 & " and pa30='" & Cbo1.Text & "' "
    End If
    strSql = strSql & " order by 6,7 "
    GRD1.FixedRows = 1
    GRD1.Rows = 2
    Call ProcDataByCase("1", strSql, intJump) 'Added by Lydia 2019/11/01
Else     '統計
        SeColPA = ",pa01||'-'||pa02||'-'||pa03||'-'||pa04 as caseno" & SeColPA 'Added by Lydia 2019/11/01
        strSql = "select a0902 as oA1,st02 as oA2,cu01||cu02 as oA3,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as oA4 ,"
        strSql = strSql & " sum(decode(cp27,null,0,19221111,0,1)) as oA5,sum(decode(cp27,null,1,19221111,0,0)) as oA6,sum(decode(cp27,null,0,19221111,1,0)) as oA7,cu13 "
        strSql = strSql & ",caseno , cust01, cust02, cust03, cust04, cust05, fcno " 'Added by Lydia 2019/11/01 利益衝突案件：加欄位
        'Modified by Lydia 2019/11/01 利益衝突案件：加欄位SeColPA
        strSql = strSql & " from customer,staff,acc090,(select distinct pa26,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
        strSql = strSql & " union select pa27,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
        strSql = strSql & " union select pa28,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
        strSql = strSql & " union select pa29,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
        strSql = strSql & " union select pa30,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
        strSql = strSql & " union select pa26,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
        strSql = strSql & " union select pa27,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
        strSql = strSql & " union select pa28,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
        strSql = strSql & " union select pa29,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
        strSql = strSql & " union select pa30,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) )  TmpTB "
        'end 2019/11/01
        strSql = strSql & " where cu01||cu02=TmpTB.pa26 and cu13=st01(+) and cu12=a0901(+) " & frm100126_1.MyStrSQL
        'Modified by Lydia 2019/11/01 2019/11/01 利益衝突案件：加欄位
        'strSql = strSql & " group by a0902,st02,cu01||cu02,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),cu13 order by 8,1,2,3 "
        strSql = strSql & " group by a0902,st02,cu01||cu02,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),cu13,caseno, cust01, cust02, cust03, cust04, cust05, fcno "
        
        GRD1.FixedRows = 2
        GRD1.Rows = 3
        Call ProcDataByCase("2", strSql) 'Added by Lydia 2019/11/01
End If

'Remove by Lydia 2019/11/01 改成模組ProcDataByCase
'rsTmp.CursorLocation = adUseClient
'rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If rsTmp.RecordCount > 0 Then
'    InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/11/16
'    Set GRD1.Recordset = rsTmp
'Else
'    InsertQueryLog (0) 'Add By Sindy 2010/11/16
'End If
'end 2019/11/01
'rsTmp.Close
'SetGrd
'end 2019/11/01
End Sub

Sub PubShowNextData()
Select Case cmdState
Case 0
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 1
        fnCloseAllFrm100
Case Else
End Select
End Sub

'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
Private Sub ProcDataByCase(ByVal pType As String, ByRef pSQL As String, Optional ByVal pJump As Integer = 0)
Dim rsAD As New ADODB.Recordset
Dim strMid As String, strGrp As String
Dim strJumpList As String '已排除的本所案號
Dim intJump As Integer '空白列數
Dim mESeqNo As String '暫存TB編號
Dim strA1 As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

    rsAD.CursorLocation = adUseClient
    rsAD.Open pSQL, cnnConnection, adOpenDynamic, adLockBatchOptimistic
    If rsAD.RecordCount > 0 Then
        dblRow = rsAD.RecordCount 'Add By Sindy 2025/9/3
        
        If pType = "2" Then '統計=>將明細資料丟到暫存檔rdatafactory
            Set RsTemp = PUB_CreateRecordset(rsAD, , , , Me.Name, mESeqNo)
        End If
        If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            rsAD.MoveFirst
            Do While rsAD.EOF = False
                strMid = "" & rsAD.Fields("CaseNo")
                '利益衝突案件：逐案號判斷
                If Len(strMid) > 9 Then
                    If strJumpList <> "" And InStr(strJumpList, strMid) > 0 Then
                        '剔除重複的本所案號
                        rsAD.Delete
                    Else
                        If strGrp <> strMid Then
                            If PUB_ChkCufaByCase(Me.Name, m_AllSys, strMid, "" & rsAD.Fields("cust01") & "," & rsAD.Fields("cust02") & "," & rsAD.Fields("cust03") & "," & rsAD.Fields("cust04") & "," & rsAD.Fields("cust05"), "" & rsAD.Fields("fcno")) = False Then
                                strJumpList = strJumpList & strMid & ","
                                intCufaCnt = intCufaCnt + 1
                                rsAD.Delete
                                If pType = "2" Then '統計
                                   strA1 = "delete from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and r009='" & strMid & "' "
                                   cnnConnection.Execute strA1
                                End If
                            End If
                        End If
                    End If
                End If
                strGrp = strMid
                rsAD.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow)
            If rsAD.RecordCount <= pJump Then  '跳過空白列
                  GoTo JumpToNoData
            End If
        Else
            InsertQueryLog (rsAD.RecordCount)
        End If
        If pType = "2" Then '統計
            If rsAD.State <> adStateClosed Then rsAD.Close
            rsAD.CursorLocation = adUseClient
            strA1 = "select R001 as OA1,R002 as OA2,R003 as OA3,R004 as OA4,SUM(R005) as OA5,SUM(R006) as OA6,SUM(R007) as OA7,R008 as CU13 from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "'"
            strA1 = strA1 & " group by r001,r002,r003,r004,r008 order by 8,1,2,3 "
            rsAD.Open strA1, cnnConnection, adOpenStatic, adLockReadOnly
            If rsAD.RecordCount = 0 Then
                GoTo JumpToNoData
            End If
        End If
        
        Set GRD1.Recordset = rsAD
    Else
        InsertQueryLog (0)
JumpToNoData:
    End If
    Set rsAD = Nothing
    SetGrd
End Sub
