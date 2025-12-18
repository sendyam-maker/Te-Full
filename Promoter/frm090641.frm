VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090641 
   BorderStyle     =   1  '單線固定
   Caption         =   "待辦案件量統計查詢"
   ClientHeight    =   4224
   ClientLeft      =   3180
   ClientTop       =   2208
   ClientWidth     =   4572
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4224
   ScaleWidth      =   4572
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3605
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   6350
      _Version        =   393216
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      ScrollTrack     =   -1  'True
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
      _Band(0).Cols   =   4
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2820
      TabIndex        =   0
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      Height          =   400
      Left            =   3600
      TabIndex        =   1
      Top             =   30
      Width           =   900
   End
End
Attribute VB_Name = "frm090641"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/22 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Create by Amy 2017/12/18
Option Explicit
Dim RsQ As New ADODB.Recordset
Dim strQ As String
Dim i As Integer, intCounter As Integer, intField As Integer
Dim strF(), intW()  '欄位名稱/大小

Private Sub cmdExcel_Click()
    Dim xlsAgentPoint As New Excel.Application
    Dim wksrpt As New Worksheet
    Dim strFileName As String, intWkName As String
    Dim intXlsSheet As Integer, intTitle As String
    
On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    
    strFileName = strSrvDate(1) & "待辦案件量統計" & ServerTime & MsgText(43)
    If Dir(strExcelPath & strFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
             MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileName
    End If
    
    intXlsSheet = 1
    xlsAgentPoint.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
    xlsAgentPoint.Workbooks.add
    If intWkName = MsgText(601) Then intWkName = Left(xlsAgentPoint.Worksheets(1).Name, Len(xlsAgentPoint.Worksheets(1).Name) - 1)
    Set wksrpt = xlsAgentPoint.Worksheets(intWkName & intXlsSheet)
    wksrpt.Activate
    
    intField = 65: intCounter = 1
    '已齊備未會稿統計
    Call SetTitle(wksrpt, 1)
    RsQ.MoveFirst
    Do While RsQ.EOF = False
        For i = LBound(strF) To UBound(strF)
            wksrpt.Range(Chr(i + intField) & intCounter).Value = "" & RsQ.Fields(i)
            If "" & RsQ.Fields(0) = "系統類別" Then
                wksrpt.Range(Chr(i + intField) & intCounter).Font.Bold = True
                wksrpt.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
            End If
        Next i
        RsQ.MoveNext
        intCounter = intCounter + 1
    Loop
    wksrpt.Name = "已齊備未會稿統計"
    intXlsSheet = intXlsSheet + 1
    
    'P新申請案已齊備未會稿明細
    strQ = GetSql(2)
    If RsQ.State <> adStateClosed Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
    If RsQ.RecordCount > 0 Then
        Set wksrpt = xlsAgentPoint.Worksheets(intWkName & intXlsSheet)
        wksrpt.Activate
        intCounter = 1
        Call SetTitle(wksrpt, 2)
        
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            For i = LBound(strF) To UBound(strF)
                wksrpt.Range(Chr(i + intField) & intCounter).Value = "" & RsQ.Fields(i)
            Next i
            RsQ.MoveNext
            intCounter = intCounter + 1
        Loop
        wksrpt.Name = "P新申請案已齊備未會稿明細"
        intXlsSheet = intXlsSheet + 1
    End If
    RsQ.Close
     
     '程序或繪圖已齊備未會稿明細
    strQ = GetSql(3)
    If RsQ.State <> adStateClosed Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
    If RsQ.RecordCount > 0 Then
        Set wksrpt = xlsAgentPoint.Worksheets(intWkName & intXlsSheet)
        wksrpt.Activate
        intCounter = 1
        Call SetTitle(wksrpt, 3)
        
        RsQ.MoveFirst
        Do While RsQ.EOF = False
            For i = LBound(strF) To UBound(strF)
                wksrpt.Range(Chr(i + intField) & intCounter).Value = "" & RsQ.Fields(i)
            Next i
            RsQ.MoveNext
            intCounter = intCounter + 1
        Loop
        wksrpt.Name = "程序或繪圖已齊備未會稿明細"
    End If
    RsQ.Close
    
    If Val(xlsAgentPoint.Version) < 12 Then
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
       xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    StatusClear
    'Modify by Amy 2021/06/22 +strExcelPathN中文字顯示
    MsgBox "Excel 檔案已產生！" & vbCrLf & "檔案存於" & strExcelPathN
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
    Set wksrpt = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub SetTitle(ByRef wksrpt As Worksheet, ByVal intChoose As Integer)
    Dim stTitle As String
    
    Select Case intChoose
        Case 1 '已齊備未會稿統計
            ReDim strF(2)
            ReDim intW(2)
            strF = Array("系統類別", "類型", "案件數")
            intW = Array(10, 10, 10)
            stTitle = "已齊備未會稿統計"
        Case 2 '新申請案已齊備未會稿明細
            ReDim strF(5)
            ReDim intW(5)
            strF = Array("本所案號", "案件性質", "申請國家", "收文日", "齊備日", "承辦人")
            intW = Array(16, 13, 10, 10, 10, 10)
            stTitle = "P新申請案已齊備未會稿明細"
        Case 3 '程序或繪圖已齊備未會稿明細
            ReDim strF(6)
            ReDim intW(6)
            strF = Array("本所案號", "總收文號", "收文日", "案件性質", "承辦人", "齊備日", "會稿")
            intW = Array(16, 11, 10, 13, 10, 10, 10)
            stTitle = "程序或繪圖已齊備未會稿明細"
    End Select
    
    wksrpt.Range(Chr(intField) & intCounter).Value = stTitle
    wksrpt.Range(Chr(intField) & intCounter).Font.Size = 16
    wksrpt.Range(Chr(intField) & intCounter).Font.Bold = True
    intCounter = intCounter + 1
   
    For i = LBound(strF) To UBound(strF)
        If intChoose > 1 Then wksrpt.Range(Chr(i + intField) & intCounter).Value = strF(i)
        wksrpt.Columns(Chr(i + intField) & ":" & Chr(i + intField)).ColumnWidth = intW(i)
        wksrpt.Range(Chr(i + intField) & intCounter).Font.Bold = True
        wksrpt.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
    Next i
    If intChoose > 1 Then intCounter = intCounter + 1
End Sub

Private Sub cmdOK_Click()
    strQ = "Select * From (" & GetSql(1) & ") Order by s,cp01,type"
    If RsQ.State <> adStateClosed Then RsQ.Close
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
    Set grdDataList.Recordset = RsQ
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    grdDataList.Clear
    SetDataListWidth
End Sub

Private Sub SetDataListWidth()
    grdDataList.col = 0: grdDataList.Text = "僅工程師"
    grdDataList.ColWidth(0) = 1200
    grdDataList.CellAlignment = flexAlignCenterCenter
    grdDataList.col = 1: grdDataList.Text = ""
    grdDataList.ColWidth(1) = 1200
    grdDataList.CellAlignment = flexAlignCenterCenter
    grdDataList.col = 2: grdDataList.Text = "剔除不計件"
    grdDataList.ColWidth(2) = 1200
    grdDataList.CellAlignment = flexAlignCenterCenter
    grdDataList.col = 3: grdDataList.Text = "Sort"
    grdDataList.ColWidth(3) = 0
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(strF)
        If UCase(strF(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function

Private Function GetSql(ByVal intChoose As Integer) As String
    GetSql = ""
    Select Case intChoose
        Case 1 '已齊備未會稿統計
            '僅工程師
            'Mododified by Lydia 2021/11/15 新申請案性質改成常數和3開頭
            'Decode(instr('101,102,103,104,105,109,110,112,113,114,115,118,120,122,125,301,302,303,304,305,306,307,308,309',cp10) ,0,'非新申請案','新申請案') =>Decode(Substr(cp10,1,1),'3','新申請案',Decode(Instr('" & NewCasePtyList & "',cp10),0, '非新申請案','新申請案')) Type
            'Memo by Lydia 2024/12/02 若未會稿的條件有異動，請一併變更frm090642-未會稿
            GetSql = "Select '僅工程師' as cp01,'' as type,'剔除不計件',1 as s From Dual " & _
                       "Union Select '系統類別' as cp01,'類型' as type,'案件數',2 as s From Dual " & _
                       "Union Select cp01,type,''||count(*),3 as s From" & _
                        "(Select cp01,cp02,cp03,cp04,cp09,cp10,cp26,ep06 齊備日,ep34 會稿, " & _
                            "Decode(Substr(cp10,1,1),'3','新申請案',Decode(Instr('" & NewCasePtyList & "',cp10),0, '非新申請案','新申請案')) Type " & _
                        "From CaseProgress, EngineerProgress, Staff " & _
                        "Where cp158=0 And cp159=0 And cp14=st01(+) And st03>='P1' And st03<='P11' And cp09=ep02(+) And cp26 is null " & _
                            "And Nvl(ep06,0) >0 And Nvl(ep07,0) =0 " & _
                        ") Group by cp01,type "
            '含程序繪圖
            'Mododified by Lydia 2021/11/15 新申請案性質改成常數和3開頭
            'Decode(instr('101,102,103,104,105,109,110,112,113,114,115,118,120,122,125,301,302,303,304,305,306,307,308,309',cp10),0,'非新申請案','新申請案') =>Decode(Substr(cp10,1,1),'3','新申請案',Decode(Instr('" & NewCasePtyList & "',cp10),0, '非新申請案','新申請案')) Type
            GetSql = GetSql & "Union Select '' as cp01,'' as type,'',4 as s From Dual " & _
                        "Union Select '含程序繪圖' as cp01,'' as type,'',5 as s From Dual " & _
                        "Union Select '系統類別' as cp01,'類型' as type,'案件數',6 as s From Dual " & _
                        "Union Select cp01,type,''||count(*),7 as s From" & _
                        "(Select cp01,cp02,cp03,cp04,cp09,cp10,cp26,ep06 齊備日,ep34 會稿," & _
                            "Decode(Substr(cp10,1,1),'3','新申請案',Decode(Instr('" & NewCasePtyList & "',cp10),0, '非新申請案','新申請案')) Type " & _
                        "From CaseProgress, EngineerProgress, Staff " & _
                        "Where cp158=0 And cp159=0 And cp14=st01(+) And st03>='P1' And st03<='P19' And cp09=ep02(+) And cp26 is null " & _
                            "And nvl(ep06,0)>0 And nvl(ep07,0)=0 " & _
                        ") Group by cp01,type"
        Case 2 'P新申請案已齊備未會稿明細
            'Mododified by Lydia 2021/11/15 新申請案性質改成常數和3開頭
            'Decode(instr('101,102,103,104,105,109,110,112,113,114,115,118,120,122,125,301,302,303,304,305,306,307,308,309',cp10) ,0,'非新申請案','新申請案') => Decode(Substr(cp10,1,1),'3','新申請案',Decode(Instr('" & NewCasePtyList & "',cp10),0, '非新申請案','新申請案')) Type
            GetSql = "Select cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,SubStr(Decode(pa09,'000',cpm03,cpm04),1,8) 案件性質,SubStr(na03,1,4) 申請國家," & _
                            "SubStr(sqldatet(cp05),1,10) 收文日,SubStr(sqldatet(ep06),1,10) 齊備日,st02 承辦人 From patent,casepropertymap,nation,staff," & _
                            "(Select cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp14,cp26,ep06,ep34," & _
                                "Decode(Substr(cp10,1,1),'3','新申請案',Decode(Instr('" & NewCasePtyList & "',cp10),0, '非新申請案','新申請案')) Type " & _
                             "From caseprogress, engineerprogress, staff " & _
                             "Where cp158=0 And cp159=0 And cp14=st01(+) And st03>='P1' And st03<='P11' And cp09=ep02(+) And cp26 is null And nvl(ep06,0) >0 And nvl(ep07,0) =0 " & _
                             ") Where cp01=pa01(+) And cp02=pa02(+) And cp03=pa03(+) And cp04=pa04(+) And pa09=na01(+) And cp01=cpm01(+) And cp10=cpm02(+) And cp14=st01(+) " & _
                                "And type='新申請案' And cp01='P' Order by cp01,cp10,cp14,pa09,cp02,cp05"
 
        Case 3 '程序或繪圖已齊備未會稿明細
            GetSql = "Select cp01||'-'||cp02||'-'||cp03||'-'||cp04 本所案號,cp09 總收文號,sqldatet(cp05) 收文日,substr(decode(cpm03,'（無）',cpm04,cpm03),1,8) 案件性質,st02 承辦人,sqldatet(ep06) 齊備日,ep34 會稿 " & _
                          "From caseprogress, engineerprogress, staff, casepropertymap " & _
                          "Where cp158=0 And cp159=0 And cp14=st01(+) And st03>='P12' And st03<='P19' And cp09=ep02(+) And cp26 is null And cp01=cpm01(+) And cp10=cpm02(+) " & _
                          "And nvl(ep06,0)>0 And nvl(ep07,0)=0 Order by cp01,cp10,cp14"
 
    End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frm090641 = Nothing
End Sub
