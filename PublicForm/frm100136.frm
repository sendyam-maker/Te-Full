VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100136 
   BorderStyle     =   1  '單線固定
   Caption         =   "不得宣傳客戶名稱資料查詢"
   ClientHeight    =   5580
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   7930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7930
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel檔"
      Height          =   372
      Index           =   2
      Left            =   4176
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   1164
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      Height          =   372
      Index           =   1
      Left            =   6576
      TabIndex        =   2
      Top             =   120
      Width           =   1164
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "畫面更新(&E)"
      Height          =   372
      Index           =   0
      Left            =   5376
      TabIndex        =   1
      Top             =   120
      Width           =   1164
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
      Height          =   4860
      Left            =   48
      TabIndex        =   0
      Top             =   576
      Width           =   7788
      _ExtentX        =   13741
      _ExtentY        =   8573
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm100136"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2023/06/16
Option Explicit
Dim IntF As Integer, strFind As String
Dim rsFD As New ADODB.Recordset
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序

Private Sub cmdok_Click(Index As Integer)

   Select Case Index
      Case 0 '畫面更新

         Call doQuery
      Case 1 '結束
         Unload Me
      Case 2 '產生Excel檔

         Call doQuery(True)
   End Select
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
       
   If InStr("F41,M51", Pub_StrUserSt03) > 0 Then
      cmdOK(2).Visible = True
   Else
      cmdOK(2).Visible = False
   End If
   
   Call doQuery
      
End Sub

Public Function ChkUseRight() As Boolean
   
   ChkUseRight = False
   
   '開放權限:各部門副理級以上之主管+業拓F41
   If Pub_StrUserSt03 <> "M51" Then
      'modify by sonia 2025/4/15 +B4013高于婷(管理部 秘書)-Elvan
      strSql = "select st01 from staff where (nvl(st20,'99')<='44' or st03='F41' or st01='B4013') and st01='" & strUserNum & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 0 Then
         MsgBox "無此使用權限...", , "警告!!"
         Exit Function
      End If
   End If
   
   ChkUseRight = True
   
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set rsFD = Nothing
   
   Set frm100136 = Nothing
End Sub

Private Sub doQuery(Optional ByVal bolExcel As Boolean = False)
   
   ClearQueryLog (Me.Name)

   Call SetGrd(True) '清空
   
   'Modifieid by Lydia 2025/04/07 若國籍為"013"或"020"則名稱抓中-->英-->日, 否則抓英-->中-->日
   '客戶
'   strFind = "select sqldatet(min(cr02)) as cdate,na03 as cnation,cu01||cu02 as appno,decode(cu05,null,nvl(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90) as appname,na01 " & _
'             "From contactrecord, customer, nation where instr(cr05,'A14')>0 and substr(cr03,1,8)=cu01(+) and substr(cr03,9,1)=cu02(+) and cu10=na01(+) and cu01 is not null " & _
'             "group by cu01||cu02, na03, decode(cu05,null,nvl(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90), na01 "
'   '代理人
'   strFind = strFind & " Union select sqldatet(min(cr02)) as cdate,na03 as cnation,fa01||fa02 as appno,decode(fa05,null,nvl(fa04,fa06),fa05||' '||fa63||' '||fa64||' '||fa65) as appname,na01 " & _
'             "From contactrecord, fagent, nation where instr(cr05,'A14')>0 and substr(cr03,1,8)=fa01(+) and substr(cr03,9,1)=fa02(+) and fa10=na01(+) and fa01 is not null " & _
'             "group by fa01||fa02, na03, decode(fa05,null,nvl(fa04,fa06),fa05||' '||fa63||' '||fa64||' '||fa65), na01 "
'   '潛在客戶
'   strFind = strFind & " Union select sqldatet(min(cr02)) as cdate,na03 as cnation,pcu01||pcu02 as appno,decode(pcu03,null,nvl(pcu08,pcu07),pcu03||' '||pcu04||' '||pcu05||' '||pcu06) as appname,na01 " & _
'            "From contactrecord, potcustomer, nation where instr(cr05,'A14')>0 and substr(cr03,1,8)=pcu01(+) and substr(cr03,9,1)=pcu02(+) and pcu09=na01(+) and pcu01 is not null " & _
'            "group by pcu01||pcu02, na03, decode(pcu03,null,nvl(pcu08,pcu07),pcu03||' '||pcu04||' '||pcu05||' '||pcu06), na01 "
'   '國內潛在客戶
'   strFind = strFind & " Union select sqldatet(min(cr02)) as cdate,na03 as cnation,poc01||poc02 as appno,decode(poc23,null,nvl(poc03,poc27),poc23||' '||poc24||' '||poc25||' '||poc26) as appname,na01 " & _
'            "From contactrecord, potcustomer1, nation where instr(cr05,'A14')>0 and substr(cr03,1,8)=poc01(+) and substr(cr03,9,1)=poc02(+) and poc04=na01(+) and poc01 is not null " & _
'            "group by poc01||poc02, na03, decode(poc23,null,nvl(poc03,poc27),poc23||' '||poc24||' '||poc25||' '||poc26), na01 "
   '客戶
   strFind = "select sqldatet(min(cr02)) as cdate,na03 as cnation,cu01||cu02 as appno,Decode(sign(instr('000,001,002,003,004,005,006,007,008,009,013,020',CU10)),1,NVL(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),Decode(CU05,NULL,NVL(CU04,CU06),CU05||' '||CU88||' '||CU89||' '||CU90 )) as appname,na01 " & _
             "From contactrecord, customer, nation where instr(cr05,'A14')>0 and substr(cr03,1,8)=cu01(+) and substr(cr03,9,1)=cu02(+) and cu10=na01(+) and cu01 is not null " & _
             "group by cu01||cu02, na03,Decode(sign(instr('000,001,002,003,004,005,006,007,008,009,013,020',CU10)),1,NVL(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),Decode(CU05,NULL,NVL(CU04,CU06),CU05||' '||CU88||' '||CU89||' '||CU90 )), na01 "
   '代理人
   strFind = strFind & " Union select sqldatet(min(cr02)) as cdate,na03 as cnation,fa01||fa02 as appno,Decode(sign(instr('000,001,002,003,004,005,006,007,008,009,013,020',FA10)),1,NVL(FA04,Decode(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),Decode(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as appname,na01 " & _
             "From contactrecord, fagent, nation where instr(cr05,'A14')>0 and substr(cr03,1,8)=fa01(+) and substr(cr03,9,1)=fa02(+) and fa10=na01(+) and fa01 is not null " & _
             "group by fa01||fa02, na03, Decode(sign(instr('000,001,002,003,004,005,006,007,008,009,013,020',FA10)),1,NVL(FA04,Decode(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),Decode(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)), na01 "
   '潛在客戶
   strFind = strFind & " Union select sqldatet(min(cr02)) as cdate,na03 as cnation,pcu01||pcu02 as appno,Decode(Sign(Instr('000,001,002,003,004,005,006,007,008,009,013,020',Pcu09)),1,Nvl(Pcu08,Decode(Pcu03,Null,Pcu07,Pcu03||' '||Pcu04||' '||Pcu05||' '||Pcu06)),Decode(Pcu03,Null,Nvl(Pcu07,Pcu08),Pcu03||' '||Pcu04||' '||Pcu05||' '||Pcu06)) as appname,na01 " & _
            "From contactrecord, potcustomer, nation where instr(cr05,'A14')>0 and substr(cr03,1,8)=pcu01(+) and substr(cr03,9,1)=pcu02(+) and pcu09=na01(+) and pcu01 is not null " & _
            "group by pcu01||pcu02, na03, Decode(Sign(Instr('000,001,002,003,004,005,006,007,008,009,013,020',Pcu09)),1,Nvl(Pcu08,Decode(Pcu03,Null,Pcu07,Pcu03||' '||Pcu04||' '||Pcu05||' '||Pcu06)),Decode(Pcu03,Null,Nvl(Pcu07,Pcu08),Pcu03||' '||Pcu04||' '||Pcu05||' '||Pcu06)), na01 "
   '國內潛在客戶
   strFind = strFind & " Union select sqldatet(min(cr02)) as cdate,na03 as cnation,poc01||poc02 as appno,Decode(Sign(Instr('000,001,002,003,004,005,006,007,008,009,013,020',Poc04)),1,Nvl(Poc03,Decode(Poc23,Null,Poc28,Poc23||' '||Poc24||' '||Poc25||' '||Poc26)),Decode(Poc23,Null,Nvl(Poc03,Poc28),Poc23||' '||Poc24||' '||Poc25||' '||Poc26)) as appname,na01 " & _
            "From contactrecord, potcustomer1, nation where instr(cr05,'A14')>0 and substr(cr03,1,8)=poc01(+) and substr(cr03,9,1)=poc02(+) and poc04=na01(+) and poc01 is not null " & _
            "group by poc01||poc02, na03, Decode(Sign(Instr('000,001,002,003,004,005,006,007,008,009,013,020',Poc04)),1,Nvl(Poc03,Decode(Poc23,Null,Poc28,Poc23||' '||Poc24||' '||Poc25||' '||Poc26)),Decode(Poc23,Null,Nvl(Poc03,Poc28),Poc23||' '||Poc24||' '||Poc25||' '||Poc26)), na01 "
            
   
   '排序:名稱,國家,日期；　名稱：英->中->日 'Memo by Lydia 2025/04/07 若國籍為"013"或"020"則名稱抓中-->英-->日, 否則抓英-->中-->日
   strFind = strFind & "order by appname,na01,cdate"
   IntF = 1
   Set rsFD = ClsLawReadRstMsg(IntF, strFind)
   If IntF = 1 Then
      pub_QL05 = pub_QL05 & ";查詢" & IIf(bolExcel = True, ";產生Excel", "")
      InsertQueryLog (rsFD.RecordCount)
      
      Set MGrid1.Recordset = rsFD
      Call SetGrd
       
      '產生Excel檔
      If bolExcel = True Then
         Call ProcExcelSave(rsFD)
      End If
   Else
      InsertQueryLog (0)
      MsgBox "查無資料!!"
   End If
   
End Sub

'產生Excel檔
Private Sub ProcExcelSave(ByRef pQuery As ADODB.Recordset)
Dim stXLSFileName As String, stXLSFullPath As String
Dim iRow As Integer, colMax As Integer, intQ As Integer
Dim xlsReport
Dim wksReport
Dim tmpArray As Variant

   stXLSFileName = strSrvDate(1) & "_不得宣傳客戶名稱清單" & MsgText(43)
   If Dir(strExcelPath & stXLSFileName) <> "" Then
       Kill strExcelPath & stXLSFileName
   End If
   stXLSFullPath = strExcelPath & stXLSFileName
   
   Set xlsReport = CreateObject("Excel.Application")
   xlsReport.Visible = False

   With pQuery
      .MoveFirst
      xlsReport.SheetsInNewWorkbook = 1
      xlsReport.Workbooks.add
      Set wksReport = xlsReport.Worksheets(1)
      wksReport.Cells.NumberFormatLocal = "@"
      wksReport.Cells.RowHeight = 18
      '設定列印邊界
      wksReport.PageSetup.PaperSize = 9 'A4
      wksReport.PageSetup.Orientation = 1 '直印
      wksReport.PageSetup.LeftMargin = xlsReport.CentimetersToPoints(1)
      wksReport.PageSetup.RightMargin = xlsReport.CentimetersToPoints(1)
      wksReport.PageSetup.TopMargin = xlsReport.CentimetersToPoints(1)
      wksReport.PageSetup.BottomMargin = xlsReport.CentimetersToPoints(1)
      wksReport.PageSetup.CenterHorizontally = True '列印頁面水平置中
      '表頭
      colMax = 4
      ReDim tmpArr(1 To colMax)
      wksReport.Range("A:A").ColumnWidth = 10
      wksReport.Range("A:A").HorizontalAlignment = xlLeft
      wksReport.Range("B:B").ColumnWidth = 12
      wksReport.Range("B:B").HorizontalAlignment = xlLeft
      wksReport.Range("C:C").ColumnWidth = 12
      wksReport.Range("C:C").HorizontalAlignment = xlLeft
      wksReport.Range("D:D").ColumnWidth = 58
      wksReport.Range("D:D").HorizontalAlignment = xlLeft
      wksReport.Range("1:1").RowHeight = 30
      wksReport.Range(Pub_NumberToSystem26(1) & "1:" & Pub_NumberToSystem26(colMax) & "1").Merge   'A1:D1
      wksReport.Range("A1").Value = "不得宣傳客戶名稱清單"
      wksReport.Range("A1").Font.Name = "標楷體"
      wksReport.Range("A1").Font.Size = 16
      wksReport.Range("A1").Font.Bold = True
      wksReport.Range("A1").HorizontalAlignment = xlCenter
      wksReport.Range("A1").VerticalAlignment = xlCenter
      wksReport.Range("D2").Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
      wksReport.Range("D2").HorizontalAlignment = xlRight
      tmpArr(1) = "建檔日期":   tmpArr(2) = "國　　籍":   tmpArr(3) = "編　　號":   tmpArr(4) = "名　　　　稱"
      wksReport.Range(Pub_NumberToSystem26(1) & "3:" & Pub_NumberToSystem26(colMax) & "3").Value = tmpArr
      wksReport.Range(Pub_NumberToSystem26(1) & "3:" & Pub_NumberToSystem26(colMax) & "3").Font.Bold = True
      wksReport.Range(Pub_NumberToSystem26(1) & "3:" & Pub_NumberToSystem26(colMax) & "3").Select
      xlsReport.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      iRow = 3
      Do While Not .EOF
         iRow = iRow + 1
         For intQ = 1 To colMax
            tmpArr(intQ) = "" & pQuery.Fields(intQ - 1)
         Next intQ
         wksReport.Range(Pub_NumberToSystem26(1) & iRow & ":" & Pub_NumberToSystem26(colMax) & iRow).Value = tmpArr
         .MoveNext
      Loop
      wksReport.Range(Pub_NumberToSystem26(1) & iRow & ":" & Pub_NumberToSystem26(colMax) & iRow).Select
      xlsReport.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
      wksReport.Range("A1").Select
      
      '判斷版本
      If Val(xlsReport.Version) < 12 Then
           xlsReport.Workbooks(1).SaveAs FileName:=stXLSFullPath, FileFormat:=-4143
      Else
           xlsReport.Workbooks(1).SaveAs FileName:=stXLSFullPath, FileFormat:=56
      End If
      
      xlsReport.Workbooks.Close
      xlsReport.Quit
   End With
   
   Set xlsReport = Nothing
   Set wksReport = Nothing
   
   MsgBox "Excel檔案產生完成！（檔案位置：" & stXLSFullPath & "）"
   
   Exit Sub
   
ErrHnd:

   MsgBox Err.Description
End Sub
Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   arrGridHeadText = Array("建檔日期", "國　　籍", "編　　號", "名　　　稱", "NA01")
   arrGridHeadWidth = Array(1100, 1200, 1100, 3500, 0)
   
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

   MGrid1.Visible = True
   
End Sub

Private Sub MGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow MGrid1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   MGrid1.col = nCol
   MGrid1.row = nRow
   If Me.MGrid1.row < 1 And Me.MGrid1.Text <> "V" Then
      '保留
      'If InStr("相似度", Me.MGrid1.Text) > 0 Then
      '   If m_blnColOrderAsc = True Then
      '      Me.MGrid1.Sort = 3  '數值昇冪
      '      m_blnColOrderAsc = False
      '   Else
      '      Me.MGrid1.Sort = 4 '數值降冪
      '      m_blnColOrderAsc = True
      '   End If
      'Else
         If m_blnColOrderAsc = True Then
            Me.MGrid1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.MGrid1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      'End If
   End If
End Sub
