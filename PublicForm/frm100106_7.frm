VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100106_7 
   BorderStyle     =   1  '單線固定
   Caption         =   "已發文未收達"
   ClientHeight    =   5720
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   9310
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   315
      Left            =   6750
      TabIndex        =   2
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "繼續(&C)"
      Height          =   315
      Left            =   8010
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5205
      Left            =   90
      TabIndex        =   0
      Top             =   420
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   9172
      _Version        =   393216
      Cols            =   4
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
      _Band(0).Cols   =   4
   End
End
Attribute VB_Name = "frm100106_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/05/24 Form2.0已修改: grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
'Add by Morgan 2005/4/12 已發文未收達,系統日-3天>=本所期限>=系統日-3天-1月
Option Explicit
Dim arrGridHeadText, arrGridHeadWidth
Dim PLeft(0 To 12) As Integer, iPrint As Integer, iPage As Integer, strTemp(1 To 8) As String
Dim m_iTitleFontSize As Single, m_iFontSize As Single
Dim m_iStartX As Integer, m_iStartY As Integer
Dim m_iPageHeight As Integer, m_iLineHeight As Integer
Dim m_iNextStep As Integer 'Added by Mogan 2018/9/26
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件

Private Sub SetDataListWidth(Optional ByVal p_bol1st As Boolean = False)

   Dim iCol As Integer
   
   If p_bol1st Then SetGridHead

   With grdDataList
      .Visible = False
      .row = 0
      For iCol = 0 To .Cols - 1
         .col = iCol
         .Text = arrGridHeadText(iCol)
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .CellAlignment = flexAlignCenterCenter
      Next
      .Refresh
      .Visible = True
   End With
End Sub

Private Sub cmdok_Click()
   'Modified by Morgan 2018/9/26 增加客戶函已齊備未判發清單
   'Me.Hide
   GoNext
   'end 2018/9/26
End Sub

Private Sub cmdPrint_Click()

   Dim iRow As Integer
   GetPleft
   With grdDataList
      If .TextMatrix(1, 1) <> "" Then
         iPage = 1
         PrintPageHeader
         PrintPageHeader1
         For iRow = 1 To .Rows - 1
            strTemp(1) = .TextMatrix(iRow, 1)
            strTemp(2) = Left(.TextMatrix(iRow, 2), 10)
            strTemp(3) = Left(.TextMatrix(iRow, 3), 7)
            strTemp(4) = .TextMatrix(iRow, 4)
            strTemp(5) = .TextMatrix(iRow, 7)
            strTemp(6) = Left(.TextMatrix(iRow, 11), 5)
            strTemp(7) = .TextMatrix(iRow, 13)
            If m_iNextStep = 1 Then
               strTemp(8) = .TextMatrix(iRow, 0)
            Else
               strTemp(8) = Left(.TextMatrix(iRow, 14), 30)
            End If
            PrintDetail
         Next
         Call PrintReportFooter(iRow - 1)
      End If
   End With
End Sub

Sub GetPleft()

   m_iTitleFontSize = 22
   m_iFontSize = 12
   m_iStartX = 500
   m_iStartY = 500
   m_iPageHeight = 10300
   m_iLineHeight = 300

   Erase PLeft
   PLeft(0) = 500
   '本所案號(2000)
   PLeft(1) = 500
   '案件名稱(2700)
   PLeft(2) = PLeft(1) + 2000
   '案件性質(1950)
   PLeft(3) = PLeft(2) + 2700
   '承辦人(1050)
   PLeft(4) = PLeft(3) + 1950
   '法定期限(1200)
   PLeft(5) = PLeft(4) + 1050
   '申請國家(1400)
   PLeft(6) = PLeft(5) + 1200
   '發文日(1200)
   PLeft(7) = PLeft(6) + 1400
   '代理人
   PLeft(8) = PLeft(7) + 1200
    
End Sub

Private Sub PrintNewLine(Optional ByVal p_bolHeader1 As Boolean = True, Optional ByVal p_iExtraLines As Integer = 1)

   iPrint = iPrint + m_iLineHeight
   If iPrint >= (m_iPageHeight - m_iLineHeight - p_iExtraLines * m_iLineHeight) Then
      Printer.CurrentX = m_iStartX
      Printer.CurrentY = iPrint
      Printer.Print String(200, "-")
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If p_bolHeader1 Then
         PrintPageHeader1
      End If
      iPrint = iPrint + m_iLineHeight
    End If
    
End Sub

Sub PrintDetail()

    Dim iCol As Integer

    PrintNewLine
    For iCol = LBound(strTemp) To UBound(strTemp)
        Printer.CurrentX = PLeft(iCol)
        Printer.CurrentY = iPrint
        Printer.Print strTemp(iCol)
    Next
    
End Sub

Sub PrintPageHeader()
    iPrint = m_iStartY
    Printer.Orientation = 2
    Printer.FontName = "細明體"
    Printer.Font.Size = m_iTitleFontSize
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    strExc(1) = Me.Caption & "清單"
    'Printer.CurrentX = 5800
    Printer.CurrentX = (Printer.Width - Printer.TextWidth(strExc(1))) / 2
    Printer.CurrentY = iPrint
    Printer.Print strExc(1)
    iPrint = iPrint + 500
    Printer.Font.Size = m_iFontSize
    Printer.Font.Bold = False
    Printer.Font.Underline = False
    PrintNewLine
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print "列印人：" & strUserName
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
    PrintNewLine
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print "系統類別：" & frm100106_1.txt5(0).Text
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "頁    次：" & str(iPage)
    PrintNewLine
End Sub

Sub PrintPageHeader1()

    Call PrintNewLine(False, 1)
    
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    Printer.Print "本所案號"
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = iPrint
    Printer.Print "案件名稱"
    Printer.CurrentX = PLeft(3)
    Printer.CurrentY = iPrint
    Printer.Print "案件性質"
    Printer.CurrentX = PLeft(4)
    Printer.CurrentY = iPrint
    Printer.Print "承辦人"
    Printer.CurrentX = PLeft(5)
    Printer.CurrentY = iPrint
    Printer.Print "法定期限"
    Printer.CurrentX = PLeft(6)
    Printer.CurrentY = iPrint
    Printer.Print "申請國家"
    Printer.CurrentX = PLeft(7)
    Printer.CurrentY = iPrint
    Printer.Print "發文日"
    Printer.CurrentX = PLeft(8)
    Printer.CurrentY = iPrint
    If m_iNextStep = 1 Then
        Printer.Print "提申期限"
    Else
      Printer.Print "代理人"
    End If
    PrintNewLine
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 2)
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    PrintNewLine
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   SetDataListWidth True
End Sub

Private Sub SetGridHead()
   'Added by Morgan 2018/9/26
   If m_iNextStep = 2 Then
      'Modified by Lydia 2019/11/01 +PA26~PA30, PA75
      arrGridHeadText = Array("齊備日", "本所案號", "案件名稱", "案件性質", "判發人", "PA26", "PA27", "PA28", "PA29", "PA30", "PA75")
      arrGridHeadWidth = Array(810, 1300, 2000, 800, 850, 0, 0, 0, 0, 0, 0)
      
   'Added by Morgan 2020/5/13
   ElseIf m_iNextStep = 1 Then
      arrGridHeadText = Array("提申期限", "本所案號", "案件名稱", "案件性質", "承辦人" _
                        , "智權人員", "收文日", "法定期限", "進度備註", "申請人" _
                        , "是否出名", "申請國家", "申請人國籍", "發文日", "代理人" _
                        , "彼所案號", "申請案號", "PA26", "PA27", "PA28", "PA29", "PA30", "PA75")
      arrGridHeadWidth = Array(1050, 1300, 800, 800, 850 _
                        , 850, 880, 880, 880, 1000 _
                        , 800, 800, 1000, 700, 810 _
                        , 800, 800, 0, 0, 0, 0, 0, 0)
   Else
   'end 2018/9/26
      
      'Modified by Lydia 2019/11/01 +PA26~PA30, PA75
      arrGridHeadText = Array("本所期限", "本所案號", "案件名稱", "案件性質", "承辦人" _
                        , "智權人員", "收文日", "法定期限", "進度備註", "申請人" _
                        , "是否出名", "申請國家", "申請人國籍", "發文日", "代理人" _
                        , "彼所案號", "申請案號", "PA26", "PA27", "PA28", "PA29", "PA30", "PA75")
      arrGridHeadWidth = Array(810, 1300, 800, 800, 850 _
                        , 850, 700, 810, 810, 1000 _
                        , 800, 800, 1000, 700, 810 _
                        , 800, 800, 0, 0, 0, 0, 0, 0)
   End If
                        
   grdDataList.Cols = UBound(arrGridHeadText) + 1
End Sub

Sub Process(Optional ByVal p_Sys As String = "")
Dim StrFa As String, stCon As String, stFDate As String, stTDate As String
Dim dblRow As Double 'Add By Sindy 2025/9/3
   
'Added by Lydia 2019/11/01
m_AllSys = IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(, frm100106_1.txt5(0).Text))
intCufaCnt = 0
'end 2019/11/01

   ClearQueryLog ("frm100106_1") 'Add By Sindy 2010/11/3 清除查詢印表記錄檔欄位
   'Add by Morgan 2005/10/7
   '統計起始日期抓系統前推3個工作天
   stFDate = TransDate(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)), -3), 2)
   stTDate = CompDate(1, -1, stFDate)
   pub_QL05 = pub_QL05 & ";[已發文未收達]:統計本所期限為系統日前3個工作天止前一個月：" & Val(stTDate) - 19110000 & "-" & Val(stFDate) - 19110000 'Add By Sindy 2010/11/3
   
   StrFa = ",DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人"
   If p_Sys <> "" Then
      stCon = " AND CP01||'' IN (" & p_Sys & ")"
      pub_QL05 = pub_QL05 & ";" & Left(frm100106_1.Label1(4), 5) & frm100106_1.txt5(0) 'Add By Sindy 2010/11/3
   End If
   'Modify by Morgan 2005/10/7 統計起始日期
   '" AND CP06<=TO_CHAR(SYSDATE-3,'YYYYMMDD') AND CP06>=TO_CHAR(ADD_MONTHS(SYSDATE-3,-1),'YYYYMMDD')"
   
   '2010/9/14 MODIFY BY SONIA 日期欄改百年日期排序問題
   'Modified by Morgan 2015/1/29 排除發文日為11111者 --慧汶
   'Modified by Lydia 2019/11/01 利益衝突案件：增加申請人1~5,FC代理人
   'strSql = "SELECT SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號" & _
      ",NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人" & _
      ",NVL(S2.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP07),-9) AS 法定期限,CP64 AS 進度備註" & _
      ",SUBSTRB(CU04,1,10) AS 申請人,DECODE(CP22,'N','否','是') AS 是否出名,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍" & _
      ",SUBSTR(' '||sqldatet(CP27),-9) AS 發文日" & StrFa & ",DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號" & _
      " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,STAFF S1,STAFF S2,CUSTOMER,FAGENT,CASEPROPERTYMAP,SYSTEMKIND " & _
      " WHERE CP57 IS NULL AND CP27>19221111 AND CP46 IS NULL AND CP47 IS NULL AND CP24 IS NULL" & _
      " AND ((CP01='P' AND CP10 IN ('205','204','107','804','408')) OR (CP01='CFP' AND CP09<'B'))" & _
      " AND CP06<=" & stFDate & " AND CP06>=" & stTDate & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57 IS NULL AND PA09<>'000'" & _
      " AND N1.NA01(+)=PA09 AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1) AND N2.NA01(+)=CU10" & _
      " AND FA01(+)=SUBSTR(CP44,1,8) AND FA02(+)=SUBSTR(CP44,9,1)" & _
      " AND S1.ST01(+)=CP14 AND S2.ST01(+)=CP13" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=TO_CHAR(CP10) AND SK01(+)=CP01" & stCon & _
      " ORDER BY 1,12,1,4,3"
   'Modified by Lydia 2021/05/24 限制欄位長度: CP64 AS 進度備註=> substr(CP64,1,500) AS 進度備註
   strSql = "SELECT SUBSTR(' '||sqldatet(CP06),-9) AS 本所期限,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號" & _
      ",NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人" & _
      ",NVL(S2.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP07),-9) AS 法定期限,substr(CP64,1,500) AS 進度備註" & _
      ",SUBSTRB(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),1,10) AS 申請人,DECODE(CP22,'N','否','是') AS 是否出名,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍" & _
      ",SUBSTR(' '||sqldatet(CP27),-9) AS 發文日" & StrFa & ",DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號" & _
      ", PA26 , PA27 , PA28 , PA29 , PA30 , PA75 " & _
      " FROM CASEPROGRESS,PATENT,NATION N1,NATION N2,STAFF S1,STAFF S2,CUSTOMER,FAGENT,CASEPROPERTYMAP,SYSTEMKIND " & _
      " WHERE CP57 IS NULL AND CP27>19221111 AND CP46 IS NULL AND CP47 IS NULL AND CP24 IS NULL" & _
      " AND ((CP01='P' AND CP10 IN ('205','204','107','804','408')) OR (CP01='CFP' AND CP09<'B'))" & _
      " AND CP06<=" & stFDate & " AND CP06>=" & stTDate & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57 IS NULL AND PA09<>'000'" & _
      " AND N1.NA01(+)=PA09 AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1) AND N2.NA01(+)=CU10" & _
      " AND FA01(+)=SUBSTR(CP44,1,8) AND FA02(+)=SUBSTR(CP44,9,1)" & _
      " AND S1.ST01(+)=CP14 AND S2.ST01(+)=CP13" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=TO_CHAR(CP10) AND SK01(+)=CP01" & stCon & _
      " ORDER BY 1,12,1,4,3"
On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      'Modified by Lydia 2019/11/01 改變型態
      '.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      .Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
      
      If .RecordCount > 0 Then
         dblRow = .RecordCount 'Add By Sindy 2025/9/3

        'Added by Lydia 2019/11/01 逐案號判斷
        If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            .MoveFirst
            Do While .EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & .Fields(1), "" & .Fields("pa26") & "," & .Fields("pa27") & "," & .Fields("pa28") & "," & .Fields("pa29") & "," & .Fields("pa30"), "" & .Fields("pa75")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    .Delete
                End If
                .MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
               pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
               MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow) 'Add By Sindy 2010/11/3
            If .RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
        Else
            InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/3
        End If
        'end 2019/11/01
       
         Set grdDataList.Recordset = adoRecordset
         SetDataListWidth
         Me.Show
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/11/3
JumpToNoData:   'Added by Lydia 2019/11/01
         'Modified by Morgan 2018/9/26 增加客戶函已齊備未判發清單
         'Me.Hide
         GoNext
         'end 2018/9/26
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description
   
End Sub

'Added by Morgan 2020/5/13
'已發文未提申
Sub Process1(Optional ByVal p_Sys As String = "")
Dim strCol1 As String, StrFa As String, stCon As String, stFDate As String, stTDate As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

m_AllSys = IIf(frm100106_1.txt5(0).Text <> "ALL", frm100106_1.txt5(0).Text, GetAllSysKind(, frm100106_1.txt5(0).Text))
intCufaCnt = 0

   ClearQueryLog ("frm100106_1") '清除查詢印表記錄檔欄位
   
   StrFa = ",DECODE(SK03,0,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),DECODE(FA05,NULL,NVL(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65)) as 代理人"
   If p_Sys <> "" Then
      stCon = " AND CP01||'' IN (" & p_Sys & ")"
      pub_QL05 = pub_QL05 & ";" & Left(frm100106_1.Label1(4), 5) & frm100106_1.txt5(0)
   End If
   
   'Modified by Morgan 2020/5/18 期限要比照查詢條件--王副總,郭雅娟
   'stFDate = CompDate(0, -1, strSrvDate(1))
   'stTDate = CompDate(2, -1, CompWorkDay(4, strSrvDate(1)))
   'Me.Caption = "已發文未提申(2個工作天後達最終/指定提申期限)"
   'stCon = stCon & " and NP09>=" & stFDate & " AND NP09<=" & stTDate
   If frm100106_1.opt1(0).Value = True Then
      stCon = stCon & " and NP08>=" & TransDate(frm100106_1.txt1(0), 2) & " AND NP08<=" & TransDate(frm100106_1.txt1(1), 2)
      Me.Caption = "已發文未提申(最終/指定提申本所期限:" & frm100106_1.txt1(0) & "-" & frm100106_1.txt1(1) & ")"
      strCol1 = "SUBSTR(' '||sqldatet(NP08),-9)||decode(np07,'995',' 指')"
   Else
      stCon = stCon & " and NP09>=" & TransDate(frm100106_1.txt2(0), 2) & " AND NP09<=" & TransDate(frm100106_1.txt2(1), 2)
      Me.Caption = "已發文未提申(最終/指定提申法定期限:" & frm100106_1.txt2(0) & "-" & frm100106_1.txt2(1) & ")"
      strCol1 = "SUBSTR(' '||sqldatet(NP09),-9)||decode(np07,'995',' 指')"
   End If
   'end 2020/5/18
   pub_QL05 = pub_QL05 & ";" & Me.Caption
   
   'Modified by Morgan 2020/5/15 要排除FMP寰華案
   strSql = "SELECT " & strCol1 & " AS 提申期限,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS 本所案號" & _
      ",NVL(PA05,NVL(PA06,PA07)) AS 案件名稱,DECODE(PA09,'000',CPM03,CPM04) AS 案件性質,NVL(S1.ST02,CP14) AS 承辦人" & _
      ",NVL(S2.ST02,CP13) AS 智權人員,SUBSTR(' '||sqldatet(CP05),-9) AS 收文日,SUBSTR(' '||sqldatet(CP07),-9) AS 法定期限,CP64 AS 進度備註" & _
      ",SUBSTRB(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),1,10) AS 申請人,DECODE(CP22,'N','否','是') AS 是否出名,N1.NA03 AS 申請國家,N2.NA03 AS 申請人國籍" & _
      ",SUBSTR(' '||sqldatet(CP27),-9) AS 發文日" & StrFa & ",DECODE(PA09,'000',PA77,CP45) AS 彼所案號,PA11 AS 申請案號" & _
      ", PA26 , PA27 , PA28 , PA29 , PA30 , PA75 " & _
      " FROM nextprogress,CASEPROGRESS,staff S3,PATENT,NATION N1,NATION N2,STAFF S1,STAFF S2,CUSTOMER,FAGENT,CASEPROPERTYMAP,SYSTEMKIND " & _
      " WHERE np06 is null AND NP02 in ('P','CFP') and np07 in ('995','996')" & _
      " and cp09(+)=np01 and CP159=0 AND CP158>0 and S3.st01(+)=cp83 and S3.st03<>'F22'" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57 IS NULL AND PA09<>'000'" & _
      " AND N1.NA01(+)=PA09 AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1) AND N2.NA01(+)=CU10" & _
      " AND FA01(+)=SUBSTR(CP44,1,8) AND FA02(+)=SUBSTR(CP44,9,1)" & _
      " AND S1.ST01(+)=CP14 AND S2.ST01(+)=CP13" & _
      " AND CPM01(+)=CP01 AND CPM02(+)=TO_CHAR(CP10) AND SK01(+)=CP01" & stCon & _
      " ORDER BY 1,12,1,4,3"
On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
      
      If .RecordCount > 0 Then
         dblRow = .RecordCount 'Add By Sindy 2025/9/3

        'Added by Lydia 2019/11/01 逐案號判斷
        If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            .MoveFirst
            Do While .EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & .Fields(1), "" & .Fields("pa26") & "," & .Fields("pa27") & "," & .Fields("pa28") & "," & .Fields("pa29") & "," & .Fields("pa30"), "" & .Fields("pa75")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    .Delete
                End If
                .MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow)
            If .RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
        Else
            InsertQueryLog (.RecordCount)
        End If
        'end 2019/11/01
       
         Set grdDataList.Recordset = adoRecordset
         SetDataListWidth
         Me.Show
      Else
         InsertQueryLog (0)
JumpToNoData:
         GoNext
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description
   
End Sub

'Added by Morgan 2018/9/26
Private Sub GoNext()
   If m_iNextStep = 2 Then
      Me.Hide
      
   'Added by Morgan 2020/5/13
   ElseIf m_iNextStep = 0 Then
      m_iNextStep = 1
      SetDataListWidth True
      Me.cmdPrint.Value = False
      Process1
      
   ElseIf m_iNextStep = 1 Then
      m_iNextStep = 2
      cmdPrint.Visible = False
      SetDataListWidth True
      Me.Caption = "客戶函已齊備逾3個工作天未判發"
      Me.cmdPrint.Value = False
      Process2
   End If
End Sub
'Added by Morgan 2018/9/26
'客戶函已齊備逾3個工作天未判發
Private Sub Process2()
   Dim StrFa As String, stCon As String, stFDate As String, stTDate As String
   Dim dblRow As Double 'Add By Sindy 2025/9/3
   
   intCufaCnt = 0 'Added by Lydia 2019/11/01
   
   ClearQueryLog ("frm100106_1") '清除查詢印表記錄檔欄位
   '統計起始日期抓系統前推3個工作天
   stFDate = TransDate(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)), -3), 2)
   pub_QL05 = pub_QL05 & ";[客戶函已齊備逾3個工作天未判發]:齊備日<=" & (Val(stFDate) - 19110000)
   
   'Modified by Morgan 2019/7/3 +判斷有發文日(非假發文但CFP已提申例外) Ex:CFP-29366(CA8038918),等正本來才要通知,先取消發文日
   'Modified by Lydia 2019/11/01 利益衝突案件：增加申請人1~5,FC代理人
   'strSql = "select sqldatet(lp03) 齊備日" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",NVL(PA05,NVL(PA06,PA07)) AS 案件名稱" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質" & _
      ",NVL(ST02,LP04) 判發人" & _
      " from letterprogress l,caseprogress,staff,casepropertymap,patent" & _
      " where lp03>0 and lp03<" & stFDate & " and lp05=0 and lp10='Y' and lp04 is not null" & _
      " and cp09(+)=lp01 and (cp27>19221111 or cp01||cp10='CFP1909') and st01(+)=lp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 order by 1,2"
   'Modified by Morgan 2020/1/8 +LP43判斷(因CFP要工程師判發的已提申不一定有客戶函)
   strSql = "select sqldatet(lp03) 齊備日" & _
      ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",NVL(PA05,NVL(PA06,PA07)) AS 案件名稱" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質" & _
      ",NVL(ST02,LP04) 判發人, PA26 , PA27 , PA28 , PA29 , PA30 , PA75" & _
      " from letterprogress l,caseprogress,staff,casepropertymap,patent" & _
      " where lp03>0 and lp03<" & stFDate & " and lp05=0 and NVL(lp10,LP43)='Y' and lp04 is not null" & _
      " and cp09(+)=lp01 and (cp27>19221111 or cp01||cp10='CFP1909') and st01(+)=lp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 order by 1,2"
      
On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      'Modified by Lydia 2019/11/01 改變型態
      '.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      .Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
      
      If .RecordCount > 0 Then
         dblRow = .RecordCount 'Add By Sindy 2025/9/3

         'Added by Lydia 2019/11/01 逐案號判斷
         If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            .MoveFirst
            Do While .EOF = False
                '利益衝突案件：逐案號判斷
                If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & .Fields(1), "" & .Fields("pa26") & "," & .Fields("pa27") & "," & .Fields("pa28") & "," & .Fields("pa29") & "," & .Fields("pa30"), "" & .Fields("pa75")) = False Then
                    intCufaCnt = intCufaCnt + 1
                    .Delete
                End If
                .MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
               pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
               MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow)
            If .RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
         Else
            InsertQueryLog (.RecordCount)
         End If
        'end 2019/11/01
         
         Set grdDataList.Recordset = adoRecordset
         SetDataListWidth
         Me.Show
      Else
         InsertQueryLog (0)
JumpToNoData:   'Added by Lydia 2019/11/01
         GoNext
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100106_7 = Nothing
End Sub


