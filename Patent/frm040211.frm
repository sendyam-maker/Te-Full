VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040211 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利處程序期限通知"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9390
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm040211.frx":0000
      Left            =   6120
      List            =   "frm040211.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   10
      Top             =   510
      Width           =   3195
   End
   Begin VB.ComboBox cboPrinter 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   915
      TabIndex        =   8
      Top             =   5025
      Width           =   3675
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   400
      Left            =   4455
      TabIndex        =   7
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度(&C)"
      Height          =   400
      Index           =   3
      Left            =   7740
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件基本資料(&B)"
      Height          =   400
      Index           =   2
      Left            =   6210
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   60
      Width           =   1500
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5355
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8550
      TabIndex        =   0
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4095
      Left            =   45
      TabIndex        =   1
      Top             =   870
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   7223
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V|本所期限 |法定期限 |智權人員 |承辦人 |本所案號　　|案件性質 |備註　　    　  　|案件名稱                 "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin MSForms.ComboBox f2Cbo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   12
      Top             =   480
      Width           =   2055
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3625;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "顏色符號說明："
      Height          =   180
      Left            =   4815
      TabIndex        =   11
      Top             =   570
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   5055
      Width           =   975
   End
   Begin VB.Label lblCnt 
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7620
      TabIndex        =   6
      Top             =   5160
      Width           =   1710
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Left            =   135
      TabIndex        =   5
      Top             =   555
      Width           =   900
   End
End
Attribute VB_Name = "frm040211"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2018/10/15 內專程序期限管制
'Memo by Lydia 2018/10/15 員工下拉選單為Form 2.0
'Memo by Lydia 2018/10/26 因與未發文案件查詢相同功能，故於10/26決定取消此新功能。
                                        '原未發文案件查詢frm040210改為所有內專程序人員(P12)進入都要自動執行。
Option Explicit
Dim bolBarShow As Boolean
Public cmdState As Integer
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_stSort As String '排序方式
Dim m_adoRst As ADODB.Recordset

Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strPrinter As String
Dim m_iCols As Integer, m_iPrtCols As Integer
Dim stDept As String

Private Sub SetRst2Grid()
   grdDataList.FixedCols = 0
   Set grdDataList.Recordset = m_adoRst
   grdDataList.FixedCols = 3
End Sub
Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Dim i As Integer, StrTag As String, lngColor As Long, ii As Integer
   Dim StrToMail(1 To 6) As String
   
On Error GoTo ErrorHandler
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   For i = 1 To grdDataList.Rows - 1
      grdDataList.col = 0
      grdDataList.row = i
      If Trim(grdDataList.Text) = "V" Then
         grdDataList.col = 0
         grdDataList.Text = ""
         grdDataList.CellBackColor = grdDataList.BackColor
         grdDataList.col = 3
         lngColor = grdDataList.CellBackColor
         For ii = 1 To 2
            grdDataList.col = ii
            grdDataList.CellBackColor = lngColor
         Next
         
         Dim Str01 As String
         
         StrTag = grdDataList.TextMatrix(i, 5)
         StrTag = Pub_RplStr(StrTag) '清掉符號
         If InStr(StrTag, "-") > 0 Then
             If InStrRev(StrTag, "-") < 6 Then StrTag = StrTag & "-0-00"
         End If
         
         Str01 = SystemNumber(StrTag, 1)
         If fnSaveParentForm(Me) = False Then
            Exit For
         End If
         Me.Show
         Select Case cmdState
            Case 2 '案件基本資料
               Select Case Pub_RplStr(Str01)
                  Case "CFP", "FCP", "P"   '專利
                     frm100101_3.Show
                     frm100101_3.Tag = StrTag
                     frm100101_3.StrMenu
                     
                  Case "FG"
                     frm100101_B.Show
                     frm100101_3.Tag = StrTag
                     frm100101_B.StrMenu
               End Select
               
            Case 3 '案件進度
               frm100101_2.Show
               frm100101_2.Tag = StrTag
               frm100101_2.StrMenu
          End Select
         Exit For
      End If
   Next i
   
ErrorHandler:
   If Err.NUMBER <> 0 Then
      MsgBox "(" & Err.NUMBER & ")" & Err.Description
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub cmdPrint_Click()
   If Val(Mid(lblCnt.Caption, 3, 1)) = 0 Then
       MsgBox "無資料可供列印 !", vbCritical
       Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   PUB_RestorePrinter cboPrinter.Text
   DoPrint
   PUB_RestorePrinter strPrinter
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Public Sub cmdQuery_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   doQuery
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'語法內有用組合欄位為條件以控制使用特定index(避掉某些不適當的)
Private Sub doQuery()
   Dim ii As Integer
   Dim stUserID As String
   Dim stConCP06 As String, stDate1 As String, stDate2 As String
   
   Call SetGrid(True) '清空
   
   'stDate1 = strSrvDate(1) - 10000
   stDate2 = CompWorkDay(4, strSrvDate(1))
   'stConCP06 = " AND CP06>=" & stDate1 & " AND CP06< " & stDate2
   stConCP06 = " AND CP06< " & stDate2
   
   stUserID = Trim(Left(f2Cbo1.Text, 6))
   stDept = GetST15(stUserID)

   
   '清除暫存檔
   strSql = "delete R060206 where R01='" & strUserNum & "'"
   cnnConnection.Execute strSql, intI

   '專利
    strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
       " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,'' as NA16,CP14,CP13,CP06,CP07,CP48,NULL,0,PA75" & _
       " From CASEPROGRESS,PATENT" & _
       " Where CP158=0 AND CP159=0 AND CP14 ='" & stUserID & "'" & stConCP06 & _
       " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA57||PA108 IS NULL AND PA01 IS NOT NULL"
    cnnConnection.Execute strSql, intI
   '服務
    strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
       " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,'' as NA16,CP14,CP13,CP06,CP07,CP48,NULL,0,SP26" & _
       " From CASEPROGRESS, SERVICEPRACTICE" & _
       " Where CP158=0 AND CP159=0 AND CP14 ='" & stUserID & "'" & stConCP06 & _
       " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP15||SP61 IS NULL AND SP01 IS NOT NULL"
    cnnConnection.Execute strSql, intI
   '法務
   strSql = "INSERT INTO R060206(R01,R02,R03,R04,R05,R06,R07,R08,R09,R10,R11,R12,R13)" & _
      " SELECT '" & strUserNum & "','A' EV1,'2' EV2,CP09,'' as NA16,CP14,CP13,CP06,CP07,CP48,NULL,0,LC22 " & _
      " From CASEPROGRESS,LawCase" & _
      " WHERE CP158=0 AND CP159=0 AND CP14 ='" & stUserID & "'" & stConCP06 & _
      " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC08||LC34 IS NULL AND LC01 IS NOT NULL"
   cnnConnection.Execute strSql, intI
   
   '讀取暫存檔
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',PA09),0,CPM03,CPM04) => DECODE(PA09,'000',CPM03,CPM04)
   strExc(0) = " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",S3.ST02 智權人員,S2.ST02 承辦人,CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(PA09,'000',CPM03,CPM04) 案件性質,CP64 備註,NVL(PA05,NVL(PA06,PA07)) 案件名稱" & _
      ",R02,R03,CP01,CP02,CP03,CP04,R05,R07,R06,R04,CP10" & _
      " FROM R060206,CASEPROGRESS,PATENT,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,FAGENT,NATION" & _
      " WHERE R01='" & strUserNum & "' AND R02='A' AND CP09(+)=R04" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND PA01 IS NOT NULL" & _
      " AND S1.ST01(+)=R05 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND FA01(+)=SUBSTR(R13,1,8) AND FA02(+)=SUBSTR(R13,9) AND NA01(+)=FA10"
   
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',SP09),0,CPM03,CPM04) => DECODE(SP09,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
      " SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",S3.ST02 智權人員,S2.ST02 承辦人,CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(SP09,'000',CPM03,CPM04) 案件性質,CP64 備註,NVL(SP05,NVL(SP06,SP07)) 案件名稱" & _
      ",R02,R03,CP01,CP02,CP03,CP04,R05,R07,R06,R04,CP10" & _
      " FROM R060206,CASEPROGRESS,SERVICEPRACTICE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,FAGENT,NATION" & _
      " WHERE R01='" & strUserNum & "' AND R02='A' AND CP09(+)=R04" & _
      " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04 AND SP01 IS NOT NULL" & _
      " AND S1.ST01(+)=R05 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND FA01(+)=SUBSTR(R13,1,8) AND FA02(+)=SUBSTR(R13,9) AND NA01(+)=FA10"
   'Modified by Lydia 2022/06/24 判斷澳門案;  + 044
   'Modified by Lydia 2022/06/28 改判斷非000抓CPM04; DECODE(INSTR('020,013,044',LC15),0,CPM03,CPM04) => DECODE(LC15,'000',CPM03,CPM04)
   strExc(0) = strExc(0) & " UNION " & _
      "SELECT '' V,NVL(lpad(SQLDateT(R08),9,' '),'2') 本所期限,NVL(lpad(SQLDateT(R09),9,' '),'2') 法定期限" & _
      ",S3.ST02 智權人員,S2.ST02 承辦人,CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) 本所案號" & _
      ",DECODE(LC15,'000',CPM03,CPM04) 案件性質,CP64 備註,NVL(LC05,NVL(LC06,LC07)) 案件名稱" & _
      ",R02,R03,CP01,CP02,CP03,CP04,R05,R07,R06,R04,CP10" & _
      " FROM R060206,CASEPROGRESS,LAWCASE,STAFF S1,STAFF S2,STAFF S3,CASEPROPERTYMAP,FAGENT,NATION" & _
      " WHERE R01='" & strUserNum & "' AND R02='A' AND CP09(+)=R04" & _
      " AND LC01(+)=CP01 AND LC02(+)=CP02 AND LC03(+)=CP03 AND LC04(+)=CP04 AND LC01 IS NOT NULL" & _
      " AND S1.ST01(+)=R05 AND S2.ST01(+)=R06 AND S3.ST01(+)=R07 AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
      " AND FA01(+)=SUBSTR(R13,1,8) AND FA02(+)=SUBSTR(R13,9) AND NA01(+)=FA10"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   
   If RsTemp Is Nothing Then Exit Sub
   If RsTemp.RecordCount = 0 Then
      Set m_adoRst = RsTemp.Clone
      MsgBox "查無資料！", vbInformation
      lblCnt.Caption = "共 0 筆"
   Else
      Set m_adoRst = PUB_CreateRecordset(RsTemp, , , 300, Me.Name)
      m_adoRst.Sort = "本所期限 asc,本所案號 asc"
      SetRst2Grid
      SetGrid
      RecordShow
      SetColor
      m_blnColOrderAsc = True
   End If
End Sub

Private Sub SetGrid(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

   With grdDataList
      .Visible = False
       arrGridHeadText = Array("V", "本所期限", "法定期限", "智權人員", "承辦人", "本所案號", "案件性質", "備註", "案件名稱")
       arrGridHeadWidth = Array(300, 900, 900, 900, 840, 1300, 1000, 1300, 1500)

       .Visible = False
       .Cols = UBound(arrGridHeadText) + 1
       If pReset = True Then
              .Clear
              .Rows = 2
       End If
       For iRow = 0 To .Cols - 1
           .row = 0
           .col = iRow
           .Text = arrGridHeadText(iRow)
           If iRow <= UBound(arrGridHeadWidth) Then
                .ColWidth(iRow) = arrGridHeadWidth(iRow)
           Else '案件名稱以後的欄位
                .ColWidth(iRow) = 0
           End If
           .CellAlignment = 0
       Next
      .ColAlignment(1) = flexAlignRightTop
      .ColAlignment(2) = flexAlignRightTop
    
      .Visible = True
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   f2Cbo1.AddItem strUserNum & " " & strUserName
   '抓同部門人員清員
   If Pub_StrUserSt15 <> "" Then
        strSql = "select st01,st02 from staff where st01<>" & CNULL(strUserNum) & " and st04='1' and st15=" & CNULL(Pub_StrUserSt15) & _
                    " and st01>='69000' and st01<'F0000' order by st05,st01"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
                 f2Cbo1.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                 RsTemp.MoveNext
            Loop
        End If
   End If
   f2Cbo1.ListIndex = 0
   
   PUB_SetPrinter Me.Name, cboPrinter, strPrinter
   
   PUB_AddExcuteLog Me.Name
   
   Combo1.Clear
   Combo1.AddItem "紅色(*)：表示逾管控期限"
   Combo1.AddItem "綠色(v): 表示當日期限"
   Combo1.AddItem "藍色: 表示點選資料"
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If Not bolUnloading Then
      If cboPrinter.Text <> cboPrinter.Tag Then
         PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cboPrinter.Name, "0", "0", Me.cboPrinter.Text
      End If
   End If
   Set frm040211 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim iCol As Integer
    iCol = grdDataList.MouseCol
    If grdDataList.MouseRow < 1 Then
      grdDataList.Visible = False
      ChgEmptyDate True
      Set grdDataList.Recordset = Nothing
      If m_blnColOrderAsc = True Then
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " desc," & m_stSort
         m_blnColOrderAsc = False
      Else
         m_adoRst.Sort = m_adoRst.Fields(iCol).Name & " asc," & m_stSort
         m_blnColOrderAsc = True
      End If
      SetRst2Grid
      SetColor
      grdDataList.Visible = True
    End If
End Sub

Private Sub ChgEmptyDate(Optional p_bolBeforeSort As Boolean)
   Dim ii As Integer, jj As Integer
   With grdDataList
   If .Rows > 1 Then
      For ii = 1 To .Rows - 1
         For jj = 1 To 3
            If p_bolBeforeSort Then
               If .TextMatrix(ii, jj) = "" Then
                  .TextMatrix(ii, jj) = "2"
               End If
            Else
               If .TextMatrix(ii, jj) = "2" Then
                  .TextMatrix(ii, jj) = ""
               End If
            End If
         Next
      Next
   End If
   End With
End Sub

Private Sub grdDataList_SelChange()
   Dim ii As Integer, lngColor As Long
   With grdDataList
      If .MouseRow > 0 Then
         .Visible = False
         .row = .MouseRow
         .col = 0
         If .Text = "V" Then
            .Text = ""
            .col = 0
            .CellBackColor = .BackColor
            .col = 3
            lngColor = .CellBackColor
            For ii = 1 To 2
               .col = ii
               .CellBackColor = lngColor
            Next
         Else
            .Text = "V"
            For ii = 0 To 2
               .col = ii
               .CellBackColor = &HFFC0C0
            Next
         End If
         .Visible = True
      End If
   End With
End Sub

Private Sub SetColor()
   Dim lngToday As Long, lngCP06 As Long
   Dim ii As Integer, jj As Integer, dblCnt As Double
   
   dblCnt = 0
   With grdDataList
   If .Rows > 1 Then
      .Visible = False
      ChgEmptyDate False
      lngToday = Val(strSrvDate(2))
      For ii = 1 To .Rows - 1
         .row = ii
         '固定欄位變回白色
         For jj = 0 To 2
            .col = jj
            .CellBackColor = &HFFFFFF
            .CellAlignment = flexAlignRightTop
            .CellFontSize = 9
         Next
        .RowHeight(ii) = 255
        lngCP06 = Val(Replace(.TextMatrix(ii, 1), "/", ""))
        
        '逾管控期限
        If lngCP06 > 0 And lngCP06 < lngToday Then
           .TextMatrix(ii, 5) = "*" & .TextMatrix(ii, 5)
           For jj = 1 To .Cols - 1
              .col = jj
              '紅
              .CellBackColor = &HFF&
           Next
        '當日期限
        ElseIf lngCP06 > 0 And lngCP06 = lngToday Then
           .TextMatrix(ii, 5) = "v" & .TextMatrix(ii, 5)
           For jj = 1 To .Cols - 1
              .col = jj
              '綠
              .CellBackColor = &HC000&
           Next
        End If
        
        If .RowHeight(ii) > 0 Then
            dblCnt = dblCnt + 1
        End If
      Next
      .TopRow = 1
      .Visible = True
   End If
   End With

   lblCnt.Caption = "共 " & dblCnt & " 筆"

End Sub

Private Sub DoPrint()

   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   Dim strTemp() As String
   
   iOrientation = Printer.Orientation
   Printer.Orientation = 2
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With grdDataList
      GetPleft
      
      m_iPrtCols = m_iCols
      ReDim strTemp(1 To m_iPrtCols)
      
      iPage = 1
      PrintPageHeader
      PrintPageHeader1
      For iRow = 1 To .Rows - 1
         For iCol = LBound(strTemp) To UBound(strTemp)
            If iCol = 6 Then '案件性質
               strTemp(iCol) = Left(.TextMatrix(iRow, iCol), 7)
            ElseIf iCol = 7 Then '備註
               strTemp(iCol) = convForm(.TextMatrix(iRow, iCol), 24)
            ElseIf iCol = 8 Then '案件名稱
               strTemp(iCol) = convForm(.TextMatrix(iRow, iCol), 30)
            Else
               strTemp(iCol) = .TextMatrix(iRow, iCol)
            End If
         Next
         PrintDetail strTemp
      Next
      Call PrintReportFooter(.Rows - 1)
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
   
End Sub

Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   m_iCols = 8
   ReDim PLeft(1 To m_iCols)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + Printer.TextWidth("本所期限") + ciColGap
   PLeft(3) = PLeft(2) + Printer.TextWidth("法定期限") + ciColGap
   PLeft(4) = PLeft(3) + Printer.TextWidth("智權人員") + ciColGap
   PLeft(5) = PLeft(4) + Printer.TextWidth("承 辦 人") + ciColGap
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(14, "A")) + ciColGap '本所案號
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(12, "A")) + ciColGap '案件性質
   PLeft(8) = PLeft(7) + Printer.TextWidth(String(24, "A")) + ciColGap '備註

End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      Printer.Print String(130, "-")
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
    
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = LBound(strData) To UBound(strData)
      If Me.grdDataList.ColWidth(iCol) > 0 Then
          Printer.CurrentX = PLeft(iCol)
          Printer.CurrentY = iPrint
          Printer.Print strData(iCol)
      End If
    Next
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = Replace(Me.Caption, "通知", "清單")
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   PrintNewLine
   strPTmp = "員工編號：" & Trim(Left(f2Cbo1.Text, 6)) & " 姓名：" & Trim(Mid(f2Cbo1.Text, 7))
   Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(9, "　"))
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print String(130, "-")
End Sub

Sub PrintPageHeader1()
   
    Call PrintNewLine(False, 1)
    For intI = 1 To m_iPrtCols
      If Me.grdDataList.ColWidth(intI) > 0 Then
         Printer.CurrentX = PLeft(intI)
         Printer.CurrentY = iPrint
         Printer.Print grdDataList.TextMatrix(0, intI)
      End If
    Next
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(130, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(130, "-")
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub


