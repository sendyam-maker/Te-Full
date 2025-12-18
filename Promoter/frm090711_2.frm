VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090711_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人工作進度資料維護－當天及前一工作天分案案件資料"
   ClientHeight    =   8880
   ClientLeft      =   -1785
   ClientTop       =   960
   ClientWidth     =   15000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   15000
   Begin VB.CommandButton cmdok 
      Caption         =   "繼續(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   12750
      TabIndex        =   4
      Top             =   330
      Width           =   990
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   7500
      Left            =   150
      TabIndex        =   1
      Top             =   1080
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   13229
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
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
      _Band(0).Cols   =   1
   End
   Begin MSForms.Label lblName 
      Height          =   225
      Left            =   1230
      TabIndex        =   3
      Top             =   540
      Width           =   1935
      VariousPropertyBits=   27
      Caption         =   "lblName"
      Size            =   "3413;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblDate 
      Caption         =   "lblDate"
      Height          =   180
      Left            =   210
      TabIndex        =   2
      Top             =   810
      Width           =   3555
   End
   Begin VB.Label Label1 
      Caption         =   "繪圖人員： "
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   540
      Width           =   915
   End
End
Attribute VB_Name = "frm090711_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (grd1,lblName)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Public TextOk As Boolean, StrGrp090711 As String
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 21) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, strDate1 As String, StrDate2 As String
Dim ChkData2 As Boolean, SWPRow As String, strCP10 As String, k As Integer, ChkNoData As Boolean, TXT090711 As TextBox
Dim NickRS As ADODB.Recordset, StrColor1 As String, StrColor2 As String, StrColor3 As String, StrColor4 As String, StrColor5 As String, StrColor6 As String
Dim ll As Integer
Dim m_dblWDBegin As Double '工作天起
Dim m_dblWDEnd As Double '工作天迄

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Unload Me
End Select
End Sub

Private Sub Form_Activate()
    If TextOk = False Then
        Unload Me
    Else
        Me.Show
    End If
End Sub

Private Sub Form_Load()
    Me.Hide
    Screen.MousePointer = vbHourglass
    MoveFormToCenter Me
    StrMenu
    SetGrd1
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm090711_2 = Nothing
    frm090711.Show
End Sub

Sub StrMenu() '代資料當月資料
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

Me.lblName.Caption = GetStaffName(strUserNum, True)
m_dblWDEnd = strSrvDate(1)
StrSQLa = "Select * From WorkDay Where WD01<" & m_dblWDEnd & " Order By WD01 Desc "
rsA.CursorLocation = adUseClient
'Add by Morgan 2003/12/31
rsA.MaxRecords = 1
   
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    m_dblWDBegin = Val("" & rsA.Fields(0).Value)
Else
    m_dblWDBegin = m_dblWDEnd
End If
If rsA.State <> adStateClosed Then rsA.Close
'Add by Morgan 2003/12/31
rsA.MaxRecords = 0

Set rsA = Nothing

Me.lblDate.Caption = ChangeTStringToTDateString(m_dblWDBegin - 19110000) & " <= 文件齊備日 <= " & ChangeTStringToTDateString(m_dblWDEnd - 19110000)

StrSQL6 = " AND EP13='" & strUserNum & "' And CP01 in ('FCP','P','CFP') "
'Modify By Sindy 2016/9/5 and cp57 is null and cp27 is null ==> and cp158=0 and cp159=0
StrSQL6 = StrSQL6 + " and cp158=0 and cp159=0 and cp05>=19980101 And (EP06>=" & m_dblWDBegin & " And EP06<=" & m_dblWDEnd & " ) "
strSQL1 = " AND EP13='" & strUserNum & "' And CP01 in ('FCP','P','CFP') "
'edit by nickc 2005/05/13
'strSQL1 = strSQL1 & " and (SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & " or SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & ") and cp05>=19980101 And (EP06>=" & m_dblWDBegin & " And EP06<=" & m_dblWDEnd & " ) "
'Modify By Sindy 2016/9/5 and cp27 is null ==> and cp158=0
strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & ") or (SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & " and cp158=0)) and cp05>=19980101 And (EP06>=" & m_dblWDBegin & " And EP06<=" & m_dblWDEnd & " ) "
'Modify by Morgan 2004/5/19
'加專利種類
'strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'             " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
'Modified by Morgan 2018/4/13 因O12會當欄位補別名後才可執行
strSql = "SELECT SUBSTR(CP09,1,1) X1," & SQLDate("CP05") & " X2,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊','') X3,NVL(PA05,NVL(PA06,PA07)) X4, Decode(pa09,'000',cpm03,cpm04) X5,s1.st02 X6, ROUND(cp18,2) X7,DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & ") X8, EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & ") X9,'' As 草期限, " & SQLDate("eP15") & " X10,0 X11, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & ") X12, '' As 墨期限, " & SQLDate("EP18") & " X13,0 X14," & SQLDate("cP06") & " X15," & SQLDate("cP27") & " X16,ep26,s3.st02 X17,CP09,DECODE(PA09,'000',PTM03,PTM04) X18," & SQLDate("CP07") & " X19," & SQLDate("CP57") & " X20,ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, pa08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
            " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
strSql = strSql & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, pa08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
             " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1

strSql = strSql + " ORDER BY 8 desc, 3 desc "
CheckOC
TextOk = True
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Set GRD1.Recordset = adoRecordset
        GRD1.Visible = False
        Screen.MousePointer = vbHourglass
        Me.Enabled = False
        For i = 1 To GRD1.Rows - 1
            GRD1.row = i
            '文齊日(將空白消除)
            Me.GRD1.TextMatrix(i, 7) = Trim(Me.GRD1.TextMatrix(i, 7))
            '草完日
            GRD1.col = 11
            strDate1 = GRD1.Text
            '草齊日
            GRD1.col = 9
            StrDate2 = GRD1.Text
            If Trim(strDate1) <> "" And Trim(StrDate2) <> "" Then
                '草天
                GRD1.col = 12
                GRD1.Text = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(StrDate2)))
            Else
                '草天
                GRD1.col = 12
                GRD1.Text = ""
            End If
            '墨完日
            GRD1.col = 16
            strDate1 = GRD1.Text
            '墨齊日
            GRD1.col = 14
            StrDate2 = GRD1.Text
            If Trim(strDate1) <> "" And Trim(StrDate2) <> "" Then
                '墨天
                GRD1.col = 17
                GRD1.Text = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(StrDate2)))
            Else
                '墨天
                GRD1.col = 17
                GRD1.Text = ""
            End If
            'Add By Cheng 2003/06/27
            '若草圖不計件
            If GRD1.TextMatrix(i, 8) = "N" Then
                GRD1.TextMatrix(i, 9) = "******"
                GRD1.TextMatrix(i, 11) = "******"
                GRD1.TextMatrix(i, 12) = ""
            End If
            '若墨圖不計件
            If GRD1.TextMatrix(i, 13) = "N" Then
                GRD1.TextMatrix(i, 14) = "******"
                GRD1.TextMatrix(i, 16) = "******"
                GRD1.TextMatrix(i, 17) = ""
            End If
            'Add By Cheng 2003/06/30
            '草圖承辦期限
            If Me.GRD1.TextMatrix(i, 9) <> "" And Me.GRD1.TextMatrix(i, 9) <> "******" Then
                '設計申請
                'Modify by Morgan 2004/5/19
                '設計改用專利種類判斷
                'If Me.grd1.TextMatrix(i, 33) = "103" Or Me.grd1.TextMatrix(i, 33) = "105" Then
                If Me.GRD1.TextMatrix(i, 34) = "3" Then
                    Me.GRD1.TextMatrix(i, 10) = ChangeTStringToTDateString(CompWorkDay(5, Replace(Me.GRD1.TextMatrix(i, 9), "/", "") + 19110000) - 19110000)
                '非設計申請
                Else
                    Me.GRD1.TextMatrix(i, 10) = ChangeTStringToTDateString(CompWorkDay(4, Replace(Me.GRD1.TextMatrix(i, 9), "/", "") + 19110000) - 19110000)
                End If
            End If
            '墨圖承辦期限
            If Me.GRD1.TextMatrix(i, 14) <> "" And Me.GRD1.TextMatrix(i, 14) <> "******" Then
                Me.GRD1.TextMatrix(i, 15) = ChangeTStringToTDateString(CompWorkDay(3, Replace(Me.GRD1.TextMatrix(i, 14), "/", "") + 19110000) - 19110000)
            End If
        Next i
        Me.Enabled = True
        Screen.MousePointer = vbDefault
        SetGrd1
        GRD1.Visible = True
        ChkNoData = False
    Else
         Me.GRD1.Clear
         Me.GRD1.Rows = 2
         SetGrd1
         ChkNoData = True
    TextOk = False
    End If
End With
CheckOC
End Sub

Private Sub SetGrd1()
    With GRD1
        .Visible = False
        'Modify by Morgan 2004/5/19
        '加專利種類
        '.Cols = 34
        .Cols = 35
        .row = 0
        .RowHeight(0) = 400
        .col = 0:   .Text = "類別"
        .ColWidth(0) = 300
        .CellAlignment = flexAlignCenterCenter
        .col = 1:   .Text = "收文日"
        .ColWidth(1) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 2:   .Text = "本所案號"
        .ColWidth(2) = 1400
        .CellAlignment = flexAlignCenterCenter
        .col = 3:   .Text = "案件名稱"
        .ColWidth(3) = 1500
        .CellAlignment = flexAlignCenterCenter
        .col = 4:   .Text = "案件性質"
        .ColWidth(4) = 800
        .CellAlignment = flexAlignCenterCenter
        .col = 5:   .Text = "承辦人"
        .ColWidth(5) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 6:   .Text = "點數"
        .ColWidth(6) = 400
        .CellAlignment = flexAlignCenterCenter
        .col = 7:   .Text = "文齊日"
        .ColWidth(7) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 8:   .Text = "草計"
        .ColWidth(8) = 400
        .CellAlignment = flexAlignCenterCenter
        .col = 9:   .Text = "草齊日"
        .ColWidth(9) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 10:   .Text = "草期限"
        .ColWidth(10) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 11:  .Text = "草完日"
        .ColWidth(11) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 12:  .Text = "草天"
        .ColWidth(12) = 400
        .CellAlignment = flexAlignCenterCenter
        .col = 13:   .Text = "墨計"
        .ColWidth(13) = 400
        .CellAlignment = flexAlignCenterCenter
        .col = 14:  .Text = "墨齊日"
        .ColWidth(14) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 15:  .Text = "墨期限"
        .ColWidth(15) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 16:  .Text = "墨完日"
        .ColWidth(16) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 17:  .Text = "墨天"
        .ColWidth(17) = 400
        .CellAlignment = flexAlignCenterCenter
        .col = 18:  .Text = "本所期限"
        .ColWidth(18) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 19:  .Text = "發文日"
        .ColWidth(19) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 20:  .Text = "備註"
        .ColWidth(20) = 800
        .CellAlignment = flexAlignCenterCenter
        .col = 21:  .Text = "智權人員"
        .ColWidth(21) = 700
        .CellAlignment = flexAlignCenterCenter
        .col = 22:  .Text = "" '收文號
        .ColWidth(22) = 0
        .CellAlignment = flexAlignCenterCenter
        .col = 23:  .Text = "" '案件性質名稱
        .ColWidth(23) = 0
        .CellAlignment = flexAlignCenterCenter
        .col = 24:  .Text = "" '法定期限
        .ColWidth(24) = 0
        .CellAlignment = flexAlignCenterCenter
        .col = 25:  .Text = "" '取消收文日
        .ColWidth(25) = 0
        .CellAlignment = flexAlignCenterCenter
        .col = 26:  .Text = "" '草圖承辦時數
        .ColWidth(26) = 0
        .CellAlignment = flexAlignCenterCenter
        .col = 27:  .Text = "" '墨圖承辦時數
        .ColWidth(27) = 0
        .CellAlignment = flexAlignCenterCenter
        .col = 28:  .Text = "" '修改時數1
        .ColWidth(28) = 0
        .CellAlignment = flexAlignCenterCenter
        .col = 29:  .Text = "" '修改時數2
        .ColWidth(29) = 0
        .CellAlignment = flexAlignCenterCenter
        .col = 30:  .Text = "" '修改時數3
        .ColWidth(30) = 0
        .CellAlignment = flexAlignCenterCenter
        .col = 31:  .Text = "" '草圖張數
        .ColWidth(31) = 0
        .CellAlignment = flexAlignCenterCenter
        .col = 32:  .Text = "" '墨圖張數
        .ColWidth(32) = 0
        .CellAlignment = flexAlignCenterCenter
        .col = 33:  .Text = "" '案件性質代號
        .ColWidth(33) = 0
        .CellAlignment = flexAlignCenterCenter
        
        'Add by Morgan 2004/5/19
        .col = 34:  .Text = "" '專利種類代號
        .ColWidth(34) = 0
        .CellAlignment = flexAlignCenterCenter
        
        .Visible = True
    End With
End Sub
