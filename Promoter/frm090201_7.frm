VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_7 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護－已達管制期限未輸入會稿日案件資料"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdok 
      Caption         =   "繼續(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6450
      TabIndex        =   3
      Top             =   120
      Width           =   850
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "輸入(&O)"
      Height          =   400
      Index           =   1
      Left            =   7290
      TabIndex        =   2
      Top             =   120
      Width           =   850
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   8130
      TabIndex        =   1
      Top             =   120
      Width           =   850
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "全部選取"
      Height          =   400
      Index           =   3
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   1185
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3630
      Left            =   90
      TabIndex        =   4
      Top             =   885
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6403
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSForms.Label lblMemo 
      Height          =   1095
      Left            =   90
      TabIndex        =   8
      Top             =   4590
      Width           =   9105
      Caption         =   "lblMemo"
      Size            =   "16060;1931"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblName 
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   330
      Width           =   1875
      VariousPropertyBits=   27
      Caption         =   "lblName"
      Size            =   "3307;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人： "
      Height          =   180
      Index           =   35
      Left            =   120
      TabIndex        =   6
      Top             =   345
      Width           =   795
   End
   Begin VB.Label lblDate 
      Caption         =   "lblDate"
      Height          =   180
      Left            =   90
      TabIndex        =   5
      Top             =   615
      Width           =   3555
   End
End
Attribute VB_Name = "frm090201_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; grd1改字型=新細明體-ExtB、lblName、lblMemo
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
'create by nickc 2006/02/08 copy from frm090201_5
Option Explicit
Public TextOk As Boolean
'Public Print1Ok As Boolean '2014/4/29 CANCEL BY SONIA
Public strContinue As Boolean    '2009/11/13 add by sonia 判斷為繼續或結束
Dim intLastRow As Integer
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String, SWPColor As String, SWPColor2 As String, SWPRow As String, SWPRow2 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 21) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String, Tmp001 As String, Tmp002 As String, Tmp003 As String, Tmp004 As String, k As Integer
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, SeekTmpBk As String, ChkData As Boolean
Dim strCP10 As String, AdoRs As ADODB.Recordset, StrNewSQL As String, Txt090201 As TextBox, ChkNoData As Boolean
Dim Fobj As FileSystemObject, ChkCp27 As Boolean, StrGrp090201 As String, ChkData2 As Boolean
Dim ADORECORDSET66 As New ADODB.Recordset
Dim Adorecordset99 As New ADODB.Recordset
Dim Intnick910123 As Integer
Dim m_blnFirstShow As Boolean '判斷表單是否第一次顯示

Private Sub cmdOK_Click(Index As Integer)
Dim KeyWord As String

   Select Case Index
      Case 0  '繼續
         strContinue = True   '2009/11/13 add by sonia
         Unload Me
      Case 1
         With grd1
            For i = 1 To .Rows - 1
               .Visible = False
               If .TextMatrix(i, 0) = "V" Then
                  Exit For
               Else
                  If i = .Rows - 1 Then
                     MsgBox "請點選欲輸入 預定會稿日 資料"
                     .Visible = True
                     Exit Sub
                  End If
               End If
            Next
            .Visible = True
         End With
   
         frm090201_8.Show
         Me.Hide
      Case 2 '結束
         strContinue = False   '2009/11/13 add by sonia
         Unload Me
      'add by nick 2004/10/28
      Case 3
            grd1.col = 0
            Screen.MousePointer = vbHourglass
            If Trim(cmdok(3).Caption) = "全部選取" Then
                KeyWord = "V"
                cmdok(3).Caption = "全部取消"
            Else
                KeyWord = ""
                cmdok(3).Caption = "全部選取"
            End If
            For i = 1 To grd1.Rows - 1
                grd1.row = i
                grd1.Text = KeyWord
            Next i
            Screen.MousePointer = vbDefault
   End Select
End Sub

Private Sub Form_Activate()
'2009/11/12 cancel by sonia
'   If TextOk = False Then
'      '2009/11/12 MODIFY BY SONIA
'      'cmdok_Click 0 '繼續
'      Unload Me '繼續
'   Else
'      Me.Show
'   End If
'    If m_blnFirstShow = True Then m_blnFirstShow = False
End Sub

Private Sub Form_Load()
   Me.Hide
   Screen.MousePointer = vbHourglass
   MoveFormToCenter Me
'   Me.Enabled = False
'   StrMenu1
'   Me.Enabled = True
   StrMenu
   SetGrd1
   lblMemo = "管制期限：1.承辦期限前 2 個工作天 ( 適用甲規則者 )" & vbCrLf & _
             "　　　　　2.P案齊備後 10 個工作天 ( 適用乙規則者 )" & vbCrLf & _
             "　　　　　3.P案所限前 10 個工作天 ( 適用乙規則者 )" & vbCrLf & _
             "　　　　　4.CFP案齊備後 15 個工作天 ( 適用乙規則者 )" & vbCrLf & _
             "　　　　　5.CFP案所限前 15 個工作天 ( 適用乙規則者 )"

   Screen.MousePointer = vbDefault
   m_blnFirstShow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090201_7 = Nothing
   If strContinue = True Then Nextstep
End Sub

'2009/11/13 add by sonia
Private Sub Nextstep()
   frm090201_9.TextOk = False
   frm090201_9.StrMenu1  '已完稿超過2天或已屆預定會稿日且尚未會稿案件資料,,無資料時由frm090201_9的nextstep執行下一畫面
   If frm090201_9.TextOk = True Then
      frm090201_9.Show
   End If
End Sub
'2009/11/13 end

Public Sub StrMenu1()
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim p_Date As String
   Dim CFP_date As String
   'Add by Morgan 2008/10/27
   Dim P_DateEP06B As String 'P適用B類規則且以齊備日計算標準日
   Dim P_DateCP06B As String 'P適用B類規則且以本所期限計算標準日
   Dim CFP_DateEP06B As String 'CFP適用B類規則且以齊備日計算標準日
   Dim CFP_DateCP06B As String 'CFP適用B類規則且以本所期限計算標準日

   'adoEng.Execute "DELETE FROM R090614 WHERE ID='" & strUserNum & "' "  '2009/11/12 cancel by sonia
   'add by nickc 2007/12/17
    adoEng.Execute "drop table R090614 "
    adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text,R110006 text,R110007 text,R110008 text,R110009 text,R110010 text,R110011 text,R110012 text,R110013 text,R110014 text,R110015 text,R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo,R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text,R110026 double,R110027 double,R110028 double,R110029 text,R110030 text)"
   
   strContinue = True   '2009/11/13 add by sonia
   DoEvents
   p_Date = CompWorkDay(2, strSrvDate(1), 1)
   CFP_date = CompWorkDay(2, strSrvDate(1), 1)
   'Add by Morgan 2008/10/27
   '齊備後10個工作天(含)
   P_DateEP06B = CompWorkDay(10, strSrvDate(1), 1)
   '所限前10個工作天(含)
   P_DateCP06B = CompWorkDay(10, strSrvDate(1))
   '齊備後15個工作天(含)
   CFP_DateEP06B = CompWorkDay(15, strSrvDate(1), 1)
   '所限前15個工作天(含)
   CFP_DateCP06B = CompWorkDay(15, strSrvDate(1))
   
   '92.6.29 end
   StrGrp090201 = ""
   StrSQL6 = ""
   strSQL1 = ""
   strSQL2 = ""
   '2009/11/6 MODIFY BY SONIA 取消完稿日條件,否則與王協理看的已發文未輸會稿完成日程式控制逾承辦期限條件不同CFP-22602
   'StrSQL6 = StrSQL6 + " and CP14='" & strUserNum & "' AND (CP27 IS NULL  and CP57 IS NULL) and cp05>=19980101 And EP28 IS NULL and ep07 is null and ep09 is null  "
   'Modify By Sindy 2016/9/5 and cp57 is null and cp27 is null ==> and cp158=0 and cp159=0
   StrSQL6 = StrSQL6 + " and CP14='" & strUserNum & "' and cp158=0 and cp159=0 and cp05>=19980101 And EP28 IS NULL and ep07 is null "
   '2009/11/6 END
   
   'Modify by Morgan 2007/8/31 加807
   'StrSQL6 = StrSQL6 & " AND CP10<>'1101' and substr(cp09,1,1) in ('A','C') and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506') "
   'Modify by Morgan 2008/10/27改抓cpm設定
   'StrSQL6 = StrSQL6 & " AND CP10<>'1101' and substr(cp09,1,1) in ('A','C') and cp10 not in ('106','121','201','202','203','204','206','207','211','212','411','416','417','421','901','902','906','909','910','911','916','917','1002','1209','1908','1902','407','920','404','215','214','408','1205','1206','401','413','429','505','506','807') "
   strSQL1 = StrSQL6
   'end 2008/10/27
   'end 2007/8/31
   
   CheckOC
   'Modify by Morgan 2008/12/2 不必管是否適用只要有設定規則的都要(原來CP112='Y'的條件去掉)
   'Modify by Morgan 2010/8/16 百年蟲 " & SQLDate("CP05") & "-->substrb(' '||sqldatet(cp05),-9)
   strSql = "SELECT CP14,EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
                   " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & StrSQL6 & _
                   " AND CP01='P' and cp48<=" & p_Date
   'Add by Morgan 2008/10/3 區分A,B類規則
   strSql = strSql & " AND CPM05='A'"
   
   
   strSql = strSql & " union SELECT CP14,EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
                   " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & StrSQL6 & _
                   " AND CP01='CFP' and cp48<=" & CFP_date
   
   'Add by Morgan 2008/10/3 區分A,B類規則
   strSql = strSql & " AND CPM05='A'"
   
   'Add by Morgan 2008/10/3 +B類規則
   strSql = strSql & " union SELECT CP14,EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02" & _
            " FROM CASEPROGRESS,CASEPROPERTYMAP,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,PATENTTRADEMARKMAP,NATION " & _
            " WHERE CP01='P' AND CP48>0 AND (ep06<=" & P_DateEP06B & " or cp06<=" & P_DateCP06B & " or cp48<=" & p_Date & ")" & _
            " and CP09=EP02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cpm05='B'" & strSQL1 & _
            " and CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+)"
   
   strSql = strSql & " union SELECT CP14,EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02" & _
            " FROM CASEPROGRESS,CASEPROPERTYMAP,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,PATENTTRADEMARKMAP,NATION " & _
            " WHERE CP01='CFP' AND CP48>0 AND (ep06<=" & CFP_DateEP06B & " or cp06<=" & CFP_DateCP06B & " or cp48<=" & CFP_date & ")" & _
            " and CP09=EP02(+) and cp01=cpm01(+) and cp10=cpm02(+) and cpm05='B'" & strSQL1 & _
            " and CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+)"
   
   strSql = strSql + " ORDER BY 1,4 "
   CheckOC
   'Print1Ok = False   '2014/4/29 CANCEL BY SONIA
   TextOk = True
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
         .MoveFirst
         k = 0
         DoEvents
         '2014/4/29 CANCEL BY SONIA
         ''判斷等級是否屬於專利
         'If (Val(CheckStr(.Fields(23))) >= 31 And Val(CheckStr(.Fields(23))) <= 39) Or (Val(CheckStr(.Fields(23))) >= 71 And Val(CheckStr(.Fields(23))) <= 89) Then
         '   Print1Ok = True
         'End If
         '2014/4/29 END
         Do While .EOF = False
            For i = 0 To 21
               'Add by Morgan 2010/8/16
               If i = 3 Then
                  strTemp(i) = "" & .Fields(i)
               Else
               'end 2010/8/16
                  strTemp(i) = CheckStr(.Fields(i))
               End If
            Next i
            'edit by nickc 2007/12/17
            'strSQL = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "','" & .Fields(28).Value & "' ) "
            strSql = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "','" & .Fields(28).Value & "',0,0,0,0,0 ) "
            adoEng.Execute strSql
            .MoveNext
            DoEvents
         Loop
      Else
          TextOk = False
          Nextstep   '2009/11/13 add by sonia
      End If
      CheckOC
   End With
End Sub

Private Sub StrMenu()
   Me.lblName.Caption = strUserName
   Me.lblDate.Caption = ""
   
   strSql = "SELECT ' ',R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024 FROM R090614 " & _
                 " WHERE ID='" & strUserNum & "' AND R110001='" & strUserNum & "' ORDER BY R110002 Desc, R110004 Desc, R110005 Desc "
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, adoEng, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
          Set grd1.Recordset = adoRecordset
          ChkNoData = False
      Else
          ChkNoData = True
          grd1.Clear
          grd1.Rows = 2
          Screen.MousePointer = vbDefault
          Exit Sub
      End If
   End With
   CheckOC
End Sub

Private Sub SetGrd1()
   With grd1
      'Modify by Morgan 2010/8/16 日期欄寬改800
      .Visible = False
      .Cols = 23
      .row = 0
      .col = 0:   .Text = "V"
      .ColWidth(0) = 200
      .CellAlignment = flexAlignCenterCenter
      .col = 1:   .Text = "目次"
      .ColWidth(1) = 350
      .CellAlignment = flexAlignCenterCenter
      .col = 2:   .Text = "收文類別"
      .ColWidth(2) = 200
      .CellAlignment = flexAlignCenterCenter
      .col = 3:   .Text = "收文日"
      .ColWidth(3) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 4:   .Text = "本所案號"
      .ColWidth(4) = 1000
      .CellAlignment = flexAlignCenterCenter
      .col = 5:   .Text = "案件名稱"
      .ColWidth(5) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 6:   .Text = "國家"
      .ColWidth(6) = 450
      .CellAlignment = flexAlignCenterCenter
      .col = 7:   .Text = "種類"
      .ColWidth(7) = 400
      .CellAlignment = flexAlignCenterCenter
      .col = 8:   .Text = "案件性質"
      .ColWidth(8) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 9:   .Text = "Y/N"
      .ColWidth(9) = 300
      .CellAlignment = flexAlignCenterCenter
      .col = 10:   .Text = "承辦期限"
      .ColWidth(10) = 700
      .CellAlignment = flexAlignCenterCenter
      .col = 11:  .Text = "本所期限"
      .ColWidth(11) = 0
      .CellAlignment = flexAlignCenterCenter
      .col = 12:  .Text = "法定期限"
      .ColWidth(12) = 0
      .CellAlignment = flexAlignCenterCenter
      .col = 13:  .Text = "齊備日"
      .ColWidth(13) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 14:  .Text = "完稿日"
      .ColWidth(14) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 15:  .Text = "會稿日"
      .ColWidth(15) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 16:  .Text = "核稿人"
      .ColWidth(16) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 17:  .Text = "會稿完成日"
      .ColWidth(17) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 18:  .Text = "發文日"
      .ColWidth(18) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 19:  .Text = "承辦天數"
      .ColWidth(19) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 20:  .Text = "備註"
      .ColWidth(20) = 2000
      .CellAlignment = flexAlignCenterCenter
      .col = 21:  .Text = "智權人員"
      .ColWidth(21) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 22:  .Text = ""  '.Text = ""
      .ColWidth(22) = 0
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
    'Modify By Cheng 2003/07/14
    '取消預設位置
'   'Add By Cheng 2002/10/23
'   '預設目前在第一筆的位置
'   With Me.grd1
'      .Row = 1
'      .col = 0
'   End With
End Sub

Private Sub grd1_SelChange()
   grd1.Visible = False
   grd1.col = 0
   grd1.row = grd1.MouseRow
   If grd1.MouseRow <> 0 Then
      If grd1.Text = "V" Then
         grd1.Text = ""
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = QBColor(15)
         Next i
      Else
         grd1.Text = "V"
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
   grd1.Visible = True
End Sub

' 回到該畫面再重新查詢一次
Public Sub RefreshData()
   TextOk = False
   StrMenu1
   If TextOk = True Then
      StrMenu
      SetGrd1
   End If
End Sub



