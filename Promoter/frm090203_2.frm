VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090203_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料查詢"
   ClientHeight    =   5760
   ClientLeft      =   -3690
   ClientTop       =   1380
   ClientWidth     =   9320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9320
   Begin VB.CommandButton Command2 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   8060
      TabIndex        =   1
      Top             =   30
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "詳細資料(&L)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6830
      TabIndex        =   0
      Top             =   30
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   5230
      Left            =   40
      TabIndex        =   2
      Top             =   460
      Width           =   9220
      _ExtentX        =   16263
      _ExtentY        =   9225
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
End
Attribute VB_Name = "frm090203_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

'Dim pemain As New ADODB.Recordset
'Public A0 As String, A1 As String, A2 As String, A3 As String, A4 As String, A5 As String, A6 As String, A7 As String, A8 As String, A9 As String, A10 As String, NUMBER As String
Dim i As Integer, strTemp(0 To 21) As String, StrSQL6 As String, k As Integer, k1 As Integer
'Dim Print1Ok As Boolean 'CANCEL BY SONIA 2014/4/29
'add by nickc 2007/08/22 紀錄作用按鍵
Public cmdState As Integer


Private Sub Command1_Click()
'edit by nickc 2007/08/22 秀玲說改跟共同查詢相同
'grd1.col = 20
'frm090203_2.Hide
   cmdState = 100
   PubShowNextData
End Sub

'add by nickc 2007/08/22 秀玲說改跟共同查詢相同
Public Sub PubShowNextData()
Dim strCP01 As String 'Add By Sindy 2012/5/21
   If cmdState = 100 Then
      grd1.row = k1
      grd1.col = 20
      If fnSaveParentForm(Me) = False Then
          Me.Enabled = True
          Exit Sub
      End If
      Screen.MousePointer = vbHourglass
      'Modify By Sindy 2012/5/21 +if,frm100101_K
      strCP01 = GetCaseProData(Trim(Pub_RplStr(grd1.Text)), "CP01")
      If strCP01 = "P" Or strCP01 = "PS" Or strCP01 = "FG" Or _
         strCP01 = "FCP" Or strCP01 = "CFP" Or strCP01 = "CPS" Or _
         Val(strSrvDate(1)) < Val(TMdebateStarDT) Then  '專利處工作進度
         frm100101_F.Show
         frm100101_F.Process Pub_RplStr(grd1.Text)
      Else
         frm100101_K.Show
         frm100101_K.Process Pub_RplStr(grd1.Text)
      End If
      '2012/5/21 End
      Screen.MousePointer = vbDefault
      Me.Enabled = True
      cmdState = 200
   End If
End Sub

Private Sub Command2_Click()
   frm090203_1.Show
   Unload Me
End Sub

Private Sub SetGrd1()
   With grd1
      .Visible = False
      .Cols = 21
      .row = 0
      .col = 0:   .Text = "目次"
      .ColWidth(0) = 600
      .CellAlignment = flexAlignCenterCenter
      .col = 1:   .Text = "收文類別"
      .ColWidth(1) = 300
      .CellAlignment = flexAlignCenterCenter
      .col = 2:   .Text = "收文日"
      .ColWidth(2) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 3:   .Text = "本所案號"
      .ColWidth(3) = 1550
      .CellAlignment = flexAlignCenterCenter
      .col = 4:   .Text = "案件名稱"
      .ColWidth(4) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 5:   .Text = "Y/N"
      .ColWidth(5) = 400
      .CellAlignment = flexAlignCenterCenter
      .col = 6:   .Text = "種類"
      .ColWidth(6) = 500
      .CellAlignment = flexAlignCenterCenter
      .col = 7:   .Text = "案件性質"
      .ColWidth(7) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 8:   .Text = "承辦期限"
      .ColWidth(8) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 9:   .Text = "本所期限"
      .ColWidth(9) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 10:  .Text = "法定期限"
      .ColWidth(10) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 11:  .Text = "齊備日"
      .ColWidth(11) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 12:  .Text = "完稿日"
      .ColWidth(12) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 13:  .Text = "會稿日"
      .ColWidth(13) = IIf(Left(PUB_GetST03(strUserNum), 2) = "F2", 0, 800) 'Modify By Sindy 2023/12/18
      .CellAlignment = flexAlignCenterCenter
      .col = 14:  .Text = "核稿人"
      .ColWidth(14) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 15:  .Text = "會稿完成日"
      .ColWidth(15) = IIf(Left(PUB_GetST03(strUserNum), 2) = "F2", 0, 800) 'Modify By Sindy 2023/12/18
      .CellAlignment = flexAlignCenterCenter
      .col = 16:  .Text = "發文日"
      .ColWidth(16) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 17:  .Text = "承辦天數"
      .ColWidth(17) = IIf(Left(PUB_GetST03(strUserNum), 1) = "F", 0, 800) 'Modify By Sindy 2023/12/18
      .CellAlignment = flexAlignCenterCenter
      .col = 18:  .Text = "備註"
      .ColWidth(18) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 19:  .Text = "智權人員"
      .ColWidth(19) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 20:  .Text = ""
      .ColWidth(20) = 0
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Modify By Cheng 2003/05/08
   'cnnConnection.Execute "DELETE FROM R090614 WHERE ID='" & strUserNum & "' "
   adoEng.Execute "DELETE FROM R090614 WHERE ID='" & strUserNum & "' "
       'add by nickc 2007/12/17
       adoEng.Execute "drop table R090614 "
       adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text,R110006 text,R110007 text,R110008 text,R110009 text,R110010 text,R110011 text,R110012 text,R110013 text,R110014 text,R110015 text,R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo,R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text)"
       
   pub_QL05 = pub_QL05 & ";" & frm090203_1.Label1 & frm090203_1.Text1 'Add By Sindy 2010/12/14
   
   StrSQL6 = ""
   'Modify By Sindy 2023/12/18
   If Left(PUB_GetST03(strUserNum), 2) = "F2" Then
      StrSQL6 = StrSQL6 + " AND (ep05='" & strUserNum & "' or (cp10='201' and ep04='" & strUserNum & "')) "
   Else
   '2023/12/18 END
      StrSQL6 = StrSQL6 + " AND ep05='" & strUserNum & "' "
   End If
   'Modify By Cheng 2003/07/21
   'StrSQL6 = StrSQL6 + " and CP27>=" & Val(frm090203_1.ClickTime) + 191100 & "01  and CP27<=" & Val(frm090203_1.ClickTime) + 191100 & "31 "
   StrSQL6 = StrSQL6 + " AND ((CP158=0 and CP159=0)" & _
                            " OR ((CP27>=" & Val(frm090203_1.ClickTime) + 191100 & "01 AND CP27<=" & Val(frm090203_1.ClickTime) + 191100 & "31 and CP159=0)" & _
                            " OR (CP57>=" & Val(frm090203_1.ClickTime) + 191100 & "01 AND CP57<=" & Val(frm090203_1.ClickTime) + 191100 & "31 and CP158=0) " & _
                            " Or (CP05>=" & Val(frm090203_1.ClickTime) + 191100 & "01 AND CP05<=" & Val(frm090203_1.ClickTime) + 191100 & "31 and CP159=0 And CP05>CP27)))" & _
                       " AND cp05>=19980101"
   CheckOC
   'Modify By Cheng 2002/04/26
   '若已閉卷, 則在本所案號後加"*"號
                        strSql = "SELECT S1.ST02,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",S3.ST02," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,S2.ST02,CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) and pA08=PTM02(+) " & StrSQL6 & " and cp01 in (" & SQLGrpStr("", 1) & ") "
   strSql = strSql + " UNION all  SELECT S1.ST02,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",S3.ST02," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,S2.ST02,CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) " & StrSQL6 & " and cp01 in (" & SQLGrpStr("", 2) & ") "
   strSql = strSql + " UNION all  SELECT S1.ST02,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',decode(lc15,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",S3.ST02," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,S2.ST02,CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND cp01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " and cp01 in (" & SQLGrpStr("", 3) & ") "
   strSql = strSql + " UNION all  SELECT S1.ST02,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',CPM03," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",S3.ST02," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,S2.ST02,CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " and cp01 in (" & SQLGrpStr("", 4) & ") "
   strSql = strSql + " UNION all  SELECT S1.ST02,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',decode(sp09,'000',cpm03,cpm04)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",S3.ST02," & SQLDate("EP08") & "," & SQLDate("CP27") & ",0,EP12,S2.ST02,CP09,S1.ST05 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP WHERE EP02=CP09(+) AND cP01=sP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) " & StrSQL6 & " and cp01 in (" & SQLGrpStr("", 5) & ") "
   strSql = strSql + " ORDER BY 1 "
   CheckOC
   'Print1Ok = False   'CANCEL BY SONIA 2014/4/29
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           .MoveFirst
           DoEvents
           'CANCEL BY SONIA 2014/4/29
           ''判斷等級是否屬於專利
           'If (Val(CheckStr(.Fields(22))) >= 31 And Val(CheckStr(.Fields(22))) <= 39) Or (Val(CheckStr(.Fields(22))) >= 71 And Val(CheckStr(.Fields(22))) <= 89) Then
           '    Print1Ok = True
           'End If
           '2014/4/29 END
           Do While .EOF = False
               For i = 0 To 21
                   strTemp(i) = CheckStr(.Fields(i))
               Next i
               '計算承辦天數
               'Modify By Sindy 2023/12/18
               If Left(PUB_GetST03(strUserNum), 1) <> "F" Then
               '2023/12/18 END
                  If Len(strTemp(14)) <> 0 And Len(strTemp(12)) <> 0 And Val(strTemp(14)) <> 0 And Val(strTemp(12)) <> 0 Then
                     strTemp(18) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(14))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(12))))))
                  Else
                     If Len(strTemp(13)) <> 0 And Len(strTemp(12)) <> 0 And Val(strTemp(13)) <> 0 And Val(strTemp(12)) <> 0 Then
                        strTemp(18) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(13))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(12))))))
                     End If
                  End If
               End If
               
               'Modify By Cheng 2002/04/18
   '            strSQL = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "') "
               'Modify By Cheng 2003/02/10
   '            strSQL = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','') "
               strSql = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','','','') "
   '            cnnConnection.Execute strSQL
               adoEng.Execute strSql
               .MoveNext
               DoEvents
           Loop
       End If
       CheckOC
   End With
   strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110007,R110008,R110009,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022 FROM R090614 WHERE ID='" & strUserNum & "' ORDER BY 1 "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   'Modify By Cheng 2003/05/08
   'adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
   adoRecordset.Open strSql, adoEng, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
       InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/14
       Set grd1.Recordset = adoRecordset
       grd1.row = 1
       k1 = 1
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/12/14
       grd1.Clear
       grd1.Rows = 2
       k1 = 0
   End If
   SetGrd1
   SetColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090203_2 = Nothing
End Sub

Private Sub Grd1_Click()
   If grd1.MouseRow <> 0 Then
      grd1.row = grd1.MouseRow
      k1 = grd1.row
   End If
   SetColor
End Sub

Private Sub SetColor()
   If k1 <> 0 Then
      grd1.row = k1
      grd1.col = 0
      grd1.CellBackColor = &HFFC0C0
      If k <> 0 And k <> k1 Then
      grd1.row = k
      grd1.col = 0
      grd1.CellBackColor = QBColor(15)
      End If
      k = k1
   End If
End Sub
