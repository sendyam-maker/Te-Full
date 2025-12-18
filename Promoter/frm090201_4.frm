VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護－已達本所期限案件資料"
   ClientHeight    =   5724
   ClientLeft      =   1656
   ClientTop       =   1512
   ClientWidth     =   9312
   ControlBox      =   0   'False
   FillColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5724
   ScaleWidth      =   9312
   Begin VB.TextBox txtWorkDays 
      Alignment       =   2  '置中對齊
      Height          =   285
      Left            =   3060
      TabIndex        =   8
      Top             =   578
      Width           =   375
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢"
      Height          =   315
      Index           =   1
      Left            =   4365
      TabIndex        =   5
      Top             =   563
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   4668
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5784
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "繼續(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   8010
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4590
      Left            =   90
      TabIndex        =   2
      Top             =   960
      Width           =   9135
      _ExtentX        =   16108
      _ExtentY        =   8086
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
      Height          =   255
      Left            =   1140
      TabIndex        =   9
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
      AutoSize        =   -1  'True
      Caption         =   "個工作天)"
      Height          =   180
      Index           =   1
      Left            =   3480
      TabIndex        =   7
      Top             =   630
      Width           =   780
   End
   Begin VB.Label Label7 
      Caption         =   "("
      Height          =   180
      Left            =   2970
      TabIndex        =   6
      Top             =   630
      Width           =   75
   End
   Begin VB.Label lblDate 
      Caption         =   "lblDate"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   630
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人： "
      Height          =   180
      Index           =   35
      Left            =   150
      TabIndex        =   3
      Top             =   360
      Width           =   1020
   End
End
Attribute VB_Name = "frm090201_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; grd1改字型=新細明體-ExtB、lblName
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Public TextOk As Boolean
Public m_bolIsDrawer As Boolean 'Added by Morgan 2016/4/18 是否為繪圖人員

'Public Print1Ok As Boolean  'CANCEL BY SONIA 2014/4/29
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String, SWPColor As String, SWPColor2 As String, SWPRow As String, SWPRow2 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 21) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String, Tmp001 As String, Tmp002 As String, Tmp003 As String, Tmp004 As String, k As Integer
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, SeekTmpBk As String, ChkData As Boolean
Dim strCP10 As String, AdoRs As ADODB.Recordset, StrNewSQL As String, Txt090201 As TextBox, ChkNoData As Boolean
Dim Fobj As FileSystemObject, ChkCp27 As Boolean, StrGrp090201 As String, ChkData2 As Boolean
Dim ADORECORDSET66 As New ADODB.Recordset
Dim Adorecordset99 As New ADODB.Recordset
Dim Intnick910123 As Integer
Dim m_dblWDBegin As Double '工作天起
Dim m_dblWDEnd As Double '工作天迄
'92.6.28 ADD BY SONIA
Dim m_ST03 As String
Dim mForm As Form 'Add By Sindy 2023/6/20
'Added by Lydia 2025/02/05
Dim bolShowCP142 As Boolean  '內專工程師增加顯示「指定日期」
Private Const mESeqNo As String = "1" '固定序號=1
Private Const mFrmName As String = "frm090201_4" '不使用Me.Name是避免在mdiMani正式Form.Show前，觸發Form_Load

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0 '繼續
   Unload Me
    
'Added by Morgan 2016/4/18
Case 1 '查詢
   Screen.MousePointer = vbHourglass
   StrMenu1 True
   Screen.MousePointer = vbDefault
'end 2016/4/18
Case Else
End Select
End Sub

Private Sub Form_Activate()
'2009/11/12 cancel by sonia
'    If TextOk = False Then
'        Unload Me
'    Else
'        Me.Show
'    End If
End Sub

Private Sub Form_Load()
   Debug.Print "Hide"
   Me.Hide
   Screen.MousePointer = vbHourglass
   MoveFormToCenter Me
'   Me.Enabled = False
'   StrMenu1
'   Me.Enabled = True

   'Added by Morgan 2016/4/18
   '專利處改列本所期限前3個工作天提醒
   If Left(Pub_StrUserSt03, 2) = "P1" Then
      Me.Caption = Replace(Me.Caption, "已達", "3個工作天內達")
      txtWorkDays = 3
   Else
      txtWorkDays = 1
   End If
   If m_bolIsDrawer Then
      Label1(35) = "繪圖人員："
   End If
   'end 2016/4/18
   
   'Added by Lydia 2025/02/05 內專工程師增加顯示「指定日期」
   If InStr("P0,P1", Left(Pub_StrUserSt03, 2)) > 0 Then
      bolShowCP142 = True
   Else
      bolShowCP142 = False
   End If
   'end 2025/02/05
   
   StrMenu
   'SetGrd1 'Mark by Lydia 2025/02/05
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '92.6.28 MODIFY BY SONIA
   'frm090201_5.Show
   '2009/11/12 modify by sonia 調整分所時間減少連線抓資料庫
   'm_ST03 = GetStaffDepartment(strUserNum)
   'If Left(m_ST03, 2) = "P2" Then

'2009/11/12 cancel by sonia 改寫至下面
'   If Left(Pub_StrUserSt03, 2) = "P2" Then
'   '2009/11/12 end
'       '2009/11/10 MODIFY BY SONIA 商標處人員直接進入工作維護
'       'frm090201_1.Show
'       frm090201_2.Show
'   Else
'       frm090201_5.Show
'       If frm090201_5.TextOk = False Then
'          'edit by nickc 2006/03/02
'          'frm090201_1.Show
'          frm090201_7.Show
'       End If
'   End If
'2009/11/12 end
   '92.6.28 END
   Set mForm = Nothing 'Add By Sindy 2023/6/20
   Set frm090201_4 = Nothing
   Nextstep   '2009/11/13 add by sonia
End Sub

'2009/11/13 add by sonia
Private Sub Nextstep()
'   If Left(Pub_StrUserSt03, 2) = "P2" Then   '商標處人員直接進入工作維護
'      frm090201_2.Show   '工作進度維護
'   Else
'      frm090201_5.TextOk = False
'      '2009/12/11 modify by sonia
'      'frm090201_5.StrMenu1  '新分案案件資料,無資料時由frm090201_5的nextstep執行下一畫面
'      frm090201_5.RefreshData
'      '2009/12/11 end
'      If frm090201_5.TextOk = True Then
'         frm090201_5.Show
'      End If
'   End If
   'Modify By Sindy 2012/5/10 非專利程式移至新程式frm090201_b單獨使用
   If Left(Pub_StrUserSt03, 2) = "P1" And _
      (InStr(UCase(App.EXEName), "PROMOTER") > 0 Or _
       (InStr(UCase(App.EXEName), "PATPRO") > 0 And InStr(UCase(App.EXEName), "PATPRO1") = 0)) Then '專利工作進度維護
      'Added by Morgan 2016/4/18
      If m_bolIsDrawer Then
         'Modified by Lydia 2025/02/05 "frm090201_4">>改常數mFrmName
         PUB_AddExcuteLog mFrmName
         'frm090711_2.Show
         'Modify By Sindy 2023/6/20
         Set mForm = Forms(0).GetForm("frm090711_2")
         mForm.Show
         '2023/6/20 END
      Else
      'end 2016/4/18
'         frm090201_5.TextOk = False
'         '2009/12/11 modify by sonia
'         'frm090201_5.StrMenu1  '新分案案件資料,無資料時由frm090201_5的nextstep執行下一畫面
'         frm090201_5.RefreshData
'         '2009/12/11 end
'         If frm090201_5.TextOk = True Then
'            frm090201_5.Show
'         End If
         'Modify By Sindy 2023/6/20
         Set mForm = Forms(0).GetForm("frm090201_5")
         mForm.TextOk = False
         mForm.RefreshData
         If mForm.TextOk = True Then
            mForm.Show
         End If
         '2023/6/20 END
      End If 'Added by Morgan 2016/4/18
      
   Else
      'Modify By Sindy 2021/9/11
      'If Val(strSrvDate(1)) >= Val(TMdebateStarDT) Then
      If Left(Pub_StrUserSt03, 2) = "P2" Or Left(Pub_StrUserSt03, 2) = "F1" Then '商標部
         'frm090201_b.Show '商標工作進度維護
         'Modify By Sindy 2023/6/20
         Set mForm = Forms(0).GetForm("frm090201_b")
         mForm.Show
         '2023/6/20 END
      'Add By Sindy 2023/6/20
      ElseIf Left(Pub_StrUserSt03, 2) = "F2" And _
         (InStr(UCase(App.EXEName), "PROMOTER") > 0 Or InStr(UCase(App.EXEName), "PATPRO1") > 0) Then '外專
         Set mForm = Forms(0).GetForm("frm090909")
         mForm.Show
         '2023/6/20 END
      ElseIf Left(Pub_StrUserSt03, 1) = "W" And InStr(UCase(App.EXEName), "LAW") > 0 Then '顧問服務組
      '2023/6/20 END
         'frm090201_d.Show '法務,顧問工作進度維護
         'Modify By Sindy 2023/6/20
         Set mForm = Forms(0).GetForm("frm090201_d")
         mForm.Show
         '2023/6/20 END
      Else
         MsgBox "此系統無您(" & strUserNum & ")的「工作進度資料維護」可使用！"
      End If
      '2021/9/11 END
   End If
   '2012/5/10 End
End Sub
'2009/11/13 end

'Mark by Lydia 2025/02/05 改寫法
'Public Sub StrMenu1(Optional pbolByUser As Boolean = False)
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'
'
'   DoEvents
'   'adoEng.Execute "DELETE FROM R090614 WHERE ID='" & strUserNum & "' "  '2009/11/12 cancel by sonia
'   'add by nickc 2007/12/17
'On Error Resume Next 'Add by Morgan 2009/12/18 若 table 不存在時跳過
'   adoEng.Execute "drop table R090614 "
'On Error GoTo 0 'Add by Morgan 2009/12/18 還原錯誤控制
'   adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text,R110006 text,R110007 text,R110008 text,R110009 text,R110010 text,R110011 text,R110012 text,R110013 text,R110014 text,R110015 text,R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo,R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text,R110026 double,R110027 double,R110028 double,R110029 text,R110030 text)"
'
'   If Not pbolByUser Or m_dblWDBegin = 0 Then 'Added by Morgan 2016/4/18
'
'      m_dblWDBegin = strSrvDate(1)
'      StrSQLa = "Select * From WorkDay Where WD01>" & m_dblWDBegin & " Order By WD01 asc"
'      rsA.CursorLocation = adUseClient
'      If Left(Pub_StrUserSt03, 2) = "P1" Then
'         rsA.MaxRecords = 3
'      Else
'         rsA.MaxRecords = 1
'      End If
'      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'         rsA.MoveLast 'Added by Morgan 2016/4/18
'         m_dblWDEnd = Val("" & rsA.Fields(0).Value)
'      Else
'          m_dblWDEnd = m_dblWDBegin
'      End If
'      If rsA.State <> adStateClosed Then rsA.Close
'      'Add by Morgan 2003/12/31
'      rsA.MaxRecords = 0
'      Set rsA = Nothing
'
'      '2011/11/3 modify by sonia 商標林副理說改3個月 T-171545
'      'm_dblWDBegin = CompDate(1, -2, m_dblWDBegin) 'Add by Morgan 2010/4/26 改抓2個月以內的都要
'      m_dblWDBegin = CompDate(1, -3, m_dblWDBegin)
'
'   End If 'Added by Morgan 2016/4/18
'
'   StrGrp090201 = ""
'   StrSQL6 = ""
'   strSQL1 = ""
'   strSQL2 = ""
'   'Modify By Cheng 2003/06/13
'   '抓未發文資料
'   'StrSQL6 = StrSQL6 + " and CP14='" & strUserNum & "' AND ((CP27 IS NULL  and CP57 IS NULL) OR ((CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp57 is null ) or (CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp27 is null) " & _
'   '                " Or (CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP05<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp57 is null And CP05>CP27 ))) and cp05>=19980101 And (CP06>=" & m_dblWDBegin & " And CP06<" & m_dblWDEnd & " ) "
'   '2006/9/15 MODIFY BY SONIA 加入未閉卷條件 CFP-13993,閉卷但實審未取消收文
'   'Modified by Morgan 2016/4/18 +繪圖人員判斷
'   'StrSQL6 = StrSQL6 + " and CP14='" & strUserNum & "' AND (CP27 IS NULL  and CP57 IS NULL) and cp05>=19980101 And (CP06>=" & m_dblWDBegin & " And CP06<" & m_dblWDEnd & " ) "
''Modify By Sindy 2024/7/30 CP14 改判斷 EP05
'   'StrSQL6 = StrSQL6 + " and " & IIf(m_bolIsDrawer, "CP29", "CP14") & "='" & strUserNum & "' and cp57 is null and cp27 is null and cp05>=19980101 And (CP06>=" & m_dblWDBegin & " And CP06<" & m_dblWDEnd & " ) "
'   StrSQL6 = StrSQL6 + " and " & IIf(m_bolIsDrawer, "CP29", "EP05") & "='" & strUserNum & "' and cp57 is null and cp27 is null and cp05>=19980101 And (CP06>=" & m_dblWDBegin & " And CP06<" & m_dblWDEnd & " ) "
''2024/7/30 END
'   'end 2016/4/18
'
'   CheckOC
'   'strSQL = "SELECT CP14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
'   '                " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & StrSQL6 & _
'   '                " AND PA57 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ") "
'   'strSQL = strSQL + " UNION all  SELECT CP14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
'   '                " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & _
'   '                " AND TM29 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") "
'   'strSQL = strSQL + " UNION all  SELECT CP14,ep01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
'   '                " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & " AND LC08 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
'   'strSQL = strSQL + " UNION all  SELECT CP14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
'   '                " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & " AND HC09 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
'   'strSQL = strSQL + " UNION all  SELECT CP14,EP01,SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
'   '                " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & " AND SP15 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") "
'   'Modify by Morgan 2010/8/16 百年蟲 " & SQLDate("CP05") & "-->substrb(' '||sqldatet(cp05),-9)
'   'Modified by Morgan 2016/4/18
'   'strSql = "SELECT CP14,EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
'   '                " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & StrSQL6 & _
'   '                " AND PA57 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ") "
'   'strSql = strSql + " UNION all  SELECT CP14,EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
'   '                " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & _
'   '                " AND TM29 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") "
'   'strSql = strSql + " UNION all  SELECT CP14,ep01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
'   '                " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & " AND LC08 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
'   'strSql = strSql + " UNION all  SELECT CP14,EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
'   '                " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & " AND HC09 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
'   'strSql = strSql + " UNION all  SELECT CP14,EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
'   '                " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & " AND SP15 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") "
'
'   'Modify By Sindy 2016/9/7
''   strSql = "SELECT " & IIf(m_bolIsDrawer, "CP29", "CP14") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
''                   " WHERE CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & StrSQL6 & _
''                   " AND PA57 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ") "
''   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "CP14") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
''                   " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & _
''                   " AND TM29 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") "
''   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "CP14") & ",ep01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
''                   " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & " AND LC08 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
''   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "CP14") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
''                   " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & " AND HC09 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
''   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "CP14") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
''                   " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & " AND SP15 IS NULL AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") "
''Modify By Sindy 2024/7/30 CP14 改判斷 EP05
'   strSql = "SELECT " & IIf(m_bolIsDrawer, "CP29", "EP05") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
'                   " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ")" & StrSQL6 & " AND CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & _
'                   " AND PA57 IS NULL"
'   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "EP05") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
'                   " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ")" & StrSQL6 & " AND CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & _
'                   " AND TM29 IS NULL"
'   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "EP05") & ",ep01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
'                   " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ")" & StrSQL6 & " AND CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) AND LC08 IS NULL"
'   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "EP05") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
'                   " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ")" & StrSQL6 & " AND CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) AND HC09 IS NULL"
'   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "EP05") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
'                   " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") " & StrSQL6 & " AND CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) AND SP15 IS NULL"
'   '2016/9/7 END
'   'end 2016/4/18
'   strSql = strSql + " ORDER BY 1,4 "
'   CheckOC
'   'Print1Ok = False
'   TextOk = True
'   With adoRecordset
'       .CursorLocation = adUseClient
'      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'       If .RecordCount <> 0 And .RecordCount > 0 Then
'           .MoveFirst
'           k = 0
'           DoEvents
'           '判斷等級是否屬於專利
'           'CANCEL BY SONIA 2014/4/29
'           'If (Val(CheckStr(.Fields(22))) >= 31 And Val(CheckStr(.Fields(22))) <= 39) Or (Val(CheckStr(.Fields(22))) >= 71 And Val(CheckStr(.Fields(22))) <= 89) Or CheckStr(.Fields("cp14")) = "94007" Then
'           '    Print1Ok = True
'           'End If
'           '2014/4/29 END
'           Do While .EOF = False
'               For i = 0 To 21
'                  'Add by Morgan 2010/8/16
'                  If i = 3 Then
'                     strTemp(3) = "" & .Fields(3)
'                  Else
'                  'end 2010/8/16
'                     strTemp(i) = CheckStr(.Fields(i))
'                  End If
'               Next i
'               'edit by nickc 2007/12/17
'               'strSQL = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "','" & .Fields(28).Value & "' ) "
'               strSql = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "','" & .Fields(28).Value & "',0,0,0,0,0) "
'               adoEng.Execute strSql
'               .MoveNext
'               DoEvents
'           Loop
'       'Modified by Morgan 2016/4/18
'       'Else
'           Debug.Print "Old:True"
'          If pbolByUser Then
'               StrMenu
'               SetGrd1
'          End If
'       ElseIf Not pbolByUser Then
'       'end 2016/4/18
'           Debug.Print "Old:False"
'           TextOk = False
'           Nextstep   '2009/11/13 add by sonia
'       End If
'       CheckOC
'   End With
'End Sub

'Mark by Lydia 2025/02/05 改寫法
'Private Sub StrMenu()
'
'   Me.lblName.Caption = strUserName
'   Me.lblDate.Caption = ChangeTStringToTDateString(m_dblWDBegin - 19110000) & " <= 本所期限 < " & ChangeTStringToTDateString(m_dblWDEnd - 19110000)
'
'   'strSQL = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024 FROM R090614 " & _
'   '              " WHERE ID='" & strUserNum & "' AND R110001='" & strUserNum & "' ORDER BY R110002 desc,R110003,R110004 "
'   strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024 FROM R090614 " & _
'                 " WHERE ID='" & strUserNum & "' AND R110001='" & strUserNum & "'"
'   'Modified by Morgan 2016/4/18 +繪圖
'   If m_bolIsDrawer Then
'      strSql = strSql & " ORDER BY R110012 ASC, R110005 Desc "
'   Else
'      strSql = strSql & " ORDER BY R110002 Desc, R110005 Desc "
'   End If
'   'end 2016/4/18
'   CheckOC
'   With adoRecordset
'       .CursorLocation = adUseClient
'       .Open strSql, adoEng, adOpenStatic, adLockReadOnly
'       If .RecordCount <> 0 And .RecordCount > 0 Then
'           Set GRD1.Recordset = adoRecordset
'           ChkNoData = False
'       Else
'           ChkNoData = True
'           GRD1.Clear
'           GRD1.Rows = 2
'           Screen.MousePointer = vbDefault
'           Exit Sub
'       End If
'   End With
'   CheckOC
'End Sub

'Mark by Lydia 2025/02/05 改寫法
'Private Sub SetGrd1()
'   With GRD1
'       'Modify by Morgan 2010/8/16 日期欄寬改800
'       .Visible = False
'       .Cols = 23
'       .row = 0
'       .col = 0:   .Text = "目次"
'       'Added by Morgan 2016/4/18
'       If m_bolIsDrawer Then
'         .ColWidth(0) = 0
'       Else
'       'end 2016/4/18
'         .ColWidth(0) = 300
'       End If 'Added by Morgan 2016/4/18
'
'       .CellAlignment = flexAlignCenterCenter
'       .col = 1:   .Text = "收文類別"
'       .ColWidth(1) = 200
'       .CellAlignment = flexAlignCenterCenter
'       .col = 2:   .Text = "收文日"
'       .ColWidth(2) = 800
'       .CellAlignment = flexAlignCenterCenter
'       .col = 3:   .Text = "本所案號"
'       .ColWidth(3) = 1000
'       .CellAlignment = flexAlignCenterCenter
'       .col = 4:   .Text = "案件名稱"
'       .ColWidth(4) = 1500
'       .CellAlignment = flexAlignCenterCenter
'       .col = 5:   .Text = "國家"
'       .ColWidth(5) = 450
'       .CellAlignment = flexAlignCenterCenter
'       .col = 6:   .Text = "種類"
'       .ColWidth(6) = 400
'       .CellAlignment = flexAlignCenterCenter
'       .col = 7:   .Text = "案件性質"
'       .ColWidth(7) = 800
'       .CellAlignment = flexAlignCenterCenter
'       .col = 8:   .Text = "Y/N"
'       .ColWidth(8) = 300
'       .CellAlignment = flexAlignCenterCenter
'
'       .col = 9:   .Text = "承辦期限"
'       .ColWidth(9) = 800
'       .CellAlignment = flexAlignCenterCenter
'       .col = 10:   .Text = "本所期限"
'       .ColWidth(10) = 800
'       .CellAlignment = flexAlignCenterCenter
'       .col = 11:  .Text = "法定期限"
'       .ColWidth(11) = 0
'       .CellAlignment = flexAlignCenterCenter
'       .col = 12:  .Text = "齊備日"
'       .ColWidth(12) = 800
'       .CellAlignment = flexAlignCenterCenter
'       .col = 13:  .Text = "完稿日"
'       .ColWidth(13) = 800
'       .CellAlignment = flexAlignCenterCenter
'       .col = 14:  .Text = "會稿日"
'       .ColWidth(14) = 800
'       .CellAlignment = flexAlignCenterCenter
'       .col = 15:  .Text = "核稿人"
'       .ColWidth(15) = 700
'       .CellAlignment = flexAlignCenterCenter
'       .col = 16:  .Text = "會稿完成日"
'       .ColWidth(16) = 800
'       .CellAlignment = flexAlignCenterCenter
'       .col = 17:  .Text = "發文日"
'       .ColWidth(17) = 800
'       .CellAlignment = flexAlignCenterCenter
'       .col = 18:  .Text = "承辦天數"
'       .ColWidth(18) = 800
'       .CellAlignment = flexAlignCenterCenter
'       .col = 19:  .Text = "備註"
'       .ColWidth(19) = 2000
'       .CellAlignment = flexAlignCenterCenter
'       .col = 20:  .Text = "智權人員"
'       .ColWidth(20) = 800
'       .CellAlignment = flexAlignCenterCenter
'       .col = 21:  .Text = ""
'       .ColWidth(21) = 0
'       .CellAlignment = flexAlignCenterCenter
'       .col = 22:  .Text = ""
'       .ColWidth(22) = 0
'       .CellAlignment = flexAlignCenterCenter
'       .Visible = True
'   End With
'   'Modify By Cheng 2003/07/14
'   '取消預設位置
'   ''Add By Cheng 2002/10/23
'   ''預設目前在第一筆的位置
'   'With Me.grd1
'   '   .Row = 1
'   '   .col = 0
'   'End With
'End Sub
'end -----'Mark by Lydia 2025/02/05 改寫法

Private Sub txtWorkDays_GotFocus()
   TextInverse txtWorkDays
   CloseIme
End Sub

Private Sub txtWorkDays_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(KeyAscii) Then
      KeyAscii = 0
   Else
   
   End If
End Sub

Private Sub txtWorkDays_Validate(Cancel As Boolean)
   If Val(txtWorkDays) > 0 Then
      m_dblWDEnd = CompWorkDay(Val(txtWorkDays) + 1, strSrvDate(1))
      Me.lblDate.Caption = ChangeTStringToTDateString(m_dblWDBegin - 19110000) & " <= 本所期限 < " & ChangeTStringToTDateString(m_dblWDEnd - 19110000)
   End If
End Sub

'Added by Lydia 2025/02/05 直接用RdataFactory
Public Sub StrMenu1(Optional pbolByUser As Boolean = False)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   If Not pbolByUser Or m_dblWDBegin = 0 Then
      m_dblWDBegin = strSrvDate(1)
      If InStr("P0,P1", Left(Pub_StrUserSt03, 2)) > 0 Then
          m_dblWDEnd = CompWorkDay(4, "" & m_dblWDBegin)
      Else
          m_dblWDEnd = CompWorkDay(2, "" & m_dblWDBegin) '加1工作天
      End If
      '2011/11/3 modify by sonia 商標林副理說改3個月 T-171545
      'm_dblWDBegin = CompDate(1, -2, m_dblWDBegin) 'Add by Morgan 2010/4/26 改抓2個月以內的都要
      m_dblWDBegin = CompDate(1, -3, m_dblWDBegin)
      
   End If
   
   StrGrp090201 = ""
   StrSQL6 = ""
   strSQL1 = ""
   strSQL2 = ""
   StrSQL6 = StrSQL6 + " and " & IIf(m_bolIsDrawer, "CP29", "EP05") & "='" & strUserNum & "' and cp57 is null and cp27 is null and cp05>=19980101 And (CP06>=" & m_dblWDBegin & " And CP06<" & m_dblWDEnd & " ) "

   '內專工程師增加顯示「指定日期
   StrSQLa = "SELECT " & IIf(m_bolIsDrawer, "CP29", "EP05") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & _
                   " " & SQLDate("CP06") & ", " & SQLDate("CP142") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0)," & _
                   " substr(EP12,1,100),nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
                   " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ")" & StrSQL6 & " AND CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+) " & _
                   " AND PA57 IS NULL"
   StrSQLa = StrSQLa + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "EP05") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & _
                   " " & SQLDate("CP06") & ", " & SQLDate("CP142") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0)," & _
                   " substr(EP12,1,100),nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
                   " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ")" & StrSQL6 & " AND CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & _
                   " AND TM29 IS NULL"
   StrSQLa = StrSQLa + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "EP05") & ",ep01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(LC15,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & _
                   " " & SQLDate("CP06") & ", " & SQLDate("CP142") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0)," & _
                   " substr(EP12,1,100),nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
                   " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ")" & StrSQL6 & " AND CP09=EP02(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) AND LC08 IS NULL"
   StrSQLa = StrSQLa + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "EP05") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10)," & SQLDate("CP48") & "," & _
                   " " & SQLDate("CP06") & ", " & SQLDate("CP142") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & _
                   " " & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0)," & _
                   " substr(EP12,1,100),nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
                   " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ")" & StrSQL6 & " AND CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) AND HC09 IS NULL"
   StrSQLa = StrSQLa + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "EP05") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & ", " & _
                   " " & SQLDate("CP06") & ", " & SQLDate("CP142") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0)," & _
                   " substr(EP12,1,100),nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
                   " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") " & StrSQL6 & " AND CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) AND SP15 IS NULL"
   StrSQLa = StrSQLa + " ORDER BY 1,4 "
      
   cnnConnection.Execute "delete from rdatafactory where id ='" & strUserNum & "' and formname='" & mFrmName & "' "
   intI = 1
   TextOk = True
   Set rsA = ClsLawReadRstMsg(intI, StrSQLa)
   If intI = 1 Then
      Set RsTemp = PUB_CreateRecordset(rsA, , , , mFrmName)
      If pbolByUser Then
          StrMenu
      End If

   ElseIf Not pbolByUser Then
       TextOk = False
       Nextstep
   End If
   Set rsA = Nothing
   
End Sub

'Added by Lydia 2025/02/05
Private Sub StrMenu()
Dim rsAD As New ADODB.Recordset

   Me.lblName.Caption = strUserName
   Me.lblDate.Caption = ChangeTStringToTDateString(m_dblWDBegin - 19110000) & " <= 本所期限 < " & ChangeTStringToTDateString(m_dblWDEnd - 19110000)

   strSql = "SELECT R002,R003,R004,R005,R006,R028,R008,R009,R007,R010,R011,R012,R013,R014,R015,R016,R017,R018,R019,R020,R021,R022,R023,R024,R025 " & _
           " FROM RDATAFACTORY WHERE FORMNAME = '" & mFrmName & "' and ID='" & strUserNum & "' AND SEQNO='" & mESeqNo & "'"

   If m_bolIsDrawer Then
      strSql = strSql & " ORDER BY R013 ASC, R005 Desc "
   Else
      strSql = strSql & " ORDER BY R002 Desc, R005 Desc "
   End If
   Call SetGrd(True)
   intI = 1
   Set rsAD = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      Set GRD1.Recordset = rsAD
      Call SetGrd
      ChkNoData = False
   Else
      ChkNoData = True
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   Set rsAD = Nothing
   
End Sub

'Added by Lydia 2025/02/05
Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("目次", "收文類別", "收文日", "本所案號", "案件名稱", "國家", "種類", "案件性質", "Y/N", "承辦期限", "本所期限", "指定日期", "法定期限", "齊備日", "完稿日", "會稿日", "核稿人", "會稿完成日", "發文日", "承辦天數", "備註", "智權人員")
   arrGridHeadWidth = Array(IIf(m_bolIsDrawer = True, 0, 300), 200, 800, 1000, 1500, 450, 400, 800, 300, 800, 800, IIf(bolShowCP142 = True, 800, 0), 800, 800, 800, 800, 700, 800, 800, 800, 2000, 800)
   
   Me.GRD1.Visible = False
   Me.GRD1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
        Me.GRD1.Clear
        Me.GRD1.Rows = 2
   End If
   For iRow = 0 To Me.GRD1.Cols - 1
      Me.GRD1.row = 0
      Me.GRD1.col = iRow
      If iRow <= UBound(arrGridHeadWidth) Then
         Me.GRD1.Text = arrGridHeadText(iRow)
         Me.GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
         Me.GRD1.CellAlignment = flexAlignCenterCenter
      Else
         Me.GRD1.Text = ""
         Me.GRD1.ColWidth(iRow) = 0
      End If
   Next
   
   Me.GRD1.Visible = True
End Sub
