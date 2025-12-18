VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060511_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專案件清單Excel-其他特定清單"
   ClientHeight    =   2580
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7932
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7932
   Begin VB.TextBox txtSales 
      Height          =   300
      Left            =   1032
      MaxLength       =   6
      TabIndex        =   13
      Top             =   192
      Width           =   708
   End
   Begin VB.Frame Frame1 
      Height          =   1836
      Left            =   24
      TabIndex        =   1
      Top             =   672
      Width           =   7788
      Begin VB.CommandButton cmdToPath 
         Height          =   300
         Index           =   0
         Left            =   6984
         Picture         =   "frm060511_2.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Top             =   960
         Width           =   350
      End
      Begin VB.CommandButton cmdToPath 
         Height          =   300
         Index           =   1
         Left            =   6456
         Picture         =   "frm060511_2.frx":0102
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   1368
         Width           =   350
      End
      Begin VB.CommandButton cmdProc1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "產生清單"
         Height          =   400
         Left            =   6288
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   168
         Width           =   1260
      End
      Begin VB.TextBox txtPath 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   2544
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1344
         Width           =   3900
      End
      Begin VB.TextBox txtPath 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3024
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   960
         Width           =   3900
      End
      Begin VB.Label Label3 
         Caption         =   "2. 將,全部取代為|     3.再另外存成CSV檔"
         ForeColor       =   &H000000FF&
         Height          =   228
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   672
         Width           =   3252
      End
      Begin VB.Label Label3 
         Caption         =   "CSV處理：1.將原本Excel前方加上SNO欄填入序號"
         ForeColor       =   &H000000FF&
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   432
         Width           =   3948
      End
      Begin VB.Label Label2 
         Caption         =   "設計案匯入CSV檔案路徑："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Index           =   1
         Left            =   48
         TabIndex        =   4
         Top             =   1368
         Width           =   2892
      End
      Begin VB.Label Label2 
         Caption         =   "發明+新型案匯入CSV檔案路徑："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Index           =   0
         Left            =   48
         TabIndex        =   3
         Top             =   960
         Width           =   2988
      End
      Begin VB.Label Label1 
         Caption         =   "X29307010 SKECHERS 案件清單 (月報告)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   228
         Left            =   120
         TabIndex        =   2
         Top             =   192
         Width           =   4356
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "回前畫面(&U)"
      Height          =   420
      Left            =   6576
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4248
      Top             =   24
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   1776
      TabIndex        =   14
      Top             =   192
      Width           =   1020
      Size            =   "1799;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "需求人員："
      Height          =   228
      Left            =   72
      TabIndex        =   12
      Top             =   216
      Width           =   948
   End
End
Attribute VB_Name = "frm060511_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2025/01/14
Option Explicit
Dim mPrevForm As Form
Dim m_FER39 As String '固定清單格式編號
Dim rsQuery As New ADODB.Recordset
Dim intQ As Integer, strQuery As String

Public Sub SetParent(ByVal fm As Form)
   Set mPrevForm = fm
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdProc1_Click()
Dim strText As String
Dim arrRow() As String
Dim arrCell() As String
Dim intA As Integer, intB As Integer, iRound As Integer
Dim strCon1 As String, strCon2 As String
Dim bol1stRow As Boolean
Dim strFileName(0 To 1) As String
   
   If Trim(txtSales) = "" Or lblSalesName.Caption = "" Then
      MsgBox "請輸入需求人員！", vbCritical
      Exit Sub
   End If
   
   For iRound = 0 To 1
      If Trim(txtPath(iRound)) = "" Or UCase(Right(txtPath(iRound), 4)) <> ".CSV" Then
         MsgBox "請選擇CSV檔案！", vbCritical
         txtPath(iRound).SetFocus
         txtPath_GotFocus iRound
         Exit Sub
      Else
         If Dir(txtPath(iRound)) = "" Then
            MsgBox "請選擇CSV檔案！", vbCritical
            txtPath(iRound).SetFocus
            txtPath_GotFocus iRound
            Exit Sub
         End If
      End If

      strFileName(iRound) = strExcelPath & strSrvDate(1) & "_X29307010" & IIf(iRound = 0, "(發明新型)", "(設計)") & MsgText(43)
      If PUB_ChkFileOpening(strFileName(iRound)) = True Then
          Exit Sub
      End If
      If Dir(strFileName(iRound)) <> "" Then
         Kill strFileName(iRound)
      End If
   Next iRound
   
   If UCase(Trim(txtPath(0))) = UCase(Trim(txtPath(1))) Then
      MsgBox "兩個檔名不可重複！", vbCritical
      Exit Sub
   End If
   
   m_FER39 = "1"
   
   cnnConnection.Execute "delete from R060511_2 where id='" & strUserNum & "' and formname = '" & Me.Name & "-" & m_FER39 & "' "
   '匯入資料
   For iRound = 0 To 1
      strText = PUB_ReadTextFile(txtPath(iRound), "UTF-8")
      arrRow = Split(strText, vbCrLf)
      bol1stRow = True
      strCon1 = "": strCon2 = ""
      For intA = LBound(arrRow) To UBound(arrRow)
         strCon1 = "": strCon2 = ""
         If arrRow(intA) <> "" Then
            '欄位名稱
            If bol1stRow = True Then
               bol1stRow = False
               If InStr(UCase(arrRow(intA)), "SNO") > 0 And ((iRound = 0 And InStr(UCase(arrRow(intA)), UCase("Patent Type")) > 0) Or (iRound = 1 And InStr(UCase(arrRow(intA)), UCase("Design Type")) > 0)) Then
               Else
                  MsgBox "【" & IIf(iRound = 0, "發明+新型案", "設計案") & "】檔案內容有誤，請確認！", vbCritical
                  Exit Sub
               End If
               
               arrCell = Split(arrRow(intA), ",")
               For intB = LBound(arrCell) To UBound(arrCell)
                  strExc(0) = Replace(arrCell(intB), "|", ",")
                  strCon1 = strCon1 & "|R" & Format(intB + 1, "000")
                  strCon2 = strCon2 & "|" & CNULL(ChgSQL(strExc(0)))
               Next intB
            '資料
            Else
               arrCell = Split(arrRow(intA), ",")
               For intB = LBound(arrCell) To UBound(arrCell)
                  strExc(0) = Replace(arrCell(intB), "|", ",") '名稱的|號還原為,號
                  strCon1 = strCon1 & "|R" & Format(intB + 1, "000")
                  strCon2 = strCon2 & "|" & CNULL(ChgSQL(Replace(strExc(0), """", "")))
               Next intB
            End If
         End If
         If strCon1 <> "" And strCon2 <> "" Then
            strSql = "Insert into R060511_2 (id, formname, SEQNO, ROWSEQ " & Replace(strCon1, "|", ",") & _
                     " ) Values ('" & strUserNum & "', '" & Me.Name & "-" & m_FER39 & "', " & iRound & ", " & intA + 1 & " " & Replace(strCon2, "|", ",") & ") "
            cnnConnection.Execute strSql
         End If
      Next intA
   Next iRound

JumpToSQL:
   '發明+新型
   '---全部案件
   strCon1 = "select 0 as rowseq,0 as sno,sqldatew(pa10) A,pa11 B,sqldatew(pa14) C,'' as C1,pa22 D ,'' as D1 " & _
             ",decode(pa57||pa108,null,decode(pa16,'1','Granted',decode(pa10,null,'Not Yet Filed','Pending')),'Inactive') E,'' as E1 " & _
             ",Z5 F ,'' as F1 ,'National Patent' G ,decode(pa08,'1','Patent','2','Utility Model') H,Y5 I " & _
             ",rtrim(pa06||' '||pa05) J,NA04 K ,decode(np09,null,'',decode(nYr,1,'1st',2,'2nd',3,'3rd',nYr||'th')||' annuity fee '||sqldatew(np09)) L,'' as L1 " & _
             ",decode(np09,null,'',decode(nYr,1,'1st',2,'2nd',3,'3rd',nYr||'th')||' annuity fee ')||sqldatew(np09) M,'' as M1 " & _
             ",sqldatew(pa12) N,'' as N1,pa13 O ,'' as O1,sqldatew(pa25) P ,'' as P1,sqldatew(pa10) Q,pa11 R " & _
             ",sqldatew(P5) S,rtrim(c1.cu05||' '||c1.cu88||' '||c1.cu89||' '||c1.cu90) T,GETINVENTOR(pa01,pa02,pa03,pa04) U " & _
             ",rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) V ,'N/A' W,'N/A' X,'N/A' Y,'N/A' Z ,GETPRIORITYX(pa01,pa02,pa03,pa04,'3') AA " & _
             ",GETPRIORITYX(pa01,pa02,pa03,pa04,'1') AB ,GETPRIORITYX(pa01,pa02,pa03,pa04,'2') AC ,'N/A' AD,'N/A' AE,'N/A' AF " & _
             ",D5 AG,'' as AG1 ,'' AH ,pa01||pa02||pa03||pa04 AI " & _
             "from patent,(select np02,np03,np04,np05,np07,np08,np09,decode(np07,'605',lastyear(pa72)+1) nYr,np15 " & _
             "From patent, nextprogress where pa26='X29307010' and pa57||pa108 is null " & _
             "and np02(+)=pa01 and np03(+)=pa02 and np04(+)=pa03 and np05(+)=pa04 and np06 is null and np07='605' " & _
             ") N,customer c1,fagent,nation "
   strCon1 = strCon1 & ",(select pa01 Z1,pa02 Z2,pa03 Z3,pa04 Z4,'Y' Z5 from patent,caseprogress " & _
             "where pa26='X29307010' and pa16='2' and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='107' ) Z " & _
             ",(select pa01 Y1,pa02 Y2,pa03 Y3,pa04 Y4,'Y' Y5 from patent,caseprogress " & _
             "where pa26='X29307010' and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='307') Y " & _
             ",(select pd01 P1,pd02 P2,pd03 P3,pd04 P4,min(pd05) P5 from patent,pridate " & _
             "where pa26='X29307010' and pd01(+)=pa01 and pd02(+)=pa02 and pd03(+)=pa03 and pd04(+)=pa04 and pd01 is not null " & _
             "group by pd01,pd02,pd03,pd04) P ,(select dc01 D1,dc02 D2,dc03 D3,dc04 D4,'Based on '||ptm05||' App No. '||p2.pa11 D5 " & _
             "from patent p1,divisioncase,patent p2,Patenttrademarkmap " & _
             "where p1.pa26='X29307010' and dc01(+)=p1.pa01 and dc02(+)=p1.pa02 and dc03(+)=p1.pa03 and dc04(+)=p1.pa04 " & _
             "and p2.pa01(+)=dc05 and p2.pa02(+)=dc06 and p2.pa03(+)=dc07 and p2.pa04(+)=dc08 and dc05 is not null " & _
             "and ptm01(+)='1' and ptm02(+)=p2.pa08) D  where pa26='X29307010' and pa08 in ('1','2') " & _
             "AND NP02(+)=PA01 AND NP03(+)=PA02 AND NP04(+)=PA03 AND NP05(+)=PA04 " & _
             "and c1.cu01(+)=substr(pa26,1,8) and c1.cu02(+)=substr(pa26,9) and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) " & _
             "and na01(+)=pa09 and Z1(+)=pa01 and Z2(+)=pa02 and Z3(+)=pa03 and Z4(+)=pa04 and Y1(+)=pa01 and Y2(+)=pa02 and Y3(+)=pa03 and Y4(+)=pa04 " & _
             "and P1(+)=pa01 and P2(+)=pa02 and P3(+)=pa03 and P4(+)=pa04 and d1(+)=pa01 and d2(+)=pa02 and d3(+)=pa03 and d4(+)=pa04 "
   strCon2 = "SELECT R060511_2.ROWSEQ,R001,R002,R003,R004,DECODE(R004,C,'',C) C1,R005,DECODE(R005,D,'',D) D1,R006,DECODE(R006,E,'',E) E1,R007,DECODE(R007,F,'',F) F1, " & _
             "R008,R009,R010,R011,R012, " & _
             "R013,DECODE(R013,L,'',L) L1,R014,DECODE(R014,M,'',M) M1,R015,DECODE(R015,N,'',N) N1,R016,DECODE(R016,O,'',O) O1,R017,DECODE(R017,P,'',P) P1, " & _
             "R018,R019,R020,R021,R022,R023,R024,R025,R026,R027,R028,R029,R030,R031,R032,R033,R034,'' AG1, " & _
             "R035 , R036, R037, R038, R039, R040, R041, R042, R043, R044, R045, R046, R047, R048, R049, R050, R051, R052, R053, R054, R055, R056, R057, R058 " & _
             "FROM (" & strCon1 & "), R060511_2 WHERE id='" & strUserNum & "' and formname = '" & Me.Name & "-" & m_FER39 & "' AND SEQNO='0' AND R036(+)=AI " & _
             "ORDER BY ROWSEQ "
   '代理人案件已排除:FCP071333000
   strCon1 = strCon1 & " and not exists(select * from r060511_2 where id='" & strUserNum & "' and formname = '" & Me.Name & "-" & m_FER39 & "' AND SEQNO='0' and R036=pa01||pa02||pa03||pa04) " & _
             " and pa01||pa02||pa03||pa04 not in ('FCP071333000') order by 2,3 "
   '抓抬頭資料
   strExc(1) = "SELECT ROWSEQ,R001,R002,R003,R004,'C1' as C1,R005,'D1' as D1,R006,'E1' as E1,R007,'F1' as F1,R008,R009,R010,R011,R012,R013," & _
             "'L1' as L1,R014,'M1' as M1,R015,'N1' as N1,R016,'O1' as O1,R017,'P1' as C1,R018,R019,R020," & _
             "R021,R022,R023,R024,R025,R026,R027,R028,R029,R030,R031,R032,R033,R034,'AG1' as AG1,R035,R036,R037,R038,R039,R040," & _
             "R041,R042,R043,R044,R045,R046,R047,R048,R049,R050,R051,R052,R053,R054,R055,R056,R057,R058 " & _
             "FROM R060511_2 where id='" & strUserNum & "' and formname = '" & Me.Name & "-" & m_FER39 & "' AND SEQNO='0' AND ROWSEQ='1'"
   If ProcExcelSave(strExc(1) & vbCrLf & strCon2 & vbCrLf & strCon1, strFileName(0)) = False Then
      GoTo EXITSUB
   End If

   '設計案
   strCon1 = "select 0 as rowseq,0 as sno, sqldatew(pa10) A,pa11 B,sqldatew(pa14) C,'' as c1,pa22 D,'' as d1 " & _
            ",decode(pa57||pa108,null,decode(pa16,'1','Registered',decode(pa10,null,'Not Yet Filed','Pending')),'Inactive') E,'' as e1 " & _
            ",'National Design' F,'Industrial Design' G,null H,rtrim(pa06||' '||pa05) I,NA04 J " & _
            ",decode(np09,null,'',decode(nYr,1,'1st',2,'2nd',3,'3rd',nYr||'th')||' annuity fee '||sqldatew(np09)) K,'' as k1 " & _
            ",null L, decode(np09,null,'',decode(nYr,1,'1st',2,'2nd',3,'3rd',nYr||'th')||' annuity fee ')||sqldatew(np09) M,'' as m1 " & _
            ",sqldatew(pa25) N,'' as n1,'N/A' O,'N/A' P,rtrim(c1.cu05||' '||c1.cu88||' '||c1.cu89||' '||c1.cu90) Q " & _
            ",rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) R " & _
            ",GETPRIORITYX(pa01,pa02,pa03,pa04,'3') S " & _
            ",GETPRIORITYX(pa01,pa02,pa03,pa04,'1') T " & _
            ",GETPRIORITYX(pa01,pa02,pa03,pa04,'2') U " & _
            ",GETDIVCASE(pa01,pa02,pa03,pa04) V,'' as V1 " & _
            ",pa01||pa02||pa03||pa04 W,'02-04' X,'' Y, '' Z "
   strCon1 = strCon1 & " from patent,(select np02,np03,np04,np05,np07,np08,np09,decode(np07,'605',lastyear(pa72)+1) nYr,np15 " & _
            "From patent, nextprogress where pa26='X29307010' and pa57||pa108 is null " & _
            "and np02(+)=pa01 and np03(+)=pa02 and np04(+)=pa03 and np05(+)=pa04 and np06 is null and np07='605' " & _
            ") N,customer c1,fagent,nation,(select dc01 D1,dc02 D2,dc03 D3,dc04 D4,'Based on '||ptm05||' App No. '||p2.pa11 D5 " & _
            "from patent p1,divisioncase,patent p2,Patenttrademarkmap " & _
            "where p1.pa26='X29307010' and dc01(+)=p1.pa01 and dc02(+)=p1.pa02 and dc03(+)=p1.pa03 and dc04(+)=p1.pa04 " & _
            "and p2.pa01(+)=dc05 and p2.pa02(+)=dc06 and p2.pa03(+)=dc07 and p2.pa04(+)=dc08 and dc05 is not null " & _
            "and ptm01(+)='1' and ptm02(+)=p2.pa08) D  where pa26='X29307010' and pa08='3' " & _
            "AND NP02(+)=PA01 AND NP03(+)=PA02 AND NP04(+)=PA03 AND NP05(+)=PA04 and c1.cu01(+)=substr(pa26,1,8) and c1.cu02(+)=substr(pa26,9) " & _
            "and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and na01(+)=pa09 and d1(+)=pa01 and d2(+)=pa02 and d3(+)=pa03 and d4(+)=pa04 "
   strCon2 = "SELECT R060511_2.ROWSEQ,R001,R002,R003,R004,DECODE(R004,C,'',C) C1, " & _
            "R005,DECODE(R005,D,'',D) D1,R006,DECODE(R006,E,'',E) E1,R007,R008,R009,R010,R011, " & _
            "R012,DECODE(R012,K,'',K) K1,R013,R014,DECODE(R014,M,'',M) M1,R015,DECODE(R015,N,'',N) N1, " & _
            "R016,R017,R018,R019,R020,R021,R022,R023,DECODE(R023,V,'',V) V1, " & _
            "R024 ,R025,R026,R027,R028,R029,R030,R031,R032,R033,R034,R035,R036,R037,R038,R039,R040,R041,R042,R043,R044,R045,R046 " & _
            "FROM (" & strCon1 & "),R060511_2 WHERE id='" & strUserNum & "' and formname = '" & Me.Name & "-" & m_FER39 & "' AND SEQNO='1' AND R024(+)=W " & _
            "ORDER BY ROWSEQ "
   strCon1 = strCon1 & " and not exists(select * from r060511_2 where id='" & strUserNum & "' and formname = '" & Me.Name & "-" & m_FER39 & "' AND SEQNO='1' and R024=pa01||pa02||pa03||pa04) " & _
             " order by 2,3 "
   '抓抬頭資料
   strExc(1) = "SELECT ROWSEQ,R001,R002,R003,R004,'C1' as C1,R005,'D1' as D1,R006,'E1' as E1,R007,R008,R009,R010,R011,R012," & _
             "'K1' as K1,R013,R014,'M1' as M1,R015,'N1' as N1,R016,R017,R018,R019,R020," & _
             "R021,R022,R023,'V1' as V1,R024,R025,R026,R027,R028,R029,R030,R031,R032,R033,R034,R035,R036,R037,R038,R039,R040," & _
             "R041,R042,R043,R044,R045,R046 " & _
             "FROM R060511_2 where id='" & strUserNum & "' and formname = '" & Me.Name & "-" & m_FER39 & "' AND SEQNO='1' AND ROWSEQ='1'"
   If ProcExcelSave(strExc(1) & vbCrLf & strCon2 & vbCrLf & strCon1, strFileName(1)) = False Then
      GoTo EXITSUB
   Else
      strExc(2) = frm060511.GetNowFER01
      If strExc(2) <> "" Then
         strSql = "Insert Into FCPEListRec(FER01,FER02,FER03,FER04,FER05,FER07,FER12,FER23,FER24,FER35,FER39) Values (" & _
                  " '" & strExc(2) & "', '" & strUserNum & "', to_char(sysdate,'yyyymmdd'), to_char(sysdate,'hh24miss')," & _
                  " '" & txtSales & "','X29307010','Y','1','1','" & Trim(Label1.Caption) & "','" & m_FER39 & "') "
         cnnConnection.Execute strSql
      End If
      MsgBox "Excel檔案產生完成！檔案位置：" & strExcelPathN
   End If
   
   
EXITSUB:
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdToPath_Click(Index As Integer)
   Dim fName As String

   cd1.Filter = "*.CSV"
   cd1.FilterIndex = 0
   
   fName = ""
   intI = InStrRev(txtPath(Index), "\")
   If intI > 0 Then
      fName = Left(txtPath(Index), intI - 1)
      If Dir(fName, vbDirectory) = "" Then
         fName = ""
      End If
   End If
   If fName <> "" Then
      cd1.InitDir = fName
   Else
      cd1.InitDir = PUB_Getdesktop
   End If

   cd1.ShowOpen
   If Trim(cd1.FileName) <> "" Then
      fName = cd1.FileName
      If UCase(Right(fName, 4)) = ".CSV" Then
         txtPath(Index) = fName
      Else
         MsgBox "請選擇CSV檔案！", vbCritical
         txtPath(Index) = ""
      End If
   End If
End Sub

Private Sub Form_Load()
Dim oObj As Control

   MoveFormToCenter Me
   
   For Each oObj In txtPath
      oObj.Text = ""
      oObj.Tag = ""
   Next
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsQuery = Nothing
   If TypeName(mPrevForm) <> "Nothing" Then
      mPrevForm.Show
   End If
   
   Set frm060511_2 = Nothing
End Sub

Private Sub txtPath_GotFocus(Index As Integer)
   TextInverse txtPath(Index)
End Sub

Private Function ProcExcelSave(ByVal pSQL As String, ByVal pFileName As String) As Boolean
Dim xlsReport
Dim wksReport
Dim iRound As Integer, nRows As Integer, MaxCols As Integer
Dim defMax As Integer
Dim tmpArr1 As Variant
Dim tmpArray As Variant
Dim bolOpenXls As Boolean
   
   ProcExcelSave = False
   tmpArr1 = Split(pSQL, vbCrLf)
   For iRound = LBound(tmpArr1) To UBound(tmpArr1)
      strQuery = Trim(tmpArr1(iRound))
      If Trim(strQuery) <> "" Then
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
         If intQ = 1 Then
            If bolOpenXls = False Then
               Set xlsReport = CreateObject("Excel.Application")
               xlsReport.SheetsInNewWorkbook = 1
               xlsReport.Workbooks.add
               Set wksReport = xlsReport.Worksheets(1)
               wksReport.Activate
               xlsReport.Visible = True
               nRows = 1
            Else
               If m_FER39 = "1" And iRound > 1 Then  '其他: 原本CSV檔以外新增的案件
                  nRows = nRows + 1
               End If
            End If
            
            rsQuery.MoveFirst
            MaxCols = rsQuery.Fields.Count
            ReDim tmpArray(1 To MaxCols)
            Do While Not rsQuery.EOF
               For intQ = 1 To MaxCols
                  tmpArray(intQ) = "" & rsQuery.Fields(intQ - 1)
               Next intQ
               wksReport.Range(Chr(65) & nRows & ":" & Pub_NumberToSystem26(MaxCols) & nRows).NumberFormatLocal = "@"
               wksReport.Range(Chr(65) & nRows & ":" & Pub_NumberToSystem26(MaxCols) & nRows).Value = tmpArray
               If bolOpenXls = False Then
                  defMax = MaxCols
                  '設定欄寬
                  For intQ = 1 To MaxCols + 1
                     'wksReport.Columns(Pub_NumberToSystem26(intQ) & ":" & Pub_NumberToSystem26(intQ)).EntireColumn.AutoFit
                     wksReport.Range(Pub_NumberToSystem26(intQ) & ":" & Pub_NumberToSystem26(intQ)).HorizontalAlignment = xlLeft
                    Next intQ
                  bolOpenXls = True
               End If
               nRows = nRows + 1
               rsQuery.MoveNext
            Loop
         End If
      End If
   Next iRound
   
   If bolOpenXls = True Then
      '設定欄寬
      For intQ = 1 To defMax + 1
         If m_FER39 = "1" Then
            wksReport.Range(Pub_NumberToSystem26(intQ) & ":" & Pub_NumberToSystem26(intQ)).Font.Size = 11
         End If
         wksReport.Range(Pub_NumberToSystem26(intQ) & ":" & Pub_NumberToSystem26(intQ)).EntireColumn.AutoFit
      Next intQ
      If m_FER39 = "1" Then
         wksReport.Range(Chr(65) & "1:" & Pub_NumberToSystem26(defMax) & "1").Interior.Color = QBColor(8) '..ColorIndex = 15     '底色:灰
         wksReport.Range(Chr(65) & "1:" & Pub_NumberToSystem26(defMax) & "1").Font.Color = QBColor(15)
         wksReport.Range("1:1").RowHeight = 20
         wksReport.Range("A2").Select
         xlsReport.ActiveWindow.FreezePanes = True '凍結窗格
      End If
      
      xlsReport.Sheets(1).Select '選擇工作表
      '判斷版本
      If Val(xlsReport.Version) < 12 Then
         xlsReport.Workbooks(1).SaveAs FileName:=pFileName, FileFormat:=-4143
      Else
         xlsReport.Workbooks(1).SaveAs FileName:=pFileName, FileFormat:=56
      End If
      xlsReport.Workbooks.Close
      xlsReport.Quit
      Set wksReport = Nothing
      Set xlsReport = Nothing
   End If
   
   ProcExcelSave = True
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox "產生清單失敗：" & pFileName & vbCrLf & Err.Description
   End If
End Function

Private Sub txtSales_GotFocus()
   TextInverse txtSales
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSales_Validate(Cancel As Boolean)
Dim strTempA As String

   strTempA = GetStaffName(txtSales.Text, True)
   If strTempA = "" Then
      MsgBox "請輸入正確的員工編號!"
      Cancel = True
      txtSales.SetFocus
      txtSales_GotFocus
      Exit Sub
   Else
      lblSalesName.Caption = strTempA
   End If

End Sub
