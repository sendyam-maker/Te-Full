VERSION 5.00
Begin VB.Form frm010009 
   BorderStyle     =   1  '單線固定
   Caption         =   "主管機關來函收文簿"
   ClientHeight    =   2064
   ClientLeft      =   2256
   ClientTop       =   1152
   ClientWidth     =   4944
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2064
   ScaleWidth      =   4944
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3096
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3924
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox txtKeyIn 
      Height          =   288
      Index           =   1
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "1"
      Top             =   1140
      Width           =   492
   End
   Begin VB.TextBox txtKeyIn 
      Height          =   288
      Index           =   0
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   0
      Top             =   750
      Width           =   1092
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PS：二份報表都不含專業部輸入的資料"
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   3060
   End
   Begin VB.Label Label1 
      Caption         =   "報表對象：            （1：專業部   2：櫃台收文）"
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1140
      Width           =   4572
   End
   Begin VB.Label Label2 
      Caption         =   "收件日期："
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   780
      Width           =   972
   End
End
Attribute VB_Name = "frm010009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/22 日期欄已修改
'Modified by Morgan 2021/8/12 智財法院-->智商法院
Option Explicit

Dim conFCTtoT  As String
Dim strLastCommandText As String
Dim str報表對象(1) As String
Dim strSql As String, iPrint As Integer, Page As Integer
Dim PLeft(0 To 19) As Integer, i As Integer, j As Integer
Dim strTemp(0 To 19) As String, SavDay1 As Integer, SavDay2 As Integer
Dim strLastCommandText_tmp As String, iPrintText As Integer
Dim m_TM01 As String, m_TM02  As String, m_TM03 As String, m_TM04 As String 'Add By Sindy 2012/9/25


Private Sub cmdok_Click(Index As Integer)
'edit by nickc 2007/02/05 不用 dll 了
'Dim objPrintDll001 As Object, varSaveCursor, bolNoData As Boolean, i As Integer, j As Integer
Dim varSaveCursor, bolNoData As Boolean, i As Integer, j As Integer
'92.04.16 nick 改善印表問題
Dim objForPrint(0 To 1, 1 To 7) As Object

   If Index = 0 Then
      For i = 0 To 1
         If CheckKeyIn(i) = False Then
            txtKeyIn(i).SetFocus
            txtKeyIn_GotFocus i
            Exit Sub
         End If
      Next
      varSaveCursor = Screen.MousePointer
      Screen.MousePointer = vbHourglass
      
      Select Case txtKeyIn(1)
         Case "1"
            'Modify By Sindy 2012/9/26 檔案室取消清單,直接以專業部的清單調卷
            'For i = 0 To 1
            For i = 1 To 1 '專業部
            '2012/9/26 End
               For j = 1 To 7
                  'If objForPrint(i, j).PrintRecieve1(txtKeyIn(0), strUserName, i, j) = False Then
                  If PrintRecieve1(txtKeyIn(0), strUserName, i, j) = False Then
                     bolNoData = True
                  End If
               Next j
            Next i
         Case "2"
            'edit by nickc 2005/06/14 從 dll 內移出
            'If objPrintDll001.PrintRecieve2(txtKeyIn(0), strUserName) = False Then
            If PrintRecieve2(txtKeyIn(0), strUserName) = False Then
               bolNoData = True
            End If
      End Select
   '   Set objPrintDll001 = Nothing
      Screen.MousePointer = varSaveCursor
      If bolNoData = False Then
         'Add By Cheng 2002/01/23
         MsgBox "列印完畢!!!", vbInformation
         Me.txtKeyIn(1).Text = Empty
         Me.txtKeyIn(0).SetFocus
      '   Unload Me
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Activate()
   conFCTtoT = "異議,評定,廢止,答辯"
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtKeyIn(0) = GetTaiwanTodayDate
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Boolean
   Select Case intIndex
      Case 0
         If CheckIsTaiwanDate(txtKeyIn(intIndex).Text) Then
             CheckKeyIn = True
         End If
      Case 1
         If Val(txtKeyIn(intIndex)) > 0 And Val(txtKeyIn(intIndex)) < 3 Then
            CheckKeyIn = True
         Else
            ShowMsg MsgText(1045)
         End If
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm010009 = Nothing
End Sub

Private Sub txtKeyIn_GotFocus(Index As Integer)
   txtKeyIn(Index).SelStart = 0
   txtKeyIn(Index).SelLength = Len(txtKeyIn(Index))
End Sub
Private Sub txtKeyIn_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = False Then
      Cancel = True
      txtKeyIn_GotFocus Index
   End If
End Sub

Public Function PrintRecieve1(ByRef strRDate As String, ByRef strUserName As String, intDept As Integer, intCounter As Integer) As Boolean
Dim strSys() As String, i As Integer, j As Integer, strFCTtoT() As String, k As Integer
Dim strTDate As String, strCmdText As String

   iPrintText = intDept
   
'MODIFY BY SONIA 2014/6/17 看不出為何要判斷IF
'   If strLastCommandText = "" Then
'      '2010/1/29 MODIFY BY SONIA 加8智商法院
'      'strLastCommandText_tmp = "select mr01,mr04,mr12,mr12||'-'||decode(mr12,'TF',substr(mr13,1,5),mr13)||decode(mr12,'TF',decode(substr(mr13,6,1),'0','','-'||substr(mr13,6,1)),'')||decode(mr14,'0','','-'||mr14)||decode(mr15,'00','',decode(mr14,'0','-0')||'-'||mr15) mr13,decode(mr05,'1','次日','2','當日','3','無期限') mr05,decode(mr06,NULL,decode(mr07,NULL,decode(mr08,NULL,'',substr(mr08,1,4)-1911||'/'||substr(mr08,5,2)||'/'||substr(mr08,7,2)),mr07||'月'),mr06||'日') mr06,decode(mr16,NULL,mr16,mr16-19110000) mr16,decode(mr10,'1','智慧局','2','內政部','3','經濟部','4','行政院','5','行政法院','6','地方法院','7','其他','8','智商法院') mr10,decode(mr10,'1',1) mrCnt,mr11,mr09  from mailrec where mr02=1999 "
'      'Modify By Sindy 2014/5/2 增加檢查create人員不可為P1X
'      strLastCommandText_tmp = "select mr01,mr04,mr12,mr12||'-'||decode(mr12,'TF',substr(mr13,1,5),mr13)||decode(mr12,'TF',decode(substr(mr13,6,1),'0','','-'||substr(mr13,6,1)),'')||decode(mr14,'0','','-'||mr14)||decode(mr15,'00','',decode(mr14,'0','-0')||'-'||mr15) mr13,decode(mr05,'1','次日','2','當日','3','無期限') mr05,decode(mr06,NULL,decode(mr07,NULL,decode(mr08,NULL,'',substr(mr08,1,4)-1911||'/'||substr(mr08,5,2)||'/'||substr(mr08,7,2)),mr07||'月'),mr06||'日') mr06,decode(mr16,NULL,mr16,mr16-19110000) mr16,decode(mr10,'1','智慧局','2','內政部','3','經濟部','4','行政院','5','行政法院','6','地方法院','7','其他','8','智商法院') mr10,decode(mr10,'1',1) mrCnt,mr11,mr09  from mailrec,staff where mr02=1999 and mr18=st01(+) and substr(st03,1,2)<>'P1'"
'      strLastCommandText = strLastCommandText_tmp
'   Else
'      strLastCommandText = strLastCommandText_tmp
'   End If
      'Modified by Morgan 2015/4/20 +pa75
      'strLastCommandText = "select mr01,mr04,mr12,mr12||'-'||decode(mr12,'TF',substr(mr13,1,5),mr13)||decode(mr12,'TF',decode(substr(mr13,6,1),'0','','-'||substr(mr13,6,1)),'')||decode(mr14,'0','','-'||mr14)||decode(mr15,'00','',decode(mr14,'0','-0')||'-'||mr15) mr13,decode(mr05,'1','次日','2','當日','3','無期限') mr05,decode(mr06,NULL,decode(mr07,NULL,decode(mr08,NULL,'',substr(mr08,1,4)-1911||'/'||substr(mr08,5,2)||'/'||substr(mr08,7,2)),mr07||'月'),mr06||'日') mr06,decode(mr16,NULL,mr16,mr16-19110000) mr16,decode(mr10,'1','智慧局','2','內政部','3','經濟部','4','行政院','5','行政法院','6','地方法院','7','其他','8','智商法院') mr10,decode(mr10,'1',1) mrCnt,mr11,mr09  from mailrec,staff where mr02=1999 and mr18=st01(+) and substr(st03,1,2)<>'P1'"
      'modify by sonia 2015/11/18 +pa26
      'strLastCommandText = "select mr01,mr04,mr12,mr12||'-'||decode(mr12,'TF',substr(mr13,1,5),mr13)||decode(mr12,'TF',decode(substr(mr13,6,1),'0','','-'||substr(mr13,6,1)),'')||decode(mr14,'0','','-'||mr14)||decode(mr15,'00','',decode(mr14,'0','-0')||'-'||mr15) mr13,decode(mr05,'1','次日','2','當日','3','無期限') mr05,decode(mr06,NULL,decode(mr07,NULL,decode(mr08,NULL,'',substr(mr08,1,4)-1911||'/'||substr(mr08,5,2)||'/'||substr(mr08,7,2)),mr07||'月'),mr06||'日') mr06,decode(mr16,NULL,mr16,mr16-19110000) mr16,decode(mr10,'1','智慧局','2','內政部','3','經濟部','4','行政院','5','行政法院','6','地方法院','7','其他','8','智商法院') mr10,decode(mr10,'1',1) mrCnt,mr11,mr09,pa75  from mailrec,staff,patent where mr02=1999 and mr18=st01(+) and substr(st03,1,2)<>'P1' and pa01(+)=MR12 and pa02(+)=MR13 and pa03(+)=MR14 and pa04(+)=MR15"
      'Modified by Morgan 2025/1/7 +pa01,pa02,pa03,pa04,salno
      'Modified by Morgan 2025/2/14 +pid
      strLastCommandText = "select mr01,mr04,mr12,mr12||'-'||decode(mr12,'TF',substr(mr13,1,5),mr13)||decode(mr12,'TF',decode(substr(mr13,6,1),'0','','-'||substr(mr13,6,1)),'')||decode(mr14,'0','','-'||mr14)||decode(mr15,'00','',decode(mr14,'0','-0')||'-'||mr15) mr13,decode(mr05,'1','次日','2','當日','3','無期限') mr05,decode(mr06,NULL,decode(mr07,NULL,decode(mr08,NULL,'',substr(mr08,1,4)-1911||'/'||substr(mr08,5,2)||'/'||substr(mr08,7,2)),mr07||'月'),mr06||'日') mr06,decode(mr16,NULL,mr16,mr16-19110000) mr16,decode(mr10,'1','智慧局','2','內政部','3','經濟部','4','行政院','5','行政法院','6','地方法院','7','其他','8','智商法院') mr10,decode(mr10,'1',1) mrCnt,mr11,mr09,pa75,pa26,pa01,pa02,pa03,pa04,'' salno,'' pid  from mailrec,staff,patent where mr02=1999 and mr18=st01(+) and substr(st03,1,2)<>'P1' and pa01(+)=MR12 and pa02(+)=MR13 and pa03(+)=MR14 and pa04(+)=MR15"
      'end 2015/11/18
      'end 2015/4/20
'END 2014/6/17
   
   'Set cnnConnection = DataEnv001.Connection
   str報表對象(0) = "檔案室"
   str報表對象(1) = "專業部"
   
   strTDate = ChangeTStringToWString(strRDate)
   If GetCkindSys(strTDate, strSys()) Then
      strFCTtoT = Split(conFCTtoT, ",")
      strCmdText = Replace(strLastCommandText, "1999", strTDate)
      Select Case intCounter
          Case 1
             Select Case intDept
                Case 0
                     strCmdText = strCmdText + " and (mr12 = 'P' or mr12='PS') order by mr12,mr13,mr14,mr15,mr01"
                Case 1
                     strCmdText = strCmdText + " and (mr12 = 'P' or mr12='PS') order by mr09,mr12,mr13,mr14,mr15,mr01"
             End Select
          Case 2
             'Modify By Sindy 2018/1/17
             'strCmdText = strCmdText + " and (mr12 = 'FCP' or mr12='FG') order by mr12,mr13,mr14,mr15,mr01"
             strCmdText = "select mr01,mr04,mr12,mr12||'-'||decode(mr12,'TF',substr(mr13,1,5),mr13)||decode(mr12,'TF',decode(substr(mr13,6,1),'0','','-'||substr(mr13,6,1)),'')||decode(mr14,'0','','-'||mr14)||decode(mr15,'00','',decode(mr14,'0','-0')||'-'||mr15) mr13,decode(mr05,'1','次日','2','當日','3','無期限') mr05,decode(mr06,NULL,decode(mr07,NULL,decode(mr08,NULL,'',substr(mr08,1,4)-1911||'/'||substr(mr08,5,2)||'/'||substr(mr08,7,2)),mr07||'月'),mr06||'日') mr06,decode(mr16,NULL,mr16,mr16-19110000) mr16,decode(mr10,'1','智慧局','2','內政部','3','經濟部','4','行政院','5','行政法院','6','地方法院','7','其他','8','智商法院') mr10,decode(mr10,'1',1) mrCnt,mr11,mr09,pa75,pa26,s2.st02 FCPEmp" & _
                          " from mailrec,staff s1,patent,fagent,nation,staff s2" & _
                          " where mr02=1999 and mr18=s1.st01(+)" & _
                          " and substr(s1.st03,1,2)<>'P1'" & _
                          " and pa01(+)=MR12 and pa02(+)=MR13 and pa03(+)=MR14 and pa04(+)=MR15" & _
                          " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9)" & _
                          " and na01(+)=fa10 and s2.st01(+)=na16"
             strCmdText = Replace(strCmdText, "1999", strTDate)
             strCmdText = strCmdText + " and (mr12 = 'FCP' or mr12='FG')"
             strLastCommandText_tmp = strCmdText & " order by FCPEmp,mr12,mr13,mr14,mr15,mr01"
             strCmdText = strCmdText & " order by mr12,mr13,mr14,mr15,mr01"
             '2018/1/17 END
          Case 3
             Select Case intDept
                Case 0
                     strCmdText = strCmdText + " and (mr12 = 'CFP' or mr12='CPS') order by mr12,mr13,mr14,mr15,mr01"
                Case 1
                     strCmdText = strCmdText + " and (mr12 = 'CFP' or mr12='CPS') order by mr09,mr12,mr13,mr14,mr15,mr01"
             End Select
          Case 4
             Select Case intDept
                Case 0
                     strCmdText = strCmdText + " and ((mr12 = 'FCT' and mr09 not in ('異議','廢止','評定','答辯','應予撤銷','註冊無效','提訴願理由','原處份撤銷','陳述意見')) or mr12='S' or mr12='CFT' or mr12='CFC') order by mr12,mr13,mr14,mr15,mr01"
                Case 1
                     strCmdText = strCmdText + " and ((mr12 = 'FCT' and mr09 not in ('異議','廢止','評定','答辯','應予撤銷','註冊無效','提訴願理由','原處份撤銷','陳述意見')) or mr12='S' or mr12='CFT' or mr12='CFC') order by mr09,mr12,mr13,mr14,mr15,mr01"
             End Select
          Case 5
             Select Case intDept
                Case 0
                     strCmdText = strCmdText + " and (mr12 = 'L' or mr12='LA') order by mr12,mr13,mr14,mr15,mr01"
                Case 1
                     strCmdText = strCmdText + " and (mr12 = 'L' or mr12='LA') order by mr09,mr12,mr13,mr14,mr15,mr01"
             End Select
          Case 6
             Select Case intDept
                Case 0
                     'Modify By Sindy 2009/07/24 增加LIN系統類別
                     'modify by sonia 2019/7/29 +ACS系統類別
                     strCmdText = strCmdText + " and (mr12 = 'FCL' or mr12='CFL' or mr12 = 'LIN' or mr12 = 'ACS') order by mr12,mr13,mr14,mr15,mr01"
                Case 1
                     'Modify By Sindy 2009/07/24 增加LIN系統類別
                     'modify by sonia 2019/7/29 +ACS系統類別
                     strCmdText = strCmdText + " and (mr12 = 'FCL' or mr12='CFL' or mr12 = 'LIN' or mr12 = 'ACS') order by mr09,mr12,mr13,mr14,mr15,mr01"
             End Select
          Case 7
             Select Case intDept
                Case 0
                     strCmdText = strCmdText + " and ((mr12 = 'FCT' and mr09 in ('異議','廢止','評定','答辯','應予撤銷','註冊無效','提訴願理由','原處份撤銷','陳述意見')) or mr12 like 'T%') order by mr12,mr13,mr14,mr15,mr01"
                Case 1
                     strCmdText = strCmdText + " and ((mr12 = 'FCT' and mr09 in ('異議','廢止','評定','答辯','應予撤銷','註冊無效','提訴願理由','原處份撤銷','陳述意見')) or mr12 like 'T%') order by mr09,mr12,mr13,mr14,mr15,mr01"
            End Select
      End Select
      '印表
      '92.06.24 nick
      'DataEnv001.Commands(1).CommandText = strCmdText
      'DataEnv001.cmd010001
      'If DataEnv001.rscmd010001.RecordCount > 0 Then
      '   datrpt010001.Sections(2).Controls("rptlblRDate").Caption = ChangeTStringToTDateString(strRDate)
      '   datrpt010001.Sections(2).Controls("rptlblPDate").Caption = ChangeTStringToTDateString(GetTaiwanTodayDate)
      '   datrpt010001.Sections(2).Controls("rptlblPPerson").Caption = strUserName
      '   datrpt010001.Sections(2).Controls("rptlblName").Caption = str報表對象(intDept)
      '   datrpt010001.Sections(2).Controls("label11").Caption = intCounter
      '   datrpt010001.Sections(2).Controls("label11").Visible = False
      '   datrpt010001.PrintReport
      '   DoEvents
      'End If
      'If DataEnv001.rscmd010001.State = adStateOpen Then
      '   DataEnv001.rscmd010001.Close
      'End If
      
      'Added by Morgan 2025/1/7
      If strSrvDate(1) >= P業務區劃分啟用日 And intCounter = 1 And intDept = 1 Then
         Call PrintDataP(strCmdText)
      Else
      'end 2025/1/7
      
         Call PrintData(strCmdText)
         
      End If 'Added by Morgan 2025/1/7
      PrintRecieve1 = True
   End If
   Exit Function
End Function

Public Function PrintRecieve2(ByRef strRDate As String, ByRef strUserName As String) As Boolean
Dim strSys() As String, i As Integer, j As Integer, strFCTtoT() As String, k As Integer
Dim strTDate As String, strCmdText As String

   iPrintText = 2
   
'MODIFY BY SONIA 2014/6/17 看不出為何要判斷IF, 6/17報表未排序故同時加入ORDER BY MR01
'   If strLastCommandText = "" Then
'      '2010/1/29 MODIFY BY SONIA 加8智商法院
'      'strLastCommandText_tmp = "select mr01,mr04,mr12,mr12||'-'||decode(mr12,'TF',substr(mr13,1,5),mr13)||decode(mr12,'TF',decode(substr(mr13,6,1),'0','','-'||substr(mr13,6,1)),'')||decode(mr14,'0','','-'||mr14)||decode(mr15,'00','',decode(mr14,'0','-0')||'-'||mr15) mr13,decode(mr05,'1','次日','2','當日','3','無期限') mr05,decode(mr06,NULL,decode(mr07,NULL,decode(mr08,NULL,'',substr(mr08,1,4)-1911||'/'||substr(mr08,5,2)||'/'||substr(mr08,7,2)),mr07||'月'),mr06||'日') mr06,decode(mr16,NULL,mr16,mr16-19110000) mr16,decode(mr10,'1','智慧局','2','內政部','3','經濟部','4','行政院','5','行政法院','6','地方法院','7','其他','8','智商法院') mr10,decode(mr10,'1',1) mrCnt,mr11,mr09  from mailrec where mr02=1999 "
'      'Modify By Sindy 2014/5/2 增加檢查create人員不可為P1X
'      strLastCommandText_tmp = "select mr01,mr04,mr12,mr12||'-'||decode(mr12,'TF',substr(mr13,1,5),mr13)||decode(mr12,'TF',decode(substr(mr13,6,1),'0','','-'||substr(mr13,6,1)),'')||decode(mr14,'0','','-'||mr14)||decode(mr15,'00','',decode(mr14,'0','-0')||'-'||mr15) mr13,decode(mr05,'1','次日','2','當日','3','無期限') mr05,decode(mr06,NULL,decode(mr07,NULL,decode(mr08,NULL,'',substr(mr08,1,4)-1911||'/'||substr(mr08,5,2)||'/'||substr(mr08,7,2)),mr07||'月'),mr06||'日') mr06,decode(mr16,NULL,mr16,mr16-19110000) mr16,decode(mr10,'1','智慧局','2','內政部','3','經濟部','4','行政院','5','行政法院','6','地方法院','7','其他','8','智商法院') mr10,decode(mr10,'1',1) mrCnt,mr11,mr09  from mailrec,staff where mr02=1999 and mr18=st01(+) and substr(st03,1,2)<>'P1'"
'      strLastCommandText = strLastCommandText_tmp
'   Else
'      strLastCommandText = strLastCommandText_tmp
'   End If
      strLastCommandText = "select mr01,mr04,mr12,mr12||'-'||decode(mr12,'TF',substr(mr13,1,5),mr13)||decode(mr12,'TF',decode(substr(mr13,6,1),'0','','-'||substr(mr13,6,1)),'')||decode(mr14,'0','','-'||mr14)||decode(mr15,'00','',decode(mr14,'0','-0')||'-'||mr15) mr13,decode(mr05,'1','次日','2','當日','3','無期限') mr05,decode(mr06,NULL,decode(mr07,NULL,decode(mr08,NULL,'',substr(mr08,1,4)-1911||'/'||substr(mr08,5,2)||'/'||substr(mr08,7,2)),mr07||'月'),mr06||'日') mr06,decode(mr16,NULL,mr16,mr16-19110000) mr16,decode(mr10,'1','智慧局','2','內政部','3','經濟部','4','行政院','5','行政法院','6','地方法院','7','其他','8','智商法院') mr10,decode(mr10,'1',1) mrCnt,mr11,mr09  from mailrec,staff where mr02=1999 and mr18=st01(+) and substr(st03,1,2)<>'P1' order by mr01"
'END 2014/6/17
   
   str報表對象(0) = "櫃檯收文"
   'On Error GoTo Err
   'DataEnv001.Commands(1).CommandText = Replace(strLastCommandText, "1999", ChangeTStringToWString(strRDate))
   'DataEnv001.cmd010001
   'datrpt010001.Sections(2).Controls("rptlblRDate").Caption = ChangeTStringToTDateString(strRDate)
   'datrpt010001.Sections(2).Controls("rptlblPDate").Caption = ChangeTStringToTDateString(GetTaiwanTodayDate)
   'datrpt010001.Sections(2).Controls("rptlblPPerson").Caption = strUserName
   'datrpt010001.Sections(2).Controls("rptlblName").Caption = "櫃台收文"
   'datrpt010001.Sections(2).Controls("label11").Visible = False
   '
   'datrpt010001.PrintReport: DoEvents
   'While datrpt010001.AsyncCount > 0
   '    DoEvents
   'Wend
   'Unload datrpt010001
   'If DataEnv001.rscmd010001.State = adStateOpen Then
   '   DataEnv001.rscmd010001.Close
   'End If
   PrintData Replace(strLastCommandText, "1999", ChangeTStringToWString(strRDate))
   PrintRecieve2 = True
   Exit Function
   'Err:
   'ErrorLog
End Function

Sub PrintTitle(strTM01 As String, strFCPEmp As String)
Dim strTitle As String
   
   GetPleft
   iPrint = 500
   Printer.Orientation = 1
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 14
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   'Modify By Sindy 2018/1/17
   'Printer.CurrentX = 4650
   strTitle = "主管機關來函收文簿" & IIf(strFCPEmp <> "", "-FCP(" & strFCPEmp & ")", "")
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTitle) / 2)
   Printer.CurrentY = iPrint
   Printer.Print strTitle
   '2018/1/17 END
   iPrint = iPrint + 500
   Printer.Font.Size = 9
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   'edit by nickc 2005/06/14
   'Printer.Print "報表對象：" & IIf(iPrintText = 1, "專業部", "檔案室")
   Printer.Print "報表對象：" & IIf(iPrintText = 1, str報表對象(1), str報表對象(0))
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人　：" & strUserName
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "製表日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "收件日期：" & ChangeTStringToTDateString(txtKeyIn(0).Text)
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   'Added by Morgan 2015/4/20
   'Move by Lydia 2016/02/02 原本在收件日期與頁次的中間，移到下一列
   If txtKeyIn(1) = "1" And (m_TM01 = "FCP" Or m_TM01 = "P") Then
      iPrint = iPrint + 300 'Added by Lydia 2016/02/02
      'Modified by Lydia 2016/02/02
      'Printer.CurrentX = 2700
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      'modify by sonia 2015/11/18 加入DOW期限來函加 v
      'Modified by Lydia 2016/02/02 加入UNIUS之有期限案件
      'Printer.Print "起算日前有 * 號者為先正達有期限之來函，有 ** 號者為 DOW 有期限之來函"
      'Modified by Lydia 2016/08/15 * 號者加上Y20656 + X702201 or X70286
      'Printer.Print "起算日前有 * 號者為先正達有期限之來函，有 ** 號者為 DOW 有期限之來函，有 *** 號者為UNIUS之有期限案件"
      'Modified by Lydia 2017/03/08 * 號者加上Y53942 'Modified by Lydia 2017/03/14 Y53942更名為Xperi
      Printer.Print "起算日前有 * 號者為先正達有期限之來函或Lerner+泰斯拉公司及英帆薩斯公司+Y53942 Xperi 的案件"
      iPrint = iPrint + 300
      Printer.CurrentX = 500
      Printer.CurrentY = iPrint
      'Modify By Sindy 2017/5/9
      'Printer.Print "有 ** 號者為 DOW 有期限之來函，有 *** 號者為UNIUS之有期限案件"
      'Printer.Print "有 ** 號者為 DOW 有期限之來函，或有 ** 號者為 Sandvik OA 需2日內報告，有 *** 號者為UNIUS之有期限案件"
      'Modify By Sindy 2017/10/20
      Printer.Print "有 ** 號者為 DOW 有期限之來函，或有 ** 號者為 需2日內報告，有 *** 號者為UNIUS或YASUTOMI之有期限案件"
      '2017/10/20 END
      '2017/5/9 END
   End If
   'end 2015/4/20
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "收件號"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "來函號數"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print IIf(strFCPEmp <> "", "  ", ""); "本所案號"
   'Add By Sindy 2018/1/17
   'If (strTM01 = "FCP" Or strTM01 = "FG") And txtKeyIn(1) = 1 Then
   If strFCPEmp <> "" Then
      Printer.CurrentX = PLeft(9)
      Printer.CurrentY = iPrint
      Printer.Print "程序分區"
   End If
   '2018/1/17 END
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "起算日"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "期限"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "本所期限"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "政府機關"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "機關文號"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "備註"
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
End Sub

Sub PrintDatil()
   Dim bolHigher As Boolean
   
   bolHigher = False
   For i = 0 To 9 '8
       If i = 1 Then
         If GetTextLength(strTemp(i)) > 12 Then
            Printer.CurrentX = PLeft(i)
            Printer.CurrentY = iPrint - 90
            Printer.Print PUB_StrToStr(strTemp(i), 12)
            Printer.CurrentX = PLeft(i)
            Printer.CurrentY = iPrint + 90
            Printer.Print Mid(strTemp(i), Len(PUB_StrToStr(strTemp(i), 12)) + 1)
            bolHigher = True
         Else
            Printer.CurrentX = PLeft(i)
            Printer.CurrentY = iPrint
            Printer.Print strTemp(i)
         End If
       Else
         Printer.CurrentX = PLeft(i)
         Printer.CurrentY = iPrint
         'Added by Morgan 2025/1/7
         If iPrintText = 1 Then
            PUB_PrintUnicodeText strTemp(i), Printer.CurrentX, Printer.CurrentY, 0
         Else
         'end 2025/1/7
            Printer.Print strTemp(i)
         End If
      End If
   Next i
   
   'Modified by Morgan 2022/8/31
   'iPrint = iPrint + 300
   If bolHigher Then
      iPrint = iPrint + 400
   Else
      iPrint = iPrint + 300
   End If
   'end 2022/8/31
End Sub

Sub GetPleft()
   Erase PLeft
   
   PLeft(0) = 500
   PLeft(1) = 1500 '1600 來函號數
   'Modified by Morgan 2022/8/31 來函號數加寬,備註縮小
   'PLeft(2) = 2600 '2750 本所案號
   'PLeft(9) = 3800 '程序分區
   'PLeft(3) = 4650 '4400 起算日
   'PLeft(4) = 5600 '5400 期限
   'PLeft(5) = 6500
   'PLeft(6) = 7400
   'PLeft(7) = 8650
   'PLeft(8) = 9650
   PLeft(2) = 2800 '本所案號
   PLeft(9) = 4000 '程序分區
   PLeft(3) = 4850 '起算日
   PLeft(4) = 5800 '期限
   PLeft(5) = 6700
   PLeft(6) = 7600
   PLeft(7) = 8850
   PLeft(8) = 9850
   'end 2022/8/31
End Sub

Sub PrintData(strSql As String)
'Add By Sindy 2012/9/25
Dim intCnt As Integer, strCaseNo(50) As String, strCP09(50) As String, strCP10(50) As String, strCP27(50) As String
Dim strCP28(50) As String 'Add By Sindy 2012/11/08
Dim iRow As Integer
Dim bolChk As Boolean 'Add By Sindy 2012/11/23
   intCnt = 0 '記錄有幾筆一申請書多件
'2012/9/25 End
'Add By Sindy 2013/1/23
Dim intDCCnt As Integer
Dim strDCase(50) As String
   intDCCnt = 0 '記錄有幾筆分割案
'2013/1/23 End
   
   'Add by Sindy 2011/11/9
   '故意設定紙張屬性以便清除印表機狀態(相同印表機驅動程式會沿用原設定值,Ex.進紙槽)
   Printer.PaperSize = 9
   Printer.EndDoc
   'End 2011/11/9
   CheckOC
   Page = 1
   SavDay1 = 0
   SavDay2 = 0
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       With adoRecordset
           .MoveFirst
           m_TM01 = "" & .Fields("MR12").Value
           Call PrintTitle(m_TM01, "")
           Do While .EOF = False
               strTemp(0) = CheckStr(.Fields("MR01").Value)
               strTemp(1) = CheckStr(.Fields("MR04").Value)
               strTemp(2) = CheckStr(.Fields("MR13").Value) '本所案號
               Call GetTMCaseNo(strTemp(2)) 'Add By Sindy 2018/7/17
               strTemp(3) = CheckStr(.Fields("MR05").Value)
               'Modify By Sindy 2017/12/11 改用Table記錄備註
               If txtKeyIn(1) = "1" Then
                  'Modify By Sindy 2018/7/17 m_TM01 ==> 改傳 m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
                  strTemp(3) = PUB_ReadIPOListMemo(m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04, "" & .Fields("pa75"), "" & .Fields("pa26"), CheckStr(.Fields("MR09").Value), "" & .Fields("MR16")) & strTemp(3)
               End If
               '2017/12/11 END
'               'Added by Morgan 2015/4/20
'               '先正達有期限來函起算日前加星號
'               'Modified by Lydia 2016/08/15 剔除核准函
'               If txtKeyIn(1) = "1" And (m_TM01 = "FCP" Or m_TM01 = "P") And CheckStr(.Fields("MR09").Value) <> "核准" Then
'                 'Modified by Lydia 2016/08/15  Y20656 (Lerner)+X7072201(Tessera, Inc.) & X70286(Invensas Corporation)之案件,來函案件性質除核准函外之有期限者,加註*
'                  'If .Fields("MR16") > 0 And Not IsNull(.Fields("pa75")) And InStr("Y4830900,Y4830901,Y4830902,Y4830903,Y4830904,Y4830905,Y4830908,Y5132600", Left("" & .Fields("pa75"), 8)) > 0 Then
'                  'Modified by Lydia 2017/03/08 +Y53942 Tessera => InStr(Left("" & .Fields("pa75"), 8), "Y20656") > 0
'                  'Modified by Lydia 2017/05/19 調整
'                  'If .Fields("MR16") > 0 And Not IsNull(.Fields("pa75")) And (InStr("Y4830900,Y4830901,Y4830902,Y4830903,Y4830904,Y4830905,Y4830908,Y5132600", Left("" & .Fields("pa75"), 8)) > 0) Or (InStr(Left("" & .Fields("pa75"), 8), "Y53942") > 0) _
'                  '   Or (InStr(Left("" & .Fields("pa75"), 8), "Y20656") > 0 And Not IsNull(.Fields("pa26")) And (InStr(Left("" & .Fields("pa26"), 8), "X7072201") > 0 Or InStr(Left("" & .Fields("pa26"), 8), "X70286") > 0)) Then
'                  'Modify By Sindy 2017/8/3 Y339400(Foley&Lardner,LLP)+申請人X48991 INTERSIL AMERICAS INC.
'                  If .Fields("MR16") > 0 _
'                     And Not IsNull(.Fields("pa75")) And ( _
'                     (InStr("Y4830900,Y4830901,Y4830902,Y4830903,Y4830904,Y4830905,Y4830908,Y5132600,Y5394200", Left("" & .Fields("pa75"), 8)) > 0) _
'                     Or (InStr(Left("" & .Fields("pa75"), 8), "Y2065600") > 0 And Not IsNull(.Fields("pa26")) And (InStr(Left("" & .Fields("pa26"), 8), "X7072201") > 0 Or InStr(Left("" & .Fields("pa26"), 8), "X70286") > 0)) _
'                     Or (InStr(Left("" & .Fields("pa75"), 8), "Y3394000") > 0 And Not IsNull(.Fields("pa26")) And InStr(Left("" & .Fields("pa26"), 8), "X4899100") > 0) _
'                     ) Then
'                  'END 2017/05/19
'                     strTemp(3) = "*" & strTemp(3)
'                  End If
'               End If
'               'end 2015/4/20
'               'modify by sonia 2015/11/18 加入DOW期限來函加 v
'               'Modified by Lydia 2016/08/15 剔除核准函
'               If txtKeyIn(1) = "1" And (m_TM01 = "FCP" Or m_TM01 = "P") And CheckStr(.Fields("MR09").Value) <> "核准" Then
'                  If .Fields("MR16") > 0 And Not IsNull(.Fields("pa75")) And InStr("Y2245700", Left("" & .Fields("pa75"), 8)) > 0 Then
'                     strTemp(3) = "**" & strTemp(3)
'                  ElseIf .Fields("MR16") > 0 And Not IsNull(.Fields("pa26")) And InStr("X6740200,X6740201,X6740202,X6050700,X6050701,X7074900", Left("" & .Fields("pa26"), 8)) > 0 Then
'                     strTemp(3) = "**" & strTemp(3)
'                  End If
'               End If
'               'end 2015/11/18
'
'               'Add By Sindy 2017/5/9 有 ** 號者 需2日內報告
'               If txtKeyIn(1) = "1" And _
'                  (m_TM01 = "FCP" Or m_TM01 = "P") And CheckStr(.Fields("MR09").Value) <> "核准" Then
'                  '有期限
'                  If .Fields("MR16") > 0 And Not IsNull(.Fields("pa75")) Then
'                     '有 ** 號者為 Sandvik OA 需2日內報告
'                     If InStr("Y5285900", Left("" & .Fields("pa75"), 8)) > 0 Or _
'                        InStr("Y5179901", Left("" & .Fields("pa75"), 8)) > 0 Then
'                        strTemp(3) = "**" & strTemp(3)
'                     End If
'                  End If
'               ElseIf txtKeyIn(1) = "1" And _
'                  (m_TM01 = "FCP" Or m_TM01 = "P") Then
'                  '有期限
'                  If .Fields("MR16") > 0 And Not IsNull(.Fields("pa75")) Then
'                     'Modify By Sindy 2017/10/20 有 ** 號者為 Shiga International Patent Office + Nippon Soda Co., Ltd. 需2日內報告
'                     '<Y47453> Shiga International Patent Office + <X55778> Nippon Soda Co., Ltd.
'                     If Not IsNull(.Fields("pa26")) Then
'                        If InStr("Y4745300", Left("" & .Fields("pa75"), 8)) > 0 And _
'                           InStr("X5577800", Left("" & .Fields("pa26"), 8)) > 0 Then
'                           strTemp(3) = "**" & strTemp(3)
'                        End If
'                     End If
'                     '2017/10/20 END
'                  End If
'               End If
'               '2017/5/9 END
'
'               'Added by Lydia 2016/02/02 加入UNIUS之有期限案件
'               'Modified by Lydia 2016/08/15 剔除核准函
'               If txtKeyIn(1) = "1" And (m_TM01 = "FCP" Or m_TM01 = "P") And CheckStr(.Fields("MR09").Value) <> "核准" Then
'                  'Modified by Lydia 2017/05/19 調整
'                  'If .Fields("MR16") > 0 And Not IsNull(.Fields("pa75")) And InStr(Left("" & .Fields("pa75"), 8), "Y51508") > 0 Then
'                  If .Fields("MR16") > 0 And Not IsNull(.Fields("pa75")) And InStr("Y5150800", Left("" & .Fields("pa75"), 8)) > 0 Then
'                     strTemp(3) = "***" & strTemp(3)
'                  End If
'               End If
'               'end 2016/02/02
               
               strTemp(4) = CheckStr(.Fields("MR06").Value)
               strTemp(5) = IIf(CheckStr(.Fields("MR16").Value) <> "", Format(CheckStr(.Fields("MR16").Value), "##/##/##"), "")
               strTemp(6) = CheckStr(.Fields("MR10").Value)
               strTemp(7) = CheckStr(.Fields("MR11").Value)
               strTemp(8) = CheckStr(.Fields("MR09").Value) '備註:移轉
               'Add By Sindy 2012/9/25
               If iPrintText = 1 Then '專業部報表才要跑下列程式
                  Call GetTMCaseNo(strTemp(2))
                  If m_TM01 = "FCT" Then
                     bolChk = False 'Add By Sindy 2012/11/23
                     'Add By Sindy 2014/4/30 櫃檯人員輸入的備註裡要有變更或移轉字樣才要管是否為一申請書多件
                     If InStr(strTemp(8), "變更") > 0 Or InStr(strTemp(8), "移轉") > 0 Then
                     '2014/4/30 END
                        '檢查是否為一申請書多件(此案號AB類最大發文日的CP148=Y)
                        'Modify By Sindy 2012/11/28 + and CP24 is null(有CP24不出清單)
                        'Modify By Sindy 2013/5/24 +tm16
                        'Modify By Sindy 2014/6/18 調整直接抓最新的變更或移轉程序
'                        strSql = "SELECT cp09,cp10,cp27,cp148,cp28,tm16 " & _
'                                 "FROM CaseProgress,trademark " & _
'                                 "WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' " & _
'                                 "AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
'                                 "AND CP09<'C' AND CP27 is not null and CP130 is not null and CP24 is null " & _
'                                 "order by CP27 desc,CP82 desc "
                        strSql = "SELECT cp09,cp10,cp27,cp148,cp28,tm16 " & _
                                 "FROM CaseProgress,trademark " & _
                                 "WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' " & _
                                 "AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) " & _
                                 "AND CP10='" & IIf(InStr(strTemp(8), "變更") > 0, "301", "501") & "' AND CP27 is not null and CP130 is not null and CP24 is null " & _
                                 "order by CP27 desc,CP82 desc "
                        '2014/6/18 END
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           If RsTemp.Fields("cp148") = "Y" Then
                              'Modify By Sindy 2013/5/24 若抓出CP148='Y'的資料為FCT的變更案,且該案之TM16 is NULL且來函輸入備註前2個字為'核准'時,則不出多件一申請書之清單
                              If RsTemp.Fields("cp10") = "301" And Trim("" & RsTemp.Fields("tm16")) = "" And Left(strTemp(8), 2) = "核准" Then
                                 GoTo ReadEnd
                              End If
                              '2013/5/24 End
                              If PUB_ChkIsOneAppMuchCase("" & RsTemp.Fields("cp28")) = True Then 'Add By Sindy 2012/11/08
                                 bolChk = True 'Add By Sindy 2012/11/23
                                 intCnt = intCnt + 1
                                 strCaseNo(intCnt) = Trim(strTemp(2))
                                 strCP09(intCnt) = RsTemp.Fields("cp09")
                                 strCP10(intCnt) = RsTemp.Fields("cp10")
                                 strCP27(intCnt) = "" & RsTemp.Fields("cp27")
                                 strCP28(intCnt) = "" & RsTemp.Fields("cp28") 'Add By Sindy 2012/11/08
                              End If
                           End If
                        End If
                        'Add By Sindy 2012/11/23 上列條件沒有抓到資料時,比對備註前2個字再讀取一次資料
                        If bolChk = False Then
                           'Modify By Sindy 2012/11/28 + and CP24 is null(有CP24不出清單)
                           strSql = "SELECT cp09,cp10,cp27,cp148,cp28 " & _
                                    "FROM CaseProgress,casepropertymap " & _
                                    "WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' " & _
                                    "AND CP09<'C' AND CP27 is not null and CP130 is not null and CP24 is null " & _
                                    "AND cp01=cpm01(+) AND cp10=cpm02(+) AND substr(cpm03,1,2)='" & Left(strTemp(8), 2) & "' " & _
                                    "order by CP27 desc,CP82 desc "
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                           If intI = 1 Then
                              If RsTemp.Fields("cp148") = "Y" Then
                                 If PUB_ChkIsOneAppMuchCase("" & RsTemp.Fields("cp28")) = True Then
                                    bolChk = True
                                    intCnt = intCnt + 1
                                    strCaseNo(intCnt) = Trim(strTemp(2))
                                    strCP09(intCnt) = RsTemp.Fields("cp09")
                                    strCP10(intCnt) = RsTemp.Fields("cp10")
                                    strCP27(intCnt) = "" & RsTemp.Fields("cp27")
                                    strCP28(intCnt) = "" & RsTemp.Fields("cp28")
                                 End If
                              End If
                           End If
                        End If
                        '2012/11/23 End
                     End If '2014/4/30 END
                  End If
ReadEnd:
                  'Add By Sindy 2013/1/23
                  If (m_TM01 = "FCT" Or m_TM01 = "T") And InStr("" & .Fields("MR09").Value, "分割") > 0 Then
                     strSql = "SELECT cp09,cp10,cp27,cp148,cp28 " & _
                              "From Trademark, CaseProgress " & _
                              "WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "' " & _
                              "AND TM01=CP01 AND TM02=CP02 AND TM03=CP03 AND TM04=CP04 " & _
                              "AND CP10='308' AND CP27 is not null and CP57 is null " & _
                              "AND TM28='1' "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        If RsTemp.RecordCount > 0 Then
                           intDCCnt = intDCCnt + 1
                           strDCase(intDCCnt) = Trim(strTemp(2))
                        End If
                     End If
                  End If
                  '2013/1/23 End
               End If
               '2012/9/25 End
               strTemp(8) = StrToStr(strTemp(8), 8)
               'Add By Sindy 2018/1/17 FCP程序分區
               If (m_TM01 = "FCP" Or m_TM01 = "FG") And txtKeyIn(1) = 1 Then
                  strTemp(9) = CheckStr(.Fields("FCPEmp").Value)
               Else
                  strTemp(9) = ""
               End If
               '2018/1/17 END
               If .Fields("MR01").Value <> "" Then
                   SavDay1 = SavDay1 + 1
               End If
               If Trim(.Fields("MRCNT").Value) = "1" Then
                   SavDay2 = SavDay2 + 1
               End If
               If iPrint > 15000 Then
                   Printer.NewPage
                   Page = Page + 1
                   Call PrintTitle(m_TM01, "")
               End If
               PrintDatil
               .MoveNext
           Loop
       End With
       PrintEnd
       Printer.EndDoc
       'Add By Sindy 2018/1/17
       If (m_TM01 = "FCP" Or m_TM01 = "FG") And txtKeyIn(1) = 1 Then
          Call PrintFCPList
       End If
       '2018/1/17 END
       'Add By Sindy 2012/9/25
       If intCnt > 0 Then
         For iRow = 1 To intCnt
            'Modify By Sindy 2012/11/08 +strCP28(iRow)
            Call PrintFCTList(strCaseNo(iRow), strCP09(iRow), strCP10(iRow), strCP27(iRow), strCP28(iRow))
         Next iRow
       End If
       '2012/9/25 End
       'Add By Sindy 2013/1/23 分割子案清單
       If intDCCnt > 0 Then
         For iRow = 1 To intDCCnt
            Call PrintDCList(strDCase(iRow))
         Next iRow
       End If
       '2013/1/23 End
   End If
   CheckOC
End Sub

Private Sub GetTMCaseNo(strCaseNo As String)
Dim MyArr As Variant
   
   MyArr = Split(strCaseNo, "-")
   m_TM01 = "": m_TM02 = "": m_TM03 = "": m_TM04 = ""
   For i = 0 To UBound(MyArr)
      If i = 0 Then m_TM01 = MyArr(i)
      If i = 1 Then m_TM02 = MyArr(i)
      If i = 2 Then m_TM03 = MyArr(i)
      If i = 3 Then m_TM04 = MyArr(i)
   Next i
   If m_TM03 = "" Then m_TM03 = "0"
   If m_TM04 = "" Then m_TM04 = "00"
End Sub

Sub PrintEnd()
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = 700
   Printer.CurrentY = iPrint
   Printer.Print "共" & Trim(SavDay1) & "筆"
   Printer.CurrentX = 2700
   Printer.CurrentY = iPrint
   Printer.Print "智慧局共" & Trim(SavDay2) & "筆"
End Sub

'Add By Sindy 2014/8/8
Sub PrintTitle2(m_CP10 As String)
   Page = Page + 1
   PLeft(0) = 1000
   PLeft(1) = 3000
   iPrint = 500
   Printer.Orientation = 1
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 14
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("主管機關來函收文簿") / 2)
   Printer.CurrentY = iPrint
   Printer.Print "主管機關來函收文簿"
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("一申請書多件" & GetPrjState4(m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04, m_CP10) & "案清單") / 2)
   Printer.CurrentY = iPrint
   Printer.Print "一申請書多件" & GetPrjState4(m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04, m_CP10) & "案清單"
   iPrint = iPrint + 500
   Printer.Font.Size = 9
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "報表對象：" & IIf(iPrintText = 1, str報表對象(1), str報表對象(0))
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人　：" & strUserName
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "製表日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "收件日期：" & ChangeTStringToTDateString(txtKeyIn(0).Text)
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "來函案件"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "未來函案件"
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
End Sub

'Add By Sindy 2018/1/17 FCP程序分區清單
Private Sub PrintFCPList()
Dim strFCPEmp As String
   
   '故意設定紙張屬性以便清除印表機狀態(相同印表機驅動程式會沿用原設定值,Ex.進紙槽)
   Printer.PaperSize = 9
   Printer.EndDoc
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strLastCommandText_tmp, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
      With adoRecordset
         .MoveFirst
         m_TM01 = "" & .Fields("MR12").Value
         'Modified by Morgan 2018/1/22
         'If strFCPEmp = "" Or strFCPEmp <> strFCPEmp Then
         Do While .EOF = False
            If strFCPEmp = "" Or strFCPEmp <> CheckStr(.Fields("FCPEmp").Value) Then
            'end 2018/1/22
              If strFCPEmp <> "" Then
                 PrintEnd
                 Printer.EndDoc
              End If
              Page = 1
              SavDay1 = 0
              SavDay2 = 0
              Call PrintTitle("", "" & .Fields("FCPEmp").Value)
            End If
        'Do While .EOF = False 'Removed by Morgan 2018/1/22
        
            strTemp(0) = CheckStr(.Fields("MR01").Value)
            strTemp(1) = CheckStr(.Fields("MR04").Value)
            strTemp(2) = CheckStr(.Fields("MR13").Value) '本所案號
            strTemp(3) = CheckStr(.Fields("MR05").Value)
            'Add By Sindy 2018/2/1
            Call GetTMCaseNo(strTemp(2))
            strSql = "SELECT substr(nvl(GetEmailFlag(CP01||CP02||CP03||CP04),' '),1,1) eMail " & _
                     "FROM CaseProgress " & _
                     "WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' " & _
                     " and rownum=1"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strTemp(2) = RsTemp.Fields("eMail") & " " & strTemp(2)
            End If
            '2018/2/1 END
            '改用Table記錄備註
            If txtKeyIn(1) = "1" Then
               'Modify By Sindy 2018/7/17 m_TM01 ==> 改傳 m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
               strTemp(3) = PUB_ReadIPOListMemo(m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04, "" & .Fields("pa75"), "" & .Fields("pa26"), CheckStr(.Fields("MR09").Value), "" & .Fields("MR16")) & strTemp(3)
            End If
            strTemp(4) = CheckStr(.Fields("MR06").Value)
            strTemp(5) = IIf(CheckStr(.Fields("MR16").Value) <> "", Format(CheckStr(.Fields("MR16").Value), "##/##/##"), "")
            strTemp(6) = CheckStr(.Fields("MR10").Value)
            strTemp(7) = CheckStr(.Fields("MR11").Value)
            strTemp(8) = CheckStr(.Fields("MR09").Value) '備註:移轉
            strTemp(8) = StrToStr(strTemp(8), 8)
            strTemp(9) = "" 'CheckStr(.Fields("FCPEmp").Value)
            strFCPEmp = CheckStr(.Fields("FCPEmp").Value)
            If .Fields("MR01").Value <> "" Then
                SavDay1 = SavDay1 + 1
            End If
            If Trim(.Fields("MRCNT").Value) = "1" Then
                SavDay2 = SavDay2 + 1
            End If
            If iPrint > 15000 Then
                Printer.NewPage
                Page = Page + 1
                Call PrintTitle("", strFCPEmp)
            End If
            PrintDatil
            .MoveNext
         Loop
       End With
       PrintEnd
       Printer.EndDoc
   End If
   CheckOC
End Sub

'Add By Sindy 2012/9/25 FCT一申請書多件(案件性質)案清單
'Modify By Sindy 2012/11/08 +m_CP28
Private Sub PrintFCTList(strCaseNo As String, m_CP09 As String, m_CP10 As String, m_CP27 As String, m_CP28 As String)
   Call GetTMCaseNo(strCaseNo)
   
   Printer.PaperSize = 9
   Printer.EndDoc
   Page = 0
   Call PrintTitle2(m_CP10)
   
   '來函案件
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04)
   
   '未來函案件
'   strSql = "SELECT c1.* " & _
'            "FROM caseprogress c1,trademark t1,(select tm01,tm02,tm03,tm04,tm23,tm78,tm79,tm80,tm81 from trademark where TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "') t2 " & _
'            "WHERE c1.CP01='" & m_TM01 & "' AND c1.CP27=" & m_CP27 & " AND c1.CP10='" & m_CP10 & "' " & _
'            "AND c1.CP01=t1.TM01(+) AND c1.CP02=t1.TM02(+) AND c1.CP03=t1.TM03(+) AND c1.CP04=t1.TM04(+) " & _
'            "AND t1.TM23||t1.TM78||t1.TM79||t1.TM80||t1.TM81=t2.TM23||t2.TM78||t2.TM79||t2.TM80||t2.TM81 " & _
'            "ORDER BY c1.CP01,c1.CP02,c1.CP03,c1.CP04 asc "
   'Modify By Sindy 2012/10/1
   'Modify By Sindy 2012/11/08 +m_CP28
   strSql = PUB_GetOneAppMuchCaseSql(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, m_CP27, m_CP28)
   '2012/10/1 End
   RsTemp.Close
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) <> RsTemp.Fields("cp01") & "-" & RsTemp.Fields("cp02") & IIf(RsTemp.Fields("cp03") & RsTemp.Fields("cp04") = "000", "", "-" & RsTemp.Fields("cp03") & "-" & RsTemp.Fields("cp04")) Then
            'Add By Sindy 2014/8/8
            If iPrint > 15000 Then
                Printer.NewPage
                Call PrintTitle2(m_CP10)
            End If
            '2014/8/8 END
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print RsTemp.Fields("cp01") & "-" & RsTemp.Fields("cp02") & IIf(RsTemp.Fields("cp03") & RsTemp.Fields("cp04") = "000", "", "-" & RsTemp.Fields("cp03") & "-" & RsTemp.Fields("cp04"))
            iPrint = iPrint + 300
         End If
         RsTemp.MoveNext
      Loop
   End If
   Printer.EndDoc
   
   Exit Sub
End Sub

'Add By Sindy 2014/8/8
Sub PrintTitle3()
   Page = Page + 1
   PLeft(0) = 1000
   PLeft(1) = 3000
   iPrint = 500
   Printer.Orientation = 1
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 14
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("主管機關來函收文簿") / 2)
   Printer.CurrentY = iPrint
   Printer.Print "主管機關來函收文簿"
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("分割子案清單") / 2)
   Printer.CurrentY = iPrint
   Printer.Print "分割子案清單"
   iPrint = iPrint + 500
   Printer.Font.Size = 9
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "報表對象：" & IIf(iPrintText = 1, str報表對象(1), str報表對象(0))
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人　：" & strUserName
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "製表日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "收件日期：" & ChangeTStringToTDateString(txtKeyIn(0).Text)
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "母案"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "子案"
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
End Sub

'Add By Sindy 2013/1/23 分割子案清單
Private Sub PrintDCList(strCaseNo As String)
   Call GetTMCaseNo(strCaseNo)
   
   Printer.PaperSize = 9
   Printer.EndDoc
   Page = 0
   Call PrintTitle3
   
   '母案
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04)
   
   '子案
   strSql = "SELECT DC01,DC02,DC03,DC04 From DivisionCase " & _
            "WHERE DC05='" & m_TM01 & "' AND DC06='" & m_TM02 & "' AND DC07='" & m_TM03 & "' AND DC08='" & m_TM04 & "'"
   RsTemp.Close
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         'Add By Sindy 2014/8/8
         If iPrint > 15000 Then
             Printer.NewPage
             Call PrintTitle3
         End If
         '2014/8/8 END
         Printer.CurrentX = PLeft(1)
         Printer.CurrentY = iPrint
         Printer.Print RsTemp.Fields("DC01") & "-" & RsTemp.Fields("DC02") & IIf(RsTemp.Fields("DC03") & RsTemp.Fields("DC04") = "000", "", "-" & RsTemp.Fields("DC03") & "-" & RsTemp.Fields("DC04"))
         iPrint = iPrint + 300
         RsTemp.MoveNext
      Loop
   End If
   Printer.EndDoc
   Exit Sub
End Sub

'Added by Morgan 2025/1/6
'Modified by Morgan 2025/2/14 +PName:程序人員
Sub PrintTitleP(pPName As String)
   Dim strTitle As String
   
   GetPleft2
   iPrint = 500
   Printer.Orientation = 1
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 14
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strTitle = "主管機關來函收文簿-P(" & pPName & ")"
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strTitle) / 2)
   Printer.CurrentY = iPrint
   Printer.Print strTitle
   iPrint = iPrint + 500
   Printer.Font.Size = 9
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "報表對象：" & str報表對象(1)
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人　：" & strUserName
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "製表日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "收件日期：" & ChangeTStringToTDateString(txtKeyIn(0).Text)
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "收件號"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "來函號數"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "智權區"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "起算日"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "期限"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "本所期限"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "政府機關"
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "備註"
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
End Sub

'Added by Morgan 2025/1/6
Private Sub GetPleft2()
   Erase PLeft
   
   PLeft(0) = 500  '收件號
   PLeft(1) = 1500 '來函號數
   PLeft(2) = 2800 '本所案號
   PLeft(3) = 4000 '智權區
   PLeft(4) = 5200 '智權人員
   PLeft(5) = 6300 '起算日
   PLeft(6) = 7200 '期限
   PLeft(7) = 8050 '本所期限
   PLeft(8) = 9200 '政府機關
   PLeft(9) = 10500 '備註
End Sub

'Added by Morgan 2025/1/7
'P案調整欄位，需另外更新智權區/同仁及改排序
Private Sub PrintDataP(strSql As String)
   Dim rsQuery As ADODB.Recordset
   Dim mSeqNo As String, stVTB As String
   Dim strPID As String, strPName As String 'Added by Morgan 2025/2/14
   
   '故意設定紙張屬性以便清除印表機狀態(相同印表機驅動程式會沿用原設定值,Ex.進紙槽)
   Printer.PaperSize = 9
   Printer.EndDoc
   CheckOC
   Page = 1
   SavDay1 = 0
   SavDay2 = 0
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   
      Set rsQuery = PUB_CreateRecordset(adoRecordset, , , 300, Me.Name, mSeqNo)
      With rsQuery
         .MoveFirst
         Do While Not .EOF
            .Fields("salno") = PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))
            .Fields("pid") = PUB_GetPHandler(.Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04"))
            .MoveNext
         Loop
         .UpdateBatch
         
         stVTB = "select R001 as " & .Fields(0).Name
         For intI = 2 To .Fields.Count
            stVTB = stVTB & ", R" & Format(intI, "000") & " as " & .Fields(intI - 1).Name
         Next
         stVTB = stVTB & " from Rdatafactory Where Id='" & strUserNum & "' And Formname='" & Me.Name & "'  And Seqno='" & mSeqNo & "'"
      End With
      'Modified by Morgan 2025/2/12 改依本所案號排序--郭(有請櫃台紙本依照本所號排)
      strSql = "Select X.*,st02 cp13,a0902 cp12 From (" & stVTB & ") X,staff,acc090" & _
         " Where st01(+)=salno and a0901(+)=st15 order by pid,pa01,pa02,pa03,pa04,mr01"
      intI = 1
      Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
      With adoRecordset
      If .RecordCount > 0 Then
           .MoveFirst
           strPID = "" & .Fields("PID")
           strPName = GetStaffName(strPID, True)
           PrintTitleP strPName
           Do While .EOF = False
               If strPID <> .Fields("PID") Then
                  PrintEnd
                  Printer.EndDoc
                  
                  Page = 1
                  SavDay1 = 0
                  SavDay2 = 0
                  strPID = "" & .Fields("PID")
                  strPName = GetStaffName(strPID, True)
                  PrintTitleP strPName
               End If
               
               strTemp(0) = CheckStr(.Fields("MR01").Value) '收件號
               strTemp(1) = CheckStr(.Fields("MR04").Value) '來函號數
               strTemp(2) = CheckStr(.Fields("MR13").Value) '本所案號
               strTemp(3) = CheckStr(.Fields("CP12").Value) '智權區
               strTemp(4) = CheckStr(.Fields("CP13").Value) '智權人員
               strTemp(5) = CheckStr(.Fields("MR05").Value) '起算日
               strTemp(6) = CheckStr(.Fields("MR06").Value) '期限
               strTemp(7) = CheckStr(.Fields("MR16").Value) '本所期限
               strTemp(8) = CheckStr(.Fields("MR10").Value) '政府機關
               strTemp(9) = CheckStr(.Fields("MR09").Value) '備註
               strTemp(9) = StrToStr(strTemp(9), 8)
               If .Fields("MR01").Value <> "" Then
                   SavDay1 = SavDay1 + 1
               End If
               If Trim(.Fields("MRCNT").Value) = "1" Then
                   SavDay2 = SavDay2 + 1
               End If
               If iPrint > 15000 Then
                   Printer.NewPage
                   Page = Page + 1
                   PrintTitleP strPName
               End If
               PrintDatil
               .MoveNext
           Loop
       End If
       End With
       PrintEnd
       Printer.EndDoc
   End If
   CheckOC
   Set rsQuery = Nothing
End Sub
