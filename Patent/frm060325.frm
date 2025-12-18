VERSION 5.00
Begin VB.Form frm060325 
   BorderStyle     =   1  '單線固定
   Caption         =   "實審通知函"
   ClientHeight    =   4560
   ClientLeft      =   4380
   ClientTop       =   2736
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4260
   Begin VB.Frame Frame2 
      Caption         =   "設定請款單及定稿"
      Height          =   660
      Left            =   285
      TabIndex        =   25
      Top             =   3810
      Width           =   3660
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   26
         Top             =   240
         Width           =   2745
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   27
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1935
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2430
      Width           =   255
   End
   Begin VB.TextBox txtCustomer 
      Height          =   264
      Index           =   2
      Left            =   2805
      MaxLength       =   9
      TabIndex        =   10
      Top             =   1710
      Width           =   1215
   End
   Begin VB.TextBox txtCustomer 
      Height          =   264
      Index           =   1
      Left            =   1470
      MaxLength       =   9
      TabIndex        =   9
      Top             =   1710
      Width           =   1215
   End
   Begin VB.TextBox txtControler 
      Height          =   264
      Left            =   1470
      MaxLength       =   6
      TabIndex        =   11
      Top             =   2040
      Width           =   1080
   End
   Begin VB.TextBox txtAgent 
      Height          =   264
      Index           =   1
      Left            =   1470
      MaxLength       =   9
      TabIndex        =   7
      Top             =   1380
      Width           =   1215
   End
   Begin VB.TextBox txtAgent 
      Height          =   264
      Index           =   2
      Left            =   2805
      MaxLength       =   9
      TabIndex        =   8
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   660
      Left            =   300
      TabIndex        =   18
      Top             =   3060
      Width           =   3660
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   14
         Top             =   240
         Width           =   2745
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   2
         Left            =   105
         TabIndex        =   19
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox txtDate 
      Height          =   264
      Index           =   2
      Left            =   2565
      MaxLength       =   7
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   264
      Index           =   4
      Left            =   3030
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1050
      Width           =   375
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   264
      Index           =   3
      Left            =   2790
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1050
      Width           =   255
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   264
      Index           =   2
      Left            =   1950
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1050
      Width           =   855
   End
   Begin VB.TextBox txtCaseNo 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "FCP"
      Top             =   1050
      Width           =   495
   End
   Begin VB.TextBox txtDate 
      Height          =   264
      Index           =   1
      Left            =   1470
      MaxLength       =   7
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   300
      TabIndex        =   2
      Top             =   1095
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "實審期限："
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   17
      Top             =   765
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3390
      TabIndex        =   16
      Top             =   120
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2595
      TabIndex        =   15
      Top             =   120
      Width           =   756
   End
   Begin VB.CheckBox Check1 
      Caption         =   "只列印承辦單"
      Height          =   225
      Left            =   60
      TabIndex        =   13
      Top             =   2910
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label6 
      Caption         =   "※選實審期限條件於結束　時會一併列印同區間內　寰華案清單"
      ForeColor       =   &H00FF00FF&
      Height          =   555
      Left            =   270
      TabIndex        =   28
      Top             =   90
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否加印 PS 內容：　　(Y)"
      Height          =   180
      Index           =   15
      Left            =   300
      TabIndex        =   24
      Top             =   2460
      Width           =   2130
   End
   Begin VB.Line Line2 
      X1              =   2445
      X2              =   3105
      Y1              =   1845
      Y2              =   1845
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   23
      Top             =   1755
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "管制人："
      Height          =   180
      Index           =   5
      Left            =   300
      TabIndex        =   22
      Top             =   2085
      Width           =   1200
   End
   Begin VB.Label lblControler 
      Height          =   180
      Left            =   2580
      TabIndex        =   21
      Top             =   2085
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "代理人："
      Height          =   180
      Index           =   4
      Left            =   300
      TabIndex        =   20
      Top             =   1425
      Width           =   1200
   End
   Begin VB.Line Line3 
      X1              =   2445
      X2              =   3105
      Y1              =   1515
      Y2              =   1515
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2460
      X2              =   2580
      Y1              =   855
      Y2              =   855
   End
End
Attribute VB_Name = "frm060325"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim strPrinter2 As String 'Add By Sindy 2015/7/3
Dim m_PrintRpt1 As Boolean, ff1 As Integer, m_strFileName1 As String 'Add By Sindy 2015/10/30


Private Function CheckConstrain() As Boolean
   '實審期限
   If Option1(0).Value = True Then
      If txtDate(1) = "" Then
         MsgBox "請輸入實審期限起日!!!", vbExclamation + vbOKOnly
         txtDate(1).SetFocus
      ElseIf txtDate(2) = "" Then
         MsgBox "請輸入實審期限迄日!!!", vbExclamation + vbOKOnly
         txtDate(2).SetFocus
      ElseIf Val(txtDate(1)) > Val(txtDate(2)) Then
         MsgBox "實審期限起日不可大於迄日!!!", vbExclamation + vbOKOnly
         txtDate(1).SetFocus
      Else
         CheckConstrain = True
      End If
      
   '本所案號
   Else
      txtCaseNo(3) = txtCaseNo(3) & "0"
      txtCaseNo(4) = txtCaseNo(4) & "00"
      If txtCaseNo(2) = "" Then
         MsgBox "請輸入案號!!!", vbExclamation, "USER 輸入錯誤"
         txtCaseNo(2).SetFocus
      ElseIf Len(txtCaseNo(2)) <> 6 Then
         MsgBox "請輸入完整案號!!!", vbExclamation, "USER 輸入錯誤"
         txtCaseNo(2).SetFocus
         txtCaseNo_GotFocus 2
      Else
         CheckConstrain = True
      End If
   End If
   If CheckConstrain = True Then
      If txtAgent(1) <> "" Or txtAgent(2) <> "" Then
         If Left(txtAgent(1), 6) <> Left(txtAgent(2), 6) Then
            CheckConstrain = False
            MsgBox "代理人前六碼必須相同!!", vbExclamation, "USER 輸入錯誤"
            txtAgent(1).SetFocus
            txtAgent_GotFocus 1
         End If
      End If
   End If
   If CheckConstrain = True Then
      If txtCustomer(1) <> "" Or txtCustomer(2) <> "" Then
         If Left(txtCustomer(1), 6) <> Left(txtCustomer(2), 6) Then
            CheckConstrain = False
            MsgBox "申請人前六碼必須相同!!", vbExclamation, "USER 輸入錯誤"
            txtCustomer(1).SetFocus
            txtCustomer_GotFocus 1
         End If
      End If
   End If
   
   If CheckConstrain = True Then
      If txtControler <> "" And lblControler = "" Then
            CheckConstrain = False
            MsgBox "管制人輸入錯誤!!", vbExclamation, "USER 輸入錯誤"
      End If
   End If
End Function

Private Function ReadData(ByRef p_adoRst As ADODB.Recordset, ByVal p_stSQL As String, Optional ByVal p_bolMsg As Boolean = True) As Boolean

On Err GoTo ErrHnd

   If p_adoRst.State = adStateOpen Then p_adoRst.Close
   With p_adoRst
      .CursorLocation = adUseClient
      .Open p_stSQL, cnnConnection, adOpenStatic, adLockReadOnly
      '若有資料
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/8
         ReadData = True
      ElseIf p_bolMsg Then
         InsertQueryLog (0) 'Add By Sindy 2010/12/8
         MsgBox "無符合條件之資料！", vbInformation
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub Process()
   Dim stCon As String, stConX As String, stLanguage As String, stReceiveNo As String, strSitu As String, stET01 As String
   Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
   Dim dblExRate As Double
   Dim strBillNo As String '待印請款單號 Add by Morgan 2011/6/23
   Dim idx As Integer
   Dim bolDNEmail As Boolean, bolDNPlusPaper As Boolean 'Added by Morgan 2014/6/3
   'Add By Sindy 2015/10/30
   Dim strLD03 As String
   Dim strFileName As String, strFullFileName As String
   Dim oFileSys As New FileSystemObject
   Dim oFile As File
   Dim strMsg As String
   Dim strNewCP09 As String
   '2015/10/30 END
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/8 清除查詢印表記錄檔欄位
   If CheckConstrain = True Then
      If Option1(0).Value = True Then
         stCon = stCon & " AND NP09 BETWEEN " & TransDate(txtDate(1), 2) & " AND " & TransDate(txtDate(2).Text, 2)
         pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & txtDate(1) & "-" & txtDate(2) 'Add By Sindy 2010/12/8
      Else
         stCon = stCon & " AND NP03='" & txtCaseNo(2) & "' AND NP04='" & txtCaseNo(3) & "' AND NP05='" & txtCaseNo(4) & "'"
         pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & txtCaseNo(1) & "-" & txtCaseNo(2) & "-" & txtCaseNo(3) & "-" & txtCaseNo(4) 'Add By Sindy 2010/12/8
      End If
      
      '代理人
      If txtAgent(1) <> "" Then
         stCon = stCon & " AND PA75>='" & txtAgent(1) & "' AND PA75<='" & txtAgent(2) & "'"
         pub_QL05 = pub_QL05 & ";" & Label1(4) & txtAgent(1) & "-" & txtAgent(2) 'Add By Sindy 2010/12/8
      End If
      '申請人
      If txtCustomer(1) <> "" Then
         stCon = stCon & " AND ((PA26>='" & txtCustomer(1) & "' AND PA26<='" & txtCustomer(2) & "')" & _
            " OR (PA27>='" & txtCustomer(1) & "' AND PA27<='" & txtCustomer(2) & "')" & _
            " OR (PA28>='" & txtCustomer(1) & "' AND PA28<='" & txtCustomer(2) & "')" & _
            " OR (PA29>='" & txtCustomer(1) & "' AND PA29<='" & txtCustomer(2) & "')" & _
            " OR (PA30>='" & txtCustomer(1) & "' AND PA30<='" & txtCustomer(2) & "'))"
         pub_QL05 = pub_QL05 & ";" & Label1(0) & txtCustomer(1) & "-" & txtCustomer(2) 'Add By Sindy 2010/12/8
      End If
      
      If txtControler <> "" Then
         stConX = stConX & " AND EXISTS( SELECT * FROM NATION WHERE NA01=X1 AND NA16='" & txtControler & "')"
         pub_QL05 = pub_QL05 & ";" & Label1(5) & txtControler & lblControler 'Add By Sindy 2010/12/8
      End If
      
      If Text1 = "Y" Then
         pub_QL05 = pub_QL05 & ";" & Left(Label1(15), 11) & Text1 'Add By Sindy 2010/12/8
      End If
      
      'Modify By Sindy 2015/7/6 +,GetEmailFlag(np02||np03||np04||np05) eMail
      'Modify By Sindy 2021/4/27 + ,np23
      strSql = "SELECT * FROM (" & _
         " SELECT np02,np03,np04,np05,np08,np09,NVL(NVL(PA85,FA31),CU64) X0,NVL(FA10,CU10) X1,PA75,pa142,fa86,cu124,GetEmailFlag(np02||np03||np04||np05) eMail,np01,np22,np23" & _
         " From Nextprogress,Patent,FAGENT,CUSTOMER" & _
         " WHERE NP02||''='FCP' and NP07=" & 實體審查 & " AND NP06 IS NULL" & stCon & _
         " AND PA01(+)=NP02 AND PA02(+)=NP03 AND PA03(+)=NP04 AND PA04(+)=NP05 AND PA57 IS NULL" & _
         " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9,1)" & _
         " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)" & _
         " ) X WHERE 1=1" & stConX & " ORDER BY eMail,np02,np03,np04,np05"
      Me.Enabled = False
      If ReadData(adoRecordset1, strSql) = True Then
         Set g_WordAp = Nothing 'Added by Morgan 2015/6/3 若有系統開啟的Word沒結束時,定稿都會開啟,故物件要先釋放
         'Add by Morgan 2011/7/8
         pub_OsPrinter = PUB_GetOsDefaultPrinter
         PUB_SetOsDefaultPrinter Combo2.Text
         PUB_SetWordActivePrinter
         'end 2011/7/8
         PUB_RestorePrinter Combo2.Text 'Add By Sindy 2015/7/3
         
         dblExRate = PUB_GetUSXRate 'Add by Morgan 2009/4/27 美金匯率
         With adoRecordset1
            stET01 = "08"
            Do While Not .EOF
'               'Add By Sindy 2015/7/3 只列印承辦單
'               If Check1.Value = 1 Then
'                  Call PUB_PrintFCPEmpBill(.Fields(0), .Fields(1), .Fields(2), .Fields(3), stET01, , 實體審查, "" & .Fields("NP09"))
'               Else

                  'Add By Sindy 2015/11/4 個案才要列印承辦單
                  If Option1(1).Value = True Then
                  '2015/11/4 END
                     '列印FCP承辦單
                     'Modified by Lydia 2019/03/04 更換類別代號;
                     'Call PUB_PrintFCPEmpBill(.Fields(0), .Fields(1), .Fields(2), .Fields(3), stET01, , 實體審查, "" & .Fields("NP09"))
                     Call PUB_PrintFCPEmpBill(.Fields(0), .Fields(1), .Fields(2), .Fields(3), "04", , 實體審查, "" & .Fields("NP09"))
                  End If
                  
                  '定稿語文
                  'Modify by Morgan 2006/6/7 改Call公用函數
                  'stLanguage = "" & .Fields("X0")
                  stLanguage = PUB_GetLanguage("" & .Fields("NP02"), "" & .Fields("NP03"), "" & .Fields("NP04"), "" & .Fields("NP05"))
                  If stLanguage = "3" Then
                     'Modified by Morgan 2014/12/23 刪除舊定稿
                     'strSitu = "04"
                     'If .Fields("NP09") >= 20130101 Then strSitu = "05" 'Added by Morgan 2012/7/9 期限>=102/1/1適用99新法定稿
                     strSitu = "05"
                     'end 2014/12/23
                  'Added by Morgan 2016/6/21
                  '中文
                  ElseIf stLanguage = "1" Then
                     strSitu = "04"
                  'end 2016/6/21
                  Else
                     'Modified by Morgan 2014/12/23 刪除舊定稿
                     'strSitu = "02"
                     'If .Fields("NP09") >= 20130101 Then strSitu = "03" 'Added by Morgan 2012/7/9 期限>=102/1/1適用99新法定稿
                     strSitu = "03"
                     'end 2014/12/23
                  End If
                  
                  stReceiveNo = "" & .Fields("NP02") & .Fields("NP03") & .Fields("NP04") & .Fields("NP05") & "&" & 實體審查
                  EndLetter stET01, stReceiveNo, strSitu, strUserNum
                  
                  idx = 1
                  'Modify By Sindy 2021/4/28
                  If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
                     strExc(idx) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & stET01 & "','" & stReceiveNo & "','" & strSitu & "','" & strUserNum & "','約定期限'," & CNULL("" & .Fields("NP23")) & ")"
                  Else
                  '2021/4/28 END
                     strExc(idx) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & stET01 & "','" & stReceiveNo & "','" & strSitu & "','" & strUserNum & "','本所期限'," & CNULL("" & .Fields("NP08")) & ")"
                  End If
                  
                  idx = idx + 1
                  strExc(idx) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & stET01 & "','" & stReceiveNo & "','" & strSitu & "','" & strUserNum & "','法定期限'," & CNULL("" & .Fields("NP09")) & ")"
                  
                  '2008/11/10 ADD BY SONIA
                  If Text1 = "Y" Then
                     idx = idx + 1
                     strExc(idx) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                              "('" & stET01 & "','" & stReceiveNo & "','" & strSitu & "','" & strUserNum & "','列印備註','p.s. Previously we have received your letter not to take further action without your specific instructions. If you want to maintain this case, please notify us immediately before the deadline.')"
                  End If
                  '2008/11/10 END
                  'Add by Morgan 2009/4/27
                  idx = idx + 1
                  'Added by Morgan 2012/7/9
                  If .Fields("NP09") >= 20130101 Then
                        strExc(idx) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & stET01 & "','" & stReceiveNo & "','" & strSitu & "','" & strUserNum & "','美金費用','" & Format(Fix(11000 / dblExRate), "##0") & "')"
                  
                        idx = idx + 1
                        strExc(idx) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & stET01 & "','" & stReceiveNo & "','" & strSitu & "','" & strUserNum & "','美金超項費','" & Format(Fix(800 / dblExRate), "##0") & "')"
                  Else
                  'end 2012/7/9
                     strExc(idx) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & stET01 & "','" & stReceiveNo & "','" & strSitu & "','" & strUserNum & "','美金費用','" & Format(Fix(12000 / dblExRate), "##0") & "')"
                  End If 'Added by Morgan 2012/7/9
                  
                  idx = idx + 1
                  strExc(idx) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & stET01 & "','" & stReceiveNo & "','" & strSitu & "','" & strUserNum & "','美金超頁費','" & Format(Fix(500 / dblExRate), "##0") & "')"
                  
                  'Add by Morgan 2011/6/23
                  If PUB_GetUnPaidBill(.Fields("NP02"), .Fields("NP03"), .Fields("NP04"), .Fields("NP05"), strBillNo) = True Then
                     idx = idx + 1
                     strExc(idx) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & stET01 & "','" & stReceiveNo & "','" & strSitu & "','" & strUserNum & "','有欠款才印','♀')"
                     idx = idx + 1
                     strExc(idx) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & stET01 & "','" & stReceiveNo & "','" & strSitu & "','" & strUserNum & "','有欠款不印','♀')"
                  End If
                  
                  'edit by nickc 2007/02/05 不用 dll 了
                  'If Not objLawDll.ExecSQL(2, strExc) Then
                  If Not ClsLawExecSQL(idx, strExc) Then
                     MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                     GoTo RunEnd
                     Exit Sub
                  End If
                  
                  'Add By Sindy 2015/10/6
                  'Modified by Lydia 2025/06/05 改成「程序大項工作整批發文」，不上發文日；+, , True, , , Me.Name
                  If PUB_AddCP1913(.Fields(0), .Fields(1), .Fields(2), .Fields(3), .Fields("np08"), .Fields("np09"), .Fields("np01"), .Fields("np22"), , , strNewCP09, , True, , , Me.Name) = False Then
                     MsgBox .Fields(0) & "-" & .Fields(1) & "-" & .Fields(2) & "-" & .Fields(3) & "新增進度檔【通知期限】失敗！作業中斷！", vbCritical
                     GoTo RunEnd
                     Exit Sub
                  End If
                  '2015/10/6 END
                  
                  'Modify by Morgan 2008/3/20 判斷是否產生電子檔
                  'NowPrint stReceiveNo, stET01, strSitu, False, strUserNum, 0
                  'Modify by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
                  'bolEmail = IsNull(.Fields("pa142") & .Fields("fa86") & .Fields("cu124"))

                  bolEmail = PUB_GetEMailFlag(.Fields(0) & .Fields(1) & .Fields(2) & .Fields(3), , , bolPlusPaper)
                  'Added by Morgan 2014/6/3
                  If bolEmail = False Then
                     bolDNEmail = PUB_GetEMailFlag(.Fields(0) & .Fields(1) & .Fields(2) & .Fields(3), , , bolDNPlusPaper, , True)
                  Else
                     bolDNEmail = bolEmail
                     bolDNPlusPaper = bolPlusPaper
                  End If
                  'end 2014/6/3
                     
                  If bolPlusPaper Then
                     iCopy = 0
                  Else
                     iCopy = 1
                  End If
                  'end 2009/10/20

                  If bolEmail Then
                     NowPrint stReceiveNo, stET01, strSitu, False, strUserNum, 0, , , , iCopy, , True, True
                     'Modify By Sindy 2015/10/30 Mark
'                     '若跑單筆時顯示訊息
'                     If Option1(1).Value = True Then
'                        MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(.Fields(0)) & " ]！"
'                     End If
                  Else
                     NowPrint stReceiveNo, stET01, strSitu, False, strUserNum, 0
                  End If
                  'end 2008/3/20
                  
                  'Add by Morgan 2011/6/23
                  '列印請款單
                  If strBillNo <> "" Then
                     'Modified by Morgan 2014/6/3
                     'PUB_PrintBill strBillNo, Combo2.Text, bolEmail, bolPlusPaper, Me.Name, , 1
                     PUB_PrintBill strBillNo, Combo2.Text, bolDNEmail, bolDNPlusPaper, Me.Name, , 1
                  End If
                  '列印通知函
                  'PUB_PrintLetter stReceiveNo, , , strLD03
                  'end 2011/6/23
                  'Modify By Sindy 2015/10/30 定稿轉PDF存卷宗區
                  strFileName = .Fields(0) & .Fields(1) & IIf(.Fields(3) <> "00", "-" & .Fields(2) & "-" & .Fields(3), IIf(.Fields(2) <> "0", "-" & .Fields(2), "")) & ".1913.CUS.PDF"
                  PUB_DelFtpFile2 strNewCP09, " and cpp02='" & strFileName & "'" '檔案改放 FTP,必須在DB資料刪除前執行
                  strSql = "delete from CasePaperPDF where cpp01='" & strNewCP09 & "' and cpp02='" & strFileName & "'"
                  cnnConnection.Execute strSql
                  If PUB_PrintLetter(stReceiveNo, , , True, strFullFileName) = True Then
                     Call PUB_ChkFileStatus(strFullFileName, False, strMsg)  'Added by Lydia 2022/10/31 判斷檔案是否存在, 超過時間就繼續;
                     Set oFile = oFileSys.GetFile(strFullFileName)
                     If SaveAttFile_PDF(strNewCP09, strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) = False Then
                        'Modified by Lydia 2022/10/31 +& ";" & strMsg
                        Call ReadTxt1(.Fields(0) & "-" & .Fields(1) & "-" & .Fields(2) & "-" & .Fields(3), strNewCP09, "定稿轉PDF失敗" & ";" & strMsg)
                     End If
                     Kill strFullFileName
                  End If
                  '2015/10/30 END
                  
                  'Modify by Morgan 2008/3/20 產生電子檔時不印地址條
                  If Not bolEmail Or bolPlusPaper Then
'                     'Add By Sindy 2015/9/21 日文定稿才要印地址條
'                     If stLanguage = "3" Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'                     '2015/9/21 END
                        '新增地址條列表資料
                        pub_AddressListSN = pub_AddressListSN + 1
                        PUB_AddNewAddressList strUserNum, "" & .Fields(0).Value, "" & .Fields(1).Value, "" & .Fields(2).Value, "" & .Fields(3).Value, "" & pub_AddressListSN, "0", 實體審查
'                     End If
                  End If

                  If Option1(0).Value = True Then
                     '新增整批定稿列印清單資料
                     PUB_AddNewLetterList "催實審通知函", txtDate(1) & "-" & txtDate(2), "" & .Fields(0).Value, "" & .Fields(1).Value, "" & .Fields(2).Value, "" & .Fields(3).Value, IIf(bolEmail, IIf(bolPlusPaper, "Ｅ", "ｅ"), "")
                  End If
'               End If
               .MoveNext
            Loop
         End With

'Removed by Morgan 2018/10/18 移到後面(沒有FCP案還是要跑P案)
'         'Modified by Morgan 2016/1/12 整批才要
'         If Option1(0).Value = True And txtAgent(1) = "" And txtCustomer(1) = "" Then
'            AddPLetterList 'Added by Morgan 2015/10/20
'         End If
'end 2018/10/18
         
         'Modifhy by Morgan 2011/6/23
         'MsgBox "定稿產生完成！", vbInformation
         PUB_SetOsDefaultPrinter pub_OsPrinter
         PUB_RestorePrinter strPrinter2 'Add By Sindy 2015/7/3
         
         'Modify By Sindy 2015/10/30
         If m_PrintRpt1 = True Then
            Close ff1
            strMsg = "請至下列位置列印檢核表：" & PUB_Getdesktop & "\" & m_strFileName1
         End If
         MsgBox "定稿列印完畢！ " & strMsg, vbInformation
         '2015/10/30 END
         'end 2011/6/23
      End If
   End If
   
   'Modified by Morgan 2016/1/12 整批才要
   If Option1(0).Value = True And txtAgent(1) = "" And txtCustomer(1) = "" Then
      AddPLetterList 'Added by Morgan 2015/10/20
   End If
   
   Me.Enabled = True
   Exit Sub
   
RunEnd:
   PUB_SetOsDefaultPrinter pub_OsPrinter
   PUB_RestorePrinter strPrinter2
End Sub

'Add By Sindy 2015/10/30
'資料檢核表
Private Sub ReadTxt1(strCaseNo As String, strRecvNo As String, strNote As String)
Dim i As Integer
Dim strTemp(1 To 7) As String
   
   If m_PrintRpt1 = False Then
      m_PrintRpt1 = True
      If ff1 > 0 Then Close #ff1
      ff1 = FreeFile
      m_strFileName1 = Me.Caption & txtDate(1) & "-" & txtDate(2) & "資料檢核表.txt"
      Open PUB_Getdesktop & "\" & m_strFileName1 For Output As ff1
      'Print #ff1, "備註：改字型Fixedsys標準11號字以橫式上下左右各10MM列印"
      Print #ff1, "本所案號        總收文號   原因"
      Print #ff1, "=============== ========== ============================================="
   End If
   For i = 1 To 3
      strTemp(i) = ""
   Next i
   strTemp(1) = convForm(CheckStr(Trim(strCaseNo)), 15)
   strTemp(2) = convForm(CheckStr(Trim(strRecvNo)), 10)
   strTemp(3) = Trim(strNote)
   Print #ff1, strTemp(1) & " " & strTemp(2) & " " & strTemp(3)
End Sub

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0 '確定
         Screen.MousePointer = vbHourglass
         Process
         Screen.MousePointer = vbDefault
      Case 1 '結束
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, Combo1
   txtControler = strUserNum
   
   'Add by Morgan 2011/6/23
   PUB_SetPrinter Me.Name, Combo2, strPrinter2
   'MsgBox "本程式已改為直接列印定稿，請先選定印表機並放好定稿紙！", vbExclamation
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '列印定稿整批列印清單
   PUB_PrintLetterList strUserNum, "6", Combo2, strPrinter2
   'Remove by Lydia 2019/05/29 FCP案和寰華案的整批清單併在一起
   'PUB_PrintLetterList strUserNum, "10", Combo2, strPrinter2, , False 'Added by Morgan 2015/10/20
   '刪除定稿整批列印資料
   'Modified by Lydia +傳入刪除條件
   'PUB_DeleteLetterList strUserNum
   PUB_DeleteLetterList strUserNum, " and LL02='催實審通知函' "
   '列印地址條
   PUB_PrintAddressList strUserNum, Me.Combo1.Text
   '刪除地址條列表資料
   PUB_DeleteAddressList strUserNum
   '初始化序號
   pub_AddressListSN = 0
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   'Add by Morgan 2011/6/23
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   
   Set frm060325 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
         txtDate(1).Enabled = True
         txtDate(2).Enabled = True
         txtCaseNo(1).Enabled = False
         txtCaseNo(2).Enabled = False: txtCaseNo(2) = ""
         txtCaseNo(3).Enabled = False: txtCaseNo(3) = ""
         txtCaseNo(4).Enabled = False: txtCaseNo(4) = ""
         txtDate(1).SetFocus
      Case 1
         txtDate(1).Enabled = False: txtDate(1) = ""
         txtDate(2).Enabled = False: txtDate(2) = ""
         txtCaseNo(1).Enabled = True: txtCaseNo(1) = "FCP"
         txtCaseNo(2).Enabled = True
         txtCaseNo(3).Enabled = True
         txtCaseNo(4).Enabled = True
         txtCaseNo(2).SetFocus
   End Select
End Sub

'2008/11/10 ADD BY SONIA
Private Sub Text1_GotFocus()
   TextInverse Me.Text1()
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      MsgBox "是否加印 PS 內容只能輸入 Y 或空白 !!!", vbExclamation + vbOKOnly
      KeyAscii = 0
   End If
End Sub
'2008/11/10 END
Private Sub txtAgent_GotFocus(Index As Integer)
   If Index = 2 And Len(txtAgent(1)) = 9 Then
      txtAgent(2) = txtAgent(1)
   End If
   TextInverse txtAgent(Index)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtAgent(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtAgent_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtAgent_Validate(Index As Integer, Cancel As Boolean)
   
   'If Len(txtAgent(Index)) = 6 Then 'Mark by Lydia 2021/03/11 因為催北京銀龍案件，程序會輸入Y5133301
   If Trim(txtAgent(Index)) <> "" And Len(txtAgent(Index)) < 9 Then 'Added by Lydia 2021/11/10 debug: 自動+000,造成FCP-058894沒有催到期限
      'Modified by Lydia 2021/11/10 補到9碼 "000"=> String(9, "0")
      txtAgent(Index) = Left(txtAgent(Index) & String(9, "0"), 9)
   End If 'Added by Lydia 2021/11/10
   'End If 'Mark by Lydia 2021/03/11
    
End Sub

Private Sub txtCaseNo_GotFocus(Index As Integer)
   TextInverse txtCaseNo(Index)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtCaseNo(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtCaseNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtControler_Change()
   If Len(txtControler) > 4 Then
      lblControler = StaffQuery(txtControler)
   Else
      lblControler = ""
   End If
End Sub

Private Sub txtControler_GotFocus()
   TextInverse txtControler
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtControler.IMEMode = 2
   CloseIme
End Sub

Private Sub txtControler_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCustomer_GotFocus(Index As Integer)
   If Index = 2 And Len(txtCustomer(1)) = 9 Then
      txtCustomer(2) = txtCustomer(1)
   End If
   TextInverse txtCustomer(Index)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtCustomer(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtCustomer_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCustomer_Validate(Index As Integer, Cancel As Boolean)
   'If Len(txtCustomer(Index)) = 6 Then 'Mark by Lydia 2021/03/11 因為催北京銀龍案件，程序會輸入Y5133301
   If Trim(txtCustomer(Index)) <> "" And Len(txtCustomer(Index)) < 9 Then  'Added by Lydia 2021/11/10 debug: 自動+000,造成FCP-058894沒有催到期限
      'Modified by Lydia 2021/11/10 補到9碼 "000"=> String(9, "0")
      txtCustomer(Index) = Left(txtCustomer(Index) & String(9, "0"), 9)
   End If 'Addec gy Lydia 2021/11/10
   'End If 'Mark by Lydia 2021/03/11
End Sub

Private Sub txtDate_GotFocus(Index As Integer)
   TextInverse txtDate(Index)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtDate(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtDate_Validate(Index As Integer, Cancel As Boolean)
   If txtDate(Index) <> "" Then
      If ChkDate(txtDate(Index)) = False Then
         Cancel = True
         txtDate(Index).SetFocus
         txtDate_GotFocus Index
      End If
   End If
End Sub

'Added by Morgan 2015/10/20
Private Sub AddPLetterList()
   Dim stCon As String
   
   On Error GoTo ErrHnd
      
   'Modified by Lydia 2015/10/26 寰華案定義改為:新案發文人員為外專程序者
'   strSql = "Insert Into LetterList (LL01,LL02,LL03,LL04,LL05,LL06,LL07,LL08)" & _
'      " select '" & strUserNum & "','寰華案實審通知函','" & txtDate(1).Text & "-" & txtDate(2).Text & "'" & _
'      ",np02,np03,np04,np05,nvl(pa75,pa26) LL08 from nextprogress A, caseprogress, patent " & _
'      " WHERE NP09 BETWEEN " & DBDATE(txtDate(1)) & " AND " & DBDATE(txtDate(2)) & " AND NP02||NP07||NP06='P416'" & _
'      " and cp01(+)=np02 and cp02(+)=np03 and cp03(+)=np04 and cp04(+)=np05 and cp31='Y' and cp12 like 'F%' and CP44='Y53374000'" & _
'      " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) " & _
'      " and not exists" & _
'      " (select * from nextprogress B where B.NP02=A.NP02 AND B.NP03=A.NP03 AND B.NP04=A.NP04 AND B.NP05=A.NP05 AND B.NP06='N' AND B.NP07=A.NP07" & _
'      " AND TO_NUMBER(TO_CHAR(ADD_MONTHS(TO_DATE(B.NP09,'YYYYMMDD'),9),'YYYYMMDD'))>A.NP09)"
   
   'Modified by Morgan 2015/11/10
   If txtControler <> "" Then
      'Modified by Lydia 2017/02/13 +FMP管制人
      If strSrvDate(1) < FMP管制人啟用日 Then
        stCon = " AND EXISTS( SELECT * FROM NATION WHERE NA01=NVL(FA10,CU10) AND NA16='" & txtControler & "')"
      Else
        stCon = " AND EXISTS( SELECT * FROM NATION WHERE NA01=NVL(FA10,CU10) AND NVL(NA79,NA16)='" & txtControler & "')"
      End If
      'end 2017/02/13
   End If
      
  'Modified by Lydia 2019/05/09 +PK: 使用者帳號@電腦名稱(pub_HostName)
  'Modified by Lydia 2019/05/29 FCP案和寰華案的整批清單併在一起( 寰華案實審通知函 => 催實審通知函)
  'Modified by Morgan 2022/1/10 要排除已閉卷案件 Ex:P-124626 (原來P案催實審就有排除)
  'Modified by Morgan 2025/10/28 實審沒有管制半年的狀況，刪除延期的判斷(應該是複製年費的語法多的)
   strSql = "Insert Into LetterList (LL01,LL02,LL03,LL04,LL05,LL06,LL07,LL08)" & _
      " select '" & strUserNum & "@" & pub_HostName & "','催實審通知函','" & txtDate(1).Text & "-" & txtDate(2).Text & "'" & _
      ",np02,np03,np04,np05,nvl(pa75,pa26) LL08 from nextprogress A, caseprogress, patent,staff,FAGENT,CUSTOMER " & _
      " WHERE NP09 BETWEEN " & DBDATE(txtDate(1)) & " AND " & DBDATE(txtDate(2)) & " AND NP02||NP07||NP06='P416'" & _
      " and cp01(+)=np02 and cp02(+)=np03 and cp03(+)=np04 and cp04(+)=np05 and cp31='Y' and cp12 like 'F%' and CP44='Y53374000'" & _
      " and a.np02=pa01(+) and a.np03=pa02(+) and a.np04=pa03(+) and a.np05=pa04(+) and pa57 is null" & _
      " and cp83=st01(+) and (st03='F22' or st01 is null) " & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9)" & _
      " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9)" & stCon
   cnnConnection.Execute strSql, intI
   Exit Sub
      
ErrHnd:
      MsgBox Err.Description, vbCritical, "寰華案清單產生失敗!!"
End Sub

