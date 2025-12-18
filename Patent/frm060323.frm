VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060323 
   BorderStyle     =   1  '單線固定
   Caption         =   "公開通知函"
   ClientHeight    =   6000
   ClientLeft      =   2796
   ClientTop       =   3948
   ClientWidth     =   6324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6324
   Begin VB.TextBox txtData 
      Height          =   280
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Top             =   840
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   345
      Left            =   5970
      TabIndex        =   13
      Top             =   2400
      Width           =   345
   End
   Begin VB.CheckBox Check2 
      Caption         =   "不印公報"
      Height          =   345
      Left            =   2520
      TabIndex        =   11
      Top             =   1980
      Width           =   1365
   End
   Begin VB.ListBox List1 
      Height          =   1488
      ItemData        =   "frm060323.frx":0000
      Left            =   75
      List            =   "frm060323.frx":0007
      TabIndex        =   20
      Top             =   4200
      Width           =   6195
   End
   Begin VB.ComboBox cmbPrinter2 
      Height          =   300
      Left            =   1845
      TabIndex        =   14
      Top             =   2760
      Width           =   4395
   End
   Begin VB.TextBox txtPDFPath 
      Height          =   315
      Left            =   1845
      TabIndex        =   15
      Text            =   "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Top             =   3120
      Width           =   4395
   End
   Begin VB.TextBox txtPath2 
      Height          =   315
      Left            =   1845
      TabIndex        =   12
      Top             =   2400
      Width           =   4125
   End
   Begin VB.CheckBox Check1 
      Caption         =   "只列印定稿清單"
      Height          =   345
      Left            =   210
      TabIndex        =   9
      Top             =   1980
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   660
      Left            =   60
      TabIndex        =   18
      Top             =   1230
      Width           =   3435
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   8
         Top             =   240
         Width           =   2520
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   19
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   4
      Left            =   2910
      MaxLength       =   2
      TabIndex        =   6
      Top             =   492
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   3
      Left            =   2655
      MaxLength       =   1
      TabIndex        =   5
      Top             =   492
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   1815
      MaxLength       =   6
      TabIndex        =   4
      Top             =   492
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   1335
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "FCP"
      Top             =   492
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   0
      Left            =   1335
      MaxLength       =   7
      TabIndex        =   1
      Top             =   156
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "公開日："
      Height          =   180
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   204
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5430
      TabIndex        =   17
      Top             =   120
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4608
      TabIndex        =   16
      Top             =   120
      Width           =   756
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   75
      TabIndex        =   21
      Top             =   3540
      Visible         =   0   'False
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4968
      Top             =   792
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check3 
      Caption         =   "只列印承辦單"
      Height          =   345
      Left            =   4290
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label4 
      Caption         =   "管制人："
      Height          =   204
      Left            =   528
      TabIndex        =   28
      Top             =   888
      Width           =   780
   End
   Begin MSForms.Label lblFM2 
      Height          =   252
      Left            =   2112
      TabIndex        =   27
      Top             =   864
      Width           =   1020
      Size            =   "1799;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "註：請用白紙列印"
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   3840
      TabIndex        =   26
      Top             =   1500
      Width           =   1515
   End
   Begin VB.Label Label6 
      Caption         =   "定稿、公報印表機："
      Height          =   180
      Left            =   75
      TabIndex        =   25
      Top             =   2820
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PDF執行檔路徑："
      Height          =   180
      Left            =   75
      TabIndex        =   24
      Top             =   3180
      Width           =   1560
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "公報PDF的存放路徑："
      Height          =   180
      Left            =   75
      TabIndex        =   23
      Top             =   2490
      Width           =   1740
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0/0 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   6135
   End
End
Attribute VB_Name = "frm060323"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim intWhere As Integer
Dim m_OriPrinterName As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer
Dim strReceiveNo As String '本所案號
Dim intPage As Integer '頁數
Dim iPrint As Integer '行數
Dim strPrinter As String, strPrinter2 As String

'Added by Morgan 2012/5/31
Dim strTPG04 As String, strTPG05 As String
Dim strTime As String

Dim m_LetterLanguage As String 'Add By Sindy 2015/9/21
Dim m_PrintRpt1 As Boolean, ff1 As Integer, m_strFileName1 As String 'Add By Sindy 2016/1/26
Dim m_StrLL10 As String 'Added by Lydia 2024/10/07 管制人

'Add By Sindy 2015/6/29
Private Sub Check3_Click()
   If Check3 = 1 Then
      Check1.Value = 0
      Check2.Value = 1
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
'edit by nickc 2007/02/06 不用 dll 了
'Dim objPrintDllPublic As New clsPrintPublic
   Dim strTmp As String, rsTemp1 As New ADODB.Recordset, rsTemp2 As New ADODB.Recordset
   Dim stET03 As String 'Add by Morgan 2004/10/13
   Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
   Dim stCon As String
   Dim strStarTime As String 'Add By Sindy 2015/6/26
   'Add By Sindy 2016/1/26
   Dim strFileName As String, strFullFileName As String
   Dim oFileSys As New FileSystemObject
   Dim oFile As File
   Dim strMsg As String
   Dim strNewCP09 As String
   '2016/1/26 END
   
    Select Case Index
    Case 0 '確定
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
        '公開日
        If Option1(0).Value = True Then
            If Text1(0).Text <> "" Then
                If Not ChkDate(Text1(0).Text) Then
                    Text1(0).SetFocus
                    TextInverse Text1(0)
                    Exit Sub
                End If
            Else
                MsgBox "公開日不得空白，請重新輸入 !", vbCritical
                Text1(0).SetFocus
                Exit Sub
            End If
            'Added by Lydia 2024/10/07
            If txtData(0).Visible = True And (Trim(txtData(0)) = "" Or lblFM2.Caption = "") Then
               MsgBox "請輸入管制人！", vbExclamation + vbOKOnly
               txtData(0).SetFocus
               Txtdata_GotFocus 0
               Exit Sub
            End If
            'end 2024/10/07
            
            'Added by Morgan 2012/5/31
            If Me.Check1.Value = Unchecked And Check2.Value = vbUnchecked Then
               If GetFilePath(DBDATE(Text1(0))) = False Then
                  Me.txtPath2.SetFocus
                  Exit Sub
               End If
            End If
            'end 2012/5/31
            
            pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Text1(0) 'Add By Sindy 2010/12/7
            stCon = " AND PA12=" & TransDate(Text1(0).Text, 2)
            '2009/12/1 add by sonia 從下面移上來
            'Modified by Morgan 2012/5/30 +pa11,pa12 及排序
            'Modify By Sindy 2015/6/17 +,GetEmailFlag(PA01||PA02||PA03||PA04) eMail
            '                          依E化排序,案號由小至大
            'Modify By Sindy 2016/1/26 +cp資料
            'Modified by Lydia 2019/06/17 +CP27,CP05,PA57,PA108
            'Modified by Lydia 2024/10/07 增加FCP管制人
            'strExc(0) = "SELECT PA01||PA02||PA03||PA04,DECODE(TPG08,'台一國際',1,0),PA01,PA02,PA03,PA04,PA75,pa142,fa86,cu124,pa11,pa12,GetEmailFlag(PA01||PA02||PA03||PA04) eMail,cp09,CP27,CP05,PA57,PA108" & _
                         " FROM PATENT,TPGAZETTE,FAGENT,CUSTOMER,CaseProgress" & _
                        " WHERE PA01='FCP'" & stCon & _
                          " AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPG01" & _
                          " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9,1)" & _
                          " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)" & _
                          " AND cp01(+)=pa01 AND cp02(+)=pa02 AND cp03(+)=pa03 AND cp04(+)=pa04 AND cp10='1229'" & _
                        " order by eMail,pa01,pa02,pa03,pa04"
            '2009/12/1 end
            strExc(0) = "SELECT PA01||PA02||PA03||PA04 as CaseNo,DECODE(TPG08,'台一國際',1,0) as pKind,PA01,PA02,PA03,PA04,PA75,pa142,fa86,cu124,pa11,pa12,GetEmailFlag(PA01||PA02||PA03||PA04) eMail,cp09,CP27,CP05,PA57,PA108" & _
                         " FROM PATENT,TPGAZETTE,FAGENT,CUSTOMER,CaseProgress,Nation" & _
                        " WHERE PA01='FCP'" & stCon & _
                          " AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPG01" & _
                          " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9,1)" & _
                          " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)" & _
                          " AND cp01(+)=pa01 AND cp02(+)=pa02 AND cp03(+)=pa03 AND cp04(+)=pa04 AND cp10='1229' and fa10=na01(+)"
            If Trim(txtData(0)) <> "" Then strExc(0) = strExc(0) & " and na16='" & Trim(txtData(0)) & "'"
            strExc(0) = strExc(0) & " order by eMail,pa01,pa02,pa03,pa04"
            'end 2024/10/07
        '本所案號
        Else
            If Text1(2) = "" Then
                MsgBox "本所案號不得空白，請重新輸入 !", vbCritical
                Text1(2).SetFocus
                Exit Sub
            End If
            strTmp = Text1(1) & Text1(2)
            pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Text1(1) & "-" & Text1(2) 'Add By Sindy 2010/12/7
            If Text1(3).Text = "" Then
                strTmp = strTmp & "0"
            Else
                strTmp = strTmp & Text1(3).Text
                pub_QL05 = pub_QL05 & "-" & Text1(3) 'Add By Sindy 2010/12/7
            End If
            If Text1(4).Text = "" Then
                strTmp = strTmp & "00"
            Else
                strTmp = strTmp & Text1(4).Text
                pub_QL05 = pub_QL05 & "-" & Text1(4) 'Add By Sindy 2010/12/7
            End If
            stCon = " AND " & ChgPatent(strTmp)
            '2009/12/1 add by sonia 從下面移上來
            'Modified by Morgan 2012/5/30 +pa11,pa12
            'Modify By Sindy 2015/6/17 +,GetEmailFlag(PA01||PA02||PA03||PA04) eMail
            'Modify By Sindy 2016/1/26 +cp資料
            'Modified by Lydia 2019/06/17 有特殊案件已閉卷尚需報告，則程序會個案下定稿，到時自動將發文日"111111"拿掉，如此案件又可自大批上發文
            'strExc(0) = "SELECT PA01||PA02||PA03||PA04,1,PA01,PA02,PA03,PA04,PA75,pa142,fa86,cu124,pa11,pa12,GetEmailFlag(PA01||PA02||PA03||PA04) eMail,cp09" & _
                         " FROM PATENT,FAGENT,CUSTOMER,CaseProgress" & _
                        " WHERE PA01='FCP'" & stCon & _
                          " AND (PA57<>'Y' OR PA57 IS NULL) AND PA12>0" & _
                          " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9,1)" & _
                          " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)" & _
                          " AND cp01(+)=pa01 AND cp02(+)=pa02 AND cp03(+)=pa03 AND cp04(+)=pa04 AND cp10='1229'"
            '2009/12/1 end
            'Modified by Lydia 2024/10/07 +CaseNo, pKind
            strExc(0) = "SELECT PA01||PA02||PA03||PA04 as CaseNo,1 as pKind,PA01,PA02,PA03,PA04,PA75,pa142,fa86,cu124,pa11,pa12,GetEmailFlag(PA01||PA02||PA03||PA04) eMail,cp09,CP27,CP05,PA57,PA108" & _
                         " FROM PATENT,FAGENT,CUSTOMER,CaseProgress" & _
                        " WHERE PA01='FCP'" & stCon & _
                          " AND PA12>0 AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9,1)" & _
                          " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)" & _
                          " AND cp01(+)=pa01 AND cp02(+)=pa02 AND cp03(+)=pa03 AND cp04(+)=pa04 AND cp10='1229'"
         End If
         
         If Check1.Value = 1 Then
            pub_QL05 = pub_QL05 & ";" & Check1.Caption 'Add By Sindy 2010/12/7
         End If
         
         Screen.MousePointer = vbHourglass
         'Add By Sindy 2015/6/25
         pub_OsPrinter = PUB_GetOsDefaultPrinter
         PUB_SetOsDefaultPrinter cmbPrinter2
         PUB_SetWordActivePrinter
         PUB_RestorePrinter cmbPrinter2
         '2015/6/25 END
'2009/12/1 cancel by sonia 移到上面,因下本所案號條件時不檢查公開公報檔,FCP-038864客戶要求公開當日先通知,無法等小真輸完資料
'         strExc(0) = "SELECT PA01||PA02||PA03||PA04,DECODE(TPG08,'台一國際',1,0),PA01,PA02,PA03,PA04,PA75,pa142,fa86,cu124 FROM PATENT,TPGAZETTE,FAGENT,CUSTOMER WHERE PA01='FCP'" & stCon & _
'                        " AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPG01" & _
'                        " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9,1)" & _
'                        " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1)"
'2009/12/1 end
         intI = 1
         Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Added by Lydia 2019/06/17 個案提醒閉卷
            rsTemp2.MoveFirst
            If Option1(1).Value = True Then
                If Trim("" & rsTemp2.Fields("PA57") & rsTemp2.Fields("PA108")) <> "" Then
                    If MsgBox("本案已閉卷/銷卷，是否繼續列印定稿？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
            End If
            
            List1.Clear 'Added by Morgan 2012/5/31
            With rsTemp2
                 InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/7
                 '.MoveFirst 'Mark by Lydia 2019/06/17
                 strStarTime = Format(ServerTime, "##:##:##") 'Add By Sindy 2015/6/26
                 
                 ProgressBar1.max = .RecordCount
                 ProgressBar1.Value = 0
                 lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
                 ProgressBar1.Visible = True
                 lblProgress.Visible = True
                 DoEvents
                 
                 Do While Not .EOF
                     'Added by Morgan 2012/5/31 單筆並且列印公報
                     If Option1(1).Value = True And Check2.Value = vbUnchecked Then
                        If GetFilePath(.Fields("pa12")) = False Then
                           Me.txtPath2.SetFocus
                           Screen.MousePointer = vbDefault
                           Exit Sub
                        End If
                     End If
                     'end 2012/5/31
                     
                     '處理定稿例外欄位
                     strReceiveNo = "" & .Fields(0).Value
'                     'Add By Sindy 2015/6/29 只列印承辦單
'                     If Check3.Value = vbChecked Then
'                        Call PUB_PrintFCPEmpBill(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), "14")
'                     Else
                        '若不只列印定稿清單或是選本所案號
                        If Me.Check1.Value = Unchecked Or Option1(1).Value = True Then
'                           'Add By Sindy 2015/6/17 列印FCP承辦單
'                           Call PUB_PrintFCPEmpBill(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), "14")
                           
                           m_LetterLanguage = PUB_GetLanguage(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04")) 'Add By Sindy 2015/9/21
                           
                           'Modify by Morgan 2006/6/7
                           'stET03 = GetSitu(strReceiveNo)
                           stET03 = GetSitu("" & .Fields("PA01"), "" & .Fields("PA02"), "" & .Fields("PA03"), "" & .Fields("PA04"), "" & .Fields("PA75"))
                           'end 2006/6/7
                            StartLetter "14", stET03
                            
                           'Modify by Morgan 2008/3/20 判斷是否產生電子檔
                           'NowPrint "" & .Fields(0) & "&000", "14", stET03, False, strUserNum, 0
                           bolEmail = PUB_GetEMailFlag(.Fields(0), , , bolPlusPaper)
                           If bolEmail Then
                              'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
                              If bolPlusPaper Then
                                 iCopy = 0
                              Else
                                 iCopy = 1
                              End If
                              'end 2009/10/20
                              NowPrint "" & .Fields(0) & "&000", "14", stET03, False, strUserNum, , , , , iCopy, , True, True
                              '若跑單筆時顯示訊息
                              If Option1(1).Value = True Then
                                 MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(.Fields("pa01")) & " ]！"
                              End If
                           Else
                              iCopy = 0
                              NowPrint "" & .Fields(0) & "&000", "14", stET03, False, strUserNum, 0
                           End If
                           'end 2008/3/20
                           
                           'PUB_PrintLetter "" & .Fields(0) & "&000" 'Add By Sindy 2015/6/17 直接列印定稿
                           'Modify By Sindy 2016/1/26 定稿轉PDF存卷宗區
                           strNewCP09 = .Fields("cp09")
                           strFileName = .Fields("pa01") & .Fields("pa02") & IIf(.Fields("pa04") <> "00", "-" & .Fields("pa03") & "-" & .Fields("pa04"), IIf(.Fields("pa03") <> "0", "-" & .Fields("pa03"), "")) & ".1229.CUS.PDF"
                           PUB_DelFtpFile2 strNewCP09, " and cpp02='" & strFileName & "'" '檔案改放 FTP,必須在DB資料刪除前執行
                           strSql = "delete from CasePaperPDF where cpp01='" & strNewCP09 & "' and cpp02='" & strFileName & "'"
                           cnnConnection.Execute strSql
                           If PUB_PrintLetter("" & .Fields(0) & "&000", , , True, strFullFileName) = True Then
                              Call PUB_ChkFileStatus(strFullFileName, False, strMsg)  'Added by Lydia 2022/10/31 判斷檔案是否存在, 超過時間就繼續;
                              Set oFile = oFileSys.GetFile(strFullFileName)
                              If SaveAttFile_PDF(strNewCP09, strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) = False Then
                                 'Modified by Lydia 2022/10/31 & ";" & strMsg
                                 Call ReadTxt1(.Fields("pa01") & "-" & .Fields("pa02") & "-" & .Fields("pa03") & "-" & .Fields("pa04"), strNewCP09, "定稿轉PDF失敗" & ";" & strMsg)
                              End If
                              Kill strFullFileName
                           End If
                           '2016/1/26 END

                           'Added by Lydia 2019/06/17 有特殊案件已閉卷尚需報告，則程序會個案下定稿，到時自動將發文日"111111"拿掉，如此案件又可自大批上發文
                           If Option1(1).Value = True And "" & .Fields("CP27") = "19221111" And Trim("" & .Fields("PA57") & .Fields("PA108")) <> "" And Check3.Value = vbUnchecked Then
                                '將發文日"111111"拿掉並且上承辦期限
                                strExc(1) = CompDate(2, 10, "" & .Fields("cp05"))
                                strSql = "Update Caseprogress set cp27=null,cp48=" & IIf(Val(strExc(1)) < strSrvDate(1), strSrvDate(1), strExc(1)) & _
                                            " where cp09='" & .Fields("cp09") & "' and cp10='1229' "
                                Pub_SeekTbLog strSql
                                cnnConnection.Execute strSql, intI
                           End If
            
                           'Added by Morgan 2012/5/30
                           If Check2.Value = vbUnchecked Then
                              PUB_GetCopySetting iCopy, .Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04") 'Added by Morgan 2012/7/5
                              'Modified by Morgan 2013/1/9 +pa12
                              'Modify By Sindy 2016/1/26 紙本,E+寄公開公報列印2份;E化不出
                              If iCopy = 1 Then
                                 '不印公開公報PDF檔,因此份數設為0
                                 GetPDFCopys .Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), .Fields("pa11"), 0, bolEmail, .Fields("pa12")
                              Else
                                 'GetPDFCopys .Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), .Fields("pa11"), IIf(iCopy = 0, 3, iCopy), bolEmail, .Fields("pa12")
                                 'Modified by Morgan 2017/9/13 Y21775,Y52922只要印一份公報
                                 If InStr("Y21775,Y52922", Left("" & .Fields("PA75"), 6)) > 0 Then
                                    GetPDFCopys .Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), .Fields("pa11"), 1, bolEmail, .Fields("pa12")
                                 Else
                                    GetPDFCopys .Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), .Fields("pa11"), 2, bolEmail, .Fields("pa12")
                                 End If
                              End If
                              '2016/1/26 END
                           End If
                           'end 2012/5/30
               
                           'Modify by Morgan 2008/3/20 產生電子檔時不印地址條
                           If Not bolEmail Or bolPlusPaper Then
'                              'Add By Sindy 2015/9/21 日文定稿才要印地址條
'                              If m_LetterLanguage = "3" Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'                              '2015/9/21 END
                                 '新增地址條列表資料
                                 pub_AddressListSN = pub_AddressListSN + 1
                                 PUB_AddNewAddressList strUserNum, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, "" & pub_AddressListSN, "0"
'                              End If
                           End If
                        End If
'                     End If
                     
                     'Add By Sindy 2015/6/29
                     If Option1(0).Value = True Then
                     'If Option1(0).Value = True And Check3.Value = vbUnchecked Then
                     '2015/6/29 END
                        '新增整批定稿列印清單資料
                        'Modify By Sindy 2015/6/4
                        'PUB_AddNewLetterList "公開通知函", Me.Text1(0).Text, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value
                        'Modified by Lydia 2024/10/07 +txtData(0)
                        PUB_AddNewLetterList "公開通知函", Me.Text1(0).Text, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, IIf(bolEmail, IIf(bolPlusPaper, "Ｅ", "ｅ"), ""), txtData(0)
                        '2015/6/4 END
                     End If
                     
                     ProgressBar1.Value = ProgressBar1.Value + 1
                     lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
                     DoEvents
                     .MoveNext
                 Loop
            End With
'            'Add By Sindy 2015/6/26
'            If Option1(0).Value = True Then '整批
'               PUB_SendMail strUserNum, "97038", "", "外專執行＜公開通知函＞整批的執行時間: " & strStarTime & " ~ " & Format(ServerTime, "##:##:##"), "如主旨"
'            End If
'            '2015/6/26 END
            
'            'Added by Morgan 2012/5/31
'            If List1.ListCount > 0 Then
'               Call PrinBatchPdf
'               MsgBox "定稿產生完成 ! (列印PDF花費時間：" & strTime & "  " & time() & ")", vbInformation
'            Else
'            'end 2012/5/31
'               MsgBox "定稿產生完成 !", vbInformation
'            End If 'Added by Morgan 2012/5/31
            
            'MsgBox "列印完成 !", vbInformation
            'Modify By Sindy 2016/1/26
            If m_PrintRpt1 = True Then
               Close ff1
               strMsg = "請至下列位置列印檢核表：" & PUB_Getdesktop & "\" & m_strFileName1
            End If
            MsgBox "定稿列印完畢！ " & strMsg, vbInformation
            '2016/1/26 END
         Else
            InsertQueryLog (0) 'Add By Sindy 2010/12/7
            MsgBox "無符合條件之資料 !", vbInformation
         End If
         'Add By Sindy 2015/6/25
         PUB_SetOsDefaultPrinter pub_OsPrinter
         PUB_RestorePrinter strPrinter2
         '2015/6/25 END
         Screen.MousePointer = vbDefault
      Case 1 '結束
         Me.Enabled = False
         Unload Me
   End Select
End Sub

'Add By Sindy 2016/1/26
'資料檢核表
Private Sub ReadTxt1(strCaseNo As String, strRecvNo As String, strNote As String)
Dim i As Integer
Dim strTemp(1 To 7) As String
   
   If m_PrintRpt1 = False Then
      m_PrintRpt1 = True
      If ff1 > 0 Then Close #ff1
      ff1 = FreeFile
      m_strFileName1 = Me.Caption & Text1(0) & "資料檢核表.txt"
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

'Add By Sindy 2015/11/13
Private Sub Command2_Click()
Dim stFileName As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.pdf"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "pdf檔案 (*.pdf)|*.pdf"
      .InitDir = PUB_Getdesktop
      '.MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         txtPath2.Text = Mid(.FileName, 1, InStrRev(.FileName, "\"))
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   Option1_Click 0
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   
   'Added by Morgan 2012/5/31
   PUB_SetPrinter Me.Name, cmbPrinter2, strPrinter2
   If Pub_StrUserSt03 <> "M51" Then
      Me.Height = 4560
   End If
   List1.Clear
   
   'Modified by Morgan 2017/11/3 改呼叫共用
   'SetFileAssociation
   txtPDFPath = PUB_SetFileAssociation
   'end 2017/11/3
   'end 2012/5/31
   
   'Add By Sindy 2015/11/13 紀錄在資料庫,否則換電腦或使用者會讀不到
   txtPath2.Text = PUB_GetLastDate(Me.Name, UCase("txtPath2"))
   If txtPath2.Text = "" Then
      txtPath2.Text = "\\Pat3\GAZETTE\PGXml\img_1\pub012012\"
   End If
   '2015/11/13 END
   
   'Added by Lydia 2024/10/07 FCP程序管制人---11/1上線
   If strSrvDate(1) >= "20241101" Then
      txtData(0) = strUserNum
      lblFM2 = strUserName
   Else
      Label4.Visible = False
      txtData(0).Visible = False
      lblFM2.Visible = False
   End If
   'end 2024/10/07
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2015/6/25
   PUB_RestorePrinter cmbPrinter2
   '2015/6/25 END
   'Copy from cmdok_Click by Morgan 2004/10/26
   '列印定稿整批列印清單
   PrintLetterList strUserNum
   '刪除定稿整批列印資料
   'Modified by Lydia +傳入刪除條件
   'PUB_DeleteLetterList strUserNum
   PUB_DeleteLetterList strUserNum, " and LL02='公開通知函' "
   'Add By Sindy 2015/6/25
   PUB_RestorePrinter strPrinter2
   '2015/6/25 END
   
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
   '2004/10/26 end
   
   'Added by Morgan 2012/5/31
   If Me.cmbPrinter2.Text <> Me.cmbPrinter2.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter2.Name, "0", "0", Me.cmbPrinter2.Text
   End If
   'end 2012/5/31
   
   Set frm060323 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
 Dim txt As TextBox, i As Integer
On Error Resume Next
   For Each txt In Text1
      txt.Enabled = False
   Next
   Select Case Index
      Case 0
         Text1(0).Enabled = True
         Text1(0).SetFocus
         Check2.Value = vbUnchecked 'Added by Morgan 2012/6/1
      Case 1
         For i = 2 To 4
            Text1(i).Enabled = True
         Next
         Text1(2).SetFocus
         Check2.Value = vbChecked 'Added by Morgan 2012/6/1
   End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   If Text1(Index) = "" Then Exit Sub
   If Option1(0).Value = True Then
      If Index = 0 Then
         If Text1(Index).Text <> "" Then
            If Not ChkDate(Text1(Index).Text) Then
               Text1(Index).SetFocus
               TextInverse Text1(Index)
            End If
         Else
            MsgBox "公開日不得空白，請重新輸入 !", vbCritical
            Text1(Index).SetFocus
         End If
      End If
   Else
      If Index = 1 Then
         If Text1(Index).Text = "" Then
            MsgBox "本所案號不得空白，請重新輸入 !", vbCritical
            Text1(Index).SetFocus
         End If
      End If
   End If
End Sub

'取得定稿處理方式
'Modify by Morgan 2006/6/6 改參數
'Private Function GetSitu(strPA0104 As String) As String
Private Function GetSitu(p_stPA01 As String, p_stPA02 As String, p_stPA03 As String, p_stPA04 As String, p_stPA75 As String) As String

Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

'Modify by Morgan 2006/6/7 改Call公用函數
'
'Dim StrSqlB As String
'Dim rsB As New ADODB.Recordset
'
'GetSitu = "00"
'StrSQLa = "Select * From PATENT WHERE " & ChgPatent(strPA0104)
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'    '若基本檔有設定定稿語文
'    If "" & rsA("PA85").Value <> "" Then
'        Select Case "" & rsA("PA85").Value
'        Case "1" '中文
'            GetSitu = "01"
'        Case "2" '英文
'            GetSitu = "02"
'        Case "3" '日文
'            GetSitu = "03"
'        End Select
'    '若基本檔未設定定稿語文
'    Else
'        '若基本檔有代理人
'        If "" & rsA("PA75").Value <> "" Then
'            StrSqlB = "Select * From FAGENT WHERE FA01='" & Mid(rsA("PA75").Value, 1, 8) & "' AND FA02='" & Mid(rsA("PA75").Value, 9, 1) & "'"
'            rsB.CursorLocation = adUseClient
'            rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsB.RecordCount > 0 Then
'                Select Case "" & rsB("FA31").Value
'                Case "1" '中文
'                    GetSitu = "01"
'                Case "2" '英文
'                    GetSitu = "02"
'                Case "3" '日文
'                    GetSitu = "03"
'                End Select
'            End If
'        '若基本檔無代理人
'        Else
'            StrSqlB = "Select * From CUSTOMER WHERE CU01='" & Mid(rsA("PA26").Value, 1, 8) & "' AND CU02='" & Mid(rsA("PA26").Value, 9, 1) & "'"
'            rsB.CursorLocation = adUseClient
'            rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsB.RecordCount > 0 Then
'                Select Case "" & rsB("CU64").Value
'                Case "1" '中文
'                    GetSitu = "01"
'                Case "2" '英文
'                    GetSitu = "02"
'                Case "3" '日文
'                    GetSitu = "03"
'                End Select
'            End If
'        End If
'    End If
'End If
'If rsB.State <> adStateClosed Then rsB.Close
'Set rsB = Nothing
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
''若無任何設定, 預設為英文一般定稿
'If GetSitu = "00" Then GetSitu = "02"

strExc(1) = PUB_GetLanguage(p_stPA01, p_stPA02, p_stPA03, p_stPA04)

'Added by Morgan 2012/8/14
'增加中文定稿 Ex. FCP-42716
If strExc(1) = "1" Then
   GetSitu = "06"
Else
'end 2012/8/14
   GetSitu = "0" & strExc(1)
   
End If 'Added by Morgan 2012/8/14
'end 2006/6/7

If GetSitu = "02" Then
    '判斷進度檔是否有實體審查或優先審查
    'Modify by Morgan 2006/6/7
    'StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(strPA0104) & " And (CP10='416' Or CP10='425') "
    'Modified by Morgan 2013/12/24 +107,435--Susan
    StrSQLa = "Select * From CaseProgress Where cp01='" & p_stPA01 & "' and cp02='" & p_stPA02 & "' and cp03='" & p_stPA03 & "' and cp04='" & p_stPA04 & "' And (CP10='416' Or CP10='425' Or CP10='107' Or CP10='435') "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        GetSitu = "02" '無實體審查期限
    Else
        GetSitu = "05" '有實體審查期限
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    
   'Add by Morgan 2006/6/12
   If p_stPA75 <> "" Then
      p_stPA75 = Left(p_stPA75, 6)
      'Modify by Morgan 2008/5/8 + Y21775 --David
      'Modified by Morgan 2013/10/9 +Y53893 --102/10/7 譚文容請作單
      'Modified by Morgan 2014/12/3 -Y48309 --Jessica
      'Modified by Morgan 2017/9/14 -Y21775 --洪郁嵐,吳若芬
      If p_stPA75 = "Y49575" Or p_stPA75 = "Y45697" Or p_stPA75 = "Y20412" Or p_stPA75 = "Y48162" Or p_stPA75 = "Y53893" Then
         If GetSitu = "02" Then
            GetSitu = "01"
         Else
            GetSitu = "04"
         End If
      End If
   End If
   'end 2006/6/12
End If
End Function

'定稿例外欄位
Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 20) As String, i As Integer, j As Integer, strTmp As String
Dim ii As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

    ii = 0
    EndLetter ET01, strReceiveNo & "&000", ET03, strUserNum
    '判斷是否不續辦但准通知
    'Modify By Cheng 2003/05/19
'    strSQLA = "Select PA89 From Patent Where " & ChgPatent(strReceiveNo)
    StrSQLa = "Select PA89, PA06 From Patent Where " & ChgPatent(strReceiveNo)
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        If "" & rsA.Fields(0).Value = "Y" Then
             ii = ii + 1
             '附註
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','附註','P.S. : This case has been allowed. If your client(s) want(s) to maintain this case, please notify us immediately.')"
        End If
        'Add By Cheng 2003/05/19
        If "" & rsA.Fields(1).Value <> "" Then
             ii = ii + 1
             '專利英文名稱
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','專利英文名稱','" & ChgSQL(rsA.Fields(1).Value) & "')"
        End If
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    'Modify by Morgan 2004/10/13 加日文定稿
    'If ET03 = "05" Then
    'Remove by Morgan 2006/6/21 都跑以免加新定稿時漏加
    'If ET03 = "05" Or ET03 = "03" Then
        StrSQLa = "Select NP09 From Nextprogress Where " & ChgNextProgress(strReceiveNo) & " And NP07='416' And NP06 Is Null "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            If "" & rsA.Fields(0).Value <> "" Then
                 ii = ii + 1
                 '法定期限
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                   "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','法定期限','" & rsA.Fields(0).Value & "') "
            End If
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    'End If
    If ii > 0 Then
        'edit by nickc 2007/02/05 不用 dll 了
        'If Not objLawDll.ExecSQL(ii, strTxt) Then
        If Not ClsLawExecSQL(ii, strTxt) Then
           MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
        End If
    End If
End Sub

'Add by Morgan 2010/4/29
'列印表頭(從PrintLetterList抽出)
Private Sub PrintHead()
   intPage = intPage + 1
   iPrint = 500
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4200
   Printer.CurrentY = iPrint
   Printer.Print "整批定稿列印清單"
   
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   'Added by Lydia 2024/10/07
   If strUserNum <> m_StrLL10 Then
      Printer.CurrentX = 4100
      Printer.CurrentY = iPrint
      Printer.Print "管制人：" & GetStaffName(m_StrLL10)
   End If
   'end 2024/10/07
   
   Printer.CurrentX = 8500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")

   
   iPrint = iPrint + 300
   '2008/11/10 ADD BY SONIA
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   'Modify by Morgan 2010/4/29
   'Printer.Print "△附寄件清單"
   'Modify by Morgan 2011/5/26 +#
   'Modified by Lydia 2022/08/02 +◎請在三天內完成
   Printer.Print "△附寄件清單  ＊交由承辦傳送  ＃直接上傳客戶EPMS之電腦系統  □在14天內完成通知  ◎在3天內完成"
   'end 2010/4/29
   '2008/11/10 END
   
   'Added by Morgan 2012/11/27
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   'Modified by Morgan 2013/10/9 +＄
   'Modified by Lydia 2015/02/05 +★公開公報函另行掃描
   'Modified by Morgan 2018/3/30 +▇優先處理
   'Modified by Lydia 2021/08/23 +◆交承辦修改公開信
   Printer.Print "※公開函只留底不E/寄客戶  ＄只ｅ公開信但不ｅ公開公報  ▇優先處理  ◆交承辦修改公開信"
   'end 2012/11/27
   
   'Modified by Morgan 2017/10/13 +☆
   'Modified by Morgan 2019/3/18 +●
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "★公開函另行掃描,再寄客戶  ☆公開信及公開函另行掃描,再寄客戶 ●另下載公開發明說明書一起寄客戶"
   'end 2017/10/13
   
   'Added by Morgan 2014/9/9 +◇
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   'Modify By Sindy 2015/6/4
   'Printer.Print "◇Email + 傳真通知客戶Email內容"
   Printer.Print "◇Email + 傳真通知客戶Email內容    ｅE化案件  ＥE化加紙本"
   '2015/6/4 END
   'end 2014/9/9
   
   Printer.CurrentX = 8500
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(intPage)
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "定稿名稱"
   Printer.CurrentX = 2100
   Printer.CurrentY = iPrint
   Printer.Print "條件"
   Printer.CurrentX = 4200
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = 6300
   Printer.CurrentY = iPrint
   Printer.Print "公開號"
   Printer.CurrentX = 7550
   Printer.CurrentY = iPrint
   Printer.Print "代理人"
   iPrint = iPrint + 300
   
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
End Sub

'Add By Cheng 2003/09/10
'列印定稿清單
Private Sub PrintLetterList(strUserNumber As String)
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
   Dim intCnt As Integer '筆數
   Dim strLetterName As String '定稿名稱
   Dim strCondition As String '條件
   
   Dim strMurgitroyd As String 'Added by Lydia 2021/01/05
   
   '2008/11/10 MODIFY BY SONIA 加PA75
   'StrSQLa = "Select LL02, LL03, LL04, LL05, LL06, LL07, PA13, Decode(FA05, Null, Nvl(FA04, FA06), FA05||' '||FA63||' '||FA64||' '||FA65) From LetterList, Fagent, Patent Where substr(LL08,1,8)=FA01 And substr(LL08,9,1)=FA02 And LL04=PA01 And LL05=PA02 And LL06=PA03 And LL07=PA04 And LL01='" & strUserNum & "' "
   'StrSQLa = StrSQLa & " Union Select LL02, LL03, LL04, LL05, LL06, LL07, PA13, Decode(CU05, Null, Nvl(CU04, CU06), CU05||' '||CU88||' '||CU89||' '||CU90) From LetterList, Customer, Patent Where substr(LL08,1,8)=CU01 And substr(LL08,9,1)=CU02 And LL04=PA01 And LL05=PA02 And LL06=PA03 And LL07=PA04 And LL01='" & strUserNum & "' "
   'Modify by Morgan 2011/5/26 +PA26,PA27,PA28,PA29,PA30
   'Modify By Sindy 2015/6/4 +LL09
   'Modified by Lydia 2019/04/30 +電腦名稱("@" & pub_HostName )
   'Modified by Lydia 2020/09/24 +程式名稱 and LL02='公開通知函'
   'Modified by Lydia 2024/10/07 +管制人LL10
   StrSQLa = "Select LL02, LL03, LL04, LL05, LL06, LL07, PA13, Decode(FA05, Null, Nvl(FA04, FA06), FA05||' '||FA63||' '||FA64||' '||FA65),PA75,PA26,PA27,PA28,PA29,PA30,LL09,LL10 From LetterList, Fagent, Patent Where substr(LL08,1,8)=FA01 And substr(LL08,9,1)=FA02 And LL04=PA01 And LL05=PA02 And LL06=PA03 And LL07=PA04 And LL01='" & strUserNum & "@" & pub_HostName & "'  and LL02='公開通知函' "
   StrSQLa = StrSQLa & " Union Select LL02, LL03, LL04, LL05, LL06, LL07, PA13, Decode(CU05, Null, Nvl(CU04, CU06), CU05||' '||CU88||' '||CU89||' '||CU90),PA75,PA26,PA27,PA28,PA29,PA30,LL09,LL10 From LetterList, Customer, Patent Where substr(LL08,1,8)=CU01 And substr(LL08,9,1)=CU02 And LL04=PA01 And LL05=PA02 And LL06=PA03 And LL07=PA04 And LL01='" & strUserNum & "@" & pub_HostName & "'  and LL02='公開通知函' "
   StrSQLa = StrSQLa & " Order By 1, 2, 3 "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   '若有整批定稿清單資料
   If rsA.RecordCount > 0 Then
      MsgBox "準備列印整批定稿清單!!!", vbExclamation + vbOKOnly
RePrint:
      strLetterName = ""
      strCondition = ""
      strMurgitroyd = Pub_GetSpecMan("外專MURGITROYD設定") 'Added by Lydia 2021/01/05
      intPage = 0
      intCnt = 0
      Printer.Orientation = 1
      
      m_StrLL10 = "" & rsA.Fields("LL10") 'Added by Lydia 2024/10/07
      
      PrintHead
      
      '移至第一筆資料
      rsA.MoveFirst
      While Not rsA.EOF
         If strLetterName <> "" & rsA.Fields(0).Value Then
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print "" & rsA.Fields(0).Value
            strLetterName = "" & rsA.Fields(0).Value
            
            Printer.CurrentX = 2100
            Printer.CurrentY = iPrint
            Printer.Print "" & rsA.Fields(1).Value
            strCondition = "" & rsA.Fields(1).Value
         ElseIf strCondition <> "" & rsA.Fields(1).Value Then
            Printer.CurrentX = 2100
            Printer.CurrentY = iPrint
            Printer.Print "" & rsA.Fields(1).Value
            strCondition = "" & rsA.Fields(1).Value
         End If
         
         'Modify by Morgan 2010/3/23 Y51774 關係企業也要附寄件清單,Y5177401 + ＊號表交由承辦傳送
         '2008/11/10 ADD BY SONIA 寄件時要附寄件清單
         'If rsA.Fields(8).Value = "Y45148000" Or rsA.Fields(8).Value = "Y51774000" Then
         '   Printer.CurrentX = 4000
         '   Printer.CurrentY = iPrint
         '   Printer.Print "△"
         'End If
         '2008/11/10 END
         strExc(1) = ""
         'Modify by Morgan 2011/6/15 --要與PUB_PrintLetterList一致
         'If rsA.Fields(8).Value = "Y45148000" Or Left(rsA.Fields(8).Value, 6) = "Y51774" Then
         '   strExc(1) = strExc(1) & "△"
         'End If
         'If Left(rsA.Fields(8).Value, 8) = "Y5177401" Then
         '   strExc(1) = strExc(1) & "＊"
         'End If
         If rsA.Fields(8).Value = "Y45148000" Or Left(rsA.Fields(8).Value, 8) = "Y5177401" Then
            strExc(1) = strExc(1) & "＊"
         ElseIf Left(rsA.Fields(8).Value, 6) = "Y51774" Then
            strExc(1) = strExc(1) & "△"
         'Added by Morgan 2012/11/27
         'Modified by Morgan 2013/1/15 +Y47794000
         'Modified by Morgan 2013/1/18 +X47794000
         'Modified by Morgan 2016/11/28 +Y46556000
         'Modified by Morgan 2017/5/31 +FCP-056858--劉又華
         'Modified by Morgan 2017/8/4 +FCP-057196--吳若芬
         'Modified by Morgan 2017/12/28 +FCP-058113--洪郁嵐
         'Modified by Lydia 2019/10/14 +Y54145 (RYUKA IP Law Firm )--洪郁嵐
         'Modified by Lydia 2021/03/22 +Y34232 (YASUTOMI & ASSOCIATES) + X73190 (OSAKA SODA CO., LTD.) --洪郁嵐
         'Modified by Lydia 2022/06/07 + 代理人 <Y52981> TANI & ABE, p.c. +申請人<X82375> APB Corporation ---Arashi
         'Modified by Morgan 2024/6/19 +<Y5228300>＋ <X76778000> --Arashi
         ElseIf "" & rsA.Fields("pa75").Value = "Y27766000" Or rsA.Fields("pa75").Value = "Y46556000" Or rsA.Fields("pa75").Value = "Y47794000" Or rsA.Fields("pa75").Value = "Y54145000" _
                    Or rsA.Fields("pa26").Value = "X47794000" Or rsA.Fields("LL04").Value & rsA.Fields("LL05").Value = "FCP056858" Or rsA.Fields("LL04").Value & rsA.Fields("LL05").Value = "FCP057196" Or rsA.Fields("LL04").Value & rsA.Fields("LL05").Value = "FCP058113" _
                    Or ("" & rsA.Fields("pa75").Value = "Y34232000" And InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X73190000") > 0) _
                    Or ("" & rsA.Fields("pa75").Value = "Y52981000" And InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X82375000") > 0) _
                    Or ("" & rsA.Fields("pa75").Value = "Y52283000" And InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X76778000") > 0) Then
            strExc(1) = strExc(1) & "※"
         'end 2012/11/27
         'Added by Morgan 2013/10/9
         ElseIf "" & rsA.Fields("pa75").Value = "Y53893000" Then
            strExc(1) = strExc(1) & "＄"
         'end 2013/10/9
         'Added by Morgan 2014/9/9
         ElseIf "" & rsA.Fields("pa75").Value = "Y48804000" Then
            strExc(1) = strExc(1) & "◇"
         'end 2014/9/9
         'Added by Lydia 2022/08/02 --Arashi
         'Modified by Lydia 2022/12/01 Y54339(Metis IP)+ X85549(SHENZHEN) ---- Teresa
         'Modified by Morgan 2024/4/11 +FCP-070422 --Arashi
         ElseIf "" & rsA.Fields("pa75").Value = "Y54047000" Or ("" & rsA.Fields("pa75").Value = "Y54339000" And InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X85549000") > 0) _
          Or rsA.Fields("LL04").Value & rsA.Fields("LL05").Value = "FCP070422" Then
            strExc(1) = strExc(1) & "◎"
         'end 2022/08/02
         End If
         'end 2011/6/15
         
         'Added by Morgan 2018/3/30
         'Modified by Lydia 2021/03/08 + X68646 (ASM IP Holding B.V.) 、X47178 (ASM AMERICA, INC.)
         'If "" & rsA.Fields("pa75").Value = "Y54116000"  Then
         'Modified by Lydia 2022/06/07 + FCP-066161個案 公開公報清單符號為 「 ■優先處理」
         'Modified by Morgan 2022/7/22 +Y45799050,Y28343010 --Arashi
         'Modified by Lydia 2022/08/10 + Y5221200 --Arashi
         'Modified by Lydia 2022/08/29 + FCP-066659 --Arashi
         'Modified by Lydia 2022/09/12 +Y55786(Tribalyte Ideas) --Arashi
         'Modified by Lydia 2022/10/13 +FCP-066796 --Arashi
         'Modified by Morgan 2022/12/5 +FCP-067322 --Arashi
         'Modified by Lydia 2023/08/16 +FCP-069056 --Arashi
         'Modified by Morgan 2024/4/11 去掉已公開的本所案號(原檢查欄位也有錯，少了系統別)
         'If "" & rsA.Fields("pa75").Value = "Y54116000" Or InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X6864600") > 0 _
             Or InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X4717800") > 0 Or rsA.Fields("LL05").Value = "FCP066161" _
             Or rsA.Fields("pa75").Value = "Y45799050" Or rsA.Fields("pa75").Value = "Y28343010" Or rsA.Fields("pa75").Value = "Y52212000" Or rsA.Fields("pa75").Value = "Y55786000" _
             Or rsA.Fields("LL05").Value = "FCP066659" Or rsA.Fields("LL05").Value = "FCP066796" Or rsA.Fields("LL05").Value = "FCP067322" Or rsA.Fields("LL05").Value = "FCP069056" Then
         'Modified by Lydia 2025/02/05 +FCP-071706 --Winfrey
         If "" & rsA.Fields("pa75").Value = "Y54116000" Or InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X6864600") > 0 _
             Or InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X4717800") > 0 _
             Or rsA.Fields("pa75").Value = "Y45799050" Or rsA.Fields("pa75").Value = "Y28343010" Or rsA.Fields("pa75").Value = "Y52212000" Or rsA.Fields("pa75").Value = "Y55786000" _
             Or rsA.Fields("LL05").Value = "FCP071706" Then
            strExc(1) = strExc(1) & "▇"  '優先處理
         End If
         'end 2018/3/30
         
         'Add by Morgan 2011/5/26
         If InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X4783301") > 0 Then
            strExc(1) = strExc(1) & "＃"
         End If
         
         'Add by Morgan 2017/3/10 □請在14天內完成通知
         If InStr("Y27856000,Y27856B30", rsA.Fields("pa75")) > 0 And (InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X72643000") > 0 Or InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X74976000") > 0) Then
            strExc(1) = strExc(1) & "□"
         End If
         'end 2017/3/10
         
         'Modified by Lydia 2015/02/05 ★公開公報函另行掃描
         'Modified by Morgan 2016/10/14 +Y42237--劉又華
         'Modified by Morgan 2018/3/30 +Y19357B10--洪郁嵐
         'Modified by Morgan 2018/12/13 +Y55101--洪郁嵐
         'Modified by Morgan 2019/5/9 +Y53675--洪郁嵐
         'Modified by Lydia 2019/06/11 +Y54372--洪郁嵐
         'Modified by Lydia 2019/07/04 +Y46957--洪郁嵐
         'Modified by Lydia 2020/02/12 +Y55263--洪郁嵐
         'Modified by Lydia 2020/09/03 +Y55421--洪郁嵐
         If InStr("Y45848000,Y52709000,Y42237000,Y19357B10,Y55101000,Y53675000,Y54372000,Y46957000,Y55263000,Y55421000", rsA.Fields("pa75")) > 0 Then
             strExc(1) = "★" & strExc(1)
         'Added by Morgan 2017/10/13 --洪郁嵐
         'Modified by Morgan 2017/11/7 +Y47034 --洪郁嵐
         ElseIf rsA.Fields("pa75") = "Y48651000" Or rsA.Fields("pa75") = "Y47034000" Then
            strExc(1) = "☆" & strExc(1)
         End If
         
         'Added by Morgan 2019/3/18 ●另下載公開發明說明書一起寄客戶 'Memo by Lydia 2021/01/05 ●另下載公開發明書一起寄客戶
         'Modified by Lydia 2021/01/05 +MURGITROYD
         'If rsA.Fields("pa75") = "Y52709000" Then
         If ("" & rsA.Fields("pa75") = "Y52709000") Or ("" & rsA.Fields("pa75") <> "" And InStr(strMurgitroyd, rsA.Fields("pa75")) > 0) Then
            strExc(1) = strExc(1) & "●"
         End If
         'end 2019/3/18
         
         'Added by Lydia 2021/08/23 ◆交承辦修改公開信
         '代理人 Y2139903 (NTD Patent & Trademark AgencyLtd.) + 申請人 X81804 (Fresenius Medical Care Deutschland GmbH)
         If "" & rsA.Fields("pa75") = "Y21399030" And InStr(rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "X81804") > 0 Then
            strExc(1) = strExc(1) & "◆"
         End If
         'end 2021/08/23
         
         strExc(1) = rsA("LL09") & strExc(1) 'Add By Sindy 2015/6/4
         If strExc(1) <> "" Then
            Printer.CurrentX = 4200 - Printer.TextWidth(strExc(1))
            Printer.CurrentY = iPrint
            Printer.Print strExc(1)
         End If
         'end 2010/3/23
         
         Printer.CurrentX = 4200
         Printer.CurrentY = iPrint
         Printer.Print "" & rsA.Fields(2).Value & "-" & rsA.Fields(3).Value & "-" & rsA.Fields(4).Value & "-" & rsA.Fields(5).Value
         Printer.CurrentX = 6300
         Printer.CurrentY = iPrint
         Printer.Print "" & rsA.Fields(6).Value
         Printer.CurrentX = 7550
         Printer.CurrentY = iPrint
         Printer.Print "" & rsA.Fields(7).Value
         iPrint = iPrint + 300
         intCnt = intCnt + 1
         rsA.MoveNext
         If rsA.EOF = True Then
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print String(200, "-")
            iPrint = iPrint + 300
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print "筆數：共" & Format(rsA.RecordCount, "#,##0") & "筆!!!"
            iPrint = iPrint + 300
         Else
            If intCnt >= 40 Then
               Printer.CurrentX = 500
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               iPrint = iPrint + 300
               Printer.NewPage
               
               intCnt = 0
               strLetterName = ""
               strCondition = ""
               PrintHead
            End If
         End If
      Wend
      Printer.EndDoc
      '可重覆列整批定清單資料
      If MsgBox("整批定稿清單已列印完畢，您是否要重新列印???", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
          GoTo RePrint
      End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Sub

'Added by Morgan 2012/5/31
Private Function GetFilePath(strDate As String) As Boolean
   
   Dim i As Integer, j As Integer
   
On Error GoTo ErrHnd
   
   If IsEmptyText(txtPath2) = True Then
      MsgBox "請輸入公開公報PDF的存放路徑！", vbExclamation + vbOKOnly
      Exit Function
   End If
   'If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   If Right(Trim(txtPath2), 1) <> "\" Then txtPath2 = txtPath2 & "\"
   
   strTPG04 = Format(Val(Val(Left(strDate, 4)) - 1911) - 91, "00")
   j = Val(Mid(strDate, 5, 2))
   i = (j - 1) * 2
   j = Val(Right(strDate, 2))
   If j >= 1 And j < 11 Then
      i = i + 1
   ElseIf j >= 11 And j < 21 Then
      i = i + 2
   End If
   '92年公報從5月開始
   If Val(strDate) < 20040000 Then i = i - 8
   strTPG05 = Format(i, "00")
   
   'Add By Sindy 2015/11/13
   If InStr(UCase(txtPath2), UCase("\img_1\pub0")) > 0 Then
      txtPath2 = Mid(txtPath2, 1, InStrRev(UCase(txtPath2), UCase("\img_1\pub0")) + 10) & strTPG04 & "0" & strTPG05 & "\"
   End If
   '2015/11/13 END
   'If Dir(txtPath2 & "\img_1\pub0" & strTPG04 & "0" & strTPG05 & "\") = "" Then
   If Dir(txtPath2) = "" Then
      MsgBox "公開公報PDF的存放路徑中無" & strTPG04 & "卷" & strTPG05 & "期資料！"
      Exit Function
   End If
   
   'Add By Sindy 2015/11/13 紀錄在資料庫,否則換電腦或使用者會讀不到
   PUB_SaveLastDate Me.Name, UCase("txtPath2"), txtPath2.Text
   txtPath2.Text = PUB_GetLastDate(Me.Name, UCase("txtPath2"))
   '2015/11/13 END
   
   GetFilePath = True
   Exit Function
   
ErrHnd:
   If Err.Number = 76 Then
      MsgBox "公開公報PDF的存放路徑中無" & strTPG04 & "卷" & strTPG05 & "期資料！"
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function

'Modified by Morgan 2013/1/9 +strPA12
Private Sub GetPDFCopys(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String, StrPA11 As String, ByRef int_Copys As Integer, ByVal pCopyFile As Boolean, Optional strPA12 As String)
   Dim strFileName As String, strToPath As String
   
   'Modify By Sindy 2013/1/4
   'strFileName = txtPath2 & "\img_1\pub0" & strTPG04 & "0" & strTPG05 & "\" & StrPA11 & "-P01.pdf"
   'Modified by Morgan 2013/1/9 102/1/1以前公開維持舊格式
   If Val(strPA12) >= "20130101" Then
      'strFileName = txtPath2 & "\img_1\pub0" & strTPG04 & "0" & strTPG05 & "\" & StrPA11 & ".pdf"
      strFileName = txtPath2 & StrPA11 & ".pdf"
   Else
      'strFileName = txtPath2 & "\img_1\pub0" & strTPG04 & "0" & strTPG05 & "\" & StrPA11 & "-P01.pdf"
      strFileName = txtPath2 & StrPA11 & "-P01.pdf"
   End If
   '2013/1/4 End
   
   'Add By Sindy 2016/1/26 紙本,E+寄公開公報列印2份;E化不印出; 因此份數若為0則代表不用印出PDF檔
   If int_Copys > 0 Then
   '2016/1/26 END
      List1.AddItem strFileName & " " & int_Copys
      Call PrinBatchPdf(List1.ListCount - 1) 'Add By Sindy 2015/6/17 直接列印出來
   End If
   
   If pCopyFile Then
      strToPath = Left(m_strFilePath, InStrRev(m_strFilePath, "\"))
      'Modify By Sindy 2015/6/5
      'strToPath = strToPath & strPA01 & strPA02 & strPA03 & strPA04 & "_" & Mid(strFileName, InStrRev(strFileName, "\") + 1)
      strToPath = strToPath & strPA01 & strPA02
      If strPA03 & strPA04 <> "000" Then
         strToPath = strToPath & strPA03 & strPA04
      End If
      strToPath = strToPath & EfileNameFCP_14
      '2015/6/5 END
      FileCopy strFileName, strToPath
   End If
End Sub

'Modify By Sindy 2015/6/17 +strPrintRow As String
'strPrintRow : A.列印全部
'            : 數字.列印筆數
Private Sub PrinBatchPdf(strPrintRow As String)
   Dim program_name As String
   Dim process_id As Long
   Dim process_handle As Long
   Dim ii As Integer, kk As Integer
   Dim strTemp As Variant
   Dim ff1 As Integer
   Dim strPrinterName As String
   Dim intFileCnt As Integer
   Dim MySize
   Dim intRow As Integer, intTotRow As Integer
   
   strTime = time()
   
'   ProgressBar1.max = List1.ListCount
'   ProgressBar1.Value = 0
'   lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
'   ProgressBar1.Visible = True
'   lblProgress.Visible = True
'   DoEvents

   program_name = txtPDFPath
   strPrinterName = cmbPrinter2

    ' Start the program.
On Error GoTo ShellError
    
    '因為第 2 個以後開啟的 Reader 才會印完後自動關閉,所以固定先開一個空的程式,全部印完後再關閉
    process_id = SHELL(program_name, vbHide)
    process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
    
On Error GoTo 0
    
   If ff1 > 0 Then Close #ff1
   ff1 = FreeFile
   Open App.path & "\FCP公開公報列印PDF時間資訊.txt" For Output As #ff1
   
   'Modify By Sindy 2015/6/17
   If strPrintRow = "A" Then
      intRow = 0
      intTotRow = List1.ListCount - 1
   Else
      intRow = CInt(strPrintRow)
      intTotRow = intRow
   End If
   'For ii = 0 To List1.ListCount - 1
   For ii = intRow To intTotRow
   '2015/6/17 END
      strTemp = Split(List1.List(ii), " ")
      For kk = 1 To Val(strTemp(1)) '列印份數
         intFileCnt = intFileCnt + 1
         mdiMain.tmrConnect.Tag = 0
         PrintOnePdf program_name, " /n /t """ & strTemp(0) & """ """ & strPrinterName & """"
      Next
      
'      ProgressBar1.Value = ProgressBar1.Value + 1
'      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
'      DoEvents
      
      MySize = FileLen(strTemp(0))
      Print #ff1, Left(ii + 1 & "     ", 5) & List1.List(ii) & " " & MySize
   Next ii
   
   TerminateProcess process_handle, 0&
   CloseHandle process_handle
   
   Print #ff1, "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   Print #ff1, "列印時間：" & strTime & "  " & time()
   Print #ff1, "檔案數量：" & intFileCnt
   Close #ff1
   
'   ProgressBar1.Visible = False
'   lblProgress.Visible = False
   Exit Sub

ShellError:
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub

Private Sub PrintOnePdf(ByVal program_name As String, parameters As String)

Dim process_id As Long
Dim process_handle As Long
    ' Start the program.
    On Error GoTo ShellError
    
    process_id = SHELL(program_name & parameters, vbHide)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
    
    Exit Sub

ShellError:
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub

'Added by Lydia 2024/10/07
Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse txtData(Index)
End Sub

'Added by Lydia 2024/10/07
Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2024/10/07
Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
Dim strTmpA As String

   Select Case Index
      Case 0
         lblFM2 = ""
         If Trim(txtData(Index)) <> "" Then
            strTmpA = GetStaffName(Trim(txtData(Index)))
            If strTmpA = "" Then
               MsgBox "請輸入管制人！", vbExclamation + vbOKOnly
               GoTo EXITSUB
            Else
               lblFM2 = strTmpA
            End If
         End If
   End Select
   
   Exit Sub
   
EXITSUB:
   Cancel = True
   txtData(Index).SetFocus
   Txtdata_GotFocus Index
End Sub
