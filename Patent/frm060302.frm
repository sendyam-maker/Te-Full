VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm060302 
   BorderStyle     =   1  '單線固定
   Caption         =   "公告通知函"
   ClientHeight    =   6144
   ClientLeft      =   2796
   ClientTop       =   3948
   ClientWidth     =   6336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6144
   ScaleWidth      =   6336
   Begin VB.CheckBox chkAddMemo 
      Caption         =   "公報有誤加註定稿"
      Height          =   225
      Left            =   1050
      TabIndex        =   36
      Top             =   1140
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   14
      Left            =   3195
      MaxLength       =   2
      TabIndex        =   31
      Top             =   510
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   13
      Left            =   2940
      MaxLength       =   1
      TabIndex        =   32
      Top             =   510
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   11
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   34
      Text            =   "FCP"
      Top             =   510
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   12
      Left            =   2115
      MaxLength       =   6
      TabIndex        =   33
      Top             =   510
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "更新核對已准專利期限(&U)"
      Height          =   585
      Index           =   2
      Left            =   4590
      TabIndex        =   30
      Top             =   570
      Width           =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   345
      Left            =   6030
      TabIndex        =   13
      Top             =   2550
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CheckBox Check4 
      Caption         =   "列印定稿"
      Height          =   345
      Left            =   3540
      TabIndex        =   10
      Top             =   2130
      Value           =   1  '核取
      Width           =   1065
   End
   Begin VB.ComboBox cmbPrinter3 
      Height          =   300
      Left            =   1905
      TabIndex        =   15
      Top             =   3240
      Width           =   4395
   End
   Begin VB.CheckBox Check3 
      Caption         =   "列印承辦單"
      Height          =   345
      Left            =   1950
      TabIndex        =   9
      Top             =   2130
      Value           =   1  '核取
      Width           =   1245
   End
   Begin VB.TextBox txtLetterDate 
      Height          =   264
      Left            =   4890
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1680
      Width           =   1035
   End
   Begin VB.CheckBox Check2 
      Caption         =   "列印公報"
      Height          =   345
      Left            =   4890
      TabIndex        =   11
      Top             =   2130
      Width           =   1335
   End
   Begin VB.TextBox txtPath2 
      Height          =   315
      Left            =   1905
      TabIndex        =   12
      Top             =   2550
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.TextBox txtPDFPath 
      Height          =   315
      Left            =   1905
      TabIndex        =   16
      Text            =   "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Top             =   3570
      Width           =   4395
   End
   Begin VB.ComboBox cmbPrinter2 
      Height          =   300
      Left            =   1905
      TabIndex        =   14
      Top             =   2910
      Width           =   4395
   End
   Begin VB.ListBox List1 
      Height          =   1308
      ItemData        =   "frm060302.frx":0000
      Left            =   105
      List            =   "frm060302.frx":0007
      TabIndex        =   22
      Top             =   4530
      Width           =   6195
   End
   Begin VB.CheckBox Check1 
      Caption         =   "列印定稿清單"
      Height          =   345
      Left            =   165
      TabIndex        =   8
      Top             =   2130
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   660
      Left            =   60
      TabIndex        =   20
      Top             =   1410
      Width           =   3435
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   7
         Top             =   240
         Width           =   2520
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   21
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   4
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   5
      Top             =   828
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   3
      Left            =   2625
      MaxLength       =   1
      TabIndex        =   4
      Top             =   828
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   1785
      MaxLength       =   6
      TabIndex        =   3
      Top             =   828
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   1305
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "FCP"
      Top             =   828
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   0
      Left            =   1305
      MaxLength       =   7
      TabIndex        =   0
      Top             =   180
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   75
      TabIndex        =   1
      Top             =   864
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      Caption         =   "公告日："
      Height          =   180
      Index           =   0
      Left            =   75
      TabIndex        =   19
      Top             =   225
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5385
      TabIndex        =   18
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4590
      TabIndex        =   17
      Top             =   90
      Width           =   756
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   105
      TabIndex        =   23
      Top             =   3930
      Visible         =   0   'False
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "重印起始案號："
      Height          =   180
      Left            =   360
      TabIndex        =   35
      Top             =   570
      Width           =   1260
   End
   Begin VB.Label Label4 
      Caption         =   "定稿、承辦單印表機:"
      Height          =   180
      Left            =   135
      TabIndex        =   29
      Top             =   3300
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "定稿日期："
      Height          =   180
      Index           =   0
      Left            =   3945
      TabIndex        =   28
      Top             =   1725
      Width           =   900
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
      Left            =   105
      TabIndex        =   27
      Top             =   4230
      Visible         =   0   'False
      Width           =   6180
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "公報PDF的存放路徑："
      Height          =   180
      Left            =   135
      TabIndex        =   26
      Top             =   2640
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PDF執行檔路徑："
      Height          =   180
      Left            =   135
      TabIndex        =   25
      Top             =   3630
      Width           =   1560
   End
   Begin VB.Label Label6 
      Caption         =   "公報印表機："
      Height          =   180
      Left            =   135
      TabIndex        =   24
      Top             =   2970
      Width           =   1755
   End
End
Attribute VB_Name = "frm060302"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim intWhere As Integer
'Add By Cheng 2003/01/28
Dim m_OriPrinterName As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer
'Add By Cheng 2003/02/12
Dim strReceiveNo As String '本所案號
'Add by Morgan 2011/3/15
Dim strPrinter As String, strPrinter2 As String, strPrinter3 As String

'Added by Morgan 2012/5/31
Dim strTPB04 As String, strTPB05 As String
Dim strTime As String
Dim strSpecNO As Boolean   'add by sonia 2014/4/25
Dim strSpecNO_2 As Boolean 'Add By Sindy 2015/11/16
Dim strEmail As String     'add by sonia 2014/4/25
'End 2012/5/31
Dim strPath As String
Dim m_LetterLanguage As String 'Add By Sindy 2017/3/15
Dim bolM926 As Boolean 'Added by Lydia 2019/03/05 記錄核對已准專利的筆數
'Added by Lydia 2019/12/03 外部呼叫(單筆) : 傳入本所案號和定稿日期
Public m_KeyCP01 As String
Public m_KeyCP02 As String
Public m_KeyCP03 As String
Public m_KeyCP04 As String
Public m_KeyDate As String
Dim m_AttachPath As String 'Added by Morgan 2021/6/25 公報PDF暫存路徑

'Modified by Lydia 2019/12/03 改成共用
'Private Sub cmdok_Click(Index As Integer)
Public Sub cmdok_Click(Index As Integer)
'edit by nickc 2007/02/06 不用 dll 了
'Dim objPrintDllPublic As New clsPrintPublic
Dim strTmp As String
'Dim rsTemp1 As New ADODB.Recordset 'Remove by Morgan 2006/6/6 沒用了
Dim rsTemp2 As New ADODB.Recordset
Dim stET03 As String
Dim iCopy As Integer
Dim Cancel As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean 'Added by Morgan 2014/9/19
'Dim bolEtype As Boolean 'Add by Lydia 2015/01/19
Dim strStarTime As String 'Add By Sindy 2015/7/7
Dim strSign As String
'Added by Morgan 2017/11/13
Dim program_name As String
Dim process_id As Long
Dim process_handle As Long
'end 2017/11/13
'Added by Lydia 2019/12/03
Dim strFileName As String
Dim fs, f, s
Dim m_iCopy As Integer 'Added by Lydia 2019/12/11

   Select Case Index
      Case 0 '確定
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
         
         'Added by Morgan 2013/4/24
         g_LetterDate = DBDATE(txtLetterDate)
         If txtLetterDate <> "" Then
            txtLetterDate_Validate Cancel
            If Cancel = True Then Exit Sub
         End If
         'end 2013/4/24
         
         If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 Then
            MsgBox "列印項目只少要勾選一項 !", vbCritical
            Check1.SetFocus
            Exit Sub
         End If
         
         '公告日
         If Option1(0).Value = True Then
            If Text1(0).Text <> "" Then
               If Not ChkDate(Text1(0).Text) Then
                  Text1(0).SetFocus
                  TextInverse Text1(0)
                  Exit Sub
               End If
            Else
               MsgBox "公告日不得空白，請重新輸入 !", vbCritical
               Text1(0).SetFocus
               Exit Sub
            End If
            
            'Modify By Cheng 2002/12/20
'            strExc(0) = "SELECT PA01||PA02||PA03||PA04,DECODE(TPB08,'台一國際專利法律事務所',1,0) FROM PATENT,TPBULLETIN WHERE PA01='FCP' AND PA14=" & TransDate(Text1(0).Text, 2) & _
'               " AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPB01(+)"
            
            pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Text1(0) 'Add By Sindy 2010/12/7
            
            If Check1.Value = 1 Then
               pub_QL05 = pub_QL05 & ";" & Check1.Caption 'Add By Sindy 2010/12/7
            End If
            
            'Modify by Morgan 2004/8/12 加排序,專利種類、證書號數
            'Modified by Morgan 2012/5/31 +pa11,改用本所號排序(同定稿)
            'modify by sonia 2014/4/25 +pa26,fa16 特定客戶/代理人信函加印代理人e-mail
            'Modify By Sindy 2015/7/7 +,GetEmailFlag(PA01||PA02||PA03||PA04) eMail
            'Modify By Sindy 2015/11/25 +,pa141
            'Modify By Sindy 2017/4/10 + order by eMail,PA01,PA02,PA03,PA04 ==> order by PA01,PA02,PA03,PA04
            strExc(0) = "SELECT PA01||PA02||PA03||PA04,DECODE(TPB08,'台一國際',1,0),PA01,PA02,PA03,PA04,PA75,pa11,pa14,pa26,fa16,GetEmailFlag(PA01||PA02||PA03||PA04) eMail,pa141 FROM PATENT,TPBULLETIN,fagent WHERE PA01='FCP' AND PA14=" & TransDate(Text1(0).Text, 2) & _
               " AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPB01(+) and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+)" & _
               " order by PA01,PA02,PA03,PA04"
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
            
            'Modified by Morgan 2012/5/31 +pa11
            'modify by sonia 2014/4/25 +pa26,fa16 特定客戶/代理人信函加印代理人e-mail
            'Modify By Sindy 2015/7/7 +,GetEmailFlag(PA01||PA02||PA03||PA04) eMail
            'Modify By Sindy 2015/11/25 +,pa141
            strExc(0) = "SELECT PA01||PA02||PA03||PA04,DECODE(TPB08,'台一國際',1,0),PA01,PA02,PA03,PA04,PA75,PA14,pa11,pa26,fa16,GetEmailFlag(PA01||PA02||PA03||PA04) eMail,pa141 FROM PATENT,TPBULLETIN,fagent WHERE " & ChgPatent(strTmp) & _
               " AND (PA57<>'Y' OR PA57 IS NULL) AND PA11=TPB01(+) and pa14>0 and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+)"
         End If
         intI = 1
         Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         Screen.MousePointer = vbHourglass
         If intI = 1 Then
            'Add By Sindy 2015/7/7
            pub_OsPrinter = PUB_GetOsDefaultPrinter
            PUB_SetOsDefaultPrinter cmbPrinter3
            PUB_SetWordActivePrinter
            PUB_RestorePrinter cmbPrinter3
            '2015/7/7 END
            
            List1.Clear 'Added by Morgan 2012/5/31
            strStarTime = Format(ServerTime, "##:##:##") 'Add By Sindy 2015/7/7
            
            ProgressBar1.max = rsTemp2.RecordCount
            ProgressBar1.Value = 0
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            ProgressBar1.Visible = True
            lblProgress.Visible = True
            DoEvents
            
            'Added by Morgan 2012/5/31
            'If Me.Check1.Value = vbUnchecked And Check2.Value = vbChecked Then
            If Check2.Value = vbChecked Then
               'Modified by Morgan 2016/6/29
               'If Text1(0) = "" Then Text1(0) = "" & rsTemp2.Fields("pa14") 'Add By Sindy 2015/7/30
               If Text1(0) = "" Then Text1(0) = TransDate("" & rsTemp2.Fields("pa14"), 1) 'Add By Sindy 2015/7/30
               'end 2016/6/29
               
               'Removed by Morgan 2021/6/25 公報改抓卷宗區，不再往pat3讀取避免當機沒開的情形
               'If GetFilePath(DBDATE(Text1(0))) = False Then
               '   Me.txtPath2.SetFocus
               '   Screen.MousePointer = vbDefault
               '   Exit Sub
               'End If
               'end 2021/6/25
               
               'Added by Morgan 2017/11/13
               program_name = txtPDFPath
               process_id = SHELL(program_name, vbHide)
               process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
               'end 2017/11/13
            End If
            'end 2012/5/31
            
            With rsTemp2
               InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/7
               Do While Not .EOF
                  'Added by Morgan 2016/5/25
                  If Option1(0) And Text1(12) <> "" Then
                     If .Fields(0) <= Text1(11) & Text1(12) & Text1(13) & Text1(14) Then
                        GoTo NextCase
                     End If
                  End If
               
                  intI = 1
                  '處理定稿例外欄位
                  strReceiveNo = "" & .Fields(0).Value
                  
                  'Add By Sindy 2017/3/15 定稿語文
                  m_LetterLanguage = PUB_GetLanguage(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))
                  
                  'add by sonia 2014/4/25
                  strSpecNO = False: strEmail = ""
                  strSpecNO_2 = False
                  'MODIFY BY SONIA 2014/5/9 取消Y52218再加入X47833,X47833020,X17901010
                  'If "" & .Fields("PA75") = "Y52218000" Or "" & .Fields("PA75") = "Y20085000" Or "" & .Fields("PA26") = "X34291000" Or "" & .Fields("PA26") = "X21382010" Then
                  'Modified by Morgan 2014/9/19
                  '公報改E化的定稿要印Email故只要控制沒有設E化的加印就好
                  'Select Case "" & .Fields("PA75")
                  '   'Modified by Morgan 2014/9/19
                  '   Case "Y20085000", "X34291000", "X21382010", "X47833000", "X47833020", "X17901010"
                  '      strSpecNO = True
                  '      strEmail = "" & .Fields("fa16")
                  'End Select
                  If .Fields("pa26") = "X47833000" Or .Fields("pa26") = "X47833020" Then
                     strSpecNO = True
                     strEmail = "" & .Fields("fa16")
                  End If
                  bolEmail = PUB_GetEMailFlag(.Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04"), , , bolPlusPaper)
                  'end 2014/9/19
                  '2014/4/25 end
                  
                  'Add By Sindy 2015/11/16
                  'Removed by Morgan 2019/2/26 定稿已修改,取消 -- Sharon,David
                  'If .Fields("pa26") = "X67402000" Or .Fields("pa26") = "X67402010" Or _
                  '   .Fields("pa26") = "X67402020" Or .Fields("pa26") = "X60507000" Or _
                  '   .Fields("pa26") = "X60507001" Or .Fields("pa26") = "X60507010" Or _
                  '   .Fields("pa26") = "X70749000" Or .Fields("pa75") = "Y22457000" Then
                  '   strSpecNO_2 = True
                  'End If
                  'end 2019/2/26
                  '2015/11/16 END
                  
                  'FCP公告通知函定稿不必調卷的案號前加印◎ 926.核對已准專利
                  strSign = ""
                  strExc(0) = "select 1 from caseprogress A where cp01='" & .Fields("pa01") & "' and cp02='" & .Fields("pa02") & "' and cp03='" & .Fields("pa03") & "' and cp04='" & .Fields("pa04") & "' and cp10='926' and cp57 is null"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 0 Then
                     strSign = "◎" '沒有核對已准專利
                  End If
                  
                  'Added by Lydia 2018/01/15 X21775 ROBERT BOSCH GMBH(不論搭配到哪一個Y編號)未來皆不做"核對已准專利",
                  If .Fields("pa26") = "X21775000" Then
                      strSign = "◎"
                  End If
                  'end 2018/01/15
                  
                  '若不只列印定稿清單或是選本所案號
                  'If Me.Check1.Value = vbUnchecked Or Option1(1).Value = True Then
                     'Add By Sindy 2015/7/7 列印承辦單
                     If Check3.Value = vbChecked Then
                        'Modified by Lydia 2019/03/04 更換類別代號;05=>02
                        Call PUB_PrintFCPEmpBill(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), "02", , , , IIf(strSpecNO, "e", IIf(bolEmail, IIf(bolPlusPaper, "E", "e"), "")) & strSign)
                     End If
                     
                     'Modified by Morgan 2017/10/17 從是否列印稿的判斷內移出上來(因為會影響公報份數,不論是否印定稿都要跑份數)
                     
                        'Modified by Morgan 2014/9/19 例外或純E化的公報也只要印1份
                        'NowPrint .Fields(0) & "&000", "05", stET03, False, strUserNum, 0
                        If strSpecNO = True Or (bolEmail = True And bolPlusPaper = False) Then
                           iCopy = 1
                        Else
                           iCopy = 0
                        End If
                        'Modify By Sindy 2015/12/8 沒有"核對已准專利"收文且是E化的定稿出2份(需與證書一同寄出)
                        'Modify By Sindy 2016/10/18 為配合證書通知函,而調整
'                        If strSign = "◎" And bolEmail = True Then
'                           iCopy = 2
'                        End If
                        If strSign = "◎" Then '沒有"核對已准專利"
                           iCopy = 2 '非E化及大E(e+寄)定稿列印2份
                           If bolEmail = True And bolPlusPaper = False Then '小e:純e化
                              iCopy = 1 '定稿列印1份
                           End If
                        'Added by Lydia 2018/09/27 印2份(Y34210000,Y34210010,Y34210020,Y34210030)
                        'Move by Lydia 2018/10/22
                        'ElseIf "" & .Fields("pa75") <> "" And InStr("Y34210000,Y34210010,Y34210020,Y34210030", .Fields("pa75")) > 0 Then
                        '      iCopy = 2
                        'end 2018/09/27
                        End If
                        '2016/10/18 END
                        '2015/12/8 END
                                                
                        m_iCopy = iCopy 'Added by Lydia 2019/12/11 先保留是否E化定稿的份數,因為會影響公報份數(核對已准專利926)的判斷
                        'Added by Lydia 2019/12/12 原本紙本iCopy=0，在列印紙本時NoPrint會再抓預設份數；現在先取得預設份數，再減一份
                        If iCopy = 0 Then
                           stET03 = GetSitu("" & .Fields("PA01"), "" & .Fields("PA02"), "" & .Fields("PA03"), "" & .Fields("PA04"), "" & .Fields("PA75"), "" & .Fields("PA14"))
                           PUB_GetCopySetting iCopy, .Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), "000", "05", stET03
                           If iCopy = 0 Then
                              iCopy = 2   '(舊制)預設紙本定稿2份，公報3份
                           End If
                        End If
                        'Added by Lydia 2019/12/03 (新制)留底的紙本直接上卷宗區 => E化不印定稿; 非E化印一份
                        iCopy = iCopy - 1
                        If iCopy < 0 Then iCopy = 0
                        
                        'Move by Lydia 2018/10/22 印2份(Y34210000,Y34210010,Y34210020,Y34210030) ; 從上方移下來( ex.FCP-50611只印到1份)
                        If "" & .Fields("pa75") <> "" And InStr("Y34210000,Y34210010,Y34210020,Y34210030", .Fields("pa75")) > 0 Then
                             'Modified by Lydia 2019/12/03 Y34210(NGB) 不印定稿1
                             'iCopy = 2
                             iCopy = 0
                             m_iCopy = iCopy 'Added by Lydia 2019/12/11 先保留定稿的份數,因為會影響公報份數(核對已准專利926)的判斷
                        End If
                        'end 2018/10/22
                        
                        'Modify By Sindy 2018/5/22 Y51982只印一份
'                        'Add By Sindy 2015/12/28 特殊定稿份數
'                        If .Fields("pa75") = "Y51982000" Then
'                           iCopy = 2
'                        End If
'                        '2015/12/28 END
                        
                     'end 2017/10/17
                     
                     'Add By Sindy 2017/3/15 Y51333010北京銀龍定稿語文1.中文時,不出公報信函
                     '列印定稿
                     'Modified by Lydia 2019/02/27 Y55054000不出中文定稿(參考Y51333010北京銀龍設定)
                     'If Check4.Value = vbChecked And _
                        Not (.Fields("pa75") = "Y51333010" And m_LetterLanguage = "1") Then
                     If Check4.Value = vbChecked And _
                        Not ("" & .Fields("pa75") <> "" And InStr("Y51333010,Y55054000", "" & .Fields("pa75")) > 0 And m_LetterLanguage = "1") Then
                        'Modify by Morgan 2006/4/10 加判斷代理人Y49575,Y45697,Y48309出不同定稿
                        'Modify by Morgan 2006/6/6
                        'stET03 = GetSitu(strReceiveNo, "" & .Fields("PA75"))
                        stET03 = GetSitu("" & .Fields("PA01"), "" & .Fields("PA02"), "" & .Fields("PA03"), "" & .Fields("PA04"), "" & .Fields("PA75"), "" & .Fields("PA14"))
                        StartLetter "05", stET03
                        If iCopy > 0 Then 'Added by Lydia 2019/12/03 先處理列印紙本
                              NowPrint .Fields(0) & "&000", "05", stET03, False, strUserNum, , , , , iCopy
                        'Added by Lydia 2019/12/03 先處理列印紙本
                              PUB_PrintLetter .Fields(0) & "&000"
                        End If
                        'end 2020/12/03
                        
                        'Added by Lydia 2019/12/03 產生PDF檔,直接上卷宗區
                        'Modified by Lydia 2019/12/04 流水號補足
                        'strFileName = "$$" & .Fields("pa01") & Val(.Fields("pa02")) & IIf(.Fields("pa03") & .Fields("pa04") <> "000", .Fields("pa03") & .Fields("pa04"), "") & ".1228.CUS.PDF"
                        strFileName = "$$" & .Fields("pa01") & .Fields("pa02") & IIf(.Fields("pa03") & .Fields("pa04") <> "000", .Fields("pa03") & .Fields("pa04"), "") & ".1228.CUS.PDF"
                        'Added by Lydia 2020/01/02 原本放在Typing2: 參考'Add By Sindy 2015/7/30 不要加confirmation的章
                        If bolEmail = True Then  'E化+E加寄
                             strUserLevel = "發FC郵件" '不要加confirmation的章
                        End If
                        'end 2020/01/02
                        'Modified by Lydia 2020/07/17 參考2019/4/3的備份程式,設定稿存DB,因為前面的"先處理列印紙本";
                                         'FCP-52350沒有收二核但是發明人有錯,需要用定稿維護Word修改,才發現定稿沒有存DB;
                        'NowPrint .Fields(0) & "&000", "05", stET03, True, strUserNum, , , , , iCopy, , True, , False, , , , , , True
                        NowPrint .Fields(0) & "&000", "05", stET03, True, strUserNum, , , , , iCopy, , True, , , , , , , , True
                        
                        strUserLevel = "" 'Added by Lydia 2020/01/02 還原設定
                        'Modified by Lydia 2023/04/27 改模組
                        'If PUB_PrintWord2PDF(g_WordAp, App.path, strFileName, "", cmbPrinter3) = True Then
                        If PUB_PrintWord2File(g_WordAp, App.path, strFileName) = True Then
                             '上卷宗區
                             strExc(0) = "SELECT CP09,CPP02 FROM CASEPROGRESS, (SELECT CPP01,CPP02 FROM CASEPAPERPDF WHERE INSTR(UPPER(CPP02),'.1228.CUS.PDF') > 0 ) X1" & _
                                              " WHERE CP01='" & .Fields("pa01") & "' AND CP02='" & .Fields("pa02") & "' AND CP03='" & .Fields("pa03") & "' AND CP04='" & .Fields("pa04") & "' AND CP159=0 AND CP10='1228' AND CP09=CPP01(+) "
                             intI = 1
                             strExc(1) = "":   strExc(2) = ""
                             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                             If intI = 1 Then
                                 strExc(1) = "" & RsTemp.Fields("cp09") '收文號1228
                                 strExc(2) = "" & RsTemp.Fields("cpp02")
                             End If
                             
                             'Modified by Lydia 2021/01/21 debug
                             'If strExc(1) <> "" And strExc(2) <> "" Then 'Added by Lydia 2021/01/20 增加有公告公報進度,才上傳檔案
                             If strExc(1) <> "" Then
                                If Dir(App.path & "\" & strFileName) <> "" Then
                                     '先刪除原檔
                                     If strExc(2) <> "" Then
                                         'Modified by Lydia 2019/12/05 直接刪除檔案 (只保留最新的通知函)
                                         If DelAttFile_PDF(.Fields("pa01") & "-" & .Fields("pa02") & "-" & .Fields("pa03") & "-" & .Fields("pa04"), strExc(1), strExc(2), , , True) = False Then
                                              MsgBox "刪除卷宗區檔案失敗：" & vbCrLf & strExc(2), vbCritical, "上傳卷宗區作業"
                                              Exit Sub
                                         End If
                                     End If
                                     Set fs = CreateObject("Scripting.FileSystemObject")
                                     Set f = fs.GetFile(App.path & "\" & strFileName)
                                     '檔案大小為 0 KB 有誤
                                     If f.Size = 0 Then
                                        ShowMsg App.path & "\" & strFileName & MsgText(9221)
                                        Exit Sub
                                     End If
                                     If SaveAttFile_PDF(strExc(1), App.path & "\" & strFileName, Replace(strFileName, "$$", ""), Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), IIf(UCase(Right(strFileName, 4)) = ".PDF", False, True), "A") = True Then
                                          Call PUB_DelPCOrgFile(App.path & "\" & strFileName)  '刪除實體檔案
                                     End If
                                End If
                             'Added by Lydia 2021/01/20
                             Else
                                     Call PUB_DelPCOrgFile(App.path & "\" & strFileName)  '刪除實體檔案
                             End If
                             'end 2021/01/20
                        End If
                        'end 2019/12/03
                        
                        'Add by Lydia 2015/01/19 E化案件信函自動產生有簽名檔的pdf檔 =>+E加寄
                        'Modified by Lydia 2015/02/05 指定英文定稿
                        'Modify By Sindy 2015/9/8 不管英日文只要E化案件,均要產生及存檔
                        'If bolEmail = True And stET03 = "04" Then 'E化+E加寄
                        'Remove by Lydia 2019/12/03 產生PDF檔,原本放在Typing2,改成直接上卷宗區
'                        If bolEmail = True Then  'E化+E加寄
'                        '2015/9/8 END
''                           bolEtype = True
'                           strUserLevel = "發FC郵件" 'Add By Sindy 2015/7/30 不要加confirmation的章
'                           'Modify By Sindy 2019/11/22 bolSave2DB = False : 必免重覆刪除定稿
'                           NowPrint .Fields(0) & "&000", "05", stET03, True, strUserNum, , , , , iCopy, , True, , False, , , , , , True
'                           strUserLevel = "" 'Add By Sindy 2015/7/30
'                           PrintWord2PDF .Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04")
'                           'Add By Sindy 2015/7/7 因上面PrintWord2PDF函數中會再切印表機
'                           'pub_OsPrinter = PUB_GetOsDefaultPrinter
'                           'Removed by Morgan 2017/11/13 移到PrintWord2PDF內
'                           'PUB_SetOsDefaultPrinter cmbPrinter3
'                           'PUB_SetWordActivePrinter
'                           'PUB_RestorePrinter cmbPrinter3
'                           'end 2017/11/13
'                           '2015/7/7 END
''                        Else
''                           bolEtype = False
'                         End If
                        'end 2015/01/19
                        'end 2014/9/19
                        'PUB_PrintLetter .Fields(0) & "&000" 'Add By Sindy 2015/7/7 直接列印定稿 'Remove by Lydia 2019/12/03
                     End If
                     
                     'Added by Morgan 2012/5/30
                     '列印公報
                     If Check2.Value = vbChecked Then
                        'Modified by Morgan 2012/7/5
                        'iCopy = 3
                        'Added by Morgan 2014/9/22 例外或純E化的公報也只要印1份
                        'iCopy = 0
                        'Add By Sindy 2015/11/25 不核對已准專利且E,e化的案件,公報均列印3份
'                        If "" & .Fields("pa141") = "N" And bolEmail = True Then
'                           iCopy = 3
'                        End If
                        'Modify By Sindy 2015/12/8 沒有"核對已准專利"收文且是E化的公報出3份(需與證書一同寄出)
                        'Modify By Sindy 2016/10/18 為配合證書通知函,而調整
'                        If strSign = "◎" And bolEmail = True Then
'                           iCopy = 3
'                        End If
                        
                        iCopy = m_iCopy 'Added by Lydia 2019/12/11 傳入先保留是否E化定稿的份數,因為會影響公報份數(核對已准專利926)的判斷
                        
                        If strSign = "◎" Then '沒有"核對已准專利"
                           iCopy = 3 '非E化及大E(e+寄)公報列印3份
                           If bolEmail = True And bolPlusPaper = False Then '小e:純e化
                              iCopy = 1 '公報列印1份
                           End If
                        End If
                        '2016/10/18 END
                        '2015/11/25 END
                        If iCopy = 0 Then
                        'end 2014/9/22
                           PUB_GetCopySetting iCopy, .Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04")
                           If iCopy = 0 Then
                              iCopy = 3
                           'Modify By Sindy 2015/12/8 Mark:淑華說取消不須加1份
'                           ElseIf iCopy = 1 Then
'                              iCopy = 2
                           End If
                        End If 'Added by Morgan 2014/9/19
                        'end 2012/7/5
                        
                        'Add By Sindy 2015/12/28 特殊公報份數
                        'Modified by Morgan 2018/3/21 Y21775000 Robert Bosch E化案件只要印1份--淑華
                        'If .Fields("pa75") = "Y21775000" Or .Fields("pa75") = "Y52922000" Then
                        'Modified by Lydia 2018/09/27 印2份(Y34210000,Y34210010,Y34210020,Y34210030)
                        'If (.Fields("pa75") = "Y21775000" And iCopy <> 1) Or .Fields("pa75") = "Y52922000" Then
                        'Modified by Lydia 2020/03/24 取消Y52922000立石國際的特殊公報份數設定，改成一般規則
                        'If (.Fields("pa75") = "Y21775000" And iCopy <> 1) Or ("" & .Fields("pa75") <> "" And InStr("Y52922000,Y34210000,Y34210010,Y34210020,Y34210030", .Fields("pa75")) > 0) Then
                        If (.Fields("pa75") = "Y21775000" And iCopy <> 1) Or ("" & .Fields("pa75") <> "" And InStr("Y34210000,Y34210010,Y34210020,Y34210030", .Fields("pa75")) > 0) Then
                           iCopy = 2
                        'Modify By Sindy 2017/3/15 Y51333010北京銀龍定稿語文1.中文時,智慧局公報要印3份
                        'Modify By Sindy 2018/5/22 Y51982只印一份
'                        ElseIf .Fields("pa75") = "Y51982000" Or _
'                           (.Fields("pa75") = "Y51333010" And m_LetterLanguage = "1") Then
                        ElseIf (.Fields("pa75") = "Y51333010" And m_LetterLanguage = "1") Then
                        '2018/5/22 END
                           iCopy = 3
                        End If
                        '2015/12/28 END
                        
                        'Added by Lydia 2019/12/03 公報整份: 原本3份減掉留底1份、特殊客戶依原設定的份數減1
                                                                 'Y34210(NGB): 設E化公報仍印出一份(公報需印出紙本用掃描給客戶)
                        If iCopy > 0 Then iCopy = iCopy - 1
                        
                        'Modified by Morgan 2013/1/9 +pa14
                        'Add by Lydia 2015/01/19 + etype
                        'Modify By Sindy 2015/9/8 不管英日文只要E化案件,均要產生及存檔
                        'GetPDFCopys .Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), .Fields("pa11"), iCopy, .Fields("pa14"), IIf((bolEmail = True And stET03 = "04"), True, False)
                        If iCopy > 0 Then 'Added by Lydia 2019/12/03 加判斷
                             GetPDFCopys .Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"), .Fields("pa11"), iCopy, .Fields("pa14"), bolEmail
                        End If
                        '2015/9/8 END
                     End If
                     'end 2012/5/30
                           
                     '新增地址條列表資料
                     'Removed by Morgan 2012/5/31 與證書一起無需再印--敏莉
                     'pub_AddressListSN = pub_AddressListSN + 1
                     'PUB_AddNewAddressList strUserNum, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, "" & pub_AddressListSN, "0"
                  'End If
                  
                  '新增整批定稿列印清單資料
                  If Option1(0).Value = True And Check1.Value = vbChecked Then 'Added by Morgan 2015/4/29 改沒勾選時不要印--敏莉
                     PUB_AddNewLetterList "公告通知函", Me.Text1(0).Text, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, IIf(strSpecNO, "ｅ", IIf(bolEmail, IIf(bolPlusPaper, "Ｅ", "ｅ"), ""))
                  End If
NextCase:
                  ProgressBar1.Value = ProgressBar1.Value + 1
                  lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
                  DoEvents
                  
                  .MoveNext
               Loop
                  
               'Added by Morgan 2017/11/13
               If process_handle <> 0 Then
                  TerminateProcess process_handle, 0&
                  CloseHandle process_handle
               End If
               'end 2017/11/13
            End With
            'Add By Sindy 2015/7/7
'            If Option1(0).Value = True Then '整批
'               PUB_SendMail strUserNum, "97038", "", "外專執行＜公告通知函＞整批的執行時間: " & strStarTime & " ~ " & Format(ServerTime, "##:##:##"), "如主旨"
'            End If
            '2015/7/7 END
            
'            'Added by Morgan 2012/5/31
'            If List1.ListCount > 0 Then
'               Call PrinBatchPdf
'               MsgBox "列印完成 ! (列印PDF花費時間：" & strTime & "  " & time() & ")", vbInformation
'            Else
'            'end 2012/5/31
               
            'Add By Sindy 2015/7/7
            PUB_SetOsDefaultPrinter pub_OsPrinter
            PUB_RestorePrinter strPrinter3
            '2015/7/7 END
            
            'Modified by Lydia 2019/12/03 排除外部呼叫
            'MsgBox "列印結束 !", vbInformation
            If m_KeyCP01 = "" Then MsgBox "列印結束 !", vbInformation
         Else
            InsertQueryLog (0) 'Add By Sindy 2010/12/7
            MsgBox "無符合條件之資料可列印 !", vbInformation
         End If
         
         Screen.MousePointer = vbDefault
         
      Case 1 '結束
         Me.Enabled = False
         Unload Me
         
      'Added by Morgan 2016/3/22
      Case 2 '更新核對已准專利期限
         '公告日
         If Option1(0).Value = True Then
            If Text1(0).Text = "" Then
               MsgBox "請輸入公告日!", vbCritical
               Text1(0).SetFocus
               Exit Sub
            ElseIf Not ChkDate(Text1(0).Text) Then
               Text1(0).SetFocus
               Exit Sub
            End If
         Else
            MsgBox "請點選公告日並輸入日期!", vbCritical
            Exit Sub
         End If
         If FormSave = True Then PrintList
   End Select
End Sub

'Added by Morgan 2016/3/22
'更新核對已准專利期限
Private Function FormSave() As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim bolOK As Boolean
   Dim strCP48 As String, strCP14 As String, strCP06 As String
   
On Error GoTo ErrHnd
   'Modified by Lydia 2016/06/17 +cp01,cp10
   'Modified by Morgan 2020/2/26 +distinct(多次列印公告清單案號也會重複)
   'Modified by Lydia 2025/04/25 排除特定代理人：特定客戶優先二核期限控管; 日代<Y4520400> SOEI、<Y5518900>TOKOSHIE
   stSQL = "select distinct cp09,cp14,st04,st02,oMan,cp01,cp10 from LetterList,caseprogress,patent,staff,SetSpecMan" & _
      " where ll01='F4102' and LL02='公告通知函-調卷清單' and ll03='" & Text1(0) & "'" & _
      " and cp01(+)=LL04 and cp02(+)=LL05 and cp03(+)=LL06 and cp04(+)=LL07 and cp10='926' and cp57||cp27 is null" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and st01(+)=cp14" & _
      " and OCODE(+)=decode(pa150,'1','T','2','R','3','S','4','T1','N') and instr(pa75,'Y4520400') = 0 and instr(pa75,'Y5518900') = 0 "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Do While Not .EOF
         '承辦期限=本所=法定=系統日+12個工作天(分案也要同步修改)
         'Modified by Lydia 2016/06/17 改成模組
         'strCP48 = CompWorkDay(12, strSrvDate(1), 0)
         strCP48 = PUB_GetFCPsetDate(.Fields("cp01"), .Fields("cp10"))
         '無承辦人 或 承辦人已離職 或 林信昌承辦 的都改為案件組別的主管
         If IsNull(.Fields("cp14")) Or .Fields("st04") = "2" Or InStr("" & .Fields("st02"), "林信昌") > 0 Then
            strCP14 = "'" & .Fields("oMan") & "'"
         Else
            strCP14 = "cp14"
         End If
         'Add By Sindy 2021/8/12 本所期限=承辦期限＋5個工作天
         strCP06 = TransDate(PUB_GetFCPOurDeadline(strCP48, , , , "N"), 2)
         '2021/8/12 END
         strSql = "update caseprogress set cp14=" & strCP14 & ",cp48=" & strCP48 & ",CP06=" & strCP06 & " where cp09='" & .Fields("cp09") & "'"
         cnnConnection.Execute strSql, intQ
         .MoveNext
      Loop
      End With
      FormSave = True
   Else
      MsgBox "無法讀取 [調卷清單] 資料，請確認該公告日是否有需調卷案件!", vbCritical
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   Set rsQuery = Nothing
End Function

'Added by Morgan 2016/3/22
'列印調卷清單(含承辦人,承辦期限,取消註記)
Private Sub PrintList()
   Dim iNo As Integer
   Dim iPrint As Integer
   
   '排序照調卷清單=專利種類+申請號
   'Modified by Lydia 2016/08/15 +本所案號 LL04~LL07
   'Modified by Morgan 2020/2/26 改寫語法(多次列印公告清單案號也會重複)
   strExc(0) = "select pa11,pa22,LL04||'-'||LL05||decode(LL06||LL07,'000','','-'||LL06||'-'||LL07) CNo,st02,decode(cp57,null,'　','◎') Tag,sqldatet(cp48) Dt,LL04,LL05,LL06,LL07" & _
      " from (select distinct L.* from LetterList L where ll01='F4102'" & _
      " and LL02='公告通知函-調卷清單' and ll03='" & Text1(0) & "') X,patent,caseprogress,staff" & _
      " where pa01(+)=LL04 and pa02(+)=LL05 and pa03(+)=LL06 and pa04(+)=LL07" & _
      " and cp01(+)=LL04 and cp02(+)=LL05 and cp03(+)=LL06 and cp04(+)=LL07 and cp10(+)='926' and st01(+)=cp14" & _
      " order by pa08,pa11"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox "準備列印調卷清單!!!", vbExclamation + vbOKOnly
      If cmbPrinter3 <> "" Then PUB_RestorePrinter cmbPrinter3
     
RePrint:
      Printer.Orientation = 1
      Printer.Font.Name = "細明體"
      Call PrintListHead(iPrint)
      iNo = 0
      
      With RsTemp
      .MoveFirst
      Do While Not .EOF
         iNo = iNo + 1
         strExc(1) = Format(iNo, "@@") & "." & .Fields("pa11") & "(" & .Fields("pa22") & ")" & "  " & .Fields("Tag") & .Fields("CNo") & "." & Left(iNo & " ", 2) & "    " & .Fields("st02") & "    " & .Fields("Dt")
         'Added by Lydia 2016/08/18 加註
         If PUB_ChkPA70kind(.Fields("LL04"), .Fields("LL05"), .Fields("LL06"), .Fields("LL07"), "N") Then
            strExc(1) = strExc(1) & "    寄證書後年費不續辦"
         End If
         'end 2016/08/15
         
         iPrint = iPrint + 300
         If iPrint > Printer.ScaleHeight - 1000 Then
            Printer.NewPage
            iPrint = 1000
         End If
         Printer.CurrentX = 500
         Printer.CurrentY = iPrint
         Printer.Print strExc(1)
         .MoveNext
      Loop
      End With
      Printer.EndDoc
      
      If MsgBox("整批定稿清單已列印完畢，您是否要重新列印???", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
          GoTo RePrint
      End If
      If strPrinter3 <> "" Then PUB_RestorePrinter strPrinter3
   End If
End Sub

Private Sub PrintListHead(ByRef iPrint As Integer)
   Dim lngX As Long
   
   iPrint = 500
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
   Printer.CurrentX = 8500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   'Modified by Morgan 2020/2/26 +(分案用)
   strExc(1) = "定稿名稱：公告通知函-調卷清單(分案用)　　條件：" & Text1(0) & "　　◎表示已取消收文"
   Printer.Print strExc(1)
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
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
   'Add By Cheng 2003/02/05
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim ii As Integer
   
   MoveFormToCenter Me
   intWhere = 國外_FC
   Option1_Click 0
   'Modify by Morgan 2011/3/15 改共用且不要排除預設印表機
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   'end 2011/3/15

   'Added by Morgan 2012/5/31
   PUB_SetPrinter Me.Name, cmbPrinter3, strPrinter3
   PUB_SetPrinter Me.Name, cmbPrinter2, strPrinter2
   If Pub_StrUserSt03 <> "M51" Then
      'Modified by Lydia 2022/04/29
      'Me.Height = 4700
      Me.Height = 4875
   End If
   List1.Clear
   
   'Modified by Morgan 2017/10/17 改呼叫共用
   'SetFileAssociation
   txtPDFPath = PUB_SetFileAssociation
   'end 2017/10/17
   'end 2012/5/31
   
   txtLetterDate = strSrvDate(2) 'Add by Morgan 2013/4/24 定稿日期
   
   'Add By Sindy 2015/11/13 紀錄在資料庫,否則換電腦或使用者會讀不到
   'Removed by Morgan 2021/6/25 公報PDF改抓卷宗區，不再往pat3讀取避免當機沒開的情形
   'txtPath2.Text = PUB_GetLastDate(Me.Name, UCase("txtPath2"))
   'If txtPath2.Text = "" Then
   '   txtPath2.Text = "\\Pat3\GAZETTE\PXml\img_1\isu012012\"
   'End If
   'end 2021/6/25
   '2015/11/13 END
   
   'Added by Morgan 2016/5/25 先設定否則若跑定稿時才設在開另一個Word時97版的會有錯誤訊息
   'If PUB_PdfCreatorNameInWord = "" Then PUB_PdfCreatorNameInWord = PUB_GetCreatorNameInWord 'Removed by Morgan 2024/11/27 沒用了
   
   'Added by Lydia 2019/12/03 外部呼叫
   If m_KeyCP01 <> "" Then
       Me.Option1(1).Value = True
       Me.Text1(1).Text = m_KeyCP01
       Me.Text1(2).Text = m_KeyCP02
       Me.Text1(3).Text = m_KeyCP03
       Me.Text1(4).Text = m_KeyCP04
       If m_KeyDate <> "" Then Me.txtLetterDate = m_KeyDate
       Me.Check1.Value = vbUnchecked
       Me.Check2.Value = vbChecked
'Removed by Morgan 2021/6/25 公報PDF改抓卷宗區，不再往pat3讀取避免當機沒開的情形
'   Else
'        'Added by Lydia 2020/01/22 提前檢查Pat3是否開啟
'        If Pub_CheckGazetteDir(txtPath2.Text) = False Then
'            Check2.Value = vbUnchecked
'            Check2.Enabled = False
'        End If
'        'end 2020/01/22
'end 2021/6/25
   End If
   
   
   'Added by Morgan 2021/6/25 公報PDF暫存路徑
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   'end 2021/6/25
End Sub

'Added by Morgan 2016/3/22
'保留調卷記錄,後續更新核對已准專利承辦人及期限時需列印清單
Private Sub SaveList()
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
      
   '本次有資料才清除前次記錄
   'Added by Morgan 2017/12/28 排除發文日為111111者--江如玉
   'Modified by Lydia 2019/04/03 PK: 使用者帳號@電腦名稱(pub_HostName)
   'strSql = "update LetterList set LL01='X'  where LL01='" & strUserNum & "' and ll02='公告通知函' and exists(select * from caseprogress where cp01=ll04 and cp02=ll05 and cp03=ll06 and cp04=ll07 and cp10='926' and cp159=0 and cp158<>19221111)"
   strSql = "update LetterList set LL01='X'  where LL01='" & strUserNum & "@" & pub_HostName & "' and ll02='公告通知函' and exists(select * from caseprogress where cp01=ll04 and cp02=ll05 and cp03=ll06 and cp04=ll07 and cp10='926' and cp159=0 and cp158<>19221111)"
   cnnConnection.Execute strSql, intI
   If intI > 0 Then
   
      strSql = "delete LetterList where LL01='F4102' and LL02='公告通知函-調卷清單' and LL03<'" & TransDate(CompDate(1, -3, strSrvDate(1)), 1) & "'"
      cnnConnection.Execute strSql, intI
      
      strSql = "delete LetterList where LL01='F4102' and LL02='公告通知函-調卷清單' and LL03='" & Text1(0) & "'"
      cnnConnection.Execute strSql, intI
      
      strSql = "update LetterList set LL01='F4102',LL02='公告通知函-調卷清單' where LL01='X' and ll02='公告通知函'"
      cnnConnection.Execute strSql, intI
   End If
      
   cnnConnection.CommitTrans
   Exit Sub
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Sub

Private Sub Form_Unload(Cancel As Integer)

   g_LetterDate = "" 'Add by Morgan 2013/4/24
   'Copy from cmdok_Click by Morgan 2004/10/26
   '列印定稿整批列印清單
   'Modify by Morgan 2007/4/3 加傳公告通知函參數
   'PUB_PrintLetterList strUserNum
   'Modified by Lydia 2019/03/05 +列印份數 ;(下載公告本:整批清批改為2份,不印核對已准清單)
   'PUB_PrintLetterList strUserNum, 4, cmbPrinter3, strPrinter3
    PUB_PrintLetterList strUserNum, 4, cmbPrinter3, strPrinter3, , , 2
    SaveList 'Added by Morgan 2016/3/22 保留調卷清單
   
   '刪除定稿整批列印資料
   'Modified by Lydia +傳入刪除條件
   'PUB_DeleteLetterList strUserNum
   PUB_DeleteLetterList strUserNum, "and LL02='公告通知函' "
   
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
   If Me.cmbPrinter3.Text <> Me.cmbPrinter3.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter3.Name, "0", "0", Me.cmbPrinter3.Text
   End If
   
   'Added by Morgan 2014/9/22
   '公報列印完不必再留
   'Removed by Morgan 2014/10/9 改回不刪--敏莉
   'Modified by Morgan 2017/11/13 又改刪除--敏莉,淑華
   If strPath <> "" Then KillTemp
   'end 2014/10/9
   'end 2014/9/22
   
'   'Added by Lydia 2019/03/05 下載公告本email通知
'   If Check1.Value = vbChecked Then
'        Call PUB_SendMailCache
'   End If
'   'end 2019/03/05
   
   Set frm060302 = Nothing
End Sub

Private Sub KillTemp()
On Error GoTo ErrHnd
   If InStr(strPath, "\$") > 0 And Dir(strPath & "\.") <> "" Then
      Kill strPath & "\*.*"
      'Added by Morgan 2017/11/3
      Dir App.path
      RmDir strPath
      'end 2017/11/13
   'Added by Lydia 2019/12/11 刪除個案的公報檔
   ElseIf Dir(strPath & "\$$*.*") <> "" Then
      Kill strPath & "\$$*.*"
   'end 2019/12/11
   End If
   Exit Sub
   
ErrHnd:
   Resume Next
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
         If Check2.Enabled = True Then  'Added by Lydia 2020/01/22  判斷
             Check2.Value = vbChecked 'Added by Morgan 2012/6/1
         End If
         'Added by Morgan 2016/5/25
         For i = 12 To 14
            Text1(i).Enabled = True
         Next
         'end 2016/5/25
         'Added by Lydia 2022/04/29
         chkAddMemo.Enabled = False
         chkAddMemo.Value = 0
      Case 1
         For i = 2 To 4
            Text1(i).Enabled = True
         Next
         Text1(2).SetFocus
         Check2.Value = vbUnchecked 'Added by Morgan 2012/6/1
         'Added by Lydia 2022/04/29
         chkAddMemo.Enabled = True
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
            MsgBox "公告日不得空白，請重新輸入 !", vbCritical
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

'Add By Cheng 2002/12/24
'取得定稿處理方式
'Modify by Morgan 2006/4/10 加 p_stPA75
'Modify by Morgan 2006/6/6 改參數
'Private Function GetSitu(strPA0104 As String, Optional p_stPA75 As String) As String
Private Function GetSitu(p_stPA01 As String, p_stPA02 As String, p_stPA03 As String, p_stPA04 As String, Optional p_stPA75 As String, Optional p_stPA14 As String) As String

Dim lngPA14 As Long 'Add by Morgan 2004/7/26
   
'預設為英文一般定稿
GetSitu = "02"

'Modify by Morgan 2006/6/6 改Call公用函數

'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'Dim StrSqlB As String
'Dim rsB As New ADODB.Recordset

'StrSQLa = "Select * From PATENT WHERE " & ChgPatent(strPA0104)
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'   'Add by Morgan 2004/7/26
'   lngPA14 = Val("" & rsA.Fields("PA14"))
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
If p_stPA14 <> "" Then
   lngPA14 = TransDate(p_stPA14, 2)
Else
   lngPA14 = TransDate(Text1(0), 2)
End If
GetSitu = "0" & PUB_GetLanguage(p_stPA01, p_stPA02, p_stPA03, p_stPA04)
'end 2006/6/6

'Add by Morgan 2004/7/26
'2004.7.1以後英文定稿更新
If GetSitu = "02" And lngPA14 >= 20040701 Then
   GetSitu = "04"
'Removed by Morgan 2014/11/7 --敏莉
'   'Add by Morgan 2006/4/10
'   If p_stPA75 <> "" Then
'      p_stPA75 = Left(p_stPA75, 6)
'      'Modify by Morgan 2006/6/12 + Y20412,Y48162
'      'Modify by Morgan 2008/5/8 + Y21775  --David
'      'Modified by Morgan 2014/6/10 -Y21775 --敏莉
'      If p_stPA75 = "Y49575" Or p_stPA75 = "Y45697" Or p_stPA75 = "Y48309" Or p_stPA75 = "Y20412" Or p_stPA75 = "Y48162" Then
'         GetSitu = "05"
'      End If
'   End If
'end 2014/11/7
End If

End Function

'Add By Cheng 203/02/12
'定稿例外欄位
Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 20) As String, i As Integer, j As Integer, strTmp As String
Dim ii As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   ii = 0
   EndLetter ET01, strReceiveNo & "&000", ET03, strUserNum
   'Add By Cheng 2003/02/12
   '判斷是否不續辦但准通知
   StrSQLa = "Select PA89 From Patent Where " & ChgPatent(strReceiveNo)
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       If "" & rsA.Fields(0).Value = "Y" Then
            ii = ii + 1
            '附註
           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
              "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','附註','P.S. : This case has been allowed. If your client(s) want(s) to maintain this case, please notify us immediately.')"
       End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   'add by sonia 2014/4/25 特定客戶/代理人電子信箱
   If strSpecNO = True Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
        "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','特定客代電子信箱','E-mail: " & strEmail & "')"
   End If
   '2014/4/25 end
   
   'Added by Lydia 2022/04/29 請在公告通知函的介面加上一欄位供勾選□公報有誤加註定稿(本所案號列印才可使用)，若有勾選此欄位，定稿請加註
   If chkAddMemo.Enabled = True And chkAddMemo.Value = 1 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
        "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','公報有誤加註', '♀')"
   End If
   'end 2022/04/29
   
   'Add By Sindy 2015/11/16
   'Removed by Morgan 2019/2/26 定稿已修改,取消 -- Sharon,David
   'If strSpecNO_2 = False Then
   '   ii = ii + 1
   '   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
   '     "('" & ET01 & "','" & strReceiveNo & "&000','" & ET03 & "','" & strUserNum & "','非特定客戶要印','♀')"
   'End If
   'end 2019/2/26
   '2015/11/16 END
   
   If ii > 0 Then
       'edit by nickc 2007/02/05 不用 dll 了
       'If Not objLawDll.ExecSQL(ii, strTxt) Then
       If Not ClsLawExecSQL(ii, strTxt) Then
          MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
       End If
   End If
End Sub

'Added by Morgan 2012/5/31
Private Function GetFilePath(strDate As String) As Boolean
   
   Dim i As Integer, j As Integer
   
On Error GoTo ErrHnd
   
   If IsEmptyText(txtPath2) = True Then
      MsgBox "請輸入公告公報PDF的存放路徑！", vbExclamation + vbOKOnly
      Exit Function
   End If
   'If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   If Right(Trim(txtPath2), 1) <> "\" Then txtPath2 = txtPath2 & "\"
   
   strTPB04 = Format(Val(Val(Left(strDate, 4)) - 1911) - 62, "00")
   j = Val(Mid(strDate, 5, 2))
   i = (j - 1) * 3
   j = Val(Right(strDate, 2))
   If j >= 1 And j < 11 Then
      i = i + 1
   ElseIf j >= 11 And j < 21 Then
      i = i + 2
   ElseIf j >= 21 Then
      i = i + 3
   End If
   strTPB05 = Format(i, "00")
   
   'Add By Sindy 2015/11/13
   If InStr(UCase(txtPath2), UCase("\img_1\isu0")) > 0 Then
      txtPath2 = Mid(txtPath2, 1, InStrRev(UCase(txtPath2), UCase("\img_1\isu0")) + 10) & strTPB04 & "0" & strTPB05 & "\"
   End If
   '2015/11/13 END
   'If Dir(txtPath2 & "\img_1\isu0" & strTPB04 & "0" & strTPB05 & "\") = "" Then
   If Dir(txtPath2) = "" Then
      MsgBox "公告公報PDF的存放路徑中無" & strTPB04 & "卷" & strTPB05 & "期資料！"
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
      MsgBox "公告公報PDF的存放路徑中無" & strTPB04 & "卷" & strTPB05 & "期資料！"
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function
'Modified by Morgan 2013/1/9 +strPA14
'Modified by Lydia 2015/01/19 + bolEtype
Private Sub GetPDFCopys(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String, StrPA11 As String, ByRef int_Copys As Integer, Optional strPA14 As String, Optional bolEtype As Boolean)
   Dim strFileName As String, strToPath As String
   
   'Added by Morgan 2021/6/25 公報改抓卷宗區，不再往pat3讀取避免當機沒開的情形
   If PUB_GetGazettePDF(strPA01, strPA02, strPA03, strPA04, True, m_AttachPath, strFileName) = True Then
      List1.AddItem strFileName & "?" & int_Copys
      Call PrinBatchPdf(List1.ListCount - 1)
   End If
   Exit Sub
   'end 2021/6/25
   
   'Modify By Sindy 2013/1/4
   'strFileName = txtPath2 & "\img_1\isu0" & strTPB04 & "0" & strTPB05 & "\" & StrPA11 & "-P01.pdf"
   'Modified by Morgan 2013/1/9 102/1/1以前公告維持舊格式
   If Val(strPA14) >= "20130101" Then
      'strFileName = txtPath2 & "\img_1\isu0" & strTPB04 & "0" & strTPB05 & "\" & StrPA11 & ".pdf"
      strFileName = txtPath2 & StrPA11 & ".pdf"
   Else
      'strFileName = txtPath2 & "\img_1\isu0" & strTPB04 & "0" & strTPB05 & "\" & StrPA11 & "-P01.pdf"
      strFileName = txtPath2 & StrPA11 & "-P01.pdf"
   End If
   'end 2013/1/9
   '2013/1/4 End
   'Modified by Morgan 2014/9/25 pdf檔印完刪除
   'Modified by Morgan 2014/10/9 改回來--敏莉
   'If strPath = "" Then
      strPath = PUB_Getdesktop
      If Option1(0).Value = True Then
         strPath = strPath & "\$公報PDF檔" & Me.Text1(0)
      'Added by Lydia 2019/12/11 個案另外放
      Else
         strPath = App.path
      'end 2019/12/11
      End If
      'strPath = App.path & "\$公報PDF檔"
      
      If Dir(strPath, vbDirectory) = "" Then
         MkDir strPath
      End If
   'End If
   'end 2014/9/25
   'end 2014/10/9
   
   If Option1(0).Value = True Then   'Added by Lydia 2019/12/11 判斷不同路徑
       strToPath = strPath & "\" & strPA01 & strPA02 & strPA03 & strPA04 & "_" & Mid(strFileName, InStrRev(strFileName, "\") + 1)
   'Added by Lydia 2019/12/11
   Else
       strToPath = strPath & "\$$" & strPA01 & strPA02 & strPA03 & strPA04 & "_" & Mid(strFileName, InStrRev(strFileName, "\") + 1)
   End If
   'end 2019/12/11
   FileCopy strFileName, strToPath
   'List1.AddItem strFileName & " " & int_Copys
   List1.AddItem strToPath & "?" & int_Copys
   
   Call PrinBatchPdf(List1.ListCount - 1) 'Add By Sindy 2015/7/7 直接列印出來
   
   'Add by Lydia 2015/01/19 將公報檔自動存入共用資料匣
   'Remove by Lydia 2019/12/03 取消存入WorkFlow,需要時從卷宗區下載
   'If bolEtype = True Then
   '   strToPath = GetPath(strPA01, strPA02, strPA03, strPA04)
   '   '更名
   '   strToPath = strToPath & "\" & strPA01 & strPA02 & IIf(strPA03 & strPA04 <> "000", strPA03 & strPA04, "") & "Patent Gazette.pdf"
   '   FileCopy strFileName, strToPath
   'End If
   'end 2015/01/19
   'end 2019/12/03
End Sub

'Modify By Sindy 2015/7/7 +strPrintRow As String
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
   Dim intRow As Integer, intTotRow As Integer 'Add By Sindy 2015/7/7
   
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
    
'Removed by Morgan 2017/11/13 移到迴圈外面只要跑一次,全部印完再關閉
'    '因為第 2 個以後開啟的 Reader 才會印完後自動關閉,所以固定先開一個空的程式,全部印完後再關閉
'    process_id = SHELL(program_name, vbHide)
'    process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
'end 2017/11/13
    
On Error GoTo 0
    
   If ff1 > 0 Then Close #ff1
   ff1 = FreeFile
   Open App.path & "\FCP公告公報列印PDF時間資訊.txt" For Output As #ff1
   
   'Modify By Sindy 2015/7/7
   If strPrintRow = "A" Then
      intRow = 0
      intTotRow = List1.ListCount - 1
   Else
      intRow = CInt(strPrintRow)
      intTotRow = intRow
   End If
   'For ii = 0 To List1.ListCount - 1
   For ii = intRow To intTotRow
   '2015/7/7 END
      strTemp = Split(List1.List(ii), "?")
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
    
'Removed by Morgan 2017/11/13 移到迴圈外面只要跑一次,全部印完再關閉
'   TerminateProcess process_handle, 0&
'   CloseHandle process_handle
'end 2017/11/13
   
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

Private Sub txtLetterDate_GotFocus()
   CloseIme
   TextInverse txtLetterDate
End Sub

Private Sub txtLetterDate_Validate(Cancel As Boolean)
   If ChkDate(txtLetterDate) = False Then
      Cancel = True
   End If
End Sub
'Add by Lydia 2015/01/19
'指定存檔路徑
Private Function GetPath(ByVal m_PA01 As String, ByVal m_PA02 As String, ByVal m_PA03 As String, ByVal m_PA04 As String) As String
Dim strSubDir As String
   strSubDir = PUB_GetEFilePath(m_PA01) & "\" & m_PA01
   If Dir(strSubDir, vbDirectory) = "" Then
      MkDir strSubDir
   End If
   strSubDir = strSubDir & "\" & Left(m_PA02, 3)
   If Dir(strSubDir, vbDirectory) = "" Then
      MkDir strSubDir
   End If
   strSubDir = strSubDir & "\" & m_PA01 & m_PA02 & IIf(m_PA03 & m_PA04 <> "000", m_PA03 & m_PA04, "")
   If Dir(strSubDir, vbDirectory) = "" Then
      MkDir strSubDir
   End If
   strSubDir = strSubDir & "\公報" '針對-E化案件信函自動產生有簽名檔的pdf檔->存放位置
   If Dir(strSubDir, vbDirectory) = "" Then
      MkDir strSubDir
   End If
   GetPath = strSubDir
End Function

'Add by Lydia 2015/01/19
'轉PDF
Private Sub PrintWord2PDF(ByVal m_PA01 As String, ByVal m_PA02 As String, ByVal m_PA03 As String, ByVal m_PA04 As String)
   Dim strFolder As String
   Dim strFileName As String
   Dim strFullName As String

   strFolder = GetPath(m_PA01, m_PA02, m_PA03, m_PA04)
   strFileName = m_PA01 & m_PA02 & IIf(m_PA03 & m_PA04 <> "000", m_PA03 & m_PA04, "") & "Letter(Patent Gazette)"
   
'Modified by Lydia 2019/03/13 改成共用模組
'   'Added by Morgan 2017/11/13
'   If pub_Word2Pdf Then
'      strFullName = strFolder & "\" & strFileName & ".pdf"
'      g_WordAp.ActiveDocument.ExportAsFixedFormat OutputFileName:=strFullName, ExportFormat:=17, OpenAfterExport:=False
'   Else
'   'end 2017/11/13
'      frmPDF.Show
'      frmPDF.StartProcess strFolder, strFileName
'      '切換印表機
'      If PUB_PdfCreatorNameInWord = "" Then PUB_PdfCreatorNameInWord = PUB_GetCreatorNameInWord
'      g_WordAp.ActivePrinter = PUB_PdfCreatorNameInWord
'      g_WordAp.ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
'      frmPDF.EndtProcess
'      Unload frmPDF
'
'      'Added by Morgan 2017/11/13
'      PUB_SetOsDefaultPrinter cmbPrinter3
'      PUB_SetWordActivePrinter
'      PUB_RestorePrinter cmbPrinter3
'      'end 2017/11/13
'
'   End If 'Added by Morgan 2017/11/13
'
'    g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
'    If g_WordAp.Documents.Count = 0 Then
'       g_WordAp.Quit wdDoNotSaveChanges
'       Set g_WordAp = Nothing 'Added by Morgan 2017/11/13
'    End If
    'Modified by Lydia 2023/04/27 改模組
    'If PUB_PrintWord2PDF(g_WordAp, strFolder, strFileName, "", cmbPrinter3) = True Then
    If PUB_PrintWord2File(g_WordAp, strFolder, strFileName) = True Then
'end 2019/03/13
        If Me.Option1(1).Value = True Then
           MsgBox "PDF檔已存於 " & strFolder & "！"
        End If
    End If '2019/03/13
End Sub

