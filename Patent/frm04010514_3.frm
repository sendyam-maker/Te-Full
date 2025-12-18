VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010514_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "初審及公佈通知來函輸入"
   ClientHeight    =   4572
   ClientLeft      =   1176
   ClientTop       =   1848
   ClientWidth     =   8772
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4572
   ScaleWidth      =   8772
   Begin VB.TextBox Text25 
      Height          =   270
      Left            =   1590
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1890
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   90
      TabIndex        =   38
      Top             =   2670
      Width           =   8520
      Begin VB.TextBox Text27 
         Height          =   270
         Left            =   3420
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   3
         Top             =   150
         Width           =   975
      End
      Begin VB.TextBox Text26 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "3"
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text15 
         Height          =   270
         Left            =   5490
         MaxLength       =   8
         TabIndex        =   4
         Top             =   150
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   7470
         MaxLength       =   8
         TabIndex        =   5
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "法定期限:"
         Height          =   180
         Index           =   4
         Left            =   2565
         TabIndex        =   42
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "本所期限:"
         Height          =   180
         Index           =   3
         Left            =   4635
         TabIndex        =   41
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "主動修正期限: 文到           月"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   40
         Top             =   195
         Width           =   2205
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "約定期限:"
         Height          =   180
         Index           =   0
         Left            =   6615
         TabIndex        =   39
         Top             =   195
         Width           =   765
      End
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   4
      Left            =   1590
      MaxLength       =   1
      TabIndex        =   11
      Top             =   4215
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   6
      Left            =   1590
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2265
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   7
      Left            =   1590
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3225
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   13
      Left            =   1590
      MaxLength       =   8
      TabIndex        =   9
      Top             =   3900
      Width           =   1092
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1590
      MaxLength       =   8
      TabIndex        =   7
      Top             =   3570
      Width           =   1092
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   4275
      MaxLength       =   20
      TabIndex        =   8
      Top             =   3570
      Width           =   1650
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   4275
      MaxLength       =   20
      TabIndex        =   10
      Top             =   3885
      Width           =   1650
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   18
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   17
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   16
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   15
      Top             =   660
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7788
      TabIndex        =   14
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5736
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6564
      TabIndex        =   13
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   960
      TabIndex        =   19
      Top             =   960
      Width           =   7635
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "13467;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "官方發文日:"
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   44
      Top             =   1935
      Width           =   945
   End
   Begin VB.Label Label22 
      Caption         =   "(1:大陸初審合格 2:大陸/香港/澳門公布 3:初審合格及進入實審) (4:公布及進入實審 5:進入實審 6:香港/澳門公告)"
      Height          =   390
      Left            =   1935
      TabIndex        =   43
      Top             =   2205
      Width           =   4920
   End
   Begin VB.Line Line1 
      DrawMode        =   16  'Merge Pen
      Index           =   1
      X1              =   135
      X2              =   8635
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   135
      X2              =   8635
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "列印客戶通知函:          (N:不印)"
      Height          =   180
      Left            =   120
      TabIndex        =   37
      Top             =   4260
      Width           =   2430
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "非台灣案通知書:"
      Height          =   180
      Left            =   120
      TabIndex        =   36
      Top             =   2310
      Width           =   1305
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "大陸實審收文:               (N:無 /  空:有)"
      Height          =   180
      Left            =   120
      TabIndex        =   35
      Top             =   3255
      Width           =   2850
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "港/澳公告日:"
      Height          =   180
      Left            =   120
      TabIndex        =   34
      Top             =   3945
      Width           =   990
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "公　開　號:"
      Height          =   180
      Index           =   1
      Left            =   3210
      TabIndex        =   33
      Top             =   3615
      Width           =   945
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "公　開　日"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   3615
      Width           =   900
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "香港公告號:"
      Height          =   180
      Index           =   2
      Left            =   3210
      TabIndex        =   31
      Top             =   3945
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Enabled         =   0   'False
      Height          =   180
      Index           =   4
      Left            =   8010
      TabIndex        =   30
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   3
      Left            =   1200
      TabIndex        =   29
      Top             =   1530
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   2
      Left            =   5340
      TabIndex        =   28
      Top             =   1290
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   1
      Left            =   1200
      TabIndex        =   27
      Top             =   1290
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   0
      Left            =   5340
      TabIndex        =   26
      Top             =   660
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   768
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   4380
      TabIndex        =   24
      Top             =   660
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   23
      Top             =   660
      Width           =   768
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   120
      TabIndex        =   22
      Top             =   1290
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   4380
      TabIndex        =   21
      Top             =   1290
      Width           =   585
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   1530
      Width           =   945
   End
End
Attribute VB_Name = "frm04010514_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/16 改成Form2.0 (Combo1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Create by Morgan 2009/11/24 自內專核准函輸入抽出
Option Explicit

Dim intWhere As Integer
Dim cp() As String
Dim pa() As String

Dim strReceiveNo As String
Dim m_HaveHK As Boolean
Dim m_HaveHKInCP As String
Dim m_HaveHKInNP As String
Dim m_SendHKMail As Boolean
Dim m_HK_CP01 As String
Dim m_HK_CP02 As String
Dim m_HK_CP03 As String
Dim m_HK_CP04 As String
Dim m_HKMailID As String
Dim m_HKMailCCID As String 'Added by Morgan 2020/3/5 FCP管制程序及其主管(寰華案CC對象)
'2006/2/7 ADD BY SONIA 香港公布選擇定稿別 1 記錄請求公佈通知書 2 政府憲報(2025/07/08改為知識產權公報)
Dim m_LetterType As VbMsgBoxResult
Dim m_bolSaveCheck As Boolean '是否為存檔前檢查
Dim m_bolFMP As Boolean '是否FMP案
Dim m_HK_110Cp06 As String
Dim m_HK_110Cp07 As String
Dim NewReceiveNo As String
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/05/17 是否為寰華案
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END
Dim m_bolNoCP27 As Boolean '不上發文 Added by Morgan 2020/1/16
Dim stCP12 As String, stCP13 As String 'Add by Morgan 2004/2/9
Dim m_bolFMPNoPrint As Boolean 'Added by Morgan 2023/4/11 FMP案是否列印中文定稿
Dim m_st245CP71 As String, m_st245CP09 As String, m_bolAdd1924 As Boolean  'Added by Lydia 2025/02/12 大陸案是否已發文延緩審查

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(1 To 23) As String, intStep As Integer, lTmp As Long, lTmp1 As Long
   Dim strTmp As String

   EndLetter ET01, strReceiveNo, ET03, strUserNum
   intStep = 1
      
   Select Case pa(9)
     
      Case "020"
          '若無大陸實審收文
         If Text5(7) = "N" Then
            
            If ET03 = "06" Then '大陸PCT案之初步審查合格 06
               'Modify by Morgan 2007/1/29 要抓最早優先權日,通知期限要抓本所期限
               'lTmp = CompDate(0, 3, TransDate(Label2(4), 1))
               '法定期限
               
'Modified by Morgan 2012/8/24 FMP所限不同改抓下一程序期限(Ex.P-102095)
'               strExc(1) = PUB_GetFirstPriDate(cp())
'               If strExc(1) <> "" Then
'                  lTmp = PUB_GetFirstPriDate(cp())
'               Else
'                  lTmp = TransDate(Label2(4), 1)
'               End If
'               lTmp = CompDate(0, 3, lTmp)
'               strExc(1) = cp(1)
'               strExc(2) = pa(9)
'               strExc(3) = lTmp
'               GetCtrlDT strExc
'               '本所期限
'               '若本所期限非工作天則抓最近的工作天
'               lTmp1 = PUB_GetWorkDay1(strExc(0), True)
'               'End 2007/1/29
               strExc(0) = "select np08,np09 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06 is null and np07='416'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  lTmp1 = RsTemp(0)
                  lTmp = RsTemp(1)
               End If
               'end 2012/8/24
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','有無實審','依大陸專利法規定 : 「發明專利申請自申請日起三年內，專利局可以根據申請人隨時提出的請求，對其申請進行實質審查 : 申請人無正當理由逾期不請求審質查的，該申請即視為撤回。」亦即本案應於" & (Left(lTmp, 4)) & "年" & Mid(lTmp, 5, 2) & "月" & Mid(lTmp, 7, 2) & "日前提出實質審查請求，貴公司所申請之本案若要提出實質審查時請於" & (Left(lTmp1, 4)) & "年" & Mid(lTmp1, 5, 2) & "月" & Mid(lTmp1, 7, 2) & "日前請通知本所，以利作業。')"
               intStep = intStep + 1
            ElseIf ET03 = "05" Then '大陸非PCT案之初步審查合格 05
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','有無實審','依大陸專利法規定，專利自申請日起三年內，申請人可隨時提出實質審查請求，若逾期不請求實質審查，該申請即視為撤回，本案若要提出實質審查請求，請儘速通知本所。')"
               intStep = intStep + 1
              'Add By Cheng 2002/12/29
            ElseIf ET03 = "07" Then '大陸/香港公布 07
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','有無實審','依大陸專利法規定，專利自申請日起三年內，申請人可隨時提出實質審查請求，若逾期不請求實質審查，該申請即視為撤回，本案若要提出實質審查請求，請儘速通知本所。')"
               intStep = intStep + 1
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','提出修正','依大陸專利法實施細則規定，申請人在提出實質審查請求時，可以對發明專利申請主動提出修正，本案若欲提出修正，請一併通知本所。')"
               intStep = intStep + 1
            End If
          '若有大陸實審收文
         Else
            'Modified by Morgan 2018/10/29 +判斷未取消收文 Ex:P116552
            strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
               " AND CP10='" & 實體審查 & "' AND CP27 IS NULL and CP57 IS NULL"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 And RsTemp.Fields(0).Value <> 0 Then    '無發文日
               'Modified by Morgan 2025/9/30
               'strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','有無實審','本案將於近日內向代理人提出實體審查請求，')"
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','有無實審','本案將於近日內向代理人提出實體審查請求。')"
               'end 2025/9/30
               intStep = intStep + 1
            ElseIf intI = 1 And RsTemp.Fields(0).Value = 0 Then   '有發文日
               'Removed by Morgan 2025/9/30 移到下面
               'strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '   "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               '   "','有無實審','本案已向大陸專利局提出實體審查請求，')"
               'intStep = intStep + 1
               'end 2025/9/30
               
               'Add By Cheng 2002/12/29
               '初步審查合格及進入實質審查程序通知書 20
               'Modified by Morgan 2016/1/22 +"33"
               If ET03 = "20" Or ET03 = "21" Or ET03 = "22" Or ET03 = "33" Then
                  'Added by Morgan 2025/9/30
                  strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','有無實審','本案已向大陸專利局提出實體審查請求，')"
                  intStep = intStep + 1
                  'end 2025/9/30
               
                  '收文日加三個月再提前十天
                  '92.5.20 MODIFY BY SONIA 改抓申請案核准日
                  'lTmp = CompDate(2, -10, CompDate(1, 3, TransDate(Label2(3), 2)))
                  'Modify by Morgan 2009/11/25 改抓畫面上的約定期限
                  'lTmp = CompDate(2, -10, CompDate(1, 3, TransDate(Text5(0), 2)))
                  lTmp = DBDATE(Text6)
                  strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                    "','提出修正','依大陸專利法實施細則規定，申請人在收到發明專利進入實質審查階段通知書之日起三個月內，可以對發明專利申請主動提出修正，本案若欲提出修正，請於" & (Left(lTmp, 4)) & "年" & Mid(lTmp, 5, 2) & "月" & Mid(lTmp, 7, 2) & "日前請通知本所，以利作業。')"
                  intStep = intStep + 1
               'Added by Morgan 2025/9/30
               Else
                  strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','有無實審','本案已向大陸專利局提出實體審查請求。')"
                  intStep = intStep + 1
               'end 2025/9/30
               End If
               
               'Added by Lydia 2025/02/12 大陸案已發文延緩審查並且准予延緩審查
               If m_st245CP71 <> "" And m_bolAdd1924 = True And ET03 = "22" Then
                  strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','准予延緩審查','♀')"
                  intStep = intStep + 1
                  strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','延緩審查日','" & m_st245CP71 & "')"
                  intStep = intStep + 1
               End If
               'end 2025/02/12
                              
            End If
         End If
         
         '若有輸入香港公告日 2006/2/7 SONIA加大陸案條件
         If Text5(13) <> "" And pa(9) = "020" Then
            'Modify By Cheng 2002/12/29
            If ET03 = "07" Or ET03 = "20" Or ET03 = "21" Or ET03 = "22" Then
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','其他公告日','依據一九九七年實施之香港專利法規定，若專利權人欲在香港取得標準專利之保護，必需在指定專利局部（如中國專利局）之申請案公開及核准後的六個月內分別向香港專利局提出專利登錄請求核准註冊請求。" & _
                  "故本案若欲於香港取得專利保護，必需於" & Mid(Val(Me.Text5(13).Text) + 19110000, 1, 4) & "年" & Val(Mid(Val(Me.Text5(13).Text) + 19110000, 5, 2)) & "月" & Val(Mid(Val(Me.Text5(13).Text) + 19110000, 7, 2)) & "日前完成第一階段的登錄請求。" & _
                  "有關此期限請特加留意，以保權益。')"
               intStep = intStep + 1
            Else
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','其他公告日','" & Text5(13) & "')"
               intStep = intStep + 1
            End If
         End If
      
         'add by sonia 2016/5/3 FMP案定稿期限會與下一程序本所不同 P-111817(原定稿設{{{<專利基本檔-公開日>+5月}-5日}/中西})
         'Modified by Morgan 2021/10/18 +98(寶齡富錦英文定稿)
         If Text7.Text <> "" And (ET03 = "07" Or ET03 = "21" Or ET03 = "98") Then
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','其他日期','" & m_HK_110Cp06 & "')"
            intStep = intStep + 1
         End If
         'end 2016/5/3
         
      '2006/2/7 ADD BY SONIA
      Case "013"
      '2006/2/7 END
   End Select
   
   'Added by Morgan 2021/10/18
   '約定期限 (寶齡富錦英文定稿會用)
   If Text6 <> "" Then
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','約定期限','" & DBDATE(Text6) & "')"
      intStep = intStep + 1
   End If
   'end 2021/10/18
   
   'Added by Lydia 2025/10/31 FMP的定稿保持原本的期限=其他日期; 參考FMP案定稿期限會與下一程序本所不同 P-111817(原定稿設{{{<專利基本檔-公開日>+5月}-5日}/中西})
   If m_bolFMP = True And m_HK_110Cp06 <> "" Then
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','香港案法限','" & m_HK_110Cp06 & "')"
      intStep = intStep + 1
   Else
   'end 2025/10/31
      'Added by Morgan 2025/7/9
      If m_HK_110Cp07 <> "" Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','香港案法限','" & m_HK_110Cp07 & "')"
         intStep = intStep + 1
      End If
      'end 2025/7/9
   End If
   If Not ClsLawExecSQL(intStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If

End Sub

Private Sub StartLetter2(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt() As String, i As Integer
   
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   i = 0
   If pa(46) = "Y" Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','PCT案','♀')"
   Else
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','非PCT案','♀')"
   End If

   If Text25 <> "" Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','官方發文日','" & DBDATE(Text25) & "')"
   End If
   
   If Text6 <> "" Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','約定期限','" & DBDATE(Text6) & "')"
   End If
   
   If m_HK_110Cp07 <> "" Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','香港案法限','" & m_HK_110Cp07 & "')"
   End If
   
   If m_HK_110Cp06 <> "" Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','香港案所限','" & m_HK_110Cp06 & "')"
   End If

   If Text5(6) = "1" Or Text5(6) = "2" Then
      strExc(0) = "select np09 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "'" & _
         " and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06||np07='416'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','實審未收文','♀')"
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','實審法定期限','" & RsTemp(0) & "')"
      '已提實審(不考慮已收未發情形)
      Else
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','實審已收文','♀')"
      End If
   End If
      
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)

   Dim strTmp As String, strTmp2 As String
   
   Select Case Index
      Case 0 '確定
         If Text5(6) = "" Then
            MsgBox "請輸入通知書選項！"
            Text5(6).SetFocus
            Exit Sub
         End If
         
         If Me.Text5(6).Text = "2" Or Me.Text5(6).Text = "4" Then
            If Me.Text7.Text = "" Then
               MsgBox "請輸入公開日!!!", vbExclamation + vbOKOnly
               Me.Text7.SetFocus
               Text7_GotFocus
               Exit Sub
            End If
            If Me.Text8.Text = "" Then
               MsgBox "請輸入公開號!!!", vbExclamation + vbOKOnly
               Me.Text8.SetFocus
               Text8_GotFocus
               Exit Sub
            End If
         End If
      
         '2006/2/7 ADD BY SONIA香港案公告
         If pa(9) = "013" And Me.Text5(6).Text = "6" Then
            If Me.Text5(13).Text = "" Then
               MsgBox "請輸入香港公告日!!!", vbExclamation + vbOKOnly
               Me.Text5(13).SetFocus
               Text7_GotFocus
               Exit Sub
            End If
            If pa(8) = "2" And Me.Text11.Text = "" Then   '短期再檢查公告號
               MsgBox "請輸入香港公告號!!!", vbExclamation + vbOKOnly
               Me.Text11.SetFocus
               Text8_GotFocus
               Exit Sub
            End If
         End If
         
         'Add By Morgan 2006/10/14  澳門公告
         If pa(9) = "044" And Me.Text5(6).Text = "6" Then
            If Me.Text5(13).Text = "" Then
               MsgBox "請輸入澳門公告日!!!", vbExclamation + vbOKOnly
               Me.Text5(13).SetFocus
               Text7_GotFocus
               Exit Sub
            End If
         End If
         
         'Added by Lydia 2025/02/12 大陸案：在進實審通知書內說明，若系統進度已發文延緩審查，請於輸入進入實審通知時彈跳視窗讓user輸入及檢核
         m_bolAdd1924 = False
         strExc(1) = ""
         If m_st245CP09 <> "" And Me.Text5(6).Text = "5" Then
            If m_st245CP71 = "" Then
               MsgBox "發文未輸入延緩審查年度！", vbCritical
               Exit Sub
            End If
            If MsgBox("是否已准予延緩審查？", vbYesNo + vbDefaultButton1) = vbYes Then
JumpToReInput:
               strExc(1) = InputBox("請輸入延緩審查日期的1~3年度", "延緩審查日期")
               If strExc(1) <> "1" And strExc(1) <> "2" And strExc(1) <> "3" Then
                  GoTo JumpToReInput
               Else
                  If m_st245CP71 <> "" And strExc(1) <> m_st245CP71 Then
                     MsgBox "發文輸入延緩審查年度為" & m_st245CP71, vbCritical
                     GoTo JumpToReInput
                  End If
               End If
               m_st245CP71 = strExc(1)
               m_bolAdd1924 = True
            Else
               MsgBox "請確認國知局是否已准予延緩審查！", vbCritical
               Exit Sub
            End If
         End If
         'end 2025/02/12
         
         '重新檢查欄位有效性
         m_bolSaveCheck = True
         If TxtValidate = False Then
            m_bolSaveCheck = False
            Exit Sub
         End If
         
         'Add By Sindy 2022/7/1
         'Mark by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail:可取消外專系統收件區，key來函承辦人掛程序人員，則按確定，信件會再打開一次的設定。
         'If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
         '   If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
         '      Exit Sub
         '   End If
         'End If
         ''2022/7/1 END
         'end 2023/05/17
         
         'add by nickc 2005/06/17 取有關大陸的香港關聯判斷
         m_HaveHK = False
         m_HaveHKInCP = ""
         m_HaveHKInNP = ""
         m_SendHKMail = False
         m_HKMailID = ""
         'Modified by Morgan 2014/9/23 大陸發明才要--郭
         'If pa(9) = "020" Then
         If pa(8) = "1" And pa(9) = "020" Then
            'Modified by Morgan 2014/9/23 香港標準專利(發明)才要--郭
            m_HaveHK = ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), m_HK_CP01, m_HK_CP02, m_HK_CP03, m_HK_CP04, , "1")
            'Remove by Morgan 2009/11/27 改到要用時再抓就好
            'If m_HaveHK = True Then
            '   m_HaveHKInCP = Chk013Have110(pa(1), pa(2), pa(3), pa(4), m_HKMailID)
            '   m_HaveHKInNP = Chk013Have110(pa(1), pa(2), pa(3), pa(4), m_HKMailID, "NP")
            'End If
         End If
                 
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         If Text5(4) <> "N" Then '通知函
         
            Select Case pa(9)
                                     
               Case 大陸國家代號
                  '大陸香港發明通知書
                  Select Case Text5(6)
                     Case "1"
                        If pa(46) = "Y" Then
                           '大陸PCT案之初步審查合格 06
                           strTmp = "06"
                        Else
                           '大陸非PCT案之初步審查合格 05
                           strTmp = "05"
                        End If
                        
                        If m_bolFMP Then
                           strTmp2 = "51"
                        End If
                        
                     Case "2" '大陸/香港公布 已提 07
                        'edit by nickc 2005/06/17
                        'strTmp = "07"
                        If m_HaveHK = True Then
                           strTmp = "34"
                        Else
                           strTmp = "07"
                        End If
                        If m_bolFMP Then
                           strTmp2 = "52"
                        End If
                     'Add By Cheng 2002/10/11
                     Case "3" '初步審查合格及進入實質審查程序通知書
                        strTmp = "20"
                     Case "4" '公布及進入實質審查程序通知書
                        'edit by nickc 2005/06/17
                        'strTmp = "21"
                        If m_HaveHK = True Then
                           strTmp = "33"
                        Else
                           strTmp = "21"
                           'Added by Morgan 2021/10/18 寶齡富錦 Y55435 案件
                           If ChangeCustomerS(pa(75)) = "Y55435" Then
                              strTmp = "98"
                           End If
                           'end 2021/10/18
                        End If
                        If m_bolFMP Then
                           strTmp2 = "53"
                        End If
                     Case "5" '進入實質審查程序通知書
                        strTmp = "22"
                        If m_bolFMP Then
                           strTmp2 = "56"
                        End If
                  End Select
               Case "013" '香港
                  Select Case Text5(6)
                     Case "2"   '標準專利記錄請求公布通知書 11或12
                        If m_LetterType = vbYes Then
                           strTmp = "11"
                        Else
                           strTmp = "12"
                        End If
                        'Added by Morgan 2013/10/15
                        If m_bolFMP Then
                           strTmp2 = "52"
                        End If
                        'end 2013/10/15
                     Case "6"   '公告
                        Select Case cp(10)
                           Case 批准記錄請求_標準專利  '標準專利核准 13
                              strTmp = "13"
                           Case 短期專利申請           '短期專利 14
                              strTmp = "14"
                        End Select
                  End Select
               'Add by Morgan 2006/10/14
               Case "044" '澳門
                  Select Case Text5(6)
                     Case "6" '公告
                        strTmp = "11"
                     'Add by Morgan 2007/8/15
                     Case "2" '公佈
                        strTmp = "12"
                     'Add by Morgan 2010/1/25(定稿未提供)
                     Case "4"
                        strTmp = "38"
                  End Select
                  
               '2006/2/6 ADD BY SONIA
               Case "056" 'ＰＣＴ只有公布通知書,無核准定稿
                  If Text5(6) = "2" Then  '公布通知書 11
                     strTmp = "08"
                  End If
               '2006/2/6 END
            End Select
            
            StartLetter "05", strTmp
            'Add by Morgan 2009/11/25
            If m_bolFMP Then
               'Modified by Morgan 2023/4/11 +m_bolFMPNoPrint
               NowPrint strReceiveNo, "05", strTmp, False, strUserNum, , , , , 1, , , , , , , , NewReceiveNo, , , , , m_bolFMPNoPrint 'Modified by Morgan 2016/6/6 +NewReceiveNo
               If strTmp2 <> "" Then
                  strUserNum = strFMPNum
                  StartLetter2 "05", strTmp2
                  NowPrint strReceiveNo, "05", strTmp2, False, strUserNum, 0
                  strUserNum = strUser1Num
               End If
            Else
            'end 2009/11/25
               NowPrint strReceiveNo, "05", strTmp, False, strUserNum, , , , , , , , , , , , , NewReceiveNo 'Modified by Morgan 2016/6/6 +NewReceiveNo
            End If

            'add by nickc 2005/06/17 發 mail
            If m_SendHKMail = True And m_HKMailID <> "" And m_HaveHKInCP <> "" Then
               'Modify by Morgan 2009/11/27
               'Call PUB_SendMail(strUserNum, m_HKMailID, m_HaveHKInCP, "大陸案(" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & ")已公布，香港案(" & m_HK_CP01 & "-" & m_HK_CP02 & "-" & m_HK_CP03 & "-" & m_HK_CP04 & ")的[批准紀錄請求]可以處理！", "大陸案(" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & ")已公布，香港案(" & m_HK_CP01 & "-" & m_HK_CP02 & "-" & m_HK_CP03 & "-" & m_HK_CP04 & ")的[批准紀錄請求]可以處理！", "")
               'Modified by Morgan 2020/3/5 寰華案同時要CC給外專程序管制人員及其主管 m_HKMailCCID
               Call PUB_SendMail(strUserNum, m_HKMailID, m_HaveHKInCP, "大陸案(" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & ")已公布，香港案(" & m_HK_CP01 & "-" & m_HK_CP02 & "-" & m_HK_CP03 & "-" & m_HK_CP04 & ")的[標準專利紀錄請求]可以處理！", "大陸案(" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & ")已公布，香港案(" & m_HK_CP01 & "-" & m_HK_CP02 & "-" & m_HK_CP03 & "-" & m_HK_CP04 & ")的[標準專利紀錄請求]可以處理！", "", , , , , m_HKMailCCID)
            End If
         End If
         'Added by Lydia 2016/02/03 大陸發明案之主動補正會在通知公布及進入實審通知時寫入期限,在此時一併發MAIL通知工程師
         If pa(9) = "020" And pa(8) = "1" And InStr("3,4,5", Me.Text5(6).Text) > 0 And Text27 <> "" Then
            strExc(0) = "select cp14,nvl(cu04,nvl(cu05,cu06)) custname from caseprogress,patent,customer where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='203' and cp27||cp57 is null" & _
                        " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(1) = pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & " 主動補正更新本所期限:" & CFDate(ChangeWStringToTString(Text15))
               strExc(2) = "本所案號: " & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & vbCrLf & _
                           "案件名稱: " & IIf(pa(5) <> "", pa(5), IIf(pa(6) <> "", pa(6), pa(7))) & vbCrLf & _
                           "案件性質: 主動補正" & vbCrLf & _
                           "申  請  人: " & RsTemp.Fields("custname") & vbCrLf & _
                           "本所期限: " & CFDate(ChangeWStringToTString(Text15))
               If "" & RsTemp.Fields("cp14") <> "" Then
                  Call PUB_SendMail(strUserNum, RsTemp.Fields("cp14"), "", strExc(1), vbCrLf & vbCrLf & strExc(2))
               End If
            End If
         End If
         'end 2016/02/03
         
         'Add By Sindy 2016/10/5
         If Me.m_strIR01 <> "" Then
            Unload frm04010514_1
            Unload frm04010514_2
            Unload Me
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         Else
         '2016/10/5 END
            strKey1 = "1"
            Unload frm04010514_2
            Unload Me
            frm04010514_1.Show
            frm04010514_1.Clear
         End If
      Case 1
         frm04010514_2.Show
         Unload Me
      Case 2
         Unload frm04010514_1
         Unload frm04010514_2
         Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean

Dim i As Integer, bStartTrans As Boolean
Dim strTmp(1 To 5) As String
Dim stCP10 As String 'Add by Morgan 2007/10/25 來函案件性質
Dim strPromoteDate As String '2010/1/19 add by sonia
'Added by Morgan 2012/10/23 香港案維持費管控
Dim iYear As Integer, strNextFeeDate As String, strNextDueDate As String
Dim stPA72 As String, stPA73 As String, stPA74 As String
Dim strCP20 As String, strCP16 As String 'Added by Morgan 2019/8/8
Dim stNP23 As String 'Added by Lydia 2025/10/29

   FormSave = False
   
On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   bStartTrans = True
   
   'Modified by Lydia 2025/09/23 改成模組
'   'Add By Cheng 2003/04/16
'   '若申請國家非台灣
'   If pa(9) <> 台灣國家代號 Then
'      '判斷大陸香港發明通知書欄位
'      Select Case Me.Text5(6).Text
'         Case "1"
'            stCP10 = "1213"
'         Case "2"
'            stCP10 = "1207"
'         Case "3"
'            stCP10 = "1214"
'         Case "4"
'            stCP10 = "1215"
'         Case "5"
'            stCP10 = "1204"
'         '2006/2/7 ADD BY SONIA香港公告
'         Case "6"
'            stCP10 = "1208"
'      End Select
'   End If
'
'   'MODIFY BY SONIA 90.10.21
'   If Text5(6) = "1" Then
'      strTmp(1) = "初步審查合格通知書"
'   End If
'   If Text5(6) = "2" Then
'      strTmp(1) = "公布通知書"
'   End If
'   'Add By Cheng 2002/06/21
'   If Me.Text5(6).Text = "3" Then
'      strTmp(1) = "初步審查合格及進入實質審查程序通知書"
'   End If
'   If Me.Text5(6).Text = "4" Then
'      strTmp(1) = "公布及進入實質審查程序通知書"
'   End If
'   '93.1.6 ADD BY SONIA 93.1.6
'   If Text5(6) = "5" Then
'      strTmp(1) = "進入實質審查程序通知書"
'   End If
'   '93.1.6 END
'   '2006/2/7 ADD BY SONIA 93.1.6
'   If Text5(6) = "6" Then
'      'Modify by Morgan 2006/10/14 加澳門
'      'strTmp(1) = "香港公告"
'      If pa(9) = "013" Then
'         strTmp(1) = "香港公告"
'      ElseIf pa(9) = "044" Then
'         strTmp(1) = "澳門公告"
'      Else
'         strTmp(1) = "公告"
'      End If
'   End If
'   '2006/2/7 END
   stCP10 = GetResult(Me.Text5(6), strTmp(1))
   'end 2025/09/23
   
   '3
   NewReceiveNo = AutoNo("C", 6)
   
   'Modified by Morgan 2019/8/8 FMP案的CP20要抓設定
   strCP20 = ""
   If m_bolFMP Then
      strCP20 = PUB_GetCP20(pa(1), stCP10, strCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
   End If
         
   'Modify by Morgan 2010/6/2 +CP133,CP134 後續程序計算期限會用(Ex.後收文的主動補正)
   'Modified by Morgan 2012/4/30 +cp119=櫃檯收文日
   'Modified by Morgan 2020/1/16 +m_bolNoCP27
   strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
     "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43,CP64,CP133,CP134,cp119) VALUES ('" & Text1 & "','" & Text2 & "','" & _
     Text3 & "','" & Text4 & "'," & TransDate(Label2(3), 2) & _
     ",'" & NewReceiveNo & "','" & stCP10 & "'," & CNULL(stCP12) & "," & _
     CNULL(stCP13) & ",'" & strUserNum & "','" & strCP20 & "','N','N'," & IIf(m_bolNoCP27, "NULL", strSrvDate(1)) & ",'" & strReceiveNo & "'" & _
     "," & CNULL(ChgSQL(strTmp(1))) & "," & CNULL(Text25, True) & "," & CNULL(Text26, True) & "," & DBDATE(Label2(3)) & ")"
     'END 2007/6/12
   'Modify end 2004/2/9
   
   cnnConnection.Execute strSql, intI
   
   'Added by Morgan 2016/6/6
   If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      'Modified by Morgan 2018/8/1
      'strExc(1) = PUB_GetLetterJudge(pa(1), stCP10, , pa(9), pa(1), pa(2), pa(3), pa(4))
      strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), stCP10, pa(9), , , m_bolFMP)
      PUB_AddLetterProgress NewReceiveNo, 2, IIf(Text5(4) = "N", False, True), strExc(1), False, pa(26), stCP10, pa(75)
   End If
   'end 2016/6/6
   
   'MODIFY BY SONIA 90.11.15向香港提新案期限只在定稿上通知,不管制期限
   
   strSql = "UPDATE PATENT SET PA12=" & CNULL(TransDate(Text7, 2)) & ",PA13=" & CNULL(Text8) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   cnnConnection.Execute strSql, intI

   '2006/2/7 ADD BY SONIA 香港公告
   'Modify by Morgan 2006/10/14 加澳門
   'If pA(9) = "013" And Text5(6) = "6" Then
   If (pa(9) = "013" Or pa(9) = "044") And Text5(6) = "6" Then
      strSql = "UPDATE PATENT SET PA14=" & CNULL(TransDate(Text5(13), 2)) & ",PA15=" & CNULL(ChgSQL(Text11)) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      cnnConnection.Execute strSql, intI
   End If
   '2006/2/7 END

   'add by nickc 2005/06/17 大陸香港關聯案
   'edit by nickc 2006/08/18 公布才作
   'If pA(9) = "020" Then
   'Modify by Morgan 2007/10/25 應判斷來函的案件性質
   'If pa(9) = "020" And cp(10) = "1207" Then
   'Modify by Morgan 2009/11/27
   'If pa(9) = "020" And stCP10 = "1207" Then
   If pa(9) = "020" And (stCP10 = "1207" Or stCP10 = "1215") Then
   'end 2007/10/25
      If Text7.Text <> "" Then '有公開日
         m_HK_110Cp07 = CompDate(1, 6, Text7.Text)
         m_HK_110Cp06 = PUB_GetWorkDay1(CompDate(2, -5, CompDate(1, -1, m_HK_110Cp07)), True)
         '檢查有無香港
         If m_HaveHK = True Then
            'Modified by Lydia 2015/09/09 改共用模組
'            m_HaveHKInCP = Chk013Have110(pa(1), pa(2), pa(3), pa(4), m_HKMailID) 'Add by Morgan 2009/11/27 從他處移來
'            '檢查有無收香港的 110
'            If m_HaveHKInCP <> "" Then
'               '更新期限，上發 mail tag
'               strSql = "Update CaseProgress Set CP06=" & m_HK_110Cp06 & ",CP07=" & m_HK_110Cp07 & " Where CP09='" & m_HaveHKInCP & "' "
'               cnnConnection.Execute strSql
'               '更新齊備日
'               strSql = "update engineerprogress set ep06=" & ServerDate & " where ep02='" & m_HaveHKInCP & "' "
'               cnnConnection.Execute strSql
'               m_SendHKMail = True
'
'               If PUB_IfSetCP48(m_HaveHKInCP) Then 'Add by Morgan 2010/10/6
'
'                  '2010/1/19 add by sonia 更新承辦期限
'                  strPromoteDate = Pub_GetHandleDay(pa(1), "013", "110", , m_HK_110Cp06)
'                  If strPromoteDate <> "" Then
'                     strSql = "Update CaseProgress Set CP48=" & CNULL(strPromoteDate) & " Where CP09='" & m_HaveHKInCP & "' "
'                     cnnConnection.Execute strSql
'                  End If
'                  '2010/1/19 end
'
'               End If 'Add by Morgan 2010/10/6
'            End If
            Call PUB_UpdCP07by020(pa, m_bolFMP, "4", strSrvDate(1), DBDATE(Text7))
            'end 2015/09/09
            'Modified by Lydia 2015/12/25 改共用模組,中間少了發mail判斷
            m_HaveHKInCP = Chk013Have110(pa(1), pa(2), pa(3), pa(4), m_HKMailID)
            If m_HaveHKInCP <> "" Then
               m_SendHKMail = True
               'Added by Morgan 2020/3/5
               '寰華案要通知管制程序及主管
               If Pub_StrUserSt03 = "F22" Then
                  strExc(1) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
                  m_HKMailCCID = strExc(1) & ";" & PUB_GetFCPProSup(strExc(1))
               End If
               'end 2020/3/5
            End If
            'end 2015/12/25
            'Added by Lydia 2021/11/10 FMP大陸案Key公布通知時，請更新衍生的香港註冊案(已先收文)之Title(請抓FMP大陸案最新之Title)
            If m_bolFMP = True And m_HK_CP01 <> "" And m_HK_CP02 <> "" Then
                strSql = "update patent set pa05=" & CNULL(pa(5)) & ", pa06=" & CNULL(pa(6)) & ", pa07=" & CNULL(pa(7)) & _
                            " where PA01='" & m_HK_CP01 & "' AND PA02='" & m_HK_CP02 & "' AND PA03='" & m_HK_CP03 & "' AND PA04='" & m_HK_CP04 & "' "
                Pub_SeekTbLog strSql
                cnnConnection.Execute strSql, intI
            End If
            'end 2021/11/10
         'Add by Morgan 2009/11/25
         'FMP案通知公佈時若未收香港案則掛大陸案標準專利紀錄請求期限。
         ElseIf m_bolFMP Then
            strSql = "UPDATE NEXTPROGRESS SET NP08=" & m_HK_110Cp06 & ",NP09=" & m_HK_110Cp07 & _
               " WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP06 IS NULL AND NP07='110'"
            cnnConnection.Execute strSql, intI
            If intI = 0 Then
               strSql = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22)" & _
                  " SELECT '" & NewReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','110'" & _
                  "," & m_HK_110Cp06 & "," & m_HK_110Cp07 & ",'" & stCP13 & "'" & _
                  ",NP22 FROM (SELECT MAX(NP22)+1 NP22 FROM NEXTPROGRESS) X"
               cnnConnection.Execute strSql, intI
            End If
         End If
      End If
   End If
         
   
   'Add by Morgan 2009/11/24
   '更新主動修正期限
   If Text27 <> "" Then
      strSql = "Update Caseprogress Set CP06=" & DBDATE(Text15) & ",CP07=" & DBDATE(Text27) & _
         " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'  and cp10='203' and cp27 is null and cp57 is null"
      cnnConnection.Execute strSql, intI
      If intI = 0 Then
         strSql = "Update Nextprogress Set NP08=" & DBDATE(Text15) & ",NP09=" & DBDATE(Text27) & _
            ",NP23=" & DBDATE(Text6) & " where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07||np06='203'"
         cnnConnection.Execute strSql, intI
         'NIKON 要掛期限
         If intI = 0 And (InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X45148") > 0 _
            Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X45149") > 0 _
            Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X47405") > 0 _
            Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X47956") > 0 _
            Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X48220") > 0 _
            Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X48340") > 0 _
            Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X51585") > 0 _
            Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X53310") > 0 _
            Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X56040") > 0 _
            Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X60049") > 0) Then
            strSql = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22,NP23)" & _
               " SELECT '" & NewReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','203'" & _
               "," & DBDATE(Text15) & "," & DBDATE(Text27) & ",'" & stCP13 & "'" & _
               ",NP22," & DBDATE(Text6) & " FROM (SELECT MAX(NP22)+1 NP22 FROM NEXTPROGRESS) X"
            cnnConnection.Execute strSql, intI
         End If
      End If
   End If
   
   'Added by Morgan 2012/10/22
   '香港通知公佈要掛維持費期限=申請日+(公開日年-申請日年+5),若公開日月>申請日月則再 +1年;繳費年度為申請日起算
   If pa(9) = "013" And Text5(6) = "2" Then
      If pa(10) <> "" And Text7 <> "" Then
         '下次繳費日
         iYear = Val(Left(DBDATE(Text7), 4)) - Val(Left(DBDATE(pa(10)), 4)) + IIf(Right(Text7, 4) - Right(pa(10), 4) > 0, 6, 5)
         If iYear > 0 Then
            strNextDueDate = CompDate(0, iYear, pa(10))
            'Added by Lydia 2025/10/29
            'stNP23 = "" 'Mark by Lydia 2025/11/05
            If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
               strNextFeeDate = PUB_GetPOurDeadline(strNextDueDate, pa(9), stNP23, pa(1), "606")
            Else
            'end 2025/10/29
               If m_bolFMP Then
                  strNextFeeDate = CompDate(2, -10, strNextDueDate)
               Else
                  strNextFeeDate = CompDate(1, -1, strNextDueDate)
                  strNextFeeDate = CompDate(2, -5, strNextFeeDate)
               End If
            End If 'Added by Lydia 2025/10/29
            strNextFeeDate = PUB_GetWorkDay1(strNextFeeDate, True)
            '已收文
            strSql = "update caseprogress set cp06=" & strNextFeeDate & ",cp07=" & strNextDueDate & " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='606' and cp27||cp57 is null"
            cnnConnection.Execute strSql, intI
            '未收文
            If intI = 0 Then
               'Modified by Lydia 2025/10/29 +NP23
               strSql = "update nextprogress set np08=" & strNextFeeDate & ",np09=" & strNextDueDate & ",np23=" & IIf(stNP23 = "", "NP23", stNP23) & " where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07='606' and np06 is null"
               cnnConnection.Execute strSql, intI
               If intI = 0 Then
                  'Modified by Lydia 2025/10/29 +NP23
                  strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09," & _
                     "NP10,NP22,NP23) VALUES ('" & NewReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','606'," & strNextFeeDate & "," & strNextDueDate & ",'" & stCP13 & "',getnp22," & CNULL(stNP23, True) & ")"
                  cnnConnection.Execute strSql, intI
               End If
            End If
            
            '繳費紀錄
            stPA72 = "1"
            stPA73 = DBDATE(Text7)
            stPA74 = ""
            For intI = 2 To iYear
               stPA72 = stPA72 & "," & intI
               stPA73 = stPA73 & "," & DBDATE(Text7)
               stPA74 = stPA74 & ","
            Next
         End If
         strSql = "update patent set pa72=" & CNULL(stPA72) & ",pa73=" & CNULL(stPA73) & ",pa74=" & CNULL(stPA74) & _
            " Where PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "'"
         cnnConnection.Execute strSql, intI
         
      End If
   End If
   'end 2012/10/22
    
   'Added by Lydia 2015/04/28 大陸案輸入主管機關來函選擇：初審及進入實審，公布及進入實體審查或進入實體審查,請系統自動將下一程序的實體審查最終提申日上Y
   If pa(9) = "020" And InStr("3,4,5", Me.Text5(6).Text) > 0 Then
       strExc(9) = ""
       If PUB_ChkCPExist(cp(), "416", , strExc(9)) Then
          strSql = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE np01='" & strExc(9) & "' and np07='996' and np06 is null "
          cnnConnection.Execute strSql, intI
       End If
       'Added by Lydia 2015/09/10 大陸發明案收到實審通知,若該案有已收文未發文未取消收文之PPH431進度, 請畫面輸入之官方發文日+1個月去更新PPH431進度的本所期限, 法定期限空白，FMP案相同；
       If pa(8) = "1" Then
          strExc(9) = ""
          If PUB_ChkCPExist(cp(), "431", 1, strExc(9)) Then
            strExc(6) = PUB_GetWorkDay1(CompDate(1, 1, DBDATE(Text25)), 1)
            strSql = "UPDATE CASEPROGRESS SET CP06=" & strExc(6) & " , CP07=NULL WHERE CP09='" & strExc(9) & "'  "
            cnnConnection.Execute strSql, intI
          End If
       End If
   End If
   'Added by Lydia 2015/09/09 大陸發明案有輸入公開日時要更新下一程序999公開期限的NP06='Y'
   If pa(9) = "020" And pa(8) = "1" And Text7.Text <> "" And InStr(NewCasePtyList, cp(10)) > 0 Then
      strSql = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE np01='" & Label2(2).Caption & "' and np07='999' and np06 is null "
      cnnConnection.Execute strSql, intI
   End If
   'end 2015/09/09
   
   'Added by Lydia 2025/02/12 大陸案已發文延緩審查並且准予延緩審查;在進度自動create一道准予延緩審查，並更新下一程序催審期限
   If pa(9) <> "000" And m_bolAdd1924 = True Then
      strExc(1) = AutoNo("C", 6)
      strExc(2) = ""
      If m_bolFMP Then
         strExc(2) = PUB_GetCP20(pa(1), "1924", strCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
      End If
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
        "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43,CP64,CP133,CP134,cp119) VALUES ('" & Text1 & "','" & Text2 & "','" & _
        Text3 & "','" & Text4 & "'," & TransDate(Label2(3), 2) & _
        ",'" & strExc(1) & "','1924'," & CNULL(stCP12) & "," & _
        CNULL(stCP13) & ",'" & strUserNum & "','" & strExc(2) & "','N','N'," & strSrvDate(1) & ",'" & NewReceiveNo & "'" & _
        "," & CNULL(ChgSQL(strTmp(1))) & "," & CNULL(Text25, True) & "," & CNULL(Text26, True) & "," & DBDATE(Label2(3)) & ")"
      cnnConnection.Execute strSql, intI
      
      '更新下一程序催審期限
      strExc(0) = "select np01,np22,np09,cf05 from caseprogress,nextprogress,casefee" & _
         " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in (" & NewCasePtyList & ")" & _
         " and np01(+)=cp09 and np07='411' and np06 is null" & _
         " and cf01(+)=cp01 and cf02='" & pa(9) & "' and cf03(+)=cp10 and cf05>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(1) = CompDate(0, Val(m_st245CP71), "" & RsTemp.Fields("np09"))
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01='" & RsTemp("np01") & "' and np07='411' and np22=" & RsTemp("np22")
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2025/02/12
   
   'Add by Sindy 2016/10/5
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", NewReceiveNo, "")
      'Modified by Lydia 2023/05/18 +不開啟附件, , , False
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010514_1", IIf(Pub_StrUserSt03 = "F22", NewReceiveNo, ""), , , False
   End If
   '2016/10/5 END
   
   'Added by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail
   'Modified by Lydia 2023/05/26 已閉卷不通知
   'Move by Lydia 2023/05/26 從commit上方移過來,
   Dim bolFMP2mail As Boolean  'Added by Lydia 2023/05/26
   If m_bolFMP = True And m_bolFMP2 = True And pa(57) = "" Then
      'Modified by Lydia 2023/10/31 傳入C類收文號 NewReceiveNo
      bolFMP2mail = Pub_SetFMP2toCMail(pa(1), pa(2), pa(3), pa(4), stCP10, strUserNum, NewReceiveNo)
   End If
   'end 2023/05/17
   
   'Added by Morgan 2020/4/10
   'FMP有期限之案件EMAIL通知
   'Modified by Morgan 2020/9/8 排除香港案的公布通知，且增加大陸案的公布通知(不管有無下一程序)--敏莉
   'If m_bolFMP = True And Left(Pub_StrUserSt03, 1) <> "F" Then
   '   PUB_FMPCaseInform NewReceiveNo
   m_bolFMPNoPrint = False 'Added by Morgan 2023/4/11
   'Modified by Morgan 2023/5/25 FMP電子化所有來函都要EMail通知
   'If m_bolFMP = True And Not (pa(9) = "013" And Text5(6) = "2") Then
   'Modified by Lydia 2023/05/26 排除-寰華案無期限之官方來函，系統自動發Mail => And bolFMP2mail = False
   If m_bolFMP And bolFMP2mail = False Then
      'Modified by Morgan 2020/9/15 +寰華案,改通知智權人員
      'If Left(Pub_StrUserSt03, 1) <> "F" Then
      '   PUB_FMPCaseInform NewReceiveNo, IIf(pa(9) = "020" And Text5(6) = "2", False, True)
      'End If
      'Modified by Morgan 2022/12/13 改無期限也通知--Joanne
      'Modified by Morgan 2022/12/14 先還原，敏莉說無期限不必通知程序，等確認規則後再改
      PUB_FMPCaseInform NewReceiveNo, IIf(pa(9) = "020" And Text5(6) = "2", False, True), True, Left(Pub_StrUserSt03, 1) = "F"
      'end 2022/12/14
      'end 2022/12/13
      'end 2020/9/15
      
      m_bolFMPNoPrint = True 'Added by Morgan 2023/4/11
   'end 2020/9/8
   End If
   'end 2020/4/10
   
   'Added by Lydia 2022/04/28  工程師命名作業收文FMP主修和告代的承辦期限：(主修)待收到進入實質審查官方發文日+3個月
   If m_bolFMP And (Text5(6) = "3" Or Text5(6) = "4" Or Text5(6) = "5") Then
       Call Pub_GetFMPbCP48("2", cp, "203")
   End If
   'end 2022/04/28
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function

ErrorHandler:
   If bStartTrans Then cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function

Private Sub Form_Initialize()
   ReDim cp(1 To TF_CP) As String
   ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
Dim strTmp As String, bolChk As Boolean
   MoveFormToCenter Me
   intWhere = 國內
   
   With frm04010514_2
      pa(1) = .Text2
      pa(2) = .Text3
      pa(3) = .Text4
      pa(4) = .Text5
      strReceiveNo = .Tag
      ReadPatent
   End With
   Label2(2) = strReceiveNo
   Label2(3) = frm04010514_1.Text5
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010514_2.m_strIR01
   m_strIR02 = frm04010514_2.m_strIR02
   m_strIR03 = frm04010514_2.m_strIR03
   m_strIR04 = frm04010514_2.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2010/4/10
   Set frm04010514_3 = Nothing
End Sub

'************************************************
' 取回專利基本資料及收文資料
'
'************************************************
Private Sub ReadPatent()
 Dim Lbl As LABEL, i As Integer, strTmp As String, bolChk As Boolean, strTemp As String, strTemp1 As String
   For Each Lbl In Label2
      Lbl = ""
   Next
   If ClsPDReadPatentDatabase(pa(), intWhere) Then
      Label2(0) = pa(11): Label2(4) = pa(10)
      AddCboName Combo1, pa(5), pa(6), pa(7)
      Text7 = DBDATE(pa(12))
      Text8 = pa(13)
      If pa(9) = "013" Then
         Text5(13) = pa(14)
         Text11 = pa(15)
      End If
   End If
   
   cp(9) = strReceiveNo
   If ClsPDReadCaseProgressDatabase(cp, intWhere) Then
      bolChk = True
      If ClsPDGetCaseProperty(pa(1), cp(10), strTmp, bolChk) Then Label2(1) = strTmp
   End If
   
   Text1 = pa(1)
   Text2 = pa(2)
   Text3 = pa(3)
   Text4 = pa(4)
   
   Text5(7).Text = ""
   If pa(9) = 大陸國家代號 Then
      intI = 1
      '抓本案是否有收文(不論是否發文)資料
      strExc(0) = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='" & 實體審查 & "' AND CP57 IS NULL"
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      '若無收文資料時, 大陸實審收文預設為"N"
      If intI = 0 Then Text5(7) = "N"
   End If
   
   Label2(3) = frm04010514_1.Text5
   
   'Modified by Morgan 2021/1/28 從 Formsave 移來以便共用
   'If Left(cp(12), 1) = "F" And pa(10) <> "000" Then
   stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
   stCP12 = GetSalesArea(stCP13)
   'Modified by Lydia 2023/06/20 pa(10)=> pa(9)
   If Left(stCP12, 1) = "F" And pa(9) <> "000" Then
   'end 2021/1/28
      m_bolFMP = True
   Else
      m_bolFMP = False
   End If
   'Added by Lydia 2023/05/17 判斷寰華案
   m_bolFMP2 = False
   If m_bolFMP = True Then
      If PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4)) = True Then
         m_bolFMP2 = True
      End If
   End If
   'end 2023/05/17
   
   'Added by Lydia 2025/02/12 大陸案是否已發文延緩審查
   If pa(1) = "P" And pa(9) <> "000" Then
      strExc(0) = "select cp09,cp71 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='245' and cp158>0 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_st245CP09 = "" & RsTemp.Fields("cp09")
         m_st245CP71 = "" & RsTemp.Fields("cp71")
      End If
   End If
   'end 2025/02/12
   
End Sub

Private Sub Text5_GotFocus(Index As Integer)
   InverseTextBox Text5(Index)
   CloseIme
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 4, 7
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 6 '大陸/香港發明通知書
         '2006/2/27 MODIFY BY SONIA 加6香港公告
         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub Text5_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If Index = 6 Then
      SetDate
   End If
End Sub

Private Sub Text5_LostFocus(Index As Integer)
   Select Case Index
      Case 6
         '2006/2/7 ADD BY SONIA
         If Me.Text5(Index).Text = "2" And pa(9) = "013" Then
            'Modofied by Lydia 2025/07/08 政府憲報=>知識產權公報
            m_LetterType = MsgBox("香港公布通知請選擇定稿別？" & vbCrLf & vbCrLf & "公佈通知書請選 YES, 知識產權公報請選 NO", vbYesNo)
         End If
         '2006/2/7 END
   End Select
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
Dim lTmp As Long, i As Integer
Dim strTmp As String
   
   Select Case Index
      Case 0 '申請案核准日
         'Add By Cheng 2003/11/21若卷宗性質不為申請, 則直接跳開本段
         If pa(23) <> "1" Then Exit Sub
        
         If IsEmptyText(Text5(Index)) = False Then
            If CheckIsDate(Text5(Index)) Then
               If Val(TransDate(Text5(Index), 2)) > Val(strSrvDate(1)) Then
                  MsgBox "申請案核准日不可大於系統日 !", vbCritical
                  Cancel = True
               End If
            Else
               Cancel = True
            End If
         End If
         
      Case 6:
         If Me.Text5(Index).Text = "6" And pa(9) <> "013" And pa(9) <> "044" Then
            MsgBox "申請國家非香港案時不可選擇香港/澳門公告通知函 !", vbCritical
            Cancel = True
         End If
         
      Case 2 '機關文號
         If CheckLengthIsOK(Text5(Index), 40) = False Then
            Cancel = True
         End If
         
      Case 9
         If Text5(Index) <> "" Then
            Cancel = Not CheckIsDate(Text5(Index))
         End If
      Case 13
         If Text5(Index) <> "" Then
            If CheckIsDate(Text5(Index)) Then
               If Val(TransDate(Text5(Index), 2)) > Val(strSrvDate(1)) Then
                  MsgBox "不可大於系統日 !", vbCritical
                  Cancel = True
               End If
            Else
               Cancel = True
            End If
         End If
      
      
   End Select
   If Cancel Then TextInverse Text5(Index)
End Sub

Private Function TxtValidate() As Boolean
   Dim objTxt As Object
   Dim ii As Integer
   Dim Cancel As Boolean

   TxtValidate = False
   For Each objTxt In Text5
      If objTxt.Enabled = True Then
         Cancel = False
         Text5_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Me.Text5(objTxt.Index).SetFocus
            Text5_GotFocus objTxt.Index
            Exit Function
         End If
      End If
   Next
   
   If Text25 = "" Then
      MsgBox "官方發文日不可空白！"
      Text25.SetFocus
      Exit Function
   End If
   
   Cancel = False
   Text25_Validate Cancel
   If Cancel = True Then
      Text25.SetFocus
      Text25_GotFocus
      Exit Function
   End If
   
   'Add by Morgan 2009/11/23
   '大陸案通知進入實審時需更新主動修正期限
   If pa(9) = "020" And (Text5(6) = "3" Or Text5(6) = "4" Or Text5(6) = "5") Then
      
      If Text26 = "" Then
         MsgBox "請輸入期限月數！"
         Text26.SetFocus
         Exit Function
      End If
      
      If Text27 = "" Then
         MsgBox "請輸入法定期限！"
         Text27.SetFocus
         Exit Function
      End If
      
      Cancel = False
      Text27_Validate Cancel
      If Cancel = True Then
         Text27.SetFocus
         Text27_GotFocus
         Exit Function
      End If
      
      If Text15 = "" Then
         MsgBox "請輸入本所期限！"
         Text15.SetFocus
         Exit Function
      End If
      
      Cancel = False
      Text15_Validate Cancel
      If Cancel = True Then
         Text15.SetFocus
         Text15_GotFocus
         Exit Function
      End If
      
      If Text6 = "" Then
         MsgBox "請輸入約定期限！"
         Text6.SetFocus
         Exit Function
      End If
      
      Cancel = False
      Text6_Validate Cancel
      If Cancel = True Then
         Text6.SetFocus
         Text6_GotFocus
         Exit Function
      End If
      
   End If
   
   'Added by Morgan 2020/1/16
   '大陸案,有通知函,程序承辦,非掛號(無期限)
   m_bolNoCP27 = False
   'Removed by Morgan 2024/1/30 取消--郭
   'If pa(9) = "020" And Text5(4) <> "N" Then
   '   If PUB_GetCustomerValue(pa(26), "CU182") = "Y" Then
   '      If MsgBox("請確認是否已收到公文正本？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
   '         m_bolNoCP27 = True
   '      End If
   '   End If
   'End If
   'end 2020/1/16
   
   TxtValidate = True
   
End Function

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   If Text7 <> "" Then
      If CheckIsDate(Text7) = False Then
         Cancel = True
      End If
      If Cancel Then TextInverse Text7
   End If
End Sub

Private Sub Text8_GotFocus()
  TextInverse Text8
End Sub

Private Sub Text15_GotFocus()
   TextInverse Text15
End Sub

Private Sub Text15_Validate(Cancel As Boolean)
   If Text15 <> "" Then
      If Len(Text15) <> 8 Then
         MsgBox "必須輸入西元年！"
         Cancel = True
      ElseIf ChkDate(Text15) Then
         If Val(DBDATE(Text15)) < Val(strSrvDate(1)) Then
            MsgBox "本所期限不可小於系統日，請重新輸入 !", vbCritical
            Cancel = True
         End If
      End If
   End If
End Sub

Private Sub Text25_GotFocus()
   TextInverse Text25
End Sub

Private Sub Text25_Validate(Cancel As Boolean)
   If Text25 = "" Then
      MsgBox "官方發文日不可空白！"
      Cancel = True
      
   ElseIf Text25.Tag <> Text25 Then
      If Len(Text25) <> 8 Then
         MsgBox "必須輸入西元年！"
         Cancel = True
      ElseIf Not ChkDate(Text25) Then
         Cancel = True
      ElseIf Val(DBDATE(Text25)) > Val(strSrvDate(1)) Then
         MsgBox "官方發文日不可晚於系統日，請重新輸入 !", vbCritical
         Cancel = True
      End If
   End If
   
   If Cancel Then
      TextInverse Text25
   Else
      Text25.Tag = Text25
   End If
End Sub

Private Sub Text26_GotFocus()
   TextInverse Text26
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text26_Validate(Cancel As Boolean)
   If Text26 <> "" And Text26.Tag <> Text26 Then
      Cancel = Not SetDate
   End If
   If Cancel = False Then
      Text26.Tag = Text26
   End If
End Sub

Private Sub Text27_GotFocus()
   TextInverse Text27
End Sub

Private Sub Text27_Validate(Cancel As Boolean)
   If Text27 <> "" And Text27.Tag <> Text27 Then
      If Len(Text27) <> 8 Then
         MsgBox "必須輸入西元年！"
         Cancel = True
      ElseIf Not ChkDate(Text27) Then
         Cancel = True
      ElseIf Val(Text27) <= Val(strSrvDate(1)) Then
         MsgBox "法定期限必須大於系統日！"
         Cancel = True
      Else
         Cancel = Not SetDate2
      End If
   End If
   If Cancel = False Then
      Text27.Tag = Text27
   End If
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 <> "" Then
      If Len(Text6) <> 8 Then
         MsgBox "必須輸入西元年！"
         Cancel = True
      ElseIf ChkDate(Text6) Then
         If Val(DBDATE(Text6)) < Val(strSrvDate(1)) Then
            MsgBox "約定期限不可小於系統日，請重新輸入 !", vbCritical
            Cancel = True
         End If
      End If
   End If
End Sub

Private Function SetDate() As Boolean
   If pa(9) = "020" And (Text5(6) = "3" Or Text5(6) = "4" Or Text5(6) = "5") Then
      Frame1.Enabled = True
      If Text25 <> "" And Text26 <> "" Then
         strExc(1) = CompDate(1, Val(Text26), Text25)
         Text27 = strExc(1)
         Text27.Tag = Text27
      End If
      SetDate = SetDate2
   Else
      Frame1.Enabled = False
      Text27 = Empty
      Text15 = Empty
      Text6 = Empty
      SetDate = True
   End If
End Function

Private Function SetDate2() As Boolean
Dim strTmp As String 'Added by Lydia 2025/10/29

   If Text27 <> "" Then
      'Added by Lydia 2025/10/29
      If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
         strTmp = GetResult(IIf(Val(Text5(6)) = 0, "1", Text5(6)))
         Text15 = PUB_GetPOurDeadline(Text27, pa(9), Text6, pa(1), strTmp)
         If Val(Text15) < Val(strSrvDate(1)) Then
            Text15 = strSrvDate(1)
         End If
         If Val(Text6) < Val(strSrvDate(1)) Then
            Text6 = strSrvDate(1)
         End If
      Else
      'end 2025/10/29
         Text15 = PUB_GetWorkDay1(CompDate(2, -7, Text27), True)
         If Val(Text15) < Val(strSrvDate(1)) Then
            Text15 = strSrvDate(1)
         End If
         If m_bolFMP Then
            Text6 = PUB_GetWorkDay1(CompDate(2, -14, Text27), True)
         Else
            Text6 = PUB_GetWorkDay1(CompDate(2, -10, Text27), True)
         End If
         If Val(Text6) < Val(strSrvDate(1)) Then
            Text6 = strSrvDate(1)
         End If
      End If 'Added by Lydia 2025/10/29
   End If
   SetDate2 = True
End Function

'Added by Lydia 2025/09/23
Private Function GetResult(ByVal pVAL As String, Optional ByRef pRTitle As String) As String
   
   If pVAL = "" Then Exit Function
   GetResult = ""
   pRTitle = ""
   
   'Add By Cheng 2003/04/16
   '若申請國家非台灣
   If pa(9) <> 台灣國家代號 Then
      '判斷大陸香港發明通知書欄位
      Select Case pVAL
         Case "1"
            GetResult = "1213"
         Case "2"
            GetResult = "1207"
         Case "3"
            GetResult = "1214"
         Case "4"
            GetResult = "1215"
         Case "5"
            GetResult = "1204"
         '2006/2/7 ADD BY SONIA香港公告
         Case "6"
            GetResult = "1208"
      End Select
   End If
 
   'MODIFY BY SONIA 90.10.21
   If pVAL = "1" Then
      pRTitle = "初步審查合格通知書"
   End If
   If pVAL = "2" Then
      pRTitle = "公布通知書"
   End If
   'Add By Cheng 2002/06/21
   If pVAL = "3" Then
      pRTitle = "初步審查合格及進入實質審查程序通知書"
   End If
   If pVAL = "4" Then
      pRTitle = "公布及進入實質審查程序通知書"
   End If
   '93.1.6 ADD BY SONIA 93.1.6
   If pVAL = "5" Then
      pRTitle = "進入實質審查程序通知書"
   End If
   '93.1.6 END
   '2006/2/7 ADD BY SONIA 93.1.6
   If pVAL = "6" Then
      If pa(9) = "013" Then
         pRTitle = "香港公告"
      ElseIf pa(9) = "044" Then
         pRTitle = "澳門公告"
      Else
         pRTitle = "公告"
      End If
   End If
   '2006/2/7 END
End Function
