VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010505_2 
   Caption         =   "證書號數輸入"
   ClientHeight    =   5880
   ClientLeft      =   132
   ClientTop       =   960
   ClientWidth     =   8736
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   8736
   Begin VB.TextBox txtPA15 
      Height          =   270
      Left            =   5310
      MaxLength       =   20
      TabIndex        =   5
      Top             =   3960
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.TextBox txtPA14 
      Height          =   270
      Left            =   7245
      MaxLength       =   7
      TabIndex        =   2
      Top             =   3585
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text15 
      Height          =   270
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   11
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox Text14 
      Height          =   270
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   12
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox Text13 
      Height          =   270
      Left            =   1260
      TabIndex        =   10
      Top             =   5190
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      Height          =   270
      Left            =   5640
      MaxLength       =   8
      TabIndex        =   9
      Top             =   4890
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   8
      Top             =   4890
      Width           =   2505
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   1665
      TabIndex        =   7
      Top             =   4575
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7680
      TabIndex        =   15
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5664
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6492
      TabIndex        =   14
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   6
      Top             =   4245
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   3192
      MaxLength       =   8
      TabIndex        =   4
      Top             =   3930
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1752
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3915
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   4770
      MaxLength       =   7
      TabIndex        =   1
      Top             =   3585
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1260
      MaxLength       =   20
      TabIndex        =   0
      Top             =   3585
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   16
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   17
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   18
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   19
      Top             =   660
      Width           =   375
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1020
      TabIndex        =   20
      Top             =   960
      Width           =   6015
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "10610;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblPA15 
      AutoSize        =   -1  'True
      Caption         =   "公告號:"
      Height          =   180
      Left            =   4590
      TabIndex        =   56
      Top             =   3990
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblPA14 
      AutoSize        =   -1  'True
      Caption         =   "公告日:"
      Height          =   180
      Left            =   6525
      TabIndex        =   55
      Top             =   3630
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "列印客戶通知函:"
      Height          =   180
      Left            =   120
      TabIndex        =   54
      Top             =   5520
      Width           =   1305
   End
   Begin VB.Label Label30 
      Caption         =   "(N:不印)"
      Height          =   255
      Left            =   2160
      TabIndex        =   53
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "是否修改通知函內容:"
      Height          =   180
      Left            =   3840
      TabIndex        =   52
      Top             =   5520
      Width           =   1665
   End
   Begin VB.Label Label28 
      Caption         =   "(Y:Word)"
      Height          =   255
      Left            =   6240
      TabIndex        =   51
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2550
      TabIndex        =   50
      Top             =   4290
      Width           =   45
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "帳單金額:"
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   49
      Top             =   5220
      Width           =   765
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "帳單日期:"
      Height          =   180
      Index           =   2
      Left            =   4770
      TabIndex        =   48
      Top             =   4920
      Width           =   765
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "代理人D/N NO:"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   47
      Top             =   4920
      Width           =   1155
   End
   Begin VB.Label Label26 
      Height          =   255
      Left            =   4320
      TabIndex        =   46
      Top             =   3600
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   180
      X2              =   8520
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   180
      X2              =   8520
      Y1              =   3510
      Y2              =   3510
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "香港/澳門領證費:"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   45
      Top             =   4590
      Width           =   1350
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   1260
      TabIndex        =   44
      Top             =   3930
      Width           =   375
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "下次繳費日:"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   43
      Top             =   4230
      Width           =   945
   End
   Begin VB.Label Label23 
      Caption         =   "~"
      Height          =   255
      Left            =   2940
      TabIndex        =   42
      Top             =   3930
      Width           =   135
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "專用期限:  "
      Height          =   180
      Left            =   180
      TabIndex        =   41
      Top             =   3930
      Width           =   870
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "發證日期:"
      Height          =   180
      Left            =   3420
      TabIndex        =   40
      Top             =   3630
      Width           =   765
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   180
      TabIndex        =   39
      Top             =   3630
      Width           =   765
   End
   Begin VB.Label Label18 
      Height          =   255
      Left            =   1260
      TabIndex        =   38
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   180
      TabIndex        =   37
      Top             =   3120
      Width           =   945
   End
   Begin MSForms.Label Label16 
      Height          =   255
      Left            =   1140
      TabIndex        =   36
      Top             =   2820
      Width           =   7335
      VariousPropertyBits=   27
      Size            =   "12938;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Left            =   180
      TabIndex        =   35
      Top             =   2820
      Width           =   675
   End
   Begin MSForms.Label Label14 
      Height          =   255
      Left            =   1140
      TabIndex        =   34
      Top             =   2520
      Width           =   7335
      VariousPropertyBits=   27
      Size            =   "12938;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Left            =   180
      TabIndex        =   33
      Top             =   2520
      Width           =   675
   End
   Begin MSForms.Label Label12 
      Height          =   255
      Left            =   1140
      TabIndex        =   32
      Top             =   2220
      Width           =   7335
      VariousPropertyBits=   27
      Size            =   "12938;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Left            =   180
      TabIndex        =   31
      Top             =   2220
      Width           =   675
   End
   Begin MSForms.Label Label10 
      Height          =   255
      Left            =   1140
      TabIndex        =   30
      Top             =   1920
      Width           =   7335
      VariousPropertyBits=   27
      Size            =   "12938;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Left            =   180
      TabIndex        =   29
      Top             =   1920
      Width           =   675
   End
   Begin MSForms.Label Label6 
      Height          =   255
      Left            =   1140
      TabIndex        =   28
      Top             =   1620
      Width           =   7335
      VariousPropertyBits=   27
      Size            =   "12938;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Left            =   180
      TabIndex        =   27
      Top             =   1620
      Width           =   675
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   1140
      TabIndex        =   26
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   180
      TabIndex        =   25
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   24
      Top             =   960
      Width           =   765
   End
   Begin VB.Label Label2 
      Height          =   252
      Left            =   4380
      TabIndex        =   23
      Top             =   660
      Width           =   2652
   End
   Begin VB.Label Label19 
      Caption         =   "申請案號:"
      Height          =   252
      Left            =   3420
      TabIndex        =   22
      Top             =   660
      Width           =   972
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   21
      Top             =   660
      Width           =   768
   End
End
Attribute VB_Name = "frm04010505_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Combo1,Label6...)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim strReceiveNo As String
Dim DATE1 As String '專用期起日
Dim DATE2 As String '專用期止日
Dim m_strNextFeeDate As String  '下次繳費日本所期限
Dim m_strNextDueDate As String  '下次繳費日法定期限
'edit by nickc 2007/02/02 改動態
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim m_bln_keyinValidate As Boolean
Dim intWhere As Integer
Dim m_blnFormFirstShow As Boolean
Dim m_CP44 As String 'CF代理人
'Add by Morgan 2004/7/6
Dim m_bolNew As Boolean '是否用新法
'Add by Morgan 2004/7/15
Dim m_str421CP09 As String '技術報告總收文號
Dim m_str421EP06 As String '技術報告文件齊備日
Dim m_str421CP48 As String '技術報告承辦期限
'Add by Morgan 2004/8/3
Dim m_bol117Exist As Boolean '是否有收文積體電路佈局
Dim m_bolNoMsg As Boolean  '是否詢問
'Add by Morgan 2005/6/14
Dim m_iYear As Integer '香港應紀錄之已繳年度
'add by nickc 2005/06/16 'Memo by Lydia 2021/11/10 大陸發明案的衍生香港案
Dim m_HaveHK As Boolean
Dim m_HaveHKInCP As String
Dim m_HaveHKInNP As String
Dim m_SendHKMail As Boolean
Dim m_HK_CP01 As String
Dim m_HK_CP02 As String
Dim m_HK_CP03 As String
Dim m_HK_CP04 As String
'Added by Lydia 2021/11/10
Dim m_Have044 As Boolean, m_CPto044(1 To 4) As String '大陸發明案的衍生澳門案
Dim m_bolFMP2 As Boolean '是否為寰華
'end 2021/11/10
'add by nickc 2006/05/05
Dim m_HK_CP09 As String
Dim cm(7) As String
Dim m_HKMailID As String
Dim i As Integer
Dim m_NP07 As String '年費案件性質
'Add by Morgan 2006/5/25
Dim m_bAnnuityInform As Boolean '大陸年費通知
Dim m_bPaperPrompt As Boolean '換紙提示
Dim m_str605NP22 As String '年費NP22
Dim m_strNextFeeYear As String '下次繳費年度
Dim m_bolDualApply As Boolean '是否為大陸發明且有一案兩請案且新型未閉卷,2013/5/31+台灣也要
Dim m_DualCaseNo(1 To 4) As String '發明一案兩請案的新型案本所號
Dim m_blnCancelClosed As Boolean '是否取消閉卷
Dim m_HK_NP08 As String '香港案111本所期限
Dim m_HK_NP09 As String '香港案111法定期限 'Added by Morgan 2012/3/2
Dim m_bln605 As Boolean         '2008/12/12 ADD BY SONIA 大陸是否管制年費
Dim m_bolFMP As Boolean 'Add by Morgan 2009/12/2
Dim m_HK1913CP09 As String 'Added by Morgan 2016/6/15
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END
Dim stCP12 As String, stCP13 As String 'Added by Morgan 2021/1/28
Dim bolAddNP445 As Boolean, str445NP08 As String 'Added by Morgan 2021/6/10 是否管制455專利權期限補償
Dim m_bolNewMedInform As Boolean, m_PA176 As String 'Added by Morgan 2021/6/29 大陸新藥是否通知專利權期限延長
'Added by Morgan 2022/12/19
Public m_DocWord As String
Public m_DocNo As String
'end 2022/12/19
Dim m_bolFMPNoPrint As Boolean 'Added by Morgan 2023/4/11 FMP案是否列印中文定稿

Private Sub Form_Activate()
   Dim strTmp As String, strTmp1 As String
   
   'Add By Cheng 2003/04/02
   '若為第一次顯示
   If m_blnFormFirstShow = True Then
       m_blnFormFirstShow = False
   '若非第一次顯示
   Else
       Exit Sub
   End If
   Combo1.Clear
   
   Text1.Text = frm04010505_1.NUMBER1
   Text2.Text = frm04010505_1.NUMBER2
   Text3.Text = frm04010505_1.NUMBER3
   Text4.Text = frm04010505_1.NUMBER4
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010505_1.m_strIR01
   m_strIR02 = frm04010505_1.m_strIR02
   m_strIR03 = frm04010505_1.m_strIR03
   m_strIR04 = frm04010505_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   '基本檔資料
   'edit by nickc 2007/02/02 不用 dll 了
   'objPublicData.ReadPatentDatabase pA(), intWhere
   ClsPDReadPatentDatabase pa(), intWhere
   
   'Add by Morgan 2004/8/3
   '2007/4/24 MODIFY BY SONIA 改判斷申請國家及申請案號第三碼為 5 者為積體電路案件
   'm_bol117Exist = PUB_ChkCPExist(pa, "117")
   m_bol117Exist = False
   'Modify by Morgan 2010/12/28 申請案號改碼數
   'If pa(9) = 台灣國家代號 And Mid(pa(11), 3, 1) = "5" Then
   If pa(9) = 台灣國家代號 And Mid(pa(11), 4, 1) = "5" Then
      m_bol117Exist = True
   End If
   '2007/4/24 END
   
   'Add by Morgan 2010/4/28
   'Modified by Morgan 2021/1/28
   'strExc(1) = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
   'strExc(2) = GetSalesArea(strExc(1))
   ''Modify by Morgan 2010/8/3
   'If Left(strExc(2), 1) = "F" And pa(10) <> "000" Then
   stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
   stCP12 = GetSalesArea(stCP13)
   'Modified by Lydia 2023/06/20 pa(10)=> pa(9)
   If Left(stCP12, 1) = "F" And pa(9) <> "000" Then
   'end 2021/1/28
      m_bolFMP = True
   Else
      m_bolFMP = False
   End If
   'Added by Lydia 2021/11/10 判斷寰華案
   m_bolFMP2 = False
   If m_bolFMP = True Then
      m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4))
   End If
   'end 2021/11/10
   
   '若申請個為台灣發證日期欄位長度設為七碼否則為八碼
   If pa(9) = "000" Then
    Me.Text5.MaxLength = 7
    Me.Text9.MaxLength = 7
   Else
    Me.Text5.MaxLength = 8
    Me.Text9.MaxLength = 8
   End If
   
   If pa(9) = "000" Then
      'Modify by Morgan 2004/10/6 改西元年
      'Label8.Caption = "民國"
      Label8.Caption = "西元"
      
      Label26.Caption = "民國"
      Label27.Caption = "(法定期限)"
   Else
      Label8.Caption = "西元"
      Label26.Caption = "西元"
      Label27.Caption = "(本所期限)"
   End If
   
   Combo1.AddItem "中: " & pa(5), 0
   Combo1.Text = "中: " & pa(5)
   Combo1.AddItem "英: " & pa(6), 1
   Combo1.AddItem "日: " & pa(6), 2
   
   '申請案號
   Label2.Caption = pa(11)
   
   '專利號數
   If pa(9) = "020" Then
      Text7.Text = "ZL" + Label2.Caption
   Else
      Text7.Text = pa(22)
   End If
   
   '申請人1
   If pa(26) = Empty Then
      Label6.Caption = ""
   Else
      Label6.Caption = GetCustomerName(ChangeCustomerL(pa(26)))
   End If
   '申請人2
   If pa(27) = Empty Then
      Label10.Caption = ""
   Else
      Label10.Caption = GetCustomerName(ChangeCustomerL(pa(27)))
   End If
   '申請人3
   If pa(28) = Empty Then
      Label12.Caption = ""
   Else
      Label12.Caption = GetCustomerName(ChangeCustomerL(pa(28)))
   End If
   '申請人4
   If pa(29) = Empty Then
      Label14.Caption = ""
   Else
      Label14.Caption = GetCustomerName(ChangeCustomerL(pa(29)))
   End If
   '申請人5
   If pa(30) = Empty Then
      Label16.Caption = ""
   Else
      Label16.Caption = GetCustomerName(ChangeCustomerL(pa(30)))
   End If
   
   '來函收文日
   Label18.Caption = frm04010505_1.Text2.Text
      
   '發證日
   '若申請國為台灣
   If pa(9) = "000" Then
      '若有發證日期
      If Val(pa(21)) > 0 Then
         Text5.Text = TransDate(pa(21), 1)
      
      'Added by Morgan 2021/11/2 紙本可能會晚一天到, 改發證日預設為公告日(同FCP案)
      ElseIf Val(pa(14)) > 0 Then
         Text5.Text = TransDate(pa(14), 1)
         
      '若無發證日期
      Else
         Text5.Text = Label18.Caption
      End If
   '若申請國不為台灣
   Else
      '若有發證日期
      If Val(pa(21)) > 0 Then
         Text5.Text = TransDate(pa(21), 2)
      End If
   End If
   
   '專用期起迄
   '台灣帶民國年
   If pa(9) = "000" Then
      If Val(pa(24)) > 0 Then
         Text6.Text = TransDate(pa(24), 1)
      End If
      
      If Val(pa(25)) > 0 Then
         Text8.Text = TransDate(pa(25), 1)
      End If
      '2007/4/24 MODIFY BY SONIA
      'EnableTextBox Text9, True
      If m_bol117Exist = False Then
         EnableTextBox Text9, True
      Else
         EnableTextBox Text9, False
      End If
      '2007/4/24 END
   '非台灣帶西元年
   Else
      If Val(pa(24)) > 0 Then
         Text6.Text = TransDate(pa(24), 2)
      End If
      
      If Val(pa(25)) > 0 Then
         Text8.Text = TransDate(pa(25), 2)
      End If
      EnableTextBox Text9, False
   End If
      
   'Modify by Morgan 2006/6/1 基本檔資料不必重複抓
   strExc(0) = "SELECT DECODE(PA09,'000',PTM03,PTM04),NP08,NP09,NP22" & _
      "  FROM PATENT,NEXTPROGRESS,PATENTTRADEMARKMAP " & _
      " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND PTM01(+)='1' AND PTM02(+)=PA08" & _
      " AND NP02(+)=PA01 AND NP03(+)=PA02 AND NP04(+)=PA03" & _
      " AND NP05(+)=PA04 AND NP07(+)=605 AND NP06(+) IS NULL"
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      Exit Sub
   End If
   
   With RsTemp
      '專利種類
      Label4.Caption = "" & .Fields(0).Value
      '下次繳費日-所限
      m_strNextFeeDate = "" & .Fields("NP08")
      '下次繳費日-法限
      m_strNextDueDate = "" & .Fields("NP09")
      m_str605NP22 = "" & .Fields("NP22")
   End With
   
   'Add by Morgan 2006/7/5
   '大陸的下次繳費日一律重算,以防止修改申請日而期限未同步更新
   'Modify by Morgan 2007/3/27 加澳門
   'If pa(9) = "020" Then
   If pa(9) = "020" Or pa(9) = "044" Then
      '紀錄原期限,存檔前如發現有異動時提醒
      Text9.Tag = m_strNextFeeDate
      m_strNextFeeDate = ""
      m_strNextDueDate = ""
   End If
   
   '下次繳費日
   Text9.Text = Empty
   '有年費期限
   If Val(m_strNextDueDate) > 0 Then
      If GetMoneyDate(Val(pa(8)) + 10, pa(9), pa(), DATE1, strTmp1, DATE2) Then   '抓專用期起止日
         '93.1.8 add by sonia
         '香港案不必減一天
         If pa(9) = "013" Then
            'DATE2 = CompDate(2, 1, DATE2)   '2008/11/6 cancel by sonia香港案改在GetMoneyDate控制
            strTmp = CompDate(0, Val(strTmp1) - 1, DATE2)
         End If
         '93.1.8 end
      End If
   Else
      If GetMoneyDate(Val(pa(8)), pa(9), pa(), DATE1, strTmp1, DATE2) Then    '抓下次繳費日
         Dim varTmp As Variant, varTmp1 As Variant
         Dim i As Integer
        'Modify By Cheng 2002/12/03
        If pa(72) <> "" Then
            strTmp = pa(72)
            Do While InStr(strTmp, ",,") > 0
               strTmp = Replace(strTmp, ",,", ",")
            Loop
            varTmp = Split(strTmp, ",")
            i = UBound(varTmp)
            varTmp1 = Split(strTmp1, ",")
            If strTmp1 = "" Then i = 0
            strTmp = Format(varTmp(i))
            strTmp = CompDate(0, Val(strTmp), DATE2)
            m_strNextDueDate = strTmp
            If pa(9) < "010" Then
               'Added by Morgan 2014/10/28
               If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                  m_strNextFeeDate = PUB_GetOurDeadline(strTmp)
               Else
               'end 2014/10/28
                  m_strNextFeeDate = CompDate(2, -2, strTmp)
               End If 'Added by Morgan 2014/10/28
            Else
               'Add by Morgan 2010/4/28
               'Modified by Morgan 2018/10/3 非FMP也改10天
               'If m_bolFMP Then
                'Added by Lydia 2025/10/29
                If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                   m_strNextFeeDate = PUB_GetPOurDeadline(strTmp, pa(9))
                Else
               'end 2025/10/29
                   m_strNextFeeDate = CompDate(2, -10, strTmp)
                End If 'Added by Lydia 2025/10/29
               'Else
               'end 2010/4/28
               '   m_strNextFeeDate = CompDate(1, -1, strTmp)
               '   m_strNextFeeDate = CompDate(2, -5, m_strNextFeeDate)
               'End If
               'end 2018/10/3
               
            End If
            If GetMoneyDate(Val(pa(8)) + 10, pa(9), pa(), DATE1, strTmp1, DATE2) Then  '抓專用期起止日
               '93.1.8 add by sonia
               '香港案不必減一天
               If pa(9) = "013" Then
                  'DATE2 = CompDate(2, 1, DATE2)     '2008/11/6 cancel by sonia香港案改在GetMoneyDate控制
                  strTmp = CompDate(0, Val(strTmp1) - 1, DATE2)
               End If
               '93.1.8 end
            End If
           
        Else
            If pa(9) = "013" Then
            
               'Modify by Morgan 2005/6/13 依照94.6.10郭副理說明規則計算第一次續期費
               'Modify by Morgan 2005/9/2 加控制標準專利才要
               If pa(8) = "1" Then
                  Text5.Text = DBDATE(pa(21))
                  SetHongKongFeeDate
               Else
                  'DATE2 = CompDate(2, 1, DATE2) '香港案不必減一天  '2008/11/6 cancel by sonia香港案改在GetMoneyDate控制
                  'Modify by Morgan 2005/9/2 控制年費案件性質為605才要減1年
                  GetNP07 pa(9), pa(8), m_NP07
                  If m_NP07 = "605" Then
                     strTmp = CompDate(0, Val(strTmp1) - 1, DATE2)
                  Else
                     strTmp = CompDate(0, Val(strTmp1), DATE2)
                  End If
                  m_strNextDueDate = strTmp
                  'Add by Morgan 2010/4/28
                  'Modified by Morgan 2018/10/3 非FMP也改10天
                  'If m_bolFMP Then
                     'Added by Lydia 2025/10/29
                     If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                        m_strNextFeeDate = PUB_GetPOurDeadline(strTmp, pa(9))
                     Else
                     'end 2025/10/29
                        m_strNextFeeDate = CompDate(2, -10, strTmp)
                     End If 'Added by Lydia 2025/10/29
                  'Else
                  'end 2010/4/28
                  '   m_strNextFeeDate = CompDate(1, -1, strTmp)
                  '   m_strNextFeeDate = CompDate(2, -5, m_strNextFeeDate)
                  'End If
                  'end 2018/10/3
                  
                  If GetMoneyDate(Val(pa(8)) + 10, pa(9), pa(), DATE1, strTmp1, DATE2) Then  '抓專用期起止日
                     'DATE2 = CompDate(2, 1, DATE2) '香港案不必減一天  '2008/11/6 cancel by sonia香港案改在GetMoneyDate控制
                  End If
               End If
            'Add by Morgan 2007/3/27 澳門發明
            'MODIFY BY SONIA 2014/9/3 加設計P-103941
            ElseIf pa(9) = "044" And (pa(8) = "1" Or pa(8) = "3") Then
               GetNP07 pa(9), pa(8), m_NP07
               If m_NP07 = "605" Then
                  strTmp = CompDate(0, Val(strTmp1) - 1, DATE2)
               Else
                  strTmp = CompDate(0, Val(strTmp1), DATE2)
               End If
               m_strNextDueDate = strTmp
               'Add by Morgan 2010/4/28
               'Modified by Morgan 2018/10/3 非FMP也改10天
               'If m_bolFMP Then
                  'Added by Lydia 2025/10/29
                  If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                     m_strNextFeeDate = PUB_GetPOurDeadline(strTmp, pa(9))
                  Else
                  'end 2025/10/29
                     m_strNextFeeDate = CompDate(2, -10, strTmp)
                  End If 'Added by Lydia 2025/10/29
               'Else
               'end 2010/4/28
               '   m_strNextFeeDate = CompDate(1, -1, strTmp)
               '   m_strNextFeeDate = CompDate(2, -5, m_strNextFeeDate)
               'End If
               'end 2018/10/3
               
               GetMoneyDate Val(pa(8)) + 10, pa(9), pa(), DATE1, strTmp1, DATE2  '抓專用期起止日
            Else
               '2007/4/24 MODIFY BY SONIA
               'MsgBox "本案無已繳年度資料!!!", vbExclamation + vbOKOnly
               If m_bol117Exist = False Then
                  MsgBox "本案無已繳年度資料!!!", vbExclamation + vbOKOnly
               End If
               '2007/4/24 END
            End If
        End If
      End If
   End If
   
   '抓工作天
   If m_strNextFeeDate <> "" Then
      m_strNextFeeDate = PUB_GetWorkDay1(m_strNextFeeDate, True)
   End If
   ' 國內案依證書輸入下次繳年費法定期限, 由電腦檢查
   ' 非國內案由電腦計算後自動帶下次繳年費本所期限顯示於畫面
   If pa(9) <> "000" Then Text9.Text = TransDate(m_strNextFeeDate, 1)
            
   '代理人
   'Modified by Morgan 2014/11/24
   'm_CP44 = GetCP44()
   ClsPDGetCasePreAgent pa(), m_CP44, False
   '香港領證費
   Text10.Enabled = False
   
   'Add by Morgan 2004/6/30
   '台灣無公告日或93.7.1以後要輸入公告日
   m_bolNew = False
   lblPA14.Visible = False
   txtPA14.Visible = False
   
   'Add by Morgan 2006/1/16
   lblPA15.Visible = False
   txtPA15.Visible = False
   
   If pa(9) = 台灣國家代號 Then
      '公告日<93.7.1以前的新型專用期為12年
      If Val(pa(14)) > 0 And Val(pa(14)) < 930701 Then
         If pa(8) = "2" Then
            DATE2 = CompDate(2, -1, CompDate(0, 12, Val(pa(10)) + 19110000))
         End If
      '公告日93.7.1以後要輸入公告日
      Else
         m_bolNew = True
         lblPA14.Visible = True
         txtPA14.Visible = True
         'Add by Morgan 2004/8/17
         '台灣新法公報會回存證書號故若無發證日時要做雙重檢查
         If pa(21) = "" Then
            Select Case pa(8)
               Case "1": If m_bol117Exist = False Then Text7.Text = "I"
               Case "2": Text7.Text = "M"
               Case "3": Text7.Text = "D"
            End Select
            Text7.SelStart = 1
            Text7.SelLength = 0
         '若已發證則不用
         Else
            txtPA14.Text = pa(14)
         End If
      End If
      
      'Add by Morgan 2005/6/7
      '非台灣才可輸帳單資料
      Text11.Enabled = False
      Text12.Enabled = False
      Text13.Enabled = False
      
   Else
      'Add by Morgan 2005/6/7
      '非台灣才可輸帳單資料
      Text11.Enabled = True
      Text12.Enabled = True
      Text13.Enabled = True
   
      'Add by Morgan 2006/1/16
      If pa(9) = 大陸國家代號 Then
         lblPA15.Visible = True
         txtPA15.Visible = True
         txtPA15.Text = pa(15)
         txtPA15.Tag = txtPA15.Text
      '只有香港才可輸香港領證費
      'Modify by Morgan 2007/3/26 加澳門
      'ElseIf pa(9) = "013" Then
      'Modify by Morgan 2010/8/3
      'ElseIf pa(9) = "013" Or pa(9) = "044" Then
      ElseIf Not m_bolFMP And (pa(9) = "013" Or pa(9) = "044") Then
         Text10.Enabled = True
         Text10 = "5000"
      End If
   End If
   
   'Add by Morgan 2004/11/10 當下次繳費日>專用期止日時，下次繳費日控制不可輸入
   If Val(m_strNextDueDate) > Val(DATE2) And Val(DATE2) > 0 Then
      Text9.Text = "" 'Added by Morgan 2022/8/16
      m_strNextDueDate = ""
      EnableTextBox Text9, False
   End If
   
   '若申請國家為大陸'游標預設在發證日欄
   If pa(9) = 大陸國家代號 Then
      SendKeys "{Tab}"
   End If

   'Add by Morgan 2006/6/1
   If Val(Text9) > 0 And Val(Text9) < Val(strSrvDate(2)) Then
      MsgBox "下次繳費日已逾期!!", vbExclamation
   End If

   'Added by Morgan 2016/6/15 電子化-台灣案定稿要轉pdf故修改只能從定稿維護作業
   If pa(9) = "000" Then
      Text14.Enabled = False
   '非臺灣案電子化
   ElseIf (內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F") Then
      Text14.Enabled = False
   End If
   'end 2016/6/15
End Sub
Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(1 To 8) As String
   Dim strKey As String
   Dim ii As Integer
           
    ii = 1
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   '有輸入下次繳費日期才做
   If IsEmptyText(Text9) = False Then
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','本所期限','" & m_strNextFeeDate & "')"
        ii = ii + 1
   
      'Added by Morgan 2023/6/14
      If pa(72) <> "" Then 'Added by Morgan 2024/5/22
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','下次年費年度','" & (Replace(Right(pa(72), 2), ",", "") + 1) & "')"
            
         ii = ii + 1
      End If
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','下次年費法限','" & DBDATE(Text9) & "')"
      ii = ii + 1
      'end 2023/6/14
   End If
   
   If IsEmptyText(Text10) = False Then
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','大陸領證費','" & Text10 & "')"
        ii = ii + 1
        
'Removed by Morgan 2014/8/15 已掛應收可取消 --玲玲
'      'Add by Morgan 2009/1/16 香港案要帶點數 --敏惠
'      '香港案
'      If Val(Text10) > 0 Then
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'            "','點數','" & Val(Text10) / 1000 & "')"
'           ii = ii + 1
'      End If
      
   End If
   If pa(9) = "020" And pa(8) = "2" Then
      '內專抓代理人Y00000001
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','服務費','" & Val(PUB_GetYF06(pa(9), pa(8), "Y00000001", "601", "1", "1")) & "')"
        ii = ii + 1
   End If

   '2010/11/16 add by sonia
   If pa(9) = "020" Then
      If ET03 = "14" Then   '大陸新型20091001以前申請,定稿"14"判斷是否收文421加一段
         If PUB_ChkCPExist(pa, "421") = False Then
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','未收技術報告','♀')"
            ii = ii + 1
            strExc(0) = "SELECT NVL(CF08,0)+(NVL(CF13,0)*1000) FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='421' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','技術報告費用'," & CNULL(RsTemp.Fields(0).Value) & ")"
               ii = ii + 1
            End If
         End If
      ElseIf ET03 = "06" Then  '大陸新型設計20091001以後申請,定稿"06"判斷是否收文423加一段
         If PUB_ChkCPExist(pa, "423") = False Then
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','未收專利權評價','♀')"
            ii = ii + 1
            strExc(0) = "SELECT NVL(CF08,0)+(NVL(CF13,0)*1000) FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='423' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                  "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','專利權評價報告費用'," & CNULL(RsTemp.Fields(0).Value) & ")"
               ii = ii + 1
            End If
         End If
      End If
      
      'Added by Morgan 2021/6/10
      If m_HaveHKInCP <> "" Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','有收香港第二階段批准記錄請求才印','♀')"
         ii = ii + 1
      End If
      
      If bolAddNP445 = True Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','專利權期限補償-本所期限','" & str445NP08 & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','下一程序','445')"
         ii = ii + 1
      End If
      'end2021/6/10
   End If
   '2010/11/16 end

   'Added by Morgan 2021/6/29
   If m_bolNewMedInform Then
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','新藥通知專利權期限延長','♀')"
      ii = ii + 1
   End If
   'end 2021/6/29
   
    If ii <> 1 Then
         'edit by nickc 2007/02/05 不用 dll 了
         'If Not objLawDll.ExecSQL(ii - 1, strTxt) Then
         If Not ClsLawExecSQL(ii - 1, strTxt) Then
             MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
         End If
    End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim strTmp As String, strTmp2 As String
Dim strKey As String
Dim bolChk As Boolean
Dim bolCancel As Boolean
Dim rsRead As New ADODB.Recordset 'Added by Lydia 2016/12/02

Select Case Index
       Case 0
       If Text7.Text = "" Then MsgBox "輸入的專利號數不可為空值": Text7.SetFocus: Exit Sub
       If Text5.Text = "" Then MsgBox "發證日期不可為空值": Text5.SetFocus: Exit Sub
         'Add by Morgan 2004/7/1
         '檢查公告日
         If txtPA14.Visible = True Then
            bolCancel = False
            txtPA14_Validate bolCancel
            If bolCancel = True Then
               Me.txtPA14.SetFocus
               txtPA14_GotFocus
               Exit Sub
            End If
         End If
         
       If Text6.Text = "" Then MsgBox "專用期間不可為空值": Text6.SetFocus: Exit Sub
       If Text8.Text = "" Then MsgBox "專用期間不可為空值": Text8.SetFocus: Exit Sub
       
      '檢查專案截止日
      bolChk = False
      Text6_Validate bolChk
      If bolChk Then Exit Sub
      
      bolChk = False
      Text8_Validate bolChk
      If bolChk Then Exit Sub
       
      '2007/4/24 MODIFY BY SONIA
      'If pa(9) < "010" Then
      If pa(9) < "010" And m_bol117Exist = False Then
         'Modify by Morgan 2004/11/15
         If Text9.Locked = False Then
            If IsEmptyText(Text9) = True Then
               MsgBox "請輸入下次繳費日期"
               Text9_GotFocus
               Text9.SetFocus
               Exit Sub
            End If
            If DBDATE(Text9) <> DBDATE(Val(m_strNextDueDate)) Then
               MsgBox "下次繳費日期應為<" & TAIWANDATE(m_strNextDueDate) & ">", vbOKOnly, "檢核資料"
               Text9_GotFocus
               Text9.SetFocus
               Exit Sub
            End If
         End If
       End If
       
      'Add by Morgan 2004/11/15
      If Text9.Locked = False Then
      
          'Add by Morgan 2004/8/3
          '積體電路佈局無下次繳費日
          If m_bol117Exist = False Then
            If DBDATE(Text9) < strSrvDate(1) Then
              MsgBox "下次繳費日期不可小於系統日期", vbOKOnly, "檢核資料"
              Text9.SetFocus
              Exit Sub
            End If
         End If
         
      End If
       '92.12.24 END
        'Add By Cheng 2002/10/31
        If (Me.Text11.Text = "" Xor Me.Text12.Text = "") Or (Me.Text11.Text = "" Xor Me.Text13.Text = "") Or (Me.Text12.Text = "" Xor Me.Text13.Text = "") Then
            MsgBox "代理人D/N NO , 帳單日期 及 帳單金額 " & Chr(10) & Chr(13) & "三欄位必須同時輸入或不輸入資料!!!", vbExclamation + vbOKOnly
            If Me.Text11.Text = "" Then Me.Text11.SetFocus:     Text11_GotFocus:     Exit Sub
            If Me.Text12.Text = "" Then Me.Text12.SetFocus:     Text12_GotFocus:     Exit Sub
            If Me.Text13.Text = "" Then Me.Text13.SetFocus:     Text13_GotFocus:     Exit Sub
        End If
              
      'Add By Cheng 2002/05/22
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
         If Text10.Enabled Then
            If pa(9) = "013" And Text10 = "" Then MsgBox "香港專利請輸入香港領證費", vbInformation: Text10.SetFocus: Exit Sub
            'Add by Morgan 2007/3/26
            If pa(9) = "044" And Text10 = "" Then MsgBox "澳門專利請輸入澳門領證費", vbInformation: Text10.SetFocus: Exit Sub
         End If
        'Add By Cheng 2002/11/05
        frm04010505_1.lblBillNo.Caption = ""
         'Add by Morgan 2005/5/20
         '非台灣 詢問是否計算結餘
        'Modified by Lydia 2015/03/03 +pa01,pa02,pa03,pa04
        Pub_EndModCashMsg pa(9), pa(1), pa(2), pa(3), pa(4)
         
         'add by nickc 2005/06/16 取有關大陸的香港關聯判斷
         m_HaveHK = False
         m_HaveHKInCP = ""
         m_HaveHKInNP = ""
         m_SendHKMail = False
         m_HKMailID = ""
         'add by nickc 2006/05/05
         m_HK_CP09 = ""
         m_Have044 = False 'Added by Lydia 2021/11/10
         
         'Modified by Morgan 2014/9/23 大陸發明才要--郭
         'If pa(9) = "020" Then
         If pa(8) = "1" And pa(9) = "020" Then
            'edit by nickc 2006/05/05
            'm_HaveHK = ChkCMIsExist013(pA(1), pA(2), pA(3), pA(4))
            'Modified by Morgan 2014/9/23 香港標準專利(發明)才要--郭
            'Modified by Morgan 2016/9/7 +判斷香港案要未閉卷(Ex.陸:P102572 港P-105613)--玲玲
            'm_HaveHK = ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), m_HK_CP01, m_HK_CP02, m_HK_CP03, m_HK_CP04, m_HK_CP09, "1")
            m_HaveHK = ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), m_HK_CP01, m_HK_CP02, m_HK_CP03, m_HK_CP04, m_HK_CP09, "1", , True)
            'end 2016/9/7
            If m_HaveHK = True Then
               'For i = 1 To 4
               '   cm(i - 1) = pA(i)
               'Next
               'If obj003.GetCaseMap(cm, 4) = True Then
               '   m_HK_CP01 = cm(4)
               '   m_HK_CP02 = cm(5)
               '   m_HK_CP03 = cm(6)
               '   m_HK_CP04 = cm(7)
               'End If
               m_HaveHKInCP = Chk013Have111(pa(1), pa(2), pa(3), pa(4), m_HKMailID)
               If m_HaveHKInCP = "" Then 'Added by Morgan 2015/12/25 '沒有CP才抓NP否則m_HKMailID會被清除
                  m_HaveHKInNP = Chk013Have111(pa(1), pa(2), pa(3), pa(4), m_HKMailID, "NP")
               End If
            End If
            
            'Added by Lydia 2021/11/10 澳門案
            m_Have044 = ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), m_CPto044(1), m_CPto044(2), m_CPto044(3), m_CPto044(4), , "1", "5")
            
            'Add by Morgan 2006/7/5
            If pa(9) = "020" Then
               If Text9.Tag <> "" And Text9.Tag <> TransDate(Text9.Text, 2) Then
                  If MsgBox("下次繳費日將更新【 " & TransDate(Text9.Tag, 1) & " --> " & Text9.Text & " 】，是否要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                     Exit Sub
                  End If
               End If
            End If
         End If
         
         'Add by Morgan 2006/9/28
         m_bolDualApply = False
         'Modified by Morgan 2012/8/21 +2009/10/1以後申請案件(應可改為所有一案兩請之發明案發證都可適用)
         'If pa(9) = "020" And pa(8) = "1" And Val(pa(10)) < 950701 Then
         'Modified by Morgan 2013/5/31 +台灣也要
         'If pa(9) = "020" And pa(8) = "1" And (Val(pa(10)) < 950701 Or Val(pa(10)) >= 981001) Then
         If pa(8) = "1" And (pa(9) = "000" Or (pa(9) = "020" And (Val(pa(10)) < 950701 Or Val(pa(10)) >= 981001))) Then
            If PUB_IsDualApply(pa(), m_DualCaseNo()) = True Then
               strExc(0) = "SELECT 1 FROM PATENT WHERE PA01='" & m_DualCaseNo(1) & "' AND PA02='" & m_DualCaseNo(2) & "' AND PA03='" & m_DualCaseNo(3) & "' AND PA04='" & m_DualCaseNo(4) & "' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  m_bolDualApply = True
               End If
            End If
         End If
         
         'Add by Morgan 2006/9/29
         If m_bolDualApply = True Then
            'Added by Morgan 2013/6/14 台灣案要檢查有發文放棄專利權429
            If pa(9) = "000" Then
               'Modified by Morgan 2015/7/9 一案兩請是否放棄新型改放PA60
               'If pa(162) <> "Y" Then 'Added by Morgan 2014/7/29 若輸審查意見時已選擇放棄新型則此處不必再確認
               If pa(60) <> "Y" Then
               'end 2015/7/9
                  If PUB_ChkCPExist(m_DualCaseNo, "429", 2) = True Then
                     MsgBox "本案與新型案(" & m_DualCaseNo(1) & m_DualCaseNo(2) & m_DualCaseNo(3) & m_DualCaseNo(4) & ")為一案兩請，新型案已發文放棄專利權將自動上閉卷！(若該新型案尚有年費期限也將會自動不續辦)", vbInformation
                  Else
                     m_bolDualApply = False
                  End If
               End If 'Added by Morgan 2014/7/29
            Else
            'end 2013/6/14
               'Modified by Morgan 2016/8/22 大陸案一案兩請的案件發明案發證時新型案沒有收文放棄專利權的案件不可上閉卷及不續辦  --玲玲
               'MsgBox "本案與新型案(" & m_DualCaseNo(1) & m_DualCaseNo(2) & m_DualCaseNo(3) & m_DualCaseNo(4) & ")為一案兩請，新型案將自動上閉卷！(若該新型案尚有年費期限也將會自動不續辦)", vbInformation
               If PUB_ChkCPExist(m_DualCaseNo, "429") = True Then
                  MsgBox "本案與新型案(" & m_DualCaseNo(1) & m_DualCaseNo(2) & m_DualCaseNo(3) & m_DualCaseNo(4) & ")為一案兩請，新型案將自動上閉卷！(若該新型案尚有年費期限也將會自動不續辦)", vbInformation
               Else
                  m_bolDualApply = False
               End If
               'end 2016/8/22
            End If
         End If
         
         'Add by Morgan 2007/5/10 若來函有期限但已閉卷
         m_blnCancelClosed = False: m_bln605 = True
         If pa(57) = "Y" And Text9.Text <> "" Then
            '2008/12/12 MODIFY BY SONIA 大陸案改為不取消閉卷仍存檔但不產生年費期限但定稿仍要印出年費期限
            'If MsgBox("本案目前為閉卷狀態，為管制期限將於存檔時取消閉卷，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
            '   Screen.MousePointer = vbDefault: Exit Sub
            'End If
            'm_blnCancelClosed = True
            If pa(9) <> "020" Then
               If MsgBox("本案目前為閉卷狀態，為管制期限將於存檔時取消閉卷，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
                  Screen.MousePointer = vbDefault: Exit Sub
               End If
               m_blnCancelClosed = True
            Else
               If MsgBox("本案目前為閉卷狀態，是否取消閉卷？（是->管制年費，否->不產生年費期限）", vbYesNo + vbDefaultButton2) = vbYes Then
                  m_blnCancelClosed = True   '取消閉卷並新增年費期限
               Else
                  m_bln605 = False           '不取消閉卷且不新增年費期限
               End If
            End If
            '2008/12/12 END
         End If
         'end 2007/5/10
         
         'Added by Morgan 2021/6/29
         '通知新藥發明專利權期限補償(大陸2021.6.1新法)
         m_bolNewMedInform = False
         m_PA176 = pa(176)
         If pa(9) = "020" And pa(8) = "1" And pa(158) = "3" And DBDATE(pa(20)) >= "20210601" Then
            '若尚未設定是否新藥
            If m_PA176 = "" Then
               intI = MsgBox("本案是否為新藥專利？", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
               If intI = vbYes Then
                  m_PA176 = "Y"
               ElseIf intI = vbNo Then
                  m_PA176 = "N"
               Else
                  Exit Sub
               End If
            End If
            
            If m_PA176 = "Y" Then
               'Modified by Morgan 2022/3/2 有設定過都要通知,不必再確認 -- 郭
               'intI = MsgBox("本案為新藥專利，請先交承辦工程師確認是否通知專利權期限延長，確認後請選擇是或否！", vbYesNoCancel + vbDefaultButton3 + vbQuestion, "是否通知專利權期限延長？")
               'If intI = vbYes Then
               '   m_bolNewMedInform = True
               'ElseIf intI = vbNo Then
               '   m_bolNewMedInform = False
               'Else
               '   Exit Sub
               'End If
               m_bolNewMedInform = True
               'end 2022/3/2
            End If
         End If
         'end 2021/6/29
         'Added by Lydia 2023/03/10 FMP大陸新藥發明專利權期限補償控管
         If m_bolFMP = True And pa(9) = "020" And pa(8) = "1" And pa(150) = "2" Then
             If m_PA176 = "Y" Then  '外專命名記錄設定(frm090902_2)
                m_bolNewMedInform = True
             End If
         End If
         'end 2023/03/10
         
         'Add By Sindy 2022/7/1
         'Mark by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail:可取消外專系統收件區，key來函承辦人掛程序人員，則按確定，信件會再打開一次的設定。
         'If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
         '   If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
         '      Screen.MousePointer = vbDefault
         '      Exit Sub
         '   End If
         'End If
         '2022/7/1 END
         'end 2023/05/17
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         If Text15 <> "N" Then '通知函
            If Text14 = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            
            strKey1 = "1"
            strTmp = "01"
            Select Case pa(8)
               Case "1"
                  If pa(9) = 台灣國家代號 Then '台灣
                     'Added by Morgan 2023/6/14 寶齡富錦 Y55435 案件
                     If pa(75) = "Y55435" Then
                        strTmp = "17"
                     'end 2023/6/14
                     '大-->台 定稿 20080915  add by Toni
                     ElseIf PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
                        strTmp = "15"
                     Else
                        strTmp = "01"
                     
                        'Add by Morgan 2004/7/14
                        '新法需加註公告日
                        If txtPA14.Visible = True Then
                           strTmp = "07"
                        End If
                     End If
                  ElseIf pa(9) = "020" Then '大陸
                     'edit by nickc 2005/06/17
                     'strTmp = "02"
                     'Modified by Morgan 2021/6/10 合併(改用例外欄位控制)
                     'If m_HaveHK = False Then
                     '   strTmp = "02"
                     'Else
                     '   If m_HaveHKInCP <> "" Then
                     '      strTmp = "11"
                     '   Else
                     '      strTmp = "02"
                     '   End If
                     'End If
                     strTmp = "02"
                     'end 2021/6/10
                  ElseIf pa(9) = "013" Then '香港 3
                     strTmp = "03"
                     'Added by Morgan 2021/3/29 寶齡富錦 Y55435 案件
                     If pa(75) = "Y55435" Then
                        strTmp = "16"
                     End If
                     
                  'Add by Morgan 2007/3/26
                  ElseIf pa(9) = "044" Then '澳門
                     strTmp = "13"
                  End If
               Case "2"
                  If pa(9) = 台灣國家代號 Then '台灣
                     '大-->台 定稿 20080915  add by Toni
                     If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
                        strTmp = "15"
                     Else
                        strTmp = "01"
                   
                     'Add by Morgan 2004/7/8
                     '新法新型發證加註可提技術報告
                     If txtPA14.Visible = True Then
                        '是否已收文技術報告
                        If PUB_ChkCPExist(pa, "421") Then
                           strTmp = "07"
                        Else
                           strTmp = "08"
                        End If
                     End If
                     
                     End If
                  ElseIf pa(9) = "020" Then '大陸
                     'Modify by Morgan 2007/7/16
                     '是否已收文檢索報告
                     'strTmp = "06"
                     '2010/11/16 modify by sonia 分20091001以前(判斷421)或以後(判斷423)申請的案件
                     'If PUB_ChkCPExist(pa, "421") Then
                     '   strTmp = "14"
                     'Else
                     '   strTmp = "06"
                     'End If
                     If Val(DBDATE(pa(10))) < 20091001 Then '未收421或423改用例外欄位加一段
                        strTmp = "14"
                     Else
                        strTmp = "06"
                     End If
                     '2010/11/16 end
                  ElseIf pa(9) = "013" Then '香港 4
                     strTmp = "04"
                  End If
               Case "3"
                  If pa(9) = 台灣國家代號 Then '台灣
                     '大-->台 定稿 20080915  add by Toni
                     If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
                        strTmp = "15"
                     Else
                        strTmp = "01"
                     
                     'Add by Morgan 2004/7/14
                     '新法需加註公告日
                     If txtPA14.Visible = True Then
                        strTmp = "07"
                     End If
                    End If
                  ElseIf pa(9) = "020" Then '大陸
                     '2010/11/16 modify by sonia 分20091001以前(02)或以後(判斷有無收文423)申請的案件
                     'strTmp = "02"
                     If Val(DBDATE(pa(10))) < 20091001 Then
                        strTmp = "02"
                     Else
                        strTmp = "06"              '未收423用例外欄位加一段
                     End If
                     '2010/11/16 end
                  ElseIf pa(9) = "013" Then '香港 5
                     strTmp = "05"
                  End If
            End Select
            'Add by Morgan 2004/8/3
            '積體電路佈局
            If m_bol117Exist = True Then
               If pa(9) = "020" Then
                  strTmp = "09"
               ElseIf pa(9) = "000" Then
                  strTmp = "10"
               End If
            End If
            StartLetter "08", strTmp
            'Add by Morgan 2009/12/2
            'Modify by Morgan 2010/3/31 +設計定稿
            If m_bolFMP Then
               'Modified by Morgan 2023/4/11 +m_bolFMPNoPrint
               NowPrint strReceiveNo, "08", strTmp, bolChk, strUserNum, , , , , 1, , , , , , , , strReceiveNo, , , , , m_bolFMPNoPrint
               
               '發證
               'Modified by Morgan 2019/12/11 取消設計案定稿(實際內容與非設計案相同)--David
               'If pa(8) = "3" Then
               '   strTmp2 = "53"
               'Else
                  'Added by Morgan 2012/7/11 +香港發明定稿
                  If pa(9) = "013" And pa(8) = "1" Then
                     strTmp2 = "56"
                  'Added by Morgan 2024/6/13 +香港設計定稿--Anny
                  ElseIf pa(9) = "013" And pa(8) = "3" Then
                     strTmp2 = "53"
                  'end 2024/6/13
                  Else
                     strTmp2 = "51"
                  End If
               'End If
               'end 2019/12/22
               strUserNum = strFMPNum
               StartLetter2 "08", strTmp2
               'Modified by Morgan 2016/5/30 不可傳LD18否則FCP承辦執行定維護時會開E化的畫面
               'NowPrint strReceiveNo, "08", strTmp2, False, strUserNum, , , , , , , , , , , , , strReceiveNo
               NowPrint strReceiveNo, "08", strTmp2, False, strUserNum
               
               'Added by Morgan 2024/6/13
               '香港設計案證書為紙本，且已有中英對照，故無需再準備譯文--Anny
               If Not (pa(9) = "013" And pa(8) = "3") Then
               'end 2024/6/13
               
                  '譯文
                  If pa(8) = "3" Then
                     strTmp2 = "54"
                  Else
                     strTmp2 = "52"
                  End If
                  StartLetter3 "08", strTmp2
                  'Modified by Morgan 2016/5/30 不可傳LD18否則FCP承辦執行定維護時會開E化的畫面
                  'NowPrint strReceiveNo, "08", strTmp2, False, strUserNum, , , , , , , , , , , , , strReceiveNo
                  NowPrint strReceiveNo, "08", strTmp2, False, strUserNum
                  
               End If 'Added by Morgan 2024/6/13
               
               strUserNum = strUser1Num
               
            Else
            'end 2009/12/2
               NowPrint strReceiveNo, "08", strTmp, bolChk, strUserNum, , , , , , , , , , , , , strReceiveNo
            End If
            
            
            'add by nickc 2005/06/17 若是大陸案，要檢查香港，若未收文且已有下一程序資料要家印此定稿
            'edit by nickc 2006/05/04  印定稿不管有沒有 NP
            'If pa(8) = "1" And pa(9) = "020" And m_HaveHK = True And m_HaveHKInCP = "" And m_HaveHKInNP <> "" Then
            If pa(8) = "1" And pa(9) = "020" And m_HaveHK = True And m_HaveHKInCP = "" Then
               'StartLetter "08", strTmp
               EndLetter "08", m_HK_CP01 & m_HK_CP02 & m_HK_CP03 & m_HK_CP04 & "&000", "12", strUserNum
               Dim strTxt(2) As String
               'Modify by Morgan 2008/6/30
               'strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('08','" & m_HK_CP01 & m_HK_CP02 & m_HK_CP03 & m_HK_CP04 & "&000" & "','12','" & strUserNum & _
                           "','本所期限','" & m_strNextFeeDate & "')"
               strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('08','" & m_HK_CP01 & m_HK_CP02 & m_HK_CP03 & m_HK_CP04 & "&000" & "','12','" & strUserNum & _
                           "','本所期限','" & m_HK_NP08 & "')"
                           
               'add by nickc 2006/05/11
               strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('08','" & m_HK_CP01 & m_HK_CP02 & m_HK_CP03 & m_HK_CP04 & "&000" & "','12','" & strUserNum & _
                   "','下一程序','111')"
               
                'edit by nickc 2007/02/05 不用 dll 了
                'If Not objLawDll.ExecSQL(2, strTxt) Then
                If Not ClsLawExecSQL(2, strTxt) Then
                    MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
                End If
                           
               'edit by nickc 2006/05/09
               'NowPrint m_HaveHKInNP, "08", strTmp, bolChk, strUserNum, 0
               'NowPrint strReceiveNo, "08", strTmp, bolChk, strUserNum, 0
               'Add by Morgan 2009/12/2
               If m_bolFMP Then
                  NowPrint m_HK_CP01 & m_HK_CP02 & m_HK_CP03 & m_HK_CP04 & "&000", "08", "12", bolChk, strUserNum, , , , , 1, , , , , , , , m_HK1913CP09
                  
                  'Added by Morgan 2012/3/2
                  strUserNum = strFMPNum
                  StartLetter4 "08", m_HK_CP01 & m_HK_CP02 & m_HK_CP03 & m_HK_CP04 & "&000", "55"
                  'Modified by Morgan 2016/5/30 不可傳LD18否則FCP承辦執行定維護時會開E化的畫面
                  NowPrint m_HK_CP01 & m_HK_CP02 & m_HK_CP03 & m_HK_CP04 & "&000", "08", "55", False, strUserNum
                  strUserNum = strUser1Num
                  
               Else
               'end 2009/12/2
                  NowPrint m_HK_CP01 & m_HK_CP02 & m_HK_CP03 & m_HK_CP04 & "&000", "08", "12", bolChk, strUserNum, , , , , , , , , , , , , m_HK1913CP09
               End If
            End If
            
'Remove by Morgan 2016/6/15 2010/9/3已取消該通知
'            'Add by Morgan 2006/5/25
'            '大陸催年費通知函
'            If m_bAnnuityInform = True Then
'               'Modify by Morgan 2009/6/3 改用新定稿
'               'strTmp = "02"
'               strTmp = "21"
'               'end 2009/6/3
'
'               EndLetter "12", pa(1) & pa(2) & pa(3) & pa(4) & "&605", strTmp, strUserNum
'               strExc(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                               "('12','" & pa(1) & pa(2) & pa(3) & pa(4) & "&605','" & strTmp & "','" & strUserNum & "','本所期限'," & CNULL(m_strNextFeeDate) & ")"
'               strExc(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                               "('12','" & pa(1) & pa(2) & pa(3) & pa(4) & "&605','" & strTmp & "','" & strUserNum & "','法定期限'," & CNULL(m_strNextDueDate) & ")"
'               strExc(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                               "('12','" & pa(1) & pa(2) & pa(3) & pa(4) & "&605','" & strTmp & "','" & strUserNum & "','下一程序','605,606')"
'
'               '取得下次繳費年度-只考慮最簡單的狀況
'               If pa(72) = "" Then
'                  m_strNextFeeYear = "1"
'               Else
'                  strExc(0) = Right(pa(72), 2)
'                  If Left(strExc(0), 1) = "," Then
'                     strExc(0) = Right(strExc(0), 1)
'                  End If
'                  m_strNextFeeYear = Val(strExc(0)) + 1
'               End If
'               strExc(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                  "('12','" & pa(1) & pa(2) & pa(3) & pa(4) & "&605','" & strTmp & "','" & strUserNum & "','費用','" & Val(PUB_GetYF0607(pa(9), pa(8), "Y00000001", "605", m_strNextFeeYear, m_strNextFeeYear)) & "')"
'               'edit by nickc 2007/02/05 不用 dll 了
'               'If Not objLawDll.ExecSQL(4, strExc) Then
'               If Not ClsLawExecSQL(4, strExc) Then
'                   MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
'               End If
'
'               'Add by Morgan 2009/12/2
'               If m_bolFMP Then
'                  NowPrint pa(1) & pa(2) & pa(3) & pa(4) & "&605", "12", strTmp, bolChk, strUserNum, , , , , 1, , , , , , , , strReceiveNo
'               Else
'               'end 2009/12/2
'                  NowPrint pa(1) & pa(2) & pa(3) & pa(4) & "&605", "12", strTmp, bolChk, strUserNum, , , , , , , , , , , , , strReceiveNo
'               End If
'            End If
'end 2016/6/15
            
            'add by nickc 2005/06/17 發 mail
            If m_SendHKMail = True And m_HKMailID <> "" And m_HaveHKInCP <> "" Then
               Call PUB_SendMail(strUserNum, m_HKMailID, m_HaveHKInCP, "大陸案(" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & ")已發證，香港案(" & m_HK_CP01 & "-" & m_HK_CP02 & "-" & m_HK_CP03 & "-" & m_HK_CP04 & ")的[批准紀錄請求]可以處理！", "大陸案(" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & ")已發證，香港案(" & m_HK_CP01 & "-" & m_HK_CP02 & "-" & m_HK_CP03 & "-" & m_HK_CP04 & ")的[批准紀錄請求]可以處理！", "")
            End If
         End If
       'Add by Lydia 2014/11/18 台灣案主管機關來函輸入，若此案有工程師未發文的程序，發E-MAIL通知工程師收到來函的內容
         'Modified by Lydia 2022/08/15 開放P大陸案
         'If pa(9) = "000" And pa(1) = "P" Then
         'Modified by Lydia 2022/10/11 經查此設定並不適用於外專及日專，故請協助排除FMP案
         'If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" Then
         If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" And m_bolFMP = False Then
            'Modified by Lydia 2022/08/16 +申請國家
            'PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), "1603", strReceiveNo
            PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), "1603", pa(9), strReceiveNo
         End If
                  
         'Added by Lydia 2016/09/26 一案兩請案件發明案發證時,若新型案的下次繳費日期大於發明案的公告日加一年請發E-MAIL通知智權同仁,可辦理退費。
         'Modified by Lydia 2017/01/10 大陸案不可退費 + And pa(9) = "000"
         If m_bolDualApply And pa(9) = "000" Then
            Dim m_PYfee As Long
            Dim m_iYearTo As String '最後繳費年度
            Dim m_iYearFrom As String 'Added by Lydia 2017/11/27 退費起始年度
            Dim m_iYearList As String '已繳費年度
            Dim m_DivNa01 As String '新型案國別
            Dim m_strDate As String '正確起算日期/專用起始日
            Dim m_Div605NP09 As String '新型案年費期限=下次繳費期限
            
            'Modified by Lydia 2016/12/02 +pa26,pa14
            'Modified by Morgan 2017/10/6 要排除新型沒有繳費記錄者否則會錯 Ex:P111488
            strSql = "select pa09,pa08,pa72,pa26,pa14 from patent where pa01='" & m_DualCaseNo(1) & "' and  pa02='" & m_DualCaseNo(2) & "' and pa03='" & m_DualCaseNo(3) & "' and  pa04='" & m_DualCaseNo(4) & "' and pa72 is not null"
            intI = 1
            'Modified by Lydia 2016/12/02 rsTemp -> rsRead
            Set rsRead = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
                '-----------------
                m_DivNa01 = "" & rsRead.Fields("pa09")
                m_iYearList = Trim("" & rsRead.Fields("pa72"))
                strExc(1) = Val("" & rsRead.Fields("pa08")) + 10
                'Move by Lydia 2016/12/02 從GetMoneyDate下面移來

                'Memo by Lydia 2016/12/02 取得專用起始日
                If GetMoneyDate(Val(strExc(1)), m_DivNa01, m_DualCaseNo, m_strDate, m_iYearTo) = True Then
                    If InStr(m_iYearList, ",") = 0 And m_iYearList <> "" Then
                       m_iYearTo = m_iYearList
                       m_iYearFrom = m_iYearList 'Added by Lydia 2017/11/27
                    Else
                       If Right(m_iYearList, 1) = "," Then m_iYearList = Mid(m_iYearList, 1, Len(m_iYearList) - 1)
                       m_iYearTo = Mid(m_iYearList, InStrRev(m_iYearList, ",") + 1)
                       'Remove by Lydia 2018/01/18 改成新型案下次繳費日期和發明案的公告日,判斷要退X年
                       ''Added by Lydia 2017/11/27 比較新型案和發明案的公告日,判斷退費起始年度
                      '                                          ' 若公告日為同一年，則退費從第2年起算到最後繳費年度。Ex.106/11/21 P-115688(發明案)已公告，P-115687(新型案)尚有預繳年費(第2-3年，金額:3,400)可辦理退費
                      ' strExc(1) = Val(PUB_DBYEAR(m_strDate)) - Val(PUB_DBYEAR(DBDATE(txtPA14)))
                      ' If Val(strExc(1)) <= 0 Then '若公告日為同一年,退費從第2年起算
                      '    m_iYearFrom = Val(Abs(strExc(1))) + 2
                      ' Else
                      '    m_iYearFrom = "1"
                      ' End If
                       'end 2017/11/27
                       'end 2018/01/18
                    End If
                    'Modified by Lydia 2018/09/14 新型案公告日(P-115682)與發明公告日(P-115681)同一天不能退第2年年費；因為退費期限要算到繳費期限前一天(專利有效期)，只能退第3年。
                    '--------參考收文號CA7026024進度備註：新型第2年年費退費不予退還事，新型專利權依專利法第32條第2項規定，自發明公告之日滅，故新型專利權於發明公告之日須維持存續狀態
                    'm_Div605NP09 = CompDate(0, m_iYearTo, m_strDate) '新型案年費期限=下次繳費期限
                    m_Div605NP09 = CompDate(0, m_iYearTo, CompDate(2, -1, m_strDate)) 'Memo by Morgan 2022/5/18 此處規則要與 frm04010306_1 退費申請書同步修改
                    'end 2018/09/14
                    
                    'Added by Lydia 2018/01/18 新型案下次繳費日期和發明案的公告日,判斷要退X年 ; ex. 107/1/11 P-116270(發明案)已公告,P-116271(新型案)尚有預繳年費(第2-3年，金額:3,400)可辦理退費。
                    m_PYfee = (m_Div605NP09 - DBDATE(txtPA14)) \ 10000 '退X年
                    m_iYearFrom = m_iYearTo - m_PYfee + 1  '求退費起始年
                    'end 2018/01/18
                    
                    'Remove by Lydia 2016/12/02
                    'm_PYfee = PUB_GetYF07(m_DivNa01, "2", ChangeCustomerL(pa(26)), 年費, m_iYearTo, m_iYearTo, "1")
                End If

                'Added by Lydia 2016/12/02 台灣案若有減免,需要扣除減免額
                If Val(m_PYfee) > 0 Then 'Added by Lydia 2018/01/18 判斷有退X年才計算
                    If m_DivNa01 = "000" Then
                        strExc(1) = PUB_GetCP81(m_DualCaseNo, strExc(2)) '是否減免/發文日
                        'Modified by Lydia 2017/11/27 計算年費不只一年
                        'PUB_GetPatentYearFee m_DivNa01, rsRead.Fields("pa08"), "Y00000001", 年費, m_iYearTo, m_iYearTo, False, strExc(1), "" & rsRead.Fields("pa14"), strExc(2), strExc(3)
                        PUB_GetPatentYearFee m_DivNa01, rsRead.Fields("pa08"), "Y00000001", 年費, m_iYearFrom, m_iYearTo, False, strExc(1), "" & rsRead.Fields("pa14"), strExc(2), strExc(3)
                        m_PYfee = Val(strExc(3))
                    Else
                       '大陸案
                        'Modified by Lydia 2017/11/27 計算年費不只一年
                        'm_PYfee = PUB_GetYF07(m_DivNa01, rsRead.Fields("pa08"), rsRead.Fields("pa26"), 年費, m_iYearTo, m_iYearTo, "1")
                        m_PYfee = PUB_GetYF07(m_DivNa01, rsRead.Fields("pa08"), rsRead.Fields("pa26"), 年費, m_iYearFrom, m_iYearTo, "1")
                    End If
                    'end 2016/12/02
                End If 'end 2018/01/18
            'end 2016/12/02 rsTemp -> rsRead
            

                Set rsRead = Nothing
                'end 2016/12/02
                
                'Memo by Lydia 2016/09/29 Morgan與玲玲確認過新型案和發明案的公告日同一天,仍可辦理退費
                'Modified by Lydia 2018/09/13 有退費才通知
                'If Val(m_Div605NP09) >= Val(CompDate(0, 1, DBDATE(txtPA14))) And Val(m_PYfee) > 0 Then
                If Val(m_PYfee) > 0 Then
                   'Modified by Lydia 2017/11/27 m_iYearTo=> IIf(m_iYearFrom <> m_iYearTo, m_iYearFrom & "-" & m_iYearTo, m_iYearTo)
                   strExc(2) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", pa(3) & "-" & pa(4)) & "(發明案)已公告," & _
                               m_DualCaseNo(1) & "-" & m_DualCaseNo(2) & IIf(m_DualCaseNo(3) & m_DualCaseNo(4) = "000", "", m_DualCaseNo(3) & "-" & m_DualCaseNo(4)) & "(新型案)" & _
                               "尚有預繳年費(第" & IIf(m_iYearFrom <> m_iYearTo, m_iYearFrom & "-" & m_iYearTo, m_iYearTo) & "年，金額:" & Format(m_PYfee, DDollar2) & ")可辦理退費"
                   'modify by sonia 2022/4/8
                   'strExc(3) = "若欲辦理請收文並取回官方繳費收據,若收據無法取回請註明原因。"
                   strExc(3) = "若欲辦理請收文；如官方繳費收據為紙本收據請向客戶取回憑辦，若收據無法取回請註明原因。"
                   'Modified by Morgan 2021/1/28
                   'strExc(4) = PUB_GetAKindSalesNo(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text)
                   strExc(4) = stCP13
                   'end 2021/1/28
                   PUB_SendMail strUserNum, strExc(4), "", strExc(2), strExc(3)
                End If
            End If
         End If
         'end 2016/09/26
         
         'Add By Sindy 2016/10/5
         If Me.m_strIR01 <> "" Then
            Unload frm04010505_1
            Unload Me
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         
         'Added by Morgan 2022/12/19
         ElseIf Me.m_DocNo <> "" Then
            Unload frm04010505_1
            Unload Me
            frm04010516.GoNext
         'end 2022/12/19
         Else
         '2016/10/5 END
            frm04010505_1.Show
            frm04010505_1.Clear
            Unload Me
         End If
       Case 1
         frm04010505_1.Show
         Unload Me
       Case 2
         Unload Me
         Unload frm04010505_1
End Select
End Sub

Private Sub Form_Initialize()
    'add by nickc 2007/02/02
    ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()

    
    Me.Height = 6285
    Me.Width = 8850
    MoveFormToCenter Me
    m_strNextDueDate = Empty
    m_strNextFeeDate = Empty
    intWhere = 國內
    'Add By Cheng 2003/04/02
    m_blnFormFirstShow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2021/11/10
   Set frm04010505_2 = Nothing
End Sub

Private Function FormSave() As Boolean
Dim autonum As String
Dim strTxt(1 To 5) As String
Dim intStep As Integer
Dim strProgressNo As String
Dim strTmp As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strBillNo As String '帳單編號
Dim stPA14 As String
'Add by Morgan 2005/6/14
Dim stPA72 As String, stPA73 As String, iPA72 As Integer, stPA74 As String
'add by nickc 2006/05/05
Dim StrlMAXbyNick As String
Dim VarlMaxByNick As Variant
Dim jjjbyNick As Integer
Dim strBillCP09 As String 'Add by Morgan 2007/6/11 帳單收文號
Dim stCP64 As String     '2008/12/12 ADD BY SONIA
Dim strPromoteDate As String '2010/1/19 add by sonia
Dim strErrMsg As String 'Added by Morgan 2016/6/30
Dim stNP23 As String 'Added by Lydia 2025/10/29

On Error GoTo ErrorHandler

   FormSave = True
   cnnConnection.BeginTrans
   
   'add by nickc 2006/05/05
   StrlMAXbyNick = ""
'   FormSave = False
   autonum = AutoNo("C", 6)
   '92.12.31 MODIFY BY SONIA
   'strTxt(1) = "update patent set PA17='Y',pa22='" & Text7.Text & "'," & _
   '   "pa21=" & TransDate(Text5.Text, 2) & ",pa24=" & TransDate(Text6.Text, 2) & "," & _
   '   "pa25=" & TransDate(Text8.Text, 2) & " " & _
   '   "Where PA01='" & frm04010505_1.NUMBER1 & "' AND PA02='" & frm04010505_1.NUMBER2 & "' AND " & _
   '   "PA03='" & frm04010505_1.NUMBER3 & "' AND PA04='" & frm04010505_1.NUMBER4 & "'"
   
   '香港案同時更新核准及准駁通知日為發證日
   '2008/8/29 modify by sonia 香港所有專利種類都是自動發證
   'If pa(9) = "013" And pa(8) = "1" Then
   If pa(9) = "013" Then
      'Add by Morgan 2005/6/14
      'Modify by Morgan 2010/3/17 +PA74
      If m_iYear > 0 Then
         'Modified by Morgan 2012/10/23
         '需考慮有繳維持費
         If pa(72) = "" Then
            stPA72 = "1"
            intI = 1
            stPA73 = Text5
            stPA74 = ""
         Else
            stPA72 = pa(72)
            stPA73 = pa(73)
            stPA74 = pa(74)
            If Left(Right(stPA72, 2), 1) <> "," Then
               intI = Right(stPA72, 2)
            Else
               intI = Right(stPA72, 1)
            End If
         End If
         'Added by Morgan 2014/7/15
         '若已繳年度超過時要刪除多餘的(維持費期限前發證會有此問題 Ex.P-101559)
         If intI > m_iYear Then
            For iPA72 = m_iYear + 1 To intI
               stPA72 = Left(stPA72, InStrRev(stPA72, ",") - 1)
               stPA73 = Left(stPA73, InStrRev(stPA73, ",") - 1)
               stPA74 = Left(stPA74, InStrRev(stPA74, ",") - 1)
            Next
         Else
         'end 2014/7/15
            intI = intI + 1
            For iPA72 = intI To m_iYear
               stPA72 = stPA72 & "," & iPA72
               stPA73 = stPA73 & "," & Text5
               stPA74 = stPA74 & ","
            Next
         End If 'Added by Morgan 2014/7/15
      End If
      'Modify by Morgan 2005/6/14 加更新pa72,pa73
      strTxt(1) = "update patent set PA16='1',PA17='Y',pa22='" & Text7.Text & "'," & _
         "pa20=" & TransDate(Text5.Text, 2) & ",pa21=" & TransDate(Text5.Text, 2) & ",pa24=" & TransDate(Text6.Text, 2) & "," & _
         "pa25=" & TransDate(Text8.Text, 2) & _
         ",pa72=" & CNULL(stPA72) & ",pa73=" & CNULL(stPA73) & ",pa74=" & CNULL(stPA74) & _
         " Where PA01='" & frm04010505_1.NUMBER1 & "' AND PA02='" & frm04010505_1.NUMBER2 & "' AND " & _
         "PA03='" & frm04010505_1.NUMBER3 & "' AND PA04='" & frm04010505_1.NUMBER4 & "'"
      'end 2010/3/17
   'Add by Morgan 2007/3/27 澳門也是主動發證--核准及准駁通知日為發證日
   ElseIf pa(9) = "044" Then
      strTxt(1) = "update patent set PA16='1',PA17='Y',pa22='" & Text7.Text & "'," & _
         "pa20=" & TransDate(Text5.Text, 2) & ",pa21=" & TransDate(Text5.Text, 2) & ",pa24=" & TransDate(Text6.Text, 2) & "," & _
         "pa25=" & TransDate(Text8.Text, 2) & _
         " Where PA01='" & frm04010505_1.NUMBER1 & "' AND PA02='" & frm04010505_1.NUMBER2 & "' AND " & _
         "PA03='" & frm04010505_1.NUMBER3 & "' AND PA04='" & frm04010505_1.NUMBER4 & "'"
   Else
      'Modify by Morgan 2004/7/2
      '台灣無公告日或公告日>9307012的同時更新公告日
      If m_bolNew = True And txtPA14 <> "" Then
         stPA14 = DBDATE(txtPA14)
         pa(14) = stPA14
      Else
         stPA14 = "PA14"
      End If
      
      'Add by Morgan 2006/1/16
      If txtPA15.Visible = True Then
         pa(15) = txtPA15.Text
         If pa(15) = "CN" Then pa(15) = "" 'Add by Morgan 2006/4/10
         '95.1.1以後大陸公告日=發證日
         If TransDate(Text5.Text, 2) > "20060101" Then
            stPA14 = DBDATE(Text5)
            pa(14) = stPA14
         End If
      End If
      
      'Modify by Morgan 2006/1/16 加pa15
      '2007/4/24 MODIFY BY SONIA
'      strTxt(1) = "update patent set PA17='Y',pa22='" & Text7.Text & "'," & _
'         " pa21=" & TransDate(Text5.Text, 2) & ",pa24=" & TransDate(Text6.Text, 2) & "," & _
'         " pa25=" & TransDate(Text8.Text, 2) & ",pa14=" & stPA14 & ",pa15=" & CNULL(ChgSQL(pa(15))) & _
'         " Where PA01='" & frm04010505_1.NUMBER1 & "' AND PA02='" & frm04010505_1.NUMBER2 & "' AND " & _
'         " PA03='" & frm04010505_1.NUMBER3 & "' AND PA04='" & frm04010505_1.NUMBER4 & "'"
      If m_bol117Exist = True Then
         strTxt(1) = "update patent set PA17='Y',pa22='" & Text7.Text & "'," & _
            " pa21=" & TransDate(Text5.Text, 2) & ",pa24=" & TransDate(Text6.Text, 2) & "," & _
            " pa25=" & TransDate(Text8.Text, 2) & ",pa16='1',pa20=" & TransDate(Text5.Text, 2) & _
            " Where PA01='" & frm04010505_1.NUMBER1 & "' AND PA02='" & frm04010505_1.NUMBER2 & "' AND " & _
            " PA03='" & frm04010505_1.NUMBER3 & "' AND PA04='" & frm04010505_1.NUMBER4 & "'"
      Else
         strTxt(1) = "update patent set PA17='Y',pa22='" & Text7.Text & "'," & _
            " pa21=" & TransDate(Text5.Text, 2) & ",pa24=" & TransDate(Text6.Text, 2) & "," & _
            " pa25=" & TransDate(Text8.Text, 2) & ",pa14=" & stPA14 & ",pa15=" & CNULL(ChgSQL(pa(15))) & _
            " Where PA01='" & frm04010505_1.NUMBER1 & "' AND PA02='" & frm04010505_1.NUMBER2 & "' AND " & _
            " PA03='" & frm04010505_1.NUMBER3 & "' AND PA04='" & frm04010505_1.NUMBER4 & "'"
      End If
      '2007/4/24 END
   End If
   '92.12.31 END
    cnnConnection.Execute strTxt(1)
   
   'Added by Morgan 2021/6/29 大陸發明生化生醫案設定是否新藥專利
   If m_PA176 <> pa(176) Then
      strSql = "update patent set pa176='" & m_PA176 & "' where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
      cnnConnection.Execute strSql, intI
   End If
   'end 2021/6/29
   
   '2008/12/12 ADD BY SONIA
   If m_bln605 = False Then
      stCP64 = "閉卷後發證不取消閉卷也不管制年費期限"
   Else
      stCP64 = ""
   End If
   '2008/12/12 END
   
   'Added by Morgan 2021/6/10
   If bolAddNP445 Then
      'Modified by Morgan 2023/2/16
      'stCP64 = "專利權期限補償服務費:6000;" & stCP64
      stCP64 = "專利權期限補償報價:6000(3p);" & stCP64
      'end 2023/2/16
   End If
   'end 2021/6/10
      
   If Text10 <> "" Then
      'Modified by Morgan 2012/4/30 +cp119=櫃檯收文日
      strTxt(2) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13," & _
         "CP14,CP16,CP17,CP18,CP26,CP27,CP64,CP119) VALUES ('" & Text1.Text & "','" & Text2.Text & "','" & _
         Text3.Text & "','" & Text4.Text & "','" & TransDate(Label18.Caption, 2) & "','" & _
         autonum & "','1603','" & stCP12 & "','" & stCP13 & "','" & strUserNum & "', " & Text10 & ", '0'," & Text10 / 1000 & ",'N','" & _
         strSrvDate(1) & "'," & CNULL(ChgSQL(stCP64)) & "," & DBDATE(Label18) & ")"
   Else
      'Modified by Morgan 2012/4/30 +cp119=櫃檯收文日
      strTxt(2) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13," & _
         "CP14,CP20,CP26,CP27,CP32,CP64,CP119) VALUES ('" & Text1.Text & "','" & Text2.Text & "','" & _
         Text3.Text & "','" & Text4.Text & "','" & TransDate(Label18.Caption, 2) & "','" & _
         autonum & "','1603','" & stCP12 & "','" & stCP13 & "','" & strUserNum & "','N','N','" & _
         strSrvDate(1) & "','N'," & CNULL(ChgSQL(stCP64)) & "," & DBDATE(Label18) & ")"
   End If
'Modify end 2004/2/9
   
   '92.12.31 end
    'Add By Cheng 2002/11/08
    cnnConnection.Execute strTxt(2)
   
   'Add by Morgan 2004/11/30 抓最新的AB類發文代理人更新
   Pub_UpdateFromMaxCP27 Text1.Text, Text2.Text, Text3.Text, Text4.Text
      
   strReceiveNo = autonum
   intStep = 3
   
   'Added by Morgan 2014/4/14 電子化-新增信函進度檔
   If pa(9) = "000" Then
      strExc(1) = ""
      If Text15 <> "N" Then
         'Modified by Morgan 2018/8/1
         'strExc(1) = PUB_GetLetterJudge(pa(1), "1603", , , pa(1), pa(2), pa(3), pa(4))
         strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), "1603")
      End If
      'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
      PUB_AddLetterProgress strReceiveNo, 1, IIf(Text15 <> "N", True, False), strExc(1), False, pa(26), "1603", pa(75)
   
   'Added by Morgan 2016/6/14 非臺灣案電子化
   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      strExc(1) = ""
      If Text15 <> "N" Then
         'Modified by Morgan 2018/8/1
         'strExc(1) = PUB_GetLetterJudge(pa(1), "1603", , pa(9), pa(1), pa(2), pa(3), pa(4))
         strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), "1603", pa(9), , , m_bolFMP)
      End If
      'Modified by Morgan 2016/7/6 證書不必存altr附件數量改1
      PUB_AddLetterProgress strReceiveNo, 1, IIf(Text15 <> "N", True, False), strExc(1), False, pa(26), "1603", pa(75)
   'end 2016/6/14
   
   End If
   'end 2014/4/14
   
   '2008/12/12 MODIFY BY SONIA 因大陸案閉卷不取消閉卷故多做控管
   'If IsEmptyText(Text9) = False Then
   If IsEmptyText(Text9) = False And m_bln605 = True Then
   '2008/12/12 END
            
      'Modify by Morgan 2010/1/8 原程式看來沒用到,FMP年費所限=法限-10天
      'strExc(1) = Text1.Text
      'strExc(2) = pa(9)
      'strExc(3) = TransDate(Text9.Text, 2)
      'GetCtrlDT strExc
      'Modified by Morgan 2018/10/3 非FMP的非台灣案也改10天
      'If m_bolFMP Then m_strNextFeeDate = PUB_GetWorkDay1(CompDate(2, -10, m_strNextDueDate), True)
      'Added by Lydia 2025/10/29
      stNP23 = ""
      If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
         m_strNextFeeDate = PUB_GetPOurDeadline(m_strNextDueDate, pa(9), stNP23, pa(1), "605")
      Else
      'end 2025/10/29
         If pa(9) <> "000" Then m_strNextFeeDate = PUB_GetWorkDay1(CompDate(2, -10, m_strNextDueDate), True)
         'end 2018/10/3
         'end 2010/1/8
      'Added by Lydia 2025/10/29
         stNP23 = m_strNextFeeDate
      End If
      'end 2025/10/29
      
      '有下一程序
      If m_str605NP22 <> Empty Then
         'Add by Morgan 2004/7/12
         '若用新法則更新下一程序年費期限
         'Modify by Morgan 2006/6/1 改都要更新
         'If m_bolNew = True Then
            'Modified by Morgan 2012/2/14 +NP23 也要更新
            'Modiiied by Lydia 2025/10/29 NP23=" & m_strNextFeeDate=>NP23=" & stNP23
            strSql = "update nextprogress set NP08=" & m_strNextFeeDate & ", np09=" & m_strNextDueDate & ",NP23=" & stNP23 & _
               " where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
               " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
            cnnConnection.Execute strSql, intI
         If m_bolNew = True Then
            'Add by Morgan 2004/7/15
            '若有未發文技術報告時更新文件齊備日及承辦期限
            If PUB_ChkCPExist(pa, "421", 1, m_str421CP09) = True Then
               '更新文件齊備日
               strSql = "Update EngineerProgress Set EP06=" & strSrvDate(1) & " Where EP02='" & m_str421CP09 & "' AND EP06 IS NULL"
               cnnConnection.Execute strSql
               
               If m_bolFMP Or PUB_IfSetCP48() Then   'Add by Morgan 2010/9/29 新規則承辦期限隔日凌晨算
                  m_str421EP06 = strSrvDate(1)
                  'Modify by Morgan 2007/10/12 承辦期限改呼叫共用函數計算
                  'm_str421CP48 = PUB_GetEngDueDate(m_str421EP06, pa(1), "000", "421")
                  m_str421CP48 = Pub_GetHandleDay(pa(1), "000", "421", m_str421EP06, , m_str421CP09)
                  'end 2007/10/12
                  If Val(m_str421CP48) > 0 Then
                     '更新承辦期限
                     strSql = "Update CaseProgress Set CP48=" & m_str421CP48 & " Where CP09='" & m_str421CP09 & "' AND CP48 IS NULL"
                     cnnConnection.Execute strSql
                  End If
               End If 'Add by Morgan 2010/9/29
            End If
         End If
         
      'Added by Morgan 2012/8/7
      '已收文年費時更新期限
      ElseIf PUB_ChkCPExist(pa, "605", 1, strExc(1)) = True Then
         strSql = "update caseprogress set CP06=" & m_strNextFeeDate & ", CP07=" & m_strNextDueDate & _
            " where cp09='" & strExc(1) & "'"
         cnnConnection.Execute strSql, intI
      'end 2012/8/7
      
      Else
        strProgressNo = GetNextProgressNo
        'Modified by Lydia 2025/10/29 +NP23
        strTxt(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09," & _
            "NP10,NP22,NP23) VALUES ('" & autonum & "','" & Text1.Text & "','" & Text2.Text & _
            "','" & Text3.Text & "','" & Text4.Text & "',605," & PUB_GetWorkDay1(m_strNextFeeDate, True) & _
            "," & m_strNextDueDate & ",'" & stCP13 & "','" & strProgressNo & "'," & CNULL(stNP23, True) & ")"
        cnnConnection.Execute strTxt(3)
        m_str605NP22 = strProgressNo 'Add by Morgan 2006/6/1
        intStep = 4
        'Add by Morgan 2004/10/1 若為中間的案子會沒有核准且無下一程序
        If m_bolNew = True Then
            '若有未發文技術報告時更新文件齊備日及承辦期限
            If PUB_ChkCPExist(pa, "421", 1, m_str421CP09) = True Then
               '更新文件齊備日
               strSql = "Update EngineerProgress Set EP06=" & strSrvDate(1) & " Where EP02='" & m_str421CP09 & "' AND EP06 IS NULL"
               cnnConnection.Execute strSql
               
               If m_bolFMP Or PUB_IfSetCP48() Then  'Add by Morgan 2010/9/29 新規則承辦期限隔日凌晨算
                  m_str421EP06 = strSrvDate(1)
                  'Modify by Morgan 2007/10/12 承辦期限改呼叫共用函數計算
                  'm_str421CP48 = PUB_GetEngDueDate(m_str421EP06, pa(1), "000", "421")
                  m_str421CP48 = Pub_GetHandleDay(pa(1), "000", "421", m_str421EP06, , m_str421CP09)
                  'end 2007/10/12
                  
                  If Val(m_str421CP48) > 0 Then
                     '更新承辦期限
                     strSql = "Update CaseProgress Set CP48=" & m_str421CP48 & " Where CP09='" & m_str421CP09 & "' AND CP48 IS NULL"
                     cnnConnection.Execute strSql
                  End If
               End If 'Add by Morgan 2010/9/29
            End If
        End If
        '2004/10/1 end
        
      End If
      
   End If
   
'   FormSave = objLawDll.ExecSQL(intStep - 1, strTxt)
   
'    'Add By Cheng 2002/10/31
'    If FormSave = True Then
        'Add By Cheng 2002/07/26
        '更新相同案號的案件進度檔案件性質為"領證及繳年費"(601)且"發文日"有值者, 其實際結果欄為"1"(准)
        strSql = "Update CASEPROGRESS SET CP24='1'," & _
           "CP25=" & TransDate(Text5.Text, 2) & " WHERE " & ChgCaseprogress(Text1 & Text2 & Text3 & Text4) & " AND CP10='601' AND CP27 IS NOT NULL "
        cnnConnection.Execute strSql
        
        'Added by Morgan 2020/3/6
        '解除領證收達提申管制
        If pa(9) <> "000" Then
            strSql = "UPDATE nextprogress SET NP06='Y' where np01=(select cp09 from caseprogress" & _
               " where " & ChgCaseprogress(Text1 & Text2 & Text3 & Text4) & " AND CP10='601' and cp27>0)" & _
               " and np07 in ('995','996','997','998') and np06 is null"
            cnnConnection.Execute strSql, intI
        End If
        'end 2020/3/6
        
        'Add By Cheng 2002/10/31
        '若有輸入代理人D/N No, 帳單日期 及 帳單金額, 則新增國外帳單資料
        If Me.Text11.Text <> "" And Me.Text12.Text <> "" And Me.Text13.Text <> "" Then
            StrSQLa = "Select CP09 From CaseProgress Where " & ChgCaseprogress(Me.Text1.Text & Me.Text2.Text & Me.Text3.Text & Me.Text4.Text) & " And  CP10='601' AND CP27 IS NOT NULL "
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               strBillCP09 = "" & rsA.Fields("CP09").Value
            'Add by Morgan 2007/6/11 沒有領證時用
            Else
               strBillCP09 = autonum
            End If
            If strBillCP09 <> "" Then
            'end 2007/6/11
                If PUB_AddNewFBillData(strBillCP09, Me.Text11.Text, Me.Text12.Text, Me.Text13.Text, strBillNo) = False Then
                    'Modified by Morgan 2016/6/30 錯誤訊息不可放在 Transaction 內
                    'MsgBox "新增國外帳單資料作業失敗!!!", vbExclamation + vbOKOnly
                    strErrMsg = "新增國外帳單資料作業失敗!!!"
                    'end 2016/6/30
                    GoTo ErrorHandler
                Else
                  'Added by Morgan 2016/6/30 非臺灣案電子化
                  'Removed by Morgan 2025/8/13 帳單已全部都電子化
                  'If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
                  'end 2025/8/13
                     '檢查帳單是否存在
                     If PUB_CheckInvoicePDF(pa(1), pa(2), pa(3), pa(4), strBillCP09, strErrMsg, , True, strBillNo) = False Then
                        GoTo ErrorHandler
                     End If
                  'End If
                  'end 2016/6/30
                    'Add By Cheng 2002/11/05
                    frm04010505_1.lblBillNo.Caption = "" & strBillNo
                End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
'    End If
    
   
   'add by nickc 2005/06/16 大陸香港關聯案
   If pa(9) = "020" Then
   Dim tmpCp06 As String
   Dim tmpCp07 As String
   '檢查有無香港
      If m_HaveHK = True Then
            'Modified by Morgan 2020/4/16 法限不必抓工作天(另目前大陸公告日=發證日)--郭
            'tmpCp07 = PUB_GetWorkDay1(CompDate(1, 6, Text5.Text), True)
            tmpCp07 = CompDate(1, 6, Text5.Text)
            'end 2020/4/16
            'Added by Lydia 2025/10/29
            stNP23 = ""
            If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
               tmpCp06 = PUB_GetPOurDeadline(tmpCp07, pa(9), stNP23, pa(1), "111")
            Else
            'end 2025/10/29
               tmpCp06 = PUB_GetWorkDay1(CompDate(2, -5, CompDate(1, -1, tmpCp07)), True)
            'Added by Lydia 2025/10/29
               stNP23 = tmpCp06
            End If
            'end 2025/10/29
            
            m_HK_NP08 = tmpCp06 'Add by Morgan 2008/6/30
            m_HK_NP09 = tmpCp07 'Added by Morgan 2012/3/2
          '檢查有無收香港的 111
          If m_HaveHKInCP <> "" Then
               '更新期限，上發 mail tag
               strSql = "Update CaseProgress Set CP06=" & tmpCp06 & ",CP07=" & tmpCp07 & " Where CP09='" & m_HaveHKInCP & "' "
               cnnConnection.Execute strSql
               '更新齊備日
               strSql = "update engineerprogress set ep06=" & ServerDate & " where ep02='" & m_HaveHKInCP & "' "
               cnnConnection.Execute strSql
               m_SendHKMail = True
               
               If PUB_IfSetCP48(m_HaveHKInCP) Then  'Add by Morgan 2010/9/29 新規則承辦期限隔日凌晨算
                  '2010/1/19 add by sonia 更新承辦期限
                  strPromoteDate = Pub_GetHandleDay(pa(1), "013", "111", , tmpCp06)
                  If strPromoteDate <> "" Then
                     strSql = "Update CaseProgress Set CP48=" & CNULL(strPromoteDate) & " Where CP09='" & m_HaveHKInCP & "' "
                     cnnConnection.Execute strSql
                  End If
                  '2010/1/19 end
               End If 'Add by Morgan 2010/9/29
          Else
               '檢查有無 np 的香港 111
               If m_HaveHKInNP <> "" Then
                  '更新期限
                  'Modify by Morgan 2008/6/30
                  'strSQL = "Update nextProgress Set nP08=" & tmpCp06 & ",nP09=" & tmpCp07 & " Where np01='" & m_HaveHKInNP & "' "
                  'Modified by Morgan 2013/12/31 核准時有預估期限但並未通知故此時約定期限(=本所期限)也要更新--David
                  'Modified by Lydia 2025/10/29 np23=" & tmpCp06=> np23=" & stNP23
                  strSql = "Update nextProgress Set nP08=" & tmpCp06 & ",nP09=" & tmpCp07 & ",np23=" & stNP23 & " Where np01='" & m_HaveHKInNP & "' and NP07='111'"
                  cnnConnection.Execute strSql
               Else
                  '新增 np 期限
                  strProgressNo = GetNextProgressNo
                  'Modified by Lydia 2025/10/29 +NP23
                  strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09," & _
                      "NP10,NP22,NP23) select '" & strReceiveNo & "','" & m_HK_CP01 & "','" & m_HK_CP02 & "','" & m_HK_CP03 & "','" & m_HK_CP04 & "',111," & tmpCp06 & "," & tmpCp07 & ",'" & _
                      PUB_GetAKindSalesNo(m_HK_CP01, m_HK_CP02, m_HK_CP03, m_HK_CP04) & "','" & strProgressNo & "'," & CNULL(stNP23, True) & _
                      " from dual "
                  StrlMAXbyNick = StrlMAXbyNick & strProgressNo & ","
                  'add by nickc 2006/05/09
                  'm_strNextFeeDate = tmpCp06 'Removed by Morgan 2013/7/15香港案期限已改用m_HK_NP08,此程式會造成大陸案的定稿期限錯誤
                  cnnConnection.Execute strSql
               End If
               
               'Added by Morgan 2016/6/15
               '非臺灣案電子化有香港案通知函要新增通知期限
               'Modified by Lydia 2022/09/30 +增加寰華案
               'If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
               If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And (Left(Pub_StrUserSt03, 1) <> "F" Or m_bolFMP2 = True) Then
                  strExc(1) = "": strExc(2) = ""
                  'Modified by Lydia 2022/09/30 +np23
                  strExc(0) = "select np01,np08,np09,np22,np23,pa09,pa26,pa75 from nextprogress,patent" & _
                     " where np02='" & m_HK_CP01 & "' and np03='" & m_HK_CP02 & "' and np04='" & m_HK_CP03 & "' and np05='" & m_HK_CP04 & "'" & _
                     " and np07=111 and np06 is null and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     With RsTemp
                     'Added by Lydia 2022/09/30 111香港標準專利批准記錄請求
                     strExc(1) = "" & .Fields("np09")  '法定期限
                     strExc(2) = "" & .Fields("np23")  '約定期限
                     'end 2022/09/30
                     If PUB_AddCP1913(m_HK_CP01, m_HK_CP02, m_HK_CP03, m_HK_CP04, .Fields("np08"), .Fields("np09"), .Fields("np01"), .Fields("NP22"), .Fields("pa09"), .Fields("pa26"), m_HK1913CP09, "" & .Fields("pa75"), , True) = False Then
                        Err.Raise 999, , "新增進度檔【通知期限】失敗！作業中斷！"
                     End If
                     End With
                  End If
                  
                  'Added by Lydia 2022/09/30 寰華案同時發Email：【通知：標準專利批准記錄請求】請報告代理人
                  If m_bolFMP2 = True Then
                      strExc(0) = "select cp13,s1.st02 as cp13n, cp14, s2.st02 as cp14n from caseprogress,staff s1, staff s2 where cp01='" & m_HK_CP01 & "' and cp02='" & m_HK_CP02 & "' and cp03='" & m_HK_CP03 & "' and cp04='" & m_HK_CP04 & "' " & _
                                       "and cp10='1913' and cp13=s1.st01(+) and cp14=s2.st01(+)"
                      intI = 1
                      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                      If intI = 1 Then
                           If "" & RsTemp.Fields("cp13") <> "" Then
                               'CC
                               strExc(4) = PUB_GetFCPProSup("" & RsTemp.Fields("cp13")) '智權人員主管
                               strExc(4) = strExc(4) & ";" & strUserNum & ";backup"
                               '主旨
                               strExc(5) = "【通知：標準專利批准記錄請求】請報告代理人 Our Ref: " & m_HK_CP01 & "-" & m_HK_CP02 & IIf(m_HK_CP03 & m_HK_CP04 <> "000", "-" & m_HK_CP03 & m_HK_CP04, "") & _
                                                "[INCOM.1913]"
                               '內文
                               strExc(6) = "大陸案" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "已收專利證書，" & _
                                                "其延伸至香港案" & m_HK_CP01 & "-" & m_HK_CP02 & IIf(m_HK_CP03 & m_HK_CP04 <> "000", "-" & m_HK_CP03 & m_HK_CP04, "") & "請通知標準專利批准記錄請求（第2階段）之期限。" & vbCrLf
                               strExc(6) = strExc(6) & "約定期限：" & ChangeWStringToTDateString(strExc(2)) & "　　法定期限：" & ChangeWStringToTDateString(strExc(1)) & vbCrLf & vbCrLf
                               strExc(6) = strExc(6) & "To " & RsTemp.Fields("cp14n") & ":" & vbCrLf & "請產生通知函至卷宗區。" & vbCrLf & vbCrLf
                               strExc(6) = strExc(6) & "To " & RsTemp.Fields("cp13n") & ":" & vbCrLf & "請報告標準專利批准記錄請求之期限。" & vbCrLf & vbCrLf
                               
                               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                     " values( '" & strUserNum & "','" & RsTemp.Fields("cp13") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                                     ",'" & strExc(5) & "','" & strExc(6) & "','" & strExc(4) & "')"
                               cnnConnection.Execute strSql, intI
                           End If
                      End If
                  End If
                  'end 2022/09/30
               End If
               'end 2016/6/15
          End If
      End If
       
      'Added by Lydia 2015/09/09
      If pa(8) = "1" Then
        '大陸發明案之核准，若該案有澳門案且發明進度尚未發文時，則同時大陸案之公告日+3個月更新至澳門發明進度之法定期限，本所期限=法定期限－1個月又5天
          Call PUB_UpdCP07by020(pa, m_bolFMP, "5", , DBDATE(Text5))
         '大陸發明案時要更新下一程序999公開期限且NP06 IS NULL的資料，更新其NP06='Y'。
         strExc(0) = "select cp09 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' " & _
                  "and cp10 in (" & NewCasePtyList & ") and cp57 is null"
         strSql = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE np01=(" & strExc(0) & ") and np07='999' and np06 is null "
         cnnConnection.Execute strSql, intI
      End If
      'end 2015/09/09
      
      'Added by Lydia 2018/07/09 為配合外專新案命名作業，針對香港案和澳門案寫入相應大陸案之名稱
      '香港案
      'Modified by Lydia 2021/11/10 + And m_bolFMP = True
      If m_HaveHK = True And m_HK_CP01 <> "" And m_HK_CP02 <> "" And m_bolFMP = True Then
            strSql = "update patent set pa05=" & CNULL(pa(5)) & ", pa06=" & CNULL(pa(6)) & ", pa07=" & CNULL(pa(7)) & _
                        " where PA01='" & m_HK_CP01 & "' AND PA02='" & m_HK_CP02 & "' AND PA03='" & m_HK_CP03 & "' AND PA04='" & m_HK_CP04 & "' "
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
      End If
      '澳門
      'Modified by Lydia 2021/11/10 改成自訂變數
      'strExc(1) = "": strExc(2) = ""
      'tmpBol = ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), strExc(1), strExc(2), strExc(3), strExc(4), strExc(5), , "5")
      'If tmpBol = True And strExc(1) <> "" And strExc(2) <> "" Then
      '     strSql = "update patent set pa05=" & CNULL(pa(5)) & ", pa06=" & CNULL(pa(6)) & ", pa07=" & CNULL(pa(7)) & _
                       " where PA01='" & strExc(1) & "' AND PA02='" & strExc(2) & "' AND PA03='" & strExc(3) & "' AND PA04='" & strExc(4) & "' "
      If m_Have044 = True And m_CPto044(1) <> "" And m_CPto044(2) <> "" And m_bolFMP = True Then
           strSql = "update patent set pa05=" & CNULL(pa(5)) & ", pa06=" & CNULL(pa(6)) & ", pa07=" & CNULL(pa(7)) & _
                       " where PA01='" & m_CPto044(1) & "' AND PA02='" & m_CPto044(2) & "' AND PA03='" & m_CPto044(3) & "' AND PA04='" & m_CPto044(4) & "' "
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql, intI
           'Added by Lydia 2022/03/17 若衍生澳門案已閉卷，則此mail可不發，謝謝
           strExc(0) = "select pa01,pa02,pa03,pa04 from patent where PA01='" & m_CPto044(1) & "' AND PA02='" & m_CPto044(2) & "' AND PA03='" & m_CPto044(3) & "' AND PA04='" & m_CPto044(4) & "' and pa57||pa108 is null "
           intI = 1
           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
           If intI = 1 Then
           'end 2022/03/17
                If m_bolFMP2 = True Then  '寰華案
                    '收件者: 程序人員;
                    strExc(1) = PUB_GetFCPHandler(m_CPto044(1), m_CPto044(2), m_CPto044(3), m_CPto044(4))
                    '副本收受者:智權人員; backup
                    strExc(2) = PUB_GetFCPSalesNo(m_CPto044(1), m_CPto044(2), m_CPto044(3), m_CPto044(4))
                ElseIf m_bolFMP = True Then
                    '收件者: 內專人員(暫時預設98012品薇);
                    'Modified by Morgan 2025/1/21
                    'strExc(1) = "98012"
                    If strSrvDate(1) >= P業務區劃分啟用日 Then
                      strExc(1) = PUB_GetPHandler(m_CPto044(1) & m_CPto044(2) & m_CPto044(3) & m_CPto044(4))
                    Else
                      strExc(1) = Pub_GetSpecMan("PS2")
                    End If
                    'end 2025/1/21
                    '副本收受者: 智權人員; backup
                    strExc(2) = PUB_GetFCPSalesNo(m_CPto044(1), m_CPto044(2), m_CPto044(3), m_CPto044(4))
                End If
                '主旨
                strExc(4) = "大陸案" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "已授權公告，" & _
                                 "澳門案" & m_CPto044(1) & "-" & m_CPto044(2) & IIf(m_CPto044(3) & m_CPto044(4) <> "000", "-" & m_CPto044(3) & "-" & m_CPto044(4), "") & _
                                 "可提出申請 Our Ref: " & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "[INCOM. 1603]"
                '內文
                strExc(0) = "select sqldatet(cp07) cp07 from caseprogress where cp01='" & m_CPto044(1) & "' and cp02='" & m_CPto044(2) & "'  and cp03='" & m_CPto044(3) & "'  and cp04='" & m_CPto044(4) & "' and cp10='101' "
                strExc(7) = ""
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                    strExc(7) = "" & RsTemp.Fields("cp07")
                End If
                strExc(5) = "澳門案" & m_CPto044(1) & "-" & m_CPto044(2) & IIf(m_CPto044(3) & m_CPto044(4) <> "000", "-" & m_CPto044(3) & "-" & m_CPto044(4), "") & _
                                 "可進行提申作業，提申期限：" & strExc(7) & "，請儘速處理。"
                If strExc(1) <> "" Then
                    strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                               " values( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                               ",'" & strExc(4) & "','" & strExc(5) & "','" & strExc(2) & ";backup')"
                    cnnConnection.Execute strSql, intI
                End If
           End If 'Added 2022/03/17
      End If
      'end 2018/07/09
   End If
   
   'Unload Me
   'frm04010505_1.Show
   
   'Added by Morgan 2022/12/16
   If m_DocNo <> "" Then
      If m_DocWord <> "" Then
         strSql = "UPDATE CASEPROGRESS set cp08='" & m_DocWord & "字第" & m_DocNo & "號'" & _
            " WHERE CP09='" & strReceiveNo & "'"
         cnnConnection.Execute strSql, intI
      End If
      PUB_UpdateEdocRec m_DocNo, strReceiveNo, pa(1), pa(2), pa(3), pa(4), "1603"
   End If
   'end 2022/12/16
   
   'Add by Sindy 2016/10/5
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", strReceiveNo, "")
      'Modified by Lydia 2023/05/18 +不開啟附件, , , False
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010505_1", IIf(Pub_StrUserSt03 = "F22", strReceiveNo, ""), , , False
   End If
   '2016/10/5 END
   
   'Add by Morgan 2005/5/20
   '非台灣 更新結餘
   Pub_UpdateEndModCash pa(1), pa(2), pa(3), pa(4)
   
   'Add by Morgan 2006/9/11
   '大陸一案兩請，若申請日為95.7.1以前則發明案發證時，新型自動上閉卷
   'Memo by Morgan 2013/5/31 +台灣案
   If m_bolDualApply = True Then
      strSql = "UPDATE PATENT SET PA57='Y',PA58=" & strSrvDate(1) & ",PA59='99', PA91=PA91||';發明案(" & pa(1) & pa(2) & pa(3) & pa(4) & ")已發證系統自動上閉卷" & "' WHERE PA01='" & m_DualCaseNo(1) & "' AND PA02='" & m_DualCaseNo(2) & "' AND PA03='" & m_DualCaseNo(3) & "' AND PA04='" & m_DualCaseNo(4) & "' AND PA57 IS NULL"
      cnnConnection.Execute strSql, intI
      'Add by Morgan 2007/10/25 下一程序年費期限自動不續辦
      strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & strSrvDate(1) & ",NP12='99', NP15=NP15||';發明案(" & pa(1) & pa(2) & pa(3) & pa(4) & ")已發證系統自動不續辦" & "' WHERE NP02='" & m_DualCaseNo(1) & "' AND NP03='" & m_DualCaseNo(2) & "' AND NP04='" & m_DualCaseNo(3) & "' AND NP05='" & m_DualCaseNo(4) & "' AND NP06 IS NULL AND NP07='605' AND NP09>=" & strSrvDate(1)
      cnnConnection.Execute strSql, intI
   End If
   
   'Add by Morgan 2007/5/10
   If m_blnCancelClosed = True Then
      strSql = "UPDATE PATENT SET PA57=NULL,PA58=NULL,PA59=NULL" & _
         " WHERE PA01 = '" & pa(1) & "' AND PA02 = '" & pa(2) & "'" & _
         " AND PA03 = '" & pa(3) & "' AND PA04 = '" & pa(4) & "' "
      cnnConnection.Execute strSql, intI
   End If
   'end 2007/5/10
   
   '2008/5/16 ADD BY SONIA 更新相關收文號為'A'類最小的總收文號,因香港短期年費解除期限會抓不到CP44
   strSql = "UPDATE CASEPROGRESS A" & _
      " SET A.CP43=(SELECT MIN(B.CP09) FROM CASEPROGRESS B WHERE B.CP01=A.CP01 AND B.CP02=A.CP02 AND B.CP03=A.CP03 AND B.CP04=A.CP04 AND B.CP09<'B')" & _
      " WHERE A.CP09='" & autonum & "' AND A.CP43 IS NULL"
   cnnConnection.Execute strSql, intI
   '2008/5/16 END
   
   '2008/8/29 add by sonia 香港澳門自動發證更新進度檔新申請案核准及核准日為發證日
   If pa(9) = "013" Or pa(9) = "044" Then
      '2010/4/6 modify by sonia 加入香港111(P-086224)
      'StrSQLa = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND ( (CP10>='101' AND CP10<='105') OR (CP10>='301' AND CP10<='307') OR CP10='107' OR CP10='110' OR CP10='112' OR CP10='204' OR CP10='501' ) AND CP27 IS NOT NULL AND CP24 IS NULL AND CP25 IS NULL ORDER BY CP27 DESC "
      StrSQLa = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND ( (CP10>='101' AND CP10<='105') OR CP10='125' OR (CP10>='301' AND CP10<='307') OR CP10='107' OR CP10='110' OR CP10='111' OR CP10='112' OR CP10='204' OR CP10='501' ) AND CP27 IS NOT NULL AND CP24 IS NULL AND CP25 IS NULL ORDER BY CP27 DESC "
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      '2010/4/6 MODIFY BY SONIA
      'If rsA.RecordCount > 0 Then
      Do While rsA.EOF = False
         strSql = "UPDATE CASEPROGRESS SET CP24='1', CP25=" & DBDATE(Me.Text5.Text) & " WHERE CP09='" & rsA("CP09") & "'"
         cnnConnection.Execute strSql
         rsA.MoveNext
      Loop
      'End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   '2008/8/29 end
   
   'Added by Morgan 2021/6/10 新增445專利權期限補償期限
   If bolAddNP445 Then
      strExc(1) = CompDate(1, 3, Text5) '法限=公告日(發證日)+3個月
      'Added by Lydia 2025/10/29
      stNP23 = ""
      If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
         str445NP08 = PUB_GetPOurDeadline(strExc(1), pa(9), stNP23, pa(1), "445")
      Else
      'end 2025/10/29
         str445NP08 = PUB_GetWorkDay1(CompDate(2, -10, strExc(1)), True) '所限=法限-10天(最近工作天)
      End If  'Added by Lydia 2025/10/29
      '相關號也要更新否則案件回覆單不會印出來
      'Modified by Lydia 2025/10/29 +np23=" & IIf(stNP23 = "", "NP23", stNP23)
      strSql = "update nextprogress set np01='" & autonum & "',np08=" & str445NP08 & ",np09=" & strExc(1) & ",np23=" & IIf(stNP23 = "", "NP23", stNP23) & " where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06 is null and np07='445'"
      cnnConnection.Execute strSql, intI
      If intI = 0 Then
         strProgressNo = GetNextProgressNo
         'Modified by Lydia 2025/10/29 +NP23
         strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09," & _
             "NP10,NP22,NP23) VALUES ('" & autonum & "','" & pa(1) & "','" & pa(2) & _
             "','" & pa(3) & "','" & pa(4) & "',445," & str445NP08 & _
             "," & strExc(1) & ",'" & stCP13 & "','" & strProgressNo & "', " & CNULL(stNP23, True) & ")"
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2021/6/10
   
   'Added by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail
   'Modified by Lydia 2023/05/26 已閉卷不通知
   'Move by Lydia 2023/05/26 從commit上方移過來,
   Dim bolFMP2mail As Boolean  'Added by Lydia 2023/05/26
   If m_bolFMP = True And m_bolFMP2 = True And pa(57) = "" Then
       'Modified by Lydia 2023/10/31 傳入C類收文號 strReceiveNo
       bolFMP2mail = Pub_SetFMP2toCMail(pa(1), pa(2), pa(3), pa(4), "1603", strUserNum, strReceiveNo)
   End If
   'end 2023/05/17
   'Modified by Lydia 2023/05/26 排除-寰華案無期限之官方來函，系統自動發Mail => And bolFMP2mail = False
   'Added by Morgan 2023/4/11
   'FMP案EMail通知承辦(智權人員)
   m_bolFMPNoPrint = False
   If m_bolFMP And Left(Pub_StrUserSt03, 1) <> "F" And bolFMP2mail = False Then
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " select '" & strUserNum & "',cp13,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'FMP案'||m1.cpm04||'核發通知:'||c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04)||'案專利證書已核發，請參考卷宗區！','如旨',st52" & _
            " from caseprogress c1,casepropertymap m1,staff " & _
            " where c1.cp09='" & strReceiveNo & "' and m1.cpm01(+)=c1.cp01 and m1.cpm02(+)=c1.cp10" & _
            " and st01(+)=cp13"
      cnnConnection.Execute strSql, intI
      
      'Modified by Morgan 2023/6/14 +電子證書才不印通知函,紙本證書的話還是要印
      If PUB_IsECertificate(pa(9), pa(1), pa(8)) = True Then
         m_bolFMPNoPrint = True
      End If
   End If
   'end 2023/4/11
   
'Add By Cheng 2002/11/08
cnnConnection.CommitTrans

   'Add by Sindy 2018/1/2
   If m_strIR01 <> "" And strBillNo <> "" Then
      MsgBox "已新增帳單【 " & strBillNo & " 】。", vbInformation
   End If
   '2018/1/2 END

'Add by Morgan 2006/5/25
m_bPaperPrompt = False
m_bAnnuityInform = False
'end 2006/5/25

'add by nickc 2006/05/05 印結案單
    If m_HaveHK = True Then
      If StrlMAXbyNick <> "" Then
        MsgBox "請更換紙張！", , "列印接洽單！"
        m_bPaperPrompt = True 'Add by Morgan 2006/5/25
        VarlMaxByNick = Split(StrlMAXbyNick, ",")
        For jjjbyNick = 0 To UBound(VarlMaxByNick)
           If Val(VarlMaxByNick(jjjbyNick)) <> 0 Then
              g_PrtForm001.PrintForm Trim(Val(VarlMaxByNick(jjjbyNick))), m_HK_CP01, m_HK_CP02, m_HK_CP03, m_HK_CP04
           End If
        Next jjjbyNick
       End If
    End If
'Remove by Morgan 2010/9/3 核准或催年費已有通知，此處改人工管控即可(因發生多次重複通知且收文狀況)--敏惠
'    'Add by morgan 2006/5/25
'    '大陸發證時如下次繳費日期與系統日相比於二個月內,自動產生年費通知定稿及接洽結案單,於存檔時出現訊息告知更換紙張列印接洽結案單--陳玲玲
'    '2008/12/12 MODIFY BY SONIA
'    'If pa(9) = "020" And Text9 <> "" And m_str605NP22 <> "" Then
'    If pa(9) = "020" And Text9 <> "" And m_str605NP22 <> "" And m_bln605 = True Then
'    '2008/12/12 END
'      If Val(CompDate(1, -2, Text9)) < Val(strSrvDate(1)) Then
'         If m_bPaperPrompt = False Then
'            MsgBox "請更換紙張！", , "列印接洽單！"
'         End If
'         g_PrtForm001.PrintForm m_str605NP22, pa(1), pa(2), pa(3), pa(4)
'         m_bAnnuityInform = True
'      End If
'    End If
'    'end 2006/5/25
    
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    FormSave = False
    If strErrMsg <> "" Then MsgBox strErrMsg, vbCritical 'Added by Morgan 2016/6/30
End Function


Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   'Modify by Morgan 2007/3/26 加澳門
   'If pa(9) = "013" And Text10 = "" Then
   If (pa(9) = "013" Or pa(9) = "044") And Text10 = "" Then
      MsgBox "香港/澳門專利請輸入香港/澳門領證費", vbInformation
      Text10.SetFocus
      Cancel = True
   End If
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2002/11/01
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
    'Add By Cheng 2002/10/31
    If Me.Text12.Text <> "" Then
        'Modify by Morgan 2005/5/4 改西元
        'If ChkDate(Text12) = False Then
        If CheckIsDate(Text12) = False Then
            Cancel = True
           TextInverse Text12
         'Add by Morgan 2006/4/25 檢查不可大於系統日
        ElseIf Val(Text12) > Val(strSrvDate(1)) Then
            MsgBox "帳單日期不可大於系統日！", vbExclamation
            Cancel = True
        End If
    End If
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
    'Add By Cheng 2002/10/31
    If Me.Text13.Text <> "" Then
        If IsNumeric(Me.Text13.Text) = False Then
            MsgBox "帳單金額輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
           TextInverse Text13
        'Add by Morgan 2004/1/30
        ElseIf Val(Text13) <> 0 Then
            If m_CP44 = "" Then
                MsgBox "該筆進度資料無代理人，不可輸入帳單!!!", vbExclamation + vbOKOnly
                Cancel = True
                TextInverse Text13
            End If
        'Add end ---------------
        End If
    End If
End Sub

'Removed by Morgan 2014/11/24 改用 ClsPDGetCasePreAgent
''Add by Morgan 2004/1/30
'Private Function GetCP44() As String
'    Dim StrSQLa As String
'    Dim rsA As New ADODB.Recordset
'On Error GoTo ErrHnd
'    'Modify by Morgan 2007/6/11
'    '改抓AB類最大發文日的最大收文號
'    'StrSQLa = "Select CP44 From CaseProgress Where " & ChgCaseprogress(Me.Text1.Text & Me.Text2.Text & Me.Text3.Text & Me.Text4.Text) & " And  CP10='601' AND CP27 IS NOT NULL "
'    StrSQLa = "Select CP44 From CaseProgress Where " & ChgCaseprogress(Me.Text1.Text & Me.Text2.Text & Me.Text3.Text & Me.Text4.Text) & " And  CP27>0 and cp44 is not null and cp09<'C' order by cp27 desc,cp09 desc"
'    '2007/6/11 end
'    rsA.CursorLocation = adUseClient
'    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'    If rsA.RecordCount > 0 Then
'        GetCP44 = "" & rsA.Fields(0).Value
'    Else
'        GetCP44 = ""
'    End If
'ErrHnd:
'    If Err.NUMBER <> 0 Then
'        MsgBox Err.Description
'    End If
'End Function

Private Sub Text14_GotFocus()
   InverseTextBox Text14
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text15_GotFocus()
   InverseTextBox Text15
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text15_Validate(Cancel As Boolean)
   If Text15 <> "" Then
      If Text15 <> "N" Then
         ShowMsg MsgText(9044)
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 = "" Then Exit Sub
   'Modify by Morgan 2005/9/2
   '台灣
   If pa(9) = "000" Then
      If ChkDate(Text5) = False Then
         Cancel = True
         
      ElseIf Val(Text5.Text) > Val(strSrvDate(2)) Then
         MsgBox "發證日不可大於系統日", vbInformation
         Cancel = True
         
      End If
      
   '非台灣
   Else
      If Len(Text5) < 8 Then
         MsgBox "請輸入西元年!"
         Cancel = True
         
      ElseIf ChkDate(Text5) = False Then
         Cancel = True
         
      '香港標準專利
      ElseIf pa(9) = "013" And pa(8) = "1" Then
         SetHongKongFeeDate
   
      '只要是發證日起算的都重設專用期起日
      Else
         GetPatDef pa(9), pa(8), strExc(1), strExc(2)
         If strExc(1) = "6" Then
            DATE1 = Text5
            Text6 = DATE1 '預設專用期起日
         End If
         
      End If
      
   End If
   'end 2007/3/29
   If Cancel Then TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = "" Then Exit Sub
   Cancel = False
   If ChkDate(Text6) Then
      If TransDate(Text6, 2) <> DATE1 Then
         'Add by Morgan 2004/8/3
         '積體電路佈局專用期可以不是申請日起算
         If m_bol117Exist = True Then
            If m_bolNoMsg = False Then
               If MsgBox("確定專用期限起日日期不為 " & TransDate(DATE1, 2) & " (申請日起算)？", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                  Cancel = True
               Else
                  DATE1 = TransDate(Text6, 2)
               End If
            End If
         Else
            If pa(9) = "000" Then
               'Modify by Morgan 2004/11/10
               'MsgBox "專用期限起日日期應為" & TransDate(DATE1, 1), vbCritical
               MsgBox "專用期限起日日期應為" & TransDate(DATE1, 2), vbCritical
            Else
               MsgBox "專用期限起日日期應為" & TransDate(DATE1, 2), vbCritical
            End If
            Cancel = True
         End If
      End If
   Else
      Cancel = True
   End If
   If Cancel Then TextInverse Text6
End Sub

Private Sub Text7_GotFocus()
   'Modify by Morgan 2004/9/7 台灣控制游標停在第二碼
   If pa(9) = "000" And Len(Text7.Text) = 1 Then
      Text7.SelStart = 1
   Else
      Text7.SelStart = 0
      Text7.SelLength = Len(Text7)
   End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   Dim bolCancel As Boolean
On Error Resume Next
   Cancel = False
   
   If m_bol117Exist = False Then       '2007/4/24 非積體電路案
      If pa(9) = "000" Then
         'Modify by Morgan 2004/8/4
         '新法證書號改7碼
         If m_bolNew = True Then
            If Not Len(Text7.Text) = 7 Then
               MsgBox "輸入的專利號數錯誤"
               Cancel = True
            ElseIf Text7.Text <> pa(22) And pa(22) <> "" Then
               MsgBox "專利號數應為【" & pa(22) & "】！", vbCritical
               Cancel = True
            'Add by Morgan 2004/8/18
            ElseIf txtPA14.Text = "" Then
               txtPA14.Text = pa(14)
               txtPA14_Validate bolCancel
            End If
         Else
            If Not Len(Text7.Text) = 6 Then
               MsgBox "輸入的專利號數錯誤"
               Cancel = True
            End If
         End If
      ElseIf pa(9) = "020" Then
         '93.5.5 modify by sonia
         'If Mid(Text7.Text, 1, 2) <> "ZL" Then
         '   MsgBox "輸入的專利號數錯誤"
         '   Cancel = True
         'ElseIf Right(((Mid(Text7.Text, 3, 1) * 2 + Mid(Text7.Text, 4, 1) * 3 + Mid(Text7.Text, 5, 1) * 4 + Mid(Text7.Text, 6, 1) * 5 + Mid(Text7.Text, 7, 1) * 6 + Mid(Text7.Text, 8, 1) * 7 + Mid(Text7.Text, 9, 1) * 8 + Mid(Text7.Text, 10, 1) * 9) Mod 11), 1) <> Val(Mid(Text7.Text, 12, 1)) Then
         If Text7.Text <> "ZL" + Label2.Caption Then
         '93.5.5 end
            MsgBox "輸入的專利號數錯誤"
            Cancel = True
            Text7.SetFocus
            Text7.SelStart = 0
            Text7.SelLength = Len(Text7.Text)
            
         End If
      ElseIf pa(9) = "013" Then
         'Add by Morgan 2007/6/5 香港設計則檢查證書號需同申請號
         If pa(8) = "3" Then
            If Text7.Text <> pa(11) Then
               MsgBox "輸入的專利號數錯誤"
               Cancel = True
               Text7.SetFocus
               Text7_GotFocus
            End If
         'end 2007/6/5
         ElseIf Mid(Text7.Text, 1, 2) <> "HK" Then
            MsgBox "輸入的專利號數錯誤"
            Cancel = True
         End If
      End If
   End If
   
   If Cancel Then Text7_GotFocus
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 = "" Then Exit Sub
   Cancel = False
   If ChkDate(Text8) Then

      If TransDate(Text8, 2) <> DATE2 Then
         'Add by Morgan 2004/8/3
         '積體電路佈局專用期可以不是申請日起算
         If m_bol117Exist = True Then
            If m_bolNoMsg = False Then
               If MsgBox("確定專用期限迄日日期不為 " & TransDate(DATE2, 2) & " (申請日起算)？", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                  Cancel = True
               Else
                  DATE2 = TransDate(Text8, 2)
               End If
            End If
         Else
            If pa(9) = "000" Then
               'Modify by Morgan 2004/10/6 專用期改西元
               'MsgBox "專用期限迄日日期應為" & TransDate(DATE2, 1), vbCritical
               MsgBox "專用期限迄日日期應為" & TransDate(DATE2, 2), vbCritical
            Else
               MsgBox "專用期限迄日日期應為" & TransDate(DATE2, 2), vbCritical
            End If
            Cancel = True
         End If
      End If

   Else
      Cancel = True
   End If
   If Cancel Then TextInverse Text8
End Sub

Private Sub Text9_GotFocus()
   TextInverse Text9
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   'Add by Morgan 2004/8/3
   If Text9.Locked = True Then Exit Sub
   
   If ChkDate(Text9) Then
      If pa(9) < "010" Then
         If IsEmptyText(m_strNextDueDate) = False Then
            If DBDATE(Text9) <> DBDATE(m_strNextDueDate) Then
               MsgBox "下次繳費日期應為<" & TAIWANDATE(m_strNextDueDate) & ">", vbOKOnly, "檢核資料"
               Cancel = True
            End If
         End If
      Else
         If IsEmptyText(m_strNextFeeDate) = False Then
            If DBDATE(Text9) <> DBDATE(m_strNextFeeDate) Then
               MsgBox "下次繳費日期應為<" & TAIWANDATE(m_strNextFeeDate) & ">", vbOKOnly, "檢核資料"
               Cancel = True
            End If
         End If
      End If
   Else
       Cancel = True
   End If
   If Cancel Then TextInverse Text9
End Sub

'取得客戶檔名稱
Private Function GetCustomerName(CUSTOMER As String)
Dim GetCustomer As New ADODB.Recordset
If GetCustomer.State = adStateOpen Then GetCustomer.Close
strExc(0) = "SELECT NVL(CU04,NVL(CU05,CU06)) FROM CUSTOMER WHERE CU01=SUBSTR('" & CUSTOMER & "' ,1,8) AND CU02= SUBSTR('" & CUSTOMER & "',9,1)"
GetCustomer.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
If GetCustomer.BOF And GetCustomer.EOF Then
    GetCustomerName = ""
Else
    If IsNull(GetCustomer.Fields(0).Value) Then
        GetCustomerName = ""
    Else
        GetCustomerName = GetCustomer.Fields(0).Value
    End If
End If
End Function

Private Function CheckDataValid() As Boolean
CheckDataValid = False
'檢查香港領證費
'Modify by Morgan 2007/3/26 加澳門
'If pa(9) = "013" And Text10 = "" Then
If (pa(9) = "013" Or pa(9) = "044") And Text10 = "" And Text10.Enabled Then
   MsgBox "香港/澳門專利請輸入香港/澳門領證費", vbInformation
   Me.Text10.SetFocus
   Text10.SetFocus
   Exit Function
End If
CheckDataValid = True
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

TxtValidate = False
If Me.Text10.Enabled = True Then
   Cancel = False
   Text10_Validate Cancel
   If Cancel = True Then
      Me.Text10.SetFocus
      Text10_GotFocus
      Exit Function
   End If
End If

If Me.Text5.Enabled = True Then
   Cancel = False
   Text5_Validate Cancel
   If Cancel = True Then
      Me.Text5.SetFocus
      Text5_GotFocus
      Exit Function
   End If
End If


If Me.Text6.Enabled = True Then
   Cancel = False
   'Add by Morgan 2004/8/3
   '最後檢查才
   m_bolNoMsg = True
   Text6_Validate Cancel
   m_bolNoMsg = False
   If Cancel = True Then
      Me.Text6.SetFocus
      Text6_GotFocus
      Exit Function
   End If
End If

If Me.Text7.Enabled = True Then
   Cancel = False
   Text7_Validate Cancel
   If Cancel = True Then
      Me.Text7.SetFocus
      Text7_GotFocus
      Exit Function
   End If
End If

If Me.Text8.Enabled = True Then
   Cancel = False
   m_bolNoMsg = True
   Text8_Validate Cancel
   m_bolNoMsg = False
   If Cancel = True Then
      Me.Text8.SetFocus
      Text8_GotFocus
      Exit Function
   End If
End If

If Me.Text9.Enabled = True Then
   Cancel = False
   Text9_Validate Cancel
   If Cancel = True Then
      Me.Text9.SetFocus
      Text9_GotFocus
      Exit Function
   End If
End If

If Me.Text12.Enabled = True Then
   Cancel = False
   Text12_Validate Cancel
   If Cancel = True Then
      Me.Text12.SetFocus
      Text12_GotFocus
      Exit Function
   End If
End If

If Me.Text15.Enabled = True Then
   Cancel = False
   Text15_Validate Cancel
   If Cancel = True Then
      Me.Text15.SetFocus
      Text15_GotFocus
      Exit Function
   End If
End If

'Add by Morgan 2006/1/16
'檢查公告號
If txtPA15.Visible = True Then
   If txtPA15 = "" Or txtPA15 = "CN" Then
      MsgBox "請輸入公告號！", vbExclamation
      txtPA15.SetFocus
      Exit Function
   Else
      Cancel = False
      txtPA15_Validate Cancel
      If Cancel = True Then
         Me.txtPA15.SetFocus
         txtPA15_GotFocus
         Exit Function
      End If
   End If
End If
         
'Add By Cheng 2003/12/22
'若有輸入代理人D/N No.或帳單日期
If Me.Text11.Text <> "" Or Me.Text12.Text <> "" Then
   'Modify By Sindy 2009/06/17 若為專利處只須以代理人+代理人D/N No.做重覆檢核
   'Modified by Morgan 2014/11/24
   'If Text1 = "P" And Left(Trim(GetStaffDepartment(strUserNum)), 2) = "P1" Then
   If Text1 = "P" And Left(Pub_StrUserSt03, 2) = "P1" Then
   'end 2014/11/24
      '若有輸入代理人D/N No.
      If Me.Text11.Text <> "" Then
         If PUB_ChkDNDup("", ChangeCustomerL(m_CP44), Text11.Text) = True Then
            Text11.SetFocus
            Text11_GotFocus
            Exit Function
         End If
      End If
   Else
   'Remove by Morgan 2004/1/30
   '為了要共用改在 Active 的時候就抓
   '    strSQLA = "Select CP44 From CaseProgress Where " & ChgCaseprogress(Me.Text1.Text & Me.Text2.Text & Me.Text3.Text & Me.Text4.Text) & " And  CP10='601' AND CP27 IS NOT NULL "
   '    rsA.CursorLocation = adUseClient
   '    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
   '    If rsA.RecordCount > 0 Then
   '        m_CP44 = "" & rsA.Fields(0).Value
   '    Else
   '        m_CP44 = ""
   '    End If
   'Remove end---
   
   'Modify by Morgan 2006/4/26 改Call共用函數
   '    If rsA.State <> adStateClosed Then rsA.Close
   '    Set rsA = Nothing
   '    StrSQLa = "Select * From ACC150 Where  A1502=" & (Val(DBDATE(Me.Text12.Text)) - 19110000) & " And A1503='" & ChangeCustomerL(m_CP44) & "' And A1504='" & Me.Text11.Text & "' "
   '    rsA.CursorLocation = adUseClient
   '    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   '    If rsA.RecordCount > 0 Then
   '        MsgBox "此帳單資料重覆，請確認!!!", vbExclamation + vbOKOnly
   '        If rsA.State <> adStateClosed Then rsA.Close
   '        Set rsA = Nothing
   '        Exit Function
   '    End If
   '    If rsA.State <> adStateClosed Then rsA.Close
   '    Set rsA = Nothing
      If PUB_ChkDNDup(Text12.Text, ChangeCustomerL(m_CP44), Text11.Text) = True Then
         Text11.SetFocus
         Text11_GotFocus
         Exit Function
      End If
   '2006/4/26 end
   End If
   
   'Added by Morgan 2016/6/30 非臺灣案電子化
   'Removed by Morgan 2025/8/13 帳單已全部都電子化
   'If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
   'end 2025/8/13
      '匯入該案帳單電子檔
      If Not PUB_ImportInvoice(pa(1), pa(2), pa(3), pa(4)) Then
         Exit Function
      End If
   'End If
   'end 2016/6/30
End If

'Add by Morgan 2007/6/11
If Me.Text13.Enabled = True Then
   Cancel = False
   Text13_Validate Cancel
   If Cancel = True Then
      Me.Text13.SetFocus
      Text13_GotFocus
      Exit Function
   End If
End If
'end 2007/6/11


'Added by Morgan 2014/4/14 電子化-檢查pdf檔
'Removed by Morgan 2014/5/15 消滅函,證書函不必檢查,程序輸入後才掃描--陳玲玲
'If pa(9) = "000" Then
'   If PUB_CheckPDF(pa(1), pa(2), pa(3), pa(4), 1) = False Then
'      Exit Function
'   End If
'End If
'end 2014/4/14

'Added by Morgan 2021/6/3
'大陸修法專利權期限補償提醒
bolAddNP445 = False: str445NP08 = ""
'Modified by Morgan 2021/11/10 更正，應為核准日(原以公告日計算) --郭
'Modified by Morgan 2021/11/15 再改為以公告日計算，另申請日PCT案為進國家階段的日期，分割案則為提交日。 --郭
strExc(1) = DBDATE(Text5)
If pa(8) = "1" And pa(9) = "020" And strExc(1) >= "20210601" Then
   If PUB_DualCaseRelationExist(pa()) = False Then 'Added by Morgan 2023/7/18 一案兩請除外  --郭

      strExc(2) = DBDATE(pa(10))
      'Added by Morgan 2021/11/15
      If pa(46) = "Y" Then
         strExc(0) = "select cp47 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in (" & NewCasePtyList & ") and cp47>19221111"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(2) = RsTemp("cp47")
         Else
            MsgBox "PCT進國家階段日期讀取失敗，無法判斷本案是否符合專利權期限補償，請確認！", vbExclamation
            Exit Function
         End If
      Else
         strExc(0) = "select cp47 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='307'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp("cp47") > 19221111 Then
               strExc(2) = RsTemp("cp47")
            Else
               MsgBox "分割提交日讀取失敗，無法判斷本案是否符合專利權期限補償，請確認！", vbExclamation
               Exit Function
            End If
         End If
      End If
      'end 2021/11/15
      
      '>申請日+4年
      strExc(2) = Val(strExc(2)) + 40000
      If strExc(1) > strExc(2) Then
         'Modified by Lydia 2021/11/22 判斷實審與發明同時提申的狀況; ex. P-116795
         'strExc(0) = "select cp47 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='416' and cp47>19221111"
         strExc(0) = "select c1.cp47 as cp47a, c1.cp27 as cp27a, c2.cp47 as cp47b, c2.cp27 as cp27b " & _
                          "from caseprogress c1, caseprogress c2 " & _
                          "where c1.cp01='" & pa(1) & "' and c1.cp02='" & pa(2) & "' and c1.cp03='" & pa(3) & "' and c1.cp04='" & pa(4) & "' and c1.cp10 in (" & NewCasePtyList & ") and c1.cp158>0 and c1.cp159=0 " & _
                          "and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and c2.cp10='416' and c2.cp159=0 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '實審提申日+3年
            'Modified by Lydia 2021/11/22 判斷實審與發明同時提申的狀況; ex. P-116795
            'strExc(2) = DBDATE(Val(RsTemp("cp47")))
            'If strExc(1) > strExc(2) Then
            strExc(2) = ""
            '實審另外提申
            If Val("" & RsTemp.Fields("cp47b")) > 19221111 Then
                strExc(2) = DBDATE(Val(RsTemp("cp47b")))
            '實審與發明同時提申
            ElseIf Val("" & RsTemp.Fields("cp47a")) > 19221111 And Val("" & RsTemp.Fields("cp27a")) = Val("" & RsTemp.Fields("cp27b")) Then
                strExc(2) = DBDATE(Val(RsTemp("cp47a")))
            End If
            If strExc(2) = "" Then
                 MsgBox "實體審查提申日讀取失敗，無法判斷本案是否符合專利權期限補償，請確認！", vbExclamation
                 Exit Function
            Else
            'end 2021/11/22
               'Added by Morgan 2023/2/16 若是實審提交日早於公開日,則要以公開日為主 --郭雅娟
               If pa(12) <> "" Then
                  strExc(3) = DBDATE(pa(12))
                  If strExc(2) < strExc(3) Then
                     strExc(2) = strExc(3)
                  End If
               End If
               'end 2023/2/16
               strExc(2) = Val(strExc(2)) + 30000
               If strExc(1) > strExc(2) Then
                  'Modified by Morgan 2021/6/10 改管制期限並帶入定稿
                  'MsgBox "本案符合專利權期限補償,請修改定稿！", vbExclamation
                  bolAddNP445 = True
                  'end 2021/6/10
               End If
            End If
         Else
              MsgBox "實體審查提申日讀取失敗，無法判斷本案是否符合專利權期限補償，請確認！", vbExclamation
              Exit Function
         End If
      End If
   End If
End If
'end 2021/6/3

TxtValidate = True
End Function

Private Sub txtPA14_GotFocus()
   TextInverse txtPA14
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtPA14.IMEMode = 2
   CloseIme
End Sub

Private Sub txtPA14_KeyPress(KeyAscii As Integer)
   '只能key倒退鍵和數字
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtPA14_Validate(Cancel As Boolean)
   If m_bol117Exist = False Then     '2007/4/24 加入非積體電路案條件
      If txtPA14.Text = "" Then
         MsgBox "公告日不可空白！", vbCritical
         Cancel = True
      ElseIf Not ChkDate(txtPA14) Then
         MsgBox "日期格式錯誤！", vbCritical
         Cancel = True
      ElseIf pa(14) <> "" Then
         If Val(txtPA14) <> Val(pa(14)) Then
            MsgBox "公告日應為【 " & ChangeTStringToTDateString(pa(14)) & " 】！", vbCritical
            txtPA14_GotFocus
            Cancel = True
         End If
      End If
      '若用新法則專用期起日=公告日
      If Cancel = False And m_bolNew = True Then
         'Modify by Morgan 2004/10/6 專用期改西元 *
         'Text6.Text = txtPA14.Text
         Text6.Text = TransDate(txtPA14.Text, 2)
         DATE1 = TransDate(Text6, 2)
         'Modify by Morgan 2005/2/17 繳10年以上會有錯
         'm_strNextDueDate = CompDate(0, Val(Right(pa(72), 1)), DATE1)
         strExc(0) = Right(pa(72), 2)
         If Left(strExc(0), 1) = "," Then strExc(0) = Right(strExc(0), 1)
         m_strNextDueDate = CompDate(0, Val(strExc(0)), DATE1)
         '2005/2/17 end
         m_strNextDueDate = CompDate(2, -1, m_strNextDueDate)
         'Added by Lydia 2025/10/29
         If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
            m_strNextFeeDate = PUB_GetPOurDeadline(m_strNextDueDate, pa(9))
         Else
         'end 2025/10/29
            'Added by Morgan 2014/10/28
            If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
               m_strNextFeeDate = PUB_GetOurDeadline(m_strNextDueDate)
            Else
            'end 2014/10/28
               m_strNextFeeDate = CompDate(2, -2, m_strNextDueDate)
               m_strNextFeeDate = PUB_GetWorkDay1(m_strNextFeeDate, True)
            End If 'Added by Morgan 2014/10/28
         End If 'Added by Lydia 2025/10/29
      End If
   End If
End Sub
'Add by Morgan 2005/6/14
'計算香港下次繳費日
Private Sub SetHongKongFeeDate()
   Dim stYear As String, strTmp1 As String
   '以發證月日-申請月日判斷繳費年,若>0則+4否則+3
   m_iYear = 0
   Text9 = ""
   If Text5 <> "" And DATE1 <> "" Then
      m_iYear = Val(Left(Text5, 4)) - Val(Left(DATE1, 4)) + IIf(Right(Text5, 4) - Right(DATE1, 4) > 0, 4, 3)
      m_strNextDueDate = CompDate(0, m_iYear, DATE1)
      'Modified by Morgan 2014/7/15
      'Modified by Morgan 2018/10/3 非FMP也改10天
      'If m_bolFMP Then
         'Added by Lydia 2025/10/29
         If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
            m_strNextFeeDate = PUB_GetPOurDeadline(m_strNextDueDate, pa(9))
         Else
         'end 2025/10/29
            m_strNextFeeDate = CompDate(2, -10, m_strNextDueDate)
         End If 'Added by Lydia 2025/10/29
      'Else
      'end 2014/7/15
      '   m_strNextFeeDate = CompDate(1, -1, m_strNextDueDate)
      '   m_strNextFeeDate = CompDate(2, -5, m_strNextFeeDate)
      'End If 'Added by Morgan 2014/7/15
      'end 2018/10/3
      
      'Add by Morgan 2006/11/2 本所期限要抓工作天
      m_strNextFeeDate = PUB_GetWorkDay1(m_strNextFeeDate, True)
      If GetMoneyDate(Val(pa(8)) + 10, pa(9), pa(), DATE1, strTmp1, DATE2) Then   '抓專用期起止日
         'DATE2 = CompDate(2, 1, DATE2)       '2008/11/6 cancel by sonia 加在GetMoneyDate控制
      End If
      If m_strNextDueDate < DATE2 Then
         Text9 = ChangeWStringToTString(m_strNextFeeDate)
      End If
   End If
End Sub

Private Sub txtPA15_GotFocus()
   'Modify by Morgan 2006/4/10 預設CN
   If txtPA15 = "" Or txtPA15 = "CN" Then
      txtPA15 = "CN"
      txtPA15.SelStart = 3
   Else
      TextInverse txtPA15
   End If
End Sub

Private Sub txtPA15_Validate(Cancel As Boolean)
   If txtPA15.Tag <> "" And txtPA15.Text <> "" And txtPA15.Text <> "CN" Then
      If txtPA15.Text <> txtPA15.Tag Then
         If MsgBox("確定要更改公告號？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
            txtPA15.Tag = Text15.Text
         Else
            Cancel = True
         End If
      End If
   End If
End Sub

'Add by Morgan 2007/3/29
'取得專利各期限起算方式
'Input
'  PA09:申請國家
'  PA08:1=發明專用期,2=新型專用期,3=設計專用期
'Output
'  pFrom:起日
'  pTo:止日
Private Sub GetPatDef(ByVal PA09 As String, ByVal PA08 As String, ByRef pFrom As String, ByRef pTo As String)
   Dim strCol As String
   Select Case PA08
      Case "1"
         strCol = " na40 , na41"
      Case "2"
         strCol = " na43 , na44"
      Case "3"
         strCol = " na46 , na47"
      Case Else
         strCol = " na40 , na41"
   End Select
   strExc(0) = "select " & strCol & " from nation where na01='" & PA09 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      pFrom = "" & RsTemp.Fields(0)
      pTo = "" & RsTemp.Fields(1)
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
   
   strExc(0) = "select lastyear(pa72),np09 from nextprogress,patent" & _
      " where np02='" & pa(1) & "' and np03='" & pa(2) & "'" & _
      " and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06||np07='605'" & _
      " and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','下次年費年度','" & (Val("" & RsTemp(0)) + 1) & "')"
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','下次年費法限','" & RsTemp(1) & "')"
   End If
         
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Sub

Private Sub StartLetter3(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt() As String, i As Integer
   
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   i = 0
   strExc(0) = ChgEngDate(DBDATE(pa(10)))
   strExc(0) = Left(strExc(0), Len(strExc(0)) - 6)
   i = i + 1
   ReDim Preserve strTxt(i)
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','申請月日','" & strExc(0) & "')"
         
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub StartLetter4(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)

   Dim strTxt() As String, ii As Integer
   
   EndLetter ET01, ET02, ET03, strUserNum
    
   ii = 1
   ReDim Preserve strTxt(ii)
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','大陸案申請號','" & pa(11) & "')"
   
   ii = ii + 1
   ReDim Preserve strTxt(ii)
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','大陸案公告日','" & pa(14) & "')"
   
   ii = ii + 1
   ReDim Preserve strTxt(ii)
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','本所期限','" & m_HK_NP08 & "')"
       
   ii = ii + 1
   ReDim Preserve strTxt(ii)
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       "','法定期限','" & m_HK_NP09 & "')"
         
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If

End Sub
