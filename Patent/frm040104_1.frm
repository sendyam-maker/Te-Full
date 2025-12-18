VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040104_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專發文"
   ClientHeight    =   5748
   ClientLeft      =   -2676
   ClientTop       =   1572
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9348
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   8388
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "延期(&D)"
      Height          =   405
      Index           =   0
      Left            =   6744
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   7560
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   120
      TabIndex        =   14
      Top             =   528
      Width           =   9072
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "申請案號"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   4440
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "P"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   6510
         TabIndex        =   7
         Top             =   180
         Width           =   800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PS: 已收文下一程序，欲辦理延期者，請以本所案號輸入搜尋後，再按延期按鈕"
         Height          =   300
         Left            =   2040
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   6216
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3615
      Left            =   120
      TabIndex        =   13
      Top             =   1980
      Width           =   9075
      _ExtentX        =   16002
      _ExtentY        =   6371
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   15
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   11
      Top             =   1500
      Width           =   7995
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14102;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   120
      X2              =   9180
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   9180
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   1500
      Width           =   765
   End
End
Attribute VB_Name = "frm040104_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (Combo1,MSHFlexGrid1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit

Dim pa(0 To 10) As String
Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer

Public MainPa9 As String
Public bolLeave As Boolean
Dim m_bolActivated As Boolean
Public bolIsEMPFlow As Boolean 'Add By Sindy 2013/5/20 是否為電子承辦簽核
Public strChoose As String 'Add by Amy 2014/04/08 記錄選擇1.補存取碼 2.優先權證明書(for 因應台日優先權證明文件)
Dim mPA11 As String 'Modified by Lydia 2014/12/9 P非新案發文改以申請案號輸入,原收文號取消
Public mCP09 As String 'Added by Morgan 2020/1/16
Public mPreForm As Form 'Added by Morgan 2020/1/16

Private Sub cmdok_Click(Index As Integer)
   Dim i As Integer, bolChk As Boolean
   ' 90.10.18 modify by louis (記錄總收文號)
   Dim strCP09 As String
   'Add By Cheng 2002/01/24
   Dim strPA23 As String
   'Add by Amy 2018/09/14
   Dim strCP31 As String
   Dim frmNext As Form
   Dim bNext As Boolean
   Dim bNoCheck As Boolean 'Added by Morgan 2019/9/11
   
   bNext = False 'Add by Amy 2018/09/14
   strChoose = "" 'Add by Amy 2015/04/02
   Select Case Index
      Case 0 '延期
         If Option1(0).Value = True Then '本所案號
            With MSHFlexGrid1
               For i = 1 To .Rows - 1
                  If .TextMatrix(i, 0) = "v" Then
                     bolChk = True
                     Me.Tag = .TextMatrix(i, 2)
                     ' 90.10.18 modify by louis (記錄總收文號)
                     strCP09 = .TextMatrix(i, 2)
                     'Add By Cheng 2002/07/15
                     '所點選的案件性質不可為"延期"
                     If PUB_CPKindDelay(strCP09, "P") Then
                        Exit Sub
                     End If
                     'Add By Cheng 2002/07/12
                     '已閉卷案件不可發文
                     If PUB_CaseClosed(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text) Then
                        Exit Sub
                     End If
                     '91.10.26 modify by sonia
                     'If .TextMatrix(i, 10) = "" Or .TextMatrix(i, 10) = "0" Then
                     '   If Not ChkShowMail Then Exit Sub
                     'End If
                     '91.10.26 end
                     
                     'Add by Morgan 2004/9/8 未輸入承辦人不可延期
                     If PUB_ChkCP14IsNull(strCP09) = True Then Exit Sub
                     Exit For
                  End If
               Next
            End With
            If bolChk = False Then
               MsgBox "請選擇資料 !", vbInformation
               Exit Sub
            End If
            Me.Tag = Me.Tag & "0"
         Else '收文號
            'Modified by Lydia 2014/12/9 P非新案發文改以申請案號輸入,原收文號取消
            If Text5 <> "" Then
                MsgBox "欲辦理延期者，請以本所案號輸入搜尋後，再按延期按鈕!!"
                Option1(0).Value = True: Text2.SetFocus
                Exit Sub
'               'Modify By Cheng 2002/04/12
''               strExc(0) = "select CP01,CP02,CP03,CP04 from caseprogress where cp09='" & Text5 & "'" & _
''                  " AND CP27 IS NULL AND CP57 IS NULL AND (SUBSTR(CP09,1,1)='A' OR SUBSTR(CP09,1,1)='B')" & _
''                  " AND (CP01='P' OR CP01='PS')"
'               strExc(0) = "select CP01,CP02,CP03,CP04 from caseprogress where cp09='" & Text5 & "'" & _
'                  " AND CP27 IS NULL AND CP57 IS NULL AND ( CP09<'C' )" & _
'                  " AND (CP01='P' OR CP01='PS')"
'               intI = 0
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  Me.Tag = Text5.Text & "0"
'                  Text1.Text = RsTemp.Fields(0)
'                  Text2.Text = RsTemp.Fields(1)
'                  Text3.Text = RsTemp.Fields(2)
'                  Text4.Text = RsTemp.Fields(3)
'               Else
'                  Exit Sub
'               End If
'               strCP09 = Text5
'            Else
'               MsgBox "總收文號不得為空值 !", vbCritical
'               Exit Sub
            End If
         End If
         'Add By Cheng 2002/07/12
         '已閉卷案件不可發文
         If PUB_CaseClosed(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text) Then
            Exit Sub
         End If
         
         'Added by Morgan 2012/7/10
         '台灣申復修正只能延期一次(FCP有例外,程式不控制--靜芳)
         'Modified by Morgan 2012/12/18 +再審也只能延期1次
         '2013/9/14 modify by sonia 取消再審107(玲玲)P-087953
         strExc(0) = "select cpm03 from caseprogress a,patent,casepropertymap where cp09='" & strCP09 & "' and cp10 in ('204','205') and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='000' and exists(select * from caseprogress b where b.cp43=a.cp09 and b.cp10='404' and b.cp27>0) and cpm01(+)=cp01 and cpm02(+)=cp10"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "台灣案" & RsTemp(0) & "只能延期一次，本程序已有延期紀錄不可再延期！"
            Exit Sub
         End If
         'end 2012/7/10
         
         'Add By Sindy 2013/11/14
         '檢查是否有承辦歷程是否有產生承辦單可以發文
         If PUB_IsEmpFlowIsSend(strCP09) = False Then
            Exit Sub
         End If
         
         'Add By Cheng 2002/06/20
         '延期記錄之資料來源為案件進度檔
         frm040104_2.m_str_DL05 = "1"
         frm040104_2.Show
         Command1.SetFocus
         Me.Hide
      Case 1 '確定
      
'Modify by Morgan 2005/1/4 不必分本所案號或總收文號
'         If Option1(0).Value = True Then '本所案號
      
            With MSHFlexGrid1
               For i = 1 To .Rows - 1
                  If .TextMatrix(i, 0) = "v" Then
                     bolChk = True
                     Me.Tag = .TextMatrix(i, 2)
                     ' 90.10.18 modify by louis (記錄總收文號)
                     strCP09 = .TextMatrix(i, 2)
                    'Modify By Cheng 2002/11/28
'                     pa(10) = .TextMatrix(i, 7)
                     pa(10) = .TextMatrix(i, 9)
                     strCP31 = .TextMatrix(i, 15) 'Add by Amy 2018/09/14
                     
                     'Modify by Amy 2015/01/22 +北所分案日cp157 有值才可發文
                     If .TextMatrix(i, 2) < "B" And .TextMatrix(i, 14) = "" Then
                        MsgBox "北所尚未分案，不可發文!!"
                        Exit Sub
                     End If
                     'end 2015/01/22
                     
                      'Modified by Lydia 2014/12/9 P非新案發文改以申請案號輸入,原收文號取消
                      'Added by Morgan 2019/9/11 非待送件發文才檢查
                      'Modified by Morgan 2020/1/16 +mCP09 = ""
                      If Not bolIsEMPFlow And mCP09 = "" Then
                      'end 2019/9/11
                        If pa(1) = "P" And Len(mPA11) > 0 And InStr(NewCasePtyList & ",601,605,421", pa(10)) = 0 Then
                           
                           'Added by Morgan 2019/9/11 +程序承辦的實審也可用本所案號發文--韻丞(淑華也同意)
                           bNoCheck = False
                           If pa(10) = "416" Then
                              strExc(0) = GetGridValue(i, "CP14_ST03")
                              If strExc(0) = "P12" Or strExc(0) = "F22" Then
                                 bNoCheck = True
                              End If
                           End If
                           If bNoCheck = False Then
                           'end 2019/9/11
                           
                              If Option1(0).Value = True Or (Option1(1).Value = True And Text5 <> mPA11) Then
                                 MsgBox "P案非新案發文請改以申請案號輸入!!"
                                 Option1(1).Value = True: Text5.SetFocus
                                 Exit Sub
                              End If
                              
                           End If 'Added by Morgan 2019/9/11
                        End If
                     End If
                    
                     strExc(0) = ""
                     intI = 1
                     If pa(10) = 領證及繳年費 Then
                        If Text1 = "P" Then
                        intI = 0
                        'Modify By Cheng 2003/05/15
'                        strExc(0) = "SELECT PA14 FROM PATENT WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
                        strExc(0) = "SELECT PA14, PA09 FROM PATENT WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                            'Modify By Cheng 2003/05/15
                            '若申請國家為台灣
                            If "" & RsTemp.Fields(1).Value = 台灣國家代號 Then
                               If Not IsNull(RsTemp.Fields(0)) Then
                                    'Modify By Cheng 2002/11/26
                                    '系統日 >= 公告日 +  三個月 + 8 天
    '                              If Val(strSrvDate(1)) < Val(CompDate(2, 7, rsTemp.Fields(0))) Then
    '                                 MsgBox "系統日必須大於等於公告日 (" & TransDate(rsTemp.Fields(0), 1) & ") + 7天 !", vbCritical
                                  If strSrvDate(1) < DBDATE(DateAdd("D", 8, DateAdd("M", 3, ChangeWStringToWDateString(RsTemp.Fields(0))))) Then
                                     MsgBox "公告尚未期滿，不可發文 !", vbCritical
                                     Exit Sub
                                  End If
                               'Remove by Morgan 2004/6/21
                               '新法先領證故不必再檢查
                               '92.5.13 add by sonia
'                               Else
'                                  MsgBox "尚未公告, 不可發文 !", vbCritical
'                                  Exit Sub
                               '92.5.13 end
                               End If
                            End If
                        End If
                     End If
                     'Add By Cheng 2002/01/24
                     '取得卷宗性質
                     intI = 1
                     strExc(0) = "SELECT PA23 FROM PATENT WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                    'Modify By Cheng 2003/02/26
                     If intI = 1 Then
                        If Not IsNull(RsTemp.Fields(0)) Then
                           strPA23 = RsTemp.Fields(0).Value
                        Else
                           strPA23 = ""
                        End If
                    Else
                        strPA23 = ""
                    End If
                     'Add By Cheng 2002/07/12
                     '已閉卷案件不可發文
                     'Add by Morgan 2006/7/31 退費(908),退證註銷(915),減免退費(919) 2007/7/20加重新委任(928)
                     'Modified by Morgan 2020/1/17 待處理呼叫除外(+mCP09="")
                     If Not (pa(10) = "908" Or pa(10) = "915" Or pa(10) = "919" Or pa(10) = "928") And mCP09 = "" Then
                        If PUB_CaseClosed(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text) Then
                           Exit Sub
                        End If
                     End If
                     'Add By Cheng 2003/09/01
                     '若為領證或年費
                    If pa(10) = "601" Or pa(10) = "605" Then
                        If PUB_ChkNP605(Me.Text1.Text & Me.Text2.Text & Me.Text3.Text & Me.Text4.Text) Then
                            MsgBox "本案下一程序有<年費>期限，不可發文!!!", vbExclamation + vbOKOnly
                           Exit Sub
                        End If
                    End If
                    
                    'Add by Morgan 2010/6/8
                    'Modified by Morgan 2024/11/18 +477再審查加速審查並改用專用模組判斷
                    'If pa(10) = "422" Then
                    '    If PUB_ChkCPExist(pa, "1204") = False Then
                    '       'Modified by Morgan 2012/3/22 因會有中間接來案件改可選擇發文 Ex.P-099833 --敏惠
                    '      'MsgBox "本案尚未接獲實審通知，【加速審查】不可發文！"
                    '       'Exit Sub
                    '       If MsgBox("本案尚未接獲實審通知，【加速審查】不該發文！是否確定要繼續??", vbYesNo + vbDefaultButton2) = vbNo Then
                    If pa(10) = "422" Or pa(10) = "447" Then
                        If PUB_Chk1204(pa) = False Then
                           If MsgBox("本案尚未接獲實審通知，【" & .TextMatrix(i, 3) & "】不該發文！是否確定要繼續??", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
                    'end 2024/11/14
                              Exit Sub
                           End If
                        End If
                    End If
                    
                    'Added by Morgan 2012/8/14
                    '台灣新型分割案母案已有准駁不可發文
                    'Removed by Morgan 2023/5/11 108.11.1 新法發明/新型准後3個月內可提分割
                    'If MainPa9 = "000" And pa(10) = "307" Then
                    '    strExc(0) = "select * from divisioncase,patent where dc01='" & pa(1) & "' and dc02='" & pa(2) & "'" & _
                    '       " and dc03='" & pa(3) & "' and dc04='" & pa(4) & "' and pa01(+)=dc05 and pa02(+)=dc06" & _
                    '       " and pa03(+)=dc07 and pa04(+)=dc08 and pa08='2' and pa09='000' and pa16 is not null"
                    '    intI = 1
                    '    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    '    If intI = 1 Then
                    '       MsgBox "台灣新型分割案母案已有准駁不可發文!!"
                    '       Exit Sub
                    '    End If
                    'End If
                    'end 2023/5/11
                    'end 2012/8/14
                    
                    'Added by Morgan 2012/8/14
                    '台灣衍生設計案母案已公告日不可發文
                    If MainPa9 = "000" And pa(10) = "125" Then
                        strExc(0) = "select sqldatet(pa14) from patent where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='0' and pa04='" & pa(4) & "' and pa14>0"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           MsgBox "本衍生設計案之母案已於 " & RsTemp(0) & " 公告，不可發文！", vbCritical
                           Exit Sub
                        End If
                    End If
                    'end 2012/8/14
                    
                    'Added by Morgan 2012/8/14
                    '台灣申請案發文若主張之國內優先權案已核准或公告時不可發文
                    'Modified by Lydia 2017/05/09 取消台灣案,改控制大陸案
'                    If MainPa9 = "000" And (pa(10) = "101" Or pa(10) = "102" Or pa(10) = "103") Then
'                       strExc(0) = "select b.pa01||'-'||b.pa02||decode(b.pa03||b.pa04,'000','','-'||b.pa03||'-'||b.pa04)||' '||decode(b.pa14,null,'已領證','已公告')" & _
'                        " from patent a,pridate,patent b,caseprogress where a.pa01='" & pa(1) & "' and a.pa02='" & pa(2) & "' and a.pa03='" & pa(3) & "' and a.pa04='" & pa(4) & "'" & _
'                        " and pd01(+)=a.pa01 and pd02(+)=a.pa02 and pd03(+)=a.pa03 and pd04(+)=a.pa04 and pd07='000'" & _
'                        " and b.pa11(+)=pd06 and b.pa09='000' and cp01(+)=b.pa01 and cp02(+)=b.pa02 and cp03(+)=b.pa03 and cp04(+)=b.pa04 and cp10(+)='601' and cp27(+)>0" & _
'                        " and (b.pa14>0 or cp09 is not null)"
                    If MainPa9 = "020" And (pa(10) = "101" Or pa(10) = "102" Or pa(10) = "103") Then
                       strExc(0) = "select b.pa01||'-'||b.pa02||decode(b.pa03||b.pa04,'000','','-'||b.pa03||'-'||b.pa04)||' '||decode(b.pa14,null,'已領證','已公告')" & _
                        " from patent a,pridate,patent b,caseprogress where a.pa01='" & pa(1) & "' and a.pa02='" & pa(2) & "' and a.pa03='" & pa(3) & "' and a.pa04='" & pa(4) & "'" & _
                        " and pd01(+)=a.pa01 and pd02(+)=a.pa02 and pd03(+)=a.pa03 and pd04(+)=a.pa04 and pd07='020'" & _
                        " and b.pa11(+)=pd06 and b.pa09='020' and cp01(+)=b.pa01 and cp02(+)=b.pa02 and cp03(+)=b.pa03 and cp04(+)=b.pa04 and cp10(+)='601' and cp27(+)>0" & _
                        " and (b.pa14>0 or cp09 is not null)"
                    'end 2017/05/09
                       intI = 1
                       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                       If intI = 1 Then
                           MsgBox "本案所主張之國內優先權案 " & RsTemp(0) & "，不可發文!!", vbCritical
                           Exit Sub
                       End If
                    End If
                    
                    'Add By Cheng 2003/10/20
                    '未輸入承辦人不可發文
                    If PUB_ChkCP14IsNull(strCP09) = True Then Exit Sub
                     '91.10.26 modify by sonia
                     'If .TextMatrix(i, 10) = "" Or .TextMatrix(i, 10) = "0" Then
                     '
                     'Else
                     '   If Not ChkShowMail Then Exit Sub
                     'End If
                     '91.10.26 end
                     Exit For
                  End If
               Next
            End With
            If bolChk = False Then
               MsgBox "請選擇資料 !", vbInformation
               Exit Sub
            End If
            
            'Added by Morgan 2012/10/23
            '香港案若有維持費期限過期不可發文標準專利批准紀錄請求
            If MainPa9 = "013" And pa(10) = "111" Then
               strExc(0) = "select cp07,1 Src from caseprogress where cp01='" & pa(1) & "' and  cp02='" & pa(2) & "' and  cp03='" & pa(3) & "' and  cp04='" & pa(4) & "'" & _
                  " and cp10='" & 維持費 & "' and cp27||cp57 is null and cp07<" & strSrvDate(1) & _
                  " union select np09,2 Src from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'" & _
                  " and np07=" & 維持費 & " and np06 is null and np09<" & strSrvDate(1)
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "本案有" & IIf(RsTemp(1) = "1", "已收文", "未收文") & "維持費已逾期，不可發文!!"
                  Exit Sub
               End If
            End If
            'end 2012/10/23
            
            
            'Added by Morgan 2020/3/5
            '申請台灣優先權證明書若申請人都不是台灣籍提醒
            If MainPa9 = "000" And pa(10) = "405" Then
               'Modified by Morgan 2025/9/11 要共用改寫成函數
               'strExc(0) = "select * from patent where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and  pa04='" & pa(4) & "'" & _
               '   " and not exists(select * from customer where cu01||cu02 in (pa26,pa27,pa28,pa29,pa30) and cu10<'010')"
               'intI = 1
               'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               'If intI = 1 Then
               '   If MsgBox("因申請人為外國籍無法用於主張中國大陸申請案，請確認此優先權證明書是否用於中國大陸。", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
               If PUB_ChkNoTWApp(pa) = True Then
                  If PUB_TWPriCertMsg() = vbYes Then
               'end 2025/9/11
                     Exit Sub
                  End If
               End If
            End If
            'end 2020/3/5
      
            
                     
            'Add by Morgan 2011/2/8
            'Modified by Morgan 2023/7/5 +cp164
            'Modified by Morgan 2024/1/26 改統一用函數檢查
'            strExc(0) = "select cp06,cp79,cp141,cp142,cp164 from caseprogress where cp09='" & strCP09 & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               If RsTemp.Fields("cp141") = "2" And RsTemp.Fields("cp79") > 0 Then
'                  If PUB_ChkPaidByCP09(strCP09) = False Then 'Added by Morgan 2016/8/23 出納繳款確認後就可送件
'                     If IsNull(RsTemp.Fields("cp06")) Or RsTemp.Fields("cp06") > strSrvDate(1) Then
'                        MsgBox "本案已設定為收款後送件且尚有未收金額，不可發文！"
'                        Exit Sub
'                     End If
'                  End If
'               'Modified by Morgan 2023/8/31
'               'ElseIf RsTemp.Fields("cp141") = "3" And RsTemp.Fields("cp142") <> strSrvDate(1) And RsTemp.Fields("cp164") = "1" Then
'               '   If MsgBox("本案已設定為指定日期送件但該日期與系統日不符，是否仍要發文？", vbYesNo + vbDefaultButton2) = vbNo Then
'               '      Exit Sub
'               '   End If
'               '指定日期(非台灣案也要提醒--玲玲)
'               ElseIf RsTemp.Fields("cp141") = "3" Then
'                  strExc(0) = ChangeWStringToTDateString(RsTemp.Fields("cp142"))
'                  '之前
'                  If RsTemp.Fields("cp164") = "2" Then
'                     If RsTemp.Fields("cp142") < strSrvDate(1) Then
'                        If MsgBox("本案已設定於指定日期(" & strExc(0) & ")之前送件但與系統日不符，是否仍要發文？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
'                           Exit Sub
'                        End If
'                     End If
'                  '之後
'                  ElseIf RsTemp.Fields("cp164") = "3" Then
'                     If RsTemp.Fields("cp142") > strSrvDate(1) Then
'                        'Added by Morgan 2023/12/28
'                        If MainPa9 = "000" Then
'                           MsgBox "本案需於指定日期(" & strExc(0) & ")之後方可發文！", vbExclamation, "指定日期檢查"
'                           Exit Sub
'                        'end 2023/12/28
'                        ElseIf MsgBox("本案已設定於指定日期(" & strExc(0) & ")之後送件但與系統日不符，是否仍要發文？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
'                           Exit Sub
'                        End If
'                     End If
'                  '當日
'                  ElseIf RsTemp.Fields("cp142") <> strSrvDate(1) Then
'                     'Added by Morgan 2023/12/28
'                     If MainPa9 = "000" And RsTemp.Fields("cp142") > strSrvDate(1) Then
'                        MsgBox "本案需於指定日期(" & strExc(0) & ")方可發文！", vbExclamation, "指定日期檢查"
'                        Exit Sub
'                     'end 2023/12/28
'                     ElseIf MsgBox("本案已設定於指定日期(" & strExc(0) & ")送件但與系統日不符，是否仍要發文？", vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then
'                        Exit Sub
'                     End If
'                  End If
'               'end 2023/8/31
'               End If
'            End If
            If PUB_ChkCP141IsSend(strCP09, True) = False Then Exit Sub
            'end 2011/2/8
            
            'Modify By Cheng 2003/06/16
            'Modify by Morgan 2006/7/31 加修正(203,204)
            'Modified by Morgan 2012/10/8 +衍生設計(125)
            If pa(10) = 發明申請 Or pa(10) = 新型申請 Or pa(10) = 設計申請 Or pa(10) = 追加申請 Or pa(10) = 聯合申請 Or pa(10) = 衍生設計 Or pa(10) = "203" Or pa(10) = "204" Then
                '2006/3/7 MODIFY BY SONIA 加 109,110,112
                'Modify by Morgan 2006/5/9 案件性質改用常數控制
'                strExc(0) = "select cm01||cm02||cm03||cm04,ST02,nvl(pa05,nvl(pa06,pa07)) FROM " & _
'                   "CASEMAP,CASEPROGRESS,PATENT,STAFF WHERE CM05='" & pA(1) & "' AND CM06='" & pA(2) & "' AND " & _
'                   "CM07='" & pA(3) & "' AND CM08='" & pA(4) & "' AND CM10='0' AND " & _
'                   "cm01=pa01 and cm02=pa02 and cm03=pa03 and cm04=pa04 AND " & _
'                   "cm01=cp01 and cm02=cp02 and cm03=cp03 and cm04=cp04 AND " & _
'                   "CP27 IS NULL and CP57 IS NULL and cp10 in (" + CNULL(發明申請) + "," + CNULL(新型申請) + "," + CNULL(設計申請) + "," + CNULL(追加申請) + "," + CNULL(聯合申請) + "," + CNULL(翻譯) + "," + CNULL(PCT申請) + "," + CNULL(記錄請求_標準專利) + "," + CNULL(短期專利申請) + ")" & _
'                   " and cp14=st01(+) ORDER BY cm01,cm02,cm03,CM04"
                strExc(0) = "select cm01||cm02||cm03||cm04,ST02,nvl(pa05,nvl(pa06,pa07)) FROM " & _
                   "CASEMAP,CASEPROGRESS,PATENT,STAFF WHERE CM05='" & pa(1) & "' AND CM06='" & pa(2) & "' AND " & _
                   "CM07='" & pa(3) & "' AND CM08='" & pa(4) & "' AND CM10='0' AND " & _
                   "cm01=pa01 and cm02=pa02 and cm03=pa03 and cm04=pa04 AND " & _
                   "cm01=cp01 and cm02=cp02 and cm03=cp03 and cm04=cp04 AND " & _
                   "CP27 IS NULL and CP57 IS NULL and cp10 in (" & CaseMapOut & ")" & _
                   " and cp14=st01(+) ORDER BY cm01,cm02,cm03,CM04"
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                   frm040104_1_1.Show vbModal
                End If
            End If
            
            'Add By Sindy 2013/11/14
            '檢查是否有承辦歷程是否有產生承辦單可以發文
            If PUB_IsEmpFlowIsSend(strCP09) = False Then
               Exit Sub
            End If
            
            Select Case pa(10)
               Case 延期
                  'Modify by Morgan 2009/12/23 延期發文必須要有法限
                  strExc(0) = "select 1 from caseprogress where cp09='" & strCP09 & "' and cp07>0"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     Me.Tag = strCP09 & "1"
                     '延期記錄之資料來源為下一程序檔
                     'Moddify by Amy 2018/09/14
                     Set frmNext = frm040104_2
                     bNext = True
                     frmNext.m_str_DL05 = "2"
                     'frm040104_2.Show
                     'end 2018/09/11
                  Else
                     MsgBox "延期發文必須要有法定期限，請重新分案！"
                     Exit Sub
                  End If
                  
               'Modified by Morgan 2013/8/29 +專屬授權 709
               Case 授權, 709
                  'Added by Morgan 2013/8/30
                  '台灣專利授權(專屬授權)檢查
                  If pa(1) = "P" And MainPa9 = "000" Then
                     '授權
                     If pa(10) = "704" Then
                        strExc(0) = "select * from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='709' and cp57 is null and cp54>=" & strSrvDate(1) & _
                           " and not exists(select * from caseprogress b where b.cp43=a.cp09 and b.cp10='705' and b.cp57 is null)"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If MsgBox("本案已辦理專屬授權，無法再辦理授權!!" & vbCrLf & "是否要繼續發文?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                              Exit Sub
                           End If
                        End If
                     '專屬授權
                     ElseIf pa(10) = "709" Then
                        strExc(0) = "select * from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in ('704','709') and cp57 is null and cp54>=" & strSrvDate(1) & _
                        " and not exists(select * from caseprogress b where b.cp43=a.cp09 and b.cp10='705' and b.cp57 is null)"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If RsTemp("cp10") = "704" Then
                              If MsgBox("本案授權期間尚未期滿，欲辦理專屬授權須待先前授權案期滿或撤回授權案，請與申請人確認!!" & vbCrLf & "是否要繼續發文?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                                 Exit Sub
                              End If
                           Else
                              If MsgBox("本案已辦理專屬授權，無法再辦理專屬授權!!" & vbCrLf & "是否要繼續發文?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                                 Exit Sub
                              End If
                           End If
                        End If
                     
                     End If
                  End If
                  'end 2013/8/30
                  
                  'Added by Morgan 2021/12/14
                  '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                  If PUB_CheckFormExist("frm040104_6") = False Then
                     Set frm040104_6 = Nothing
                  End If
                  'end 2021/12/14
                  
                  'Moddify by Amy 2018/09/14
                  'frm040104_6.Show
                  Set frmNext = frm040104_6
                  bNext = True
                  'end 2018/09/14
               Case 領證及繳年費
                     'Added by Morgan 2013/7/29
                     If PUB_CheckFormExist("frm040104_i") Then
                        MsgBox "【內專發文-領證及繳年費】不可與【" & frm040104_i.Caption & "】同時執行！", vbExclamation
                        Exit Sub
                     End If
                     'end 2013/7/29
         
                    'Modify by Morgan 2004/3/19
                    '加基本檔核准檢查
                    'frm040104_7.Show
                    If PUB_ApproveCheck(strCP09) Then
                    
                        'Added by Morgan 2021/12/14
                        '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                        If PUB_CheckFormExist("frm040104_7") = False Then
                           Set frm040104_7 = Nothing
                        End If
                        'end 2021/12/14
                        
                        'Moddify by Amy 2018/09/14
                        'frm040104_7.Show
                        Set frmNext = frm040104_7
                        bNext = True
                        'end 2018/09/14
                    Else
                        Exit Sub
                    End If
               'Add by Morgan 2007/8/10 加延展費
               'Modify by Amy 2018/03/20 +612 年費移作次年
               Case 年費, 維持費, 延展費, "612"
                     'Added by Morgan 2013/7/29
                     If PUB_CheckFormExist("frm040104_i") Then
                        MsgBox "【內專發文-年費、維持費、延展費】不可與【" & frm040104_i.Caption & "】同時執行！", vbExclamation
                        Exit Sub
                     End If
                     'end 2013/7/29
                     
                    'Add by Morgan 2004/6/29
                    If MainPa9 = "000" Then
                        'Modified by Morgan 2021/5/6
                        'If PUB_ChkCPExist(pa, 減免退費, 1) = True Then
                        '   If MsgBox("本案有【減免退費】未發文，若要同時發文請改選【減免退費】發文！確定只發文【年費】？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                        '      Exit Sub
                        '   End If
                        'End If
                        'end 2021/5/6
                    End If
                    
                     'Added by Morgan 2021/12/15
                     '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                     If PUB_CheckFormExist("frm040104_a") = False Then
                        Set frm040104_a = Nothing
                     End If
                     'end 2021/12/15
                     
                    'Modify by Morgan 2004/3/19
                    '加基本檔核准檢查
                    'frm040104_a.Show
                    '94.1.25 MODIFY BY SONIA 還原
                    'If PUB_ApproveCheck(strCP09) Then
                    '    frm040104_a.Show
                    'Else
                    '    Exit Sub
                    'End If
                    'Moddify by Amy 2018/09/14
                    'frm040104_a.Show
                    Set frmNext = frm040104_a
                    bNext = True
                    'end 2018/09/14
                    '94.1.25 END
               Case 設定質權, 終止設定質權
                  'Added by Morgan 2021/12/15
                  '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                  If PUB_CheckFormExist("frm040104_8") = False Then
                     Set frm040104_8 = Nothing
                  End If
                  'end 2021/12/15
                  'Moddify by Amy 2018/09/14
                  'frm040104_8.Show
                  Set frmNext = frm040104_8
                  bNext = True
                  'end 2018/09/14
               Case 異議_專, 舉發
                  'Added by Morgan 2021/12/15
                  '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                  If PUB_CheckFormExist("frm040104_9") = False Then
                     Set frm040104_9 = Nothing
                  End If
                  'end 2021/12/15
                  'Moddify by Amy 2018/09/14
                  'frm040104_9.Show
                  Set frmNext = frm040104_9
                  bNext = True
                  'end 2018/09/14
               'Modified by Morgan 2023/2/20 合併專利權讓與,+合併,繼承
               'Case 讓與
               Case 讓與, 專利權讓與, 合併, 繼承
                  'Added by Morgan 2021/12/15
                  '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                  If PUB_CheckFormExist("frm040104_c") = False Then
                     Set frm040104_c = Nothing
                  End If
                  'end 2021/12/15
                  'Moddify by Amy 2018/09/14
                  'frm040104_c.Show
                   Set frmNext = frm040104_c
                  bNext = True
                  'end 2018/09/14
                  
'Removed by Morgan 2023/2/20 併入讓與
'               'Add By Cheng 2002/01/11
'               Case 專利權讓與
'                  'Added by Morgan 2021/12/15
'                  '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
'                  If PUB_CheckFormExist("frm040104_c") = False Then
'                     Set frm040104_c = Nothing
'                  End If
'                  'end 2021/12/15
'                  'Moddify by Amy 2018/09/14
'                  'frm040104_c.Show
'                  Set frmNext = frm040104_c
'                  bNext = True
'                  'end 2018/09/14
'end 2023/2/20

               Case 減免退費
                  'Added by Morgan 2021/5/6
                  If PUB_ChkCPExist(pa, 年費, 1) = True Then
                     MsgBox "本案有【年費】未發文，改選【年費】發文！", vbInformation
                     Exit Sub
                  End If
                  'end 2021/5/6
                  
                  'Added by Morgan 2021/12/15
                  '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                  If PUB_CheckFormExist("frm040104_e") = False Then
                     Set frm040104_e = Nothing
                  End If
                  'end 2021/12/15
                  
                  'Moddify by Amy 2018/09/14
                  'frm040104_e.Show
                  Set frmNext = frm040104_e
                  bNext = True
                  'end 2018/09/14
               'Add by Morgan 2005/1/4
               Case 延緩公告
                  If PUB_ChkCPExist(pa, 領證及繳年費, 1) = True Then
                     MsgBox "本案有【領證及繳年費】未發文，請改選【領證及繳年費】發文！", vbExclamation, "延緩公告提醒"
                     Exit Sub
                  End If
                  
                  'Added by Morgan 2021/12/15
                  '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                  If PUB_CheckFormExist("frm040104_g") = False Then
                     Set frm040104_g = Nothing
                  End If
                  'end 2021/12/15
   
                  'Moddify by Amy 2018/09/14
                  'frm040104_g.Show
                  Set frmNext = frm040104_g
                  bNext = True
                  'end 2018/09/14
               Case Else
                  'Modify by Morgan 2004/6/3 台灣改請案,延緩公告檢查
                  'frm040104_3.Show
                  If MainPa9 = "000" Then
                     Select Case Val(pa(10))
                        Case 301 To 306
                           If PUB_Check301(pa(1), pa(2), pa(3), pa(4)) = False Then Exit Sub
                           'Add by Morgan 2006/4/25
                           If PUB_ChkCPExist(pa, 其他, 1) = True Then
                              MsgBox "本案有【其他】未發文，請先選【其他】發文！"
                              Exit Sub
                           End If
                           '2006/4/25 end
                        'Remove by Morgan 2005/1/4 改為可單獨發文但用新form
'                        Case 412
'                           MsgBox "延緩公告不可單獨發文！", vbInformation: Exit Sub
                           
                        'Modify by Morgan 2007/8/30 加第三人申請技術報告807
                        Case 421, 807
                           If PUB_Check421(pa(1), pa(2), pa(3), pa(4)) = False Then
                              Exit Sub
                           End If
                     End Select
                     'Add by Morgan 2004/8/16 台灣技術報告要印申請書故改用新form
                     Select Case Val(pa(10))
                     'Modify by Morgan 2007/8/30 加第三人申請技術報告807
                     Case 421, 807
                        'Added by Morgan 2021/12/15
                        '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                        If PUB_CheckFormExist("frm040104_f") = False Then
                           Set frm040104_f = Nothing
                        End If
                        'end 2021/12/15
                  
                        'Moddify by Amy 2018/09/14
                        'frm040104_f.Show
                        Set frmNext = frm040104_f
                        bNext = True
                        'end 2018/09/14
                     Case 928    '2007/6/28 add by sonia 加重新委任
                        'Added by Morgan 2021/12/15
                        '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                        If PUB_CheckFormExist("frm040104_h") = False Then
                           Set frm040104_h = Nothing
                        End If
                        'end 2021/12/15
                        'Moddify by Amy 2018/09/14
                        'frm040104_h.Show
                        Set frmNext = frm040104_h
                        bNext = True
                        'end 2018/09/14
                     Case Else
                        'Add by Amy 2014/04/08 因應台日優先權證明文件控管 2014/04/14 +判斷不為PCT案且無分割案
                        If InStr("1,2", pa(8)) > 0 And (pa(10) = "106" Or pa(10) = "124") And pa(7) = "" And PUB_ChkCPExist(pa, "307") = False Then
                            'Mark by Amy 2015/04/02 案發文會有問題,變數未清
                            'If strChoose = "" Then
                                If ChkPD09Null(pa(1), pa(2), pa(3), pa(4)) = True Then
                                    frm040104_3_1.Show vbModal
                                    strChoose = strPublicTemp
                                    strPublicTemp = ""
                                    If strChoose = "" Then Exit Sub
                                End If
                            'End If
                        End If
                        'end 2014/04/08
                        
                        'Added by Morgan 2021/12/14
                        '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                        If PUB_CheckFormExist("frm040104_3") = False Then
                           Set frm040104_3 = Nothing
                        End If
                        'end 2021/12/14
   
                        'Moddify by Amy 2018/09/14
                        'frm040104_3.Show
                        Set frmNext = frm040104_3
                        bNext = True
                        'end 2018/09/14
                     End Select
                     
                  Else
                     'Add by Amy 2014/04/08 因應台日優先權證明文件控管 2014/04/14 +判斷不為PCT案且無分割案
                     If MainPa9 = "020" And InStr("1,2", pa(8)) > 0 And (pa(10) = "106" Or pa(10) = "124") And pa(7) = "" And PUB_ChkCPExist(pa, "307") = False Then
                        'Mark by Amy 2015/04/02 案發文會有問題,變數未清
                        'If strChoose = "" Then
                            If ChkPD09Null(pa(1), pa(2), pa(3), pa(4)) = True Then
                                frm040104_3_1.Show vbModal
                                strChoose = strPublicTemp
                                strPublicTemp = ""
                                If strChoose = "" Then Exit Sub
                            End If
                        'End If
                     End If
                     'end 2014/04/08
                     
                     'Added by Morgan 2021/12/14
                     '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
                     If PUB_CheckFormExist("frm040104_3") = False Then
                        Set frm040104_3 = Nothing
                     End If
                     'end 2021/12/14
                     
                     'Moddify by Amy 2018/09/14
                     'frm040104_3.Show
                     Set frmNext = frm040104_3
                     bNext = True
                     'end 2018/09/14
                  End If
                  
            End Select
            Command1.SetFocus
            Me.Hide
            'Modify By Cheng 2002/01/24
            '若卷宗性質為申請(1)
            If strPA23 = "1" Then
               ' 90.10.18 modify by louis (顯示專利基本檔)
               '91.11.27 MODIFY BY SONIA
               'ShowMaintainForm strCP09
               '91.12.9 加 提早公開 91.12.24加 調卷 92.1.16加 回覆代理人
               'Modify by Morgan 2004/5/14
               '加 117
               'If pa(10) <> 實體審查 And pa(10) <> 翻譯 And pa(10) <> 其他 And pa(10) <> 專利調查 And pa(10) <> 補收款 And pa(10) <> 主動修正 And pa(10) <> 主張優先權 And pa(10) <> 提早公開 And pa(10) <> 調卷 And pa(10) <> 回覆代理人 Then
               '2005/4/6 MODIFY BY SONIA 加 急件費920,超項費917
               'If pa(10) <> "117" And pa(10) <> 實體審查 And pa(10) <> 翻譯 And pa(10) <> 其他 And pa(10) <> 專利調查 And pa(10) <> 補收款 And pa(10) <> 主動修正 And pa(10) <> 主張優先權 And pa(10) <> 提早公開 And pa(10) <> 調卷 And pa(10) <> 回覆代理人 Then
               'Modify by Morgan 2007/12/14 加121
               'Modify by Morgan 2009/12/23 +936,404
               '2010/1/6 modify by sonia 加938超頁費,939超項費
               If pa(10) <> "117" And pa(10) <> 實體審查 And pa(10) <> 翻譯 And pa(10) <> 其他 And pa(10) <> 專利調查 And pa(10) <> 補收款 And pa(10) <> 主動修正 And pa(10) <> 主張優先權 And pa(10) <> 提早公開 And pa(10) <> 調卷 And pa(10) <> 回覆代理人 And pa(10) <> 920 And pa(10) <> 917 And pa(10) <> 938 And pa(10) <> 939 And pa(10) <> "121" And pa(10) <> "936" And pa(10) <> "404" Then
                  ShowMaintainForm strCP09
               End If
               '91.11.27 END
            End If
'Remove by Morgan 2005/1/4 不必分本所案號或總收文號
'         Else
'            Command1_Click '收文號
'         End If
'2005/1/4
            'Add by Amy 2018/09/14 +新案號 cp31=Y顯示申請人地址讓user修改
            If bNext = True Then
                Me.Hide
                frmNext.Show
                If strCP31 = "Y" Then
                   frm020102_23.Hide
                   Set frm020102_23.UpForm = frmNext
                   frm020102_23.m_CP09 = strCP09
                   frm020102_23.QueryData
                   frm020102_23.Show vbModal
                End If
            End If
            'end 2018/09/14
      Case 2 '離開
         Unload Me
   End Select
End Sub

Public Sub Command1_Click()
   Dim ii As Integer
 
   'Modify by Morgan 2005/1/5 為避免與選本所案號的程式不一致改用相同程式
   '選擇本所案號
   'If Option1(0).Value = True Then
   If Option1(1).Value = True Then
      'Modified by Lydia 2014/12/9 P非新案發文改以申請案號輸入,原收文號取消
      'strExc(0) = "select cp01,cp02,cp03,cp04 from caseprogress where cp09='" & Text5 & "'"
      strExc(0) = "select pa01,pa02,pa03,pa04 from patent where pa11='" & Text5 & "'" & _
                  " union select sp01,sp02,sp03,sp04 from servicepractice where sp11='" & Text5 & "'"
      'Added by Morgan 2017/1/19
      'Modify By Sindy 2018/12/21 + and CP158=0
      strExc(0) = strExc(0) & " union select cp01,cp02,cp03,cp04 from caseprogress where cp36='" & Text5 & "' and CP158=0"
      intI = 0
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
'         Text1 = "" & RsTemp.Fields("cp01")
'         Text2 = "" & RsTemp.Fields("cp02")
'         Text3 = "" & RsTemp.Fields("cp03")
'         Text4 = "" & RsTemp.Fields("cp04")
         Text1 = "" & RsTemp.Fields("pa01")
         Text2 = "" & RsTemp.Fields("pa02")
         Text3 = "" & RsTemp.Fields("pa03")
         Text4 = "" & RsTemp.Fields("pa04")
      Else
'         MsgBox "無該收文號資料！", vbExclamation, "發文"
         Exit Sub
       'end 'Modified by Lydia 2014/12/9
      End If
   End If
   '2005/1/5 end
 
      If Text3 = "" Then Text3 = "0"
      If Text4 = "" Then Text4 = "00"
    
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
        If FMP2open = True Then
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text1, Text2, Text3, Text4) = False Then
             ' Text1 = "P": Text2 = "": Text3 = "": Text4 = "": Text5 = "" '無權限清空
              If Option1(0).Value = True Then Text2.SetFocus
              If Option1(1).Value = True Then Text5.SetFocus
              Exit Sub
           End If
        End If
        
      pa(1) = Text1
      pa(2) = Text2
      pa(3) = Text3
      pa(4) = Text4
      
      Combo1.Clear
      'Modified by Lydia 2014/12/9 P非新案發文改以申請案號輸入,原收文號取消 + 收文號 (+pa11,sp11)
      mPA11 = ""
      If Text1 = "P" Then
         'Modify by Amy 2014/04/08 因應台日優先權證明文件判斷+pa08 2014/04/14 +pa46
         strExc(0) = "SELECT PA05,PA06,PA07,PA09,PA23,PA14,PA08,PA46,PA11 FROM PATENT WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      ElseIf Text1 = "PS" Then
         strExc(0) = "SELECT SP05,SP06,SP07,SP09,SP11 as PA11 FROM SERVICEPRACTICE WHERE " & ChgService(pa(1) & pa(2) & pa(3) & pa(4))
      End If
      intI = 1
      strExc(1) = "CPM03,"
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
            Combo1.AddItem "中 : " & .Fields(0)
            Combo1.AddItem "英 : " & .Fields(1)
            Combo1.AddItem "日 : " & .Fields(2)
            'Modified by Morgan 2017/1/19 舉發案可能會有同一申請號對兩個本所案，改放對照號數　Ex.P113312, P108253
            'mPA11 = "" & RsTemp!PA11
            If Option1(1).Value = True Then
               mPA11 = Text5
            Else
               mPA11 = "" & RsTemp!PA11
            End If
            'end 2017/1/19
            
            Combo1.ListIndex = 0
            If IsNull(.Fields(3)) = False Then
               MainPa9 = .Fields(3)
               If .Fields(3) = 台灣國家代號 Then
                  strExc(1) = "CPM03,"
               Else
                  strExc(1) = "CPM04,"
               End If
            End If
            If Text1 = "P" Then pa(7) = "" & .Fields("PA46"): pa(8) = "" & .Fields("PA08") 'Modify by Amy 2013/04/14
         End With
      End If
      'Modify By Cheng 2002/04/12
'      strExc(0) = "select ''," & SQLDate("CP05") & ",cp09," & strExc(1) & "staff.st02 as st1," & _
'         "staff1.st02 as st2,cp64,cp10,cp12,cp13,CP73 from caseprogress, casepropertymap," & _
'         "staff,staff staff1 where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " AND CP27 IS NULL AND CP57 IS NULL AND (SUBSTR(CP09,1,1)='A' OR SUBSTR(CP09,1,1)='B')" & _
'         " AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)"
        'Modify By Cheng 2002/11/28
        '加本所期限及相關人
'      strExc(0) = "select ''," & SQLDate("CP05") & ",cp09," & strExc(1) & "staff.st02 as st1," & _
'         "staff1.st02 as st2,cp64,cp10,cp12,cp13,CP79 from caseprogress, casepropertymap," & _
'         "staff,staff staff1 where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " AND CP27 IS NULL AND CP57 IS NULL AND ( CP09<'C' )" & _
'         " AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)"
      'Modify by Morgan 2009/10/2 + 改C類也可發文
      'MODIFY BY SONIA 2014/5/13 +cp43
      'Modify by Amy 2015/01/22 +cp157
      'modify by sonia 2017/9/27 +ORDER BY CP05,CP09
      'Modify by Amy 2018/09/14 +CP31
      'Modified by Morgan 2019/9/11 +CP14_ST03
      'Modified by Morgan 2020/1/16 +cp09條件
      strExc(0) = "select ''," & SQLDate("CP05") & ",cp09," & strExc(1) & "staff.st02 as st1," & _
         "staff1.st02 as st2,Decode(CP06,Null,Null,CP06 - 19110000),DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
         "cp64,cp10,cp12,cp13,CP79,CP43,CP157,CP31,staff.st03 CP14_ST03  from caseprogress, casepropertymap," & _
         "staff,staff staff1,Customer where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " AND CP27 IS NULL AND CP57 IS NULL" & IIf(mCP09 <> "", " and cp09='" & mCP09 & "'", "") & _
         " AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) ORDER BY CP05,CP09 "
   
      'Add by Morgan 2005/1/5
      'Modified by Lydia 2014/12/9 P非新案發文改以申請案號輸入,原收文號取消
      'If Option1(1).Value = True Then strExc(0) = strExc(0) & " and cp09='" & Text5 & "'"
   
      intI = 0
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      Set MSHFlexGrid1.Recordset = RsTemp
      'ADD BY SONIA 2014/5/13 加相關總收號案件性質
      For ii = 1 To Me.MSHFlexGrid1.Rows - 1
          Me.MSHFlexGrid1.TextMatrix(ii, 3) = Me.MSHFlexGrid1.TextMatrix(ii, 3) & PUB_GetRelateCasePropertyName(Me.MSHFlexGrid1.TextMatrix(ii, 2), "1")
      Next ii
      'END 2014/5/13
      GridHead
      
      'Added by Morgan 2012/8/13 102新法
      '發明,新型申請檢查是否有一案兩請,提醒是否要同時送件(大陸案一併檢查)
      If RsTemp.RecordCount > 0 Then
         With RsTemp
         Do While Not .EOF
            If .Fields("cp10") = "101" Or .Fields("cp10") = "102" Then
               strExc(0) = "select cp01||'-'||cp02 from caseprogress where (cp01,cp02,cp03,cp04) in (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "'" & _
                  " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & pa(1) & "' and cm06='" & pa(2) & "' and cm07='" & pa(3) & "' and cm08='" & pa(4) & "') and cp10 in ('101','102') and cp57||cp27 is null"
               intI = 1
               Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "本案與 " & adoRecordset(0) & " 案為一案兩請，請檢查二案是否同時送件！", vbExclamation
               End If
               Exit Do
            End If
            .MoveNext
         Loop
         
         'Added by Morgan 2015/9/14
         '擬制喪失新穎性同時發文檢查
         .MoveFirst
         Do While Not .EOF
            If .Fields("cp10") = "101" Or .Fields("cp10") = "102" Or .Fields("cp10") = "103" Then
               strExc(0) = "select cp01||'-'||cp02 from caseprogress where (cp01,cp02,cp03,cp04) in (select cm05,cm06,cm07,cm08 from casemap where cm10='6' and cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "'" & _
                  " union select cm01,cm02,cm03,cm04 from casemap where cm10='6' and cm05='" & pa(1) & "' and cm06='" & pa(2) & "' and cm07='" & pa(3) & "' and cm08='" & pa(4) & "') and cp10 in ('101','102','103') and cp57||cp27 is null"
               intI = 1
               Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "本案與 " & adoRecordset(0) & " 案為擬制喪失新穎性關聯，請檢查二案是否同時送件！", vbExclamation
               End If
               Exit Do
            End If
            .MoveNext
         Loop
         'end 2015/9/14
         
         End With
      End If
      'end 2012/8/13
      
      'Add By Cheng 2002/05/10
      '若只搜尋到一筆時直接勾選
      If Me.MSHFlexGrid1.Rows = 2 Then
        Me.MSHFlexGrid1.row = 1
        GridClick MSHFlexGrid1, 1, 0
        cmdok(1).SetFocus
        If mCP09 <> "" Then cmdok(1).Value = True 'Added by Morgan 2020/1/16
        
        'Add by Morgan 2005/1/5
        'Modified by Lydia 2014/12/9 P非新案發文改以申請案號輸入,原收文號取消
        'If Option1(1).Value = True Then cmdOK_Click 1
        'Debug.Print Now
      End If
   
'Remove by Morgan 2005/1/5 為避免與選本所案號的程式不一致改用相同程式
'
'   '選擇收文號
'   Else
'      GridHead
'      'Modify By Cheng 2002/04/12
''      strExc(0) = "select CP10,CP12,CP13,CP73,PA09,CP01,CP02,CP03,CP04,PA14 from caseprogress,PATENT where cp09='" & Text5 & "'" & _
''         " AND CP27 IS NULL AND CP57 IS NULL AND (SUBSTR(CP09,1,1)='A' OR SUBSTR(CP09,1,1)='B')" & _
''         " AND (CP01='P' OR CP01='PS') AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04"
'      strExc(0) = "select CP10,CP12,CP13,CP79,PA09,CP01,CP02,CP03,CP04,PA14 from caseprogress,PATENT where cp09='" & Text5 & "'" & _
'         " AND CP27 IS NULL AND CP57 IS NULL AND ( CP09<'C' )" & _
'         " AND (CP01='P' OR CP01='PS') AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04"
'      intI = 1
'      Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'
'      If intI = 1 Then
'
'         With rsTemp
'            Me.Tag = Text5
'            If Not IsNull(.Fields(5)) Then Text1.Text = .Fields(5)
'            If Not IsNull(.Fields(6)) Then Text2.Text = .Fields(6)
'            If Not IsNull(.Fields(7)) Then Text3.Text = .Fields(7)
'            If Not IsNull(.Fields(8)) Then Text4.Text = .Fields(8)
'            If Not IsNull(.Fields(4)) Then MainPa9 = .Fields(4)
'            'Add By Cheng 2003/02/18
'            '若案年性質為領證及繳年費
'            If "" & rsTemp("CP01").Value = "P" And "" & rsTemp("CP10") = 領證及繳年費 Then
'                '若有公告日
'                If Not IsNull(rsTemp.Fields("PA14").Value) Then
'                    '系統日 >= 公告日 +  三個月 + 8 天
'                    If ServerDate < DBDATE(DateAdd("D", 8, DateAdd("M", 3, ChangeWStringToWDateString(rsTemp.Fields("PA14").Value)))) Then
'                        MsgBox "公告尚未期滿，不可發文 !", vbCritical
'                        Exit Sub
'                    End If
'                End If
'            End If
'            'Add By Cheng 2002/07/12
'            '已閉卷案件不可發文
'            If PUB_CaseClosed(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text) Then
''               Me.MSHFlexGrid1.Clear
''               GridHead
''               Me.MSHFlexGrid1.Rows = 2
''               Me.MSHFlexGrid1.FixedRows = 1
'               Exit Sub
'            End If
'             'Add By Cheng 2003/09/01
'             '若為領證或年費
'            If "" & rsTemp("CP10").Value = "601" Or "" & rsTemp("CP10").Value = "605" Then
'                If PUB_ChkNP605(Me.Text1.Text & Me.Text2.Text & Me.Text3.Text & Me.Text4.Text) Then
'                    MsgBox "本案下一程序有<年費>期限，不可發文!!!", vbExclamation + vbOKOnly
''                    Me.MSHFlexGrid1.Clear
''                    GridHead
''                    Me.MSHFlexGrid1.Rows = 2
''                    Me.MSHFlexGrid1.FixedRows = 1
'                    Exit Sub
'                End If
'            End If
'            'Add By Cheng 2003/10/20
'            '未輸入承辦人不可發文
'            If PUB_ChkCP14IsNull(Me.Text5.Text) = True Then Exit Sub
'            'Add By Cheng 2003/06/16
'            If "" & .Fields(0).Value = 發明申請 Or "" & .Fields(0).Value = 新型申請 Or "" & .Fields(0).Value = 設計申請 Or "" & .Fields(0).Value = 追加申請 Or "" & .Fields(0).Value = 聯合申請 Then
'                strExc(0) = "select cm01||cm02||cm03||cm04,ST02,nvl(pa05,nvl(pa06,pa07)) FROM " & _
'                   "CASEMAP,CASEPROGRESS,PATENT,STAFF WHERE CM05='" & .Fields(5).Value & "' AND CM06='" & .Fields(6).Value & "' AND " & _
'                   "CM07='" & .Fields(7).Value & "' AND CM08='" & .Fields(8).Value & "' AND CM10='0' AND " & _
'                   "cm01=pa01 and cm02=pa02 and cm03=pa03 and cm04=pa04 AND " & _
'                   "cm01=cp01 and cm02=cp02 and cm03=cp03 and cm04=cp04 AND " & _
'                   "CP27 IS NULL and CP57 IS NULL and cp10 in (" + CNULL(發明申請) + "," + CNULL(新型申請) + "," + CNULL(設計申請) + "," + CNULL(追加申請) + "," + CNULL(聯合申請) + "," + CNULL(翻譯) + ")" & _
'                   " and cp14=st01(+) ORDER BY cm01,cm02,cm03,CM04"
'                intI = 1
'                Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'                If intI = 1 Then
'                   frm040104_1_1.Show vbModal
'                End If
'            End If
'            '91.10.26 modify by sonia
'            'If IsNull(.Fields(3)) Or .Fields(3) = 0 Then
'            '
'            'Else
'            '   If Not ChkShowMail Then Exit Sub
'            'End If
'            '91.10.26 end
'            Select Case .Fields(0)
'               Case 延期
'                  Me.Tag = Me.Tag & "1"
'                  'Add By Cheng 2002/06/20
'                  '延期記錄之資料來源為下一程序檔
'                  frm040104_2.m_str_DL05 = "2"
'                  frm040104_2.Show
'               Case 授權
'                  frm040104_6.Show
'               Case 領證及繳年費
'                  If Text1 = "P" Then
'                     If Not IsNull(rsTemp.Fields(9)) Then
'                        If Val(strSrvDate(1)) < Val(CompDate(2, 7, .Fields(9))) Then
'                           MsgBox "系統日必須大於等於公告日(" & TransDate(rsTemp.Fields(9), 1) & ") + 7天 !", vbCritical
'                           Exit Sub
'                        End If
'                     End If
'                  End If
'                    'Modify by Morgan 2004/3/19
'                    '加基本檔核准檢查
'                    'frm040104_7.Show
'                    If PUB_ApproveCheck(Text5) Then
'                        frm040104_7.Show
'                    Else
'                        Exit Sub
'                    End If
'               Case 設定質權, 終止設定質權
'                  frm040104_8.Show
'               Case 異議_專, 舉發
'                  frm040104_9.Show
'               Case 年費
'                    'Modify by Morgan 2004/3/19
'                    '加基本檔核准檢查
'                    'frm040104_a.Show
'                    If PUB_ApproveCheck(Text5) Then
'                        frm040104_a.Show
'                    Else
'                        Exit Sub
'                    End If
'               Case 讓與
'                  frm040104_c.Show
'               'Add By Cheng 2002/01/11
'               Case 專利權讓與
'                  frm040104_c.Show
'               Case Else
'                  'Add by Morgan 2004/6/3 台灣改請案,延緩公告檢查
'                  If ("" & .Fields("PA09")) = "000" Then
'                     Select Case Val("" & .Fields("CP10"))
'                        Case 301 To 306
'                           If PUB_Check301(.Fields(5), .Fields(6), .Fields(7), .Fields(8)) = False Then Exit Sub
'                        Case 421
'                           MsgBox "延緩公告不可單獨發文！", vbInformation: Exit Sub
'                     End Select
'                  End If
'                  'Add End
'                  frm040104_3.Show
'            End Select
'            Me.Hide
'         End With
'      Else
'         MsgBox "無符合發文條件之資料 !", vbCritical
'      End If
'   End If
End Sub

Private Function ChkShowMail() As Boolean
   frm040104_1_2.Show vbModal
   If bolLeave = True Then
      ChkShowMail = False
   Else
      ChkShowMail = True
   End If
End Function

Private Sub Form_Activate()
   Dim i As Integer, j As Integer
   
   'Added by Morgan 2020/1/16
   If m_bolActivated = True And Me.mCP09 <> "" Then
      mPreForm.QueryData
      Unload Me
      Exit Sub
   End If
   'end 2020/1/16
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = ""
        If .CellBackColor = &HFFC0C0 Then
            For j = 0 To .Cols - 1
               .col = j
               .CellBackColor = .BackColor
            Next
         End If
      Next
   End With
   
   'Add by Morgan 2010/4/1 改預設本所號
   If Not m_bolActivated Then
      m_bolActivated = True
      Option1(0).Value = True
   End If
End Sub

Private Sub Form_Load()
' On Error Resume Next
   MoveFormToCenter Me
   intWhere = 國內
   'Combo1.ListIndex = 0 'Removed by Morgan 2021/12/14
   Text1.Enabled = False
   Text2.Enabled = False
   Text3.Enabled = False
   Text4.Enabled = False
    'Modify By Cheng 2002/11/28
'   InitGrid 11, MSHFlexGrid1
   'Modify by Amy 2015/01/22 原:13
   InitGrid 15, MSHFlexGrid1
   GridHead
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   bolIsEMPFlow = False 'Add By Sindy 2013/5/20
   Set frm040104_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdok(1).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
 On Error Resume Next
   Select Case Index
      Case 0
         Text1.Enabled = True
         Text2.Enabled = True
         Text3.Enabled = True
         Text4.Enabled = True
         Text5.Enabled = False
        'Modify By Cheng 2002/11/04
'         Text1.SetFocus
        Me.Text2.SetFocus
      Case 1
         Text1.Enabled = False
         Text2.Enabled = False
         Text3.Enabled = False
         Text4.Enabled = False
         Text5.Enabled = True
         Text5.SetFocus
   End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "P" And Text1 <> "PS" And Text1 <> "" Then
      MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
      TextInverse Text1
      Cancel = True
   End If
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
        'Modify By Cheng 2002/11/28
'      .Col = 1: .ColWidth(1) = 1200: .Text = "收文日"
      .col = 1: .ColWidth(1) = 900: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
        'Modify By Cheng 2002/11/28
'      .Col = 4: .ColWidth(4) = 1200: .Text = "承辦人"
      .col = 4: .ColWidth(4) = 900: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
        'Modify By Cheng 2002/11/28
'      .Col = 5: .ColWidth(5) = 1400: .Text = "智權人員"
      .col = 5: .ColWidth(5) = 900: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
    'Add By Cheng 2002/11/28
      .col = 6: .ColWidth(6) = 900: .Text = "本所期限"
      .col = 7: .ColWidth(7) = 1400: .Text = "相關人"
      
      'Modify by Morgan 2008/1/3
      '.col = 8: .ColWidth(8) = 1400: .Text = "案件備註"
      .col = 8: .ColWidth(8) = 1400: .Text = "進度備註"
      'Modify by Amy 2015/02/03 隱藏cp43及cp157 欄位 原12
      For i = 9 To .Cols - 1
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Public Sub ReQuery()
    'Add By Cheng 2003/03/26
    If Me.Option1(1).Value Then
        Me.Option1(0).Value = True
    End If
    Command1_Click
End Sub

Public Sub Clear()
    'Modify By Cheng 2002/10/30
    ' 保留系統類別
'   Text1 = Empty
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   Text5 = Empty
   Option1(0).Value = False
   'Modify by Morgan 2010/4/16 改預設本所案號
   'Option1(1).Value = True
   Option1(0).Value = True
   
   Combo1.Clear
   'MSHFlexGrid1.Clear
   'MSHFlexGrid1.Rows = 1
    'Modify By Cheng 2002/11/28
'   InitGrid 11, MSHFlexGrid1
   InitGrid 13, MSHFlexGrid1
   GridHead
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text3_GotFocus()
   InverseTextBox Text3
End Sub

Private Sub Text4_GotFocus()
   InverseTextBox Text4
End Sub

Private Sub Text5_GotFocus()
   InverseTextBox Text5
End Sub

'Added by Morgan 2019/9/11
'以欄位名稱抓值
Private Function GetGridValue(pRow As Integer, pFieldName As String) As String
   Dim ii As Integer
   With MSHFlexGrid1
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pFieldName) Then
         GetGridValue = .TextMatrix(pRow, ii)
         Exit For
      End If
   Next
   End With
   
End Function
