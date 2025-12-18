VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050102_9 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文（繳年費、維持費、延展費、年費移作次年）"
   ClientHeight    =   5772
   ClientLeft      =   516
   ClientTop       =   1020
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   8520
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   3390
      MaxLength       =   4
      TabIndex        =   13
      Top             =   4365
      Width           =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "(美、加、法國案)"
      Height          =   525
      Left            =   60
      TabIndex        =   50
      Top             =   3180
      Width           =   4200
      Begin VB.OptionButton optChoose 
         Caption         =   "微個體"
         Height          =   300
         Index           =   2
         Left            =   2580
         TabIndex        =   53
         Top             =   180
         Width           =   1575
      End
      Begin VB.OptionButton optChoose 
         Caption         =   "小個體"
         Height          =   300
         Index           =   1
         Left            =   1215
         TabIndex        =   52
         Top             =   180
         Width           =   1365
      End
      Begin VB.OptionButton optChoose 
         Caption         =   "大個體"
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   51
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   2220
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   6120
      MaxLength       =   5
      TabIndex        =   7
      Top             =   2550
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   4
      Top             =   2550
      Width           =   405
   End
   Begin VB.TextBox textNum 
      Height          =   270
      Left            =   4920
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2550
      Width           =   435
   End
   Begin VB.TextBox textYear 
      Height          =   270
      Left            =   720
      MaxLength       =   5
      TabIndex        =   3
      Top             =   2550
      Width           =   405
   End
   Begin VB.OptionButton optSel 
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   5
      Top             =   2610
      Width           =   252
   End
   Begin VB.OptionButton optSel 
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2610
      Value           =   -1  'True
      Width           =   252
   End
   Begin VB.CommandButton cmdCountry 
      Caption         =   "指定國家"
      Height          =   300
      Left            =   660
      TabIndex        =   12
      Top             =   4380
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7620
      TabIndex        =   20
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5580
      TabIndex        =   18
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6396
      TabIndex        =   19
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   8
      Left            =   1680
      TabIndex        =   9
      Top             =   3720
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   9
      Left            =   6120
      TabIndex        =   10
      Top             =   3690
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   4
      Left            =   6120
      TabIndex        =   8
      Top             =   2880
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   6
      Left            =   6120
      TabIndex        =   14
      Top             =   4350
      Width           =   1095
      VariousPropertyBits=   671107099
      MaxLength       =   9
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   5
      Left            =   1320
      TabIndex        =   11
      Top             =   4050
      Width           =   1035
      VariousPropertyBits=   671107099
      MaxLength       =   9
      Size            =   "1826;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   3
      Left            =   1680
      TabIndex        =   17
      Top             =   2880
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   1890
      Width           =   1095
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   810
      Index           =   7
      Left            =   90
      TabIndex        =   15
      Top             =   4950
      Width           =   8355
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14737;1429"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數："
      Height          =   180
      Index           =   18
      Left            =   2430
      TabIndex        =   54
      Top             =   4410
      Width           =   900
   End
   Begin VB.Label Label12 
      Caption         =   "是否列印通知函：            （N:不印）"
      Height          =   180
      Left            =   120
      TabIndex        =   49
      Top             =   3780
      Width           =   2895
   End
   Begin VB.Label Label19 
      Caption         =   "是否修改通知函內容：             （Y:Word）"
      Height          =   180
      Left            =   4290
      TabIndex        =   48
      Top             =   3765
      Width           =   3465
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "巳繳年費："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   47
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label lblCaseField 
      AutoSize        =   -1  'True
      Height          =   225
      Index           =   6
      Left            =   1080
      TabIndex        =   46
      Top             =   1620
      Width           =   7335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否修改指示信內容：            （Y:Word）"
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   45
      Top             =   2940
      Width           =   3270
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "第                次至第            次費用"
      Height          =   180
      Index           =   3
      Left            =   4620
      TabIndex        =   44
      Top             =   2610
      Width           =   2520
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "下次繳費期限："
      Height          =   180
      Left            =   4770
      TabIndex        =   43
      Top             =   4410
      Width           =   1260
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "EPC："
      Height          =   180
      Left            =   120
      TabIndex        =   42
      Top             =   4440
      Width           =   495
   End
   Begin MSForms.Label lblNotify 
      Height          =   255
      Left            =   2400
      TabIndex        =   41
      Top             =   4050
      Width           =   5940
      VariousPropertyBits=   27
      Size            =   "10477;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "年費通知人："
      Height          =   180
      Left            =   120
      TabIndex        =   40
      Top             =   4080
      Width           =   1080
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "第              年至第            年費用"
      Height          =   180
      Index           =   1
      Left            =   420
      TabIndex        =   16
      Top             =   2610
      Width           =   2430
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   2220
      TabIndex        =   39
      Top             =   2250
      Width           =   6135
      VariousPropertyBits=   27
      Size            =   "10821;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否列印指示信：            （N:不印）"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   38
      Top             =   2940
      Width           =   2865
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   120
      TabIndex        =   37
      Top             =   2250
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Top             =   1950
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   "進度備註："
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   35
      Top             =   4710
      Width           =   975
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Top             =   570
      Width           =   2535
   End
   Begin MSForms.Label lblSalesName 
      Height          =   255
      Left            =   6000
      TabIndex        =   27
      Top             =   930
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   26
      Top             =   1290
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   25
      Top             =   930
      Width           =   615
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   24
      Top             =   570
      Width           =   495
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   23
      Top             =   1290
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   22
      Top             =   930
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   21
      Top             =   570
      Width           =   3135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   34
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   32
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      Height          =   180
      Left            =   120
      TabIndex        =   30
      Top             =   570
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   29
      Top             =   1290
      Width           =   900
   End
End
Attribute VB_Name = "frm050102_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/6 改成Form2.0 (txtCaseField,lblSalesName...)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'2005/7/11整理
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'Modify By Cheng 2002/11/22
''intLeaveKind離開時，是0:結束  1:回上一畫面
'intLeaveKind離開時，是0:結束  1:回上一畫面  2:EPC指定國家子國皆閉卷, 回上一畫面
Dim intLeaveKind As Integer
'StrCountry存放指定國家  strMoneyCountry存放繳費國家 strMoney存放費用
Dim strCountry As String, strMoneyCountry As String, strMoney As String
'dobDateAdd存放專利年度,strCaseProperty存放下一程序之案件性質,以供存檔時取用
Dim dobDateAdd As Double, strCaseProperty As String
' 開始繳費日期
Dim m_StartDate As String
' 繳費年度字串資料 (year1,year2,year3,year4,...)
Dim m_FeeYears As String
'910815 Sieg
Dim varFeeYears As Variant
Dim varPA72 As Variant
'Add By Cheng 2002/10/02
Dim m_strCountryEngName As String '指定國家的英文名稱
'92.1.16 add by sonia
Dim old_Entity As String   '原大小個體
Dim new_Entity As String   '新大小個體
'2005/4/15 ADD BY SONIA
Dim old_cp44 As String     '原發文代理人
Dim m_iFixNo As Integer '修法第次
Dim strCP09List  As String 'Add by Amy 2013/08/30 子案指示信-子案總收文號
Dim strNCP09List As String 'Added by Lydia 2017/06/28 EPC未勾選國家上結案，萁出結案指示信
'Add By Sindy 2018/1/8
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2018/1/8 END
Dim m_strAF01 As String, m_strChildAF01 As String, m_strLD18 As String 'Added by Morgan 2018/8/22
Dim m_strCP81 As String, m_strJpMemo As String 'Added by Morgan 2019/5/2
Dim m_strEPC_DE_BCP09 As String 'Added by Morgan 2024/4/25 EPC德國子案年費B類收文號

Private Sub cmdCountry_Click()
    'Modify By Cheng 2003/09/22
    '與領證用相同的畫面
'   ModifyMoneyCountry strCountry, strMoneyCountry, strMoney
   'Modified by Morgan 2020/8/14 +"2"
   'Modified by Morgan 2023/3/7 改傳入案件性質
    ModifyLicenceCountry strCountry, strMoneyCountry, , cp(10)
End Sub

Private Sub cmdok_Click(Index As Integer)
 Dim stLetter As String 'Add by Morgan 2004/9/27
 Dim i As Integer, strTmp As String
'Add By Cheng 2002/07/31
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strPrintCP09() As String 'Add by Amy 2013/08/30 產生各國給各代的年費指示信-案總收文號
'Added by Lydia 2017/06/29
Dim tmpArr As Variant
Dim tmpCp10 As String 'Add by Amy 2018/03/20 指示信使用

   Select Case Index
      Case 0
      
         'Added by Morgan 2015/8/7
         If PUB_ChkFileNP(cp(9)) Then
            MsgBox "下一程序已有提申或收達期限，不可發文！"
            Exit Sub
         End If
         'end 2015/8/7
   
         ' 90.10.15 modify by louis
         If CheckDataValid() = False Then
            Exit Sub
         End If
         'Modify by Morgan 2006/12/25 不必判斷發證日但必須是准
         'If field(9) = EPC指定國家 And field(20) <> "" And field(21) <> "" And strMoneyCountry = "" Then
         'Modify by Morgan 2009/8/3 加判斷有公告日(已核准未公告仍繳給EPO Ex.CFP-18628)
         'If field(9) = EPC指定國家 And field(20) <> "" And field(16) = "1" And strMoneyCountry = "" Then
         'Modify by Morgan 2009/10/2 改判斷是否可點選否則控制會不一致
         'If field(9) = EPC指定國家 And field(14) <> "" And field(20) <> "" And field(16) = "1" And strMoneyCountry = "" Then
         If cmdCountry.Enabled = True And strMoneyCountry = "" Then
            MsgBox "未輸入繳費之指定國家 !", vbCritical
            Exit Sub
         End If
         '92.1.16 add by sonia
         'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
         If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
            'Modified by Morgan 2013/3/20 +微個體
            If Not OptChoose(0).Value And Not OptChoose(1).Value And Not OptChoose(2).Value Then
               If OptChoose(2).Enabled = True Then
                  MsgBox "請選擇" & OptChoose(0).Caption & "、" & OptChoose(1).Caption & "或" & OptChoose(2).Caption & "資料 !", vbCritical
               Else
                  MsgBox "請選擇" & OptChoose(0).Caption & "或" & OptChoose(1).Caption & "資料 !", vbCritical
               End If
               Exit Sub
            End If
         End If
         '92.1.16 end
         'Add By Cheng 2002/10/02
         m_strCountryEngName = ""
         
         Screen.MousePointer = vbHourglass
         For i = 0 To 6
            '910917 nick 沒有 陣列 1 物件
            If i <> 1 Then
                ' 90.10.15 modify by louis
                If i <> 2 Then
                   If txtCaseField(i).Enabled Then
                        If CheckKeyIn(i) = -1 Then
                           txtCaseField(i).SetFocus
                           txtCaseField_GotFocus (i)
                           Exit For
                        End If
                   End If
                End If
            'Add By Cheng 2002/08/19
            Else
               If CheckKeyIn(i) <> 1 Then
                  Me.Combo1.SetFocus
                  Exit For
               End If
            End If
         Next
         If i = 7 Then
'            If cmdCountry.Enabled And strMoney = "" Then
'               ShowMsg MsgText(9191)
'            Else
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
            
               'Add by Morgan 2005/5/24
               '詢問是否計算結餘
               If txtCaseField(6) = "" Then
                  '2011/11/8 modify by sonia TF子案不可結餘故加傳本所案號
                  'Pub_EndModCashMsg field(9)
                  Pub_EndModCashMsg field(9), field(1), field(2), field(3), field(4)
               End If
            
               'Add by Morgan 2006/8/21 馬來西亞新型延展費提醒
               'Modify by Morgan 2007/7/31 加俄羅斯新型
               'If field(9) = "018" And field(8) = "2" And cp(10) = "607" Then
               'Modified by Morgan 2015/6/9 +俄羅斯設計延展
               'If (field(9) = "018" Or field(9) = "233") And field(8) = "2" And cp(10) = "607" Then
               'Modified by Morgan 2015/10/15 俄羅斯2015/1/1修法所有未發證未延展新型案專用期都為10年不得延展
               'If ((field(9) = "018" And field(8) = "2") Or (field(9) = "233" And (field(8) = "2" Or field(8) = "3"))) And cp(10) = "607" Then
               'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
               'Modified by Morgan 2022/6/15 俄羅斯設計案 2015/1/1以前提申案件除了延展費外仍要繳年費，2015/1/1以後提申案件僅須繳延展費--禧佩
               'If ((field(9) = "018" And field(8) = "2") Or (field(9) = "023" And field(8) = "3")) And cp(10) = "607" Then
               If ((field(9) = "018" And field(8) = "2") Or (field(9) = "023" And field(8) = "3" And Val(field(10)) < 20150101)) And cp(10) = "607" Then
               'end 2022/6/15
               'end 2015/10/15
                  strExc(0) = GetPrjNationName(field(9))
                  If MsgBox(strExc(0) & lblTrademarkKind & "延展費無法檢查繳費紀錄，是否確定次數資料無誤？", vbYesNo + vbDefaultButton2) = vbNo Then
                     Screen.MousePointer = vbDefault 'Added by Morgan 2022/6/15
                     Exit Sub
                  End If
               End If
               
               If SaveDatabase Then
               
                  'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
                  PUB_CheckEMail cp(44), cp(116)
                  PUB_CheckEMail field(75), field(144)
                  If field(145) <> "" Then
                     PUB_CheckEMail field(75), field(145)
                  End If
                  'end 2008/2/20
               
                  '指示信
                  If txtCaseField(3) <> "N" Then
                     'Add by Amy 2018/03/20 612年費移作次年,指示信使用
                     If cp(10) = "612" Then
                        '輸[年]->年費指示信
                        If optSel(0).Value = 1 Then
                            tmpCp10 = 年費
                        '輸[次]且為美國->維持費指示信
                        ElseIf field(9) = "101" Then
                            tmpCp10 = 維持費
                        '其他->延展費指示信
                        Else
                            tmpCp10 = 延展費
                        End If
                     Else
                        tmpCp10 = cp(10)
                     End If
                     'Select Case cp(10)
                     Select Case tmpCp10
                     'end 2018/03/20
                        Case 年費
                           Select Case field(9)
                              'Add By Cheng 2003/03/14
                              '加申請國家日本
                              Case "011"
                                strTmp = "34"
                              Case "221" 'EPC年費'(核准後)
                                 'Modify by Morgan 2008/1/24 不必判斷發證日但必須是准
                                 'If field(20) <> "" And field(21) <> "" Then
                                 If field(20) <> "" And field(16) = "1" Then
                                    strTmp = "31"
                                 Else
                                    strTmp = "30"
                                 End If
                                 
                                 'Added by Morgan 2024/6/21 上網繳納改以代理人判斷且指示信是寄給財務處
                                 If InStr(Pub_GetSpecMan("各國專利局Y編號"), Combo1.Text) > 0 Then
                                    strTmp = "38"
                                 End If
                                 'end 2024/6/21
                              Case "102" '加拿大年費
                                 strTmp = "32"
                              'Add by Morgan 2008/8/18
                              Case "013" '香港
                                 strTmp = "35"
                              'Added by Morgan 2013/1/11
                              Case "126" '智利
                                 strTmp = "33"
                              'Added by Morgan 2024/4/25
                              Case "231"
                                 'Modified by Morgan 2024/6/21 上網繳納改以代理人判斷且指示信是寄給財務處
                                 'strTmp = "36"
                                 If InStr(Pub_GetSpecMan("各國專利局Y編號"), Combo1.Text) > 0 Then
                                    strTmp = "36"
                                 Else
                                    strTmp = "30"
                                 End If
                                 'end 2024/6/21
                              Case Else '一般
                                 strTmp = "30"
                           End Select
                        Case 維持費
                           Select Case field(9)
                              Case "101" '美國
                                 'Modified by Morgan 2024/6/21
                                 '至下一程序檔中找下一程序代號是繳年費及是否續辦為空，是則一般，若空的則是最後一次年費
                                 'strExc(0) = "SELECT COUNT(*) FROM NEXTPROGRESS WHERE " & ChgNextProgress(cp(1) & cp(2) & cp(3) & cp(4)) & _
                                 '   " AND NP07=" & 維持費 & " AND NP06 IS NULL"
                                 'intI = 1
                                 'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                                 'If intI = 1 Then
                                 '   If RsTemp.Fields(0) = 0 Then
                                 '      '最後一次
                                 '      strTmp = "33"
                                 '   Else
                                 '      '第一,二次
                                 '      strTmp = "32"
                                 '   End If
                                 'End If
                                 '
                                 ''92'10'7 add by sonia
                                 'If txtCaseField(2) = "" Then
                                 '   strTmp = "30"
                                 'End If
                                 ''92.10.7 end
                                 '上網繳納改以代理人判斷且指示信是寄給財務處
                                 If InStr(Pub_GetSpecMan("各國專利局Y編號"), Combo1.Text) > 0 Then
                                    strTmp = "38"
                                 Else
                                    strTmp = "30"
                                 End If
                                 'end 2024/6/21
                              Case Else '一般
                           
                           End Select
                           
                        Case 延展費
                           strTmp = "30"
                           'Added by Morgan 2024/6/21 上網繳納改以代理人判斷且指示信是寄給財務處
                           If InStr(Pub_GetSpecMan("各國專利局Y編號"), Combo1.Text) > 0 Then
                              'Added by Morgan 2025/8/8
                              If field(9) = "231" Then '德國
                                 strTmp = "36"
                              Else
                              'end 2025/8/8
                                 strTmp = "38"
                              End If
                           End If
                           'end 2024/6/21
                     End Select
                     
                     
                     If strTmp <> "" Then
                        
                        If Len(strCP09List) = 0 Then 'Added by Morgan 2020/8/19 EPC年費母案不用出指示信
                        
                           StartLetter "01", strTmp
                           'Modify by Morgan 2004/9/27
                           '年費,維持費,延展費加印傳真封面
                           'Modify by Morgan 2004/10/22
                           'NowPrint cp(9), "01", "99", False, strUserNum, , , True, stLetter
                           'Removed by Morgan 2018/10/22 取消傳真封面--慧汶
                           'NowPrint cp(9), "01", "89", False, strUserNum, , , True, stLetter, , , , , , , , , m_strAF01
                           'If m_strAF01 <> "" Then Sleep 1000 '等1秒以確保letterdemand不會發生dupe錯誤 Added by Morgan 2018/8/20
                           'end 2018/10/22
                           NowPrint cp(9), "01", strTmp, True, strUserNum, 0, stLetter, , , , , , , , , , , m_strAF01
                           
                           'Added by Morgan 2018/8/22 CFP電子化
                           If m_strAF01 <> "" Then
                              frm1105_1.m_RecNo = m_strAF01
                              frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".DATA.PDF"
                              frm1105_1.Show
                              If txtCaseField(9).Text = "Y" Then
                                 MsgBox "指示信編輯中，客戶函請至定稿維護修改！", vbExclamation
                                 txtCaseField(9).Text = ""
                              End If
                           End If
                           'end 2018/8/22
                  
                        'Add by Amy 2013/08/30 產生各國給各代的年費指示信
                        'Modifie by Morgan 2020/8/19
                        'If Len(strCP09List) > 0 Then
                        Else
                        'end 2020/8/19
                           strPrintCP09 = Split(strCP09List, ",")
                           For i = 0 To UBound(strPrintCP09)
                              'Added by Morgan 2024/4/25 EPC德國子案指示信不同
                              If strPrintCP09(i) = m_strEPC_DE_BCP09 Then
                                 StartLetter "01", "37", m_strEPC_DE_BCP09
                                 NowPrint m_strEPC_DE_BCP09, "01", "37", False, strUserNum, , , , , , , , , , , , , m_strEPC_DE_BCP09
                              Else
                              'end 2024/4/25
                              
                                 StartLetter "01", strTmp, strPrintCP09(i) 'Added by Morgan 2013/9/10
                                 'Added by Morgan 2018/8/22 CFP電子化
                                 'If m_strAF01 <> "" Then 'Removed by Morgan 2020/8/19
                                    m_strChildAF01 = strPrintCP09(i)
                                    'Removed by Morgan 2018/10/22 取消傳真封面--慧汶
                                    'NowPrint strPrintCP09(i), "01", "89", False, strUserNum, , , , , , , , , , , , , m_strChildAF01
                                    'Sleep 1000
                                    'end 2018/10/22
                                    NowPrint strPrintCP09(i), "01", strTmp, False, strUserNum, , , , , , , , , , , , , m_strChildAF01
                                    
                                 'Removed by Morgan 2020/8/19
                                 'Else
                                 ''end 2018/8/22
                                 '   'Removed by Morgan 2018/10/22 取消傳真封面--慧汶
                                 '   'NowPrint strPrintCP09(i), "01", "89", False, strUserNum, , , True, stLetter
                                 '   'end 2018/10/22
                                 '   NowPrint strPrintCP09(i), "01", strTmp, True, strUserNum, 0, stLetter
                                 'End If 'Added by Morgan 2018/8/22
                                 'end 2020/8/19
                                 
                              End If
                           Next
                           
                           'Modified by Morgan 2020/8/19
                           'If m_strAF01 <> "" Then MsgBox "子案【年費】指示信請至待處理區作業！", vbExclamation 'Added by Morgan 2018/8/22 CFP電子化
                           MsgBox "子案【年費】指示信請至待處理區作業！", vbExclamation
                           'end 2020/8/19
                        End If
                        'end 2013/08/30
                     End If
                     
                     
'Removed by Morgan 2012/3/7 不必詢問,需要時程序自行列印--甄妮
'                     StrSQLa = "Select FA04,FA05||' '||FA63||' '||FA64||' '||FA65,FA06,FA01||FA02 From CASEPROGRESS, FAGENT WHERE SUBSTR(CP44,1,8)=FA01(+) AND SUBSTR(CP44,9,1)=FA02(+) AND CP09='" & cp(9) & "'"
'                     rsA.CursorLocation = adUseClient
'                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'                     If rsA.RecordCount > 0 Then
'                       If MsgBox("代理人名稱(中)：" & rsA.Fields(0).Value & Chr(10) & Chr(13) & _
'                                 "　　　　　(英)：" & rsA.Fields(1).Value & Chr(10) & Chr(13) & _
'                                 "　　　　　(日)：" & rsA.Fields(2).Value & Chr(10) & Chr(13) & Chr(10) & Chr(13) & _
'                                 "是否列印代理人小信封？", vbExclamation + vbYesNo) = vbYes Then
'                           '列印地址條
'                           'Modify by Morgan 2006/10/17 改Call公用函數
'                           'PrintCase "" & rsA.Fields(3).Value
'                           PUB_PrintCase "" & rsA.Fields(3).Value
'                       End If
'                     End If
'                     If rsA.State <> adStateClosed Then rsA.Close
'                     Set rsA = Nothing
                  
                  End If
                  
                  '通知函
                  If txtCaseField(8) <> "N" Then
                     strTmp = "00"
                     'Add by Morgan 2007/1/18 新加坡發明案年費屆滿通知
                     'Removed by Morgan 2022/11/29 此為提申定稿,發文用一般就好 Ex:CFP-25469
                     'If field(9) = "014" And field(8) = "1" And txtCaseField(6) = "" Then
                     '   strTmp = "01"
                     'End If
                     'end 2022/11/29
                     'end 2007/1/18
                     
                     'Add by Morgan 2008/8/12 香港年費
                     If field(9) = "013" Then
                        strTmp = "02"
                     End If
                     
                     StartLetter "01", strTmp, , True 'Modified by Morgan 2018/5/24 + , True
                     NowPrint cp(9), "01", strTmp, IIf(Me.txtCaseField(9).Text = "Y", True, False), strUserNum, 0, , , , , , , , , , , , m_strLD18
                     
                     'Added by Morgan 2018/8/22 CFP電子化
                     If txtCaseField(9).Text = "Y" And m_strLD18 <> "" Then
                        frm1105_1.m_RecNo = m_strLD18
                        frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".CUS.PDF"
                        frm1105_1.Show
                     End If
                     'end 2018/8/22
                     
                  End If
                  
                  'Added by Lydia 2017/06/29 EPC案若未勾選國家，並出結案指示信(解除期限-一般定稿)
                  If strNCP09List <> "" Then
                     tmpArr = Empty
                     tmpArr = Split(strNCP09List, ",")
                     For i = 0 To UBound(tmpArr)
                        If Trim(tmpArr(i)) <> "" Then
                           ''Added by Morgan 2018/8/22 CFP電子化
                           'If m_strAF01 <> "" Then 'Removed by Morgan 2020/10/23
                              m_strChildAF01 = Trim(tmpArr(i))
                              NowPrint Trim(tmpArr(i)), "01", "89", False, strUserNum, , , , , , , , , , , , , m_strChildAF01
                              Sleep 1000
                              NowPrint Trim(tmpArr(i)), "14", "30", False, strUserNum, , , , , , , , , , , , , m_strChildAF01
                           
                           'Removed by Morgan 2020/10/23
                           'Else
                           'end 2018/8/22
                           '
                           '   NowPrint Trim(tmpArr(i)), "01", "89", False, strUserNum, , , True, stLetter
                           '   NowPrint Trim(tmpArr(i)), "14", "30", True, strUserNum, 0, stLetter
                           '
                           'End If 'Added by Morgan 2018/8/22
                           'end 2020/10/23
                        End If
                     Next i
                     If m_strAF01 <> "" Then MsgBox "子案【結案】指示信請至待處理區作業！", vbExclamation 'Added by Morgan 2018/8/22 CFP電子化
                  End If
                  'end 2017/06/29
                  
                  bolLeave = True
                  intLeaveKind = 1
                  'Add By Cheng 2002/04/30
                  '若有未發文資料顯示警告
                  PUB_GetCPunIssueDatas "" & Me.lblCaseField(1).Caption
                  Unload Me
                'Add By Cheng 2002/11/11
                Else
                    MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
               End If
            'End If
         End If
         Screen.MousePointer = vbDefault
      Case 1, 2
         'Add By Sindy 2013/5/28
         If frm050102_1.bolIsEMPFlow = True Then
            intLeaveKind = 0
            'Unload frm050102_1
            frm090202_4.Show
            frm090202_4.QueryData
         '2013/5/28 End
         'Add By Sindy 2018/1/8
         ElseIf Me.m_strIR01 <> "" Then
            intLeaveKind = 0
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         '2018/1/8 END
         Else
            If Index = 2 Then
               intLeaveKind = 0
            Else
               intLeaveKind = 1
            End If
         End If
         Unload Me
   End Select
   
    ' 發文回前畫面時
   Select Case Index
      Case 0:
         ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
         'Add By Sindy 2013/5/28
         If frm050102_1.bolIsEMPFlow = True Then
            intLeaveKind = 0
            'Unload frm050102_1
            frm090202_4.Show
            frm090202_4.QueryData
         '2013/5/28 End
         'Add By Sindy 2018/1/8
         ElseIf Me.m_strIR01 <> "" Then
            intLeaveKind = 0
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         '2018/1/8 END
         Else
            frm050102_1.Clear
         End If
   End Select
End Sub

'Add By Cheng 2002/10/02
'Modified by Morgan 2013/9/10 +ET02,原固定用cp(9)改可傳入(EPC子案要用)
Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String, Optional ByVal ET02 As String, Optional ByVal pIsCustLetter As Boolean = False)
Dim strTxt(1 To 6) As String, i As Integer, s As Integer, iStep As Integer, strTmp As String
Dim strAnnuity As String
'Add By Sindy 2009/07/16
Dim strFeeType As String
Dim strTmp1 As String
Dim strKey(5) As String
Dim strCaseFee(1 To 2) As String
Dim bFind As Boolean
Dim varRef As Variant
'2009/07/16 End
Dim StrYear1 As String, StrYear2 As String
 
   StrYear1 = textYear
   StrYear2 = Text1(0)
   
   'Removed by Morgan 2013/8/23 繳費年度改與實際相同,期限則另外控管
   ''Modified by Morgan 2013/1/16 印尼年費為事後繳
   'If field(9) = "017" And cp(10) = 年費 Then
   '   If StrYear1 > 1 Then StrYear1 = StrYear1 - 1
   '   If StrYear2 > 1 Then StrYear2 = StrYear2 - 1
   'End If
   
   If ET02 = "" Then ET02 = cp(9)
   
   EndLetter ET01, ET02, ET03, strUserNum
   Dim Jjj As Integer
   Jjj = 1
   
   If ET02 <> cp(9) Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) "
      strTxt(Jjj) = strTxt(Jjj) & "select '" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','指定國家',na04 from caseprogress,patent,nation where cp09='" & ET02 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and na01(+)=pa09"
      Jjj = Jjj + 1
   ElseIf m_strCountryEngName <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','指定國家','" & m_strCountryEngName & "')"
      Jjj = Jjj + 1
   End If
   strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','EPC合計費用','費用總金額')"
   Jjj = Jjj + 1
   
   'Add By Sindy 2009/07/16
   '年度說明
   'Modified by Morgan 2022/6/13 +field(10)
   strFeeType = PUB_GetNa20Na22Na24(field(9), field(8), field(10))
   If optSel(0).Value = True Then '繳費年度
      'Modify by Morgan 2009/11/5 若有修法時改抓修法次數對應的代理人編號
      'strTmp = PUB_GetYF15(field(9), field(8), "Y0000000", strFeeType, CDbl(strYear1))
      'strTmp1 = PUB_GetYF15(field(9), field(8), "Y0000000", strFeeType, CDbl(strYear2))
      strTmp = PUB_GetYF15(field(9), field(8), "Y000000" & m_iFixNo, strFeeType, CDbl(StrYear1))
      strTmp1 = PUB_GetYF15(field(9), field(8), "Y000000" & m_iFixNo, strFeeType, CDbl(StrYear2))
      If StrYear1 <> StrYear2 Then
         strTmp = strTmp & "至" & strTmp1
      End If
   Else '繳費次數
      strKey(0) = ""
      strKey(1) = cp(1)
      strKey(2) = cp(2)
      strKey(3) = cp(3)
      strKey(4) = cp(4)
      bFind = GetMoneyDate(field(8), field(9), strKey, strCaseFee(1), strCaseFee(2))
      If bFind Then
         If IsEmptyText(strCaseFee(2)) = False Then
            varRef = Split(strCaseFee(2), ",")
            'Modify by Morgan 2009/11/5 若有修法時改抓修法次數對應的代理人編號
            'strTmp = PUB_GetYF15(field(9), field(8), "Y0000000", strFeeType, Val(varRef(Val(textNum) - 1)))
            strTmp = PUB_GetYF15(field(9), field(8), "Y000000" & m_iFixNo, strFeeType, Val(varRef(Val(textNum) - 1)))
         End If
      End If
   End If
   '2009/07/16 End
   
   If optSel(0).Value = True Then
      Select Case Val(StrYear1)
         Case 1
            strAnnuity = StrYear1 & "st"
         Case 2
            strAnnuity = StrYear1 & "nd"
         Case 3
            strAnnuity = StrYear1 & "rd"
         Case Else
            strAnnuity = StrYear1 & "th"
      End Select
      '2008/8/22 add by sonia 019泰國新型第7年年費發文,指示信改為第7-8年,因第8年不用繳
      If field(9) = "019" And field(8) = "2" Then
         Select Case StrYear1
            Case 7
               strAnnuity = strAnnuity & " to 8th"
            Case 9             '2008/9/2 add by sonia 019泰國新型第9年年費發文,指示信改為第9-10年
               strAnnuity = strAnnuity & " to 10th"
         End Select
      End If
      '2008/8/22 end
      '2008/9/2 add by sonia 012韓國新型第2年年費發文,指示信改為第2-3年,因第3年不用繳
      If field(9) = "012" And field(8) = "2" Then
         Select Case StrYear1
            Case 2
               strAnnuity = strAnnuity & " to 3th"
         End Select
      End If
      '2008/9/2 end
      '2009/11/2 add by sonia 222波蘭新型第4年年費發文,指示信改為第4-5年,第6年年費發文,指示信改為第6-8年,第9年年費發文,指示信改為第9-10年
      If field(9) = "222" And field(8) = "2" Then
         Select Case StrYear1
            Case 4
               strAnnuity = strAnnuity & " to 5th"
            Case 6
               strAnnuity = strAnnuity & " to 8th"
            Case 9
               strAnnuity = strAnnuity & " to 10th"
         End Select
      End If
      '2009/11/2 end
      If StrYear1 = StrYear2 Then
'         strTmp = "第" & strYear1 & "年"
'         '2008/8/22 add by sonia 019泰國新型第7年年費發文,通知函改為第7-8年,因第8年不用繳
'         If field(9) = "019" And field(8) = "2" Then
'            Select Case strYear1
'               Case 7
'                  strTmp = "第" & strYear1 & "至8年"
'               Case 9          '2008/9/2 add by sonia 019泰國新型第9年年費發文,指示信改為第9-10年
'                  strTmp = "第" & strYear1 & "至10年"
'            End Select
'         End If
'         '2008/8/22 end
'         '2008/9/2 add by sonia 012韓國新型第2年年費發文,通知函改為第2-3年,因第3年不用繳
'         If field(9) = "012" And field(8) = "2" Then
'            Select Case strYear1
'               Case 2
'                  strTmp = "第" & strYear1 & "至3年"
'            End Select
'         End If
'         '2008/9/2 end
      Else
'         strTmp = "第" & strYear1 & "至" & strYear2 & "年"
         Select Case Val(StrYear2)
            Case 1
               strAnnuity = strAnnuity & " to " & StrYear2 & "st"
            Case 2
               strAnnuity = strAnnuity & " to " & StrYear2 & "nd"
            Case 3
               strAnnuity = strAnnuity & " to " & StrYear2 & "rd"
            Case Else
               strAnnuity = strAnnuity & " to " & StrYear2 & "th"
         End Select
      End If
   Else
      Select Case Val(textNum)
         Case 1
            strAnnuity = textNum & "st"
         Case 2
            strAnnuity = textNum & "nd"
         Case 3
            strAnnuity = textNum & "rd"
         Case Else
            strAnnuity = textNum & "th"
      End Select
      If textNum = Text1(1) Then
'         strTmp = "第" & textNum & "次"
'         '92.9.17 add by sonia
'         If field(9) = "101" Then
'            Select Case Val(textNum)
'               Case 1
'                  strTmp = "第1次(第5~8年)"
'               Case 2
'                  strTmp = "第2次(第9~12年)"
'               Case 3
'                  strTmp = "第3次(第13年~專用期屆滿)"
'            End Select
'         End If
'         '92.9.17 end
'         '2007/6/1 ADD BY SONIA
'         If field(9) = "231" Then
'            Select Case field(8)
'               Case "2"
'                  Select Case Val(textNum)
'                     Case 1
'                        strTmp = "第1次(第4~6年)"
'                     Case 2
'                        strTmp = "第2次(第7~8年)"
'                     Case 3
'                        strTmp = "第3次(第9年~專用期屆滿)"
'                  End Select
'               Case "3"
'                  Select Case Val(textNum)
'                     Case 1
'                        strTmp = "第1次(第6~10年)"
'                     Case 2
'                        strTmp = "第2次(第11~15年)"
'                     Case 3
'                        strTmp = "第3次(第16~20年)"
'                     Case 4
'                        strTmp = "第4次(第20年~專用期屆滿)"
'                  End Select
'            End Select
'         End If
'         '2007/6/1 end
      Else
'         strTmp = "第" & textNum & "至" & Text1(1) & "次"
'         '92.9.17 add by sonia
'         If field(9) = "101" Then
'            Select Case Val(textNum)
'               Case 1
'                  strTmp = "第1次(第5~8年)至"
'               Case 2
'                  strTmp = "第2次(第9~12年)至"
'               Case 3
'                  strTmp = "第3次(第13年~專用期屆滿)至"
'            End Select
'            Select Case Val(Text1(1))
'               Case 1
'                  strTmp = "第1次(第5~8年)"
'               Case 2
'                  strTmp = "第2次(第9~12年)"
'               Case 3
'                  strTmp = "第3次(第13年~專用期屆滿)"
'            End Select
'         End If
'         '92.9.17 end
'         '2007/6/1 ADD BY SONIA
'         If field(9) = "231" Then
'            Select Case field(8)
'               Case "2"
'                  Select Case Val(textNum)
'                     Case 1
'                        strTmp = "第1次(第4~6年)"
'                     Case 2
'                        strTmp = "第2次(第7~8年)"
'                     Case 3
'                        strTmp = "第3次(第9年~專用期屆滿)"
'                  End Select
'                  Select Case Val(Text1(1))
'                     Case 1
'                        strTmp = "第1次(第4~6年)"
'                     Case 2
'                        strTmp = "第2次(第7~8年)"
'                     Case 3
'                        strTmp = "第3次(第9年~專用期屆滿)"
'                  End Select
'               Case "3"
'                  Select Case Val(textNum)
'                     Case 1
'                        strTmp = "第1次(第6~10年)"
'                     Case 2
'                        strTmp = "第2次(第11~15年)"
'                     Case 3
'                        strTmp = "第3次(第16~20年)"
'                     Case 4
'                        strTmp = "第4次(第20年~專用期屆滿)"
'                  End Select
'                  Select Case Val(Text1(1))
'                     Case 1
'                        strTmp = "第1次(第6~10年)"
'                     Case 2
'                        strTmp = "第2次(第11~15年)"
'                     Case 3
'                        strTmp = "第3次(第16~20年)"
'                     Case 4
'                        strTmp = "第4次(第20年~專用期屆滿)"
'                  End Select
'            End Select
'         End If
'         '2007/6/1 end
         Select Case Val(Text1(1))
            Case 1
               strAnnuity = strAnnuity & " to " & Text1(1) & "st"
            Case 2
               strAnnuity = strAnnuity & " to " & Text1(1) & "nd"
            Case 3
               strAnnuity = strAnnuity & " to " & Text1(1) & "rd"
            Case Else
               strAnnuity = strAnnuity & " to " & Text1(1) & "th"
         End Select
      End If
   End If
   
   'Modify by Moran 2008/12/19 印尼發明不用說明繳費年度
   'If Not (field(9) = "017" And field(8) = "1") Then 'Removed by Morgan 2016/10/19 比照其他國家改為先繳納,不必再例外控制 --禧佩
   'end 2008/12/19
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','第幾年至幾年費','" & strTmp & "')"
      Jjj = Jjj + 1
   
   If Not (field(9) = "015" And (field(8) = "1" Or field(8) = "2")) Then 'Added by Morgan 2019/5/2 澳洲發明及新型案年費的指示信不要帶出年度--慧汶
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','列印備註'," & CNULL(strAnnuity) & ")"
       Jjj = Jjj + 1
   End If 'Added by Morgan 2019/5/2
   'End If
   
   '92.1.16 add by sonia
   'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
   If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
      If OptChoose(0).Value = True Then
         strTmp = "(Large Entity)"
      ElseIf OptChoose(1).Value = True Then
         strTmp = "(Small Entity)"
      'Added by Morgan 2013/3/20
      ElseIf OptChoose(2).Value = True Then
         strTmp = "(Micro Entity)"
      'end 2013/3/20
      End If
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','大小個體','" & strTmp & "')"
      Jjj = Jjj + 1
   End If
   '92.1.16 end
   
   'Added by Morgan 2012/4/10
   If field(9) = "231" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','無年費收據','♀')"
      Jjj = Jjj + 1
   End If
   
   'Added by Lydia 2016/06/15 EPC指定國家
   'Modified by Lydia 2016/06/17 准後(發證日)才秀EPC指定國家
   If pIsCustLetter Then 'Added by Morgan 2018/5/24 +判斷客戶函才要跑,
      If field(9) = "221" And Val(field(21)) > 0 Then
         'Modified by Morgan 2018/5/24 +判斷子案有年費進度
         'If ClsPDReadCountry(intCaseKind, field, strExc(0), True, , True) Then
         If ClsPDReadCountry(intCaseKind, field, strExc(0), True, True, True) Then
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','EPC指定國家','" & strExc(0) & "')"
            Jjj = Jjj + 1
         End If
      End If
   End If
   
   If Jjj = 1 Then Exit Sub
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(Jjj - 1, strTxt) Then
   If Not ClsLawExecSQL(Jjj - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   '*******************************************
End Sub

Private Function SaveDatabase() As Boolean
Dim strSql As String
Dim strPA72 As String
Dim strPA73 As String
Dim strPA74 As String
Dim strNP08 As String
Dim strNP09 As String
Dim strNP22 As String
Dim rsTmp As ADODB.Recordset
Dim nPosition As Integer
Dim strTxt(1 To 40) As String, iStep As Integer
Dim varTmp As Variant, varTmp1 As Variant, pa04 As String
Dim strCP04 As String, strCP09 As String, intR As Integer 'Add by Amy 2013/08/30
Dim str930Date As String 'Added by Lydia 2017/11/02
Dim strCP10 As String 'Add by Amy 2018/03/20
Dim strLetterJudge As String, strSubject As String, strChildCP09 As String, strChildCP45 As String '指示信判發人/主旨/子案收文號/彼號 Added by Morgan 2018/8/22
Dim ii As Integer
Dim strChildCP44 As String 'Added by Lydia 2019/01/15 子案最新/最後代理人

'911106 nick transation
SaveDatabase = True
On Error GoTo CheckingErr
cnnConnection.BeginTrans
   'SaveDatabase = False
   
   iStep = 1
   
   'Modify by Morgan 2008/2/21
   'cp(44) = Combo1
   intI = InStr(Combo1, "-")
   If intI > 0 Then
      cp(44) = Left(Combo1, intI - 1)
      cp(116) = Mid(Combo1, intI + 1)
   Else
      cp(44) = Combo1
      cp(116) = ""
   End If
   'end 2008/2/21
   cp(44) = ChangeCustomerL(cp(44))
   
   'Added by Morgan 2020/1/8 --秀玲
   If optSel(0).Value = True Then
      cp(53) = textYear
      cp(54) = Text1(0)
   Else
      cp(53) = textNum
      cp(54) = Text1(1)
   End If
   'end 2020/1/8
   
   ' 搜尋彼所案號
   'Modified by Morgan 2012/2/15 改呼叫共用函數
   'cp(45) = ""
   'strSql = "SELECT CP45 FROM CASEPROGRESS " & _
   '         "WHERE CP09 = (SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01 = '" & cp(1) & "' AND " & _
   '         "CP02 = '" & cp(2) & "' AND CP03 = '" & cp(3) & "' AND CP04 = '" & cp(4) & "' AND " & _
   '         "CP44 = '" & cp(44) & "' AND CP45 IS NOT NULL) "
   'Set rsTmp = New ADODB.Recordset
   'rsTmp.CursorLocation = adUseClient
   'rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'If rsTmp.RecordCount > 0 Then
   '   If Not IsNull(rsTmp.Fields("CP45")) Then
   '      cp(45) = rsTmp.Fields("CP45")
   '   End If
   'End If
   'rsTmp.Close
   'Set rsTmp = Nothing
   If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   'end 2012/2/15
   
   ' 更新案件進度檔
   cp(27) = txtCaseField(0)
   
   '92.1.16 modify by sonia
   'cp(64) = txtCaseField(7)
   'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
   'Modified by Morgan 2023/3/25
   'If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
   If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
      If strSrvDate(1) >= PA179啟用日 Then
         If OptChoose(0).Value = True Then
            new_Entity = OptChoose(0).Caption
         ElseIf OptChoose(1).Value = True Then
            new_Entity = OptChoose(1).Caption
         ElseIf OptChoose(2).Value = True Then
            new_Entity = OptChoose(2).Caption
         End If
         
      Else
   'end 2023/3/25
   
         If OptChoose(0).Value = True Then
            new_Entity = "大個體"
         ElseIf OptChoose(1).Value = True Then
            new_Entity = "小個體"
         'Added by Morgan 2013/3/20
         ElseIf OptChoose(2).Value = True Then
            new_Entity = "微個體"
         'end 2013/3/20
         End If
         
      End If 'Added 2023/3/25
      
      If old_Entity <> new_Entity And old_Entity <> "" Then  '改大小個體時
         If txtCaseField(7) = "" Then
            cp(64) = "原大小個體為" & old_Entity
         Else
            cp(64) = "原大小個體為" & old_Entity & "，" & Me.txtCaseField(7).Text
         End If
      Else
         cp(64) = txtCaseField(7)
      End If
   Else
      cp(64) = txtCaseField(7)
   End If
   '92.1.12 end
   '2005/4/15 ADD BY SONIA
   If old_cp44 = "Y49572" Then
      cp(46) = strSrvDate(1)
   End If
   '2005/4/15 END
   
   If m_strJpMemo <> "" Then cp(64) = m_strJpMemo & ";" & cp(64) 'Added by Morgan 2019/5/2 日本年費可減免備註加減免身分
   cp(81) = m_strCP81 'Added by Morgan 2019/5/2
   cp(113) = txtCP113 'Added by Lydia 2021/05/25 工作時數
   
   strTxt(iStep) = GetCPSQL(cp())
   
   '911106 nick transation
   cnnConnection.Execute strTxt(iStep)
   
   iStep = iStep + 1
   
   ' 更新基本檔的繳費年度及繳費日期
   strSql = Empty
   'Modify by Morgan 2006/8/21 馬來西亞新型延展費不用更新
   'Modify by Morgan 2007/7/31 加俄羅斯新型
   'If field(9) = "018" And field(8) = "2" And cp(10) = "607" Then
   'Modified by Morgan 2015/6/9 +俄羅斯設計延展
   'If (field(9) = "018" Or field(9) = "233") And field(8) = "2" And cp(10) = "607" Then
   'Modified by Morgan 2015/10/15 俄羅斯2015/1/1修法所有未發證未延展新型案專用期都為10年不得延展
   'If ((field(9) = "018" And field(8) = "2") Or (field(9) = "233" And (field(8) = "2" Or field(8) = "3"))) And cp(10) = "607" Then
   'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
   'Modified by Morgan 2022/6/15 俄羅斯設計案 2015/1/1以前提申案件除了延展費外仍要繳年費，2015/1/1以後提申案件僅須繳延展費--禧佩
   'If ((field(9) = "018" And field(8) = "2") Or (field(9) = "023" And field(8) = "3")) And cp(10) = "607" Then
   If ((field(9) = "018" And field(8) = "2") Or (field(9) = "023" And field(8) = "3" And Val(field(10)) < 20150101)) And cp(10) = "607" Then
   'end 2022/6/15
   'end 2015/10/15
      strPA72 = field(72)
      strPA73 = field(73)
      strPA74 = field(74)
   Else
      Dim i As Integer, iStart As Integer, iEnd As Integer
      
      If optSel(0).Value = True Then
         '910815 Sieg
         iStart = InStr(m_FeeYears, textYear)
         For i = 0 To UBound(varFeeYears)
            If Format(varFeeYears(i)) = textYear Then
               iStart = i
            End If
            If Format(varFeeYears(i)) = Text1(0) Then
               iEnd = i
               Exit For
            End If
         Next
      Else
         '910815 Sieg
         iStart = textNum - 1
         iEnd = Text1(1) - 1
      End If
      
      strExc(1) = ""
      strExc(2) = ""
      strExc(3) = ""
      For i = iStart To iEnd
         strExc(1) = strExc(1) & Format(varFeeYears(i)) & ","
         strExc(2) = strExc(2) & TransDate(txtCaseField(0), 2) & ","
         strExc(3) = strExc(3) & ","
      Next
      
      If Right(strExc(1), 1) = "," Then strExc(1) = Left(strExc(1), Len(strExc(1)) - 1)
      If Right(strExc(2), 1) = "," Then strExc(2) = Left(strExc(2), Len(strExc(2)) - 1)
      If Right(strExc(3), 1) = "," Then strExc(3) = Left(strExc(3), Len(strExc(3)) - 1)
      
      If field(72) <> "" Then
         strPA72 = field(72) & "," & strExc(1)
         strPA73 = field(73) & "," & strExc(2)
         strPA74 = field(74) & "," & strExc(3)
      Else
         strPA72 = strExc(1)
         strPA73 = strExc(2)
         strPA74 = strExc(3)
      End If
   End If
   
   '92.1.12 add by sonia 改大小個體時
   'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
   'Modified by Morgan 2023/3/25
   'If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Then
   If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
      If strSrvDate(1) >= PA179啟用日 Then
         If OptChoose(0).Value = True Then
            field(179) = "1"
         ElseIf OptChoose(1).Value = True Then
            field(179) = "2"
         ElseIf OptChoose(2).Value = True Then
            field(179) = "3"
         End If
      Else
   'end 2023/3/25
   
         If old_Entity <> new_Entity Then
            If InStr(1, field(91), old_Entity, 1) > 0 Then
               field(91) = Replace(field(91), old_Entity, new_Entity, InStr(1, field(91), old_Entity, 1), , 1)
            Else
               If field(91) = "" Then
                  field(91) = new_Entity
               Else
                  field(91) = new_Entity & "，" & field(91)
               End If
            End If
         End If
         
      End If 'Added by Morgan 2023/3/25
   End If
   '92.1.12 end
   field(72) = strPA72
   field(73) = strPA73
   field(74) = strPA74
   field(76) = txtCaseField(5)
   strTxt(iStep) = GetPASQL(field())
   
   '911106 nick transation
   cnnConnection.Execute strTxt(iStep)
   
   iStep = iStep + 1
   
   varTmp = Split(strCountry, ",")
   varTmp1 = Split(strMoneyCountry, ",")
   
   ' 若該案為母案則一併更新子案的案件
   If field(9) = "221" And field(3) = "0" And field(4) = "00" Then
      'Add by Morgan 2011/2/11 繳給EPO時子案也要更新繳費年度，否則若各國沒有繳過但有逾期通知時會無法帶出逾期的年度 Ex.CP-018628
      If strMoneyCountry = "" Then
         strTxt(iStep) = "UPDATE PATENT SET PA72 = '" & strPA72 & "', PA73 = '" & strPA73 & "', PA74 = '" & strPA74 & "' " & _
            "WHERE PA01 = '" & field(1) & "' AND PA02 = '" & field(2) & "' AND " & _
            "PA03 = '" & field(3) & "' AND PA04 <> '00' and pa57 is null"
         cnnConnection.Execute strTxt(iStep)
         iStep = iStep + 1
      Else
      'End 2011/2/11
         For i = 0 To UBound(varTmp1)
            If Format(varTmp1(i)) <> "" Then
               'Modify by Morgan 2005/7/13 PA74也要
               'Modified by Morgan 2013/9/10 +pa22EPC子案無證書號時更新為母案的證書號(有例外時會人工修改故不可一率更新)--甄妮
               strTxt(iStep) = "UPDATE PATENT SET PA72 = '" & strPA72 & "', PA73 = '" & strPA73 & "', PA74 = '" & strPA74 & "' " & _
                  ",PA22=NVL(PA22,'" & field(22) & "') WHERE PA01 = '" & field(1) & "' AND PA02 = '" & field(2) & "' AND " & _
                  "PA03 = '" & field(3) & "' AND PA04 <> '00' AND PA09 ='" & Format(varTmp1(i)) & "'"
               '911106 nick transation
               cnnConnection.Execute strTxt(iStep)
               iStep = iStep + 1
            End If
         Next
      End If
   End If
   
   ' 寫下一程序檔
   If txtCaseField(6) <> "" Then
      strNP09 = DBDATE(txtCaseField(6))
      strExc(1) = field(1)
      strExc(2) = field(9)
      strExc(3) = TransDate(strNP09, 2)
      GetCtrlDT strExc
      strNP08 = DBDATE(strExc(0))
      strNP22 = GetNextProgressNo()
        'Modify By Cheng 2003/11/24
        '重抓智權人員
'      strTxt(iStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & cp(10) & "," & strNP08 & "," & strNP09 & ",'" & cp(13) & "'," & strNP22 & ") "
        'Modify By Cheng 2003/12/08
        '若本所期限非工作天則抓最近的工作天
'      strTxt(iStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & cp(10) & "," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "'," & strNP22 & ") "
      'Modified by Lydia 2015/06/12 緬甸案的年費備註為廣告費
    '  strTxt(iStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & cp(10) & "," & PUB_GetWorkDay1(strNP08, True) & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "'," & strNP22 & ") "
      If field(9) = "048" Then
         strExc(5) = "刊登廣告"
      Else
         strExc(5) = ""
      End If
      'Add by Amy 2018/03/20
      strCP10 = cp(10)
      If cp(10) = "612" Then strCP10 = "605"
      strTxt(iStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22) " & _
               "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strCP10 & "," & PUB_GetWorkDay1(strNP08, True) & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "','" & strExc(5) & "'," & strNP22 & ") "
      'end 2018/03/20
      'end 2015/06/12
      '911106 nick transation
      cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
   End If
   
   'Added by Lydia 2017/06/01 印度催商業使用聲明,自動產生繳費期間內的商業使用聲明(930)期限
   'modify by sonia 2024/7/18 印度發明修改商業使用聲明每年呈報改為每三年呈報一次
'   If field(9) = "040" And optSel(0).Value = True And Trim(textYear) <> "" Then
'      If Trim(Text1(0)) = "" Or Trim(textYear) = Trim(Text1(0)) Then
'         strExc(0) = "1"
'      Else
'         strExc(0) = Val(Text1(0)) - Val(textYear) + 1
'      End If
'      'Added by Lydia 2017/11/02 抓專利起用日期來計算商業使用聲明期限的年度
'      If Trim(textYear) > "1" Then
'         str930Date = CompDate(0, Val(textYear) - 1, m_StartDate)
'      Else
'         str930Date = m_StartDate
'      End If
'      'end 2017/11/02
'
'      'Added by Morgan 2021/7/14 第1次商業使用聲期限
'      '3/31 以前核准->次年的9/30 Ex:2020/3/31-->2021/9/30 (提2020/4/1~2021/3/31之商業使用聲明)
'      '4/01 以後核准->後年的9/30 Ex:2020/4/01-->2022/9/30 (提2021/4/1~2022/3/31之商業使用聲明)
'      strExc(2) = ""
'      If field(20) <> "" Then
'         If Right(field(20), 4) <= "0331" Then
'            strExc(2) = Left(DBDATE(field(20)), 4) + 1
'         Else
'            strExc(2) = Left(DBDATE(field(20)), 4) + 2
'         End If
'         strExc(2) = strExc(2) & "0930"
'      End If
'      'end 2021/7/14
'
'      If str930Date <> "" Then 'Added by Lydia 2017/11/02
'        For ii = 1 To Val(strExc(0))
'            If ii = 1 Then
'               'Modified by Lydia 2017/11/02 從要繳年度起算
'               'strExc(1) = Mid(CompDate(0, 1, strSrvDate(1)), 1, 4) & "0331" '法限=明年3/31
'               'Modified by Morgan 2020/12/10 2020/10/20印度新法:法限改為9/30,所限=法限-1個月--禧佩
'               'strExc(1) = Mid(CompDate(0, 1, str930Date), 1, 4) & "0331" '法限=明年3/31
'               strExc(1) = Mid(CompDate(0, 1, str930Date), 1, 4) & "0930" '法限=明年9/30
'               'end 2020/12/10
'            Else
'               strExc(1) = CompDate(0, 1, strExc(1))
'            End If
'            'Modified by Morgan 2021/7/14 +判斷要大於第1次商業使用聲期限 Ex:CFP-025681
'            'If strExc(1) > strSrvDate(1) Then  'Added by Lydia 2017/11/02 若計算結果的法定期限<系統日則該期限不產生
'            If strExc(1) > strSrvDate(1) And strExc(1) >= strExc(2) Then
'            'end 2021/7/14
'                'Added by Lydia 2018/06/05 判斷下一程序期限是否存在
'                strSql = "select np01,np22,np06,np07 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' " & _
'                            "and np07='930' and np06 is null and np09=" & CNULL(strExc(1), True)
'                intI = 1
'                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                If intI = 0 Then
'                'end 2018/06/05
'                    'Modified by Morgan 2020/12/10 2020/10/20印度新法:法限改為9/30,所限=法限-1個月--禧佩
'                    'strExc(2) = PUB_GetWorkDay1(Mid(strExc(1), 1, 4) & "0131", True) '所限=明年1/31(抓工作天)
'                    strExc(2) = PUB_GetWorkDay1(CompDate(1, -1, strExc(1)), True) '所限=法限-1個月
'                    'end 2020/12/10
'                    strNP22 = GetNextProgressNo()
'                    strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                             "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','930'," & strExc(2) & "," & strExc(1) & ",'" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "'," & strNP22 & ") "
'                    cnnConnection.Execute strSql
'                End If 'end 2018/06/05
'            End If 'end 2017/11/02
'        Next ii
'      End If 'end 2017/11/02
'   End If
'   'end 2017/06/01
   If field(9) = "040" And field(8) = "1" Then    '印度發明2024/3/15修法：商業使用聲明原每年呈報改為每三年呈報一次
      str930Date = ""
      strSql = "select np01 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np07='930' and np06 is null union " & _
               "select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='930' and cp158=0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 0 Then                             '中間接進來案件才要補商業使用聲明930期限
         If field(20) <= 20230331 Then               '核准日 <=20230331，第一次期限為2026/09/30
            strExc(0) = 20260930                     '2023/4/1~20240331，第一次期限為2027/09/30
         Else                                        '2024/4/1~20250331，第一次期限為2028/09/30
            strExc(0) = field(20)                    '依此類推
            If Right(strExc(0), 4) >= "0401" Then    '依上述規則，核准日的日期在04/01~12/31之間的第一期限是4年後的09/30，故先加1年
               strExc(0) = strExc(0) + 10000
            End If
         End If
         strExc(1) = strExc(0)
         If strExc(0) > field(25) Then    '若第一次之法定已大於專利檔的專用期限則掛第一次期限
            GoTo AddNew
         End If
NextTime:
         If strExc(1) <= cp(7) Or strExc(1) < 20260930 Then
            strExc(1) = CompDate(0, 3, strExc(1))  '每3年一次
            GoTo NextTime
         End If
   
         If strExc(1) > field(25) Then       '新期限大於專利檔的專用止日時則掛專用期止日當年或隔年之9/30
            If Right(field(25), 4) <= "0930" Then       '判斷專用期止日是否<=當年9/30
               strExc(1) = Left(DBDATE(field(25)), 4)         '掛當年之9/30
            Else
               strExc(1) = Left(DBDATE(field(25)), 4) + 1     '掛次年之9/30
            End If
         Else
            strExc(1) = Left(DBDATE(strExc(1)), 4)            '只抓年度
         End If
         strExc(1) = strExc(1) & "0930"
AddNew:
         If strExc(1) <> "" Then
            strExc(2) = PUB_GetWorkDay1(CompDate(1, -1, strExc(1)), True) '所限=法限-1個月
            strSql = "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     " select '" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','930'," & strExc(2) & "," & strExc(1) & ",'" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "',newNP22 from dual,(select nvl(max(np22),0)+1 newNP22 from nextprogress)"
            cnnConnection.Execute strSql, intI
            str930Date = strExc(2)                            '存檔後彈訊息用
         End If
      End If
   End If
   'end 2024/7/18
   
   strCP09List = ""
   If cmdCountry.Enabled Then
      'Modify by Morgan 2006/12/25
      'If objPublicData.SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strMoneyCountry) Then
      'Modified by Morgan 2020/8/19 +傳cp(10)
      If PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strMoneyCountry, , , cp(10)) Then
      'end 2006/12/25
         
         iStep = 1
         'Memo by Morgan 2009/6/1 因為要控制該年度需繳費而不繳的才要閉卷所以要逐案更新
         For i = 0 To UBound(varTmp)
'--------------Memo by Lydia 2019/01/15 子案不繳費=>閉卷
            If InStr(strMoneyCountry, Format(varTmp(i))) = 0 Then
               pa04 = GetPA04(field(1), field(2), field(3), Format(varTmp(i)))
               'Added by Lydia 2017/06/29 EPC案若未勾選國家，並出結案指示信
               strExc(0) = AutoNo("B", 6)
               strExc(1) = PUB_GetAKindSalesNo(field(1), field(2), field(3), pa04)
               strExc(2) = ""
               'Added by Lydia 2017/12/01 抓子案的最後代理人/彼所案號 (EPC案母案解除期限和上閉卷，子案一併上閉卷和出結案指信，相關總收文號放母案的閉卷收文號，但是CF代理人要代子案的最後代理人　ex.CFP-26333)
               'Modified by Lydia 2019/01/15
               'If PUB_GetCP44(field(1), field(2), field(3), PA04, strExc(4), strExc(5), strExc(6)) = False Then '抓子案的最後代理人
               '   strExc(4) = cp(44) '沒有,就代母案代理人
               '   strExc(6) = cp(45)
               'End If
               'end 2017/12/01
               'Modified by Lydia 2019/01/31 EPC子案代理人依發文性質排順序
               'If PUB_GetCP44(field(1), field(2), field(3), PA04, strChildCP44, strExc(5), strChildCP45) = False Then '抓子案的最後代理人
               If PUB_GetEPCtoCP44(field(1), field(2), field(3), pa04, strChildCP44, strExc(5), strChildCP45) = False Then
                   strChildCP44 = cp(44) '沒有,就代母案代理人
                   strChildCP45 = cp(45) '彼所案號
               End If
               'end 2019/01/15
               
               '取消收文原因15:接獲指示不辦理
               'Modified by Lydia 2017/12/01
               'strExc(2) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP44,CP57,CP58) " & _
                           "VALUES ('" & field(1) & "','" & field(2) & "','" & field(3) & "','" & PA04 & "'," & strSrvDate(1) & ",'" & strExc(0) & "', " & _
                           "'913','" & PUB_GetStaffST15(strExc(1), 1) & "','" & strExc(1) & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'N'," & _
                           "'" & cp(9) & "','" & cp(44) & "'," & strSrvDate(1) & ",'15') "
               'Modified by Lydia 2019/01/15
               'strExc(2) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP44,CP45,CP57,CP58) " & _
                           "VALUES ('" & field(1) & "','" & field(2) & "','" & field(3) & "','" & PA04 & "'," & strSrvDate(1) & ",'" & strExc(0) & "', " & _
                           "'913','" & PUB_GetStaffST15(strExc(1), 1) & "','" & strExc(1) & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'N'," & _
                           "'" & cp(9) & "','" & strExc(4) & "','" & strExc(6) & "'," & strSrvDate(1) & ",'15') "
               strExc(2) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP44,CP45,CP57,CP58) " & _
                           "VALUES ('" & field(1) & "','" & field(2) & "','" & field(3) & "','" & pa04 & "'," & strSrvDate(1) & ",'" & strExc(0) & "', " & _
                           "'913','" & PUB_GetStaffST15(strExc(1), 1) & "','" & strExc(1) & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'N'," & _
                           "'" & cp(9) & "','" & strChildCP44 & "','" & strChildCP45 & "'," & strSrvDate(1) & ",'15') "
               cnnConnection.Execute strExc(2), intI
               'add by sonia 2017/7/12 更新cp44,cp45,否則指示信會帶錯(子案的代理人請設與子案指定國註冊或領證(後發者)同樣代理人，若有彼所案號請一併代出)
               'Mark by Lydia 2019/01/15 統一用子案最新/最後代理人
'               strExc(1) = "Select cp44,cp45,cp27,cp82 From CaseProgress,Patent  Where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp03='" & field(3) & "' and cp04='" & PA04 & "' " & _
'                                 "and (cp10='224' or cp10='601') and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 Order by cp27 Desc, cp82 Desc"
'               intR = 1
'               Set rsTmp = ClsLawReadRstMsg(intR, strExc(1))
'               If intR > 0 Then
'                   rsTmp.MoveFirst
'                   strExc(2) = "Update CaseProgress Set cp44=" & CNULL(IIf(IsNull(rsTmp.Fields("cp44")), "", rsTmp.Fields("cp44"))) & ",cp45=" & CNULL(IIf(IsNull(rsTmp.Fields("cp45")), "", rsTmp.Fields("cp45"))) & " " & _
'                                     "Where cp09='" & strExc(0) & "' "
'                   cnnConnection.Execute strExc(2)
'
'                   strChildCP45 = "" & rsTmp.Fields("cp45") 'Added by Morgan 2018/8/22
'               End If
'               'end 2017/7/12
               'end 2019/01/15
               strNCP09List = strNCP09List & strExc(0) & ","
               'end 2017/06/29
               
               strTxt(iStep) = "UPDATE PATENT SET PA57='Y' WHERE PA01='" & field(1) & "' AND PA02='" & field(2) & "' AND PA03='" & field(3) & "' AND PA04='" & pa04 & "'"
               '911106 nick transation
               cnnConnection.Execute strTxt(iStep)
               iStep = iStep + 1
               
               'Added by Morgan 2018/8/22 CFP電子化,子案結案指示信
               If strSrvDate(1) >= CFP指示信電子化啟用日 Then
                  strChildCP09 = strExc(0) '要傳給另一變數,否則內容會變
                  strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), "913", Format(varTmp(i)))
                  strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), pa04, "913", field(11), strChildCP45, Format(varTmp(i)))
                  PUB_AddAppForm strChildCP09, True, strLetterJudge, strSubject
               End If
               'end 2018/8/22
   
            '92.12.20 modify by sonia
            'End If
            ''Add By Cheng 2002/10/02
            'm_strCountryEngName = m_strCountryEngName & ", " & PUB_GetNationEngName("" & varTmp(i))
            'Modify by Morgan 2004/4/27
            '排除空字串
            'Else
            'Remove by Morgan 2009/6/1 下面重新組字串
            'ElseIf varTmp(i) <> "" Then
            '    m_strCountryEngName = m_strCountryEngName & ", " & PUB_GetNationEngName("" & varTmp(i))
            'Add by Amy 2013/08/30 +else 年費發文，子案的代理人請設與子案指定國註冊或領證(後發者)同樣代理人，若有彼所案號請一併代出
'--------------Memo by Lydia 2019/01/15 子案繳費
            Else
                If cp(10) = "605" Then
                   
                   strExc(0) = "Select cp04,cp09 From CaseProgress,Patent  Where cp43='" & cp(9) & "' and cp10='605' and pa09='" & varTmp(i) & "'  and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 "
                   strExc(0) = strExc(0) & " order by cp09 desc" 'Added by Morgan 2018/8/30 測試母案重新發文時避免抓到前次的子案進度
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    If intI > 0 Then
                        strCP04 = RsTemp.Fields("cp04")
                        strCP09 = RsTemp.Fields("cp09")
                        
                        'Modified by Lydia 2019/01/15 抓子案最新/最後代理人,排除本次產生的子案年費收文 (ex.CFP-28620-0-08)
                        'strExc(1) = "Select cp44,cp45,cp27,cp82 From CaseProgress,Patent  Where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & strCP04 & "' " & _
                        '                  "and (cp10='224' or cp10='601') and pa09='" & varTmp(i) & "'  and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 Order by cp27 Desc, cp82 Desc"
                        'intR = 1
                        'Set rsTmp = ClsLawReadRstMsg(intR, strExc(1))
                        'If intR > 0 Then
                        '    rsTmp.MoveFirst
                        '    strExc(2) = "Update CaseProgress Set cp44=" & CNULL(IIf(IsNull(rsTmp.Fields("cp44")), "", rsTmp.Fields("cp44"))) & ",cp45=" & CNULL(IIf(IsNull(rsTmp.Fields("cp45")), "", rsTmp.Fields("cp45"))) & " " & _
                        '                      "Where cp09='" & strCP09 & "' "
                        '
                        '    cnnConnection.Execute strExc(2)
                        '    strCP09List = strCP09List & strCP09 & ","
                        '
                        '    strExc(6) = "" & rsTmp.Fields("cp45") 'Added by Morgan 2018/8/22
                        'End If
                        'Modified by Lydia 2019/01/31 EPC子案代理人依發文性質排順序
                        'If PUB_GetCP44(field(1), field(2), field(3), strCP04, strChildCP44, strExc(5), strChildCP45, , strCP09) = False Then '抓子案的最後代理人
                        If PUB_GetEPCtoCP44(field(1), field(2), field(3), strCP04, strChildCP44, strExc(5), strChildCP45, strCP09) = False Then
                            strChildCP44 = cp(44) '沒有,就代母案代理人
                            strChildCP45 = cp(45) '彼所案號
                        End If
                        
                        'Added by Morgan 2024/4/25
                        If varTmp(i) = "231" Then
                           m_strEPC_DE_BCP09 = strCP09
                           'Added by Morgan 2024/6/21 固定帶德國專利局 --玫音
                           strChildCP44 = "Y55766000"
                           strChildCP45 = ""
                           'end 2024/6/21
                        End If
                        'end 2024/4/25
                        
                        strExc(2) = "Update CaseProgress Set cp44=" & CNULL(strChildCP44) & ",cp45=" & CNULL(strChildCP45) & " Where cp09='" & strCP09 & "' "
                        cnnConnection.Execute strExc(2)
                        strCP09List = strCP09List & strCP09 & ","
                        'end 2019/01/15
                        
                        
                        
                        'Added by Morgan 2018/8/22 CFP電子化,子案年費指示信
                        If strSrvDate(1) >= CFP指示信電子化啟用日 Then
                           If txtCaseField(3) <> "N" Then
                              strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), cp(10), Format(varTmp(i)))
                              'Modified by Lydia 2019/01/15
                              'strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), strCP04, cp(10), field(11), strExc(6), Format(varTmp(i)))
                              strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), strCP04, cp(10), field(11), strChildCP45, Format(varTmp(i)))
                              PUB_AddAppForm strCP09, True, strLetterJudge, strSubject
                           End If
                        End If
                        'end 2018/8/22
                        
                        'Added by Morgan 2020/8/19 子案年費各別管控提申收達
                        '提申期限
                        PUB_SetApplyDate cp(1), cp(2), cp(3), strCP04, cp(7), strCP09, cp(10), txtCaseField(0), Format(varTmp(i))
                        '收達
                        PUB_SetArriveDate strCP09
                        'end 2020/8/19
                    End If
                End If
            'end 2013/08/30
            End If
            '92.12.20 end
         Next
         strCP09List = IIf(Len(strCP09List) > 0, Left(strCP09List, Len(strCP09List) - 1), "")
         
         'Add by Morgan 2009/6/1 指定國家排序要用英文字母開頭(因為選國家的畫面有排序過所以照傳回來的國家代碼的順序就可以)
         m_strCountryEngName = ""
         For i = 0 To UBound(varTmp1)
            If varTmp1(i) <> "" Then
               m_strCountryEngName = m_strCountryEngName & ", " & PUB_GetNationEngName("" & varTmp1(i))
            End If
         Next
         'END 2009/6/1
         
         If m_strCountryEngName <> "" Then
            m_strCountryEngName = Mid(m_strCountryEngName, 3)
         End If
         
         'Add by Morgan 2004/4/27
         '最後一個逗號改用 and
         i = InStrRev(m_strCountryEngName, ",")
         If i > 0 Then
            m_strCountryEngName = Left(m_strCountryEngName, i - 1) & " and" & Mid(m_strCountryEngName, i + 1)
         End If
         
         '911106 nick transation
         'SaveDatabase = objLawDll.ExecSQL(iStep - 1, strTxt)
      
      End If
   End If
   
   ' 90.12.05 modify by louis 若案件國家收費表存在代理人收達天數則新增一筆收達的下一程序檔
   If SaveDatabase = True Then
      'Modify by Morgan 2015/8/7 發文收達期限管控改呼叫公用函式
      If cmdCountry.Enabled = False Then 'Added by Morgan 2020/8/19 EPC年費改管制子案，母案不用
         PUB_SetArriveDate cp(9)
      End If 'Added by Morgan 2020/8/19
      'end 2015/8/7
   End If
   
   'Added by Morgan 2015/8/7
   '提申管制
   If cmdCountry.Enabled = False Then 'Added by Morgan 2020/8/19 EPC年費改管制子案，母案不用
      PUB_SetApplyDate cp(1), cp(2), cp(3), cp(4), cp(7), cp(9), cp(10), txtCaseField(0), field(9)
   End If 'Added by Morgan 2020/8/19
   'end 2015/8/7
   
   'Add by Morgan 2005/5/20
   '更新結餘
   If txtCaseField(6) = "" Then
      Pub_UpdateEndModCash field(1), field(2), field(3), field(4)
   End If
   
   'Add by Sindy 2018/1/8
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050102_1"
   End If
   '2018/1/8 END
   
   'Added by Morgan 2018/8/22 CFP電子化
   If strSrvDate(1) >= CFP指示信電子化啟用日 Then
      If txtCaseField(3) <> "N" Then
         If cmdCountry.Enabled = False Then 'Added by Morgan 2020/8/19 EPC年費改管制子案，母案不用
            strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), cp(10), field(9))
            strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), cp(4), cp(10), field(11), cp(45), field(9))
            PUB_AddAppForm cp(9), True, strLetterJudge, strSubject
            m_strAF01 = cp(9)
         End If
      End If
   End If
   If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
      If txtCaseField(8) <> "N" Then
         strLetterJudge = PUB_GetLetterJudgeNew("1", field(1), cp(10), field(9))
         PUB_AddLetterProgress cp(9), 0, True, strLetterJudge, False, field(26), cp(10), field(75)
         m_strLD18 = cp(9)
      End If
   End If
   'end 2018/8/22
   
   cnnConnection.CommitTrans
   'add by sonia 2024/7/19
   If str930Date <> "" Then
      MsgBox "印度發明中間接進來案件，已補掛商業使用聲明期限，請至下一程序確認期限是否正確 !", vbCritical
   End If
   'end 2024/7/19
   
   Exit Function

EXITSUB:
   '911106 nick transation
     Exit Function
CheckingErr:
   SaveDatabase = False
   cnnConnection.RollbackTrans
End Function

Private Sub ReadAllData()
Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String, strTemp1 As String, j As Integer
Dim adoRecord As Object, strSameName As String
Dim strKey(0 To 4) As String

On Error GoTo HndErr
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'Modify by Morgan 2006/10/19 改不Call Dll
'If objPublicData.ReadAllData(frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.Row, 5), cp(), field(), intCaseKind, intPWhere) Then
ReDim cp(TF_CP) As String

   cp(9) = frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.row, 5)
   If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
   'end 2006/10/19
      lblCaseField(0) = cp(9)
      lblCaseField(1) = cp(1) + " - " + cp(2) + _
         IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
         IIf(cp(4) = "00", "", " - " + cp(4))
      lblCaseField(2) = TransDate(cp(6), 1)
      lblCaseField(4) = cp(13)
      lblCaseField(5) = TransDate(cp(7), 1)
      lblCaseField(3) = field(8)
      lblCaseField(6) = field(72)
      txtCaseField(7) = cp(64)
      '2005/4/15 ADD BY SONIA
      old_cp44 = cp(44)
      '2005/4/15 END
      
      'Add by Amy 2018/04/10 +612 年費移作次年
      If cp(10) = "612" Then
        optSel(0).Enabled = True
        optSel(1).Enabled = True
        optSel(0).Value = True
      '92.2.7 ADD BY SONIA
      ElseIf cp(10) = "605" Then
         optSel(1).Value = False
         optSel(1).Enabled = False
         optSel(0).Value = False
         optSel(0).Value = True
      Else
         optSel(0).Value = False
         optSel(0).Enabled = False
         optSel(1).Value = True
      End If
      '92.2.7 END
     
      '92.1.12 add by sonia
      'Modify by Morgan 2006/9/21 加法國
      'Modified by Morgan 2023/3/25
      'If field(9) = "101" Or field(9) = "102" Or field(9) = "203" Then
      '   'Added by Morgan 2013/3/20
      '   If field(9) = "101" Then
      '      optChoose(2).Enabled = True
      '   Else
      '      optChoose(2).Enabled = False
      '   End If
      '   'end 2013/3/20
      PUB_SetEntityOpt field(1), field(9), field(8), OptChoose
      If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
         If strSrvDate(1) >= PA179啟用日 Then
            If field(179) = "1" Then
               OptChoose(0).Value = True
               old_Entity = OptChoose(0).Caption
            ElseIf field(179) = "2" Then
               OptChoose(1).Value = True
               old_Entity = OptChoose(1).Caption
            ElseIf field(179) = "3" Then
               OptChoose(2).Value = True
               old_Entity = OptChoose(2).Caption
            Else
               old_Entity = ""
            End If
         Else
      'end 2023/3/25
      
            If InStr(1, field(91), "大個體", 1) > 0 Then
               OptChoose(0).Value = True
               old_Entity = "大個體"
            ElseIf InStr(1, field(91), "小個體", 1) > 0 Then
               OptChoose(1).Value = True
               old_Entity = "小個體"
            'Added by Morgan 2013/3/20
            ElseIf InStr(1, field(91), "微個體", 1) > 0 Then
               OptChoose(2).Value = True
               old_Entity = "微個體"
            'end 2013/3/20
            Else
               old_Entity = ""
            End If
            
         End If 'Added by Morgan 2023/3/25
         
         'Added by Morgan 2024/12/10 個體別順序會因國家有所不同,且客戶設定目前只設定是否可減免,原預設規則只適用於1,2選項為大小個體時
         If OptChoose(0).Caption = "大個體" And OptChoose(1).Caption = "小個體" Then
         'end 2024/12/10
         
            'Add by Morgan 2004/9/24
            If OptChoose(0).Value = False And OptChoose(1).Value = False Then
               Dim stAD03 As String
               For i = 1 To 5
                  If field(25 + i) <> "" Then
                     stAD03 = PUB_GetAD03(field(25 + i), field(9))
                     If stAD03 = "N" Then
                        OptChoose(0).Value = True
                        Exit For
                     '只要有未設定減免身分的公司申請人則不預設大小個體
                     ElseIf stAD03 = "" Then
                        Exit For
                     End If
                  End If
               Next
               '若五個申請人檢查完都不是大個體則為小個體
               If OptChoose(2).Enabled = False Then 'Added by Morgan 2013/3/20 不可選微個體時才預設
                  If OptChoose(0).Value = False And i = 6 Then OptChoose(1).Value = True
               End If
             End If
             
         End If 'Added by Morgan 2024/12/9
      End If
      '92.1.12 end
      
      '2005/7/8 MODIFY BY SONIA
      'If field(76) = "" Then
      '   txtCaseField(5) = field(26)
      'Else
      '   txtCaseField(5) = field(76)
      'End If
      txtCaseField(5) = field(76)
      '2005/7/8 END
      CheckKeyIn 5
      'Modify By Cheng 2002/08/19
   '   If objPublicData.GetCasePreAgent(cp(), strTemp) Then
   '      txtCaseField(1) = strTemp
   '      CheckKeyIn 1
   '   End If
      Set adoRecord = CreateObject("ADODB.Recordset")
      
      'Added by Morgan 2023/10/30 已有設定時不必再重新設定(IDS分案會先設,且抓預設代理人時也會剔除)
      If cp(44) <> "" Then
         Combo1 = cp(44) & IIf(cp(116) <> "", "-" & cp(116), "")
         CheckKeyIn 1
      'end 2023/10/30
      
      'Added by Morgan 2024/6/19 本所自行繳納設定
      '美國發明維持費
      ElseIf field(9) = "101" Then
         Combo1.AddItem "Y49572000"
         Combo1 = "Y49572000"
         CheckKeyIn 1
      '德國發明年費、新型及設計延展費
      ElseIf field(9) = "231" Then
         Combo1.AddItem "Y55766000"
         Combo1 = "Y55766000"
         CheckKeyIn 1
      'EPC核准前年費
      ElseIf field(9) = "221" And field(16) <> "1" Then
         Combo1.AddItem "Y55627000"
         Combo1 = "Y55627000"
         CheckKeyIn 1
      '歐盟設計延展費
      ElseIf field(9) = "239" Then
         Combo1.AddItem "Y55645000"
         Combo1 = "Y55645000"
         CheckKeyIn 1
      'end 2024/6/19
      
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.SelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
      '2007/4/23 MODIFY BY SONIA 加發文日降冪排序
      'If ClsPDSelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
      'Modify by Morgan 2008/2/21 加聯絡人
      'Added by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設
      ElseIf cp(31) = "Y" Then
         AddAgent Combo1, cp, , , , cp(9), field(9), field(26)
         If Combo1 <> "" Then CheckKeyIn 1
         
      Else '非新案照原本
      
         If ClsPDSelectTable("select cp44||decode(cp116,null,null,'-'||cp116) from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "' and cp09<'C' and cp44 is not null order by cp27 desc", adoRecord) Then
         '2007/4/23 END
            Do While adoRecord.EOF = False
               If IsNull(adoRecord.Fields(0).Value) = False Then
                  If strSameName <> adoRecord.Fields(0).Value Then
                     Combo1.AddItem adoRecord.Fields(0).Value
                     strSameName = adoRecord.Fields(0).Value
                  End If
               End If
               adoRecord.MoveNext
            Loop
            Combo1 = Combo1.List(0)
         End If
         
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCasePreAgent(cp(), strTemp) Then
         'Modified by Morgan 2024/6/24
         'If ClsPDGetCasePreAgent(cp(), strTemp) Then
         If ClsPDGetCasePreAgent(cp(), strTemp, , , , False) Then
         'end 2024/6/24
            Combo1 = strTemp
            CheckKeyIn 1
         End If
         'add by sonia 2018/4/23 日本Y20755不再代繳年費,預設改Y20332
         'modify by sonia 2018/12/4 +Y45397
         If Left(Combo1, 6) = "Y20755" Or Left(Combo1, 6) = "Y45397" Then
            Combo1.AddItem "Y20332000"
            Combo1 = "Y20332000"
            CheckKeyIn 1
         End If
         'end 2018/4/23
      End If
      'end 2016/10/27
      
      If field(9) <> EPC指定國家 Then
         cmdCountry.Enabled = False
      Else
         '2007/9/27 不必判斷發證日但必須是准 CFP-017891
         'If field(20) <> "" And field(21) <> "" Then
         'Modify by morgan 2009/10/2 加判斷有公告日(已核准未公告仍繳給EPO Ex.CFP-18628)
         'If field(20) <> "" And field(16) = "1" Then
         'Modified by Morgan 2017/11/29  繳費期限大於公告日才要繳到各國 Ex.CFP-28143--玫音
         If field(14) <> "" And field(20) <> "" And field(16) = "1" And DBDATE(cp(7)) > DBDATE(field(14)) Then
            cmdCountry.Enabled = True
         Else
            cmdCountry.Enabled = False
         End If
      End If
      If cmdCountry.Enabled Then
         'If objPublicData.ReadCountry(intCaseKind, cp(), strCountry, True) = True Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.ReadCountry(intCaseKind, cp(), strCountry, False) = True Then
         'Modify by Morgan 2007/12/24 閉卷不要
         'If ClsPDReadCountry(intCaseKind, cp(), strCountry, False) = True Then
         If ClsPDReadCountry(intCaseKind, cp(), strCountry, True) = True Then
         'end 2007/12/24
            If strCountry = "" Then
               MsgBox "所有子案已閉卷 !", vbCritical
               'Add By Cheng 2002/11/22
               intLeaveKind = 2
               GoTo err1
            End If
         End If
      Else
         strCountry = ""
      End If
      
      ' 依國家檔所記錄的資料取得開始繳費的日期及繳費年度串列字串
      'Modify by Morgan 2009/11/5 改呼叫共用函式
      'GetFeeStartDate m_StartDate, m_FeeYears
      strKey(0) = cp(9)
      strKey(1) = cp(1)
      strKey(2) = cp(2)
      strKey(3) = cp(3)
      strKey(4) = cp(4)
   
      GetMoneyDate field(8), field(9), strKey, m_StartDate, m_FeeYears, , cp(10), m_iFixNo
      'Added by Morgan 2016/1/8
      '印度,馬來西亞設計案延展期限從最早優先權日起算
      'Modified by Morgan 2016/10/19 巴基斯坦發明及設計之專利期間及年費均從優先權日起算--禧佩
      If (field(8) = "3" And (field(9) = "018" Or field(9) = "040")) Or (field(9) = "038" And (field(8) = "1" Or field(8) = "3")) Then
         strExc(0) = PUB_GetFirstPriDate(field)
         If strExc(0) <> "" Then
            m_StartDate = strExc(0)
         End If
      End If
      'end 2011/9/15
      'end 2016/1/8
      'end 2009/11/5
      varPA72 = Split(field(72), ",")
      varFeeYears = Split(m_FeeYears, ",")
      'Add by Morgan 2007/6/12
      '義大利發明已繳過第4年的繳費年度要加4才不會抓錯
      If UBound(varPA72) <> -1 And UBound(varFeeYears) <> -1 Then
         If field(9) = "204" And field(8) = "1" Then
            If varPA72(0) = 4 And varFeeYears(0) = 5 Then
               m_FeeYears = "4," & m_FeeYears
               varFeeYears = Split(m_FeeYears, ",")
            End If
         End If
      End If
      'End 2007/6/12
      '2008/7/18 add by sonia 荷蘭發明新型原自第5年起繳年費,2008年修法自第4年起
      '                       因無日期可判斷, 故以是否曾繳過年費判斷, 已繳過且從第5年起則仍以第5年起算,未繳過則以第4年起算
   '2010/4/21 CANCEL BY SONIA 改在 GetMoneyDate做
   '   If field(9) = "207" And field(72) <> "" Then
   '      If varPA72(0) = 5 Then
   '         m_FeeYears = Mid(m_FeeYears, 3)
   '         varFeeYears = Split(m_FeeYears, ",")
   '      End If
   '   End If
   '   '2008/7/18 end
      
      '2008/8/4 add by sonia 2008年修法自第4年起繳年費,但若第4年已過期則仍自第5年起繳 CFP-016429
   '2010/4/21 CANCEL BY SONIA 改在 GetMoneyDate做
   '   If field(9) = "207" And field(72) = "" Then
   '      If DBDATE(GetFeeNextDate(m_StartDate, 3, field(9), field(8))) < strSrvDate(1) Then
   '         m_FeeYears = Mid(m_FeeYears, 3)
   '         varFeeYears = Split(m_FeeYears, ",")
   '      End If
   '   End If
      '2008/8/4 END
      
      txtCaseField(4) = "Y"
      txtCaseField(4).Enabled = False 'Add by Amy 2018/04/11
   
      If field(9) = "101" And cp(10) = "606" Then
         txtCaseField(8) = "N"
         'Removed by Morgan 2024/6/19 取消--玫音
         'txtCaseField(2) = "Y"
         'txtCaseField(3) = "N" 'Added by Morgan 2018/9/5 美國維持費發文請預設不產生指示信(上N)--玫音
         'end 2024/6/19
      End If
      
      'Modify by Morgan 2008/1/24 EPC進各國要照各國年費規定，故需排除該年度不需繳費的國家
      If strCountry <> "" Then
         strExc(1) = ""
         If field(72) = "" Then
            If m_FeeYears = "" Then
               strExc(1) = 3
            Else
               strExc(1) = Val(m_FeeYears) 'VB會回傳第一個數字
            End If
         Else
            strExc(1) = Val(varPA72(UBound(varPA72))) + 1
         End If
         
         If strExc(1) <> "" Then
            strSql = "select na01 from nation where na01 in (" & strCountry & ") and instr(','||na21||',','," & strExc(1) & ",')>0"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strExc(2) = RsTemp.GetString(adClipString, , , ",")
               strCountry = Left(strExc(2), Len(strExc(2)) - 1)
            End If
         End If
      End If
      'end 2008/1/24
      
      'Added by Lydia 2021/05/25
      txtCP113 = ""
      If cp(113) <> "" Then txtCP113 = cp(113)
      'end 2021/05/25
         
      'add by sonia 2024/7/18 印度發明案無核准日無法計算商業使用聲明930期限故設定不能發文
      If field(9) = "040" And field(8) = "1" And field(20) = "" Then
         MsgBox "印度發明案無核准日，無法計算商業使用聲明期限 !", vbCritical
         cmdok(0).Enabled = False
      End If
      'end 2024/7/18
      
   Else
err1:
       'Modify By Cheng 2002/11/22
   '   bolLeave = True
   '   intLeaveKind = 1
   '   Unload Me
   End If
   Screen.MousePointer = varSaveCursor
   Exit Sub

HndErr:
ErrorMsg
Screen.MousePointer = varSaveCursor
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   Dim strNo As String, iPos As Integer
   If Combo1.Text <> "" Then
      If CheckKeyIn(1) = -1 Then
         Cancel = True
      End If
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If Cancel = False Then
         strNo = Combo1.Text
         'Add by Morgan 2008/2/21 加聯絡人判斷
         iPos = InStr(Combo1.Text, "-")
         If iPos > 0 Then
            strNo = Left(Combo1.Text, iPos - 1)
         End If
         'end 2008/2/21
         
         If PUB_CheckStatus(strNo) = False Then
            Cancel = True
         'add by sonia 2018/4/23
         ElseIf Left(Combo1, 6) = "Y20755" Then
            MsgBox "Y20755 不再代繳年費，請更換代理人！", vbExclamation
            Cancel = True
         'end 2018/4/23
         'add by sonia 2018/12/4
         ElseIf Left(Combo1, 6) = "Y45397" Then
            MsgBox "Y45397 不再代繳年費，請更換代理人！", vbExclamation
            Cancel = True
         'end 2018/12/4
         'Added by Morgan 2012/3/7 發文都要顯示代理人備註--甄妮
         Else
            strExc(0) = "select FA29 from Fagent where " & ChgFagent(strNo) & " and FA29 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "" & RsTemp(0), vbExclamation, "代理人備註"
            End If
         'end 2012/3/7
         End If
      End If
      
      If Cancel Then Combo1.SetFocus
   End If
End Sub

Private Sub Form_Activate()
    'Add By Cheng 2002/11/22
    If intLeaveKind = "2" Then
        intLeaveKind = 1
        bolLeave = True
        Unload Me
    End If
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String

   Select Case Index
      Case 3
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
            lblTrademarkKind = strTemp
         End If
      Case 4
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
         If ClsPDGetStaff(lblCaseField(Index), strTemp) Then
            lblSalesName = strTemp
         Else
            lblSalesName = ""
         End If
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   bolLeave = False
   intLeaveKind = 1
   strMoneyCountry = ""
   strCountry = ""
   txtCaseField(0) = strSrvDate(2)
   ReadAllData
   'Add by Morgan 2006/7/4
   '加控制只有香港可改下次繳費日
   If field(9) = "013" Then
      txtCaseField(6).Locked = False
   Else
      txtCaseField(6).Locked = True
   End If
   
'Removed by Morgan 2016/10/19 比照其他國家改為先繳納,不必再例外控制 --禧佩
'   'Added by Morgan 2013/1/16
'   If field(9) = "017" Then
'      'Modified by Morgan 2013/8/23
'      'MsgBox "印尼年費非預先繳納,指示信會比實際輸入年度少一年", vbExclamation, "印尼年費提醒"
'      MsgBox "印尼年費為使用後繳納,期限為該年度專利期止日後的發證月日！", vbExclamation, "印尼年費提醒"
'   End If
'end 2016/10/19
   
   If PUB_ChkFileNP(cp(9)) Then MsgBox "下一程序已有提申或收達期限，若為重新發文時需要先刪除後才可作業！" 'Added by Morgan 2015/8/7
   
   'Add By Sindy 2018/1/8
   m_strIR01 = frm050102_1.m_strIR01
   m_strIR02 = frm050102_1.m_strIR02
   m_strIR03 = frm050102_1.m_strIR03
   m_strIR04 = frm050102_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2018/1/8 END
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bolLeave = False Then
      If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Cancel = 1
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If intLeaveKind = 1 Then
      frm050102_1.Show
   ElseIf intLeaveKind = 0 Then
     Unload frm050102_1
   End If
   ShowEditForm 'Added by Morgan 2018/8/22
   
   'Set frm050102_9 = Nothing 'Removed by Morgan 2021/12/10 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub optSel_Click(Index As Integer)
On Error Resume Next
   If Index = 0 Then
      textYear.Enabled = True
      Text1(0).Enabled = True
      'Add by Amy 2018/03/20 612年費移作次年需顯示(第x年費用或第x次費用需擇一輸入)
      If cp(10) <> "612" Then
        textYear.Visible = True
        Text1(0).Visible = True
        textNum.Visible = False
        Text1(1).Visible = False
      End If
      textNum.Enabled = False
      Text1(1).Enabled = False
      textYear.SetFocus
   Else
      textYear.Enabled = False
      Text1(0).Enabled = False
      'Add by Amy 2018/03/20 612年費移作次年需顯示(第x年費用或第x次費用需擇一輸入)
      If cp(10) <> "612" Then
        textYear.Visible = False
        Text1(0).Visible = False
         textNum.Visible = True
        Text1(1).Visible = True
      End If
      textNum.Enabled = True
      Text1(1).Enabled = True
      textNum.SetFocus
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strNext As String, nNext As Integer
Dim strDate As String
 
    'Modify By Cheng 2002/11/11
'   If Text1(Index) = "" Then
   If Index = 0 Then
      If Val(textYear) > Val(Text1(Index)) Then
         MsgBox "繳費年度錯誤，請重新輸入 !", vbCritical
         Cancel = True
      'Add by Morgan 2004/9/29 香港或日本19940101以前申請不控制
      ElseIf field(9) = "013" Or (field(9) = "011" And Val(field(10)) < 19940101) Then
         '基本資料可能有錯，不作控制。
      Else
         'Modify by Morgan 2010/2/24
         'If InStr(m_FeeYears, Text1(Index)) = 0 Then
         If InStr("," & m_FeeYears & ",", "," & Text1(Index) & ",") = 0 Then
            MsgBox "繳費年度錯誤，請重新輸入 !", vbCritical
            Cancel = True
         End If
      End If
   Else
      If Val(textNum) > Val(Text1(Index)) Then
         MsgBox "繳費次數錯誤，請重新輸入 !", vbCritical
         Cancel = True
      'Add by Morgan 2006/8/21 馬來西亞新型除年費外還需繳延展費
      ElseIf field(9) = "018" And field(8) = "2" And cp(10) = "607" Then
         '因延展費無繳費記錄所以無法正確檢查，存檔時會再提醒
         If Val(Text1(Index)) > 2 Then
            MsgBox "馬來西亞延展次數不可超過 2!", vbCritical
            Cancel = True
         End If
      'END 2006/8/21
      
      'Add by Morgan 2007/7/31 俄羅斯新型除年費外還需繳延展費
      'Modified by Morgan 2015/10/15 俄羅斯2015/1/1修法所有未發證未延展新型案專用期都為10年不得延展
      'ElseIf field(9) = "233" And field(8) = "2" And cp(10) = "607" Then
      '   '因延展費無繳費記錄所以無法正確檢查，存檔時會再提醒
      '   If Val(Text1(Index)) > 1 Then
      '      MsgBox "俄羅斯延展次數不可超過 1!", vbCritical
      '      Cancel = True
      '   End If
      'end 2015/10/15
      'END 2007/7/31
      
      'Add by Morgan 2015/6/9 俄羅斯設計除年費外還需繳延展費(自申請起5年,期滿可延4次,每次5年,共25年)
      'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
      'Modified by Morgan 2022/6/15 2015 2015/1/1以前提申案件除了延展費外仍要繳年費，2015/1/1以後提申案件僅須繳延展費--禧佩
      'ElseIf field(9) = "023" And field(8) = "3" And cp(10) = "607" Then
      ElseIf field(9) = "023" And field(8) = "3" And cp(10) = "607" And Val(field(10)) < 20150101 Then
         '因延展費無繳費記錄所以無法正確檢查，存檔時會再提醒
         '20141001以前申請者自申請起15年,期滿可延1次
         If Val(field(10)) < 20141001 And Val(Text1(Index)) > 1 Then
            MsgBox "俄羅斯延展次數不可超過 1!", vbCritical
            Cancel = True
         '20141001以後申請者自申請起5年,期滿可延4次,每次5年,共25年
         ElseIf Val(field(10)) >= 20141001 And Val(Text1(Index)) > 5 Then
            MsgBox "俄羅斯延展次數不可超過 5!", vbCritical
            Cancel = True
         End If
      'END 2015/6/9
      
      Else
         If Val(Text1(Index)) - 1 > UBound(Split(m_FeeYears, ",")) Then
            MsgBox "繳費次數錯誤，請重新輸入 !", vbCritical
            Cancel = True
         End If
      End If
   End If
   
   If Cancel Then
      TextInverse Text1(Index)
      Exit Sub
   End If
   
   If Text1(Index) <> "" Then
        'Remove by Morgan 2006/11/9 香港也要檢查專用期止日
        ''若申請國家為香港
        'If field(9) = "013" Then
        '    If Index = 0 Then
        '       txtCaseField(6) = TransDate(DBDATE(GetFeeNextDate(m_StartDate, Me.Text1(Index))), 1)
        '    Else
        '       txtCaseField(6) = TransDate(DBDATE(GetFeeNextDate(m_StartDate, varFeeYears(Me.Text1(Index)))), 1)
        '    End If
        'Else
        'end 2006/11/9
            If Index = 0 Then
               'Add by Morgan 2009/4/6 還有繳費年度時才要掛期限
               txtCaseField(6) = ""
               If Val(varFeeYears(UBound(varFeeYears))) > Val(Text1(Index)) Then
               'end 2009/4/6
                  
'Removed by Morgan 2016/10/19 比照其他國家改為先繳納,不必再例外控制 --禧佩
'                  'Added by Morgan 2013/8/23
'                  '印尼年費為使用後繳納,期限為該年度專利期止日後的發證月日 Ex.CFP-23499 申請日 99/10/20 發證日 101/7/12 繳1-4年(99/10/20-103/10/19)則第5年(103/10/20-104/10/19)期限應為 105/7/12
'                  If field(9) = "017" And cp(10) = "605" Then
'                     If Right(field(21), 4) < Right(field(10), 4) Then
'                        strDate = CompDate(0, Text1(Index) + 2, field(10))
'                     Else
'                        strDate = CompDate(0, Text1(Index) + 1, field(10))
'                     End If
'                     strDate = (strDate \ 10000) & Right(field(21), 4)
'                     txtCaseField(6) = TransDate(strDate, 1)
'                  Else
'                  'end 2013/8/23
'end 2016/10/19
                  
                     '2009/11/2 modify by sonia 以下次繳費年度-1計算 CFP-015442繳第7年為第7-8年故以第9年-1計算
                     'txtCaseField(6) = TransDate(DBDATE(GetFeeNextDate(m_StartDate, Me.Text1(Index))), 1)
                     txtCaseField(6) = TransDate(DBDATE(GetFeeNextDate(m_StartDate, GetNextValue(m_FeeYears, Text1(0), nNext) - 1, field(9), field(8))), 1)
                     '2009/11/2 END
                     '2009/1/16 add by sonia CFP-019667
                     If field(9) = "012" And field(8) = "2" And field(10) > 19990701 And field(10) < 20061001 And Text1(Index) = "2" Then
                        txtCaseField(6) = TransDate(DBDATE(GetFeeNextDate(m_StartDate, 3, field(9), field(8))), 1)
                     End If
                     '2009/1/16 end
                     
'                  End If 'Added by Morgan 2013/8/23
                  
                  If Val(TransDate(txtCaseField(6), 2)) >= Val(field(25)) And field(25) <> "" Then
                     txtCaseField(6) = ""
                  End If
               End If
               
            'Added by Morgan 2015/10/16
            '2014/10/01之後提出申請的俄羅斯設計案,專用期間為自申請起5年,期滿可延4次,每次5年,共25年
            'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
            'Removed by Morgan 2022/6/15 2015 2015/1/1以前提申案件除了延展費外仍要繳年費，2015/1/1以後提申案件僅須繳延展費--禧佩
            'ElseIf field(9) = "023" And field(8) = "3" And cp(10) = "607" And Val(field(10)) >= 20141001 And Val(Text1(Index)) < 4 Then
            '   txtCaseField(6) = TransDate(CompDate(0, (Val(Text1(Index)) + 1) * 5, field(10)), 1)
            'end 2022/6/15
            'end 2015/10/16
            
            Else
               If Me.Text1(Index) <= UBound(varFeeYears) Then
                  txtCaseField(6) = TransDate(DBDATE(GetFeeNextDate(m_StartDate, varFeeYears(Me.Text1(Index)), field(9), field(8))), 1)
                  
                  'Added by Morgan 2023/2/4 Ex:CFP-025393-2
                  If Val(TransDate(txtCaseField(6), 2)) >= Val(field(25)) And field(25) <> "" Then
                     txtCaseField(6) = ""
                  End If
                  'end 2023/2/4
               Else
                  txtCaseField(6) = ""
               End If
            End If
        'End If
    End If
End Sub

Private Sub txtCaseField_Change(Index As Integer)
   Select Case Index
      Case 5
         lblNotify = ""
   End Select
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 1, 3, 4, 5, 8, 9
                 KeyAscii = UpperCase(KeyAscii)
      'Add By Cheng 2002/07/31
      Case 2, 4, 9
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> 89 Then
            KeyAscii = 0
         End If
   End Select
End Sub

'92.10.7 add by sonia
'Removed by Morgan 2024/6/21 沒用了,上網繳納改以代理人判斷
'Private Sub txtCaseField_LostFocus(Index As Integer)
'   Select Case Index
'       Case 2
'          If txtCaseField(2) = "Y" Then
'             txtCaseField(8) = "N"
'             txtCaseField(3) = "N" 'Added by Morgan 2019/4/15 指示信也要預設--禧佩
'          Else
'             txtCaseField(8) = ""
'             txtCaseField(3) = "" 'Added by Morgan 2019/4/15 指示信也要預設--禧佩
'          End If
'   End Select
'End Sub
'end 2024/6/21
'92.10.7 end

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = -1 Then
      Cancel = True
   End If
   
   'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
   If Cancel = False And Index = 5 Then
      If PUB_CheckStatus(txtCaseField(Index).Text) = False Then Cancel = True
   End If
      
   If Cancel Then txtCaseField_GotFocus (Index)
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strCusTemp As String, strTemp1 As String, strTemp2 As String
Dim varTemp As Variant, strStartDate As String

   CheckKeyIn = -1
   Select Case intIndex
      Case 0
         If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
            CheckKeyIn = 1
         End If
         
      Case 1 '代理人
         lblAgent = ""
         If Combo1.Text = "" Then
            MsgBox "代理人欄不可空白!!!", vbExclamation
         Else
            strCusTemp = Combo1
            'Add by Morgan 2008/2/21 加判斷是否為聯絡人
            If InStr(strCusTemp, "-") > 0 Then
               If ClsPDGetContact(strCusTemp, strTemp) Then
                  Combo1 = strCusTemp
                  lblAgent.Caption = strTemp
                  CheckKeyIn = 1
               End If
            
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetAgent(strCusTemp, strTemp) Then
            ElseIf ClsPDGetAgent(strCusTemp, strTemp) Then
               Combo1 = strCusTemp
               lblAgent.Caption = strTemp
               CheckKeyIn = 1
            End If
         End If
         
      Case 3, 8
         If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
            CheckKeyIn = 1
         Else
            ShowMsg MsgText(1038)
         End If
         
      Case 5
         '2005/7/8 MODIFY BY SONIA 加判斷有值才做
         If txtCaseField(intIndex) <> "" Then
            strCusTemp = txtCaseField(intIndex)
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetCustomer(strCusTemp, strTemp, strTemp1) Then
            If ClsPDGetCustomer(strCusTemp, strTemp, strTemp1) Then
               txtCaseField(intIndex) = strCusTemp
               lblNotify = strTemp
               CheckKeyIn = 1
            End If
         Else
            CheckKeyIn = 1
         End If
         
      Case 6
         If txtCaseField(intIndex) = "" Then
            CheckKeyIn = 0
         ElseIf CheckIsTaiwanDate(txtCaseField(intIndex)) Then
            If CheckReKey(txtCaseField(intIndex)) Then
               CheckKeyIn = 1
            Else
               CheckKeyIn = 0
            End If
         End If
         
      Case Else
         CheckKeyIn = 1
         
   End Select
   
End Function

Private Sub txtCaseField_GotFocus(Index As Integer)
   txtCaseField(Index).SelStart = 0
   txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
   '儲存未修改前之值至Tag中,供再確認時使用
   txtCaseField(Index).Tag = txtCaseField(Index)
End Sub

Private Sub textYear_Validate(Cancel As Boolean)
   If optSel(0).Value = True Then
      '香港(013)或日本(011)19940101以前申請的不抓國家檔設定
      If field(9) = "013" Or (field(9) = "011" And Val(field(10)) < 19940101) Then
         If "" & field(72) <> "" Then
            '香港(013)或日本(011)19940101以前申請的不抓國家檔設定
            If Val(textYear) <> varPA72(UBound(varPA72)) + 1 Then
               MsgBox "請輸入正確的繳費年度 !", vbCritical
               Cancel = True
            End If
         End If
      'Add by Amy 2018/03/20 +612年費移作次年判斷為空時不檢查,留至最後檢查
      ElseIf cp(10) = "612" And Trim(textYear) = MsgText(601) Then
        Exit Sub
      Else
         If Val(textYear) <> varFeeYears(UBound(varPA72) + 1) Then
            MsgBox "請輸入正確的繳費年度 !", vbCritical
            Cancel = True
         End If
      End If
   End If
   If Cancel Then TextInverse textYear
End Sub

Private Sub textNum_Validate(Cancel As Boolean)
   If optSel(1).Value = True Then
      'Add by Morgan 2006/8/21 馬來西亞新型除年費外還需繳延展費
      If field(9) = "018" And field(8) = "2" And cp(10) = "607" Then
         '因延展費無繳費記錄所以無法正確檢查，存檔時會再提醒
         If Val(textNum) > 2 Then
            MsgBox "馬來西亞延展次數不可超過 2!", vbCritical
            Cancel = True
         End If
      'END 2006/8/21
      
      'Add by Morgan 2006/8/21 俄羅斯新型除年費外還需繳延展費
      ''Modified by Morgan 2015/10/15 俄羅斯2015/1/1修法所有未發證未延展新型案專用期都為10年不得延展
      'ElseIf field(9) = "233" And field(8) = "2" And cp(10) = "607" Then
      '   '因延展費無繳費記錄所以無法正確檢查，存檔時會再提醒
      '   If Val(textNum) > 1 Then
      '      MsgBox "俄羅斯延展次數不可超過 1!", vbCritical
      '      Cancel = True
      '   End If
      'end 2015/10/15
      'END 2006/8/21
      
      'Add by Morgan 2015/6/9 俄羅斯設計除年費外還需繳延展費
      'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
      'Modified by Morgan 2022/6/15 2015 2015/1/1以前提申案件除了延展費外仍要繳年費，2015/1/1以後提申案件僅須繳延展費--禧佩
      'ElseIf field(9) = "023" And field(8) = "3" And cp(10) = "607" Then
      ElseIf field(9) = "023" And field(8) = "3" And cp(10) = "607" And Val(field(10)) < 20150101 Then
      'end 2022/6/15
         '因延展費無繳費記錄所以無法正確檢查，存檔時會再提醒
         '20141001以前申請者自申請起15年,期滿可延1次
         If Val(field(10)) < 20141001 And Val(textNum) > 1 Then
            MsgBox "俄羅斯延展次數不可超過 1!", vbCritical
            Cancel = True
         '20141001以後申請者自申請起5年,期滿可延4次,每次5年,共25年
         ElseIf Val(field(10)) >= 20141001 And Val(textNum) > 5 Then
            MsgBox "俄羅斯延展次數不可超過 5!", vbCritical
            Cancel = True
         End If
      'END 2015/6/9
      'Add by Amy 2018/03/20 +612年費移作次年判斷為空時不檢查,留至最後檢查
      ElseIf cp(10) = "612" And Trim(textNum) = MsgText(601) Then
        Exit Sub
      ElseIf Val(textNum) <> UBound(varPA72) + 2 Then
         MsgBox "請輸入正確的次數 !", vbCritical
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse textNum
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 結合原含逗號的字串與新的資料
' Input : strData ==> 原始含逗號的字串
'         strValue ==> 所要插入的資料
'         nPosition ==> 所要插入的位置 (若 < 0 則表示依順序自動插入到新位置, 並傳回新位置)
' Output : 傳回結合後的字串
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CombineData(ByVal strData As String, ByVal strValue As String, ByRef nPosition As Integer) As String
Dim nDotCount As Integer
Dim nDot As Integer
Dim nPos As Integer
Dim strRet As String
Dim bFind As Boolean
Dim aryData
Dim aryRet
   
   ' 計算逗號的總數(幾格)
   nDotCount = 0
   For nPos = 1 To Len(strData)
      If Mid(strData, nPos, 1) = "," Then nDotCount = nDotCount + 1
   Next nPos
   
   ' 將原始資料轉成排好位置的字串
   aryData = Split(strData, ",")
   For nPos = 0 To UBound(aryData)
      If Not IsEmptyText(aryData(nPos)) Then
         If Not IsEmptyText(strRet) Then strRet = strRet & ","
         strRet = strRet & aryData(nPos)
      End If
   Next nPos
   
   ' 補足原始資料中逗號的總數
   nDot = 0
   For nPos = 1 To Len(strRet)
      If Mid(strRet, nPos, 1) = "," Then nDot = nDot + 1
   Next nPos
   
   If nDot < nDotCount Then
      strRet = strRet & String(nDotCount - nDot, ",")
   End If
   
   If nPosition > 0 Then
      nDot = 0
      For nPos = 1 To Len(strRet)
         If Mid(strRet, nPos, 1) = "," Then nDot = nDot + 1
      Next nPos
      If nDot < nPosition Then
         strRet = strRet & String(nPosition - nDot, ",")
      End If
   End If
   
   If nPosition >= 0 Then
      aryRet = Split(strRet, ",")
      If nPosition <= UBound(aryRet) Then
         aryRet(nPosition) = strValue
         strRet = Empty
         For nPos = 0 To UBound(aryRet)
            If nPos > 0 Then strRet = strRet & ","
            strRet = strRet & aryRet(nPos)
         Next nPos
      Else
         If Not IsEmptyText(strRet) Then strRet = strRet & ","
         strRet = strRet & strValue
      End If
   Else
      ' 將新的資料插入到回傳的資料中
      bFind = False
      aryRet = Split(strRet, ",")
      For nPos = 0 To UBound(aryRet)
         If IsEmptyText(aryRet(nPos)) Then
            bFind = True
            aryRet(nPos) = strValue
            ' 所插入的位置
            nPosition = nPos
            Exit For
         End If
      Next nPos
      If bFind Then
         strRet = Empty
         For nPos = 0 To UBound(aryRet)
            If nPos > 0 Then strRet = strRet & ","
            strRet = strRet & aryRet(nPos)
         Next nPos
      Else
         nPosition = UBound(aryRet) + 1
         If Not IsEmptyText(strRet) Then strRet = strRet & ","
         strRet = strRet & strValue
      End If
   End If
   
   CombineData = strRet
End Function

' 依所輸入的數值取得下一個數值
Private Function GetNextValue(ByVal strData As String, ByVal strCurrData As String, ByRef nNext As Integer) As String
Dim aryData
Dim strNextData As String
Dim nPos As Integer
   
   strNextData = Empty
   nNext = 0
   aryData = Split(strData, ",")
   For nPos = 0 To UBound(aryData)
      If aryData(nPos) = strCurrData Then
         If nPos < UBound(aryData) Then
            strNextData = aryData(nPos + 1)
            nNext = nPos + 1
         End If
      ElseIf Val(aryData(nPos)) > Val(strCurrData) Then
         strNextData = aryData(nPos)
         Exit For
      End If
   Next nPos
   
   nNext = nNext + 1
   GetNextValue = strNextData
End Function

' 取得以逗號分隔的字串中最大數值的一筆資料及其所在的位置
Private Sub GetLastData(ByVal strData As String, ByRef strVal As String, ByRef nLastIndex As Integer)
Dim aryData
Dim nPos As Integer
Dim strMax As String
Dim nCount As Integer
   
   nCount = 0
   aryData = Split(strData, ",")
   For nPos = 0 To UBound(aryData)
      If Not IsEmptyText(aryData(nPos)) Then
         If aryData(nPos) > strMax Then
            strMax = aryData(nPos)
         End If
         nCount = nCount + 1
      End If
   Next nPos
   strVal = strMax
   nLastIndex = nCount
End Sub

'Removed by Morgan 2015/6/9 沒再用了
'' 取得開始繳費的日期及繳費年度串列
'Private Function GetFeeStartDate(ByRef strStartDate As String, ByRef strFeeYears As String) As Boolean
'   Dim strSql As String
'   Dim rsTmp As ADODB.Recordset
'   Dim strNA01 As String
'
'   '
'   GetFeeStartDate = False
'   ' 申請國家
'   strNA01 = field(9)
'   ' 開始計算的日期
'   strStartDate = Empty
'   ' 繳費年度
'   strFeeYears = Empty
'   Select Case field(8)
'      ' 發明
'      Case "1":
'         strSql = "SELECT NA06,NA21 FROM NATION " & _
'                  "WHERE NA01 = '" & strNA01 & "' "
'      ' 新型
'      Case "2":
'         strSql = "SELECT NA08,NA23 FROM NATION " & _
'                  "WHERE NA01 = '" & strNA01 & "' "
'      ' 設計
'      Case "3":
'         strSql = "SELECT NA10,NA25 FROM NATION " & _
'                  "WHERE NA01 = '" & strNA01 & "' "
'      Case Else:
'         Exit Function
'   End Select
'   Set rsTmp = New ADODB.Recordset
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      ' 取得開始計算繳費的日期
'      If Not IsNull(rsTmp.Fields(0)) Then
'         Select Case rsTmp.Fields(0)
'            ' 收文日
'            Case "1": strStartDate = DBDATE(cp(5))
'            ' 申請日
'            Case "2": strStartDate = DBDATE(field(10))
'            ' 發文日
'            Case "3": strStartDate = DBDATE(cp(27))
'            ' 准駁日
'            Case "4": strStartDate = DBDATE(field(20))
'            ' 公告日
'            Case "5": strStartDate = DBDATE(field(14))
'            ' 發證日
'            Case "6": strStartDate = DBDATE(field(21))
'            ' 公開日
'            Case "7": strStartDate = DBDATE(field(12))
'         End Select
'      End If
'      ' 取得繳費的年度字串
'      If IsNull(rsTmp.Fields(1)) = False Then
'         strFeeYears = rsTmp.Fields(1)
'      End If
'      '2005/9/23 ADD BY SONIA 各國修法
'      Select Case strNA01
'         Case "042"   '越南
'            Select Case field(8)
'               Case "1"
'                  If Val(DBDATE(field(21))) < 20031127 Then strStartDate = DBDATE(field(10))
'               Case "2"
'                  If Val(DBDATE(field(21))) < 20031127 Then strStartDate = DBDATE(field(10))
'            End Select
'         Case "012"   '韓國
'            Select Case field(8)
'               Case "2"
'                  If Val(DBDATE(field(10))) < 19990701 Then
'                     strFeeYears = "2,3,4,5,6,7,8,9,10,11,12,13,14,15"
'                  '2008/9/25 ADD BY SONIA
'                  ElseIf Val(DBDATE(field(10))) < 20061001 Then
'                     strFeeYears = "2,4,5,6,7,8,9,10"
'                  End If
'                  '2008/9/25 END
'               Case "3"
'                  If Val(DBDATE(field(10))) < 19980301 Then strFeeYears = "4,5,6,7,8,9,10"
'                  If Val(DBDATE(field(10))) < 20140701 Then strFeeYears = "4,5,6,7,8,9,10,11,12,13,14,15"   '2015/4/16 ADD BY SONIA 舊法為發證日起15年
'            End Select
'         Case "011"   '日本
'            Select Case field(8)
'               Case "2"
'                  If Val(DBDATE(field(10))) < 20050331 Then strFeeYears = "4,5,6"
'               'Add by Morgan 2007/4/20
'               '日本設計自申請日2007/4/1起改用新法,舊法為發證日起15年,第4年起逐年繳交。
'               Case "3"
'                  If Val(DBDATE(field(10))) < 20070401 Then strFeeYears = "4,5,6,7,8,9,10,11,12,13,14,15"
'            End Select
'         Case "211"   '西班牙
'            Select Case field(8)
'               Case "3"
'                  If Val(DBDATE(field(10))) < 20030709 Then strFeeYears = "5,10,15"
'            End Select
'         Case "015"   '澳洲
'            Select Case field(8)
'               Case "3"
'                  If Val(DBDATE(field(10))) < 20040617 Then
'                     'Modify by Morgan 2009/11/5 還原,尚有4個案件但都只剩第3次
'                     ''Modify by Morgan 2005/6/13
'                     ''strFeeYears = "1,6,11"
'                     'strFeeYears = "1,5,10"
'                     'strStartDate = DBDATE(field(21))
'                     strFeeYears = "1,6,11"
'                  End If
'            End Select
'         'Add by Morgan 2006/8/21
'         Case "018" '馬來西亞
'            '新型延展費
'            If field(8) = "2" And cp(10) = "607" Then
'               'Modify by Morgan 2007/7/18 新型不管何時提申均適用上述新法(2003年8月14日實施)
'               ''舊法(發證日<2001/8/1)自發證日起5年，期滿可延展2次，每次5年；
'               'If Val(DBDATE(field(21))) < 20010801 Then
'               '   strFeeYears = "5,10"
'               '   strStartDate = DBDATE(field(21))
'               'Else
'               '新法自申請日起10年，期滿可延展2次，每次5年；
'               '   strFeeYears = "10,15"
'               '   strStartDate = DBDATE(field(10))
'               'End If
'               strFeeYears = "10,15"
'               strStartDate = DBDATE(field(10))
'               'end 2007/7/18
'
'            End If
'
'         'Add by Morgan 2007/7/31
'         Case "233" '俄羅斯
'            '新型延展費:自申請日起5年,可延展3年
'            If field(8) = "2" And cp(10) = "607" Then
'               '2008/10/6 MODIFY BY SONIA 2008/1/1修法改申請日起10年
'               'strFeeYears = "5"
'               strFeeYears = "10"
'               strStartDate = DBDATE(field(10))
'            End If
'      End Select
'      '2005/9/23 END
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'
'   If Not IsEmptyText(strStartDate) And Not IsEmptyText(strFeeYears) Then
'      GetFeeStartDate = True
'   End If
'End Function

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim checkSw As Boolean
Dim nResponse
   
   'add by nickc 2008/05/01
   If IsDebt(field(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        Exit Function
   End If
   checkSw = False
   CheckDataValid = False
      If optSel(0).Value Then
         If IsEmptyText(textYear) Then
            strTit = "檢核資料"
            strMsg = "請輸入第幾年"
            nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
            textYear.SetFocus
            GoTo EXITSUB
         Else
            textYear_Validate checkSw
            If checkSw = True Then
               textYear.SetFocus
               GoTo EXITSUB
            End If
         End If
      Else
         If IsEmptyText(textNum) Then
            strTit = "檢核資料"
            strMsg = "請輸入第幾次"
            nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
            textNum.SetFocus
            GoTo EXITSUB
         Else
            textNum_Validate checkSw
            If checkSw = True Then
               textNum.SetFocus
               GoTo EXITSUB
            End If
         End If
      End If

   CheckDataValid = True
EXITSUB:
End Function

Private Sub textYear_GotFocus()
   InverseTextBox textYear
End Sub

Private Sub textNum_GotFocus()
   InverseTextBox textNum
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   If Me.textNum.Enabled = True Then
      Cancel = False
      textNum_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      '2005/7/19 ADD BY SONIA
      Text1_Validate 1, Cancel
      If Cancel = True Then
         Exit Function
      End If
      '2005/7/19 END
   End If

   If Me.textYear.Enabled = True Then
      Cancel = False
      textYear_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      '2005/7/19 ADD BY SONIA
      Text1_Validate 0, Cancel
      If Cancel = True Then
         Exit Function
      End If
      '2005/7/19 END
   End If

   For Each objTxt In Me.txtCaseField
      If objTxt.Enabled = True Then
         Cancel = False
         txtCaseField_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   'Add by Morgan 2004/9/14
   If Combo1.Enabled = True Then
      Cancel = False
      Combo1_Validate Cancel
      If Cancel = True Then
         Combo1.SetFocus
         Exit Function
      End If
   End If

   'Added by Morgan 2018/9/12 CFP電子化-接洽單檢查
   If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
      If cp(9) < "B" And Left(cp(12), 1) <> "F" Then
         If PUB_CheckPDF3(cp(1), cp(2), cp(3), cp(4)) = False Then
            Exit Function
         End If
      End If
   End If
   'end 2018/9/12
   
   'Added by Morgan 2019/5/2
   '日本發案明年費(第4-10年)發文要設定減免身分
   m_strCP81 = ""
   m_strJpMemo = ""
   If field(9) = "011" And field(8) = "1" And Val(textYear) <= 10 Then
      If PUB_ChkJpDiscount(cp(1), cp(2), cp(3), cp(4), True) = True Then
         Dim stAD10 As String, stAD15 As String
         For ii = 1 To 5
            If field(25 + ii) <> "" Then
               strExc(1) = PUB_GetAD03(field(25 + ii), "011", stAD10, , stAD15)
               m_strJpMemo = m_strJpMemo & PUB_GetJpDiscountDesc(stAD10, stAD15) & ";"
               If strExc(1) = "" Then
                  'Modified by Morgan 2019/6/19 改詢問是否不可減免,若是則系統自動設定--禧佩
                  'MsgBox "申請人【" & field(25 + ii) & " " & GetCustomerNameAndState(field(25 + ii)) & "】尚未設定減免身分不可發文！", vbCritical, "日本年費發文減免身分檢查"
                  'Exit Function
                  If MsgBox("申請人【" & field(25 + ii) & " " & GetCustomerNameAndState(field(25 + ii)) & "】尚未設定減免身分！" & vbCrLf & vbCrLf & "是否要設定為【不可減免】？", vbYesNo + vbDefaultButton2 + vbExclamation, "日本實審發文減免身分檢查") = vbYes Then
                     PUB_SetNoDisc field(25 + ii), field(9)
                     m_strCP81 = "N"
                  Else
                     Exit Function
                  End If
                  'end 2019/6/19
               ElseIf m_strCP81 <> "N" Then
                  m_strCP81 = strExc(1)
               End If
            End If
         Next
         If m_strCP81 <> "Y" Then m_strJpMemo = ""
      End If
   End If
   'end 2019/5/2
   
   'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
   If Pub_ChkACS112isNull(field(1), field(2), field(3), field(4), txtCP113) = True Then
       txtCP113.SetFocus
       txtCP113_GotFocus
       Exit Function
   End If
   'end 2021/05/25

   TxtValidate = True
End Function

'Added by Lydia 2021/05/25
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
