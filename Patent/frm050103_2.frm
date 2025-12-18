VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050103_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人案件提申"
   ClientHeight    =   4596
   ClientLeft      =   -2688
   ClientTop       =   3216
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4596
   ScaleWidth      =   8520
   Begin VB.CheckBox Check1 
      Caption         =   "工程師判發"
      Height          =   225
      Left            =   6795
      TabIndex        =   51
      Top             =   2670
      Width           =   1590
   End
   Begin VB.TextBox txtFavDt 
      Height          =   270
      Left            =   7500
      MaxLength       =   7
      TabIndex        =   45
      Top             =   3870
      Width           =   885
   End
   Begin VB.CommandButton cmdPriority 
      Caption         =   "輸入(&P)"
      Height          =   270
      Left            =   1215
      TabIndex        =   44
      Top             =   4230
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7596
      TabIndex        =   18
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5544
      TabIndex        =   16
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6372
      TabIndex        =   17
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.CheckBox chkChoose 
      Height          =   300
      Index           =   1
      Left            =   5670
      TabIndex        =   11
      Top             =   3270
      Width           =   735
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "1296;529"
      Value           =   "0"
      Caption         =   "圖式"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.CheckBox chkChoose 
      Height          =   300
      Index           =   2
      Left            =   90
      TabIndex        =   8
      Top             =   3270
      Width           =   1980
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3492;529"
      Value           =   "0"
      Caption         =   "受理通知書 / 收據"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   17
      Left            =   7530
      TabIndex        =   56
      Top             =   2340
      Width           =   270
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "476;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   16
      Left            =   5340
      TabIndex        =   54
      Top             =   2340
      Width           =   270
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "476;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   15
      Left            =   3360
      TabIndex        =   52
      Top             =   2340
      Width           =   270
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "476;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   14
      Left            =   7110
      TabIndex        =   50
      Top             =   4200
      Width           =   1275
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "2249;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   13
      Left            =   7110
      TabIndex        =   12
      Top             =   2940
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   12
      Left            =   5400
      TabIndex        =   1
      Top             =   1710
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
      Height          =   300
      Index           =   11
      Left            =   7470
      TabIndex        =   13
      Top             =   3270
      Visible         =   0   'False
      Width           =   870
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1535;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   10
      Left            =   4965
      TabIndex        =   38
      Top             =   3870
      Visible         =   0   'False
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
      Height          =   300
      Index           =   9
      Left            =   4860
      TabIndex        =   7
      Top             =   2970
      Width           =   270
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "476;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   8
      Left            =   1320
      TabIndex        =   0
      Top             =   1710
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
      Height          =   300
      Index           =   7
      Left            =   1560
      TabIndex        =   15
      Top             =   3870
      Width           =   495
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "873;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   5
      Left            =   2340
      TabIndex        =   10
      Top             =   3270
      Width           =   2415
      VariousPropertyBits=   671107097
      Size            =   "4260;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1140
      TabIndex        =   27
      Top             =   840
      Width           =   7335
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12938;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   4
      Left            =   1920
      TabIndex        =   6
      Top             =   2970
      Width           =   270
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "476;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   3
      Left            =   5400
      TabIndex        =   5
      Top             =   2670
      Width           =   1095
      VariousPropertyBits=   671107101
      MaxLength       =   8
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
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
      Height          =   300
      Index           =   1
      Left            =   5400
      TabIndex        =   3
      Top             =   2040
      Width           =   2415
      VariousPropertyBits=   671107099
      MaxLength       =   25
      Size            =   "4260;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   4
      Top             =   2670
      Width           =   3135
      VariousPropertyBits=   671107099
      MaxLength       =   50
      Size            =   "5530;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   6
      Left            =   720
      TabIndex        =   14
      Top             =   3570
      Width           =   7650
      VariousPropertyBits=   671107099
      Size            =   "13494;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      Caption         =   "實審及指定費：     （Y/N)"
      Height          =   255
      Left            =   6270
      TabIndex        =   57
      Top             =   2370
      Width           =   2115
   End
   Begin VB.Label Label12 
      Caption         =   "主張優先權：     （Y/N)"
      Height          =   255
      Left            =   4260
      TabIndex        =   55
      Top             =   2370
      Width           =   1965
   End
   Begin VB.Label Label4 
      Caption         =   "是否已一併提出：實審/合併實審及檢索：    （Y/N)"
      Height          =   255
      Left            =   90
      TabIndex        =   53
      Top             =   2370
      Width           =   4125
   End
   Begin VB.Label lbl118Fee 
      AutoSize        =   -1  'True
      Caption         =   "提出英譯本及聲明書之費用："
      Height          =   180
      Left            =   4680
      TabIndex        =   49
      Top             =   4290
      Width           =   2340
   End
   Begin VB.Label lblFavDt 
      AutoSize        =   -1  'True
      Caption         =   "優惠期日期："
      Height          =   180
      Left            =   6435
      TabIndex        =   48
      Top             =   3915
      Width           =   1080
   End
   Begin VB.Label lblItemCnt 
      Caption         =   "項數："
      Height          =   255
      Left            =   6120
      TabIndex        =   47
      Top             =   2970
      Width           =   975
   End
   Begin VB.Label Label37 
      Caption         =   "優先權資料："
      Height          =   255
      Left            =   90
      TabIndex        =   46
      Top             =   4230
      Width           =   1095
   End
   Begin VB.Label Label20 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "PCT國家階段提交日："
      Height          =   180
      Left            =   3600
      TabIndex        =   43
      Top             =   1755
      Width           =   1755
   End
   Begin VB.Label lbl416Fee 
      Caption         =   "實審費用："
      Height          =   255
      Left            =   6525
      TabIndex        =   42
      Top             =   3270
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblCaseField 
      AutoSize        =   -1  'True
      Height          =   225
      Index           =   3
      Left            =   1140
      TabIndex        =   41
      Top             =   1440
      Width           =   7245
   End
   Begin VB.Label Label6 
      Caption         =   "巳繳年費："
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   40
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "美國預定公開日："
      Height          =   255
      Index           =   3
      Left            =   3465
      TabIndex        =   39
      Top             =   3870
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "是否修改通知函內容：     （Y:Word）"
      Height          =   255
      Left            =   3060
      TabIndex        =   37
      Top             =   2970
      Width           =   2970
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "櫃台收文日："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   36
      Top             =   1755
      Width           =   1080
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   5310
      TabIndex        =   24
      Top             =   540
      Width           =   405
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "提申日："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   35
      Top             =   2085
      Width           =   720
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Index           =   0
      Left            =   4455
      TabIndex        =   34
      Top             =   2085
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號："
      Height          =   180
      Left            =   120
      TabIndex        =   33
      Top             =   2715
      Width           =   900
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "下次繳費日："
      Height          =   180
      Left            =   4275
      TabIndex        =   32
      Top             =   2715
      Width           =   1080
   End
   Begin VB.Label Label9 
      Caption         =   "附件："
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   3570
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "是否列印客戶通知函：     （N:不印）"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2970
      Width           =   2925
   End
   Begin VB.Label Label11 
      Caption         =   "有無前案可提供：          （N：無）"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3870
      Width           =   2895
   End
   Begin VB.Label Label24 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   975
   End
   Begin MSForms.Label lblCountryName 
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   540
      Width           =   2565
      VariousPropertyBits=   27
      Size            =   "4524;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   540
      Width           =   945
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1140
      TabIndex        =   23
      Top             =   540
      Width           =   3105
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   1140
      Width           =   2535
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1140
      TabIndex        =   19
      Top             =   1140
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "專利種類："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   540
      Width           =   975
   End
   Begin MSForms.CheckBox chkChoose 
      CausesValidation=   0   'False
      Height          =   285
      Index           =   0
      Left            =   2100
      TabIndex        =   9
      Top             =   3278
      Width           =   3375
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "5953;503"
      Value           =   "0"
      Caption         =   "                                                      說明書"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm050103_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/6 改成Form2.0 (txtCaseField,cboCaseName,lblCountryName,chkChoose...)
'Modified by Morgan 2021/12/7 chkChoose 改 Form2.0 後 Value 判斷也要改為 Trure 或 False
'Memo By Morgan 2012/12/13 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'2005/7/11整理
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
Dim strReceiveNo As String
'Add By Cheng 2002/02/08
Dim m_strCP09ByCheng As String
'91.12.31 add by sonia 是否新增實體審查期限
Dim m_416YN As String
'91.12.31 end
Dim m_NP07 As String
'Add by Morgan 2005/6/17 控制只觸發一次
Dim m_bolActive As Boolean
'Add by Morgan 2005/10/27
Dim m_bolNewApp As Boolean '申請案
'strPriority存放優先權
'Add by Morgan 2005/11/17
Dim strPriority(1 To 5) As String '優先權資料
Dim m_bolRePriDate As Boolean '優先權資料需重新輸入
'Modified by Morgan 2025/3/12 +125
Const m_RefCP10List As String = "101,102,103,104,105,109,113,114,115,120,122,125,301,302,303,305,306,307" '通知申請案號相關案件性質 Add by Morgan 2010/6/14
Dim m_bol118C As Boolean, m_str118Date As String 'Added by Lydia 2015/05/20 美國暫時申請案(118)確認是否為中文送件
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim strCP10 As String 'Add by Morgan 2016/5/30 'Modified by Morgan 2018/6/8 改全域
Dim m_NewCP09 As String, m_bolAddLP As Boolean 'Added by Morgan 2018/6/8
Dim bolLastTime As Boolean 'Added by Moragn 2019/7/10 是否最後一次年費/維持費/延展費
Dim m_bolXCACase As Boolean 'Added by Morgan 2019/11/28 母案非本所之接續案
Dim m_OtherCP10 As String 'Added by Morgan 2020/3/11 一併提出實審性質(416/427)
Dim m_605NP08 As String 'Added by Morgan 2021/3/15 下次繳費日所限
Dim m_bolEPC_DE_CHK As Boolean 'Added by Morgan 2024/4/25

Private Sub chkChoose_Click(Index As Integer)
   Dim strTemp As String

   If Index = 0 Then
      txtCaseField(5).Enabled = chkChoose(Index).Value
      If chkChoose(Index).Value Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(lblCaseField(1), , strTemp) Then
         If ClsPDGetNation(lblCaseField(1), , strTemp) Then
            txtCaseField(5) = strTemp
         End If
      Else
         txtCaseField(5) = ""
      End If
   'Add by Morgan 2005/10/27
   '申請案且勾收據時預設要出通知函
   'Modified by Morgan 2020/8/17 +年費
   'ElseIf Index = 2 And m_bolNewApp Then
   ElseIf Index = 2 And (m_bolNewApp Or (cp(10) >= "605" And cp(10) <= "607")) Then
   'end 2020/8/17
      If chkChoose(Index).Value Then
         txtCaseField(4) = ""
         
      'Added by Morgan 2020/8/17 年費沒收據不出定稿
         txtCaseField(4).Enabled = True
      ElseIf cp(10) >= "605" And cp(10) <= "607" Then
         txtCaseField(4) = "N"
         txtCaseField(4).Enabled = False
      'end 2020/8/17
      End If
   End If

End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 'Modified by Lydia 2015/05/20 strTxt 10 = > 11
 Dim strTxt(1 To 14) As String, iStep As Integer, strTmp As String
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   Dim Jjj As Integer
   
   'Modified by Morgan 2016/10/28 收據改為申請受理通知書--玫音
   strTmp = ""
   If chkChoose(2).Value Then strTmp = strTmp & "申請受理通知書、"
   If chkChoose(0).Value Then strTmp = strTmp & txtCaseField(5) & "說明書、"
   'Modify by Morgan 2009/12/21 圖式放最後
   If chkChoose(1).Value Then strTmp = strTmp & "圖式"
   '邱小姐說一般的格式時沒有例外欄位
   '********* 90.11.14   nick
   Jjj = 1
   'Modify by Morgan 2005/12/26
   'If ET03 < "20" Then
   If ET03 < "20" Or ET03 >= "30" Then
      '************** 90.11.14   nick
      '錯誤型態
      'If Right(strTmp, "、") Then strTmp = Left(strTmp, Len(strTmp) - 1)
      If Right(strTmp, 1) = "、" Then strTmp = Left(strTmp, Len(strTmp) - 1)
      'Added by Lydia 2015/05/20 +官方通知,約定期限=控管正式申請案之約定期限(本所期限-28天且須為工作天)
      If m_bol118C = True Then
         strTmp = strTmp & "及官方通知"
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','約定期限','" & m_str118Date & "')"
         Jjj = Jjj + 1
      End If
      
      If CheckStr(strTmp) <> "" Then
         '91.11.10 modify by sonia
         'strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         '   "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         '   "','是否有說明書、圖式、申請收據','" & strTmp & "')"
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','是否有說明書、圖式、申請收據','茲隨函檢附本案" & strTmp & "各一份，敬請查收備存。')"
          '91.11.10 end
        Jjj = Jjj + 1
      End If
      If CheckStr(txtCaseField(6)) <> "" Then
         '91.11.10 modify by sonia
         'strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         '   "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         '   "','附件內容'," & CNULL(txtCaseField(6)) & ")"
         'Modify by Morgan 2009/6/23
         'strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','附件內容','茲隨函檢附" & CNULL(txtCaseField(6)) & "，以供查照。')"
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','附件內容'," & CNULL("茲隨函檢附" & txtCaseField(6) & "，以供查照。") & ")"
         'end 2009/6/23
         '91.11.10 end
         Jjj = Jjj + 1
      End If

      If CheckStr(txtCaseField(10)) <> "" Then
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','其他公告日'," & CNULL(DBDATE(txtCaseField(10))) & ")"
         Jjj = Jjj + 1
      End If
      
   End If
   
   If txtCaseField(3) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','下次繳年費日'," & CNULL(TransDate(txtCaseField(3), 2)) & ")"
      Jjj = Jjj + 1
   
      'Added by Morgan 2021/3/15
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','下次繳年費日所限'," & CNULL(m_605NP08) & ")"
      Jjj = Jjj + 1
      'end 2021/3/15
      
   'Added by Morgan 2019/3/6
   'EPC子案要抓母案的期限 Ex:CFP-012758-03--玫音
   ElseIf cp(4) <> "00" Then
      strExc(0) = "select np09 from nextprogress where np02=" + CNULL(cp(1)) + " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05='00' and np07='" & cp(10) & "' AND np06 IS null order by np09 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','下次繳年費日'," & RsTemp(0) & ")"
         Jjj = Jjj + 1
      End If
   'end 2019/3/6
   End If
   
   If cp(10) = 年費 Or cp(10) = 維持費 Or cp(10) = 延展費 Then
      
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','第幾年至幾年費','" & GetPayYear & "')"
      
      'Added by Morgan 2021/3/15 寶齡富錦 Y55435 案件
      If field(75) = "Y55435" Then
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','第幾年至幾年費/英','" & GetPayYear(True) & "')"
      End If
                        
      '2008/9/19 ADD BY SONIA
      '019泰國新型第7年年費發文,通知函改為第7-8年,因第8年不用繳,第9年年費發文,通知函改為第9-10年
      If field(9) = "019" And field(8) = "2" Then
         Select Case GetPayYear
            Case "第7年"
               strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','第幾年至幾年費','第7至8年')"
            Case "第9年"
               strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','第幾年至幾年費','第9至10年')"
         End Select
      End If
      '012韓國新型第2年年費發文,通知函改為第2-3年,因第3年不用繳
      If field(9) = "012" And field(8) = "2" Then
         Select Case GetPayYear
            Case "第2年"
               strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','第幾年至幾年費','第2至3年')"
         End Select
      End If
      '2008/9/19 END
      '92.9.17 ADD BY SONIA
      If field(9) = "101" Then
         Select Case GetPayYear
            Case "第1次"
               strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','第幾年至幾年費','" & GetPayYear & "(第5~8年)')"
            Case "第2次"
               strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','第幾年至幾年費','" & GetPayYear & "(第9~12年)')"
            Case "第3次"
               strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','第幾年至幾年費','" & GetPayYear & "(第13年~專用期屆滿)')"
         End Select
      End If
      '92.9.17 END
      '2007/6/1 ADD BY SONIA
      If field(9) = "231" Then
         Select Case field(8)
            Case "2"
               Select Case GetPayYear
                  Case "第1次"
                     strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                        "','第幾年至幾年費','" & GetPayYear & "(第4~6年)')"
                  Case "第2次"
                     strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                        "','第幾年至幾年費','" & GetPayYear & "(第7~8年)')"
                  Case "第3次"
                     strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                        "','第幾年至幾年費','" & GetPayYear & "(第9年~專用期屆滿)')"
               End Select
            Case "3"
               Select Case GetPayYear
                  Case "第1次"
                     strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                        "','第幾年至幾年費','" & GetPayYear & "(第6~10年)')"
                  Case "第2次"
                     strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                        "','第幾年至幾年費','" & GetPayYear & "(第11~15年)')"
                  Case "第3次"
                     strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                        "','第幾年至幾年費','" & GetPayYear & "(第16~20年)')"
                  Case "第4次"
                     strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                        "','第幾年至幾年費','" & GetPayYear & "(第21年~專用期屆滿)')"
               End Select
         End Select
      End If
      '2007/6/1 END
      Jjj = Jjj + 1
      
      'Add by Morgan 2005/1/4 EPC指定國家
      If lblCaseField(1) = "221" Then
         'Modified by Morgan 2018/5/24 要抓有繳年費的國家
         'strExc(0) = GetFeeCountry
         If ClsPDReadCountry(intCaseKind, field, strExc(0), True, True, True) Then
         'end 2018/5/24
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','EPC指定國家','" & strExc(0) & "')"
               Jjj = Jjj + 1
         End If
      End If
   End If
   
   '實審報價敘述
   If m_416YN = "Y" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " SELECT '" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','實審法限',NP09 FROM NEXTPROGRESS WHERE NP02='" & field(1) & "'" & _
         " AND NP03='" & field(2) & "' AND NP04='" & field(3) & "' AND NP05='" & field(4) & "'" & _
         " AND NP06 IS NULL AND NP07='416'"
      Jjj = Jjj + 1
         
      'Modifty by Amy 2013/05/22 實審費用不需於定稿帶出且不需輸入
      'Add by Morgan 2010/6/14
      'strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      '   "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      '   "','是否印實審報價','♀')"
      '   Jjj = Jjj + 1
         
            'Add by Morgan 2005/6/17
      'If txtCaseField(11).Visible = True Then
      '      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      '         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      '         "','實審費用'," & Val(txtCaseField(11)) & ")"
      '      Jjj = Jjj + 1
      'Else
      '   'Add By Cheng 2002/07/31
      '   'Modify by Morgan 2005/6/17代理人補9碼
      '   'strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & field(9) & "' AND YF03='" & cp(44) & "' AND YF04='416' AND YF05='1' "
      '   'Modify by Morgan 2007/3/5 加判斷專利種類
      '   'strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & field(9) & "' AND YF03='" & GetNewFagent(cp(44)) & "' AND YF04='416' AND YF05='1' "
      '   strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & field(9) & "' AND YF03='" & GetNewFagent(cp(44)) & "' AND YF04='416' AND YF05='1' AND YF02='" & field(8) & "' "
      '   'end 2007/3/5
      '   intI = 1
      '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      '   If intI = 1 Then
      '      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      '         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      '         "','實審費用'," & CNULL(RsTemp.Fields(0).Value) & ")"
      '      Jjj = Jjj + 1
      '   Else
      '      '92.4.3 Add By SONIA
      '      'Modify by Morgan 2007/3/5 加判斷專利種類
      '      'strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & field(9) & "' AND YF03='Y00000000' AND YF04='416' AND YF05='1' "
      '      strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & field(9) & "' AND YF03='Y00000000' AND YF04='416' AND YF05='1' AND YF02='" & field(8) & "' "
      '      'end 2007/3/5
      '      intI = 1
      '      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      '      If intI = 1 Then
      '         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      '            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      '            "','實審費用'," & CNULL(RsTemp.Fields(0).Value) & ")"
      '         Jjj = Jjj + 1
      '      End If
      '      '92.4.3 END
      '   End If
      'End If
      'end 2013/05/22
   End If
   
   'Added by Morgan 2015/12/16
   If txtCaseField(14).Enabled Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','費用','" & txtCaseField(14).Text & "')"
         Jjj = Jjj + 1
   End If
   'end 2015/12/16
   
   'Added by Morgan 2019/7/10
   If bolLastTime = True Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','" & lblCountryName & "最後一年要印','♀')"
         Jjj = Jjj + 1
   End If
   'end 2019/7/10
   
   'Added by Morgan 2019/10/2
   'EPC回覆檢索報告
   If field(9) = "221" And cp(10) = "218" And txtCaseField(17) = "Y" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','一併提申性質','、請求實體審查及繳交指定費')"
      Jjj = Jjj + 1
      
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','有一併提申','♀')"
      Jjj = Jjj + 1
      
   End If
   'end 2019/10/2
   'Added by Morgan 2020/3/11
   strExc(1) = ""
   If txtCaseField(15) = "Y" Then
      strExc(1) = IIf(InStr(m_OtherCP10, "427") > 0, "合併檢索及實審", "實審")
   End If
   If txtCaseField(16) = "Y" Then
      strExc(1) = "主張" & IIf(InStr(m_OtherCP10, "121") > 0, "國內", "國際") & "優先權" & IIf(strExc(1) <> "", "、", "") & strExc(1)
   End If
   If strExc(1) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','一併提申事宜','且已提出" & strExc(1) & "事宜')"
      Jjj = Jjj + 1
   End If
   'end 2020/3/11
   
    'Added by Morgan 2020/9/23
   '緬甸刊登廣告(通知已登報函)
   If field(9) = "048" And cp(10) = "951" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " select '" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','刊登廣告期限',np09" & _
         " from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "'" & _
         " and np05='" & cp(4) & "' and np06 is null and np07='951'"
      Jjj = Jjj + 1
   End If
   'end 2020/9/23
   
   'Added by Morgan 2024/1/15
   '英國設計案
   If field(9) = "201" And field(8) = "3" Then
      If PUB_ChkCPExist(cp, "1608") Then
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','再註冊才印','♀')"
         Jjj = Jjj + 1
      End If
   End If
   'end 2024/1/15
   
   'Added by Morgan 2025/9/26
   '阿拉伯發明新型不印實審期限
   If field(9) = "034" And (field(8) = "1" Or field(8) = "2") Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','阿拉伯發明新型不印','♀')"
      Jjj = Jjj + 1
   End If
   'end 2025/9/26
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(Jjj - 1, strTxt) Then
   If Not ClsLawExecSQL(Jjj - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   '*******************************************
End Sub

Private Sub cmdok_Click(Index As Integer)
 Dim i As Integer, strTmp As String, bolChk As Boolean

   Select Case Index
      Case 0, 3
      
         '2005/7/11 MODIFY BY SONIA 判斷PCT案
         'cp(47) = txtCaseField(0)
         'Modify by Morgan 2005/12/20 加控制新案
         'If field(46) <> "" Then
         'Modify by Morgan 2011/7/1
         'If m_bolNewApp And field(46) <> "" Then
         If txtCaseField(12).Visible = True Then
            cp(47) = txtCaseField(12)
         Else
            cp(47) = txtCaseField(0)
         End If
         '2005/7/11 END
         
         'Add by Morgan 2007/3/14 多國案若有其他相同案已核准、發證、公開時不可分案、發文、提申
         If cp(47) <> "" And (cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103") Then
            'Modify by Morgan 2008/2/21 +cp(47)
            If PUB_SameCaseCheck(cp, cp(47)) = False Then
               Exit Sub
            End If
         End If
         'end 2007/3/14
         Screen.MousePointer = vbHourglass
         'Modify By Cheng 2002/07/31
'         For i = 0 To 7
         For i = 0 To 10
            If txtCaseField(i).Visible Then
               If CheckKeyIn(i) <> 1 Then
                  '2005/7/11 MODIFY BY SONIA
                  'txtCaseField(i).SetFocus
                  'txtCaseField_GotFocus (i)
                  If i <> 3 Then
                     txtCaseField(i).SetFocus
                     txtCaseField_GotFocus (i)
                  End If
                  '2005/7/11 END
                  Exit For
               End If
            End If
         Next
         'Modify By Cheng 2002/07/31
'         If i = 8 Then
         If i = 11 Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
            
            'Add By Sindy 2020/7/20
            If m_strIR01 <> "" Then
               '下載信件檔,檢查信件是否開啟中,以免後面上傳卷宗區會無法儲存
               'Modify By Sindy 2022/11/10 + IIf(field(9) <> 台灣國家代號, "PAT", "RX")
               If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", IIf(field(9) <> 台灣國家代號, "PAT", "RX"), , True) = False Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            '2020/7/20 END
            
            'Add by Morgan 2005/11/17
            '有勾收據，且需輸優先權資料時
            If chkChoose(2).Value And cmdPriority.Enabled = True Then
               If strPriority(1) = "" Or m_bolRePriDate = False Then
                  MsgBox "本案有優先權資料,請重新輸入以便與原資料檢核！", vbCritical
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            'Add by Morgan 2009/5/13
            '檢查提申日期是否不同於指定提申或晚於最終提申
            If txtCaseField(0) <> "" Then
               strExc(0) = "select * from nextprogress where np01='" & cp(9) & "' and np06 is null and np07 in ('995','996')"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp("NP07") = "995" Then
                     If DBDATE(txtCaseField(0)) <> RsTemp("NP09") Then
                        If MsgBox("所輸入的提申日期【" & txtCaseField(0) & "】與指定提申期限【" & TransDate(RsTemp("NP09"), 1) & "】不同，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                           Screen.MousePointer = vbDefault
                           Exit Sub
                        End If
                     End If
                  Else
                     If Val(DBDATE(txtCaseField(0))) > Val(RsTemp("NP09")) Then
                        If MsgBox("所輸入的提申日期【" & txtCaseField(0) & "】晚於最終提申期限【" & TransDate(RsTemp("NP09"), 1) & "】，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                           Screen.MousePointer = vbDefault
                           Exit Sub
                        End If
                     End If
                  End If
               End If
            End If
            
            If SaveDatabase Then
               If txtCaseField(4) <> "N" Then
                  bolLastTime = False 'Added by Morgan 2019/7/10
                  Select Case cp(10)
                     '93.11.10 MODIFY BY SONIA
                     'Case 發明申請, CIP申請
                     'Modify by Morgan 2005/1/3 加改請發明301
                     '2006/9/12 MODIFY BY SONIA 加CA申請122
                     'Case 發明申請, CIP申請, 美國暫時申請
                     'Modify by Morgan 2009/3/18 +PCT申請
                     'Modify by Morgan 2011/6/27
                     'Case 發明申請, CIP申請, 美國暫時申請, 改請發明, "122", "109"
                     Case 美國暫時申請
                        'Added by Lydia 2015/05/20 美國暫時申請案(118)確認是否為中文送件
                        'Modified by Lydia 2015/11/23 官方通知是與收據一同發出
                        'If MsgBox("美國暫時申請案是否為中文送件?", vbYesNo + vbExclamation) = vbYes Then
                        strTmp = "00" '預設
                        If chkChoose(2).Value Then
                           'Modified by Morgan 2015/12/16 改在存檔前詢問以便檢查費用是否有輸入
                           'If MsgBox("美國暫時申請案是否為中文送件?", vbYesNo + vbExclamation) = vbYes Then
                           '   m_bol118C = True
                           If m_bol118C Then
                           'end 2015/12/16
                              strTmp = "01"
                           End If
'                        Else
'                           strTmp = "00" '原-預設
                        End If
                        'end 2015/11/23
                     'Modified by Morgan 2018/3/31 + CPA申請--禧佩
                     Case 發明申請, CIP申請, 改請發明, "122", PCT申請, CPA申請
                     '93.11.10 END
                        '若為PCT案
                        'Modify by Morgan 2008/1/24 希臘用一般定稿
                        'If field(46) = "Y" Then
                        If field(46) = "Y" And field(9) <> "212" Then
                           'Add by Morgan 2007/4/20 PCT進EPC
                           If field(9) = "221" Then
                              strTmp = "08"
                           'Add by Morgan 2007/9/5 PCT進美國
                           ElseIf field(9) = "101" Then
                              strTmp = "12"
                           'PCT進各國
                           Else
                              strTmp = "06"
                           End If
                        Else
                           Select Case field(9)
                              'Add by Morgan 2007/8/28
                              Case "101" '美國
                                 strTmp = "09"
                                 '若有輸入美國預定公開日
                                 If txtCaseField(10) <> "" Then
                                    strTmp = "05"
                                 End If
                                 
                              Case "207" '荷蘭
                                 strTmp = "01"
                                 
                              'Remove by Morgan 2010/6/15 定稿改共用(改用例外欄位控制)
                              'Case "011" '日本
                              '   If m_416YN = "Y" Then
                              '      strTmp = "02"  '未提實審
                              '   Else
                              '      strTmp = "00"  '已提實審
                              '   End If
                              'Case "012" '韓國
                              '   If m_416YN = "Y" Then
                              '      strTmp = "03"  '未提實審
                              '   Else
                              '      strTmp = "00"  '已提實審
                              '   End If
                              
                              'Case "102" '加拿大 '與慧汶確認後也可改用通函
                              '
                              '   If m_416YN = "Y" Then
                              '      strTmp = "04"  '未提實審
                              '   Else
                              '      strTmp = "07"  '已提實審
                              '   End If
                              'end 2010/6/15
                                 
                              'ADD BY SONIA 2014/11/4 改EPC定稿故獨立
                              Case "221"  'EPC
                                 strTmp = "02"
                              'END 2014/11/4
                              
                              Case Else '一般
                                 strTmp = "00"
                                 
                           End Select
                        End If
                        
                     'Modify by Morgan 2005/1/3 加改請新型302
                     'Modified by Morgan 2025/7/29 +集體設計申請105 Ex:CFP-035264-1 --禧佩
                     Case 新型申請, 改請新型, "105"
                        Select Case field(9)
                           'Add by Morgan 2007/8/28
                           Case "101" '美國
                              strTmp = "09"
                           'Case "207" '荷蘭  92.1.14 CALCEL BY SONIA
                           '   strTmp = "01"
                           'Case "012" '韓國  90.12.25 MODIFY BY SONIA
                           '   strTmp = "03"
                           'Case "102" '加拿大  92.1.16 CALCEL BY SONIA
                           '   strTmp = "07"
                           
                           'Remove by Morgan 2010/6/15 定稿改共用(改用例外欄位控制)
                           'Case "012" '韓國 'Add by Morgan 2006/12/12
                           '   If m_416YN = "Y" Then
                           '      strTmp = "03" '未提實審
                           '   Else
                           '      strTmp = "00" '已提實審
                           '   End If
                           'end 2010/6/15
                           
                           Case Else '一般
                              strTmp = "00"
                        End Select
                        'Modify By Cheng 2002/07/31
                        '若為PCT案
                        'Modify by Morgan 2008/1/24 希臘用一般定稿
                        'If field(46) = "Y" Then
                        If field(46) = "Y" And field(9) <> "212" Then
                           strTmp = "06"
                           
                           'Add by Morgan 2007/4/20 加EPC定稿
                           If field(9) = "221" Then
                              strTmp = "08"
                           End If
                           
                           'Removed by Morgan 2015/12/16 已抽出
                           'If cp(10) = "118" Then
                           '   strTmp = "01"
                           'End If
                           
                        End If
                        
                     'Modify by Morgan 2005/1/3 加改請設計303
                     'Modified by Morgan 2025/3/12 +衍生設計125
                     Case 設計申請, 改請設計, 衍生設計
                        'If field(9) = "102" Then  92.1.16 CALCEL BY SONIA
                        '   '加拿大 1
                        '   strTmp = "07"
                        'Else
                        
                        'Add by Morgan 2007/8/28
                        If field(9) = "101" Then
                           strTmp = "09"
                        Else
                        'end 2007/8/28
                           '一般
                           strTmp = "00"
                        End If
                        
                        'Modify By Cheng 2002/07/31
                        '若為PCT案
                        If field(46) = "Y" Then
                           strTmp = "06"
                           
                           'Removed by Morgan 2015/12/16 已抽出
                           'Select Case cp(10)
                           'Case "118"
                           '   strTmp = "01"
                           'End Select
                           
                        End If
                        
                     '92.12.18 MODIFY BY SONIA
                     'Case 答辯, 選取, 修正, 補充說明, 變更
                     '   strTmp = "10"
                     'Modified by Morgan 2016/3/3 +126 期末拋棄
                     'Modified by Lydia 2016/08/26 再考量試行計畫(AFCP 2.0)
                     Case 答辯, 修正, 補充說明, 變更, "126", "438"
                        strTmp = "10"
                        
                        'Added by Morgan 2021/4/16 寶齡富錦 Y55435 案件
                        If field(75) = "Y55435" And cp(10) = 答辯 Then
                           strTmp = "99"
                        End If
                        'end 2021/4/16
                        
                     Case 選取
                        strTmp = "10"
                        If field(9) = "101" Then
                           strTmp = "12"
                        End If
                    '92.12.18 END
                     Case 提供前案資料
                        strTmp = "10"
                        '93.12.13 cancel by sonia 因找不到定稿暫時取消
                        'If txtCaseField(7) = "Y" Then
                        '   strTmp = "11"
                        'End If
                        '93.12.13 end
                     'Modify By Cheng 2003/01/27
                     '加案件性質
'                        Case 年費
                     Case 年費, 維持費, 延展費
                        strTmp = "20"
                        
                        'Added by Morgan 2021/3/15 寶齡富錦 Y55435 案件
                        If field(75) = "Y55435" Then
                           strTmp = "99"
                        End If
                        'end 2021/3/15
                        
                        'Add by Morgan 2005/1/3 CFP-02-605-21
                        '94.2.15 MODIFY BY SONIA
                        'If field(9) = "221" Then 'EPC
                        'Modify by Morgan 2005/11/14 加判斷核准
                        'If field(9) = "221" And field(20) <> "" Then 'EPC
                        'Modified by Morgan 2023/9/7 改判斷已發證--禧佩 Ex.CFP-031017
                        'If field(9) = "221" And field(20) <> "" And field(16) = "1" Then 'EPC
                        If field(9) = "221" And field(22) <> "" And field(16) = "1" Then 'EPC
                           strTmp = "21"
                        End If
                        
                     'Added by Morgan 2012/12/17
                     '馬來西亞,俄羅斯新型延展費
                     'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
                     If (field(9) = "018" Or field(9) = "023") And field(8) = "2" And cp(10) = "607" Then
                        strTmp = "23"
                     Else
                        
'Remove by Morgan 2011/7/25 改用一般定稿,原定稿只有605並無607,且內容並無不同,已刪除
'                        'Add by Morgan 2008/8/12 香港年費
'                        If field(9) = "013" Then
'                           strTmp = "22"
'                        End If
                        
                        '至下一程序檔中找下一程序代號是繳年費及是否續辦為空，是則一般，若空的則是最後一次年費
                         'Modify By Cheng 2003/01/27
'                           strExc(0) = "SELECT COUNT(*) FROM NEXTPROGRESS WHERE " & ChgNextProgress(cp(1) & cp(2) & cp(3) & cp(4)) & _
'                              " AND NP07=" & 年費 & " AND NP06 IS NULL"
                        'Modified by Morgan 2019/3/6 EPC子案要抓母案的期限 Ex:CFP-012758-03--玫音
                        strExc(0) = "SELECT COUNT(*) FROM NEXTPROGRESS WHERE " & ChgNextProgress(cp(1) & cp(2) & cp(3) & IIf(cp(4) <> "00", "00", cp(4))) & _
                           " AND NP07=" & cp(10) & " AND NP06 IS NULL"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If RsTemp.Fields(0) = 0 Then
                              bolLastTime = True 'Added by Morgan 2019/7/10 延展費期滿定稿應可共用 29,將內容參照印度案加入定稿設計並依國家專利種類設定定稿代碼
                              strTmp = "" '最後一年若沒期滿定稿則不要出,遇到時逐一新增--禧佩
                              
                              Select Case field(9)
'2008/4/23 cancel by sonia 日本新型,法國新型,比利時新型改用23同一種定稿
'                                 Case "011"
'                                    '日本新型最後一次年費 6
'                                    strTmp = "26"
'2008/4/23 end
                                 
                                 'Added by Morgan 2012/12/19
                                 Case "011"
                                    'Modified by Morgan 2021/5/26 +發明(並改通用定稿)--玫音
                                    If field(8) = "2" Or field(8) = "1" Then
                                       strTmp = "25"
                                    
                                    '日本設計自申請日2007/4/1起改用新法,舊法為發證日起15年,第4年起逐年繳交。
                                    ElseIf field(8) = "3" Then
                                       If Val(DBDATE(field(10))) < 20070401 Then
                                          strTmp = "23"
                                       Else
                                          strTmp = "24"
                                       End If
                                    End If
                                    
                                 'Add by Morgan 2010/8/20
                                 Case "012" '韓國新型最後一年
                                    
                                    If field(8) = "2" Then
                                       'Modified by Morgan 2023/2/21 已無案件,舊法定稿刪除
                                       'If DBDATE(field(10)) < 19990701 Then
                                       '   strTmp = "26"
                                       'Else
                                       '   'Modified by Morgan 2012/11/7 改通用定稿
                                       '   'strTmp = "24"
                                       '   strTmp = "25"
                                       'End If
                                       strTmp = "25"
                                       'end 2023/2/21
                                    'Added by Morgan 2013/1/16
                                    ElseIf field(8) = "1" Then
                                       strTmp = "28"
                                    End If
                                 
                                 'Added by Morgan 2016/8/10
                                 Case "013" '香港
                                    If field(8) = "3" Then
                                       strTmp = "25"
                                    End If
                                    
                                 'Added by Morgan 2013/10/11
                                 Case "014" '新加坡
                                    'If field(8) = "1" Then    'cancel by sonia 2015/4/16 加設計期滿定稿
                                       strTmp = "30"
                                    'End If                    'cancel by sonia 2015/4/16
                                    
                                 Case "015"
                                    '澳洲新型最後一次年費 5
                                    'Modified by Morgn 2012/11/7 +發明(改通用定稿)
                                    If field(8) = "1" Or field(8) = "2" Then
                                       strTmp = "25"
                                    'Added by Morgan 2012/9/6
                                    '澳洲設計
                                    ElseIf field(8) = "3" Then
                                       strTmp = "39"
                                       
                                    End If
                                 'Added by Morgan 2023/2/20
                                 Case "017" '印尼
                                    strTmp = "25" '通用定稿
                                    
                                 'Add by Morgan 2011/6/14
                                 Case "018" '馬來西亞
                                    'Added by Lydia 2024/02/20 +發明/新型(改通用定稿)
                                    If field(8) = "1" Or field(8) = "2" Then
                                       strTmp = "25"
                                    '設計延展期滿
                                    'If field(8) = "3" Then
                                    ElseIf field(8) = "3" Then
                                    'end 2024/02/20
                                       strTmp = "22"
                                    End If
                                    
                                 'Add by Morgan 2009/5/8
                                 Case "019" '泰國
                                    If field(8) = "2" Then   '新型最後一年
                                       strTmp = "29"
                                    'Modified by Morgan 2023/2/22 發明設計可用通用定稿
                                    'End If
                                    ''add by sonia 2014/5/6 CFP-017351
                                    'If field(8) = "3" Then   '設計最後一年
                                    '   strTmp = "31"
                                    'End If
                                    ''end 2014/5/6
                                    Else
                                       strTmp = "25" '通用定稿
                                    End If
                                    'end 2023/2/22
                                    
                                 'Added by Morgan 2012/11/27
                                 'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
                                 Case "023"
                                    strTmp = "25"
                                    
                                 'Added by Morgan 2023/8/7
                                 '菲律賓發明年費最後一年
                                 Case "030"
                                    If field(8) = "1" Then
                                       strTmp = "25" '通用定稿
                                    End If
                                    
                                 'Added by Morgan 2019/7/10
                                 Case "040" '印度設計最後一年
                                    If field(8) = "3" Then
                                       strTmp = "29"
                                    'Added by Morgan 2022/11/29 印度發明
                                    ElseIf field(8) = "1" Then
                                       strTmp = "25" '通用定稿
                                    End If
                                 
                                 'Add by Morgan 2011/3/8
                                 Case "042" '越南新型最後一年
                                    If field(8) = "2" Then
                                       'Modified by Morgan 2012/11/7 改抓通用定稿
                                       'strTmp = "24"
                                       'strTmp = "25"
                                       'Add by Lydia 2014/10/13 新增-越南新型期滿通知
                                       strTmp = "32"
                                    End If
                                    'Add by Lydia 2015/01/28 新增越南設計期滿CFP-017058
                                    If field(8) = "3" Then   '延展費最後一年
                                       strTmp = "33"
                                    End If
                                 Case 美國國家代號
                                    '美國發明最後一次年費 2
                                    If field(8) = "1" Then
                                       strTmp = "22"
                                    End If
                                    
                                 'Added by Morgan 2024/2/1
                                 Case "102" '加拿大
                                    If field(8) = "1" Then
                                       strTmp = "25"
                                    End If
                                    
                                 Case "104" '墨西哥最後一次年費 Add by Morgan 2010/11/9
                                    'Modified by Morgan 2019/9/11 +設計
                                    If field(8) = "2" Or field(8) = "3" Then
                                       'Modified by Morgan 2012/11/7 改抓通用定稿
                                       'strTmp = "30"
                                       strTmp = "25"
                                    End If
                                    
                                 Case "117"
                                    '巴西新型最後一次年費 5
                                    'Modified by Morgan 2012/12/17 +發明
                                    'If field(8) = "2" Then
                                    If field(8) = "1" Or field(8) = "2" Then
                                       'Modified by Morgan 2012/11/7 改抓通用定稿
                                       'strTmp = "27"
                                       strTmp = "25"
                                    'Added by Morgan 2025/5/9 巴西設計
                                    ElseIf field(8) = "3" Then
                                       strTmp = "32"
                                    End If
                                 
                                 'add by sonia 2016/11/18
                                 Case "118"    '阿根廷118設計 CFP-017348
                                    If field(8) = "3" Then
                                       strTmp = "27"
                                    End If
                                 'end 2016/11/18
                                 
                                 'Added by Morgan 2013/1/7
                                 Case "126" '智利
                                    If field(8) = "3" Then
                                       strTmp = "27"
                                    End If
                                    
                                 '2008/4/23 modify by sonia 加日本新型,比利時新型
                                 'Case "203"
                                 'Modified by Morgan 2012/12/19 日本抽出來(要加設計)
                                 Case "203", "209"
                                    '法國,比利時新型最後一次年費 3
                                    'Modified by Morgan 2024/12/19 +法國發明--玫音
                                    If field(8) = "2" Or (field(9) = "203" And field(8) = "1") Then
                                       'Modified by Morgan 2012/11/7 改抓通用定稿
                                       'strTmp = "23"
                                       strTmp = "25"
                                    End If
                                 
                                 'Added by Morgan 2012/11/7
                                 Case "201"
                                    If field(8) = "1" Then
                                       strTmp = "25"
                                    'Added by Morgan 2024/1/15 再註冊與一般共用--玫音 Ex:CFP-034217
                                    ElseIf field(8) = "3" Then
                                       strTmp = "31"
                                    'end 2024/1/15
                                    End If
                                    
                                 Case "204"
                                    '義大利新型最後一次年費 1
                                    If field(8) = "2" Then
                                       strTmp = "21"
                                    'Added by Morgan 2012/11/7 +發明
                                    ElseIf field(8) = "1" Then
                                       strTmp = "25"
                                    End If
                                    
                                 'Add by Morgan 2018/4/9
                                 Case "210"
                                    '荷比盧設計最後一次年費
                                    If field(8) = "3" Then
                                       strTmp = "34"
                                    End If
                                    
                                 'Add by Morgan 2006/11/24
                                 Case "217"
                                    '芬蘭新型最後一次年費
                                    If field(8) = "2" Then
                                       strTmp = "28"
                                    End If
                                 
                                 'Added by Morgan 2022/8/5
                                 Case "221" 'EPC
                                    strTmp = "21"
                                    
                                 'Added by Morgan 2012/10/11
                                 Case "222"
                                    '波蘭新型最後一次年費
                                    If field(8) = "2" Then
                                       strTmp = "22"
                                    End If
                                 
                                 Case "231"
                                    '德國新型最後一次年費 4
                                    If field(8) = "2" Then
                                       strTmp = "24"
                                    'add by sonia 2016/11/18 +設計CFP-007090
                                    ElseIf field(8) = "3" Then
                                       strTmp = "26"
                                    'end 2016/11/18
                                    'Added by Morgan 2018/5/15 +發明
                                    ElseIf field(8) = "1" Then
                                       strTmp = "25"
                                    End If
                                    

                                 'add by sonia 2016/11/18
                                 Case "235"    '土耳其235新型 CFP-019982
                                    If field(8) = "2" Then
                                       strTmp = "25"
                                    End If
                                 'end 2016/11/18
                                 
                                 'Added by Morgan 2012/10/11
                                 Case "239"
                                    '歐盟設計最後一次年費
                                    If field(8) = "3" Then
                                       strTmp = "40"
                                    End If
                                 
                                 'Added by Morgan 2023/11/28
                                 Case "301"
                                    If field(8) = "1" Or field(8) = "3" Then
                                       strTmp = "25"
                                    End If
                                    
                              End Select
                           End If
                        End If
                     End If
                     '******** 90.11.14   nick
                     '邱小姐說沒有的案件性質傳一般
                     Case Else
                        'Modified by Morgan 2018/6/15 +EPC領證定稿
                        If field(9) = "221" And cp(10) = "601" Then
                           strTmp = "13"
                        Else
                           '一般
                           strTmp = "11"
                           
                           'Added by Morgan 2021/3/15 寶齡富錦 Y55435 案件
                           If field(75) = "Y55435" Then
                              strTmp = "99"
                              txtCaseField(9) = "Y" '案件性質帶中文需人工改英文
                           End If
                           'end 2021/3/15
                        End If
                     '**************************
                  End Select
                  
                  'Add by Morgan 2005/12/26
                  'If chkChoose(0).Value = vbChecked Then
                  'Modify by Morgan 2006/11/7 植物新品種有勾收據時開窗
                  If cp(10) = "120" And chkChoose(2).Value Then
                     strTmp = Val(strTmp) + 30
                     '有客戶案件案號
                     If field(48) <> "" Then
                        strTmp = Val(strTmp) + 30
                     End If
                  'Modify by Morgan 2006/9/13 勾圖示也要印開窗
                  '有勾申請書時印有開窗的格式
                  ElseIf chkChoose(0).Value Or chkChoose(1).Value Then
                     strTmp = Val(strTmp) + 30
                     '有客戶案件案號
                     If field(48) <> "" Then
                        strTmp = Val(strTmp) + 30
                     End If
                  End If
                  
                  If strTmp <> "" Then
                     strReceiveNo = cp(9)
                     StartLetter "02", strTmp
                     bolChk = IIf(Me.txtCaseField(9).Text = "Y", True, False)
                     'Modified by Morgan 2018/6/13  CFP電子化
                     NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , m_NewCP09
                     If m_bolAddLP And bolChk Then
                        frm1105_1.m_RecNo = m_NewCP09
                        frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & strCP10 & ".CUS.PDF"
                        frm1105_1.Show
                     End If
                     'end 2018/6/13
                     
                  End If
               End If
               
               'Add By Sindy 2016/10/7
               If Me.m_strIR01 <> "" Then
                  bolLeave = True
                  intLeaveKind = 0
                  Unload frm050103_1
                  'Modify By Sindy 2022/5/20
                  'frm04010519.GoNext
                  Forms(0).Tmpfrm04010519.GoNext
                  Set Forms(0).Tmpfrm04010519 = Nothing
                  '2022/5/20 END
                  Unload Me
               Else
               '2016/10/7 END
                  bolLeave = True
                  intLeaveKind = 1
                  Unload Me
               End If
            'Add By Cheng 2002/11/29
            Else
                MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
            End If
         End If
         Screen.MousePointer = vbDefault
      Case 1, 2
         If Index = 2 Then
            intLeaveKind = 0
         Else
            intLeaveKind = 2
         End If
         Unload Me
   End Select
End Sub

Private Function SaveDatabase() As Boolean
'Add By Cheng 2002/07/31
Dim strDateS(0 To 5) As String
Dim strNP22  As String
Dim strTemp As String, strTemp1 As String, strTemp2 As String, dobDateAdd As Double
Dim varTemp As Variant
Dim strDate As String, strDate1 As String, strStartDate As String
Dim strTxt(1 To 30) As String, iStep As Integer
Dim lMax As Long
Dim bolNP22 As Boolean, NP22(1 To 3) As String, iNP22 As Integer
'edit by nickc 2007/02/02
'Dim strDataTemp(1 To T_CP) As String
Dim strDataTemp() As String
ReDim strDataTemp(1 To TF_CP) As String
Dim strCRList As String
Dim strPA12 As String
Dim strPrePA12 As String
'Add by Morgan 2004/6/8 指定費用
Dim m_NP08 As String, m_NP09 As String
Dim str1205CP09 As String 'Add by Morgan 2010/3/5 通知提供前案收文號(印C類接洽單用)
Dim str1205CP14 As String 'Added by Morgan 2018/10/2
Dim bolYrFee As Boolean 'Add by Morgan 2010/3/10 是否更新年費期限
Dim strEP06 As String 'Add by Morgan 2010/10/4 齊備日
Dim bolSavPdf As Boolean 'Added by Morgan 2018/10/2
Dim strFullFileName As String, oContext As String, oCustName As String 'Add By Sindy 2019/7/24
Dim iNo As Integer 'Added by Morgan 2021/3/15

m_416YN = ""

'Added by Lydia 2015/05/20
m_str118Date = ""
'm_bol118C = False 'Removed by Morgan 2015/12/16 移到存檔前
'end 2015/05/20

 On Error GoTo CheckingErr
 
 '911105 nick transation
 cnnConnection.BeginTrans
 
   cp(45) = txtCaseField(2)
   Select Case cp(10)
      'Modify by Morgan 2004/5/14
      'Case 發明申請, 新型申請, 設計申請, 追加申請, 聯合申請, CIP申請
      '加 分割, CPA申請,  再發行
      '2006/9/12 加CA申請 BY SONIA
      'Modify by Morgan 2006/10/26 加120植物新品種
      'Modify by Morgan 2009/3/18 + 109 PCT申請
      'Modified by Morgan 2025/3/12 +125 衍生設計
      Case 發明申請, 新型申請, 設計申請, 追加申請, 聯合申請, CIP申請, 分割, CPA申請, 再發行, "122", "120", "109", 衍生設計
         field(10) = txtCaseField(0)
         field(11) = txtCaseField(1)
      '20080919 edit by Toni 原申請案號寫到cp30
      Case "301", "302", "303", "305", "306"
         'modify by sonia 2023/12/22 有申請案號才更新CP30
         'cp(30) = field(11)
         If field(11) <> "" Then cp(30) = field(11)
         'end 2023/12/22
         'field(10) = txtCaseField(0) 'Remove by Morgan 2010/6/3 改請不可回寫申請日
         field(11) = txtCaseField(1)
      '20080919 end
      Case Else
         If cp(64) = "" Then
            cp(64) = "代理人提申櫃台收文日：" & Me.txtCaseField(8).Text
         Else
            cp(64) = cp(64) & ",代理人提申櫃台收文日：" & Me.txtCaseField(8).Text
         End If
         '2008/8/26 add by sonia 櫃台收文日
         cp(119) = ChangeTStringToWString(Me.txtCaseField(8).Text)
         '2008/8/26 end
   End Select
   
   strTxt(1) = GetCPSQL(cp())
   
   '911106 nick transation
   cnnConnection.Execute strTxt(1)
   
   '2009/12/30 ADD BY SONIA CFP-022728
   '新增CP or NP時智權人員一律Call Function 抓最新收文智權人員
   cp(13) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
   cp(12) = GetSalesArea(cp(13))
   '2009/12/30 END
   
   iStep = 2
   
   lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax
   bolNP22 = False
   iNP22 = 1
   '若為申請案
   'If cp(10) = 發明申請 Or cp(10) = 新型申請 Or cp(10) = 設計申請 Or cp(10) = 追加申請 Or cp(10) = 聯合申請 Or cp(10) = CIP申請 Then
   '加 分割, CPA申請,  再發行
   '2006/9/12 加CA申請 BY SONIA
   '2008/9/22 加改請案件性質 by sonia
   'Modify by Morgan 2009/3/18 + 109 PCT申請
   'Modify by Morgan 2010/6/14 改用常數
   'If cp(10) = 發明申請 Or cp(10) = 新型申請 Or cp(10) = 設計申請 Or cp(10) = 追加申請 Or cp(10) = 聯合申請 Or cp(10) = CIP申請 Or cp(10) = 分割 Or cp(10) = CPA申請 Or cp(10) = 再發行 Or cp(10) = "122" Or cp(10) = "120" Or cp(10) = "301" Or cp(10) = "302" Or cp(10) = "303" Or cp(10) = "305" Or cp(10) = "306" Or cp(10) = "109" Then
   If InStr(m_RefCP10List, cp(10)) > 0 Then
      strTxt(iStep) = GetPASQL(field())
      '911106 nick transation
      cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
      
      '有輸入申請案號時才新增進度檔
      'Modify by Morgan 2005/10/27
      '只輸申請日時也會要寄通知函所以一律新增進度檔
      'If field(11) <> "" Then
      
         'Add by Morgan 2016/5/30
         If txtCaseField(1) <> "" Then
            strCP10 = 通知申請案號
         Else
            strCP10 = 通知申請日
         End If
         'end 2016/5/30
         
         strDataTemp(1) = cp(1)
         strDataTemp(2) = cp(2)
         strDataTemp(3) = cp(3)
         strDataTemp(4) = cp(4)
         strDataTemp(5) = strSrvDate(1)
         strDataTemp(27) = strSrvDate(1)
         strDataTemp(9) = 主管機關來函
         'Modify by Morgan 2016/5/30
         'strDataTemp(10) = 通知申請案號
         strDataTemp(10) = strCP10
         'end 2016/5/30
         'Modify by Morgan 2010/3/5
         'strDataTemp(13) = cp(13)
         'strDataTemp(12) = cp(12)
         strDataTemp(13) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
         strDataTemp(12) = GetSalesArea(strDataTemp(13))
         'end 2010/3/5
         strDataTemp(14) = strUserNum
         strDataTemp(20) = "N"
         strDataTemp(26) = "N"
         strDataTemp(32) = "N"
         strDataTemp(43) = cp(9)
         '2008/8/26 modify by sonia 櫃台收文日改存 cp119
         '2008/10/24 MODIFY BY SONIA CP64仍存
         strDataTemp(64) = "櫃台收文日：" & Me.txtCaseField(8).Text
         strDataTemp(119) = ChangeTStringToWString(Me.txtCaseField(8).Text)
         '2008/8/26 end
         
         If chkChoose(0).Value Or txtCaseField(6) <> "" Then strDataTemp(145) = "Y"  'Added by Morgan 2018/9/27 說明書或附件都算有副本
         
         strTxt(iStep) = GetCPSQL(strDataTemp(), False)
         
         cnnConnection.Execute strTxt(iStep)
         iStep = iStep + 1
         
         'Add by Morgan 2010/3/5 印度發明CP新增通知提供前案1205,NP新增提供前案資料207
         If field(9) = "040" And field(8) = "1" And Val(cp(47)) >= 0 Then
            '法限=申請日6個月,所限=法限-14天,約定=所限-28天
            strExc(1) = CompDate(1, 6, Val(cp(47))) '法限
            strExc(2) = PUB_GetWorkDay1(CompDate(2, -14, Val(strExc(1))), True) '所限
            strExc(3) = PUB_GetWorkDay1(CompDate(2, -28, Val(strExc(2))), True) '約定
            '承辦人,承辦期限
            strExc(5) = "NULL"
            If GetStaffName(cp(14)) = "" Then
               str1205CP14 = ""
            Else
               str1205CP14 = cp(14)
               If PUB_GetST06(cp(14)) <> "1" Then '分所人員以系統日的下一個工作天計算
                  strEP06 = CompWorkDay(2, strSrvDate(1), 0)
               Else
                  strEP06 = strSrvDate(1)
               End If
               If PUB_IfSetCP48() Then  'Add by Morgan 2010/10/4
                  strExc(5) = Pub_GetHandleDay(field(1), field(9), "1205", strEP06, strExc(2))
               End If
            End If
            'Add by Morgan 2010/5/7 +考慮第2次輸申請案號
            strSql = "select cp09,cp27 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='207' and cp57 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            '已收文
            If intI = 1 Then
               '未發文
               If IsNull(RsTemp("cp27")) Then
                  strSql = "update caseprogress set cp06=" & strExc(2) & ",cp07=" & strExc(1) & " where cp09='" & RsTemp("cp09") & "'"
                  cnnConnection.Execute strSql, intI
               End If
               
            Else
               strSql = "select cp09,cp27,cp07 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='1205' and cp57 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               '已收文
               If intI = 1 Then
                  If RsTemp("cp07") <> strExc(1) Then
                     intI = 0 '期限不同重新產生來函
                  End If
               End If
               
               If intI <> 1 Then
            'end 2010/5/7
            
                  '收文號
                  str1205CP09 = AutoNo("C", 6)
                  strExc(7) = strDataTemp(64) & "，約定期限：" & TransDate(strExc(3), 1)
                  
                  'Modified by Morgan 2019/1/19 +cp121
                  strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10" & _
                     ",cp12,cp13,cp14,cp20,cp26,cp32,cp43,cp45,cp48,cp64,cp121 ) values ('" & cp(1) & "'" & _
                     ",'" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & _
                     "," & strExc(2) & "," & strExc(1) & ",'" & str1205CP09 & "','1205','" & strDataTemp(12) & "'" & _
                     ",'" & strDataTemp(13) & "','" & str1205CP14 & "','N','N','N'" & _
                     ",'" & cp(9) & "','" & ChgSQL(cp(45)) & "'," & strExc(5) & ",'" & strExc(7) & "','Y')"
                     
                  cnnConnection.Execute strSql, intI
                  
                  'Added by Morgan 2019/1/19 自動收文的通知提供前案,沒有附件,工程師寫信
                  PUB_AddLetterProgress str1205CP09, 0, False, , , field(26), "1205", field(75)
                  PUB_UpdateLP03 str1205CP09
                  'end 2019/1/19
                  
                  'Add by Morgan 2010/10/4
                  If strEP06 <> "" Then
                     strSql = "update engineerprogress set ep06=" & strEP06 & " where ep02='" & str1205CP09 & "'"
                     cnnConnection.Execute strSql, intI
                  End If
                  
                  strSql = "insert into nextprogress(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22)" & _
                     " values ('" & str1205CP09 & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "'" & _
                     ",'" & cp(4) & "','207'," & strExc(2) & "," & strExc(1) & ",'" & strDataTemp(13) & "'" & _
                     ",getnp22)"
                  cnnConnection.Execute strSql, intI
               End If
            End If
         End If
         
         'Add by Morgan 2007/1/23
         '起算日統一抓一次就好
         '若有主張優先權時用申請日起算的期限都改抓最早優先權日--慧汶
         'strStartDate = GetStartDate(申請日, strDataTemp(), field())
         strStartDate = PUB_GetFirstPriDate2(strPriority(2))
         
         'Added by Morgan 2012/11/23
         '比利時發明設定"申請檢索報告"421期限為優先權日(申請日)18個月
         'Modified by Morgan 2018/2/8 +土耳其發明案檢索報告法定期限為申請日起12個月
         If (field(9) = "209" Or field(9) = "235") And cp(10) = "101" Then
            '比利時
            If field(9) = "209" Then
               If strStartDate <> "" Then
                  strExc(0) = strStartDate
               Else
                  strExc(0) = TransDate(txtCaseField(0), 2)
               End If
               '法限
               strExc(1) = CompDate(1, 18, strExc(0))
            '土耳其
            Else
               strExc(1) = CompDate(1, 12, txtCaseField(0))
            End If
         'end 2018/2/8
          
            strDateS(1) = cp(1)
            strDateS(2) = field(9)
            strDateS(3) = strExc(1)
            GetCtrlDT strDateS
            '所限
            strExc(2) = strDateS(0)
            strExc(2) = PUB_GetWorkDay1(strExc(2), True)
            
            strSql = "Select cp09,cp27 From caseprogress Where cp01='" & cp(1) & "' AND cp02='" & cp(2) & "' AND cp03='" & cp(3) & "' AND cp04='" & cp(4) & "' AND cp10='421' and cp57 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If IsNull(RsTemp("cp27")) Then
                  strSql = "update caseprogress set cp06=" & strExc(2) & ",cp07=" & strExc(1) & " where cp09='" & RsTemp("cp09") & "'"
                  cnnConnection.Execute strSql
               End If
            Else
               strSql = "Select NP01,NP07,NP22 From Nextprogress Where NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP07='421' AND NP06 IS NULL  ORDER BY NP22 DESC"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " WHERE NP01='" & RsTemp("NP01") & "' AND NP22=" & RsTemp("NP22")
               Else
                  strSql = "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     " select '" & strDataTemp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','421'," & strExc(2) & "," & strExc(1) & ",'" & strDataTemp(13) & "',newNP22 from dual,(select nvl(max(np22),0)+1 newNP22 from nextprogress)"
               End If
               cnnConnection.Execute strSql, intI
            End If
         End If
         'end 2012/11/23
         
         'Add by Morgan 2009/7/28 若申請國家為PCT時掛119(進入國家階段)期限,法限=申請日(最早優先權日)+30月;所限=法限-2月(同P案)
         If cp(10) = "109" And txtCaseField(0) <> "" Then
            '起算日
            If strStartDate <> "" Then
               strExc(0) = strStartDate
            Else
               strExc(0) = TransDate(txtCaseField(0), 2)
            End If
            '法限
            strExc(1) = CompDate(1, 30, strExc(0))
            '所限
            strExc(2) = CompDate(1, -2, strExc(1))
            strExc(2) = PUB_GetWorkDay1(strExc(2), True)
            
            strSql = "Select NP01,NP07,NP22 From Nextprogress Where NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP07='119' AND NP06 IS NULL  ORDER BY NP22 DESC"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strSql = "update nextprogress set np08=" + strExc(2) + ",np09=" + strExc(1) + " WHERE NP01='" & RsTemp("NP01") & "' AND NP22=" & RsTemp("NP22")
            Else
               strSql = "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  " select '" & strDataTemp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','119'," & strExc(2) & "," & strExc(1) & ",'" & strDataTemp(13) & "',newNP22 from dual,(select nvl(max(np22),0)+1 newNP22 from nextprogress)"
            End If
            cnnConnection.Execute strSql
         End If
         'end 2009/7/28
         
'Modify by Morgan 2008/10/20 改呼叫共用函數
         PUB_UpdCfpDate2 field(1), field(2), field(3), field(4), strStartDate, strDataTemp(9), m_416YN, IIf(cp(10) = "307", True, False)
          
         'Added by Morgan 2023/5/23
         '泰國發明及設計輸申請日時下一程序新增公開費(217)，管制期限同公開(999)--禧佩
         If field(9) = "019" And (field(8) = "1" Or field(8) = "3") Then
            PUB_GetOpenDate field(1), field(2), field(3), field(4), strTemp1, strTemp2, strDataTemp(9), strStartDate
            
            '設計沒有管制公開預設為申請日(最早優先權日)+18個月
            If strTemp2 = "" Then
               If strStartDate = "" Then strStartDate = DBDATE(txtCaseField(0))
               strTemp2 = CompDate(1, 18, strStartDate)
               strTemp1 = PUB_GetWorkDay1(CompDate(2, -14, strTemp2), True)
            End If
            
            If strTemp2 <> "" Then
               If PUB_ChkCPExist(field, "217", , strTemp, , , strDate) Then
                  If strDate = "" Then
                     strSql = "update caseprogress set cp06=" & strTemp1 & ",cp07=" & strTemp2 & " where cp09='" & strTemp & "'"
                     cnnConnection.Execute strSql, intI
                  End If
               ElseIf PUB_ChkNPExist(cp, "217", 0, strNP22, strTemp) Then
                  strSql = "Update NextProgress Set NP08=" & strTemp1 & ",NP09=" & strTemp2 & " Where NP22=" & strNP22 & " and NP01='" & strTemp & "'"
                  cnnConnection.Execute strSql, intI
               Else
                  strSql = "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     " Values ('" & strDataTemp(9) & "','" & field(1) & "','" & field(2) & "','" & field(3) & "','" & field(4) & "','217'," & strTemp1 & "," & strTemp2 & ",'" & strDataTemp(13) & "',GETNP22) "
                  cnnConnection.Execute strSql, intI
               End If
            End If
         End If
         'end 2023/5/23

'               '2005/12/8 ADD BY SONIA 印度發明申請日2004/12/31(含)以前實審期限為申請日起48個月
'               '2005/12/21 ADD BY SONIA 新加坡發明申請日2004/6/30(含)以前實審期限為申請日起28個月
'               'Add by Morgan 2008/5/2 馬來西亞發明,新型要新增(更新) 提前案資料207 期限
'               'Add by Morgan 2004/6/8 若申請國家為 EPC 時更新[215-指定費]進度檔(下一程序)或新增指定費下一程序，期限同[416-實體審查]
'
''''舊程式已刪除'''

         
'2008/10/15 CANCEL BY SONIA馬來西亞新型於發證時才掛年費期限,俄羅斯新型為准後繳年費,故取消此處改至發證時控管
'         '更新延展費期限
'         'Add by Morgan 2006/8/21 馬來西亞新型新法(發證日>=2001/8/1)延展費自申請日起10年,期滿延展2次,每次5年
'         'Modify by Morgan 2007/7/18 新型不管何時提申均適用上述新法(2003年8月14日實施)
'         'Modify by Morgan 2007/7/31 加俄羅斯新型也要掛延展費期限
'
''''舊程式已刪除'''


         '更新年費期限
         '讀取年費設定
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, m_NP07, strTemp2) = 0 Then
         'Modified by Morgan 2013/10/23
         'If ClsPDGetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, m_NP07, strTemp2) = 0 Then
         If ClsPDGetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, m_NP07, strTemp2, , , field(10), field(21), field(72)) = 0 Then
            varTemp = Split(strTemp1, ",")
            
            'Add by Morgan 2010/3/3 '起算日為公開日且公開日為申請日起算者也要掛年費期限
            If Val(strTemp) = 公開日 Then
               strExc(0) = "select 1 from nation where decode('" & field(8) & "','1',na32,'2',na34,'3',na36)=2"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If blnUpdateNP(cp(1) & cp(2) & cp(3) & cp(4)) = True Then
                     strTemp = 申請日
                  End If
               End If
            End If
               
            'Add by Morgan 2004/9/7 加判斷申請日起算且有帶出下次繳費日的才更新，不再重算
            If Val(strTemp) = 申請日 And txtCaseField(3) <> "" Then
               bolYrFee = True
            End If
            
         End If
   '若非申請案
   Else
      '若有下次繳費日
      If txtCaseField(3) <> "" Then
         bolYrFee = True
      End If
   End If
   
   'Modify by Morgan 2010/3/10 將上面相同程式合併到此
   If bolYrFee Then
      strDate = DBDATE(txtCaseField(3)) '下次繳費日(法限)
      strDateS(1) = cp(1)
      strDateS(2) = field(9)
      strDateS(3) = TransDate(strDate, 2)
      GetCtrlDT strDateS
      strDate1 = strDateS(0) '下次繳費日(所限)
      strDate1 = PUB_GetWorkDay1(strDate1, True)
      m_605NP08 = strDate1 'Added by Morgan 2021/3/15
      
      'Add by Morgan 2010/3/10 要加判斷已收文狀況
      strExc(0) = "SELECT CP09 FROM CASEPROGRESS WHERE Cp10=" & m_NP07 & " and Cp01=" + CNULL(cp(1)) + _
         " and Cp02=" + CNULL(cp(2)) + " and Cp03=" + CNULL(cp(3)) + " and Cp04=" + CNULL(cp(4)) & " and Cp27||CP57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "update caseprogress set cp06=" & strDate1 & ",cp07=" & strDate & " where cp09='" & RsTemp("cp09") & "'"
         cnnConnection.Execute strSql, intI
      Else
      'End 2010/3/10
      
         strExc(0) = "SELECT NP22 FROM NEXTPROGRESS WHERE np07=" & m_NP07 & " and np02=" + CNULL(cp(1)) + _
            " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05=" + CNULL(cp(4)) & " and np06 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            lMax = ClsLawGetMax
            '判斷是否要更新下一程序
            If blnUpdateNP(cp(1) & cp(2) & cp(3) & cp(4)) = True Then
                strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22) values (" & _
                   CNULL(strDataTemp(9)) & "," & CNULL(cp(1)) & "," & CNULL(cp(2)) & "," & CNULL(cp(3)) & _
                   "," & CNULL(cp(4)) & "," & CNULL(m_NP07) & "," & strDate1 & "," & strDate & _
                   "," & CNULL(strDataTemp(13)) & "," & lMax & ")"
                 cnnConnection.Execute strTxt(iStep)
            End If
   
            bolNP22 = True
            NP22(iNP22) = lMax
            iNP22 = iNP22 + 1
            iStep = iStep + 1
            lMax = lMax + 1
         Else
            strTxt(iStep) = "update nextprogress set np08=" & strDate1 & ",np09=" & strDate & " WHERE np07=" & m_NP07 & _
               " and np02=" & CNULL(cp(1)) & _
               " and np03=" & CNULL(cp(2)) & " and np04=" & CNULL(cp(3)) & " and np05=" & CNULL(cp(4)) & " and np06 is null"
            cnnConnection.Execute strTxt(iStep)
   
            iStep = iStep + 1
         End If
      End If
   End If
         
   'Add by Morgan 2010/4/2
   '集體設計申請日號同母案
   'modify by sonia 2017/3/15 剔除韓國012設計案CFP-028750
   'Modified by Morgan 2025/7/29 英國集體子案也要個別輸入申請號--禧佩
   If strDataTemp(9) <> "" And cp(10) = "103" And cp(3) = "0" And field(9) <> "012" And field(9) <> "201" Then
      strExc(0) = "select cp03,cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03<>'0' and cp04='" & cp(4) & "' and cp10='105' and cp27>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         iNo = 1 'Added by Morgan 2021/3/15
         With RsTemp
         Do While Not .EOF
            '新增來函
            strExc(1) = AutoNo("C", 6)
            strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp43,cp64,cp119) " & _
               " select cp01,cp02,'" & .Fields("cp03") & "',cp04,cp05,'" & strExc(1) & "',cp10,cp12,cp13,cp14,cp20,cp26,cp27,cp32,'" & .Fields("cp09") & "',cp64,cp119" & _
               " from caseprogress where cp09='" & strDataTemp(9) & "'"
            cnnConnection.Execute strSql, intI
            '更進度檔
            strSql = "update caseprogress a set (cp45,cp47)=(select b.cp45,b.cp47 from caseprogress b where b.cp09='" & cp(9) & "')" & _
               " where cp09='" & .Fields("cp09") & "'"
            cnnConnection.Execute strSql, intI
            '更新基本檔
            'Added by Morgan 2021/3/15 EU申請號比照證書號規則
            If Right(txtCaseField(1), 5) = "-0001" And field(9) = "239" Then
               iNo = iNo + 1
               strSql = "update patent a set (pa10,pa11)=(select b.pa10,replace(b.pa11,'-0001','-" & Format(iNo, "0000") & "') from patent b where b.pa01=a.pa01" & _
                  " and b.pa02=a.pa02 and b.pa03='0' and b.pa04=a.pa04)" & _
                  " where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & .Fields("cp03") & "' and pa04='" & cp(4) & "'"
               cnnConnection.Execute strSql, intI
            Else
            'end 2021/3/15
               strSql = "update patent a set (pa10,pa11)=(select b.pa10,b.pa11 from patent b where b.pa01=a.pa01" & _
                  " and b.pa02=a.pa02 and b.pa03='0' and b.pa04=a.pa04)" & _
                  " where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & .Fields("cp03") & "' and pa04='" & cp(4) & "'"
               cnnConnection.Execute strSql, intI
            End If 'Added by Morgan 2021/3/15
            
            '集體設計年費期限同母案
            If bolYrFee Then
               strSql = "update caseprogress set cp06=" & strDate1 & ",cp07=" & strDate & _
                  " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & .Fields("cp03") & "' and cp04='" & cp(4) & "'" & _
                  " and cp10='" & m_NP07 & "' and cp27 is null and cp57 is null"
               cnnConnection.Execute strSql, intI
               If intI = 0 Then
                  strSql = "update nextprogress set np08=" & strDate1 & ",np09=" & strDate & _
                     " where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & .Fields("cp03") & "' and np05='" & cp(4) & "'" & _
                     " and np07='" & m_NP07 & "' and np06 is null"
                  cnnConnection.Execute strSql, intI
                  If intI = 0 Then
                     strSql = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
                        " values ('" & strExc(1) & "','" & cp(1) & "','" & cp(2) & "','" & .Fields("cp03") & "','" & cp(4) & "'" & _
                        ",'" & m_NP07 & "'," & strDate1 & "," & strDate & _
                        ",'" & strDataTemp(13) & "',getnp22)"
                     cnnConnection.Execute strSql, intI
                  End If
               End If
            End If
            
            'Added by Morgan 2019/7/2 提申期限也要解除 Ex:CFP-031097-1-00--禧佩
            strSql = "update nextprogress set np06='Y' where np01='" & .Fields("cp09") & "' and np02='" & cp(1) & _
              "' and np03='" & cp(2) & "' and np04='" & .Fields("cp03") & "' and np05='" & cp(4) & "' and np07 IN (" & 提申 & "," & 收達 & ",995,996) and np06 is null"
            cnnConnection.Execute strSql, intI
            'end 2019/7/2
            
            .MoveNext
         Loop
         End With
      End If
   End If
   'end 2010/4/2
   
Nextstep:
   
   'Modify by Morgan 2009/5/12 +指定提申995,最後提申996
   strTxt(iStep) = "update nextprogress set np06='Y' where np01='" & cp(9) & "' and np02='" & cp(1) & _
      "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np07 IN (" & 提申 & "," & 收達 & ",995,996)"
    cnnConnection.Execute strTxt(iStep)

      
   iStep = iStep + 1
   
   '93.11.10 MODIFY BY SONIA
   'If cp(1) = "CPS" And cp(10) = 美國暫時申請 Then
   If cp(10) = 美國暫時申請 Then
   '93.11.10 END
      strTxt(iStep) = "UPDATE CASEPROGRESS SET CP47=" & CNULL(TransDate(txtCaseField(0), 2)) & ", CP45=" & CNULL(txtCaseField(2)) & " WHERE CP09='" & cp(9) & "'"
      '911105 nick transation
      cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
      '93.11.10 MODIFY BY SONIA
      'strTxt(iStep) = "UPDATE SERVICEPRACTICE SET SP10=" & CNULL(TransDate(txtCaseField(0), 2)) & ", SP11=" & CNULL(txtCaseField(1)) & " WHERE CP09='" & cp(9) & "'"
      strTxt(iStep) = "UPDATE PATENT SET PA10=" & CNULL(TransDate(txtCaseField(0), 2)) & ", PA11=" & CNULL(txtCaseField(1)) & " WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'"
      '93.11.8 END
      
      cnnConnection.Execute strTxt(iStep)
      
      '有輸入申請案號時才新增進度檔 Mem by Morgan 2004/6/8
      'Modify by Morgan 2007/3/2 改判斷有輸申請日時 -- 慧汶
      'If txtCaseField(1) <> "" Then
      If txtCaseField(0) <> "" Then
      'end 2007/3/2
      
         'Add by Morgan 2016/5/30
         If txtCaseField(1) <> "" Then
            strCP10 = 通知申請案號
         Else
            strCP10 = 通知申請日
         End If
         'end 2016/5/30
         
         iStep = iStep + 1
         lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax
         strDataTemp(1) = cp(1)
         strDataTemp(2) = cp(2)
         strDataTemp(3) = cp(3)
         strDataTemp(4) = cp(4)
         strDataTemp(5) = strSrvDate(1)
         strDataTemp(27) = strSrvDate(1)
         strDataTemp(9) = 主管機關來函
         'Modify by Morgan 2016/5/30
         'strDataTemp(10) = 通知申請案號
         strDataTemp(10) = strCP10
         'end 2016/5/30
         strDataTemp(13) = cp(13)
         strDataTemp(12) = cp(12)
         strDataTemp(14) = strUserNum
         strDataTemp(20) = "N"
         strDataTemp(26) = "N"
         strDataTemp(32) = "N"
         strDataTemp(43) = cp(9)
         
         '2008/8/26 modify by sonia 櫃台收文日改存 cp119
         '2008/10/24 MODIFY BY SONIA CP64仍存
         strDataTemp(64) = "櫃台收文日：" & Me.txtCaseField(8).Text
         strDataTemp(119) = ChangeTStringToWString(Me.txtCaseField(8).Text)
         '2008/8/26 end
         
         If chkChoose(0).Value Then strDataTemp(145) = "Y"  'Added by Morgan 2018/9/27
         
         strTxt(iStep) = GetCPSQL(strDataTemp(), False)
         cnnConnection.Execute strTxt(iStep)
         
         iStep = iStep + 1
         '93.11.10 期限為暫時申請之申請日+1年
         strDate = Format(TransDate(Me.txtCaseField(0).Text, 2)) + 10000
         strDateS(0) = ""
         strDateS(1) = cp(1)
         strDateS(2) = field(9)
         strDateS(3) = strDate
         GetCtrlDT strDateS()
         'Add by Morgan 2008/1/24 若已掛期限時更新 -- 禧佩
         strExc(0) = "select np22 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06 is null and np07=910 and np15='美國發明申請' order by np01 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strDateS(0) = PUB_GetWorkDay1(strDateS(0), True)
            strSql = "update nextprogress set np08=" & strDateS(0) & ",np09=" & strDateS(3) & " where np22=" & RsTemp.Fields(0) & " and np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "'"
            cnnConnection.Execute strSql, intI
            
         Else
         'end 2008/1/24
            strTxt(iStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP15,NP22,NP10) " & _
                  "VALUES ('" & strDataTemp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "',910," & PUB_GetWorkDay1(strDateS(0), True) & "," & strDateS(3) & ",'美國發明申請'," & lMax & ",'" & cp(13) & "') "
               
            cnnConnection.Execute strTxt(iStep)
         End If
         
         'Added by Lydia 2015/05/20 約定期限=所限-28天
         m_str118Date = CompDate(2, -28, strDateS(0))
         m_str118Date = PUB_GetWorkDay1(m_str118Date, True)
         'end 2015/05/20
         
      End If
      
      bolNP22 = True
      NP22(iNP22) = lMax
      iNP22 = iNP22 + 1
      iStep = iStep + 1
      lMax = lMax + 1
   End If
   
   'Add by Morgan 2004/7/28
   '美國預定公開日若有輸時更新基本檔公開日PA12
   If txtCaseField(10).Text <> "" Then
      strSql = "UPDATE PATENT SET PA12='" & DBDATE(txtCaseField(10).Text) & "' WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'"
      cnnConnection.Execute strSql
   End If
   
   'Added by Morgan 2024/4/25
   If m_bolEPC_DE_CHK Then
      strSql = "UPDATE PATENT SET PA11='" & txtCaseField(1) & "' WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'"
      cnnConnection.Execute strSql
   End If
   'end 2024/4/25

   'Add by Morgan 2007/8/20 當有美國案時,以預定公開日更新其他多國案期限
   'Modify by Morgan 2007/12/31 發明案才要
   'If field(9) = "101" Then
   'Modified by Morgan 2012/1/31 改判斷案件性質 101 才更新(Ex. CFP-22261)--郭
   'If field(9) = "101" And field(8) = "1" Then
   If field(9) = "101" And cp(10) = "101" Then
      If txtCaseField(10).Text <> "" Then
         strTemp = DBDATE(txtCaseField(10))
      ElseIf field(12) <> "" Then
         strTemp = DBDATE(field(12))
      Else
         strTemp = DBDATE(strPrePA12)
      End If
      If strTemp <> "" Then
         PUB_UpdCP07byPA12 field, strTemp
      End If
   End If
   'end 2007/8/20
   
   'Add by Morgan 2005/11/17 有優先權資料且有重輸才要
   If strPriority(1) <> Empty And m_bolRePriDate = True Then
      'Modify by Morgan 2007/4/25 加, strPriority(4)
      'Moidfy by Amy 2014/04/17 + strPriority(5)
      ClsPDSavePriority field, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)
   End If
   '2005/12/19 ADD BY SONIA 更新相同本所案號之相同代理人的彼所案號，若是彼所案號空的話
   If txtCaseField(2) <> "" Then
      'Modified by Morgan 2012/2/15 取消 cp09<'C' 條件(C類也會有發文作業,有代理人就要更新彼號,資料才會一致)
      strSql = "update caseprogress set cp45=" & CNULL(txtCaseField(2)) & " where cp09 in (select cp09 from caseprogress where rtrim(cp45) is null and CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND cp44 in (select cp44 from caseprogress where cp09='" & cp(9) & "' ))"
      cnnConnection.Execute strSql
   End If
   '2005/12/19 END
   
   'Add by Morgan 2010/8/9 若有被主張優先權時需更新相關期限資料
   If txtCaseField(0) <> "" And txtCaseField(1) <> "" Then
      strExc(0) = "select pd01,pd02,pd03,pd04 from pridate where PD06='" & cp(1) & cp(2) & cp(3) & cp(4) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "update pridate set pd05=" & DBDATE(txtCaseField(0)) & ",pd06='" & txtCaseField(1) & "',pd07='" & field(9) & "',pd08='" & field(8) & "'" & _
            " where pd06='" & cp(1) & cp(2) & cp(3) & cp(4) & "'"
         cnnConnection.Execute strSql, intI
         With RsTemp
            .MoveFirst
            Do While Not .EOF
               PUB_UpdCfpDate1 .Fields("pd01"), .Fields("pd02"), .Fields("pd03"), .Fields("pd04")
               .MoveNext
            Loop
         End With
      End If
   End If
   
   'Added by Lydia 2016/08/26 +438 再考量試行計畫(AFCP 2.0)要更新下一程序提申期限為收達日+7天
   If cp(10) = "438" Then
      strExc(1) = CompDate(2, 14, TransDate(txtCaseField(0), 2))
      strSql = "update nextprogress set np08=" & CNULL(strExc(1), True) & " ,np09=" & CNULL(strExc(1), True) & _
               " where np01='" & cp(9) & "' and np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "'" & _
               " and nvl(np06,' ')=' ' and np07='" & 催審 & "' "
      cnnConnection.Execute strSql, intI
   End If
   
   'ADD BY SONIA 2016/9/6 美國暫時申請案自動上可結餘日為申請日一年後
   If cp(10) = "118" Then
      strSql = "update caseprogress set cp109=" & TransDate(txtCaseField(0), 2) + 10000 & _
                " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' " & _
                " and cp27>0 and cp59 is null "
      cnnConnection.Execute strSql, intI
   End If
   'END 2016/9/6
   
   'Add By Sindy 2019/7/24 消催審期限
   If cp(10) = "958" Then '958.代理人撰稿
      strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
               "WHERE NP01 = '" & cp(9) & "' AND " & _
                     "NP07 = '411' AND NP06 IS NULL "
      cnnConnection.Execute strSql
   End If
   '2019/7/24 END
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      'Add By Sindy 2019/7/24
      'Modify By Sindy 2022/11/10 + IIf(field(9) <> 台灣國家代號, "PAT", "RX")
      Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", IIf(field(9) <> 台灣國家代號, "PAT", "RX"), strFullFileName, True)
      '2019/7/24 END
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050103_1"
   End If
   '2016/10/7 END
   
   'Added by Morgan 2016/10/7
   '美國發明案提申時若臺灣案有TW-SUPA未發文時設定法限為申請日+6個月,所限=法限-2工作天(跟郭確認過只需要判斷發明申請)
   'Modified by Morgan 2021/5/13 +日本 --玲玲,郭
   If (field(9) = "101" Or field(9) = "011") And field(8) = "1" And cp(10) = "101" Then
      strExc(1) = CompDate(1, 6, field(10))
      'modify by sonia 2021/7/9 所限=法限-2工作天改為CFP申請日+1個月
      'strExc(2) = PUB_GetOurDeadline(strExc(1))
      strExc(2) = CompDate(1, 1, field(10))
      strExc(2) = PUB_GetWorkDay1(strExc(2), True)
      'end
      strSql = "update caseprogress set cp06=" & strExc(2) & ",cp07=" & strExc(1) & " where (cp01,cp02,cp03,cp04) in" & _
         " (select pa01,pa02,pa03,pa04 from pridate,patent where pd01='" & cp(1) & "' and pd02='" & cp(2) & "' and pd03='" & cp(3) & "' and pd04='" & cp(4) & "'" & _
         " and pd07='000' and pa11(+)=pd06) and cp10='434' and cp27||cp57 is null and cp07 is null"
      cnnConnection.Execute strSql, intI
      
      'Added by Morgan 2021/5/13
      '第1案提申時 EMail通知 TW-SUPA 承辦工程師 --玲玲,郭
      If intI = 1 Then
         strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
         strExc(2) = "(434)TW-SUPA請送件，因相對應的" & strExc(1) & "(" & lblCountryName & ")"
         If txtCaseField(1) <> "" Then
            strExc(2) = strExc(2) & "，已取得申請案號:" & txtCaseField(1)
         Else
            strExc(2) = strExc(2) & "已提申！"
         End If
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " select '" & strUserNum & "',cp14,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)||'" & ChgSQL(strExc(2)) & "'" & _
            ",'如旨' from pridate,patent,caseprogress where pd01='" & cp(1) & "' and pd02='" & cp(2) & "'" & _
            " and pd03='" & cp(3) & "' and pd04='" & cp(4) & "' and pd07='000' and pa11(+)=pd06" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
            " and cp10='434' and cp27||cp57 is null and cp14 is not null"
         cnnConnection.Execute strSql, intI
      End If
      'end 2021/5/13
   End If
   'end 2016/10/7
   
   'Modified by Morgan 2017/9/12 從Transaction外移進來
   'Add by Morgan 2009/8/17
   '若歐盟設計的其他多國皆有申請日時提醒該案已可發文
   chk103in239OK cp
               
   'Added by Morgan 2018/6/8 CFP電子化
   'Modified by Morgan 2020/8/18 年費無收據不新增進度也不出定稿
   'If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) Then
   If Not (chkChoose(2).Value = False And cp(10) >= "605" And cp(10) <= "607") Then
      '年費有收據時回寫副本欄位
      If chkChoose(2).Value And cp(10) >= "605" And cp(10) <= "607" Then
         strSql = "update caseprogress set cp145='Y' where cp09='" & cp(9) & "'"
         cnnConnection.Execute strSql, intI
      End If
   'end 2020/8/18
      If strDataTemp(9) <> "" Then
         m_NewCP09 = strDataTemp(9)
      Else
         '非申請案要新增已提申
         strCP10 = "1909"
         m_NewCP09 = AutoNo("C", 6)
         strDataTemp(13) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
         strDataTemp(12) = GetSalesArea(strDataTemp(13))
         strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp20,cp26," & _
            "cp32,cp27,cp43,cp145) values('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'" & _
            "," & strSrvDate(1) & ",'" & m_NewCP09 & "','" & strCP10 & "','" & strDataTemp(12) & "','" & strDataTemp(13) & "','" & strUserNum & "','N','N','N'," & strSrvDate(1) & _
            ",'" & cp(9) & "','" & IIf(chkChoose(0).Value, "Y", "") & "')"
         cnnConnection.Execute strSql, intI
      End If
      'Modified by Morgan 2018/9/25 +工程師判發控制
      If Check1.Value = vbChecked Then
         'Modified by Morgan 2018/11/26 Ex:CFP-28457
         '考慮判發人離職問題:工程師>>核稿人>>判發人(若離職最後還是會給王副總)--郭
         'strExc(1) = cp(14)
         strExc(1) = GetEngCaseJudge(cp(9))
         'end 2018/11/26
      Else
         strExc(1) = PUB_GetLetterJudgeNew("1", field(1), strCP10, field(9), cp(10))
      End If
      
      '附件數
      intI = 1
      If chkChoose(0).Value Then intI = intI + 1
      'If chkChoose(1).Value = vbChecked Then intI = intI + 1 'Removed by Morgan 2018/10/11 圖式不會單獨一個檔案--甄妮
      If chkChoose(2).Value Then intI = intI + 1
      If CheckStr(txtCaseField(6)) <> "" Then intI = intI + 1
      'Modified by Morgan 2020/1/7 +傳是否工程師判發參數
      PUB_AddLetterProgress m_NewCP09, intI, IIf(txtCaseField(4) = "N", False, True), strExc(1), False, field(26), strCP10, field(75), IIf(chkChoose(2).Value, True, False), , , IIf(Check1.Value = vbChecked, True, False)
      m_bolAddLP = True
      
      'Added by Morgan 2018/9/18 報價(目前僅美國暫時申請且中文送件有,固定5點)
      If m_bol118C Then
         PUB_UpdateLP2930 m_NewCP09, txtCaseField(14), "5"
      End If
      'end 2018/9/18
      
      'Added by Morgan 2018/10/2
      If str1205CP09 <> "" And str1205CP14 <> "" And Left(cp(12), 1) <> "F" Then
         Pub_COrderInform str1205CP09
         bolSavPdf = True
      End If
      'end 2018/10/2
   End If
   'end 2018/6/8
   
   'Added by Morgan 2019/10/2
   'Modified by Morgan 2020/3/11 +427
   If txtCaseField(15) = "Y" Then
      strSql = "update caseprogress set cp47=" & DBDATE(cp(47)) & " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('416','427') and cp27>19221111 and cp57 is null and cp47 is null"
      cnnConnection.Execute strSql, intI
      'Added by Morgan 2019/10/8
      strSql = "update nextprogress set np06='Y' where (np01,np02,np03,np04,np05) in (select cp09,cp01,cp02,cp03,cp04 from caseprogress where cp01='" & cp(1) & _
         "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('416','427') and cp47>0) and np07 IN (" & 提申 & "," & 收達 & ",995,996) and np06 is null"
      cnnConnection.Execute strSql, intI
      'end 2019/10/8
   End If
   If txtCaseField(16) = "Y" Then
      strSql = "update caseprogress set cp47=" & DBDATE(cp(47)) & " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('106','121') and cp27>19221111 and cp57 is null and cp47 is null"
      cnnConnection.Execute strSql, intI
      'Added by Morgan 2019/10/8
      strSql = "update nextprogress set np06='Y' where (np01,np02,np03,np04,np05) in (select cp09,cp01,cp02,cp03,cp04 from caseprogress where cp01='" & cp(1) & _
         "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('106','121') and cp47>0) and np07 IN (" & 提申 & "," & 收達 & ",995,996) and np06 is null"
      cnnConnection.Execute strSql, intI
      'end 2019/10/8
   End If
   If txtCaseField(17) = "Y" Then
      strSql = "update caseprogress set cp47=" & DBDATE(cp(47)) & " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('416','215') and cp27>19221111 and cp57 is null and cp47 is null"
      cnnConnection.Execute strSql, intI
      'Added by Morgan 2019/10/8
      strSql = "update nextprogress set np06='Y' where (np01,np02,np03,np04,np05) in (select cp09,cp01,cp02,cp03,cp04 from caseprogress where cp01='" & cp(1) & _
         "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('416','215') and cp47>0) and np07 IN (" & 提申 & "," & 收達 & ",995,996) and np06 is null"
      cnnConnection.Execute strSql, intI
   End If
   'end 2019/10/2
   
   'Added by Morgan 2020/3/4
   '美國IDS輸入申請案號時若有IDS未發文則設定所限為系統日+1週,法限不設
   'Removed by Morgan 2022/2/21 取消,已有其他管控--郭
   'Modified by Lydia 2024/06/28 恢復此本所期限之設定--郭
   If field(9) = "101" And txtCaseField(1) <> "" Then
      strExc(1) = PUB_GetWorkDay1(CompDate(2, 7, strSrvDate(1)), True)
      'Modified by Lydia 2024/06/28 原IDS的本限早於此預設本限，則請以原有本限為主
      'strSql = "update caseprogress set cp06=" & strExc(1) & " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='214' and cp06 is null and cp27||cp57 is null"
      strSql = "update caseprogress set cp06=" & strExc(1) & " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='214' and (cp06 is null or cp06>" & strExc(1) & ") and cp27||cp57 is null"
      cnnConnection.Execute strSql, intI
   End If
   'end 2024/06/28
   'end 2022/2/21
   'end 2020/3/4
   
   'Added by Morgan 2020/9/23 緬甸
   '後續刊登廣告期限：「前一次刊登日("刊登廣告"的提申日) +3年」為下一次法定期限，本所期限為法定期限提早1個月。
   If field(9) = "048" And cp(10) = "951" Then
      strExc(1) = CompDate(0, 3, txtCaseField(0)) '法限
      strExc(2) = PUB_GetWorkDay1(CompDate(1, -1, strExc(1)), True) '所限
      strSql = "Select NP01,NP07,NP22 From Nextprogress Where NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP07='951' AND NP06 IS NULL  ORDER BY NP22 DESC"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " WHERE NP01='" & RsTemp("NP01") & "' AND NP22=" & RsTemp("NP22")
      Else
         strSql = "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
            " select '" & m_NewCP09 & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','951'," & strExc(2) & "," & strExc(1) & ",'" & strDataTemp(13) & "',newNP22 from dual,(select nvl(max(np22),0)+1 newNP22 from nextprogress)"
      End If
      cnnConnection.Execute strSql, intI
   End If
   'end 2020/9/23
   
   cnnConnection.CommitTrans
   
   'Add By Sindy 2019/7/24 已提申代理人撰稿發E-Mail通知工程師
   If cp(10) = "958" Then
      strExc(0) = "select cp14 from caseprogress where cp09=(select cp43 from caseprogress where cp09='" & cp(9) & "' and cp43 is not null)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If ClsPDGetCustomerNameAndAddress(GetPrjPeopleNum1(cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)), oCustName) Then
         End If
         oContext = "本所案號：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & vbCrLf & vbCrLf & "專利名稱：" & cboCaseName & vbCrLf & vbCrLf & "申請人　：" & oCustName & vbCrLf & vbCrLf & "案件性質：" & GetPrjState4(cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4), cp(10)) & vbCrLf & vbCrLf & "提申日　：" & ChangeWStringToWDateString(DBDATE(txtCaseField(0)))
         PUB_SendMail strUserNum, RsTemp.Fields("cp14"), strReceiveNo, "通知 " & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & "「" & cboCaseName & "」 " & GetPrjState4(cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4), cp(10)) & "已完成！", oContext, , strFullFileName
      End If
   End If
   '2019/7/24 END
   
   SaveDatabase = True
   
   If str1205CP09 <> "" Then
      g_PrtForm001.PrintCFForm str1205CP09, , bolSavPdf
   End If
   
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   SaveDatabase = False
   
End Function

'傳回起算日之日期
Private Function GetStartDate(ByRef strTemp As String, cp() As String, field() As String) As String
   Dim strDate As String
   
   Select Case strTemp
             Case 收文日
                        GetStartDate = cp(5)
             Case 申請日
                        'Modify by Morgan 2011/9/15 英國,EPC 的分割案要從母案的申請日起算
                        'GetStartDate = field(10)
                        If (field(9) = "201" Or field(9) = "221") And cp(10) = "307" Then
                           GetStartDate = PUB_DivAppDate(cp(1), cp(2), cp(3), cp(4))
                        'Added by Morgan 2016/1/8
                        '印度,馬來西亞設計案延展期限從最早優先權日起算
                        ElseIf field(8) = "3" And (field(9) = "018" Or field(9) = "040") Then
                           strDate = PUB_GetFirstPriDate(field)
                           If strDate = "" Then
                              GetStartDate = field(10)
                           Else
                              GetStartDate = strDate
                           End If
                        Else
                           GetStartDate = field(10)
                        End If
                        'end 2011/9/15
             Case 發文日
                        GetStartDate = cp(27)
             Case 准駁日
                        GetStartDate = cp(25)
             Case 公告日
                        GetStartDate = field(14)
             Case 發證日
                        GetStartDate = field(21)
             'Add by Morgan 2010/3/3
             Case 公開日
                        GetStartDate = field(12)
                        If GetStartDate = "" And field(10) <> "" Then
                           If PUB_GetOpenDate(cp(1), cp(2), cp(3), cp(4), , strDate, , , field(10)) Then
                              GetStartDate = strDate
                           End If
                        End If
   End Select
   If GetStartDate = "" Then
   '   MsgBox "找不到案件起算日!!", vbCritical
   End If
End Function

Private Sub ReadAllData()
   Dim varSaveCursor, strTemp As String, strTemp1 As String

On Error GoTo HndErr
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   'Modify by Morgan 2006/10/19 改不Call Dll
   'If objPublicData.ReadAllData(frm050103_1.grdDataList.TextMatrix(frm050103_1.grdDataList.Row, 0), cp(), field(), intCaseKind, intPWhere) Then
   ReDim cp(TF_CP) As String
   cp(9) = frm050103_1.grdDataList.TextMatrix(frm050103_1.grdDataList.row, 0)
   If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
   'end 2006/10/19
   
      'Add by Morgan 2005/12/20
      'Modify by Morgan 2006/2/8 加118
      'Modify by Morgan 2006/10/26 加120植物新品種
      'Modify by Morgan 2010/5/24 +109
      'Modified by Morgan 2016/8/10 +122CA申請 Ex.CFP-25889-1 --禧佩
      If cp(10) = 發明申請 Or cp(10) = 新型申請 Or cp(10) = 設計申請 Or cp(10) = 追加申請 Or cp(10) = 聯合申請 Or cp(10) = CIP申請 Or cp(10) = 分割 Or cp(10) = CPA申請 Or cp(10) = 再發行 Or cp(10) = "118" Or cp(10) = "120" Or cp(10) = "109" Or cp(10) = "122" Then
         m_bolNewApp = True
      Else
         m_bolNewApp = False
      End If
   
      'Added by Morgan 2019/11/28
      If cp(3) = "0" And (cp(10) = "113" Or cp(10) = "122") Then
         m_bolXCACase = True
      Else
         m_bolXCACase = False
      End If
      'end 2019/11/28
   
      lblCaseField(0) = MergeString(cp(1), cp(2), cp(3), cp(4))
      lblCaseField(1) = field(9)
      'Modify by Morgan 2005/11/21 專利須與CPS分開
      If intCaseKind = 專利 Then
         lblCaseField(2) = field(8)
         lblCaseField(3) = field(72)
         '2005/7/11 MODIFY BY SONIA判斷是否PCT案
         'txtCaseField(0) = TransDate(cp(47), 1)
         
         'Modify by Morgan 2005/12/20 加控制新案
         'If field(46) <> "" Then
         If m_bolNewApp And field(46) <> "" Then
         '2005/12/20
            txtCaseField(0) = TransDate(field(10), 1)
            'Modify by Morgan 2011/7/8
            'txtCaseField(12) = TransDate(cp(47), 1)
            If txtCaseField(0) <> "" Then
               txtCaseField(0).Enabled = False
            End If
            txtCaseField(12).Tag = TransDate(cp(47), 1)
            'end 2011/7/8
            txtCaseField(12).Enabled = True
            txtCaseField(12).Visible = True
            Label20.Visible = True
            Label5(1) = "申請日："
            
         'Add by Morgan 2011/7/1
         ElseIf cp(10) = 分割 Then
            Label5(1) = "申請日："
            txtCaseField(0) = TransDate(field(10), 1)
            If txtCaseField(0) = "" Then
               txtCaseField(0) = TransDate(PUB_DivAppDate(field(1), field(2), field(3), field(4), True), 1)
            'Modified by Morgan 2024/4/1 有抓到母案申請日時要控制不可修改，否則後面檢查新案時又會被清除 Ex:CFP-034275
            'Else
            End If
            If txtCaseField(0) <> "" Then
            'end 2024/4/1
               txtCaseField(0).Enabled = False
            End If
            Label20.Caption = "分割案提交日："
            Label20.Visible = True
            txtCaseField(12).Tag = TransDate(cp(47), 1)
            txtCaseField(12).Enabled = True
            txtCaseField(12).Visible = True
         
         'Added by Morgan 2019/11/28
         ElseIf m_bolXCACase = True Then
            Label5(1) = "申請日："
            txtCaseField(0) = TransDate(field(10), 1)
            If txtCaseField(0) <> "" Then
               txtCaseField(0).Enabled = False
            End If
            Label20.Caption = "接續案提交日："
            Label20.Visible = True
            txtCaseField(12).Tag = TransDate(cp(47), 1)
            txtCaseField(12).Enabled = True
            txtCaseField(12).Visible = True
         'end 2019/11/28
         Else
            txtCaseField(0) = TransDate(cp(47), 1)
            txtCaseField(12).Enabled = False
            txtCaseField(12).Visible = False
            txtCaseField(12) = ""
            Label20.Visible = False
         End If
         '2005/7/11 END
         
         m_NP07 = ""
         
         If InStr(NewCasePtyList & "601,605,606,607", cp(10)) > 0 Then 'Added by Morgan 2013/10/1 新案或領證年費提申才要帶下次繳費日,否則若期限已過期會無法作業 Ex.CFP-23182 答辯--慧汶
            If GetNP07(field(9), field(8), strTemp) Then
               m_NP07 = strTemp
               strTemp = ""
               strTemp1 = ""
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, , , False) = 0 Then
               'Modified by Morgan 2013/10/23
               'If ClsPDGetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, , , False) = 0 Then
               If ClsPDGetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, , , False, , field(10), field(21), field(72)) = 0 Then
               
                  'Add by Morgan 2010/3/3 '起算日為公開日且公開日為申請日起算者也要掛年費期限
                  If Val(strTemp) = 公開日 Then
                     strExc(0) = "select 1 from nation where decode('" & field(8) & "','1',na32,'2',na34,'3',na36)=2"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If blnUpdateNP(cp(1) & cp(2) & cp(3) & cp(4)) = True Then
                           strTemp = 申請日
                        End If
                     End If
                  End If
                     
                  If Val(strTemp) = 申請日 Then
                     strExc(0) = "select np09 from nextprogress where np02=" + CNULL(cp(1)) + " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05=" + CNULL(cp(4)) + " and np07=" & m_NP07 + " AND np06 IS null"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If intI = 1 And Not IsNull(RsTemp.Fields("np09")) Then
                        txtCaseField(3) = TransDate(RsTemp.Fields("np09"), 1)
                     Else
                        '2007/5/8 modify by sonia CFP-19619分割
                        'If cp(10) <= "105" Then
                        If cp(10) = 發明申請 Or cp(10) = 新型申請 Or cp(10) = 設計申請 Or cp(10) = 追加申請 Or cp(10) = 聯合申請 Or cp(10) = CIP申請 Or cp(10) = 分割 Or cp(10) = CPA申請 Or cp(10) = 再發行 Or cp(10) = "118" Or cp(10) = "120" Then
                        '2007/5/8 end
                           'Add by Morgan 2007/3/16 加判斷非准後繳年費的才要預設下次繳費日
                           'Modified by Morgan 2016/1/13
                           'PCT進紐西蘭之發明案若PCT申請日是在2014/9/13新法實施前則年費為准後繳且期限為自發證日起4個月
                           If field(46) = "Y" And field(9) = "016" And field(8) = "1" And DBDATE(field(10)) < "20140913" Then
                              txtCaseField(3) = ""
                           ElseIf blnUpdateNP(cp(1) & cp(2) & cp(3) & cp(4)) = True Then
                           'end 2007/3/16
                              GetNextDate
                           End If
                        End If
                     End If
                     
                  '92.1.11 add by sonia
                  Else
                     If cp(10) >= "605" And cp(10) <= "607" Then
                        strExc(0) = "select np09 from nextprogress where np02=" + CNULL(cp(1)) + " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05=" + CNULL(cp(4)) + " and np07=" & m_NP07 + " AND np06 IS null"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                        If intI = 1 And Not IsNull(RsTemp.Fields("np09")) Then
                           txtCaseField(3) = TransDate(RsTemp.Fields("np09"), 1)
                        End If
                     End If
                  '92.1.11 end
                  End If
               End If
            End If
         End If
         
'Remove by Morgan 2007/3/16 搬到上面做 因為准後繳的國家年費提申要帶下次繳費日(定稿要用)
'         '93.8.26 add by sonia 判斷是否准後繳年費, 准後繳年費不預設下次繳費日
'         If blnUpdateNP(cp(1) & cp(2) & cp(3) & cp(4)) = False Then
'            txtCaseField(3) = ""
'         End If
'         '93.8.26 end
'end 2007/3/16
         
         Select Case cp(10)
            Case 答辯
               txtCaseField(6) = "答辯理由書影本乙份"
               '92.11.18 CANCEL BY SONIA 改為預設來函收文日
               'If txtCaseField(0) = "" Then txtCaseField(0) = strSrvDate(2)
            'Added by Morgan 2016/3/3 +126 期末拋棄
            Case "126"
               txtCaseField(6) = "期末拋棄理由書影本乙份"
            'Add by Morgan 2009/6/23
            Case 提供前案資料
               txtCaseField(6) = frm050103_1.grdDataList.TextMatrix(frm050103_1.grdDataList.row, 3) & "影本乙份"

            'Modify by Morgan 2009/6/23 提供前案資料抽出來,因為資料
            'Case 修正, 補充說明, 提供前案資料, 變更
            Case 修正, 補充說明, 變更
               txtCaseField(6) = frm050103_1.grdDataList.TextMatrix(frm050103_1.grdDataList.row, 3) & "資料影本乙份"
               '92.11.18 CANCEL BY SONIA 改為預設來函收文日
               'If txtCaseField(0) = "" Then txtCaseField(0) = strSrvDate(2)
            'Modify by Morgan 2006/10/26 加120植物新品種
            Case 發明申請, 新型申請, 設計申請, 追加申請, 聯合申請, CIP申請, "120"
               txtCaseField(0) = TransDate(field(10), 1)
               txtCaseField(1) = field(11)
            '92.12.18 ADD BY SONIA
            Case 選取
               txtCaseField(6) = "選取理由書影本乙份"
            '92.12.18 END
            'Added by Lydia 2016/08/26 再考量試行計畫(AFCP 2.0)
            Case 438
               txtCaseField(6) = "再考量試行計畫影本乙份"
            Case Else
               '92.11.18 CANCEL BY SONIA 改為預設來函收文日
               'If txtCaseField(0) = "" Then txtCaseField(0) = strSrvDate(2)
               
               'Added by Morgan 2024/4/25
               'EPC德國子案指定國註冊費提申時要輸入國家階段案號-->更新子案申請號
               If field(9) = "231" And field(4) <> "00" And cp(10) = "224" Then
                  m_bolEPC_DE_CHK = True
                  Label6(0) = "國家階段案號："
                  strExc(0) = "select pa11 from patent where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & cp(3) & "' and pa04='00'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     '若申請案號與母案不相同代表尚已經輸入
                     If field(11) <> "" & RsTemp(0) Then
                        txtCaseField(1) = field(11)
                     End If
                  End If
               End If
               'end 2024/4/25
         End Select
         'Add By Cheng 2002/07/31
         'Modify by Morgan 2006/11/14 加CA申請
         'If field(9) = 美國國家代號 And (cp(10) = 發明申請 Or cp(10) = CIP申請 Or cp(10) = 分割) Then
         If field(9) = 美國國家代號 And (cp(10) = 發明申請 Or cp(10) = CIP申請 Or cp(10) = 分割 Or cp(10) = "122") Then
            Me.Label5(3).Visible = True
            Me.txtCaseField(10).Visible = True
            If field(12) <> "" Then txtCaseField(10) = TransDate(field(12), 1) '若有公開日時預設 Add by Morgan 2011/10/14--禧佩
         End If
                  
         'Add by Morgan 2005/6/17日韓的發明申請若未收文實審則可輸實審費用且預設PATENTYEARFEE的費用
         Set416Obj
      Else
         txtCaseField(0) = TransDate(cp(47), 1)
         txtCaseField(12).Enabled = False
         txtCaseField(12).Visible = False
         txtCaseField(12) = ""
         Label20.Visible = False
      End If
   
      If cp(45) = "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseThatCode(cp()) = False Then GoTo Err1
         If ClsPDGetCaseThatCode(cp()) = False Then GoTo err1
      End If
      txtCaseField(2) = cp(45)
      SetNameToCombo cboCaseName, field(5), field(6), field(7)
      
      'Added by Morgan 2012/4/24
      If cp(1) = "CFP" And cp(10) = "123" Then
         txtFavDt.Enabled = True
      Else
         txtFavDt.Enabled = False
      End If
      'end 2012/4/24
      
      'Added by Morgan 2015/12/16
      If cp(10) = 美國暫時申請 Then
         txtCaseField(14).Enabled = True
      Else
         txtCaseField(14).Enabled = False
      End If
      'end 2015/12/16
      
      'Added by Morgan 2018/9/27
      If cp(145) = "Y" Then
         'Added by Morgan 2020/8/18
         If cp(10) >= "605" And cp(10) <= "607" Then
            chkChoose(2).Value = True
         Else
         'end 2020/8/18
            chkChoose(0).Value = True
         End If
      End If
      'end 2018/9/27
                        
   Else
err1:
      bolLeave = True
      intLeaveKind = 1
      Unload Me
   End If
HndErr:
   ErrorMsg
   Screen.MousePointer = varSaveCursor
End Sub

Private Function GetLast(ByVal strValue As String) As String
   Dim strTemp As String
   Dim nIndex As Integer
   Dim aryDate
   
   strTemp = Empty
   aryDate = Split(strValue, ",")
   For nIndex = 0 To UBound(aryDate)
      If Not IsEmptyText(aryDate(nIndex)) Then
         strTemp = aryDate(nIndex)
      End If
   Next nIndex
   GetLast = strTemp
End Function

Private Function GetNext(ByVal strValue As String, ByVal strLast As String) As String
   Dim strTemp As String
   Dim nIndex As Integer
   Dim aryDate
   strTemp = Empty
   aryDate = Split(strValue, ",")
   For nIndex = 0 To UBound(aryDate)
      If Not IsEmptyText(aryDate(nIndex)) Then
         If Val(aryDate(nIndex)) > Val(strLast) Then
            strTemp = aryDate(nIndex)
            Exit For
         End If
      End If
   Next nIndex
   GetNext = strTemp
End Function


'計算下次繳費日
Private Sub GetNextDate()
   Dim strTemp1 As String, strTemp2 As String, strCaseProperty As String, strTemp As String
   Dim varTemp As Variant, dobDateAdd As Double, strStartDate As String, intYears As Integer
   Dim strReduceOne As String 'Added by Morgan 2019/11/29
   
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetNationTax(Val(field(8)), field(9), strTemp, strTemp1, strTemp2, strCaseProperty) Then
   'Modified by Morgan 2019/12/3 +strReduceOne
   If ClsPDGetNationTax(Val(field(8)), field(9), strTemp, strTemp1, strTemp2, strCaseProperty, , strReduceOne) Then
      Dim strLast As String
      Dim strNext As String
      strLast = GetLast(field(72))
      strNext = GetNext(strTemp1, strLast)
      If Not IsEmptyText(strNext) Then
         dobDateAdd = CDbl(strNext)
         strStartDate = GetStartDate(strTemp, cp(), field())
         'end 2011/9/15
         If strStartDate <> "" Then
            If m_NP07 = "605" Then dobDateAdd = dobDateAdd - 1
             
             '91.11.10 MODIFY BY SONIA
             'txtCaseField(3) = ChangeWStringToTString(CompDate(2, -1, strStartDate))
             'Modified by Morgan 2019/12/4 計算專用期止日要減1天者年費/延展費期限也要減1天,Ex:新加坡設計案
             'txtCaseField(3) = ChangeWStringToTString(CompDate(0, dobDateAdd, strStartDate))
             txtCaseField(3) = ChangeWStringToTString(PUB_GetEndDate(strStartDate, dobDateAdd, strReduceOne, field(9)))
             'end 2019/12/3
             '91.11.10 END
             
             '2005/10/5 ADD BY SONIA沙烏地阿拉伯發明自申請日第二年起須逐年繳費,年費期限為每年3/31
             '2008/3/21 modiby by sonia 沙烏地阿拉伯設計也是
             'If field(9) = "021" And field(8) = "1" Then
             If field(9) = "021" Then
                txtCaseField(3) = ChangeWStringToTString(Mid(ChangeTStringToWString(txtCaseField(3)), 1, 4) + "0331")
             End If
             '2005/10/5 END
         Else
            If Not IsEmptyText(txtCaseField(0)) Then
               If m_NP07 = "605" Then dobDateAdd = dobDateAdd - 1
               
               '91.11.10 MODIFY BY SONIA
               'txtCaseField(3) = ChangeWStringToTString(CompDate(2, -1, strStartDate))
               'Modified by Morgan 2019/12/4 計算專用期止日要減1天者年費/延展費期限也要減1天,Ex:新加坡設計案
               'txtCaseField(3) = ChangeWStringToTString(CompDate(0, dobDateAdd, txtCaseField(0)))
               txtCaseField(3) = ChangeWStringToTString(PUB_GetEndDate(txtCaseField(0), dobDateAdd, strReduceOne, field(9)))
               '91.11.10 END
            End If
         End If
      End If
   End If
      
End Sub

'Add by Morgan 2005/11/17
Private Sub cmdPriority_Click()
   'Modify by Amy 2014/04/17 +, strPriority(5)
   If m_bolRePriDate = True Then
      '第二次不再檢查
      ModifyPriority strPriority(1), strPriority(2), strPriority(3), field(8), , field(1) & field(2) & field(3) & field(4), field(9), , strPriority(4), strPriority(5)
   Else
      m_bolRePriDate = True
      ModifyPriority strPriority(1), strPriority(2), strPriority(3), field(8), m_bolRePriDate, field(1) & field(2) & field(3) & field(4), field(9), , strPriority(4), strPriority(5)
   End If
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String

Select Case Index
   Case 1
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetNation(lblCaseField(Index), strTemp) Then
      If ClsPDGetNation(lblCaseField(Index), strTemp) Then
         lblCountryName.Caption = strTemp
      End If
   Case 2
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
      If ClsPDGetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
         lblTrademarkKind = strTemp
      End If
End Select
End Sub
Private Sub Form_Activate()
   Dim oActObj As Control 'Added by Morgan 2023/9/5
   
   'Add by Morgan 2005/6/17
   If m_bolActive = True Then Exit Sub
   m_bolActive = True
   '2005/6/17 end
   'Modify By Cheng 2002/07/31
   'txtCaseField(0).SetFocus
   txtCaseField(8).SetFocus
   ReadAllData
   
   'Add by Morgan 2005/10/27
   If m_bolNewApp = True Then
      If txtCaseField(0).Enabled = True Then
         txtCaseField(0).Tag = TransDate(field(10), 1)
         txtCaseField(0).Text = Empty
         '不是第一次輸申請日時預設不印通知函,勾收據時再改要印
         If txtCaseField(0).Tag <> Empty Then
            txtCaseField(4) = "N"
         End If
      End If
     
      '讀取優先權資料
      'Modify by Morgan 2007/4/25 加strPriority(4)
      'Modify by Amy 2014/04/17 +, strPriority(5)
      ClsPDReadPriority field, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)
      ' 優先權輸入控制
      If strPriority(1) <> "" Then
         m_bolRePriDate = False
         cmdPriority.Enabled = True
      Else
         cmdPriority.Enabled = False
      End If
      '2005/11/17 END
   'Add by Morgan 2006/2/8
   Else
      cmdPriority.Enabled = False
   End If
   '2005/10/27 end
   
   'Modified by Morgan 2020/8/17 取消預設，改自行勾選
   'If cp(10) = "605" Then chkChoose(2).Value = vbChecked 'Added by Morgan 2018/6/27 CFP電子化 -年費提申預設有收據
   '年費沒收據不出定稿
   If cp(10) >= "605" And cp(10) <= "607" Then
      If chkChoose(2).Value = False Then
         txtCaseField(4) = "N"
         txtCaseField(4).Enabled = False
      End If
   '指定國註冊費提申不出定稿，輸入1607註冊登記才要通知客戶
   ElseIf cp(10) = "224" Then
      txtCaseField(4) = "N"
      txtCaseField(4).Enabled = False
   End If
   'end 2020/8/17
   
   'Added by Morgan 2018/9/26 工程師判發
   If InStr("107、126、203、204、205、206、207、208、214、218、402、424、431、438、501、502、503、801、802、803、804、805、806", cp(10)) > 0 Then
      Check1.Value = vbChecked
   End If
   'end 2018/9/26

   'add by sonia 2019/7/30 PCT進入國家階段提醒
   If InStr(NewCasePtyList, cp(10)) > 0 And field(46) = "Y" Then
      MsgBox "本案為PCT國際申請案進入國家階段案件，請確認代理人提交專利局的PCT相關資料是否正確無誤。"
   End If
   'end 2019/7/30
   
   'Added by Morgan 2023/9/5
   '內部收文詢問是否要通知客戶
   If txtCaseField(4) = "" And Left(cp(9), 1) = "B" Then
      strExc(0) = PUB_AskBKindLetter(cp(1), cp(9), cp(10), 2)
      If txtCaseField(4) <> strExc(0) Then
         txtCaseField(4) = strExc(0)
         If txtCaseField(8).Enabled Then txtCaseField(8).SetFocus
      End If
   End If
   'end 2023/9/5
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   'Add By Sindy 2017/12/28
   m_strIR01 = frm050103_1.m_strIR01
   m_strIR02 = frm050103_1.m_strIR02
   m_strIR03 = frm050103_1.m_strIR03
   m_strIR04 = frm050103_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/28 END
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2009/8/17
   If m_strIR01 <> "" Then intLeaveKind = 0 'Add By Sindy 2017/10/17
'   'Add By Sindy 2016/10/13
'   If Me.m_strIR01 = "" Then
'   '2016/10/13 END
      If intLeaveKind = 1 Then
         frm050103_1.Show
         frm050103_1.Clear
      ElseIf intLeaveKind = 0 Then
        Unload frm050103_1
      ElseIf intLeaveKind = 2 Then
         frm050103_1.Show
      End If
'   End If
   'Add By Cheng 2002/07/18
   Set frm050103_2 = Nothing
   'Modify By Sindy 2017/10/17
   'frm050103_1.cmdOK(2).Default = True
   If m_strIR01 = "" Then frm050103_1.cmdok(2).Default = True
   '2017/10/17 END
End Sub

'Add by Morgan 2009/6/23
Private Sub txtCaseField_Change(Index As Integer)
   If Index = 7 Then
      If cp(10) = 提供前案資料 Then
         If txtCaseField(Index) = "N" Then
            txtCaseField(6) = "代理人回覆該國專利局「本案並無任何前案或其審查結果可資提供」影本乙份"
         Else
            txtCaseField(6) = frm050103_1.grdDataList.TextMatrix(frm050103_1.grdDataList.row, 3) & "影本乙份"
         End If
      End If
   End If
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
            'Modify by Morgan 2004/4/13
            '申請案號可輸入小寫
            'Case 1, 2, 4, 5, 7
            Case 2, 4, 5, 7
               KeyAscii = UpperCase(KeyAscii)
            'Add By Cheng 2002/07/31
            Case 9
               KeyAscii = UpperCase(KeyAscii)
               If KeyAscii <> 8 And KeyAscii <> 89 Then
                  KeyAscii = 0
               End If
            'Add by Morgan 2005/6/17
            Case 11
               If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                  Beep
                  KeyAscii = 0
               End If
            'Added by Morgan 2019/10/1
            Case 15, 16, 17
               KeyAscii = UpperCase(KeyAscii)
               If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
                  KeyAscii = 0
               End If
End Select
End Sub
Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = -1 Then
      Cancel = True
   End If
   If Cancel Then txtCaseField_GotFocus (Index)

   'Add by Morgan 2011/7/8
   If Cancel = False And Index = 12 Then
      If CheckReKey(txtCaseField(Index)) = True Then
         txtCaseField(Index).Tag = txtCaseField(Index)
      Else
         Cancel = True
      End If
   End If
   
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, strCusTemp As String

CheckKeyIn = -1
Select Case intIndex
             Case 0 '提申日
               'Add by Morgan 2008/1/24 存檔前再檢查
               If txtCaseField(intIndex) = "" Then
                  CheckKeyIn = 1
                  Set416Obj 'Add by Morgan 2010/7/13
                  Exit Function
               End If
               
                  If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                     'Add by Morgan 2007/1/10 加控制不可大於系統日
                     If Val(DBDATE(txtCaseField(intIndex).Text)) > Val(strSrvDate(1)) Then
                        MsgBox "提申日不可大於系統日!!!", vbExclamation + vbOKOnly
                     Else
                     'end 2007/1/10
                           'Add by Morgan 2004/5/13
                           '非新申請案不必確認
                           CheckKeyIn = 1
                           '2007/5/8 modify by sonia
                           'If cp(10) <= "105" Then
                           If cp(10) = 發明申請 Or cp(10) = 新型申請 Or cp(10) = 設計申請 Or cp(10) = 追加申請 Or cp(10) = 聯合申請 Or cp(10) = CIP申請 Or cp(10) = 分割 Or cp(10) = CPA申請 Or cp(10) = 再發行 Or cp(10) = "118" Or cp(10) = "120" Then
                           '2007/5/8 end
                              If CheckReKey(txtCaseField(intIndex)) Then
                                 CheckKeyIn = 1
                                 'Add by Morgan 2006/8/3 馬來西亞發明舊法(申請日<2001/8/1)年費由發證日起算
                                 If field(9) = "018" And field(8) = "1" And Val(TransDate(txtCaseField(0), 2)) < 20010801 Then
                                    txtCaseField(3) = ""
                                 Else
                                    '91.10.27 ADD BY SONIA
                                    'edit by nickc 2007/02/02 不用 dll 了
                                    'If cp(10) <= "105" And objPublicData.GetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, , , False) = 0 Then
                                    '2007/5/8 modify by sonia
                                    'If cp(10) <= "105" And ClsPDGetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, , , False) = 0 Then
                                    If cp(10) = 發明申請 Or cp(10) = 新型申請 Or cp(10) = 設計申請 Or cp(10) = 追加申請 Or cp(10) = 聯合申請 Or cp(10) = CIP申請 Or cp(10) = 分割 Or cp(10) = CPA申請 Or cp(10) = 再發行 Or cp(10) = "118" Or cp(10) = "120" Then
                                       'Modified by Morgan 2013/10/23
                                       'If ClsPDGetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, , , False) = 0 Then
                                       If ClsPDGetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, , , False, , field(10), field(21), field(72)) = 0 Then
                                    '2007/5/8 end
                                    
                                          'Add by Morgan 2010/3/3 '起算日為公開日且公開日為申請日起算者也要掛年費期限
                                          If Val(strTemp) = 公開日 Then
                                             strExc(0) = "select 1 from nation where decode('" & field(8) & "','1',na32,'2',na34,'3',na36)=2"
                                             intI = 1
                                             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                                             If intI = 1 Then
                                                If blnUpdateNP(cp(1) & cp(2) & cp(3) & cp(4)) = True Then
                                                   strTemp = 申請日
                                                End If
                                             End If
                                          End If
                                          
                                          If Val(strTemp) = 申請日 Then
                                             'Add by Morgan 2006/8/10 申請日須用變更後的提申日
                                             field(10) = TransDate(txtCaseField(intIndex), 2)
                                             
                                             '93.8.26 modify by sonia 判斷是否准後繳年費, 准後繳年費不預設下次繳費日
                                             'GetNextDate
                                             'Modified by Morgan 2016/1/13
                                             'PCT進紐西蘭之發明案若PCT申請日是在2014/9/13新法實施前則年費為准後繳且期限為自發證日起4個月
                                             If field(46) = "Y" And field(9) = "016" And field(8) = "1" And DBDATE(field(10)) < "20140913" Then
                                                txtCaseField(3) = ""
                                             ElseIf blnUpdateNP(cp(1) & cp(2) & cp(3) & cp(4)) = False Then
                                                txtCaseField(3) = ""
                                             Else
                                                GetNextDate
                                             End If
                                             '93.8.26 end
                                          End If
                                       End If
                                    End If
                                    '91.10.27 END
                                 End If
                              Else
                                 CheckKeyIn = 0
                              End If
                              
                           End If
                           'Add by Morgan 2010/7/13
                           If CheckKeyIn = 1 And txtCaseField(0).Tag <> txtCaseField(0).Text Then
                              Set416Obj
                           End If
                           'end 2010/7/13
                           
                           'Add by Morgan 2005/10/27
                           If txtCaseField(0).Text <> Empty Then txtCaseField(0).Tag = txtCaseField(0).Text
                     End If
                  End If
             Case 1 '申請案號
                        If CheckReKey(txtCaseField(intIndex)) Then
                           CheckKeyIn = 1
                        End If
                        
                        'Added by Morgan 2012/8/21 檢查申請案號是否重複
                        If CheckKeyIn = 1 Then
                           If PUB_ChkAppNo(txtCaseField(1).Text, field(1), field(2), field(9)) = True Then
                              CheckKeyIn = 1
                           Else
                              CheckKeyIn = 0
                           End If
                        End If
                        'end 2012/8/21
            
             Case 3 '下次繳費日
                         If txtCaseField(intIndex) = "" Then
                            CheckKeyIn = 1
                            Exit Function
                         End If
                         If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           If Val(txtCaseField(intIndex)) < Val(strSrvDate(2)) Then
                              MsgBox "下次繳費日不可小於系統日期!!!", vbExclamation
                           Else
                              If CheckReKey(txtCaseField(intIndex)) Then
                                 CheckKeyIn = 1
                              End If
                           End If
                         End If
             Case 4
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           '91.7.18 MODIFY BY SONIA
                           'If txtCaseField(intIndex) = "" And txtCaseField(1) = "" Then
                           '   ShowMsg MsgText(1057)
                           '   CheckKeyIn = 0
                           'Else
                              CheckKeyIn = 1
                           'End If
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 5
                        If chkChoose(0).Value And txtCaseField(intIndex) = "" Then
                           ShowMsg MsgText(1055)
                        Else
                           CheckKeyIn = 1
                        End If
             Case 7
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9174)
                        End If
             'Add By Cheng 2002/07/31
             Case 8 '櫃台收文日
                        If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           If Val(txtCaseField(8)) > Val(strSrvDate(2)) Then
                              MsgBox "櫃台收文日不可大於系統日!!!", vbExclamation + vbOKOnly
                           Else
                              CheckKeyIn = 1
                              '92.11.18 ADD BY SONIA
                              Select Case cp(10)
                                 'Modify by Morgan 2006/10/26 加120植物新品種
                                 'Modified by Morgan 2020/2/20 +CA申請
                                 Case 發明申請, 新型申請, 設計申請, 追加申請, 聯合申請, CIP申請, 分割, CPA申請, 再發行, "118", "120", CA申請
                                 Case Else
                                    If txtCaseField(0) = "" Then txtCaseField(0) = txtCaseField(8)
                              End Select
                              '92.11.18 END
                           End If
                        End If
             'Add By Cheng 2002/07/31
             Case 10 '美國預定公開日
                        If Me.txtCaseField(10).Visible = True Then
                            'Modify By Cheng 2002/11/26
'                           If Me.txtCaseField(10).Text <> "" Or (Me.txtCaseField(0).Text > "20001129" And chkChoose(2).Value = vbChecked) Then
                           If Me.txtCaseField(10).Text <> "" Or (Val(Me.txtCaseField(0).Text) > 891129 And chkChoose(2).Value) Then
                              'Add by Morgan 2005/10/27
                              '可不輸但要確認 -- 慧汶
                              If Me.txtCaseField(10).Text = "" Then
                                 If MsgBox("確定不輸美國預定公開日？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                                    CheckKeyIn = 1
                                 End If
                              Else
                              '2005/10/27 end
                                 If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                                    CheckKeyIn = 1
                                 End If
                              End If
                            'Modify By Cheng 2002/11/26
                            Else
                                 CheckKeyIn = 1
                           End If
                        End If
             '2005/7/11 ADD BY SONIA
             'Modify by Morgan 2011/7/1 +與分割案提交日共用
             Case 12 'PCT提交日
                        'Add by Morgan 2008/1/24 存檔前再檢查
                        If txtCaseField(intIndex) = "" Then
                           CheckKeyIn = 1
                           Exit Function
                        End If
                        
                        If Me.txtCaseField(12).Visible = True Then
                              If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                                 'Modify by Morgan 2008/1/24 希臘例外處理
                                 If field(9) = "212" And cp(10) <> 分割 Then
                                    If Val(txtCaseField(12)) <> Val(txtCaseField(0)) And txtCaseField(0) <> "" Then
                                       MsgBox "希臘的PCT提交日必須等於申請日!!!"
                                       
                                    Else
                                       CheckKeyIn = 1
                                    End If
                                 Else
                                 'end 2008/1/24
                                    If Val(txtCaseField(12)) <= Val(txtCaseField(0)) Then
                                       MsgBox Replace(Label20, "：", "") & "不可小於等於申請日!!!", vbExclamation + vbOKOnly
                                    'Add by Morgan 2007/1/10 加控制不可大於系統日
                                    ElseIf Val(DBDATE(txtCaseField(intIndex).Text)) > Val(strSrvDate(1)) Then
                                       MsgBox Replace(Label20, "：", "") & "不可大於系統日!!!", vbExclamation + vbOKOnly
                                    'end 2007/1/10
                                    Else
                                       CheckKeyIn = 1
                                    End If
                                 End If
                              End If
                        End If
             'Modify by Amy 2013/05/22 實審費用不需於定稿帶出且不需輸入,項數也不輸
             'Add by Morgan 2010/6/17
             'Case 13 '項數
             '     CheckKeyIn = 1
             '     If txtCaseField(13).Tag <> txtCaseField(13) Then
             '        If Val(txtCaseField(13)) > 0 Then
             '           Alert416Fee txtCaseField(13), field(9), field(157), field(8), cp(44)
             '        End If
             '     End If
             '     txtCaseField(13).Tag = txtCaseField(13)
             ''2005/7/11 END
             Case Else
                     CheckKeyIn = 1
End Select
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
'edit by nickc 2007/07/11 切換輸入法改用API
'txtCaseField(Index).IMEMode = 2
CloseIme
txtCaseField(Index).SelStart = 0
txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
'儲存未修改前之值至Tag中,供再確認時使用
'Modify by Morgan 2005/10/27
'Modify by Morgan 2011/7/8 +提申日要另外控制
If Not (m_bolNewApp And (Index = 0 Or Index = 12)) Then
   txtCaseField(Index).Tag = txtCaseField(Index)
End If

'Add by Morgan 2010/6/17
'項數離開都提示金額
If Index = 13 Then
   txtCaseField(Index).Tag = ""
End If
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

   'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If

For Each objTxt In Me.txtCaseField
   If objTxt.Visible = True Then
      'Add by Morgan 2008/1/24 必要欄位檢查
      If objTxt.Text = "" And (objTxt.Index = 0 Or objTxt.Index = 12) Then
         MsgBox "必要欄位不可空白！"
         txtCaseField(objTxt.Index).SetFocus
         Exit Function
      End If
      'end 2008/1/24
      Cancel = False
      txtCaseField_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next


   'Added by Morgan 2012/4/24
   If txtFavDt.Enabled Then
      If txtFavDt = "" Then
         MsgBox "請輸入優惠期日期！", vbExclamation
         txtFavDt.SetFocus
         Exit Function
      ElseIf DBDATE(txtFavDt) <> DBDATE(field(140)) Then
         MsgBox "優惠期日期與分案不同！", vbExclamation
         txtFavDt.SetFocus
         Exit Function
      End If
   End If
   'end 2012/4/24
   
   'Added by Morgan 2015/12/16
   m_bol118C = False
   If cp(10) = 美國暫時申請 Then
      If txtCaseField(4) <> "N" Then
         If MsgBox("美國暫時申請案是否為中文送件?", vbYesNo + vbExclamation) = vbYes Then
            m_bol118C = True
         End If
         If m_bol118C And txtCaseField(14) = "" Then
            MsgBox "請輸入提出英譯本及聲明書之費用！", vbExclamation
            txtCaseField(14).SetFocus
            Exit Function
         End If
      End If
   End If
   'end 2015/12/16
   
   'Added by Morgan 2019/10/1 --禧佩
   '輸入通知申請日案號檢查是否有實審、主張優先權已發文未提申
   '提醒且要在畫面上輸入Y/N以避免點錯
   m_OtherCP10 = "" 'Added by Morgan 2020/3/11
   If InStr(m_RefCP10List, cp(10)) > 0 Then
      'Modified by Morgan 2020/12/21
      'If chkChoose(2).Value = 1 Then 若有勾選時也要檢查否則定稿會帶錯 Ex:CFP-032040
      If chkChoose(2).Value Or txtCaseField(15).Text <> "" Or txtCaseField(16) <> "" Or txtCaseField(17) <> "" Then
         'Modified by Morgan 2020/3/11 +加427
         strExc(0) = "select cp09,cp10,cp47 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('416','106','121','427') and cp27>19221111 and cp57 is null order by cp10 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            Do While Not .EOF
               m_OtherCP10 = m_OtherCP10 & .Fields("cp10") & ";" 'Added by Morgan 2020/3/11
               If IsNull(.Fields("cp47")) Then
                  If .Fields("cp10") = "416" Or .Fields("cp10") = "427" Then
                     If txtCaseField(15) = "" Then
                        MsgBox "請一併確認是否同時提出" & IIf(.Fields("cp10") = "416", "實審", "合併檢索及實審") & "！", vbQuestion
                        txtCaseField(15).SetFocus
                        Exit Function
                     End If
                  '不會同時主張國際及國內優先權
                  ElseIf .Fields("cp10") = "106" Or .Fields("cp10") = "121" Then
                     If txtCaseField(16) = "" And IsNull(.Fields("cp47")) Then
                        MsgBox "請一併確認是否同時提出主張優先權！", vbQuestion
                        txtCaseField(16).SetFocus
                        Exit Function
                     End If
                  End If
               End If
               .MoveNext
            Loop
            End With
         End If
      End If
   
   '輸入EPC的回覆檢索報告檢查是否有實審、指定費已發文未提申(都會同時收發文)
   '提醒且要在畫面上輸入Y/N以避免點錯
   ElseIf field(9) = "221" And cp(10) = "218" Then
      strExc(0) = "select cp09,cp10 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('416','215') and cp27>19221111 and cp57 is null and cp47 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If txtCaseField(17) = "" Then
            MsgBox "請一併確認是否同時提出實審及指定費！"
            txtCaseField(17).SetFocus
            Exit Function
         End If
      End If
   End If
   'end 2019/10/1
   
   'Added by Morgan 2024/4/25
   If m_bolEPC_DE_CHK Then
      If txtCaseField(1) = "" Then
         MsgBox "德國子案須輸入國家階段案號！", vbExclamation + vbOKOnly
         txtCaseField(1).SetFocus
         Exit Function
      End If
   End If
   'end 2024/4/25
   
'Modify by Amy 2013/05/22 實審費用不需於定稿帶出且不需輸入
'Add by Morgan 2010/6/14
'If txtCaseField(4) <> "N" Then
   'Add by Morgan 2010/6/17
   'If txtCaseField(13).Visible = True And txtCaseField(13) = "" Then
      'If MsgBox("是否要輸入項數以便估算實審費用？", vbYesNo + vbDefaultButton1) = vbYes Then
         'txtCaseField(13).SetFocus
         'Exit Function
      'End If
   'End If

   'If m_416YN = "Y" And txtCaseField(11) = "" And txtCaseField(11).Visible = True Then
   '   MsgBox "本案客戶通知函要做實審報價，實審費用不可空白！"
   '   If txtCaseField(11).Enabled Then txtCaseField(11).SetFocus
   '   Exit Function
   'End If
'End If
'end 2010/6/14
'end 2013/05/22

TxtValidate = True
End Function

'Modified by Morgan 2021/3/15 +pInEnglish
Private Function GetPayYear(Optional pInEnglish As Boolean) As String
Dim ArrYear
Dim arrDay
Dim ii As Integer
Dim intBegin As Single
Dim intEnd As Single
Dim strPayYear As String

   GetPayYear = ""
   intBegin = -1
   intEnd = 0
   '若繳年費日期, 年度及發文日有資料
   If "" & field(72) <> "" And cp(27) <> "" Then
      ArrYear = Split(field(72), ",")
      arrDay = Split(field(73), ",")
      '若陣列數相同
      If UBound(ArrYear) = UBound(arrDay) Then
         For ii = LBound(arrDay) To UBound(arrDay)
            '判斷同一個發文日繳幾年度
            If DBDATE(arrDay(ii)) = DBDATE(cp(27)) Then
                If intBegin = -1 Then intBegin = ii
                intEnd = ii
            End If
         Next ii
         '若無繳費日與發文日相同時, 則直接抓繳費年度最後一年
         If intEnd = 0 Then
            intEnd = Val(ArrYear(UBound(ArrYear)))
            intBegin = intEnd
         '若有繳費日與發文日相同者, 則抓相同日的相對應的繳費年度
         Else
            intBegin = Val(ArrYear(intBegin))
            intEnd = Val(ArrYear(intEnd))
         End If
      Else
         intEnd = Val(ArrYear(UBound(ArrYear)))
         intBegin = intEnd
      End If
   End If
    
   If cp(10) <> 年費 Then
      For ii = 0 To UBound(ArrYear)
          If Format(ArrYear(ii)) = intBegin Then
             intBegin = ii + 1
          End If
          If Format(ArrYear(ii)) = intEnd Then
             intEnd = ii + 1
             Exit For
          End If
      Next
   End If
   
   
   If intBegin = intEnd Then
      'Modified by Morgan 2015/9/14
      '加拿大繳費年度表示特別 CFP-21305
      'If cp(10) <> 年費 Then
      '   GetPayYear = "第" & intBegin & "次"
      'Else
      '   GetPayYear = "第" & intBegin & "年"
      'End If
      
      ii = 0
      'Modified by Morgan 2015/11/4
      'PUB_GetSpecAnnuityRef field(9), field(8), field(10), cp(10), field(21), field(72), CStr(intBegin), ii
      GetMoneyDate field(8), field(9), field, "", "", , cp(10), ii
      'end 2015/11/4
      strPayYear = PUB_GetYF15(field(9), field(8), "Y000000" & ii, cp(10), CStr(ArrYear(UBound(ArrYear))))
      'end 2015/9/14
   Else
      If cp(10) <> 年費 Then
         strPayYear = "第" & intBegin & "至" & intEnd & "次"
      Else
         strPayYear = "第" & intBegin & "至" & intEnd & "年"
      End If
   End If
   
   'Added by Morgan 2021/3/15
   If pInEnglish Then
      If InStr(strPayYear, "週年") > 0 Then
         strPayYear = "the " & PUB_GetEngPayYear(CStr(intBegin), CStr(intEnd)) & " annuity (" & PUB_GetEngPayYear(CStr(intBegin - 1), CStr(intEnd - 1)) & " anniversary annuity)"
      Else
         strPayYear = "the " & PUB_GetEngPayYear(CStr(intBegin), CStr(intEnd)) & " annuity"
      End If
   End If
   'end 2021/3/15
   
   GetPayYear = strPayYear
End Function

Private Function blnUpdateNP(strCaseNo As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strPA08 As String '專利種類
Dim strPA09 As String '申請國家
Dim strPA16 As String '目前准駁
Dim strPA20 As String '准駁通知日

blnUpdateNP = False
StrSQLa = "Select * From Patent Where " & ChgPatent(strCaseNo)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    strPA08 = "" & rsA("PA08").Value
    strPA09 = "" & rsA("PA09").Value
    strPA16 = "" & rsA("PA16").Value
    strPA20 = "" & rsA("PA20").Value
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    StrSQLa = "Select * From Nation Where NA01='" & strPA09 & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        Select Case strPA08
        Case "1" '發明
            '若非准後繳費者
            If "" & rsA("NA56").Value <> "Y" Then
                '新增下一程序
                blnUpdateNP = True
            Else
                '若有准駁通知日且目前准駁為准的案件
                If "" & strPA20 <> "" And "" & strPA16 = "1" Then
                    '新增下一程序
                    blnUpdateNP = True
                Else
                    '不新增下一程序
                    blnUpdateNP = False
                End If
            End If
        Case "2" '新型
            '若非准後繳費者
            If "" & rsA("NA57").Value <> "Y" Then
                '新增下一程序
                blnUpdateNP = True
            Else
                '若有准駁通知日且目前准駁為准的案件
                If "" & strPA20 <> "" And "" & strPA16 = "1" Then
                    '新增下一程序
                    blnUpdateNP = True
                Else
                    '不新增下一程序
                    blnUpdateNP = False
                End If
            End If
        Case "3" '設計
            '若非准後繳費者
            If "" & rsA("NA58").Value <> "Y" Then
                '新增下一程序
                blnUpdateNP = True
            Else
                '若有准駁通知日且目前准駁為准的案件
                If "" & strPA20 <> "" And "" & strPA16 = "1" Then
                    '新增下一程序
                    blnUpdateNP = True
                Else
                    '不新增下一程序
                    blnUpdateNP = False
                End If
            End If
        End Select
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Removed by Morgan 2018/5/24 沒用了
''Add by Morgan 2005/1/3 EPC指定國家中文字串
'Private Function GetFeeCountry() As String
'
'   Dim stTemp As String, i As Integer, iPos As Integer
'
'On Error GoTo HndErr
'
'   strSql = "Select PA09 From Patent Where PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04<>'" & cp(4) & "' AND PA57 IS NULL"
'   CheckOC3
'   With AdoRecordSet3
'      .CursorLocation = adUseClient
'      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If .RecordCount > 0 Then
'         stTemp = .GetString(adClipString, , , ",")
'         '去掉最後一個逗號
'         stTemp = Left(stTemp, Len(stTemp) - 1)
'         stTemp = PUB_GetNationName(stTemp)
'         stTemp = Replace(stTemp, ",", "、")
'         i = 0: iPos = 0
'         Do
'            iPos = i
'            i = i + 1
'            i = InStr(i, stTemp, "、")
'         Loop While i > 0
'         If iPos > 0 Then stTemp = Left(stTemp, iPos - 1) & "及" & Mid(stTemp, iPos + 1)
'      End If
'   End With
'   GetFeeCountry = stTemp
'
'HndErr:
'   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
'
'End Function

Private Sub Set416Obj()
   If intCaseKind = 專利 Then
'Modify by Morgan 2010/6/14
'      lbl416Fee.Visible = False
'      txtCaseField(11).Visible = False
'      txtCaseField(11) = ""
'      'Modify by Morgan 2006/12/12 韓國新型也要
'      'If (field(9) = "011" Or field(9) = "012") And cp(10) = "101" Then
'      If ((field(9) = "011" Or field(9) = "012") And cp(10) = "101") Or (field(9) = "012" And cp(10) = "102") Then
'         strExc(0) = "SELECT CP09,CP27 FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
'            " AND CP10=" & 實體審查 & " AND CP57 IS NULL"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'         If intI <> 1 Then
'            lbl416Fee.Visible = True
'            txtCaseField(11).Visible = True
'            'Modify by Morgan 2006/12/12 加判斷專利種類
'            'strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & field(9) & "' AND (YF03='" & GetNewFagent(cp(44)) & "' or YF03='Y00000000') AND YF04='416' AND YF05='1' order by YF03 desc"
'            strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & field(9) & "' AND YF02='" & field(8) & "' AND (YF03='" & GetNewFagent(cp(44)) & "' or YF03='Y00000000') AND YF04='416' AND YF05='1' order by YF03 desc"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               '20080924 取消預設實審費用 modify by Toni
'               'txtCaseField(11) = RsTemp.Fields(0)
'               'end 20080924
'            End If
'         End If
'      End If
      
      m_416YN = ""
      'Add by Morgan 2010/6/17
      lblItemCnt.Visible = False
      txtCaseField(13).Visible = False
      txtCaseField(13) = ""
      'end 2010/6/17
      lbl416Fee.Visible = False
      txtCaseField(11).Visible = False
      txtCaseField(11) = ""
      If InStr(m_RefCP10List, cp(10)) > 0 Then
         PUB_UpdExamDate cp(1), cp(2), cp(3), cp(4), , , m_416YN, True, txtCaseField(0)
         'Modify by Amy 2013/05/22 實審費用不需於定稿帶出且不需輸入
         'If m_416YN = "Y" Then
         '   lbl416Fee.Visible = True
         '   txtCaseField(11).Visible = True
         '   'Add by Morgan 2010/6/17
         '   If field(9) = "012" Then
         '      lblItemCnt.Visible = True
         '      txtCaseField(13).Visible = True
         '   End If
         '   'end 2010/6/17
         'End If
         'end 2013/05/22
      End If
'end 2010/6/14
   End If
End Sub

Private Sub txtFavDt_GotFocus()
   TextInverse txtFavDt
   CloseIme
End Sub

Private Sub txtFavDt_Validate(Cancel As Boolean)
   If txtFavDt <> "" Then
      Cancel = Not ChkDate(txtFavDt)
   End If
End Sub

'Added by Morgan 2018/11/26
Private Function GetEngCaseJudge(pCP09 As String) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select EP05,s1.st04 x1,EP04,s2.st04 x2,EP40,s3.st04 x3" & _
      " from engineerprogress,staff s1,staff s2,staff s3" & _
      " where ep02='" & pCP09 & "' and s1.st01(+)=EP05 and s2.st01(+)=EP04 and s3.st01(+)=EP40"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      '外翻人員>>判發人
      If Left(rsQuery("EP05"), 1) = "F" Then
         If rsQuery("x2") = "1" Then
            GetEngCaseJudge = rsQuery("EP04")
         Else
            'Added by Lydia 2023/04/24 修改王副總退休之相關控制
            If strSrvDate(1) >= "20230501" Then
               GetEngCaseJudge = "99050"
            Else
            'end 2023/04/24
               GetEngCaseJudge = "71011"
            End If 'Added by Lydia 2023/04/24
         End If
      Else
         If rsQuery("x1") = "1" Then
            GetEngCaseJudge = rsQuery("EP05")
         ElseIf rsQuery("x2") = "1" Then
            GetEngCaseJudge = rsQuery("EP04")
         ElseIf rsQuery("x3") = "1" Then
            GetEngCaseJudge = rsQuery("EP40")
         Else
            'Modified by Morgan 2020/5/12 粘竺儒退休案件轉由蔡順興處理--王副總
            If rsQuery("EP05") = "84012" Then
               GetEngCaseJudge = "94019"
            Else
               'Added by Lydia 2023/04/24 修改王副總退休之相關控制
               If strSrvDate(1) >= "20230501" Then
                  GetEngCaseJudge = "99050"
               Else
               'end 2023/04/24
                  GetEngCaseJudge = "71011"
               End If 'Added by Lydia 2023/04/24
            End If
            'end 2020/5/12
         End If
      End If
   End If
   Set rsQuery = Nothing
End Function
