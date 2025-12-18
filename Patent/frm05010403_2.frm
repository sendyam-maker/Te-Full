VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm05010403_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "證書號數輸入"
   ClientHeight    =   5880
   ClientLeft      =   2160
   ClientTop       =   2208
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form25"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8520
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   18
      Left            =   7320
      TabIndex        =   19
      Top             =   5560
      Width           =   975
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   17
      Left            =   5280
      TabIndex        =   18
      Top             =   5560
      Width           =   975
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   16
      Left            =   2880
      TabIndex        =   17
      Top             =   5560
      Width           =   975
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   15
      Left            =   1200
      TabIndex        =   16
      Top             =   5560
      Width           =   975
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   14
      Left            =   5280
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   13
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   15
      Top             =   5250
      Width           =   360
   End
   Begin VB.CommandButton cmdCountry 
      Caption         =   "指定國家"
      Height          =   300
      Left            =   5280
      TabIndex        =   14
      Top             =   4860
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   0
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   0
      Top             =   3060
      Width           =   2205
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   1
      Left            =   5280
      MaxLength       =   8
      TabIndex        =   1
      Top             =   3060
      Width           =   975
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   5
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3960
      Width           =   372
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   6
      Left            =   1080
      MaxLength       =   8
      TabIndex        =   9
      Top             =   4260
      Width           =   972
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   10
      Left            =   1080
      TabIndex        =   11
      Top             =   4560
      Width           =   972
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   9
      Left            =   5280
      TabIndex        =   8
      Top             =   3960
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   11
      Left            =   5280
      TabIndex        =   12
      Top             =   4560
      Width           =   852
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   12
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   13
      Top             =   4875
      Width           =   360
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   2
      Left            =   1080
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3360
      Width           =   972
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   3
      Left            =   2310
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3360
      Width           =   972
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   4
      Left            =   5280
      MaxLength       =   8
      TabIndex        =   6
      Top             =   3660
      Width           =   972
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   8
      Left            =   1080
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3660
      Width           =   972
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   7
      Left            =   5280
      MaxLength       =   30
      TabIndex        =   10
      Top             =   4260
      Width           =   2652
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6348
      TabIndex        =   21
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5520
      TabIndex        =   20
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7572
      TabIndex        =   22
      Top             =   70
      Width           =   800
   End
   Begin VB.ComboBox cboCaseName 
      CausesValidation=   0   'False
      Height          =   300
      ItemData        =   "frm05010403_2.frx":0000
      Left            =   1080
      List            =   "frm05010403_2.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   23
      Top             =   930
      Width           =   7335
   End
   Begin VB.Label lblAD 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   3
      Left            =   6360
      TabIndex        =   74
      Top             =   5595
      Width           =   900
   End
   Begin VB.Label lblAD 
      AutoSize        =   -1  'True
      Caption         =   "點數："
      Height          =   180
      Index           =   2
      Left            =   2280
      TabIndex        =   73
      Top             =   5595
      Width           =   540
   End
   Begin VB.Label lblAD 
      AutoSize        =   -1  'True
      Caption         =   "刊登廣告費："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   72
      Top             =   5595
      Width           =   1080
   End
   Begin VB.Label lblAD 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   71
      Top             =   5595
      Width           =   900
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   11
      Left            =   6345
      TabIndex        =   69
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "年費金額："
      Height          =   180
      Index           =   2
      Left            =   4320
      TabIndex        =   68
      Top             =   3360
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "是否修改通知函內容：             （Y:Word）"
      Height          =   180
      Left            =   120
      TabIndex        =   67
      Top             =   5280
      Width           =   3315
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   10
      Left            =   5160
      TabIndex        =   66
      Top             =   2370
      Width           =   1095
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請日："
      Height          =   180
      Index           =   2
      Left            =   4320
      TabIndex        =   65
      Top             =   2370
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "EPC："
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   64
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   7
      Left            =   5160
      TabIndex        =   63
      Top             =   2010
      Width           =   855
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   62
      Top             =   2010
      Width           =   855
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   61
      Top             =   1650
      Width           =   855
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   60
      Top             =   2010
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   59
      Top             =   2010
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   58
      Top             =   1650
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   57
      Top             =   1650
      Width           =   720
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   8
      Left            =   960
      TabIndex        =   56
      Top             =   2370
      Width           =   855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   55
      Top             =   2370
      Width           =   720
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   54
      Top             =   1650
      Width           =   855
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   53
      Top             =   1650
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   52
      Top             =   2010
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   51
      Top             =   2370
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   255
      Index           =   1
      Left            =   6120
      TabIndex        =   50
      Top             =   1650
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   255
      Index           =   3
      Left            =   6120
      TabIndex        =   49
      Top             =   2010
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   48
      Top             =   3060
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "發證日："
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   47
      Top             =   3060
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "專用期間："
      Height          =   180
      Left            =   120
      TabIndex        =   46
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "下次繳費日："
      Height          =   180
      Left            =   4140
      TabIndex        =   45
      Top             =   3660
      Width           =   1080
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "有無證書：         （Y/N）"
      Height          =   180
      Left            =   120
      TabIndex        =   44
      Top             =   3960
      Width           =   1950
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "公告號："
      Height          =   180
      Left            =   4320
      TabIndex        =   43
      Top             =   4260
      Width           =   720
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "公告日："
      Height          =   180
      Left            =   120
      TabIndex        =   42
      Top             =   4260
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "核准日："
      Height          =   180
      Left            =   120
      TabIndex        =   41
      Top             =   3660
      Width           =   720
   End
   Begin VB.Label lblFee1 
      AutoSize        =   -1  'True
      Caption         =   "領證費："
      Height          =   180
      Left            =   120
      TabIndex        =   40
      Top             =   4560
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "影本份數："
      Height          =   180
      Left            =   4320
      TabIndex        =   39
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "點數："
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   38
      Top             =   4560
      Width           =   540
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "是否列印客戶通知函：             （N:不印）"
      Height          =   180
      Left            =   120
      TabIndex        =   37
      Top             =   4920
      Width           =   3315
   End
   Begin VB.Line Line2 
      X1              =   2115
      X2              =   2235
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   36
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   4320
      TabIndex        =   35
      Top             =   1290
      Width           =   900
   End
   Begin MSForms.Label lblNation 
      Height          =   255
      Left            =   6120
      TabIndex        =   34
      Top             =   1290
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   1560
      TabIndex        =   33
      Top             =   1290
      Width           =   2655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Left            =   120
      TabIndex        =   31
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   30
      Top             =   1290
      Width           =   375
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   9
      Left            =   1320
      TabIndex        =   29
      Top             =   2730
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   28
      Top             =   570
      Width           =   2655
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "櫃台收文日："
      Height          =   180
      Left            =   120
      TabIndex        =   26
      Top             =   2730
      Width           =   1080
   End
   Begin VB.Label Label8 
      Caption         =   "申請案號："
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   25
      Top             =   570
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   24
      Top             =   570
      Width           =   2775
   End
   Begin VB.Label lblFee1s 
      BackColor       =   &H00808080&
      Height          =   195
      Left            =   120
      TabIndex        =   70
      Top             =   4575
      Width           =   735
   End
End
Attribute VB_Name = "frm05010403_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/8 改成Form2.0 (cboCaseName,lblPetitionName..)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,pa()存放基本資料檔
Dim cp() As String, pa() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
'StrCountry存放指定國家  strMoneyCountry存放繳費國家 strMoney存放費用
Dim strCountry As String, strMoneyCountry As String, strMoney As String
'bolUpdate是否已有發證書之案件進度
Dim bolUpdate As Boolean
Dim strReceiveNo As String
Dim m_NP22 As String
Dim m_CaseType As String
Dim strNP08 As String
Dim strNP09 As String
'Add By Cheng 2002/02/15
Dim m_strCP09ByCheng As String
'Add By Cheng 2003/04/23
Dim m_blnFirstShow As Boolean
'92.5.27 ADD BY SONIA
Dim m_PayType As String
'Dim m_strNPName As String '下一程序名稱
Dim strPA25 As String 'Add by Morgan 2004/12/10
'93.12.22 ADD BY SONIA
Dim m_strStartDate As String
Dim m_blnCompNextDate As Boolean '是否繼續計算下一次的期限
Dim m_strDate As String
'93.12.22 END
Dim m_NP09_Old As String '原下次繳費期限
Dim stCP12 As String, stCP13 As String 'Add by Morgan 2004/2/6
'Add by Morgan 2007/4/26 香港案控制
Dim m_HKPA01 As String, m_HKPA02 As String, m_HKPA03 As String, m_HKPA04 As String '本所案號
Dim m_HKCP09 As String '收文號
Dim m_HKCP10 As String '案件性質
Dim m_HKCP14 As String '承辦人
Dim m_HKNP22 As String 'NP22
Dim m_HKNP08 As String 'NP08
Dim m_blnCancelClosed As Boolean '是否取消閉卷
Dim m_oPA14 As String, m_oPA20 As String, m_oPA21 As String
Dim m_bolEpcRegNotPaid As Boolean '未繳繳EPC註冊費
Dim m_strEpcRegDueDay As String 'EPC註冊費所限
Dim m_bolAutoIssue As Boolean 'Add by Morgan 2011/8/8 是否自動發證
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim m_bolAddLP As Boolean, m_strCP10 As String, m_strLD18 As String 'Added by Morgan 2018/7/16
Dim m_HK1913CP09 As String, m_HKNP01 As String, m_HKNP09 As String 'Added by Morgan 2018/7/16
Dim m_930DueDate As String 'Added by Morgan 2018/10/30 商業使用聲明法定期限
Dim m_bolNewPlant As Boolean 'Added by Morgan 2025/8/7 是否為植物新品種保護

'Add by Lydia 2014/10/29 , copy frm05010401_3
'Add by Morgan 2008/5/9
'設定費用及點數
Private Sub SetFee()
Dim strMsg As String
   
   'Modified by Mogan 2015/1/14  +lock判斷
   'If txtCaseField(10) = "" Then  'Added by Morgan 2014/11/17
   If txtCaseField(10).Locked = False And txtCaseField(10) = "" Then
   'end 2015/1/14
   
   lblFee1.Tag = ""
   lblFee1.BackColor = &H8000000F
   lblFee1s.Visible = False
            
      strExc(0) = "select yf06,yf07,YF08 from patentyearfee where yf01='" & pa(9) & "'" & _
         " and yf02='" & pa(8) & "' and yf04='601' and yf05='1'"
      'Added by Morgan 2023/3/25
      If strSrvDate(1) >= PA179啟用日 Then
         If pa(179) = "1" Then '大個體
            strExc(0) = strExc(0) & " and yf03='Y00000002'"
         ElseIf pa(179) = "3" Then '微個體
            strExc(0) = strExc(0) & " and yf03='Y00000003'"
         Else
            strExc(0) = strExc(0) & " and yf03='Y00000000'"
         End If
      Else
      'end 2023/3/25
      
         If InStr(pa(91), "大個體") > 0 Then
            strExc(0) = strExc(0) & " and yf03='Y00000002'"
         'Added by Morgan 2013/10/18
         ElseIf InStr(pa(91), "微個體") > 0 Then
            strExc(0) = strExc(0) & " and yf03='Y00000003'"
         'end 2013/10/18
         Else
            strExc(0) = strExc(0) & " and yf03='Y00000000'"
         End If
         
      End If 'Added by Morgan 2023/3/25
      
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '領證費
         txtCaseField(10) = Val(Format("" & RsTemp(0))) + Val(Format("" & RsTemp(1)))
         '點數
         txtCaseField(11) = Val(Format("" & RsTemp(0))) / 1000
         If Not IsNull(RsTemp(2)) Then
            strMsg = RsTemp(2) & vbCrLf & vbCrLf & strMsg
         End If
      End If
      
      If strMsg <> "" Then
         MsgBox strMsg, , "報價提醒！"
      End If
      
      'Add by Morgan 2008/5/16
      If PUB_GetOldPrice(pa(26), pa(9), pa(8), "1603", , , , "3") = True Then '自動發證國家之半年內相同客戶報價
         lblFee1.Tag = "Y"
         lblFee1.BackColor = &HC0FFC0
         lblFee1s.Visible = True
      End If
      
   End If 'Added by Morgan 2014/11/17
End Sub



Private Sub lblFee1_Click()
   If lblFee1.Tag = "Y" Then
      PUB_LabelActive lblFee1, lblFee1s, False
      If PUB_GetOldPrice(pa(26), pa(9), pa(8), "1603", RsTemp, , , "3") = True Then
         PUB_LabelActive lblFee1, lblFee1s
         Set frm880014.grdDataList.Recordset = RsTemp
         Set frm880014.fmParent = Me
         frm880014.Show vbModal
      End If
   End If
End Sub

Private Sub lblFee1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblFee1, lblFee1s
End Sub

Private Sub lblFee1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblFee1, lblFee1s
End Sub

'Add by Morgan 2008/5/12
'Modified by Lydia 2014/11/10 因為單純增加領證費,在轉稿時會缺少欄位,因此與StarLetter整合
Private Sub StartLetter1(ByVal ET03 As String, Optional strNP01 As String, Optional strNP22 As String)
Dim strTxt(1 To 99) As String, iStep As Integer, strTmp As Variant
Dim strTemp1 As String, strStartDate As String, strTemp As Variant
Dim bolTmp As Boolean, StrExt1 As String, StrExt2 As String, i As Integer
Dim iEPC As Integer 'EPC 指定國家順序
Dim iPos As Integer '字元搜尋位置
Dim Jjj As Integer
   
   Jjj = 1
   
   If Val(txtCaseField(10)) > 0 Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'領證費','" & Val(txtCaseField(10)) & "','Y')" 'LCV05=Y=>報價是否顯示=智權可改
      Jjj = Jjj + 1
      
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'領證費點數','" & Val(txtCaseField(11)) & "','')"
      Jjj = Jjj + 1
   End If
 
   strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
      "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用總計','" & Val(txtCaseField(10)) & "','')"
   Jjj = Jjj + 1
   
   strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
      "VALUES ('" & strNP01 & "'," & strNP22 & ",'點數合計','" & Val(txtCaseField(11)) & "','')"
   Jjj = Jjj + 1
 
   If Not ClsLawExecSQL(Jjj - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
'.end 'Add by Lydia 2014/10/29 , copy frm05010401_3
End Sub

Private Sub cmdCountry_Click()
   ModifyMoneyCountry strCountry, strMoneyCountry, strMoney
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String, Optional LCvar As Boolean)
 '********* 90.11.14   nickc
 'Modified by Lydia 2014/11/10
 'Dim strTxt(1 To 8) As String, iStep As Integer, strTmp As String
 'Added by Lydia 2015/05/18
' Dim strTxt(1 To 13) As String, iStep As Integer, strTmp As String
 Dim strTxt(1 To 16) As String, iStep As Integer, strTmp As String
 Dim strTemp1 As String, strStartDate As String, strTemp As String
 Dim bolTmp As Boolean
 Dim rsTmp As New ADODB.Recordset
 '********************************
   EndLetter ET01, m_strCP09ByCheng, ET03, strUserNum
   '********* 90.11.14   nickc
   Dim Jjj As Integer
   

   Jjj = 1
   
   If CheckStr(txtCaseField(4)) <> "" Then
      'Modify by Morgan 2004/3/31
      '土耳其用西元年表示
      'strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
            "','年費法定期限'," & CNULL(strNP09) & ")"
            
      'Modify by Morgan 2004/9/13 '加巴基斯坦 038, 應該不必轉存檔時就已經轉過, 但留著以防萬一。
      
      'Modify by Morgan 2005/1/3 一律轉西元以方便測試
      'If pa(9) = "235" Or pa(9) = "038" Or pa(9) = "301" Then
        'Modified by Lydia 2014/11/10 因為在LetterCacheVar 單純增加領證費,在轉稿時會缺少欄位,因此與StarLetter整合
        If LCvar = True Then
          strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & m_strCP09ByCheng & "',0,'年費法定期限'," & CNULL(DBDATE(strNP09)) & ",'')"
        Else
          strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
            "','年費法定期限'," & CNULL(DBDATE(strNP09)) & ")"
        End If
'      Else
'         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
'            "','年費法定期限'," & CNULL(strNP09) & ")"
'      End If
      'Modify end ---
      
      Jjj = Jjj + 1
   End If
   If CheckStr(txtCaseField(4)) <> "" Then
        'Modified by Lydia 2014/11/10 因為在LetterCacheVar 單純增加領證費,在轉稿時會缺少欄位,因此與StarLetter整合
        If LCvar = True Then
          strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & m_strCP09ByCheng & "',0,'年費本所期限'," & CNULL(strNP08) & ",'')"
        Else
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
               "','年費本所期限'," & CNULL(strNP08) & ")"
        End If
      Jjj = Jjj + 1
   End If
   
   'Added by Morgan 2018/10/30
   If m_930DueDate <> "" Then
      If LCvar = True Then
        strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
          "VALUES ('" & m_strCP09ByCheng & "',0,'商業使用聲明法定期限'," & m_930DueDate & ",'')"
      Else
          strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
             "','商業使用聲明法定期限'," & m_930DueDate & ")"
      End If
      Jjj = Jjj + 1
   End If
   'end 2018/10/30
   
   strTemp = ""
   strTemp1 = ""


   '2005/6/14 MODIFY BY SONIA
   'strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '      "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
   '      "','專用年度'," & CNULL(GetPatentYear(pa(9), pa(8))) & ")"
   If pa(9) = "011" And pa(8) = "2" And pa(10) <> "" And Val(pa(10)) < 20050331 Then '日本新型舊法專用期6年
      'Modified by Lydia 2014/11/10
      If LCvar = True Then
        strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
          "VALUES ('" & m_strCP09ByCheng & "',0,'專用年度'," & 6 & ",'')"
      Else
        strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
              "','專用年度'," & 6 & ")"
      End If
        Jjj = Jjj + 1
   End If

   'ADD BY SONIA 2014/4/30 美國設計舊法(申請日20131218以前)專用期14年, 2015/1/19 慧汶通知取消, 仍為發證日起14年
   'add by sonia 2015/11/30 2015/5/8 慧汶通知確定於2015/5/13起生效,改為15年
   If pa(9) = "101" And pa(8) = "3" And pa(10) <> "" Then
      If Val(pa(10)) < 20150513 Then
         If LCvar = True Then
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & m_strCP09ByCheng & "',0,'專用年度'," & 14 & ",'')"
         Else
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
                  "','專用年度'," & 14 & ")"
         End If
      Else
         If LCvar = True Then
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & m_strCP09ByCheng & "',0,'專用年度'," & CNULL(GetPatentYear(pa(9), pa(8))) & ",'')"
         Else
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
                  "','專用年度'," & CNULL(GetPatentYear(pa(9), pa(8))) & ")"
         End If
      End If
      Jjj = Jjj + 1
   End If
   'end 2015/11/30

   '2006/1/27 MODIFY BY SONIA 菲律賓新法新型不必繳年費
   'If strNP09 <> "" Then BolTmp = objPublicData.GetNationTax(Val(pa(8)), pa(9), strTemp, strTemp1, 年費, strExc(0))
   If strNP09 <> "" Then
      If pa(9) = "030" And pa(8) = "2" And pa(10) <> "" And Val(pa(10)) < 19980101 Then
      Else
         'edit by nickc 2007/02/02 不用 dll 了
         'BolTmp = objPublicData.GetNationTax(Val(pA(8)), pA(9), strTemp, strTemp1, 年費, strExc(0))
         bolTmp = ClsPDGetNationTax(Val(pa(8)), pa(9), strTemp, strTemp1, 年費, strExc(0))
      End If
   End If
   '2006/1/27 END
   '91.12.26 END
   If bolTmp And CheckStr(strTemp1) <> "" And CheckStr(strTemp) <> "" Then
      Set rsTmp = New ADODB.Recordset
      iStep = InStr(1, strTemp1, ",")
      If iStep = 0 Then
         iStep = 1
      Else
         iStep = Mid(strTemp1, 1, (InStr(1, strTemp1, ",")) - 1)
      End If
      strSql = "select decode(yf06,null,0,yf06)+decode(yf07,null,0,yf07) from patentyearfee where yf02='" & lblCaseField(2) & "' and yf01='" & lblCaseField(3) & "' and yf05='" & iStep & "' and yf03='" & CheckStr(pa(26)) & "' and yf04='" & CheckStr(strTemp) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If Not rsTmp.EOF And Not rsTmp.BOF Then
        'Modified by Lydia 2014/11/10
        If LCvar = True Then
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & m_strCP09ByCheng & "',0,'年費金額','" & CheckStr(rsTmp.Fields(0).Value) & "','')"
        Else
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
               "','年費金額','" & CheckStr(rsTmp.Fields(0).Value) & "')"
        End If
         Jjj = Jjj + 1
      Else
         Set rsTmp = New ADODB.Recordset
         strSql = "select decode(yf06,null,0,yf06)+decode(yf07,null,0,yf07) from patentyearfee where yf02='Y00000000' and yf01='" & lblCaseField(3) & "' and yf05='" & iStep & "' and yf03='" & CheckStr(pa(26)) & "' and yf04='" & CheckStr(strTemp) & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If Not rsTmp.EOF And Not rsTmp.BOF Then
            'Modified by Lydia 2014/11/10
            If LCvar = True Then
                strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                  "VALUES ('" & m_strCP09ByCheng & "',0,'年費金額','" & CheckStr(rsTmp.Fields(0).Value) & "','')"
            Else
                strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
                   "','年費金額','" & CheckStr(rsTmp.Fields(0).Value) & "')"
            End If
            Jjj = Jjj + 1
         Else
            'Modified by Lydia 2014/11/10
            If LCvar = True Then
                strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                  "VALUES ('" & m_strCP09ByCheng & "',0,'年費金額',0,'')"
            Else
                strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
                   "','年費金額','0')"
            End If
            Jjj = Jjj + 1
         End If
      End If
   Else
        'Modified by Lydia 2014/11/10
        If LCvar = True Then
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & m_strCP09ByCheng & "',0,'年費金額',0,'')"
        Else
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
               "','年費金額','0')"
        End If
      Jjj = Jjj + 1
   End If
   'Modify by Morgan 2005/2/17 領證費>0才存因為要控制定稿整句不印
   'If CheckStr(CNULL(txtCaseField(10))) <> "" Then
   If Val(txtCaseField(10)) > 0 Then
        'Modified by Lydia 2014/11/10
        If LCvar = True Then
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & m_strCP09ByCheng & "',0,'領證費','" & Val(txtCaseField(10)) & "','Y')" 'LCV05=Y=>報價是否顯示=智權可改
            Jjj = Jjj + 1
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
               "VALUES ('" & m_strCP09ByCheng & "',0,'領證費點數','" & Val(txtCaseField(11)) & "','')"
            Jjj = Jjj + 1
            'Modified by Lydia 2015/05/18
'            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
'               "VALUES ('" & m_strCP09ByCheng & "',0,'費用總計','" & Val(txtCaseField(10)) & "','')"
'            Jjj = Jjj + 1
'            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
'               "VALUES ('" & m_strCP09ByCheng & "',0,'點數合計','" & Val(txtCaseField(11)) & "','')"
'            Jjj = Jjj + 1
        Else
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
               "','領證費'," & CNULL(txtCaseField(10)) & ")"
            'Modified by Lydia 2015/05/18
            Jjj = Jjj + 1
        End If
      'Modified by Lydia 2015/05/18
      'Jjj = Jjj + 1
   End If
   
   If LCvar = True Then 'Added by Morgan 2019/11/7
   
      'Added by Lydia 2015/05/18 +048緬甸專利必填廣告費
      If Val(txtCaseField(18)) > 0 Then '所限=刊登廣告期限
          strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
           "VALUES ('" & m_strCP09ByCheng & "',0,'刊登廣告期限'," & Val(txtCaseField(18)) & ",'')"
          Jjj = Jjj + 1
      End If
      If Val(txtCaseField(15)) > 0 Then
          strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
             "VALUES ('" & m_strCP09ByCheng & "',0,'廣告費','" & Val(txtCaseField(15)) & "','Y')"
          Jjj = Jjj + 1
      End If
      If Val(txtCaseField(16)) > 0 Then
          strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
             "VALUES ('" & m_strCP09ByCheng & "',0,'廣告費點數','" & Val(txtCaseField(16)) & "','')"
          Jjj = Jjj + 1
      End If
      If Val(txtCaseField(10)) > 0 Or Val(txtCaseField(15)) > 0 Then
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
               "VALUES ('" & m_strCP09ByCheng & "',0,'費用總計','" & Val(txtCaseField(10)) + Val(txtCaseField(15)) & "','')"
         Jjj = Jjj + 1
      End If
      If Val(txtCaseField(11)) > 0 Or Val(txtCaseField(16)) > 0 Then
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
               "VALUES ('" & m_strCP09ByCheng & "',0,'點數總計','" & Val(txtCaseField(11)) + Val(txtCaseField(16)) & "','')"
         Jjj = Jjj + 1
      End If
      'end 2015/05/18
      
   End If 'Added by Morgan 2019/11/7
   
   'Add by Morgan 2004/12/10
   If CheckStr(CNULL(txtCaseField(14))) <> "" Then
        'Modified by Lydia 2014/11/10
        If LCvar = True Then
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & m_strCP09ByCheng & "',0,'調整期'," & CNULL(txtCaseField(14)) & ",'')"
        Else
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
               "','調整期'," & CNULL(txtCaseField(14)) & ")"
        End If
      Jjj = Jjj + 1
   End If
    '若申請國家為EPC時
    If pa(9) = "221" Then
        'Modified by Lydia 2014/11/10
        'Modified by Morgan 2023/8/14
        'strExc(0) = GetEPCNations(pa(1), pa(2), pa(3), pa(4))
        'If strExc(0) <> "" Then
        If ClsPDReadCountry(專利, pa(), strExc(0), True, False, True) = True Then
        'end 2023/8/14
            If LCvar = True Then
                strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                  "VALUES ('" & m_strCP09ByCheng & "',0,'指定國家','" & strExc(0) & "','')"
            Else
                strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
                    "','指定國家','" & strExc(0) & "')"
            End If
        End If
        Jjj = Jjj + 1
    End If
    
    'Add by Morgan 2006/8/28
    'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
    If pa(9) = "023" And pa(8) = "2" Then '俄羅斯新型
      If pa(72) <> "" Then
         strExc(0) = Right(pa(72), 2)
         If Left(strExc(0), 1) = "," Then strExc(0) = Right(strExc(0), 1)
         If strExc(0) = "1" Then
            strExc(1) = "一"
         ElseIf strExc(0) = "2" Then
            strExc(1) = "一、二"
         Else
            strExc(1) = "一～" & PUB_ChgNumber2Chinese(strExc(0))
         End If
         'Modified by Lydia 2014/11/10
        If LCvar = True Then
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & m_strCP09ByCheng & "',0,'已繳年度','" & strExc(1) & "','')"
        Else
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & "','已繳年度','" & strExc(1) & "')"
        End If
         Jjj = Jjj + 1
      End If
    End If
    'edit by nickc 2007/02/05 不用 dll 了
    'If Not objLawDll.ExecSQL(Jjj - 1, strTxt) Then
    'Add by Morgan 2007/7/20
    If pa(3) <> "0" Then
        'Modified by Lydia 2014/11/10
        If LCvar = True Then
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & m_strCP09ByCheng & "',0,'母案','母案','')"
        Else
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
                "','母案','母案')"
        End If
        Jjj = Jjj + 1
    End If
    
    'Add by Morgan 2008/5/15
    If m_strEpcRegDueDay <> "" Then
        'Modified by Lydia 2014/11/10
        If LCvar = True Then
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
              "VALUES ('" & m_strCP09ByCheng & "',0,'註冊費期限','" & m_strEpcRegDueDay & "','')"
        Else
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_strCP09ByCheng & "','" & ET03 & "','" & strUserNum & _
               "','註冊費期限','" & m_strEpcRegDueDay & "')"
        End If
      Jjj = Jjj + 1
    End If
    'end 2008/5/15
    
    'end 2007/7/20
    If Not ClsLawExecSQL(Jjj - 1, strTxt) Then
        MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
    End If
End Sub

Private Sub cmdok_Click(Index As Integer)
 Dim i As Integer, strTmp As String
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         
         'Removed by Morgan 2019/11/7 與 TxtValidate 檢查重複
         'For i = 0 To 13
         '   If txtCaseField(i).Enabled Then
         '      If CheckKeyIn(i) <> 1 Then
         '         txtCaseField(i).SetFocus
         '         txtCaseField_GotFocus (i)
         '         Exit For
         '      End If
         '   End If
         'Next
         'If i <> 14 Then Screen.MousePointer = vbDefault: Exit Sub
         'end 2019/11/7
         
         '重新檢查欄位有效性
         If TxtValidate = True Then
            'Add by Amy 2018/05/25  判斷有未收款彈訊息
            If Pub_B911NotPay(pa(1), pa(2), pa(3), pa(4)) = True Then
                MsgBox "此案有未收款！", vbExclamation
            End If
            'end 2018/05/25
            'Add by Morgan 2006/7/5
            '期限更新時提醒
            If m_NP09_Old <> "" And m_NP09_Old <> txtCaseField(4).Text Then
               If MsgBox("下次繳費日將更新【 " & m_NP09_Old & " --> " & txtCaseField(4).Text & " 】，是否要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            'end 2006/7/5

            'Add by Morgan 2005/5/20
            '非台灣 詢問是否計算結餘
            If Trim(txtCaseField(2)) <> Empty Then
               '2011/11/8 modify by sonia TF子案不可結餘故加傳本所案號
               'Pub_EndModCashMsg pa(9)
               Pub_EndModCashMsg pa(9), pa(1), pa(2), pa(3), pa(4)
            End If
            
            'Add by Morgan 2007/5/10 若來函有期限但已閉卷
            m_blnCancelClosed = False
            '2012/2/23 modify by sonia 加控制下次繳費日未過期條件 CFP-020624
            'If pa(57) = "Y" And txtCaseField(4) <> "" Then
            If pa(57) = "Y" And txtCaseField(4) <> "" And Val(Me.txtCaseField(4).Text) >= Val(strSrvDate(2)) Then
               If MsgBox("本案目前為閉卷狀態，為管制期限將於存檔時取消閉卷，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
                  Screen.MousePointer = vbDefault: Exit Sub
               End If
               m_blnCancelClosed = True
            End If
            'end 2007/5/10
            
            If SaveData Then
               'Add by Morgan 2007/4/30 通知香港案承辦人
               If m_HKCP14 <> "" Then
                  Call PUB_SendMail(strUserNum, m_HKCP14, m_HKCP09, "香港的關聯案(" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & ")已發證，香港案(" & m_HKPA01 & "-" & m_HKPA02 & "-" & m_HKPA03 & "-" & m_HKPA04 & ")的[標準專利批准記錄請求]可以處理！", "如旨")
               End If
               'end 2007/4/30
               
               If txtCaseField(12) <> "N" Then
                  'Modify by Morgan 2004/9/10
                  '無證書(無專用期)定稿共用 CFP-05-000-01
                  If txtCaseField(5) = "N" Then
                     strTmp = "01"
                     
                  'Add by Morgan 2009/9/22
                  'EPC子案發證也要出定稿 CFP-018284-0-20 -- 禧佩
                  ElseIf pa(4) <> "00" Then
                     strTmp = "40"
                  
'Modified by Morgan 2020/10/15 定稿合併(加例外欄位控制), +025伊朗
'                  'Added by Morgan 2020/8/28
'                  '通用定稿(專利證書正本,無維持費用)
'                  '017印尼設計、030菲律賓新型
'                  ElseIf (pa(9) = "017" And pa(8) = "3") Or (pa(9) = "030" And pa(8) = "2") Then
'                     strTmp = "51"
'
'                  '通用定稿(專利證書正本)
'                  '017印尼非設計,018馬來西亞,023俄羅斯,030菲律賓非新型,301南非
'                  ElseIf (pa(9) = "017" And pa(8) <> "3") Or pa(9) = "018" Or pa(9) = "023" Or (pa(9) = "030" And pa(8) <> "2") Or pa(9) = "301" Then
'                     strTmp = "24"
'
'                  '通用定稿(專利證書電子檔印出)
'                  '014新加坡、040印度、126智利
'                  ElseIf pa(9) = "014" Or pa(9) = "040" Or pa(9) = "126" Then
'                     strTmp = "25"
'
'                  'end 2020/8/28
                  'Modified by Morgan 2020/10/27 +204 義大利發明，其他專利種類尚未確認先不出定稿
                  'Modified by Morgan 2020/10/28 +118 阿根廷
                  'Modified by Morgan 2020/12/11 +204 義大利新型設計
                  'Modified by Morgan 2021/1/28 +117 巴西發明
                  'Modified by Morgan 2021/3/2 +016紐西蘭??、027以色列、031斯里蘭卡、038巴基斯坦、074歐亞專利聯盟
                  '、104墨西哥、116秘魯、117巴西、222波蘭、235土耳其、303埃及
                  '、304非洲聯盟(OAPI)、338非洲區域專利組織(ARIPO)
                  'Modified by Morgan 2021/9/28 +021沙烏地阿拉伯
                  'Modified by Morgan 2021/11/26 +046 柬埔寨 --慧汶
                  'Modified by Morgan 2021/12/9 +213 葡萄牙 --慧汶
                  ElseIf InStr("014,016,017,018,021,023,025,027,030,031,038,040,046,074,104,116,117,118,126,204,213,222,235,301,303,304,338", pa(9)) > 0 Then
                     strTmp = "24"
'end 2020/10/15
                  'Added by Morgan 2020/9/23
                  '048緬甸
                  ElseIf pa(9) = "048" Then
                     strTmp = "C5"
                  'end 2020/9/23
                  '有證書
                  Else
                     'strTmp = "00"  '2008/10/8 CANCEL BY SONIA 無此處理狀況
                     Select Case pa(8)
                        Case "1" '發明
                           Select Case pa(9)
                              Case "011" '日本
                                 'Modified by Morgan 2024/6/7 除 Y20332000,Y51555000 兩家代理人案件，其餘改電子專利證書定稿
                                 'strTmp = "20"
                                 strExc(1) = GetCaseProData(m_strLD18, "cp44")
                                 If strExc(1) = "Y20332000" Or strExc(1) = "Y51555000" Then
                                    strTmp = "20"
                                 Else
                                    strTmp = "13"
                                 End If
                                 'end 2024/6/7
                                 'Added by Morgan 2015/8/25
                                 '植物新品種
                                 'Modified by Morgan 2025/8/7
                                 'If PUB_ChkCPExist(cp, "120") = True Then
                                 If m_bolNewPlant Then
                                 'end 2025/8/7
                                    strTmp = "C7"
                                 End If
                                 'end 2015/8/25
                              Case "012" '韓國
                                 strTmp = "06"
                              Case "013" '香港
                                 strTmp = "21"
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "014" '新加坡
                              '   strTmp = "03"
                              Case "015" '澳洲
                                 'Added by Morgan 2025/8/7
                                 If m_bolNewPlant Then
                                    strTmp = "03"
                                 Else
                                 'end 2025/8/7
                                    strTmp = "22"
                                 End If
                                 
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "016" '紐西蘭
                                 'strTmp = "23"
                                 
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "017" '印尼
                              '   strTmp = "07"
                              'Case "018" '馬來西亞
                              '   strTmp = "24"
                              
                              Case "019" '泰國
                                 strTmp = "05"
                              
                              'Removed by Morgan 2020/10/15 移到上面改通用定稿24
                              'Case "025" '伊朗 'Add by Morgan 2006/6/7
                              '   strTmp = "35"
                              
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "027" '以色列 Add by Morgan 2010/11/9
                              '   strTmp = "68"
                              
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "030" '菲律賓
                              '   'Modify by Morgan 2009/2/26 區分新舊法
                              '   If pa(10) <> "" And Val(pa(10)) < 19980101 Then
                              '      strTmp = "09"
                              '   Else
                              '      strTmp = "39"
                              '   End If
                              
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "031" '斯里蘭卡
                              '   strTmp = "04"
                              'Case "038" '巴基斯坦 'Add by Morgan 2004/9/13
                              '   strTmp = "32"
                              
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿25
                              'Case "040" '印度
                              '   strTmp = "25"
                              Case "042" '越南
                                 strTmp = "16"
                              Case "101" '美國
                                 strTmp = "02"
                                 'Add by Morgan 2004/12/10 有調整期
                                 'Removed by Morgan 2023/4/25 併入 02
                                 'If Val(txtCaseField(14).Text) > 0 Then
                                 '   strTmp = "21"
                                 'End If
                                 'end 2023/4/25
                              Case "102" '加拿大
                                 strTmp = "11"
                              
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "104" '墨西哥
                              '   strTmp = "10"
                              
                              'Removed by Morgan 2021/1/28 移到上面改通用定稿24
                              'Case "117" '巴西
                              '   strTmp = "12"
                              
                              'Removed by Morgan 2020/10/28 移到上面改通用定稿24
                              'Case "118" '阿根廷
                              '   strTmp = "13"
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿25
                              'Case "126" '智利 'Added by Morgan 2017/3/1
                              '   strTmp = "C9"
                              Case "201" '英國
                                 strTmp = "18"
                                 'Add By Cheng 2002/07/31
                                 If Me.txtCaseField(14).Visible Then
                                    If Me.txtCaseField(14).Text <> "" Then
                                       strTmp = "19"   '含年費
                                    Else
                                       strTmp = "18"   '無年費
                                    End If
                                 End If
                              Case "203" '法國
                                 strTmp = "08"
                              'Removed by Morgan 2020/10/28 移到上面改通用定稿24
                              'Case "204" '義大利
                              '   strTmp = "26"
                              Case "205" '瑞士
                                 strTmp = "14"
                              Case "206" '奧地利
                                 strTmp = "15"
                              Case "207" '荷蘭
                                 strTmp = "17"
                              Case "209" '比利時
                                 strTmp = "27"
                              'Add by Morgan 2007/1/29
                              Case "211" '西班牙
                                 strTmp = "36"
                              Case "214" '瑞典
                                 strTmp = "28"
                              Case "215" '挪威
                                 strTmp = "C3"
                              Case "221" 'ＥＰＣ
                                 strTmp = "29"
                                 'Modify by Morgan 2008/5/15 +未繳註冊費定稿
                                 If m_bolEpcRegNotPaid = True Then
                                    strTmp = "37"
                                 End If
                              'Add by Morgan 2008/8/25
                              Case "223" '捷克
                                 strTmp = "38"
                              'Add by Morgan 2010/11/15
                              Case "228" '羅馬尼亞
                                 strTmp = "38"
                              Case "231" '德國
                                 strTmp = "30"
                              'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "023" '俄羅斯
                              '   strTmp = "31"
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "301" '南非 'Add by Morgan 2004/11/8
                              '   strTmp = "33"
                                                            
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "303" '埃及 '2005/12/15 ADD BY SONIA
                              '   strTmp = "34"
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "304" '非洲聯盟 'Added by Morgan 2017/2/20
                              '   strTmp = "C8"
                                 
                              'Added by Lydia 2015/05/18
                              'Removed by Morgan 2020/9/23 改定稿內容並移到上面(不分專利種類)
                              'Case "048" '緬甸
                              '   strTmp = "C5"
                              'end 2020/9/23
                                                            
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "074" '歐亞專利聯盟 'Added by Morgan 2018/7/31
                              '   strTmp = "A0"
                              'Case "235" '土耳其 'Added by Morgan 2018/10/30
                              '   strTmp = "A1"
                              
                           End Select
                           
                        Case "2" '新型
                           Select Case pa(9)
                              Case "011" '日本
                                 'Modified by Morgan 2024/6/7 除 Y20332000,Y51555000 兩家代理人案件，其餘改電子專利證書定稿
                                 'strTmp = "47"
                                 strExc(1) = GetCaseProData(m_strLD18, "cp44")
                                 If strExc(1) = "Y20332000" Or strExc(1) = "Y51555000" Then
                                    strTmp = "47"
                                 Else
                                    strTmp = "25"
                                 End If
                                 'end 2024/6/7
                              Case "012" '韓國
                                 strTmp = "65"  '新法  2006/10/1(含)以後提申者
                                 
                                 'Removed by Morgan 2021/6/28 專用期已過，刪除
                                 'If pa(10) <> "" And Val(pa(10)) < 19990701 Then
                                 '   strTmp = "45"   '舊法
                                 'ElseIf pa(10) <> "" And Val(pa(10)) < 20061001 Then
                                 '   strTmp = "46"   '舊法
                                 'End If
                                 'end 2021/6/28
                              Case "015" '澳洲
                                 strTmp = "48"
                                 
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "017" '印尼
                              '   strTmp = "57"  '新法  2001/8/1(含)以後提申者
                              '   If pa(10) <> "" And Val(pa(10)) < 20010801 Then
                              '      strTmp = "41"   '舊法
                              '   End If
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "018" '馬來西亞
                              '   strTmp = "49"
                              Case "019" '泰國
                                 strTmp = "50"
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿51
                              'Case "030" '菲律賓
                              '   strTmp = "51"
                              Case "042" '越南
                                 strTmp = "44"
                              
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "104" '墨西哥
                              '   strTmp = "42"
                              
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "117" '巴西 Added by Morgan 2012/4/20
                              '   'strTmp = "69" 'Removed by Morgan 2021/1/28 未確認先不出
                                 
                              Case "203" '法國
                                 strTmp = "52"
                              'Removed by Morgan 2020/10/28 移到上面改通用定稿24
                              'Case "204" '義大利
                              '   strTmp = "53"
                              Case "206" '奧地利 'Add by Morgan 2009/1/15
                                 strTmp = "67"
                              Case "207" '荷蘭
                                 strTmp = "54"
                              Case "209" '比利時 'Add by Morgan 2005/4/21
                                 strTmp = "58"
                              Case "211" '西班牙
                                 strTmp = "55"
                              Case "212" '希臘 'Add by Morgan 2008/1/22
                                 strTmp = "63"
                              'Add by Morgan 2009/4/17
                              'Mark by Lydia 2025/03/14 移到上面改通用定稿24--- from 'Modified by Morgan 2021/12/9 +213 葡萄牙 --慧汶
                              'Case "213" '葡萄牙
                              '   strTmp = "00"
                              'end 2025/03/14
                              Case "217" '芬蘭
                                 strTmp = "43"
                              Case "219" '匈牙利 'Add by Morgan 2008/8/26
                                 strTmp = "64"
                              
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "222" '波蘭 'Add by Morgan 2007/8/20
                              '   strTmp = "62"
                              
                              Case "223" '捷克 'Add by Morgan 2007/8/10
                                 strTmp = "61"
                              Case "226" '保加利亞 'Add by Morgan 2008/10/8
                                 strTmp = "66"
                              Case "231" '德國
                                 strTmp = "56"
                              'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "023" '俄羅斯 'Add by Morgan 2006/8/28
                              '   strTmp = "59"
                              
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "235" '土耳其 'Add by Morgan 2006/11/1
                              '   strTmp = "60"
                              
                           End Select
                           
                        Case "3" '設計
                           Select Case pa(9)
                              Case "011" '日本
                                 'Modified by Morgan 2024/6/7 除 Y20332000,Y51555000 兩家代理人案件，其餘改電子專利證書定稿
                                 'strTmp = "76"
                                 strExc(1) = GetCaseProData(m_strLD18, "cp44")
                                 If strExc(1) = "Y20332000" Or strExc(1) = "Y51555000" Then
                                    strTmp = "76"
                                 Else
                                    strTmp = "26"
                                 End If
                                 'end 2024/6/7
                              Case "012" '韓國
                                 '2015/4/16 MODIFY BY SONIA 舊法為發證日起15年
                                 'strTmp = "75"
                                 'cancel by sonia 2025/5/29 不會再有舊法未發證案件
                                 'If Val(pa(10)) < 20140701 Then
                                 '   strTmp = "C4"
                                 'Else
                                    strTmp = "75"
                                 'End If
                                 '2015/4/16 END
                              '2005/12/27 ADD BY SONIA
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿25
                              'Case "014" '新加坡
                              '   strTmp = "90"
                              '2005/12/27 END
                              Case "015" '澳洲
                                 'Modify by Morgan 2005/6/14 加新法定稿
                                 If Val(pa(10)) < 20040617 Then
                                    'strTmp = "84" 'Removed by Morgan 2013/10/18 定稿已刪除，因已無案件適用。
                                 Else
                                    strTmp = "87"
                                 End If
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿51
                              'Case "017" '印尼
                              '   strTmp = "C1"
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "018" '馬來西亞
                              '   strTmp = "81"
                              Case "019" '泰國
                                 strTmp = "99"
                              
                              'Removed by Morgan 2021/9/28 移到上面改通用定稿24
                              'Case "021" '沙烏地阿拉伯 Added by Morgan 2011/12/1
                               '  strTmp = "C2"
                                 
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "027" '以色列 'Add by Morgan 2008/12/12
                              '   strTmp = "96"
                              
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "030" '菲律賓 'Add by Morgan 2009/3/6
                              '   strTmp = "97"
                              
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "038" '巴基斯坦 'Add by Morgan 2008/9/8
                              '   strTmp = "95"
                              
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿25
                              'Case "040" '印度
                              '   strTmp = "82"
                              Case "042" '越南
                                 strTmp = "71"
                              Case "101" '美國
                                 strTmp = "77"
                              Case "102" '加拿大
                                 strTmp = "72"
                              
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "104" '墨西哥 'Add by Morgan 2006/9/7
                              '   strTmp = "91"
                              'Case "116" '祕魯
                              '   strTmp = "98"
                                                            
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "117" '巴西 'Add by Morgan 2007/3/20
                              '   'strTmp = "92" 'Removed by Morgan 2021/1/28 未確認先不出
                              
                              '2005/12/27 ADD BY SONIA
                              'Removed by Morgan 2020/10/28 移到上面改通用定稿24
                              'Case "118" '阿根廷
                              '   strTmp = "89"
                              'Add by Morgan 2010/7/19
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿25
                              'Case "126" '智利
                              '   strTmp = "70"
                              '2005/12/27 END
                              Case "201" '英國
                                 strTmp = "78"
                              Case "203" '法國 'Add by Morgan 2005/8/2
                                 strTmp = "88"
                              'Removed by Morgan 2020/10/28 移到上面改通用定稿24
                              'Case "204" '義大利
                              '   strTmp = "79"
                              Case "210" '荷比盧
                                 strTmp = "74"
                              Case "211" '西班牙
                                 strTmp = "83"
                              Case "214" '瑞典   '2007/7/9 ADD BY SONIA
                                 strTmp = "93"
                              Case "231" '德國
                                 strTmp = "80"
                              'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "023" '俄羅斯 'Add by Morgan 2008/2/20
                              '   strTmp = "94"
                              '   If Val(pa(10)) >= 20141001 Then
                              '      strTmp = "C6"
                              '   End If
                              
                              'Removed by Morgan 2021/3/2 移到上面改通用定稿24
                              'Case "235" '土耳其 'Add by Morgan 2004/3/31
                              '   strTmp = "85"
                              
                              Case "239" '歐盟 'Add by Morgan 2005/1/3
                                 strTmp = "86"
                              'Removed by Morgan 2020/8/28 移到上面改通用定稿24
                              'Case "301" '南非
                              '   strTmp = "73"
                           End Select
                     End Select
                  End If
                  
                   'Added by Morgan 2020/5/25
                   '下列專利權期間尚未確定的國家先取消定稿
                   strExc(1) = Pub_GetSpecMan("CFP證書專用期未確定國家", True)
                   If InStr(strExc(1), pa(9)) > 0 Then strTmp = ""
                   'end 2020/5/25
                   
                  'Modified by Morgan 2015/2/9
                  '要報價但沒有定稿時提醒
                  'If strTmp <> "" Then
                  'end 2015/2/9
                  
                        'Add by Lydia 2014/10/29 自動發證國家目前寄證書程序發文後，業務常要求改費用同舊案報價。修改系統在告准輸入(證書號輸入)時，會自動提供半年內之領證費報價(類一般來函輸入領證費)。
                        'copy copy frm05010401_3
                        'Add by Morgan 2008/5/7 新增領證報價通知
                        'Modified by Morgan 2014/11/17 +判斷有證書
                        'Modified by Moragn 2015/1/14 +判斷有領證費才出報價定稿
                        If m_bolAutoIssue = True And txtCaseField(5) = "Y" And Val(txtCaseField(10)) > 0 Then
                              'Modified by Morgan 2015/2/9
                              '要報價但沒有定稿時提醒
                              If strTmp = "" Then
                                 MsgBox "本案要報價但沒有系統的定稿，請注意！", vbExclamation
                                 SetNoReceipt 'Added by Morgan 2020/12/16 設為不請款以避免財處先開收據
                              Else
                              'end 2015/2/9
                              
                                 '若是自動發證國家，因為無領證程序所以無下一程序NP22
                                  PUB_AddLetterCache m_strCP09ByCheng, "0", m_strCP09ByCheng, "05", strTmp, , m_strLD18
                                ' StartLetter1 strTmp, m_strCP09ByCheng, "0" 'strTmp=定稿號碼
                                  StartLetter "05", strTmp, True 'Modified by Lydia ,LCvar="Y" 將新增資料在LetterCacheVar
                              'end 'Add by Lydia 2014/10/29 自動發證
                              
                        'Modified by Morgan 2015/2/9
                        'else
                              End If
                        ElseIf strTmp <> "" Then
                        'end 2015/2/9
                        
                           'Added by Morgan 2022/5/19 寶齡富錦 Y55435 案件
                           If ChangeCustomerS(pa(75)) = "Y55435" Then
                              strTmp = "12"
                           End If
                           'end 2022/5/19
                              
                            StartLetter "05", strTmp
                            NowPrint m_strCP09ByCheng, "05", strTmp, IIf(Me.txtCaseField(13).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                            
                           'Added by Morgan 2018/7/16 CFP電子化
                           If m_bolAddLP And txtCaseField(13).Text = "Y" Then
                              If m_HKNP22 <> "" And m_HKCP10 = "111" Then
                                 MsgBox "為配合轉PDF檔至卷宗區，香港案[" & m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04 & "]的標準專利批准紀錄請求通知函請到定稿維護修改內容!!", vbInformation, "CFP電子化"
                                 txtCaseField(13).Text = ""
                              End If
                              frm1105_1.m_RecNo = m_strLD18
                              frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & m_strCP10 & ".CUS.PDF"
                              frm1105_1.Show
                           End If
                           'end 2018/7/16
               
                        End If
                        
                  'End If 'Removed by Morgan 2015/2/9
                  
                  'Add by Morgan 2007/4/30 若有未收文的香港案的"標準專利批准紀錄請求"時要出定稿
                  If m_HKNP22 <> "" And m_HKCP10 = "111" Then
                     RunHKInform IIf(Me.txtCaseField(13).Text = "Y", True, False)
                  End If
                  'end 2007/4/30
               End If
               'Add by Morgan 2007/3/22
               'Modify by Morgan 2007/6/6 EU(329)發證日=申請日，改以公告日管控;另需考慮也有可能輸公告日或核准日
               'If txtCaseField(1).Text <> "" And DBDATE(txtCaseField(1)) <> DBDATE(pa(21)) Then
               '   PUB_SameCaseCheck1 cp(), 2, DBDATE(txtCaseField(1).Text)
               strExc(1) = "": strExc(2) = "": strExc(3) = ""
               If pa(9) <> "239" And txtCaseField(1) <> "" And DBDATE(txtCaseField(1)) <> DBDATE(m_oPA21) Then
                  strExc(1) = DBDATE(txtCaseField(1))
                  strExc(2) = 2
               End If
               If txtCaseField(6) <> "" And DBDATE(txtCaseField(6)) <> DBDATE(m_oPA14) Then
                  strExc(3) = DBDATE(txtCaseField(6))
                  If strExc(1) = "" Or strExc(3) < strExc(1) Then
                     strExc(1) = strExc(3)
                     strExc(2) = 4
                  End If
               End If
               If txtCaseField(8) <> "" And DBDATE(txtCaseField(8)) <> DBDATE(m_oPA20) Then
                  strExc(3) = DBDATE(txtCaseField(8))
                  If strExc(1) = "" Or strExc(3) < strExc(1) Then
                     strExc(1) = strExc(3)
                     strExc(2) = 1
                  End If
               End If
               If strExc(1) <> "" Then
                  PUB_SameCaseCheck1 cp(), Val(strExc(2)), strExc(1)
               End If
               'end 2007/6/6
               'end 2007/3/22
               
               bolLeave = True
               'Add By Sindy 2016/10/7
               If Me.m_strIR01 <> "" Then
                  intLeaveKind = 2
                  'Unload frm05010402_1
                  Unload Me
                  'Modify By Sindy 2022/5/20
                  'frm04010519.GoNext
                  Forms(0).Tmpfrm04010519.GoNext
                  Set Forms(0).Tmpfrm04010519 = Nothing
                  '2022/5/20 END
               Else
               '2016/10/7 END
                  intLeaveKind = 0
                  Unload Me
               End If
            '911202 nick
            Else
                MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
            End If
         End If
         Screen.MousePointer = vbDefault
      Case 1, 2
         If Index = 2 Then
            intLeaveKind = 2
         Else
            intLeaveKind = 1
         End If
         Unload Me
   End Select
End Sub

Private Function SaveData() As Boolean

Dim strSql As String
Dim strNP22 As String
Dim strTxt(1 To 30) As String, iStep As Integer
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim strDateS(0 To 5) As String
Dim strRefCP14 As String
Dim strTemp As String, strTemp1 As String  '2010/7/16 ADD BY SONIA
'edit by nickc 2007/02/02
'Dim strDataTemp(1 To T_CP) As String
Dim strDataTemp() As String
ReDim strDataTemp(1 To TF_CP) As String
Dim m_bolEndModCash As Boolean             '2011/8/16 ADD BY SONIA
Dim iNo As Integer 'Added by Morgan 2011/11/3
Dim Caseno235(1 To 4) As String            'add by sonia 2018/10/12 EPC土耳其子案案號
 
 On Error GoTo CheckingErr

   cnnConnection.BeginTrans
            
   pa(22) = txtCaseField(0)
   pa(21) = txtCaseField(1)
   pa(24) = txtCaseField(2)
   pa(25) = txtCaseField(3)
   'Modify by Morgan 2004/12/10 若有調整其則存調整後的專用期止日
   If lblCaseField(11).Caption <> "" Then
      pa(25) = lblCaseField(11).Caption
   End If
   
   pa(20) = txtCaseField(8)
   pa(14) = txtCaseField(6)
   pa(15) = txtCaseField(7)
   pa(17) = "Y"
   pa(16) = "1"
   cp(14) = strUserNum
   strTxt(1) = GetPASQL(pa)
   
   cnnConnection.Execute strTxt(1)
   
   iStep = 2
   If bolUpdate = False Then
      
      strDataTemp(1) = pa(1)
      strDataTemp(2) = pa(2)
      strDataTemp(3) = pa(3)
      strDataTemp(4) = pa(4)
      strDataTemp(5) = strSrvDate(1)
      strDataTemp(9) = 主管機關來函
      If pa(24) = "" Then
         strDataTemp(10) = 通知證書號數
      Else
         strDataTemp(10) = 專利證書
      End If
      strDataTemp(12) = stCP12
      strDataTemp(13) = stCP13
      strDataTemp(14) = cp(14)
      strDataTemp(16) = txtCaseField(10)
      strDataTemp(18) = txtCaseField(11)
      strDataTemp(17) = Val(txtCaseField(10)) - (Val(txtCaseField(11)) * 1000)
      If txtCaseField(10) = "" Then
         strTemp = "N"
         strDataTemp(17) = ""
      Else
         strTemp = ""
      End If
      strDataTemp(20) = strTemp
      strDataTemp(26) = "N"
      
      'Add by Lydia 2014/11/6 自動發證國家之發文日(CP27),保留到業務確認
      'Modified by Morgan 2014/11/17 +判斷有證書
      'Modified by Morgan 2019/2/13 +領證費判斷 Ex:CFP-30806 (非報價定稿)
      If m_bolAutoIssue = True And txtCaseField(5) = "Y" And Val(txtCaseField(10)) > 0 Then
         strDataTemp(27) = ""
         strDataTemp(144) = strSrvDate(2) & " 領證費" & txtCaseField(10) & "(" & txtCaseField(11) & ");" 'Added by Morgan 2020/12/16 沒有報價定稿時可查詢原輸入金額
      Else
         strDataTemp(27) = strSrvDate(1)
      End If
      'end 2014/11/6
      
      strDataTemp(32) = strTemp
      '2008/8/26 modify by sonia 櫃台收文日改存 cp119
      '2008/10/24 MODIFY BY SONIA CP64仍存
      strDataTemp(64) = "櫃台收文日：" & lblCaseField(9).Caption
      strDataTemp(119) = ChangeTStringToWString(lblCaseField(9).Caption)
      '2008/8/26 end
      
      strTxt(iStep) = GetCPSQL(strDataTemp(), False)
      
      cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
      
      'Add by Morgan 2004/11/30 抓最新的AB類發文代理人更新
      Pub_UpdateFromMaxCP27 pa(1), pa(2), pa(3), pa(4)

      m_strCP09ByCheng = strDataTemp(9)
      
   Else
      cp(5) = strSrvDate(1)
      cp(16) = txtCaseField(10)
      cp(18) = txtCaseField(11)
      cp(17) = Val(txtCaseField(10)) - (Val(txtCaseField(11)) * 1000)
      cp(26) = "N"
      'Add by Lydia 2014/11/6 自動發證國家之發文日(CP27),保留到業務確認
      'Modified by Morgan 2014/11/17 +判斷有證書
      'Modified by Moragn 2015/11/12 +判斷有領證費才出報價定稿
      If m_bolAutoIssue = True And txtCaseField(5) = "Y" And Val(txtCaseField(10)) > 0 Then
         'Modified by Morgan 2018/7/17
         'strDataTemp(27) = ""
         cp(27) = ""
         'end 2018/7/17
         cp(144) = strSrvDate(2) & " 領證費" & txtCaseField(10) & "(" & txtCaseField(11) & ");" & cp(144)  'Added by Morgan 2020/12/16 沒有報價定稿時可查詢原輸入金額
      Else
         cp(27) = strSrvDate(1)
      End If
      
      '2008/8/26 modify by sonia 櫃台收文日改存 cp119
      '2008/10/24 MODIFY BY SONIA CP64仍存
      If cp(64) = "" Then
         cp(64) = "櫃台收文日：" & lblCaseField(9).Caption
      Else
         cp(64) = cp(64) & ",櫃台收文日：" & lblCaseField(9).Caption
      End If
      cp(119) = ChangeTStringToWString(lblCaseField(9).Caption)
      '2008/8/26 end
      
      'Add by Morgan 2004/7/28
      '若有費用CP20=CP32=NULL
      If Val(txtCaseField(10)) > 0 Then
         cp(20) = "": cp(32) = ""
      Else
         cp(20) = "N": cp(32) = "N"
      End If
      
      strTxt(iStep) = GetCPSQL(cp())
      
      cnnConnection.Execute strTxt(iStep)
      
      iStep = iStep + 1
      m_strCP09ByCheng = cp(9)
      strDataTemp(9) = cp(9)
   End If
      
   '2012/2/23 modify by sonia 加控制下次繳費日未過期條件 CFP-020624
   'If txtCaseField(4) <> "" Then
   If txtCaseField(4) <> "" And Val(Me.txtCaseField(4).Text) >= Val(strSrvDate(2)) Then
      strNP09 = TransDate(txtCaseField(4), 2)
      Dim strDate(0 To 3) As String
      strDate(1) = cp(1)
      strDate(2) = pa(9)
      strDate(3) = strNP09
      GetCtrlDT strDate
      strNP08 = strDate(0)
      If Not IsEmptyText(m_NP22) Then
         strTxt(iStep) = "UPDATE NEXTPROGRESS SET NP08 = " & CNULL(PUB_GetWorkDay1(strNP08, True)) & ",NP09 = " & CNULL(strNP09) & _
            " WHERE NP22 = " & m_NP22 & " and np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "'"
         cnnConnection.Execute strTxt(iStep)
         
         iStep = iStep + 1
      Else
         strNP22 = GetNextProgressNo()
         If IsEmptyText(m_CaseType) Then
            m_CaseType = 年費
         End If
         
         '2005/6/21 MODIFY BY SONIA 先檢查是否已收文
         '搜尋案件性質介於"605"-"607"且未發文者
         StrSQLa = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND  (CP10>='605' AND CP10<='607') AND CP27 IS NULL AND CP57 IS NULL ORDER BY CP09 DESC "
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            strTxt(iStep) = "UPDATE CASEPROGRESS SET CP06=" & CNULL(PUB_GetWorkDay1(strNP08, True)) & " ,CP07=" & CNULL(strNP09) & " WHERE CP09='" & rsA("CP09") & "'"
            cnnConnection.Execute strTxt(iStep)
         Else
            strTxt(iStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & strDataTemp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & m_CaseType & "," & _
               CNULL(PUB_GetWorkDay1(strNP08, True)) & "," & CNULL(strNP09) & ",'" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "'," & strNP22 & ") "
               cnnConnection.Execute strTxt(iStep)
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         iStep = iStep + 1
         '2005/6/21 END
      End If
      
      'Add by Morgan 2006/8/21 馬來西亞的新型
      'Modify by Morgan 2007/7/31 加俄羅斯新型也要掛延展費期限,2008/10/15因是准後繳年費所以從代理人案件提申轉過來
      'If pa(9) = "018" And pa(8) = "2" Then
      'Modified by Morgan 2015/5/28 2014/10/01之前提出申請的俄羅斯新型案,專用期間為自申請日起算10年，可延展1次(+3年),新法不可延展
      'If (pa(9) = "018" Or pa(9) = "233") And pa(8) = "2" Then
      'Modified by Morgan 2015/10/15 俄羅斯2015/1/1修法所有未發證未延展新型案專用期都為10年不得延展
      'If (pa(9) = "018" Or (pa(9) = "233" And Val(pa(10)) < 20141001)) And pa(8) = "2" Then
      If pa(9) = "018" And pa(8) = "2" Then
      'end 2015/10/15
         strDate(1) = cp(1)
         strDate(2) = pa(9)
         '若國家檔設年費則補掛延展費,否則補掛延展費
         If m_CaseType = 年費 Then
            '延展費
            strExc(9) = "607"
            Select Case pa(9)
            Case "018"
               'Modify by Morgan 2007/7/18 新型不管何時提申均適用上述新法(2003年8月14日實施)
               ''舊法(發證日<2001/8/1)自發證日起5年，期滿可延展2次，每次5年；
               'If Val(DBDATE(pa(21))) < 20010801 Then
               '   strDate(3) = CompDate(0, 5, pa(21))
               ''新法自申請日起10年，期滿可延展2次，每次5年；
               'Else
               '   strDate(3) = CompDate(0, 10, pa(10))
               'End If
               strDate(3) = CompDate(0, 10, pa(10))
               'end 2007/7/18
'Removed by Morgan 2015/10/15 俄羅斯2015/1/1修法所有未發證未延展新型案專用期都為10年
'            '2008/10/15因是准後繳年費所以從代理人案件提申轉過來,原為5年2008/1/1起修法改10年
'            Case "233"
'               strDate(3) = CompDate(0, 10, pa(10))
'end 2015/10/15

            End Select
            '2008/10/15 end
'2008/10/15 cancel by sonia 因國家檔設定年費故取消
'         Else
'            '年費：自第3年起逐年繳交(新舊法都一樣)
'            strExc(9) = "605"
'            strDate(3) = CompDate(0, 2, pa(21))
'2008/10/15 end
         End If
         GetCtrlDT strDate
         strExc(0) = "select cp09 from CASEPROGRESS where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' AND CP10='" & strExc(9) & "' AND CP27 IS NULL AND CP57 IS NULL"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "UPDATE CASEPROGRESS SET CP06=" & PUB_GetWorkDay1(strDate(0), True) & " ,CP07=" & strDate(3) & " WHERE CP09='" & RsTemp("CP09") & "'"
            cnnConnection.Execute strSql
         Else
            strExc(0) = "select NP01,NP07,NP22 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06 is null and np07='" & strExc(9) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strSql = "UPDATE NEXTPROGRESS SET NP08 = " & PUB_GetWorkDay1(strDate(0), True) & ",NP09 = " & strDate(3) & _
                  " WHERE NP01='" & RsTemp("NP01") & "' AND NP07='" & RsTemp("NP07") & "' AND NP22=" & RsTemp("NP22")
               cnnConnection.Execute strSql
            Else
               strNP22 = GetNextProgressNo()
               strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & strDataTemp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strExc(9) & "," & _
                  PUB_GetWorkDay1(strDate(0), True) & "," & strDate(3) & ",'" & stCP13 & "'," & strNP22 & ") "
               cnnConnection.Execute strSql
            End If
         End If
      End If
      'end 2006/8/21
      '2011/9/27 俄羅斯設計掛第16年延展費期限,國家檔設年費,年費期限由發文時處理
      'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
      'Modified by Morgan 2022/6/10 2015 2015/1/1以前提申案件除了延展費外仍要繳年費，2015/1/1以後提申案件僅須繳延展費--禧佩
      'If pa(9) = "023" And pa(8) = "3" Then
      If pa(9) = "023" And pa(8) = "3" And Val(pa(10)) < 20150101 Then
      'end 2022/6/10
         strDate(1) = cp(1)
         strDate(2) = pa(9)
         If m_CaseType = 年費 Then
            strExc(9) = "607"
            'Modified by Morgan 2015/5/28 2014/10/01之後提出申請的俄羅斯設計案,專用期間為自申請起5年,期滿可延4次,每次5年,共25年,惟自申請日起第3年(包括延展期間)須逐年繳年費(准後始須繳交)
            'Memoed by Morgan 2015/10/15 俄羅斯2015/1/1修法設計案不變
            If Val(pa(10)) >= 20141001 Then
               strDate(3) = CompDate(0, 5, pa(10))
            Else
               strDate(3) = CompDate(0, 15, pa(10))   '第16年即 +15
            End If
            'end 2015/5/28
         End If
         GetCtrlDT strDate
         strExc(0) = "select cp09 from CASEPROGRESS where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' AND CP10='" & strExc(9) & "' AND CP27 IS NULL AND CP57 IS NULL"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "UPDATE CASEPROGRESS SET CP06=" & PUB_GetWorkDay1(strDate(0), True) & " ,CP07=" & strDate(3) & " WHERE CP09='" & RsTemp("CP09") & "'"
            cnnConnection.Execute strSql
         Else
            strExc(0) = "select NP01,NP07,NP22 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06 is null and np07='" & strExc(9) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strSql = "UPDATE NEXTPROGRESS SET NP08 = " & PUB_GetWorkDay1(strDate(0), True) & ",NP09 = " & strDate(3) & _
                  " WHERE NP01='" & RsTemp("NP01") & "' AND NP07='" & RsTemp("NP07") & "' AND NP22=" & RsTemp("NP22")
               cnnConnection.Execute strSql
            Else
               strNP22 = GetNextProgressNo()
               strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & strDataTemp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strExc(9) & "," & _
                  PUB_GetWorkDay1(strDate(0), True) & "," & strDate(3) & ",'" & stCP13 & "'," & strNP22 & ") "
               cnnConnection.Execute strSql
            End If
         End If
      End If
      '2011/9/27 end
   End If
      
   If txtCaseField(2).Text <> "" And txtCaseField(3).Text <> "" Then
      '更新相同案號的案件進度檔案件性質為"領證及繳年費"(601)且"發文日"有值者, 其實際結果欄為"1"(准)
      'Modfiy by Morgan 2005/4/14 加更新公開費的結果
      'strTxt(iStep) = "Update CASEPROGRESS SET CP24='1' WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='601' AND CP27 IS NOT NULL "
      strTxt(iStep) = "Update CASEPROGRESS SET CP24='1' WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 in ('601','217') AND CP27 IS NOT NULL "
      cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
   End If
   
   '判斷申請國家專利種類是否自動發證
   'Modify by Morgan 2011/8/8 改先存變數不必每次都重抓
   'If AutoIssue(pa(9), pa(8)) = True Then
   If m_bolAutoIssue = True Then
   
        'Modify By Cheng 2002/12/09
        '自動發證國家, 若未輸入核准日時, 則以發證日為核准日
'      '更新專利基本檔"目前准駁"欄為"准"及"准駁通知日"欄為發證日
'      strTxt(iStep) = "UPDATE PATENT SET PA16='1', PA20=" & DBDATE(Me.txtCaseField(1).Text) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        If Me.txtCaseField(8).Text = "" Then
        'Modify by Morgan 2009/9/17 上面已有上准駁及核准日,此處只需判斷無核准日用發證日以免混淆
        '    strTxt(iStep) = "UPDATE PATENT SET PA16='1', PA20=" & DBDATE(Me.txtCaseField(1).Text) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        'Else
        '    strTxt(iStep) = "UPDATE PATENT SET PA16='1', PA20=" & DBDATE(Me.txtCaseField(8).Text) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
            strTxt(iStep) = "UPDATE PATENT SET PA20=" & DBDATE(Me.txtCaseField(1).Text) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
            cnnConnection.Execute strTxt(iStep)
            iStep = iStep + 1
        End If
      '911106 nick transation
      
      If txtCaseField(2).Text <> "" And txtCaseField(3).Text <> "" Then  'add by sonia 2015/12/2 為管制專利證書,若只是通知證書號數(無專用期間)則不更新進度之核准,否則trigger會自動將催審消掉
         '搜尋案件性質介於"101"-"105"或"301"-"307"且發文日最大者
         '2005/7/20 MODIFY BY SONIA 郭說只更新無准駁者,已有准駁進度檔保留原狀態
         'strSQLA = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND ( (CP10>='101' AND CP10<='105') OR (CP10>='301' AND CP10<='307')) AND CP27 IS NOT NULL ORDER BY CP27 DESC "
         'Modify by Morgan 2007/2/13 加107,113,114,204,501也要
         'StrSQLa = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND ( (CP10>='101' AND CP10<='105') OR (CP10>='301' AND CP10<='307')) AND CP27 IS NOT NULL AND CP24 IS NULL AND CP25 IS NULL ORDER BY CP27 DESC "
         'Modify by Morgan 2009/9/16 取消204否則申請程序不會更新到Ex.CFP-21797(資料已修正) --> 郭:原來可能是考慮中間接來案件,但畢竟為少數,遇到時人工處理
         'StrSQLa = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND ( (CP10>='101' AND CP10<='105') OR (CP10>='301' AND CP10<='307') OR CP10='107' OR CP10='113' OR CP10='114' OR CP10='204' OR CP10='501' ) AND CP27 IS NOT NULL AND CP24 IS NULL AND CP25 IS NULL ORDER BY CP27 DESC "
         'Modify by Morgan 2009/10/19 要排除集體設計105 CFP-17245
         'StrSQLa = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND ( (CP10>='101' AND CP10<='105') OR (CP10>='301' AND CP10<='307') OR CP10='107' OR CP10='113' OR CP10='114' OR CP10='501' ) AND CP27 IS NOT NULL AND CP24 IS NULL AND CP25 IS NULL ORDER BY CP27 DESC "
         StrSQLa = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND ( (CP10>='101' AND CP10<='104') OR (CP10>='301' AND CP10<='307') OR CP10='107' OR CP10='113' OR CP10='114' OR CP10='501' ) AND CP27 IS NOT NULL AND CP24 IS NULL AND CP25 IS NULL ORDER BY CP27 DESC "
         'End 2007/2/13
         '2005/7/20 END
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            '2005/7/14 MODIFY BY SONIA 自動發證國家, 若未輸入核准日時, 則以發證日為核准日
            'strTxt(iStep) = "UPDATE CASEPROGRESS SET CP24=" & IIf(IsNull(rsA("CP24")), "'1'", "CP24") & " ,CP25=" & DBDATE(Me.txtCaseField(1).Text) & " WHERE CP09='" & rsA("CP09") & "'"
            If Me.txtCaseField(8).Text = "" Then
               strTxt(iStep) = "UPDATE CASEPROGRESS SET CP24='1' ,CP25=" & DBDATE(Me.txtCaseField(1).Text) & " WHERE CP09='" & rsA("CP09") & "'"
            Else
               strTxt(iStep) = "UPDATE CASEPROGRESS SET CP24='1', CP25=" & DBDATE(Me.txtCaseField(8).Text) & " WHERE CP09='" & rsA("CP09") & "'"
            End If
            '2005/7/14 END
            cnnConnection.Execute strTxt(iStep)
            iStep = iStep + 1
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If   'add by sonia 2015/12/2
      
      '2013/5/15 add by sonia 日本新型自動發證,若該案有讓渡701時同時上核准(Trigger會自動更新其催審上Y)
      If pa(9) = "011" Then
         strTxt(iStep) = "update caseprogress set cp24='1',cp25=" & IIf(txtCaseField(8) = "", CNULL(TransDate(txtCaseField(1), 2)), CNULL(TransDate(txtCaseField(8), 2), True)) & " where cp10='701' and cp27 is not null and cp24 is null and " & _
                         "cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'"
         cnnConnection.Execute strTxt(iStep)
         iStep = iStep + 1
      End If
      '2013/5/15 end
   End If
   
   'Add by Morgan 2004/6/7 更新相關收文號為最小的新案總收文號
   'Modify by Morgan 2006/4/19改抓'A'類的最小收文號(因有些案子會沒上到新案)
   'strTxt(iStep) = "UPDATE CASEPROGRESS A" & _
      " SET A.CP43=(SELECT MIN(B.CP09) FROM CASEPROGRESS B WHERE B.CP01=A.CP01 AND B.CP02=A.CP02 AND B.CP03=A.CP03 AND B.CP04=A.CP04 AND B.CP31='Y')" & _
      " WHERE A.CP09='" & m_strCP09ByCheng & "' AND A.CP43 IS NULL"
   strTxt(iStep) = "UPDATE CASEPROGRESS A" & _
      " SET A.CP43=(SELECT MIN(B.CP09) FROM CASEPROGRESS B WHERE B.CP01=A.CP01 AND B.CP02=A.CP02 AND B.CP03=A.CP03 AND B.CP04=A.CP04 AND B.CP09<'B')" & _
      " WHERE A.CP09='" & m_strCP09ByCheng & "' AND A.CP43 IS NULL"
   cnnConnection.Execute strTxt(iStep)
   iStep = iStep + 1
   
   m_bolEndModCash = bolEndModCash  '2011/8/16 ADD BY SONIA 記錄是否上結餘,以便歐盟設計子案更新結餘
   'Add by Morgan 2005/5/20
   '非台灣 更新結餘
   Pub_UpdateEndModCash pa(1), pa(2), pa(3), pa(4)
   
   'Add by Morgan 2007/4/30
   'EPC或英國案須檢查是否有香港案
   m_HKCP14 = "": m_HKCP09 = "": m_HKCP10 = "": m_HKNP22 = ""
   If pa(9) = "221" Or pa(9) = "201" Then
      'Add by Morgan 2008/5/15
      'EPC 要更新指定國註冊費的期限
      m_bolEpcRegNotPaid = False
      If pa(9) = "221" And txtCaseField(1) <> "" Then
         '法限=發證日+3個月 所限=發證日+1個月
         strExc(1) = CompDate(1, 3, txtCaseField(1))
         
         strExc(2) = PUB_GetWorkDay1(CompDate(1, 1, txtCaseField(1)), True)
         
         strExc(0) = "select cp09,cp27 from caseprogress where " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
            " and cp10='224' and cp57 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '已收文
         If intI = 1 Then
            '未發文
            If IsNull(RsTemp("cp27")) Then
               m_bolEpcRegNotPaid = True
               m_strEpcRegDueDay = "" '已收文只需更新期限，定稿不必再通知
               strSql = "update caseprogress set cp06=" & strExc(2) & ",cp07=" & strExc(1) & " where cp09='" & RsTemp("cp09") & "'"
               cnnConnection.Execute strSql, intI
            End If
         '未收文
         Else
            strExc(0) = "select np01,np22 from nextprogress where " & ChgNextProgress(cp(1) & cp(2) & cp(3) & cp(4)) & _
               " and np07='224' and np06 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_bolEpcRegNotPaid = True
               m_strEpcRegDueDay = strExc(2)
               strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01='" & RsTemp("np01") & "' and np22=" & RsTemp("np22")
               cnnConnection.Execute strSql, intI
            End If
         End If

         'Added by Morgan 2023/3/8 249UP註冊期限
         strExc(1) = CompDate(1, 1, txtCaseField(1))
         strExc(2) = PUB_GetWorkDay1(CompDate(2, -14, strExc(1)), True)
         strExc(0) = "select cp09,cp27 from caseprogress where " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
            " and cp10='249' and cp57 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '已收文
         If intI = 1 Then
            '未發文
            If IsNull(RsTemp("cp27")) Then
               strSql = "update caseprogress set cp06=" & strExc(2) & ",cp07=" & strExc(1) & " where cp09='" & RsTemp("cp09") & "'"
               cnnConnection.Execute strSql, intI
            End If
         '未收文
         Else
            strExc(0) = "select np01,np22 from nextprogress where " & ChgNextProgress(cp(1) & cp(2) & cp(3) & cp(4)) & _
               " and np07='249' and np06 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01='" & RsTemp("np01") & "' and np22=" & RsTemp("np22")
               cnnConnection.Execute strSql, intI
            End If
         End If
         'end 2023/3/8
         
         
         'Added by Morgan 2020/8/18
         '更新子案224指定國註冊費的提申及催審期限
         strExc(3) = txtCaseField(1)
         '最終提申期限
         strExc(1) = PUB_Get224CtrlDate(2, strExc(3), cp)
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01 in (select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04<>'00' and cp10='224' and cp27>0) and np07='996' and np06 is null"
         cnnConnection.Execute strSql, intI
         '一般提申期限
         strExc(1) = PUB_Get224CtrlDate(3, strExc(3), cp)
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01 in (select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04<>'00' and cp10='224' and cp27>0) and np07='998' and np06 is null"
         cnnConnection.Execute strSql, intI
         '催審期限
         strExc(1) = PUB_Get224CtrlDate(1, strExc(3), cp)
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01 in (select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04<>'00' and cp10='224' and cp27>0) and np07='411' and np06 is null"
         cnnConnection.Execute strSql, intI
         'end 2020/8/18
         
         'Added by Morgan 2023/3/9
         '更新子案249 UP註冊的最終提申及催審期限(一般提申發文時用通則管制不必再更新)
         '最終提申期限
         strExc(1) = PUB_Get224CtrlDate(2, strExc(3), cp, True)
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01 in (select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04<>'00' and cp10='249' and cp27>0) and np07='996' and np06 is null"
         cnnConnection.Execute strSql, intI
         '催審期限
         strExc(1) = PUB_Get224CtrlDate(1, strExc(3), cp, True)
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01 in (select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04<>'00' and cp10='249' and cp27>0) and np07='411' and np06 is null"
         cnnConnection.Execute strSql, intI
         'end 2023/3/9
      End If
      'end 2008/5/15
      '有香港案
      If ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), m_HKPA01, m_HKPA02, m_HKPA03, m_HKPA04) = True Then
         '法限=公告(發證)日+6個月
         '公告日
         If txtCaseField(6) <> "" Then
            m_HKCP10 = "111"
            strDateS(3) = CompDate(1, 6, txtCaseField(6))
         '發證日
         ElseIf txtCaseField(1) <> "" Then
            m_HKCP10 = "111"
            strDateS(3) = CompDate(1, 6, txtCaseField(1))
         End If
         If m_HKCP10 <> "" Then
            strDateS(0) = ""
            strDateS(1) = m_HKPA01
            strDateS(2) = "013"
            GetCtrlDT strDateS
            '所限
            strDateS(4) = PUB_GetWorkDay1(strDateS(0), True)
         
            strExc(0) = "select cp09,cp14,cp27,EP06,cf04 from patent,caseprogress,engineerprogress,casefee" & _
               " where pa01='" & m_HKPA01 & "' and pa02='" & m_HKPA02 & "' and pa03='" & m_HKPA03 & "' and pa04='" & m_HKPA04 & "' and pa57 is null" & _
               " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='" & m_HKCP10 & "' and cp57 is null" & _
               " and ep02(+)=cp09 and cf01(+)=cp01 and cf02(+)='013' and cf03(+)=cp10"
         
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            '已收文
            If intI = 1 Then
               With RsTemp
               '未發文
               If IsNull(.Fields("cp27")) Then
                  m_HKCP09 = "" & .Fields("cp09")
                  '未齊備
                  If IsNull(.Fields("EP06")) Then
                     m_HKCP14 = "" & .Fields("cp14")
                     '更新齊備日
                     strSql = "update engineerprogress set ep06=" & strSrvDate(1) & " where ep02='" & m_HKCP09 & "' and ep06 is null"
                     cnnConnection.Execute strSql, intI
                     
                     If PUB_IfSetCP48(m_HKCP09) Then  'Add by Morgan 2010/10/4
                        '承辦期限
                        'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
                        'strDates(5) = CompWorkDay(Val("" & .Fields("cf04")), strSrvDate(1))
                        strDateS(5) = Pub_GetHandleDay(m_HKPA01, "013", m_HKCP10, , , m_HKCP09)
                        'end 2007/10/11
                        '更新承辦期限
                        strSql = "Update CaseProgress Set CP48=" & strDateS(5) & " Where CP09='" & m_HKCP09 & "' AND ( CP48 IS NULL OR CP48>" & strDateS(5) & ")"
                        cnnConnection.Execute strSql, intI
                     End If
                  End If
                  '更新期限
                  strSql = "Update CaseProgress Set CP06=" & strDateS(4) & ",CP07=" & strDateS(3) & " Where CP09='" & m_HKCP09 & "' and (cp07 is null or CP07>" & strDateS(3) & ") and cp27 is null"
                  cnnConnection.Execute strSql, intI
               End If
               End With
            '未收文
            Else
               m_HKNP08 = strDateS(4)
               m_HKNP09 = strDateS(3) 'Added by Morgan 2018/7/16
               strExc(0) = "select np22,np01 from nextprogress where  np02='" & m_HKPA01 & "' and np03='" & m_HKPA02 & "' and np04='" & m_HKPA03 & "' and np05='" & m_HKPA04 & "' and np07='" & m_HKCP10 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  m_HKNP01 = RsTemp.Fields("np01")  'Added by Morgan 2018/7/16
                  m_HKNP22 = RsTemp.Fields("np22")
                  strSql = "update nextprogress set np08=" & strDateS(4) & ",np09=" & strDateS(3) & " Where np22=" & m_HKNP22 & " and np01='" & RsTemp.Fields("np01") & "'"
               Else
                  m_HKNP01 = strDataTemp(9)  'Added by Morgan 2018/7/16
                  m_HKNP22 = GetNextProgressNo
                  strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22)" & _
                           " VALUES ('" & strDataTemp(9) & "','" & m_HKPA01 & "','" & m_HKPA02 & "','" & m_HKPA03 & "','" & m_HKPA04 & "'" & _
                           ",'" & m_HKCP10 & "'," & strDateS(4) & "," & strDateS(3) & ",'" & PUB_GetAKindSalesNo(m_HKPA01, m_HKPA02, m_HKPA03, m_HKPA04) & "'," & m_HKNP22 & ")"
               End If
               cnnConnection.Execute strSql, intI
            End If
         End If
      End If
   End If
   
   'Add by Morgan 2007/5/10
   If m_blnCancelClosed = True Then
      strSql = "UPDATE PATENT SET PA57=NULL,PA58=NULL,PA59=NULL" & _
         " WHERE PA01 = '" & pa(1) & "' AND PA02 = '" & pa(2) & "'" & _
         " AND PA03 = '" & pa(3) & "' AND PA04 = '" & pa(4) & "' "
      cnnConnection.Execute strSql, intI
   End If
   'end 2007/5/10

   'Add by Morgan 2010/3/11
   'EPC案需同時更新子案的相關資料
   If pa(9) = "221" Then
      strSql = "update patent a set (pa10,pa11,pa12,pa13,pa16,pa20,pa24,pa25)=(select b.pa10,b.pa11,b.pa12,b.pa13,b.pa16,b.pa20,b.pa24,b.pa25 from patent b where b.pa01=a.pa01 and b.pa02=a.pa02 and b.pa03=a.pa03 and b.pa04='00') where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04<>'00' and pa57 is null"
      cnnConnection.Execute strSql, intI
   End If
      
   'Add by Morgan 2010/5/12
   '台灣案加速審查通知
   If txtCaseField(8) <> "" And txtCaseField(8) <> txtCaseField(8).Tag Then
      strExc(0) = "select na28,na29,na49,na53 from nation where na01='" & pa(9) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '發明或新型有訂實審期限的國家
         If pa(8) = "1" Or (pa(8) = "2" And Not IsNull(RsTemp("na28")) And RsTemp("na29") > 0) Then
            If ((pa(8) = "1" And RsTemp("na49") = "Y") Or (pa(8) = "2" And RsTemp("na53") = "Y")) Then
               strExc(0) = "select cp14 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "'" & _
                  " and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and instr('" & NewCasePtyList & ",107',cp10)>0 and cp27>0 order by cp27 desc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strRefCP14 = "" & RsTemp(0)
               End If
               '台灣案已收文通知實審日且相關收文號尚無結果(不必管是否曾收文加速審查)
               '承辦人相同
               strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo" & _
                  " from casemap,patent,caseprogress a" & _
                  " where cm01='" & cp(1) & "' and cm02='" & cp(2) & "'" & _
                  " and cm03='" & cp(3) & "' and cm04='" & cp(4) & "'" & _
                  " and pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08" & _
                  " and (pa16 is null or pa16='2') and pa09='000' and pa08='1'" & _
                  " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='1204'" & _
                  " and exists(select * from caseprogress b where b.cp09=a.cp43 and b.cp24 is null" & _
                  " and b.cp14='" & strRefCP14 & "')"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
                  strExc(2) = strExc(1) & " 已核准，台灣發明案 " & RsTemp(0) & " 符合提出加速審查之條件.."
                  strExc(3) = "台灣發明案 " & RsTemp(0) & " 仍在審查中，惟其相對應 " & strExc(1) & " 已核准" & _
                     "，故台灣發明案符合提出加速審查之條件，若欲辦理，請洽承辦工程師 '||st02||'" & _
                     "('||st01||') 研討內容及評估費用。"
                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                     " select '" & strUserNum & "','" & strDataTemp(13) & "',to_char(sysdate,'yyyymmdd')" & _
                     ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','" & strExc(3) & "',st01" & _
                     " from staff where st01='" & strRefCP14 & "'"
                  cnnConnection.Execute strSql, intI
               End If
            End If
         End If
      End If
   End If
   'end 2010/5/12
   
   'Add by Morgan 2010/5/26
   '歐盟設計更新集體案基本資料
   'Modified by Morgan 2012/4/25 其他國家也要
   'If pa(8) = "3" And pa(9) = "239" Then
   'modify by sonia 2017/3/15 剔除韓國012設計案CFP-028750
   'Modified by Morgan 2023/11/23 剔除日本011設計案CFP-033709--禧佩
   If pa(8) = "3" And pa(3) = "0" And pa(9) <> "012" And pa(9) <> "011" Then
      'Modified by Lydia 2017/01/12 語法與Pub_Cache2Letter共用
      'strExc(0) = "select cp03,cp09 from caseprogress,patent where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp04='" & cp(4) & "' and cp57 is null and cp27>0 and cp10='105' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa57 is null order by cp03"
      strExc(0) = Pub_GetCFPQuery105(cp(1), cp(2), cp(3), cp(4))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         iNo = 1
         With RsTemp
         Do While Not .EOF
            intI = 0
            If bolUpdate Then
               strSql = "update caseprogress a set (cp05,cp20,cp26,cp27,cp32,cp64,cp119)" & _
                  "=(select b.cp05,b.cp20,b.cp26,b.cp27,b.cp32,b.cp64,b.cp119 from caseprogress b where b.cp09='" & strDataTemp(9) & "')" & _
                  " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & .Fields("cp03") & "' and cp04='" & cp(4) & "' and cp10='" & 專利證書 & "'"
               cnnConnection.Execute strSql, intI
            End If
            If intI = 0 Then
               strExc(1) = "cp01"
               strExc(2) = "cp01"
               For intI = 2 To TF_CP
                  Select Case intI
                     '2011/8/16 MODIFY BY SONIA 加入CP59,CP109
                     Case 16, 17, 18, 60, 65, 66, 67, 68, 69, 70, 61, 62, 63, 87, 88, 59, 109
                     Case Else
                        strExc(1) = strExc(1) & ",cp" & Format(intI, "00")
                        If intI = 3 Then
                           strExc(2) = strExc(2) & ",'" & .Fields("cp03") & "'"
                        ElseIf intI = 9 Then
                           strExc(2) = strExc(2) & ",'" & AutoNo("C", 6) & "'"
                        ElseIf intI = 43 Then
                           strExc(2) = strExc(2) & ",'" & .Fields("cp09") & "'"
                        Else
                           strExc(2) = strExc(2) & ",cp" & Format(intI, "00")
                        End If
                  End Select
               Next
               strSql = "Insert into caseprogress(" & strExc(1) & ") select " & strExc(2) & " from caseprogress where cp09='" & strDataTemp(9) & "'"
               cnnConnection.Execute strSql, intI
            End If
            
            '2011/8/16 ADD BY SONIA 集體設計子案也要更新結餘 CFP-023564 (但因各國不知是否相同故只先做歐盟設計)
            '非台灣 更新結餘
            'Removed by Morgan 2012/4/25 改在 Pub_UpdateEndModCash 內批次做
            'If m_bolEndModCash Then
            '   bolEndModCash = m_bolEndModCash
            '   Pub_UpdateEndModCash cp(1), cp(2), .Fields("cp03"), cp(4)
            'End If
            'End 2012/4/25
            '2011/8/16 END


            'Modify by Morgan 2010/7/14 +pa16,pa17
            'Modify by Morgan 2011/4/28 集體子案公告號規則同證書號(實際上兩者號數相同)
            'Modified by Morgan 2011/11/3 子案會超過10且會有部分子案閉卷情形所以證書號就會不是 pa03+1 Ex.CFP-24311
            'Modified by Morgan 2019/11/14 德國比照歐盟 Ex:CFP-25387 --玫音
            If pa(9) = "239" Or pa(9) = "231" Then 'Added by Morgan 2012/4/25
               iNo = iNo + 1
               strSql = "update patent a set (pa14,pa15,pa16,pa17,pa20,pa21,pa22,pa24,pa25)" & _
                  "=(select b.pa14,replace(b.pa15,'-0001','-" & Format(iNo, "0000") & "'),b.pa16,b.pa17,b.pa20,b.pa21,replace(b.pa22,'-0001','-" & Format(iNo, "0000") & "'),b.pa24,b.pa25" & _
                  " from patent b where b.pa01=a.pa01 and b.pa02=a.pa02 and b.pa03='0' and b.pa04=a.pa04)" & _
                  " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa04='" & pa(4) & "' and pa03='" & .Fields("cp03") & "'"
               
            'Added by Morgan 2012/4/25 非EU的子案號數與母案相同
            Else
               strSql = "update patent a set (pa14,pa15,pa16,pa17,pa20,pa21,pa22,pa24,pa25)" & _
                  "=(select b.pa14,b.pa15,b.pa16,b.pa17,b.pa20,b.pa21,b.pa22,b.pa24,b.pa25" & _
                  " from patent b where b.pa01=a.pa01 and b.pa02=a.pa02 and b.pa03='0' and b.pa04=a.pa04)" & _
                  " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa04='" & pa(4) & "' and pa03='" & .Fields("cp03") & "'"
               cnnConnection.Execute strSql, intI
            End If
            'end 2012/4/25
            
            cnnConnection.Execute strSql, intI
            
            'Added by Morgan 2012/11/20 進度檔也要更新
            strSql = "update caseprogress set (cp24,cp25)=(select pa16,pa20 from patent where pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04) where cp09='" & .Fields("cp09") & "'"
            cnnConnection.Execute strSql, intI
            'end 2012/11/20
            .MoveNext
         Loop
         End With
      End If
   End If
   'end 2010/5/26
   
   'Added by Morgan 2012/5/25
   '若有專用期則下一程序 1603 上'Y'
   If txtCaseField(3) <> "" Then
      strSql = "update nextprogress set np06='Y' where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07='1603' and np06 is null"
      cnnConnection.Execute strSql, intI
   End If
   'end 2012/5/25
   
   'Added by Lydia 2015/05/18 申請國家為048緬甸時,點選案件性質為新申請案之案件,下一程序增加刊登廣告費
   If txtCaseField(15).Visible = True And txtCaseField(16).Visible = True And txtCaseField(17).Visible = True And Val(txtCaseField(10)) > 0 Then
      strExc(8) = DBDATE(txtCaseField(18))
      strExc(9) = DBDATE(txtCaseField(17))
      
      'Added by Morgan 2020/9/23
      '檢查是否已收文刊登廣告
      strExc(0) = " select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='951' and cp57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '未發文則更新期限,若已發文則會在輸提申時新增下次的期限
         If IsNull(RsTemp("cp27")) Then
            strSql = "update caseprogress set cp06=" & strExc(8) & ",cp07=" & strExc(9) & " where cp09='" & RsTemp("cp09") & "'"
            cnnConnection.Execute strSql
         End If
      Else
      'end 2020/9/23
         
         strSql = " select * from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np07='951' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 0 Then
           strNP22 = GetNextProgressNo()
           'strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                    "SELECT np01,np02,np03,np04,np05,951,np08,np09,np10," & CNULL(strNP22) & " from nextprogress " & _
                    "where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np07='607' "
           'Modified by Morgan 2015/7/8 延展費改為公告日起算
           'strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                    "SELECT np01,np02,np03,np04,np05,951," & CNULL(strExc(8), True) & "," & CNULL(strExc(9), True) & ",np10," & CNULL(strNP22) & " from nextprogress " & _
                    "where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np07='607' "
            strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                    " values('" & m_strCP09ByCheng & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "',951," & CNULL(strExc(8), True) & "," & CNULL(strExc(9), True) & ",'" & stCP13 & "'," & strNP22 & ")"
         Else
           strSql = "update nextprogress set np08=" & CNULL(strExc(8), True) & ",np09=" & CNULL(strExc(9), True) & _
                   "where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np07='951' "
         
         End If
         cnnConnection.Execute strSql
         
      End If 'Added by Morgan 2020/9/23
   End If
   
   'add by sonia 2018/10/12土耳其235發明案或EPC221案 輸"證書號數"時更新下一程序"商業使用聲明"為公告日+3年為法限,本所=法定-2月,CFP-027741
   'modify by sonia 2020/4/4 +cp(4)="00"即EPC子案輸證書不可更新CFP-029945-0-39
   'modify by sonia 2020/7/24 土耳其加新型案
   If txtCaseField(6) <> "" And (pa(8) = "1" Or pa(8) = "2") And pa(9) = "235" And cp(4) = "00" Then
      strExc(1) = CompDate(0, 3, txtCaseField(6))   '法限
      strExc(2) = CompDate(1, -2, strExc(1))        '本所
      strExc(2) = PUB_GetWorkDay1(strExc(2), True)
      strSql = "select np22 from nextprogress" & _
         " where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "'" & _
         " and np07='930' and np06 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & _
            " where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np22=" & RsTemp.Fields("np22")
      Else
         strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
            "NP07,NP08,NP09,NP10,NP22) select '" & strDataTemp(9) & "','" & cp(1) & "'" & _
            ",'" & cp(2) & "','" & cp(3) & "','" & cp(4) & "',930," & strExc(2) & "," & strExc(1) & _
            "," & CNULL(PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))) & ",NP22 from dual,(select nvl(max(np22),0)+1 NP22 from nextprogress)"
      End If
      cnnConnection.Execute strSql, intI
      m_930DueDate = strExc(1) 'Added by Morgan 2018/10/30
      
   ElseIf txtCaseField(6) <> "" And pa(9) = "221" Then  '期限更新土耳其子案
      '先抓土耳其子案案號
      strSql = "select pa01,pa02,pa03,pa04 from patent where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & cp(3) & "'" & _
               " and pa09='235' and pa57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         Caseno235(1) = RsTemp.Fields("pa01")
         Caseno235(2) = RsTemp.Fields("pa02")
         Caseno235(3) = RsTemp.Fields("pa03")
         Caseno235(4) = RsTemp.Fields("pa04")
         strExc(1) = CompDate(0, 3, txtCaseField(6))   '法限
         strExc(2) = CompDate(1, -2, strExc(1))        '本所
         strExc(2) = PUB_GetWorkDay1(strExc(2), True)
         strSql = "select np22 from nextprogress" & _
            " where np02='" & Caseno235(1) & "' and np03='" & Caseno235(2) & "' and np04='" & Caseno235(3) & "' and np05='" & Caseno235(4) & "'" & _
            " and np07='930' and np06 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & _
               " where np02='" & Caseno235(1) & "' and np03='" & Caseno235(2) & "' and np04='" & Caseno235(3) & "' and np05='" & Caseno235(4) & "' and np22=" & RsTemp.Fields("np22")
         Else
            strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
               "NP07,NP08,NP09,NP10,NP22) select '" & strDataTemp(9) & "','" & Caseno235(1) & "'" & _
               ",'" & Caseno235(2) & "','" & Caseno235(3) & "','" & Caseno235(4) & "',930," & strExc(2) & "," & strExc(1) & _
               "," & CNULL(PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))) & ",NP22 from dual,(select nvl(max(np22),0)+1 NP22 from nextprogress)"
         End If
         cnnConnection.Execute strSql, intI
         End If
   End If
   'end 2018/10/12
   
   'Added by Morgan 2020/12/10 印度發明案商業使用聲明期限
   '3/31 以前核准->次年的9/30 Ex:2020/3/31-->2021/9/30 (提2020/4/1~2021/3/31之商業使用聲明)
   '4/01 以後核准->後年的9/30 Ex:2020/4/01-->2022/9/30 (提2021/4/1~2022/3/31之商業使用聲明)
   'modify by sonia 2024/6/19 印度修法商業使用聲明期限改為每三年呈報一次
   If pa(9) = "040" And pa(8) = "1" And txtCaseField(8) <> "" Then
      If Right(txtCaseField(8), 4) <= "0331" Then
         'modify by sonia 2024/6/19 印度修法商業使用聲明期限改為每三年呈報一次
         'strExc(1) = Left(DBDATE(txtCaseField(8)), 4) + 1
         strExc(1) = Left(DBDATE(txtCaseField(8)), 4) + 3
      Else
         'modify by sonia 2024/6/19 印度修法商業使用聲明期限改為每三年呈報一次
         'strExc(1) = Left(DBDATE(txtCaseField(8)), 4) + 2
         strExc(1) = Left(DBDATE(txtCaseField(8)), 4) + 4
      End If
      strExc(1) = strExc(1) & "0930"
      strSql = "select np01,np22,np06,np07 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' " & _
                  "and np07='930' and np06 is null and np09=" & CNULL(strExc(1), True)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 0 Then
          strExc(2) = PUB_GetWorkDay1(CompDate(1, -1, strExc(1)), True) '所限=法限-1個月
          strSql = "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                      " select '" & strDataTemp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','930'," & strExc(2) & "," & strExc(1) & ",'" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "',newNP22 from dual,(select nvl(max(np22),0)+1 newNP22 from nextprogress)"
          cnnConnection.Execute strSql, intI
      End If
            
   End If
   'end 2020/12/10
   
   'Added by Morgan 2023/4/28
   '美國通知證書號數更新下一程序尚未收文之IDS法定及本限(較晚的才要更新，只能提前不可延後)--郭
   If pa(9) = "101" And strDataTemp(10) = 通知證書號數 Then
      'Modified by Morgan 2023/7/18
      'strExc(1) = DBDATE(txtCaseField(1)) '法限=發證日
      'strExc(2) = CompDate(2, -7, strExc(1)) '所限=法限-1週
      'strExc(2) = PUB_GetWorkDay1(strExc(2), True)
      strExc(1) = PUB_GetWorkDay1(CompDate(2, -1, txtCaseField(1)), True) '法限=發證日-1天(再抓工作日)
      strExc(2) = CompWorkDay(2, CompDate(2, -1, strExc(1)), 1) '所限=法限提前2個工作天
      'end 2023/7/18
      If strExc(2) < strSrvDate(1) Then strExc(2) = strSrvDate(1)
      strSql = "update nextprogress set np09=" & strExc(1) & ",np08=" & strExc(2) & " where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06 is null and np07='214' and np09>" & strExc(1)
      cnnConnection.Execute strSql, intI
   End If
   'end 2023/4/28
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm05010402_1"
   End If
   '2016/10/7 END
   
   
   'Added by Morgan 2018/7/16 CFP電子化
   If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) Then
      If Not bolUpdate Or txtCaseField(12) <> "N" Then
         m_strLD18 = strDataTemp(9)
         If bolUpdate = False Then
            m_strCP10 = strDataTemp(10)
         Else
            strSql = "update caseprogress set cp68='" & strUserNum & "',cp69=to_number(to_char(sysdate,'YYYYMMDD')),cp70=to_number(to_char(sysdate,'HH24MI')) where cp09='" & m_strLD18 & "'"
            cnnConnection.Execute strSql, intI
            m_strCP10 = cp(10)
         End If
         strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), m_strCP10, pa(9))
         '若第二次輸證書則原信函進度將會被刪除
         PUB_AddLetterProgress m_strLD18, IIf(m_strCP10 = 通知證書號數, 1, 2), IIf(txtCaseField(12) <> "N", True, False), strExc(1), False, pa(26), m_strCP10, pa(75)
         If m_HKNP22 <> "" Then
            strExc(0) = "select pa26,pa75 from patent where " & ChgPatent(m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If PUB_AddCP1913(m_HKPA01, m_HKPA02, m_HKPA03, m_HKPA04, m_HKNP08, m_HKNP09, m_HKNP01, m_HKNP22, "013", "" & RsTemp.Fields("pa26"), m_HK1913CP09, "" & RsTemp.Fields("pa75"), , True) = False Then
               Err.Raise 999, , "新增進度檔【通知期限】失敗！作業中斷！"
            End If
         End If
         m_bolAddLP = True
      End If
   End If
   'end 2018/7/16
   
   cnnConnection.CommitTrans
   SaveData = True
   
   'Add by Morgan 2007/4/30 印結案單
   If m_HKNP22 <> "" Then
      MsgBox "請更換紙張！", , "列印接洽單！"
      g_PrtForm001.PrintForm m_HKNP22, m_HKPA01, m_HKPA02, m_HKPA03, m_HKPA04
   End If

CheckingErr:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
   End If
End Function
'Modified by Lydia 2017/04/27 改成Function
'Private Sub ReadAllData()
Private Function ReadAllData() As Boolean
Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String

On Error GoTo HndErr
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
ReadAllData = False 'Added by Lydia 2017/04/27

'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetReceiveCode(frm05010402_1.txtSystem, frm05010402_1.txtCode(0), _
   IIf(frm05010402_1.txtCode(1) = "", "0", frm05010402_1.txtCode(1)), _
   IIf(frm05010402_1.txtCode(2) = "", "00", frm05010402_1.txtCode(2)), strTemp) Then
If ClsPDGetReceiveCode(frm05010402_1.txtSystem, frm05010402_1.txtCode(0), _
   IIf(frm05010402_1.txtCode(1) = "", "0", frm05010402_1.txtCode(1)), _
   IIf(frm05010402_1.txtCode(2) = "", "00", frm05010402_1.txtCode(2)), strTemp) Then
   
   'Modify by Morgan 2006/10/19 改不Call Dll
   'If objPublicData.ReadAllData(strTemp, cp(), pA(), intCaseKind, intPWhere) Then
   ReDim cp(TF_CP) As String
   cp(9) = strTemp
   If PUB_ReadAllData(cp(), pa(), intCaseKind, intPWhere) Then
      'Modified by Morgan 2012/4/24
      'm_bolAutoIssue = PUB_AutoIssue(pa(9), pa(8), pa(10)) 'Add by Morgan 2011/8/8
      m_bolAutoIssue = PUB_AutoIssue(pa(9), pa(8), pa(10), pa)
      'end 2012/4/24
      
      'Added by Morgan 2025/8/7
      m_bolNewPlant = PUB_ChkCPExist(cp(), "120", "2")
      'end 2025/8/7
      
   'end 2006/10/19
      'Add by Morgan 2007/6/6
      m_oPA14 = pa(14)
      m_oPA20 = pa(20)
      m_oPA21 = pa(21)
      'end 2007/6/6
      stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
      stCP12 = GetSalesArea(stCP13)
      'edit by nickc 2007/02/05 不用 dll 了
      'i = obj003.ReadCCaseProgressDatabase(cp(), 專利證書, 國外_CF)
      i = Cls003ReadCCaseProgressDatabase(cp(), 專利證書, 國外_CF)
      Select Case i
                   Case 0
                              GoTo err1
                   Case 1
                              bolUpdate = True
                              txtCaseField(10) = cp(16)
                              txtCaseField(11) = cp(18)
                   Case 2
                              bolUpdate = False
      End Select
      lblCaseField(0) = pa(1) + " - " + pa(2) + _
      IIf(pa(4) = "00" And pa(3) = "0", "", " - " + pa(3)) + _
      IIf(pa(4) = "00", "", " - " + pa(4))
      lblCaseField(1) = pa(11)
      lblCaseField(2) = pa(8)
      lblCaseField(3) = pa(9)
      SetNameToCombo cboCaseName, pa(5), pa(6), pa(7)
      For i = 0 To 4
             lblCaseField(i + 4) = pa(26 + i)
      Next
      lblCaseField(9) = frm05010402_1.txtReceivedDay
      '申請日
      Me.lblCaseField(10).Caption = TransDate(pa(10), 1)
      
      txtCaseField(0) = pa(22)
      txtCaseField(1) = TransDate(pa(21), 1)
      txtCaseField(2) = pa(24)
      txtCaseField(3) = pa(25)
      If pa(16) = "1" Then
         txtCaseField(8) = TransDate(pa(20), 1)
         txtCaseField(8).Tag = txtCaseField(8)
      End If
      
      '2008/3/20 add by sonia 為防自動發證未設定改為鎖住核准日欄,自動發證或前已核駁者才可輸入
      txtCaseField(8).Enabled = False
      
      'Modify by Morgan 2011/8/8
      'If AutoIssue(pa(9), pa(8)) = True Then
      '   txtCaseField(8).Enabled = True
      txtCaseField(10).Enabled = False
      txtCaseField(11).Enabled = False
      If m_bolAutoIssue = True Then
         txtCaseField(8).Enabled = True
         txtCaseField(10).Enabled = True
         txtCaseField(11).Enabled = True
      'end 2011/8/8
      
      'Add by Lydia 2014/10/29 自動發證國家目前寄證書程序發文後，業務常要求改費用同舊案報價。修改系統在告准輸入(證書號輸入)時，會自動提供半年內之領證費報價(類一般來函輸入領證費)。
      '要注意CP64備註會不斷記錄櫃台收文日,同時CP05會UPDATE為最後輸入日期
         'Removed by Morgan 2014/11/17 改到設定有無證書時
         'SetFee
      End If
      If pa(16) = "2" Then txtCaseField(8).Enabled = True
      '2008/3/20 end
      txtCaseField(6) = TransDate(pa(14), 1)
      txtCaseField(7) = pa(15)
      strTemp = ""
      
      
      'Add by Morgan 2006/7/5
      '下次繳費日一律重算,以防止基本資料有修改而期限未同步更新
      'Modify by Morgan 2008/12/17 改呼叫公用函式
      'strNP09 = TransDate(GetMoneyDay(), 1)
      'Added by Morgan 2020/9/23 緬甸無年費制度，以刊登廣告為維護權利手段
      If pa(9) = "048" Then
         If txtCaseField(1) <> "" Then
            CheckKeyIn 1
         End If
      Else
      'end 2020/9/23
         strNP09 = TransDate(PUB_GetNextYearFeeDate(cp, pa, m_CaseType, m_PayType), 1)
      End If
      
      If ReadNextProgress() = False Then
        '2009/7/8 ADD BY SONIA EPC子案不掛年費期限 CFP-018284-0-25
        If cp(4) <> "00" Then
           m_NP09_Old = ""
           strNP09 = ""
        End If
        '2009/7/8 END
         
      End If
      
   End If
   If pa(9) <> EPC指定國家 Then
      cmdCountry.Enabled = False
   End If
   If cmdCountry.Enabled Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadCountry(intCaseKind, cp(), strCountry, True) = True Then
      If ClsPDReadCountry(intCaseKind, cp(), strCountry, True) = True Then
         If strCountry = "" Then
            ShowMsg MsgText(1051)
            GoTo err1
         End If
      End If
   Else
      strCountry = ""
   End If
   
   strPA25 = GetSDateTo 'Add by Morgan 2004/12/10
   
   'Add By Cheng 2002/07/31
   '若申請國家為"英國"且專利種類為"發明"時
   If pa(9) = "201" And pa(8) = "1" Then
      Me.Label6(2).Caption = "年費金額："
      Me.Label6(2).Visible = True
      Me.txtCaseField(14).Visible = True
      Me.txtCaseField(14).MaxLength = 0
   'Add by Morgan 2004/12/8 美國發明可輸入調整期
   ElseIf pa(9) = "101" And pa(8) = "1" Then
      Me.Label6(2).Caption = "調整期："
      Me.Label6(2).Visible = True
      Me.txtCaseField(14).Visible = True
      Me.txtCaseField(14).MaxLength = 4
      If txtCaseField(3).Text <> "" Then
         lblCaseField(11).Caption = txtCaseField(3).Text
         txtCaseField(3).Text = strPA25
         txtCaseField(14).Text = DateDiff("d", ChangeWStringToWDateString(txtCaseField(3).Text), ChangeWStringToWDateString(lblCaseField(11).Caption))
      End If
   End If
   
   'Add by Morgan 2010/3/17 若有收文領證則不可輸領證費及點數欄位
   strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='601' and cp57 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      txtCaseField(10) = "0"    '2010/7/9 ADD BY SONIA CFP-021117先收領證又是自動發證會無法操作
      txtCaseField(10).Locked = True
      txtCaseField(10).BackColor = &H8000000F
      txtCaseField(11) = "0"    '2010/7/9 ADD BY SONIA CFP-021117先收領證又是自動發證會無法操作
      txtCaseField(11).Locked = True
      txtCaseField(11).BackColor = &H8000000F
   End If
   ReadAllData = True 'Added by Lydia 2017/04/27
Else
err1:
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If
Screen.MousePointer = varSaveCursor
Exit Function
HndErr:
ErrorMsg
Screen.MousePointer = varSaveCursor
End Function
Private Sub Form_Activate()
'Add By Cheng 2003/04/23
If m_blnFirstShow = True Then
    m_blnFirstShow = False
    'Modified by Lydia 2017/04/27 改成Function ,回傳boolean判斷是否繼續作業
    'ReadAllData
    If ReadAllData = True Then
        'Added by Lydia 2015/05/18 申請國家為048緬甸時,點選案件性質為新申請案之案件,增加刊登廣告費(必填)
        If pa(9) <> "048" Then
           lblAD(0).Visible = False: lblAD(1).Visible = False: lblAD(2).Visible = False: lblAD(3).Visible = False
           txtCaseField(15).Visible = False: txtCaseField(16).Visible = False: txtCaseField(17).Visible = False: txtCaseField(18).Visible = False
           Me.Height = 5895
        End If
        'end 2015/05/18
        
        'Add by Morgan 2004/9/10
        '若無專用期間則有無證書設'N'，否則'Y'。
        txtCaseField(5).Enabled = False
        If txtCaseField(2).Text = "" And txtCaseField(3).Text = "" Then
           txtCaseField(5).Text = "N"
           txtCaseField(10) = "" 'Added by Morgan 2014/11/17
        Else
           txtCaseField(5).Text = "Y"
           If m_bolAutoIssue = True Then SetFee 'Added by Morgan 2014/11/17
        End If
    End If 'end by Lydia 2017/04/27
End If
End Sub
Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   Me.Caption = frm05010402_1.Caption
   m_blnFirstShow = True
   
   'Add By Sindy 2017/12/28
   m_strIR01 = frm05010402_1.m_strIR01
   m_strIR02 = frm05010402_1.m_strIR02
   m_strIR03 = frm05010402_1.m_strIR03
   m_strIR04 = frm05010402_1.m_strIR04
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
   PUB_SendMailCache 'Add by Morgan 2010/5/12
   Select Case intLeaveKind
      Case 0
         frm05010402_1.Show
         frm05010402_1.Clear
      Case 1
         frm05010402_1.Show
      Case 2
         Unload frm05010402_1
   End Select
   Set frm05010403_2 = Nothing
End Sub

Private Sub lblCaseField_Change(Index As Integer)
   Dim strTemp As String, strCusTemp As String
   
   Select Case Index
      Case 2
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , pA(9)) Then
         If ClsPDGetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , pa(9)) Then
            lblTrademarkKind = strTemp
         End If
      Case 3
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(lblCaseField(Index), strTemp) Then
         If ClsPDGetNation(lblCaseField(Index), strTemp) Then
            lblNation.Caption = strTemp
         End If
      Case 4, 5, 6, 7, 8
         If lblCaseField(Index) <> "" Then
            strCusTemp = lblCaseField(Index)
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetCustomer(strCusTemp, strTemp) Then
            If ClsPDGetCustomer(strCusTemp, strTemp) Then
               lblCaseField(Index) = strCusTemp
               lblPetitionName(Index - 4).Caption = strTemp
            End If
         End If
   End Select
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 5, 12:
         KeyAscii = UpperCase(KeyAscii)
      Case 13
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> 89 Then
            KeyAscii = 0
         End If
      'Add by Morgan 2004/12/9
      'Added by Lydia 2015/05/18 +048緬甸廣告費
      Case 14, 10, 15, 16
         If Not (KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57)) Then
            KeyAscii = 0
         End If
      Case Else:
   End Select
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = -1 Then
      Cancel = True
   End If
   If Cancel Then txtCaseField_GotFocus (Index)
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, strStartDate As String
Dim strTmp1 As String
Dim intPos As Integer

   CheckKeyIn = -1
   Select Case intIndex
      Case 0
         If txtCaseField(intIndex) = "" Then
            ShowMsg MsgText(1058)
         Else
            '2007/10/25 add by sonia 美國發明預設(專利號數+1格空白)至公告號
            If pa(9) = "101" And pa(8) = "1" And txtCaseField(7) = "" Then
               txtCaseField(7) = txtCaseField(0) & " "
            End If
            '2007/10/25 end
            '2011/4/27 add by sonia 歐盟設計預設專利號數至公告號CFP-023898
            If pa(9) = "239" And pa(8) = "3" Then
               txtCaseField(7) = txtCaseField(0)
            End If
            '2011/4/27 end
            CheckKeyIn = 1
         End If
         
      Case 1 '發證日
         If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
            CheckKeyIn = 1
            pa(21) = txtCaseField(1)
            '2007/10/25 add by sonia 美國發明預設發證日至公告日
            If pa(9) = "101" And pa(8) = "1" And txtCaseField(6) = "" Then
               txtCaseField(6) = txtCaseField(1)
            End If
            '2007/10/25 end
            '2011/4/27 add by sonia 歐盟設計預設發證日至公告日CFP-023898
            If pa(9) = "239" And pa(8) = "3" Then
               txtCaseField(6) = txtCaseField(1)
            End If
            '2011/4/27 end
             
            'Add by Morgan 2006/8/18 馬來西亞舊法(申請日<20010801)發明為發證日起15年
            If pa(9) = "018" And pa(8) = "1" And Val(pa(10)) > 0 And Val(TransDate(pa(10), 2)) < 20010801 Then
               m_PayType = "6"
            End If
            'end 2006/8/18
             
             '92.5.27 ADD BY SONIA 計算發證日起算之下次繳費日
             If m_PayType = "6" Then
               'Modify by Morgan 2008/12/17 改呼叫公用函式
                'strNP09 = TransDate(GetMoneyDay(), 1)
                strNP09 = TransDate(PUB_GetNextYearFeeDate(cp, pa, m_CaseType, m_PayType), 1)
                strPA25 = GetSDateTo 'Add by Morgan 2007/4/20 若為發證日起算時須重新計算
                
             'Added by Morgan 2018/7/31 --禧佩
             '歐亞專利聯盟年費為准後始須繳交,若遇到之第1次年費期限距發證日少於2個月,原年費期限可延長2個月
             ElseIf (pa(9) = "074" And pa(8) = "1") Then
               strNP09 = TransDate(PUB_GetNextYearFeeDate(cp, pa, m_CaseType, m_PayType), 1)
             'end 2018/7/31
             End If
             '93.12.22 ADD BY SONIA
             If strPA25 = "" Then strPA25 = GetSDateTo
             '93.12.22 END
             
             '92.5.27 END
             '92.2.5 ADD BY SONIA1. 澳洲設計證書計算第一年延展費期限不依國家檔，改為註冊日+1年-1個月
             'Modify by Morgan 2005/6/13 申請日<2040617才不依國家檔
             'If pa(9) = "015" And pa(8) = "3" Then
             If pa(9) = "015" And pa(8) = "3" And Val(pa(10)) < 20040617 Then
                '92.5.23 MODIFY BY SONIA 不預設, 由使用者輸入再檢查
                'txtCaseField(4) = CompDate(0, 1, txtCaseField(1))
                'txtCaseField(4) = TransDate(CompDate(1, -1, txtCaseField(4)), 1)
                strNP09 = CompDate(0, 1, txtCaseField(1))
                strNP09 = TransDate(CompDate(1, -1, strNP09), 1)
                '92.5.23 END
             End If
             '92.2.5 END
             
             'Added by Morgan 2016/10/14 新加坡發明若下次繳費日小於發證日+3個月時更新為該日期--禧佩
             '印度發明案應該規則相同但先不改等有案例再說(發證日<原期限<發證日+3月時是否更新問題)--禧佩
             If pa(9) = "014" And pa(8) = "1" Then
               strExc(1) = TransDate(CompDate(1, 3, txtCaseField(1)), 1)
               If Val(strNP09) < Val(strExc(1)) Then
                  strNP09 = strExc(1)
               End If
             End If
             'end 2016/10/14
             
             'add by toni 20080924 印度(040)下次繳費日小於發証日時改為發証日+3個月
             'Modify by Morgan 2010/8/11 百年蟲
             'If pa(9) = "040" And pa(8) = "1" And strNP09 < txtCaseField(1) Then
             If pa(9) = "040" And pa(8) = "1" And Val(strNP09) < Val(txtCaseField(1)) Then
               strNP09 = TransDate(CompDate(1, 3, txtCaseField(1)), 1)
             End If
             
             '英國(201)下次繳費日之年月<=發証日之年月時改為發証日+3個月
             If pa(9) = "201" And pa(8) = "1" And Mid(TransDate(strNP09, 2), 1, 6) <= Mid(TransDate(txtCaseField(1), 2), 1, 6) Then
               '2009/2/4 MODIFY BY SONIA 再改為核准日+3個月   2019/6/27慧汶說再改回發證日+3個月
               strNP09 = TransDate(CompDate(1, 3, txtCaseField(1)), 1)
               'If txtCaseField(8) <> "" Then strNP09 = TransDate(CompDate(1, 3, txtCaseField(8)), 1)  'cancel by sonia 2019/6/27慧汶說再改回發證日+3個月
               '2009/2/4 END
             End If
             '20080924 end
             '2010/8/2 add by sonia 以色列發明以發證日+3個月更新年費期限
             If pa(9) = "027" And pa(8) = "1" Then
                strNP09 = TransDate(CompDate(1, 3, txtCaseField(1)), 1)
             End If
             '2010/8/2 END
             
            'Added by Morgan 2016/1/13
            'PCT進紐西蘭之發明案若PCT申請日是在2014/9/13新法實施前則年費為准後繳且期限為自發證日起4個月
            If pa(46) = "Y" And pa(9) = "016" And pa(8) = "1" And DBDATE(pa(10)) < "20140913" Then
               'Modified by Morgan 2016/2/16 若原期限較晚時保留--禧佩 CFP-25891
               'strNP09 = TransDate(CompDate(1, 4, txtCaseField(1)), 1)
               strExc(1) = TransDate(CompDate(1, 4, txtCaseField(1)), 1)
               If Val(strExc(1)) > Val(strNP09) Then
                  strNP09 = strExc(1)
               End If
               'end 2016/2/16
            End If
            'end 2016/1/13
   
            'Added by Morgan 2016/10/26
            '印尼發明及新型的年費期限第1次為發證日+6個月,第2次以後為屆滿前1個月
            If pa(9) = "017" And (pa(8) = "1" Or pa(8) = "2") And pa(72) = "" Then
               'Removed by Morgan 2021/11/5 已修正為核准日6個月內須繳交自申請日起算累計至核准日次年之年費
               'strNP09 = TransDate(CompDate(1, 6, txtCaseField(1)), 1)
               'end 2021/11/5
               strNP09 = TransDate(GetIDN1st605FeeDate(txtCaseField(8)), 1) 'Added by Morgan 2022/3/22
            End If
            'end 2016/10/26
      
            'Added by Morgan 2020/9/23
            '緬甸第一次刊登廣告期限：以「發證日+6個月」為法定期限，本所期限為法定期限提早1個月。
            If pa(9) = "048" Then
               txtCaseField(17) = TransDate(CompDate(1, 6, txtCaseField(1)), 1)
               txtCaseField(18) = TransDate(PUB_GetWorkDay1(CompDate(1, -1, txtCaseField(17)), True), 1)
            End If
            'end 2020/9/23
         End If
      Case 2
         If txtCaseField(intIndex) = "" Then
            CheckKeyIn = 1
         Else
            Dim strPA24 As String
            strPA24 = GetSDateFrom
            If strPA24 <> txtCaseField(2) Then
               MsgBox "專用期間起日應為<" & strPA24 & "> !"
            Else
               CheckKeyIn = 1
               If strPA25 <> "" And txtCaseField(3) = "" Then txtCaseField(3) = strPA25 'Added by Morgan 2019/11/14 預設專用期止日
            End If
         End If
         'Add by Morgan 2004/9/10
         If txtCaseField(2) = "" And txtCaseField(3) = "" Then
            txtCaseField(5) = "N"
            txtCaseField(10) = "" 'Added by Morgan 2014/11/17
         Else
            txtCaseField(5) = "Y"
            If m_bolAutoIssue = True Then SetFee 'Added by Morgan 2014/11/17
         End If
         
      Case 3
         If txtCaseField(intIndex) = "" Then
            CheckKeyIn = 1
         Else
            'Modify by Morgan 2004/12/10 移到全域變數
            'Dim strPA25 As String
            'Modify by Morgan 2004/12/10 移到ReadAllData
            'strPA25 = GetSDateTo
            
            If strPA25 <> txtCaseField(3) Then
               MsgBox "專用期間止日應為<" & strPA25 & "> !"
            Else
               CheckKeyIn = 1
            End If
         End If
         'Add by Morgan 2004/9/10
         If txtCaseField(2) = "" And txtCaseField(3) = "" Then
            txtCaseField(5) = "N"
            txtCaseField(10) = "" 'Added by Morgan 2014/11/17
         Else
            txtCaseField(5) = "Y"
            If m_bolAutoIssue = True Then SetFee  'Added by Morgan 2014/11/17
         End If
      
      Case 4 '下次繳費日
         If Me.txtCaseField(intIndex).Text <> "" Then
            '2009/7/8 ADD BY SONIA EPC子案不掛年費期限 CFP-018284-0-25
            If cp(4) <> "00" Then
               MsgBox "EPC子案不必輸入下次繳費日!!!", vbExclamation
               CheckKeyIn = -1
               Exit Function
            End If
            '2009/7/8 END
            'Add By Cheng 2002/03/11
            If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                If CheckReKey(txtCaseField(intIndex)) Then
                   CheckKeyIn = 1
                End If
            End If
'2012/2/23 cancel by sonia 移至TxtValidate檢查,否則會重覆出現訊息
'            If Val(Me.txtCaseField(intIndex).Text) < Val(strSrvDate(2)) Then
'               MsgBox "下次繳費日不可小於系統日期!!!", vbExclamation
'               CheckKeyIn = -1
'               Exit Function
'            End If
'2012/2/23 end
            '92.5.23 ADD BY SONIA
            If Val(DBDATE(Me.txtCaseField(intIndex).Text)) <> Val(DBDATE(strNP09)) Then
               MsgBox "下次繳費日法定期限應為 " & strNP09, vbCritical
               CheckKeyIn = -1
               Exit Function
            End If
            '92.5.23 END
         'Add by Morgan 2008/3/19 下次繳費日不可大於專用期止日
         ElseIf txtCaseField(3) <> "" And Val(DBDATE(txtCaseField(3))) < Val(DBDATE(strNP09)) Then
            CheckKeyIn = 1
         Else
            '92.5.23 ADD BY SONIA
            If Val(strNP09) <> 0 Then
               MsgBox "下次繳費日法定期限應為 " & strNP09, vbCritical
               CheckKeyIn = -1
               Exit Function
            Else
               CheckKeyIn = 1
            End If
         End If
                  
      Case 5
         If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Or txtCaseField(intIndex) = "Y" Then
            If txtCaseField(intIndex).Text = "" And pa(9) = 美國國家代號 Then
               ShowMsg MsgText(1060)
            Else
               CheckKeyIn = 1
            End If
         Else
            ShowMsg MsgText(9177)
         End If
                  
      Case 6 '公告日
         If Me.txtCaseField(intIndex).Text <> "" Then
            If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                CheckKeyIn = 1
            End If
         Else
            CheckKeyIn = 1
         End If
               
      '2007/10/25 add by sonia
      Case 7 '公告號
         If pa(9) = "101" And pa(8) = "1" And txtCaseField(2) <> "" Then
            intPos = InStr("" & txtCaseField(7), "B1")
            If intPos = 0 Then
               intPos = InStr("" & txtCaseField(7), "B2")
               If intPos = 0 Then
                  MsgBox "美國發明案公告號應輸入B1或B2,以便判斷是否退公開費 ! ", vbCritical
                  CheckKeyIn = -1
                  Exit Function
               Else
                  CheckKeyIn = 1
               End If
            Else
               CheckKeyIn = 1
            End If
         Else
            CheckKeyIn = 1
         End If
      '2007/10/25 end
      
      Case 8 '核准日
         If Me.txtCaseField(intIndex).Text <> "" Then
             If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                 CheckKeyIn = 1
             End If
'cancel by sonia 2019/6/27慧汶說再改回發證日+3個月
'             '2009/2/4 ADD BY SONIA英國(201)下次繳費日之年月<=發証日之年月時改為核准日+3個月
'             If pa(9) = "201" And pa(8) = "1" And Mid(TransDate(strNP09, 2), 1, 6) <= Mid(TransDate(txtCaseField(1), 2), 1, 6) Then
'                strNP09 = TransDate(CompDate(1, 3, txtCaseField(8)), 1)
'             End If
'             '2009/2/4 end
'end 2019/6/27
            
            'Added by Morgan 2022/3/22
            '印尼發明/新型第1次年費期限
            If pa(9) = "017" And (pa(8) = "1" Or pa(8) = "2") And pa(72) = "" Then
               strNP09 = TransDate(GetIDN1st605FeeDate(txtCaseField(8)), 1)
            End If
            'end 2022/3/22
         Else
             CheckKeyIn = 1
         End If
         
      Case 11
         'Modify By Sindy 2024/9/10
         'If txtCaseField(10) <> "" And txtCaseField(intIndex) = "" Then
         If txtCaseField(10) <> "" And Val(txtCaseField(intIndex)) = 0 Then
         '2024/9/10 END
            MsgBox "請同時輸入領證費及點數!!!", vbExclamation + vbOKOnly
         Else
            'add by sonia 2023/12/13
            'Modify By Sindy 2024/9/10
            'If txtCaseField(10) = "" And txtCaseField(intIndex) <> "" Then
            If txtCaseField(10) = "" And Val(txtCaseField(intIndex)) > 0 Then
            '2024/9/10 END
               MsgBox "無領證費時不可輸入點數!!!", vbExclamation + vbOKOnly
            Else
            'end 2023/12/13
               CheckKeyIn = 1
            End If
         End If
            
      Case 12
         If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
            CheckKeyIn = 1
         Else
            ShowMsg MsgText(9177)
         End If
            
      Case 14 '年費金額
         If Me.txtCaseField(14).Visible Then
            If Me.txtCaseField(14).Text <> "" Then
               If IsNumeric(Me.txtCaseField(14).Text) = False Then
                  MsgBox "年費金額輸入錯誤!!!", vbExclamation + vbOKOnly
               Else
                  CheckKeyIn = 1
                  'Add by Morgan 2004/12/8
                  If pa(9) = "101" And txtCaseField(3).Text <> "" Then
                     lblCaseField(11).Caption = CompDate(2, Val(Me.txtCaseField(14).Text), txtCaseField(3).Text)
                  End If
               End If
            Else
               CheckKeyIn = 1
            End If
         Else
            CheckKeyIn = 1
         End If
      'Added by Lydia 2015/05/18  申請國家為048緬甸時,點選案件性質為新申請案之案件,增加刊登廣告費(必填)
      Case 15
         If txtCaseField(intIndex).Visible = True Then
            If txtCaseField(intIndex).Text = "" Then
               MsgBox "緬甸專利案有刊登廣告費！", vbExclamation + vbOKOnly
            ElseIf IsNumeric(txtCaseField(intIndex).Text) = False Then
                 MsgBox "刊登廣告費金額輸入錯誤!!!", vbExclamation + vbOKOnly
            Else
                 CheckKeyIn = 1
            End If
            If CheckKeyIn <> 1 Then txtCaseField(intIndex).SetFocus 'Added by Morgan 2020/9/23
         Else
            CheckKeyIn = 1
         End If
      Case 16
         If txtCaseField(intIndex).Visible = True Then
            If IsNumeric(txtCaseField(intIndex).Text) = False Then
               MsgBox "刊登廣告費點數輸入錯誤!!!", vbExclamation + vbOKOnly
            Else
               CheckKeyIn = 1
            End If
            If CheckKeyIn <> 1 Then txtCaseField(intIndex).SetFocus 'Added by Morgan 2020/9/23
         Else
            CheckKeyIn = 1
         End If
      Case 17, 18
         If txtCaseField(intIndex).Visible = True Then
            If txtCaseField(intIndex).Text = "" Then
                MsgBox "緬甸專利案有刊登廣告期限！", vbExclamation + vbOKOnly
            Else
               If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                   If intIndex = 17 Then
                      CheckKeyIn = 1
                        '本所期限=法限-14天,若非工作天則抓最近工作天
                        strExc(7) = CompDate(2, -14, DBDATE(txtCaseField(17).Text))
                        txtCaseField(18).Text = TransDate(PUB_GetWorkDay1(strExc(7), True), 1)
                   Else
                     'Modified by Lydia 2017/07/31 改為預設和檢查
                     'If PUB_CheckCP0607(0, txtCaseField(18).Text, txtCaseField(17).Text) Then
                     'Modified by Lyddia 2023/11/08 傳入必需欄位
                     'If PUB_CheckCP0607(0, txtCaseField(18), txtCaseField(17)) Then
                     If PUB_CheckCP0607(0, txtCaseField(18), txtCaseField(17), "", pa(9), cp(1), IIf(pa(24) = "", 通知證書號數, 專利證書)) Then
                        If ChkWorkDay(DBDATE(txtCaseField(intIndex).Text)) = True Then
                           CheckKeyIn = 1
                        Else
                           MsgBox "刊登廣告期限請輸入工作日！", vbExclamation + vbOKOnly
                        End If
                     End If
                   End If
               End If
            End If
         Else
            CheckKeyIn = 1
         End If
      'end 2015/05/18
      Case Else
         CheckKeyIn = 1
   End Select
   
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
Dim intPos As Integer

   txtCaseField(Index).SelStart = 0
   txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
   '儲存未修改前之值至Tag中,供再確認時使用
   txtCaseField(Index).Tag = txtCaseField(Index)
   '2007/10/25 add by sonia
   If Index = 7 Then
      If pa(9) = "101" And pa(8) = "1" And txtCaseField(7) <> "" Then
         intPos = InStr("" & txtCaseField(7), "B")
         If intPos - 1 >= 0 Then
            txtCaseField(7).SelStart = intPos - 1
            txtCaseField(7).SelLength = 0
         Else
            txtCaseField(7).SelStart = Len(txtCaseField(7)) + 2
            txtCaseField(7).SelLength = 0
         End If
      End If
   End If
   '2007/10/25 end
   
   'txtCaseField(Index).SetFocus 'Added by Morgan 2021/12/8 'Removed by Morgan 2023/5/15 原來是為了要能顯示游標，但重輸期限後會造成無窮迴圈，故取消。
End Sub

Private Function ReadNextProgress() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   ReadNextProgress = False
   m_NP22 = Empty
   strSql = "SELECT * FROM NEXTPROGRESS " & _
            "WHERE " & _
                  "NP02 = '" & cp(1) & "' AND " & _
                  "NP03 = '" & cp(2) & "' AND " & _
                  "NP04 = '" & cp(3) & "' AND " & _
                  "NP05 = '" & cp(4) & "' AND " & _
                  "(NP06 IS NULL OR NP06 = '' OR NP06 = ' ')"
                  
   'Modify by Morgan 2006/8/3 考慮當為延展期限,年費期限同時存在的國家(馬來西亞)時抓國家檔設定
   'strSQL = strSQL & " AND (NP07 = '605' OR NP07 = '606' OR NP07 = '607') "
   If m_CaseType <> "" Then
      strSql = strSql & " AND NP07 = " & m_CaseType
   Else
      strSql = strSql & " AND (NP07 = '605' OR NP07 = '606' OR NP07 = '607') "
   End If
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ReadNextProgress = True
      If Not IsNull(rsTmp.Fields("NP22")) Then
         m_NP22 = rsTmp.Fields("NP22")
      End If
      If Not IsNull(rsTmp.Fields("NP09")) Then
         '92.5.23 MODIFY BY SONIA 不預設, 由使用者輸入再檢查
         m_NP09_Old = TransDate(DBDATE(rsTmp.Fields("NP09")), 1)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Remove by Morgan 2011/4/28 不再使用
'Private Function GetMoneyDay() As String
'Dim strTemp As String, strTemp1 As String, strTemp2 As String, dobDateAdd As Double
'Dim varTemp As Variant, strStartDate As String, strDate As String, StrDate1 As String
'Dim i As Integer, yearTemp As String
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'Dim stDateType As String, stYear As String
'
'On Error GoTo HndErr
''modify by sonia 91.1.29(先抓已繳年度, 若無再抓第一年, 未繳時不要放,,,,,)
'' 91.01.28 modify by louis (證書固定是第一年)
''strTemp = GetMoneyYears(pa(72))
''strTemp = 1
'yearTemp = GetMoneyYears(pa(72))
'
''Modify by Morgan 2005/3/21 改GetMoneyYears
'''93.10.29 modify by sonia 應抓下次繳費年度,以發證日計算者仍為第一年
'''If yearTemp <> 1 Then yearTemp = yearTemp + 1
''If pa(72) = "" Then
''   yearTemp = 1
''Else
''   yearTemp = yearTemp + 1
''End If
'''93.10.29 end
''2003/3/21 end
'
'm_CaseType = Empty: m_PayType = Empty
'If GetNationTaxEx(Val(pa(8)), pa(9), strTemp, strTemp1, 年費, strTemp2, m_CaseType) Then
'   '2006/1/27 ADD BY SONIA菲律賓修法
'   If pa(9) = "030" And pa(10) <> "" And Val(pa(10)) < 19980101 Then
'      m_PayType = "6"
'      strTemp = m_PayType
'      Select Case pa(8)
'         Case "1"
'            strTemp1 = "5,6,7,8,9,10,11,12,13,14,15,16,17"
'            strTemp2 = 17
'         Case "2"
'            strTemp1 = "5,10,15"
'            strTemp2 = 15
'         Case "3"
'            strTemp1 = "5,10,15"
'            strTemp2 = 15
'      End Select
'
'   'Add by Morgan 2006/8/3 馬來西亞修法,舊法發明為發證日起15年,新型為發證日起15年
'   ElseIf pa(9) = "018" Then
'      Select Case pa(8)
'         Case "1"
'            '發明申請日2001/8/1以後適用新法
'            If Val(pa(10)) > 0 And Val(TransDate(pa(10), 2)) < 20010801 Then
'               m_PayType = "6"
'            End If
'
'         Case "2"
'            '新型發證日2001/8/1以後適用新法
'            'Remove by Morgan 2007/7/26 新型不管何時提申均適用新法(2003年8月14日實施)
'            'If Val(pa(21)) > 0 And Val(TransDate(pa(21), 2)) < 20010801 Then
'            '   m_PayType = "6"
'            '   strTemp = "6"
'            '   strTemp1 = "5,10"
'            '   strTemp2 = 10
'            'End If
'      End Select
'
'   'Add by Morgan 2007/4/20 日本設計自申請日2007/4/1起改用新法,舊法為發證日起15年,第4年起逐年繳交。
'   ElseIf pa(9) = "011" Then
'      If pa(8) = "3" Then
'         If Val(pa(10)) > 0 And Val(TransDate(pa(10), 2)) < 20070401 Then
'            strTemp1 = "4,5,6,7,8,9,10,11,12,13,14,15"
'            strTemp2 = 15
'         End If
'      End If
'   '2008/9/25 韓國新型自申請日2006/10/1起改用新法,舊法為發證日起第2,4,5,6,7,8,9,10年起逐年繳交。
'   ElseIf pa(9) = "012" Then
'      If pa(8) = "2" Then
'         If Val(pa(10)) > 0 And Val(TransDate(pa(10), 2)) < 20061001 Then
'            strTemp1 = "2,4,5,6,7,8,9,10"
'         End If
'      End If
'   '2008/9/25 END
'   End If
'   '2006/1/27 END
'
'   varTemp = Split(strTemp1, ",")
''modify by sonia 91.1.29(先抓已繳年度, 若無再抓第一年)
'   'dobDateAdd = varTemp(0)
'   dobDateAdd = varTemp(yearTemp - 1)
'
'   strStartDate = GetStartDate(strTemp, cp(), pa())
'   'Add by Morgan 2007/8/7 加考慮年費起算日與繳年費起算日不同時 CFP-14786
'   If strStartDate <> "" Then
'      strExc(0) = "select * from nation where na01='" & pa(9) & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         Select Case pa(8)
'            Case "1"
'               stDateType = RsTemp.Fields("na42")
'            Case "2"
'               stDateType = RsTemp.Fields("na45")
'            Case "3"
'               stDateType = RsTemp.Fields("na48")
'         End Select
'      End If
'      If stDateType <> strTemp Then
'         '考慮民國年或空白
'         Select Case stDateType
'            Case 收文日
'               If Len(cp(5)) > 4 Then
'                  stYear = Left(cp(5), Len(cp(5)) - 4)
'               End If
'            Case 申請日
'               If Len(pa(10)) > 4 Then
'                  stYear = Left(pa(10), Len(pa(10)) - 4)
'               End If
'            Case 公開日
'               If Len(pa(12)) > 4 Then
'                  stYear = Left(pa(12), Len(pa(12)) - 4)
'               End If
'            Case 准駁日
'               If Len(pa(20)) > 4 Then
'                  stYear = Left(pa(20), Len(pa(20)) - 4)
'               End If
'            Case 公告日
'               If Len(pa(14)) > 4 Then
'                  stYear = Left(pa(14), Len(pa(14)) - 4)
'               End If
'            Case 發證日
'               If Len(pa(21)) > 4 Then
'                  stYear = Left(pa(21), Len(pa(21)) - 4)
'               End If
'            Case 發文日
'               If Len(cp(27)) > 4 Then
'                  stYear = Left(cp(27), Len(cp(27)) - 4)
'               End If
'         End Select
'         If stYear <> "" Then
'            strStartDate = stYear & Right(strStartDate, 4)
'         End If
'      End If
'   End If
'
'   If strStartDate <> "" Then
'      '91.12.22 ADD BY SONIA
'      If pa(9) = "012" And pa(8) = "2" And pa(10) <> "" And Val(pa(10)) < 19990701 Then dobDateAdd = 3
'      '91.12.22 END
'      If m_CaseType = "605" Then
'         'Modify by Morgan 2005/9/4
'         'strStartDate = CompDate(0, (dobDateAdd - 1), strStartDate)
'         strDate = CompDate(0, (dobDateAdd - 1), strStartDate)
'      Else
'         'Modify by Morgan 2005/9/4
'         'strStartDate = CompDate(0, dobDateAdd, strStartDate)
'         strDate = CompDate(0, dobDateAdd, strStartDate)
'      End If
'      'Modify by Morgan 2005/9/4
'      'strDate = strStartDate
'   End If
'   'strDate1 = ChangeWDateStringToWString(DateAdd("M", -1, ChangeWStringToWDateString(strDate)))
'   GetMoneyDay = strDate
'
'   'Modify by Morgan 2005/9/4
'   'If GetMoneyDay > strSrvDate(2) Then
'   If GetMoneyDay > strSrvDate(1) Then
'      Exit Function
'   End If
'
'   '2008/9/25 ADD BY SONIA 英國,印度發明第一年若過期改為發證日+三個月故不抓下一年
'   If (pa(9) = "201" Or pa(9) = "040") And pa(8) = "1" Then
'      Exit Function
'   End If
'   '2008/9/25 END
'
'   '93.12.22 add 准後繳年費者
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   StrSQLa = "Select * From Nation Where NA01='" & pa(9) & "' "
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'       Select Case pa(8)
'       Case "1" '發明
'           '若非准後繳費者
'           If "" & rsA("NA56").Value <> "Y" Then
'               If rsA.State <> adStateClosed Then rsA.Close
'               Set rsA = Nothing
'               Exit Function
'           End If
'       Case "2" '新型
'           '若非准後繳費者
'           If "" & rsA("NA57").Value <> "Y" Then
'               If rsA.State <> adStateClosed Then rsA.Close
'               Set rsA = Nothing
'               Exit Function
'           End If
'       Case "3" '設計
'           '若非准後繳費者
'           If "" & rsA("NA58").Value <> "Y" Then
'               If rsA.State <> adStateClosed Then rsA.Close
'               Set rsA = Nothing
'               Exit Function
'           End If
'       End Select
'   End If
'
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'
'CompNextDate:
'
'   yearTemp = yearTemp + 1
'   If yearTemp > UBound(varTemp) Then Exit Function
'
'   dobDateAdd = varTemp(yearTemp - 1)
'   '法定期限
'   If m_CaseType = "605" Then
'      GetMoneyDay = CompDate(0, (dobDateAdd - 1), strStartDate)
'   Else
'      GetMoneyDay = CompDate(0, dobDateAdd, strStartDate)
'   End If
'
'   '若法定期限小於系統日期
'   If GetMoneyDay < strSrvDate(1) Then
'      '重算
'      GoTo CompNextDate
'   '若法定期限大於等於系統日期
'   Else
'      Exit Function
'   End If
'   '93.12.22 end
'End If
'HndErr:
'Exit Function
'End Function

'' 1 : 發明年費
'' 2 : 新型年費
'' 3 : 設計年費
'' 4 : 發明實體審查
'' 5 : 新型實體審查
'' 6 : 設計實體審查
'' 7 : 發明公開
'' 8 : 新型公開
'' 9 : 設計公開
''10 : 商標
''取得年費等期限
'Public Function GetNationTaxEx(ByRef intChoose As Integer, ByRef strNation As String, ByRef strStartUpDay As String, ByRef strYears As String, Optional strProperty As String, Optional strPayYears As String, Optional strCaseType As String) As Boolean
'Dim strSql As String, rsRecordset As New ADODB.Recordset, strTemp As String
'
'On Error GoTo HndErr
'Select Case intChoose
'   Case 1
'      strSql = "select na06,na21,na07,na20"
'   Case 2
'      strSql = "select na08,na23,na09,na22"
'   Case 3
'      strSql = "select na10,na25,na11,na24"
'   Case 4
'      strSql = "select na26,na27"
'   Case 5
'      strSql = "select na28,na29"
'   Case 6
'      strSql = "select na30,na31"
'   Case 7
'      strSql = "select na32,na33"
'   Case 8
'      strSql = "select na34,na35"
'   Case 9
'      strSql = "select na36,na37"
'   Case 10
'      strSql = "select na12,na13"
'End Select
'strSql = strSql + " from nation where na01=" + CNULL(strNation)
'rsRecordset.CursorLocation = adUseClient
'rsRecordset.Open strSql, cnnConnection
'If rsRecordset.RecordCount > 0 Then
'   strStartUpDay = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
'   strYears = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
'   If rsRecordset.Fields.Count = 4 Then
'      strPayYears = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
'      strTemp = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
'      strCaseType = strTemp
'      m_PayType = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
'   End If
'End If
''91.12.22 MODIFY BY SONIA
''If (rsRecordset.Fields.Count = 2 And strYears = "") Or (rsRecordset.Fields.Count = 4 And strYears = "" And strTemp <> strProperty) Then
'If (rsRecordset.Fields.Count = 2 And strPayYears = "") Or (rsRecordset.Fields.Count = 4 And strPayYears = "") Then
''91.12.22 END
'   ShowMsg MsgText(9104)
'Else
'   GetNationTaxEx = True
'End If
'rsRecordset.Close
'Exit Function
'HndErr:
'End Function

Private Function GetSDateFrom() As String
   
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strOpt As String
   Dim stPA10Old As String
   
   GetSDateFrom = Empty
   strOpt = Empty
   strSql = "SELECT * FROM NATION " & _
            "WHERE NA01 = '" & pa(9) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Select Case pa(8)
         ' 發明
         Case "1":
            If Not IsNull(rsTmp.Fields("NA40")) Then
               strOpt = rsTmp.Fields("NA40")
            End If
         ' 新型
         Case "2":
            If Not IsNull(rsTmp.Fields("NA43")) Then
               strOpt = rsTmp.Fields("NA43")
            End If
         ' 設計
         Case "3":
            If Not IsNull(rsTmp.Fields("NA46")) Then
               strOpt = rsTmp.Fields("NA46")
            End If
      End Select
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   If Not IsEmptyText(strOpt) Then
      '2015/4/16 ADD BY SONIA 舊法為發證日起15年
      If pa(9) = "012" And pa(8) = "3" And pa(10) <> "" Then
         If Val(pa(10)) < 20140701 Then
            strOpt = "6"
         End If
      End If
      '2015/4/16 END
      '93.9.29 ADD BY SONIA 舊法為發證日起10年
      If pa(9) = "017" And pa(8) = "2" Then
         If pa(10) <> "" And Val(pa(10)) < 20010801 Then
            strOpt = "6"
         End If
      End If
      '93.9.29 END
      '2006/1/27 ADD BY SONIA菲律賓修法
      If pa(9) = "030" And pa(10) <> "" And Val(pa(10)) < 19980101 Then
         strOpt = "6"
      End If
      '2006/1/27 END
      
      'Add by Morgan 2006/8/3 馬來西亞修法,舊法發明為發證日起15年,新型為發證日起5年
      If pa(9) = "018" Then
         Select Case pa(8)
            Case "1"
               '發明申請日2001/8/1以後適用新法
               If Val(pa(10)) > 0 And Val(TransDate(pa(10), 2)) < 20010801 Then
                  strOpt = "6"
               End If
            Case "2"
               '新型發證日2001/8/1以後適用新法
               'Remove by Morgan 2007/7/26 新型不管何時提申均適用新法(2003年8月14日實施)
               'If Val(pa(21)) > 0 And Val(TransDate(pa(21), 2)) < 20010801 Then
               '   strOpt = "6"
               'End If
         End Select
      End If
      'end 2006/8/3
   
      '91.12.22 ADD BY SONIA
      If strOpt = "6" Then pa(21) = TransDate(txtCaseField(1), 2)
      If strOpt = "4" Then pa(20) = TransDate(txtCaseField(8), 2)
      '91.12.22 END
            
      'Added by Morgan 2013/10/2 印度,馬來西亞設計專用期間為優先權日起算
      'Modified by Morgan 2013/12/31 +印尼設計--禧佩
      'Removed by Morgan 2014/1/13 取消印尼設計--禧佩
      'modify by sonia 2018/9/18 巴基斯坦發明及設計之專利期間及年費均從優先權日起算--禧佩2016/10/19Morgan少改起日的控制
      stPA10Old = pa(10)
      'Removed by Morgan 2021/3/3 都已改為發證日起算(國家檔設定)，此處取消以免混淆
      'If ((pa(9) = "040" Or pa(9) = "018") And pa(8) = "3") Or pa(9) = "038" Then
      '   strExc(1) = PUB_GetFirstPriDate(pa())
      '   If strExc(1) <> "" Then
      '      pa(10) = strExc(1)
      '   End If
      ''Added by Morgan 2020/9/2
      ''301南非 起日=發證日起算9個月
      'ElseIf pa(9) = "301" Then
      '   pa(10) = CompDate(1, 9, pa(21))
      ''end 2020/9/2
      'End If
      'end 2021/3/3
      'end 2013/10/2
      
      GetSDateFrom = GetStartDate(strOpt, cp(), pa())
      
      pa(10) = stPA10Old 'Added by Morgan 2013/10/2
   End If
End Function

Private Function CheckDivision(ByRef p_PA25 As String, Optional p_PA10 As String) As Boolean
   
On Error GoTo ErrHnd
   'Modify by Morgan 2011/8/17
   '分割子案只能是獨立案
   strSql = "SELECT PA25,PA10 FROM DIVISIONCASE,PATENT" & _
      " WHERE DC01='" & pa(1) & "' AND DC02='" & pa(2) & "' AND DC03='0' AND DC04='" & pa(4) & "'" & _
      " AND PA01(+)=DC05 AND PA02(+)=DC06 AND PA03(+)=DC07 AND PA04(+)=DC08"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         p_PA25 = "" & .Fields(0)
         p_PA10 = "" & .Fields(1)
         CheckDivision = True
      End If
   End With
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Function
Private Function GetSDateTo() As String
   
   'Add by Morgan 2005/8/10
   '判斷是否為分割案，若是則抓母案專用期止日
   Dim stPA25 As String, stPA10 As String, stPA10Old As String
   Dim stReduecOneDay As String 'Added by Morgan 2019/11/4
   Dim stToDate As String, stToDate2 As String 'Added by Morgan 2020/7/28
   
   'Remove by Morgan 2006/9/11 修正為以母案申請日計算原專用期止日(不含調整期) --CFP018255
   '移到下面
   'If CheckDivision(stPA25) = True Then
      'GetSDateTo = stPA25
      'Exit Function
   'End If
   'End 2006/9/11
   
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strOpt As String
   Dim strDate As String
   Dim strYear As String
   GetSDateTo = Empty
   strOpt = Empty
   strDate = Empty
   strYear = "0"
   strSql = "SELECT * FROM NATION " & _
            "WHERE NA01 = '" & pa(9) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Select Case pa(8)
         ' 發明
         Case "1":
            If Not IsNull(rsTmp.Fields("NA41")) Then
               strOpt = rsTmp.Fields("NA41")
            End If
            If Not IsNull(rsTmp.Fields("NA07")) Then
               strYear = rsTmp.Fields("NA07")
            End If
            stReduecOneDay = "" & rsTmp.Fields("NA82") 'Added by Morgan 2019/11/4
         ' 新型
         Case "2":
            If Not IsNull(rsTmp.Fields("NA44")) Then
               strOpt = rsTmp.Fields("NA44")
            End If
            If Not IsNull(rsTmp.Fields("NA09")) Then
               strYear = rsTmp.Fields("NA09")
            End If
            stReduecOneDay = "" & rsTmp.Fields("NA83") 'Added by Morgan 2019/11/4
         ' 設計
         Case "3":
            If Not IsNull(rsTmp.Fields("NA47")) Then
               strOpt = rsTmp.Fields("NA47")
            End If
            If Not IsNull(rsTmp.Fields("NA11")) Then
               strYear = rsTmp.Fields("NA11")
            End If
            stReduecOneDay = "" & rsTmp.Fields("NA84") 'Added by Morgan 2019/11/4
      End Select
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   'Added by Morgan 2015/5/28 2014/10/01之前提出申請的俄羅斯新型案,專用期間為自申請日起算13年
   'Removed by Morgan 2015/10/15 俄羅斯2015/1/1修法所有未發證未延展新型案專用期都為10年不得延展
   'If pa(9) = "233" And pa(8) = "2" And Val(pa(10)) < 20141001 Then
   '   strYear = 13
   'End If
   'end 2015/10/15
   'end 2015/5/28
   
   'Added by Morgan 2022/12/5 --禧佩
   '阿拉伯聯合大公國設計2021/12/31以前提申的為自申請日起10年，須自申請日起第2年逐年繳年費。
   If pa(9) = "034" And pa(8) = "3" And pa(10) <> "" Then
      If Val(pa(10)) <= 20211231 Then
         strYear = 10
      End If
   End If
   'end 2022/12/5
   
   '91.12.22 ADD BY SONIA
   If pa(9) = "012" And pa(8) = "2" And pa(10) <> "" Then
      If Val(pa(10)) < 19990701 Then
         strYear = 15
      End If
   End If
   '91.12.22 END
   '2015/4/16 ADD BY SONIA 舊法為發證日起15年
   If pa(9) = "012" And pa(8) = "3" And pa(10) <> "" Then
      If Val(pa(10)) < 20140701 Then
         strYear = 15
         strOpt = "6"
      End If
   End If
   '2015/4/16 END
   '93.9.29 ADD BY SONIA 舊法為發證日起10年
   If pa(9) = "017" And pa(8) = "2" Then
      If pa(10) <> "" And Val(pa(10)) < 20010801 Then
         strOpt = "6"
      End If
   End If
   '93.9.29 END
   'Add by Morgan 2005/4/28
   If pa(9) = "011" Then
      Select Case pa(8)
         '日本新型專用期限2005年3月31以前提申案件自申請日起6年,自發證日第4年起應繳年費
         Case "2"
            If Val(pa(10)) <= 20050331 Then
               '2006/1/27 MODIFY BY SONIA
               'GetSDateTo = CompDate(2, -1, CompDate(0, 6, TransDate(pa(10), 2)))
               'Exit Function
               strYear = 6
               '2006/1/27 END
            End If
         'Add by Morgan 2007/4/20
         '日本設計自申請日2007/4/1起改用新法,舊法為發證日起15年,第4年起逐年繳交。
         Case "3"
            If Val(pa(10)) < 20070401 Then
               strYear = 15
            'Added by Morgan 2020/3/9
            '日本設計自申請日2020/4/1起改用新法,舊法為發證日起20年,第4年起逐年繳交。
            ElseIf Val(pa(10)) < 20200401 Then
               strOpt = "6"
               strYear = 20
            End If
      End Select
   End If
   '2005/4/28 end

   '2006/1/27 ADD BY SONIA菲律賓修法
   If pa(9) = "030" And pa(10) <> "" And Val(pa(10)) < 19980101 Then
      strOpt = "6"
      Select Case pa(8)
         Case "1"
            strYear = 17
         Case Else
            strYear = 15
      End Select
   End If
   '2006/1/27 END
   
   'Add by Morgan 2006/8/3 馬來西亞修法,舊法發明為發證日起15年,新型為發證日起5年
   If pa(9) = "018" Then
      Select Case pa(8)
         Case "1"
            '發明申請日2001/8/1以後適用新法
            If Val(pa(10)) > 0 And Val(pa(10)) < 20010801 Then
               strOpt = "6"
               strYear = 15
            End If
         Case "2"
            '新型發證日2001/8/1以後適用新法
            'Remove by Morgan 2007/7/26 新型不管何時提申均適用新法(2003年8月14日實施)
            'If Val(pa(21)) > 0 And Val(pa(21)) < 20010801 Then
            '   strOpt = "6"
            '   strYear = 15
            'End If
      End Select
   End If
   'end 2006/8/3
   
   'Add by sonia 2014/4/30 美國設計案2013年12月18日修法,舊法為發證日起14年,2013年12月18日當日或以後為發證日起15年,2015/1/19 慧汶通知取消, 仍為發證日起14年
   'add by sonia 2015/11/30 2015/5/8 慧汶通知確定於2015/5/13起生效,改為15年
   If pa(9) = "101" Then
      Select Case pa(8)
         Case "3"
            '設計申請日2015/5/13以前為舊法
            If Val(pa(10)) > 0 And Val(pa(10)) < 20150513 Then
               strYear = 14
            End If
      End Select
   End If
   'end 2015/11/30
   
   'add by sonia 2019/3/14 以色列
   If pa(9) = "027" Then
      Select Case pa(8)
         Case "3"
            '設計於2018/8/7修法,舊法為自申請日起算18年,自申請日起第5、10、15年應繳延展費。
            If Val(pa(10)) > 0 And Val(pa(10)) < 20180807 Then
               strYear = 18
            End If
      End Select
   End If
   'end 2019/3/14
   
   'Added by Morgan 2021/5/12 墨西哥
   If pa(9) = "104" Then
      Select Case pa(8)
         Case "2"
            '墨西哥新型改為自申請日起15年(2020年11月5日當日或以後提申之案件適用)
            '原為申請日起10年
            If Val(pa(10)) > 0 And Val(pa(10)) < 20201105 Then
               strOpt = "2"
               strYear = 10
            End If
      End Select
   End If
   'end 2021/5/12
   
   'Added by Morgan 2025/8/7
   '澳洲植物新品種保護的專利權期間為自發證日起算25年，自發證日起第2年須逐年繳年費--禧佩
   If pa(9) = "015" And m_bolNewPlant Then
      strOpt = "6"
      strYear = 25
   End If
   'end 2025/8/7
   
   If Not IsEmptyText(strOpt) And Not IsEmptyText(strYear) Then
      '91.12.22 ADD BY SONIA
      If strOpt = "6" Then
         pa(21) = TransDate(txtCaseField(1), 2)
      End If
      '91.12.22 END
      
      'Modify by Morgan 2006/9/11 分割案以母案申請日計算原專用期止日(不含調整期) --CFP018255
      'strDate = GetStartDate(strOpt, cp(), pA())
      'Modify by Morgan 2007/7/20 接續案的專用期止日同母案--CFP-015676-1
      stPA10Old = pa(10)
      If pa(3) <> "0" Then
         strExc(0) = "select pa10 from patent where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='0' and pa04='" & pa(4) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            pa(10) = "" & RsTemp.Fields("pa10")
         End If
      'Modify by Morgan 2011/8/17 兩段獨立檢查
      'ElseIf CheckDivision(stPA25, stPA10) = True Then
      End If
      If CheckDivision(stPA25, stPA10) = True Then
      'end 2011/8/17
         pa(10) = stPA10
      
      'Added by Morgan 2013/7/2
      '040印度新式樣以最早優先權日起算
      'Modified by Morgan 2013/10/2 +018馬來西亞設計
      'Modified by Morgan 2013/12/31 +印尼設計--禧佩
      'Removed by Morgan 2014/1/13 取消印尼設計--禧佩
      'Modified by Morgan 2016/10/19 038巴基斯坦發明及設計之專利期間及年費均從優先權日起算--禧佩
      'Modified by Morgan 2021/3/3  +016紐西蘭
      'Modified by Morgan 2021/5/12 016紐西蘭改只有設計案 --慧汶
      ElseIf ((pa(9) = "040" Or pa(9) = "018" Or pa(9) = "016") And pa(8) = "3") Or (pa(9) = "038" And (pa(8) = "1" Or pa(8) = "3")) Then
         strExc(1) = PUB_GetFirstPriDate(pa())
         If strExc(1) <> "" Then
            pa(10) = strExc(1)
         End If
         
      End If
      strDate = GetStartDate(strOpt, cp(), pa())
      pa(10) = stPA10Old
      'end 2006/9/11
       
      If Not IsEmptyText(strDate) Then
         'Modify by Morgan 2004/3/10
         '修正DateSerial函數年月加減問題
         'GetSDateTo = DBDATE(DateSerial(Val(DBYEAR(strDate)) + Val(strYear), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) - 1))
         'Modified by Morgan 2019/11/7 加判斷國家檔是否要減1天的設定(另考慮起算日為2/29問題，改先減1天再加上專用年度)
         'GetSDateTo = DBDATE(DateAdd("d", -1, DateAdd("yyyy", Val(strYear), ChangeWStringToWDateString(DBYEAR(strDate) & DBMONTH(strDate) & DBDAY(strDate)))))
         strDate = DBDATE(strDate)
         'Modified by Morgan 2019/11/12 改用函數
         'If stReduecOneDay = "N" Then
         '   GetSDateTo = DBDATE(DateAdd("yyyy", Val(strYear), ChangeWStringToWDateString(strDate)))
         'Else
         '   GetSDateTo = DBDATE(DateAdd("yyyy", Val(strYear), DateAdd("d", -1, ChangeWStringToWDateString(strDate))))
         'End If
         stToDate = PUB_GetEndDate(strDate, Val(strYear), stReduecOneDay, pa(9))
         GetSDateTo = stToDate
         'end 2019/11/7
         
         'Added by Morgan 2020/7/28
         '加拿大設計申請日在2018年11月05日之後的案件，專利權是(1)申請日起算15年或(2)註冊日起算10年，以兩者中較晚到期者為準
         If pa(9) = "102" And pa(8) = "3" And DBDATE(pa(10)) > "20181105" Then
            stToDate2 = PUB_GetEndDate(pa(10), 15, stReduecOneDay, pa(9))
            If stToDate2 > stToDate Then
               GetSDateTo = stToDate2
            End If
         End If
         'end 2020/7/28
      End If
   End If
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False

   '92.12.25 add by sonia
   '2008/3/20 由下方提上來 by sonia
   If Me.txtCaseField(8).Text = "" And pa(9) = "201" And pa(8) = "1" Then
       MsgBox "此案為英國發明案, 請輸入核准日!!!", vbExclamation + vbOKOnly
       Exit Function
   End If
   If Me.txtCaseField(8).Text = "" And pa(16) = "2" Then
       MsgBox "此案曾核駁, 請輸入核准日!!!", vbExclamation + vbOKOnly
       Me.txtCaseField(8).SetFocus
       TextInverse Me.txtCaseField(8)
       Exit Function
   End If
   '92.12.25 end
    
    'Add By Cheng 2002/10/01
    '若申請國家專利種類非自動發證, 核准日不可空白
    'Modify by Morgan 2011/8/8 改先存變數不必每次都重抓
    'If AutoIssue(pa(9), pa(8)) = False Then
    If m_bolAutoIssue = False Then
    
        If Me.txtCaseField(8).Text = "" Then
           '2008/3/20 modify by sonia
            'MsgBox "申請國家非自動發證國家，請輸入核准日!!!", vbExclamation + vbOKOnly
            'Me.txtCaseField(8).SetFocus
            'TextInverse Me.txtCaseField(8)
            MsgBox "未輸入核准，是否為自動發證？否則應先做核准輸入!!!", vbExclamation + vbOKOnly
            '2008/3/20 end
            Exit Function
        End If
    'Modify by Morgan 2005/1/27 加判斷有證書才要
    'Add by Morgan 2005/1/19 自動發證時檢查一定要輸入領證費
    'Modify by Morgan 2005/2/17 可輸入0
    'ElseIf Val(txtCaseField(10)) = 0 And txtCaseField(5) = "Y" Then
    'Modified by Morgan 2015/1/14 +lock 判斷
    'ElseIf txtCaseField(10) = "" And txtCaseField(5) = "Y" Then
    ElseIf txtCaseField(10).Locked = False And txtCaseField(10) = "" And txtCaseField(5) = "Y" Then
      MsgBox "申請國家為自動發證國家，請輸入領證費!!!" & vbCrLf & vbCrLf & "( 若不請款請輸入 0 )", vbExclamation + vbOKOnly
      Me.txtCaseField(10).SetFocus
      TextInverse Me.txtCaseField(10)
      Exit Function
    '2005/1/19 end
    End If
    
   '92.4.16 Add By sonia
   If pa(9) = "221" And Me.txtCaseField(6).Text = "" Then
      MsgBox "申請國家為ＥＰＣ，請輸入公告日!!!", vbExclamation + vbOKOnly
      Me.txtCaseField(6).SetFocus
      TextInverse Me.txtCaseField(6)
      Exit Function
   End If
   '92.4.16 END
   
   'Add by Morgan 2005/1/26 若已開收據且費用不同時提示不可存檔
   If bolUpdate = True And cp(60) <> "" Then
      If cp(16) <> Val(txtCaseField(10)) Or cp(18) <> Val(txtCaseField(11)) Then
         MsgBox "已開收據，費用或點數不可修改。請通知財務處作廢該收據後再作業！"
         txtCaseField(10).SetFocus
         txtCaseField_GotFocus 10
         Exit Function
      End If
   End If
   
   '2012/2/23 add by sonia CFP-020624下次繳費日已過期(發證較晚)
   If Me.txtCaseField(4).Text <> "" And Val(Me.txtCaseField(4).Text) < Val(strSrvDate(2)) Then
      If MsgBox("下次繳費日不可小於系統日期!!，是否要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         txtCaseField(4).SetFocus
         txtCaseField_GotFocus 4
         Exit Function
      End If
   End If
   '2012/2/23 end
   
   'Modified by Morgan 2019/11/7 從最上面移下來(要先檢查，否則領證費清除又會預設導致沒有檢查到!!) Ex:CFP-026567
   For Each objTxt In Me.txtCaseField
      If objTxt.Enabled = True Then
         Cancel = False
         txtCaseField_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   'end 2019/11/7

   TxtValidate = True
End Function

'Add By Cheng 2002/10/01
'取得專利種類的專利年度
Private Function GetPatentYear(ByVal strNA01 As String, ByVal strPA08 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

StrSQLa = "Select * From Nation Where NA01='" & strNA01 & "'"
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic
If rsA.RecordCount > 0 Then
   Select Case strPA08
   Case 1 '發明
      GetPatentYear = "" & rsA("NA07").Value
   Case 2 '新型
      GetPatentYear = "" & rsA("NA09").Value
   Case 3 '設計
      GetPatentYear = "" & rsA("NA11").Value
   End Select
Else
   GetPatentYear = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Remove by Morgan 2011/8/8 改共用 PUB_AutoIssue
''Add By Cheng 2002/10/01
''判斷專利種類是否自動發證
'Private Function AutoIssue(ByVal strNA01 As String, ByVal strPA08 As String) As Boolean
'Dim rsA As New ADODB.Recordset
'Dim StrSQLa As String
'
'StrSQLa = "Select * From Nation Where NA01='" & strNA01 & "'"
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic
'If rsA.RecordCount > 0 Then
'   Select Case strPA08
'   Case 1 '發明
'      AutoIssue = IIf("" & rsA("NA49").Value = "", False, True)
'   Case 2 '新型
'      AutoIssue = IIf("" & rsA("NA53").Value = "", False, True)
'      '2009/1/20 add by sonia CFP-021410
'      If strNA01 = "012" And pa(10) < 20061001 Then
'         AutoIssue = True
'      End If
'      '2009/1/20 end
'   Case 3 '設計
'      AutoIssue = IIf("" & rsA("NA54").Value = "", False, True)
'   End Select
'Else
'   AutoIssue = False
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'
'End Function

'取得EPC國家
'Removed by Morgan 2023/8/14 改用 ClsPDReadCountry
'Private Function GetEPCNations(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String) As String
'Dim StrSQLa As String
'Dim rsA As New ADODB.Recordset
'
'GetEPCNations = ""
''Modify by Morgan 2006/7/21
''StrSQLa = "Select PA09 From Patent Where PA01='" & strPA01 & "' And PA02='" & strPA02 & "' And PA03='" & strPA03 & "' And PA04<>'00'  Order By PA04 "
''Modify by Morgan 2008/9/19 改抓領證有發文的 CFP-15807
''StrSQLa = "Select PA09 From Patent Where PA01='" & strPA01 & "' And PA02='" & strPA02 & "' And PA03='" & strPA03 & "' And PA04<>'00' and pa57 is null  Order By PA04 "
'StrSQLa = "Select PA09,PA04 From caseprogress,Patent Where cp01='" & strPA01 & "' And cp02='" & strPA02 & "' And cp03='" & strPA03 & "' And cp04<>'00' and cp10='601' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp27>0"
'StrSQLa = StrSQLa & " union Select PA09,PA04 From caseprogress,Patent Where cp01='" & strPA01 & "' And cp02='" & strPA02 & "' And cp03='" & strPA03 & "' And cp04<>'00' and cp10='224' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cp27>0  Order By PA04 "
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'    While Not rsA.EOF
'        GetEPCNations = GetEPCNations & GetNationName("" & rsA.Fields(0).Value, 0) & "、"
'        rsA.MoveNext
'    Wend
'    If GetEPCNations <> "" Then
'        GetEPCNations = Left(GetEPCNations, Len(GetEPCNations) - 1)
'    End If
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'
'End Function

'Add by Morgan 2007/4/30
Private Sub RunHKInform(p_bolEdit As Boolean)
   Dim strTxt(1 To 3) As String
   EndLetter "08", m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04 & "&000", "12", strUserNum
   strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('08','" & m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04 & "&000" & "','12','" & strUserNum & _
               "','本所期限','" & m_HKNP08 & "')"
   strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('08','" & m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04 & "&000" & "','12','" & strUserNum & _
       "','下一程序','" & m_HKCP10 & "')"
   
   'Added by Morgan 2021/1/8
   strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('08','" & m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04 & "&000" & "','12','" & strUserNum & _
       "','香港母案國家','" & lblNation & "')"
   'end 2021/1/8
   If Not ClsLawExecSQL(3, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   NowPrint m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04 & "&000", "08", "12", p_bolEdit, strUserNum, , , , , , , , , , , , , m_HK1913CP09
   
   'Added by Morgan 2018/7/17 CFP電子化
   If m_bolAddLP And p_bolEdit Then
      frm1105_1.m_RecNo = m_HK1913CP09
      frm1105_1.m_PdfName = PUB_CaseNo2FileName(m_HKPA01, m_HKPA02, m_HKPA03, m_HKPA04) & ".1913.CUS.PDF"
      frm1105_1.Show
   End If
   'end 2018/7/17
End Sub

'Added by Morgan 2020/12/16
Private Sub SetNoReceipt()
   cnnConnection.Execute "update caseprogress set cp20='N' where cp09='" & m_strLD18 & "' and cp16>0 and cp60 is null", intI
End Sub
