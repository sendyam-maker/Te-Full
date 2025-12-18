VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010603_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人其他來函輸入                    "
   ClientHeight    =   5772
   ClientLeft      =   168
   ClientTop       =   960
   ClientWidth     =   8508
   ControlBox      =   0   'False
   LinkTopic       =   "Form25"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   8508
   Begin VB.TextBox txtFiles 
      Height          =   270
      Left            =   7845
      MaxLength       =   2
      TabIndex        =   10
      Top             =   3510
      Width           =   375
   End
   Begin VB.CommandButton CmdAFID03 
      Caption         =   "申5"
      Height          =   270
      Index           =   4
      Left            =   7680
      TabIndex        =   64
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdAFID03 
      Caption         =   "申4"
      Height          =   270
      Index           =   3
      Left            =   7200
      TabIndex        =   63
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdAFID03 
      Caption         =   "申3"
      Height          =   270
      Index           =   2
      Left            =   6720
      TabIndex        =   62
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdAFID03 
      Caption         =   "申2"
      Height          =   270
      Index           =   1
      Left            =   6240
      TabIndex        =   61
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdAFID03 
      Caption         =   "申請人1"
      Height          =   270
      Index           =   0
      Left            =   5400
      TabIndex        =   60
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6320
      TabIndex        =   17
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5508
      TabIndex        =   16
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7560
      TabIndex        =   18
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   975
      Left            =   1080
      TabIndex        =   14
      Top             =   4140
      Width           =   7335
      _ExtentX        =   12933
      _ExtentY        =   1715
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
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
      _Band(0).Cols   =   2
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   3
      Left            =   5775
      TabIndex        =   5
      Top             =   2940
      Width           =   360
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "635;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   13
      Left            =   7260
      TabIndex        =   13
      Top             =   3810
      Width           =   855
      VariousPropertyBits=   671107099
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   11
      Left            =   5400
      TabIndex        =   12
      Top             =   3840
      Width           =   975
      VariousPropertyBits=   671107099
      Size            =   "1720;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   50
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
      Index           =   1
      Left            =   5400
      TabIndex        =   1
      Top             =   2340
      Width           =   1935
      VariousPropertyBits=   671107099
      MaxLength       =   40
      Size            =   "3413;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   570
      Index           =   12
      Left            =   1080
      TabIndex        =   15
      Top             =   5160
      Width           =   7335
      VariousPropertyBits=   -1467987941
      ScrollBars      =   2
      Size            =   "12938;1005"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   2640
      Width           =   615
      VariousPropertyBits=   671107099
      MaxLength       =   4
      Size            =   "1085;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   8
      Left            =   1080
      TabIndex        =   8
      Top             =   3540
      Width           =   975
      VariousPropertyBits=   671107099
      MaxLength       =   9
      Size            =   "1720;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   4
      Left            =   1080
      TabIndex        =   3
      Top             =   2940
      Width           =   975
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1720;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   5
      Left            =   3240
      TabIndex        =   4
      Top             =   2940
      Width           =   975
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1720;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   7
      Left            =   6180
      TabIndex        =   6
      Top             =   3240
      Width           =   360
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "635;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   6
      Left            =   1080
      TabIndex        =   7
      Top             =   3240
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
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   2340
      Width           =   615
      VariousPropertyBits=   671107099
      MaxLength       =   4
      Size            =   "1085;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   9
      Left            =   5400
      TabIndex        =   9
      Top             =   3540
      Width           =   975
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1720;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   10
      Left            =   1440
      TabIndex        =   11
      Top             =   3840
      Width           =   495
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "873;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "附件檔案數量："
      Height          =   180
      Left            =   6570
      TabIndex        =   67
      Top             =   3555
      Width           =   1260
   End
   Begin VB.Label Label10 
      Caption         =   "是否列印通知函：        （N：不印）"
      Height          =   255
      Left            =   4320
      TabIndex        =   66
      Top             =   2940
      Width           =   2910
   End
   Begin VB.Label Label7 
      Caption         =   "國外ID號數："
      Height          =   255
      Left            =   4320
      TabIndex        =   65
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "點數："
      Height          =   255
      Left            =   6600
      TabIndex        =   59
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label112 
      Caption         =   "費　　用："
      Height          =   255
      Left            =   4320
      TabIndex        =   58
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label lblNation 
      Height          =   255
      Left            =   6000
      TabIndex        =   57
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   9
      Left            =   5280
      TabIndex        =   56
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   4320
      TabIndex        =   55
      Top             =   2040
      Width           =   975
   End
   Begin MSForms.Label lblSales 
      Height          =   255
      Left            =   6120
      TabIndex        =   54
      Top             =   1740
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   1800
      TabIndex        =   53
      Top             =   1140
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
      Index           =   6
      Left            =   1080
      TabIndex        =   52
      Top             =   1740
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   51
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblProperty 
      Height          =   255
      Left            =   1800
      TabIndex        =   49
      Top             =   2340
      Width           =   2415
   End
   Begin VB.Label Label12 
      Caption         =   "本案期限："
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   4140
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "進度備註："
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   5220
      Width           =   975
   End
   Begin VB.Label Label19 
      Caption         =   "承辦期限："
      Height          =   255
      Left            =   4320
      TabIndex        =   46
      Top             =   3540
      Width           =   975
   End
   Begin VB.Label Label28 
      Caption         =   "下一程序："
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label29 
      Caption         =   "承辦人："
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   3540
      Width           =   975
   End
   Begin VB.Label lblNextCaseProperty 
      Height          =   255
      Left            =   1800
      TabIndex        =   43
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label Label31 
      Caption         =   "本所期限："
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label32 
      Caption         =   "法定期限："
      Height          =   255
      Left            =   2160
      TabIndex        =   41
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "機關文號："
      Height          =   255
      Left            =   4320
      TabIndex        =   40
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "申請人："
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   38
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   8
      Left            =   1320
      TabIndex        =   37
      Top             =   2040
      Width           =   1095
   End
   Begin MSForms.Label lblCaseProperty 
      Height          =   255
      Left            =   1680
      TabIndex        =   36
      Top             =   1740
      Width           =   2535
      VariousPropertyBits=   27
      Size            =   "4471;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   35
      Top             =   540
      Width           =   2655
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   34
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   5160
      TabIndex        =   33
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   7
      Left            =   5280
      TabIndex        =   32
      Top             =   1740
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "智權人員："
      Height          =   255
      Left            =   4320
      TabIndex        =   31
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label lblIssue 
      Caption         =   "收文日："
      Height          =   255
      Left            =   4320
      TabIndex        =   30
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收文號："
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "來函收文日："
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "案件性質："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "申請案號："
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   540
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   24
      Top             =   540
      Width           =   2775
   End
   Begin VB.Label Label14 
      Caption         =   "來函性質："
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "是否修改通知函內容：            (Y:Word)"
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   3240
      Width           =   3225
   End
   Begin VB.Label Label18 
      Caption         =   "是否閉卷：           （Y：閉卷）"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3240
      Width           =   2670
   End
   Begin MSForms.Label lblPromoter 
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   3540
      Width           =   2055
      VariousPropertyBits=   27
      Size            =   "3625;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label34 
      Caption         =   "是否算案件數：            （N：不算）"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   2895
   End
End
Attribute VB_Name = "frm02010603_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/10 改成Form2.0 (txtCaseField,cboCaseName,lblAgent...)
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
'Memo By Morgan 2012/12/17 智權人員欄已修改
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，2:結束  1:回上一畫面  0:確定
Dim intLeaveKind As Integer
' 90.07.02 modify by louis
' 系統別
Dim m_PA01 As String
' 國家
Dim m_PA09 As String
' 91.01.22 modify by louis
Dim m_NewCP09 As String
'Add By Cheng 2002/12/09
Dim m_blnClosed As Boolean '是否閉卷
Dim m_blnCancelClosed As Boolean '是否取銷閉卷
Dim m_strCloseDate As String '閉卷日期
'Add By Cheng 2003/04/16
'edit by nickc 2007/02/02
'Dim Ncp(T_CP) As String
Dim Ncp() As String

Dim strAutoNumber As String
'Add by Morgan 2004/1/13
Dim m_blnCustReturnSheet As Boolean '判斷是否列印案件回覆單
'Add by Morgan 2004/2/18
'若承辦人是王協理且未發文則要發EMail通知
Dim stCP09 As String, stCP14 As String
Dim m_bolActive As Boolean 'Add by Morgan 2005/5/18 Active事件是否已觸發
'Add by Morgan 2006/6/27
Dim m_901CP09 As String '901內部收文之總收文號
Dim m_901CP12 As String '901內部收文之業務區
Dim m_901CP13 As String '901內部收文之智權人員
Dim stCP12 As String, stCP13 As String '最新收文智權人員,業務區
Dim m_bMala1205 As Boolean
Dim m_CP14ST06 As String '2010/1/20 add by sonia 承辦人所別
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim m_PrevForm As Form 'Add By Sindy 2016/10/11
Dim m_bolAddLP As Boolean, m_strCP10 As String, m_strLD18 As String 'Added by Morgan 2018/7/18 CFP電子化
Dim m_bolValidateErr As Boolean 'Added by Morgan 2021/12/10

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 '********* 90.11.14   nickc
 Dim strTxt(1 To 6) As String, iStep As Integer, strTmp As Variant
 Dim strTemp1 As String, strStartDate As String, strTemp As Variant
 Dim bolTmp As Boolean, StrExt1 As String, StrExt2 As String, i As Integer
 Dim rsTmp As New ADODB.Recordset
 '********************************
   EndLetter ET01, lblCaseField(4), ET03, strUserNum
   EndLetter ET01, m_NewCP09, ET03, strUserNum
   Dim Jjj As Integer
   Jjj = 1
   
   '92.6.24 add by sonia 代理人請款
   If txtCaseField(0) = "1908" Then
      If CheckStr(txtCaseField(11)) <> "" Then
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
            "','費用','" & txtCaseField(11) & "')"
         Jjj = Jjj + 1
      End If
   End If
   
   '若有輸入下一程序
   If Me.txtCaseField(2).Text <> "" Then
       strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
          "','下一程序','" & txtCaseField(2) & "')"
       Jjj = Jjj + 1
   End If
   If Me.lblNextCaseProperty.Caption <> "" Then
       strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
          "','下一程序名稱','" & Me.lblNextCaseProperty.Caption & "')"
       Jjj = Jjj + 1
   End If
   'End
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(Jjj - 1, strTxt) Then
   If Not ClsLawExecSQL(Jjj - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim i As Integer, strTmp As String
 
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         For i = 0 To 10
            'Added by Lydia 2015/04/10
            If i <> 3 Then
                If txtCaseField(i).Enabled Then
                   If CheckKeyIn(i) <> 1 Then
                      txtCaseField(i).SetFocus
                      txtCaseField_GotFocus (i)
                      Exit For
                   End If
                End If
            End If
         Next
         If i = 11 Then
            If txtCaseField(6) = "Y" Then
               If CheckCloseFile = False Then
                  GoTo EXITSUB
               End If
            End If
            '重新檢查欄位有效性
            If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
            'Add By Sindy 2022/7/21
            If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
               If PUB_ChkFileOpening2(m_PrevForm.m_strFullFileName, "後續才能一併歸卷！") = True Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            '2022/7/21 END
            If SaveDatabase Then
               PUB_CtrlDateAlert m_NewCP09 'Add by Morgan 2011/7/14
                '若承辦人是王協理且未發文則要發EMail通知
                'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
                If stCP14 = "99050" Then
                    Call PUB_SendMail(strUserNum, "99050", stCP09, "分案通知")
                End If
            
               '92.1.10 add by sonia  來函性質1908(代理人請款)套空白定稿
               'Modified by Morgan 2019/2/19 改畫面控制
               'If cp(1) = "CFP" And txtCaseField(0) = "1908" Then
               If txtCaseField(3) <> "N" Then
               'end 2019/2/19
                  StartLetter "06", "00"
                  NowPrint m_NewCP09, "06", "00", IIf(txtCaseField(7) = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                  'Added by Morgan 2018/7/19 CFP電子化
                  If m_bolAddLP And txtCaseField(7) = "Y" Then
                     frm1105_1.m_RecNo = m_strLD18
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(field(1), field(2), field(3), field(4)) & "." & m_strCP10 & ".CUS.PDF"
                     frm1105_1.Show
                  End If
                  'end 2018/7/19
               '列印案件回覆單
               ElseIf m_blnCustReturnSheet = True Then
                  'Modified by Morgan 2018/7/19 CFP電子化, 配合轉pdf到卷宗區(LP41='Y')改共用一般來函的定稿CFP-03-000-12(內容相同)
                  'StartLetter "06", "01"
                  'NowPrint m_NewCP09, "06", "01", False, strUserNum
                  StartLetter "03", "12"
                  NowPrint m_NewCP09, "03", "12", False, strUserNum, , , , , , , , , , , , , m_strLD18
                  'end 2018/7/19
               End If

               bolLeave = True
               intLeaveKind = 0
               Unload Me
            Else
                MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
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
EXITSUB:
End Sub

Private Function SaveDatabase() As Boolean
   Dim i As Integer, m_PrintCForm As String
   Dim strTxt(1 To 30) As String, iStep As Integer
   Dim lMax As Long
   Dim bolNP22 As Boolean, NP22(1 To 3) As String, iNP22 As Integer
   Dim bolNewNP As Boolean
   'Add By Cheng 2002/12/10
   Dim strMsg As String
   Dim strSalesNo As String
   Dim bolSavPdf As Boolean 'Added by Morgan 2018/10/2
 
'Add By Cheng 2002/11/05
On Error GoTo ErrorHandler
SaveDatabase = True
'Add By Cheng 2002/12/09
'若本案已閉卷
m_blnCancelClosed = False
If m_blnClosed Then
    If m_strCloseDate = "" Then
        strMsg = "此案已於 ??年 ??月 ??日 閉卷，是否取消閉卷? "
    Else
        strMsg = "此案已於 " & Left(m_strCloseDate, 4) - 1911 & "年" & Mid(m_strCloseDate, 5, 2) & "月" & Right(m_strCloseDate, 2) & "日閉卷，是否取消閉卷? "
    End If
    If MsgBox(strMsg, vbExclamation + vbYesNo) = vbYes Then
        m_blnCancelClosed = True
    Else
        m_blnCancelClosed = False
    End If
End If

cnnConnection.BeginTrans

   m_PrintCForm = ""
   iStep = 1
   bolNewNP = False
   If txtCaseField(2) <> "" Then 'Added by Morgan 2023/11/20 有下一程序才要更新(否則誤點選時期限會被清除)--秀玲
      For i = 1 To grdDataList.Rows - 1
         If grdDataList.TextMatrix(i, 0) = "ˇ" Then
            bolNewNP = True
            'Modify by Morgan 2006/1/24 加NP01
            strTxt(iStep) = "UPDATE NEXTPROGRESS SET NP08=" & CNULL(TransDate(txtCaseField(4), 2)) & _
               ",NP09=" & CNULL(TransDate(txtCaseField(5), 2)) & " WHERE NP22 = " & grdDataList.TextMatrix(i, 10) & " and np01='" & grdDataList.TextMatrix(i, 7) & "'"
           'Add By Cheng 2002/11/05
           cnnConnection.Execute strTxt(iStep)
            iStep = iStep + 1
         End If
      Next
   End If
   
   '取得機關文號
   strTxt(iStep) = "update caseprogress set cp08=" & CNULL(txtCaseField(1)) & " where cp09='" & cp(9) & "'"
    'Add By Cheng 2002/11/05
    cnnConnection.Execute strTxt(iStep)
   iStep = iStep + 1
   
   'edit by nickc 2007/02/02
   'Dim strDataTemp(1 To T_CP) As String
   Dim strDataTemp() As String
   ReDim strDataTemp(1 To TF_CP) As String
   
   strDataTemp(1) = cp(1)
   strDataTemp(2) = cp(2)
   strDataTemp(3) = cp(3)
   strDataTemp(4) = cp(4)
   
   If intPWhere = 國外_CF Then
      strDataTemp(5) = strSrvDate(1)
      '2008/8/26 modify by sonia 櫃台收文日改存 cp119
      '2008/10/24 MODIFY BY SONIA CP64仍存
      If txtCaseField(12) = "" Then
         strDataTemp(64) = "櫃台收文日：" & Me.lblCaseField(8).Caption
      Else
         strDataTemp(64) = "櫃台收文日：" & Me.lblCaseField(8).Caption & "，" & txtCaseField(12).Text
      End If
      strDataTemp(119) = ChangeTStringToWString(Me.lblCaseField(8).Caption)
      '2008/8/26 end
      m_PrintCForm = "Y"
   ElseIf intPWhere = 國內 Then
      '91.8.22 MODIFY BY SONIA
      'strDataTemp(5) = ChangeTDateStringToTString(lblCaseField(8))
      strDataTemp(5) = lblCaseField(8)
      '91.8.22 END
      'Modify by Morgan 2009/12/14 +代理人來函備註與一般來函區別
      strDataTemp(64) = txtCaseField(12) & ";代理人來函;"
   End If
      
   strDataTemp(6) = txtCaseField(4)
   strDataTemp(7) = txtCaseField(5)
   strDataTemp(9) = 主管機關來函
   strDataTemp(10) = txtCaseField(0)
   strDataTemp(13) = stCP13
   strDataTemp(12) = stCP12
   strDataTemp(14) = txtCaseField(8)
   strDataTemp(26) = txtCaseField(10)
   strDataTemp(48) = txtCaseField(9)
   strDataTemp(43) = cp(9)
   strDataTemp(20) = "N"
   strDataTemp(32) = "N"
   
   'P的來函承辦掛操作人員,直接上發文日
   If cp(1) = "P" Then
      strDataTemp(14) = strUserNum
      strDataTemp(27) = strSrvDate(1)
   End If
   
   '92.6.24 ADD BY SONIA
   If txtCaseField(0) = "1908" Then
      m_PrintCForm = "N"
      strDataTemp(27) = strSrvDate(1)
      strDataTemp(26) = "N"
      strDataTemp(32) = ""
      strDataTemp(20) = ""
      strDataTemp(16) = Val(txtCaseField(11))
      strDataTemp(17) = Val(txtCaseField(11)) - (Val(txtCaseField(13)) * 1000)
      strDataTemp(18) = Val(txtCaseField(13))
   End If
   '92.6.24 END

   strDataTemp(119) = DBDATE(lblCaseField(8)) 'Added by Morgan 2012/4/30 +cp119=櫃檯收文日
   
   strTxt(iStep) = GetCPSQL(strDataTemp(), False)
   'Add By Cheng 2002/11/05
   cnnConnection.Execute strTxt(iStep)
   
   'Added by Morgan 2016/6/6
   If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And cp(1) = "P" And Left(Pub_StrUserSt03, 1) <> "F" Then
      PUB_AddLetterProgress strDataTemp(9), 1, False
   End If
   'end 2016/6/6
   
   '2010/1/20 add by sonia 承辦人為分所人員以系統日的下一個工作天上齊備日
   If m_CP14ST06 <> "1" And strDataTemp(27) = "" Then
      strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & CompWorkDay(2, strSrvDate(1), 0) & " WHERE EP02='" & strDataTemp(9) & "'"
      cnnConnection.Execute strSql
   'Add by Morgan 2010/10/1
   Else
      strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & strDataTemp(9) & "'"
      cnnConnection.Execute strSql
   End If
   '2010/1/20 end
   iStep = iStep + 1
   iNP22 = 1
   m_NewCP09 = strDataTemp(9)
   
   bolNP22 = False
   Dim a As Single
   If txtCaseField(2) <> "" And Not bolNewNP Then
      If txtCaseField(2) = 催審 Or txtCaseField(2) = 提申 Or txtCaseField(2) = 收達 Then
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
        If strDataTemp(1) <> "P" Then
            strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np15,np22) values (" + _
               CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
               "," + CNULL(strDataTemp(4)) + "," + CNULL(txtCaseField(2)) + "," & CNULL(TransDate(txtCaseField(4), 2)) & "," & _
               CNULL(TransDate(txtCaseField(5), 2)) & _
               "," + CNULL(PUB_GetAKindSalesNo(strDataTemp(1), strDataTemp(2), strDataTemp(3), strDataTemp(4))) + "," + CNULL(txtCaseField(12)) + "," & lMax & ")"
        Else
            'Modify By Cheng 2003/04/16
            '是否續辦直接上"Y"
             strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np06,np07,np08,np09,np10,np15,np22) values (" + _
                CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
                "," + CNULL(strDataTemp(4)) + ",'Y'," + CNULL(txtCaseField(2)) + "," & CNULL(TransDate(txtCaseField(4), 2)) & "," & _
                CNULL(TransDate(txtCaseField(5), 2)) & _
                "," + CNULL(PUB_GetAKindSalesNo(strDataTemp(1), strDataTemp(2), strDataTemp(3), strDataTemp(4))) + "," + CNULL(txtCaseField(12)) + "," & lMax & ")"
        End If
      Else
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
        If strDataTemp(1) <> "P" Then
            strSalesNo = PUB_GetAKindSalesNo(strDataTemp(1), strDataTemp(2), strDataTemp(3), strDataTemp(4))
            'Add by Morgan 2008/5/5
            '通知提供前案
            If txtCaseField(0) = "1205" And txtCaseField(2) = "207" And txtCaseField(4) <> "" Then
               'Modify by Morgan 2011/2/16 改若已收文且已發文時仍要新增下一程序(改判斷是否有未發文的)--禧佩
               'strExc(0) = "SELECT CP09, CP27 FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
                  " AND CP10='207' AND CP57 IS NULL"
               strExc(0) = "SELECT CP09, CP27 FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
                  " AND CP10='207' AND CP57 IS NULL AND CP27 IS NULL"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If IsNull(RsTemp("CP27")) Then
                     '更新已收文期限
                     strSql = "update caseprogress set cp06=" & TransDate(txtCaseField(4), 2) & ",cp07=" & TransDate(txtCaseField(5), 2) & _
                        " WHERE cp09='" & RsTemp("cp09") & "'"
                     cnnConnection.Execute strSql, intI
                     '發文日上系統日
                     strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & strDataTemp(9) & "' and cp27 is null"
                     cnnConnection.Execute strSql, intI
                  End If
               Else
                  strExc(0) = "SELECT NP22,NP01 FROM NEXTPROGRESS WHERE " & ChgNextProgress(cp(1) & cp(2) & cp(3) & cp(4)) & _
                     " and np07='207' and np06 is null"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     '更新NP期限
                     strSql = "update nextprogress set np08=" & TransDate(txtCaseField(4), 2) & ",np09=" & TransDate(txtCaseField(5), 2) & _
                        " WHERE np01='" & RsTemp("NP01") & "' and np22=" & RsTemp("NP22")
                     cnnConnection.Execute strSql, intI
                     'Add by Morgan 2009/11/18 更新也要印接洽單
                     NP22(iNP22) = RsTemp("NP22")
                     iNP22 = iNP22 + 1
                     'end 2009/11/18
                  Else
                     '新增NP
                     lMax = ClsLawGetMax
                     bolNP22 = True
                     strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
                        " values ('" & strDataTemp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'" & _
                        ",207," & TransDate(txtCaseField(4), 2) & "," & TransDate(txtCaseField(5), 2) & ",'" & strSalesNo & "'," & lMax & ")"
                  End If
               End If
            'end 2008/5/5
            Else
               lMax = ClsLawGetMax
               bolNP22 = True
               strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np15,np22) values (" + _
                  CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
                  "," + CNULL(strDataTemp(4)) + "," + CNULL(txtCaseField(2)) + "," & CNULL(TransDate(txtCaseField(4), 2)) & "," & _
                  CNULL(TransDate(txtCaseField(5), 2)) & _
                  ",'" & strSalesNo & "'," + CNULL(txtCaseField(12)) + "," & lMax & ")"
            End If
         Else
            lMax = ClsLawGetMax
            bolNP22 = True
            '是否續辦直接上"Y"
            strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np06,np07,np08,np09,np10,np15,np22) values (" + _
               CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
               "," + CNULL(strDataTemp(4)) + ",'Y'," + CNULL(txtCaseField(2)) + "," & CNULL(TransDate(txtCaseField(4), 2)) & "," & _
               CNULL(TransDate(txtCaseField(5), 2)) & _
               "," + CNULL(PUB_GetAKindSalesNo(strDataTemp(1), strDataTemp(2), strDataTemp(3), strDataTemp(4))) + "," + CNULL(txtCaseField(12)) + "," & lMax & ")"
         End If
      End If
      
      If bolNP22 = True Then
         cnnConnection.Execute strTxt(iStep)
         NP22(iNP22) = lMax
         iNP22 = iNP22 + 1
      End If
      
         iStep = iStep + 1
        '新增B類收文
        If strDataTemp(1) = "P" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetAutoNumber("B", strAutoNumber, True, True) Then
            If ClsPDGetAutoNumber("B", strAutoNumber, True, True) Then
                'Modify By Sindy 2010/8/18 比對自動編號年度
                'Ncp(9) = "B" & Right(DBYEAR(ServerDate) - 1911, 2) & strAutoNumber
                Ncp(9) = "B" & CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) & strAutoNumber
                Ncp(1) = strDataTemp(1)
                Ncp(2) = strDataTemp(2)
                Ncp(3) = strDataTemp(3)
                Ncp(4) = strDataTemp(4)
                Ncp(5) = strSrvDate(1)
                Ncp(6) = ChangeTStringToWString(txtCaseField(4).Text)
                Ncp(7) = ChangeTStringToWString(txtCaseField(5).Text)
                Ncp(10) = Me.txtCaseField(2).Text
                Ncp(11) = "90"
                Ncp(13) = stCP13
                Ncp(12) = stCP12
                '承辦人
                Ncp(14) = Me.txtCaseField(8).Text
                '是否算案件數
                cp(26) = Me.txtCaseField(10).Text
                Ncp(43) = strDataTemp(9)
                Ncp(48) = TransDate(txtCaseField(9), 2) 'Add by Morgan 2006/8/2
                If PUB_AddNewCaseProgress(Ncp) = False Then GoTo ErrorHandler
               '2010/1/20 add by sonia 承辦人為分所人員以系統日的下一個工作天上齊備日
               If m_CP14ST06 <> "1" Then
                  strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & CompWorkDay(2, strSrvDate(1), 0) & " WHERE EP02='" & Ncp(9) & "'"
                  cnnConnection.Execute strSql
                  
               'Add by Morgan 2010/10/1
               ElseIf Ncp(48) = "" Then
                  strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & Ncp(9) & "'"
                  cnnConnection.Execute strSql
               End If
               '2010/1/20 end
                'Add by Morgan 2004/2/18
                '若承辦人是王協理且未發文則要發EMail通知
                stCP09 = Ncp(9): stCP14 = Ncp(14)
            Else
                GoTo ErrorHandler
            End If
            'Add by Morgan 2006/6/27
            '國外部收文若有期限則自動內部收文901告知代理人,承辦人固定為78063黃得峻並列印內部收文接洽單
            If Left(stCP12, 1) = "F" Then
               m_901CP09 = AutoNo("B", 6)
               '2008/12/2 modify by sonia 改FMP控管方式
               'm_901CP13 = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
               m_901CP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
               '2008/12/2 END
               m_901CP12 = GetSalesArea(m_901CP13)
               strExc(1) = GetWorkDays(field(1), field(9), "901")
               If strExc(1) = Empty Then strExc(1) = 7
               'Add by Morgan 2008/5/26 若來函期限超過(含)3個月則告代的承辦期限為14天--阮威立
               If Val(strExc(1)) < 14 Then
                  If DBDATE(txtCaseField(5)) >= CompDate(1, 3, strSrvDate(1)) Then
                     strExc(1) = 14
                  End If
               End If
               'Modify by Morgan 2006/8/4 不必抓工作天--郭
               'strExc(0) = CompWorkDay(Val(strExc(1)), strSrvDate(1), 0)
               strExc(0) = CompDate(2, Val(strExc(1)), strSrvDate(1))
               'Modify by Morgan 2008/5/26 78063離職改85030阮威立--郭
               '2008/12/3 MODIFY BY SONIA 依FC代理人國籍抓預設承辦人
               'strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
                  "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
                  "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & "," & strExc(0) & "," & strExc(0) & _
                  ",'" & m_901CP09 & "','" & 告知代理人 & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
                  ",'85030','N','N','N','" & strDataTemp(9) & "'," & strExc(0) & ") "
               'Modified by Morgan 2017/10/11 承辦人　CNULL(PUB_GetFmpCP14(cp)) ->txtCaseField(8)
               strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
                  "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
                  "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & "," & strExc(0) & "," & strExc(0) & _
                  ",'" & m_901CP09 & "','" & 告知代理人 & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
                  "," & CNULL(txtCaseField(8)) & ",'N','N','N','" & strDataTemp(9) & "'," & strExc(0) & ") "
               '2008/12/3 END
               cnnConnection.Execute strSql
            End If
        End If
   End If
   'Added by Lydia 2015/04/10 申請人可在該國有多筆識別番號->call frm880021
'   If txtCaseField(3) <> "" Then
'      strExc(1) = Left(ChangeCustomerL(field(26)), 8)
'      strExc(0) = "select afid03 from applicantforeignid where afid01=" + CNULL(strExc(1)) + _
'         " and afid02=" + CNULL(field(9))
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'      If intI = 0 Then
'         strTxt(iStep) = "insert into applicantforeignid values (" + CNULL(strExc(1)) + "," + CNULL(field(9)) + "," + CNULL(txtCaseField(3)) + ")"
'      Else
'         strTxt(iStep) = "update applicantforeignid set afid03=" + CNULL(txtCaseField(3)) + " where afid01=" + CNULL(strExc(1)) + _
'            " and afid02=" + CNULL(field(9))
'      End If
'        'Add By Cheng 2002/11/05
'        cnnConnection.Execute strTxt(iStep)
'      iStep = iStep + 1
'   End If
   
   'Add By Sindy 2012/3/5 原基本檔未閉卷時,才可更新
   If txtCaseField(6) = "Y" And m_blnClosed = False Then
      'Add By Sindy 2012/3/5 +PA59
      strTxt(iStep) = "UPDATE PATENT SET PA57='Y',PA58=" & strSrvDate(1) & ",PA59='99' WHERE " & ChgPatent(field(1) & field(2) & field(3) & field(4))
        'Add By Cheng 2002/11/05
        cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
   End If
    'Add By Cheng 2002/12/09
    '若要取消閉卷, 則更新基本檔閉卷及其相關欄位為NULL
    If m_blnCancelClosed = True Then
        strSql = "Update Patent Set PA57=Null,PA58=Null,PA59=Null Where  " & ChgPatent(field(1) & field(2) & field(3) & field(4))
        cnnConnection.Execute strSql
    End If
    
    'Modify By Cheng 2002/11/05
'   SaveDatabase = objLawDll.ExecSQL(iStep - 1, strTxt())
    
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", strDataTemp(9), "")
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010603_1", IIf(Pub_StrUserSt03 = "F22", strDataTemp(9), "")
   End If
   '2016/10/7 END
   
   'Added by Morgan 2018/7/18 CFP電子化
   If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) And field(1) = "CFP" Then
      m_strLD18 = strDataTemp(9)
      m_strCP10 = strDataTemp(10)
      strExc(1) = PUB_GetLetterJudgeNew("1", field(1), m_strCP10, field(9), cp(10))
      'Modified by Morgan 2019/2/19 代理人請款沒有附件,預設不通知客戶
      'PUB_AddLetterProgress m_strLD18, 1 + Val(txtFiles), IIf(m_strCP10 = "1908", True, False), strExc(1), False, field(26), m_strCP10, field(75)
      PUB_AddLetterProgress m_strLD18, 1 + Val(txtFiles), IIf(txtCaseField(3) <> "N", True, False), strExc(1), False, field(26), m_strCP10, field(75)
      'end 2019/2/19
      m_bolAddLP = True
      
      If m_PrintCForm = "Y" And txtCaseField(8) <> "" And Left(cp(12), 1) <> "F" Then
         Pub_COrderInform strDataTemp(9)
         bolSavPdf = True
      End If
   End If
   'end 2018/7/18
   
   cnnConnection.CommitTrans
    
   'Add by Morgan 2004/1/13
   '先預設不列印案件回覆單
   m_blnCustReturnSheet = False
   'Modify by Morgan 2009/11/18
   'If SaveDatabase And bolNP22 Then
   If SaveDatabase And iNP22 > 1 Then
      For i = 1 To iNP22 - 1
            'Modify By Cheng 2003/04/16
'         g_PrtForm001.PrintForm NP22(i), cp(1), cp(2), cp(3), cp(4)
        If cp(1) = "P" Then
            g_PrtForm001.PrintForm NP22(i), cp(1), cp(2), cp(3), cp(4), Ncp(9)
        Else
            g_PrtForm001.PrintForm NP22(i), cp(1), cp(2), cp(3), cp(4)
        End If
         'Add By Cheng 2002/08/27
         '若智權人員所屬部門別為"F"開頭時, 再列印一張案件性質為901"告知代理人"的接洽結案單
         If Left("" & GetStaffDepartment(Me.lblCaseField(7).Caption), 1) = "F" Then
            'Modify By Cheng 2003/04/16
'            g_PrtForm001.PrintForm NP22(i), cp(1), cp(2), cp(3), cp(4)
            If cp(1) = "P" Then
               'Modify by Morgan 2006/6/27
                'g_PrtForm001.PrintForm NP22(i), cp(1), cp(2), cp(3), cp(4), Ncp(9)
                g_PrtForm001.PrintForm "", cp(1), cp(2), cp(3), cp(4), m_901CP09
            Else
                g_PrtForm001.PrintForm NP22(i), cp(1), cp(2), cp(3), cp(4)
            End If
         End If
      Next
      'Add by Morgan 2004/1/13
      '設定要列印案件回覆單
      If cp(1) = "CFP" Then m_blnCustReturnSheet = True
   End If
   
   '列印C類接洽記錄單 92.1.28 ADD BY SONIA
   If m_PrintCForm = "Y" Then g_PrtForm001.PrintCFForm strDataTemp(9), , bolSavPdf

'Add By Cheng 2002/11/05
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    SaveDatabase = False
End Function

Private Sub ReadAllData()
Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String

On Error GoTo ErrHnd
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.ReadAllData(frm02010603_2.grdDataList.TextMatrix(frm02010603_2.grdDataList.Row, 0), cp(), field(), intCaseKind, intPWhere) Then
   ReDim cp(TF_CP) As String
   cp(9) = frm02010603_2.grdDataList.TextMatrix(frm02010603_2.grdDataList.row, 0)
   If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
    '智權人員存最近收文A類接洽記錄單的智權人員
    stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
    stCP12 = GetSalesArea(stCP13)
   'Add by Morgan 2010/9/30 新規則承辦期限隔日凌晨算
   If Not PUB_IfSetCP48(cp(9)) Then
      txtCaseField(9).Text = ""
      txtCaseField(9).Enabled = False
   End If
   'end 2010/9/30
    
   If cp(1) = 馬德里案 Then
      lblCaseField(0) = cp(1) + " - " + Left(cp(2), 5) + _
         IIf(Right(cp(2), 1) = "0", "", " - " + Right(cp(2), 1)) + _
         IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
         IIf(cp(4) = "00", "", " - " + cp(4))
   Else
      lblCaseField(0) = MergeString(cp(1), cp(2), cp(3), cp(4))
   End If
   Select Case intPCaseKind
                Case 專利
                           lblCaseField(1) = field(11)
                           lblCaseField(2) = field(26)
                           lblCaseField(9) = field(9)
                Case 商標
                           lblCaseField(1) = field(12)
                           lblCaseField(2) = field(23)
                           lblCaseField(9) = field(10)
                Case Else
                           lblCaseField(1) = field(11)
                           lblCaseField(2) = field(8)
                           lblCaseField(9) = field(9)
   End Select
   ' 90.07.02 modify by louis
   m_PA01 = field(1)
   m_PA09 = field(9)
   lblCaseField(4) = cp(9)
   lblCaseField(6) = cp(10)
   lblCaseField(7) = cp(13)
   txtCaseField(1) = cp(8)
   
   'Add by Morgan 2008/5/2 若點選收文有期限且未逾期時預設到本來函
   If cp(1) = "CFP" And DBDATE(cp(7)) >= strSrvDate(1) Then
      txtCaseField(4) = TransDate(cp(6), 1)
      txtCaseField(5) = TransDate(cp(7), 1)
   End If
   'end 2008/5/2
   
   'Add by Morgan 2006/8/1 國外部收文承辦人預設黃得峻78063 -- 郭
   If cp(1) = "P" Then
      If Left(cp(12), 1) = "F" Then
         '2009/9/23 modify by sonia
         'txtCaseField(8) = "78063"
         'Modified by Morgan 2012/8/20
         'txtCaseField(8) = PUB_GetFMCASECP14(cp(1), cp(2), cp(3), cp(4))
         'Modified by Morgan 2017/10/11 FMP預設承辦人比照FCP(來函性質先用936跑,若需要區分時要改寫到輸入來函的事件,且要增加來函性質到函數內)
         'txtCaseField(8) = PUB_GetFmpCP14(cp)
         txtCaseField(8) = PUB_GetFCPPromoterNo(cp(9), "936", cp(14))
         'end 2017/10/11
         '2009/9/23 end
      Else
         'Add by Morgan 2006/8/2 點申請程序時要檢查若有國內案帶國內案的承辦人 --郭
         If InStr(CaseMapIn, cp(10)) > 0 Then
            txtCaseField(8) = PUB_GetInCaseCP14(cp(1), cp(2), cp(3), cp(4))
         End If
         If txtCaseField(8) = "" Then
            txtCaseField(8) = cp(14)
         End If
      End If
   Else
      txtCaseField(8) = cp(14)
   End If
   'add by sonia 2024/7/15 A7010柯昱安調離也要改為李柏翰經理99050
   If GetStaffDepartment(txtCaseField(8)) >= "P10" And GetStaffDepartment(txtCaseField(8)) <= "P11" Then
   Else
      txtCaseField(8) = "99050"
   End If
   'end 2024/7/15
   CheckKeyIn 8
   lblCaseField(8) = frm02010603_1.txtCaseCode(3)
   lblCaseField(5) = ChangeTStringToTDateString(cp(5))
   SetNameToCombo cboCaseName, field(5), field(6), field(7)
   GetCaseDeadLineData grdDataList, intLastRow, cp(1), cp(2), cp(3), cp(4), True
   
   'Added by Lydia 2015/04/10 申請人可在該國有多筆識別番號
    CmdAFID03(0).Enabled = True

    For intI = 1 To 4
        If Len(field(26 + intI)) > 0 Then
           CmdAFID03(intI).Visible = True
        Else
           CmdAFID03(intI).Visible = False
        End If
    Next intI
      
    'Add By Cheng 2002/12/09
    If "" & field(57) = "Y" Then
        m_blnClosed = True
    Else
        m_blnClosed = False
    End If
    m_strCloseDate = "" & field(58)

End If
Screen.MousePointer = varSaveCursor
Exit Sub
ErrHnd:
ErrorMsg
Screen.MousePointer = varSaveCursor
End Sub

Private Sub Form_Activate()
'Add by Morgan 2005/5/18 控制只執行一次
If m_bolActive = True Then Exit Sub
m_bolActive = True

blnOKtoShow = True
ReadAllData
txtCaseField(0).SetFocus
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim Ncp(TF_CP) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   
   txtCaseField(3).Enabled = False 'Added by Morgan 2019/2/19
   txtCaseField(3).Text = "N" 'Added by Morgan by Morgan 2022/1/7 目前只有CFP案有定稿
   txtCaseField(7).Enabled = False
   If intPCaseKind = 專利 And intPWhere = 國外_CF Then
      'Added by Lydia 2015/04/10
      'lblID.Visible = True
      'txtCaseField(3).Visible = True
      txtCaseField(7).Enabled = True
      txtCaseField(3).Enabled = True 'Added by Morgan 2019/2/19
      txtCaseField(3).Text = "" 'Added by Morgan by Morgan 2022/1/7
   End If
   
   'If intPWhere <> 國外_CF Then
      txtCaseField(4).MaxLength = 7
      txtCaseField(5).MaxLength = 7
      txtCaseField(9).MaxLength = 7
   'End If
   txtCaseField_Change 2
   Me.Caption = frm02010603_2.Caption
   'Add By Cheng 2002/07/24
   If intPCaseKind = 專利 And intPWhere = 國外_CF Then
      Label3.Caption = "櫃台收文日："
   End If
   'Add by Amy 2014/09/17 承辦人期限隱藏
   Label19.Visible = False
   txtCaseField(9).Enabled = False
   txtCaseField(9).Visible = False
   'end 2014/09/17
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub

'Add By Sindy 2016/10/11
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2018/10/2
   'Add By Sindy 2016/10/7
   If Me.m_strIR01 <> "" Then
      Unload frm02010603_1
      Unload frm02010603_2
      'Add By Sindy 2016/10/11
      If Not m_PrevForm Is Nothing Then
         Call m_PrevForm.GoNext
      End If
      '2016/10/11 END
   Else
   '2016/10/7 END
      If intLeaveKind = 1 Then
         frm02010603_2.Show
      Else
         Unload frm02010603_2
         If intLeaveKind = 2 Then
            Unload frm02010603_1
         End If
      End If
   End If

   'Add By Sindy 2016/10/11
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If

   'Add By Cheng 2002/07/18
   Set frm02010603_3 = Nothing
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String, strCusTemp As String

Select Case Index
             Case 2
                        strCusTemp = lblCaseField(Index)
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCustomer(strCusTemp, strTemp) Then
                        If ClsPDGetCustomer(strCusTemp, strTemp) Then
                           lblCaseField(Index) = strCusTemp
                           lblAgent.Caption = strTemp
                           'Added by Lydia 2015/04/10
'                           If lblID.Visible Then
'                              'edit by nickc 2007/02/02 不用 dll 了
'                              'If objPublicData.GetAgentID(strCusTemp, field(9), strTemp) Then
'                              If ClsPDGetAgentID(strCusTemp, field(9), strTemp) Then
'                                 txtCaseField(3) = strTemp
'                              End If
'                           End If
                        End If
             Case 6
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCaseProperty(cp(1), lblCaseField(Index), strTemp) Then
                        If ClsPDGetCaseProperty(cp(1), lblCaseField(Index), strTemp) Then
                           lblCaseProperty = strTemp
                        End If
             Case 7
                        ' 91.07.23 modify by sonia
                        'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
                        '   lblSales = strTemp
                        'End If
                        lblSales = GetStaffName(lblCaseField(Index), True)
             Case 9
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetNation(lblCaseField(Index), strTemp) Then
                        If ClsPDGetNation(lblCaseField(Index), strTemp) Then
                           lblNation.Caption = strTemp
                        End If
End Select
End Sub
Private Sub grdDataList_GotFocus()
GridGotFocus grdDataList
End Sub
Private Sub grdDataList_LostFocus()
GridLostFocus grdDataList
End Sub
Private Sub grdDataList_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then GrdDataList_Click
End Sub
Private Sub GrdDataList_Click()
Dim i As Integer

If grdDataList.TextMatrix(grdDataList.row, 0) = "ˇ" Then
   grdDataList.TextMatrix(grdDataList.row, 0) = ""
Else
   For i = 0 To grdDataList.Rows - 1
          grdDataList.TextMatrix(i, 0) = ""
   Next
   grdDataList.TextMatrix(grdDataList.row, 0) = "ˇ"
End If
End Sub
Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      ShowBar grdDataList, intLastRow, 6
      blnOKtoShow = True
   End If
End If
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 6, 8, 10, 11
         KeyAscii = UpperCase(KeyAscii)
      'Add By Cheng 2002/07/22
      Case 7
         '若為是否修改通知函內容
         If Me.txtCaseField(Index).MaxLength = 1 Then
            KeyAscii = UpperCase(KeyAscii)
            If KeyAscii <> 8 And KeyAscii <> 89 Then
               KeyAscii = 0
            End If
         End If
      'Added by Morgan 2019/2/19
      Case 3
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
         End If
   End Select
End Sub
Private Sub txtCaseField_Change(Index As Integer)
   Dim strDays As String
   
   Select Case Index
      Case 2
         lblNextCaseProperty = ""
         'Add by Morgan 2008/5/26 從 Validate 移過來,否則存檔時期限會重設
         If Len(txtCaseField(Index)) = 3 Then
            'Add by Morgan 2006/8/2
            If cp(1) = "P" Then
               strDays = Empty
               strDays = GetWorkDays(m_PA01, m_PA09, txtCaseField(Index))  '2010/1/20 因為下一程序自動內部收文,因要掛下一程序內部收文後之期限故以下一程序抓工作天
               If strDays <> Empty Then
                  'Modify by Morgan 2006/8/4 不必抓工作天--郭
                  'Remark by Morgan 2007/10/16 來函的承辦期限為何用下一程序的性質計算?--郭也不記得了，故暫時保留
                  '2010/1/20 因為下一程序自動內部收文,因要掛下一程序內部收文後之期限故以下一程序抓工作天
                  'txtCaseField(9) = TransDate(CompWorkDay(Val(strDays), strSrvDate(1), 0), 1)
                  '2010/1/20 modify by sonia 承辦人為北所人員以系統日計算承辦期限,分所人員以系統日的下一個工作天計算
                  'txtCaseField(9) = TransDate(CompDate(2, Val(strDays), strSrvDate(1)), 1)
                  If m_CP14ST06 <> "1" Then
                     txtCaseField(4) = TransDate(CompDate(2, Val(strDays), CompWorkDay(2, strSrvDate(1), 0)), 1)
                  Else
                     txtCaseField(4) = TransDate(CompDate(2, Val(strDays), strSrvDate(1)), 1)
                  End If
                  '2010/1/20 end
                  txtCaseField(5) = txtCaseField(4)
                  If txtCaseField(9).Enabled Then 'Add by Morgan 2010/9/30 配合新規則判斷
                     txtCaseField(9) = txtCaseField(4)
                  End If
               End If
            End If
         End If
   End Select
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
   Dim strDays As String
      
   Select Case Index
      Case 12
         cmdOK(0).Default = True
         cmdOK(0).CausesValidation = True
   End Select
   If CheckKeyIn(Index) = -1 Then
      Cancel = True
   End If
   '90.07.01 modify by louis
   Select Case Index
      Case 0:
         '2010/1/20 modify by soniacp(1) = "P"
         'If txtCaseField(9) = "" Then
         If txtCaseField(9).Enabled And (txtCaseField(9) = "" Or cp(1) <> "P") Then
            '2010/1/20 modify by sonia 承辦人為北所人員以系統日計算承辦期限,分所人員以系統日的下一個工作天計算
            'txtCaseField(9) = TransDate(Pub_GetHandleDay(m_PA01, m_PA09, txtCaseField(0)), 1)
            If m_CP14ST06 <> "1" Then
               txtCaseField(9) = TransDate(Pub_GetHandleDay(m_PA01, m_PA09, txtCaseField(0), CompWorkDay(2, strSrvDate(1), 0), IIf(txtCaseField(4) = "", "", TransDate(txtCaseField(4), 2))), 1)
            Else
               txtCaseField(9) = TransDate(Pub_GetHandleDay(m_PA01, m_PA09, txtCaseField(0), , IIf(txtCaseField(4) = "", "", TransDate(txtCaseField(4), 2))), 1)
            End If
            '2010/1/20 end
         End If
         
         'Add by Morgan 2008/5/2
         If field(1) = "CFP" Then
            Select Case txtCaseField(Index)
               Case "1902" '其他來函
                  txtCaseField(10) = "N" '預設不算案件數
                  
               Case "1205"
                  '馬來西亞發明新型的 通知提供前案(1205) 要帶下一程序 提供前案(207) 並計算期限
                  If field(9) = "018" And (field(8) = "1" Or field(8) = "2") Then
                     If txtCaseField(2) <> "207" Then
                        txtCaseField(2) = "207"
                        Set207DueDate
                     End If
                  End If
               
               'Added by Morgan 2022/1/11
               '代理人請款預設不出定稿，若改要出定稿時預設要修改--禧佩
               Case "1908"
                  txtCaseField(3).Text = "N"
                  txtCaseField(7).Text = "Y"
               'end 2022/1/11
            End Select
         End If
   End Select
   If Cancel Then
      TextInverse txtCaseField(Index)
      txtCaseField(Index).SetFocus 'Added by Morgan 2021/12/10 Form2.0要再設駐點否則游標會不見
   End If
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
   Dim strTemp As String, strTemp1 As String, bolIsChina As Boolean, strCusTemp As String

   CheckKeyIn = -1
   Select Case intIndex
             Case 0 '來函性質
                        'If Len(Me.txtCaseField(0).Text) > 0 Then 'Removed by Morgan 2019/2/19
                           If Len(Me.txtCaseField(0).Text) <> 4 Then
                              MsgBox "來函性質欄位值必須為四碼 !", vbExclamation
                              Exit Function
                           End If
                        'End If
                        '2009/10/22 ADD BY SONIA
                        'Modified by Morgan 2019/1/31 +1225 --郭
                        If txtCaseField(0) <> "" And InStr("1001,1002,1006,1225,1503", txtCaseField(0)) > 0 Then
                           MsgBox "此來函性質不可由本畫面輸入資料", vbExclamation
                           Exit Function
                        End If
                        '2009/10/22 END
      
                        If lblCaseField(9) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCaseProperty(cp(1), txtCaseField(intIndex), strTemp, bolIsChina) Then
                        If ClsPDGetCaseProperty(cp(1), txtCaseField(intIndex), strTemp, bolIsChina) Then
                           lblProperty = strTemp
                           CheckKeyIn = 1
                        End If
                        '92.8.6 ADD BY SONIA
                        'Modify by Morgan 2005/7/28 加"代理人請款"1908(承辦人會有無法更改的問題)--甄妮
                        'If Me.txtCaseField(0).Text = 檢索報告 Then
                        If Me.txtCaseField(0).Text = 檢索報告 Or Me.txtCaseField(0).Text = "1908" Then
                           Me.txtCaseField(8).Text = strUserNum
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetStaff(strUserNum, strTemp) Then
                           If ClsPDGetStaff(strUserNum, strTemp) Then
                              lblPromoter = strTemp
                           End If
                        End If
                        
                        '92.8.6 END
             Case 2 '下一程序
                       If txtCaseField(intIndex) = "" Then
                          txtCaseField(4) = ""
                          txtCaseField(5) = ""
                          CheckKeyIn = 1
                       Else
                           'Add By Cheng 2002/01/04
                           If Len(Me.txtCaseField(2).Text) <> 3 Then
                              MsgBox "下一程序欄位值必須為三碼 !", vbExclamation
                              Exit Function
                           End If
                          
                          If lblCaseField(9) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
'                          txtCaseField(9) = ""
                          'edit by nickc 2007/02/02 不用 dll 了
                          'If objPublicData.GetCaseProperty(cp(1), txtCaseField(2), strTemp, bolIsChina) Then
                          If ClsPDGetCaseProperty(cp(1), txtCaseField(2), strTemp, bolIsChina) Then
                              lblNextCaseProperty = strTemp
'                              If objPublicData.GetCaseWorkDays(cp(1), lblCaseField(9), txtCaseField(intIndex), strTemp) Then
'                                 If strTemp <> "" Then
'                                    strTemp1 = DateAdd("D", Val(strTemp), ChangeTStringToWDateString(Replace(lblCaseField(8), "/", "")))
'                                    txtCaseField(9) = ChangeWDateStringToTString(strTemp1)
'                                 Else
'                                    txtCaseField(9) = ""
'                                 End If
                                 CheckKeyIn = 1
'                              End If
                           End If
                        End If
             Case 4 '本所期限
                        If txtCaseField(2) = "" Then
                              CheckKeyIn = 1
                        Else
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              CheckKeyIn = 1
                           End If
                        End If
                        'Add By Cheng 2002/03/11
                        '若本所期限有輸入, 則不可小於系統日期
                        With Me.txtCaseField(4)
                           If .Text <> "" Then
                                'Modify By Cheng 2003/12/08
'                              If Val(.Text) + 19110000 < ServerDate Then
                                '若本所期限小於系統日
                                If Val(.Text) + 19110000 < strSrvDate(1) Then
                                    MsgBox "本所期限不可小於系統日期!!!", vbExclamation
                                    CheckKeyIn = -1
                                '若本所期限非工作天則直接調整至最近的工作天
                                Else
                                    .Text = TransDate(PUB_GetWorkDay1(.Text, True), 1)
                                End If
                           End If
                        End With
             Case 5 '法定期限
                        If txtCaseField(2) = "" Then
                              CheckKeyIn = 1
                        Else
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              'Modify by Morgan 2010/8/18 百年蟲
                              'If txtCaseField(4) <= txtCaseField(5) Then
                              If Val(txtCaseField(4)) <= Val(txtCaseField(5)) Then
                                 CheckKeyIn = 1
                              Else
                                 ShowMsg MsgText(1033)
                              End If
                           End If
                        End If
             Case 6
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "Y" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9174)
                        End If
            Case 8
                        m_CP14ST06 = "1" '2010/1/20 add by sonia
                        If txtCaseField(intIndex) = "" Then
                           CheckKeyIn = 1
                        'edit by nickc 2007/02/02 不用 dll 了
                        'ElseIf objPublicData.GetStaff(txtCaseField(intIndex), strTemp) Then
                        ElseIf ClsPDGetStaff(txtCaseField(intIndex), strTemp) Then
                           lblPromoter = strTemp
                           CheckKeyIn = 1
                           m_CP14ST06 = PUB_GetST06(txtCaseField(intIndex))  '2010/1/20 add by sonia
                           '92.5.8 ADD BY SONIA
                           strExc(0) = "SELECT ST03 FROM STAFF WHERE ST01='" & txtCaseField(intIndex) & "'"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              If Not IsNull(RsTemp.Fields("ST03")) And RsTemp.Fields("ST03") = "P12" Then
                                 txtCaseField(10) = "N"
                              End If
                           End If
                        End If
                        '2010/1/20 add by sonia 重新依承辦人所別以系統日或下一個工作天計算承辦期限
                        If txtCaseField(9).Enabled And (txtCaseField(9) = "" Or cp(1) <> "P") Then
                           If m_CP14ST06 <> "1" Then
                              txtCaseField(9) = TransDate(Pub_GetHandleDay(m_PA01, m_PA09, txtCaseField(0), CompWorkDay(2, strSrvDate(1), 0), IIf(txtCaseField(4) = "", "", TransDate(txtCaseField(4), 2))), 1)
                           Else
                              txtCaseField(9) = TransDate(Pub_GetHandleDay(m_PA01, m_PA09, txtCaseField(0), , IIf(txtCaseField(4) = "", "", TransDate(txtCaseField(4), 2))), 1)
                           End If
                        End If
                       '2010/1/20 end
            Case 9 '承辦期限
                        '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
                        If Len(Me.txtCaseField(4).Text) > 0 And Len(Me.txtCaseField(9).Text) > 0 Then
                           If Val(Me.txtCaseField(4).Text) < Val(Me.txtCaseField(9).Text) Then
                              MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        
                        If txtCaseField(intIndex) = "" Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetCaseWorkDays(cp(1), lblCaseField(9), txtCaseField(2), strTemp) Then
                           'Modify by Morgan 2007/10/16 工作天函數統一
                           'If ClsPDGetCaseWorkDays(cp(1), lblCaseField(9), txtCaseField(2), strTemp) Then
                           '2010/1/20 MODIFY BY SONIA 應以來函性質判斷有無工作天
                           'strTemp = GetWorkDays(cp(1), lblCaseField(9), txtCaseField(2))
                           strTemp = GetWorkDays(cp(1), lblCaseField(9), txtCaseField(0))
                           '2010/1/20 END
                           'end 2007/10/16
                              If strTemp <> "" And txtCaseField(2) <> "" Then
                                 ShowMsg MsgText(1049)
                              Else
                                 CheckKeyIn = 1
                              End If
                           'End If
                        Else
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                               CheckKeyIn = 1
                           End If
                        End If
             Case 10
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case Else
                        CheckKeyIn = 1
   End Select
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
   txtCaseField(Index).SelStart = 0
   txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
   Select Case Index
      Case 12
         cmdOK(0).Default = False
         cmdOK(0).CausesValidation = False
   End Select
End Sub

Private Function TxtValidate() As Boolean
   Dim objTxt As Object
   Dim ii As Integer
   Dim Cancel As Boolean
   
   TxtValidate = False
   
   'Added by Morgan 2021/12/10 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/9
   
   For Each objTxt In Me.txtCaseField
      If objTxt.Enabled = True Then
         Cancel = False
         txtCaseField_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   'Add by Morgan 2008/5/26
   If txtCaseField(2) <> "" Then
      If txtCaseField(4) = "" Then
         MsgBox "本所期限不可空白！"
         txtCaseField(4).SetFocus
         Exit Function
      End If
      If txtCaseField(5) = "" Then
         MsgBox "法定期限不可空白！"
         txtCaseField(5).SetFocus
         Exit Function
      End If
      If txtCaseField(9) = "" And txtCaseField(9).Enabled Then
         MsgBox "承辦期限不可空白！"
         txtCaseField(9).SetFocus
         Exit Function
      End If
   End If
   'end 2008/5/26
   
   'Added by Morgan 2018/7/19 CFP電子化
   If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) And field(1) = "CFP" Then
      If txtFiles = "" Then
         MsgBox "請輸入附件案檔案數量!!", vbExclamation
         txtFiles.SetFocus
         Exit Function
      End If
   End If
   'end 2018/7/19

   TxtValidate = True
End Function
'Add by Morgan 2008/5/5
'計算並設定馬來西亞提供前案期限(同實體審查)
Private Sub Set207DueDate()
   Dim strStartDate As String, strTemp As String, strTemp1 As String
   Dim stDate(3) As String
'Modify by Morgan 2009/11/18 改呼叫公用函式以免不一致
'   Dim strDates(0 To 3) As String, iMonthAdd As Integer
'   If field(10) <> "" Then
'      strStartDate = PUB_GetFirstPriDate(field)
'      If strStartDate = "" Then
'         strStartDate = field(10)
'      End If
'      If ClsPDGetNationTaxEx(Val(field(8)) + 3, field(9), strTemp, strTemp1, , , False) = 0 Then
'         If Val(strTemp) = 申請日 And strStartDate <> "" Then
'            iMonthAdd = Val(strTemp1)
'            strTemp = CompDate(1, iMonthAdd, strStartDate)
'            strDates(1) = cp(1)
'            strDates(2) = field(9)
'            strDates(3) = TransDate(strTemp, 2)
'            GetCtrlDT strDates
'            strTemp1 = strDates(0)
'            txtCaseField(4) = TransDate(PUB_GetWorkDay1(strTemp1, True), 1)
'            txtCaseField(5) = TransDate(strTemp, 1)
'         End If
'      End If
'   End If
   strStartDate = PUB_GetFirstPriDate(field)
   PUB_GetExamDate cp(1), cp(2), cp(3), cp(4), cp(9), strStartDate, strTemp1, strTemp
   If strTemp1 <> "" Then
      'Add by Morgan 2010/6/14 改實審期限-6個月
      'strTemp = CompDate(1, -6, strTemp) Remove by Morgan 2011/4/11 禧佩
      stDate(0) = ""
      stDate(1) = field(1)
      stDate(2) = field(9)
      stDate(3) = strTemp
      GetCtrlDT stDate
      strTemp1 = PUB_GetWorkDay1(stDate(0), True)
      'end 2010/6/14
      txtCaseField(4) = TransDate(strTemp1, 1)
      txtCaseField(5) = TransDate(strTemp, 1)
   End If
'end 2009/11/18
End Sub
'Added by Lydia 2015/04/10 呼叫~申請人國外ID資料維護
Private Sub CmdAFID03_Click(Index As Integer)

   Set frm880021.m_PrevF = Me
   frm880021.strXNo = field(26 + Index)
   frm880021.strXNation = m_PA09
   frm880021.lblTitle.Caption = "申請人" & str(Index + 1)
   frm880021.lblCust.Caption = lblAgent.Caption '客戶名稱
   frm880021.lblNation.Caption = lblNation.Caption  '國名
   frm880021.Show
End Sub
'end 2015/04/10

