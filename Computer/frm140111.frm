VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140111 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標日文資料維護作業"
   ClientHeight    =   6390
   ClientLeft      =   570
   ClientTop       =   975
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8955
   Begin VB.CommandButton Command2 
      Caption         =   "商品及服務(&I)"
      Height          =   375
      Left            =   4560
      TabIndex        =   53
      Top             =   0
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消(&C)"
      Height          =   375
      Index           =   3
      Left            =   6960
      TabIndex        =   28
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   29
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Height          =   375
      Index           =   0
      Left            =   6000
      TabIndex        =   27
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   350
      Index           =   2
      Left            =   3330
      TabIndex        =   4
      Top             =   30
      Width           =   800
   End
   Begin MSForms.TextBox textTM131 
      Height          =   600
      Left            =   1440
      TabIndex        =   54
      Top             =   1020
      Width           =   7455
      VariousPropertyBits=   -1467989989
      MaxLength       =   140
      ScrollBars      =   2
      Size            =   "13150;1058"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   1680
      Width           =   1092
      VariousPropertyBits=   671105055
      MaxLength       =   8
      Size            =   "1926;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM45 
      Height          =   300
      Left            =   1440
      TabIndex        =   8
      Top             =   1986
      Width           =   2772
      VariousPropertyBits=   671105051
      MaxLength       =   50
      Size            =   "4890;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM53 
      Height          =   300
      Left            =   7080
      TabIndex        =   7
      Top             =   1986
      Visible         =   0   'False
      Width           =   372
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "656;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   600
      Left            =   1440
      TabIndex        =   5
      Top             =   390
      Width           =   7455
      VariousPropertyBits=   -1467989989
      MaxLength       =   140
      ScrollBars      =   2
      Size            =   "13150;1058"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM52 
      Height          =   300
      Left            =   5580
      TabIndex        =   18
      Top             =   4740
      Width           =   3345
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "5900;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM49 
      Height          =   300
      Left            =   1110
      TabIndex        =   17
      Top             =   4740
      Width           =   3345
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "5900;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM96 
      Height          =   300
      Left            =   1110
      TabIndex        =   19
      Top             =   5046
      Width           =   3345
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "5900;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM99 
      Height          =   300
      Left            =   5580
      TabIndex        =   20
      Top             =   5046
      Width           =   3345
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "5900;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM102 
      Height          =   300
      Left            =   1110
      TabIndex        =   21
      Top             =   5352
      Width           =   3345
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "5900;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM105 
      Height          =   300
      Left            =   5580
      TabIndex        =   22
      Top             =   5352
      Width           =   3345
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "5900;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM108 
      Height          =   300
      Left            =   1110
      TabIndex        =   23
      Top             =   5658
      Width           =   3345
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "5900;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM111 
      Height          =   300
      Left            =   5580
      TabIndex        =   24
      Top             =   5658
      Width           =   3345
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "5900;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM114 
      Height          =   300
      Left            =   1110
      TabIndex        =   25
      Top             =   5970
      Width           =   3345
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "5900;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM117 
      Height          =   300
      Left            =   5580
      TabIndex        =   26
      Top             =   5970
      Width           =   3345
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "5900;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM76 
      Height          =   300
      Left            =   1440
      TabIndex        =   16
      Top             =   4434
      Width           =   7455
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "13150;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM43 
      Height          =   300
      Left            =   1440
      TabIndex        =   15
      Top             =   4128
      Width           =   7455
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "13150;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM40 
      Height          =   300
      Left            =   1440
      TabIndex        =   14
      Top             =   3822
      Width           =   7455
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "13150;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM26 
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   2292
      Width           =   7455
      VariousPropertyBits=   671105051
      MaxLength       =   70
      Size            =   "13150;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM90 
      Height          =   300
      Left            =   1440
      TabIndex        =   10
      Top             =   2598
      Width           =   7455
      VariousPropertyBits=   671105051
      MaxLength       =   70
      Size            =   "13150;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM91 
      Height          =   300
      Left            =   1440
      TabIndex        =   11
      Top             =   2904
      Width           =   7455
      VariousPropertyBits=   671105051
      MaxLength       =   70
      Size            =   "13150;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM92 
      Height          =   300
      Left            =   1440
      TabIndex        =   12
      Top             =   3210
      Width           =   7455
      VariousPropertyBits=   671105051
      MaxLength       =   70
      Size            =   "13150;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM93 
      Height          =   300
      Left            =   1440
      TabIndex        =   13
      Top             =   3516
      Width           =   7455
      VariousPropertyBits=   671105051
      MaxLength       =   70
      Size            =   "13150;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM03 
      Height          =   300
      Left            =   2640
      TabIndex        =   2
      Top             =   75
      Width           =   255
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "450;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM04 
      Height          =   300
      Left            =   2880
      TabIndex        =   3
      Top             =   75
      Width           =   300
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "529;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM02 
      Height          =   300
      Left            =   1920
      TabIndex        =   1
      Top             =   75
      Width           =   735
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1296;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM01 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Top             =   75
      Width           =   495
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "873;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44_2 
      Height          =   285
      Left            =   2580
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   1688
      Width           =   6135
      VariousPropertyBits=   671105055
      BackColor       =   -2147483633
      MaxLength       =   20
      Size            =   "10821;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "定稿商標名稱 :"
      Height          =   240
      Index           =   1
      Left            =   195
      TabIndex        =   55
      Top             =   1110
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "FC代理人 :"
      Height          =   270
      Index           =   65
      Left            =   195
      TabIndex        =   52
      Top             =   1695
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   180
      Index           =   66
      Left            =   195
      TabIndex        =   51
      Top             =   2046
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "定稿語文:                        ( 1:中 2:英 3:日 )"
      Height          =   180
      Index           =   18
      Left            =   5835
      TabIndex        =   50
      Top             =   2046
      Visible         =   0   'False
      Width           =   3090
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件名稱 :"
      Height          =   180
      Index           =   4
      Left            =   195
      TabIndex        =   49
      Top             =   450
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人2(日) :"
      Height          =   180
      Index           =   82
      Left            =   4530
      TabIndex        =   48
      Top             =   4800
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人1(日) :"
      Height          =   180
      Index           =   79
      Left            =   60
      TabIndex        =   47
      Top             =   4800
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人3(日) :"
      Height          =   180
      Index           =   85
      Left            =   60
      TabIndex        =   46
      Top             =   5106
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人4(日) :"
      Height          =   180
      Index           =   88
      Left            =   4530
      TabIndex        =   45
      Top             =   5106
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人5(日) :"
      Height          =   180
      Index           =   91
      Left            =   60
      TabIndex        =   44
      Top             =   5412
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人6(日) :"
      Height          =   180
      Index           =   94
      Left            =   4530
      TabIndex        =   43
      Top             =   5412
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人7(日) :"
      Height          =   180
      Index           =   97
      Left            =   60
      TabIndex        =   42
      Top             =   5718
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人8(日) :"
      Height          =   180
      Index           =   100
      Left            =   4530
      TabIndex        =   41
      Top             =   5718
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人9(日) :"
      Height          =   180
      Index           =   103
      Left            =   60
      TabIndex        =   40
      Top             =   6030
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人10(日) :"
      Height          =   180
      Index           =   106
      Left            =   4530
      TabIndex        =   39
      Top             =   6030
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人部門(日) :"
      Height          =   180
      Index           =   75
      Left            =   105
      TabIndex        =   38
      Top             =   4494
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "聯絡人2(日) :"
      Height          =   180
      Index           =   74
      Left            =   195
      TabIndex        =   37
      Top             =   4188
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "聯絡人1(日) :"
      Height          =   180
      Index           =   71
      Left            =   195
      TabIndex        =   36
      Top             =   3882
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請地址1(日) :"
      Height          =   180
      Index           =   52
      Left            =   195
      TabIndex        =   35
      Top             =   2352
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請地址2(日) :"
      Height          =   180
      Index           =   55
      Left            =   195
      TabIndex        =   34
      Top             =   2658
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請地址3(日) :"
      Height          =   180
      Index           =   58
      Left            =   195
      TabIndex        =   33
      Top             =   2964
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請地址4(日) :"
      Height          =   180
      Index           =   61
      Left            =   195
      TabIndex        =   32
      Top             =   3270
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請地址5(日) :"
      Height          =   180
      Index           =   64
      Left            =   195
      TabIndex        =   31
      Top             =   3576
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   195
      TabIndex        =   30
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frm140111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/13 改成Form2.0 ; 全部TextBox
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_TM09 As String 'Add By Sindy 2012/11/14
Public ChkTG As Boolean '檢查是否已經有商品及服務


'Add By Sindy 2012/11/14
Private Sub Command2_Click()
   If textTM01 = "" Or textTM02 = "" Or textTM03 = "" Or textTM04 = "" Then Exit Sub
   
   frm03010303_04.Hide
   Set frm03010303_04.UpForm = Me
   frm03010303_04.TGKey = textTM01 & "-" & textTM02 & "-" & textTM03 & "-" & textTM04
   frm03010303_04.AllClass = m_TM09
   frm03010303_04.cmdOK(2).Visible = True
   
'   If m_EditMode <> 1 And m_EditMode <> 2 Then
'       frm03010303_04.Label2.Visible = False
'       frm03010303_04.CmdOK(0).Visible = False
'       frm03010303_04.CmdOK(2).Visible = False
'       frm03010303_04.cmd.Visible = False
'       frm03010303_04.cmd2.Visible = False
'       frm03010303_04.txt2(0).Visible = False
'       frm03010303_04.txt2(1).Visible = False
'       frm03010303_04.txt2(2).Visible = False
'       frm03010303_04.txt2(3).Visible = False
'       frm03010303_04.Line1.Visible = False
'   End If
   If m_TM09 <> "" Then '有商品類別才可進入 T-113511團體標章
      Me.Hide
      frm03010303_04.QueryData
      frm03010303_04.Show vbModal
   Else
      MsgBox ("無商品類別，不可使用此按鈕 !")
   End If
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bUpdate = IsUserHasRightOfFunction("frm140111", strEdit, False)
   
   MoveFormToCenter Me
   CmdLock 1
End Sub

Private Sub CmdLock(TF As Integer)
   Select Case TF
      Case 0
         Command1(2).Enabled = False
         Command1(0).Enabled = True
         Command1(3).Enabled = True
         Command2.Enabled = True 'Add By Sindy 2012/11/14
         textTM01.Locked = True
         textTM02.Locked = True
         textTM03.Locked = True
         textTM04.Locked = True
      Case 1
         Command1(2).Enabled = True
         Command1(0).Enabled = False
         Command1(3).Enabled = False
         Command2.Enabled = False 'Add By Sindy 2012/11/14
         textTM01.Locked = False
         textTM02.Locked = False
         textTM03.Locked = False
         textTM04.Locked = False
   End Select
End Sub

Private Sub ClearAll(bClearPk As Boolean)
   If bClearPk = True Then
      textTM01 = Empty
      textTM02 = Empty
      textTM03 = Empty
      textTM04 = Empty
   End If
   textTM05 = Empty
   textTM131 = Empty 'Add By Sindy 2015/7/13
   textTM26 = Empty
   textTM40 = Empty
   textTM43 = Empty
   textTM49 = Empty
   textTM52 = Empty
   textTM76 = Empty
   textTM90 = Empty
   textTM91 = Empty
   textTM92 = Empty
   textTM93 = Empty
   textTM96 = Empty
   textTM99 = Empty
   textTM102 = Empty
   textTM105 = Empty
   TextTM108 = Empty
   TextTM111 = Empty
   TextTM114 = Empty
   TextTM117 = Empty
   'Add By Sindy 2012/4/16
   textTM44 = Empty
   textTM44_2 = Empty
   textTM45 = Empty
   textTM53 = Empty
   '2012/4/16 End
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo ErrHand
   Select Case Index
      Case 0 '確定
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         On Error GoTo ErrorHandler
         cnnConnection.BeginTrans
         'Modify By Sindy 2012/4/16 +TM45,TM53
         'Modify By Sindy 2015/7/14 +TM131
         'Modified by Lydai 2018/12/28 -TM53(定稿語文) ;刪外商\檔案維護\商標日文資料維護作業之定稿語文欄by 阿蓮
         strExc(1) = "UPDATE Trademark " & _
                                    "SET TM05='" & ChgSQL(textTM05) & "'," & _
                                            "TM26='" & ChgSQL(textTM26) & "'," & _
                                            "TM40='" & ChgSQL(textTM40) & "'," & _
                                            "TM43='" & ChgSQL(textTM43) & "'," & _
                                            "TM49='" & ChgSQL(textTM49) & "'," & _
                                            "TM52='" & ChgSQL(textTM52) & "'," & _
                                            "TM76='" & textTM76 & "'," & _
                                            "TM90='" & textTM90 & "'," & _
                                            "TM91='" & textTM91 & "'," & _
                                            "TM92='" & textTM92 & "'," & _
                                            "TM93='" & textTM93 & "'," & _
                                            "TM96='" & ChgSQL(textTM96) & "'," & _
                                            "TM99='" & ChgSQL(textTM99) & "'," & _
                                            "TM102='" & ChgSQL(textTM102) & "'," & _
                                            "TM105='" & ChgSQL(textTM105) & "'," & _
                                            "TM108='" & ChgSQL(TextTM108) & "'," & _
                                            "TM111='" & ChgSQL(TextTM111) & "'," & _
                                            "TM114='" & ChgSQL(TextTM114) & "'," & _
                                            "TM117='" & ChgSQL(TextTM117) & "'," & _
                                            "TM45='" & textTM45 & "'," & _
                                            "TM131='" & ChgSQL(textTM131) & "' " & _
                           "WHERE tm01='" & textTM01 & "' and tm02='" & textTM02 & "' and tm03='" & textTM03 & "' and tm04='" & textTM04 & "' "
         Pub_SeekTbLog strExc(1)
         cnnConnection.Execute strExc(1)
         cnnConnection.CommitTrans
         'Modified by Lydia 2018/02/21 vbCritical=> vbInformation
         MsgBox "存檔完成 !", vbInformation
         CmdLock 1
         Call ClearAll(True)
         textTM01.SetFocus
      Case 1 '結束
         Unload frm140111
         Set frm140111 = Nothing
      Case 2 '尋找
         If textTM01 = "" Or textTM02 = "" Then
            MsgBox "請輸入本所案號 !", vbCritical
            Exit Sub
         End If
         If textTM03 = "" Then textTM03 = "0"
         If textTM04 = "" Then textTM04 = "00"
         Call ClearAll(False)
         intI = 1
         strExc(0) = "SELECT * FROM TradeMark WHERE tm01='" & textTM01 & "' and tm02='" & textTM02 & "' and tm03='" & textTM03 & "' and tm04='" & textTM04 & "' "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.RecordCount > 0 Then
               If Not IsNull(RsTemp.Fields("TM05")) Then textTM05 = RsTemp.Fields("TM05")
               If Not IsNull(RsTemp.Fields("TM131")) Then textTM131 = RsTemp.Fields("TM131") 'Add By Sindy 2015/7/14
               'Add By Sindy 2012/11/14
               m_TM09 = ""
               If Not IsNull(RsTemp.Fields("TM09")) Then m_TM09 = RsTemp.Fields("TM09")
               '2012/11/14 End
               If Not IsNull(RsTemp.Fields("TM26")) Then textTM26 = RsTemp.Fields("TM26")
               If Not IsNull(RsTemp.Fields("TM40")) Then textTM40 = RsTemp.Fields("TM40")
               If Not IsNull(RsTemp.Fields("TM43")) Then textTM43 = RsTemp.Fields("TM43")
               If Not IsNull(RsTemp.Fields("TM49")) Then textTM49 = RsTemp.Fields("TM49")
               If Not IsNull(RsTemp.Fields("TM52")) Then textTM52 = RsTemp.Fields("TM52")
               If Not IsNull(RsTemp.Fields("TM76")) Then textTM76 = RsTemp.Fields("TM76")
               If Not IsNull(RsTemp.Fields("TM90")) Then textTM90 = RsTemp.Fields("TM90")
               If Not IsNull(RsTemp.Fields("TM91")) Then textTM91 = RsTemp.Fields("TM91")
               If Not IsNull(RsTemp.Fields("TM92")) Then textTM92 = RsTemp.Fields("TM92")
               If Not IsNull(RsTemp.Fields("TM93")) Then textTM93 = RsTemp.Fields("TM93")
               If Not IsNull(RsTemp.Fields("TM96")) Then textTM96 = RsTemp.Fields("TM96")
               If Not IsNull(RsTemp.Fields("TM99")) Then textTM99 = RsTemp.Fields("TM99")
               If Not IsNull(RsTemp.Fields("TM102")) Then textTM102 = RsTemp.Fields("TM102")
               If Not IsNull(RsTemp.Fields("TM105")) Then textTM105 = RsTemp.Fields("TM105")
               If Not IsNull(RsTemp.Fields("TM108")) Then TextTM108 = RsTemp.Fields("TM108")
               If Not IsNull(RsTemp.Fields("TM111")) Then TextTM111 = RsTemp.Fields("TM111")
               If Not IsNull(RsTemp.Fields("TM114")) Then TextTM114 = RsTemp.Fields("TM114")
               If Not IsNull(RsTemp.Fields("TM117")) Then TextTM117 = RsTemp.Fields("TM117")
               'Add By Sindy 2012/4/16
               If Not IsNull(RsTemp.Fields("TM44")) Then textTM44 = RsTemp.Fields("TM44"): textTM44_2 = GetPrjName1(RsTemp.Fields("TM44"))
               If Not IsNull(RsTemp.Fields("TM45")) Then textTM45 = RsTemp.Fields("TM45")
               If Not IsNull(RsTemp.Fields("TM53")) Then textTM53 = RsTemp.Fields("TM53")
               '2012/4/16 End
            End If
            CmdLock 0
         Else
            MsgBox "本所案號錯誤，請重新輸入 !", vbCritical
            textTM01.SetFocus
         End If
      Case 3 '取消
         If MsgBox("你並未存檔，確定離開嗎 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
         CmdLock 1
         Call ClearAll(True)
         textTM01.SetFocus
   End Select
   Exit Sub
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbInformation
    Exit Sub
ErrorHandler:
    cnnConnection.RollbackTrans
    MsgBox "更新資料失敗，請洽系統管理員 !", vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140111 = Nothing
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If IsEmptyText(textTM05) = True Then
      MsgBox "案件名稱不可空白", vbOKOnly, "檢核資料"
      textTM05.SetFocus
      Exit Function
   End If
      
   If Me.textTM05.Enabled = True Then
      Cancel = False
      textTM05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2015/7/14
    If Me.textTM131.Enabled = True Then
       Cancel = False
       textTM131_Validate Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
    '2015/7/14 END
   
   If Me.textTM26.Enabled = True Then
      Cancel = False
      textTM26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM40.Enabled = True Then
      Cancel = False
      textTM40_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
      
   If Me.textTM43.Enabled = True Then
      Cancel = False
      textTM43_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM49.Enabled = True Then
      Cancel = False
      textTM49_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM52.Enabled = True Then
      Cancel = False
      textTM52_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM76.Enabled = True Then
      Cancel = False
      textTM76_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM90.Enabled = True Then
      Cancel = False
      textTM90_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM91.Enabled = True Then
      Cancel = False
      textTM91_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM92.Enabled = True Then
      Cancel = False
      textTM92_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM93.Enabled = True Then
      Cancel = False
      textTM93_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM96.Enabled = True Then
      Cancel = False
      textTM96_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM99.Enabled = True Then
      Cancel = False
      textTM99_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM102.Enabled = True Then
      Cancel = False
      textTM102_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM105.Enabled = True Then
      Cancel = False
      textTM105_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.TextTM108.Enabled = True Then
      Cancel = False
      textTM108_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.TextTM111.Enabled = True Then
      Cancel = False
      textTM111_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.TextTM114.Enabled = True Then
      Cancel = False
      textTM114_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.TextTM117.Enabled = True Then
      Cancel = False
      textTM117_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2012/4/16
   If Me.textTM53.Enabled = True Then
      Cancel = False
      textTM53_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Added by Lydia 2021/10/13 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If

   TxtValidate = True
End Function

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
End Sub

'Modified by Lydia 2021/10/13 改成Form 2.0
'Private Sub textTM01_KeyPress(KeyAscii As Integer)
Private Sub textTM01_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 系統類別
Private Sub textTM01_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      If Not IsCorrectSysKind(textTM01) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "系統類別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         Me.textTM01.Text = ""
         GoTo EXITSUB
      End If
      
      ' 檢查使用者是否有使用該系統類別的權限
      If IsUserHasRightOfSystem(strUserNum, textTM01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使有此系統別的權限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub

Private Sub textTM05_GotFocus()
   OpenIme
   InverseTextBox textTM05
End Sub

'Add By Sindy 2015/7/14
Private Sub textTM131_GotFocus()
   InverseTextBox textTM131
   '切換輸入法改用API
   OpenIme
End Sub
' 定稿商標名稱
Private Sub textTM131_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM131, textTM131.MaxLength) = False Then
      Cancel = True
      textTM131_GotFocus
   End If
   '切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub
'2015/7/14 END

Private Sub textTM26_GotFocus()
   OpenIme
   InverseTextBox textTM26
End Sub

Private Sub textTM40_GotFocus()
   OpenIme
   InverseTextBox textTM40
End Sub

Private Sub textTM43_GotFocus()
   OpenIme
   InverseTextBox textTM43
End Sub

'Add By Sindy 2012/4/16
'Modified by Lydia 2021/10/13 改成Form 2.0
'Private Sub textTM45_KeyPress(KeyAscii As Integer)
Private Sub textTM45_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM49_GotFocus()
   OpenIme
   InverseTextBox textTM49
End Sub

Private Sub textTM52_GotFocus()
   OpenIme
   InverseTextBox textTM52
End Sub

Private Sub textTM76_GotFocus()
   OpenIme
   InverseTextBox textTM76
End Sub

Private Sub textTM90_GotFocus()
   OpenIme
   InverseTextBox textTM90
End Sub

Private Sub textTM91_GotFocus()
   OpenIme
   InverseTextBox textTM91
End Sub

Private Sub textTM92_GotFocus()
   OpenIme
   InverseTextBox textTM92
End Sub

Private Sub textTM93_GotFocus()
   OpenIme
   InverseTextBox textTM93
End Sub

Private Sub textTM96_GotFocus()
   OpenIme
   InverseTextBox textTM96
End Sub

Private Sub textTM99_GotFocus()
   OpenIme
   InverseTextBox textTM99
End Sub

Private Sub textTM102_GotFocus()
   OpenIme
   InverseTextBox textTM102
End Sub

Private Sub textTM105_GotFocus()
   OpenIme
   InverseTextBox textTM105
End Sub

Private Sub textTM108_GotFocus()
   OpenIme
   InverseTextBox TextTM108
End Sub

Private Sub textTM111_GotFocus()
   OpenIme
   InverseTextBox TextTM111
End Sub

Private Sub textTM114_GotFocus()
   OpenIme
   InverseTextBox TextTM114
End Sub

Private Sub textTM117_GotFocus()
   OpenIme
   InverseTextBox TextTM117
End Sub

'Add By Sindy 2012/4/16
Private Sub textTM45_GotFocus()
   InverseTextBox textTM45
End Sub

'Add By Sindy 2012/4/16
Private Sub textTM53_GotFocus()
   InverseTextBox textTM53
End Sub

' 案件名稱(中)
Private Sub textTM05_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM05, 140) = False Then
      Cancel = True
      textTM05_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 申請地址(日)
Private Sub textTM26_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM26, 70) = False Then
      Cancel = True
      textTM26_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 聯絡人1(日)
Private Sub textTM40_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM40, textTM40.MaxLength) = False Then
      Cancel = True
      textTM40_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 聯絡人2(日)
Private Sub textTM43_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM43, textTM43.MaxLength) = False Then
      Cancel = True
      textTM43_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 代表人1(日)
Private Sub textTM49_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM49, 40) = False Then
   If CheckLengthIsOK(textTM49, textTM49.MaxLength) = False Then
      Cancel = True
      textTM49_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 代表人2(日)
Private Sub textTM52_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM52, 40) = False Then
   If CheckLengthIsOK(textTM52, textTM52.MaxLength) = False Then
      Cancel = True
      textTM52_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM76_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM76, textTM76.MaxLength) = False Then
      Cancel = True
      textTM76_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM90_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM90, 70) = False Then
      Cancel = True
      textTM90_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM91_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM91, 70) = False Then
      Cancel = True
      textTM91_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM92_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM92, 70) = False Then
      Cancel = True
      textTM92_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM93_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM93, 70) = False Then
      Cancel = True
      textTM93_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 代表人3(日)
Private Sub textTM96_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM96, 40) = False Then
   If CheckLengthIsOK(textTM96, textTM96.MaxLength) = False Then
      Cancel = True
      textTM96_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 代表人4(日)
Private Sub textTM99_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM99, 40) = False Then
   If CheckLengthIsOK(textTM99, textTM102.MaxLength) = False Then
      Cancel = True
      textTM99_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 代表人5(日)
Private Sub textTM102_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM102, 40) = False Then
   If CheckLengthIsOK(textTM102, textTM102.MaxLength) = False Then
      Cancel = True
      textTM102_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 代表人6(日)
Private Sub textTM105_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM105, 40) = False Then
   If CheckLengthIsOK(textTM105, textTM105.MaxLength) = False Then
      Cancel = True
      textTM105_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 代表人7(日)
Private Sub textTM108_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM108, 40) = False Then
   If CheckLengthIsOK(TextTM108, TextTM108.MaxLength) = False Then
      Cancel = True
      textTM108_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 代表人8(日)
Private Sub textTM111_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM111, 40) = False Then
   If CheckLengthIsOK(TextTM111, TextTM111.MaxLength) = False Then
      Cancel = True
      textTM111_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 代表人9(日)
Private Sub textTM114_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM114, 40) = False Then
   If CheckLengthIsOK(TextTM114, TextTM114.MaxLength) = False Then
      Cancel = True
      textTM114_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 代表人10(日)
Private Sub textTM117_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM117, 40) = False Then
   If CheckLengthIsOK(TextTM117, TextTM117.MaxLength) = False Then
      Cancel = True
      textTM117_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

'Add By Sindy 2012/4/16
' 定稿語文
Private Sub textTM53_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM53) = False Then
      Select Case textTM53
         Case "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入 1 或 2 或 3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM53_GotFocus
      End Select
   End If
End Sub
