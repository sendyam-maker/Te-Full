VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140110 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利日文資料維護作業"
   ClientHeight    =   6045
   ClientLeft      =   570
   ClientTop       =   975
   ClientWidth     =   9210
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9210
   Visible         =   0   'False
   Begin VB.CheckBox ChkPA174 
      Caption         =   "有特殊字"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   60
      TabIndex        =   51
      Top             =   540
      Width           =   1035
   End
   Begin VB.CommandButton CmdPA174 
      BackColor       =   &H00C0FFFF&
      Caption         =   "特殊字"
      Height          =   280
      Left            =   300
      Style           =   1  '圖片外觀
      TabIndex        =   50
      Top             =   720
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   350
      Index           =   2
      Left            =   3800
      TabIndex        =   4
      Top             =   240
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   5700
      TabIndex        =   25
      Top             =   120
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   7620
      TabIndex        =   27
      Top             =   120
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消(&C)"
      Height          =   400
      Index           =   3
      Left            =   6660
      TabIndex        =   26
      Top             =   120
      Width           =   912
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   7485
      VariousPropertyBits=   679495707
      Size            =   "13203;503"
      FontName        =   "細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   132
      Left            =   5640
      TabIndex        =   24
      Top             =   5475
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   126
      Left            =   5640
      TabIndex        =   22
      Top             =   5154
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   120
      Left            =   5640
      TabIndex        =   20
      Top             =   4836
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   114
      Left            =   5640
      TabIndex        =   18
      Top             =   4518
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   84
      Left            =   5640
      TabIndex        =   16
      Top             =   4200
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   129
      Left            =   1080
      TabIndex        =   23
      Top             =   5475
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   123
      Left            =   1080
      TabIndex        =   21
      Top             =   5160
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   117
      Left            =   1080
      TabIndex        =   19
      Top             =   4845
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   111
      Left            =   1080
      TabIndex        =   17
      Top             =   4515
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   81
      Left            =   1080
      TabIndex        =   15
      Top             =   4200
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   100
      Left            =   1440
      TabIndex        =   14
      Top             =   3840
      Width           =   3060
      VariousPropertyBits=   679495707
      Size            =   "5397;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   139
      Left            =   1440
      TabIndex        =   13
      Top             =   3520
      Width           =   3060
      VariousPropertyBits=   679495707
      Size            =   "5397;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   56
      Left            =   1440
      TabIndex        =   12
      Top             =   3195
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   53
      Left            =   1440
      TabIndex        =   11
      Top             =   2880
      Width           =   3375
      VariousPropertyBits=   679495707
      Size            =   "5953;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   45
      Left            =   1440
      TabIndex        =   10
      Top             =   2520
      Width           =   7485
      VariousPropertyBits=   679495707
      Size            =   "13203;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   44
      Left            =   1440
      TabIndex        =   9
      Top             =   2205
      Width           =   7485
      VariousPropertyBits=   679495707
      Size            =   "13203;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   43
      Left            =   1440
      TabIndex        =   8
      Top             =   1890
      Width           =   7485
      VariousPropertyBits=   679495707
      Size            =   "13203;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   42
      Left            =   1440
      TabIndex        =   7
      Top             =   1590
      Width           =   7485
      VariousPropertyBits=   679495707
      Size            =   "13203;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   41
      Left            =   1440
      TabIndex        =   6
      Top             =   1275
      Width           =   7485
      VariousPropertyBits=   679495707
      Size            =   "13203;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   495
      VariousPropertyBits=   679495707
      Size            =   "873;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   375
      VariousPropertyBits=   679495707
      Size            =   "661;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   855
      VariousPropertyBits=   679495707
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   280
      Index           =   1
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   615
      VariousPropertyBits=   679495707
      Size            =   "1085;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Left            =   1440
      TabIndex        =   49
      Top             =   600
      Width           =   7485
      Size            =   "13203;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人10(日):"
      Height          =   180
      Index           =   6
      Left            =   4530
      TabIndex        =   48
      Top             =   5527
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人8(日):"
      Height          =   180
      Index           =   5
      Left            =   4530
      TabIndex        =   47
      Top             =   5206
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人6(日):"
      Height          =   180
      Index           =   4
      Left            =   4530
      TabIndex        =   46
      Top             =   4888
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人4(日):"
      Height          =   180
      Index           =   2
      Left            =   4530
      TabIndex        =   45
      Top             =   4570
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人2(日):"
      Height          =   180
      Index           =   1
      Left            =   4530
      TabIndex        =   44
      Top             =   4252
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人9(日):"
      Height          =   180
      Index           =   125
      Left            =   90
      TabIndex        =   43
      Top             =   5527
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人7(日):"
      Height          =   180
      Index           =   124
      Left            =   90
      TabIndex        =   42
      Top             =   5206
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人5(日):"
      Height          =   180
      Index           =   123
      Left            =   90
      TabIndex        =   41
      Top             =   4888
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人3(日):"
      Height          =   180
      Index           =   122
      Left            =   90
      TabIndex        =   40
      Top             =   4570
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代表人1(日):"
      Height          =   180
      Index           =   121
      Left            =   90
      TabIndex        =   39
      Top             =   4252
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人(日)2:"
      Height          =   180
      Index           =   44
      Left            =   420
      TabIndex        =   38
      Top             =   3247
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人(日)1:"
      Height          =   180
      Index           =   31
      Left            =   420
      TabIndex        =   37
      Top             =   2932
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "實體聯絡人(日):"
      Height          =   180
      Index           =   79
      Left            =   150
      TabIndex        =   36
      Top             =   3870
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人部門(日):"
      Height          =   180
      Index           =   76
      Left            =   150
      TabIndex        =   35
      Top             =   3555
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人1地址(日):"
      Height          =   180
      Index           =   102
      Left            =   60
      TabIndex        =   34
      Top             =   1324
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2地址(日):"
      Height          =   180
      Index           =   107
      Left            =   60
      TabIndex        =   33
      Top             =   1636
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3地址(日):"
      Height          =   180
      Index           =   108
      Left            =   60
      TabIndex        =   32
      Top             =   1948
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4地址(日):"
      Height          =   180
      Index           =   109
      Left            =   60
      TabIndex        =   31
      Top             =   2260
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5地址(日):"
      Height          =   180
      Index           =   110
      Left            =   60
      TabIndex        =   30
      Top             =   2572
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱(外):"
      Height          =   180
      Index           =   3
      Left            =   330
      TabIndex        =   29
      Top             =   1012
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   630
      TabIndex        =   28
      Top             =   292
      Width           =   765
   End
End
Attribute VB_Name = "frm140110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2018/11/20 改成Form2.0 (lblCaseName和Textbox)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
Dim oTxt  As MSForms.TextBox  'Added by Lydia 2018/11/21
Dim bolMsgRight As Boolean 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效
Dim SyxMsg As String 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效(記錄前一位置)

'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim rsDefineSize As New ADODB.Recordset, intWhere As Integer
'Added by Lydia 2020/02/21
Public bolAskPA174 As Boolean '存檔前檢查有修改案件名稱，將原始檔之維護word檔自動打開，是否有上傳
Dim m_PA07

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bUpdate = IsUserHasRightOfFunction("frm140110", strEdit, False)
   
   MoveFormToCenter Me
   CmdLock 1
   
   strExc(0) = "SELECT * FROM PATENT WHERE ROWNUM<1"
   intI = 1
   Set rsDefineSize = ClsLawReadRstMsg(intI, strExc(0))
   
   'Added by Lydia 2018/11/20 模組-抓DB中的欄位實際長度
   For Each oTxt In Text1
          oTxt.MaxLength = PUB_GetFieldDefSize("PATENT", "PA" & Format(oTxt.Index, "00"))
   Next
   
   Call ClearAll(True)  'Added by Lydia 2020/02/21
End Sub

Private Sub CmdLock(TF As Integer)
   Select Case TF
      Case 0
         Command1(2).Enabled = False
         Command1(0).Enabled = True
         Command1(3).Enabled = True
         Text1(1).Locked = True
         Text1(2).Locked = True
         Text1(3).Locked = True
         Text1(4).Locked = True
      Case 1
         Command1(2).Enabled = True
         Command1(0).Enabled = False
         Command1(3).Enabled = False
         Text1(1).Locked = False
         Text1(2).Locked = False
         Text1(3).Locked = False
         Text1(4).Locked = False
   End Select
End Sub

Private Sub ClearAll(bClearPk As Boolean)
   'Modified by Lydia 2018/11/20
'   If bClearPk = True Then
'      Text1(1) = Empty
'      Text1(2) = Empty
'      Text1(3) = Empty
'      Text1(4) = Empty
'   End If
'   Text1(7) = Empty
'   Text1(41) = Empty
'   Text1(42) = Empty
'   Text1(43) = Empty
'   Text1(44) = Empty
'   Text1(45) = Empty
'   Text1(53) = Empty
'   Text1(56) = Empty
'   Text1(81) = Empty
'   Text1(84) = Empty
'   Text1(100) = Empty
'   Text1(111) = Empty
'   Text1(114) = Empty
'   Text1(117) = Empty
'   Text1(120) = Empty
'   Text1(123) = Empty
'   Text1(126) = Empty
'   Text1(129) = Empty
'   Text1(132) = Empty
'   Text1(139) = Empty
   For Each oTxt In Text1
       If oTxt.Index <= 4 Then
           If bClearPk = True Then
               oTxt.Text = Empty
           End If
       Else
           oTxt.Text = Empty
       End If
   Next
   'end 2018/11/20
   
   'Modified by Lydia 2018/11/20
   'Label1(7).Caption = Empty
   lblCaseName.Caption = Empty

   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   ChkPA174.Visible = False
   ChkPA174.Value = vbUnchecked: ChkPA174.Tag = ""
   CmdPA174.Visible = False
   bolAskPA174 = False
   m_PA07 = ""
   'end 2020/02/21
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo ErrHand
   Select Case Index
      Case 0 '確定
         
         'Added by Lydia 2021/04/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me, , True) = False Then
             Exit Sub
         End If
         'end 2021/04/14
         
         If TxtValidate = False Then Exit Sub 'Added by Lydia 2020/02/21
         
         On Error GoTo ErrorHandler
         cnnConnection.BeginTrans
         'Modified by Lydia 2020/02/21 +PA174
         'strExc(1) = "UPDATE Patent " & _
                                    "SET PA07='" & Text1(7) & "'," & _
                                            "PA41='" & Text1(41) & "'," & _
                                            "PA42='" & Text1(42) & "'," & _
                                            "PA43='" & Text1(43) & "'," & _
                                            "PA44='" & Text1(44) & "'," & _
                                            "PA45='" & Text1(45) & "'," & _
                                            "PA53='" & Text1(53) & "'," & _
                                            "PA56='" & Text1(56) & "'," & _
                                            "PA81='" & Text1(81) & "'," & _
                                            "PA84='" & Text1(84) & "'," & _
                                            "PA100='" & Text1(100) & "'," & _
                                            "PA111='" & Text1(111) & "'," & _
                                            "PA114='" & Text1(114) & "'," & _
                                            "PA117='" & Text1(117) & "'," & _
                                            "PA120='" & Text1(120) & "'," & _
                                            "PA123='" & Text1(123) & "'," & _
                                            "PA126='" & Text1(126) & "'," & _
                                            "PA129='" & Text1(129) & "'," & _
                                            "PA132='" & Text1(132) & "'," & _
                                            "PA139='" & Text1(139) & "' " & _
                           "WHERE pa01='" & Text1(1) & _
                                 "' and pa02='" & Text1(2) & _
                                 "' and pa03='" & Text1(3) & _
                                 "' and pa04='" & Text1(4) & "' "
         strExc(1) = "UPDATE Patent " & _
                                    "SET PA07='" & Text1(7) & "'," & _
                                            "PA41='" & Text1(41) & "',PA42='" & Text1(42) & "',PA43='" & Text1(43) & "',PA44='" & Text1(44) & "',PA45='" & Text1(45) & "'," & _
                                            "PA53='" & Text1(53) & "',PA56='" & Text1(56) & "'," & _
                                            "PA81='" & Text1(81) & "',PA84='" & Text1(84) & "'," & _
                                            "PA100='" & Text1(100) & "',PA111='" & Text1(111) & "'," & _
                                            "PA114='" & Text1(114) & "',PA117='" & Text1(117) & "'," & _
                                            "PA120='" & Text1(120) & "',PA123='" & Text1(123) & "'," & _
                                            "PA126='" & Text1(126) & "',PA129='" & Text1(129) & "'," & _
                                            "PA132='" & Text1(132) & "',PA139='" & Text1(139) & "', " & _
                                            "PA174='" & IIf(ChkPA174.Value = 0, "", "Y") & "' " & _
                           "WHERE pa01='" & Text1(1) & _
                                 "' and pa02='" & Text1(2) & _
                                 "' and pa03='" & Text1(3) & _
                                 "' and pa04='" & Text1(4) & "' "
                                 
         'Modified by Lydia 2021/04/27 更新來源的表單名稱
         'Pub_SeekTbLog strExc(1)
         Pub_SeekTbLog strExc(1), , , , Me.Caption & "(" & Me.Name & ")"
         cnnConnection.Execute strExc(1)
         cnnConnection.CommitTrans
         
         MsgBox "存檔完成 !", vbInformation
         CmdLock 1
         Call ClearAll(True)
         Text1(1).SetFocus
      Case 1 '結束
         Unload frm140110
         Set frm140110 = Nothing
      Case 2 '尋找
         If Text1(1) = "" Or Text1(2) = "" Then
            MsgBox "請輸入本所案號 !", vbCritical
            Exit Sub
         End If
         If Text1(3) = "" Then Text1(3) = "0"
         If Text1(4) = "" Then Text1(4) = "00"
         
         'Added by Morgan 2012/5/16
         If CheckSR09(strUserNum, Text1(1), "Y", False, Text1(1), Text1(2), Text1(3), Text1(4)) = False Then
            MsgBox "本所案號輸入錯誤 !", vbCritical
            Text1(1).SetFocus
            Text1_GotFocus 1
            Exit Sub
         End If
         'end 2012/5/16
         
         Call ClearAll(False)
         intI = 1
         strExc(0) = "SELECT * FROM Patent WHERE pa01='" & Text1(1) & "' and pa02='" & Text1(2) & "' and pa03='" & Text1(3) & "' and pa04='" & Text1(4) & "' "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.RecordCount > 0 Then
               'Modified by Lydia 2018/11/20
               'Label1(7).Caption = "" & RsTemp.Fields("PA05") & "" & RsTemp.Fields("PA06")
               lblCaseName.Caption = "" & RsTemp.Fields("PA05") & "" & RsTemp.Fields("PA06")
               
               If Not IsNull(RsTemp.Fields("PA07")) Then Text1(7) = RsTemp.Fields("PA07")
               m_PA07 = "" & RsTemp.Fields("PA07") 'Added by Lydia 2020/02/21
               If Not IsNull(RsTemp.Fields("PA41")) Then Text1(41) = RsTemp.Fields("PA41")
               If Not IsNull(RsTemp.Fields("PA42")) Then Text1(42) = RsTemp.Fields("PA42")
               If Not IsNull(RsTemp.Fields("PA43")) Then Text1(43) = RsTemp.Fields("PA43")
               If Not IsNull(RsTemp.Fields("PA44")) Then Text1(44) = RsTemp.Fields("PA44")
               If Not IsNull(RsTemp.Fields("PA45")) Then Text1(45) = RsTemp.Fields("PA45")
               If Not IsNull(RsTemp.Fields("PA53")) Then Text1(53) = RsTemp.Fields("PA53")
               If Not IsNull(RsTemp.Fields("PA56")) Then Text1(56) = RsTemp.Fields("PA56")
               If Not IsNull(RsTemp.Fields("PA81")) Then Text1(81) = RsTemp.Fields("PA81")
               If Not IsNull(RsTemp.Fields("PA84")) Then Text1(84) = RsTemp.Fields("PA84")
               If Not IsNull(RsTemp.Fields("PA100")) Then Text1(100) = RsTemp.Fields("PA100")
               If Not IsNull(RsTemp.Fields("PA111")) Then Text1(111) = RsTemp.Fields("PA111")
               If Not IsNull(RsTemp.Fields("PA114")) Then Text1(114) = RsTemp.Fields("PA114")
               If Not IsNull(RsTemp.Fields("PA117")) Then Text1(117) = RsTemp.Fields("PA117")
               If Not IsNull(RsTemp.Fields("PA120")) Then Text1(120) = RsTemp.Fields("PA120")
               If Not IsNull(RsTemp.Fields("PA123")) Then Text1(123) = RsTemp.Fields("PA123")
               If Not IsNull(RsTemp.Fields("PA126")) Then Text1(126) = RsTemp.Fields("PA126")
               If Not IsNull(RsTemp.Fields("PA129")) Then Text1(129) = RsTemp.Fields("PA129")
               If Not IsNull(RsTemp.Fields("PA132")) Then Text1(132) = RsTemp.Fields("PA132")
               If Not IsNull(RsTemp.Fields("PA139")) Then Text1(139) = RsTemp.Fields("PA139")
                'Added by Lydia 2020/02/21 預設「名稱有特殊字」
                If Text1(1) = "FCP" Or Text1(1) = "P" Then
                   ChkPA174.Visible = True
                   CmdPA174.Visible = True
                   If "" & RsTemp.Fields("PA174") = "Y" Then
                       ChkPA174.Value = vbChecked
                       ChkPA174.Tag = "" & RsTemp.Fields("PA174")
                   End If
                End If
                'end 2020/02/21
            End If
            CmdLock 0
         Else
            MsgBox "本所案號錯誤，請重新輸入 !", vbCritical
            Text1(1).SetFocus
         End If
      Case 3 '取消
         If MsgBox("你並未存檔，確定離開嗎 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
         CmdLock 1
         Call ClearAll(True)
         Text1(1).SetFocus
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
   Set frm140110 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   Select Case Index
      Case 7, 42, 43, 45, 53, 56, 81, 84, 100, 111, 114, 117, 120, 123, 126, 129, 132, 139
         OpenIme
      Case Else
         CloseIme
   End Select
End Sub

'Modified by Lydia 2018/11/20 改成Form2.0
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 1, 3
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub
Private Sub Text1_LostFocus(Index As Integer)
'Memo by Lydia 2018/11/20 測試
'   Select Case Index
'      Case 56
'        Text1(Index).Text = ConvBig5(txtTmp, Text1(Index).Text)
'   End Select
End Sub

'Remove by Lydia 2020/04/01 FCP-61869會發生堆疊空間不足的錯誤,先移除
''Memo by Lydia 2018/11/20 測試
'Private Sub Text1_Change(Index As Integer)
'   Select Case Index
'      Case 1, 2, 3, 4
'      Case 53
'
'      Case Else
'          Text1(Index).Text = ConvBig5_2(Me.txtTmp, Text1(Index).Text)
'   End Select
'
'End Sub
'
''Memo by Lydia 2018/11/20 測試
'Private Function ConvBig5(ByRef pLoc As TextBox, ByVal pStr As String) As String
'Dim tmpText As TextBox
'
'     'strT = pStr
''     tmpText.Caption = pStr
''     ConvBig5 = tmpText.Caption
'End Function
''Memo by Lydia 2018/11/20 測試
''Private Function ConvBig5_2(ByRef pLoc As TextBox, ByVal pStr As String) As String
''Dim TT As Controls
''     With TT
''         .Text = pStr
''     End With
''     'pLoc.Text = pStr
''     ConvBig5_2 = TT.Text
''End Function
'
'Private Function ConvBig5_2(ByRef pLoc As Control, ByVal pStr As String) As String
''Tag 仍是以?儲存Unicode
'     pLoc.Text = pStr
'     ConvBig5_2 = pLoc.Text
'End Function
'end 2020/04/01

' 系統類別
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim i As Integer
   Cancel = False
   Select Case Index
      Case 1
         If IsEmptyText(Text1(Index)) = False Then
            If Not IsCorrectSysKind(Text1(Index)) Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "系統類別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               Call Text1_GotFocus(Index)
               Me.Text1(Index).Text = ""
               GoTo EXITSUB
            End If
            
'Removed by Morgan 2012/5/16 改按尋找時檢查
'            ' 檢查使用者是否有使用該系統類別的權限
'            If IsUserHasRightOfSystem(strUserNum, Text1(Index)) = False Then
'               Cancel = True
'               strTit = "資料檢核"
'               strMsg = "您沒有使有此系統別的權限"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               Call Text1_GotFocus(Index)
'               GoTo EXITSUB
'            End If
         End If
   End Select
   
   '檢查中文欄位長度是否過長
   If CheckLengthIsOK(Text1(Index).Text, rsDefineSize.Fields(Index - 1).DefinedSize) = False Then
      Cancel = True
   End If
   If Cancel = True Then TextInverse Text1(Index)
EXITSUB:
End Sub

'Added by Lydia 2018/11/21
Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If SyxMsg <> "Text1_" & Format(Index, "00") Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "Text1_" & Format(Index, "00")
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
    
End Sub

'Added by Lydia 2020/02/21 檢核資料
Private Function TxtValidate() As Boolean
   TxtValidate = False
            
   '檢查「名稱有特殊字」
   If Text1(1) = "P" Or Text1(1) = "FCP" Then
       If Pub_GetPA174toFile("2", Text1(1), Text1(2), Text1(3), Text1(4), Me, frm100101_M_1) = True Then
           strExc(1) = "Y"
       Else
           strExc(1) = "N"
       End If
       If ChkPA174.Value = vbUnchecked And strExc(1) = "Y" Then
           If MsgBox("原始檔區已有案件名稱Word檔，請問是否取消「名稱有特殊字」？", vbInformation + vbYesNo + vbDefaultButton2, "檢查資料") = vbNo Then
               Exit Function
           End If
       End If
       If ChkPA174.Value = vbChecked And strExc(1) = "N" Then
           If MsgBox("原始檔區沒有案件名稱Word檔，請問是否繼續作業？", vbInformation + vbYesNo + vbDefaultButton2, "檢查資料") = vbNo Then
               Exit Function
           End If
       End If
       '當「名稱有特殊字」有勾選，並且有修改案件名稱，將原始檔之維護word檔自動打開，並彈訊息提醒。
       If ChkPA174.Value = vbChecked And bolAskPA174 = False Then  '不用再次彈訊息
           If Text1(7) <> m_PA07 Then
               MsgBox "名稱有特殊字，案件名稱有修改，請一併修改案件名稱Word檔。", vbInformation, "檢查資料"
               Call ProcPA174toFile("Y")
               Exit Function
           End If
       End If
   End If
   
   TxtValidate = True
End Function

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟/維護FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()
    Call ProcPA174toFile("N")

End Sub

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟/維護FCP0xxxxx.新案性質.案件名稱.doc
Private Sub ProcPA174toFile(ByVal pKind As String)
Dim strKind As String

    If Text1(1).Locked = False Or Text1(2).Locked = False Then
        MsgBox "請先查詢正確的案號!", vbExclamation, "檢核資料"
    Else
        If ChkPA174.Value = vbUnchecked Then
            MsgBox "請先勾選「有特殊字」!", vbInformation + vbOKOnly, Me.Caption
        Else
            If pKind = "Y" Then 'bolAskPA174
                strKind = "3"
            Else
                strKind = "1"
            End If
            If Pub_GetPA174toFile(strKind, Me.Text1(1), Me.Text1(2), Me.Text1(3), Me.Text1(4), Me, frm100101_M_1) = True Then
            End If
        End If
    End If
    
End Sub

'Added by Lydia 2020/02/21
Public Sub PubShowNextData()
   '原始檔Word檔維護，上傳後直接進入存檔
   If bolAskPA174 = True Then
        Call Command1_Click(0) '確定->存檔
   End If
End Sub

