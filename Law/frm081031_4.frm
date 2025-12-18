VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081031_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "TIPS自動內部收文-本所期限"
   ClientHeight    =   4512
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6852
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4512
   ScaleWidth      =   6852
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&E)"
      Height          =   330
      Left            =   4530
      TabIndex        =   41
      Top             =   90
      Width           =   1005
   End
   Begin VB.TextBox textCP06 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   2730
      MaxLength       =   7
      TabIndex        =   31
      Top             =   4080
      Width           =   1000
   End
   Begin VB.TextBox textCP06 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   2730
      MaxLength       =   7
      TabIndex        =   30
      Top             =   3765
      Width           =   1000
   End
   Begin VB.TextBox textCP06 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   2730
      MaxLength       =   7
      TabIndex        =   29
      Top             =   3435
      Width           =   1000
   End
   Begin VB.TextBox textCP06 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   2730
      MaxLength       =   7
      TabIndex        =   28
      Top             =   3120
      Width           =   1000
   End
   Begin VB.TextBox textCP06 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   2730
      MaxLength       =   7
      TabIndex        =   27
      Top             =   2790
      Width           =   1000
   End
   Begin VB.TextBox textCP06 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   2730
      MaxLength       =   7
      TabIndex        =   26
      Top             =   2475
      Width           =   1000
   End
   Begin VB.TextBox textCP06 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   2730
      MaxLength       =   7
      TabIndex        =   25
      Top             =   2145
      Width           =   1000
   End
   Begin VB.TextBox textCP06 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   2730
      MaxLength       =   7
      TabIndex        =   24
      Top             =   1830
      Width           =   1000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "回前畫面(&U)"
      Height          =   330
      Left            =   5580
      TabIndex        =   1
      Top             =   90
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "當年度8月31日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   4080
      TabIndex        =   40
      Top             =   4090
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "當年度8月20日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   4080
      TabIndex        =   39
      Top             =   3770
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "當年度8月15日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   4080
      TabIndex        =   38
      Top             =   3450
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "當年度7月15日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   4080
      TabIndex        =   37
      Top             =   3130
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "當年度5月31日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   4080
      TabIndex        =   36
      Top             =   2810
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "當年度6月30日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   4080
      TabIndex        =   35
      Top             =   2490
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "當年度5月31日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   4080
      TabIndex        =   34
      Top             =   2170
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "收文日+30日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   4080
      TabIndex        =   33
      Top             =   1850
      Width           =   1545
   End
   Begin VB.Label Label3 
      Caption         =   "當年度申請驗證"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   4080
      TabIndex        =   32
      Top             =   1530
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "本所期限"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2820
      TabIndex        =   23
      Top             =   1530
      Width           =   885
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  '靠右對齊
      Caption         =   "218 驗證申請："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   7
      Left            =   360
      TabIndex        =   22
      Top             =   4080
      Width           =   2205
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  '靠右對齊
      Caption         =   "216 自評報告："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   6
      Left            =   360
      TabIndex        =   21
      Top             =   3765
      Width           =   2205
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  '靠右對齊
      Caption         =   "215 管理審查："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   5
      Left            =   360
      TabIndex        =   20
      Top             =   3435
      Width           =   2205
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  '靠右對齊
      Caption         =   "213 內部稽核："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   4
      Left            =   360
      TabIndex        =   19
      Top             =   3120
      Width           =   2205
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  '靠右對齊
      Caption         =   "211 文件修制訂："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   3
      Left            =   360
      TabIndex        =   18
      Top             =   2790
      Width           =   2205
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  '靠右對齊
      Caption         =   "2092 全體員工教育訓練："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   2
      Left            =   360
      TabIndex        =   17
      Top             =   2475
      Width           =   2205
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  '靠右對齊
      Caption         =   "2091 權責人員教育訓練："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   1
      Left            =   360
      TabIndex        =   16
      Top             =   2145
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   15
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   225
      Index           =   2
      Left            =   90
      TabIndex        =   14
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   225
      Index           =   3
      Left            =   3150
      TabIndex        =   13
      Top             =   1170
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "收文日期："
      Height          =   225
      Index           =   4
      Left            =   90
      TabIndex        =   12
      Top             =   1170
      Width           =   900
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1050
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   480
      Width           =   5715
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "10081;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   225
      Index           =   5
      Left            =   2400
      TabIndex        =   10
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(0)"
      Height          =   285
      Index           =   0
      Left            =   1050
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(1)"
      Height          =   285
      Index           =   1
      Left            =   3300
      TabIndex        =   8
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(2)"
      Height          =   285
      Index           =   2
      Left            =   1050
      TabIndex        =   7
      Top             =   1170
      Width           =   1635
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(3)"
      Height          =   285
      Index           =   3
      Left            =   4080
      TabIndex        =   6
      Top             =   1170
      Width           =   675
   End
   Begin MSForms.Label lblFM2 
      Height          =   285
      Index           =   0
      Left            =   4800
      TabIndex        =   5
      Top             =   1170
      Width           =   1845
      VariousPropertyBits=   27
      Caption         =   "lblFM2(0)"
      Size            =   "3254;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "當事人："
      Height          =   225
      Index           =   17
      Left            =   90
      TabIndex        =   4
      Top             =   840
      Width           =   900
   End
   Begin VB.Label lblData 
      Caption         =   "lblData(4)"
      Height          =   285
      Index           =   4
      Left            =   1050
      TabIndex        =   3
      Top             =   840
      Width           =   825
   End
   Begin MSForms.Label lblFM2 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   4815
      VariousPropertyBits=   27
      Caption         =   "lblFM2(1)"
      Size            =   "8493;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  '靠右對齊
      Caption         =   "208 啟始會議："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1830
      Width           =   2205
   End
End
Attribute VB_Name = "frm081031_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/04/14 Form2.0已修改 lblFM2
'Create by Lydia 2023/04/14 TIPS自動內部收文-本所期限
Option Explicit
Dim m_PrevForm As Form  '前一畫面
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String '本所案號
Dim m_CP09 As String  '收文號
Dim m_Kind As String 'Added by Lydia 2023/06/21 輸入選項
Dim strTmpQ As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim oObj

Public Sub SetParent(ByVal pFrm As Form, ByVal pCP09 As String, ByVal pKind As String)
    Set m_PrevForm = pFrm
    m_CP09 = pCP09
    m_Kind = pKind
End Sub

Private Sub cmdExit_Click()
    m_PrevForm.Show
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim tmpBol As Boolean
    
    strTmpQ = ""
    For Each oObj In textCP06
        Call textCP06_Validate(oObj.Index, tmpBol)
        If tmpBol = True Then
            Exit Sub
        End If
        If oObj.Visible = True Then 'Added by Lydia 2023/06/21
           strTmpQ = strTmpQ & "," & DBDATE(textCP06(oObj.Index))
        End If 'Added by Lydia 2023/06/21
    Next
    If strTmpQ <> "" Then
        m_PrevForm.strBCP06List = Mid(strTmpQ, 2)
        Call cmdExit_Click
    End If
End Sub

Private Sub Form_Load()
  MoveFormToCenter Me
  Call ClearForm
  Call doQuery
  'Added by Lydia 2023/06/21
  If m_Kind = "F" Then '隱藏216 自評報告,218 驗證申請
     lblTitle(6).Visible = False: lblTitle(7).Visible = False
     textCP06(6).Visible = False: textCP06(7).Visible = False
     Label3(7).Visible = False: Label3(8).Visible = False
  Else
     lblTitle(6).Visible = True: lblTitle(7).Visible = True
     textCP06(6).Visible = True: textCP06(7).Visible = True
     Label3(7).Visible = True: Label3(8).Visible = True
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm081031_4 = Nothing
End Sub

Private Sub ClearForm()

    For Each oObj In lblFM2
       oObj.Caption = ""
    Next
    For Each oObj In lblData
       oObj.Caption = ""
    Next
    For Each oObj In textCP06
       oObj.Text = ""
       oObj.Tag = ""
    Next
    
End Sub

Private Sub doQuery()

    strTmpQ = "select cp01||'-'||cp02||decode(cp03,'0',null,'-'||cp03)||decode(cp04,'00',null,'-'||cp04) caseno,cp09," & _
                     "lc05,lc06,lc07,lc11, nvl(cu04,nvl(cu05,cu06)) lc11n,cp05,cp13,st02 as cp13n," & _
                     "cp16,cp18,cp15,cp01,cp02,cp03,cp04,cp53,cp54 " & _
                     "from caseprogress, lawcase, staff, customer " & _
                     "where cp09='" & m_CP09 & "' and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp13=st01(+) " & _
                     "and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) "
    intQ = 0
    Set rsQuery = ClsLawReadRstMsg(intQ, strTmpQ)
    If intQ = 0 Then
        Exit Sub
    End If
    intQ = 0
    Combo1.AddItem "中：" & rsQuery.Fields("lc05"), 0
    If rsQuery.Fields("lc05") <> "" Then intQ = 1
    Combo1.AddItem "英：" & rsQuery.Fields("lc06"), 1
    If rsQuery.Fields("lc06") <> "" Then intQ = 2
    Combo1.AddItem "日：" & rsQuery.Fields("lc07"), 2
    If rsQuery.Fields("lc07") <> "" Then intQ = 3
    Combo1.ListIndex = intQ - 1
    
    '本所案號
    lblData(0).Caption = "" & rsQuery.Fields("caseno")
    m_CP01 = "" & rsQuery.Fields("cp01")
    m_CP02 = "" & rsQuery.Fields("cp02")
    m_CP03 = "" & rsQuery.Fields("cp03")
    m_CP04 = "" & rsQuery.Fields("cp04")
    '收文號
    lblData(1).Caption = "" & rsQuery.Fields("cp09")
    '當事人
    lblData(4).Caption = "" & rsQuery.Fields("lc11")
    lblFM2(1).Caption = "" & rsQuery.Fields("lc11n")
    '收文日期
    lblData(2).Caption = ChangeWStringToTDateString("" & rsQuery.Fields("cp05"))
    '智權人員
    lblData(3).Caption = "" & rsQuery.Fields("cp13")
    lblFM2(0).Caption = "" & rsQuery.Fields("cp13n")

End Sub

Private Sub textCP06_GotFocus(Index As Integer)
    TextInverse textCP06(Index)
End Sub

Private Sub textCP06_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textCP06_Validate(Index As Integer, Cancel As Boolean)
   If textCP06(Index) <> "" Then
      If CheckIsTaiwanDate(textCP06(Index)) Then
         '倒推工作日
         strExc(1) = PUB_GetWorkDay1(DBDATE(textCP06(Index)), True)
         textCP06(Index) = TransDate(strExc(1), 1)
      Else
         GoTo EXITSUB
      End If
   ElseIf textCP06(Index).Visible = True Then
      MsgBox "本所期限不可空白", vbCritical
      GoTo EXITSUB
   End If
   Exit Sub
   
EXITSUB:
   Cancel = True
   textCP06(Index).SetFocus
   textCP06_GotFocus Index
End Sub
