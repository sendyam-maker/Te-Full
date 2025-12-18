VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020108 
   BorderStyle     =   1  '單線固定
   Caption         =   "主管機關來電處理記錄"
   ClientHeight    =   5604
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9312
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5604
   ScaleWidth      =   9312
   Begin VB.CommandButton cmdInput 
      Caption         =   "多案案號輸入"
      Height          =   400
      Left            =   5610
      TabIndex        =   37
      Top             =   105
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   8
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   13
      Top             =   2820
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   7
      Left            =   1290
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2478
      Width           =   405
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   375
      Left            =   6840
      TabIndex        =   27
      Top             =   1802
      Width           =   1935
      Begin VB.OptionButton Opt1 
         Caption         =   "小姐"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   50
         Width           =   735
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "先生"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   50
         Width           =   735
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "本所案號："
      Height          =   180
      Left            =   4215
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   795
      Width           =   1230
   End
   Begin VB.OptionButton Option1 
      Caption         =   "審定號/申請案號："
      Height          =   180
      Left            =   405
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   795
      Value           =   -1  'True
      Width           =   1830
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   690
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   5
      Left            =   1290
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1839
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   6
      Left            =   5115
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1839
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   1
      Top             =   735
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7470
      TabIndex        =   16
      Top             =   105
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8310
      TabIndex        =   17
      Top             =   105
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   7155
      MaxLength       =   2
      TabIndex        =   6
      Top             =   735
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   6915
      MaxLength       =   1
      TabIndex        =   5
      Top             =   735
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   6075
      MaxLength       =   6
      TabIndex        =   4
      Top             =   735
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   5595
      MaxLength       =   3
      TabIndex        =   3
      Top             =   735
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frm020108.frx":0000
      Height          =   1530
      Left            =   135
      TabIndex        =   15
      Top             =   3960
      Width           =   8925
      _ExtentX        =   15748
      _ExtentY        =   2709
      _Version        =   393216
      Cols            =   18
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   18
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   5
      Left            =   1290
      TabIndex        =   36
      Top             =   2218
      Width           =   1455
      VariousPropertyBits=   27
      Caption         =   "lblFM2(5)"
      Size            =   "2566;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   3
      Left            =   5130
      TabIndex        =   35
      Top             =   2218
      Width           =   1455
      VariousPropertyBits=   27
      Caption         =   "lblFM2(3)"
      Size            =   "2566;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   2
      Left            =   1290
      TabIndex        =   34
      Top             =   1542
      Width           =   1455
      VariousPropertyBits=   27
      Caption         =   "lblFM2(2)"
      Size            =   "2566;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   1
      Left            =   5100
      TabIndex        =   33
      Top             =   1542
      Width           =   1455
      VariousPropertyBits=   27
      Caption         =   "lblFM2(1)"
      Size            =   "2566;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1290
      TabIndex        =   32
      Top             =   1200
      Width           =   7785
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13732;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCont 
      Height          =   690
      Left            =   150
      TabIndex        =   14
      Top             =   3180
      Width           =   8940
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "15769;1217"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "副本給最後承辦人(非操作者)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   2
      Left            =   2440
      TabIndex        =   31
      Top             =   300
      Width           =   3195
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "回覆期限："
      Height          =   180
      Index           =   13
      Left            =   240
      TabIndex        =   30
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(1.補正 2.放棄專用權 3.延期 4.檢送同意書)"
      Height          =   180
      Index           =   12
      Left            =   1800
      TabIndex        =   29
      Top             =   2538
      Width           =   3315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "下一程序："
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   28
      Top             =   2556
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "來電內容："
      Height          =   180
      Index           =   0
      Left            =   8130
      TabIndex        =   26
      Top             =   2880
      Width           =   900
   End
   Begin VB.Shape Shape1 
      Height          =   540
      Left            =   135
      Top             =   620
      Width           =   8880
   End
   Begin VB.Label Label1 
      Caption         =   "存檔時會自動發郵件：正本給智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   11
      Left            =   180
      TabIndex        =   25
      Top             =   60
      Width           =   5475
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   10
      Left            =   225
      TabIndex        =   24
      Top             =   2232
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "最後承辦人："
      Height          =   180
      Index           =   8
      Left            =   4005
      TabIndex        =   23
      Top             =   2218
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "分機號碼："
      Height          =   180
      Index           =   7
      Left            =   225
      TabIndex        =   22
      Top             =   1908
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "來電人員："
      Height          =   180
      Index           =   6
      Left            =   4005
      TabIndex        =   21
      Top             =   1899
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   5
      Left            =   225
      TabIndex        =   20
      Top             =   1584
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "分所案號："
      Height          =   180
      Index           =   3
      Left            =   4005
      TabIndex        =   19
      Top             =   1579
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   4
      Left            =   225
      TabIndex        =   18
      Top             =   1260
      Width           =   900
   End
End
Attribute VB_Name = "frm020108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/07 改成Form2.0 ; Combo1、Text1(9)=>textCont、Label2(index)=>lblFM2(index)
'Created by Lydia 2015/11/10 (商標處)主管機關來電處理記錄
Option Explicit

Dim m_CP13 As String '智權人員
Dim m_LCP14 As String  '最後承辦人
Dim m_CaseName As String
Dim m_NP07 As String
Dim m_NP07CPM As String
Dim m_CaseType As String '審定號/申請案號
Dim bolClose As Boolean '是否閉卷
Dim oLabel As Control
Dim oText As Control
Dim m_TM12 As String, m_TM15 As String 'Added by Lydia 2022/09/27 申請案號、審定號

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         If bolClose Then
            MsgBox "此案件己閉卷，請再請認!", vbCritical
            Exit Sub
         End If
         cmdok(0).Enabled = False
         Me.cmdInput.Enabled = False 'Added by Lydia 2022/09/27
         If TxtValidate = True Then
            If FormSave = True Then
               PUB_SendMailCache
               FormClear True
               If Option1.Value Then
                  Text1(0).SetFocus
               Else
                  Text1(1).SetFocus
               End If
            Else
               cmdok(0).Enabled = True
               Me.cmdInput.Enabled = True 'Added by Lydia 2022/09/27
            End If
         Else
            cmdok(0).Enabled = True
            Me.cmdInput.Enabled = True 'Added by Lydia 2022/09/27
         End If
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Command1_Click()
   If GetCaseData = False Then
      If Option1.Value Then
         Text1(0).SetFocus
         Text1_GotFocus 0
      Else
         Text1(1).SetFocus
         Text1_GotFocus 1
      End If
   Else
      cmdok(0).Enabled = True
      Me.cmdInput.Enabled = True 'Added by Lydia 2022/09/27
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   FormClear True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If MsgBox("是否確定要結束？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm020108 = Nothing
End Sub

Private Sub Option1_Click()
   Text1(1).Enabled = False
   Text1(2).Enabled = False
   Text1(3).Enabled = False
   Text1(4).Enabled = False
   Text1(0).Enabled = True
   Text1(0).SetFocus
End Sub

Private Sub Option2_Click()
   Text1(1).Enabled = True
   Text1(2).Enabled = True
   Text1(3).Enabled = True
   Text1(4).Enabled = True
   Text1(0).Enabled = False
   Text1(1).SetFocus
End Sub

Private Sub Text1_Change(Index As Integer)
   If Index < 5 And Text1(Index).Tag <> "" Then
      FormClear False
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 5
          If Text1(Index) = "" Then
             MsgBox "請輸入分機號碼!", vbCritical
             GoTo ExitMode
          End If
      Case 6
          If Text1(Index) = "" Then
             MsgBox "請輸入來電人員!", vbCritical
             GoTo ExitMode
          End If
      Case 7
          'Modified by Lydia 2015/12/16
'          If Text1(Index) = "" Or (Text1(Index) <> "" And InStr("1,2,3", Text1(Index)) = 0) Then
'             MsgBox "下一程序請輸入1-3!", vbCritical
          If Text1(Index) = "" Or (Text1(Index) <> "" And InStr("1,2,3,4", Text1(Index)) = 0) Then
             MsgBox "下一程序請輸入1-4!", vbCritical
             GoTo ExitMode
          End If
      Case 8 '回覆期限：不可小於系統日且必須為工作日
          If Text1(Index) = "" Then
             MsgBox "請輸入回覆日期!", vbCritical
             GoTo ExitMode
          ElseIf CheckIsTaiwanDate(Text1(Index)) = False Then
               GoTo ExitMode
          ElseIf Text1(Index) < strSrvDate(2) Then
                MsgBox "回覆日期不可小於系統日!", vbCritical
                GoTo ExitMode
          ElseIf ChkWorkDay(ChangeTStringToWString(Text1(Index))) = False Then
                  MsgBox "回覆日期必須為工作天!", vbCritical
                  GoTo ExitMode
          End If
      'Remove by Lydia 2021/10/07 Text1(9) => textCont
      'Case 9
      '    If Text1(Index) = "" Then
      '       MsgBox "請輸入來電內容!", vbCritical
      '       GoTo ExitMode
      '    End If
      'end 2021/10/07
   End Select
   
   Exit Sub
   
ExitMode:
   Text1(Index).SetFocus
   Cancel = True

End Sub

Private Function FormSave() As Boolean
   Dim strCP09 As String, strCP12 As String
   Dim strCP64 As String, strDate As String
   Dim strReceiver As String, strCC As String
   Dim strNP22 As String
   'Added by Lydia 2022/09/27
   Dim strCase(1 To 4) As String
   Dim tmpArr As Variant
   Dim intR As Integer
   Dim tmpCP13 As String, tmpLCP14 As String
   'end 2022/09/27
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

   strCP09 = AutoNo("C", 6)
   strDate = ChangeTStringToWString(Text1(8))
   strCP12 = GetSalesArea(m_CP13)
   strCP64 = "來電人員：" & Text1(6) & " " & IIf(opt1(0).Value = True, "先生", "小姐") & ", 分機號碼：" & Text1(5) & ", 來電內容：" & textCont
   
   Select Case Text1(7).Text
   'Modified by Lydia 2015/12/16 嘉雯提出修改
'       Case "1"
'            m_NP07 = "313": m_NP07CPM = "減縮商品"
'       Case "2"
'            m_NP07 = "201": m_NP07CPM = "補正"
'       Case "3"
'            m_NP07 = "206": m_NP07CPM = "放棄專用權"
       Case "1"
            m_NP07 = "201": m_NP07CPM = "補正"
       Case "2"
            m_NP07 = "206": m_NP07CPM = "放棄專用權"
       Case "3"
            m_NP07 = "303": m_NP07CPM = "延期"
       Case "4"
            m_NP07 = "211": m_NP07CPM = "檢送同意書"
   End Select

  '存檔時新增C類來函，收文日=發文日=系統日，案件性質=電話通知1727；
  'CP13：T案以PUB_GetAKindSalesNo抓智權人員，FCT案以PUB_GetFCTSalesNo抓業務承辦；CP14為操作人員，本所期限=法定期限=回覆期限
   strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
      "CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP27,cp64) VALUES " & _
      "('" & Text1(1) & "','" & Text1(2) & "','" & Text1(3) & "','" & Text1(4) & "'," & strSrvDate(1) & "," & strDate & "," & strDate & _
      ",'" & strCP09 & "','1727'," & CNULL(strCP12) & "," & CNULL(m_CP13) & _
      ",'" & strUserNum & "','N','N','N'," & strSrvDate(1) & ",'" & ChgSQL(strCP64) & "') "
   cnnConnection.Execute strSql, intI
  '新增下一程序檔，本所期限=法定期限=回覆期限
   strNP22 = GetNextProgressNo
   strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) VALUES ('" & strCP09 & "','" & Text1(1) & "','" & Text1(2) & "','" & Text1(3) & "','" & Text1(4) & _
         "'," & m_NP07 & "," & strDate & "," & strDate & "," & CNULL(m_CP13) & "," & strNP22 & ")"
   cnnConnection.Execute strSql, intI
   
   strReceiver = m_CP13
   If strUserNum <> m_LCP14 Then
      strCC = m_LCP14
   End If

   strExc(1) = Text1(1) & "-" & Text1(2) & IIf(Text1(3) & Text1(4) = "000", "", "-" & Text1(3) & "-" & Text1(4)) & " (" & strCP09 & ") 主管機關來電通知"
   'Modified by Lydia 2022/09/27 改模組
   'strExc(2) = GetMailText
   strExc(2) = GetMailText_New
   strExc(3) = PUB_LeftB(strExc(2), 4000)
   strExc(4) = Mid(strExc(2), Len(strExc(3)) + 1)
   strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc10,mc11)" & _
      " values ('" & strUserNum & "','" & strReceiver & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
      ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(3)) & "','" & strCC & "','Y','" & ChgSQL(strExc(4)) & "')"
   cnnConnection.Execute strSql, intI
   
   'Added by Lydia 2022/09/27 多案案號輸入
   If Me.cmdInput.Tag <> "" Then
       tmpArr = Split(Me.cmdInput.Tag, ",")
       For intR = 0 To UBound(tmpArr)
           If Trim(tmpArr(intR)) <> "" Then
               Sleep 1000
               Call ChgCaseNo(tmpArr(intR), strCase)
               If strCase(1) <> "" And strCase(2) <> "" Then
                    '存檔時新增C類來函，收文日=發文日=系統日，案件性質=電話通知1727；
                    strCP09 = AutoNo("C", 6)
                    strExc(2) = GetMailText_New(strCase(1), strCase(2), strCase(3), strCase(4), tmpCP13, tmpLCP14)
                    strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
                    "CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP27,cp64) VALUES " & _
                    "('" & strCase(1) & "','" & strCase(2) & "','" & strCase(3) & "','" & strCase(4) & "'," & strSrvDate(1) & "," & strDate & "," & strDate & _
                    ",'" & strCP09 & "','1727'," & CNULL(strCP12) & "," & CNULL(tmpCP13) & _
                    ",'" & strUserNum & "','N','N','N'," & strSrvDate(1) & ",'" & ChgSQL(strCP64) & "') "
                    cnnConnection.Execute strSql, intI
                    '新增下一程序檔，本所期限=法定期限=回覆期限
                    strNP22 = GetNextProgressNo
                    strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) VALUES ('" & strCP09 & "','" & strCase(1) & "','" & strCase(2) & "','" & strCase(3) & "','" & strCase(4) & _
                              "'," & m_NP07 & "," & strDate & "," & strDate & "," & CNULL(tmpCP13) & "," & strNP22 & ")"
                    cnnConnection.Execute strSql, intI
                    
                    strExc(1) = strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) = "000", "", "-" & strCase(3) & "-" & strCase(4)) & " (" & strCP09 & ") 主管機關來電通知"
                    strReceiver = tmpCP13
                    strCC = ""
                    If strUserNum <> tmpLCP14 Then
                        strCC = tmpLCP14
                    End If

                    strExc(3) = PUB_LeftB(strExc(2), 4000)
                    strExc(4) = Mid(strExc(2), Len(strExc(3)) + 1)
                    strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc10,mc11)" & _
                       " values ('" & strUserNum & "','" & strReceiver & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                       ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(3)) & "','" & strCC & "','Y','" & ChgSQL(strExc(4)) & "')"
                    cnnConnection.Execute strSql, intI
               End If
           End If
       Next intR
   End If
   'end 2022/09/27

   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
   
End Function

Private Function GetCaseData() As Boolean
   Dim bolFound As Boolean
   Dim ii As Integer
   
   m_CaseName = ""
   m_CaseType = ""
   m_TM12 = "": m_TM15 = ""  'Added by Lydia 2022/09/27
   
   If Option1.Value Then
       strExc(0) = "select '1' ord1 ,tm01,tm02,tm03,tm04,tm05,tm06,tm07,tm15,tm12, tm34,nvl(cu04,nvl(cu06,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))) CuName,tm29 " & _
                   "from trademark,customer where substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and tm10='000' and tm15='" & Trim(Text1(0).Text) & "' " & _
                   "Union select '2' ord1 ,tm01,tm02,tm03,tm04,tm05,tm06,tm07,tm15,tm12,tm34,nvl(cu04,nvl(cu06,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))) CuName,tm29 " & _
                   "from trademark,customer where substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and tm10='000' and tm12='" & Trim(Text1(0).Text) & "' order by 1"
   Else
      If Text1(3) = "" Then Text1(3) = "0"
      If Text1(4) = "" Then Text1(4) = "00"
      strExc(0) = "select '0' ord1 ,tm01,tm02,tm03,tm04,tm05,tm06,tm07,tm15,tm12,tm34,tm23,nvl(cu04,nvl(cu06,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))) CuName,tm29 " & _
                 "from trademark,customer where substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and tm10='000' and tm01='" & Trim(Text1(1).Text) & "' and tm02='" & Trim(Text1(2).Text) & "' and tm03='" & Trim(Text1(3).Text) & "' and tm04='" & Trim(Text1(4).Text) & "' "
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then bolFound = True
   
   If Not bolFound Then
      MsgBox IIf(Option1.Value = True, "審定號/申請案號", "本所案號") & "輸入錯誤！", vbExclamation
   Else
      With RsTemp
      
      If Option2.Value Then Text1(0) = "" & IIf(IsNull(.Fields("tm15")), .Fields("tm12"), .Fields("tm15"))
      If Not IsNull(.Fields("tm15")) Then
          m_CaseType = "審定號"
      Else
          m_CaseType = "申請案號"
      End If
      'Added by Lydia 2022/09/27
      m_TM12 = "" & .Fields("tm12")
      m_TM15 = "" & .Fields("tm15")
      'end 2022/09/27
      
      Text1(1) = "" & .Fields("tm01")
      Text1(1).Tag = Text1(1)
      Text1(2) = "" & .Fields("tm02")
      Text1(2).Tag = Text1(2)
      Text1(3) = "" & .Fields("tm03")
      Text1(3).Tag = Text1(3)
      Text1(4) = "" & .Fields("tm04")
      Text1(4).Tag = Text1(4)
      Text1(0).Tag = Text1(0)
      
      If Not IsNull(.Fields("tm29")) Then
         MsgBox "此案件己閉卷，請再請認!", vbCritical
         bolClose = True
         Exit Function
      End If
      
      If Not IsNull(.Fields("tm05")) Then
         m_CaseName = .Fields("tm05")
      ElseIf Not IsNull(.Fields("tm06")) Then
         m_CaseName = .Fields("tm06")
      ElseIf Not IsNull(.Fields("tm07")) Then
         m_CaseName = .Fields("tm07")
      End If
      AddCboName Combo1, "" & .Fields("tm05"), "" & .Fields("tm06"), "" & .Fields("tm07")

      lblFM2(1) = "" & .Fields("tm34")
      lblFM2(2) = "" & .Fields("CuName")
      End With

      '最後承辦人: 不可修改
       lblFM2(3).Caption = "": m_LCP14 = ""
        '抓該案號之最後承辦人非程序(ST03<>'P22')者，若離職則抓部門主管
        'Added by Lydia 2023/12/26
        If strSrvDate(1) >= 新部門啟用日 Then
            strExc(0) = "select cp14,s1.st04 stype,s1.st02 s1name,nvl(a0924,a0909) s2no,getstaffnamelist(nvl(a0924,a0909)) as s2name  " & _
                        "from caseprogress,acc090,acc090new,staff s1 where cp01='" & Text1(1) & "' and cp02='" & Text1(2) & "' and cp03='" & Text1(3) & "' and cp04='" & Text1(4) & "' " & _
                        "and cp14=s1.st01(+) and s1.st03=a0901(+) and s1.st93=a0921(+) and cp57 is null and s1.st03 <>'P22' order by cp05 desc,cp09 desc"
        Else
        'end 2023/12/26
            strExc(0) = "select cp14,s1.st04 stype,s1.st02 s1name,a0909 s2no,s2.st02 s2name " & _
                        "from caseprogress,acc090,staff s1,staff s2 where cp01='" & Text1(1) & "' and cp02='" & Text1(2) & "' and cp03='" & Text1(3) & "' and cp04='" & Text1(4) & "' " & _
                        "and cp14=s1.st01(+) and s1.st03=a0901(+) and a0909=s2.st01 and cp57 is null and s1.st03 <>'P22' order by cp05 desc,cp09 desc"
        End If
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
          If RsTemp("stype") = "1" Then
             m_LCP14 = "" & RsTemp("cp14")
             lblFM2(3).Caption = "" & RsTemp("s1name")
          Else
             m_LCP14 = "" & RsTemp("s2no")
             lblFM2(3).Caption = "" & RsTemp("s2name")
          End If
        End If
      
      '智權人員: 不可修改
      lblFM2(5).Caption = "": m_CP13 = ""
      Select Case Trim(Text1(1).Text)
          Case "T"
              'T案，則以PUB_GetAKindSalesNo抓智權人員
              m_CP13 = PUB_GetAKindSalesNo(Text1(1), Text1(2), Text1(3), Text1(4))
          Case "FCT"
              'FCT案，則以PUB_GetFCTSalesNo抓智權人員
              m_CP13 = PUB_GetFCTSalesNo(Text1(1), Text1(2), Text1(3), Text1(4))
      End Select
      lblFM2(5).Caption = GetStaffName(m_CP13)
      MSHFlexGrid1.Visible = False

      strExc(0) = "select ' ' AS V,SQLDATET2(CP05) as 收文日,CP09 as 總收文號,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10) as 案件性質" & _
         ",NVL(S1.ST02,CP14) as 承辦人,NVL(S2.ST02,CP13) as 智權人員" & _
         ",SQLDATET2(CP06) as 本所期限,SQLDATET2(CP07) as 法定期限,SQLDATET2(CP27) as 發文日,SQLDATET2(CP57) as 取消收文日" & _
         " from caseprogress,trademark,casepropertymap,staff s1,staff s2" & _
         " WHERE cp01='" & Text1(1) & "' and cp02='" & Text1(2) & "' and cp03='" & Text1(3) & "' and cp04='" & Text1(4) & "'" & _
         " and TM01(+)=cp01 and TM02(+)=cp02 and TM03(+)=cp03 and TM04(+)=cp04" & _
         " and cp01=cpm01(+) and cp10=cpm02(+) and s1.st01(+)=cp14 and s2.st01(+)=cp13" & _
         " ORDER BY SQLDatet2(CP05) DESC,CP66 DESC,CP67 DESC,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3'),CP09 DESC"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      GridHead
      If intI = 1 Then
         With MSHFlexGrid1
         For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 4) <> "" Then
                .TextMatrix(ii, 3) = .TextMatrix(ii, 3) & PUB_GetRelateCasePropertyName(.TextMatrix(ii, 2), "1")
            End If
         Next
         
         End With
      End If
      MSHFlexGrid1.Visible = True
      
      GetCaseData = True
   End If
End Function

Private Sub FormClear(Optional pbolAll As Boolean)

   If pbolAll Then
        Text1(1).Text = ""
        Text1(2).Text = ""
        Text1(3).Text = ""
        Text1(4).Text = ""
        Text1(0) = ""
   Else
      If Option1.Value Then
        Text1(1).Text = ""
        Text1(2).Text = ""
        Text1(3).Text = ""
        Text1(4).Text = ""
      Else
         Text1(0) = ""
      End If
   End If
   
   bolClose = False

   For Each oText In Text1
       oText.Tag = ""
       If oText.Index > 4 Then
          oText.Text = ""
       End If
   Next
   textCont.Text = "": textCont.Tag = ""  'Added by Lydia 2021/10/07
   Text1(0).Tag = ""
   Combo1.Clear
   For Each oLabel In lblFM2
      oLabel.Caption = ""
   Next
   MSHFlexGrid1.Clear
   MSHFlexGrid1.Rows = 2
   GridHead
   Me.cmdok(0).Enabled = False
   'Added by Lydia 2022/09/27
   Me.cmdInput.Tag = ""
   Me.cmdInput.Enabled = False
   'end 2022/09/27
End Sub

Private Function GetMailText() As String
   Dim strText As String
   '要有 &nbsp; 字串空白才不會被再轉換一次
   strText = ""
   strText = strText & "<TABLE BORDER CELLSPACING=2 CELLPADDING=2 WIDTH=600 STYLE=""border:2px solid;"">"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" COLSPAN=4 HEIGHT=37>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=5><P ALIGN=""CENTER"">主&nbsp;管&nbsp;機&nbsp;關&nbsp;來&nbsp;電&nbsp;通&nbsp;知&nbsp;單"
   strText = strText & "</FONT></TD></TR>"
   strText = strText & "<TR><TD WIDTH=""20%"" VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">日　期</FONT></TD>"
   strText = strText & "<TD WIDTH=""30%"" VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & ChangeTStringToTDateString(strSrvDate(2)) & "</FONT></TD>"
   strText = strText & "<TD WIDTH=""15%"" VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">接話人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & strUserName & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">" & m_CaseType & " </FONT></TD>" '審定號/申請案號
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text1(0) & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">申請人</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & lblFM2(2) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">本所案號</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text1(1) & "-" & Text1(2) & IIf(Text1(3) & Text1(4) = "000", "", "-" & Text1(3) & "-" & Text1(4)) & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P>分所案號</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & IIf(lblFM2(1) = "", "　", lblFM2(1)) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">案件名稱</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"" COLSPAN=3>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & m_CaseName & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">來電人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text1(6) & " " & IIf(opt1(0).Value = True, "先生", "小姐") & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">分機號碼</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text1(5) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">最後承辦人</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & lblFM2(3) & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">智權人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & GetStaffName(m_CP13) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">下一程序</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & m_NP07CPM & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">回覆期限</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & ChangeTStringToTDateString(Text1(8)) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""TOP"" COLSPAN=4 HEIGHT=200>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P>來電內容：<BR>" & textCont & "</FONT></TD></TR>"
   strText = strText & "</TABLE>"
   strText = strText & "<TABLE WIDTH=600>"
   strText = strText & "</TABLE>"
   
   GetMailText = strText
End Function

Private Sub GridHead()
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .Cols = 11
      .row = 0
      .col = 0: .Text = "V"
      .ColWidth(0) = 0
      .col = 1: .Text = "收文日"
      .ColWidth(1) = 800
      .CellAlignment = flexAlignRightCenter
      .col = 2: .Text = "總收文號"
      .ColWidth(2) = 1000
      .CellAlignment = flexAlignLeftCenter
      .col = 3: .Text = "案件性質"
      .ColWidth(3) = 2000
      .CellAlignment = flexAlignLeftCenter
      .col = 4: .Text = "承辦人"
      .ColWidth(4) = 650
      .CellAlignment = flexAlignLeftCenter
      .col = 5: .Text = "智權人員"
      .ColWidth(5) = 650
      .CellAlignment = flexAlignLeftCenter
      .col = 6: .Text = "本所期限"
      .ColWidth(6) = 820
      .CellAlignment = flexAlignRightCenter
      .col = 7: .Text = "法定期限"
      .ColWidth(7) = 820
      .CellAlignment = flexAlignRightCenter
      .col = 8: .Text = "發文日"
      .ColWidth(8) = 800
      .CellAlignment = flexAlignRightCenter
      .col = 9: .Text = "取消收文日"
      .ColWidth(9) = 1000
      .CellAlignment = flexAlignLeftCenter
      .Visible = True
   End With
End Sub

Private Function TxtValidate() As Boolean
Dim tmpBol As Boolean
Dim jj As Integer

   jj = 0
   For Each oText In Text1
      Text1_Validate jj, tmpBol
      If tmpBol = True Then
         Exit Function
      End If
      jj = jj + 1
   Next
      
   'Added by Lydia 2021/10/07
   textCont_Validate tmpBol
   If tmpBol = True Then Exit Function
   'Added by Lydia 2021/10/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         Exit Function
   End If
   'end 2021/10/07
   
   If opt1(0).Value = 0 And opt1(1).Value = 0 Then
       MsgBox "請選擇來電人員的稱謂!", vbExclamation
       opt1(0).Value = 1
       Exit Function
   End If
   TxtValidate = True
   
End Function

'Added by Lydia 2021/10/07
Private Sub textCont_GotFocus()
    TextInverse textCont
End Sub
'Added by Lydia 2021/10/07
Private Sub textCont_Validate(Cancel As Boolean)
     If textCont.Text = "" Then
         MsgBox "請輸入來電內容!", vbCritical
         textCont.SetFocus
         Cancel = True
     End If
     
End Sub

'Added by Lydia 2022/09/27 多案案號輸入：與FCT->電話通知-輸入期限frm030209_02共用
Private Sub cmdInput_Click()
   Set frm880004.mPreForm = Me
   frm880004.iStiu = 8
   frm880004.m_LCV01 = Text1(1) & Text1(2) & Text1(3) & Text1(4) & "," & m_TM15 & "," & m_TM12
   frm880004.m_TempList = Me.cmdInput.Tag
   frm880004.Show vbModal
End Sub

'Added by Lydia 2022/09/27 配合多案案號輸入,可傳入案號
Private Function GetMailText_New(Optional ByVal pCP01 As String, Optional ByVal pCP02 As String, Optional ByVal pCP03 As String, Optional ByVal pCP04 As String, Optional ByRef pCP13 As String, Optional ByRef pLCP14 As String) As String
Dim strText As String
Dim strTemp(1 To 8) As String
   
   If pCP01 = "" And pCP02 = "" Then '直接用畫面的資料
       strTemp(1) = m_CaseType
       strTemp(2) = Text1(0)
       strTemp(3) = lblFM2(2) '申請人
       strTemp(4) = Text1(1) & "-" & Text1(2) & IIf(Text1(3) & Text1(4) = "000", "", "-" & Text1(3) & "-" & Text1(4))   '本所案號
       strTemp(5) = IIf(lblFM2(1) = "", "　", lblFM2(1)) '分所案號
       strTemp(6) = m_CaseName '案件名稱
       strTemp(7) = lblFM2(3)  '最後承辦人
       strTemp(8) = lblFM2(5) '智權人員
   Else
       strExc(0) = "select '0' ord1 ,tm01,tm02,tm03,tm04,nvl(tm05,nvl(tm06,tm07)) casename,tm15,tm12,tm34,tm23,nvl(cu04,nvl(cu06,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))) CuName,tm29 " & _
                  "from trademark,customer where substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and tm10='000' and tm01='" & pCP01 & "' and tm02='" & pCP02 & "' and tm03='" & pCP03 & "' and tm04='" & pCP04 & "' "
       pCP13 = "": pLCP14 = ""
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
            If "" & RsTemp.Fields("TM15") <> "" Then
                 strTemp(1) = "審定號"
                 strTemp(2) = "" & RsTemp.Fields("tm15")
            Else
                 strTemp(1) = "申請案號"
                 strTemp(2) = "" & RsTemp.Fields("tm12")
            End If
            strTemp(3) = "" & RsTemp.Fields("CuName")
            strTemp(4) = pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 = "000", "", "-" & pCP03 & "-" & pCP04)
            strTemp(5) = "" & RsTemp.Fields("tm34")
            strTemp(6) = "" & RsTemp.Fields("casename")
       End If
       '抓該案號之最後承辦人非程序(ST03<>'P22')者，若離職則抓部門主管
        strExc(0) = "select cp14,s1.st04 stype,s1.st02 s1name,a0909 s2no,s2.st02 s2name " & _
                    "from caseprogress,acc090,staff s1,staff s2 where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' " & _
                    "and cp14=s1.st01(+) and s1.st03=a0901(+) and a0909=s2.st01 and cp57 is null and s1.st03 <>'P22' order by cp05 desc,cp09 desc"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            If "" & RsTemp("stype") = "1" Then
                pLCP14 = "" & RsTemp.Fields("cp14")
                strTemp(7) = "" & RsTemp.Fields("s1name")
            Else
                pLCP14 = "" & RsTemp.Fields("s2no")
                strTemp(7) = "" & RsTemp.Fields("s2name")
            End If
        End If
        '智權人員: 不可修改
        strExc(0) = "": strTemp(8) = ""
        Select Case pCP01
           Case "T"
              strExc(0) = PUB_GetAKindSalesNo(pCP01, pCP02, pCP03, pCP04)
           Case "FCT"
              strExc(0) = PUB_GetFCTSalesNo(pCP01, pCP02, pCP03, pCP04)
        End Select
        If strExc(0) <> "" Then
            pCP13 = strExc(0)
            strTemp(8) = GetStaffName(strExc(0))
        End If
   End If
   
   '要有 &nbsp; 字串空白才不會被再轉換一次
   strText = ""
   strText = strText & "<TABLE BORDER CELLSPACING=2 CELLPADDING=2 WIDTH=600 STYLE=""border:2px solid;"">"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" COLSPAN=4 HEIGHT=37>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=5><P ALIGN=""CENTER"">主&nbsp;管&nbsp;機&nbsp;關&nbsp;來&nbsp;電&nbsp;通&nbsp;知&nbsp;單"
   strText = strText & "</FONT></TD></TR>"
   strText = strText & "<TR><TD WIDTH=""20%"" VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">日　期</FONT></TD>"
   strText = strText & "<TD WIDTH=""30%"" VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & ChangeTStringToTDateString(strSrvDate(2)) & "</FONT></TD>"
   strText = strText & "<TD WIDTH=""15%"" VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">接話人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & strUserName & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">" & strTemp(1) & " </FONT></TD>" '審定號/申請案號
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & strTemp(2) & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">申請人</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & strTemp(3) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">本所案號</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & strTemp(4) & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P>分所案號</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & strTemp(5) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">案件名稱</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"" COLSPAN=3>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & strTemp(6) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">來電人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text1(6) & " " & IIf(opt1(0).Value = True, "先生", "小姐") & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">分機號碼</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & Text1(5) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">最後承辦人</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & strTemp(7) & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">智權人員</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & strTemp(8) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""MIDDLE"" HEIGHT=30>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">下一程序</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & m_NP07CPM & "</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""CENTER"">回覆期限</FONT></TD>"
   strText = strText & "<TD VALIGN=""MIDDLE"">"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P ALIGN=""LEFT"">" & ChangeTStringToTDateString(Text1(8)) & "</FONT></TD></TR>"
   strText = strText & "<TR><TD VALIGN=""TOP"" COLSPAN=4 HEIGHT=200>"
   strText = strText & "<FONT FACE=""標楷體"" SIZE=4><P>來電內容：<BR>" & textCont & "</FONT></TD></TR>"
   strText = strText & "</TABLE>"
   strText = strText & "<TABLE WIDTH=600>"
   strText = strText & "</TABLE>"
   
   GetMailText_New = strText
End Function
