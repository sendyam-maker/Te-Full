VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_d 
   Appearance      =   0  '平面
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文-異議舉發"
   ClientHeight    =   2280
   ClientLeft      =   -1212
   ClientTop       =   3156
   ClientWidth     =   10272
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   10272
   Begin VB.CommandButton Command1 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   8595
      TabIndex        =   10
      Top             =   135
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7770
      TabIndex        =   9
      Top             =   135
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Caption         =   "補件期限"
      Height          =   1425
      Left            =   120
      TabIndex        =   11
      Top             =   660
      Width           =   10005
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   0
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   23
         Top             =   330
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   1
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   22
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   2
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   21
         Top             =   930
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   2
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   7
         Top             =   930
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   1
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   4
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   0
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   1
         Top             =   330
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   2
         Left            =   960
         MaxLength       =   8
         TabIndex        =   6
         Top             =   930
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   960
         MaxLength       =   8
         TabIndex        =   3
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   960
         MaxLength       =   8
         TabIndex        =   0
         Top             =   330
         Width           =   855
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Index           =   2
         Left            =   6420
         TabIndex        =   8
         Top             =   930
         Width           =   3435
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6059;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Index           =   1
         Left            =   6420
         TabIndex        =   5
         Top             =   630
         Width           =   3435
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6059;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Index           =   0
         Left            =   6420
         TabIndex        =   2
         Top             =   330
         Width           =   3435
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6059;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "約定期限"
         Height          =   180
         Index           =   2
         Left            =   3720
         TabIndex        =   26
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "約定期限"
         Height          =   180
         Index           =   1
         Left            =   3720
         TabIndex        =   25
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "約定期限"
         Height          =   180
         Index           =   0
         Left            =   3720
         TabIndex        =   24
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補件內容"
         Height          =   180
         Index           =   11
         Left            =   5550
         TabIndex        =   20
         Top             =   930
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補件內容"
         Height          =   180
         Index           =   10
         Left            =   5550
         TabIndex        =   19
         Top             =   630
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補件內容"
         Height          =   180
         Index           =   9
         Left            =   5550
         TabIndex        =   18
         Top             =   330
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限"
         Height          =   180
         Index           =   8
         Left            =   1920
         TabIndex        =   17
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限"
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限"
         Height          =   180
         Index           =   6
         Left            =   1920
         TabIndex        =   15
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   630
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限"
         Height          =   180
         Index           =   4
         Left            =   1920
         TabIndex        =   13
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm060104_d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/17 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String, pa(1 To 5) As String, intWhere As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 設定本所期限及法定期限
' Input :  strDate1 ==> 本所期限
'          strDate2 ==> 法定期限
'          strDate3 ==> 約定期限 'Add By Sindy 2021/8/17
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetData(ByVal strDate1 As String, ByVal strDate2 As String, ByVal strDate3 As String)
   Dim nIndex As Integer
   For nIndex = 0 To 2
      Text1(nIndex) = TAIWANDATE(strDate1)
   Next nIndex
   For nIndex = 0 To 2
      Text2(nIndex) = TAIWANDATE(strDate2)
   Next nIndex
   'Add By Sindy 2021/8/17
   For nIndex = 0 To 2
      Text3(nIndex) = TAIWANDATE(strDate3)
   Next nIndex
   '2021/8/17 END
End Sub

Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      If CheckDataValid() = True Then
         'Add by Sindy 2021/11/17 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If

         If FormSave = False Then
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         End If
      Else
         Exit Sub
      End If
   End If
   frm060104_9.Show
   Unload Me
End Sub

Private Sub Form_Load()
 Dim i As Integer
   MoveFormToCenter Me
   intWhere = 國外_FC
   For i = 0 To 2
      Combo1(i).Clear
   Next
   With frm060104_9
      strReceiveNo = .Label2(0)
      pa(1) = .Text1
      pa(2) = .Text2
      pa(3) = .Text3
      pa(4) = .Text4
      'pa(5) = .Text6 'Removed by Morgan 2013/1/14 沒再用
   End With
   Select Case frm060104_9.Tag
      Case "1" '異議_專
         For i = 0 To 2
            Combo1(i).AddItem "異議理由書一式三份"
            Combo1(i).AddItem "代理人委任書正本"
            Combo1(i).AddItem "法人地位證明書正本"
         Next
      Case "2" '舉發
         For i = 0 To 2
            Combo1(i).AddItem "舉發理由書一式四份"
            Combo1(i).AddItem "代理人委任書正本"
            Combo1(i).AddItem "法人地位證明書正本"
         Next
   End Select
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_d = Nothing
End Sub

Private Function FormSave() As Boolean
 Dim i As Integer, intMax As Long, strTxt(1 To 5) As String, j As Integer
 
'911105 nick transation
FormSave = True
 On Error GoTo CheckingErr
cnnConnection.BeginTrans

   strTxt(1) = "DELETE FROM NEXTPROGRESS WHERE NP01='" & strReceiveNo & _
      "' AND NP07=" & 補文件 & " AND (NP06 IS NULL OR NP06='')"
      
   '911105 nick transation
   cnnConnection.Execute strTxt(1)
   
      'edit by nickc 2007/02/02 不用 dll 了
   'intMax = objPublicData.GetNextProgressNo
   intMax = GetNextProgressNo
   j = 2
   For i = 0 To 2
      ' 90.07.24 modify by louis
      'If Text1(i) <> "" And Text2(i) <> "" And Combo1(i) <> "" Then
      If IsEmptyText(Text1(i).Text) = False And IsEmptyText(Text2(i).Text) = False And IsEmptyText(Combo1(i).Text) = False Then
         'Modify By Cheng 2002/07/05
'         strTxt(j) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'            "NP09,NP10,NP15,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & _
'            "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 補文件 & _
'            "," & TransDate(Text1(i), 2) & "," & TransDate(Text2(i), 2) & ",'" & _
'            pa(5) & "','" & Combo1(i) & "'," & intMax & ")"
         '911105 nick
            'edit by nickc 2007/02/02 不用 dll 了
   'intMax = objPublicData.GetNextProgressNo
   intMax = GetNextProgressNo
         'Modify By Sindy 2021/8/17 + ,NP23
         strTxt(j) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
            "NP09,NP10,NP15,NP22,NP23) VALUES ('" & strReceiveNo & "','" & pa(1) & _
            "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 補文件 & _
            "," & TransDate(Text1(i), 2) & "," & TransDate(Text2(i), 2) & ",'" & _
            PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "','" & Combo1(i) & "'," & intMax & "," & TransDate(Text3(i), 2) & ")"
            
            '911105 nick transation
            cnnConnection.Execute strTxt(j)
            
            
         j = j + 1
         intMax = intMax + 1
      End If
   Next
   
   '911105 nick transation
   'FormSave = objLawDll.ExecSQL(j - 1, strTxt)
   cnnConnection.CommitTrans
'911105 nick
   Exit Function
CheckingErr:
   cnnConnection.RollbackTrans
   FormSave = False
   
End Function

Private Sub Text1_GotFocus(Index As Integer)
   InverseTextBox Text1(Index)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Text1(Index) <> "" Then
      If Not ChkDate(Text1(Index)) Then
         Cancel = True
         Text1_GotFocus Index
      'Added by Lydia 2025/11/13 改抓最近工作天
      Else
         Text1(Index) = TransDate(PUB_GetWorkDay1(DBDATE(Text1(Index)), True), 1)
      'end 2025/11/13
      End If
   End If
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   InverseTextBox Text2(Index)
End Sub

Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
   If IsEmptyText(Text2(Index)) = False Then
      If ChkDate(Text2(Index)) = False Then
         Cancel = True
         Text2_GotFocus Index
      End If
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   
   CheckDataValid = False
   For nIndex = 0 To 2
      If IsEmptyText(Text1(nIndex)) = True And IsEmptyText(Text2(nIndex)) = False Then
         strTit = "檢核資料"
         strMsg = "本所期限與法定期限必須同時有值!"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         Text1(nIndex).SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(Text1(nIndex)) = False And IsEmptyText(Text2(nIndex)) = True Then
         strTit = "檢核資料"
         strMsg = "本所期限與法定期限必須同時有值!"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         Text2(nIndex).SetFocus
         GoTo EXITSUB
      End If
      'Add By Sindy 2021/8/17
      If IsEmptyText(Text3(nIndex)) = True And IsEmptyText(Text2(nIndex)) = False Then
         strTit = "檢核資料"
         strMsg = "約定期限與法定期限必須同時有值!"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         Text3(nIndex).SetFocus
         GoTo EXITSUB
      End If
      '2021/8/17 END
      If IsEmptyText(Combo1(nIndex).Text) = False Then
         If IsEmptyText(Text1(nIndex)) = True Or IsEmptyText(Text2(nIndex)) = True Or IsEmptyText(Text3(nIndex)) = True Then
            strTit = "檢核資料"
            strMsg = "補件內容有設定,本所期限與法定期限與約定期限不可空白!"
            nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
            Text2(nIndex).SetFocus
            GoTo EXITSUB
         End If
      End If
      'Added by Lydia 2025/11/13
      If Text1(nIndex).Text <> "" Then
         Text1(nIndex).Text = TransDate(PUB_GetWorkDay1(DBDATE(Text1(nIndex)), True), 1)
      End If
      If Text3(nIndex).Text <> "" Then
         Text3(nIndex).Text = TransDate(PUB_GetWorkDay1(DBDATE(Text3(nIndex)), True), 1)
      End If
      'end 2025/11/13
   Next nIndex
   CheckDataValid = True
EXITSUB:
End Function

'Added by Lydia 2025/11/13
Private Sub Text3_GotFocus(Index As Integer)
   TextInverse Text3(Index)
End Sub

Private Sub Text3_Validate(Index As Integer, Cancel As Boolean)
   If Text3(Index) <> "" Then
      If Not ChkDate(Text3(Index)) Then
         Cancel = True
         Text3_GotFocus Index
      Else
         Text3(Index) = TransDate(PUB_GetWorkDay1(DBDATE(Text3(Index)), True), 1)
      End If
   End If
End Sub
