VERSION 5.00
Begin VB.Form frm040104_d 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文-異議舉發"
   ClientHeight    =   2160
   ClientLeft      =   -240
   ClientTop       =   3132
   ClientWidth     =   7212
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   7212
   Begin VB.CommandButton Command1 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   5868
      TabIndex        =   10
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5040
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Caption         =   "補件期限"
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   660
      Width           =   6975
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   2
         ItemData        =   "frm040104_d.frx":0000
         Left            =   4920
         List            =   "frm040104_d.frx":0002
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         ItemData        =   "frm040104_d.frx":0004
         Left            =   4920
         List            =   "frm040104_d.frx":0006
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   0
         ItemData        =   "frm040104_d.frx":0008
         Left            =   4920
         List            =   "frm040104_d.frx":000A
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   2
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   1
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   0
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   2
         Left            =   960
         MaxLength       =   8
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   960
         MaxLength       =   8
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   960
         MaxLength       =   8
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補件內容"
         Height          =   180
         Index           =   11
         Left            =   3960
         TabIndex        =   20
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補件內容"
         Height          =   180
         Index           =   10
         Left            =   3960
         TabIndex        =   19
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補件內容"
         Height          =   180
         Index           =   9
         Left            =   3960
         TabIndex        =   18
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限"
         Height          =   180
         Index           =   8
         Left            =   2040
         TabIndex        =   17
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限"
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限"
         Height          =   180
         Index           =   6
         Left            =   2040
         TabIndex        =   15
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限"
         Height          =   180
         Index           =   4
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm040104_d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (無)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit

Dim strReceiveNo As String, pa(1 To 5) As String, intWhere As Integer

Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      If ChkNull Then
         If FormSave Then
            frm040104_9.Show
            Unload Me
         Else
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         End If
      End If
   Else
      frm040104_9.Show
      Unload Me
   End If
End Sub

Private Sub Form_Load()
 Dim i As Integer
   MoveFormToCenter Me
   intWhere = 國內
   For i = 0 To 2
      Combo1(i).Clear
   Next
   With frm040104_9
      strReceiveNo = .Label2(0)
      pa(1) = .Text1
      pa(2) = .Text2
      pa(3) = .Text3
      pa(4) = .Text4
      'pa(5) = .Text6 'Removed by Morgan 2013/1/14 沒再用
   End With
   Select Case frm040104_9.Tag
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
   strExc(0) = "SELECT NP08,NP09,NP15 FROM NEXTPROGRESS WHERE NP01='" & strReceiveNo & _
      "' AND NP07='" & 補文件 & "' AND (NP06 IS NULL OR NP06='')"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      i = 0
      Do While Not .EOF
         If Not IsNull(.Fields(0)) Then Text1(i) = TransDate(.Fields(0), 1)
         If Not IsNull(.Fields(1)) Then Text2(i) = TransDate(.Fields(1), 1)
         If Not IsNull(.Fields(2)) Then Combo1(i).Text = .Fields(2)
         i = i + 1
         .MoveNext
      Loop
   End If
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040104_d = Nothing
End Sub

Private Function FormSave() As Boolean
 Dim i As Integer, intMax As Long, strTxt(1 To 5) As String, j As Integer
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
FormSave = True
cnnConnection.BeginTrans

   strTxt(1) = "DELETE FROM NEXTPROGRESS WHERE NP01='" & strReceiveNo & _
      "' AND NP07=" & 補文件 & " AND (NP06 IS NULL OR NP06='')"
    'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(1)
    
      'edit by nickc 2007/02/02 不用 dll 了
   'intMax = objPublicData.GetNextProgressNo
   intMax = GetNextProgressNo
   j = 2
   For i = 0 To 2
      If Text1(i) <> "" And Text2(i) <> "" And Combo1(i) <> "" Then
        'Modify By Cheng 2003/11/24
        '重抓智權人員
'      strTxt(j) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'         "NP09,NP10,NP15,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & _
'         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 補文件 & _
'         "," & TransDate(Text1(i), 2) & "," & TransDate(Text2(i), 2) & ",'" & _
'         pa(5) & "','" & Combo1(i) & "'," & intMax & ")"
      strTxt(j) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
         "NP09,NP10,NP15,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & _
         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 補文件 & _
         "," & TransDate(Text1(i), 2) & "," & TransDate(Text2(i), 2) & ",'" & _
         PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4)) & "','" & Combo1(i) & "'," & intMax & ")"
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(j)
      j = j + 1
'      intMax = intMax + 1
           
   
   intMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了  intMax = objPublicData.GetNextProgressNo
      End If
   Next
'   FormSave = objLawDll.ExecSQL(j - 1, strTxt)
    cnnConnection.CommitTrans
    Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    FormSave = False
End Function

Private Function ChkNull() As Boolean
 Dim i As Integer, bolChk As Boolean
   ChkNull = True
   bolChk = True
   For i = 0 To 2
      If (Text1(i) <> "" And (Text2(i) = "" Or Combo1(i) = "")) Or _
         (Text2(i) <> "" And (Text1(i) = "" Or Combo1(i) = "")) Or _
         (Combo1(i) <> "" And (Text1(i) = "" Or Text1(i) = "")) Then
         bolChk = False
         Exit For
      End If
   Next
   If Not bolChk Then
      MsgBox "本所期限、法定期限及備註必須同時輸入 !", vbCritical
      ChkNull = False
   End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    If Text1(Index) <> "" Then
        If Not ChkDate(Text1(Index)) Then
            Cancel = True
            TextInverse Text1(Index)
        Else
            'Add By Cheng 2003/12/08
            Select Case Index
            Case 0, 1, 2 '本所期限
                '若本所期限非工作天則直接調整至最近的工作天
                Me.Text1(Index).Text = TransDate(PUB_GetWorkDay1(Me.Text1(Index).Text, True), 1)
            End Select
            'End
        End If
    End If
    If Cancel = True Then Text1_GotFocus Index
End Sub

Private Sub Text2_GotFocus(Index As Integer)
  TextInverse Text2(Index)
End Sub

Private Sub Text2_LostFocus(Index As Integer)
   If ChkDate(Text1(Index)) Then
      If Not ChkRange(Text1(Index), Text2(Index), "期限") Then
         Text1(Index).SetFocus
      End If
   Else
      Text1(Index).SetFocus
   End If
End Sub
