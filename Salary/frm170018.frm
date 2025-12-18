VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm170018 
   BorderStyle     =   1  '單線固定
   Caption         =   "勞保勞退健保費率資料"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5925
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   18
      Left            =   2772
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "99.99"
      Top             =   1800
      Width           =   700
   End
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   17
      Left            =   2775
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "99.99"
      Top             =   1488
      Width           =   700
   End
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   16
      Left            =   2900
      MaxLength       =   4
      TabIndex        =   11
      Text            =   "99.9"
      Top             =   3840
      Width           =   600
   End
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   8
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   9
      Text            =   "99.99"
      Top             =   3180
      Width           =   700
   End
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   7
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   8
      Text            =   "99.99"
      Top             =   2835
      Width           =   700
   End
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   6
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   7
      Text            =   "99.99"
      Top             =   2505
      Width           =   700
   End
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   5
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "99.99"
      Top             =   2115
      Width           =   700
   End
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   4
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "99.99"
      Top             =   1140
      Width           =   700
   End
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   3
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "99.99"
      Top             =   828
      Width           =   700
   End
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   2
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "99.99"
      Top             =   480
      Width           =   700
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   1
      Left            =   4950
      TabIndex        =   13
      Top             =   150
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "存檔(&S)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   3990
      TabIndex        =   12
      Top             =   150
      Width           =   915
   End
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   9
      Left            =   3030
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "99"
      Top             =   3495
      Width           =   435
   End
   Begin VB.TextBox txtIR 
      Alignment       =   1  '靠右對齊
      Height          =   270
      Index           =   1
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   0
      Text            =   "99.99"
      Top             =   144
      Width           =   700
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   45
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4200
      Width           =   5880
      VariousPropertyBits=   671105055
      Size            =   "7223;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "公司負擔職災保險費率２：                   ％ (適用智慧所以外)"
      Height          =   180
      Left            =   570
      TabIndex        =   29
      Top             =   1830
      Width           =   4620
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "公司負擔職災保險費率：                   ％ (適用智慧所)"
      Height          =   180
      Left            =   750
      TabIndex        =   28
      Top             =   1515
      Width           =   4080
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "公司負擔平均眷口數：                   人"
      Height          =   180
      Left            =   945
      TabIndex        =   27
      Top             =   3885
      Width           =   2835
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "勞退："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   600
      TabIndex        =   25
      Top             =   2130
      Width           =   765
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "公司提撥費率：                   ％"
      Height          =   180
      Left            =   1470
      TabIndex        =   24
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "就業保險費率：                   ％"
      Height          =   180
      Left            =   1476
      TabIndex        =   23
      Top             =   504
      Width           =   2292
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "最 高 眷 口 數：                   人"
      Height          =   180
      Left            =   1470
      TabIndex        =   22
      Top             =   3540
      Width           =   2295
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "費　　　　率：                   ％"
      Height          =   180
      Left            =   1470
      TabIndex        =   21
      Top             =   2550
      Width           =   2295
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "公司負擔比例：                   ％"
      Height          =   180
      Left            =   1470
      TabIndex        =   20
      Top             =   3225
      Width           =   2295
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "個人負擔比例：                   ％"
      Height          =   180
      Left            =   1470
      TabIndex        =   19
      Top             =   2865
      Width           =   2295
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "勞保費率：                           ％"
      Height          =   180
      Left            =   1476
      TabIndex        =   18
      Top             =   192
      Width           =   2292
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "公司負擔比例：                   ％"
      Height          =   180
      Left            =   1476
      TabIndex        =   17
      Top             =   1188
      Width           =   2292
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "個人負擔比例：                   ％"
      Height          =   180
      Left            =   1476
      TabIndex        =   16
      Top             =   864
      Width           =   2292
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "健保："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   570
      TabIndex        =   15
      Top             =   2520
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "勞保："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   576
      TabIndex        =   14
      Top             =   168
      Width           =   768
   End
End
Attribute VB_Name = "frm170018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/22 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2008/12/27 add by sonia
Option Explicit

Dim m_FieldList() As FIELDITEM
Dim TF_IR As Integer '欄位數
Dim oText As Object
Dim idx As Integer


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0  '存檔
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               InitialField
               MsgBox "更新勞保勞退健保費率資料完成，請即刻執行「勞、健保保費重算」功能！", vbExclamation
            Else
               MsgBox "更新勞保勞退健保費率資料錯誤！", vbInformation
            End If
         End If
      Case 1  '結束
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   MsgBox "請務必於「當」月份進行調整!!", vbExclamation
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   textCUID.BackColor = &H8000000F
   InitialField
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170018 = Nothing
End Sub

' 初始化欄位陣列及抓資料
Private Sub InitialField()
Dim CUID(1 To 6) As String
   
   strExc(0) = "select * from InsuranceRate "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      ClearField
      With RsTemp
         TF_IR = .Fields.Count
         ReDim m_FieldList(TF_IR) As FIELDITEM
         For Each oText In txtIR
            idx = oText.Index
            m_FieldList(idx).fiName = "IR" & Format(idx, "00")
            'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
            'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
               m_FieldList(idx).fiType = 0
            'Else
            '   m_FieldList(idx).fiType = 1
            'End If
            'end 2017/06/29
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
            m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
            oText.Text = m_FieldList(idx).fiOldData
         Next
         CUID(1) = "" & .Fields("ir10")
         CUID(2) = "" & .Fields("ir11")
         CUID(3) = "" & .Fields("ir12")
         CUID(4) = "" & .Fields("ir13")
         CUID(5) = "" & .Fields("ir14")
         CUID(6) = "" & .Fields("ir15")
      End With
   End If
   UpdateCUID CUID, textCUID
   If Me.Visible = True Then
      txtIR(1).SetFocus
      txtIR_GotFocus 1
   End If
   
End Sub

Private Sub ClearField()
   For Each oText In txtIR
      oText.Text = Empty
   Next
   
   For intI = 1 To TF_IR
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = ""

End Sub

Private Sub UpdateFieldNewData()
   For Each oText In txtIR
      idx = oText.Index
      m_FieldList(idx).fiNewData = oText.Text
   Next
End Sub

Private Function ModRecord() As Boolean
Dim stSQL As String, stSet As String, stCols As String, stValues As String
Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "begin user_data.user_enabled:=1; UPDATE InsuranceRate SET "
   stSet = ""
   For Each oText In txtIR
      idx = oText.Index
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where ir10 is not null; end; "
      
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   cnnConnection.CommitTrans
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Function TxtValidate() As Boolean
   
   If txtIR(1) = "" Then
      ShowMsg "請輸入勞保費率－本國人 !"
      txtIR(1).SetFocus
      txtIR_GotFocus 1
      Exit Function
   End If
   If txtIR(2) = "" Then
      ShowMsg "請輸入勞保費率－外國人 !"
      txtIR(2).SetFocus
      txtIR_GotFocus 2
      Exit Function
   End If
   If txtIR(3) = "" Then
      ShowMsg "請輸入勞保個人負擔比例 !"
      txtIR(3).SetFocus
      txtIR_GotFocus 3
      Exit Function
   End If
   If txtIR(4) = "" Then
      ShowMsg "請輸入勞保公司負擔比例 !"
      txtIR(4).SetFocus
      txtIR_GotFocus 4
      Exit Function
   End If
   If txtIR(5) = "" Then
      ShowMsg "請輸入勞退公司提撥費率 !"
      txtIR(5).SetFocus
      txtIR_GotFocus 5
      Exit Function
   End If
   If txtIR(6) = "" Then
      ShowMsg "請輸入健保費率 !"
      txtIR(6).SetFocus
      txtIR_GotFocus 6
      Exit Function
   End If
   If txtIR(7) = "" Then
      ShowMsg "請輸入健保個人負擔比例 !"
      txtIR(7).SetFocus
      txtIR_GotFocus 7
      Exit Function
   End If
   If txtIR(8) = "" Then
      ShowMsg "請輸入健保公司負擔比例!"
      txtIR(8).SetFocus
      txtIR_GotFocus 8
      Exit Function
   End If
   If txtIR(9) = "" Then
      ShowMsg "請輸入健保最高眷口數 !"
      txtIR(9).SetFocus
      txtIR_GotFocus 9
      Exit Function
   End If
   If txtIR(16) = "" Then
      ShowMsg "請輸入健保公司負擔平均眷口數 !"
      txtIR(16).SetFocus
      txtIR_GotFocus 16
      Exit Function
   End If
   
   If txtIR(17) = "" Then
      ShowMsg "請輸入公司負擔職災保險費率!"
      txtIR(17).SetFocus
      txtIR_GotFocus 17
      Exit Function
   End If
   
   If txtIR(18) = "" Then
      ShowMsg "請輸入公司負擔職災保險費率２!"
      txtIR(18).SetFocus
      txtIR_GotFocus 18
      Exit Function
   End If
   
   TxtValidate = True
    
End Function

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

Private Sub txtIR_GotFocus(Index As Integer)
   TextInverse txtIR(Index)
   CloseIme
End Sub

Private Sub txtIR_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub
