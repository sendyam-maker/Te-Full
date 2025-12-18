VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040102_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "催審提申期限設定"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   8295
   Begin VB.TextBox txtCF 
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   1215
      TabIndex        =   11
      Top             =   990
      Width           =   1185
   End
   Begin VB.TextBox txtCF 
      Alignment       =   2  '置中對齊
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   3060
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox txtCF 
      Height          =   315
      Index           =   0
      Left            =   1215
      TabIndex        =   8
      Top             =   1650
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6930
      TabIndex        =   7
      Top             =   90
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Height          =   400
      Index           =   0
      Left            =   5670
      TabIndex        =   6
      Top             =   90
      Width           =   1200
   End
   Begin VB.TextBox txtCF 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   1215
      TabIndex        =   0
      Top             =   660
      Width           =   1185
   End
   Begin VB.TextBox txtCF 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   1215
      TabIndex        =   1
      Top             =   1320
      Width           =   1185
   End
   Begin MSForms.Label lblCF 
      Height          =   285
      Index           =   2
      Left            =   2445
      TabIndex        =   13
      Top             =   1020
      Width           =   2385
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4207;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Index           =   3
      Left            =   270
      TabIndex        =   12
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "審查時間：                              個月 (                         天)"
      Height          =   180
      Index           =   2
      Left            =   270
      TabIndex        =   9
      Top             =   1710
      Width           =   4080
   End
   Begin VB.Label lblMemo 
      Caption         =   "審查時間以 12 個月計算年度(每年以365 天計算),餘數再以每個月 30 天計算。(系統存天數)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   4950
      TabIndex        =   5
      Top             =   990
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統別："
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   750
      Width           =   720
   End
   Begin MSForms.Label lblCF 
      Height          =   285
      Index           =   3
      Left            =   2445
      TabIndex        =   3
      Top             =   1350
      Width           =   2385
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4207;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   1
      Left            =   270
      TabIndex        =   2
      Top             =   1380
      Width           =   900
   End
End
Attribute VB_Name = "frm12040102_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/29 Form2.0已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
'Create by Morgan 2009/8/3
Option Explicit

Dim bolActivated As Boolean


Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         If TxtValidate = True Then
            If FormSave = True Then
               Unload Me
            End If
         End If
      Case 1
         Unload Me
   End Select
   
End Sub

Private Function FormSave() As Boolean
   Dim StrMailContent As String
   Dim strMailSubject As String
   If txtCF(5) <> txtCF(5).Tag Then
      cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

      strSql = "select * from casefee where cf01='" & txtCF(1) & "' and cf02='" & txtCF(2) & "' and cf03='" & txtCF(3) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strSql = "update casefee set cf05=" & CNULL(txtCF(5), True) & " where cf01='" & txtCF(1) & "' and cf02='" & txtCF(2) & "' and cf03='" & txtCF(3) & "'"
      Else
         strSql = "insert into casefee (cf01,cf02,cf03,cf05)" & _
            " values ('" & txtCF(1) & "','" & txtCF(2) & "','" & txtCF(3) & "'," & CNULL(txtCF(5), True) & ")"
      End If
      
      Pub_SeekTbLog strSql
      strSql = "begin user_data.user_enabled:=1;  " & strSql & "; end; "
      cnnConnection.Execute strSql, intI
      
      strMailSubject = txtCF(1) & "系統" & lblCF(2) & "(" & txtCF(2) & ")案的" & lblCF(3) & "(" & txtCF(3) & ")的審查時間已變更！"
      StrMailContent = txtCF(5).Tag & " --> " & txtCF(5) & " 天"
      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
         " values ('" & strUserNum & "','79075',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(strMailSubject) & "','" & ChgSQL(StrMailContent) & "')"
      cnnConnection.Execute strSql, intI
      cnnConnection.CommitTrans
   End If
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub Form_Activate()
   If bolActivated = False Then
      bolActivated = True
      SetValue
      txtCF(0).SetFocus
      txtCF_GotFocus 0
   End If
End Sub

Private Sub SetValue()
   strSql = "select cf05,trunc(Cf05/365)*12+round(MOD(CF05,365)/30,1) Mn from casefee where cf01='" & txtCF(1) & "' and cf02='" & txtCF(2) & "' and cf03='" & txtCF(3) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      txtCF(5) = "" & RsTemp(0)
      txtCF(5).Tag = txtCF(5)
      txtCF(0) = "" & RsTemp(1)
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040102_2 = Nothing
End Sub

Private Sub txtCF_Change(Index As Integer)
   Dim strTempName As String
   Select Case Index
      Case 0 '審查期限(月)
         If txtCF(0) = "" Then
            txtCF(5) = ""
         Else
            txtCF(5) = GetDays(Val(txtCF(Index)))
         End If
      Case 2 '申請國家
         lblCF(2) = GetPrjNationName(txtCF(2))
      Case 3 '案件性質
         If ClsPDGetCaseProperty(txtCF(1), txtCF(3), strTempName) Then
            lblCF(3) = strTempName
         End If
   End Select
End Sub

Private Function GetDays(pMonths As Double) As String
   '12個月以365天計,餘數每1個月以30天計(四捨五入)
   GetDays = 365 * (Int(pMonths) \ 12) + Round(30 * (pMonths - 12 * (Int(pMonths) \ 12)))
End Function

Private Sub txtCF_GotFocus(Index As Integer)
   CloseIme
   TextInverse txtCF(Index)
End Sub

Private Sub txtCF_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Index = 0 Then
      If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(".") Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Function TxtValidate() As Boolean
   If Len(txtCF(5)) > 4 Then
      txtCF(0).SetFocus
      txtCF_GotFocus 0
      MsgBox "審查時間超過限制！", vbExclamation
      Exit Function
   End If
   TxtValidate = True
End Function
