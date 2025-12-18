VERSION 5.00
Begin VB.Form frm060105_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請案號輸入"
   ClientHeight    =   1560
   ClientLeft      =   255
   ClientTop       =   990
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5340
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4344
      TabIndex        =   6
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3528
      TabIndex        =   5
      Top             =   48
      Width           =   800
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1224
      TabIndex        =   4
      Top             =   912
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2808
      MaxLength       =   2
      TabIndex        =   3
      Top             =   468
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2544
      MaxLength       =   1
      TabIndex        =   2
      Top             =   468
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1704
      MaxLength       =   6
      TabIndex        =   1
      Top             =   468
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1224
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "FCP"
      Top             =   468
      Width           =   495
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   144
      TabIndex        =   8
      Top             =   912
      Width           =   948
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   468
      Width           =   768
   End
End
Attribute VB_Name = "frm060105_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/21 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim intWhere As Integer
'Add By Sindy 2022/5/11
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String
Dim m_Done As Boolean
Dim m_PrevForm As Form
'2022/5/11 END


'Add By Sindy 2022/5/11
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdOK_Click(Index As Integer)
 Dim strTmp As String
 Dim bolChk As Boolean
 
   Select Case Index
      Case 0
         'Add By Sindy 2019/5/10
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> Text1 & Text2 & Text3 & Text4 Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 & ")一致！"
               Exit Sub
            End If
         End If
         '2019/5/10 END
         
         bolChk = False
         Text1_Validate bolChk
         If bolChk Then Exit Sub
         
         If Text2 = "" Then
            MsgBox "本所案號不可空白，請重新輸入 !", vbCritical
            Text2.SetFocus
            Exit Sub
         End If
         If CheckCP02 = False Then Exit Sub 'Add by Morgan 2004/10/21
         bolChk = False
         Text5_Validate bolChk
         If bolChk Then Exit Sub
         
         Text4_LostFocus
         strTmp = Text1 & Text2 & Text3 & Text4
         Select Case Text1.Text
            Case "FCP"
               strExc(0) = "SELECT PA01,PA02,PA03,PA04 FROM PATENT WHERE " & ChgPatent(strTmp)
            Case "FG"
               strExc(0) = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE WHERE " & ChgService(strTmp)
         End Select
         intI = 0
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            frm060105_2.SetData strTmp
            
            'Add By Sindy 2022/5/11
            If Not m_PrevForm Is Nothing Then
               Call frm060105_2.SetParent(m_PrevForm)
            End If
            frm060105_2.m_strIR01 = m_strIR01
            frm060105_2.m_strIR02 = m_strIR02
            frm060105_2.m_strIR03 = m_strIR03
            frm060105_2.m_strIR04 = m_strIR04
            '2022/5/11 END
            
            frm060105_2.Show
            Me.Hide
         Else
            Text2.SetFocus
         End If
      Case 1
         Unload Me
   End Select
End Sub

'Added by Sindy 2022/5/11
Private Sub Form_Activate()
   If m_strIR01 <> "" And m_Done = False Then
      Text1.Text = m_strCP01
      Text2.Text = m_strCP02
      Text3.Text = m_strCP03
      Text4.Text = m_strCP04
      cmdOK(0).Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
End Sub
'2022/5/11 END

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   Text5.Text = strSrvDate(2)
   'Add By Cheng 2002/12/11
   SendKeys "{Tab}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2022/5/11
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2022/5/11 END
   
   Set frm060105_1 = Nothing
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1.Text <> "FCP" And Text1.Text <> "FG" Then
      MsgBox "本所案號不正確，請重新輸入 !", vbCritical
      Cancel = True
      InverseTextBox Text1
   End If
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text3_GotFocus()
   InverseTextBox Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   InverseTextBox Text4
End Sub

Private Sub Text4_LostFocus()
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 = "" Then
      MsgBox "來函收文日不可空白，請重新輸入 !", vbCritical
      Cancel = True
   Else
      If ChkDate(Text5) Then
         If Val(Text5) > Val(strSrvDate(2)) Then
            MsgBox "來函收文日不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
'   If Cancel Then TextInverse Text5
   If Cancel Then Text5_GotFocus
End Sub
'Add by Morgan 2004/10/21 檢查本所號
Private Function CheckCP02() As Boolean
   If Len(Text2.Text) <> 6 Then
      MsgBox "本所案號輸入錯誤！"
      Text2.SetFocus
      Text2_GotFocus
      CheckCP02 = False
      Exit Function
   End If
   CheckCP02 = True
End Function
