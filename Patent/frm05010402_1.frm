VERSION 5.00
Begin VB.Form frm05010402_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "公開公告資料輸入"
   ClientHeight    =   2205
   ClientLeft      =   480
   ClientTop       =   1905
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5115
   Begin VB.TextBox txtAppNo 
      Height          =   264
      Left            =   1860
      MaxLength       =   25
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtCode 
      Height          =   288
      Index           =   0
      Left            =   2580
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1170
      Width           =   1212
   End
   Begin VB.TextBox txtCode 
      Height          =   288
      Index           =   1
      Left            =   3780
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1170
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   288
      Index           =   2
      Left            =   4140
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1170
      Width           =   492
   End
   Begin VB.TextBox txtSystem 
      Height          =   288
      Left            =   1860
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1170
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3810
      TabIndex        =   7
      Top             =   165
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2955
      TabIndex        =   6
      Top             =   165
      Width           =   800
   End
   Begin VB.TextBox txtReceivedDay 
      Height          =   264
      Left            =   1860
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1530
      Width           =   1092
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   570
      TabIndex        =   10
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "櫃台收文日："
      Height          =   315
      Left            =   570
      TabIndex        =   9
      Top             =   1530
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   570
      TabIndex        =   8
      Top             =   1170
      Width           =   1215
   End
End
Attribute VB_Name = "frm05010402_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/8 改成Form2.0 (無)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'intChoose 1: 公開公告資料輸入 2: 證書號數輸入
Public intChoose As Integer
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String, m_AppNo As String
Dim m_Done As Boolean
'2016/10/7 END


Private Sub cmdOK_Click(Index As Integer)
If Index = 0 Then
   If CheckDataValid() = False Then
      GoTo EXITSUB
   End If
   
   If CheckKeyIn(2) Then
      If CheckReceivedDay(txtReceivedDay) Then
         Select Case intChoose
            Case 1
               'Add By Sindy 2017/12/28
               If m_strIR01 <> "" Then
                  If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtSystem & txtCode(0) & txtCode(1) & txtCode(2) Then
                     MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
                     Exit Sub
                  End If
               End If
               '2017/12/28 END
               'Add By Sindy 2016/10/7
               frm05010402_2.m_strIR01 = m_strIR01
               frm05010402_2.m_strIR02 = m_strIR02
               frm05010402_2.m_strIR03 = m_strIR03
               frm05010402_2.m_strIR04 = m_strIR04
               '2016/10/7 END
               frm05010402_2.Show
            Case 2
               'Add By Sindy 2017/12/28
               If m_strIR01 <> "" Then
                  If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtSystem & txtCode(0) & txtCode(1) & txtCode(2) Then
                     MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
                     Exit Sub
                  End If
               End If
               '2017/12/28 END
               'Add By Sindy 2016/10/7
               frm05010403_2.m_strIR01 = m_strIR01
               frm05010403_2.m_strIR02 = m_strIR02
               frm05010403_2.m_strIR03 = m_strIR03
               frm05010403_2.m_strIR04 = m_strIR04
               '2016/10/7 END
               frm05010403_2.Show
         End Select
         Me.Hide
      Else
         txtReceivedDay.SetFocus
      End If
   End If
Else
   Unload Me
End If

EXITSUB:
End Sub

'Add By Sindy 2009/06/10
' 90.07.12 modify by louis (檢查資料是否輸入完整)
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDate As String
   CheckDataValid = False
   
   ' 來函收文日
   If IsEmptyText(txtReceivedDay) = False Then
      strDate = txtReceivedDay
      If CheckIsTaiwanDate(strDate, False) = False Then
         strTit = "檢核資料"
         strMsg = "來函收文日日期格式不正確!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtReceivedDay.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add by Morgan 2008/5/22 申請案號檢查
   If txtAppNo = "" Then
      If txtSystem = "CFP" Then
         MsgBox "CFP案必須輸入申請案號！"
         txtAppNo.SetFocus
         GoTo EXITSUB
      End If
   Else
      strExc(1) = "N"
      strExc(0) = "select pa02,pa03,pa04 from patent where pa01='CFP' and pa11='" & txtAppNo & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            If txtCode(0) = RsTemp("pa02") And Left(txtCode(1) & "0", 1) = RsTemp("pa03") Then
               strExc(1) = "Y"
               Exit Do
            End If
            .MoveNext
         Loop
         End With
         If strExc(1) = "N" Then
            MsgBox "本所案號輸入錯誤！"
            txtCode(0).SetFocus
            txtCode_GotFocus (0)
            GoTo EXITSUB
         End If
      Else
         MsgBox "申請案號不存在！"
         txtAppNo.SetFocus
         txtAppNo_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Activate()
   'Added by Sindy 2016/10/7
   If m_strIR01 <> "" And m_Done = False Then
      txtSystem.Text = m_strCP01
      txtCode(0).Text = m_strCP02
      txtCode(1).Text = m_strCP03
      txtCode(2).Text = m_strCP04
      txtAppNo.Text = m_AppNo
      txtReceivedDay = m_RDate
      'cmdOK(0).Value = True
      m_Done = True
      'Add By Sindy 2017/12/28
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
      '2017/12/28 END
   End If
   '2016/10/7 END
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm05010402_1 = Nothing
End Sub

'Add By Sindy 2009/06/10
Private Sub txtAppNo_GotFocus()
   TextInverse txtAppNo
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   TextInverse txtCode(Index)
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   Cancel = Not CheckKeyIn(Index)
   If Cancel Then TextInverse txtCode(Index)
End Sub

Private Sub txtReceivedDay_GotFocus()
   TextInverse txtReceivedDay
End Sub

Private Sub txtReceivedDay_Validate(Cancel As Boolean)
If CheckReceivedDay(txtReceivedDay) = False Then
   TextInverse txtReceivedDay
   Cancel = True
End If
End Sub

Public Sub Clear()
On Error Resume Next
   txtCode(0) = ""
   txtCode(1) = ""
   txtCode(2) = ""
   txtSystem = ""
   
   'Modify By Sindy 2009/06/10
   'txtSystem.SetFocus
   If Me.txtSystem.Enabled Then
      txtAppNo = Empty
      txtAppNo.SetFocus
   End If
End Sub

Private Sub txtSystem_GotFocus()
   TextInverse txtSystem
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
If txtSystem <> "CFP" Then
   ShowMsg MsgText(1056)
   Cancel = True
   txtSystem_GotFocus
End If
End Sub

Private Function CheckKeyIn(ByRef intIndex As Integer) As Boolean
If Len(txtCode(intIndex)) > 0 And Len(txtCode(intIndex)) < txtCode(intIndex).MaxLength Then
   ShowMsg MsgText(33)
ElseIf intIndex = 2 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) Then
   If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) Then
   End If
   CheckKeyIn = True
Else
   CheckKeyIn = True
End If
End Function

Private Function CheckReceivedDay(ByRef strReceivedDay As String) As Boolean
If CheckIsTaiwanDate(strReceivedDay) Then
   If Val(ChangeWDateStringToWString(strReceivedDay)) > Val(strSrvDate(1)) Then
      ShowMsg MsgText(1050)
   Else
      CheckReceivedDay = True
   End If
End If
End Function
