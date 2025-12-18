VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010604_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "分割案件關係維護"
   ClientHeight    =   4065
   ClientLeft      =   435
   ClientTop       =   1635
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7695
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6756
      TabIndex        =   14
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4704
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5532
      TabIndex        =   13
      Top             =   70
      Width           =   1200
   End
   Begin VB.Frame fraIn 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   372
      Left            =   1080
      TabIndex        =   22
      Top             =   1725
      Width           =   2535
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         Height          =   264
         Index           =   9
         Left            =   1380
         MaxLength       =   1
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   4
         Left            =   0
         MaxLength       =   3
         TabIndex        =   7
         Top             =   0
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   7
         Left            =   2100
         MaxLength       =   2
         TabIndex        =   11
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   6
         Left            =   1740
         MaxLength       =   1
         TabIndex        =   10
         Top             =   0
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   5
         Left            =   600
         MaxLength       =   6
         TabIndex        =   8
         Top             =   0
         Width           =   1035
      End
   End
   Begin VB.Frame fraOut 
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   372
      Left            =   1080
      TabIndex        =   21
      Top             =   555
      Width           =   2535
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         Height          =   264
         Index           =   8
         Left            =   1380
         MaxLength       =   1
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   600
         MaxLength       =   6
         TabIndex        =   1
         Top             =   0
         Width           =   1035
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   1740
         MaxLength       =   1
         TabIndex        =   3
         Top             =   0
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   3
         Left            =   2100
         MaxLength       =   2
         TabIndex        =   4
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   0
         MaxLength       =   3
         TabIndex        =   0
         Top             =   0
         Width           =   492
      End
   End
   Begin MSForms.ComboBox cboPromoterIn 
      Height          =   300
      Left            =   1050
      TabIndex        =   34
      Top             =   2430
      Width           =   6495
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "11456;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboIn 
      Height          =   300
      Left            =   1050
      TabIndex        =   33
      Top             =   2100
      Width           =   6495
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "11456;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboPromoterOut 
      Height          =   300
      Left            =   1050
      TabIndex        =   6
      Top             =   1260
      Width           =   6495
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "11456;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboOut 
      Height          =   300
      Left            =   1050
      TabIndex        =   5
      Top             =   930
      Width           =   6495
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "11451;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   3
      Left            =   4380
      TabIndex        =   32
      Top             =   3630
      Width           =   1650
      BackColor       =   -2147483644
      VariousPropertyBits=   27
      Size            =   "2910;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   1140
      TabIndex        =   31
      Top             =   3630
      Width           =   1650
      BackColor       =   -2147483644
      VariousPropertyBits=   27
      Size            =   "2910;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   4380
      TabIndex        =   30
      Top             =   3330
      Width           =   1650
      BackColor       =   -2147483644
      VariousPropertyBits=   27
      Size            =   "2910;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   1140
      TabIndex        =   29
      Top             =   3330
      Width           =   1650
      BackColor       =   -2147483644
      VariousPropertyBits=   27
      Size            =   "2910;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請日："
      Height          =   180
      Index           =   0
      Left            =   4230
      TabIndex        =   28
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Update Time:"
      Height          =   180
      Index           =   4
      Left            =   3360
      TabIndex        =   27
      Top             =   3667
      Width           =   948
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Update Name:"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   3667
      Width           =   996
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Time:"
      Height          =   180
      Index           =   2
      Left            =   3360
      TabIndex        =   25
      Top             =   3367
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Create Name:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   3367
      Width           =   948
   End
   Begin VB.Label lblSendDay 
      Height          =   255
      Left            =   5190
      TabIndex        =   23
      Top             =   1770
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "分割案號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   990
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "母案案號："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   1725
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   2490
      Width           =   720
   End
End
Attribute VB_Name = "frm02010604_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/11/30 改成Form2.0 ;Label3(index)、cboOut、cboPromoterOut、cboIn、cboPromoterIn
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/23 日期欄已修改
Option Explicit

'intLeaveKind離開時，是0:結束1:回上一畫面
Dim intLeaveKind As Integer
Public intWhereToGo As Integer '0從frm02010604_1來,1從frm02010604_3來
Public strCode1 As String, strCode2 As String, strCode3 As String, strCode4 As String
Public strCode5 As String, strCode6 As String, strCode7 As String, strCode8 As String
Public strCode18 As String
Public intChoose As String
Dim m_blnFirstShow As Boolean
Dim m_strSystemKindForUser As String '記錄使用者可使用的系統類別

Private Sub cmdOK_Click(Index As Integer)
Dim strCode() As String, i As Integer, bolSave As Boolean
Dim StrSQLa As String
   
On Error GoTo ErrorHandler
    bolSave = True
    Select Case Index
    Case 0 '確定
        Screen.MousePointer = vbHourglass
        '檢查資料輸入的完整性
        For i = 4 To 9
            If i <> 8 Then
                If CheckKeyIn(i) = False Then
                    If i = 3 Then
                        i = 0
                    ElseIf i = 7 Then
                        i = 4
                    End If
                    Me.txtCode(i).SetFocus
                    txtCode_GotFocus i
                    Exit For
                End If
            End If
        Next i
        If i <> 10 Then Screen.MousePointer = vbDefault: Exit Sub
        If CheckDataValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
        '重新檢查欄位有效性
        If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
        Select Case intChoose
        Case 1 '新增
            If Me.txtCode(2).Text = "" Then Me.txtCode(2).Text = "0"
            If Me.txtCode(3).Text = "" Then Me.txtCode(3).Text = "00"
            If Me.txtCode(6).Text = "" Then Me.txtCode(6).Text = "0"
            If Me.txtCode(7).Text = "" Then Me.txtCode(7).Text = "00"
            ReDim strCode(7) As String
            For i = 0 To 7
                strCode(i) = txtCode(i)
            Next i
            If strCode(0) = "TF" Then strCode(1) = Me.txtCode(8).Text
            If strCode(4) = "TF" Then strCode(5) = Me.txtCode(9).Text
            cnnConnection.BeginTrans
            bolSave = False
            StrSQLa = "Insert Into DivisionCase(DC01, DC02, DC03, DC04, DC05, DC06, DC07, DC08, DC09, DC10, DC11) " & _
                                " values ('" & strCode(0) & "','" & strCode(1) & "','" & strCode(2) & "','" & strCode(3) & "','" & strCode(4) & "','" & strCode(5) & "','" & strCode(6) & "','" & strCode(7) & "','" & strUserNum & "'," & strSrvDate(1) & "," & ServerTime & " )"
            cnnConnection.Execute StrSQLa
            cnnConnection.CommitTrans
            bolSave = True
            If Me.intWhereToGo = 1 Then
                frm02010604_3.m_blnFirstShow = True
            End If
        Case 2 '修改
            If Me.txtCode(2).Text = "" Then Me.txtCode(2).Text = "0"
            If Me.txtCode(3).Text = "" Then Me.txtCode(3).Text = "00"
            If Me.txtCode(6).Text = "" Then Me.txtCode(6).Text = "0"
            If Me.txtCode(7).Text = "" Then Me.txtCode(7).Text = "00"
            ReDim strCode(7) As String
            For i = 0 To 7
                strCode(i) = txtCode(i)
            Next i
            If strCode(0) = "TF" Then strCode(1) = Me.txtCode(8).Text
            If strCode(4) = "TF" Then strCode(5) = Me.txtCode(9).Text
            cnnConnection.BeginTrans
            bolSave = False
            StrSQLa = "Update DivisionCase Set DC05='" & strCode(4) & "', DC06='" & strCode(5) & "', DC07='" & strCode(6) & "', DC08='" & strCode(7) & "', DC12='" & strUserNum & "', DC13=" & strSrvDate(1) & ", DC14=" & ServerTime & " Where DC01='" & strCode(0) & "' And DC02='" & strCode(1) & "' And DC03='" & strCode(2) & "' And DC04='" & strCode(3) & "' "
            cnnConnection.Execute StrSQLa
            cnnConnection.CommitTrans
            bolSave = True
            If Me.intWhereToGo = 1 Then
                frm02010604_3.m_blnFirstShow = True
            End If
        Case 4 '刪除
            If Me.txtCode(2).Text = "" Then Me.txtCode(2).Text = "0"
            If Me.txtCode(3).Text = "" Then Me.txtCode(3).Text = "00"
            If Me.txtCode(6).Text = "" Then Me.txtCode(6).Text = "0"
            If Me.txtCode(7).Text = "" Then Me.txtCode(7).Text = "00"
            ReDim strCode(7) As String
            For i = 0 To 7
                strCode(i) = txtCode(i)
            Next i
            If strCode(0) = "TF" Then strCode(1) = Me.txtCode(8).Text
            If strCode(4) = "TF" Then strCode(5) = Me.txtCode(9).Text
            If MsgBox("您是否確定刪除???", vbExclamation + vbYesNo) = vbYes Then
                cnnConnection.BeginTrans
                bolSave = False
                StrSQLa = "Delete From DivisionCase Where DC01='" & strCode(0) & "' And DC02='" & strCode(1) & "' And DC03='" & strCode(2) & "' And DC04='" & strCode(3) & "' "
                cnnConnection.Execute StrSQLa
                cnnConnection.CommitTrans
                bolSave = True
                If Me.intWhereToGo = 1 Then
                    frm02010604_3.m_blnFirstShow = True
                End If
            Else
                bolSave = False
            End If
        Case 5 '查詢
            '無動作
        End Select
        Screen.MousePointer = vbDefault
        If bolSave Then
            intLeaveKind = 1
            Unload Me
        End If
        Case 1 '回前畫面
            intLeaveKind = 1
            Unload Me
        Case 2 '結束
            intLeaveKind = 0
            Unload Me
    End Select
    Exit Sub
ErrorHandler:
    Screen.MousePointer = vbDefault
    If bolSave = False Then cnnConnection.RollbackTrans
    If Err.Number <> 0 Then MsgBox "(" & Err.Number & ")" & Err.Description
End Sub

Private Sub Form_Activate()
Dim bolGoOn As Boolean
Dim Lbl As Object
Dim strTxt(1 To 17) As String, i As Integer
 
    If m_blnFirstShow = True Then
        Me.Caption = Me.Caption & IIf(Me.intChoose = "1", "--新增", IIf(Me.intChoose = "2", "--修改", IIf(Me.intChoose = "4", "--刪除", "--查詢")))
        Me.fraOut.Enabled = False
        If Me.intChoose = "1" Or Me.intChoose = "2" Then
            Me.fraIn.Enabled = True
        Else
            Me.fraIn.Enabled = False
        End If
        txtCode(0) = strCode1
        If Me.txtCode(0).Text = "TF" Then
            txtCode(1) = Mid(strCode2, 1, 5)
            Me.txtCode(8).Text = Mid(strCode2, 6, 1)
        Else
            txtCode(1) = strCode2
            Me.txtCode(8).Text = ""
        End If
        txtCode(2) = strCode3
        txtCode(3) = strCode4
        txtCode(4) = strCode5
        If Me.txtCode(0).Text = "TF" Then
            txtCode(5) = Mid(strCode6, 1, 5)
            Me.txtCode(9).Text = Mid(strCode6, 6, 1)
        Else
            txtCode(5) = strCode6
            Me.txtCode(9).Text = ""
        End If
        txtCode(6) = strCode7
        txtCode(7) = strCode8
        For Each Lbl In Label3
            Lbl.Caption = ""
        Next
        Me.cboPromoterOut.Clear
        Me.cboPromoterIn.Clear
        GetDivisionCase Me.txtCode(0).Text, Me.txtCode(1).Text & Me.txtCode(8).Text, Me.txtCode(2).Text, Me.txtCode(3).Text
        If Me.txtCode(0).Text <> "" Then GetCaseData "D", Me.txtCode(0).Text, Me.txtCode(1).Text & Me.txtCode(8).Text, Me.txtCode(2).Text, Me.txtCode(3).Text
        If Me.txtCode(4).Text <> "" Then GetCaseData "M", Me.txtCode(4).Text, Me.txtCode(5).Text & Me.txtCode(9).Text, Me.txtCode(6).Text, Me.txtCode(7).Text
        m_blnFirstShow = False
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    m_blnFirstShow = True
    m_strSystemKindForUser = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If intWhereToGo = 0 Then
        If intLeaveKind = 1 Then
            frm02010604_1.Show
        Else
            Unload frm02010604_1
        End If
    Else
        If intLeaveKind = 1 Then
            frm02010604_3.Show
        Else
            Unload frm02010604_3
        End If
    End If
    intLeaveKind = 0
    Set frm02010604_2 = Nothing
End Sub

Private Sub txtCode_Change(Index As Integer)
    Select Case Index
    Case 0 '分割案系統類別
        If Me.txtCode(Index).Text = "TF" Then
            Me.txtCode(8).Visible = True
            Me.txtCode(8).Enabled = True
            Me.txtCode(1).MaxLength = 5
        Else
            Me.txtCode(8).Visible = False
            Me.txtCode(8).Enabled = False
            Me.txtCode(1).MaxLength = 6
        End If
    Case 4 '母案系統類別
        If Me.txtCode(Index).Text = "TF" Then
            Me.txtCode(9).Visible = True
            Me.txtCode(8).Enabled = True
            Me.txtCode(5).MaxLength = 5
        Else
            Me.txtCode(9).Visible = False
            Me.txtCode(9).Enabled = False
            Me.txtCode(5).MaxLength = 6
        End If
    End Select
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
    TextInverse txtCode(Index)
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_LostFocus(Index As Integer)
    Select Case Index
    Case 7:
        If CheckKeyIn(Index) = False Then
            txtCode(4).SetFocus
        End If
    End Select
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = False Then
      '本所案號錯誤時,讓Cursor繼續往下跳
      If Index <> 3 And Index <> 7 Then
         Cancel = True
         txtCode_GotFocus Index
      End If
   End If
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Boolean
Dim intCaseKind As Integer, intWhere As Integer, strTemp As String
Dim arrSystemKind, arrSystemKind1
Dim ii As Integer, jj As Integer
Dim blnNoRight As Boolean
Dim strNoRightSK As String
Dim strCode(7) As String

Select Case intIndex
Case 0, 4 '系統類別
    If Me.txtCode(intIndex).Enabled = True Then
        If Me.txtCode(intIndex).Text <> "" Then
            If m_strSystemKindForUser <> "" Then
                blnNoRight = True
                arrSystemKind = Split(m_strSystemKindForUser, ",")
                For ii = LBound(arrSystemKind) To UBound(arrSystemKind)
                    If Me.txtCode(intIndex).Text = arrSystemKind(ii) Then
                        blnNoRight = False
                        Exit For
                    End If
                Next ii
                If blnNoRight = True Then
                    MsgBox "您無權使用 " & Me.txtCode(intIndex).Text & " 系統類別!!!", vbExclamation + vbOKOnly
                Else
                    CheckKeyIn = True
                End If
            Else
                MsgBox "您無權使用 " & Me.txtCode(intIndex).Text & " 系統類別!!!", vbExclamation + vbOKOnly
            End If
        Else
            CheckKeyIn = True
        End If
    Else
        CheckKeyIn = True
    End If
Case 3, 7
    If Me.txtCode(intIndex).Enabled = True Then
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.CheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2) & IIf(intIndex = 3, Me.txtCode(8).Text, Me.txtCode(9).Text), _
             IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
        If ClsPDCheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2) & IIf(intIndex = 3, Me.txtCode(8).Text, Me.txtCode(9).Text), _
             IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
            If intIndex = 3 Then
                If Me.txtCode(2).Text = "" Then Me.txtCode(2).Text = "0"
                If Me.txtCode(3).Text = "" Then Me.txtCode(3).Text = "00"
                GetCaseData "D", Me.txtCode(0).Text, Me.txtCode(1).Text & Me.txtCode(8).Text, Me.txtCode(2).Text, Me.txtCode(3).Text
            Else
                If Me.txtCode(6).Text = "" Then Me.txtCode(6).Text = "0"
                If Me.txtCode(7).Text = "" Then Me.txtCode(7).Text = "00"
                GetCaseData "M", Me.txtCode(4).Text, Me.txtCode(5).Text & Me.txtCode(9).Text, Me.txtCode(6).Text, Me.txtCode(7).Text
            End If
            CheckKeyIn = True
        End If
    Else
        CheckKeyIn = True
    End If
Case Else
    CheckKeyIn = True
End Select
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In txtCode
   If objTxt.Enabled = True Then
      Cancel = False
      txtCode_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.txtCode(objTxt.Index).SetFocus
         txtCode_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function

Private Sub GetDivisionCase(strDC01 As String, strDC02 As String, strDC03 As String, strDC04 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select DC01, DC02, DC03, DC04, DC05, DC06, DC07, DC08, DC09, DC10, DC11, DC12, DC13, DC14, S1.ST02 As S1ST02, S2.ST02 As S2ST02 From DivisionCase, Staff S1, Staff S2 Where DC09=S1.ST01(+) And DC12=S2.ST01(+) And DC01='" & strDC01 & "' And DC02='" & strDC02 & "' And DC03='" & strDC03 & "' And DC04='" & strDC04 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    Me.txtCode(4).Text = "" & rsA("DC05").Value
    Me.txtCode(5).Text = IIf(Me.txtCode(4).Text = "TF", Left("" & rsA("DC06").Value, 5), "" & rsA("DC06").Value)
    Me.txtCode(9).Text = IIf(Me.txtCode(4).Text = "TF", Mid("" & rsA("DC06").Value, 6, 1), "")
    Me.txtCode(6).Text = "" & rsA("DC07").Value
    Me.txtCode(7).Text = "" & rsA("DC08").Value
    Me.Label3(0).Caption = "" & rsA("S1ST02").Value
    Me.Label3(1).Caption = ChangeTStringToTDateString(ChangeWStringToTString("" & rsA("DC10").Value)) & "    " & Format(rsA("DC11").Value, "##:##:##")
    Me.Label3(2).Caption = "" & rsA("S2ST02").Value
    Me.Label3(3).Caption = ChangeTStringToTDateString(ChangeWStringToTString("" & rsA("DC13").Value)) & "    " & Format(rsA("DC14").Value, "##:##:##")
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub

Private Sub GetCaseData(strKind As String, strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String)
'strKind : D分割案, M母案
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer

StrSQLa = "Select PA05, PA06, PA07, ''||PA10, PA26, Nvl(C1.CU04, Decode(C1.CU05, Null, C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)), PA27, Nvl(C2.CU04, Decode(C2.CU05, Null, C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)), PA28, Nvl(C3.CU04, Decode(C3.CU05, Null, C3.CU06, C3.CU05||' '||C3.CU88||' '||C3.CU89||' '||C3.CU90)), PA29, Nvl(C4.CU04, Decode(C4.CU05, Null, C4.CU06, C4.CU05||' '||C4.CU88||' '||C4.CU89||' '||C4.CU90)), PA30, Nvl(C5.CU04, Decode(C5.CU05, Null, C5.CU06, C5.CU05||' '||C5.CU88||' '||C5.CU89||' '||C5.CU90)) From Patent , Customer C1, Customer C2, Customer C3, Customer C4, Customer C5 " & _
                " Where substr(PA26,1,8)=C1.CU01(+) And substr(PA26,9,1)=C1.CU02(+) And substr(PA27,1,8)=C2.CU01(+) And substr(PA27,9,1)=C2.CU02(+) And substr(PA28,1,8)=C3.CU01(+) And substr(PA28,9,1)=C3.CU02(+) And substr(PA29,1,8)=C4.CU01(+) And substr(PA29,9,1)=C4.CU02(+) And substr(PA30,1,8)=C5.CU01(+) And substr(PA30,9,1)=C5.CU02(+) And PA01='" & strCP01 & "' And PA02='" & strCP02 & "' And PA03='" & strCP03 & "' And PA04='" & strCP04 & "' "
StrSQLa = StrSQLa & " Union Select TM05, TM06, TM07, ''||TM11, TM23, Nvl(C1.CU04, Decode(C1.CU05, Null, C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)), '', '', '', '', '', '', '', '' From Trademark , Customer C1 " & _
                " Where substr(TM23,1,8)=C1.CU01(+) And substr(TM23,9,1)=C1.CU02(+) And TM01='" & strCP01 & "' And TM02='" & strCP02 & "' And TM03='" & strCP03 & "' And TM04='" & strCP04 & "' "
StrSQLa = StrSQLa & " Union Select LC05, LC06, LC07, '', LC11, Nvl(C1.CU04, Decode(C1.CU05, Null, C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)), '', '', '', '', '', '', '', '' From Lawcase , Customer C1 " & _
                " Where substr(LC11,1,8)=C1.CU01(+) And substr(LC11,9,1)=C1.CU02(+) And LC01='" & strCP01 & "' And LC02='" & strCP02 & "' And LC03='" & strCP03 & "' And LC04='" & strCP04 & "' "
StrSQLa = StrSQLa & " Union Select HC06, '', '', '', HC05, Nvl(C1.CU04, Decode(C1.CU05, Null, C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)), '', '', '', '', '', '', '', '' From Hirecase , Customer C1 " & _
                " Where substr(HC05,1,8)=C1.CU01(+) And substr(HC05,9,1)=C1.CU02(+) And HC01='" & strCP01 & "' And HC02='" & strCP02 & "' And HC03='" & strCP03 & "' And HC04='" & strCP04 & "' "
StrSQLa = StrSQLa & " Union Select SP05, SP06, SP07, ''||SP10, SP08, Nvl(C1.CU04, Decode(C1.CU05, Null, C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)), SP58, Nvl(C2.CU04, Decode(C2.CU05, Null, C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)), SP59, Nvl(C3.CU04, Decode(C3.CU05, Null, C3.CU06, C3.CU05||' '||C3.CU88||' '||C3.CU89||' '||C3.CU90)), '', '', '', '' From Servicepractice , Customer C1, Customer C2, Customer C3 " & _
                " Where substr(SP08,1,8)=C1.CU01(+) And substr(SP08,9,1)=C1.CU02(+) And substr(SP58,1,8)=C2.CU01(+) And substr(SP58,9,1)=C2.CU02(+) And substr(SP59,1,8)=C3.CU01(+) And substr(SP59,9,1)=C3.CU02(+) And SP01='" & strCP01 & "' And SP02='" & strCP02 & "' And SP03='" & strCP03 & "' And SP04='" & strCP04 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    '分割案
    If UCase(strKind) = "D" Then
        Me.cboOut.Clear
        If "" & rsA.Fields(0).Value <> "" Then
            Me.cboOut.AddItem "中：" & rsA.Fields(0).Value
        End If
        If "" & rsA.Fields(1).Value <> "" Then
            Me.cboOut.AddItem "英：" & rsA.Fields(0).Value
        End If
        If "" & rsA.Fields(2).Value <> "" Then
            Me.cboOut.AddItem "日：" & rsA.Fields(2).Value
        End If
        If Me.cboOut.ListCount > 0 Then Me.cboOut.ListIndex = 0
        Me.cboPromoterOut.Clear
        For ii = 4 To 11 Step 2
            If "" & rsA.Fields(ii).Value <> "" Then
                Me.cboPromoterOut.AddItem "" & rsA.Fields(ii).Value & " " & rsA.Fields(ii + 1).Value
            End If
        Next ii
        If Me.cboPromoterOut.ListCount > 0 Then Me.cboPromoterOut.ListIndex = 0
    '母案
    Else
        Me.cboIn.Clear
        If "" & rsA.Fields(0).Value <> "" Then
            Me.cboIn.AddItem "中：" & rsA.Fields(0).Value
        End If
        If "" & rsA.Fields(1).Value <> "" Then
            Me.cboIn.AddItem "英：" & rsA.Fields(0).Value
        End If
        If "" & rsA.Fields(2).Value <> "" Then
            Me.cboIn.AddItem "日：" & rsA.Fields(2).Value
        End If
        If Me.cboIn.ListCount > 0 Then Me.cboIn.ListIndex = 0
        Me.cboPromoterIn.Clear
        For ii = 4 To 11 Step 2
            If "" & rsA.Fields(ii).Value <> "" Then
                Me.cboPromoterIn.AddItem "" & rsA.Fields(ii).Value & " " & rsA.Fields(ii + 1).Value
            End If
        Next ii
        If Me.cboPromoterIn.ListCount > 0 Then Me.cboPromoterIn.ListIndex = 0
        Me.lblSendDay.Caption = ChangeTStringToTDateString(ChangeWStringToTString("" & rsA.Fields(3).Value))
    End If
Else
    '分割案
    If UCase(strKind) = "D" Then
        Me.cboOut.Clear
        Me.cboPromoterOut.Clear
    '母案
    Else
        Me.cboIn.Clear
        Me.cboPromoterIn.Clear
        Me.lblSendDay.Caption = ""
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub

Private Function CheckDataValidate() As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCode() As String
Dim i As Integer
    
CheckDataValidate = False
'若為新增或修改
If Me.intChoose = "1" Or Me.intChoose = "2" Then
    If Me.txtCode(4).Text = "" Or Me.txtCode(5).Text = "" Or (Me.txtCode(4).Text = "TF" And Me.txtCode(9).Text = "") Then
        MsgBox "母案案號輸入不完整!!!", vbExclamation + vbOKOnly
        Me.txtCode(4).SetFocus
        Exit Function
    End If
    If Me.txtCode(2).Text = "" Then Me.txtCode(2).Text = "0"
    If Me.txtCode(3).Text = "" Then Me.txtCode(3).Text = "00"
    If Me.txtCode(6).Text = "" Then Me.txtCode(6).Text = "0"
    If Me.txtCode(7).Text = "" Then Me.txtCode(7).Text = "00"
    If Me.txtCode(0).Text & Me.txtCode(1).Text & IIf(Me.txtCode(0).Text = "TF", Me.txtCode(8).Text, "") & Me.txtCode(2).Text & Me.txtCode(3).Text = _
        Me.txtCode(4).Text & Me.txtCode(5).Text & IIf(Me.txtCode(4).Text = "TF", Me.txtCode(9).Text, "") & Me.txtCode(6).Text & Me.txtCode(7).Text Then
        MsgBox "分割案及母案案號不可相同!!!", vbExclamation + vbOKOnly
        Me.txtCode(4).SetFocus
        Exit Function
    End If
    ReDim strCode(7) As String
    For i = 0 To 7
        strCode(i) = txtCode(i)
    Next i
    If strCode(0) = "TF" Then strCode(1) = Me.txtCode(8).Text
    If strCode(4) = "TF" Then strCode(5) = Me.txtCode(9).Text
    If ChkCaseReleate(strCode()) = False Then
        Me.txtCode(4).SetFocus
        Exit Function
    End If
End If
CheckDataValidate = True
End Function

Private Function ChkCaseReleate(ByRef strCode() As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strNation As String
Dim strNation_1 As String
Dim strCust(0 To 4) As String
Dim strCust_1(0 To 4) As String
Dim strPA08 As String
Dim strPA08_1 As String
Dim ii As Integer
Dim jj As Integer

    ChkCaseReleate = True
    StrSQLa = "Select PA09, PA26, PA27, PA28, PA29, PA30, PA08 From Patent Where PA01='" & strCode(0) & "' And PA02='" & strCode(1) & "' And PA03='" & strCode(2) & "' And PA04='" & strCode(3) & "' "
    StrSQLa = StrSQLa & " Union Select TM10, TM23, '', '', '', '', '' From Trademark Where TM01='" & strCode(0) & "' And TM02='" & strCode(1) & "' And TM03='" & strCode(2) & "' And TM04='" & strCode(3) & "' "
    StrSQLa = StrSQLa & " Union Select LC15, LC11,'', '', '', '', '' From Lawcase Where LC01='" & strCode(0) & "' And LC02='" & strCode(1) & "' And LC03='" & strCode(2) & "' And LC04='" & strCode(3) & "' "
    StrSQLa = StrSQLa & " Union Select '000', HC05, '', '', '', '', '' From Hirecase Where HC01='" & strCode(0) & "' And HC02='" & strCode(1) & "' And HC03='" & strCode(2) & "' And HC04='" & strCode(3) & "' "
    StrSQLa = StrSQLa & " Union Select SP09, SP08, '', '', '', '', '' From Servicepractice Where SP01='" & strCode(0) & "' And SP02='" & strCode(1) & "' And SP03='" & strCode(2) & "' And SP04='" & strCode(3) & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        strNation = "" & rsA.Fields(0).Value
        For ii = 0 To 4
            strCust(ii) = "" & rsA.Fields(ii + 1).Value
        Next ii
        strPA08 = "" & rsA.Fields(6).Value
    Else
        MsgBox "查無此分割案號資料!!!", vbExclamation + vbOKOnly
        ChkCaseReleate = False
        GoTo ExitFunction
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
        
    StrSQLa = "Select PA09, PA26, PA27, PA28, PA29, PA30, PA08 From Patent Where PA01='" & strCode(4) & "' And PA02='" & strCode(5) & "' And PA03='" & strCode(6) & "' And PA04='" & strCode(7) & "' "
    StrSQLa = StrSQLa & " Union Select TM10, TM23, '', '', '', '', '' From Trademark Where TM01='" & strCode(4) & "' And TM02='" & strCode(5) & "' And TM03='" & strCode(6) & "' And TM04='" & strCode(7) & "' "
    StrSQLa = StrSQLa & " Union Select LC15, LC11, '', '', '', '', '' From Lawcase Where LC01='" & strCode(4) & "' And LC02='" & strCode(5) & "' And LC03='" & strCode(6) & "' And LC04='" & strCode(7) & "' "
    StrSQLa = StrSQLa & " Union Select '000', HC05, '', '', '', '', '' From Hirecase Where HC01='" & strCode(4) & "' And HC02='" & strCode(5) & "' And HC03='" & strCode(6) & "' And HC04='" & strCode(7) & "' "
    StrSQLa = StrSQLa & " Union Select SP09, SP08, '', '', '' ,'', '' From Servicepractice Where SP01='" & strCode(4) & "' And SP02='" & strCode(5) & "' And SP03='" & strCode(6) & "' And SP04='" & strCode(7) & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        strNation_1 = "" & rsA.Fields(0).Value
        For ii = 0 To 4
            strCust_1(ii) = "" & rsA.Fields(ii + 1).Value
        Next ii
        strPA08_1 = "" & rsA.Fields(6).Value
    Else
        MsgBox "查無此母案案號資料!!!", vbExclamation + vbOKOnly
        ChkCaseReleate = False
        GoTo ExitFunction
    End If
    If strNation <> strNation_1 Then
        MsgBox "您輸入的分割案及母案申請國家不同!!!", vbExclamation + vbOKOnly
        ChkCaseReleate = False
        GoTo ExitFunction
    End If
    ChkCaseReleate = False
    For ii = 0 To 4
        For jj = 0 To 4
            If strCust(ii) <> "" And strCust_1(jj) <> "" Then
                If Left(strCust(ii), 6) = Left(strCust_1(jj), 6) Then
                    ChkCaseReleate = True
                End If
            End If
            If ChkCaseReleate = True Then Exit For
        Next jj
        If ChkCaseReleate = True Then Exit For
    Next ii
    If ChkCaseReleate = False Then
        MsgBox "您輸入的分割案及母案申請人非關係企業!!!", vbExclamation + vbOKOnly
        GoTo ExitFunction
    End If
    If InStr(strCode(0), "P") > 0 And InStr(strCode(4), "P") > 0 And strPA08 <> strPA08_1 Then
        MsgBox "您輸入的分割案及母案專利種類不同!!!", vbExclamation + vbOKOnly
        ChkCaseReleate = False
        GoTo ExitFunction
    End If
            
ExitFunction:
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Function
