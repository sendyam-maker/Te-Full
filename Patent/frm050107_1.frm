VERSION 5.00
Begin VB.Form frm050107_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "美國IDS資料對照維護"
   ClientHeight    =   3945
   ClientLeft      =   330
   ClientTop       =   2205
   ClientWidth     =   5970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5970
   Begin VB.OptionButton optChoose 
      Caption         =   "多筆查詢條件"
      CausesValidation=   0   'False
      Height          =   252
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   2580
      Width           =   1455
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "單筆維護"
      CausesValidation=   0   'False
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   780
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4176
      TabIndex        =   14
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5004
      TabIndex        =   15
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame fraChoose 
      Enabled         =   0   'False
      Height          =   975
      Index           =   1
      Left            =   180
      TabIndex        =   20
      Top             =   2820
      Width           =   5652
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   10
         Left            =   1800
         MaxLength       =   8
         TabIndex        =   12
         Top             =   600
         Width           =   972
      End
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   9
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   11
         Top             =   240
         Width           =   972
      End
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   11
         Left            =   3120
         MaxLength       =   8
         TabIndex        =   13
         Top             =   600
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "(西元年月日)"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblEnginer 
         Height          =   252
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Width           =   2172
      End
      Begin VB.Label Label3 
         Caption         =   "EPC,英國案收文日："
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1692
      End
      Begin VB.Label Label3 
         Caption         =   "美國案工程師："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   2880
         X2              =   3000
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame fraChoose 
      Height          =   1332
      Index           =   0
      Left            =   180
      TabIndex        =   16
      Top             =   1020
      Width           =   5652
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   2
         Top             =   240
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   3
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   5
         Top             =   240
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   4
         Top             =   240
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   3
         Top             =   240
         Width           =   852
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   5
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   7
         Top             =   600
         Width           =   852
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   6
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   8
         Top             =   600
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   7
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   9
         Top             =   600
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   4
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   6
         Top             =   600
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   270
         Index           =   8
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   10
         Top             =   960
         Width           =   372
      End
      Begin VB.Label Label1 
         Caption         =   "美國案號："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "EPC,英國案號："
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "功能代號：           (1.新增  2.修改  4.刪除  5.查詢 )"
         Height          =   252
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   3972
      End
   End
End
Attribute VB_Name = "frm050107_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/3 改成Form2.0 (無)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Public intWhereToGo As Integer

Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer, varSaveCursor, strCode(7) As String

Select Case Index
             Case 0 '確定
                        varSaveCursor = Screen.MousePointer
                        Screen.MousePointer = vbHourglass
                        For i = 0 To 11
                               If txtCode(i).Enabled Then
                                  If CheckKeyIn(i) = False Then
                                     '本所案號錯誤時,將Cursor跳回系統別欄位
                                     If i = 3 Or i = 7 Then i = i - 3
                                     txtCode(i).SetFocus
                                     txtCode_GotFocus i
                                     Exit For
                                  End If
                               End If
                        Next
                        If i = 12 Then
                           If optChoose(0).Value Then
                              If txtCode(8) = "4" Then
                                 For i = 0 To 7
                                    strCode(i) = txtCode(i)
                                 Next
                                 'edit by nickc 2007/02/05 不用 dll 了
                                 'If obj003.ChkExist(strCode(), 1) Then
                                 If Cls003ChkExist(strCode(), 1) Then
                                    If MsgBox("是否要刪除此筆資料 ?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
                                       '911105 nick transation
                                       cnnConnection.BeginTrans
                                       'edit by nickc 2007/02/05 不用 dll 了
                                       'If obj003.DeleteCaseRelation(strCode(), 1) Then
                                       If Cls003DeleteCaseRelation(strCode(), 1) Then
                                          '911105 nick transation
                                          cnnConnection.CommitTrans
                                          For i = 0 To 8
                                             txtCode(i) = ""
                                          Next
                                          txtCode(0).SetFocus
                                       Else
                                          '911105 nick transation
                                          cnnConnection.RollbackTrans
                                       End If
                                    End If
                                 End If
                              ElseIf txtCode(8) = "1" Then
                                 For i = 0 To 7
                                    strCode(i) = txtCode(i)
                                 Next
                                 'edit by nickc 2007/02/05 不用 dll 了
                                 'If obj003.ChkCaseMap(strCode, 1) Then
                                 If Cls003ChkCaseMap(strCode, 1) Then
                                    GoTo A0
                                 Else
                                    txtCode(0).SetFocus
                                 End If
                              Else
A0:                              frm050107_2.intWhereToGo = 0
                                 frm050107_2.strCode1 = txtCode(0)
                                 frm050107_2.strCode2 = txtCode(1)
                                 frm050107_2.strCode3 = txtCode(2)
                                 frm050107_2.strCode4 = txtCode(3)
                                 frm050107_2.strCode5 = txtCode(4)
                                 frm050107_2.strCode6 = txtCode(5)
                                 frm050107_2.strCode7 = txtCode(6)
                                 frm050107_2.strCode8 = txtCode(7)
                                 frm050107_2.intChoose = Val(txtCode(8))
                                 frm050107_2.Show
                                 Me.Hide
                              End If
                           Else
                              frm050107_3.lblEnginer = txtCode(9)
                              frm050107_3.lblEnginerName = lblEnginer
                              frm050107_3.lblDate(0) = ChangeWStringToWDateString(txtCode(10))
                              frm050107_3.lblDate(1) = ChangeWStringToWDateString(txtCode(11))
                              Me.Hide
                           End If
                        End If
                        Screen.MousePointer = varSaveCursor
             Case 1
                        Unload Me
End Select
End Sub
Private Sub Form_Activate()
If optChoose(0) Then
   txtCode(0).SetFocus
   txtCode(8) = "1"
Else
   txtCode(9).SetFocus
End If
End Sub
Private Sub Form_Load()
MoveFormToCenter Me
'txtCode(10) = GetTodayDate
'txtCode(11) = GetTodayDate
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Select Case intWhereToGo
      Case 1
          frm050101_2.Show
      ' 91.09.11 modify by louis
      Case 2
         frm010012_05.Show
   End Select
   'Add By Cheng 2002/07/18
   Set frm050107_1 = Nothing
End Sub

Private Sub optChoose_Click(Index As Integer)
fraChoose(Index).Enabled = True
fraChoose((Index + 1) Mod 2).Enabled = False
If Index = 0 Then
   txtCode(0).SetFocus
Else
   txtCode(9).SetFocus
End If
End Sub
Private Sub txtCode_Change(Index As Integer)
Select Case Index
             Case 9
                       lblEnginer = ""
End Select
End Sub
Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index))
End Sub
Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
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

Select Case intIndex
             Case 0
                  If txtCode(intIndex) = "CFP" Then
                     CheckKeyIn = True
                  Else
                     MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
                  End If
             Case 4
                  If txtCode(intIndex) = "CFP" Then
                     CheckKeyIn = True
                  Else
                     MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
                  End If
             Case 3, 7
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.CheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2), _
                             IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
                        If ClsPDCheckCaseCodeIsExist(txtCode(intIndex - 3), txtCode(intIndex - 2), _
                             IIf(txtCode(intIndex - 1) = "", "0", txtCode(intIndex - 1)), IIf(txtCode(intIndex) = "", "00", txtCode(intIndex))) Then
                           CheckKeyIn = True
                        End If
             Case 8
                        If Val(txtCode(intIndex)) = 1 Or Val(txtCode(intIndex)) = 2 Or Val(txtCode(intIndex)) = 4 Or Val(txtCode(intIndex)) = 5 Then
                           CheckKeyIn = True
                        Else
                           ShowMsg MsgText(9198)
                        End If
             Case 9
                        If txtCode(intIndex) = "" Then
                           CheckKeyIn = True
                        'edit by nickc 2007/02/02 不用 dll 了
                        'ElseIf objPublicData.GetStaff(txtCode(intIndex).Text, strTemp) Then
                        ElseIf ClsPDGetStaff(txtCode(intIndex).Text, strTemp) Then
                           lblEnginer = strTemp
                           CheckKeyIn = True
                        End If
             Case 10, 11
                        If txtCode(intIndex) = "" Then
                           CheckKeyIn = True
                        ElseIf CheckIsDate(txtCode(intIndex)) Then
                           CheckKeyIn = True
                        End If
                        If intIndex = 11 And txtCode(10) <> "" And txtCode(11) = "" Then
                           ShowMsg MsgText(9169)
                           CheckKeyIn = False
                        ElseIf txtCode(11) <> "" And Val(txtCode(10)) > Val(txtCode(11)) Then
                           ShowMsg MsgText(9170)
                           CheckKeyIn = False
                        End If
             Case Else
                        CheckKeyIn = True
End Select
End Function
