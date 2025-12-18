VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210133_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "簽核人員資料"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   3765
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Height          =   345
      Index           =   1
      Left            =   1590
      TabIndex        =   6
      Top             =   150
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   345
      Index           =   0
      Left            =   2640
      TabIndex        =   7
      Top             =   150
      Width           =   975
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Index           =   6
      Left            =   1590
      TabIndex        =   5
      Top             =   3240
      Width           =   1725
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3043;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Index           =   5
      Left            =   1590
      TabIndex        =   4
      Top             =   2868
      Width           =   1725
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3043;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Index           =   4
      Left            =   1590
      TabIndex        =   3
      Top             =   2496
      Width           =   1725
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3043;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Index           =   3
      Left            =   1590
      TabIndex        =   2
      Top             =   2124
      Width           =   1725
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3043;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Index           =   2
      Left            =   1590
      TabIndex        =   1
      Top             =   1752
      Width           =   1725
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3043;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Index           =   1
      Left            =   1590
      TabIndex        =   0
      Top             =   1380
      Width           =   1725
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3043;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblST02 
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   840
      Width           =   1635
      VariousPropertyBits=   27
      Size            =   "2884;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "人員："
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   540
   End
   Begin VB.Label lblST01 
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "簽核人員1："
      Height          =   255
      Index           =   10
      Left            =   570
      TabIndex        =   13
      Top             =   1410
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "簽核人員2："
      Height          =   255
      Index           =   9
      Left            =   570
      TabIndex        =   12
      Top             =   1776
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "簽核人員3："
      Height          =   255
      Index           =   8
      Left            =   570
      TabIndex        =   11
      Top             =   2142
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "簽核人員4："
      Height          =   255
      Index           =   7
      Left            =   570
      TabIndex        =   10
      Top             =   2508
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "簽核人員5："
      Height          =   255
      Index           =   0
      Left            =   570
      TabIndex        =   9
      Top             =   2874
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "簽核人員6："
      Height          =   255
      Index           =   2
      Left            =   570
      TabIndex        =   8
      Top             =   3240
      Width           =   990
   End
End
Attribute VB_Name = "frm210133_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/07 改成Form2.0 ; lblST02、Combo2(index)
'Create by Sindy 2015/1/8
Option Explicit

' 變數宣告區
Dim i As Integer
Public m_SetFlowKind As String '1.結案單 2.銷案銷帳單 3.接洽單


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 1
         If TxtValidate = False Then Exit Sub
         frm210133.m_SetFlowEmp1 = Left(Combo2(1).Text, 5) '設定簽核人員1
   End Select
   Unload Me
End Sub

Private Sub Combo2_GotFocus(Index As Integer)
   InverseTextBox Combo2(Index)
End Sub

'Modified by Lydia 2021/10/07 改成Form 2.0
'Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
   If Combo2(Index).Text > "" And Len(Trim(Combo2(Index).Text)) = 5 Then
      '抓取員工姓名
      Combo2(Index).Text = SetCboStaffName(Combo2(Index).Text)
   End If
End Sub

Private Sub Combo2_Validate(Index As Integer, Cancel As Boolean)
   If Combo2(Index) <> "" Then
      If Left(Combo2(Index), 5) = lblST01 Then
         MsgBox "不可為本人！", vbExclamation
         Call Combo2_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(Combo2(Index), 5)) = True Then
         Call Combo2_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查 員工不可為”不寄信”
      If ChkStaffST14(Left(Combo2(Index), 5)) = True Then
         Call Combo2_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查輸入順序
      If (Trim(Combo2(2)) <> "" And Trim(Combo2(1)) = "") Or _
         (Trim(Combo2(3)) <> "" And Trim(Combo2(2)) = "") Or _
         (Trim(Combo2(4)) <> "" And Trim(Combo2(3)) = "") Or _
         (Trim(Combo2(5)) <> "" And Trim(Combo2(4)) = "") Or _
         (Trim(Combo2(6)) <> "" And Trim(Combo2(5)) = "") Then
         MsgBox "請依序輸入簽核人員！", vbExclamation
         Combo2(Index).SetFocus
         Call Combo2_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      For i = 1 To 6
         If i <> Index Then
            If Trim(Combo2(i)) <> "" And Left(Trim(Combo2(i)), 5) = Left(Trim(Combo2(Index)), 5) Then
               MsgBox "資料重覆！", vbExclamation
               Combo2(Index).SetFocus
               Call Combo2_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Next i
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   For i = 1 To Combo2.UBound
      Combo2(i).Enabled = False
   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210133_1 = Nothing
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean

TxtValidate = False

For i = 1 To Combo2.UBound
   If Combo2(i).Enabled = True Then
      Cancel = False
      Combo2_Validate i, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next i

TxtValidate = True
End Function

' 將資料庫中的資料更新到所有欄位中
Public Sub doQuery()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   If InStr(Flow_可改簽核人員1, lblST01) > 0 Then
      Combo2(1).Enabled = True
      cmdOK(1).Visible = True
   Else
      cmdOK(1).Visible = False
   End If
   
   strSql = "select FLOW001.*" & _
            ",s1.ST02 s1_ST02,s2.ST02 s2_ST02,s3.ST02 s3_ST02,s4.ST02 s4_ST02,s5.ST02 s5_ST02,s6.ST02 s6_ST02 " & _
            "from FLOW001" & _
            ",STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6 " & _
            "where F0101='" & lblST01 & "' and F0102='" & m_SetFlowKind & "' " & _
            "and F0103=s1.ST01(+) " & _
            "and F0104=s2.ST01(+) " & _
            "and F0105=s3.ST01(+) " & _
            "and F0106=s4.ST01(+) " & _
            "and F0107=s5.ST01(+) " & _
            "and F0108=s6.ST01(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If frm210133.m_SetFlowEmp1 <> "" Then
         Combo2(1).Text = SetCboStaffName(frm210133.m_SetFlowEmp1)
      Else
         If IsNull(rsTmp.Fields("F0103")) = False Then Combo2(1).Text = Left(Trim(rsTmp.Fields("F0103")) & Space(5), 7) & rsTmp.Fields("s1_ST02")
      End If
      If IsNull(rsTmp.Fields("F0104")) = False Then Combo2(2).Text = Left(Trim(rsTmp.Fields("F0104")) & Space(5), 7) & rsTmp.Fields("s2_ST02")
      If IsNull(rsTmp.Fields("F0105")) = False Then Combo2(3).Text = Left(Trim(rsTmp.Fields("F0105")) & Space(5), 7) & rsTmp.Fields("s3_ST02")
      If IsNull(rsTmp.Fields("F0106")) = False Then Combo2(4).Text = Left(Trim(rsTmp.Fields("F0106")) & Space(5), 7) & rsTmp.Fields("s4_ST02")
      If IsNull(rsTmp.Fields("F0107")) = False Then Combo2(5).Text = Left(Trim(rsTmp.Fields("F0107")) & Space(5), 7) & rsTmp.Fields("s5_ST02")
      If IsNull(rsTmp.Fields("F0108")) = False Then Combo2(6).Text = Left(Trim(rsTmp.Fields("F0108")) & Space(5), 7) & rsTmp.Fields("s6_ST02")
   End If
   rsTmp.Close
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault

EXITSUB:
   Set rsTmp = Nothing
End Sub
