VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090801_9 
   BorderStyle     =   1  '單線固定
   Caption         =   "對造為本所客戶陳報資料"
   ClientHeight    =   3020
   ClientLeft      =   2790
   ClientTop       =   3720
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3020
   ScaleWidth      =   5460
   Begin VB.CommandButton cmdCancl 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   405
      Left            =   4380
      TabIndex        =   6
      Top             =   45
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   3330
      TabIndex        =   5
      Top             =   45
      Width           =   930
   End
   Begin MSForms.TextBox txtCRL135 
      Height          =   1830
      Left            =   720
      TabIndex        =   4
      Top             =   1050
      Width           =   4635
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "8176;3228"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo134 
      Height          =   300
      Left            =   570
      TabIndex        =   1
      Top             =   600
      Width           =   1500
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "2646;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "原因 ："
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "(未向主管陳報時可不填)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3150
      TabIndex        =   2
      Top             =   600
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "已向                                     陳報可辦理"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3000
   End
End
Attribute VB_Name = "frm090801_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/22 改成Form2.0 (Combo134,txtCRL135)
'Create by Amy 2016/09/01
Option Explicit

Public m_blnCallQuery As Boolean  'Add By Sindy 2022/9/23 外部呼叫查詢
Dim m_PrevForm As Form '前一畫面


'Add By Sindy 2022/11/4
'取消
Private Sub cmdCancl_Click()
   Unload Me
End Sub

'確定
Private Sub cmdok_Click()
   If FormCheck = False Then Exit Sub
   
   'If Combo134 <> MsgText(601) Then m_PrevForm.m_stCRL134 = Combo134
   'If txtCRL135 <> MsgText(601) Then m_PrevForm.m_stCRL135 = txtCRL135
   m_PrevForm.m_stCRL134 = Combo134
   m_PrevForm.m_stCRL135 = txtCRL135
    
   Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
   
   'Me.Move 0, 0
   Screen.MousePointer = vbDefault
   Call SetCombo134 'Added by Lydia 2020/05/05
   'Add By Sindy 2022/9/13
   If m_PrevForm.m_stCRL134 <> MsgText(601) Then
      For i = 0 To Combo134.ListCount - 1
         If InStr(Combo134.List(i), m_PrevForm.m_stCRL134) > 0 Then
            Combo134.ListIndex = i
            Exit For
         End If
      Next i
   End If
   If m_PrevForm.m_stCRL135 <> MsgText(601) Then txtCRL135 = m_PrevForm.m_stCRL135
   'Sindy 2022/9/13 END
   
   'Add By Sindy 2022/9/23 外部呼叫查詢,鎖住欄位
   If m_blnCallQuery = True Then
      Combo134.Enabled = False
      'Combo134.Locked = True
      txtCRL135.Locked = True
      cmdOK.Visible = False: cmdCancl.Caption = "離開"
   End If
   '2022/9/23 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090801_9 = Nothing
End Sub

Private Function FormCheck() As Boolean
    FormCheck = False
   'Added by Morgan 2022/1/22 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2022/1/22
   
    If Combo134 = MsgText(601) And txtCRL135 = MsgText(601) Then
        MsgBox "主管及原因不可同時空白！", vbExclamation + vbOKOnly
        Exit Function
    End If
    FormCheck = True
End Function

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'Added by Lydia 2020/05/05 改成用特殊設定預設下拉選單
Private Sub SetCombo134()
Dim strA As String
Dim intA As Integer
Dim tmpArr As Variant

    Me.Combo134.Clear
    strA = Pub_GetSpecMan("對造客戶之陳報主管")
    If strA <> "" Then
        tmpArr = Split(strA, ";")
        For intA = 0 To UBound(tmpArr)
            If Trim(tmpArr(intA)) <> "" Then
                Me.Combo134.AddItem Trim(tmpArr(intA)), intA
            End If
        Next intA
    End If
End Sub
