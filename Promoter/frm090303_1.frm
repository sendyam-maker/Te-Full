VERSION 5.00
Begin VB.Form frm090303_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料查詢"
   ClientHeight    =   885
   ClientLeft      =   2370
   ClientTop       =   2505
   ClientWidth     =   2580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   2580
   Begin VB.CommandButton Cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Index           =   1
      Left            =   1428
      TabIndex        =   3
      Top             =   48
      Width           =   1092
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   636
      TabIndex        =   2
      Top             =   48
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   0
      Top             =   516
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "發文年月："
      Height          =   180
      Left            =   72
      TabIndex        =   1
      Top             =   516
      Width           =   960
   End
End
Attribute VB_Name = "frm090303_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/28 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim pemain As New ADODB.Recordset
Public UserStaff As String
Public ClickTime As String
Dim TmpStates As String
Public TextOk As Boolean
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
       Case 0
          If Text1.Text = "" Then MsgBox "未輸入發文年月", vbInformation: Exit Sub
         'Add By Cheng 2002/03/21
         If PUB_CheckKeyInYYMM(Me.Text1) = -1 Then
            Me.Text1.SetFocus
            Text1_GotFocus
            Exit Sub
         End If
          ClickTime = Text1.Text
          TmpStates = ProState
          ProState = 4
          Screen.MousePointer = vbHourglass
          Me.Enabled = False
          ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/16 清除查詢印表記錄檔欄位
          pub_QL05 = pub_QL05 & ";" & Label1 & Text1.Text 'Add By Sindy 2010/12/16
          frm090711.Show
          Me.Hide
          Me.Enabled = True
          Screen.MousePointer = vbDefault
       Case 1
          If TmpStates <> "" Then
            ProState = TmpStates
          End If
          Unload Me
End Select
End Sub

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_Load()
m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
TmpStates = ""
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090303_1 = Nothing
End Sub

Private Sub Text1_GotFocus()
Me.Text1.SelStart = 0
Me.Text1.SelLength = Len(Me.Text1.Text)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Text1 <> "" Then
If CheckIsTaiwanDate(Text1.Text + "01") = False Then
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Cancel = True
Else
Cancel = False
End If
End If
End Sub

Public Sub Process()
frm090711.Show
frm090711.Hide
frm090711.Combo1.Clear
frm090711.Combo1.AddItem strUserNum, 0
frm090711.Combo1.Text = frm090711.Combo1.List(0)
End Sub
