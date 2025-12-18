VERSION 5.00
Begin VB.Form frm090203_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料查詢"
   ClientHeight    =   850
   ClientLeft      =   2760
   ClientTop       =   2390
   ClientWidth     =   2640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   850
   ScaleWidth      =   2640
   Begin VB.CommandButton Cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   48
      Width           =   1092
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   648
      TabIndex        =   2
      Top             =   48
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   1068
      MaxLength       =   5
      TabIndex        =   0
      Top             =   516
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "發文年月："
      Height          =   180
      Left            =   108
      TabIndex        =   1
      Top             =   516
      Width           =   972
   End
End
Attribute VB_Name = "frm090203_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/12/23 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim pemain As New ADODB.Recordset
Public UserStaff As String
Public ClickTime As String


Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
       Case 0
          If Text1.Text = "" Then MsgBox "未輸入發文年月", vbInformation: Text1.SetFocus: Exit Sub
          'Add By Cheng 2002/03/21
         If PUB_CheckKeyInYYMM(Me.Text1) = -1 Then
            Text1_GotFocus
            Me.Text1.SetFocus
            Exit Sub
         End If
          
          ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/14 清除查詢印表記錄檔欄位
          ClickTime = Text1.Text
          Me.Enabled = False
          Screen.MousePointer = vbHourglass
          frm090203_2.Show
          Screen.MousePointer = vbDefault
          Me.Enabled = True
          frm090203_1.Hide
       Case 1
          Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   If pemain.State = adStateOpen Then pemain.Close
   pemain.CursorLocation = adUseClient: pemain.CursorType = adOpenDynamic: pemain.LockType = adLockBatchOptimistic
   
   strExc(0) = "SELECT ST01 FROM STAFF WHERE ST02='" & strUserName & "'"
   pemain.Open strExc(0), cnnConnection
   If pemain.BOF And pemain.EOF Then MsgBox "無此LOGIN人員之資料", vbInformation: Unload Me
   
   UserStaff = pemain.Fields(0).Value
   pemain.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090203_1 = Nothing
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
