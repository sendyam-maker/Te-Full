VERSION 5.00
Begin VB.Form frm081016 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人已收達/已提申"
   ClientHeight    =   2100
   ClientLeft      =   2790
   ClientTop       =   2130
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4575
   Begin VB.CommandButton ComSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2712
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton ComBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3540
      TabIndex        =   5
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox Text8 
      Height          =   300
      Left            =   1368
      MaxLength       =   3
      TabIndex        =   0
      Top             =   960
      Width           =   550
   End
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   1972
      MaxLength       =   6
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   2761
      MaxLength       =   1
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   3192
      MaxLength       =   2
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   1008
      Width           =   900
   End
End
Attribute VB_Name = "frm081016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/09/22 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Private Sub ComBack_Click()
   Unload Me
End Sub

Private Sub ComSure_Click()
  Dim St As String
  Dim strCP01 As String
  Dim strCP02 As String
  Dim strCP03 As String
  Dim strCP04 As String
  
  If Text8.Text = "" Or Text7.Text = "" Then
     MsgBox "本所案號不可空白!", vbInformation, "代理人已收達/已提申"
     Text8.SetFocus
     Exit Sub
  End If
  strCP01 = Text8.Text
  strCP02 = Text7.Text
  If Text6.Text = "" Then
     strCP03 = "0"
  Else
      strCP03 = Text6.Text
  End If
      
  If Text5.Text = "" Then
     strCP04 = "00"
  Else
      strCP04 = Text5.Text
  End If
   
  St = strCP01 & strCP02 & strCP03 & strCP04
  strExc(0) = "SELECT LC11,LC05,LC06,LC07,LC15 FROM LAWCASE WHERE " & _
              "LC01 ='" & strCP01 & "'" & _
              " AND LC02 ='" & strCP02 & "'" & _
              " AND LC03 ='" & strCP03 & "'" & _
              " AND LC04 ='" & strCP04 & "'"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 1 Then Exit Sub
   frm081017.Tag = St
   Me.Hide
   frm081017.ShowData
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm081016 = Nothing
End Sub

Private Sub Text8_GotFocus()
   TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8.Text <> "" Then
      If Text8.Text <> "CFL" Then
         DataErrorMessage 1, "系統類別"
         Cancel = True
         TextInverse Text8
      End If
   End If
End Sub
