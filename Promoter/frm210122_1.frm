VERSION 5.00
Begin VB.Form frm210122_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶應收帳款明細-預定收款日期更改"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdok 
      Caption         =   "取消(&X)"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   3630
      TabIndex        =   2
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Height          =   375
      Index           =   0
      Left            =   2580
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   0
      Top             =   420
      Width           =   945
   End
   Begin VB.Label Label4 
      Caption         =   "注意！輸入後按 確定 存檔，僅主管可更改　　　若不輸入請按  取消。"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   30
      TabIndex        =   6
      Top             =   2730
      Width           =   4560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(請輸入民國日期不含 /)"
      Height          =   180
      Left            =   2250
      TabIndex        =   5
      Top             =   480
      Width           =   1830
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "預定收款日："
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1875
      Left            =   90
      TabIndex        =   3
      Top             =   780
      Width           =   4515
   End
End
Attribute VB_Name = "frm210122_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/11 Form2.0已檢查 (無需修改的物件); 已不使用
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Public UpForm As Form
Dim oKey As String

Private Sub cmdOK_Click(Index As Integer)
Dim rtCnt As Integer
Dim m_rs As New ADODB.Recordset

   Select Case Index
   Case 0
      '2008/7/10 ADD BY SONIA
      If ChkDate(txtDate) = False Then
         txtDate_GotFocus
         txtDate.SetFocus
         Exit Sub
      '2011/6/29 add by sonia
      ElseIf DBDATE(txtDate) < strSrvDate(1) Then
         MsgBox "預定收款日不可小於系統日！", vbCritical, "輸入錯誤！"
         txtDate_GotFocus
         txtDate.SetFocus
         Exit Sub
      '2011/6/29 end
      End If
      '2008/7/10
      If MsgBox("確定修改???", vbYesNo) = vbYes Then
            Set m_rs = New ADODB.Recordset
            If m_rs.State = 1 Then m_rs.Close
            m_rs.CursorLocation = adUseClient
            'Modified by Morgan 2011/10/31 考慮多對多收據情形改用收文號抓
            'm_rs.Open "select * from caseprogress where cp60='" & oKey & "' ", cnnConnection, adOpenStatic, adLockReadOnly
            m_rs.Open "select * from caseprogress where cp09 in (select a0j01 from acc0j0 where a0j13='" & oKey & "') ", cnnConnection, adOpenStatic, adLockReadOnly
            If Not m_rs.EOF And Not m_rs.BOF Then
                m_rs.MoveFirst
                Do While Not m_rs.EOF
                   If txtDate = "" Then
                      cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CheckStr(m_rs.Fields("CP09")) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "', null from receivablesday where rd01='" & CheckStr(m_rs.Fields("CP09")) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CheckStr(m_rs.Fields("CP09")) & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & Val(DBDATE(txtDate)) & " ", rtCnt
                   Else
                      cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CheckStr(m_rs.Fields("CP09")) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtDate) & " from receivablesday where rd01='" & CheckStr(m_rs.Fields("CP09")) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CheckStr(m_rs.Fields("CP09")) & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & Val(DBDATE(txtDate)) & " ", rtCnt
                   End If
                   If rtCnt = 0 Then
                      If txtDate = "" Then
                         cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CheckStr(m_rs.Fields("CP09")) & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "', null from dual "
                      Else
                         cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CheckStr(m_rs.Fields("CP09")) & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & Val(DBDATE(txtDate)) & " from dual "
                      End If
                   End If
                   m_rs.MoveNext
                Loop
            End If
            UpForm.SetDate = txtDate
        End If
        Unload Me
   Case 1
        'UpForm.SetDate = ""   '2015/8/24 cancel by sonia 否則按取消時會清掉前畫面grid中的原來日期
        Unload Me
   Case Else
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtDate = UpForm.SetDate
   Label1.Caption = UpForm.SetData
   oKey = UpForm.SetKey
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm210122_1 = Nothing
End Sub

Private Sub txtDate_GotFocus()
   InverseTextBox txtDate
End Sub


