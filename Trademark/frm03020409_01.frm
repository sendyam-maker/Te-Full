VERSION 5.00
Begin VB.Form frm03020409_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務結果輸入"
   ClientHeight    =   1875
   ClientLeft      =   -90
   ClientTop       =   3525
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4515
   Begin VB.TextBox textCP05 
      Height          =   264
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1200
      Width           =   2892
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3480
      TabIndex        =   7
      Top             =   72
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2544
      TabIndex        =   6
      Top             =   72
      Width           =   912
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   0
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   3
      Top             =   720
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   4
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   1
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "來函收文日 :"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   972
   End
End
Attribute VB_Name = "frm03020409_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/09/13 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 總收文號
Dim m_CP09 As String

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
'   textCP05 = TAIWANDATE(SystemDate())
   textCP05 = strSrvDate(2)
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   ' 本所案號的系統別
   If IsEmptyText(textTM01) = True Then
      strTit = "檢核資料"
      strMsg = "請先本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   If IsEmptyText(textTM02) = True Then
      strTit = "檢核資料"
      strMsg = "請先本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   ' 檢查欄位的資料是否都已經輸入正確
   If CheckDataValid = False Then
      GoTo EXITSUB
   End If
   
   strTM01 = Trim(textTM01)
   strTM02 = Trim(textTM02)
   strTM03 = Trim(textTM03)
   If IsEmptyText(strTM03) = True Then: strTM03 = "0"
   strTM04 = Trim(textTM04)
   If IsEmptyText(strTM04) = True Then: strTM04 = "00"
   
   ' 檢查所輸入的資料是否合乎資料庫的條件
   Select Case textTM01
      Case "S":
         ' 讀取商標基本檔
         strSql = "SELECT * FROM ServicePractice " & _
                  "WHERE SP01 = '" & strTM01 & "' AND " & _
                        "SP02 = '" & strTM02 & "' AND " & _
                        "SP03 = '" & strTM03 & "' AND " & _
                        "SP04 = '" & strTM04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無符合的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         rsTmp.Close
         
         ' 讀取案件進度檔(案件性質為申請)
         strSql = "SELECT * FROM CaseProgress " & _
                  "WHERE CP01 = '" & strTM01 & "' AND " & _
                        "CP02 = '" & strTM02 & "' AND " & _
                        "CP03 = '" & strTM03 & "' AND " & _
                        "CP04 = '" & strTM04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無符合的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         m_CP09 = rsTmp.Fields("CP09")
         rsTmp.Close
      Case Else:
         GoTo EXITSUB
   End Select
         
   ' 顯示下一個畫面
   DisplayNextForm
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub DisplayNextForm()
   ' 本所案號
   frm03020409_02.SetData 0, Trim(textTM01), True
   frm03020409_02.SetData 1, Trim(textTM02), False
   If IsEmptyText(textTM03) = True Then
      frm03020409_02.SetData 2, "0", False
   Else
      frm03020409_02.SetData 2, Trim(textTM03), False
   End If
   If IsEmptyText(textTM04) = True Then
      frm03020409_02.SetData 3, "00", False
   Else
      frm03020409_02.SetData 3, Trim(textTM04), False
   End If
   frm03020409_02.SetData 4, textCP05, False
   Me.Hide
   frm03020409_02.Show
   frm03020409_02.QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm03020409_01 = Nothing
End Sub

' 本所案號中的系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM02_2.Visible = False
   textTM02_2.Locked = True
   textTM02_2.TabStop = False
   textTM02.MaxLength = 6
      
   If IsEmptyText(textTM01) = False Then
      Select Case textTM01
         Case "S":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "本所案號中的系統別不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM01_GotFocus
      End Select
   End If
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As TextBox)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Private Sub textTM01_GotFocus()
   InverseAll textTM01
End Sub

Private Sub textTM02_GotFocus()
   InverseAll textTM02
End Sub

Private Sub textTM03_GotFocus()
   InverseAll textTM03
End Sub

Private Sub textTM04_GotFocus()
   InverseAll textTM04
End Sub



