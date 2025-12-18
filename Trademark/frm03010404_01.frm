VERSION 5.00
Begin VB.Form frm03010404_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "發回補理由/發回補答辨"
   ClientHeight    =   1770
   ClientLeft      =   -170
   ClientTop       =   3990
   ClientWidth     =   4810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4810
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3780
      MaxLength       =   2
      TabIndex        =   4
      Top             =   720
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3420
      MaxLength       =   1
      TabIndex        =   3
      Top             =   720
      Width           =   372
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   0
      Top             =   720
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2820
      TabIndex        =   6
      Top             =   70
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3792
      TabIndex        =   7
      Top             =   70
      Width           =   912
   End
   Begin VB.TextBox textCP05 
      Height          =   264
      Left            =   1620
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1080
      Width           =   2892
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2340
      MaxLength       =   6
      TabIndex        =   1
      Top             =   720
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Left            =   180
      TabIndex        =   9
      Top             =   720
      Width           =   1332
   End
   Begin VB.Label Label2 
      Caption         =   "來函收文日 :"
      Height          =   252
      Left            =   180
      TabIndex        =   8
      Top             =   1080
      Width           =   1452
   End
End
Attribute VB_Name = "frm03010404_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/08/13 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

'Add By Sindy 2023/5/2
Public m_RDate As String
Public strTM01 As String
Public strTM02 As String
Public strTM03 As String
Public strTM04 As String
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
Public m_Done As Boolean
'2023/5/2 END


'Add By Sindy 2023/5/2
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2003/07/09
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
    'Modify By Cheng 2003/08/20
    '移至Form_Load時做
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
'    '初始化序號
'    pub_AddressListSN = 0
    Unload Me
End Sub

Private Sub Form_Activate()
   'Added by Sindy 2023/5/2
   If m_strIR01 <> "" And m_Done = False Then
      textTM01.Text = strTM01
      If strTM01 = "TF" Then
         textTM02.Text = Left(strTM02, 5)
         textTM02_2.Text = Mid(strTM02, 6)
      Else
         textTM02.Text = strTM02
      End If
      textTM03.Text = strTM03
      textTM04.Text = strTM04
      textCP05.Text = DBDATE(m_RDate)
      cmdOK.Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2023/5/2 END
End Sub

Private Sub Form_Load()
    ' 來函收文日預設為系統日
'    textCP05 = DBDATE(SystemDate())
    textCP05 = strSrvDate(1)
    MoveFormToCenter Me
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
   
   ' 來函收文日不可為空白
   If IsEmptyText(textCP05) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入來函收文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
    'add by nickc 2006/03/17 加入驗證
    Dim Cancel As Boolean
    Cancel = False
    textCP05_Validate Cancel
    If Cancel = True Then GoTo EXITSUB
    
   CheckDataValid = True
EXITSUB:
End Function

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
'   Dim strTM01 As String
'   Dim strTM02 As String
'   Dim strTM03 As String
'   Dim strTM04 As String

   'Add By Sindy 2023/5/2
   If m_strIR01 <> "" Then
      If strTM01 & strTM02 & strTM03 & strTM04 <> textTM01 & IIf(textTM01 = "TF", textTM02 & textTM02_2, textTM02) & textTM03 & textTM04 Then
         MsgBox "信件輸入必須與信件本所案號(" & strTM01 & strTM02 & strTM03 & strTM04 & ")一致！"
         Exit Sub
      End If
   End If
   '2023/5/2 END
   
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
      Case "CFT":
         ' 讀取商標基本檔
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "TM02 = '" & strTM02 & "' AND " & _
                        "TM03 = '" & strTM03 & "' AND " & _
                        "TM04 = '" & strTM04 & "' "
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
      Case Else:
         GoTo EXITSUB
   End Select
         
   ' 顯示下一個畫面
   DisplayNextForm
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub DisplayNextForm()
   frm03010404_02.SetData 0, Trim(textTM01), True
   frm03010404_02.SetData 1, Trim(textTM02), False
   If IsEmptyText(textTM03) = True Then
      frm03010404_02.SetData 2, "0", False
   Else
      frm03010404_02.SetData 2, Trim(textTM03), False
   End If
   If IsEmptyText(textTM04) = True Then
      frm03010404_02.SetData 3, "00", False
   Else
      frm03010404_02.SetData 3, Trim(textTM04), False
   End If
   frm03010404_02.SetData 4, textCP05, False
   'Add By Sindy 2023/5/2
   If Not m_PrevForm Is Nothing Then
      Call frm03010404_02.SetParent(m_PrevForm)
   End If
   frm03010404_02.m_strIR01 = m_strIR01
   frm03010404_02.m_strIR02 = m_strIR02
   frm03010404_02.m_strIR03 = m_strIR03
   frm03010404_02.m_strIR04 = m_strIR04
   '2023/5/2 END
   Me.Hide
   frm03010404_02.Show
   frm03010404_02.QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '列印接洽接案單
   PUB_PrintCaseCloseSheet strUserNum
   '刪除暫存資料
   PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2023/5/2
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2023/5/2 END
   
   'Add By Cheng 2002/07/19
   Set frm03010404_01 = Nothing
End Sub

' 來函收文日
Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If CheckIsDate(textCP05, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的來函收文日"
         strTit = "來函收文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
         GoTo EXITSUB
      End If
      'edit by nickc  2006/03/17
      'If Val(textCP05) > Val(ChangeWDateStringToWString(Date)) Then
      If Val(textCP05) > Val(strSrvDate(1)) Then
         Cancel = True
         strMsg = "來函收文日不可超過系統日"
         strTit = "來函收文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

' 本所案號中的系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      Select Case textTM01
         Case "CFT":
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "本所案號中的系統別不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM01_GotFocus
      End Select
   Else
      textTM02_2.Visible = False
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02.MaxLength = 6
   End If
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 將所有的文字反白
Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub


