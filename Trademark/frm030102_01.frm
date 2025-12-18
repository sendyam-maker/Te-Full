VERSION 5.00
Begin VB.Form frm030102_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標申請案號輸入"
   ClientHeight    =   1690
   ClientLeft      =   4470
   ClientTop       =   3440
   ClientWidth     =   4790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1690
   ScaleWidth      =   4790
   Begin VB.TextBox textCP05 
      Height          =   264
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1020
      Width           =   2892
   End
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   660
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   4
      Top             =   660
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   3
      Top             =   660
      Width           =   372
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   0
      Top             =   660
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2820
      TabIndex        =   6
      Top             =   48
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3780
      TabIndex        =   7
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   1
      Top             =   660
      Width           =   1092
   End
   Begin VB.Label Label2 
      Caption         =   "來函收文日 :"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   1020
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   660
      Width           =   1332
   End
End
Attribute VB_Name = "frm030102_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/08/10 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit
'Add By Sindy 2023/4/24
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
'2023/4/24 END


'Add By Sindy 2023/4/24
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   'Added by Sindy 2023/4/24
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
   '2023/4/24 END
End Sub

Private Sub Form_Load()
   '來函收文日預設為系統日
   'textCP05 = DBDATE(SystemDate())
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
   'Add By Cheng 2004/02/20
   Dim strTM10 As String '申請國家
   'End
       
   'Add By Sindy 2023/4/24
   If m_strIR01 <> "" Then
      If strTM01 & strTM02 & strTM03 & strTM04 <> textTM01 & IIf(textTM01 = "TF", textTM02 & textTM02_2, textTM02) & textTM03 & textTM04 Then
         MsgBox "信件輸入必須與信件本所案號(" & strTM01 & strTM02 & strTM03 & strTM04 & ")一致！"
         Exit Sub
      End If
   End If
   '2023/4/24 END
   
   strTM10 = ""
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
      Case "CFT", "FCT":
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
        strTM10 = "" & rsTmp("TM10").Value
         rsTmp.Close
      Case Else:
         ' 讀取服務業務基本檔
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
        strTM10 = "" & rsTmp("SP09").Value
         rsTmp.Close
      'Case Else:
         'GoTo ExitSub
   End Select
   'Modify By Cheng 2002/07/11
   '原判斷CP31='Y', 現改為CP10='101'
'   strSQL = "SELECT * FROM CaseProgress " & _
'            "WHERE CP01 = '" & strTM01 & "' AND " & _
'                  "CP02 = '" & strTM02 & "' AND " & _
'                  "CP03 = '" & strTM03 & "' AND " & _
'                  "CP04 = '" & strTM04 & "' AND " & _
'                  "CP31 = '" & "Y" & "' "
    'Add By Cheng 2004/02/20
    strSql = ""
    '93.12.1 恢復 BY SONIA  CFT-009900
    '    'edit by nick 2004/09/01 不限制國家
    ''Select Case strTM10
    ''Case "014" '新加坡
    '    StrSql = " And (CP10='101' Or CP10='107') "
    ''Case Else '其他
    ''    strSQL = " And CP10='101' "
    ''End Select
    Select Case strTM10
    Case "014" '新加坡
        strSql = " And (CP10='101' Or CP10='107') "
    Case Else '其他
        '2006/1/26 MODIFY BY SONIA
        'strSQL = " And CP10='101' "
        Select Case strTM01
           Case "CFC"
              strSql = " AND CP10='806' "
           Case Else
              strSql = " And CP10='101' "
        End Select
        '2006/1/26 END
    End Select
    '93.12.1 end
    'End
    'Modify By Cheng 2004/02/20
'   strSQL = "SELECT * FROM CaseProgress " & _
'            "WHERE CP01 = '" & strTM01 & "' AND " & _
'                  "CP02 = '" & strTM02 & "' AND " & _
'                  "CP03 = '" & strTM03 & "' AND " & _
'                  "CP04 = '" & strTM04 & "' AND " & _
'                  "CP10 = '101' "
   'modify by sonia 2023/8/18 加入已發文條件 CP27>0 (CFT-023826)
   strSql = "SELECT * FROM CaseProgress WHERE CP01 = '" & strTM01 & "' AND " & _
                  "CP02 = '" & strTM02 & "' AND CP03 = '" & strTM03 & "' AND " & _
                  "CP04 = '" & strTM04 & "' and CP27>0 " & strSql
    'End
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then
      rsTmp.Close
      strMsg = "資料庫中無符合的記錄"
      strTit = "檢核資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
'edit by nickc 2007/08/21 因為新加坡的跨類，可以收在同一案號裡面  CFT-011435
'   ElseIf rsTmp.RecordCount > 1 Then
'      rsTmp.Close
'      strMsg = "資料庫中的記錄不正確, 請聯絡電腦中心的人員"
'      strTit = "檢核資料"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      GoTo EXITSUB
   End If
         
   ' 顯示下一個畫面
   DisplayNextForm
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub DisplayNextForm()
   frm030102_02.SetData 0, Trim(textTM01), True
   frm030102_02.SetData 1, Trim(textTM02), False
   If IsEmptyText(textTM03) = True Then
      frm030102_02.SetData 2, "0", False
   Else
      frm030102_02.SetData 2, Trim(textTM03), False
   End If
   If IsEmptyText(textTM04) = True Then
      frm030102_02.SetData 3, "00", False
   Else
      frm030102_02.SetData 3, Trim(textTM04), False
   End If
   frm030102_02.SetData 4, textCP05, False
   
   'Add By Sindy 2023/4/24
   If Not m_PrevForm Is Nothing Then
      Call frm030102_02.SetParent(m_PrevForm)
   End If
   frm030102_02.m_strIR01 = m_strIR01
   frm030102_02.m_strIR02 = m_strIR02
   frm030102_02.m_strIR03 = m_strIR03
   frm030102_02.m_strIR04 = m_strIR04
   '2023/4/24 END
   
   Me.Hide
   frm030102_02.Show
   frm030102_02.QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2023/4/24
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2023/4/24 END
   
   'Add By Cheng 2002/07/19
   Set frm030102_01 = Nothing
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
      If Val(textCP05) > Val(strSrvDate(1)) Then 'edit by nickc 2006/03/17 原先為 ChangeWDateStringToWString(Date)) Then
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
         Case "CFT", "FCT", "CFC":
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
Private Sub InverseAll(ByRef tb As TextBox)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Private Sub textTM01_GotFocus()
   InverseAll textTM01
   CloseIme
End Sub

Private Sub textTM02_GotFocus()
   InverseAll textTM02
End Sub

Private Sub textTM03_GotFocus()
   InverseAll textTM03
   CloseIme
End Sub

Private Sub textTM03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM04_GotFocus()
   InverseAll textTM04
End Sub

Private Sub textCP05_GotFocus()
   InverseAll textCP05
End Sub

