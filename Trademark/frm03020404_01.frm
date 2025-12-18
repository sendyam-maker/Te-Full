VERSION 5.00
Begin VB.Form frm03020404_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標發註冊證輸入"
   ClientHeight    =   2976
   ClientLeft      =   -96
   ClientTop       =   3456
   ClientWidth     =   4896
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2976
   ScaleWidth      =   4896
   Begin VB.TextBox textTM12 
      Height          =   264
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   3
      Top             =   750
      Width           =   2892
   End
   Begin VB.OptionButton radio 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   750
      Value           =   -1  'True
      Width           =   1452
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   660
      Left            =   210
      TabIndex        =   15
      Top             =   2160
      Width           =   3915
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   240
         Width           =   2940
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   16
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1080
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3600
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1080
      Width           =   372
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1080
      Width           =   732
   End
   Begin VB.TextBox textTM15 
      Height          =   264
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1410
      Width           =   2892
   End
   Begin VB.TextBox textCP05 
      Height          =   264
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   10
      Top             =   1770
      Width           =   2892
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2940
      TabIndex        =   12
      Top             =   72
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3912
      TabIndex        =   13
      Top             =   72
      Width           =   912
   End
   Begin VB.OptionButton radio 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1452
   End
   Begin VB.OptionButton radio 
      Caption         =   "審定號數 :"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1410
      Width           =   1452
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1770
      Width           =   1455
   End
End
Attribute VB_Name = "frm03020404_01"
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

'Add By Sindy 2019/7/22 暫存公告日
Public m_TM14 As String

Dim m_KeySel As Integer
'Add By Cheng 2002/05/08
Public m_blnTM16Is1 As Boolean '商標基本檔的是否准駁欄是否為准
'Add By Cheng 2003/02/17
Dim SeekPrint As Integer, SeekPrintL As Integer
'Added by Morgan 2023/1/17 電子公文
Public m_RDate As String
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_RegNo As String
Public m_DeadLine As String
Public m_NewCP10 As String
Dim m_Done As Boolean
'end 2023/1/17

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Modify By Cheng 2003/06/26
    '取消列印接洽結案單
'    'Add By Cheng 2003/06/23
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
    'Add By Cheng 2003/02/17
    '列印地址條
'move to unload by nick 2004/10/22
'    PUB_PrintAddressList strUserNum, Me.Combo1.Text
'    '刪除地址條列表資料
'    PUB_DeleteAddressList strUserNum
'    '初始化序號
'    pub_AddressListSN = 0
'    '若印表機變動, 則更新列印設定
'    If Me.Combo1.Text <> Me.Combo1.Tag Then
'        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
'    End If
    Unload Me
End Sub

Private Sub Form_Activate()
   'Added by Morgan 2023/1/17 電子公文
   If m_AppNo & m_RegNo <> "" And m_Done = False Then
      If m_AppNo <> "" Then
         radio(2).Value = True
         textTM12.Text = m_AppNo
         m_KeySel = 2
      Else
         radio(0).Value = True
         textTM15.Text = m_RegNo
         m_KeySel = 0
      End If
      textCP05.Text = m_RDate
      cmdOK.Value = True
      m_Done = True
   End If
   'end 2023/1/17
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'textCP05 = TAIWANDATE(SystemDate())
   textCP05 = strSrvDate(2)
   
   'Add By Sindy 2019/7/22 暫存公告日
   'm_TM14 = "" 'Removed by Morgan 2023/6/15 取消,否則電子公文設定後會被清除

   Initial
   UpdateCtrlState
   
   SeekPrintL = Printer.Orientation
   PUB_SetPrinter Me.Name, Combo1, , False, SeekPrint     'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除
   
   SendKeys "{Tab}"
End Sub

Private Sub Initial()
   ' 預設由申請案號來取得資料
   m_KeySel = 2
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   ' 檢查輸入的欄位
   Select Case m_KeySel
      Case 0:
         If IsEmptyText(textTM15) = True Then
            strMsg = "請輸入審定號數"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
      Case 1:
         If IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True Then
            strMsg = "請輸入本所案號"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
      Case 2:
         If IsEmptyText(textTM12) = True Then
            strMsg = "請輸入申請案號"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
   End Select
   ' 檢查來函收文日
   If IsEmptyText(textCP05) = True Then
      strMsg = "請輸入來函收文日"
      strTit = "檢核資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   Else
      If CheckIsTaiwanDate(textCP05, False) = False Then
         strMsg = "請輸入正確的來函收文日"
         strTit = "檢核資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      'edit by nickc 2006/03/17
      'If Val(textCP05) > Val(TAIWANDATE(Date)) Then
      If Val(textCP05) > Val(strSrvDate(2)) Then
         strMsg = "來函收文日不可超過系統日"
         strTit = "檢核資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   'Add By Cheng 2002/05/08
   m_blnTM16Is1 = True
   
   ' 檢查欄位的資料是否都已經輸入正確
   If CheckDataValid = False Then
      GoTo EXITSUB
   End If
   ' 檢查所輸入的資料是否合乎資料庫的條件
   Select Case m_KeySel
      Case 0:
         ' 設定SQL語法
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM15 = '" & textTM15 & "' AND " & _
                        "TM10 < '010' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無符合的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         'Add By Cheng 2002/05/08
         If rsTmp("TM16") = "1" Then
            m_blnTM16Is1 = True
         Else
            m_blnTM16Is1 = False
         End If
         rsTmp.Close
      Case 1:
         strTM01 = textTM01
         strTM02 = textTM02
         strTM03 = textTM03
         If IsEmptyText(strTM03) = True Then: strTM03 = "0"
         strTM04 = textTM04
         If IsEmptyText(strTM04) = True Then: strTM04 = "00"
         
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "TM02 = '" & strTM02 & "' AND " & _
                        "TM03 = '" & strTM03 & "' AND " & _
                        "TM04 = '" & strTM04 & "' "
         ' 若申請國家為台灣, 則須設定申請案號或審定號
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無符合的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         'Add By Cheng 2002/05/08
         If rsTmp("TM16") = "1" Then
            m_blnTM16Is1 = True
         Else
            m_blnTM16Is1 = False
         End If
         rsTmp.Close
      Case 2:
         ' 設定SQL語法
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM12 = '" & textTM12 & "' AND " & _
                        "TM10 < '010' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無符合的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         'Add By Cheng 2002/05/08
         If rsTmp("TM16") = "1" Then
            m_blnTM16Is1 = True
         Else
            m_blnTM16Is1 = False
         End If
         rsTmp.Close
   End Select
   ' 顯示下一個畫面
   DisplayNextForm
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub DisplayNextForm()
   Select Case m_KeySel
      Case 0:
         frm03020404_02.SetData 4, textTM15, True
         frm03020404_02.SetData 5, textCP05, False
         'Add By Sindy 2019/7/22 暫存公告日
         frm03020404_02.SetData 7, m_TM14, False
      Case 1:
         frm03020404_02.SetData 0, Trim(textTM01), True
         frm03020404_02.SetData 1, Trim(textTM02), False
         frm03020404_02.SetData 2, Trim(textTM03), False
         frm03020404_02.SetData 3, Trim(textTM04), False
         frm03020404_02.SetData 5, textCP05, False
         'Add By Sindy 2019/7/22 暫存公告日
         frm03020404_02.SetData 7, m_TM14, False
      Case 2:
         frm03020404_02.SetData 6, textTM12, True
         frm03020404_02.SetData 5, textCP05, False
         'Add By Sindy 2019/7/22 暫存公告日
         frm03020404_02.SetData 7, m_TM14, False
   End Select
   Me.Hide
   frm03020404_02.Show
   frm03020404_02.QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PUB_PrintAddressList strUserNum, Me.Combo1.Text
    '刪除地址條列表資料
    PUB_DeleteAddressList strUserNum
    '初始化序號
    pub_AddressListSN = 0
    '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    
    'Add By Cheng 2003/02/17
    '還原預設印表機
    Set Printer = Printers(SeekPrint)
    Printer.Orientation = SeekPrintL
    'Add By Cheng 2002/07/19
    Set frm03020404_01 = Nothing
End Sub

Private Sub radio_Click(Index As Integer)
   m_KeySel = Index
   UpdateCtrlState
   ' 設定游標停留的位置
   Select Case Index
      Case 0: textTM15.SetFocus
      Case 1: textTM01.SetFocus
      Case 2: textTM12.SetFocus
   End Select
End Sub

Private Sub UpdateCtrlState()
   Select Case m_KeySel
      Case 0:
         EnableTextBox textTM15, True
         EnableTextBox textTM01, False
         EnableTextBox textTM02, False
         EnableTextBox textTM03, False
         EnableTextBox textTM04, False
         EnableTextBox textTM12, False
         textTM02_2.Visible = False
      Case 1:
         EnableTextBox textTM15, False
         EnableTextBox textTM01, True
         EnableTextBox textTM02, True
         EnableTextBox textTM03, True
         EnableTextBox textTM04, True
         EnableTextBox textTM12, False
         textTM01_Validate False
      Case 2:
         EnableTextBox textTM15, False
         EnableTextBox textTM01, False
         EnableTextBox textTM02, False
         EnableTextBox textTM03, False
         EnableTextBox textTM04, False
         EnableTextBox textTM12, True
         textTM02_2.Visible = False
   End Select
End Sub

' 來函收文日
Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If CheckIsTaiwanDate(textCP05, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的來函收文日"
         strTit = "來函收文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
         GoTo EXITSUB
      End If
      'edit by nickc 2006/03/17
      'If Val(textCP05) > Val(TAIWANDATE(Date)) Then
      If Val(textCP05) > Val(strSrvDate(2)) Then
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

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      Select Case textTM01
         Case "FCT":
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

Private Sub textTM12_GotFocus()
    TextInverse Me.textTM12
End Sub

Private Sub textTM15_GotFocus()
   InverseTextBox textTM15
End Sub

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

