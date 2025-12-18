VERSION 5.00
Begin VB.Form frm030401 
   BorderStyle     =   1  '單線固定
   Caption         =   "FC公告通知函"
   ClientHeight    =   2480
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2480
   ScaleWidth      =   5280
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   660
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   3915
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   240
         Width           =   2940
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   15
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.OptionButton radio 
      Caption         =   "本所案號："
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   1452
   End
   Begin VB.OptionButton radio 
      Caption         =   "公告日："
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   1452
   End
   Begin VB.TextBox textPrt 
      Height          =   264
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4260
      TabIndex        =   9
      Top             =   48
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3330
      TabIndex        =   8
      Top             =   48
      Width           =   912
   End
   Begin VB.TextBox textTM14 
      Height          =   264
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   1
      Top             =   960
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3600
      MaxLength       =   1
      TabIndex        =   4
      Top             =   960
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   5
      Top             =   960
      Width           =   732
   End
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   2
      Top             =   960
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "(1:通知函 2:翻譯函 3:全部)"
      Height          =   252
      Left            =   2640
      TabIndex        =   13
      Top             =   1320
      Width           =   2412
   End
   Begin VB.Label Label3 
      Caption         =   "列印別："
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   1332
   End
End
Attribute VB_Name = "frm030401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim m_KeySel As Integer
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 商標種類
Dim m_TM08 As String
' 審定號
Dim m_TM15 As String
' 正商標號數
Dim m_TM27 As String
'Add By Cheng 2003/02/17
Dim SeekPrint As Integer, SeekPrintL As Integer
'Add By Cheng 2003/02/20
Dim m_TM67 As String '放棄專用權
'Add By Cheng 2003/03/13
Dim m_blnPriDate As Boolean '判斷是否有優先權
'Add By Cheng 2003/12/16
Dim m_TM11 As String '申請日
Dim m_blnPrintAddress As Boolean '是否要列印地條

Private Sub Clear()
   textTM14 = Empty
   textTM01 = Empty
   textTM02 = Empty
   textTM02_2 = Empty
   textTM03 = Empty
   textTM04 = Empty
   textPrt = Empty
End Sub

Private Sub SetInputEntry()
   radio(0).Value = True
   radio(1).Value = False
   textTM14.SetFocus
   m_KeySel = 0
   UpdateCtrlState
End Sub

Private Sub cmdExit_Click()
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
    'Add By Cheng 2002/03/21
    Unload Me
End Sub

Private Sub cmdok_Click()
   If CheckDataValid() = True Then
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 清除查詢印表記錄檔欄位
      If Not OnProcess Then
         MsgBox "無符合條件之資料 !", vbInformation
      End If
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
'      Clear
'      SetInputEntry
   End If
End Sub

Private Sub Form_Load()
    
MoveFormToCenter Me
radio(0).Value = True

SeekPrintL = Printer.Orientation
PUB_SetPrinter Me.Name, Combo1, , False, SeekPrint     'Modified by Morgan 2017/11/21 設定印表機改呼叫公用函數,原程式移除

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
    Set frm030401 = Nothing
End Sub

Private Sub radio_Click(Index As Integer)
On Error Resume Next
   m_KeySel = Index
   UpdateCtrlState
   ' 設定游標停留的位置
   Select Case Index
      Case 0: textTM14.SetFocus
      Case 1: textTM01.SetFocus
   End Select
End Sub

' 更新控制項的顯示狀態
Private Sub UpdateCtrlState()
   Select Case m_KeySel
      Case 0:
         EnableTextBox textTM14, True
         EnableTextBox textTM01, False
         EnableTextBox textTM02, False
         EnableTextBox textTM03, False
         EnableTextBox textTM04, False
      Case 1:
         EnableTextBox textTM14, False
         EnableTextBox textTM01, True
         EnableTextBox textTM02, True
         EnableTextBox textTM03, True
         EnableTextBox textTM04, True
   End Select
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號的系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
         ' 檢查系統類別
      If IsCorrectSysKind(textTM01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "本所案號中的系統別不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
      ' 檢查使用者權限
      If IsUserHasRightOfSystem(strUserNum, textTM01) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "您沒有使用該系統類別的權限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01_GotFocus
         GoTo EXITSUB
      End If
      
      Select Case textTM01
         Case "FCT":
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
         'Case "TF":
         '   textTM02_2.Visible = True
         '   textTM02_2.Locked = False
         '   textTM02_2.TabStop = True
         '   textTM02.MaxLength = 5
         Case Else:
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
EXITSUB:
End Sub

' 公告日
Private Sub textTM14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM14) = False Then
      If CheckIsTaiwanDate(textTM14, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的公告日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14_GotFocus
      End If
   End If
End Sub

' 列印別
Private Sub textPrt_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPrt) = False Then
      Select Case textPrt
         Case "1", "2", "3":
         Case Else:
            Cancel = True
            strMsg = "請輸入正確的列印別"
            strTit = "資料檢核"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrt_GotFocus
      End Select
   End If
End Sub

Private Function OnProcess() As Boolean
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   
   OnProcess = True
   Select Case m_KeySel
      Case 0:
         pub_QL05 = pub_QL05 & ";" & radio(0).Caption & textTM14 'Add By Sindy 2010/10/22
         strSql = "SELECT * FROM TRADEMARK " & _
                  "WHERE TM14 = " & DBDATE(textTM14) & " AND " & _
                        "TM01 = 'FCT' AND " & _
                        "TM29 IS NULL "
      Case 1:
         ' 設定本所案號
         strTM01 = textTM01
         strTM02 = textTM02
         strTM03 = textTM03
         If IsEmptyText(strTM03) = True Then: strTM03 = "0"
         strTM04 = textTM04
         If IsEmptyText(strTM04) = True Then: strTM04 = "00"
         pub_QL05 = pub_QL05 & ";" & radio(1).Caption & strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04  'Add By Sindy 2010/10/22
            'Modify By Cheng 2003/01/10
'         strSQL = "SELECT * FROM TRADEMARK " & _
'                  "WHERE TM01 = '" & strTM01 & "' AND " & _
'                        "TM02 = '" & strTM02 & "' AND " & _
'                        "TM03 = '" & strTM03 & "' AND " & _
'                        "TM04 = '" & strTM04 & "' AND TM29 <> 'Y'"
         strSql = "SELECT * FROM TRADEMARK " & _
                  "WHERE TM01 = '" & strTM01 & "' AND " & _
                        "TM02 = '" & strTM02 & "' AND " & _
                        "TM03 = '" & strTM03 & "' AND " & _
                        "TM04 = '" & strTM04 & "' AND ( TM29 IS NULL OR TM29 <> 'Y' ) "
      Case Else:
   End Select
   If Len(textPrt) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label3 & textPrt & Label1 'Add By Sindy 2010/10/22
   End If
   
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/10/22
      OnProcess = True
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         m_TM01 = rsTmp.Fields("TM01")
         m_TM02 = rsTmp.Fields("TM02")
         m_TM03 = rsTmp.Fields("TM03")
         m_TM04 = rsTmp.Fields("TM04")
         m_TM08 = Empty
         m_TM15 = Empty
         m_TM27 = Empty
         If IsNull(rsTmp.Fields("TM08")) = False Then
            m_TM08 = rsTmp.Fields("TM08")
         End If
         If IsNull(rsTmp.Fields("TM15")) = False Then
            m_TM15 = rsTmp.Fields("TM15")
         End If
         If IsNull(rsTmp.Fields("TM27")) = False Then
            m_TM27 = rsTmp.Fields("TM27")
         End If
         'Add By Cheng 2003/02/20
         '取得放棄專用權
         m_TM67 = "" & rsTmp("TM67").Value
        'Add By Cheng 2003/12/16
         '取得申請日
         m_TM11 = "" & rsTmp("TM11").Value
        'End
         PrintLetter
        
         '新增地址條列表資料
         'Modify By Sindy 2025/10/2 取消地址條
'         If m_blnPrintAddress = True Then
'            pub_AddressListSN = pub_AddressListSN + 1
'            PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
'         End If
         rsTmp.MoveNext
      Loop
      MsgBox "列印結束 !", vbInformation
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/10/22
      OnProcess = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 檢查資料輸入是否正確
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   Select Case m_KeySel
      Case 0:
         ' 公告日不可空白
         If IsEmptyText(textTM14) = True Then
            strTit = "資料檢核"
            strMsg = "請輸入公告日"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         'Add By Cheng 2002/03/21
         If PUB_CheckKeyInDate(Me.textTM14) = -1 Then
            Me.textTM14.SetFocus
            textTM14_GotFocus
            GoTo EXITSUB
         End If
               
      Case 1:
         ' 本所案號
         If IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True Then
            strTit = "資料檢核"
            strMsg = "請輸入本所案號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         ' 本所案號
         If textTM01 = "TF" Then
            If IsEmptyText(textTM02_2) = True Then
               strTit = "資料檢核"
               strMsg = "本所案號輸入不完全"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               GoTo EXITSUB
            End If
         End If
      Case Else:
   End Select
   
   ' 列印別不可空白
   If IsEmptyText(textPrt) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入列印別"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub

Private Sub textTM14_GotFocus()
   InverseTextBox textTM14
End Sub

Private Sub textPrt_GotFocus()
   InverseTextBox textPrt
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
Dim strSql As String
Dim strTemp As String
Dim strKey As String
Dim rsTmp As ADODB.Recordset
'Add By Cheng 2003/03/13
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
    strKey = m_TM01 & m_TM02 & m_TM03 & m_TM04 & "&" & "101"
    'Add  By Cheng 2003/01/23
    '判斷是否有優先權資料
    StrSQLa = "Select Count(*) From PriDate Where PD01='" & m_TM01 & "' And PD02='" & m_TM02 & "' And PD03='" & m_TM03 & "' And PD04='" & m_TM04 & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.Fields(0).Value > 0 Then
        m_blnPriDate = True
    Else
        m_blnPriDate = False
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
   ' 定稿語文
   Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
      ' 中文
      Case "1":
         ' 清除定稿例外欄位檔原有資料
         EndLetter "03", strKey, "01", strUserNum
         If IsEmptyText(m_TM15) = False And IsEmptyText(m_TM08) = False Then
            'Modify By Cheng 2003/12/08
'            strSQL = "SELECT * FROM TMBULLETINMASTER " & _
'                     "WHERE TMBM01 = '" & m_TM15 & "' AND " & _
'                           "TMBM02 = '" & m_TM08 & "' "
            strSql = "SELECT * FROM TMBULLETIN " & _
                     "WHERE TMBM01 = '" & m_TM15 & "' AND " & _
                           "TMBM02 = '" & m_TM08 & "' "
            Set rsTmp = New ADODB.Recordset
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
            If rsTmp.RecordCount > 0 Then
               If IsNull(rsTmp.Fields("TMBM07")) = False Then
                  ' 卷
                  strTemp = Mid(rsTmp.Fields("TMBM07"), 1, 2)
                  If IsEmptyText(strTemp) = False Then
                     ' 卷數
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & strKey & "','" & "01" & "','" & strUserNum & _
                              "','卷數','" & strTemp & "')"
                     cnnConnection.Execute strSql
                  End If
                  '  期
                  strTemp = Mid(rsTmp.Fields("TMBM07"), 3, 3)
                  If IsEmptyText(strTemp) = False Then
                     ' 期數
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & strKey & "','" & "01" & "','" & strUserNum & _
                              "','期數','" & strTemp & "')"
                     cnnConnection.Execute strSql
                  End If
               End If
            End If
            rsTmp.Close
            Set rsTmp = Nothing
         End If
      ' 英文
      Case "2":
         Select Case textPrt
            'Modify By Cheng 2003/01/28
'            Case "1", "3"
            Case "1" '通知函
                '若申請日小於921128
                If Val(m_TM11) < 20031128 Then
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "03", strKey, "99", strUserNum
                '若申請日大於等於921128
                Else
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "03", strKey, "02", strUserNum
                End If
'            Case "2", "3":
            Case "2" '翻譯函
                '若申請日小於921128
                If Val(m_TM11) < 20031128 Then
                    '無動作
                '若申請日大於等於921128
                Else
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "03", strKey, IIf(m_blnPriDate = True, "03", "04"), strUserNum
                     ' 例外欄位--正商標號數
                     If m_TM08 = "2" Or m_TM08 = "5" Then
                         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "03" & "','" & strKey & "','" & IIf(m_blnPriDate = True, "03", "04") & "','" & strUserNum & _
                                  "','正商標號數','" & vbCrLf & ChgSQL("Its Principal Trademark No. : " & m_TM27) & "')"
                         cnnConnection.Execute strSql
                     End If
                     ' 例外欄位--放棄專用權
                     If m_TM67 <> "" Then
                         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "03" & "','" & strKey & "','" & IIf(m_blnPriDate = True, "03", "04") & "','" & strUserNum & _
                                  "','放棄專用權','" & vbCrLf & ChgSQL("The following part disclaimed : " & m_TM67) & "')"
                         cnnConnection.Execute strSql
                     End If
                End If
            Case "3" '通知函+翻譯函
                '若申請日小於921128
                If Val(m_TM11) < 20031128 Then
                       ' 清除定稿例外欄位檔原有資料
        '               EndLetter "03", strKey, "02", strUserNum
        '               EndLetter "03", strKey, "03", strUserNum
                        ' 清除定稿例外欄位檔原有資料
                        EndLetter "03", strKey, "99", strUserNum
'                       EndLetter "03", strKey, IIf(m_blnPriDate = True, "03", "04"), strUserNum
'                        ' 例外欄位--正商標號數
'                        If m_TM08 = "2" Or m_TM08 = "5" Then
'                            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                     "VALUES ('" & "03" & "','" & strKey & "','" & IIf(m_blnPriDate = True, "03", "04") & "','" & strUserNum & _
'                                     "','正商標號數','" & vbCrLf & ChgSQL("Its Principal Trademark No. : " & m_TM27) & "')"
'                            cnnConnection.Execute strSQL
'                        End If
'                        ' 例外欄位--放棄專用權
'                        If m_TM67 <> "" Then
'                            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                     "VALUES ('" & "03" & "','" & strKey & "','" & IIf(m_blnPriDate = True, "03", "04") & "','" & strUserNum & _
'                                     "','放棄專用權','" & vbCrLf & ChgSQL("The following part disclaimed : " & m_TM67) & "')"
'                            cnnConnection.Execute strSQL
'                        End If
                '若申請日大於等於921128
                Else
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "03", strKey, "02", strUserNum
                End If
         End Select
      ' 日文
      Case "3":
         Select Case textPrt
            Case "1", "3"
               ' 清除定稿例外欄位檔原有資料
               EndLetter "03", strKey, "04", strUserNum
            Case "2", "3":
               ' 清除定稿例外欄位檔原有資料
               EndLetter "03", strKey, "05", strUserNum
               ' 聯合商標
               If IsEmptyText(m_TM27) = False Then
                  ' 聯合商標
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & strKey & "','" & "05" & "','" & strUserNum & _
                           "','聯合商標','" & "依存 正商標 登錄番號 : (" & m_TM27 & ")" & "')"
                  cnnConnection.Execute strSql
               End If
               ' 商品區分
               If m_TM08 = "4" Then
                  ' 商品區分
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & strKey & "','" & "05" & "','" & strUserNum & _
                           "','商品區分','" & "服務區分" & "')"
                  cnnConnection.Execute strSql
               Else
                  ' 商品區分
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & strKey & "','" & "05" & "','" & strUserNum & _
                           "','商品區分','" & "商品區分" & "')"
                  cnnConnection.Execute strSql
               End If
               ' 指定商品
               If m_TM08 = "4" Then
                  ' 指定商品
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & strKey & "','" & "05" & "','" & strUserNum & _
                           "','指定商品','" & "指定役務" & "')"
                  cnnConnection.Execute strSql
               Else
                  ' 指定商品
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & strKey & "','" & "05" & "','" & strUserNum & _
                           "','指定商品','" & "指定商品" & "')"
                  cnnConnection.Execute strSql
               End If
            Case Else
         End Select
   End Select
End Sub

Private Sub PrintLetter()
   Dim ET01 As String, ET02 As String, ET03 As String, ET03_1 As String, stContent As String
   Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
   
   ET01 = "03"
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   ET02 = m_TM01 & m_TM02 & m_TM03 & m_TM04 & "&" & "101"
   ' 定稿語文
   Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
      ' 中文
      Case "1":
         ET03 = "01"
      ' 英文
      Case "2":
         Select Case textPrt
            Case "1"
                '若申請日小於921128
                If Val(m_TM11) < 20031128 Then
                    ET03 = "99"
                '若申請日大於等於921128
                Else
                    ET03 = "02"
                End If
            Case "2"
                If Val(m_TM11) >= 20031128 Then
                  If m_blnPriDate = True Then
                     ET03 = "03"
                  Else
                     ET03 = "04"
                  End If
                End If
            Case "3"
                '若申請日小於921128
                If Val(m_TM11) < 20031128 Then
                     ET03 = "99"
                '若申請日大於等於921128
                Else
                     ET03 = "02"
                     If m_blnPriDate = True Then
                        ET03_1 = "03"
                     Else
                        ET03_1 = "04"
                     End If
                End If
         End Select
      ' 日文
      Case "3":
         Select Case textPrt
            Case "1", "3"
               ET03 = "04"
            Case "2", "3":
               ET03 = "05"
            Case Else
         End Select
   End Select
   If ET03 <> "" Then
      'Add by Morgan 2008/6/13
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      If bolEmail Then
         'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'end 2009/10/20
         If ET03_1 <> "" Then
            NowPrint ET02, ET01, ET03, False, strUserNum, , , , , iCopy
            NowPrint ET02, ET01, ET03_1, False, strUserNum, , , , , iCopy
            NowPrint ET02, ET01, ET03, False, strUserNum, , , True, stContent, , , , True
            NowPrint ET02, ET01, ET03_1, False, strUserNum, , stContent, , , , , True, True
         Else
            NowPrint ET02, ET01, ET03, False, strUserNum, , , , , iCopy, , True, True
         End If
         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
      Else
      'end 2008/6/13
         NowPrint ET02, ET01, ET03, False, strUserNum
         If ET03_1 <> "" Then
            NowPrint ET02, ET01, ET03_1, False, strUserNum
         End If
      End If
      
      If Not bolEmail Or bolPlusPaper Then
         m_blnPrintAddress = True
      Else
         m_blnPrintAddress = False
      End If
   End If
End Sub

