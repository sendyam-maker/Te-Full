VERSION 5.00
Begin VB.Form frm02010411_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "智慧局註冊費通知函輸入"
   ClientHeight    =   2805
   ClientLeft      =   435
   ClientTop       =   2490
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5115
   Begin VB.OptionButton radio 
      Caption         =   "延展"
      Height          =   252
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   1000
      Width           =   1452
   End
   Begin VB.OptionButton radio 
      Caption         =   "註冊費"
      Height          =   252
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.OptionButton radio 
      Caption         =   "第二期註冊費(已隱藏)"
      Height          =   252
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   2500
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3984
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3132
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textCP05 
      Height          =   264
      Left            =   1860
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2280
      Width           =   2892
   End
   Begin VB.TextBox textTM15 
      Height          =   264
      Left            =   1860
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1980
      Width           =   2892
   End
   Begin VB.TextBox textTM12 
      Height          =   264
      Left            =   1860
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1680
      Width           =   2892
   End
   Begin VB.Label Label4 
      Caption         =   "審定號數："
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號："
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   1725
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "通知函性質： "
      Height          =   255
      Left            =   580
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日："
      Height          =   255
      Left            =   650
      TabIndex        =   8
      Top             =   2355
      Width           =   1095
   End
End
Attribute VB_Name = "frm02010411_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/29 Form2.0已修改 (無需修改)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/6 日期欄已修改
'2007/9/6 ADD BY SONIA
Option Explicit

Dim m_KeySel As Integer
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_NP06 As String
Dim m_NP07 As String
Dim m_NP09 As String
'Added by Morgan 2017/4/21 電子公文
Public m_RDate As String
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_RegNo As String
Public m_DeadLine As String
Public m_NewCP10 As String
Dim m_Done As Boolean
'end 2017/4/21

Private Sub cmdExit_Click()
   Me.Enabled = False
   Unload Me
End Sub

Private Sub Form_Activate()
   'Added by Morgan 2017/4/24 電子公文
   If m_AppNo & m_RegNo <> "" And m_Done = False Then
      textTM12.Text = m_AppNo
      textTM15.Text = m_RegNo
      textCP05.Text = m_RDate
      m_Done = True
   End If
   'end 2017/4/24
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   textCP05 = strSrvDate(2)
   UpdateCtrlState
End Sub

Private Sub Initial()
   ' 預設由申請案號來取得資料
   m_KeySel = 0
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   CheckDataValid = False
   ' 檢查輸入的欄位
   Select Case m_KeySel
      Case 0:
         If IsEmptyText(textTM12) = True Then
            strMsg = "請輸入申請案號"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
      Case 1, 2:
         If IsEmptyText(textTM15) = True Then
            strMsg = "請輸入審定號數"
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
      If Val(textCP05) > Val(strSrvDate(1)) Then
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
   
   ' 檢查欄位的資料是否都已經輸入正確
   If CheckDataValid = False Then
      GoTo EXITSUB
   End If
   
   m_TM01 = "": m_TM02 = "": m_TM03 = "": m_TM04 = "": m_NP06 = "": m_NP09 = ""

   ' 檢查所輸入的資料是否合乎資料庫的條件
   Select Case m_KeySel
      Case 0:  '第一期註冊費
         '商標基本資料
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM12 = '" & textTM12 & "' AND " & _
                        "TM10 < '010' AND TM16 = '1' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無此申請案號的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM12_GotFocus
            GoTo EXITSUB
         ElseIf rsTmp.RecordCount > 1 Then    '申請案號重覆
            rsTmp.Close
            MultiData
            GoTo EXITSUB
         End If
         m_TM01 = rsTmp.Fields("TM01")
         m_TM02 = rsTmp.Fields("TM02")
         m_TM03 = rsTmp.Fields("TM03")
         m_TM04 = rsTmp.Fields("TM04")
         If m_TM01 = "FCT" Then
            rsTmp.Close
            strMsg = "此申請案號為 FCT案件 !" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM12_GotFocus
            GoTo EXITSUB
         End If
         rsTmp.Close
         '先檢查是否重覆輸入
         '2012/12/19 MODIFY BY SONIA 通知繳納第一期註冊費1715改為1720通知繳納註冊費
         strSql = "SELECT * FROM CASEPROGRESS " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "' AND " & _
                        "CP10 = '1720' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.Close
            strMsg = "此案件已輸過通知註冊費來函, 不可重覆輸入 !"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM12_GotFocus
            GoTo EXITSUB
         End If
         rsTmp.Close
         '第一期註冊費715期限資料,收文可能收全期註冊費717
         '2012/12/19 MODIFY BY SONIA 已無第一期註冊費715
         'strSql = "SELECT NP06,NP09,CP07,CP57 FROM NEXTPROGRESS,CASEPROGRESS " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 in ('" & m_NP07 & "','717') AND NP01=CP43(+) AND NP07=CP10 UNION " & _
                  "SELECT NP06,NP09,CP07,CP57 FROM NEXTPROGRESS,CASEPROGRESS " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 in ('" & m_NP07 & "','717') AND NP01=CP43(+) AND '717'=CP10(+) "
         strSql = "SELECT NP06,NP09,CP07,CP57 FROM NEXTPROGRESS,CASEPROGRESS " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = '" & m_NP07 & "' AND NP01=CP43(+) AND NP07=CP10(+) "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            '2012/12/19 MODIFY BY SONIA 已無第一期註冊費715
            'strMsg = "此案件無第一期註冊費期限的記錄"
            strMsg = "此案件無註冊費期限的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM12_GotFocus
            GoTo EXITSUB
         Else
            '抓法定期限
            If rsTmp.Fields("NP06") = "Y" Then
               m_NP09 = rsTmp.Fields("CP07")
               '2013/1/28 modify by sonia
               'm_NP06 = "Y"
               If IsNull(rsTmp.Fields("CP57")) Then
                  m_NP06 = "Y"
               Else
                  m_NP06 = ""
               End If
               '2013/1/28 END
            Else
               m_NP09 = rsTmp.Fields("NP09")
               m_NP06 = ""
            End If
         End If
         rsTmp.Close
      Case 1:  '第二期註冊費
         '商標基本資料
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM15 = '" & textTM15 & "' AND " & _
                        "TM10 < '010' AND TM16 = '1' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無此審定號數的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM15_GotFocus
            GoTo EXITSUB
         ElseIf rsTmp.RecordCount > 1 Then    '審定號重覆
            rsTmp.Close
            MultiData
            GoTo EXITSUB
         End If
         m_TM01 = rsTmp.Fields("TM01")
         m_TM02 = rsTmp.Fields("TM02")
         m_TM03 = rsTmp.Fields("TM03")
         m_TM04 = rsTmp.Fields("TM04")
         If m_TM01 = "FCT" Then
            rsTmp.Close
            strMsg = "此申請案號為 FCT案件 !" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM15_GotFocus
            GoTo EXITSUB
         End If
         rsTmp.Close
         '先檢查是否重覆輸入
         strSql = "SELECT * FROM CASEPROGRESS " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "' AND " & _
                        "CP10 = '1716' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.Close
            strMsg = "此案件已輸過通知第二期註冊費來函, 不可重覆輸入 !"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM12_GotFocus
            GoTo EXITSUB
         End If
         rsTmp.Close
         '第二期註冊費資料
         strSql = "SELECT NP06,NP09,CP07,CP57 FROM NEXTPROGRESS,CASEPROGRESS " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = '" & m_NP07 & "' AND NP01=CP43(+) AND NP07=CP10(+)"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "此案件無第二期註冊費期限的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM15_GotFocus
            GoTo EXITSUB
         Else
            '抓法定期限
            If rsTmp.Fields("NP06") = "Y" Then
               m_NP09 = rsTmp.Fields("CP07")
               '2013/1/28 modify by sonia
               'm_NP06 = "Y"
               If IsNull(rsTmp.Fields("CP57")) Then
                  m_NP06 = "Y"
               Else
                  m_NP06 = ""
               End If
               '2013/1/28 END
            Else
               m_NP09 = rsTmp.Fields("NP09")
               m_NP06 = ""
            End If
         End If
         rsTmp.Close
      Case 2:  '延展
         '商標基本資料
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM15 = '" & textTM15 & "' AND " & _
                        "TM10 < '010' AND TM16 = '1' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "資料庫中無此審定號數的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM15_GotFocus
            GoTo EXITSUB
         ElseIf rsTmp.RecordCount > 1 Then    '審定號重覆
            rsTmp.Close
            MultiData
            GoTo EXITSUB
         End If
         m_TM01 = rsTmp.Fields("TM01")
         m_TM02 = rsTmp.Fields("TM02")
         m_TM03 = rsTmp.Fields("TM03")
         m_TM04 = rsTmp.Fields("TM04")
         If m_TM01 = "FCT" Then
            rsTmp.Close
            strMsg = "此申請案號為 FCT案件 !" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM15_GotFocus
            GoTo EXITSUB
         End If
         rsTmp.Close
         '先檢查是否重覆輸入
         strSql = "SELECT * FROM CASEPROGRESS " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "' AND " & _
                        "CP10 = '1717' AND CP05>= " & strSrvDate(1) - 10000
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.Close
            strMsg = "此案件已輸過通知延展來函, 不可重覆輸入 !"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM12_GotFocus
            GoTo EXITSUB
         End If
         rsTmp.Close
         '延展期限資料, 考慮NP有多次延展期限情形
         strSql = "SELECT NP06,NP09,CP07,CP57 FROM NEXTPROGRESS,CASEPROGRESS " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = '" & m_NP07 & "' AND NP01=CP43(+) AND NP07=CP10(+) ORDER BY NP09 DESC"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "此案件無延展期限的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM15_GotFocus
            GoTo EXITSUB
         Else
            rsTmp.MoveFirst
            Do While rsTmp.EOF = False
               If rsTmp.Fields("np09") > ServerDate + 30000 Then
                  rsTmp.MoveNext
               Else
                  '抓法定期限
                  If rsTmp.Fields("NP06") = "Y" Then
                     m_NP09 = "" & rsTmp.Fields("CP07")
                     If IsNull(rsTmp.Fields("CP57")) Then
                        m_NP06 = "Y"
                     Else
                        m_NP06 = ""
                     End If
                  Else
                     m_NP09 = rsTmp.Fields("NP09")
                     m_NP06 = ""
                  End If
                  Exit Do
               End If
            Loop
         End If
         rsTmp.Close
   End Select
   ' 顯示下一個畫面
   frm02010411_3.SetData 0, m_TM01, True
   frm02010411_3.SetData 1, m_TM02, False
   frm02010411_3.SetData 2, m_TM03, False
   frm02010411_3.SetData 3, m_TM04, False
   frm02010411_3.SetData 4, textCP05, False
   frm02010411_3.SetData 5, m_NP06, False
   ' 通知函性質
   Select Case m_NP07
      '2012/12/19 MODIFY BY SONIA 第一期註冊費715改為717註冊費,通知繳納第一期註冊費1715改為1720通知繳納註冊費
      Case "717"
         frm02010411_3.SetData 6, "1720", False
      Case "716"
         frm02010411_3.SetData 6, "1716", False
      Case "102"
         frm02010411_3.SetData 6, "1717", False
   End Select
   frm02010411_3.SetData 7, m_NP09, False
   'Added by Morgan 2017/4/24 電子公文
   frm02010411_3.m_DocWord = frm02010411_1.m_DocWord
   frm02010411_3.m_DocNo = frm02010411_1.m_DocNo
   frm02010411_3.m_AppNo = frm02010411_1.m_AppNo
   frm02010411_3.m_DeadLine = frm02010411_1.m_DeadLine
   'end 2017/4/24
   Me.Hide
   frm02010411_3.Show
   frm02010411_3.QueryData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm02010411_1 = Nothing
End Sub

Private Sub radio_Click(Index As Integer)
   m_KeySel = Index
   UpdateCtrlState
   ' 設定游標停留的位置
   Select Case Index
      Case 0:
         textTM12.SetFocus
         '2012/12/19 MODIFY BY SONIA 第一期註冊費715改為717註冊費
         m_NP07 = "717"
      Case 1:
         textTM15.SetFocus
         m_NP07 = "716"
      Case 2:
         textTM15.SetFocus
         m_NP07 = "102"
   End Select
End Sub

Private Sub UpdateCtrlState()
   Select Case m_KeySel
      Case 0:
         EnableTextBox textTM12, True
         EnableTextBox textTM15, False
      Case 1, 2:
         EnableTextBox textTM12, False
         EnableTextBox textTM15, True
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
Private Sub textTM12_GotFocus()
   InverseTextBox textTM12
End Sub

Private Sub textTM15_GotFocus()
   InverseTextBox textTM15
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub

Private Function MultiData() As Boolean
   ' 顯示下一個畫面
   frm02010411_2.SetData 0, textTM12, True
   frm02010411_2.SetData 1, textTM15, False
   frm02010411_2.SetData 2, textCP05, False
   frm02010411_2.SetData 3, m_NP07, False
   Me.Hide
   frm02010411_2.Show
   frm02010411_2.QueryData
End Function
