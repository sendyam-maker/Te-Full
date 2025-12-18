VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm02010411_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "智慧局註冊費通知函輸入"
   ClientHeight    =   5688
   ClientLeft      =   156
   ClientTop       =   1008
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5688
   ScaleWidth      =   9324
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7212
      TabIndex        =   1
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6384
      TabIndex        =   0
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8436
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5580
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   660
      Width           =   2292
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   660
      Width           =   2292
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4692
      Left            =   96
      TabIndex        =   7
      Top             =   936
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   8276
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   4620
      TabIndex        =   6
      Top             =   660
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "審定號數 :"
      Height          =   252
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   852
   End
End
Attribute VB_Name = "frm02010411_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/29 Form2.0已修改 grdList
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/6 日期欄已修改
'2007/9/6 ADD BY SONIA
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 來函收文日
Dim m_CP05 As String
Dim m_NP07 As String
Dim m_NP06 As String
Dim m_NP09 As String

Private Sub cmdCancel_Click()
   Unload Me
   frm02010411_1.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm02010411_1
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   If grdList.Rows > 0 Then
      If IsEmptyText(m_TM01) = True Or IsEmptyText(m_TM02) = True Or IsEmptyText(m_TM03) = True Or IsEmptyText(m_TM04) = True Then
         strMsg = "請先選取一筆記錄"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   Else
      strMsg = "無符合的資料"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   m_NP06 = "": m_NP09 = ""

   ' 檢查所輸入的資料是否合乎資料庫的條件
   Select Case m_NP07
      '2012/12/19 MODIFY BY SONIA 715第一期註冊費改為717註冊費,1715改為1720
      'Case "715":  '第一期註冊費
      Case "717":  '註冊費
         If m_TM01 = "FCT" Then
            strMsg = "此申請案號為 FCT案件 !" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         '先檢查是否重覆輸入
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
            GoTo EXITSUB
         End If
         rsTmp.Close
         '第一期註冊費715期限資料,收文可能收全期註冊費717
         '2012/12/19 MODIFY BY SONIA 已無第一期註冊費715
         'strSql = "SELECT NP06,NP09,CP07 FROM NEXTPROGRESS,CASEPROGRESS " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 in ('" & m_NP07 & "','717') AND NP01=CP43(+) AND NP07=CP10 AND CP57 IS NULL UNION " & _
                  "SELECT NP06,NP09,CP07 FROM NEXTPROGRESS,CASEPROGRESS " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 in ('" & m_NP07 & "','717') AND NP01=CP43(+) AND '717'=CP10(+) AND CP57 IS NULL"
         strSql = "SELECT NP06,NP09,CP07,CP57 FROM NEXTPROGRESS,CASEPROGRESS " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = '" & m_NP07 & "' AND NP01=CP43(+) AND NP07=CP10(+) "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            rsTmp.Close
            strMsg = "此案件無註冊費期限的記錄"
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         Else
            '抓法定期限
            If rsTmp.Fields("NP06") = "Y" Then
               m_NP09 = rsTmp.Fields("CP07")
               '2013/1/28 MODIFY BY SONIA
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
      Case "716":  '第二期註冊費
         m_TM01 = rsTmp.Fields("TM01")
         m_TM02 = rsTmp.Fields("TM02")
         m_TM03 = rsTmp.Fields("TM03")
         m_TM04 = rsTmp.Fields("TM04")
         If m_TM01 = "FCT" Then
            rsTmp.Close
            strMsg = "此申請案號為 FCT案件 !" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
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
            GoTo EXITSUB
         End If
         rsTmp.Close
         '第二期註冊費資料
         '2013/1/28 MODIFY BY SONIA
         'strSql = "SELECT NP06,NP09,CP07 FROM NEXTPROGRESS,CASEPROGRESS " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = '" & m_NP07 & "' AND NP01=CP43(+) AND NP07=CP10(+) AND CP57 IS NULL"
         strSql = "SELECT NP06,NP09,CP07 FROM NEXTPROGRESS,CASEPROGRESS " & _
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
      Case "102":  '延展
         If m_TM01 = "FCT" Then
            rsTmp.Close
            strMsg = "此申請案號為 FCT案件 !" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
            strTit = "檢核資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
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
            GoTo EXITSUB
         End If
         rsTmp.Close
         '延展期限資料, 考慮NP有多次延展期限情形
         '2013/1/28 modify by sonia T-127226收文後取消收文
         'strSql = "SELECT NP06,NP09,CP07 FROM NEXTPROGRESS,CASEPROGRESS " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = '" & m_NP07 & "' AND NP01=CP43(+) AND NP07=CP10(+) AND CP57 IS NULL ORDER BY NP09 DESC"
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
            GoTo EXITSUB
         Else
            rsTmp.MoveFirst
            Do While rsTmp.EOF = False
               If rsTmp.Fields("np09") > ServerDate + 30000 Then
                  rsTmp.MoveNext
               Else
                  '抓法定期限
                  If rsTmp.Fields("NP06") = "Y" Then
                     m_NP09 = rsTmp.Fields("CP07")
                     '2013/1/28 modify by sonia T-127226收文後取消收文
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
                  Exit Do
               End If
            Loop
         End If
         rsTmp.Close
   End Select
   
   DisplayNextForm

EXITSUB:
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTM15.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   
   MoveFormToCenter Me
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      textTM12 = Empty
      textTM15 = Empty
      m_CP05 = Empty
      m_NP07 = Empty
   End If
   
   Select Case nType
      ' 申請案號
      Case 0: textTM12 = strData
      ' 審定號
      Case 1: textTM15 = strData
      ' 來函收文日
      Case 2: m_CP05 = strData
      ' 通知函性質
      Case 3: m_NP07 = strData
         
   End Select
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 11
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "本所案號"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "商標名稱"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "商標種類"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "商品類別"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "申請人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "申請國家"
   grdList.ColWidth(6) = 1200
   ' 本所案號 欄位一
   grdList.col = 7
   grdList.Text = Empty
   grdList.ColWidth(7) = 0
   ' 本所案號 欄位二
   grdList.col = 8
   grdList.Text = Empty
   grdList.ColWidth(8) = 0
   ' 本所案號 欄位三
   grdList.col = 9
   grdList.Text = Empty
   grdList.ColWidth(9) = 0
   ' 本所案號 欄位四
   grdList.col = 10
   grdList.Text = Empty
   grdList.ColWidth(10) = 0
End Sub

' 顯示下一個畫面
Private Sub DisplayNextForm()
   ' 顯示下一個畫面
   frm02010411_3.SetData 0, m_TM01, True
   frm02010411_3.SetData 1, m_TM02, False
   frm02010411_3.SetData 2, m_TM03, False
   frm02010411_3.SetData 3, m_TM04, False
   frm02010411_3.SetData 4, frm02010411_1.textCP05, False
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
End Sub

' 列出所有資料
Private Sub ListData(ByRef rsTmp As ADODB.Recordset)
   Dim strNationCode As String
   
   InitialGrdList
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      strNationCode = Empty
      grdList.Rows = grdList.Rows + 1
      grdList.row = grdList.Rows - 1
      ' 本所案號欄位
      grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("TM01") & rsTmp.Fields("TM02") & rsTmp.Fields("TM03") & rsTmp.Fields("TM04")
      ' 商標名稱欄位
      If IsNull(rsTmp.Fields("TM05")) = False Then
         grdList.TextMatrix(grdList.row, 2) = rsTmp.Fields("TM05")
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         strNationCode = rsTmp.Fields("TM10")
         grdList.TextMatrix(grdList.row, 6) = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      ' 商標種類欄位
      If IsNull(rsTmp.Fields("TM08")) = False Then
         If strNationCode < "010" Then
            grdList.TextMatrix(grdList.row, 3) = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
         Else
            grdList.TextMatrix(grdList.row, 3) = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
         End If
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         grdList.TextMatrix(grdList.row, 4) = rsTmp.Fields("TM09")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         grdList.TextMatrix(grdList.row, 5) = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 隱藏起來的本所案號
      grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("TM01")
      grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("TM02")
      grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("TM03")
      grdList.TextMatrix(grdList.row, 10) = rsTmp.Fields("TM04")
      
      rsTmp.MoveNext
   Loop
   
   'Added by Lydia 2023/10/18
   If grdList.Rows >= 2 Then
      grdList.FixedRows = 1
   End If
   'end 2023/10/18
   
   ' 設定第一列為選取的狀態
   grdList_SetSelection 1
      
   ' 當只有一筆記錄時, 則直接跳至下一畫面
   If grdList.Rows = 2 Then
      DisplayNextForm
   Else
      Me.Show
   End If

End Sub
' 搜尋資料
Public Sub QueryData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nType As Integer
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   
   ' 檢查所輸入的資料是否合乎資料庫的條件
   Select Case m_NP07
      '2012/12/19 MODIFY BY SONIA 第一期註冊費715改為717註冊費
      Case "717":  '註冊費
         '商標基本資料
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM12 = '" & textTM12 & "' AND " & _
                        "TM10 < '010' AND TM16 = '1' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            ListData rsTmp
         End If
         rsTmp.Close
      Case "716", "102": '第二期註冊費,延展
         '商標基本資料
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM15 = '" & textTM15 & "' AND " & _
                        "TM10 < '010' AND TM16 = '1' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            ListData rsTmp
         End If
         rsTmp.Close
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frm02010411_2.textTM12 = "": frm02010411_2.textTM15 = ""
   Set frm02010411_2 = Nothing
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      grdList.col = 7
      m_TM01 = grdList.Text
      grdList.col = 8
      m_TM02 = grdList.Text
      grdList.col = 9
      m_TM03 = grdList.Text
      grdList.col = 10
      m_TM04 = grdList.Text
   End If
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nRow As Integer
   Dim nCol As Integer
   Dim nCurrSel As Integer
   nCurrSel = grdList.row
   For nRow = 1 To grdList.Rows - 1
      grdList.row = nRow
      If nRow = nCurrSel Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            grdList.CellBackColor = &H8000000D
            grdList.CellForeColor = &H80000005
         Next nCol
      Else
         grdList.col = 1
         If grdList.CellBackColor <> &H80000005 Then
            For nCol = 1 To grdList.Cols - 1
               grdList.col = nCol
               If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
               If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            Next nCol
         End If
      End If
   Next nRow
   grdList.row = nCurrSel
   grdList.col = 0
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If grdList.Rows >= 2 Then
      grdList.row = nSel
      grdList_SelChange
      grdList_ShowSelection
   End If
End Sub
