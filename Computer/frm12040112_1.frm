VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040112_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "資料刪除記錄檔"
   ClientHeight    =   6012
   ClientLeft      =   888
   ClientTop       =   696
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6012
   ScaleWidth      =   7500
   Begin VB.TextBox textDD08 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   5112
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1632
      Width           =   2052
   End
   Begin VB.TextBox textDD07 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1632
      Width           =   1932
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4824
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   5652
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6480
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textDD02_2 
      Height          =   264
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   552
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textDD04 
      Height          =   264
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   4
      Top             =   552
      Width           =   732
   End
   Begin VB.TextBox textDD03 
      Height          =   264
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   3
      Top             =   552
      Width           =   372
   End
   Begin VB.TextBox textDD01 
      Height          =   264
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   0
      Top             =   552
      Width           =   732
   End
   Begin VB.TextBox textDD02 
      Height          =   264
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   1
      Top             =   552
      Width           =   1092
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3720
      Left            =   210
      TabIndex        =   17
      Top             =   2100
      Width           =   7050
      _ExtentX        =   12425
      _ExtentY        =   6562
      _Version        =   393216
      Cols            =   10
      AllowUserResizing=   3
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
      _Band(0).Cols   =   10
   End
   Begin MSForms.TextBox textDD06 
      Height          =   285
      Left            =   1200
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1230
      Width           =   5895
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "10398;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textDD05 
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   900
      Width           =   5895
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "10398;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   4008
      TabIndex        =   13
      Top             =   1632
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家 :"
      Height          =   252
      Left            =   168
      TabIndex        =   11
      Top             =   1632
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   168
      TabIndex        =   10
      Top             =   900
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   168
      TabIndex        =   9
      Top             =   1230
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Left            =   168
      TabIndex        =   8
      Top             =   552
      Width           =   996
   End
End
Attribute VB_Name = "frm12040112_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/15 改成Form2.0 ; textDD05、textDD06、grdList改字型=新細明體-ExtB(MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Dim m_KeySel As Integer
Dim m_CurrSel As Integer

Dim m_DD01 As String
Dim m_DD02 As String
Dim m_DD03 As String
Dim m_DD04 As String

' 90.07.16 modify by Ken (執行各項功能的權限)
Dim m_bDelete As Boolean
Public strDD07 As String   'modify by sonia 2021/10/28 改為Public，frm12040112_2才能使用

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(textDD01) = True Then
      strTit = "檢核資料"
      strMsg = "請先輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If IsEmptyText(textDD02) = True Then
      strTit = "檢核資料"
      strMsg = "請先輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If textDD01 = "TF" And IsEmptyText(textDD02_2) = True Then
      strTit = "檢核資料"
      strMsg = "請先輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   DisplayNextForm
EXITSUB:
End Sub

Private Sub Form_Load()
   
   textDD05.BackColor = &H8000000F
   textDD06.BackColor = &H8000000F
   textDD07.BackColor = &H8000000F
   textDD08.BackColor = &H8000000F
   
   MoveFormToCenter Me

   ' Ken 90/07/16 -- Start
   m_bDelete = IsUserHasRightOfFunction("frm12040112_1", strDel, False)
   If m_bDelete Then
       cmdOK.Enabled = True
   Else
       cmdOK.Enabled = False
   End If
   ' Ken 90/07/16 -- End
End Sub

Private Sub ClearField()
   textDD05 = Empty
   textDD06 = Empty
   textDD07 = Empty
   textDD08 = Empty
End Sub

Private Sub cmdQuery_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   m_DD01 = Empty
   m_DD02 = Empty
   m_DD03 = Empty
   m_DD04 = Empty
   
   If IsEmptyText(textDD01) = True Or IsEmptyText(textDD02) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If textDD01 = "TF" Then
      If IsEmptyText(textDD02_2) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入本所案號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   ' 查詢資料
   If QueryData() = False Then
      ClearField
      strTit = "查詢資料"
      strMsg = "無資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   
EXITSUB:
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm12040112_1 = Nothing
End Sub

Private Sub textDD01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號的系統別
Private Sub textDD01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textDD01) = False Then
      Select Case textDD01
         Case "TF":
            textDD02_2.Visible = True
            textDD02_2.Locked = False
            textDD02_2.TabStop = True
            textDD02.MaxLength = 5
         Case Else:
            textDD02_2.Visible = False
            textDD02_2.Locked = True
            textDD02_2.TabStop = False
            textDD02.MaxLength = 6
      End Select
   Else
      textDD02_2.Visible = False
      textDD02_2.Locked = True
      textDD02_2.TabStop = False
      textDD02.MaxLength = 6
   End If
End Sub

Private Function QueryData() As Boolean
   Dim strDD01 As String
   Dim strDD02 As String
   Dim strDD03 As String
   Dim strDD04 As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nRow As Integer
   Dim strTM10 As String
   
   QueryData = False
   
   ClearField
   InitialGridList
   
   ' 組成本所案號
   strDD01 = textDD01
   strDD02 = textDD02
   If strDD01 = "TF" Then: strDD02 = strDD02 & textDD02_2
   strDD03 = textDD03
   If IsEmptyText(strDD03) = True Then: strDD03 = "0"
   strDD04 = textDD04
   If IsEmptyText(strDD04) = True Then: strDD04 = "00"
   
   '2010/7/7 add by sonia 先抓基本資料,判斷應抓刪除記錄的或原基本檔的,T-157136抓TRADEMARK,CFT-015137抓DATADELETERECORD
   strSql = "SELECT DD05,NVL(NVL(CU04,CU05),CU06) AS DD06,DD07,DD08 FROM DATADELETERECORD,CUSTOMER " & _
            "WHERE DD01 = '" & strDD01 & "' AND DD02 = '" & strDD02 & "' AND " & _
                  "DD03 = '" & strDD03 & "' AND DD04 = '" & strDD04 & "' AND " & _
                  "SUBSTR(DD06,1,8) = CU01(+) AND " & _
                  "SUBSTR(DD06,9,1) = CU02(+) AND DD07 IS NOT NULL "
   strSql = strSql & "UNION SELECT TM05 DD05,NVL(NVL(CU04,CU05),CU06) AS DD06,TM10 DD07,TM12 DD08 FROM TRADEMARK,CUSTOMER " & _
            "WHERE TM01 = '" & strDD01 & "' AND TM02 = '" & strDD02 & "' AND " & _
                  "TM03 = '" & strDD03 & "' AND TM04 = '" & strDD04 & "' AND " & _
                  "SUBSTR(TM23,1,8) = CU01(+) AND " & _
                  "SUBSTR(TM23,9,1) = CU02(+) "
   strSql = strSql & "UNION SELECT NVL(NVL(PA05,PA06),PA07) DD05,NVL(NVL(CU04,CU05),CU06) AS DD06,PA09 DD07,PA11 DD08 FROM PATENT,CUSTOMER " & _
            "WHERE PA01 = '" & strDD01 & "' AND PA02 = '" & strDD02 & "' AND " & _
                  "PA03 = '" & strDD03 & "' AND PA04 = '" & strDD04 & "' AND " & _
                  "SUBSTR(PA26,1,8) = CU01(+) AND " & _
                  "SUBSTR(PA26,9,1) = CU02(+) "
   strSql = strSql & "UNION SELECT NVL(NVL(LC05,LC06),LC07) DD05,NVL(NVL(CU04,CU05),CU06) AS DD06,LC15 DD07,'' DD08 FROM LAWCASE,CUSTOMER " & _
            "WHERE LC01 = '" & strDD01 & "' AND LC02 = '" & strDD02 & "' AND " & _
                  "LC03 = '" & strDD03 & "' AND LC04 = '" & strDD04 & "' AND " & _
                  "SUBSTR(LC11,1,8) = CU01(+) AND " & _
                  "SUBSTR(LC11,9,1) = CU02(+) "
   strSql = strSql & "UNION SELECT NVL(NVL(SP05,SP06),SP07) DD05,NVL(NVL(CU04,CU05),CU06) AS DD06,SP09 DD07,SP11 DD08 FROM SERVICEPRACTICE,CUSTOMER " & _
            "WHERE SP01 = '" & strDD01 & "' AND SP02 = '" & strDD02 & "' AND " & _
                  "SP03 = '" & strDD03 & "' AND SP04 = '" & strDD04 & "' AND " & _
                  "SUBSTR(SP08,1,8) = CU01(+) AND " & _
                  "SUBSTR(SP08,9,1) = CU02(+) "
   strSql = strSql & "UNION SELECT HC06 DD05,NVL(NVL(CU04,CU05),CU06) AS DD06,'000' DD07,'' DD08 FROM HIRECASE,CUSTOMER " & _
            "WHERE HC01 = '" & strDD01 & "' AND HC02 = '" & strDD02 & "' AND " & _
                  "HC03 = '" & strDD03 & "' AND HC04 = '" & strDD04 & "' AND " & _
                  "SUBSTR(HC05,1,8) = CU01(+) AND " & _
                  "SUBSTR(HC05,9,1) = CU02(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ' 申請國家
      strDD07 = "000"
      If IsNull(rsTmp.Fields("DD07")) = False Then
         strDD07 = rsTmp.Fields("DD07")
         textDD07 = GetNationName(strDD07, 0)
      End If
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("DD05")) = False Then
         If IsEmptyText(textDD05) = True Then
            textDD05 = rsTmp.Fields("DD05")
         End If
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("DD06")) = False Then
         If IsEmptyText(textDD06) = True Then
            textDD06 = rsTmp.Fields("DD06")
         End If
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("DD08")) = False Then
         If IsEmptyText(textDD08) = True Then
            textDD08 = rsTmp.Fields("DD08")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   '2010/7/7 END
   
   ' 設定查詢資料庫的SQL語法
'2010/7/7 MODIFY BY SONIA
'   strSql = "SELECT DD05,C1.CU04 AS DD06,N1.NA03 AS DD07,DD08,DD14,NVL(DECODE(DECODE(DD07,'000',C2.CPM03,'',C2.CPM03,C2.CPM04),'（無）',C2.CPM04,C2.CPM03), DD15) AS DD15,SUBSTR(' '||sqldatet(DD16),-9) AS DD16,SUBSTR(' '||sqldatet(DD17),-9) AS DD17,SUBSTR(' '||sqldatet(DD18),-9) AS DD18,S1.ST02 AS DD19,S2.ST02 AS DD23,SUBSTR(' '||sqldatet(DD27),-9) AS DD27,DD28 FROM DATADELETERECORD, NATION N1, STAFF S1, STAFF S2, CUSTOMER C1, CASEPROPERTYMAP C2 " & _
'            "WHERE DD01 = '" & strDD01 & "' AND " & _
'                  "DD02 = '" & strDD02 & "' AND " & _
'                  "DD03 = '" & strDD03 & "' AND " & _
'                  "DD04 = '" & strDD04 & "' AND " & _
'                  "DD19 = S1.ST01(+) AND " & _
'                  "DD23 = S2.ST01(+) AND " & _
'                  "SUBSTR(DD06,1,8) = C1.CU01(+) AND " & _
'                  "SUBSTR(DD06,9,1) = C1.CU02(+) AND " & _
'                  "DD01 = C2.CPM01(+) AND " & _
'                  "DD15 = C2.CPM02(+) AND " & _
'                  "DD07 = N1.NA01(+) "
   '2010/9/14 MODIFY BY SONIA 日期欄改百年日期排序問題
   strSql = "SELECT DD14,DD15,SUBSTR(' '||sqldatet(DD16),-9) AS DD16,SUBSTR(' '||sqldatet(DD17),-9) AS DD17,SUBSTR(' '||sqldatet(DD18),-9) AS DD18,S1.ST02 AS DD19,S2.ST02 AS DD23,SUBSTR(' '||sqldatet(DD27),-9) AS DD27,DD28 FROM DATADELETERECORD, STAFF S1, STAFF S2 " & _
            "WHERE DD01 = '" & strDD01 & "' AND DD02 = '" & strDD02 & "' AND " & _
                  "DD03 = '" & strDD03 & "' AND DD04 = '" & strDD04 & "' AND " & _
                  "DD19 = S1.ST01(+) AND DD23 = S2.ST01(+) "
   strSql = strSql & " order by dd25,dd27 "    '2010/8/3 add by sonia
'2010/7/7 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryData = True
      ' 更新本所案號
      m_DD01 = strDD01
      m_DD02 = strDD02
      m_DD03 = strDD03
      m_DD04 = strDD04
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
'2010/7/7 cancel by sonia 移至上面
'         ' 申請國家
'         strDD07 = "000"
'         If IsNull(rsTmp.Fields("DD07")) = False Then
'            'strDD07 = rsTmp.Fields("DD07")
'            'textDD07 = GetNationName(rsTmp.Fields("DD07"), 0)
'            If IsEmptyText(textDD07) = True Then
'               textDD07 = rsTmp.Fields("DD07")
'            End If
'         End If
'         ' 案件中文名稱
'         If IsNull(rsTmp.Fields("DD05")) = False Then
'            If IsEmptyText(textDD05) = True Then
'               textDD05 = rsTmp.Fields("DD05")
'            End If
'         End If
'         ' 申請人
'         If IsNull(rsTmp.Fields("DD06")) = False Then
'            'textDD06 = GetCustomerName(rsTmp.Fields("DD06"), 0)
'            If IsEmptyText(textDD06) = True Then
'               textDD06 = rsTmp.Fields("DD06")
'            End If
'         End If
'         ' 申請案號
'         If IsNull(rsTmp.Fields("DD08")) = False Then
'            If IsEmptyText(textDD08) = True Then
'               textDD08 = rsTmp.Fields("DD08")
'            End If
'         End If
'2010/7/7 end
         grdList.Rows = grdList.Rows + 1
         nRow = grdList.Rows - 1
         ' 總收文號
         If IsNull(rsTmp.Fields("DD14")) = False Then
            grdList.TextMatrix(nRow, 1) = rsTmp.Fields("DD14")
         End If
         ' 案件性質
         If IsNull(rsTmp.Fields("DD15")) = False Then
            If strDD07 < "010" Then
               grdList.TextMatrix(nRow, 2) = GetCaseTypeName(strDD01, rsTmp.Fields("DD15"), 0)
            Else
               grdList.TextMatrix(nRow, 2) = GetCaseTypeName(strDD01, rsTmp.Fields("DD15"), 1)
            End If
            'grdList.TextMatrix(nRow, 2) = rsTmp.Fields("DD15")   '2010/7/7 CANCEL BY SONIA
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("DD16")) = False Then
            'If rsTmp.Fields("DD16") <> "0" Then
            '   grdList.TextMatrix(nRow, 3) = TAIWANDATE(rsTmp.Fields("DD16"))
            'End If
            grdList.TextMatrix(nRow, 3) = rsTmp.Fields("DD16")
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("DD17")) = False Then
            'If rsTmp.Fields("DD17") <> "0" Then
            '   grdList.TextMatrix(nRow, 4) = TAIWANDATE(rsTmp.Fields("DD17"))
            'End If
            grdList.TextMatrix(nRow, 4) = rsTmp.Fields("DD17")
         End If
         ' 收文日
         If IsNull(rsTmp.Fields("DD18")) = False Then
            'If rsTmp.Fields("DD18") <> "0" Then
            '   grdList.TextMatrix(nRow, 5) = TAIWANDATE(rsTmp.Fields("DD18"))
            'End If
            grdList.TextMatrix(nRow, 5) = rsTmp.Fields("DD18")
         End If
         ' 智權人員
         If IsNull(rsTmp.Fields("DD19")) = False Then
            'grdList.TextMatrix(nRow, 6) = GetStaffName(rsTmp.Fields("DD19"))
            grdList.TextMatrix(nRow, 6) = rsTmp.Fields("DD19")
         End If
         ' 失誤人員
         If IsNull(rsTmp.Fields("DD23")) = False Then
            'grdList.TextMatrix(nRow, 7) = GetStaffName(rsTmp.Fields("DD23"))
            grdList.TextMatrix(nRow, 7) = rsTmp.Fields("DD23")
         End If
         ' 刪除日期
         If IsNull(rsTmp.Fields("DD27")) = False Then
            'If rsTmp.Fields("DD27") <> "0" Then
            '   grdList.TextMatrix(nRow, 8) = TAIWANDATE(rsTmp.Fields("DD27"))
            'End If
            grdList.TextMatrix(nRow, 8) = rsTmp.Fields("DD27")
         End If
         ' 序號
         If IsNull(rsTmp.Fields("DD28")) = False Then
            grdList.TextMatrix(nRow, 9) = rsTmp.Fields("DD28")
         End If
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2022/01/11 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
      If grdList.Rows > 1 Then
         grdList.FixedRows = 1
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 初始化列表
Public Sub InitialGridList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 10

   grdList.ColWidth(0) = 300
   grdList.row = 0

   grdList.col = 0
   grdList.ColAlignment(0) = flexAlignCenterCenter
   grdList.col = 1
   grdList.Text = "總收文號"
   grdList.ColWidth(1) = 950
   grdList.ColAlignment(1) = flexAlignLeftCenter
   grdList.col = 2
   grdList.Text = "案件性質"
   grdList.ColWidth(2) = 1000
   grdList.ColAlignment(2) = flexAlignLeftCenter
   grdList.col = 3
   grdList.Text = "本所期限"
   grdList.ColWidth(3) = 840
   grdList.ColAlignment(3) = flexAlignCenterCenter
   grdList.col = 4
   grdList.Text = "法定期限"
   grdList.ColWidth(4) = 840
   grdList.ColAlignment(4) = flexAlignCenterCenter
   grdList.col = 5
   grdList.Text = "收文日"
   grdList.ColWidth(5) = 840
   grdList.ColAlignment(5) = flexAlignCenterCenter
   grdList.col = 6
   grdList.Text = "智權人員"
   grdList.ColWidth(6) = 800
   grdList.ColAlignment(6) = flexAlignLeftCenter
   grdList.col = 7
   grdList.Text = "失誤人員"
   grdList.ColWidth(7) = 800
   grdList.ColAlignment(7) = flexAlignLeftCenter
   grdList.col = 8
   grdList.Text = "刪除日期"
   grdList.ColWidth(8) = 850
   grdList.ColAlignment(8) = flexAlignCenterCenter
   grdList.col = 9
   grdList.Text = "序號"
   grdList.ColWidth(9) = 1200
   grdList.ColAlignment(9) = flexAlignLeftCenter
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
   ' 案件性質必須為延期的才可以選取
   If KeyCode = vbKeySpace Then
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
            grdList.Text = "V"
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub grdList_Click()
   ' 案件性質必須為延期的才可以選取
   If grdList.row > 0 Then
      grdList.col = 0
      If grdList.Text = "V" Then
         grdList.Text = Empty
      Else
         grdList.Text = "V"
      End If
   End If
End Sub

Private Sub grdList_SelChange()
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

Private Sub DisplayNextForm()
   Dim nRow As Integer
   Dim strDD01 As String
   Dim strDD02 As String
   Dim strDD03 As String
   Dim strDD04 As String
   Dim bFind As Boolean
   bFind = False
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 0) = "V" Then
         bFind = True
         Exit For
      End If
   Next nRow
   
   If bFind = True Then
      frm12040112_2.SetData 0, m_DD01, True
      frm12040112_2.SetData 1, m_DD02, False
      frm12040112_2.SetData 2, m_DD03, False
      frm12040112_2.SetData 3, m_DD04, False
   Else
      ' 組成本所案號
      strDD01 = textDD01
      strDD02 = textDD02
      If strDD01 = "TF" Then: strDD02 = strDD02 & textDD02_2
      strDD03 = textDD03
      If IsEmptyText(strDD03) = True Then: strDD03 = "0"
      strDD04 = textDD04
      If IsEmptyText(strDD04) = True Then: strDD04 = "00"
      frm12040112_2.SetData 0, strDD01, True
      frm12040112_2.SetData 1, strDD02, False
      frm12040112_2.SetData 2, strDD03, False
      frm12040112_2.SetData 3, strDD04, False
   End If
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 0) = "V" Then
         frm12040112_2.SetData 4, grdList.TextMatrix(nRow, 9), False
      End If
   Next nRow
   frm12040112_2.Show
   frm12040112_2.QueryDB
   Me.Hide
End Sub

' 刪除該筆記錄
Public Sub DelRecord(ByVal strData As String)
   Dim nRow As Integer
   Dim bFind As Boolean
   bFind = False
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 9) = strData Then
         bFind = True
         Exit For
      End If
   Next nRow
   
   If bFind = True Then
      If grdList.Rows <= 2 Then
         InitialGridList
      Else
         grdList.RemoveItem nRow
      End If
   End If
   
End Sub

' 更新該筆記錄
Public Sub ModRecord(ByVal strData As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nRow As Integer
   Dim bFind As Boolean
   Dim nCol As Integer
   bFind = False
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 9) = strData Then
         bFind = True
         Exit For
      End If
   Next nRow
   
   ' 設定查詢資料庫的SQL語法
   strSql = "SELECT * FROM DataDeleteRecord " & _
            "WHERE DD01 = '" & m_DD01 & "' AND " & _
                  "DD02 = '" & m_DD02 & "' AND " & _
                  "DD03 = '" & m_DD03 & "' AND " & _
                  "DD04 = '" & m_DD04 & "' AND " & _
                  "DD28 = '" & strData & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      
      ' 若找不到時新增一筆記錄
      If bFind = False Then
         grdList.Rows = grdList.Rows + 1
         nRow = grdList.Rows - 1
      End If
      ' 先清除成空白
      For nCol = 1 To grdList.Cols - 1
         grdList.TextMatrix(nRow, nCol) = Empty
      Next nCol
      
      ' 申請國家
      strDD07 = "000"
      If IsNull(rsTmp.Fields("DD07")) = False Then
         strDD07 = rsTmp.Fields("DD07")
      End If
      ' 設定成被選取的狀態
      grdList.TextMatrix(nRow, 0) = "V"
      ' 總收文號
      If IsNull(rsTmp.Fields("DD14")) = False Then
         grdList.TextMatrix(nRow, 1) = rsTmp.Fields("DD14")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("DD15")) = False Then
         If strDD07 < "010" Then
            grdList.TextMatrix(nRow, 2) = GetCaseTypeName(m_DD01, rsTmp.Fields("DD15"), 0)
         Else
            grdList.TextMatrix(nRow, 2) = GetCaseTypeName(m_DD01, rsTmp.Fields("DD15"), 1)
         End If
      End If
      ' 本所期限
      If IsNull(rsTmp.Fields("DD16")) = False Then
         If rsTmp.Fields("DD16") <> "0" Then
            grdList.TextMatrix(nRow, 3) = TAIWANDATE(rsTmp.Fields("DD16"))
         End If
      End If
      ' 法定期限
      If IsNull(rsTmp.Fields("DD17")) = False Then
         If rsTmp.Fields("DD17") <> "0" Then
            grdList.TextMatrix(nRow, 4) = TAIWANDATE(rsTmp.Fields("DD17"))
         End If
      End If
      ' 收文日
      If IsNull(rsTmp.Fields("DD18")) = False Then
         If rsTmp.Fields("DD18") <> "0" Then
            grdList.TextMatrix(nRow, 5) = TAIWANDATE(rsTmp.Fields("DD18"))
         End If
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("DD19")) = False Then
         grdList.TextMatrix(nRow, 6) = GetStaffName(rsTmp.Fields("DD19"))
      End If
      ' 失誤人員
      If IsNull(rsTmp.Fields("DD23")) = False Then
         grdList.TextMatrix(nRow, 7) = GetStaffName(rsTmp.Fields("DD23"))
      End If
      ' 刪除日期
      If IsNull(rsTmp.Fields("DD27")) = False Then
         If rsTmp.Fields("DD27") <> "0" Then
            grdList.TextMatrix(nRow, 8) = TAIWANDATE(rsTmp.Fields("DD27"))
         End If
      End If
      ' 序號
      If IsNull(rsTmp.Fields("DD28")) = False Then
         grdList.TextMatrix(nRow, 9) = rsTmp.Fields("DD28")
      End If
   End If
   Set rsTmp = Nothing
End Sub

Public Sub ClearRemark()
   Dim nRow As Integer
   For nRow = 0 To grdList.Rows - 1
      grdList.TextMatrix(nRow, 0) = Empty
   Next nRow
End Sub

Public Sub RefreshList()
   QueryData
End Sub

Private Sub textDD01_GotFocus()
   InverseTextBox textDD01
End Sub

Private Sub textDD02_GotFocus()
   InverseTextBox textDD02
End Sub

Private Sub textDD02_2_GotFocus()
   InverseTextBox textDD02_2
End Sub

Private Sub textDD03_GotFocus()
   InverseTextBox textDD03
End Sub

Private Sub textDD03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textDD04_GotFocus()
   InverseTextBox textDD04
End Sub


