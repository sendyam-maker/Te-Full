VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm030201_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "分案"
   ClientHeight    =   5760
   ClientLeft      =   5172
   ClientTop       =   1656
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9348
   Begin VB.OptionButton optSelect 
      Caption         =   "電子收文未分案 :"
      Height          =   252
      Index           =   3
      Left            =   2100
      TabIndex        =   9
      Top             =   1320
      Width           =   1752
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   3915
      Left            =   120
      TabIndex        =   28
      Top             =   1770
      Width           =   9135
      _ExtentX        =   16108
      _ExtentY        =   6900
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
   Begin VB.TextBox TextSys 
      Height          =   270
      Left            =   6030
      MaxLength       =   3
      TabIndex        =   14
      Top             =   1440
      Width           =   510
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "取消數量(&C)"
      Height          =   400
      Index           =   1
      Left            =   1380
      TabIndex        =   26
      Top             =   60
      Width           =   1152
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   6030
      MaxLength       =   1
      TabIndex        =   12
      Top             =   1170
      Width           =   300
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   13
      Top             =   1170
      Width           =   300
   End
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.OptionButton optSelect 
      Caption         =   "收文日 :"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   21
      Top             =   720
      Value           =   -1  'True
      Width           =   1332
   End
   Begin VB.OptionButton optSelect 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1020
      Width           =   1332
   End
   Begin VB.OptionButton optSelect 
      Caption         =   "以前未分案 :"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1332
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8232
      TabIndex        =   20
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7295
      TabIndex        =   19
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1020
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1020
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1020
      Width           =   732
   End
   Begin VB.TextBox textCP05_1 
      Height          =   264
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   0
      Top             =   720
      Width           =   1212
   End
   Begin VB.TextBox textCP05_2 
      Height          =   264
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   1
      Top             =   720
      Width           =   1212
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   4905
      TabIndex        =   22
      Top             =   540
      Width           =   4332
      Begin VB.OptionButton optType 
         Caption         =   "接洽及內部收文單"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   1755
      End
      Begin VB.OptionButton optType 
         Caption         =   "主管機關來函"
         Height          =   252
         Index           =   1
         Left            =   2160
         TabIndex        =   11
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "所有資料(&L)"
      Height          =   400
      Left            =   6118
      TabIndex        =   18
      Top             =   60
      Width           =   1152
   End
   Begin VB.CommandButton cmdChkCP14 
      Caption         =   "未分案(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4941
      TabIndex        =   17
      Top             =   60
      Width           =   1152
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "全部選取(&A)"
      Height          =   400
      Left            =   2587
      TabIndex        =   15
      Top             =   60
      Width           =   1152
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "全部清除(&D)"
      Height          =   400
      Left            =   3764
      TabIndex        =   16
      Top             =   60
      Width           =   1152
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1020
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Height          =   1092
      Left            =   120
      TabIndex        =   23
      Top             =   540
      Width           =   4692
      Begin VB.Line Line1 
         X1              =   2880
         X2              =   3120
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.Label Label3 
      Caption         =   "系統類別："
      Height          =   240
      Left            =   5070
      TabIndex        =   27
      Top             =   1470
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "收文所別："
      Height          =   240
      Left            =   5070
      TabIndex        =   25
      Top             =   1200
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   6390
      X2              =   6630
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Label Label2 
      Caption         =   "(1.北  2.中  3.南  4.高)"
      Height          =   240
      Left            =   7050
      TabIndex        =   24
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "frm030201_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/03/18 grdList : MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
'Memo by Lydia 2021/08/30 改成Form2.0 ; grdList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Dim m_SelQry As Integer
Dim m_SelType As String

' 宣告報表表頭的欄位其資料型態
'Private Type TMTYPEITEM
'   tiTMType As String
'   tiCaseList() As String
'   tiCaseCount As Integer
'End Type
'Dim m_TMTypeList() As TMTYPEITEM
'Dim m_TMTypeCount As Integer
'
Dim m_CurrSel As Integer
Dim intOrderQty As Integer 'Add By Sindy 2020/12/17 接洽單案件性質數量
Dim lngX As Long, lngY As Long 'Add By Sindy 2020/12/17

Private Sub cmdChkCP14_Click()
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   ' 讀取資料
   If CheckInputValid() = True Then
      QueryDB True
   End If
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
   'Add By Cheng 2002/04/24
   'Modify by Morgan 2003/12/23
   'If Me.grdList.Rows = 2 Then
   If Me.grdList.Rows = 2 And Me.Visible = True Then
   'Modify end 2003/12/23
   
      cmdSelAll_Click
      cmdOK_Click
   End If
   
   '910718 Sieg
   If grdList.Rows = 1 Then
      MsgBox "無符合條件資料 !", vbInformation
'      textTM01.SetFocus
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid() = True Then
      DisplayNextForm
   End If
End Sub

' 設定該筆收文資料已做完存檔的工作
Public Sub SetDataComplete(ByVal strCP09 As String)
   Dim nIndex As Integer
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 3) = strCP09 Then
         grdList.TextMatrix(nIndex, 0) = Empty
         UpdateCurrRecord nIndex
         Exit For
      End If
   Next nIndex
End Sub

Private Sub cmdQuery_Click()
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   ' 讀取資料
   If CheckInputValid() = True Then
      QueryDB False
   End If
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
   If grdList.Rows = 1 Then MsgBox "無符合條件資料 !", vbInformation
End Sub

'Add By Sindy 2020/12/17
'取消數量
Private Sub cmdReceive_Click(Index As Integer)
   With grdList
      '有數量才需要清空白
      If .TextMatrix(m_CurrSel, 0) = "V" And Val(.TextMatrix(m_CurrSel, 1)) > 0 Then
         strSql = "Update Caseprogress Set CP156=null Where CP09='" & .TextMatrix(m_CurrSel, 3) & "'"
         cnnConnection.Execute strSql
         .TextMatrix(m_CurrSel, 1) = ""
         .TextMatrix(m_CurrSel, 0) = ""
      End If
   End With
End Sub

Private Sub cmdSelAll_Click()
   Dim nIndex As Integer
   For nIndex = 1 To grdList.Rows - 1
      grdList.TextMatrix(nIndex, 0) = "V"
   Next nIndex
End Sub

Private Sub cmdClearAll_Click()
   Dim nIndex As Integer
   For nIndex = 1 To grdList.Rows - 1
      grdList.TextMatrix(nIndex, 0) = Empty
   Next nIndex
End Sub

'Private Sub AddTMTypeCase(ByVal strTM As String, ByVal strCase As String)
'   Dim nX As Integer
'   Dim nY As Integer
'   Dim bFind As Boolean
'
'   bFind = False
'   For nX = 0 To m_TMTypeCount - 1
'      If strTM = m_TMTypeList(nX).tiTMType Then
'         bFind = False
'         For nY = 0 To m_TMTypeList(nX).tiCaseCount - 1
'            If strCase = m_TMTypeList(nX).tiCaseList(nY) Then
'               bFind = True
'               Exit For
'            End If
'         Next nY
'         If bFind = False Then
'            ReDim Preserve m_TMTypeList(nX).tiCaseList(m_TMTypeList(nX).tiCaseCount + 1)
'            m_TMTypeList(nX).tiCaseList(m_TMTypeList(nX).tiCaseCount) = strCase
'            m_TMTypeList(nX).tiCaseCount = m_TMTypeList(nX).tiCaseCount + 1
'         End If
'         bFind = True
'         Exit For
'      End If
'   Next nX
'   If bFind = False Then
'      ReDim Preserve m_TMTypeList(m_TMTypeCount + 1)
'      nX = m_TMTypeCount
'      m_TMTypeList(nX).tiTMType = strTM
'      m_TMTypeList(nX).tiCaseCount = 1
'      ReDim Preserve m_TMTypeList(nX).tiCaseList(1)
'      m_TMTypeList(nX).tiCaseList(0) = strCase
'      m_TMTypeCount = m_TMTypeCount + 1
'   End If
'End Sub

'Private Sub QueryTMType()
'   Dim bFind As Boolean
'   Dim nX As Integer
'   Dim nY As Integer
'   Dim rsTmp As New ADODB.Recordset
'   Dim strSQL As String
'   strSQL = "SELECT ST11 FROM Staff " & _
'            "WHERE ST01 = '" & strUserNum & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      strSQL = Empty
'      If IsNull(rsTmp.Fields("ST11")) = False Then
'         strSQL = "SELECT * FROM Staff_Group " & _
'                  "WHERE SG01 = '" & rsTmp.Fields("ST11") & "' "
'      End If
'   End If
'   rsTmp.Close
'
'   If IsEmptyText(strSQL) = False Then
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount > 0 Then
'         rsTmp.MoveFirst
'         Do While rsTmp.EOF = False
'            If IsEmptyText(rsTmp.Fields("SG02")) = False And IsEmptyText(rsTmp.Fields("SG03")) = False Then
'               AddTMTypeCase rsTmp.Fields("SG02"), rsTmp.Fields("SG03")
'            End If
'            rsTmp.MoveNext
'         Loop
'      End If
'   End If
'End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   ' 設定初始化的資料
   InitialData
   InitialGrdList
End Sub

' 初始化資料
Private Sub InitialData()
   ' 預設為收文日
   'optSelect(0).Value = True
   m_SelQry = 0
   ' 收文日的預設值為系統日
   'Modified by Morgan 2022/12/29 系統日改抓變數
   'textCP05_1.Text = TAIWANDATE(SystemDate())
   'textCP05_2.Text = TAIWANDATE(SystemDate())
   'Modify By Sindy 2023/3/8
   'textCP05_1.Text = strSrvDate(2)
   'textCP05_2.Text = textCP05_1.Text
   textCP05_1.Text = TransDate(CompWorkDay(-2, strSrvDate(1), 1), 1)
   textCP05_2.Text = strSrvDate(2)
   '2023/3/8 END
   'end 2022/12/29
   ' 預設為接洽內部收文單
   optType(0).Value = True
   m_SelType = 0
   '93.6.27 add by sonia
   Text1 = PUB_GetST06(strUserNum)
   Text2 = 4 'PUB_GetST06(strUserNum) Modify By Sindy 2023/1/3
   '93.6.27 END
End Sub

' 檢查畫面上的資料輸入是否正確
Private Function CheckInputValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckInputValid = False
   
   If optSelect(0).Value = True Then
      If IsEmptyText(textCP05_1) Or IsEmptyText(textCP05_2) Then
         strTit = "檢核資料"
         strMsg = "請輸入收文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      If Val(textCP05_1) > Val(textCP05_2) Then
         strTit = "檢核資料"
         strMsg = "收文日範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   
   If optSelect(1).Value = True Then
      If IsEmptyText(textTM01) Then
         strTit = "檢核資料"
         strMsg = "請輸入本所案號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      If IsEmptyText(textTM02) Then
         strTit = "檢核資料"
         strMsg = "請輸入本所案號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      If textTM01 = "TF" And IsEmptyText(textTM02_2) Then
         strTit = "檢核資料"
         strMsg = "請輸入本所案號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   
   CheckInputValid = True
   
EXITSUB:
End Function

Public Sub QueryDB(ByVal bCP14 As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strSubSQL As String
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   Dim nX As Integer
   Dim nY As Integer
   Dim bFirst As Boolean
   Dim strTMType() As String
   Dim strGroup As String
   Dim strSG02 As String

   strSql = "SELECT ST11 FROM Staff " & _
            "WHERE ST01 = '" & strUserNum & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ST11")) = False Then
         strGroup = rsTmp.Fields("ST11")
      End If
   End If
   rsTmp.Close
   strSql = Empty

   strSG02 = Empty
   strSql = "SELECT DISTINCT SG02 FROM STAFF_GROUP " & _
            "WHERE SG01 = '" & strGroup & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         If IsNull(rsTmp.Fields("SG02")) = False Then
            If strSG02 <> Empty Then: strSG02 = strSG02 & ","
            strSG02 = strSG02 & "'" & rsTmp.Fields("SG02") & "'"
         End If
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   strSG02 = "(" & strSG02 & ")"
   
   'strSQL = "SELECT CP01,CP02,CP03,CP04,NVL(CP05 - 19110000, NULL) AS CP05,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07,CP09,NVL(DECODE(NVL(TM10,SP09),'000',C1.CPM03,C1.CPM04),CP10) AS CP10,S1.ST02 AS CP13,S2.ST02 AS CP14,CP57,TM05,TM06,TM07,TM29,SP05,SP06,SP07,SP15,cp27 FROM CASEPROGRESS, STAFF S1, STAFF S2, TRADEMARK, SERVICEPRACTICE, CASEPROPERTYMAP C1 "
   'Modify By Sindy 2020/12/17 + 數量 => ,decode(CP140||substr(cp09,1,1),'A',''||CP156,'-') as cntCP156,CP140,NVL(TM10,SP09) TM10
   'Modify by Amy 2022/11/10 +CP122
   'Modify by Sindy 2023/6/8 +Flow003
   strSql = "SELECT CP01,CP02,CP03,CP04,NVL(CP05 - 19110000, NULL) AS CP05,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07,CP09,NVL(DECODE(NVL(TM10,SP09),'000',C1.CPM03,C1.CPM04),CP10) AS CP10,S1.ST02 AS CP13,S2.ST02 AS CP14,CP27,CP57,TM05,TM06,TM07,TM29,SP05,SP06,SP07,SP15,SG01,SG02,SG03,decode(CP140||substr(cp09,1,1),'A',''||CP156,'-') as cntCP156,CP140,NVL(TM10,SP09) TM10,CP122 FROM CASEPROGRESS, STAFF S1, STAFF S2, TRADEMARK, SERVICEPRACTICE, CASEPROPERTYMAP C1, STAFF_GROUP, Flow003 "

   'strSubSQL = "CP01 = SG02 AND " & _
   '            "CP10 = SG03 AND " & _

   strSubSQL = "CP01 IN " & strSG02 & " AND " & _
               "'" & strGroup & "' = SG01(+) AND " & _
               "CP01 = SG02(+) AND " & _
               "CP10 = SG03(+) AND " & _
               "CP01 = TM01(+) AND " & _
               "CP02 = TM02(+) AND " & _
               "CP03 = TM03(+) AND " & _
               "CP04 = TM04(+) AND " & _
               "CP01 = SP01(+) AND " & _
               "CP02 = SP02(+) AND " & _
               "CP03 = SP03(+) AND " & _
               "CP04 = SP04(+) AND " & _
               "CP13 = S1.ST01(+) AND " & _
               "CP14 = S2.ST01(+) AND " & _
               "CP01 = C1.CPM01(+) AND " & _
               "CP10 = C1.CPM02(+) AND CP10<>'703' AND CP10<>'704' "

   'strSubSQL = "CP01 = SG02 AND CP10 = SG03 "
   Select Case m_SelQry
      ' 以收文日的範圍來查詢
      Case 0:
         If IsEmptyText(textCP05_1) = False Then
            If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
            strSubSQL = strSubSQL & " CP05 >= " & DBDATE(textCP05_1) & " "
         End If
         If IsEmptyText(textCP05_2) = False Then
            If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
            strSubSQL = strSubSQL & " CP05 <= " & DBDATE(textCP05_2) & " "
         End If
         If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
         strSubSQL = strSubSQL & " (CP01 = 'CFT' OR CP01 = 'FCT' OR CP01 = 'CFC' OR CP01 = 'S' ) "
      ' 以本所案號的範圍來查詢
      Case 1:
         ' 依本所案號的系統類別來判斷
         strTM01 = textTM01
         strTM02 = textTM02
         If textTM01 = "TF" Then: strTM02 = strTM02 & textTM02_2
         '2011/10/20 add by sonia FCT-006704會帶出FCT-006704-T-00
         If textTM03 = "" Then textTM03 = "0"
         If textTM04 = "" Then textTM04 = "00"
         '2011/10/20 END
         strTM03 = textTM03
         strTM04 = textTM04
         
         If IsEmptyText(strTM01) = False Then
            If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
            strSubSQL = strSubSQL & "CP01 = '" & strTM01 & "' "
         End If
         If IsEmptyText(strTM02) = False Then
            If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
            strSubSQL = strSubSQL & "CP02 = '" & strTM02 & "' "
         End If
         If IsEmptyText(strTM03) = False Then
            If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
            strSubSQL = strSubSQL & "CP03 = '" & strTM03 & "' "
         End If
         If IsEmptyText(strTM04) = False Then
            If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
            strSubSQL = strSubSQL & "CP04 = '" & strTM04 & "' "
         End If
      Case 2, 3: 'Modify By Sindy 2023/6/8
         If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
         strSubSQL = strSubSQL & " (CP01 = 'CFT' OR CP01 = 'FCT' OR CP01 = 'CFC' OR CP01 = 'S') "
         'Add By Sindy 2023/6/8 電子收文未分案
         If m_SelQry = 3 Then
            strSubSQL = strSubSQL & " And F0309='" & Flow_已收文 & "' And F0301 IS NOT NULL "
         End If
         '2023/6/8 END
   End Select

   Select Case m_SelType
      Case 0:
         If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
         strSubSQL = strSubSQL & "(CP09 LIKE 'A%' OR CP09 LIKE 'B%') "
      Case 1:
         If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
         strSubSQL = strSubSQL & "CP09 LIKE 'C%' "
   End Select
   
   'Modify By Sindy 2023/6/8 +電子收文未分案
   If bCP14 = True Or m_SelQry = 3 Then
      If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
      strSubSQL = strSubSQL & "(CP14 IS NULL OR CP14 = '' ) "
   End If
   
   'Add By Sindy 2021/1/18 查詢系統別XXX
   If Trim(TextSys.Text) <> "" Then
      If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
      strSubSQL = strSubSQL & "CP01 = '" & TextSys & "' "
   End If
   '2021/1/18 END
   
   '93.6.27 add by sonia 加收文所別
   If m_SelQry = 0 Then
      If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
      If Text1 <> "1" And Text2 <> "1" Then
         strSubSQL = strSubSQL & "S1.ST06 >= '" & Text1 & "' AND S1.ST06 <= '" & Text2 & "' "
      Else
         strSubSQL = strSubSQL & "((S1.ST06 >= '" & Text1 & "' AND S1.ST06 <= '" & Text2 & "') OR S1.ST06='5') "
      End If
   End If
   '93.6.27 END
   
   'Modify By Sindy 2023/6/8 + And CP140=F0301(+)
   If IsEmptyText(strSubSQL) = False Then strSql = strSql & "WHERE " & strSubSQL & " And CP140=F0301(+)"
   '91.12.12 MODIFY BY SONIA
   'strSQL = strSQL & " " & "ORDER BY CP01, CP02, CP03, CP04 "
   'Modify By Sindy 2012/12/18 阿蓮提選收文日條件時才依總收文號排序,選本所案號或以前未分案時 , 改依收文日 + 總收文號排序
   'strSql = strSql & " " & "ORDER BY CP09 "
   Select Case m_SelQry
      '以收文日的範圍來查詢
      Case 0:
         strSql = strSql & " " & "ORDER BY CP09 "
      '其他
      Case Else
         strSql = strSql & " " & "ORDER BY CP05,CP09 "
   End Select
   '2012/12/18 End
   '91.12.12 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ListData rsTmp
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub UpdateCurrRecord(ByVal nIndex As String)
   Dim rsSrcTmp As New ADODB.Recordset
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTM10 As String
   Dim bTM29 As Boolean
   Dim nCol As Integer

   If nIndex > 0 And nIndex <= grdList.Rows - 1 Then
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP09 = '" & grdList.TextMatrix(nIndex, 3) & "' "
      rsSrcTmp.CursorLocation = adUseClient
      rsSrcTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSrcTmp.RecordCount > 0 Then
         rsSrcTmp.MoveFirst
         grdList.TextMatrix(nIndex, 9) = "0"
         ' 案件名稱
         Select Case rsSrcTmp.Fields("CP01")
            Case "CFT", "FCT":
               strSql = "SELECT TM05,TM06,TM07,TM10,TM29 FROM TradeMark " & _
                        "WHERE TM01 = '" & rsSrcTmp.Fields("CP01") & "' AND " & _
                              "TM02 = '" & rsSrcTmp.Fields("CP02") & "' AND " & _
                              "TM03 = '" & rsSrcTmp.Fields("CP03") & "' AND " & _
                              "TM04 = '" & rsSrcTmp.Fields("CP04") & "' "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount > 0 Then
                  If IsNull(rsTmp.Fields("TM10")) = False Then
                     strTM10 = rsTmp.Fields("TM10")
                  End If
                  If IsNull(rsTmp.Fields("TM05")) = False Then
                     grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("TM05")
                  ElseIf IsNull(rsTmp.Fields("TM06")) = False Then
                     grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("TM06")
                  ElseIf IsNull(rsTmp.Fields("TM07")) = False Then
                     grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("TM07")
                  End If
                  If IsNull(rsTmp.Fields("TM29")) = False Then
                     If rsTmp.Fields("TM29") = "Y" Then
                        bTM29 = True
                     End If
                  End If
               End If
               rsTmp.Close
            Case Else:
               strSql = "SELECT SP05,SP06,SP07,SP09,SP15 FROM ServicePractice " & _
                        "WHERE SP01 = '" & rsSrcTmp.Fields("CP01") & "' AND " & _
                              "SP02 = '" & rsSrcTmp.Fields("CP02") & "' AND " & _
                              "SP03 = '" & rsSrcTmp.Fields("CP03") & "' AND " & _
                              "SP04 = '" & rsSrcTmp.Fields("CP04") & "' "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount > 0 Then
                  If IsNull(rsTmp.Fields("SP09")) = False Then
                     strTM10 = rsTmp.Fields("SP09")
                  End If
                  If IsNull(rsTmp.Fields("SP05")) = False Then
                     grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("SP05")
                  ElseIf IsNull(rsTmp.Fields("SP06")) = False Then
                     grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("SP06")
                  ElseIf IsNull(rsTmp.Fields("SP07")) = False Then
                     grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("SP07")
                  End If
                  If IsNull(rsTmp.Fields("SP15")) = False Then
                     If rsTmp.Fields("SP15") = "Y" Then
                        bTM29 = True
                     End If
                  End If
               End If
               rsTmp.Close
         End Select
         ' 收文日
         If IsNull(rsSrcTmp.Fields("CP05")) = False Then
            grdList.TextMatrix(nIndex, 2) = TAIWANDATE(rsSrcTmp.Fields("CP05"))
         End If
         ' 收文號
         If IsNull(rsSrcTmp.Fields("CP09")) = False Then
            grdList.TextMatrix(nIndex, 3) = rsSrcTmp.Fields("CP09")
         End If
         ' 案件性質
         If IsNull(rsSrcTmp.Fields("CP10")) = False Then
            If strTM10 < "010" Then
               grdList.TextMatrix(nIndex, 5) = GetCaseTypeName(rsSrcTmp.Fields("CP01"), rsSrcTmp.Fields("CP10"), 0)
            Else
               grdList.TextMatrix(nIndex, 5) = GetCaseTypeName(rsSrcTmp.Fields("CP01"), rsSrcTmp.Fields("CP10"), 1)
            End If
         End If
         ' 本所案號
         If rsSrcTmp.Fields("CP01") = "TF" Then
            grdList.TextMatrix(nIndex, 6) = rsSrcTmp.Fields("CP01") & "-" & Mid(rsSrcTmp.Fields("CP02"), 1, 5) & "-" & Mid(rsSrcTmp.Fields("CP02"), 6, 1) & "-" & rsSrcTmp.Fields("CP03") & "-" & rsSrcTmp.Fields("CP04")
         Else
            grdList.TextMatrix(nIndex, 6) = rsSrcTmp.Fields("CP01") & "-" & rsSrcTmp.Fields("CP02") & "-" & rsSrcTmp.Fields("CP03") & "-" & rsSrcTmp.Fields("CP04")
         End If
         ' 智權人員
         If IsNull(rsSrcTmp.Fields("CP13")) = False Then
            grdList.TextMatrix(nIndex, 7) = GetStaffName(rsSrcTmp.Fields("CP13"))
         End If
         ' 承辦人
         If IsNull(rsSrcTmp.Fields("CP14")) = False Then
            grdList.TextMatrix(nIndex, 8) = GetStaffName(rsSrcTmp.Fields("CP14"))
         End If
         ' 顏色
         grdList.TextMatrix(nIndex, 9) = Empty
         'Modify by Amy 2022/11/10 +CP122=Y 顯示紅色
         'If IsNull(rsSrcTmp.Fields("CP06")) = False Then
         If "" & rsSrcTmp.Fields("CP06") <> MsgText(601) Or "" & rsSrcTmp.Fields("CP122") = "Y" Then
            If ("" & rsSrcTmp.Fields("CP06") <> MsgText(601) And Val(DBDATE("" & rsSrcTmp.Fields("CP06"))) < Val(DBDATE(SystemDate())) And IsNull(rsSrcTmp.Fields("CP27"))) _
              Or "" & rsSrcTmp.Fields("CP122") = "Y" Then
               grdList.TextMatrix(nIndex, 9) = "1"
            End If
         End If
         If grdList.TextMatrix(nIndex, 9) = "0" Then
            If bTM29 = True Then
               grdList.TextMatrix(nIndex, 9) = "2"
            ElseIf IsNull(rsSrcTmp.Fields("CP57")) = False Then
               If rsSrcTmp.Fields("CP57") <> "0" Then
                  grdList.TextMatrix(nIndex, 9) = "3"
               End If
            End If
         End If

         Select Case grdList.TextMatrix(nIndex, 9)
            Case "1":
               For nCol = 1 To grdList.Cols - 1
                  grdList.row = nIndex
                  grdList.col = nCol
                  grdList.CellBackColor = &H8080FF '紅色
               Next nCol
            Case "2":
               For nCol = 1 To grdList.Cols - 1
                  grdList.row = nIndex
                  grdList.col = nCol
                  grdList.CellBackColor = &HFFFF&
               Next nCol
            Case "3":
               For nCol = 1 To grdList.Cols - 1
                  grdList.row = nIndex
                  grdList.col = nCol
                  grdList.CellBackColor = &HE0E0E0
               Next nCol
            Case Else:
               For nCol = 1 To grdList.Cols - 1
                  grdList.row = nIndex
                  grdList.col = nCol
                  grdList.CellBackColor = &H80000005
                  grdList.CellForeColor = vbBlack
               Next nCol
         End Select
      End If
      rsSrcTmp.Close
   End If

   Set rsSrcTmp = Nothing
   Set rsTmp = Nothing
End Sub

Private Sub ListData(ByRef rsSrcTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTM10 As String
   Dim bTM29 As Boolean
   Dim nCol As Integer
      
   InitialGrdList
   
   bTM29 = False
   If rsSrcTmp.RecordCount > 0 Then
      rsSrcTmp.MoveFirst
      Do While rsSrcTmp.EOF = False
         ' 若SG01,SG02,SG03其中無值表示該使用者的群組無此系統別及案件性質的使用權限
         If IsNull(rsSrcTmp.Fields("SG01")) Or IsNull(rsSrcTmp.Fields("SG02")) Or IsNull(rsSrcTmp.Fields("SG03")) Then
            GoTo NextRecord
         End If
         
         grdList.Rows = grdList.Rows + 1
         nIndex = grdList.Rows - 1
         bTM29 = False
         grdList.TextMatrix(nIndex, 9) = "0"
         ' 案件名稱
         Select Case rsSrcTmp.Fields("CP01")
            Case "CFT", "FCT":
               'strSQL = "SELECT TM05,TM06,TM07,TM10,TM29 FROM TradeMark " & _
               '         "WHERE TM01 = '" & rsSrcTmp.Fields("CP01") & "' AND " & _
               '               "TM02 = '" & rsSrcTmp.Fields("CP02") & "' AND " & _
               '               "TM03 = '" & rsSrcTmp.Fields("CP03") & "' AND " & _
               '               "TM04 = '" & rsSrcTmp.Fields("CP04") & "' "
               'rsTmp.CursorLocation = adUseClient
               'rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
               'If rsTmp.RecordCount > 0 Then
               '   If IsNull(rsTmp.Fields("TM10")) = False Then
               '      strTM10 = rsTmp.Fields("TM10")
               '   End If
               '   If IsNull(rsTmp.Fields("TM05")) = False Then
               '      grdList.TextMatrix(nIndex, 3) = rsTmp.Fields("TM05")
               '   ElseIf IsNull(rsTmp.Fields("TM06")) = False Then
               '      grdList.TextMatrix(nIndex, 3) = rsTmp.Fields("TM06")
               '   ElseIf IsNull(rsTmp.Fields("TM07")) = False Then
               '      grdList.TextMatrix(nIndex, 3) = rsTmp.Fields("TM07")
               '   End If
               '   If IsNull(rsTmp.Fields("TM29")) = False Then
               '      If rsTmp.Fields("TM29") = "Y" Then
               '         bTM29 = True
               '      End If
               '   End If
               'End If
               'rsTmp.Close
               If IsNull(rsSrcTmp.Fields("TM05")) = False Then
                  grdList.TextMatrix(nIndex, 4) = rsSrcTmp.Fields("TM05")
               ElseIf IsNull(rsSrcTmp.Fields("TM06")) = False Then
                  grdList.TextMatrix(nIndex, 4) = rsSrcTmp.Fields("TM06")
               ElseIf IsNull(rsSrcTmp.Fields("TM07")) = False Then
                  grdList.TextMatrix(nIndex, 4) = rsSrcTmp.Fields("TM07")
               End If
               If IsNull(rsSrcTmp.Fields("TM29")) = False Then
                  If rsSrcTmp.Fields("TM29") = "Y" Then
                     bTM29 = True
                  End If
               End If
            Case Else:
               'strSQL = "SELECT SP05,SP06,SP07,SP09,SP15 FROM ServicePractice " & _
               '         "WHERE SP01 = '" & rsSrcTmp.Fields("CP01") & "' AND " & _
               '               "SP02 = '" & rsSrcTmp.Fields("CP02") & "' AND " & _
               '               "SP03 = '" & rsSrcTmp.Fields("CP03") & "' AND " & _
               '               "SP04 = '" & rsSrcTmp.Fields("CP04") & "' "
               'rsTmp.CursorLocation = adUseClient
               'rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
               'If rsTmp.RecordCount > 0 Then
               '   If IsNull(rsTmp.Fields("SP09")) = False Then
               '      strTM10 = rsTmp.Fields("SP09")
               '   End If
               '   If IsNull(rsTmp.Fields("SP05")) = False Then
               '      grdList.TextMatrix(nIndex, 3) = rsTmp.Fields("SP05")
               '   ElseIf IsNull(rsTmp.Fields("SP06")) = False Then
               '      grdList.TextMatrix(nIndex, 3) = rsTmp.Fields("SP06")
               '   ElseIf IsNull(rsTmp.Fields("SP07")) = False Then
               '      grdList.TextMatrix(nIndex, 3) = rsTmp.Fields("SP07")
               '   End If
               '   If IsNull(rsTmp.Fields("SP15")) = False Then
               '      If rsTmp.Fields("SP15") = "Y" Then
               '         bTM29 = True
               '      End If
               '   End If
               'End If
               'rsTmp.Close
               If IsNull(rsSrcTmp.Fields("SP05")) = False Then
                  grdList.TextMatrix(nIndex, 4) = rsSrcTmp.Fields("SP05")
               ElseIf IsNull(rsSrcTmp.Fields("SP06")) = False Then
                  grdList.TextMatrix(nIndex, 4) = rsSrcTmp.Fields("SP06")
               ElseIf IsNull(rsSrcTmp.Fields("SP07")) = False Then
                  grdList.TextMatrix(nIndex, 4) = rsSrcTmp.Fields("SP07")
               End If
               If IsNull(rsSrcTmp.Fields("SP15")) = False Then
                  If rsSrcTmp.Fields("SP15") = "Y" Then
                     bTM29 = True
                  End If
               End If
         End Select
         
         'Add By Sindy 2020/12/17 + 數量
         If IsNull(rsSrcTmp.Fields("cntCP156")) = False Then
            grdList.TextMatrix(nIndex, 1) = rsSrcTmp.Fields("cntCP156")
         End If
         '申請國家
         If IsNull(rsSrcTmp.Fields("TM10")) = False Then
            grdList.TextMatrix(nIndex, 10) = rsSrcTmp.Fields("TM10")
         End If
         '2020/12/17 END
         
         ' 收文日
         If IsNull(rsSrcTmp.Fields("CP05")) = False Then
            'grdList.TextMatrix(nIndex, 1) = TAIWANDATE(rsSrcTmp.Fields("CP05"))
            grdList.TextMatrix(nIndex, 2) = rsSrcTmp.Fields("CP05")
         End If
         ' 收文號
         If IsNull(rsSrcTmp.Fields("CP09")) = False Then
            grdList.TextMatrix(nIndex, 3) = rsSrcTmp.Fields("CP09")
         End If
         ' 案件性質
         If IsNull(rsSrcTmp.Fields("CP10")) = False Then
            'If strTM10 < "010" Then
            '   grdList.TextMatrix(nIndex, 4) = GetCaseTypeName(rsSrcTmp.Fields("CP01"), rsSrcTmp.Fields("CP10"), 0)
            'Else
            '   grdList.TextMatrix(nIndex, 4) = GetCaseTypeName(rsSrcTmp.Fields("CP01"), rsSrcTmp.Fields("CP10"), 1)
            'End If
            grdList.TextMatrix(nIndex, 5) = rsSrcTmp.Fields("CP10")
         End If
         ' 本所案號
         If rsSrcTmp.Fields("CP01") = "TF" Then
            grdList.TextMatrix(nIndex, 6) = rsSrcTmp.Fields("CP01") & "-" & Mid(rsSrcTmp.Fields("CP02"), 1, 5) & "-" & Mid(rsSrcTmp.Fields("CP02"), 6, 1) & "-" & rsSrcTmp.Fields("CP03") & "-" & rsSrcTmp.Fields("CP04")
         Else
            grdList.TextMatrix(nIndex, 6) = rsSrcTmp.Fields("CP01") & "-" & rsSrcTmp.Fields("CP02") & "-" & rsSrcTmp.Fields("CP03") & "-" & rsSrcTmp.Fields("CP04")
         End If
         ' 智權人員
         If IsNull(rsSrcTmp.Fields("CP13")) = False Then
            'grdList.TextMatrix(nIndex, 6) = GetStaffName(rsSrcTmp.Fields("CP13"))
            grdList.TextMatrix(nIndex, 7) = rsSrcTmp.Fields("CP13")
         End If
         ' 承辦人
         If IsNull(rsSrcTmp.Fields("CP14")) = False Then
            'grdList.TextMatrix(nIndex, 7) = GetStaffName(rsSrcTmp.Fields("CP14"))
            grdList.TextMatrix(nIndex, 8) = rsSrcTmp.Fields("CP14")
         End If
         ' 顏色
         'Modify by Amy 2022/11/10 +CP122=Y 顯示紅色
         'If IsNull(rsSrcTmp.Fields("CP06")) = False Then
         If "" & rsSrcTmp.Fields("CP06") <> MsgText(601) Or "" & rsSrcTmp.Fields("CP122") = "Y" Then
            If ("" & rsSrcTmp.Fields("CP06") <> MsgText(601) And Val(DBDATE("" & rsSrcTmp.Fields("CP06"))) < Val(DBDATE(SystemDate())) And IsNull(rsSrcTmp.Fields("CP27"))) _
              Or "" & rsSrcTmp.Fields("CP122") = "Y" Then
               grdList.TextMatrix(nIndex, 9) = "1"
            End If
         End If
         If grdList.TextMatrix(nIndex, 9) = "0" Then
            If bTM29 = True Then
               grdList.TextMatrix(nIndex, 9) = "2"
            ElseIf IsNull(rsSrcTmp.Fields("CP57")) = False Then
               If rsSrcTmp.Fields("CP57") <> "0" Then
                  grdList.TextMatrix(nIndex, 9) = "3"
               End If
            End If
         End If
         
         Select Case grdList.TextMatrix(nIndex, 9)
            Case "1":
               For nCol = 1 To grdList.Cols - 1
                  grdList.row = nIndex
                  grdList.col = nCol
                  grdList.CellBackColor = &H8080FF
               Next nCol
            Case "2":
               For nCol = 1 To grdList.Cols - 1
                  grdList.row = nIndex
                  grdList.col = nCol
                  grdList.CellBackColor = &HFFFF&
               Next nCol
            Case "3":
               For nCol = 1 To grdList.Cols - 1
                  grdList.row = nIndex
                  grdList.col = nCol
                  grdList.CellBackColor = &HE0E0E0
               Next nCol
         End Select
NextRecord:
         rsSrcTmp.MoveNext
      Loop
      'Added by Lydia 2022/03/18 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
      If grdList.Rows > 1 Then
         grdList.FixedRows = 1
      End If
   End If
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030201_01 = Nothing
End Sub

Private Sub optSelect_Click(Index As Integer)
On Error Resume Next
   m_SelQry = Index
   
   '910718 Sieg
   textCP05_1.Enabled = False
   textCP05_2.Enabled = False
   textTM01.Enabled = False
   textTM02.Enabled = False
   textTM02_2.Enabled = False
   textTM03.Enabled = False
   textTM04.Enabled = False
   Select Case Index
      Case 0
         textCP05_1.Enabled = True
         textCP05_2.Enabled = True
         textCP05_1.SetFocus
      Case 1
         textTM01.Enabled = True
         textTM02.Enabled = True
         textTM02_2.Enabled = True
         textTM03.Enabled = True
         textTM04.Enabled = True
         textTM01.SetFocus
   End Select
End Sub

Private Sub optType_Click(Index As Integer)
   m_SelType = Index
End Sub

' 初始化 GridList
'Modifed by Lydia 2022/03/18
Private Sub InitialGrdList()
   grdList.Clear
   
   grdList.Rows = 1
   grdList.Cols = 11 '9
   
   grdList.ColWidth(0) = 300
   
   'Add By Sindy 2020/12/17 + 數量
   grdList.TextMatrix(0, 1) = "數量"
   grdList.ColWidth(1) = 450
   grdList.ColAlignment(1) = flexAlignCenterCenter
   '2020/12/17 END
   
   grdList.TextMatrix(0, 2) = "收文日"
   grdList.ColWidth(2) = 1000
   grdList.ColAlignment(2) = flexAlignCenterCenter
   
   grdList.TextMatrix(0, 3) = "收文號"
   grdList.ColWidth(3) = 1000
   grdList.ColAlignment(3) = flexAlignLeftCenter
   
   grdList.TextMatrix(0, 4) = "案件名稱"
   grdList.ColWidth(4) = 2000
   grdList.ColAlignment(4) = flexAlignLeftCenter
      
   grdList.TextMatrix(0, 5) = "案件性質"
   grdList.ColWidth(5) = 1200
   grdList.ColAlignment(5) = flexAlignLeftCenter
      
   grdList.TextMatrix(0, 6) = "本所案號"
   grdList.ColWidth(6) = 1600
   grdList.ColAlignment(6) = flexAlignLeftCenter
      
   grdList.TextMatrix(0, 7) = "智權人員"
   grdList.ColWidth(7) = 1200
   grdList.ColAlignment(7) = flexAlignLeftCenter
      
   grdList.TextMatrix(0, 8) = "承辦人"
   grdList.ColWidth(8) = 1200
   grdList.ColAlignment(8) = flexAlignLeftCenter
      
   grdList.TextMatrix(0, 9) = "特殊"
   grdList.ColWidth(9) = 0
   
   'Add By Sindy 2020/12/17 + 申請國家
   grdList.TextMatrix(0, 10) = "申請國家"
   grdList.ColWidth(10) = 0

End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If nSel > 0 And nSel < grdList.Rows And grdList.Rows >= 2 Then
      grdList.row = nSel
'      grdList_SelChange
   End If
End Sub

Private Sub grdList_Click()
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         If grdList.TextMatrix(grdList.row, 0) = "V" Then
            grdList.TextMatrix(grdList.row, 0) = Empty
         Else
            grdList.TextMatrix(grdList.row, 0) = "V"
            cmdOK.SetFocus
         End If
      End If
   End If
   grdList_ShowSelection
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      If grdList.row > 0 Then
         If grdList.TextMatrix(grdList.row, 0) = "V" Then
            grdList.TextMatrix(grdList.row, 0) = Empty
         Else
            grdList.TextMatrix(grdList.row, 0) = "V"
         End If
      End If
   End If
End Sub

'Private Sub grdList_SelChange()
'   grdList_ShowSelection
'End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   Dim intCurrCol As Integer 'Add By Sindy 2020/12/17
   Dim strHeight As String
   Dim i As Integer
   
   nCurrSel = grdList.row
   intCurrCol = grdList.col 'Add By Sindy 2020/12/17
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      Dim nOldCol As Integer
      nOldCol = grdList.col
      grdList.col = 1
      If grdList.CellBackColor <> &H8000000D Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
         Next nCol
      End If
      grdList.col = nOldCol
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
   
EXITSUB:
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      'Add By Sindy 2020/12/17 A類非電子收文者，均要輸入接洽單數量
      If grdList.TextMatrix(m_CurrSel, 0) = "V" Then
         'Add By Sindy 2020/12/17 + And intCurrCol = 1
         If grdList.TextMatrix(m_CurrSel, 1) <> "-" And Val(grdList.TextMatrix(m_CurrSel, 1)) = 0 _
            And intCurrCol = 1 Then
            '彈出表單位置控制
            frm040101_3.Label3.Caption = grdList.TextMatrix(m_CurrSel, 6) & " (" & grdList.TextMatrix(m_CurrSel, 3) & ")"
            strHeight = mdiMain.Top + Me.Top + grdList.Top + lngY + (mdiMain.Height - mdiMain.ScaleHeight) + (Me.Height - Me.ScaleHeight)
            If Val(strHeight) + frm040101_3.Height > Val(Val(mdiMain.Top + mdiMain.Height)) Then
               strHeight = Val(strHeight) - frm040101_3.Height - Val(grdList.RowHeight(1))
            End If
            frm040101_3.Move mdiMain.Left + Me.Left + grdList.Left + lngX, Val(strHeight)
            frm040101_3.Show vbModal
            intOrderQty = Val(strPublicTemp)
            strPublicTemp = ""
            If intOrderQty = 0 Then
               grdList.TextMatrix(grdList.row, 1) = ""
               For i = 1 To grdList.Cols - 1
                  grdList.col = i
                  grdList.CellBackColor = grdList.BackColor
               Next
            Else
               grdList.TextMatrix(grdList.row, 1) = intOrderQty
            End If
            strSql = "Update Caseprogress Set CP156=" & IIf(Val(intOrderQty) > 0, intOrderQty, "Null") & " Where CP09='" & grdList.TextMatrix(m_CurrSel, 3) & "'"
            cnnConnection.Execute strSql
         End If
      End If
      '2019/7/18 END
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
'EXITSUB:
End Sub
'93.6.27 ADD BY SONIA
' 收文所別(起)
Private Sub Text1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(Text1) = True Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "收文所別(起)不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      If Text1 < "1" Or Text1 > "4" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文所別(起)只可為 '1'~'4' "
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub
' 收文所別(止)
Private Sub Text2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(Text2) = True Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "收文所別(止)不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      If Text2 < "1" Or Text2 > "4" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文所別(止)只可為 '1'~'4' "
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Else
         If Text2 < Text1 Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "收文所別範圍錯誤 "
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End If
      End If
   End If
End Sub
'93.6.27 END

' 收文日(起)
Private Sub textCP05_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05_1) = False Then
      If CheckIsTaiwanDate(textCP05_1, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文日(起)日期不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

' 收文日(止)
Private Sub textCP05_2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05_2) = False Then
      If CheckIsTaiwanDate(textCP05_2, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文日(止)日期不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

Private Sub textSys_GotFocus()
   CloseIme
   TextSys.SelStart = 0
   TextSys.SelLength = Len(TextSys)
End Sub

Private Sub textSys_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
   textTM02_2.Visible = False
   textTM02_2.Locked = True
   textTM02_2.TabStop = False
   textTM02.MaxLength = 6
            
   If IsEmptyText(textTM01) = False Then
      Select Case textTM01
         Case "CFT", "FCT", "CFC", "S":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "本所案號中的系統別只可為CFT、FCT、CFC或S"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End Select
   End If
   If Cancel Then TextInverse textTM01
EXITSUB:
End Sub

Private Sub DisplayNextForm()
   Dim nIndex As Integer
   Dim bClear As Boolean
   bClear = True
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         'Add By Sindy 2020/12/17 數量預設為1
         If grdList.TextMatrix(nIndex, 1) <> "-" And Val(grdList.TextMatrix(nIndex, 1)) = 0 Then
            strSql = "Update Caseprogress Set CP156=1 Where CP09='" & grdList.TextMatrix(nIndex, 3) & "'"
            cnnConnection.Execute strSql
            grdList.TextMatrix(nIndex, 1) = "1"
         End If
         '2020/12/17 END
         frm030201_02.SetData 0, grdList.TextMatrix(nIndex, 3), bClear
         bClear = False
      End If
   Next nIndex
   Me.Hide
   frm030201_02.Show
   frm030201_02.QueryData
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   CheckDataValid = False
   If grdList.Rows <= 1 Then
      strTit = "檢核資料"
      strMsg = "無資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   bFind = False
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      strTit = "檢核資料"
      strMsg = "請先選取一筆資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCP05_1_GotFocus()
   InverseTextBox textCP05_1
End Sub

Private Sub textCP05_2_GotFocus()
   InverseTextBox textCP05_2
End Sub

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
   CloseIme
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
   CloseIme
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
   CloseIme
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
   CloseIme
End Sub

Private Sub textTM03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
   CloseIme
End Sub
'93.6.27 ADD BY SONIA
Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub
'93.6.27 END
