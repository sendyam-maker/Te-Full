VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm020101_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "分案"
   ClientHeight    =   5750
   ClientLeft      =   -460
   ClientTop       =   1790
   ClientWidth     =   9340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   9340
   Begin VB.OptionButton optSelect 
      Caption         =   "電子收文未分案:"
      Height          =   252
      Index           =   3
      Left            =   2100
      TabIndex        =   3
      Top             =   1320
      Width           =   1812
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "取消數量(&C)"
      Height          =   400
      Index           =   1
      Left            =   1500
      TabIndex        =   25
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   6840
      MaxLength       =   1
      TabIndex        =   14
      Top             =   1270
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   6120
      MaxLength       =   1
      TabIndex        =   13
      Top             =   1270
      Width           =   300
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "全部清除(&D)"
      Height          =   400
      Left            =   3952
      TabIndex        =   22
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelAll 
      Caption         =   "全部選取(&A)"
      Height          =   400
      Left            =   2726
      TabIndex        =   15
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdChkCP14 
      Caption         =   "未分案(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5178
      TabIndex        =   16
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "所有資料(&L)"
      Height          =   400
      Left            =   6404
      TabIndex        =   17
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox textCP05_2 
      Height          =   264
      Left            =   3300
      MaxLength       =   20
      TabIndex        =   5
      Top             =   720
      Width           =   1212
   End
   Begin VB.TextBox textCP05_1 
      Height          =   264
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   4
      Top             =   720
      Width           =   1212
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1020
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1020
      Width           =   372
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1020
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7630
      TabIndex        =   18
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8460
      TabIndex        =   19
      Top             =   70
      Width           =   800
   End
   Begin VB.OptionButton optSelect 
      Caption         =   "以前未分案 :"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1332
   End
   Begin VB.OptionButton optSelect 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1020
      Width           =   1332
   End
   Begin VB.OptionButton optSelect 
      Caption         =   "收文日 :"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1332
   End
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   3120
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1020
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Height          =   1092
      Left            =   120
      TabIndex        =   20
      Top             =   540
      Width           =   4812
      Begin VB.Line Line1 
         X1              =   2880
         X2              =   3120
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.Frame Frame2 
      Height          =   555
      Left            =   5040
      TabIndex        =   21
      Top             =   540
      Width           =   4212
      Begin VB.OptionButton optType 
         Caption         =   "主管機關來函"
         Height          =   252
         Index           =   1
         Left            =   2040
         TabIndex        =   12
         Top             =   180
         Width           =   1572
      End
      Begin VB.OptionButton optType 
         Caption         =   "接洽內部收文單"
         Height          =   252
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   180
         Width           =   1572
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4032
      Left            =   96
      TabIndex        =   26
      Top             =   1656
      Width           =   9132
      _ExtentX        =   16104
      _ExtentY        =   7108
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
      Caption         =   "(1.北  2.中  3.南  4.高)"
      Height          =   240
      Left            =   7320
      TabIndex        =   24
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Line Line2 
      X1              =   6480
      X2              =   6720
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "收文所別："
      Height          =   240
      Left            =   5160
      TabIndex        =   23
      Top             =   1320
      Width           =   900
   End
End
Attribute VB_Name = "frm020101_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/11/30 Form2.0已修改 grdList
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
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
Dim m_CurrSel As Integer
'Add by Morgan 2003/12/05
'前次是否為未分案查詢
Dim bolIsChk As Boolean
Dim strGroup As String
Dim strSG02 As String
Dim intOrderQty As Integer  'Add By Sindy 2019/7/18 接洽單案件性質數量
Dim lngX As Long, lngY As Long 'Add By Sindy 2019/7/18
Dim arrF, arrW  'Add by Amy 2022/10/06
Dim stDefArea1 As String, stDefArea2 As String 'Add by Amy 2022/11/17

'未分案
Private Sub cmdChkCP14_Click()
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   ' 讀取資料
   QueryDB True
   
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
   'Add By Cheng 2002/04/23
   '若只搜尋到一筆資料時, 直接進入下一畫面
   'Modify by Morgan 2003/12/23
   'If Me.grdList.Rows = 2 Then
   If Me.grdList.Rows = 2 And Me.Visible = True Then
   'Modify end 2003/12/23
   
      cmdSelAll_Click
      cmdOK_Click
   End If
   If grdList.Rows = 1 Then MsgBox "無符合條件資料 !", vbInformation
   
   'Modify by Morgan 2003/12/04
   bolIsChk = True
   'End 2003/12/05
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
      'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 3)
      If grdList.TextMatrix(nIndex, GetValue("收文號")) = strCP09 Then
         grdList.TextMatrix(nIndex, 0) = Empty
         UpdateCurrRecord nIndex
         Exit For
      End If
   Next nIndex
End Sub

'所有資料
Private Sub cmdQuery_Click()
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   ' 讀取資料
   QueryDB False
   
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
   
   'Add By Morgan 2003/12/05
   '若只搜尋到一筆資料時, 直接進入下一畫面
   If Me.grdList.Rows = 2 Then
      cmdSelAll_Click
      cmdOK_Click
   End If
   bolIsChk = False
   'End 2003/12/05
   
   If grdList.Rows = 1 Then MsgBox "無符合條件資料 !", vbInformation
End Sub

'Add By Sindy 2019/7/19
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
   'Modify by Amy 2022/11/07 +CRL59
   'Modify by Amy 2022/11/09 CRL59 改為CP122
   arrF = Split("|數量|收文日|收文號|案件名稱|案件性質|本所案號|智權人員|目前表單狀態|承辦人|特殊|申請國家|F0301|CP122", "|")
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
        arrW = Split("300;450;1000;1000;2000;1200;1600;1200;1200;1200;0;0;0;0", ";")
   Else
        arrW = Split("300;450;1000;1000;2000;1200;1600;1200;0;1200;0;0;0;0", ";")
   End If
   'end 2022/11/07
   MoveFormToCenter Me
   ' 設定初始化的資料
   InitialData
   InitialGrdList
   'cmdChkCP14_Click
    'Add By Cheng 2004/02/18
    strGroup = ""
    strSG02 = ""
    'End
    'Add by Amy 2022/11/17
    stDefArea1 = Text1
    stDefArea2 = Text2
End Sub

' 初始化資料
Private Sub InitialData()
   ' 預設為收文日
   optSelect(0).Value = True
   m_SelQry = 0
   ' 收文日的預設值為系統日
    'Modify By Cheng 2004/01/29
'   textCP05_1.Text = TAIWANDATE(SystemDate())
'   textCP05_2.Text = TAIWANDATE(SystemDate())
   'Modify By Sindy 2023/3/8
   'textCP05_1.Text = strSrvDate(2)
   'textCP05_2.Text = textCP05_1.Text
   'modify by sonia 2024/1/11 起日改為前7個工作天
   'textCP05_1.Text = TransDate(CompWorkDay(-2, strSrvDate(1), 1), 1)
   textCP05_1.Text = TransDate(CompWorkDay(-7, strSrvDate(1), 1), 1)
   textCP05_2.Text = strSrvDate(2)
   '2023/3/8 END
    'End
   ' 預設為接洽內部收文單
   optType(0).Value = True
   m_SelType = 0
   '93.6.27 add by sonia
   Text1 = PUB_GetST06(strUserNum)
   'Modify by Amy 2022/11/02 接洽單電子收文上線後無紙本, 北所進入預設顯示全部
   If Text1 = "1" And strSrvDate(1) >= 接洽單電子收文啟用日 Then
        Text2 = "4"
   Else
        Text2 = PUB_GetST06(strUserNum)
   End If
   '93.6.27 END

End Sub
'Modify by Morgan 2003/12/05
'Public Sub QueryDB(ByVal bCP14 As Boolean)
Public Sub QueryDB(Optional ByVal bCP14 As Boolean)
Dim strField As String 'Add By Sindy 2022/8/19
Dim strWhere As String 'Add by Amy 2022/11/15

   If Me.Visible = False Then bCP14 = bolIsChk
   
'End 2003/12/05
   
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
    'Modify By Cheng 2004/02/18
    '改成Module層級的變數
'   Dim strGroup As String
'   Dim strSG02 As String
    'End
    'Modify By Cheng 2004/02/18
    '若使用者群組別變數無值
    If strGroup = "" Then
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
    End If
   strSql = Empty
    'Modify By Cheng 2004/02/18
    '若使用者可使用的系統類別變數無值
'   strSG02 = Empty
    If strSG02 = "" Then
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
    End If
       
   'strSQL = "SELECT CP01,CP02,CP03,CP04,NVL(CP05 - 19110000, NULL) AS CP05,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07,CP09,NVL(DECODE(NVL(TM10,SP09),'000',C1.CPM03,C1.CPM04),CP10) AS CP10,S1.ST02 AS CP13,S2.ST02 AS CP14,CP57,TM05,TM06,TM07,TM29,SP05,SP06,SP07,SP15,CP27 FROM CASEPROGRESS, STAFF S1, STAFF S2, TRADEMARK, SERVICEPRACTICE, CASEPROPERTYMAP C1 "
   'Modify By Sindy 2019/7/18 + 數量 => ,decode(CP140||substr(cp09,1,1),'A',''||CP156,'-') as cntCP156,CP140,NVL(TM10,SP09) TM10
   'Add By Sindy 2022/8/19 北所分案日有值才顯示承辦人,因為電子收文後分案主管會先填入承辦人
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      strField = " S2.ST02 AS CP14,"
   Else
      strField = " Decode(Nvl(cp157,0),0,'',S2.ST02) AS CP14,"
   End If
   '2022/8/19 END
   'Modify by Amy 2022/10/06 +目前表單狀態,F0301
   'Modify by Amy 2022/10/19 原:Decode(F0309,null,'',Decode(F0309," & ShowFlow表單狀態中文 & ")),接洽單案件性質不止一筆,若某筆已分案其他筆,不可顯示已分案
   'Modify by Amy 2022/11/09  +CP122
   'Modify by Amy 2022/11/18 原:Decode(F0309,'" & Flow_已分案 & "',Decode(cp157,null,'處理中',Decode(F0309,null,''," & ShowFlow表單狀態中文 & "))," & ShowFlow表單狀態中文 & "),改為接洽單多筆案件性質,最後一筆才上Flow相關資料
   strSql = "SELECT CP01,CP02,CP03,CP04,NVL(CP05 - 19110000, NULL) AS CP05,NVL(CP06 - 19110000, NULL) AS CP06,NVL(CP07 - 19110000, NULL) AS CP07,CP09,NVL(DECODE(NVL(TM10,SP09),'000',C1.CPM03,C1.CPM04),CP10) AS CP10,S1.ST02 AS CP13" & _
                ",Decode(F0309,'" & Flow_處理中 & "',Decode(Decode(Decode(cp157,null,0,1),1,'已分案',F0309),'已分案','已分案'," & ShowFlow表單狀態中文 & ")," & ShowFlow表單狀態中文 & ") as Status,CP27," & strField & "CP57,TM05,TM06,TM07,TM29,SP05,SP06,SP07,SP15,SG01,SG02,SG03,decode(CP140||substr(cp09,1,1),'A',''||CP156,'-') as cntCP156,CP140,NVL(TM10,SP09) TM10,F0301,CP122 " & _
                "FROM CASEPROGRESS, STAFF S1, STAFF S2, TRADEMARK, SERVICEPRACTICE, CASEPROPERTYMAP C1, STAFF_GROUP,Flow003,ConsultRecordList "
   
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
      ' 以本所案號的範圍來查詢
      Case 1:
         ' 依本所案號的系統類別來判斷
         strTM01 = textTM01
         strTM02 = textTM02
         If textTM01 = "TF" Then: strTM02 = strTM02 & textTM02_2
         If textTM03 = "" Then textTM03 = "0"
         If textTM04 = "" Then textTM04 = "00"
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
      Case 2:
      Case 3:
         'Add By Sindy 2023/6/8 電子收文未分案
         If m_SelQry = 3 Then
            If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
            strSubSQL = strSubSQL & "F0308='A7' And F0309='" & Flow_處理中 & "' And F0301 IS NOT NULL "
         End If
         '2023/6/8 END
   End Select
   
   Select Case m_SelType
      Case 0:
         If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
        'Modify By Cheng 2004/02/18
'         strSubSQL = strSubSQL & "(CP09 LIKE 'A%' OR CP09 LIKE 'B%') "
         strSubSQL = strSubSQL & " CP09 < 'C' "
        'End
      Case 1:
         If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
        'Modify By Cheng 2004/02/18
'         strSubSQL = strSubSQL & "CP09 LIKE 'C%' "
         strSubSQL = strSubSQL & " CP09 > 'C' "
        'End
   End Select
   
   'Modify By Sindy 2023/6/8 +電子收文未分案
   If bCP14 = True Or m_SelQry = 3 Then
      If IsEmptyText(strSubSQL) = False Then: strSubSQL = strSubSQL & " AND "
      'Modify By Sindy 2022/8/19 增加判斷是否有北所分案日,因為電子收文後分案主管會先填入承辦人
'      If strSrvDate(1) >= 20220901 Then
         strSubSQL = strSubSQL & "(CP14 IS NULL OR CP14 = '' OR CP157 IS NULL)"
'      Else
'      '2022/8/19 END
'         strSubSQL = strSubSQL & "(CP14 IS NULL OR CP14 = '')"
'      End If
   End If
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
   'Add by Amy 2022/10/06 +接洽單電子收文流程(Flow003)
   'Modify by Amy 2022/11/07 +串 接洽單主檔
   'Modify by Amy 2022/11/15 +if 未分案才加判斷條件
   'Modify By Sindy 2023/6/8 + And m_SelQry <> 3
   If bCP14 = True And m_SelQry <> 3 Then
      strWhere = "And (F0308='A7' Or F0309='" & Flow_已分案 & "' Or F0301 IS NULL) "
   End If
   strSubSQL = strSubSQL & " And CP140=F0301(+) " & strWhere & " And CP140=CRL01(+) "
   '91.10.25 MODIFY BY SONIA
   'If IsEmptyText(strSubSQL) = False Then strSQL = strSQL & "WHERE " & strSubSQL
   If IsEmptyText(strSubSQL) = False Then strSql = strSql & "WHERE " & strSubSQL & " ORDER BY CP05, CP09"
   '91.10.25 END
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
      'Modif by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 3)
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP09 = '" & grdList.TextMatrix(nIndex, GetValue("收文號")) & "' "
      rsSrcTmp.CursorLocation = adUseClient
      rsSrcTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSrcTmp.RecordCount > 0 Then
         rsSrcTmp.MoveFirst
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 9)
         grdList.TextMatrix(nIndex, GetValue("特殊")) = "0"
         ' 案件名稱
         Select Case rsSrcTmp.Fields("CP01")
            Case "T", "TF", "FC":
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
                  'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 4)
                  If IsNull(rsTmp.Fields("TM05")) = False Then
                     grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsTmp.Fields("TM05")
                  ElseIf IsNull(rsTmp.Fields("TM06")) = False Then
                     grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsTmp.Fields("TM06")
                  ElseIf IsNull(rsTmp.Fields("TM07")) = False Then
                     grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsTmp.Fields("TM07")
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
                  'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 4)
                  If IsNull(rsTmp.Fields("SP05")) = False Then
                     grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsTmp.Fields("SP05")
                  ElseIf IsNull(rsTmp.Fields("SP06")) = False Then
                     grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsTmp.Fields("SP06")
                  ElseIf IsNull(rsTmp.Fields("SP07")) = False Then
                     grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsTmp.Fields("SP07")
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
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 2)
         If IsNull(rsSrcTmp.Fields("CP05")) = False Then
            grdList.TextMatrix(nIndex, GetValue("收文日")) = TAIWANDATE(rsSrcTmp.Fields("CP05"))
         End If
         ' 收文號
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 3)
         If IsNull(rsSrcTmp.Fields("CP09")) = False Then
            grdList.TextMatrix(nIndex, GetValue("收文號")) = rsSrcTmp.Fields("CP09")
         End If
         ' 案件性質
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 5)
         If IsNull(rsSrcTmp.Fields("CP10")) = False Then
            If strTM10 < "010" Then
               grdList.TextMatrix(nIndex, GetValue("案件性質")) = GetCaseTypeName(rsSrcTmp.Fields("CP01"), rsSrcTmp.Fields("CP10"), 0)
            Else
               grdList.TextMatrix(nIndex, GetValue("案件性質")) = GetCaseTypeName(rsSrcTmp.Fields("CP01"), rsSrcTmp.Fields("CP10"), 1)
            End If
         End If
         ' 本所案號
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 6)
         If rsSrcTmp.Fields("CP01") = "TF" Then
            grdList.TextMatrix(nIndex, GetValue("本所案號")) = rsSrcTmp.Fields("CP01") & "-" & Mid(rsSrcTmp.Fields("CP02"), 1, 5) & "-" & Mid(rsSrcTmp.Fields("CP02"), 6, 1) & "-" & rsSrcTmp.Fields("CP03") & "-" & rsSrcTmp.Fields("CP04")
         Else
            grdList.TextMatrix(nIndex, GetValue("本所案號")) = rsSrcTmp.Fields("CP01") & "-" & rsSrcTmp.Fields("CP02") & "-" & rsSrcTmp.Fields("CP03") & "-" & rsSrcTmp.Fields("CP04")
         End If
         ' 智權人員
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 7)
         If IsNull(rsSrcTmp.Fields("CP13")) = False Then
            grdList.TextMatrix(nIndex, GetValue("智權人員")) = GetStaffName(rsSrcTmp.Fields("CP13"))
         End If
         ' 承辦人
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 8)
         If IsNull(rsSrcTmp.Fields("CP14")) = False Then
            grdList.TextMatrix(nIndex, GetValue("承辦人")) = GetStaffName(rsSrcTmp.Fields("CP14"))
         End If
         ' 顏色
         'Modify by Amy 2022/11/09 +CP122=Y 顯示紅色
         'If IsNull(rsSrcTmp.Fields("CP06")) = False  then
         If "" & rsSrcTmp.Fields("CP06") <> MsgText(601) Or "" & rsSrcTmp.Fields("CP122") = "Y" Then
            If ("" & rsSrcTmp.Fields("CP06") <> MsgText(601) And Val(DBDATE("" & rsSrcTmp.Fields("CP06"))) < Val(DBDATE(SystemDate())) And IsNull(rsSrcTmp.Fields("CP27"))) _
              Or "" & rsSrcTmp.Fields("CP122") = "Y" Then
         'end 2022/11/09
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 9)
               grdList.TextMatrix(nIndex, GetValue("特殊")) = "1"
            End If
         End If
         'end 2022/11/07
         If grdList.TextMatrix(nIndex, GetValue("特殊")) = "0" Then
            If bTM29 = True Then
               grdList.TextMatrix(nIndex, GetValue("特殊")) = "2"
            ElseIf IsNull(rsSrcTmp.Fields("CP57")) = False Then
               If rsSrcTmp.Fields("CP57") <> "0" Then
                  grdList.TextMatrix(nIndex, GetValue("特殊")) = "3"
               End If
            End If
         End If
         'Add By Sindy 2022/10/31
         If grdList.TextMatrix(nIndex, GetValue("特殊")) = "0" Then
            'Add By Sindy 2018/7/18 自動收文變綠色
            If Trim("" & rsSrcTmp.Fields("CP140")) <> "" Then
               grdList.TextMatrix(nIndex, GetValue("特殊")) = "4"
            End If
         End If
         Select Case grdList.TextMatrix(nIndex, GetValue("特殊"))
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
                  grdList.CellBackColor = &HFFFF& '黃色
               Next nCol
            Case "3":
               For nCol = 1 To grdList.Cols - 1
                  grdList.row = nIndex
                  grdList.col = nCol
                  grdList.CellBackColor = &HE0E0E0 '灰色
               Next nCol
            'Add By Sindy 2018/7/18
            Case "4":
               For nCol = 1 To grdList.Cols - 1
                  grdList.row = nIndex
                  grdList.col = nCol
                  grdList.CellBackColor = &HFF7F& '綠色
               Next nCol
            '2018/7/18 END
         End Select
      End If
      'Add by Amy 2022/12/23 需補件者,案件性質顯示 粉紅色
      If grdList.TextMatrix(nIndex, GetValue("目前表單狀態")) = "程序補件" Then
         grdList.row = nIndex
         grdList.col = GetValue("案件性質")
         grdList.CellBackColor = &HFF80FF     '粉紅色
      End If
      rsSrcTmp.Close
   End If
   
   Set rsSrcTmp = Nothing
   Set rsTmp = Nothing
End Sub

Private Sub ListData(ByRef rsSrcTmp As ADODB.Recordset)
   Dim nIndex As Long 'Integer
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
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 9)
         grdList.TextMatrix(nIndex, GetValue("特殊")) = "0"
         ' 案件名稱
         Select Case rsSrcTmp.Fields("CP01")
            Case "T", "TF", "FCT":
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
               '      grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("TM05")
               '   ElseIf IsNull(rsTmp.Fields("TM06")) = False Then
               '      grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("TM06")
               '   ElseIf IsNull(rsTmp.Fields("TM07")) = False Then
               '      grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("TM07")
               '   End If
               '   If IsNull(rsTmp.Fields("TM29")) = False Then
               '      If rsTmp.Fields("TM29") = "Y" Then
               '         bTM29 = True
               '      End If
               '   End If
               'End If
               'rsTmp.Close
               'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 4)
               If IsNull(rsSrcTmp.Fields("TM05")) = False Then
                  grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsSrcTmp.Fields("TM05")
               ElseIf IsNull(rsSrcTmp.Fields("TM06")) = False Then
                  grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsSrcTmp.Fields("TM06")
               ElseIf IsNull(rsSrcTmp.Fields("TM07")) = False Then
                  grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsSrcTmp.Fields("TM07")
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
               '      grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("SP05")
               '   ElseIf IsNull(rsTmp.Fields("SP06")) = False Then
               '      grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("SP06")
               '   ElseIf IsNull(rsTmp.Fields("SP07")) = False Then
               '      grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("SP07")
               '   End If
               '   If IsNull(rsTmp.Fields("SP15")) = False Then
               '      If rsTmp.Fields("SP15") = "Y" Then
               '         bTM29 = True
               '      End If
               '   End If
               'End If
               'rsTmp.Close
               If IsNull(rsSrcTmp.Fields("SP05")) = False Then
                  grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsSrcTmp.Fields("SP05")
               ElseIf IsNull(rsSrcTmp.Fields("SP06")) = False Then
                  grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsSrcTmp.Fields("SP06")
               ElseIf IsNull(rsSrcTmp.Fields("SP07")) = False Then
                  grdList.TextMatrix(nIndex, GetValue("案件名稱")) = rsSrcTmp.Fields("SP07")
               End If
               If IsNull(rsSrcTmp.Fields("SP15")) = False Then
                  If rsSrcTmp.Fields("SP15") = "Y" Then
                     bTM29 = True
                  End If
               End If
         End Select
         
         'Add By Sindy 2019/7/18 + 數量
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 1)
         If IsNull(rsSrcTmp.Fields("cntCP156")) = False Then
            grdList.TextMatrix(nIndex, GetValue("數量")) = rsSrcTmp.Fields("cntCP156")
         End If
         '申請國家
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 10)
         If IsNull(rsSrcTmp.Fields("TM10")) = False Then
            grdList.TextMatrix(nIndex, GetValue("申請國家")) = rsSrcTmp.Fields("TM10")
         End If
         '2019/7/18 END
         
         ' 收文日
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 2)
         If IsNull(rsSrcTmp.Fields("CP05")) = False Then
            'grdList.TextMatrix(nIndex, 2) = TAIWANDATE(rsSrcTmp.Fields("CP05"))
            grdList.TextMatrix(nIndex, GetValue("收文日")) = rsSrcTmp.Fields("CP05")
         End If
         ' 收文號
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 3)
         If IsNull(rsSrcTmp.Fields("CP09")) = False Then
            grdList.TextMatrix(nIndex, GetValue("收文號")) = rsSrcTmp.Fields("CP09")
         End If
         ' 案件性質
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 5)
         If IsNull(rsSrcTmp.Fields("CP10")) = False Then
            'If strTM10 < "010" Then
            '   grdList.TextMatrix(nIndex, 5) = GetCaseTypeName(rsSrcTmp.Fields("CP01"), rsSrcTmp.Fields("CP10"), 0)
            'Else
            '   grdList.TextMatrix(nIndex, 5) = GetCaseTypeName(rsSrcTmp.Fields("CP01"), rsSrcTmp.Fields("CP10"), 1)
            'End If
            grdList.TextMatrix(nIndex, GetValue("案件性質")) = rsSrcTmp.Fields("CP10")
         End If
         ' 本所案號
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 6)
         If rsSrcTmp.Fields("CP01") = "TF" Then
            grdList.TextMatrix(nIndex, GetValue("本所案號")) = rsSrcTmp.Fields("CP01") & "-" & Mid(rsSrcTmp.Fields("CP02"), 1, 5) & "-" & Mid(rsSrcTmp.Fields("CP02"), 6, 1) & "-" & rsSrcTmp.Fields("CP03") & "-" & rsSrcTmp.Fields("CP04")
         Else
            grdList.TextMatrix(nIndex, GetValue("本所案號")) = rsSrcTmp.Fields("CP01") & "-" & rsSrcTmp.Fields("CP02") & "-" & rsSrcTmp.Fields("CP03") & "-" & rsSrcTmp.Fields("CP04")
         End If
         ' 智權人員
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 7)
         If IsNull(rsSrcTmp.Fields("CP13")) = False Then
            'grdList.TextMatrix(nIndex, 7) = GetStaffName(rsSrcTmp.Fields("CP13"))
            grdList.TextMatrix(nIndex, GetValue("智權人員")) = rsSrcTmp.Fields("CP13")
         End If
         ' 承辦人
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 8)
         If IsNull(rsSrcTmp.Fields("CP14")) = False Then
            'grdList.TextMatrix(nIndex, 8) = GetStaffName(rsSrcTmp.Fields("CP14"))
            grdList.TextMatrix(nIndex, GetValue("承辦人")) = rsSrcTmp.Fields("CP14")
         End If
         'Add by Amy 2022/10/06 目前表單狀態/F0301
         If IsNull(rsSrcTmp.Fields("Status")) = False Then
            grdList.TextMatrix(nIndex, GetValue("目前表單狀態")) = rsSrcTmp.Fields("Status")
         End If
         If IsNull(rsSrcTmp.Fields("F0301")) = False Then
            grdList.TextMatrix(nIndex, GetValue("F0301")) = rsSrcTmp.Fields("F0301")
         End If
         'end 2022/10/06
         
         ' 顏色
         'Modify by Amy 2022/11/09 +CP122=Y 顯示紅色
         'If IsNull(rsSrcTmp.Fields("CP06")) = False then
         If "" & rsSrcTmp.Fields("CP06") <> MsgText(601) Or "" & rsSrcTmp.Fields("CP122") = "Y" Then
            'edit by nickc 2006/03/17
            'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 9)
            'If Val(DBDATE(rsSrcTmp.Fields("CP06"))) < Val(DBDATE(Date)) And IsNull(rsSrcTmp.Fields("CP27")) Then
            If ("" & rsSrcTmp.Fields("CP06") <> MsgText(601) And Val(DBDATE("" & rsSrcTmp.Fields("CP06"))) < Val(strSrvDate(1)) And IsNull(rsSrcTmp.Fields("CP27"))) _
              Or "" & rsSrcTmp.Fields("CP122") = "Y" Then
               grdList.TextMatrix(nIndex, GetValue("特殊")) = "1"
'            'Add By Sindy 2018/7/18 自動收文變綠色
'            ElseIf Trim("" & rsSrcTmp.Fields("CP140")) <> "" Then
'               grdList.TextMatrix(nIndex, GetValue("特殊")) = "4"
            End If
            '2018/7/18 END
         End If
         If grdList.TextMatrix(nIndex, GetValue("特殊")) = "0" Then
            If bTM29 = True Then '閉卷
               grdList.TextMatrix(nIndex, GetValue("特殊")) = "2"
            ElseIf IsNull(rsSrcTmp.Fields("CP57")) = False Then '已取消收文
               If rsSrcTmp.Fields("CP57") <> "0" Then
                  grdList.TextMatrix(nIndex, GetValue("特殊")) = "3"
               End If
            End If
         End If
         'Modify By Sindy 2022/10/31
         If grdList.TextMatrix(nIndex, GetValue("特殊")) = "0" Then
            'Add By Sindy 2018/7/18 自動收文變綠色
            If Trim("" & rsSrcTmp.Fields("CP140")) <> "" Then
               grdList.TextMatrix(nIndex, GetValue("特殊")) = "4"
            End If
         End If
         Select Case grdList.TextMatrix(nIndex, GetValue("特殊"))
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
                  grdList.CellBackColor = &HFFFF& '黃色
               Next nCol
            Case "3":
               For nCol = 1 To grdList.Cols - 1
                  grdList.row = nIndex
                  grdList.col = nCol
                  grdList.CellBackColor = &HE0E0E0 '灰色
               Next nCol
            'Add By Sindy 2018/7/18
            Case "4":
               For nCol = 1 To grdList.Cols - 1
                  grdList.row = nIndex
                  grdList.col = nCol
                  grdList.CellBackColor = &HFF7F& '綠色
               Next nCol
            '2018/7/18 END
         End Select
         'Add by Amy 2022/12/23 需補件者,案件性質顯示 粉紅色
         If grdList.TextMatrix(nIndex, GetValue("目前表單狀態")) = "程序補件" Then
            grdList.row = nIndex
            grdList.col = GetValue("案件性質")
            grdList.CellBackColor = &HFF80FF     '粉紅色
         End If
NextRecord:
         rsSrcTmp.MoveNext
      Loop
     
      If grdList.Rows > 1 Then 'Added by lydia 2022/04/22 以111/4/1~111/4/22查未分案會因為無資料而出錯;
          grdList.FixedRows = 1  'Added by Amy 2022/04/21 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
      End If
   End If
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm020101_01 = Nothing
End Sub

Private Sub optSelect_Click(Index As Integer)
   m_SelQry = Index
   'Add by Amy 2022/11/17 避免資料量太多,造成溢位,選 以前未分案 帶User所別
    If Index = 2 Then
        Text1 = PUB_GetST06(strUserNum)
        Text2 = Text1
    Else
        Text1 = stDefArea1
        Text2 = stDefArea2
    End If
    'end 2022/11/17
End Sub

Private Sub optType_Click(Index As Integer)
   m_SelType = Index
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
    Dim ii As Integer
    
    grdList.Clear
    grdList.Rows = 1
    grdList.Cols = UBound(arrF) + 1
    
    For ii = LBound(arrF) To UBound(arrF)
        grdList.TextMatrix(0, ii) = arrF(ii)
        grdList.ColWidth(ii) = arrW(ii)
        grdList.ColAlignment(ii) = flexAlignLeftCenter 'Add By Sindy 2023/5/8
    Next ii
End Sub

'Mark by Amy 2022/10/06 改寫法,避免加欄位有未改的欄位
Private Sub InitialGrdList_Old()
'   grdList.Clear
'
'   grdList.Rows = 1
'   grdList.Cols = 11 '9
'
'   grdList.ColWidth(0) = 300
'
'   'Add By Sindy 2019/7/18 + 數量
'   grdList.TextMatrix(0, 1) = "數量"
'   grdList.ColWidth(1) = 450
'   grdList.ColAlignment(1) = flexAlignCenterCenter
'   '2019/7/18 END
'
'   grdList.TextMatrix(0, 2) = "收文日"
'   grdList.ColWidth(2) = 1000
'   grdList.ColAlignment(2) = flexAlignCenterCenter
'
'   grdList.TextMatrix(0, 3) = "收文號"
'   grdList.ColWidth(3) = 1000
'   grdList.ColAlignment(3) = flexAlignLeftCenter
'
'   grdList.TextMatrix(0, 4) = "案件名稱"
'   grdList.ColWidth(4) = 2000
'   grdList.ColAlignment(4) = flexAlignLeftCenter
'
'   grdList.TextMatrix(0, 5) = "案件性質"
'   grdList.ColWidth(5) = 1200
'   grdList.ColAlignment(5) = flexAlignLeftCenter
'
'   grdList.TextMatrix(0, 6) = "本所案號"
'   grdList.ColWidth(6) = 1600
'   grdList.ColAlignment(6) = flexAlignLeftCenter
'
'   grdList.TextMatrix(0, 7) = "智權人員"
'   grdList.ColWidth(7) = 1200
'   grdList.ColAlignment(7) = flexAlignLeftCenter
'
'   grdList.TextMatrix(0, 8) = "承辦人"
'   grdList.ColWidth(8) = 1200
'   grdList.ColAlignment(8) = flexAlignLeftCenter
'
'   grdList.TextMatrix(0, 9) = "特殊"
'   grdList.ColWidth(9) = 0
'
'   'Add By Sindy 2019/7/18 + 申請國家
'   grdList.TextMatrix(0, 10) = "申請國家"
'   grdList.ColWidth(10) = 0
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If nSel > 0 And nSel < grdList.Rows And grdList.Rows >= 2 Then
      grdList.row = nSel
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

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
Dim nCurrSel As Integer
Dim nCol As Integer
Dim strHeight As String
Dim i As Integer
Dim intCurrCol As Integer 'Add By Sindy 2019/7/24
   
   nCurrSel = grdList.row
   intCurrCol = grdList.col 'Add By Sindy 2019/7/24
   
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
      'Add By Sindy 2019/7/18 A類非電子收文者，均要輸入接洽單數量
      If grdList.TextMatrix(m_CurrSel, 0) = "V" Then
         'Add By Sindy 2019/7/24 + And intCurrCol = 1
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(m_CurrSel, 1)
         If grdList.TextMatrix(m_CurrSel, GetValue("數量")) <> "-" And Val(grdList.TextMatrix(m_CurrSel, GetValue("數量"))) = 0 _
            And intCurrCol = 1 Then
            '彈出表單位置控制
            'Modify by Amy 2022/10/06 原:grdList.TextMatrix(m_CurrSel, 6)/grdList.TextMatrix(m_CurrSel, 3)
            frm040101_3.Label3.Caption = grdList.TextMatrix(m_CurrSel, GetValue("本所案號")) & " (" & grdList.TextMatrix(m_CurrSel, GetValue("收文號")) & ")"
            strHeight = mdiMain.Top + Me.Top + grdList.Top + lngY + (mdiMain.Height - mdiMain.ScaleHeight) + (Me.Height - Me.ScaleHeight)
            If Val(strHeight) + frm040101_3.Height > Val(Val(mdiMain.Top + mdiMain.Height)) Then
               strHeight = Val(strHeight) - frm040101_3.Height - Val(grdList.RowHeight(1))
            End If
            frm040101_3.Move mdiMain.Left + Me.Left + grdList.Left + lngX, Val(strHeight)
            frm040101_3.Show vbModal
            intOrderQty = Val(strPublicTemp)
            strPublicTemp = ""
            If intOrderQty = 0 Then
               'Modify by Amy 2022/10/06 原:grdList.TextMatrix(grdList.row, 1)
               grdList.TextMatrix(grdList.row, GetValue("數量")) = ""
               For i = 1 To grdList.Cols - 1
                  grdList.col = i
                  grdList.CellBackColor = grdList.BackColor
               Next
            Else
               grdList.TextMatrix(grdList.row, GetValue("數量")) = intOrderQty
            End If
            'Modify by Amy 2022/10/06 原:grdList.TextMatrix(m_CurrSel, 3)
            strSql = "Update Caseprogress Set CP156=" & IIf(Val(intOrderQty) > 0, intOrderQty, "Null") & " Where CP09='" & grdList.TextMatrix(m_CurrSel, GetValue("收文號")) & "'"
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
      If textTM01 <> "FCT" Then
         If Mid(textTM01, 1, 1) <> "T" Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "本所案號中的系統別只可為T類及FCT"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End If
      End If
      
      Select Case textTM01
         Case "TF":
            textTM02_2.Visible = True
            textTM02_2.Locked = False
            textTM02_2.TabStop = True
            textTM02.MaxLength = 5
         Case Else
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
      End Select
   Else
      textTM02_2.Visible = False
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02.MaxLength = 6
   End If
   
EXITSUB:
End Sub

Private Sub DisplayNextForm()
   Dim nIndex As Integer
   Dim bClear As Boolean
   bClear = True
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         'add by nickc 2006/04/26
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 5)
         If Trim(grdList.TextMatrix(nIndex, GetValue("案件性質"))) = "領土延伸" Then
            'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 3)
            UpdateTmData grdList.TextMatrix(nIndex, GetValue("收文號"))
         End If
         'Add By Sindy 2019/7/24 數量預設為1
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 1)
         If grdList.TextMatrix(nIndex, GetValue("數量")) <> "-" And Val(grdList.TextMatrix(nIndex, GetValue("數量"))) = 0 Then
            'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 3)
            strSql = "Update Caseprogress Set CP156=1 Where CP09='" & grdList.TextMatrix(nIndex, GetValue("收文號")) & "'"
            cnnConnection.Execute strSql
            'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 1)
            grdList.TextMatrix(nIndex, GetValue("數量")) = "1"
         End If
         '2019/7/24 END
         'Modify by Amy 2022/10/06 原:grdList.TextMatrix(nIndex, 3)
         'frm020101_02.txtF0309 = Trim(grdList.TextMatrix(nIndex, GetValue("目前表單狀態"))) 'Add By Sindy 2022/11/22
         frm020101_02.SetData 0, grdList.TextMatrix(nIndex, GetValue("收文號")), bClear
         bClear = False
      End If
   Next nIndex
   Me.Hide
   frm020101_02.SetParent Me
   frm020101_02.Show
   frm020101_02.QueryData
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

'add by nickc 2006/04/26 TF 014 將母案資料複製過來
Sub UpdateTmData(oStr As String)
Dim tmpRs As New ADODB.Recordset
Dim tmpSQL As String
Dim iCount As Integer
Dim TmpUpSQL As String
Set tmpRs = New ADODB.Recordset
tmpSQL = "select tm.* from trademark TM,caseprogress where cp09='" & oStr & "' and cp01=tm01(+) and substr(cp02,1,5) ||'0'=tm02(+) and cp03=tm03(+) and cp04=tm04(+) "
TmpUpSQL = ""
With tmpRs
    .CursorLocation = adUseClient
    .Open tmpSQL, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        For iCount = 1 To .Fields.Count
            Select Case iCount
            Case 1, 2, 3, 4, 10, 59 To 64
            Case Else
                If Len(Trim(TmpUpSQL)) <> 0 Then
                    TmpUpSQL = TmpUpSQL & ","
                End If
                'TmpUpSQL = TmpUpSQL & .Fields.Item(iCount - 1).Name & "=" & IIf(Len(Trim(ChgSQL(CheckStr(.Fields.Item(iCount - 1).Value)))) = 0, "null", IIf(.Fields.Item(iCount - 1).Type = adVarChar, "'", "") & ChgSQL(CheckStr(.Fields.Item(iCount - 1).Value)) & IIf(.Fields.Item(iCount - 1).Type = adVarChar, "'", ""))
                TmpUpSQL = TmpUpSQL & .Fields.Item(iCount - 1).Name & "=" & IIf(Len(Trim(ChgSQL(CheckStr(.Fields.Item(iCount - 1).Value)))) = 0, "null", "'" & ChgSQL(CheckStr(.Fields.Item(iCount - 1).Value)) & "'")
            End Select
        Next iCount
    End If
End With
If Trim(TmpUpSQL) <> "" Then
   cnnConnection.Execute "update trademark set " & TmpUpSQL & " where tm01||tm02||tm03||tm04 in (select distinct cp01||cp02||cp03||cp04 from caseprogress where cp09='" & oStr & "' ) "
   'add by nickc 2006/11/16  順便複製商品
   'Modify By Sindy 2012/3/27 +,tg15,tg16,tg17
   'cnnConnection.Execute "insert into tmgoods select tm01,tm02,tm03,tm04,tg05,tg06,tg07,tg08,'" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI')),null,null,null from trademark,tmgoods where tm01||tm02||tm03||tm04 in (select distinct cp01||cp02||cp03||cp04 from caseprogress where cp09='" & oStr & "' ) and tm01=tg01(+) and substr(tm02,1,5) ||'0'=tg02(+) and tm03=tg03(+) and tm04=tg04(+) "
   'Add By Sindy 2012/3/27 檢查是否有商品資料
   tmpRs.Close
   tmpSQL = "select count(*) from tmgoods,caseprogress where cp09='" & oStr & "' and cp01=tg01 and substr(cp02,1,5) ||'0'=tg02 and cp03=tg03 and cp04=tg04"
   tmpRs.CursorLocation = adUseClient
   tmpRs.Open tmpSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If tmpRs.RecordCount <> 0 Then
      If tmpRs.Fields(0) > 0 Then
   '2012/3/27 End
         '2012/11/8 ADD BY SONIA 領土延伸案之商品資料先刪除,否則會重覆 TF-000122-0-00
         cnnConnection.Execute "delete tmgoods where (tg01,tg02,tg03,tg04) in (select cp01,cp02,cp03,cp04 from caseprogress where cp09='" & oStr & "')"
         '2012/11/8 END
         cnnConnection.Execute "insert into tmgoods select tm01,tm02,tm03,tm04,tg05,tg06,tg07,tg08,'" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI')),null,null,null,tg15,tg16,tg17 from trademark,tmgoods where tm01||tm02||tm03||tm04 in (select distinct cp01||cp02||cp03||cp04 from caseprogress where cp09='" & oStr & "' ) and tm01=tg01(+) and substr(tm02,1,5) ||'0'=tg02(+) and tm03=tg03(+) and tm04=tg04(+) "
      End If
   End If
End If
End Sub

'Add by Amy 2022/10/06
Private Function GetValue(pFieldN As String) As Integer
   Dim jj As Integer
 
    For jj = LBound(arrF) To UBound(arrF)
        If UCase(arrF(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function
