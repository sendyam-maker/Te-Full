VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm04010506_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "異議/舉發受理函輸入"
   ClientHeight    =   5748
   ClientLeft      =   1692
   ClientTop       =   1512
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9336
   Begin VB.Frame Frame1 
      Height          =   612
      Left            =   120
      TabIndex        =   11
      Top             =   336
      Width           =   7410
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(F)"
         Default         =   -1  'True
         Height          =   300
         Left            =   6192
         TabIndex        =   6
         Top             =   216
         Width           =   1065
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1224
         MaxLength       =   20
         TabIndex        =   0
         Top             =   216
         Width           =   1572
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   4140
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "P"
         Top             =   216
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   4620
         MaxLength       =   6
         TabIndex        =   3
         Top             =   216
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   5460
         MaxLength       =   1
         TabIndex        =   4
         Top             =   216
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   5700
         MaxLength       =   2
         TabIndex        =   5
         Top             =   216
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "對造號數:"
         Height          =   204
         Left            =   96
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1128
      End
      Begin VB.OptionButton Option2 
         Caption         =   "本所案號:"
         Height          =   204
         Left            =   3024
         TabIndex        =   1
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   7
      Top             =   1032
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7536
      TabIndex        =   8
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8364
      TabIndex        =   9
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4152
      Left            =   144
      TabIndex        =   13
      Top             =   1440
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   7324
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
      Caption         =   "來函收文日:"
      Height          =   252
      Left            =   144
      TabIndex        =   10
      Top             =   1056
      Width           =   972
   End
End
Attribute VB_Name = "frm04010506_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Morgan 2021/12/20 改成Form2.0 (grdList)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Modify by Morgan 2008/8/18 已改開窗定稿，地址條列印功能取消
Option Explicit
Public NUMBER As String
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_CP09 As String
'Added by Morgan 2014/1/14
Public m_DocWord As String 'Added by Morgan 2014/4/17
Public m_DocNo As String
Public m_AppNo As String
Public m_RDate As String
Dim m_Done As Boolean
Dim m_Retry As Boolean 'Added by Morgan 2014/6/12
'end 2014/1/14
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
'2016/10/5 END


Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0: '確定
        'Modify By Cheng 2003/01/03
        If Me.grdList.Rows > 1 Then
            DisplayNextForm
        End If
      Case 1: '結束
         Unload Me
   End Select
End Sub

Private Sub Command1_Click()
   Dim bQuery As Boolean
   
   If Option2.Value = True Then
      If Text1 = "" Then
         MsgBox "請輸入本所案號", vbCritical
         Text1.SetFocus
         Exit Sub
      End If
      If Text2 = "" Then
         MsgBox "請輸入本所案號", vbCritical
         Text2.SetFocus
         Exit Sub
      End If
      If Text3.Text = "" Then Text3.Text = "0"
      If Text4.Text = "" Then Text4.Text = "00"
      
      bQuery = False
        'Add by Lydia 2014/10/31 先判斷外專程序人員權限。
        If FMP2open = True And FMP2openSQL <> "" Then
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text1, Text2, Text3, Text4) = False Then
            Me.Text2.SetFocus
            Exit Sub
           End If
        End If
      strExc(0) = "SELECT * FROM PATENT " & _
               "WHERE PA01 = '" & Text1 & "' AND " & _
                     "PA02 = '" & Text2 & "' AND " & _
                     "PA03 = '" & Text3 & "' AND " & _
                     "PA04 = '" & Text4 & "' "
        
      intI = 1
      bQuery = False
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If IsNull(RsTemp.Fields("PA09")) = False Then
            If RsTemp.Fields("PA09") < "010" Then
               MsgBox "申請國家為台灣, 請以對造號數查詢 !", vbInformation, "檢核資料"
               Option1.Value = True
               Option1_Click
            Else
               bQuery = True
            End If
         Else
            bQuery = True
         End If
      Else
        MsgBox "本所案號不存在 !", vbOKOnly + vbCritical, "檢核資料"
         Text2.SetFocus
      End If
      
      If bQuery Then QueryData
   Else
      ClearResult
      If Text7 <> "" Then
         QueryData
      Else
         MsgBox "請輸入對照號數 !", vbCritical
         Text7.SetFocus
      End If
   End If
   
End Sub

Private Sub Form_Activate()
   'Added by Morgan 2014/1/14
   If m_strIR01 <> "" And m_Done = False Then
      Option1.Value = True
      Text5.Text = m_RDate
      'Modified by Morgan 2017/10/25
      '先用完整的查一次(Ex.P-117994 對方延期)，若無資料再用前9碼查
'      Text7.Text = m_AppNo
'      m_Retry = False
'      Command1.Value = True
'      If m_Retry = True Then
'         Text7.Text = Left(m_AppNo, 9)
'         Command1.Value = True
'      End If
'      'end 2017/10/25
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   'Added by Morgan 2014/1/14
   ElseIf m_AppNo <> "" And m_Done = False Then
      Option1.Value = True
      Text5.Text = m_RDate
      'Modified by Morgan 2017/10/25
      '先用完整的查一次(Ex.P-117994 對方延期)，若無資料再用前9碼查
      Text7.Text = m_AppNo
      m_Retry = False
      Command1.Value = True
      If m_Retry = True Then
         Text7.Text = Left(m_AppNo, 9)
         Command1.Value = True
      End If
      'end 2017/10/25
      m_Done = True
   End If
   'end 2014/1/14
End Sub

Private Sub Form_Load()
  
   MoveFormToCenter Me
   EnableTextBox Text7, True
   EnableTextBox Text1, False
   EnableTextBox Text2, False
   EnableTextBox Text3, False
   EnableTextBox Text4, False
   Text5.Text = GetTaiwanTodayDate
   ClearResult
   InitialGrdList
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010506_1 = Nothing
End Sub

Private Sub Option1_Click()
   EnableTextBox Text7, True
   EnableTextBox Text1, False
   EnableTextBox Text2, False
   EnableTextBox Text3, False
   EnableTextBox Text4, False
   Text7.SetFocus
End Sub

Private Sub Option2_Click()
   EnableTextBox Text1, True
   EnableTextBox Text2, True
   EnableTextBox Text3, True
   EnableTextBox Text4, True
   EnableTextBox Text7, False
   Text1.SetFocus
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If IsEmptyText(Text1) = False Then
      If Text1.Text <> "P" Then
          MsgBox "只可為P案件", vbInformation
          Text1_GotFocus
          Cancel = True
          Exit Sub
      Else
          Cancel = False
      End If
    End If
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text3_GotFocus()
   InverseTextBox Text3
End Sub

Private Sub Text4_GotFocus()
   InverseTextBox Text4
End Sub

Private Sub Text4_LostFocus()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
End Sub

Private Sub Text5_GotFocus()
   InverseTextBox Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> "" Then
      If ChkDate(Text5) Then
         Text5 = TransDate(Text5, 1) 'Add by Morgan 2009/7/31 改可輸西元年但自動轉民國年
         If Val(Text5) > Val(strSrvDate(2)) Then
            MsgBox "來函收文日不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   
   If Text5 = "" Then
      MsgBox "來函收文日不可空白 !", vbCritical
      Text5.SetFocus
      Exit Function
      
   'Add by Morgan 2009/7/31
   Else
      Text5_Validate Cancel
      If Cancel = True Then
         Text5.SetFocus
         Text5_GotFocus
         Exit Function
      End If
      
   End If
   TxtValidate = True
   
End Function

Private Sub Text7_GotFocus()
   InverseTextBox Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub ClearResult()
   m_CP01 = Empty
   m_CP02 = Empty
   m_CP03 = Empty
   m_CP04 = Empty
   m_CP09 = Empty
   InitialGrdList
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
   grdList.ColWidth(1) = 1600
   grdList.ColAlignment(1) = flexAlignLeftCenter
   grdList.col = 2
   grdList.Text = "對造案件名稱"
   grdList.ColWidth(2) = 1600
   grdList.ColAlignment(2) = flexAlignLeftCenter
   grdList.col = 3
   grdList.Text = "對造名稱"
   grdList.ColWidth(3) = 1600
   grdList.ColAlignment(3) = flexAlignLeftCenter
   grdList.col = 4
   grdList.Text = "案件性質"
   grdList.ColWidth(4) = 1600
   grdList.ColAlignment(4) = flexAlignLeftCenter
   grdList.col = 5
   grdList.Text = "發文日"
   grdList.ColWidth(5) = 1200
   grdList.ColAlignment(5) = flexAlignCenterCenter
   grdList.col = 6
   grdList.Text = "收文號"
   grdList.ColWidth(6) = 0
   grdList.col = 7
   grdList.Text = "PA01"
   grdList.ColWidth(7) = 0
   grdList.col = 8
   grdList.Text = "PA02"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "PA03"
   grdList.ColWidth(9) = 0
   grdList.col = 10
   grdList.Text = "PA04"
   grdList.ColWidth(10) = 0
End Sub

Private Sub QueryData()
   Dim rsTmp As ADODB.Recordset

   ClearResult
   If Option1.Value = True Then
      'Modify By Cheng 2002/04/15
'      strSQL = "SELECT (CP01||'-'||CP02||'-'||CP03||'-'||CP04||'N') as FLD1, " & _
'                  "NVL(CP37,NVL(CP38,NVL(CP39,''))) AS FLD2, " & _
'                  "NVL(CP40,NVL(CP41,NVL(CP42,''))) AS FLD3, " & _
'                  "CP09,CP01,CP02,CP03,CP04,DECODE(PA09,'000',CPM03,CPM04) AS FLD4, " & _
'                  "NVL(CP27-19110000, NULL) AS FLD5,PA09 FROM CASEPROGRESS, PATENT,CASEPROPERTYMAP " & _
'               "WHERE CP36='" & Text7.Text & "' AND " & _
'                     "(SUBSTR(CP09,1,1)='A' OR SUBSTR(CP09,1,1)='B') AND " & _
'                     "CP01='P' AND " & _
'                     "CP01=PA01 AND " & _
'                     "CP02=PA02 AND " & _
'                     "CP03=PA03 AND " & _
'                     "CP04=PA04 AND " & _
'                     "(CP10 = '801' OR CP10 = '803') AND " & _
'                     "PA23 <> '1' AND " & _
'                     "CP01=CPM01(+) AND " & _
'                     "CP10=CPM02(+) "
' 91.09.13 modify by louis (排序)
'      strSQL = "SELECT (CP01||'-'||CP02||'-'||CP03||'-'||CP04||'N') as FLD1, " & _
'                  "NVL(CP37,NVL(CP38,NVL(CP39,''))) AS FLD2, " & _
'                  "NVL(CP40,NVL(CP41,NVL(CP42,''))) AS FLD3, " & _
'                  "CP09,CP01,CP02,CP03,CP04,DECODE(PA09,'000',CPM03,CPM04) AS FLD4, " & _
'                  "NVL(CP27-19110000, NULL) AS FLD5,PA09 FROM CASEPROGRESS, PATENT,CASEPROPERTYMAP " & _
'               "WHERE CP36='" & Text7.Text & "' AND " & _
'                     "( CP09<'C' ) AND CP27 IS NOT NULL AND " & _
'                     "CP01='P' AND " & _
'                     "CP01=PA01 AND " & _
'                     "CP02=PA02 AND " & _
'                     "CP03=PA03 AND " & _
'                     "CP04=PA04 AND " & _
'                     "(CP10 = '801' OR CP10 = '803') AND " & _
'                     "PA23 <> '1' AND " & _
'                     "CP01=CPM01(+) AND " & _
'                     "CP10=CPM02(+) "
      'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
        'Add by Lydia 2014/10/31 先判斷外專程序人員權限。
        If FMP2open = True And FMP2openSQL <> "" Then
           strExc(0) = "select cp01,cp02,cp03,cp04,cp05,cp09,cp10 FROM CASEPROGRESS f0 " & _
                  "WHERE CP36='" & Text7.Text & "' " & FMP2openSQL
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, "CHANGE_SQL", strExc(0)) = False Then
            Me.Text7.SetFocus
            Text7_GotFocus
            Exit Sub
           End If
        End If

        '設別名f0,+FMP2openSQL
      strSql = "SELECT (CP01||'-'||CP02||'-'||CP03||'-'||CP04||'N') as FLD1, " & _
                  "NVL(CP37,NVL(CP38,NVL(CP39,''))) AS FLD2, " & _
                  "NVL(CP40,NVL(CP41,NVL(CP42,''))) AS FLD3, " & _
                  "CP09,CP01,CP02,CP03,CP04,DECODE(PA09,'000',CPM03,CPM04) AS FLD4, " & _
                  "NVL(CP27-19110000, NULL) AS FLD5,PA09,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
                  "FROM CASEPROGRESS f0, PATENT,CASEPROPERTYMAP " & _
               "WHERE CP36='" & Text7.Text & "' AND " & _
                     "( CP09<'C' ) AND CP27 IS NOT NULL AND " & _
                     "CP01='P' AND " & _
                     "CP01=PA01 AND " & _
                     "CP02=PA02 AND " & _
                     "CP03=PA03 AND " & _
                     "CP04=PA04 " & _
                     "AND (CP10 = '801' OR CP10 = '803') AND " & _
                     "PA23 <> '1' AND " & _
                     "CP01=CPM01(+) AND " & _
                     "CP10=CPM02(+) ORDER BY SORTFIELD DESC "
   Else
'      strSQL = "select (cp01||'-'||cp02||'-'||cp03||'-'||cp04||'N') as FLD1, " & _
'                  "nvl(cp37,nvl(cp38,nvl(cp39,''))) as FLD2," & _
'                  "nvl(cp40,nvl(cp41,nvl(cp42,''))) as FLD3," & _
'                  "CP09,CP01,CP02,CP03,CP04,decode(pa09,'000',cpm03,cpm04) as FLD4," & _
'                  "NVL(cp27-19110000,NULL) as FLD5,PA09 From caseprogress, patent,casepropertymap " & _
'               "where CP01 = '" & Text1 & "' AND " & _
'                     "CP02 = '" & Text2 & "' AND " & _
'                     "CP03 = '" & Text3 & "' AND " & _
'                     "CP04 = '" & Text4 & "' AND " & _
'                     "(CP10 = '801' OR CP10 = '803') AND " & _
'                     "(substr(cp09,1,1)='A' or substr(cp09,1,1)='B') and " & _
'                     "cp01=pa01 and " & _
'                     "cp02=pa02 and " & _
'                     "cp03=pa03 and " & _
'                     "cp04=pa04 and " & _
'                     "pa23 <> '1' and " & _
'                     "CP01 = CPM01(+) and " & _
'                     "CP10 = CPM02(+)"
' 91.09.13 modify by louis
'      strSQL = "select (cp01||'-'||cp02||'-'||cp03||'-'||cp04||'N') as FLD1, " & _
'                  "nvl(cp37,nvl(cp38,nvl(cp39,''))) as FLD2," & _
'                  "nvl(cp40,nvl(cp41,nvl(cp42,''))) as FLD3," & _
'                  "CP09,CP01,CP02,CP03,CP04,decode(pa09,'000',cpm03,cpm04) as FLD4," & _
'                  "NVL(cp27-19110000,NULL) as FLD5,PA09 From caseprogress, patent,casepropertymap " & _
'               "where CP01 = '" & Text1 & "' AND " & _
'                     "CP02 = '" & Text2 & "' AND " & _
'                     "CP03 = '" & Text3 & "' AND " & _
'                     "CP04 = '" & Text4 & "' AND " & _
'                     "(CP10 = '801' OR CP10 = '803') AND " & _
'                     "( cp09<'C' ) and " & _
'                     "cp01=pa01 and " & _
'                     "cp02=pa02 and " & _
'                     "cp03=pa03 and " & _
'                     "cp04=pa04 and " & _
'                     "pa23 <> '1' and " & _
'                     "CP01 = CPM01(+) and " & _
'                     "CP10 = CPM02(+)"
                '設別名f0,+FMP2openSQL
      strSql = "SELECT (cp01||'-'||cp02||'-'||cp03||'-'||cp04||'N') as FLD1, " & _
                  "nvl(cp37,nvl(cp38,nvl(cp39,''))) as FLD2," & _
                  "nvl(cp40,nvl(cp41,nvl(cp42,''))) as FLD3," & _
                  "CP09,CP01,CP02,CP03,CP04,decode(pa09,'000',cpm03,cpm04) as FLD4," & _
                  "NVL(cp27-19110000,NULL) as FLD5,PA09,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
                  "From caseprogress f0, patent,casepropertymap " & _
               "WHERE CP01 = '" & Text1 & "' AND " & _
                     "CP02 = '" & Text2 & "' AND " & _
                     "CP03 = '" & Text3 & "' AND " & _
                     "CP04 = '" & Text4 & "' AND " & _
                     "(CP10 = '801' OR CP10 = '803') AND " & _
                     "( cp09<'C' ) and " & _
                     "cp01=pa01 and " & _
                     "cp02=pa02 and " & _
                     "cp03=pa03 and " & _
                     "cp04=pa04 " & _
                     "and pa23 <> '1' and " & _
                     "CP01 = CPM01(+) and " & _
                     "CP10 = CPM02(+) ORDER BY SORTFIELD DESC "
   End If
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ListData rsTmp
   Else
      'Added by Morgan 2014/6/12
      If m_AppNo <> "" And m_Retry = False Then
         m_Retry = True
      Else
      'end 2014/6/12
         MsgBox "沒有符合條件的資料", vbOKOnly + vbInformation, "查詢資料"
      End If 'Added by Morgan 2014/6/12
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   If grdList.Rows = 2 Then
      grdList_SetSelection 1
      DisplayNextForm
   End If
End Sub

Private Sub ListData(ByRef rsTmp As ADODB.Recordset)
   Dim nRow As Integer
   
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      grdList.Rows = grdList.Rows + 1
      nRow = grdList.Rows - 1
      ' 本所案號
      If IsNull(rsTmp.Fields("FLD1")) = False Then
         grdList.TextMatrix(nRow, 1) = rsTmp.Fields("FLD1")
      End If
      ' 對造案件名稱
      If IsNull(rsTmp.Fields("FLD2")) = False Then
         grdList.TextMatrix(nRow, 2) = rsTmp.Fields("FLD2")
      End If
      ' 對造名稱
      If IsNull(rsTmp.Fields("FLD3")) = False Then
         grdList.TextMatrix(nRow, 3) = rsTmp.Fields("FLD3")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("FLD4")) = False Then
         grdList.TextMatrix(nRow, 4) = rsTmp.Fields("FLD4")
      End If
      ' 發文日
      If IsNull(rsTmp.Fields("FLD5")) = False Then
         grdList.TextMatrix(nRow, 5) = rsTmp.Fields("FLD5")
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         grdList.TextMatrix(nRow, 6) = rsTmp.Fields("CP09")
      End If
      ' 本所案號
      If IsNull(rsTmp.Fields("CP01")) = False Then
         grdList.TextMatrix(nRow, 7) = rsTmp.Fields("CP01")
      End If
      If IsNull(rsTmp.Fields("CP02")) = False Then
         grdList.TextMatrix(nRow, 8) = rsTmp.Fields("CP02")
      End If
      If IsNull(rsTmp.Fields("CP03")) = False Then
         grdList.TextMatrix(nRow, 9) = rsTmp.Fields("CP03")
      End If
      If IsNull(rsTmp.Fields("CP04")) = False Then
         grdList.TextMatrix(nRow, 10) = rsTmp.Fields("CP04")
      End If
      rsTmp.MoveNext
   Loop
   'Added by Lydia 2023/10/17
   If grdList.Rows >= 2 Then
      grdList.FixedRows = 1
   End If
   'end 2023/10/17
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

Private Sub grdList_Click()
   grdList_SelChange
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      m_CP01 = grdList.TextMatrix(grdList.row, 7)
      m_CP02 = grdList.TextMatrix(grdList.row, 8)
      m_CP03 = grdList.TextMatrix(grdList.row, 9)
      m_CP04 = grdList.TextMatrix(grdList.row, 10)
      m_CP09 = grdList.TextMatrix(grdList.row, 6)
      grdList_ShowSelection
   End If
End Sub

' 檢查來函記錄檔
Private Function PromptIfTaiwanNoResult() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strNation As String
   Dim bPrompt As Boolean
   
   bPrompt = False
   PromptIfTaiwanNoResult = True
   strNation = "111"
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & m_CP01 & "' AND " & _
                  "PA02 = '" & m_CP02 & "' AND " & _
                  "PA03 = '" & m_CP03 & "' AND " & _
                  "PA04 = '" & m_CP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("PA09")) = False Then
         strNation = rsTmp.Fields("PA09")
      End If
   End If
   rsTmp.Close

   If strNation < "010" Then
      If m_DocNo = "" Then 'Added by Morgan 2014/5/5 排除無期限電子公文
         strSql = "SELECT * FROM MailRec " & _
                  "WHERE MR12 = '" & m_CP01 & "' AND " & _
                        "MR13 = '" & m_CP02 & "' AND " & _
                        "MR14 = '" & m_CP03 & "' AND " & _
                        "MR15 = '" & m_CP04 & "' AND " & _
                        "MR02 = " & DBDATE(Text5) & " "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenDynamic
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("MR16")) = False Then
               If rsTmp.Fields("MR16") <> "0" Then
                  bPrompt = True
               End If
            End If
         Else
            bPrompt = True
         End If
         rsTmp.Close
      
         If bPrompt = True Then
            strTit = "資料檢核"
            strMsg = "與櫃台之來函收文記錄不符, 請確認"
            nResponse = MsgBox(strMsg, vbOKCancel, strTit)
            If nResponse = vbCancel Then
               PromptIfTaiwanNoResult = False
            End If
         End If
      End If 'Added by Morgan 2014/5/5
   End If
   
   Set rsTmp = Nothing
End Function

Private Sub DisplayNextForm()
   Dim strSql As String
   
   If TxtValidate = False Then Exit Sub


   If PromptIfTaiwanNoResult() = True Then
      'Add By Sindy 2017/12/27
      If m_strIR01 <> "" Then
         If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> m_CP01 & m_CP02 & m_CP03 & m_CP04 Then
            MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
            Exit Sub
         End If
      End If
      '2017/12/27 END
      
      frm04010506_2.SetData m_CP09
      Me.Hide
      'Added by Morgan 2014/1/14
      frm04010506_2.m_AppNo = m_AppNo
      frm04010506_2.m_DocNo = m_DocNo
      frm04010506_2.m_DocWord = m_DocWord 'Added by Morgan 2015/9/8
      'end 2014/1/14
      'Add By Sindy 2016/10/5
      frm04010506_2.m_strIR01 = m_strIR01
      frm04010506_2.m_strIR02 = m_strIR02
      frm04010506_2.m_strIR03 = m_strIR03
      frm04010506_2.m_strIR04 = m_strIR04
      '2016/10/5 END
      frm04010506_2.Show
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   If IsEmptyText(Text5) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入來函收文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   If grdList.Rows < 2 Then
      strTit = "資料檢核"
      strMsg = "請先選取資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   If grdList.row < 2 Then
      strTit = "資料檢核"
      strMsg = "請先選取資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   If IsEmptyText(m_CP09) = True Then
      strTit = "資料檢核"
      strMsg = "請先選取資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

'Add By Cheng 2003/01/03
Public Sub Clear()
    '若選擇對造號數
    If Me.Option1.Value Then
        TextInverse Me.Text7
    '若選擇本所案號
    Else
        TextInverse Me.Text2
    End If
    Me.grdList.Rows = 1
    Me.Command1.Default = True
End Sub
