VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160012_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "新增打卡資料"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdDel 
      Caption         =   "刪除"
      Default         =   -1  'True
      Height          =   375
      Left            =   5490
      TabIndex        =   19
      Top             =   2820
      Width           =   975
   End
   Begin VB.ComboBox cboB1403 
      Height          =   260
      ItemData        =   "frm160012_2.frx":0000
      Left            =   2070
      List            =   "frm160012_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   1500
      Width           =   1125
   End
   Begin VB.TextBox textB1404 
      Height          =   270
      Left            =   2070
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1860
      Width           =   1035
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "打卡明細"
      Height          =   375
      Left            =   5580
      TabIndex        =   6
      Top             =   540
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox textB1402 
      Height          =   270
      Left            =   2070
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1170
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "新增"
      Height          =   375
      Left            =   5580
      TabIndex        =   4
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面"
      Height          =   375
      Left            =   7290
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   375
      Left            =   6660
      TabIndex        =   5
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox textB1401 
      Height          =   270
      Left            =   2070
      MaxLength       =   6
      TabIndex        =   0
      Top             =   840
      Width           =   1035
   End
   Begin VB.CommandButton cmdABS 
      Caption         =   "請假、出差、加班資料"
      Height          =   375
      Left            =   6660
      TabIndex        =   9
      Top             =   540
      Visible         =   0   'False
      Width           =   2145
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   1155
      Left            =   6810
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   2028
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      ScrollTrack     =   -1  'True
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
      _Band(0).Cols   =   14
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   2600
      Left            =   2070
      TabIndex        =   16
      Top             =   2670
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   4568
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|刷卡日期|刷卡時間|人事補登"
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
      _Band(0).Cols   =   4
   End
   Begin MSForms.Label Label12 
      Height          =   230
      Left            =   3180
      TabIndex        =   20
      Top             =   870
      Width           =   1400
      BackColor       =   12632256
      VariousPropertyBits=   27
      Size            =   "2461;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "（有異常資料一併刪除。若不刪，請把時段清空白）"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   200
      Left            =   3270
      TabIndex        =   18
      Top             =   1560
      Width           =   4830
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "當日打卡明細："
      Height          =   180
      Left            =   780
      TabIndex        =   17
      Top             =   2730
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "時　　段："
      Height          =   180
      Left            =   1110
      TabIndex        =   15
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label13 
      Caption         =   "備註：會發E-Mail通知當事人。"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   210
      TabIndex        =   14
      Top             =   5430
      Width           =   4125
   End
   Begin VB.Label Label2 
      Caption         =   "例.HHMMSS（ex.93000 : 9點半）"
      Height          =   270
      Left            =   3180
      TabIndex        =   13
      Top             =   1860
      Width           =   2570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "打卡時間："
      Height          =   180
      Left            =   1110
      TabIndex        =   12
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      Height          =   180
      Left            =   1110
      TabIndex        =   11
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   1
      Left            =   1110
      TabIndex        =   10
      Top             =   870
      Width           =   930
   End
End
Attribute VB_Name = "frm160012_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/17 Form2.0已修改
'Create By Sindy 2013/6/21
Option Explicit

' 變數宣告區
Public m_B1401 As String '員工代號
Public m_B1402 As String '日期
Public m_B1403 As String '打卡類別
Public bolClose As Boolean
Dim m_B1405 As String
'Dim m_PR03 As String
Dim intB14Cnt As Integer


Private Sub cmdBack_Click()
   Unload Me
   frm160012.Show
End Sub

'Add By Sindy 2020/11/24
Private Sub cmdDel_Click()
Dim i As Integer
Dim strPR01 As String, strPR02 As String, strPR03 As String

On Error GoTo ErrHand
   
   strPR01 = "": strPR02 = "": strPR03 = ""
   For i = 1 To grdList.Rows - 1
      If grdList.TextMatrix(i, 0) = "V" And grdList.TextMatrix(i, 1) <> "" Then
         strPR01 = grdList.TextMatrix(i, 4)
         strPR02 = grdList.TextMatrix(i, 5)
         strPR03 = grdList.TextMatrix(i, 6)
         Exit For
      End If
   Next i
   
   If strPR01 <> "" And strPR02 <> "" And strPR03 <> "" Then
      If MsgBox("確定要刪除這筆刷卡記錄嗎？", vbInformation + vbYesNo + vbDefaultButton1, "刪除資料") = vbNo Then
         Exit Sub
      End If
   Else
      Exit Sub
   End If
   
   cnnConnection.BeginTrans
   
   '刪除刷卡資料
   strSql = "DELETE FROM pollrecord WHERE pr01=" & strPR01 & " and pr02=" & strPR02 & " and pr03='" & strPR03 & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   
   PollRecordQueryData
   MsgBox "刪除成功 !!!"
   
   Exit Sub
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "資料刪除失敗！" & vbCrLf & Err.Description
End Sub

'打卡明細
Private Sub cmdDetail_Click()
   If textB1401 = "" Then
      MsgBox "員工代號不可空白 !!!"
      textB1401.SetFocus
      Exit Sub
   End If
   If textB1402 = "" Then
      MsgBox "日期不可空白 !!!"
      textB1402.SetFocus
      Exit Sub
   End If
   
   'If textB1401 = "" Or textB1402 = "" Then Exit Sub
   Call frm180303_1.SetParent(Me)
   frm180303_1.m_B1401 = textB1401
   frm180303_1.m_B1402 = DBDATE(textB1402)
   If frm180303_1.QueryData = True Then
      frm180303_1.Show vbModal '強制回應表單
   Else
      Unload frm180303_1
   End If
End Sub

' 當日打卡明細初始化列表
Private Sub InitialGridList()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2020/11/24 + "V", "pr01", "pr02", "pr03"
   '                        0    1           2           3           4       5       6
   arrGridHeadText = Array("V", "刷卡日期", "刷卡時間", "人事補登", "pr01", "pr02", "pr03")
   arrGridHeadWidth = Array(300, 900, 900, 800, 0, 0, 0)
   grdList.Visible = False
   grdList.Cols = UBound(arrGridHeadText) + 1
   grdList.Rows = 2
   For iRow = 0 To grdList.Cols - 1
      grdList.row = 0
      grdList.col = iRow
      grdList.Text = arrGridHeadText(iRow)
      grdList.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grdList.CellAlignment = flexAlignCenterCenter
   Next
   grdList.Visible = True
End Sub

' 當日打卡明細查詢
Public Function PollRecordQueryData() As Boolean
   Dim stSQL As String
   
   PollRecordQueryData = False
   
   Screen.MousePointer = vbHourglass
   Me.grdList.MousePointer = flexHourglass
   InitialGridList
   
   '檢查是否有異常資料
   cboB1403.ListIndex = 0: intB14Cnt = 0
   Label5.Visible = False
   cboB1403.Visible = False
   Label6.Visible = False
   stSQL = "select * from abs014 where b1401='" & textB1401 & "' and b1402=" & DBDATE(textB1402) & " and b1411 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      If RsTemp.RecordCount > 0 Then
         Label5.Visible = True
         cboB1403.Visible = True
         Label6.Visible = True
         If RsTemp.Fields("b1403") = "A" Then
            cboB1403.ListIndex = 1
         ElseIf RsTemp.Fields("b1403") = "P" Then
            cboB1403.ListIndex = 2
         End If
         intB14Cnt = RsTemp.RecordCount
      End If
   End If
   
   '當日打卡明細資料
   stSQL = "select '' as V,sqldatet(pr01) as 刷卡日期,sqltime6(pr02) as 刷卡時間,decode(pr08,999,'Y','') as 人事補登,pr01,pr02,pr03"
   stSQL = stSQL & " from staff,staffcarddata,pollrecord where scd01(+)=st01 and pr03(+)=scd02 and pr01>0" & _
                    " and st01='" & textB1401 & "' and pr01=" & DBDATE(textB1402)
   stSQL = stSQL & " order by pr02 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      Set grdList.Recordset = RsTemp
      grdList.row = 1
      PollRecordQueryData = True
'   Else
'      ShowNoData
'      Me.grdList.MousePointer = flexDefault
'      Screen.MousePointer = vbDefault
'      Exit Function
   End If
   Me.grdList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Function

Private Sub cmdExit_Click()
   Unload Me
   'Unload frm160012
   If frm160012.m_QueryType = 3 Then
      frm160012.cmdQuery3_Click
   ElseIf frm160012.m_QueryType = 2 Then
      frm160012.cmdQuery2_Click
   Else
      frm160012.cmdQuery_Click
   End If
   frm160012.Show
End Sub

'請假、出差、加班資料
Private Sub cmdABS_Click()
Dim rsTmp As New ADODB.Recordset
   
   GRD2.Clear
   SetGrd2
   If PUB_QueryData_ABS(textB1401, textB1402, rsTmp) = True Then
      Set GRD2.Recordset = rsTmp
      Call PubShowNextData
      Exit Sub
   End If
End Sub

Private Sub SetGrd2()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("V", "員工代號", "表單編號", "TableID", "SA02", "SA03")
   arrGridHeadWidth = Array(800, 800, 800, 800, 800, 800)
   'grd2.Visible = False
   GRD2.Cols = UBound(arrGridHeadText) + 1
   GRD2.Rows = 2
   For iRow = 0 To GRD2.Cols - 1
      GRD2.row = 0
      GRD2.col = iRow
      GRD2.Text = arrGridHeadText(iRow)
      GRD2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD2.CellAlignment = flexAlignCenterCenter
   Next
   'grd2.Visible = True
End Sub

'查詢出缺勤明細資料
Public Sub PubShowNextData()
Dim i As Integer
   
   Me.Enabled = False
   For i = 1 To GRD2.Rows - 1
      GRD2.col = 0
      GRD2.row = i
      If Trim(GRD2.Text) = "V" Then
         GRD2.Text = ""
         GRD2.col = 2 '表單編號
         Screen.MousePointer = vbHourglass
         Me.Hide
         Call frm180301_03.SetParent(Me)
         If GRD2.TextMatrix(i, 3) = "1" Then '出缺勤
            frm180301_03.txtB1001 = Pub_RplStr(GRD2.Text)
            frm180301_03.QueryData
         Else
            frm180301_03.txtB1003 = Pub_RplStr(GRD2.TextMatrix(i, 1))
            frm180301_03.m_SA02 = Pub_RplStr(GRD2.TextMatrix(i, 4))
            frm180301_03.m_SA03 = Pub_RplStr(GRD2.TextMatrix(i, 5))
            If GRD2.TextMatrix(i, 3) = "2" Then '請假
               frm180301_03.QueryData_2
            ElseIf GRD2.TextMatrix(i, 3) = "3" Then '加班
               frm180301_03.QueryData_3
            ElseIf GRD2.TextMatrix(i, 3) = "4" Then '出差
               frm180301_03.QueryData_4
            End If
         End If
         frm180301_03.Show
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         Exit Sub
      End If
   Next i
   Me.Enabled = True
End Sub

'確定
Private Sub cmdok_Click()
Dim strSubject As String, strContent As String, strTo As String
   
   If CheckDataValid = False Then Exit Sub
   If txtValidate = False Then Exit Sub
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   'Add By Sindy 2013/8/2
   '刪除異常資料
   If Trim(cboB1403) <> "" Then
      strSql = "delete from abs014 where b1401='" & textB1401 & "' and b1402=" & DBDATE(textB1402) & " and b1403='" & Left(Trim(cboB1403), 1) & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   
   '新增刷卡資料
   'Modify By Sindy 2022/5/2 m_PR03 ==> textB1401
   strSql = "insert into pollrecord(pr01,pr02,pr03,pr08)" & _
            " VALUES(" & CNULL(DBDATE(textB1402)) & "," & CNULL(textB1404) & "," & _
            CNULL(textB1401) & ",999)"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   MsgBox "新增成功 !!!"
   
   strTo = PUB_GetST59(textB1401)
   If IsNull(strTo) Or strTo = "" Then
      strTo = textB1401
   End If
   strSubject = Label12 & " " & ChangeWStringToTDateString(DBDATE(textB1402)) & " " & "人事處補輸打卡資料通知！"
   strContent = "員工姓名：" & Label12 & vbCrLf
   strContent = strContent & "日　　期：" & ChangeTStringToTDateString(textB1402) & vbCrLf
   strContent = strContent & "打卡時間：" & Format(textB1404, "##:##:##") & vbCrLf
   'Modify By Sindy 2017/7/28 劉經理說:人事處補輸同仁打卡時間, 請取消系統通知當事人功能(系統不公告)
   'PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True
   
   'Call ClearData
   'textB1404.Text = ""
   '清空欄位值
   textB1404 = ""
   Call PollRecordQueryData
   Exit Sub
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "資料新增失敗！" & vbCrLf & Err.Description
End Sub

'檢查 pollrecord 資料是否存在
Private Function IsRecordExist() As Boolean
Dim Rs As ADODB.Recordset
   
   IsRecordExist = False
   
'   '讀取員工的卡號
'   m_PR03 = ""
'   'Modify By Sindy 2022/5/2 + order by scd07 desc
'   strExc(0) = "select scd01,scd02 from staffcarddata where scd01='" & textB1401 & "' order by scd07 desc"
'   intI = 1
'   Set rs = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If rs.RecordCount > 0 Then
'         m_PR03 = rs.Fields("scd02")
'      End If
'   End If
   
   '檢查資料是否已存在
   strExc(0) = "select pollrecord.* from pollrecord,staffcarddata where scd01='" & textB1401 & "' and scd02=pr03 and pr01=" & DBDATE(textB1402) & " and pr02=" & Val(textB1404)
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Rs.RecordCount > 0 Then
         IsRecordExist = True
         Rs.Close
         Set Rs = Nothing
         Exit Function
      End If
   End If
   
   'Add By Sindy 2022/6/6 未建指紋的狀況下
   strSql = "select * from pollrecord" & _
            " where pr01=" & DBDATE(textB1402) & " And pr02=" & Val(textB1404) & " and pr03='" & textB1401 & "'"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Rs.RecordCount > 0 Then
         IsRecordExist = True
         Rs.Close
         Set Rs = Nothing
         Exit Function
      End If
   End If
   '2022/6/6 END
   
   Set Rs = Nothing
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   Call ClearData
   Label5.Visible = False
   cboB1403.Visible = False
   Label6.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160012_2 = Nothing
End Sub

'清除欄位值
Private Sub ClearData()
   textB1401.Text = "": Label12.Caption = ""
   textB1402.Text = ""
   textB1404.Text = ""
'   m_PR03 = ""
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   CheckDataValid = False
   
   If textB1401 = "" Then
      MsgBox "員工代號不可空白 !!!"
      textB1401.SetFocus
      Exit Function
   End If
   
   If textB1402 = "" Then
      MsgBox "日期不可空白 !!!"
      textB1402.SetFocus
      Exit Function
   End If
   
   If textB1404 = "" Then
      MsgBox "打卡時間不可空白 !!!"
      textB1404.SetFocus
      Exit Function
   End If
   
   '檢查資料是否已存在
   If IsRecordExist() = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textB1404.SetFocus
      Exit Function
   End If
   
'   If m_PR03 = "" Then
'      MsgBox "讀取不到員工卡號，請洽電腦中心 !!!"
'      Exit Function
'   End If
   
   '檢查是否有異常資料
   If Trim(cboB1403) <> "" Then
      strSql = "select * from abs014 where b1401='" & textB1401 & "' and b1402=" & DBDATE(textB1402) & " and b1403='" & Left(Trim(cboB1403), 1) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount <> 1 Then
            MsgBox "無異常資料可刪除，請重新確認 !!!"
            Exit Function
         End If
      Else
         MsgBox "無異常資料可刪除，請重新確認 !!!"
         Exit Function
      End If
   End If
   
   'Add By Sindy 2013/8/16
   If intB14Cnt = 2 And cboB1403.ListIndex = 0 Then
      MsgBox "請選擇一筆要刪除的異常時段 !!!"
      Exit Function
   End If
   
   CheckDataValid = True
End Function

Private Function txtValidate() As Boolean
Dim Cancel As Boolean

   txtValidate = False
   
'   If Me.textB1401.Enabled = True Then
'      Cancel = False
'      textB1401_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
'   If Me.textB1402.Enabled = True Then
'      Cancel = False
'      textB1402_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
   If Me.textB1404.Enabled = True Then
      Cancel = False
      textB1404_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   txtValidate = True
End Function

'Add By Sindy 2020/11/24
Private Sub grdList_Click()
Dim j As Integer, i As Integer

grdList.Visible = False
If grdList.MouseRow <> 0 Then
   '先清空全部已選取的資料列
   For j = 1 To grdList.Rows - 1
      If grdList.TextMatrix(j, 1) <> "" Then
         If grdList.TextMatrix(j, 0) = "V" Then
            grdList.col = 0
            grdList.row = j
            grdList.Text = ""
            For i = 0 To grdList.Cols - 1
               grdList.col = i
               grdList.CellBackColor = QBColor(15)
            Next i
         End If
      End If
   Next j
   '該筆資料列變成已選取
   grdList.col = 0
   grdList.row = grdList.MouseRow
   If grdList.TextMatrix(grdList.MouseRow, 1) <> "" Then
'      If grdList.Text = "V" Then
'         grdList.Text = ""
'         For i = 0 To grdList.Cols - 1
'            grdList.col = i
'            grdList.CellBackColor = QBColor(15)
'         Next i
'      Else
         grdList.Text = "V"
         For i = 0 To grdList.Cols - 1
            grdList.col = i
            grdList.CellBackColor = &HFFC0C0
         Next i
'      End If
   End If
End If
grdList.Visible = True
End Sub
'Private Sub grdList_Click()
'Dim i As Integer
'
'grdList.Visible = False
'If grdList.MouseRow <> 0 Then
'   grdList.col = 0
'   grdList.row = grdList.MouseRow
'   If grdList.TextMatrix(grdList.MouseRow, 1) <> "" Then
'      If grdList.Text = "V" Then
'         grdList.Text = ""
'         For i = 0 To grdList.Cols - 1
'            grdList.col = i
'            grdList.CellBackColor = QBColor(15)
'         Next i
'      Else
'         grdList.Text = "V"
'         For i = 0 To grdList.Cols - 1
'            grdList.col = i
'            grdList.CellBackColor = &HFFC0C0
'         Next i
'      End If
'   End If
'End If
'grdList.Visible = True
'End Sub

Private Sub textB1401_GotFocus()
   InverseTextBox textB1401
End Sub

Private Sub textB1401_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textB1401_Validate(Cancel As Boolean)
Dim Rs As New ADODB.Recordset
   
   If textB1401.Text = "" Then Label12 = ""
   
   If textB1401 <> "" Then
      ' 檢查員工編號規則
      If ChkStaffST04(textB1401) Then
         Call textB1401_GotFocus
         Cancel = True
         Exit Sub
      End If
      Label12 = GetStaffName(textB1401, True)
      If Label12 = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         Call textB1401_GotFocus
         Cancel = True
         Exit Sub
      End If
      If textB1402 <> "" Then
         Call PollRecordQueryData
      End If
   End If
End Sub

Private Sub textB1402_GotFocus()
   InverseTextBox textB1402
End Sub

Private Sub textB1402_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textB1402_Validate(Cancel As Boolean)
   If textB1402 = "" Then Exit Sub
   If textB1402 <> "" Then
      If ChkDate(textB1402) = False Then
         Call textB1402_GotFocus
         Cancel = True
         Exit Sub
      End If
      '日期檢查不可大於系統日
      If Val(textB1402) > Val(strSrvDate(2)) Then
         MsgBox "日期不可大於系統日！"
         Call textB1402_GotFocus
         Cancel = True
         Exit Sub
      End If
      '只可輸當月及前一個月的日期
      'Modify By Sindy 2017/12/13 可以輸入前六個月的資料
'      If Left(DBDATE(textB1402), 6) <> Left(strSrvDate(1), 6) And _
'         Left(DBDATE(textB1402), 6) <> Left(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##"))), 6) Then
      If Left(DBDATE(textB1402), 6) <> Left(strSrvDate(1), 6) And _
         Left(DBDATE(textB1402), 6) < Left(DBDATE(DateAdd("m", -6, Format(strSrvDate(1), "####/##/##"))), 6) Then
         'MsgBox "只可輸當月及前一個月的日期！"
         MsgBox "只可輸當月及前六個月的日期！"
         Call textB1402_GotFocus
         Cancel = True
         Exit Sub
      End If
      '必須為工作天
      If ChkWorkDay(DBDATE(textB1402)) = False Then
         MsgBox "日期必須為工作天！"
         Call textB1402_GotFocus
         Cancel = True
         Exit Sub
      End If
      If textB1401 <> "" Then
         Call PollRecordQueryData
      End If
   End If
End Sub

Private Sub textB1404_GotFocus()
   InverseTextBox textB1404
End Sub

Private Sub textB1404_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textB1404_Validate(Cancel As Boolean)
   If textB1404 = "" Then Exit Sub
   If textB1404 <> "" Then
      If Len(textB1404) > 6 Then
         MsgBox "時間輸入錯誤 !!!"
         Call textB1404_GotFocus
         Cancel = True
         Exit Sub
      End If
      If Len(textB1404) < 5 Then
         MsgBox "打卡時間必須輸入至秒數 !!!"
         Call textB1404_GotFocus
         Cancel = True
         Exit Sub
      End If
      'Modify By Sindy 2021/6/8
'      If Not (Right(Trim(textB1404), 2) >= 0 And Right(Trim(textB1404), 2) <= 59) Then
'         MsgBox "秒數必須介於 00 ~ 59 !!!"
'         Call textB1404_GotFocus
'         Cancel = True
'         Exit Sub
'      End If
      If IsDate(Format(textB1404.Text, "00:00:00")) = False Then
         MsgBox "非時間時態 !!!"
         Call textB1404_GotFocus
         Cancel = True
         Exit Sub
      End If
      '2021/6/8 END
   End If
End Sub
