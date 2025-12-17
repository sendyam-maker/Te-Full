VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm160018_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "下班逾30分鐘原因確認 - 明細資料"
   ClientHeight    =   5200
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   7990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5200
   ScaleWidth      =   7990
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改"
      Default         =   -1  'True
      Height          =   345
      Left            =   2970
      TabIndex        =   20
      Top             =   180
      Width           =   795
   End
   Begin VB.CommandButton cmdQueryNext 
      Caption         =   "查詢下一筆(&N)"
      Height          =   345
      Left            =   5130
      TabIndex        =   19
      Top             =   180
      Width           =   1365
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認"
      Height          =   345
      Left            =   4260
      TabIndex        =   18
      Top             =   180
      Width           =   795
   End
   Begin VB.ComboBox cboWorkTime 
      Height          =   260
      ItemData        =   "frm160018_1.frx":0000
      Left            =   4980
      List            =   "frm160018_1.frx":000D
      TabIndex        =   15
      Text            =   "cboWorkTime"
      Top             =   1530
      Width           =   1550
   End
   Begin VB.TextBox textB1505 
      Height          =   810
      Left            =   1560
      MaxLength       =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   7
      Top             =   2250
      Width           =   5900
   End
   Begin VB.TextBox textB1501 
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   6
      Top             =   810
      Width           =   735
   End
   Begin VB.TextBox textB1502 
      Height          =   270
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1170
      Width           =   1010
   End
   Begin VB.ComboBox Combo1 
      Height          =   260
      ItemData        =   "frm160018_1.frx":0038
      Left            =   1560
      List            =   "frm160018_1.frx":003A
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1530
      Width           =   2110
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   270
      Left            =   4980
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1140
      Width           =   1010
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   270
      Left            =   6120
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1140
      Width           =   1010
   End
   Begin VB.CommandButton cmdOvertime 
      Caption         =   "加班單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   290
      Left            =   1560
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   1890
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   345
      Left            =   6570
      TabIndex        =   0
      Top             =   180
      Width           =   795
   End
   Begin MSForms.Label Label23 
      Height          =   200
      Left            =   120
      TabIndex        =   17
      Top             =   4920
      Width           =   7700
      VariousPropertyBits=   27
      Caption         =   "CREATE :                                                    UPDATE : "
      Size            =   "13582;353"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "上班時段："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   5
      Left            =   4050
      TabIndex        =   16
      Top             =   1550
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   630
      TabIndex        =   14
      Top             =   860
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      Height          =   180
      Left            =   630
      TabIndex        =   13
      Top             =   1230
      Width           =   900
   End
   Begin MSForms.Label Label12 
      Height          =   230
      Left            =   2340
      TabIndex        =   12
      Top             =   840
      Width           =   1400
      BackColor       =   12632256
      VariousPropertyBits=   27
      Size            =   "2461;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "備　　註："
      Height          =   180
      Index           =   17
      Left            =   630
      TabIndex        =   11
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "打卡起迄時間："
      Height          =   180
      Left            =   3690
      TabIndex        =   10
      Top             =   1190
      Width           =   1260
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "逾時原因："
      Height          =   180
      Left            =   630
      TabIndex        =   9
      Top             =   1590
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   5730
      X2              =   6420
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label5 
      Caption         =   $"frm160018_1.frx":003C
      ForeColor       =   &H000000FF&
      Height          =   700
      Left            =   630
      TabIndex        =   8
      Top             =   3540
      Width           =   6520
   End
End
Attribute VB_Name = "frm160018_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2025/10/29
Option Explicit

' 變數宣告區
Public m_B1501 As String '員工代號
Public m_B1502 As String '日期
Dim m_EditMode As Integer
Dim m_B1001 As String, m_SO02 As String, m_So03 As String


Private Sub cboWorkTime_Click()
   Call ChkWorkTime
End Sub
Private Function ChkWorkTime() As Boolean
   ChkWorkTime = False
   '檢查是否有在正常上班時段
   If cboWorkTime.Text <> "" Then
      If DateDiff("n", Format(Val(Format(Mid(cboWorkTime, InStr(cboWorkTime, "~") + 1), "HHMM") & "00"), "00:00:00"), Text2) < 30 Then
         If cmdOK.Caption = "確認" And cmdOK.Enabled = True Then Combo1.ListIndex = 0 '1.正常上班時段 (預帶)
         ChkWorkTime = True
      End If
   End If
End Function

Private Sub cmdModify_Click()
   cmdOK.Visible = True
   cmdOK.Enabled = True
   cmdOK.Caption = "存檔"
   textB1505.Locked = False
   textB1505.BackColor = &H80000005
   cboWorkTime.Locked = False
   cboWorkTime.BackColor = &H80000005
   Combo1.Locked = False
   Combo1.BackColor = &H80000005
   cmdModify.Enabled = False
   Exit Sub
End Sub

Private Sub cmdOvertime_Click()
   If IsStaffOvertimeExist(m_B1501, m_B1502) = True Then
      Call frm180301_03.SetParent(Me)
      frm180301_03.txtB1003 = m_B1501
      frm180301_03.m_SA02 = m_SO02
      frm180301_03.m_SA03 = m_So03
      '加班
      frm180301_03.QueryData_3
      frm180301_03.Show
      Me.Hide
      If cmdOK.Caption = "確認" And cmdOK.Enabled = True Then Combo1.ListIndex = 1 '2.處理公務(預帶)
   
   ElseIf IsABS010Exist(m_B1501, m_B1502) = True Then
      Call frm180301_03.SetParent(Me)
      '出缺勤
      frm180301_03.txtB1001 = m_B1001
      frm180301_03.QueryData
      frm180301_03.Show
      Me.Hide
      If cmdOK.Caption = "確認" And cmdOK.Enabled = True Then Combo1.ListIndex = 1 '2.處理公務(預帶)
   
   Else
      '尚未輸入加班單,請填寫
      If Trim(Combo1.Text) = "" Or Combo1.ListIndex = 1 Then
InputOverTime:
         Call frm180102.SetParent(Me)
         frm180102.Hide
         '新增表單
         frm180102.txtB1001 = ""
         frm180102.txtB1003 = m_B1501
         frm180102.CboB1002 = "02 加班"
         Call frm180102.CboB1002_Click
         frm180102.txtB1004 = TransDate(m_B1502, 1) '民國年
         frm180102.Show
         'Me.Hide 畫面開著人員才能看打卡時間,方便填加班單
      Else
         If MsgBox("逾時原因非處理公務，是否要調整為處理公務並且開啟加班單填寫？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            GoTo InputOverTime
         End If
      End If
   End If
End Sub

Private Sub cmdQueryNext_Click()
   Unload Me
   frm160018.Show
   frm160018.PubShowNextData
End Sub

Private Sub cmdExit_Click()
   Unload Me
   frm160018.QueryData
   frm160018.Show
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly()
   If m_EditMode = 1 Then '新增
      textB1501.Locked = False
      textB1501.BackColor = &H80000005 '白色
      textB1502.Locked = False
      textB1502.BackColor = &H80000005
      textB1505.Locked = False
      textB1505.BackColor = &H80000005
      cboWorkTime.Locked = False
      cboWorkTime.BackColor = &H80000005
      Combo1.Locked = False
      Combo1.BackColor = &H80000005
   Else 'm_EditMode = 2 '修改
      textB1501.Locked = True
      textB1501.BackColor = &H8000000F '灰色
      textB1502.Locked = True
      textB1502.BackColor = &H8000000F
      '已輸入逾時原因後,不能修改
      If Trim(Combo1.Text) <> "" Or textB1501 <> strUserNum Then '非自己的資料
         textB1505.Locked = True
         textB1505.BackColor = &H8000000F
         cboWorkTime.Locked = True
         cboWorkTime.BackColor = &H8000000F
         Combo1.Locked = True
         Combo1.BackColor = &H8000000F
      Else
         textB1505.Locked = False
         textB1505.BackColor = &H80000005 '白色
         cboWorkTime.Locked = False
         cboWorkTime.BackColor = &H80000005
         Combo1.Locked = False
         Combo1.BackColor = &H80000005
      End If
   End If
End Sub

'查詢資料
Public Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer

On Error GoTo ErrHand

   Screen.MousePointer = vbHourglass
   
   strSql = "select abs015.*,B1504||' '||ac03 reason from abs015,allcode" & _
            " where B1501='" & m_B1501 & "' and B1502=" & DBDATE(m_B1502) & _
            " and ac01(+)='18' and ac02(+)=B1504"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_EditMode = 2 '修改
      textB1501 = "" & rsTmp.Fields("B1501")
      Label12.Caption = GetPrjSalesNM(textB1501)
      textB1502 = ChangeWStringToTString("" & rsTmp.Fields("B1502"))
      
      Call GetPr02 '取得打卡時間
      '用上班打卡時間檢查適合的上班時段
      For intI = intByPassArea To 1 Step -1
         If Val(Replace(Text1, ":", "")) > Val(Format(strByPassStarTime(intI), "HHMM") & "00") Then
            cboWorkTime.RemoveItem intI - 1
         End If
      Next intI
      If cboWorkTime.ListCount = 0 Then
         For intI = 1 To intByPassArea
            cboWorkTime.AddItem strByPassStarTime(intI) & "~" & strByPassEndTime(intI)
         Next intI
      End If
      '上班時段
      If "" & rsTmp.Fields("B1503") <> "" Then
         cboWorkTime.Text = "" & rsTmp.Fields("B1503")
         cboWorkTime.Tag = Trim(cboWorkTime.Text)
      Else
         cboWorkTime.ListIndex = 0 '預設值
      End If
      
      Combo1.Text = "" & rsTmp.Fields("reason")
      Combo1.Tag = Trim(Combo1.Text)
      cmdModify.Visible = False '隱藏修改鍵
      If Trim(Combo1.Text) <> "" Then
         cmdOK.Enabled = False
         '只有人事才能修改資料
         If GetStaffDepartment(strUserNum) = "M21" Then cmdModify.Visible = True
      Else
         '尚未輸入逾時原因,人員才能輸入資料做確認
         cmdOK.Enabled = True
      End If
      If textB1501 <> strUserNum Then cmdOK.Visible = False '只有當事人才能顯示出確認鍵
      
      textB1505 = "" & rsTmp.Fields("B1505")
      textB1505.Tag = Trim(textB1505.Text)
      
      If IsStaffOvertimeExist(m_B1501, m_B1502) = True Or IsABS010Exist(m_B1501, m_B1502) = True Then
         cmdOvertime.BackColor = &H80FF80 '綠色
         If Trim("" & rsTmp.Fields("reason")) = "" Then Combo1.ListIndex = 1 '2.處理公務(預帶)
      Else
         cmdOvertime.BackColor = &H8000000F '灰色 按鈕表面
         If textB1501 <> strUserNum Or Trim(Combo1.Text) <> "" Then
            cmdOvertime.FontBold = False
            cmdOvertime.Enabled = False '只有當事人才能填寫自己的加班單
         End If
      End If
      
      Call UpdateCUID(rsTmp)
   Else
      m_EditMode = 1 '新增
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   rsTmp.Close
   Call SetCtrlReadOnly
   
ErrHand:
   Set rsTmp = Nothing
   Screen.MousePointer = vbDefault
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String, strUName As String
Dim strCDate As String, strUDate As String
Dim strCTime As String, strUTime As String

   If IsNull(rsSrcTmp.Fields("B1506")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1506")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("B1506"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1507")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1507")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("B1507"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1508")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1508")) = False Then
         strTemp = Right("000000" & rsSrcTmp.Fields("B1508"), 6)
         strCTime = Format(strTemp, "00:00:00")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1509")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1509")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("B1509"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1510")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1510")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("B1510"))
         strUDate = Format(strTemp, "###/##/##")
      End If
      cmdOK.Tag = rsSrcTmp.Fields("B1510")
   End If
   If IsNull(rsSrcTmp.Fields("B1511")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1511")) = False Then
         strTemp = rsSrcTmp.Fields("B1511")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   '設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

'確認/存檔
Private Sub cmdok_Click()
Dim strSubject As String, strContent As String, strTo As String
Dim strCon As String
   
   If cmdOK.Caption = "存檔" Then
      strCon = ""
      If cboWorkTime.Tag <> cboWorkTime.Text Then
         strCon = strCon & ",B1503=" & CNULL(Trim(cboWorkTime.Text))
      End If
      If Combo1.Tag <> Combo1.Text Then
         strCon = strCon & ",B1504=" & CNULL(Trim(Left(Combo1.Text, 1)))
      End If
      If textB1505.Tag <> textB1505.Text Then
         strCon = strCon & ",B1505=" & CNULL(Trim(textB1505.Text))
      End If
      If strCon <> "" Then
         strCon = Mid(strCon, 2)
      Else
         MsgBox "無異動資料 !!!"
         Exit Sub
      End If
   Else
      strCon = "B1503=" & CNULL(Trim(cboWorkTime.Text)) & _
               ",B1504=" & CNULL(Trim(Left(Combo1.Text, 1))) & _
               ",B1505=" & CNULL(Trim(textB1505.Text))
   End If
   
   If CheckDataValid = False Then Exit Sub
   If TxtValidate = False Then Exit Sub

On Error GoTo ErrHand

   cnnConnection.BeginTrans
   strSql = "update ABS015 set " & strCon & _
            " where B1501='" & textB1501 & "' and B1502=" & DBDATE(textB1502)
   If Val(cmdOK.Tag) > 0 Then Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   
   If cmdOK.Caption = "存檔" Then
      '發EMail通知當事者
      strSubject = GetPrjSalesNM(textB1501) & " " & ChangeWStringToTDateString(DBDATE(textB1502)) & " 下班逾30分鐘原因確認，人事處修改通知！"
      strContent = "員工姓名：" & Label12 & vbCrLf
      strContent = strContent & "日　　期：" & ChangeTStringToTDateString(textB1502) & vbCrLf
      strContent = strContent & "打卡時間：" & Text1 & "~" & Text2 & vbCrLf
      strContent = strContent & "上班時段：" & cboWorkTime.Text & vbCrLf
      strContent = strContent & "逾時原因：" & Combo1.Text & vbCrLf
      strContent = strContent & "備　　註：" & textB1505 & vbCrLf
      strContent = strContent & vbCrLf & "請至案件管理系統的一般作業\出缺勤作業\表單\下班逾30分鐘原因確認，做確認。"
      
      strTo = PUB_GetST59(textB1501)
      If IsNull(strTo) Or strTo = "" Then
         strTo = textB1501
      End If
      PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True
   End If
   
   Call cmdQueryNext_Click
   Exit Sub

ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "資料" & cmdOK.Caption & "失敗！" & vbCrLf & Err.Description
End Sub

'檢查 加班 資料是否存在
Private Function IsStaffOvertimeExist(StrST01 As String, strDate As String) As Boolean
Dim Rs As ADODB.Recordset

   IsStaffOvertimeExist = False
   strExc(0) = "select * from Staff_Overtime where sO01='" & StrST01 & "' and sO02=" & DBDATE(strDate)
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_SO02 = Rs.Fields("So02")
      m_So03 = Rs.Fields("So03")
      m_B1001 = Rs.Fields("So13")
      IsStaffOvertimeExist = True
   End If
   Rs.Close
   Set Rs = Nothing
End Function
'檢查 加班簽核 資料是否存在
Private Function IsABS010Exist(StrST01 As String, strDate As String) As Boolean
Dim Rs As ADODB.Recordset

   IsABS010Exist = False
   strExc(0) = "select * from abs010 WHERE B1002 in('02')" & _
               " and B1018 not in('" & 註銷 & "','" & 已核准 & "')" & _
               " and B1003='" & StrST01 & "' and B1004=" & DBDATE(strDate)
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   m_B1001 = ""
   If intI = 1 Then
      m_B1001 = Rs.Fields("B1001")
      IsABS010Exist = True
   End If
   Rs.Close
   Set Rs = Nothing
End Function

'檢查 ABS015 資料是否存在
Private Function IsABS015Exist(StrST01 As String, strDate As String) As Boolean
Dim Rs As ADODB.Recordset

   IsABS015Exist = False
   strExc(0) = "select * from ABS015 where b1501='" & StrST01 & "' and b1502=" & DBDATE(strDate)
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      IsABS015Exist = True
   End If
   Rs.Close
   Set Rs = Nothing
End Function

Private Sub Form_Load()
Dim MyRs As New ADODB.Recordset
   
   MoveFormToCenter Me
   
   '逾時原因
   Combo1.Clear
   Set MyRs = New ADODB.Recordset
   If MyRs.State = 1 Then MyRs.Close
   strSql = "select ac02||' '||ac03 from allcode where ac01='18' order by ac02"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
      While Not MyRs.EOF
         Combo1.AddItem "" & MyRs.Fields(0).Value
         MyRs.MoveNext
      Wend
   End If
   '上班時段
   SetB102829Combo cboWorkTime, 1, , , True
   
   Call ClearData
   Call QueryData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160018_1 = Nothing
End Sub

'清除欄位值
Private Sub ClearData()
   textB1501.Text = "": Label12.Caption = ""
   textB1502.Text = ""
   Text1.Text = "": Text2.Text = ""
   Combo1.ListIndex = -1
   cboWorkTime.ListIndex = -1
   textB1505.Text = ""
   Label23 = Empty
End Sub

'取得打卡時間
Private Sub GetPr02()
Dim stSQL As String

   Screen.MousePointer = vbHourglass
   
   stSQL = "select min(sqltime6(pr02)),max(sqltime6(pr02))" & _
           " from staff,staffcarddata,pollrecord where scd01(+)=st01 and pr03(+)=scd02 and pr01>0" & _
           " and st01='" & textB1501 & "' and pr01=" & DBDATE(textB1502) & _
           " order by pr02 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      Text1.Text = RsTemp.Fields(0)
      Text2.Text = RsTemp.Fields(1)
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim objCbo As Object
Dim bolChk As Boolean
   
   CheckDataValid = False

'   If textB1501 = "" Then
'       MsgBox "員工代號不可空白 !!!"
'       textB1501.SetFocus
'       Exit Function
'    End If
'    If textB1502 = "" Then
'       MsgBox "日期不可空白 !!!"
'       textB1502.SetFocus
'       Exit Function
'    End If
'   '檢查資料是否已存在
'   If IsABS015Exist(textB1501, DBDATE(textB1502)) = True Then
'      strTit = "新增資料"
'      strMsg = "該筆記錄已存在"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textB1502.SetFocus
'      Exit Function
'   End If
   
   If cboWorkTime.Text = "" Then
      MsgBox "上班時段不可空白 !!!"
      If cboWorkTime.Enabled = True Then cboWorkTime.SetFocus
      Exit Function
   End If
   bolChk = False: strExc(1) = cboWorkTime.Text
   For intI = 0 To Me.cboWorkTime.ListCount - 1
      If strExc(1) = Me.cboWorkTime.List(intI) Then
         bolChk = True
      End If
   Next intI
   If bolChk = False Then
      MsgBox "上班時段有誤，請重新點選 !!!"
      If cboWorkTime.Enabled = True Then cboWorkTime.SetFocus
      Exit Function
   End If
   
   '2025/10/30 人事處修改時,逾時原因才可清空
   If Not (cmdOK.Caption = "存檔" And cmdOK.Enabled = True And Trim(Combo1.Text) = "") Then
   '2025/10/30 END
      If Trim(Combo1.Text) = "" Then
         MsgBox "逾時原因不可空白 !!!"
         If Combo1.Enabled = True Then Combo1.SetFocus
         Exit Function
      End If
   End If
   If Trim(Combo1.Text) <> "" Then
      bolChk = False: strExc(1) = Trim(Combo1.Text)
      For intI = 0 To Me.Combo1.ListCount - 1
         If strExc(1) = Me.Combo1.List(intI) Then
            bolChk = True
         End If
      Next intI
      If bolChk = False Then
         MsgBox "逾時原因有誤，請重新點選 !!!"
         If Combo1.Enabled = True Then Combo1.SetFocus
         Exit Function
      End If
      
      If cmdOvertime.BackColor = &H80FF80 And Left(Trim(Combo1.Text), 1) <> "2" Then '綠色:有加班單
         MsgBox "有加班單，請點選 2.處理公務 !!!"
         If Combo1.Enabled = True Then Combo1.SetFocus
         Exit Function
      End If
      
      If Left(Trim(Combo1.Text), 1) = "2" Then '2.處理公務
         If IsStaffOvertimeExist(m_B1501, m_B1502) = False _
            And IsABS010Exist(m_B1501, m_B1502) = False Then
            '尚未輸入加班單,請填寫
            Call cmdOvertime_Click
            Exit Function
         End If
      End If
      
      If InStr(Trim(Combo1.Text), "其他") > 0 And Trim(textB1505) = "" Then
         MsgBox "逾時原因為其他時，請在備註欄填寫事由 !!!"
         If textB1505.Enabled = True Then textB1505.SetFocus
         Exit Function
      End If
      If ChkWorkTime = False And Left(Trim(Combo1.Text), 1) = "1" Then
         MsgBox "不在正常上班時段裡，請重新點選逾時原因 或 重選上班時段!!!"
         cboWorkTime.SetFocus
         Exit Function
      End If
   End If
   
   CheckDataValid = True
End Function

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean

   TxtValidate = False

   If Me.textB1501.Enabled = True Then
      Cancel = False
      textB1501_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textB1502.Enabled = True Then
      Cancel = False
      textB1502_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   '檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   
   TxtValidate = True
End Function

Private Sub textB1501_GotFocus()
   InverseTextBox textB1501
End Sub

Private Sub textB1501_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textB1501_Validate(Cancel As Boolean)
Dim Rs As New ADODB.Recordset

   If textB1501.Text = "" Then Label12 = ""

   If textB1501 <> "" Then
      ' 檢查員工編號規則
      If ChkStaffST04(textB1501) Then
         Call textB1501_GotFocus
         Cancel = True
         Exit Sub
      End If
      Label12 = GetStaffName(textB1501, True)
      If Label12 = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         Call textB1501_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub textB1502_GotFocus()
   InverseTextBox textB1502
End Sub

Private Sub textB1502_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textB1502_Validate(Cancel As Boolean)
   If textB1502 = "" Then Exit Sub
   If textB1502 <> "" Then
      If ChkDate(textB1502) = False Then
         Call textB1502_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub textB1505_GotFocus()
   InverseTextBox textB1505
   CloseIme
End Sub
