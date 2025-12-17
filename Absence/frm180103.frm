VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm180103 
   BorderStyle     =   1  '單線固定
   Caption         =   "職代/簽核主管代填表單"
   ClientHeight    =   4350
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   7120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7120
   Tag             =   "加班資料"
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   360
      Left            =   4650
      TabIndex        =   28
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5460
      TabIndex        =   13
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox txtB1001 
      Height          =   300
      Left            =   2010
      MaxLength       =   8
      TabIndex        =   0
      Top             =   930
      Width           =   1185
   End
   Begin VB.TextBox txtB1007_1 
      Height          =   300
      Left            =   4350
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2490
      Width           =   585
   End
   Begin VB.TextBox txtB1007_2 
      Height          =   300
      Left            =   5220
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2490
      Width           =   585
   End
   Begin VB.TextBox txtB1006 
      Height          =   300
      Left            =   4860
      MaxLength       =   7
      TabIndex        =   6
      Top             =   2130
      Width           =   945
   End
   Begin VB.ComboBox CboB1002 
      Height          =   300
      ItemData        =   "frm180103.frx":0000
      Left            =   2010
      List            =   "frm180103.frx":0002
      TabIndex        =   2
      Top             =   1530
      Width           =   1695
   End
   Begin VB.TextBox txtB1005_2 
      Height          =   300
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2490
      Width           =   585
   End
   Begin VB.TextBox txtB1004 
      Height          =   300
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   3
      Top             =   2130
      Width           =   945
   End
   Begin VB.TextBox txtB1005_1 
      Height          =   300
      Left            =   2010
      MaxLength       =   2
      TabIndex        =   4
      Top             =   2490
      Width           =   585
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "新增(&N)"
      Height          =   360
      Left            =   1410
      TabIndex        =   9
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "存檔(&S)"
      Height          =   360
      Left            =   3840
      TabIndex        =   12
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "刪除(&D)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   3030
      TabIndex        =   11
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改(&M)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   2220
      TabIndex        =   10
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   6270
      TabIndex        =   14
      Top             =   60
      Width           =   800
   End
   Begin MSForms.Label Label26 
      Height          =   195
      Left            =   210
      TabIndex        =   31
      Top             =   4080
      Width           =   6825
      VariousPropertyBits=   27
      Caption         =   "CREATE :                                                    UPDATE : "
      Size            =   "12039;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboB1003 
      Height          =   285
      Left            =   2010
      TabIndex        =   1
      Top             =   1230
      Width           =   1680
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2963;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(共7碼)"
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   5
      Left            =   3510
      TabIndex        =   30
      Top             =   2190
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "註：按取消進入查詢狀態"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   4
      Left            =   4650
      TabIndex        =   29
      Top             =   570
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "註：可以輸入表單編號或員工代號(＋表單類別＋起始日期)做查詢"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   27
      Top             =   3420
      Width           =   5160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "表單編號："
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   26
      Top             =   960
      Width           =   900
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   4200
      X2              =   6000
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3900
      TabIndex        =   25
      Top             =   2490
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "時"
      Height          =   180
      Left            =   4980
      TabIndex        =   24
      Top             =   2550
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      Height          =   180
      Index           =   2
      Left            =   4440
      TabIndex        =   23
      Top             =   2190
      Width           =   360
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "            迄"
      Height          =   180
      Left            =   4620
      TabIndex        =   22
      Top             =   1860
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "分"
      Height          =   180
      Left            =   5850
      TabIndex        =   21
      Top             =   2550
      Width           =   180
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   1980
      X2              =   6030
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "表單類別："
      Height          =   180
      Left            =   1080
      TabIndex        =   20
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   19
      Top             =   1290
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   2010
      X2              =   3810
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "分"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   3510
      TabIndex        =   18
      Top             =   2550
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "時間：                 起                    "
      Height          =   180
      Left            =   1440
      TabIndex        =   17
      Top             =   1860
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   17
      Left            =   2100
      TabIndex        =   16
      Top             =   2190
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "時"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2640
      TabIndex        =   15
      Top             =   2550
      Width           =   180
   End
End
Attribute VB_Name = "frm180103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/8/24
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
Dim i As Integer
Dim m_B1017 As String


Private Sub cmdCancel_Click()
   m_EditMode = "0" '取消
   Call ControlCmd
End Sub

Private Sub cmdQuery_Click()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
   
   If txtB1001 = "" Then
      If CboB1003 = "" Then
         If CboB1003 = "" Or CboB1002 = "" Or txtB1004 = "" Then
            MsgBox "請輸入查詢條件！"
            Exit Sub
         End If
      End If
   End If
   
   m_EditMode = 4 '查詢
   Screen.MousePointer = vbHourglass
   
   If txtB1001 <> "" Then
      strCon = " and B1001='" & Me.txtB1001 & "' "
   ElseIf CboB1003 <> "" And CboB1002 <> "" And txtB1004 <> "" Then
      strCon = " and B1003='" & Left(CboB1003, 5) & "' " & _
               " and B1002='" & Left(CboB1002, 2) & "' " & _
               " and B1004=" & DBDATE(txtB1004)
   ElseIf CboB1003 <> "" Then
      strCon = " and B1003='" & Left(CboB1003, 5) & "' "
   End If
   
   '出缺勤電子簽核主檔
   strSql = "Select B1001,B1002,B1003,B1004,substr(ltrim(to_char('0000'||to_char(B1005),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),3,2) B1005,B1006,substr(ltrim(to_char('0000'||to_char(B1007),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),3,2) B1007,B1008||' '||AC03 B1008,B1009,B1010,B1011,B1012,B1013,B1014,B1015,B1016,B1017," & B1018CName & " B1018,B1019,B1020,B1021,B1022,B1023,B1024,B1025,B1026,B1027 " & _
            "From ABS010, allcode " & _
            "Where ac01(+)='04' and B1008=ac02(+) and B1018='" & 主管代填 & "' " & strCon
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   m_B1017 = ""
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("B1001")) Then txtB1001 = rsTmp.Fields("B1001")
      If Not IsNull(rsTmp.Fields("B1002")) Then CboB1002 = GetB1002Value(rsTmp.Fields("B1002"))
      If Not IsNull(rsTmp.Fields("B1003")) Then CboB1003 = rsTmp.Fields("B1003") & "  " & GetPrjSalesNM(rsTmp.Fields("B1003"))
      If Not IsNull(rsTmp.Fields("B1004")) Then txtB1004 = ChangeWStringToTString(rsTmp.Fields("B1004"))
      If Not IsNull(rsTmp.Fields("B1005")) Then txtB1005_1 = Left(rsTmp.Fields("B1005"), 2): txtB1005_2 = Right(rsTmp.Fields("B1005"), 2)
      If Not IsNull(rsTmp.Fields("B1006")) Then txtB1006 = ChangeWStringToTString(rsTmp.Fields("B1006"))
      If Not IsNull(rsTmp.Fields("B1007")) Then txtB1007_1 = Left(rsTmp.Fields("B1007"), 2): txtB1007_2 = Right(rsTmp.Fields("B1007"), 2)
      If Not IsNull(rsTmp.Fields("B1017")) Then m_B1017 = rsTmp.Fields("B1017")
      
      Call CboB1002_Click
      Call UpdateCUID(rsTmp)
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
   '檢查人事系統裡是否已有表單編號
   If ChkPerSysB1001Exist(txtB1001, Left(Trim(CboB1003), 5)) = True Then Exit Sub
   
   If m_B1017 <> "" Then
      MsgBox "此表單已在簽核中！"
      Exit Sub
   End If
   
   Call ControlCmd
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub ClearField()
   txtB1001 = Empty
   CboB1002 = Empty
   CboB1003 = Empty
   txtB1004 = strSrvDate(2)
   txtB1005_1 = "8"
   txtB1005_2 = "00"
   txtB1006 = strSrvDate(2)
   txtB1007_1 = "17"
   txtB1007_2 = "00"
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   txtB1001.Locked = bEnable
   If bEnable Then txtB1001.BackColor = &H8000000F Else txtB1001.BackColor = &H80000005
   CboB1003.Locked = bEnable
   If bEnable Then CboB1003.BackColor = &H8000000F Else CboB1003.BackColor = &H80000005
   CboB1002.Locked = bEnable
   If bEnable Then CboB1002.BackColor = &H8000000F Else CboB1002.BackColor = &H80000005
   txtB1004.Locked = bEnable
   If bEnable Then txtB1004.BackColor = &H8000000F Else txtB1004.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   txtB1005_1.Enabled = bEnable
   If bEnable = False Then txtB1005_1.BackColor = &H8000000F Else txtB1005_1.BackColor = &H80000005
   txtB1005_2.Enabled = bEnable
   If bEnable = False Then txtB1005_2.BackColor = &H8000000F Else txtB1005_2.BackColor = &H80000005
   txtB1006.Enabled = bEnable
   If bEnable = False Then txtB1006.BackColor = &H8000000F Else txtB1006.BackColor = &H80000005
   txtB1007_1.Enabled = bEnable
   If bEnable = False Then txtB1007_1.BackColor = &H8000000F Else txtB1007_1.BackColor = &H80000005
   txtB1007_2.Enabled = bEnable
   If bEnable = False Then txtB1007_2.BackColor = &H8000000F Else txtB1007_2.BackColor = &H80000005
End Sub

Private Sub cmdNew_Click()
   m_EditMode = 1 '新增
   Call ControlCmd
   Me.CboB1003.SetFocus
End Sub

Private Sub cmdModify_Click()
   m_EditMode = 2 '修改
   Call ControlCmd
End Sub

Private Sub ControlCmd()
   Me.txtB1001.Enabled = True
   Me.txtB1001.BackColor = &H80000005
   
   If m_EditMode = 0 Then '取消
      Call ClearField
      Call SetKeyReadOnly(False)
      Call SetCtrlReadOnly(False)
      Me.cmdNew.Enabled = True
      Me.cmdModify.Enabled = False
      Me.cmdDel.Enabled = False
      Me.cmdSave.Enabled = False
      Me.cmdQuery.Enabled = True
      Me.CboB1003.SetFocus
      
   ElseIf m_EditMode = 1 Or m_EditMode = 3 Then '新增,刪除
      Call ClearField
      Call SetKeyReadOnly(False)
      Call SetCtrlReadOnly(True)
      Me.cmdNew.Enabled = False
      Me.cmdModify.Enabled = False
      Me.cmdDel.Enabled = False
      Me.cmdSave.Enabled = True
      Me.cmdQuery.Enabled = False
      Me.txtB1001.Enabled = False
      Me.txtB1001.BackColor = &H8000000F
      'Me.CboB1003.SetFocus
      
   ElseIf m_EditMode = 2 Then '修改
      Call SetKeyReadOnly(True)
      Call SetCtrlReadOnly(True)
      Me.cmdNew.Enabled = False
      Me.cmdModify.Enabled = False
      Me.cmdDel.Enabled = True
      Me.cmdSave.Enabled = True
      Me.cmdQuery.Enabled = False
      Me.txtB1005_1.SetFocus
      
   ElseIf m_EditMode = 4 Then '查詢
      Call SetKeyReadOnly(True)
      Call SetCtrlReadOnly(False)
      Me.cmdNew.Enabled = False
      Me.cmdModify.Enabled = True
      Me.cmdDel.Enabled = True
      Me.cmdSave.Enabled = False
      Me.cmdQuery.Enabled = True
      Me.CboB1003.SetFocus
   End If
End Sub

Private Sub cmdDel_Click()
   
On Error GoTo ErrHand
   
   If MsgBox("確定是否要刪除資料？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then Exit Sub
   
   m_EditMode = 3 '刪除
   
   Screen.MousePointer = vbHourglass
   
   cnnConnection.BeginTrans
   
   strSql = "DELETE FROM ABS010 WHERE B1001 = '" & txtB1001 & "' "
   Pub_SeekTbLog strSql '記錄刪除Log
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   
   MsgBox "已刪除！"
   Screen.MousePointer = vbDefault
   
   Call ControlCmd
   Me.CboB1003.SetFocus
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
Dim strB1006 As String
   
On Error GoTo ErrHand
   
   '檢查條件
   If txtValidate = False Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   
   cnnConnection.BeginTrans
   
   '迄止日期
   If txtB1006.Visible = True Then
      strB1006 = DBDATE(txtB1006)
   End If
   
   If m_EditMode = "1" Then '新增
      '表單編號自動給號
      txtB1001 = AutoNo_ABS("ABS", 5)
      
      '檢查是否還有自動給號資料不一致的問題
      strSql = "select AU03 from autonumber where AU01='ABS'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Val(RsTemp.Fields("AU03")) <> Val(Right(txtB1001, Len(txtB1001.Text) - 3)) Then
            MsgBox "系統自動給號(" & txtB1001 & ")更新有誤，請洽電腦中心！", vbInformation, "系統錯誤"
            txtB1001 = ""
            GoTo ErrHand
            Exit Sub
         End If
      End If
      
      strSql = "insert into ABS010(B1001,B1002,B1003,B1004,B1005,B1006,B1007,B1018) " & _
               "values(" & CNULL(txtB1001) & "," & CNULL(Left(CboB1002, 2)) & "," & CNULL(Left(CboB1003, 5)) & "," & _
               CNULL(DBDATE(txtB1004)) & "," & CNULL(txtB1005_1 & Format("00" & txtB1005_2, "00")) & "," & CNULL(strB1006) & "," & _
               CNULL(txtB1007_1 & Format("00" & txtB1007_2, "00")) & ",'" & 主管代填 & "') "
   ElseIf m_EditMode = "2" Then '修改
      strSql = "update ABS010 set " & _
               "B1002=" & CNULL(Left(CboB1002, 2)) & _
               ",B1004=" & CNULL(DBDATE(txtB1004)) & _
               ",B1005=" & CNULL(txtB1005_1 & Format("00" & txtB1005_2, "00")) & _
               ",B1006=" & CNULL(strB1006) & _
               ",B1007=" & CNULL(txtB1007_1 & Format("00" & txtB1007_2, "00")) & _
               " where B1001=" & CNULL(txtB1001)
   End If
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   
   'Add By Sindy 2018/11/16 檢查在此請假區間中是否有幫他人做職代,若有,發Mail通知其他職代
   If m_EditMode = "1" And _
      (Left(CboB1002, 2) = 表單類別_請假 Or Left(CboB1002, 2) = 表單類別_出差) Then
      Call CheckIsPersonRestSectorMail(txtB1001)
   End If
   
   MsgBox "存檔成功！"
   Screen.MousePointer = vbDefault
   
   Call cmdQuery_Click
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox " 存檔失敗！" & vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
Dim strA0911 As String, strA0925 As String
   
   MoveFormToCenter Me
   
   strA0911 = GetStaffA0911(strUserNum, strA0925) 'Modify By Sindy 2023/12/20
   
   '預設值
   Call SetB1003Combo(CboB1003, strA0911, strA0925)
   SetB1002Combo CboB1002
   m_EditMode = 1 '新增狀態
   
   Call ControlCmd
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180103 = Nothing
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("B1022")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1022")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("B1022"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1023")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1023")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("B1023"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1024")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1024")) = False Then
         strTemp = rsSrcTmp.Fields("B1024")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1025")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1025")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("B1025"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1026")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1026")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("B1026"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1027")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1027")) = False Then
         strTemp = rsSrcTmp.Fields("B1027")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label26.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Function txtValidate() As Boolean
Dim Cancel As Boolean

txtValidate = False

If CboB1003.Text = "" Then
    MsgBox "員工代號不可以空白！", vbExclamation
    CboB1003.SetFocus
    Exit Function
End If
If CboB1002.Text = "" Then
    MsgBox "表單類別不可以空白！", vbExclamation
    CboB1002.SetFocus
    Exit Function
End If
If txtB1004.Text = "" Then
    MsgBox "日期起不可以空白！", vbExclamation
    txtB1004.SetFocus
    Exit Function
End If
If txtB1005_1.Text = "" Or txtB1005_1.Text = "00" Then
    MsgBox "必須輸入起始(時)！", vbExclamation
    txtB1005_1.SetFocus
    Exit Function
End If
If txtB1006.Visible = True Then
   If txtB1006.Text = "" Then
       MsgBox "日期迄不可以空白！", vbExclamation
       txtB1006.SetFocus
       Exit Function
   End If
End If
If txtB1007_1.Text = "" Or txtB1007_1.Text = "00" Then
    MsgBox "必須輸入迄止(時)！", vbExclamation
    txtB1007_1.SetFocus
    Exit Function
End If

'檢查起迄日期時間區間是否有重覆
If m_EditMode = 1 Or m_EditMode = 2 Then
   If Left(CboB1002, 2) = 表單類別_加班 Then
      If IsRecordExist(Left(CboB1003, 5), DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1004), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
         txtB1004.SetFocus
         Exit Function
      End If
   Else
      If IsRecordExist(Left(CboB1003, 5), DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1006), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
         txtB1006.SetFocus
         Exit Function
      End If
   End If
End If

If Me.txtB1004.Enabled = True Then
   Cancel = False
   txtB1004_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.txtB1005_1.Enabled = True Then
   Cancel = False
   txtB1005_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.txtB1005_2.Enabled = True Then
   Cancel = False
   txtB1005_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If txtB1006.Visible = True Then
   If Me.txtB1006.Enabled = True Then
      Cancel = False
      txtB1006_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
End If
If Me.txtB1007_1.Enabled = True Then
   Cancel = False
   txtB1007_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.txtB1007_2.Enabled = True Then
   Cancel = False
   txtB1007_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

txtValidate = True
End Function

Private Sub txtB1001_GotFocus()
   InverseTextBox txtB1001
End Sub

Private Sub txtB1001_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub CboB1003_GotFocus()
   InverseTextBox CboB1003
End Sub

Private Sub CboB1003_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboB1003_LostFocus()
   If CboB1003.Text > "" And Len(Trim(CboB1003.Text)) = 5 Then
      '抓取員工姓名
      CboB1003.Text = SetCboStaffName(CboB1003.Text)
   End If
End Sub

Private Sub CboB1003_Validate(Cancel As Boolean)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If CboB1003 <> "" Then
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(CboB1003, 5)) = True Then
         Call CboB1003_GotFocus
         Cancel = True
         Exit Sub
      End If
      '檢查 員工不可為”不寄信”
      If ChkStaffST14(Left(CboB1003, 5), False) = True Then
         Call CboB1003_GotFocus
         MsgBox "此人員需走紙本假單流程，請通知人事處該人員休假未填假單!", vbExclamation + vbOKOnly
         Cancel = True
         Exit Sub
      End If
      '檢查是否有代填表單的權限
      If ChkLimitsIsOk() = False Then
         Call CboB1003_GotFocus
         MsgBox "無權限代填此人員表單!", vbExclamation + vbOKOnly
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub CboB1002_Click()
   If Left(CboB1002.Text, 2) = 表單類別_請假 Then
      Label1(2).Visible = True
      txtB1006.Visible = True
   ElseIf Left(CboB1002.Text, 2) = 表單類別_加班 Then
      Label1(2).Visible = False
      txtB1006.Visible = False
   ElseIf Left(CboB1002.Text, 2) = 表單類別_出差 Then
      Label1(2).Visible = True
      txtB1006.Visible = True
   End If
End Sub

Private Sub CboB1002_GotFocus()
   InverseTextBox CboB1002
End Sub

Private Sub CboB1002_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboB1002_LostFocus()
   If CboB1002.Text > "" Then
      For i = 0 To CboB1002.ListCount - 1
         If Left(CboB1002.List(i), 2) = CboB1002.Text Then CboB1002.Text = CboB1002.List(i): Exit For
      Next i
   End If
End Sub

Private Sub CboB1002_Validate(Cancel As Boolean)
Dim bolComp As Boolean
   
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If CboB1002 <> "" Then
      bolComp = False
      For i = 0 To CboB1002.ListCount
         If Left(CboB1002, 2) = Left(CboB1002.List(i), 2) Then
            bolComp = True
            Exit For
         End If
      Next i
      If bolComp = False Then
         MsgBox "表單類別有誤!!!", vbExclamation + vbOKOnly
         Call CboB1002_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub txtB1004_GotFocus()
   InverseTextBox txtB1004
   CloseIme
End Sub

Private Sub txtB1004_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1004_Validate(Cancel As Boolean)
If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub

If txtB1004 <> "" Then
   If CheckIsTaiwanDate(txtB1004, False) = False Then
      MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
      Call txtB1004_GotFocus
      Cancel = True
      Exit Sub
   End If
   If Left(CboB1002, 2) = 表單類別_請假 Then
      If ChkWorkDay(DBDATE(txtB1004)) = False Then
         MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
         Call txtB1004_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   If txtB1004 <> "" And txtB1006 <> "" Then
'      '檢查不能跨月
'      If Left(DBDATE(txtB1004), 6) <> Left(DBDATE(txtB1006), 6) Then
'         Call txtB1004_GotFocus
'         Cancel = True
'         MsgBox "跨月請另填新表單！", vbInformation, "輸入日期錯誤"
'         Exit Sub
'      End If
      If Val(txtB1004) > Val(txtB1006) Then
         txtB1006 = ""
      Else
         If RunNick2(txtB1004, txtB1006) Then
            Call txtB1004_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End If
End Sub

Private Sub txtB1005_1_GotFocus()
   InverseTextBox txtB1005_1
End Sub

Private Sub txtB1005_1_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1005_1_Validate(Cancel As Boolean)
If txtB1005_1 = "" Then txtB1005_1 = "00"

If txtB1005_1 <> "" Then
   If CheckLengthIsOK(txtB1005_1, txtB1005_1.MaxLength) = False Then
      Call txtB1005_1_GotFocus
      Cancel = True
      Exit Sub
   End If
   If Val(txtB1005_1.Text) = 0 And txtB1004 <> "" Then
      MsgBox "請輸入時分!", vbExclamation + vbOKOnly
      Call txtB1005_1_GotFocus
      Cancel = True
      Exit Sub
   End If
   If txtB1005_1.Text > 24 Then
      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
      Call txtB1005_1_GotFocus
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
End Sub

Private Sub txtB1005_2_GotFocus()
   InverseTextBox txtB1005_2
End Sub

Private Sub txtB1005_2_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1005_2_Validate(Cancel As Boolean)
If txtB1005_2 = "" Then txtB1005_2 = "00"

If txtB1005_2 <> "" Then
   If CheckLengthIsOK(txtB1005_2, txtB1005_2.MaxLength) = False Then
      Call txtB1005_2_GotFocus
      Cancel = True
      Exit Sub
   End If
   If txtB1005_2.Text > 59 Then
      Call txtB1005_2_GotFocus
      MsgBox "不可超過59分!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If Left(CboB1002, 2) = 表單類別_加班 Then
         If IsRecordExist(Left(CboB1003, 5), DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1004), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
            Call txtB1005_2_GotFocus
            'Cancel = True
            Exit Sub
         End If
      Else
         If IsRecordExist(Left(CboB1003, 5), DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1006), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
            Call txtB1005_2_GotFocus
            'Cancel = True
            Exit Sub
         End If
      End If
   End If
End If
CloseIme
End Sub

Private Sub txtB1006_GotFocus()
   InverseTextBox txtB1006
End Sub

Private Sub txtB1006_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1006_Validate(Cancel As Boolean)
If txtB1006 <> "" Then
   If CheckIsTaiwanDate(txtB1006, False) = False Then
      MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
      Call txtB1006_GotFocus
      Cancel = True
      Exit Sub
   End If
   If Left(CboB1002, 2) = 表單類別_請假 Then
      If ChkWorkDay(DBDATE(txtB1006)) = False Then
         Call txtB1006_GotFocus
         Cancel = True
         MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
         Exit Sub
      End If
   End If
   If txtB1004 <> "" And txtB1006 <> "" Then
'      '檢查不能跨月
'      If Left(DBDATE(txtB1004), 6) <> Left(DBDATE(txtB1006), 6) Then
'         Call txtB1006_GotFocus
'         Cancel = True
'         MsgBox "跨月請另填新表單！", vbInformation, "輸入日期錯誤"
'         Exit Sub
'      End If
      If RunNick2(txtB1004, txtB1006) Then
         Call txtB1006_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End If
End Sub

Private Sub txtB1007_1_GotFocus()
   InverseTextBox txtB1007_1
End Sub

Private Sub txtB1007_1_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1007_1_Validate(Cancel As Boolean)
If txtB1007_1 = "" Then txtB1007_1 = "00"

If txtB1007_1 <> "" Then
   If CheckLengthIsOK(txtB1007_1, txtB1007_1.MaxLength) = False Then
      Call txtB1007_1_GotFocus
      Cancel = True
      Exit Sub
   End If
   If Val(txtB1007_1.Text) = 0 And _
      ((txtB1004 <> "" And Left(CboB1002, 2) = 表單類別_加班) Or (txtB1006 <> "" And Left(CboB1002, 2) <> 表單類別_加班)) Then
      Call txtB1007_1_GotFocus
      MsgBox "請輸入時分!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   If txtB1007_1.Text > 24 Then
      Call txtB1007_1_GotFocus
      MsgBox "不可超過24時!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
End Sub

Private Sub txtB1007_2_GotFocus()
   InverseTextBox txtB1007_2
End Sub

Private Sub txtB1007_2_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtB1007_2_Validate(Cancel As Boolean)
If txtB1007_2 = "" Then txtB1007_2 = "00"

If txtB1007_2 <> "" Then
   If CheckLengthIsOK(txtB1007_2, txtB1007_2.MaxLength) = False Then
      Call txtB1007_2_GotFocus
      Cancel = True
      Exit Sub
   End If
   If txtB1007_2.Text > 59 Then
      Call txtB1007_2_GotFocus
      MsgBox "不可超過59分!", vbExclamation + vbOKOnly
      Cancel = True
      Exit Sub
   End If
   'If Trim(txtB1004) <> "" And Trim(txtB1005_1) <> "" And Trim(txtB1005_2) <> "" And Trim(txtB1006) <> "" And Trim(txtB1007_1) <> "" And Trim(txtB1007_2) <> "" Then
      If CheckIsTaiwanDate(txtB1004, False) = True And CheckIsTaiwanDate(txtB1006, False) = True Then
         If CompDateTime(txtB1004 & Format(txtB1005_1, "00") & Format(txtB1005_2, "00"), txtB1006 & Format(txtB1007_1, "00") & Format(txtB1007_2, "00")) = False Then
            Call txtB1007_2_GotFocus
            MsgBox "日期時間設定錯誤！", vbInformation, "輸入錯誤！"
            'Cancel = True
            Exit Sub
         End If
      End If
   'End If
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If Left(CboB1002, 2) = 表單類別_加班 Then
         If IsRecordExist(Left(CboB1003, 5), DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1004), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
            Call txtB1005_2_GotFocus
            'Cancel = True
            Exit Sub
         End If
      Else
         If IsRecordExist(Left(CboB1003, 5), DBDATE(txtB1004), Trim(txtB1005_1.Text & ":" & txtB1005_2.Text), DBDATE(txtB1006), Trim(txtB1007_1.Text & ":" & txtB1007_2.Text)) = True Then
            Call txtB1005_2_GotFocus
            'Cancel = True
            Exit Sub
         End If
      End If
   End If
End If
CloseIme
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String, ByVal strKEY04 As String, ByVal strKEY05 As String) As Boolean
   IsRecordExist = False
   
   If IsNull(strKEY01) Or strKEY01 = "" Then Exit Function
   If IsNull(strKEY02) Or strKEY02 = "" Then Exit Function
   If IsNull(strKEY03) Or strKEY03 = "" Then Exit Function
   If IsNull(strKEY04) Or strKEY04 = "" Then
      strKEY04 = strKEY02
      strKEY05 = strKEY03
   End If
   
   If CheckIsAbsenceExist(strKEY01, strKEY02, strKEY03, strKEY04, strKEY05, txtB1001, Left(Trim(CboB1002), 2)) = True Then IsRecordExist = True
   If IsRecordExist = True Then
      MsgBox "該筆記錄已存在", vbOKOnly, "新增資料"
   End If
End Function

'檢查是否有代填表單的權限
Private Function ChkLimitsIsOk() As Boolean
Dim rsTmp As New ADODB.Recordset
   
   ChkLimitsIsOk = False
   
   '人事處及電腦中心可以代填
   If GetStaffDepartment(strUserNum) = "M51" Or _
      GetStaffDepartment(strUserNum) = "M21" Then
      ChkLimitsIsOk = True
      Exit Function
   End If
   
   '開放當事人的職代及審核主管才可以代填表單
   strSql = "SELECT * FROM ABS001 " & _
            "WHERE B0101='" & Left(Trim(CboB1003), 5) & "' " & _
            "and (B0102='" & strUserNum & "' or B0103='" & strUserNum & "' or B0104='" & strUserNum & "' or B0105='" & strUserNum & "' or B0106='" & strUserNum & "' or B0107='" & strUserNum & "' or B0108='" & strUserNum & "' or B0109='" & strUserNum & "' or B0110='" & strUserNum & "' or B0111='" & strUserNum & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ChkLimitsIsOk = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
