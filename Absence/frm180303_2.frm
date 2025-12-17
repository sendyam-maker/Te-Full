VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm180303_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "異常處理結果資料"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.TextBox textB1401 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   35
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox textB1402 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   34
      Top             =   390
      Width           =   1035
   End
   Begin VB.ComboBox cboB1403 
      Height          =   300
      ItemData        =   "frm180303_2.frx":0000
      Left            =   1380
      List            =   "frm180303_2.frx":000A
      Style           =   2  '單純下拉式
      TabIndex        =   33
      Top             =   690
      Width           =   1125
   End
   Begin VB.ComboBox cboSTime 
      Height          =   300
      ItemData        =   "frm180303_2.frx":001E
      Left            =   4860
      List            =   "frm180303_2.frx":002E
      Style           =   2  '單純下拉式
      TabIndex        =   30
      Top             =   990
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox textB1404 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1020
      Width           =   1035
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "下班資料"
      Height          =   375
      Left            =   5535
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "打卡明細"
      Height          =   375
      Left            =   7830
      TabIndex        =   4
      Top             =   3660
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox textB1411 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      Top             =   3060
      Width           =   405
   End
   Begin VB.TextBox textB1409 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   18
      Top             =   2730
      Width           =   405
   End
   Begin VB.TextBox textB1408 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   16
      Top             =   2430
      Width           =   915
   End
   Begin VB.TextBox textB1405 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1350
      Width           =   405
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   375
      Left            =   7770
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdABS 
      Caption         =   "查詢當日請假資料"
      Height          =   375
      Left            =   6750
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   2145
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   1155
      Left            =   7020
      TabIndex        =   3
      Top             =   4500
      Visible         =   0   'False
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   2037
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   2085
      Left            =   4860
      TabIndex        =   31
      Top             =   3300
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   3678
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "刷卡日期|刷卡時間|人事補登"
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
      _Band(0).Cols   =   3
   End
   Begin MSForms.Label Label23 
      Height          =   195
      Left            =   240
      TabIndex        =   40
      Top             =   5460
      Width           =   6735
      VariousPropertyBits=   27
      Caption         =   "CREATE :                                                    UPDATE : "
      Size            =   "11880;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textB1407 
      Height          =   690
      Left            =   1380
      TabIndex        =   39
      Top             =   1710
      Width           =   7485
      VariousPropertyBits=   -1466939361
      BackColor       =   -2147483633
      ScrollBars      =   3
      Size            =   "13203;1217"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblB1408 
      Height          =   255
      Left            =   2340
      TabIndex        =   38
      Top             =   2430
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2302;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   255
      Left            =   2340
      TabIndex        =   37
      Top             =   90
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2302;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   " 7.颱風假 8.其他)"
      Height          =   180
      Left            =   1830
      TabIndex        =   36
      Top             =   1500
      Width           =   1320
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "當日打卡明細："
      Height          =   180
      Left            =   3570
      TabIndex        =   32
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "批示日期："
      Height          =   180
      Left            =   3930
      TabIndex        =   29
      Top             =   2430
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "上班時段："
      Height          =   180
      Left            =   3930
      TabIndex        =   28
      Top             =   1050
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label LblB1410 
      Caption         =   "LblB1410"
      Height          =   255
      Left            =   4860
      TabIndex        =   27
      Top             =   2430
      Width           =   1305
   End
   Begin VB.Label LblB1413 
      Caption         =   "LblB1413"
      Height          =   255
      Left            =   1380
      TabIndex        =   26
      Top             =   3690
      Width           =   1305
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "確認時間："
      Height          =   180
      Left            =   450
      TabIndex        =   25
      Top             =   3690
      Width           =   900
   End
   Begin VB.Label LblB1412 
      Caption         =   "LblB1412"
      Height          =   255
      Left            =   1380
      TabIndex        =   24
      Top             =   3390
      Width           =   1305
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "確認日期："
      Height          =   180
      Left            =   450
      TabIndex        =   23
      Top             =   3390
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "處理結果："
      Height          =   180
      Left            =   450
      TabIndex        =   22
      Top             =   3060
      Width           =   900
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "（A.系統確認 B.人事室先確認 C.人事室已確認）"
      Height          =   180
      Left            =   1830
      TabIndex        =   21
      Top             =   3060
      Width           =   3825
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "（Y.同意 N.不同意）"
      Height          =   180
      Left            =   1830
      TabIndex        =   20
      Top             =   2730
      Width           =   1635
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "主管批示："
      Height          =   180
      Left            =   450
      TabIndex        =   19
      Top             =   2730
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主管代號："
      Height          =   180
      Index           =   0
      Left            =   450
      TabIndex        =   17
      Top             =   2430
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "未打卡原因："
      Height          =   180
      Left            =   270
      TabIndex        =   15
      Top             =   1710
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "(1.請假 2.遲到/早退 3.忘打卡 4.洽公請主管批示 5.指紋異常 6.因公未打卡 "
      Height          =   180
      Left            =   1830
      TabIndex        =   14
      Top             =   1290
      Width           =   5685
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "個人確認："
      Height          =   180
      Left            =   450
      TabIndex        =   13
      Top             =   1350
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "打卡時間："
      Height          =   180
      Left            =   450
      TabIndex        =   12
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "時　　段："
      Height          =   180
      Left            =   450
      TabIndex        =   11
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      Height          =   180
      Left            =   450
      TabIndex        =   10
      Top             =   390
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   1
      Left            =   420
      TabIndex        =   9
      Top             =   90
      Width           =   930
   End
End
Attribute VB_Name = "frm180303_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Create By Sindy 2013/7/9
Option Explicit

' 變數宣告區
Public m_B1401 As String '員工代號
Public m_B1402 As String '日期
'因一天最多會有二筆異常
Public m_B1403_A As Boolean '上班異常
Public m_B1403_P As Boolean '下班異常
Dim m_PrevForm As Form '前一畫面
Public bolClose As Boolean


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdBack_Click()
   m_PrevForm.Show
   Unload Me
   m_PrevForm.cmdB14_Click
End Sub

'打卡明細
Private Sub cmdDetail_Click()
   If textB1401 = "" Or textB1402 = "" Then Exit Sub
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

   arrGridHeadText = Array("刷卡日期", "刷卡時間", "人事補登", "刷卡機")
   arrGridHeadWidth = Array(800, 800, 800, 1000)
   GrdList.Visible = False
   GrdList.Cols = UBound(arrGridHeadText) + 1
   GrdList.Rows = 2
   For iRow = 0 To GrdList.Cols - 1
      GrdList.row = 0
      GrdList.col = iRow
      GrdList.Text = arrGridHeadText(iRow)
      GrdList.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GrdList.CellAlignment = flexAlignCenterCenter
   Next
   GrdList.Visible = True
End Sub

' 當日打卡明細查詢
Public Function PollRecordQueryData() As Boolean
   Dim stSQL As String
   
   PollRecordQueryData = False
   
   Screen.MousePointer = vbHourglass
   Me.GrdList.MousePointer = flexHourglass
   InitialGridList
   
   stSQL = "select sqldatet(pr01) as 刷卡日期,sqltime6(pr02) as 刷卡時間,decode(pr08,999,'Y','') as 人事補登,decode(OMAN,null,pr09,OMAN) 刷卡機"
   stSQL = stSQL & " from staff,staffcarddata,pollrecord,setSpecMan where scd01(+)=st01 and pr03(+)=scd02 and pr01>0" & _
                    " and st01='" & textB1401 & "' and pr01=" & DBDATE(textB1402) & " and ocode(+)=pr09"
   stSQL = stSQL & " order by pr02 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      Set GrdList.Recordset = RsTemp
      GrdList.row = 1
      PollRecordQueryData = True
'   Else
'      ShowNoData
'      Me.grdList.MousePointer = flexDefault
'      Screen.MousePointer = vbDefault
'      Exit Function
   End If
   Me.GrdList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Function

Private Sub cmdExit_Click()
   Unload Me
   Unload m_PrevForm
End Sub

'請假、出差、加班資料
Private Sub cmdABS_Click()
Dim rsTmp As New ADODB.Recordset
   
   grd2.Clear
   SetGrd2
   If PUB_QueryData_ABS(textB1401, textB1402, rsTmp) = True Then
      Set grd2.Recordset = rsTmp
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
   grd2.Cols = UBound(arrGridHeadText) + 1
   grd2.Rows = 2
   For iRow = 0 To grd2.Cols - 1
      grd2.row = 0
      grd2.col = iRow
      grd2.Text = arrGridHeadText(iRow)
      grd2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd2.CellAlignment = flexAlignCenterCenter
   Next
   'grd2.Visible = True
End Sub

'查詢出缺勤明細資料
Public Sub PubShowNextData()
Dim i As Integer
   
   Me.Enabled = False
   For i = 1 To grd2.Rows - 1
      grd2.col = 0
      grd2.row = i
      If Trim(grd2.Text) = "V" Then
         grd2.Text = ""
         grd2.col = 2 '表單編號
         Screen.MousePointer = vbHourglass
         Me.Hide
         Call frm180301_03.SetParent(Me)
         If grd2.TextMatrix(i, 3) = "1" Then '出缺勤
            frm180301_03.txtB1001 = Pub_RplStr(grd2.Text)
            frm180301_03.QueryData
         Else
            frm180301_03.txtB1003 = Pub_RplStr(grd2.TextMatrix(i, 1))
            frm180301_03.m_SA02 = Pub_RplStr(grd2.TextMatrix(i, 4))
            frm180301_03.m_SA03 = Pub_RplStr(grd2.TextMatrix(i, 5))
            If grd2.TextMatrix(i, 3) = "2" Then '請假
               frm180301_03.QueryData_2
            ElseIf grd2.TextMatrix(i, 3) = "3" Then '加班
               frm180301_03.QueryData_3
            ElseIf grd2.TextMatrix(i, 3) = "4" Then '出差
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

'查詢資料
Private Sub QueryData(strB1403 As String)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   
   '打卡異常資料
   strSql = "select *" & _
            " from abs014" & _
            " where B1401='" & m_B1401 & "'" & _
            " and B1402=" & DBDATE(m_B1402) & _
            " and B1403='" & strB1403 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      textB1401 = "" & rsTmp.Fields("B1401")
      Label12.Caption = GetPrjSalesNM(textB1401)
      textB1402 = ChangeWStringToTString("" & rsTmp.Fields("B1402"))
      If "" & rsTmp.Fields("B1403") = "A" Then
         cboB1403.ListIndex = 0
      Else
         cboB1403.ListIndex = 1
      End If
      If "" & rsTmp.Fields("B1404") <> "" Then
         textB1404.Text = Format(Right("000000" & rsTmp.Fields("B1404"), 6), "00:00:00")
      End If
      textB1405.Text = "" & rsTmp.Fields("B1405")
      If IsNull(rsTmp.Fields("B1406")) Or "" & rsTmp.Fields("B1406") = "" Then
         cboSTime.ListIndex = 0
      Else
         If "" & rsTmp.Fields("B1406") = "800" Then
            cboSTime.ListIndex = 1
         ElseIf "" & rsTmp.Fields("B1406") = "830" Then
            cboSTime.ListIndex = 2
         ElseIf "" & rsTmp.Fields("B1406") = "900" Then
            cboSTime.ListIndex = 3
         End If
      End If
      textB1407.Text = "" & rsTmp.Fields("B1407")
      textB1408.Text = "" & rsTmp.Fields("B1408")
      LblB1408.Caption = GetPrjSalesNM(textB1408)
      LblB1410.Caption = ChangeWStringToTString("" & rsTmp.Fields("B1410"))
      textB1409.Text = "" & rsTmp.Fields("B1409")
      textB1411.Text = "" & rsTmp.Fields("B1411")
      LblB1412.Caption = ChangeWStringToTString("" & rsTmp.Fields("B1412"))
      LblB1413.Caption = Format("" & rsTmp.Fields("B1413"), "00:00:00")
      Call UpdateCUID(rsTmp)
      cboB1403.Enabled = False
      cboSTime.Enabled = False
      
      '當日打卡明細
      Call PollRecordQueryData
   Else
      'ShowNoData
      MsgBox "當日無打卡異常資料！"
      rsTmp.Close
      Set rsTmp = Nothing
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   rsTmp.Close
   Set rsTmp = Nothing
   Screen.MousePointer = vbDefault
   Call GetcmdSelectCaption
   Exit Sub
   
ErrHand:
   Set rsTmp = Nothing
   Screen.MousePointer = vbDefault
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String

   If IsNull(rsSrcTmp.Fields("b1414")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("b1414")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("b1414"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("b1415")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("b1415")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("b1415"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("b1416")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("b1416")) = False Then
         strTemp = Right("000000" & rsSrcTmp.Fields("b1416"), 6)
         strCTime = Format(strTemp, "00:00:00")
      End If
   End If
   
   '設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ")
End Sub

Private Sub GetcmdSelectCaption()
   cmdSelect.Visible = False
   If Left(Trim(cboB1403.Text), 1) = "A" Then
      cmdSelect.Caption = "下班資料"
      If m_B1403_P = True Then cmdSelect.Visible = True
   ElseIf Left(Trim(cboB1403.Text), 1) = "P" Then
      cmdSelect.Caption = "上班資料"
      If m_B1403_A = True Then cmdSelect.Visible = True
   End If
End Sub

Private Sub cmdSelect_Click()
   Call ClearData
   If Trim(cmdSelect.Caption) = "上班資料" Then
      Call QueryData("A")
   ElseIf Trim(cmdSelect.Caption) = "下班資料" Then
      Call QueryData("P")
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Call ClearData
   If m_B1403_A = True Then
      Call QueryData("A")
   Else
      Call QueryData("P")
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Set m_PrevForm = Nothing
   Set frm180303_2 = Nothing
End Sub

'清除欄位值
Private Sub ClearData()
   textB1401.Text = "": Label12.Caption = ""
   textB1402.Text = ""
   cboB1403.ListIndex = 0
   textB1404.Text = ""
   textB1405.Text = ""
   cboSTime.ListIndex = 0
   textB1407.Text = ""
   textB1408.Text = "": LblB1408.Caption = ""
   textB1409.Text = ""
   LblB1410.Caption = ""
   textB1411.Text = ""
   LblB1412.Caption = ""
   LblB1413.Caption = ""
   Label23 = Empty
   grd2.Clear
End Sub
