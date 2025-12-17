VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm180105_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "確認處理方式"
   ClientHeight    =   5730
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8950
   Begin VB.TextBox textB1404 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   270
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1230
      Width           =   1035
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "打卡明細"
      Height          =   375
      Left            =   7470
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox textB1402 
      Height          =   270
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   1
      Top             =   510
      Width           =   1035
   End
   Begin VB.ComboBox cboB1403 
      Height          =   300
      ItemData        =   "frm180105_1.frx":0000
      Left            =   1800
      List            =   "frm180105_1.frx":000A
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   840
      Width           =   1125
   End
   Begin VB.ComboBox cboSTime 
      Height          =   300
      ItemData        =   "frm180105_1.frx":001E
      Left            =   5280
      List            =   "frm180105_1.frx":002E
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1260
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox textB1409 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   22
      Top             =   3210
      Width           =   405
   End
   Begin VB.TextBox textB1408 
      Height          =   270
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2880
      Width           =   1035
   End
   Begin VB.TextBox textB1405 
      Height          =   285
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1560
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   375
      Left            =   4350
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面"
      Height          =   375
      Left            =   6980
      TabIndex        =   10
      Top             =   120
      Width           =   945
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   375
      Left            =   7980
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox textB1401 
      Height          =   270
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
   Begin VB.CommandButton cmdABS 
      Caption         =   "查詢當日請假資料"
      Height          =   375
      Left            =   5260
      TabIndex        =   9
      Top             =   120
      Width           =   1665
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   1155
      Left            =   7020
      TabIndex        =   12
      Top             =   4230
      Visible         =   0   'False
      Width           =   1830
      _ExtentX        =   3246
      _ExtentY        =   2028
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   1725
      Left            =   1800
      TabIndex        =   29
      Top             =   3570
      Width           =   3210
      _ExtentX        =   5644
      _ExtentY        =   3052
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
      Left            =   60
      TabIndex        =   33
      Top             =   5490
      Width           =   3945
      VariousPropertyBits=   27
      Caption         =   "CREATE :                      UPDATE : "
      Size            =   "6959;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textB1407 
      Height          =   900
      Left            =   1800
      TabIndex        =   6
      Top             =   1920
      Width           =   7095
      VariousPropertyBits=   -1466939365
      ScrollBars      =   3
      Size            =   "12515;1587"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblB1408 
      Height          =   255
      Left            =   2880
      TabIndex        =   31
      Top             =   2910
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
      Left            =   2880
      TabIndex        =   32
      Top             =   240
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2302;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "當日打卡明細："
      Height          =   180
      Left            =   510
      TabIndex        =   30
      Top             =   3630
      Width           =   1260
   End
   Begin VB.Label Label13 
      Caption         =   "備註：若個人確認為”4”將發送E-Mail通知主管。"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4080
      TabIndex        =   28
      Top             =   5460
      Width           =   4845
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "批示日期："
      Height          =   180
      Left            =   4380
      TabIndex        =   27
      Top             =   2910
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "上班時段："
      Height          =   180
      Left            =   4350
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label LblB1410 
      Caption         =   "LblB1410"
      Height          =   255
      Left            =   5310
      TabIndex        =   25
      Top             =   2880
      Width           =   1305
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "（Y.同意 N.不同意）"
      Height          =   180
      Left            =   2250
      TabIndex        =   24
      Top             =   3270
      Width           =   1635
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "主管批示："
      Height          =   180
      Left            =   870
      TabIndex        =   23
      Top             =   3270
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主管代號："
      Height          =   180
      Index           =   0
      Left            =   870
      TabIndex        =   21
      Top             =   2910
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "未打卡原因："
      Height          =   180
      Left            =   690
      TabIndex        =   20
      Top             =   1980
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "（1.請假 2.遲到/早退 3.忘打卡 4.洽公請主管批示 8.其他）"
      Height          =   180
      Left            =   2250
      TabIndex        =   19
      Top             =   1620
      Width           =   4500
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "個人確認："
      Height          =   180
      Left            =   870
      TabIndex        =   18
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "打卡時間："
      Height          =   180
      Left            =   870
      TabIndex        =   17
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "時　　段："
      Height          =   180
      Left            =   870
      TabIndex        =   16
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      Height          =   180
      Left            =   870
      TabIndex        =   15
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   1
      Left            =   840
      TabIndex        =   14
      Top             =   210
      Width           =   930
   End
End
Attribute VB_Name = "frm180105_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Create By Sindy 2013/6/21
Option Explicit

' 變數宣告區
Public m_B1401 As String '員工代號
Public m_B1402 As String '日期
Public m_B1403 As String '打卡類別
Public bolClose As Boolean
Dim m_B1408 As String


Private Sub cmdBack_Click()
   'Add By Sindy 2025/11/11
   If cmdOK.Tag = "Y" Then
      intI = MsgBox("尚未完成 [確認] 動作！資料是否已輸入完畢，要進行確認嗎？" & vbCrLf & vbCrLf & "【是】：執行確認鍵" & vbCrLf & "【否】：回畫面繼續操作" & vbCrLf & "【取消】：結束離開", vbYesNoCancel)
      If intI = vbYes Then
         Call cmdok_Click
         Exit Sub
      ElseIf intI = vbNo Then
         Exit Sub
      End If
   End If
   '2025/11/11 END
   Unload Me
   frm180105.Show
   frm180105.cmdok_Click
End Sub

Private Sub cmdExit_Click()
   'Add By Sindy 2025/11/11
   If cmdOK.Tag = "Y" Then
      intI = MsgBox("尚未完成 [確認] 動作！資料是否已輸入完畢，要進行確認嗎？" & vbCrLf & vbCrLf & "【是】：執行確認鍵" & vbCrLf & "【否】：回畫面繼續操作" & vbCrLf & "【取消】：結束離開", vbYesNoCancel)
      If intI = vbYes Then
         Call cmdok_Click
         Exit Sub
      ElseIf intI = vbNo Then
         Exit Sub
      End If
   End If
   '2025/11/11 END
   Unload Me
   Unload frm180105
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

   arrGridHeadText = Array("刷卡日期", "刷卡時間", "人事補登")
   arrGridHeadWidth = Array(1000, 1000, 800)
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
   
   stSQL = "select sqldatet(pr01) as 刷卡日期,sqltime6(pr02) as 刷卡時間,decode(pr08,999,'Y','') as 人事補登"
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
Private Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strData As String, strTemp As Variant, i As Integer
Dim strB1104 As String, m_bolIsRest1Day As Boolean
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   
   '打卡異常資料
   strSql = "select *" & _
            " from abs014" & _
            " where B1401='" & m_B1401 & "'" & _
            " and B1402=" & DBDATE(m_B1402) & _
            " and B1403='" & m_B1403 & "'"
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
         textB1404.Text = Format(Right("000000" & rsTmp.Fields("B1404"), 6), "##:##:##")
      End If
      textB1405.Text = "" & rsTmp.Fields("B1405")
      textB1407.Text = "" & rsTmp.Fields("B1407")
'      If IsNull(rsTmp.Fields("B1406")) Or "" & rsTmp.Fields("B1406") = "" Then
'         cboSTime.ListIndex = 0
'      Else
'         If "" & rsTmp.Fields("B1406") = "800" Then
'            cboSTime.ListIndex = 1
'         ElseIf "" & rsTmp.Fields("B1406") = "830" Then
'            cboSTime.ListIndex = 2
'         ElseIf "" & rsTmp.Fields("B1406") = "900" Then
'            cboSTime.ListIndex = 3
'         End If
'      End If
      '主管
      m_B1408 = "" & rsTmp.Fields("B1408")
      textB1408.Text = "" & rsTmp.Fields("B1408")
      '*** 主管預設為第一階主管 ***
      If textB1408.Text = "" Then
         strData = GetABS001_2(textB1401, 1)
         If strData <> "" Then
            strTemp = Split(strData, ",")
            For i = 0 To UBound(strTemp)
               textB1408 = strTemp(i)
               strB1104 = strTemp(i)
               'Modify By Sindy 2018/9/7 Mark 慧汶:打卡異常處理,王副總及雅娟不在,應是由游經理處理為何顯示主管代號為玲玲,CFP程序又不歸她管
'               If CheckPerCurrRestReturnPer(strB1104, textB1401, m_bolIsRest1Day) = True Then '主管為休假時，轉職代
'                  textB1408 = Left(Trim(strB1104), 5)
'               End If
               Exit For
            Next i
         End If
      End If
      '*** END
      LblB1408.Caption = GetPrjSalesNM(textB1408)
      LblB1410.Caption = ChangeWStringToTString("" & rsTmp.Fields("B1410"))
      textB1409.Text = "" & rsTmp.Fields("B1409")
      '已輸入個人確認為4時,不可修改
      If Trim("" & rsTmp.Fields("B1405")) = "4" Then
         textB1405.Enabled = False
      End If
      '預設值 : 如果上班打卡時間是大於等於9:30時,則一定要請假
      'Modify By Sindy 2014/10/7 決策主管除外 +And m_B1401 <> "81040" And m_B1401 <> "68001" And m_B1401 <> "94007" And m_B1401 <> "68006" And m_B1401 <> "68009" And m_B1401 <> "71011" And m_B1401 <> "67001"
      'Modify By Sindy 2023/7/26 +清潔人員B2024劉美英
      If textB1405.Text = "" And textB1404 <> "" Then
         If "" & rsTmp.Fields("B1403") = "A" And Val(Replace(Left(textB1404, Len(textB1404) - 2), ":", "")) >= 930 _
            And (m_B1401 <> "99029" And m_B1401 <> "B2024" And m_B1401 <> "96006" And m_B1401 <> "81040" And m_B1401 <> "68001" And m_B1401 <> "94007" And m_B1401 <> "68006" And m_B1401 <> "68009" And m_B1401 <> "71011" And m_B1401 <> "67001") Then
            textB1405.Text = "1"
            textB1405.Enabled = False
         End If
      End If
'      textB1411.Text = "" & rsTmp.Fields("B1411")
'      LblB1412.Caption = ChangeWStringToTString("" & rsTmp.Fields("B1412"))
'      LblB1413.Caption = "" & rsTmp.Fields("B1413")
      Call UpdateCUID(rsTmp)
      If textB1405.Enabled = True Then
         textB1405.TabIndex = 0
      Else
         textB1407.TabIndex = 0
      End If
      '已有批示結果時,不可再異動資料
      If Trim(textB1409.Text) <> "" Then
'         cboSTime.Enabled = False
         textB1405.Enabled = False
         textB1407.Enabled = False
         textB1408.Enabled = False
         cmdOK.Enabled = False
         cmdBack.TabIndex = 0
      Else
         Call textB1405_Validate(True)
      End If
      
      '當日打卡明細
      Call PollRecordQueryData
   Else
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   rsTmp.Close
   Set rsTmp = Nothing
   Screen.MousePointer = vbDefault
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

'確定
Private Sub cmdok_Click()
Dim strSubject As String, strContent As String
   
   If CheckDataValid = False Then Exit Sub
   If TxtValidate = False Then Exit Sub
   
On Error GoTo ErrHand
   
   cmdOK.Tag = "" 'Add By Sindy 2025/11/11
   cnnConnection.BeginTrans
   '",B1406=" & CNULL(Format(cboSTime.Text, "hhmm"))
   strSql = "UPDATE ABS014 SET B1405='" & textB1405 & "'" & _
                             ",B1407='" & textB1407 & "'"
   If textB1405 = "4" Then '主管批示
      strSql = strSql & ",B1408='" & textB1408 & "'"
   End If
   strSql = strSql & " WHERE B1401='" & textB1401 & "' and B1402=" & DBDATE(textB1402) & " and B1403='" & Left(Trim(cboB1403), 1) & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   
   If textB1405 = "4" And textB1408 <> "" And m_B1408 <> textB1408 Then
      '發Mail通知主管批示
      strSubject = Label12 & " " & ChangeWStringToTDateString(DBDATE(textB1402)) & " " & IIf(Left(Trim(cboB1403), 1) = "A", "上", "下") & "班未打卡，請主管簽核！"
      strContent = strSubject & vbCrLf & vbCrLf
'      strContent = strContent & "員　　　工：" & textB1401 & " " & Label12 & vbCrLf
'      strContent = strContent & "未打卡日期：" & ChangeWStringToTDateString(DBDATE(textB1402)) & vbCrLf
      strContent = strContent & "未打卡原因：" & textB1407 & vbCrLf
      strContent = strContent & "處　　　理：請至案件管理系統（一般作業->出缺勤作業->簽核->打卡異常主管處理）中，進行處理。" & vbCrLf
      PUB_SendMail strUserNum, textB1408, "", strSubject, strContent, , , , , , , , , , True
   End If
   
   Call cmdBack_Click
   Exit Sub
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "資料存檔失敗！" & vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Call ClearData
   Call QueryData
   textB1401.Enabled = False
   textB1402.Enabled = False
   cboB1403.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180105_1 = Nothing
End Sub

'清除欄位值
Private Sub ClearData()
   textB1401.Text = "": Label12.Caption = ""
   textB1402.Text = ""
   cboB1403.ListIndex = 0
   textB1404.Text = ""
   textB1405.Text = ""
'   cboSTime.ListIndex = 0
   textB1407.Text = ""
   textB1408.Text = "": LblB1408.Caption = "": m_B1408 = ""
   textB1409.Text = ""
   LblB1410.Caption = ""
'   textB1411.Text = ""
'   LblB1412.Caption = ""
'   LblB1413.Caption = ""
   Label23 = Empty
   grd2.Clear
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   CheckDataValid = False
   
   If textB1405 = "" Then
      MsgBox "個人確認不可空白 !!!"
      textB1405.SetFocus
      Exit Function
   End If
'   If Left(Trim(cboB1403), 1) = "A" Then
'      If cboSTime = "" Then
'         MsgBox "上班時段不可空白 !!!"
'         cboSTime.SetFocus
'         Exit Function
'      End If
'   End If
   If textB1405 = "4" And textB1407 = "" Then
      MsgBox "未打卡原因不可空白 !!!"
      textB1407.SetFocus
      Exit Function
   'Add By Sindy 2013/8/30
   ElseIf textB1405 = "8" And textB1407 = "" Then
      MsgBox "請輸入未打卡原因 !!!"
      textB1407.SetFocus
      Exit Function
   '2013/8/30 END
   End If
   
   cmdOK.Tag = "Y" 'Add By Sindy 2025/11/11
   If textB1405 = "1" Then '填請假,是否有假單存在
      If CheckIsPersonRestSector(textB1401, DBDATE(textB1402), "00:00", DBDATE(textB1402), "24:00", "") = False Then
         MsgBox "無假單，請輸入 !!!"
         If textB1405.Enabled = True Then textB1405.SetFocus
         Call textB1405_LostFocus
         Exit Function
      End If
   ElseIf textB1405 = "4" Then '請主管批示,必須輸入主管的員工代號
      If textB1408 = "" Then
         MsgBox "請輸入批示主管的員工代號 !!!"
         textB1408.SetFocus
         Exit Function
      ElseIf textB1408 = textB1401 Then
         MsgBox "批示主管不可為自己 !!!"
         textB1408.SetFocus
         Exit Function
      End If
   End If
   
   CheckDataValid = True
End Function

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean

   TxtValidate = False
   
   If Me.textB1405.Enabled = True Then
      Cancel = False
      textB1405_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textB1408.Enabled = True Then
      Cancel = False
      textB1408_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add by Sindy 2021/5/28 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/5/28 END
   
   TxtValidate = True
End Function

Private Sub textB1405_GotFocus()
   InverseTextBox textB1405
   CloseIme
End Sub

Private Sub textB1405_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textB1405_LostFocus()
Dim strB1001 As String
Dim Cancel As Boolean 'Add By Sindy 2018/9/7
   
   If textB1405 = "1" Then '請假
      '檢查是否有主管代填假單
      strSql = "select * from abs010 where b1003='" & textB1401 & "' and " & DBDATE(textB1402) & ">=b1004 and b1006<=" & DBDATE(textB1402) & " and b1018='" & 主管代填 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strB1001 = RsTemp.Fields("b1001")
         Call frm180102.SetParent(Me)
         frm180102.Hide
         frm180102.txtB1001 = strB1001
         frm180102.Show
         frm180102.QueryData
         If frm180102.txtB1001 <> "" And frm180102.m_B1017 = "" Then
            frm180102.TxtValidate
         End If
         cmdOK.Tag = "Y" 'Add By Sindy 2025/11/11
         Me.Hide
         Exit Sub
      End If
      '檢查無請假資料時,自動開啟請假表單
      If CheckIsPersonRestSector(textB1401, DBDATE(textB1402), "00:00", DBDATE(textB1402), "24:00", "") = False Then
         Call frm180102.SetParent(Me)
         frm180102.Hide
         frm180102.txtB1001 = ""
         frm180102.txtB1004 = textB1402
         frm180102.txtB1006 = textB1402
         frm180102.Show
         cmdOK.Tag = "Y" 'Add By Sindy 2025/11/11
         Me.Hide
      End If
   'Add By Sindy 2018/9/7
   ElseIf textB1405 = "4" Then '洽公請主管批示
      If Me.textB1408.Enabled = True Then
         Cancel = False
         textB1408_Validate Cancel
         If Cancel = True Then
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub textB1405_Validate(Cancel As Boolean)
   If textB1405 = "" Then Exit Sub
   If textB1405 <> "" Then
      Select Case textB1405
         'Modify By Sindy 2013/8/30 +8
         Case 1, 2, 3, 4, 8
            If textB1405 <> "4" Then
               textB1408.Enabled = False
               textB1408.BackColor = &H8000000F
'               textB1408 = ""
'               LblB1408.Caption = ""
            Else
               textB1408.Enabled = True
               textB1408.BackColor = &H80000005
            End If
         Case Else
            'Modify By Sindy 2013/8/30 +8
            MsgBox "個人確認只可輸入1~4 或 8 !!!"
            Call textB1405_GotFocus
            Cancel = True
            Exit Sub
      End Select
   End If
   CloseIme
End Sub

Private Sub textB1407_GotFocus()
   InverseTextBox textB1407
   OpenIme
End Sub

Private Sub textB1407_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 textB1407
End Sub

'Private Sub cboSTime_GotFocus()
''   InverseTextBox cboSTime
'End Sub
'
'Private Sub cboSTime_KeyPress(KeyAscii As Integer)
'   KeyAscii = Pub_NumAscii(KeyAscii)
'End Sub
'
'Private Sub cboSTime_Validate(Cancel As Boolean)
'Dim bolChkOK As Boolean
'Dim i As Integer
'
'   If cboSTime <> "" Then
'      bolChkOK = False
'      For i = 0 To cboSTime.ListCount - 1
'        If cboSTime.Text = cboSTime.List(i) Then
'           bolChkOK = True
'           Exit For
'        End If
'      Next i
'      If bolChkOK = False Then
'         Call cboSTime_GotFocus
'         MsgBox "輸入錯誤 !!!", vbExclamation + vbOKOnly
'         Cancel = True
'         Exit Sub
'      End If
'   End If
'End Sub

Private Sub textB1408_GotFocus()
   InverseTextBox textB1408
End Sub

Private Sub textB1408_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textB1408_Validate(Cancel As Boolean)
Dim Rs As New ADODB.Recordset
   
   If textB1408.Text = "" Then LblB1408 = ""
   
   If textB1408 <> "" Then
      ' 檢查員工編號規則
      If ChkStaffST04(textB1408) Then
         Call textB1408_GotFocus
         Cancel = True
         Exit Sub
      End If
      LblB1408 = GetStaffName(textB1408, True)
      If Label12 = "" Then
         MsgBox "主管的員工編號錯誤！查無此員工！", vbInformation
         Call textB1408_GotFocus
         Cancel = True
         Exit Sub
      End If
      
'      'Add By Sindy 2016/10/4 若簽核主管休假,彈訊息讓當事人知道
'      If textB1405 = "4" Then '洽公請主管批示
'         If textB1408.Tag <> textB1408.Text Then
'            If CheckIsPersonRest(textB1408, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = True Then
'               MsgBox LblB1408 & "目前休假中！"
'               textB1408.Tag = textB1408.Text
'            End If
'         End If
'      End If
      'Add By Sindy 2016/10/4 若簽核主管休假,彈訊息讓當事人知道
      If textB1405 = "4" Then '洽公請主管批示
         If textB1408.Tag <> textB1408.Text Then
            If CheckIsPersonRest(textB1408, strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = True Then
               'Add By Sindy 2018/9/7
               If MsgBox(LblB1408 & "目前休假中，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2, "提醒！") = vbNo Then
                  textB1408.Text = ""
                  textB1408.SetFocus
                  Exit Sub
               End If
               textB1408.Tag = textB1408.Text
               '2018/9/7 END
            End If
         End If
      End If
   End If
End Sub
