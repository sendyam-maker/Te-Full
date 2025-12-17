VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180203_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "每月出缺勤統計確認"
   ClientHeight    =   5750
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8960
   Begin VB.CommandButton cmdOK 
      Caption         =   "全選(&A)"
      Height          =   360
      Index           =   0
      Left            =   4830
      TabIndex        =   3
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "快速簽核(&E)"
      Height          =   360
      Index           =   4
      Left            =   3000
      TabIndex        =   4
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "畫面更新(&Q)"
      Height          =   360
      Index           =   1
      Left            =   5670
      TabIndex        =   0
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "逐筆簽核(&O)"
      Height          =   360
      Index           =   2
      Left            =   6870
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Index           =   3
      Left            =   8055
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm180203_1.frx":0000
      Height          =   4725
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   8835
      _ExtentX        =   15593
      _ExtentY        =   8326
      _Version        =   393216
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
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label5 
      Caption         =   "註3．有出缺勤資料”相關欄位”才會顯示出來。(一樣需要確認)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   690
      Width           =   5805
   End
   Begin VB.Label Label2 
      Caption         =   "註2．按＜畫面更新＞按鈕，顯示最新資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   450
      Width           =   5805
   End
   Begin VB.Label Label3 
      Caption         =   "淺橘色"
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   2250
      TabIndex        =   8
      Top             =   180
      Width           =   645
   End
   Begin VB.Label Label4 
      Caption         =   "淺綠色"
      ForeColor       =   &H0000C000&
      Height          =   225
      Left            =   1050
      TabIndex        =   7
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "註1．當月：淺綠色   年度：淺橘色"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   60
      TabIndex        =   6
      Top             =   180
      Width           =   2835
   End
End
Attribute VB_Name = "frm180203_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/11/17
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Public cmdState As Integer '紀錄作用按鍵


Public Sub PubShowNextData()
Dim bolSelect As Boolean, strB1303 As String, intB1303Seqno As Integer
   
   Select Case cmdState
      Case 0 '全選
         GRD1.Visible = False
         If GRD1.Rows > 1 Then
            If GRD1.TextMatrix(1, 1) <> "" Then
               For j = 1 To GRD1.Rows - 1
                  GRD1.col = 0
                  GRD1.row = j
                  GRD1.Text = "V"
                  For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = &HFFC0C0
                  Next i
               Next j
            End If
         End If
         GRD1.Visible = True
      Case 1 '查詢
         If QueryData = False Then ShowNoData
      Case 2 '逐筆簽核
         bolSelect = False
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then
               bolSelect = True
               GRD1.col = 0
               'GRD1.Text = "" 'Modify By Sindy 2012/2/13
               For j = 0 To GRD1.Cols - 1
                  GRD1.col = j
                  GRD1.CellBackColor = QBColor(15)
               Next j
               frm160201.intChoose = 1
               frm160201.Hide
               'Modify By Sindy 2020/2/5
               'frm160201.txt1(0) = Val(GRD1.TextMatrix(i, 48)) - 191100
               frm160201.txt1(0) = Val(GRD1.TextMatrix(i, 50)) - 191100
               '2020/2/5 END
               frm160201.txt1(1) = ""
               frm160201.txt1(2) = ""
               frm160201.txt1(3) = GRD1.TextMatrix(i, 2)
               frm160201.txt1(4) = GRD1.TextMatrix(i, 2)
               Me.Hide
               Call frm160201.cmdok_Click(0)
            End If
         Next i
         Me.Show
         If bolSelect = True Then chkClearGrd1 'Add By Sindy 2012/2/13
         Me.Enabled = True
'         Call QueryData
      Case 3 '結束
         Unload Me
      Case 4 '一次簽核
         bolSelect = False
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then
               bolSelect = True
               '逐筆簽核
               Screen.MousePointer = vbHourglass
               cnnConnection.BeginTrans
               
               'Modify By Sindy 2020/2/5
               'strB1303 = GRD1.TextMatrix(i, 49) '目的：為讓目前處理人員,可以Run到Do While迴圈
               'Modify By Sindy 2025/11/12
               'strB1303 = GRD1.TextMatrix(i, 51) '目的：為讓目前處理人員,可以Run到Do While迴圈
               strB1303 = GRD1.TextMatrix(i, 53) '目的：為讓目前處理人員,可以Run到Do While迴圈
               '2020/2/5 END
               '檢查下一處理人員是否離職,若是,則移轉下一處理人員
               'Modify By Sindy 2020/2/5
               'Do While strB1303 = GRD1.TextMatrix(i, 49) Or (strB1303 <> "" And ChkStaffST04(strB1303, False) = True)
               'Modify By Sindy 2025/11/12
               'Do While strB1303 = GRD1.TextMatrix(i, 51) Or (strB1303 <> "" And ChkStaffST04(strB1303, False) = True)
               Do While strB1303 = GRD1.TextMatrix(i, 53) Or (strB1303 <> "" And ChkStaffST04(strB1303, False) = True)
               '2020/2/5 END
                  '審核主管確認
                  intB1303Seqno = GetCurrB1303Seqno("04", GRD1.TextMatrix(i, 2), strB1303)
                  If intB1303Seqno = 1 Then
                     strSql = "Update ABS013 set B1304='Y' WHERE B1301='04' and B1302='" & GRD1.TextMatrix(i, 2) & "' "
                     cnnConnection.Execute strSql
                  ElseIf intB1303Seqno = 2 Then
                     strSql = "Update ABS013 set B1305='Y' WHERE B1301='04' and B1302='" & GRD1.TextMatrix(i, 2) & "' "
                     cnnConnection.Execute strSql
                  ElseIf intB1303Seqno = 3 Then
                     strSql = "Update ABS013 set B1306='Y' WHERE B1301='04' and B1302='" & GRD1.TextMatrix(i, 2) & "' "
                     cnnConnection.Execute strSql
                  ElseIf intB1303Seqno = 4 Then
                     strSql = "Update ABS013 set B1307='Y' WHERE B1301='04' and B1302='" & GRD1.TextMatrix(i, 2) & "' "
                     cnnConnection.Execute strSql
                  End If
                  '讀取下一處理人員
                  strB1303 = GetNextB1303("04", GRD1.TextMatrix(i, 2))
               Loop
               
               If strB1303 = "" Then
                  '已無下一處理人員時,代表已確認完畢即可刪除該筆資料
                  strSql = "DELETE FROM ABS013 WHERE B1301='04' and B1302='" & GRD1.TextMatrix(i, 2) & "' "
               Else
                  '送至下一處理人員
                  strSql = "Update ABS013 set B1303='" & strB1303 & "' WHERE B1301='04' and B1302='" & GRD1.TextMatrix(i, 2) & "' "
               End If
               cnnConnection.Execute strSql
               
               cnnConnection.CommitTrans
               Screen.MousePointer = vbDefault
            End If
         Next i
         'Add By Sindy 2012/2/13
         Me.Show
         If bolSelect = True Then chkClearGrd1
         Me.Enabled = True
         '2012/2/13 End
         'If bolSelect = True Then Call QueryData
      Case Else
   End Select
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox "全部簽核失敗！" & vbCrLf & Err.Description
End Sub

'Add By Sindy 2012/2/13
Private Sub chkClearGrd1()
   '以防簽核人員已做確認,重新檢核一下目前資料列裡的資料,已簽核資料必須隱藏
   'Modify By Sindy 2023/12/22
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "Select A0922 部門別,ST01 員工代號,ST02 姓名,B1314,B1303 " & _
               "From ABS013, Staff, ACC090NEW " & _
               "Where B1301='04' and B1303='" & strUserNum & "' and B1302=ST01(+) and ST93=A0921(+) " & _
               "order by ST01 asc "
   Else
   '2023/12/22 END
      strSql = "Select A0902 部門別,ST01 員工代號,ST02 姓名,B1314,B1303 " & _
               "From ABS013, Staff, ACC090 " & _
               "Where B1301='04' and B1303='" & strUserNum & "' and B1302=ST01(+) and ST03=A0901(+) " & _
               "order by ST01 asc "
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      For i = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(i, 0) = "V" Then
            GRD1.TextMatrix(i, 0) = ""
            RsTemp.MoveFirst
            For j = 1 To RsTemp.RecordCount
               If GRD1.TextMatrix(i, 2) = RsTemp.Fields(1) Then
                  GoTo ExitReadNextRow
               End If
               RsTemp.MoveNext
            Next j
            GRD1.RowHeight(i) = 0
         End If
ExitReadNextRow:
      Next i
      Call SetColor
   Else
      GRD1.Clear
      SetGrd
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   
   PubShowNextData
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, StrST01 As String
Dim strSDate As String, strEDate As String, strSYDate As String, strEYDate As String
'Dim dblHour(18) As Double, dblCnt(18) As Double
'Dim dblHour(22) As Double, dblCnt(22) As Double...移到basPerson共用變數區
Dim intItem As Integer, strText As String, dblDay As Double, dblHours As Double
Dim strAdd1 As String, strAdd2 As String
   
   m_blnColOrderAsc = True
   QueryData = True
   GRD1.Clear
   SetGrd
   GRD1.FixedCols = 0
   
   Screen.MousePointer = vbHourglass
   '             4        5      6      7      8      9      10       11     12     13     14     15       16     17       18     19     20           21             22          23       24       25           26           27
   strAdd1 = ",0 忘打卡,0 遲到,0 曠職,0 事假,0 病假,0 公假,0 特別假,0 出差,0 加班,0 婚假,0 產假,0 流產假,0 喪假,0 公傷假,0 補休,0 其他,0 扣年終產假,0 扣年終流產假,0 陪產檢及陪產假,0 生理假,0 產檢假,0 家庭照顧假,0 防疫照顧假,0 天災不給薪"
   '             28       29     30     31     32     33     34       35     36     37     38     39       40     41       42     43     44           45             46          47       48       49           50           51
   strAdd2 = ",0 忘打卡,0 遲到,0 曠職,0 事假,0 病假,0 公假,0 特別假,0 出差,0 加班,0 婚假,0 產假,0 流產假,0 喪假,0 公傷假,0 補休,0 其他,0 扣年終產假,0 扣年終流產假,0 陪產檢及陪產假,0 生理假,0 產檢假,0 家庭照顧假,0 防疫照顧假,0 天災不給薪"
   'Modify By Sindy 2023/12/22
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "Select ' ' as V,A0922 部門別,ST01 員工代號,ST02 姓名" & strAdd1 & strAdd2 & ",B1314,B1303 " & _
               " From ABS013,Staff,ACC090NEW" & _
               " Where B1301='04' and B1303='" & strUserNum & "' and B1302=ST01(+) and ST93=A0921(+)" & _
               " order by ST01 asc"
   Else
   '2023/12/22 END
      strSql = "Select ' ' as V,A0902 部門別,ST01 員工代號,ST02 姓名" & strAdd1 & strAdd2 & ",B1314,B1303 " & _
               " From ABS013,Staff,ACC090" & _
               " Where B1301='04' and B1303='" & strUserNum & "' and B1302=ST01(+) and ST03=A0901(+)" & _
               " order by ST01 asc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      GRD1.FixedCols = 4
      
      strSDate = Val(rsTmp.Fields("B1314")) * 100 + 1
      strEDate = Val(rsTmp.Fields("B1314")) * 100 + 31
      strSYDate = Left(Trim(strSDate), 4) & "0101"
      strEYDate = strEDate
      
      For i = 1 To rsTmp.RecordCount
         StrST01 = CheckStr(GRD1.TextMatrix(i, 2))
         strSql = " and ST01='" & StrST01 & "' "
         Call Pub_GetSpecWorkHour(StrST01, strSDate) 'Add By Sindy 2012/7/9 上班時數為特殊者
         '取得各假別時數-當月統計
         If PUB_GetAbsenceHour(strSql, strSDate, strEDate, dblHour(), dblCnt()) = True Then
            GRD1.TextMatrix(i, 4) = IIf(Val(dblHour(1)) > 0, dblHour(1) & "次", "") '忘打卡
            GRD1.TextMatrix(i, 5) = IIf(Val(dblHour(2)) > 0, dblHour(2) & "次", "") '遲到
            GRD1.TextMatrix(i, 6) = IIf(Val(dblHour(3)) > 0, dblHour(3) & "分", "") '曠職 'Add By Sindy 2025/11/17
            GRD1.TextMatrix(i, 12) = IIf(Val(dblHour(16)) > 0, dblHour(16) & "時", "") '加班
            'For j = 6 To 27 '26 '25
            For j = 7 To 27 '26 '25
               intItem = 0 'Add By Sindy 2019/10/9
               ' ”\”整除運算子不可使用於有小數位之數值, 因此先將數值*10做運算
               'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
               'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
               'If j = 6 Then intItem = 3 'Modify By Sindy 2025/11/17 mark
               If j = 7 Then intItem = 5
               If j = 8 Then intItem = 6
               If j = 9 Then intItem = 7
               If j = 10 Then intItem = 8
               If j = 11 Then intItem = 4
               'Add By Sindy 2019/10/9 24~36
               If j = 13 Then intItem = 9 '婚假
               If j = 14 Then intItem = 10 '產假
               If j = 15 Then intItem = 11 '流產假
               If j = 16 Then intItem = 12 '喪假
               If j = 17 Then intItem = 13 '公傷假
               If j = 18 Then intItem = 14 '補休
               If j = 19 Then intItem = 15 '其他
               If j = 20 Then intItem = 17 '扣年終產假
               If j = 21 Then intItem = 18 '扣年終流產假
               If j = 22 Then intItem = 19 '陪產檢及陪產假
               If j = 23 Then intItem = 20 '生理假
               If j = 24 Then intItem = 21 '產檢假
               If j = 25 Then intItem = 22 '家庭照顧假
               If j = 26 Then intItem = 24 '防疫照顧假 Add By Sindy 2020/2/5
               If j = 27 Then intItem = 25 '天災不給薪 Add By Sindy 2025/11/13
               'Modify By Sindy 2012/7/9 上班時數為特殊者
'               If strST01 = "99029" Then
'                  dblDay = (dblHour(intItem) * 10) \ (5 * 10)
'                  dblHours = Round(dblHour(intItem) - (dblDay * 5), 1)
               If intItem > 0 Then
               '2019/10/9 END
                  dblDay = (dblHour(intItem) * 10) \ (PUB_intWkHour * 10)
                  dblHours = Round(dblHour(intItem) - (dblDay * PUB_intWkHour), 1)
                  strText = IIf(dblDay > 0, dblDay & "日", "")
                  strText = strText & IIf(dblHours > 0, dblHours & "時", "")
                  GRD1.TextMatrix(i, j) = strText
               End If
            Next j
         End If
         '取得各假別時數-年度統計
         If PUB_GetAbsenceHour(strSql, strSYDate, strEYDate, dblHour(), dblCnt()) = True Then
            GRD1.TextMatrix(i, 28) = IIf(Val(dblHour(1)) > 0, dblHour(1) & "次", "") '忘打卡
            GRD1.TextMatrix(i, 29) = IIf(Val(dblHour(2)) > 0, dblHour(2) & "次", "") '遲到
            GRD1.TextMatrix(i, 30) = IIf(Val(dblHour(3)) > 0, dblHour(3) & "分", "") '曠職 'Add By Sindy 2025/11/17
            GRD1.TextMatrix(i, 36) = IIf(Val(dblHour(16)) > 0, dblHour(16) & "時", "") '加班
            'For j = 28 To 47
            'For j = 29 To 49
            'For j = 30 To 51
            For j = 31 To 51
               intItem = 0 'Add By Sindy 2019/10/9
               ' ”\”整除運算子不可使用於有小數位之數值, 因此先將數值*10做運算
               'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
               'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
               'If j = 30 Then intItem = 3 'Modify By Sindy 2025/11/17 mark
               If j = 31 Then intItem = 5
               If j = 32 Then intItem = 6
               If j = 33 Then intItem = 7
               If j = 34 Then intItem = 8
               If j = 35 Then intItem = 4
               'Add By Sindy 2019/10/9 37~49
               If j = 37 Then intItem = 9 '婚假
               If j = 38 Then intItem = 10 '產假
               If j = 39 Then intItem = 11 '流產假
               If j = 40 Then intItem = 12 '喪假
               If j = 41 Then intItem = 13 '公傷假
               If j = 42 Then intItem = 14 '補休
               If j = 43 Then intItem = 15 '其他
               If j = 44 Then intItem = 17 '扣年終產假
               If j = 45 Then intItem = 18 '扣年終流產假
               If j = 46 Then intItem = 19 '陪產檢及陪產假
               If j = 47 Then intItem = 20 '生理假
               If j = 48 Then intItem = 21 '產檢假
               If j = 49 Then intItem = 22 '家庭照顧假
               If j = 50 Then intItem = 24 '防疫照顧假 Add By Sindy 2020/2/5
               If j = 51 Then intItem = 25 '天災不給薪 Add By Sindy 2025/11/13
               'Modify By Sindy 2012/7/9 上班時數為特殊者
'               If strST01 = "99029" Then
'                  dblDay = (dblHour(intItem) * 10) \ (5 * 10)
'                  dblHours = Round(dblHour(intItem) - (dblDay * 5), 1)
               If intItem > 0 Then
               '2019/10/9 END
                  If PUB_bSpecY <> True Then '非過渡期
                     dblDay = (dblHour(intItem) * 10) \ (PUB_intWkHour * 10)
                     dblHours = Round(dblHour(intItem) - (dblDay * PUB_intWkHour), 1)
                  Else
                     dblDay = 0
                     dblHours = dblHour(intItem)
                  End If
                  strText = IIf(dblDay > 0, dblDay & "日", "")
                  strText = strText & IIf(dblHours > 0, dblHours & "時", "")
                  GRD1.TextMatrix(i, j) = strText
               End If
            Next j
         End If
      Next i
   Else
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
   
   'Modify By Sindy 2012/2/13
'   GRD1.Visible = False
'   GRD1.col = 0
'   GRD1.row = 1
'   If rsTmp.RecordCount > 0 Then
''      GRD1.Text = "V"
''      For i = 0 To GRD1.Cols - 1
''         GRD1.col = i
''         GRD1.CellBackColor = &HFFC0C0
''      Next i
'      '當月:淺綠色 年度:淺橘色
'      For j = 1 To GRD1.Rows - 1
'         GRD1.row = j
'         For i = 0 To GRD1.Cols - 1
'            GRD1.col = i
'            If i >= 1 And i <= 3 Then
'               GRD1.CellBackColor = &H8000000F
'            ElseIf i >= 4 And i <= 12 Then
'               GRD1.CellBackColor = &H80FF80
'            ElseIf i >= 13 And i <= 21 Then
'               GRD1.CellBackColor = &H80C0FF
'            Else
'               GRD1.CellBackColor = QBColor(15)
'            End If
'         Next i
'      Next j
'   End If
'   GRD1.Visible = True
   Call SetColor
   
   'Add By Sindy 2019/10/9 有資料時才要顯示欄位
   Dim ii As Integer, jj As Integer, kk As Integer
   Dim bolFind As Boolean
   GRD1.Visible = False
   If GRD1.Rows > 1 Then
      If GRD1.TextMatrix(1, 2) <> "" Then
'         For ii = 26 To 47 '年度假別(行數)
'            bolFind = False
'            For jj = 1 To GRD1.Rows - 1 '列數
'               If GRD1.TextMatrix(jj, ii) <> "" Then
'                  bolFind = True
'                  Exit For
'               End If
'            Next jj
'            If bolFind = False Then
'               kk = ii - 22
'               GRD1.ColWidth(ii) = 0 '年度
'               GRD1.ColWidth(kk) = 0 '當月
'            End If
'         Next ii
         For ii = 4 To 51 '49 '47 '行數
            bolFind = False
            For jj = 1 To GRD1.Rows - 1 '列數
               If GRD1.TextMatrix(jj, ii) <> "" Then
                  bolFind = True
                  Exit For
               End If
            Next jj
            If bolFind = False Then
               GRD1.ColWidth(ii) = 0
            End If
         Next ii
      End If
   End If
   GRD1.Visible = True
   '2019/10/9 END
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'Add By Sindy 2012/2/13
Private Sub SetColor()
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   If GRD1.Rows > 1 Then
      If GRD1.TextMatrix(1, 2) <> "" Then
   '      GRD1.Text = "V"
   '      For i = 0 To GRD1.Cols - 1
   '         GRD1.col = i
   '         GRD1.CellBackColor = &HFFC0C0
   '      Next i
         '當月:淺綠色 年度:淺橘色
         For j = 1 To GRD1.Rows - 1
            GRD1.row = j
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               If i >= 1 And i <= 3 Then
                  GRD1.CellBackColor = &H8000000F
               ElseIf i >= 4 And i <= 27 Then '當月
                  GRD1.CellBackColor = &H80FF80
               ElseIf i >= 28 And i <= 51 Then '年度
                  GRD1.CellBackColor = &H80C0FF
               Else
                  GRD1.CellBackColor = QBColor(15)
               End If
            Next i
         Next j
      End If
   End If
   GRD1.Visible = True
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   frmpic003.Label1.Caption = "資料統計中, 請稍候..."
   frmpic003.Show
   DoEvents
   QueryData
   Unload frmpic003
   DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strText As String
   
   'Me.Form=C
   '一進入系統,檢查是否有須要開啟此作業
   If pub_CallNextABSForm = True Then
      strText = ChkIsAbsenceMustPro
      Me.Hide
      If InStr(1, strText, "D") > 0 Then
         'Set frm160102 = Nothing
         frm160102.intChoose = 1
         frm160102.Hide
         Call frm160102.cmdok_Click(0)
      'Add By Sindy 2015/7/2
      ElseIf InStr(1, strText, "G") > 0 Then
         If TypeName(Tmpfrm210148) <> "Nothing" Then
            Tmpfrm210148.Show
         End If
      ElseIf InStr(1, strText, "H") > 0 Then
         If TypeName(Tmpfrm210147) <> "Nothing" Then
            Tmpfrm210147.Show
         End If
      '2015/7/2 END
      Else
         pub_CallNextABSForm = False
      End If
   End If
   
   Set frm180203_1 = Nothing
   If pub_CallNextABSForm = False Then
      Call Forms(0).SysStartCallForm 'Add By Sindy 2011/10/7
   End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2020/2/5 + 防疫照顧假
   'Modify By Sindy 2025/11/12 + 天災不給薪
   arrGridHeadText = Array("V", "部門別", "員工代號", "姓名", _
                           "忘打卡", "遲到", "曠職", "事假", "病假", "公假", "特別假", "出差", "加班", "婚假", "產假", "流產假", "喪假", "公傷假", "補休", "其他", "扣年終產假", "扣年終流產假", "陪產檢及陪產假", "生理假", "產檢假", "家庭照顧假", "防疫照顧假", "天災不給薪", _
                           "忘打卡", "遲到", "曠職", "事假", "病假", "公假", "特別假", "出差", "加班", "婚假", "產假", "流產假", "喪假", "公傷假", "補休", "其他", "扣年終產假", "扣年終流產假", "陪產檢及陪產假", "生理假", "產檢假", "家庭照顧假", "防疫照顧假", "天災不給薪", _
                           "B1314", "B1303")
   arrGridHeadWidth = Array(200, 0, 800, 800, _
                           800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, _
                           800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, _
                           0, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignGeneral 'flexAlignRightCenter 'flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   If GRD1.TextMatrix(GRD1.MouseRow, 1) <> "" Then
      If GRD1.Text = "V" Then
         GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            If i >= 1 And i <= 3 Then
               GRD1.CellBackColor = &H8000000F
            ElseIf i >= 4 And i <= 27 Then
               GRD1.CellBackColor = &H80FF80
            ElseIf i >= 28 And i <= 51 Then
               GRD1.CellBackColor = &H80C0FF
            Else
               GRD1.CellBackColor = QBColor(15)
            End If
         Next i
      Else
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
End If
GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If Me.GRD1.Text = "員工代號" Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub
