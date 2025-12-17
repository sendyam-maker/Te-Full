VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmacc42c0 
   BorderStyle     =   1  '單線固定
   Caption         =   "旅遊補助付款通知"
   ClientHeight    =   5730
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9110
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9110
   Begin VB.TextBox txt1 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   750
      Index           =   1
      Left            =   150
      Locked          =   -1  'True
      MaxLength       =   7
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Frmacc42c0.frx":0000
      Top             =   4950
      Width           =   8895
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "執行(&E)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   345
      Index           =   1
      Left            =   5460
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   90
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "畫面更新(&Q)"
      Height          =   345
      Index           =   0
      Left            =   6405
      TabIndex        =   1
      Top             =   90
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1770
      MaxLength       =   7
      TabIndex        =   0
      Top             =   150
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   7740
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   90
      Width           =   780
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "Frmacc42c0.frx":00AB
      Height          =   4335
      Left            =   90
      TabIndex        =   6
      Top             =   570
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   7638
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "員工編號|姓名|申請日|申請金額|補助年度|補助額度|補助金額|旅遊期間|備註"
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
      _Band(0).Cols   =   9
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "付款日期："
      Height          =   180
      Left            =   840
      TabIndex        =   5
      Top             =   210
      Width           =   900
   End
End
Attribute VB_Name = "frmacc42c0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/05 Form2.0已修改 GRD1
'Create by Sindy 2020/3/11
Option Explicit


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 2 '結束
         Unload Me
      Case 0 '畫面更新
         GetData
      Case 1 '執行
         SaveData
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   GetData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmacc42c0 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow GRD1, x, y, nCol, nRow
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim i As Integer

GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
'   '上一筆資料列清除反白
'   If dblPrevRow > 0 Then
'      GRD1.col = 0
'      GRD1.row = dblPrevRow
'      GRD1.Text = ""
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = QBColor(15)
'      Next i
'   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
'   dblPrevRow = GRD1.row
   If GRD1.Text = "V" Then
      GRD1.Text = ""
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   Else
      If GRD1.TextMatrix(GRD1.row, 1) <> "" Then
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

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If Val(txt1(0)) = 0 Then
      MsgBox "付款日期不可以空白！", vbExclamation
      txt1(0).SetFocus
      Exit Function
   End If
   
   Cancel = False
   txt1_Validate 0, Cancel
   If Cancel = True Then
      Exit Function
   End If
   
   TxtValidate = True
End Function

Private Function SaveData() As Boolean
Dim strEmp As String, strEMP_Tel As String
Dim intRow As Integer
Dim strTo As String, strCC As String
Dim strSubject As String, strContent As String
Dim strSql As String
Dim i As Integer
Dim strCompName As String 'Add By Sindy 2020/3/30
   
   '檢查欄位有效性
   If TxtValidate = False Then Exit Function
   
On Error GoTo ErrHand
   
   'Modify By Sindy 2021/11/1 Pub_GetSpecMan("財務處出納人員") ==> strUserNum
   '辜苑琪(財務.主任): 改設定為輸入的人 原因:辜跟瑞婷都有可能發MAIL
   strExc(0) = "select st02,ed01" & _
               " from staff,ExtensionData" & _
               " where ST01=ED02(+)" & _
               " and st01='" & strUserNum & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strEmp = RsTemp.Fields("st02")
      strEMP_Tel = "" & RsTemp.Fields("ed01")
   End If
   
   For intRow = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(intRow, 0) = "V" Then
         strSql = "UPDATE staff_TravelData SET STD13=" & DBDATE(txt1(0)) & _
                  " WHERE STD01=" & CNULL(GRD1.TextMatrix(intRow, 1)) & " and STD02=" & DBDATE(GRD1.TextMatrix(intRow, 5))
         cnnConnection.Execute strSql
         GRD1.RowHeight(intRow) = 0
         strTo = "": strSubject = ""
         If Val(txt1(0)) >= 1090313 Then
            '寄發E-Mail
            'modify by sonia 2023/10/25 加入5其他所
            'If GRD1.TextMatrix(intRow, 12) = "2" Or _
               GRD1.TextMatrix(intRow, 12) = "3" Or _
               GRD1.TextMatrix(intRow, 12) = "4" Then '北所
            If GRD1.TextMatrix(intRow, 12) <> "1" Then '非北所
               '由北所輸入付款日期後發email 通知分所管理人員付款並cc當事人
               '管理人員： 中所 87027.陳淑芳/ 南所 71002.唐惠琴 / 高所 68008.余玉瑛
               If GRD1.TextMatrix(intRow, 12) = "2" Then
                  strTo = Pub_GetSpecMan("出納人員-中所")
               ElseIf GRD1.TextMatrix(intRow, 12) = "3" Then
                  strTo = Pub_GetSpecMan("出納人員-南所")
               ElseIf GRD1.TextMatrix(intRow, 12) = "4" Then
                  'Modified by Morgan 2023/7/4 玉瑛留停3個月,暫改 A8029 呂麗君
                  strTo = Pub_GetSpecMan("出納人員-高所")
                  'strTo = "A8029"    'cancel by sonia 2023/10/2玉瑛復職
                  'end 2023/7/4
               'add by sonia 2023/10/25 5其他所
               ElseIf GRD1.TextMatrix(intRow, 12) = "5" Then
                  'Modify by Amy 2024/05/13 財務2個特殊設定拆成3個
                  If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
                     strTo = Pub_GetSpecMan("財務處應收處理人員")
                  Else
                     strTo = Pub_GetSpecMan("財務處出納人員")
                  End If
                  'end 2024/05/13
               'end 2023/10/25
               End If
               strCC = GRD1.TextMatrix(intRow, 1)
               'Add By Sindy 2024/2/17
               If ChkStaffST04(strCC, False) = True Then
                  strCC = strUserNum
                  strSubject = "(已離職)"
               End If
               '2024/2/17 END
               'Modify By Sindy 2024/2/17 加入員工編號及姓名
               strSubject = strSubject & GRD1.TextMatrix(intRow, 1) & _
                            GRD1.TextMatrix(intRow, 2) & "旅遊補助款 NT" & GRD1.TextMatrix(intRow, 6) & " 已核准，請付款！"
            Else 'If GRD1.TextMatrix(intRow, 12) = "1" Then '北所
               '輸入付款日期後發email通知當事人
               If GRD1.TextMatrix(intRow, 12) = "1" Then '北所
                  strTo = GRD1.TextMatrix(intRow, 1)
               'Add By Sindy 2021/12/1 所別非1~4的發信，改抓員工檔的外部信箱ST18
               Else
                  strSql = "SELECT ST18 FROM STAFF WHERE ST01='" & GRD1.TextMatrix(intRow, 1) & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     strTo = "" & RsTemp.Fields("ST18")
                  End If
               End If
               If strTo = "" Then
                  strTo = strUserNum
               Else
                  'Add By Sindy 2024/2/17
                  If ChkStaffST04(strTo, False) = True Then
                     strTo = strUserNum
                     strSubject = "(已離職)"
                  End If
                  '2024/2/17 END
               End If
               '2021/12/1 END
               strCC = ""
               'Modify By Sindy 2024/2/17 加入員工編號及姓名
               strSubject = strSubject & GRD1.TextMatrix(intRow, 1) & GRD1.TextMatrix(intRow, 2) & _
                            "旅遊補助款 NT" & GRD1.TextMatrix(intRow, 6) & " 已於" & ChangeTStringToTDateString(txt1(0)) & "存入您銀行帳戶，請查收！"
            End If
            'Modify By Sindy 2020/3/30
            '"台一國際專利法律事務所" & vbCrLf => a0802query(strCompName)
            Call ClsPDGetStaffComp(GRD1.TextMatrix(intRow, 1), strCompName)  'Add By Sindy 2020/3/30
            strContent = vbCrLf & "同主旨" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                        "財務處　" & strEmp & vbCrLf & _
                        A0802Query(strCompName) & vbCrLf & _
                        "台北市長安東路２段１１２號９樓" & vbCrLf & _
                        "電話：０２－２５０６１０２３" & IIf(strEMP_Tel <> "", "（" & strEMP_Tel & "）", "") & vbCrLf & _
                        "傳真：０２－２５０１１６６６"
            '請假不必發給職代
            PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , True, , strCC, , , , True
            If bolMailSendOk = False Then Exit Function
         End If
      End If
   Next intRow
   GetData
   
   Exit Function

ErrHand:
   MsgBox Err.Number & ":" & Err.Description
End Function

Public Sub GetData()
Dim rsTmp As New ADODB.Recordset
Dim ii As Integer
Dim strKey1 As String, StrKey2 As String
   
   '抓取資料
   'Modify By Sindy 2025/4/24 +所別,部門別
   strSql = "SELECT '' as V,STD01,st02,decode(st06,'1','北','2','中','3','南','4','高')||'所',a0922,sqldateT(STD02),to_char(STD03,'999,999'),STF03-1911," & _
            "to_char(STF04,'999,999'),to_char(STF05,'999,999'),sqldateT(STD04)||'~'||sqldateT(STD05),STD06,st06" & _
            " FROM staff_TravelFee,staff,staff_TravelData,acc090new" & _
            " where STD01=st01(+) and STD01=STF01(+) and STD02=STF02(+) and STD13 is null and st93=a0921(+)" & _
            " order by STD02 desc,STD01 desc,STF03 asc"
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'GRD1.FixedCols = 0 'Add By Sindy 2019/10/23
   Set GRD1.Recordset = rsTmp
   SetGrd
   'GRD1.FixedCols = 2 'Add By Sindy 2019/10/23
   For ii = 1 To GRD1.Rows - 1
      If strKey1 <> "" And GRD1.TextMatrix(ii, 1) <> "" And _
         (strKey1 = GRD1.TextMatrix(ii, 1) And StrKey2 = GRD1.TextMatrix(ii, 5)) Then
         GRD1.TextMatrix(ii, 1) = ""
         GRD1.TextMatrix(ii, 2) = ""
         GRD1.TextMatrix(ii, 5) = ""
         GRD1.TextMatrix(ii, 6) = ""
         GRD1.TextMatrix(ii, 10) = ""
         GRD1.TextMatrix(ii, 11) = ""
      Else
         strKey1 = GRD1.TextMatrix(ii, 1)
         StrKey2 = GRD1.TextMatrix(ii, 5)
      End If
   Next ii
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   'Modify By Sindy 2025/4/24 +所別,部門別
   '                        0    1           2       3       4         5         6           7           8           9           10          11      12
   arrGridHeadText = Array("V", "員工編號", "姓名", "所別", "部門別", "申請日", "申請金額", "補助年度", "補助額度", "補助金額", "旅遊期間", "備註", "st06")
   arrGridHeadWidth = Array(300, 800, 750, 500, 800, 800, 800, 850, 850, 800, 1700, 2000, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         'KeyAscii = UpperCase(KeyAscii)
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0 '付款日期
         If CheckIsTaiwanDate(txt1(Index), False) = False Then
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Cancel = True
            Exit Sub
         End If
         If Not ChkWorkDay(DBDATE(txt1(Index))) Then
            MsgBox "付款日期必須是工作天 !", vbExclamation
            Cancel = True
            Exit Sub
         End If
         If Val(DBDATE(txt1(Index))) > Val(strSrvDate(1)) Then
            MsgBox "付款日期不可大於系統日！"
            Cancel = True
            Exit Sub
         End If
      Case Else
   End Select
End Sub
