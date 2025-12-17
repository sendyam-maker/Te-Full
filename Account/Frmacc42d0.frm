VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc42d0 
   BorderStyle     =   1  '單線固定
   Caption         =   "出庭費發放通知"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9108
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9108
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認明細(&E)"
      Height          =   345
      Index           =   3
      Left            =   5208
      TabIndex        =   7
      Top             =   90
      Width           =   1155
   End
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
      Top             =   4950
      Width           =   8895
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "執行(&E)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   345
      Index           =   1
      Left            =   4320
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
      Bindings        =   "Frmacc42d0.frx":0000
      Height          =   4335
      Left            =   90
      TabIndex        =   6
      Top             =   570
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   7641
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|員工編號|姓名|律所案號|智慧所案號|出庭費總金額|確認領取日期"
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
      _Band(0).Cols   =   7
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "發放日期："
      Height          =   180
      Left            =   840
      TabIndex        =   5
      Top             =   216
      Width           =   900
   End
End
Attribute VB_Name = "Frmacc42d0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/09/30 (113/11/01上線)
Option Explicit
Public cmdState As Integer
Dim intLastRow As Integer '記錄GRD1勾選最後一筆
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim intQ As Integer, strQ1 As String
Dim rsQD As New ADODB.Recordset
Dim colCL01 As Integer, colCL02 As Integer, colCL04 As Integer '收文號,員工編號,確認領取日期

Private Sub cmdOK_Click(Index As Integer)
Dim intA As Integer, intChk As Integer
Dim m_strFilePath As String
   cmdState = Index
   Select Case Index
      Case 2 '結束
         Unload Me
      Case 0 '畫面更新
         GetData
      Case 3  '確認明細
         If PUB_CheckFormExist("frm075013_1") Then
            MsgBox "請先關閉〔出庭費確認維護明細〕畫面！"
            Exit Sub
         End If
         PubShowNextData
      Case 1 '執行
         '檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         m_strFilePath = strExcelPath & strSrvDate(2) & "_" & "出庭費發放通知" & MsgText(43)
         If Dir(m_strFilePath) <> "" Then
            If PUB_ChkFileOpening(m_strFilePath) = True Then
               Exit Sub
            End If
            Kill m_strFilePath
         End If
         
         intChk = 0
         For intA = 1 To GRD1.Rows - 1
            GRD1.col = 0
            GRD1.row = intA
            If Trim(GRD1.Text) = "V" And Trim(GRD1.TextMatrix(intA, colCL04)) = "" Then
               intChk = intChk + 1
            End If
         Next
         If intChk > 0 Then
            If MsgBox("尚有" & intChk & "筆無確認領取日期，是否繼續執行？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
         
         Me.Enabled = False
         Screen.MousePointer = vbHourglass
         Call SaveData(m_strFilePath)
         Screen.MousePointer = vbDefault
         Me.Enabled = True
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   GetData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsQD = Nothing
   Set Frmacc42d0 = Nothing
End Sub

Private Sub Grd1_Click()

   If "" & GRD1.TextMatrix(GRD1.row, colCL01) <> "" Then
      GridClick GRD1, intLastRow, 0, 2, , "V"
   End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   getGrdColRow GRD1, x, y, nCol, nRow
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If InStr("出庭費總金額", Me.GRD1.Text) > 0 Then
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

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If Val(txt1(0)) = 0 Then
      MsgBox "發放日期不可以空白！", vbExclamation
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

Private Function SaveData(ByRef strFileName As String) As Boolean
Dim strEmp As String, strEMP_Tel As String
Dim intRow As Integer
Dim strTo As String, strCC As String
Dim strSubject As String, strContent As String
Dim strCompName As String
Dim strExSql As String, strNoList As String


On Error GoTo ErrHand

   '使用輸入的人 原因:辜跟瑞婷都有可能發MAIL
   strEmp = strUserName
   strEMP_Tel = Pub_GetStaffExtn(strUserNum)
   
   For intRow = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(intRow, 0) = "V" Then
         strExSql = "Update CaseLawer Set CL06=" & DBDATE(txt1(0)) & " Where CL01='" & GRD1.TextMatrix(intRow, colCL01) & "' and CL02='" & GRD1.TextMatrix(intRow, colCL02) & "' "
         cnnConnection.Execute strExSql
         
         GRD1.RowHeight(intRow) = 0
         strNoList = strNoList & "," & GRD1.TextMatrix(intRow, colCL01) & GRD1.TextMatrix(intRow, colCL02)
      End If
   Next intRow
   If strNoList <> "" Then
      'Modified by Lydia 2025/03/06 +ST06
      strQ1 = "SELECT CL02,ST02,ST06,LISTAGG(CL01||CL02||ST06,',') WITHIN GROUP (ORDER BY CL01) GRPNO " & _
              "FROM (SELECT CL01,CL02 FROM CASELAWER WHERE INSTR ('" & strNoList & "', CL01||CL02)>0), STAFF WHERE CL02=ST01(+) GROUP BY CL02,ST02,ST06 ORDER BY CL02 "
      intQ = 1
      Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 1 Then
         rsQD.MoveFirst
         Do While Not rsQD.EOF
            Call PUB_GetFrm075013toXls(Me.Name, "", "Y", "", "", strFileName, "" & rsQD.Fields("GRPNO"))
            If strFileName <> "" Then
               If Dir(strFileName) <> "" Then
                  strSubject = Mid(txt1(0), 1, 3) & "年" & Mid(txt1(0), 4, 2) & "月出庭費已發放完成"
                  strContent = rsQD.Fields("ST02") & "　您好，" & vbCrLf & _
                               Mid(txt1(0), 1, 3) & "年" & Mid(txt1(0), 4, 2) & "月的出庭費已在今日發放完成，發放明細請詳如附件。" & vbCrLf & _
                               "提醒您，實際收到款項可能較少是因為已扣除所得稅、補充保費。" & vbCrLf
                  Call ClsPDGetStaffComp(rsQD.Fields("CL02"), strCompName)
                  strContent = strContent & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                              "財務處　" & strEmp & vbCrLf & _
                              A0802Query(strCompName) & vbCrLf & _
                              "台北市長安東路２段１１２號９樓" & vbCrLf & _
                              "電話：０２－２５０６１０２３" & IIf(strEMP_Tel <> "", "（" & strEMP_Tel & "）", "") & vbCrLf & _
                              "傳真：０２－２５０１１６６６"
                  'Added by Lydia 2025/03/06 比照旅遊補助付款通知：為分所同仁時，新增發放通知副本予分所會計
                  Select Case "" & rsQD.Fields("st06")
                      Case "2" '中所
                         strExc(1) = Pub_GetSpecMan("出納人員-中所")
                      Case "3" '中所
                         strExc(1) = Pub_GetSpecMan("出納人員-南所")
                      Case "4" '高所
                         strExc(1) = Pub_GetSpecMan("出納人員-高所")
                      Case "5" '其他所
                         strExc(1) = Pub_GetSpecMan("財務處應收處理人員")
                      Case Else
                         strExc(1) = ""
                  End Select
                  'end 2025/03/06
                  
                  '請假不必發給職代
                  'Modified by Lydia 2025/03/06 + CC >> strExc(1)
                  PUB_SendMail strUserNum, "" & rsQD.Fields("CL02"), "", strSubject, strContent, , strFileName, , True, , strExc(1), , , , True
                  If bolMailSendOk = False Then Exit Function
On Error Resume Next
                  Kill strFileName
               End If
            End If
            rsQD.MoveNext
         Loop
      End If
      Set rsQD = Nothing
   End If
   GetData

   Exit Function

ErrHand:
   MsgBox Err.Number & ":" & Err.Description
End Function

Public Sub GetData()
   
   Call SetGrd(True) '清空
   '抓取資料：所有未發放的出庭費>> 目前有設定律師出庭費CL03並且已發文的資料
   'Modified by Lydia 2025/04/07 +CL09財務確認律師不領取出庭費日期
   strQ1 = "SELECT '' AS V,A.CL02 AS Y01, S1.ST02 AS Y02,C1.CP01||'-'||C1.CP02||'-'||C1.CP03||'-'||C1.CP04 AS CASENO " & _
           ",DECODE(C2.CP01,NULL,NULL,'TT',NULL,C2.CP01||'-'||C2.CP02||'-'||C2.CP03||'-'||C2.CP04) AS PCASE " & _
           ",A.CL03 AS Y03,SQLDATET(A.CL04) AS CL04T,A.CL01 " & _
           "FROM CASELAWER A, STAFF S1, CASEPROGRESS C1, LAWOFFICESOURCE, CASEPROGRESS C2 " & _
           "WHERE A.CL02=S1.ST01(+) AND NVL(A.CL03,0)>0  AND A.CL06||A.CL09 IS NULL AND C1.CP158> 0 AND C1.CP159=0 " & _
           "AND A.CL01=C1.CP09(+) AND C1.CP162=LOS15(+) AND LOS01=C2.CP09(+) "
   strQ1 = strQ1 & " ORDER BY A.CL04,A.CL02,C1.CP158 "
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      Set GRD1.Recordset = rsQD
      Call SetGrd
   End If

End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

   '                       0     1           2       3          4              5        6               7
   arrGridHeadText = Array("V", "員工編號", "姓名", "律所案號", "智慧所案號", "出庭費", "確認領取日期", "CL01")
   arrGridHeadWidth = Array(300, 800, 800, 1400, 1400, 1300, 1300, 0)
        
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
      GRD1.Clear
      GRD1.Rows = 2
   End If
       
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next

   For intI = 1 To GRD1.Rows - 1
      GRD1.row = intI
      For iRow = 0 To GRD1.Cols - 1
         GRD1.col = iRow
         If InStr("05,", Format(iRow, "00")) > 0 Then
            GRD1.CellAlignment = flexAlignCenterCenter
         End If
      Next iRow
   Next intI
   
   If colCL01 = 0 Then
      colCL01 = PUB_MGridGetId("CL01", GRD1)
      colCL02 = PUB_MGridGetId("員工編號", GRD1)
      colCL04 = PUB_MGridGetId("確認領取日期", GRD1)
   End If
      
   GRD1.Visible = True
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0 '發放日期
         If CheckIsTaiwanDate(txt1(Index), False) = False Then
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Cancel = True
            Exit Sub
         End If
         If Not ChkWorkDay(DBDATE(txt1(Index))) Then
            MsgBox "發放日期必須是工作天 !", vbExclamation
            Cancel = True
            Exit Sub
         End If
         If Val(DBDATE(txt1(Index))) > Val(strSrvDate(1)) Then
            MsgBox "發放日期不可大於系統日！"
            Cancel = True
            Exit Sub
         End If
      Case Else
   End Select
End Sub

Public Sub PubShowNextData(Optional bolRefresh As Boolean = False)
Dim intA As Integer, intB As Integer
Dim strKeyNo As String

On Error GoTo ErrorHandler
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   For intA = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = intA
      If Trim(GRD1.Text) = "V" Then
         bolRefresh = False
         GRD1.col = 0
         GRD1.Text = ""
         For intB = 0 To GRD1.Cols - 1
            GRD1.col = intB
            GRD1.CellBackColor = &H80000005
         Next

         strKeyNo = GRD1.TextMatrix(intA, colCL01)
         
         Me.Show
         Select Case cmdState
            Case 3 '確認明細
               If Len(strKeyNo) = 9 Then
                  Call frm075013_1.SetParent(Me, strKeyNo, GRD1.TextMatrix(intA, colCL02))
                  frm075013_1.Show
                  Me.Hide
               End If
         End Select
         Exit For
      End If
   Next intA
   
   If bolRefresh = True Then
      GetData
   End If
   
ErrorHandler:
   If Err.Number <> 0 Then
      MsgBox "(" & Err.Number & ")" & Err.Description
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

