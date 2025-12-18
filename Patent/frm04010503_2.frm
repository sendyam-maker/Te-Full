VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010503_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "核駁函輸入"
   ClientHeight    =   5748
   ClientLeft      =   120
   ClientTop       =   960
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9336
   Begin VB.CommandButton cmdOK 
      Caption         =   "內部收文(&E)"
      Height          =   400
      Index           =   3
      Left            =   5130
      TabIndex        =   17
      Top             =   60
      Width           =   1200
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   720
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "1"
      Top             =   5340
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7188
      TabIndex        =   8
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6360
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8412
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "frm04010503_2.frx":0000
      Left            =   1080
      List            =   "frm04010503_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   6420
      TabIndex        =   4
      Top             =   720
      Width           =   1332
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   1
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   0
      Top             =   720
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3852
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   6795
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1:核駁, 2:改變原處分, 3:部分准駁)"
      Height          =   180
      Left            =   1080
      TabIndex        =   16
      Top             =   5340
      Width           =   2685
   End
   Begin VB.Label Label3 
      Caption         =   "結果:"
      Height          =   252
      Left            =   120
      TabIndex        =   15
      Top             =   5340
      Width           =   492
   End
   Begin MSForms.Label Label8 
      Height          =   270
      Left            =   1800
      TabIndex        =   14
      Top             =   1110
      Width           =   6060
      VariousPropertyBits=   27
      Caption         =   "Label8"
      Size            =   "10689;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   4980
      TabIndex        =   12
      Top             =   720
      Width           =   888
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   768
   End
End
Attribute VB_Name = "frm04010503_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Label8)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim strReceiveNo As String, strTemp As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 1
         frm04010503_1.Show
         Unload Me
      Case 2
         Unload frm04010503_1
         Unload Me
      'Add By Cheng 2002/06/21
      Case 3 '內部收文
         mdiMain.mnu1102_Click 1
   End Select
End Sub

' 確認鈕
Private Sub FormConfirm()
 Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" Then
            bolChk = True
            Me.Tag = .TextMatrix(i, 1)
            strExc(5) = .TextMatrix(i, 4)
            strExc(6) = .TextMatrix(i, 12) 'Added by Morgan 2022/8/24
            strExc(7) = .TextMatrix(i, 2) 'Added by Morgan 2025/3/25
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If
   
   'Added by Morgan 2022/8/24
   If Text6 = "" Then
      MsgBox "請輸入結果！", vbExclamation
      Exit Sub
   ElseIf Text6 = "3" Then
      'Modified by Morgan 2025/3/25 +804舉發答辯且不再限制台灣案
      'If Not (pa(9) = "000" And strExc(6) = "803") Then
      If Not (strExc(6) = "803" Or strExc(6) = "804") Then
         'MsgBox "目前部分准駁只能選擇台灣舉發！", vbExclamation
         MsgBox "[" & strExc(7) & "] 不可選擇部分准駁！", vbExclamation
         Exit Sub
      End If
   End If
   'end 2022/8/24
   
   'Added by Morgan 2021/12/20
   '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
   If PUB_CheckFormExist("frm04010503_3") = False Then
      Set frm04010503_3 = Nothing
   End If
   'end 2021/12/20
   
   'Added by Morgan 2014/1/14
   frm04010503_3.m_AppNo = frm04010503_1.m_AppNo
   frm04010503_3.m_DocNo = frm04010503_1.m_DocNo
   frm04010503_3.m_DocWord = frm04010503_1.m_DocWord
   frm04010503_3.m_DeadLine = frm04010503_1.m_DeadLine
   'end 2014/1/14
   'Add By Sindy 2016/10/5
   frm04010503_3.m_strIR01 = m_strIR01
   frm04010503_3.m_strIR02 = m_strIR02
   frm04010503_3.m_strIR03 = m_strIR03
   frm04010503_3.m_strIR04 = m_strIR04
   '2016/10/5 END
   frm04010503_3.Show
   Me.Hide
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label8 = pa(5)
      Case "英"
         Label8 = pa(6)
      Case "日"
         Label8 = pa(7)
   End Select
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   
   pa(1) = strExc(1)
   pa(2) = strExc(2)
   pa(3) = strExc(3)
   pa(4) = strExc(4)
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010503_1.m_strIR01
   m_strIR02 = frm04010503_1.m_strIR02
   m_strIR03 = frm04010503_1.m_strIR03
   m_strIR04 = frm04010503_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
   ReadPatent 1
End Sub

Private Sub ReadPatent(ByVal iSitu As Integer)
 Dim strTmp As String
   Label8 = ""
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label8 = pa(5)
      Text1 = pa(11)
   End If
   If iSitu = 1 Then
      'Modify By Cheng 2002/04/15
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is null and " & _
'         "(substr(cp09,1,1)='A' or substr(cp09,1,1)='B' or " & _
'         "(substr(cp09,1,1)='C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))"
        'Modify By Cheng 2003/07/25
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is null and " & _
'         "( cp09<'C' or " & _
'         "( cp09>'C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))"
      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is null and " & _
         "( cp09<'C' or " & _
         "( cp09>'C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "' Or CP10='1504' Or CP10='1505' Or CP10='1211' Or CP10='1210' )))"
   Else
      'Modify By Cheng 2002/04/15
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is not null and " & _
'         "(substr(cp09,1,1)='A' or substr(cp09,1,1)='B' or " & _
'         "(substr(cp09,1,1)='C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))"
        'Modify By Cheng 2003/07/25
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is not null and " & _
'         "( cp09<'C' or " & _
'         "( cp09>'C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))"
      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is not null and " & _
         "( cp09<'C' or " & _
         "( cp09>'C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "' Or CP10='1504' Or CP10='1505' Or CP10='1211' Or CP10='1210' )))"
   End If
   
   If pa(9) = 台灣國家代號 Then
      strTmp = "CPM03"
   Else
      strTmp = "CPM04"
   End If
   
   'Modify By Cheng 2002/06/21
'   strExc(2) = "'',CP09," & strTmp & "," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP19,CP64 " & _
'      "from caseprogress,casepropertymap,CUSTOMER"
' 91.09.13 modify by louis
'strExc(2) = "'',CP09," & strTmp & ", CP43," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP19,CP64,CP10 " & _
'      "from caseprogress,casepropertymap,CUSTOMER"
    'Modify By Cheng 2003/01/27
    '加顯示對造號數欄位
'   strExc(2) = "'',CP09," & strTmp & ", CP43," & _
'      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
'      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP19,CP64,CP10 " & _
'      ", DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
'      "from caseprogress,casepropertymap,CUSTOMER"
   strExc(2) = "'',CP09," & strTmp & ", CP43," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40),CP36," & _
      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP19,CP64,CP10 " & _
      ", DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
      "from caseprogress,casepropertymap,CUSTOMER"
'   strExc(0) = "select " & strExc(2) & " where substr(cp10,1,1)='1' and " & strExc(1) & " and " & _
'      "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)" & _
'      " and (cp01,cp02,cp03,cp04) not in " & _
'      "(select cp01,cp02,cp03,cp04 from caseprogress where substr(cp10,1,1)='3' and " & _
'      strExc(1) & ") union " & _
'      "select " & strExc(2) & " where substr(cp10,1,1)<>'1' and " & strExc(1) & " and " & _
'      "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)"
' 91.09.13 modify by louis
   '94.2.2 modify by sonia
   'strExc(0) = "select " & strExc(2) & " where substr(cp10,1,1)='1' and " & strExc(1) & " and " & _
   '   "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)" & _
   '   " and (cp01,cp02,cp03,cp04) not in " & _
   '   "(select cp01,cp02,cp03,cp04 from caseprogress where substr(cp10,1,1)='3' and " & _
   '   strExc(1) & ") union " & _
   '   "select " & strExc(2) & " where substr(cp10,1,1)<>'1' and " & strExc(1) & " and " & _
   '   "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) " & _
   '   "ORDER BY SORTFIELD DESC "
   strExc(0) = "select " & strExc(2) & " where substr(cp10,1,1)='1' and " & strExc(1) & " and " & _
      "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)" & _
      " and (cp01,cp02,cp03,cp04) not in " & _
      "(select cp01,cp02,cp03,cp04 from caseprogress where substr(cp10,1,1)='3' and " & _
      strExc(1) & ") union " & _
      "select " & strExc(2) & " where (substr(cp10,1,1)<>'1' or cp10='107') and " & strExc(1) & " and " & _
      "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) " & _
      "ORDER BY SORTFIELD DESC "
   '94.2.2 end
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   Combo1.ListIndex = 0
   
   ' 若只有一筆資料時自動選取第一筆
   If MSHFlexGrid1.Rows = 2 Then
      MSHFlexGrid1.row = 1
      'Add by Morgan 2003/11/25
      If GridDataCheck() = False Then Exit Sub
      '---End
      GridClick MSHFlexGrid1, intLastRow, 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010503_2 = Nothing
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      'Add By Cheng 2002/06/21
      '加相關總收文號
      .col = 3: .ColWidth(3) = 1000: .Text = "相關總收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1000: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
        'Add By Cheng 2003/01/27
        '加對造號數
      .col = 5: .ColWidth(5) = 1000: .Text = "對造號數"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 800: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 800: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 800: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 9: .ColWidth(9) = 600: .Text = "結果"
      .CellAlignment = flexAlignCenterCenter
      .col = 10: .ColWidth(10) = 800: .Text = "後金"
      .CellAlignment = flexAlignCenterCenter
      .col = 11: .ColWidth(11) = 1000: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      'Add By Cheng 2002/06/21
      '加案件性質代號
      .col = 12: .ColWidth(12) = 1000: .Text = "案件性質代號"
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()

   'Add by Morgan 2003/11/25
   If GridDataCheck() = False Then Exit Sub
   '---End
   
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK(0).SetFocus

End Sub
'Add by Morgan 2003/11/25
Private Function GridDataCheck() As Boolean
   
   Dim strSql As String, strTemp As String, bolRtn As Boolean
   
   bolRtn = False
   If (MSHFlexGrid1.row = 0) Then
      bolRtn = True
   ElseIf (pa(9) <> "000") Then
      bolRtn = True
   Else
      'Modified by Morgan 2021/10/5 RsTemp 為共用若中間有被使用會導致錯誤，改抓Grid內的值
      'RsTemp.Move MSHFlexGrid1.row - 1, 1
      'strTemp = RsTemp.Fields("CP10")
      strTemp = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 12)
      'end 2021/10/5
      If (Len(strTemp) = 3 And strTemp >= "101" And strTemp <= "105") Then
         strTemp = pa(11)
         If (Trim(strTemp) = Empty) Then
            bolRtn = True
         Else
            'Modify by Morgan 2004/6/16
            Dim stCaseNo As String
            If PUB_ChkPriDate(strTemp, stCaseNo) Then
               MsgBox "此案已被 " & stCaseNo & " 主張國內優先權且自申請日起逾15個月，不可輸入准駁！", vbCritical
            Else
               bolRtn = True
            End If
            'end
         End If
      Else
         bolRtn = True
      End If
   End If
   GridDataCheck = bolRtn
   
End Function
Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub
'Modified by Morgan 2022/8/24 +3
Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   Else
      If KeyAscii = 49 Or KeyAscii = 51 Then
         ReadPatent 1
      Else
         ReadPatent 2
      End If
   End If
End Sub
