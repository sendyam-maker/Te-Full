VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010603_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "核駁函輸入"
   ClientHeight    =   5730
   ClientLeft      =   -1810
   ClientTop       =   1120
   ClientWidth     =   9350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9350
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   9
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   8
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4440
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      ItemData        =   "frm06010603_2.frx":0000
      Left            =   1080
      List            =   "frm06010603_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1020
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8436
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6384
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7212
      TabIndex        =   2
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   720
      MaxLength       =   1
      TabIndex        =   0
      Text            =   "1"
      Top             =   5400
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3912
      Left            =   120
      TabIndex        =   10
      Top             =   1380
      Width           =   9072
      _ExtentX        =   15998
      _ExtentY        =   6897
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
   Begin MSForms.Label LblFM2 
      Height          =   195
      Left            =   1740
      TabIndex        =   16
      Top             =   1050
      Width           =   7455
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "13150;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3360
      TabIndex        =   14
      Top             =   720
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1020
      Width           =   768
   End
   Begin VB.Label Label3 
      Caption         =   "結果:"
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   5400
      Width           =   492
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1:核駁, 2:改變原處分, 3:裁定駁回, 4:部分准駁 )"
      Height          =   180
      Left            =   1080
      TabIndex        =   11
      Top             =   5400
      Width           =   3660
   End
End
Attribute VB_Name = "frm06010603_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/4/23 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String, strTemp As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim sp() As String    'add by sonia 2024/11/21

Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer
'Added by Lydia 2023/09/25
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'end 2023/09/25

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 1
         frm06010603_1.Show
         Unload Me
      Case 2
         Unload frm06010603_1
         Unload Me
   End Select
End Sub

' 確認鈕
Private Sub FormConfirm()
 Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
 
   'Add by Morgan 2006/4/20
   If Text6 = "" Then
      MsgBox "請輸入結果", vbInformation
      Exit Sub
   End If
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" Then
            bolChk = True
            Me.Tag = .TextMatrix(i, 1)
            strExc(5) = .TextMatrix(i, 3)
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If
   
   'Added by Morgan 2017/5/10 電子公文
   frm06010603_3.m_DocWord = frm06010603_1.m_DocWord
   frm06010603_3.m_DocNo = frm06010603_1.m_DocNo
   frm06010603_3.m_DocDate = frm06010603_1.m_DocDate
   frm06010603_3.m_AppNo = frm06010603_1.m_AppNo
   frm06010603_3.m_DeadLine = frm06010603_1.m_DeadLine
   'end 2017/5/10
   'Added By Lydia 2023/09/25
   frm06010603_3.m_strIR01 = m_strIR01
   frm06010603_3.m_strIR02 = m_strIR02
   frm06010603_3.m_strIR03 = m_strIR03
   frm06010603_3.m_strIR04 = m_strIR04
   'end 2023/09/25
   frm06010603_3.Show
   Me.Hide
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         LblFM2 = pa(5)
      Case "英"
         LblFM2 = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         LblFM2 = pa(7)
   End Select
End Sub

Private Sub Form_Initialize()
   'add by nickc 2007/02/02
   ReDim pa(1 To TF_PA) As String
   ReDim sp(1 To tf_SP) As String   'add by sonia 2024/11/21
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   intWhere = 國外_FC
   pa(1) = strExc(1)
   pa(2) = strExc(2)
   pa(3) = strExc(3)
   pa(4) = strExc(4)
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   
   'Added by Lydia 2023/09/25
   m_strIR01 = frm06010603_1.m_strIR01
   m_strIR02 = frm06010603_1.m_strIR02
   m_strIR03 = frm06010603_1.m_strIR03
   m_strIR04 = frm06010603_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   'end 2023/09/25
   
   'modify by sonia 2024/11/21 +服務業務 FG-001323植物新品種保護
   'ReadPatent 1
   If pa(1) = "FCP" Then
      ReadPatent 1
   Else
      sp(1) = pa(1)
      sp(2) = pa(2)
      sp(3) = pa(3)
      sp(4) = pa(4)
      ReadServicePractice
   End If
   'en 2024/11/21
End Sub

Private Sub ReadPatent(ByVal iSitu As Integer)
 Dim Lbl As LABEL, txt As TextBox, i As Integer
 Dim strTmp(0 To 5) As String
   LblFM2 = ""
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      LblFM2 = pa(5)
      Text1 = pa(11)
   End If
   '駁回
   '2008/5/19 MODIFY BY SONIA 加506參加訴訟(FCP-013693)
   If iSitu = 1 Then
      '2010/11/23 modify by sonia C類來函被異議理由1801,被舉發理由1802,通知參加訴願1504,通知參加訴訟1505只抓下一程序未續辦者,改在下面另外抓
      'strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is null and " & _
         "(( cp09<'C' And ((cp10>='101' and cp10<='105') or cp10 ='107' or (cp10>='301' and cp10<='307') or (cp10>='401' and cp10<='403') or (cp10>='413' and cp10<='415') or (cp10>='501' and cp10<='504') or cp10='506' or cp10='507' or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804'))) Or " & _
         "( cp09>'C' And (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "' Or CP10='1504' Or CP10='1505' Or CP10='1211' Or CP10='1210' ))) "
      'Modified by Morgan 2014/6/25 +125 衍生設計
      'modify by sonia 2019/8/22 +439專利權部分拋棄,440申請權部分拋棄
      'modify by sonia 2024/3/15 +508行政上訴答辯
      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is null and " & _
         "(( cp09<'C' And ((cp10>='101' and cp10<='105') or cp10 ='107' or cp10 ='125' or (cp10>='301' and cp10<='307') or (cp10>='401' and cp10<='403') or (cp10>='413' and cp10<='415') or (cp10>='501' and cp10<='504') or cp10='506' or cp10='507' or cp10='439' or cp10='440' or cp10='508' or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804'))) Or " & _
         "( cp09>'C' And (CP10='1211' Or CP10='1210' ))) "
   'Add by Morgan 2006/4/21
   '裁定駁回
   ElseIf iSitu = 3 Then
      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is null and cp09<'C' and cp10 in ('503','504','507')"
         
   Else
      '2010/11/23 modify by sonia C類來函被異議理由1801,被舉發理由1802,通知參加訴願1504,通知參加訴訟1505只抓下一程序未續辦者,改在下面另外抓
      'strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is not null and " & _
         "(( cp09<'C' And ((cp10>='101' and cp10<='105') or cp10 ='107' or (cp10>='301' and cp10<='307') or (cp10>='401' and cp10<='403') or (cp10>='413' and cp10<='415') or (cp10>='501' and cp10<='504') or cp10='506' or cp10='507' or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804'))) Or " & _
         "( cp09>'C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "' Or CP10='1504' Or CP10='1505' Or CP10='1211' Or CP10='1210' ))) "
      'Modified by Morgan 2014/6/25 +125 衍生設計
      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is not null and " & _
         "(( cp09<'C' And ((cp10>='101' and cp10<='105') or cp10 ='107' or cp10 ='125' or (cp10>='301' and cp10<='307') or (cp10>='401' and cp10<='403') or (cp10>='413' and cp10<='415') or (cp10>='501' and cp10<='504') or cp10='506' or cp10='507' or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804'))) Or " & _
         "( cp09>'C' and (CP10='1211' Or CP10='1210' ))) "
   End If
   strExc(2) = "'',CP09,CPM03," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64 " & _
      ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD,CP10 " & _
      "from caseprogress,casepropertymap,CUSTOMER"
   '2010/11/23 add by sonia 已有專用期間者不帶出新申請案件性質
   If pa(25) <> "" Then
      strExc(3) = " and instr('" & NewCasePtyList & "',cp10)=0 "
   Else
      strExc(3) = ""
   End If
   '2010/11/23 end
   strExc(0) = "select " & strExc(2) & " where substr(cp10,1,1)='1' and " & strExc(1) & strExc(3) & " and " & _
      "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)" & _
      " and (cp01,cp02,cp03,cp04) not in " & _
      "(select cp01,cp02,cp03,cp04 from caseprogress where substr(cp10,1,1)='3' and " & _
      strExc(1) & ") union " & _
      "select " & strExc(2) & " where (substr(cp10,1,1)<>'1' or cp10='107') and " & strExc(1) & strExc(3) & " and " & _
      "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) "
   '2010/11/23 add by sonia C類來函被異議理由1801,被舉發理由1802,通知參加訴願1504,通知參加訴訟1505只抓下一程序未續辦者
   If iSitu <> 2 Then
      strExc(0) = strExc(0) & " union " & _
         "select " & strExc(2) & ",nextprogress where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         "and cp09>'C' And (CP10='1801' Or CP10='1802' Or CP10='1504' Or CP10='1505') " & _
         "and cp09=np01(+) and (np06 is null or np06='N') " & _
         "and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) "
      '駁回
      If iSitu = 1 Then
         strExc(0) = strExc(0) & " and cp24 is null "
      Else
         strExc(0) = strExc(0) & " and cp24 is not null "
      End If
   End If
   '2010/11/23 end
   strExc(0) = strExc(0) & "ORDER BY SORTFIELD DESC "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   Combo1.ListIndex = 0
End Sub

'add by sonia 2024/11/21 +服務業務 FG-001323植物新品種保護120
Private Sub ReadServicePractice()
 Dim Lbl As LABEL, txt As TextBox, i As Integer
 Dim strTmp(0 To 5) As String
   If ClsPDReadServicePracticeDatabase(sp(), intWhere) Then
      Text1 = sp(11)
   End If
   strExc(2) = "'',CP09,CPM03,CP40," & _
      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64 " & _
      ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD,CP10 " & _
      "from caseprogress,casepropertymap,CUSTOMER"
   strExc(0) = "select " & strExc(2) & " where cp10='120' and " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp27 is not null and cp24 is null" & _
      " and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) ORDER BY SORTFIELD DESC "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   Combo1.ListIndex = 0
End Sub
'end 2024/11/21

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010603_2 = Nothing
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1500: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1500: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1500: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1500: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1500: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 1500: .Text = "結果"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 1500: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()

   'GridClick MSHFlexGrid1, intLastRow, 0
  
   'Add by Morgan 2003/11/25
   If GridDataCheck() = False Then Exit Sub
   '---End
   
'Modify by Morgan 2003/11/26

'   Dim nOldCol As Integer
'   Dim nCol As Integer
'   If MSHFlexGrid1.Row > 0 Then
'      If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = "v" Then
'         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = Empty
'      Else
'         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0) = "v"
'      End If
'      nOldCol = MSHFlexGrid1.Col
'      For nCol = 0 To MSHFlexGrid1.Cols - 1
'         MSHFlexGrid1.Col = nCol
'         MSHFlexGrid1.CellBackColor = &HFFC0C0
'      Next nCol
'      MSHFlexGrid1.Col = nOldCol
'   End If
GridClick MSHFlexGrid1, intLastRow, 0
MSHFlexGrid1.SetFocus
'---End
   
End Sub
'Add by Morgan 2003/11/25
Private Function GridDataCheck() As Boolean
   
   
   Dim strSql As String, strTemp As String, bolRtn As Boolean
   Dim rsCheck As New ADODB.Recordset
   
   bolRtn = False
   If (MSHFlexGrid1.row = 0) Then
      bolRtn = True
   ElseIf (pa(9) <> "000") Then
      bolRtn = True
   Else
      MSHFlexGrid1.Recordset.Move MSHFlexGrid1.row - 1, 1
      strTemp = MSHFlexGrid1.Recordset.Fields("CP10")
      If (Len(strTemp) = 3 And strTemp >= "101" And strTemp <= "105") Then
         strTemp = pa(11)
         If (Trim(strTemp) = Empty) Then
            bolRtn = True
         Else
            strSql = "SELECT PD01||'-'||PD02||PD03||PD04 FROM PRIDATE WHERE PD06='" & strTemp & "' AND PD07='000'"
            rsCheck.CursorLocation = adUseClient
            rsCheck.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If Not (rsCheck.BOF And rsCheck.EOF) Then
               MsgBox "此案已被 " & rsCheck.Fields(0).Value & " 主張國內優先權！", vbCritical
            Else
               bolRtn = True
            End If
            rsCheck.Close
         End If
      Else
         bolRtn = True
      End If
   End If
   GridDataCheck = bolRtn
   
End Function
Private Sub Text6_GotFocus()
   InverseTextBox Text6
End Sub

'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'   If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
'      KeyAscii = 0
'      Beep
'   Else
'      If KeyAscii = 49 Then
'         ReadPatent 1
'      Else
'         ReadPatent 2
'      End If
'   End If
'End Sub
' 90.06.26 modify by louis 結果改為離開欄位才檢查
Private Sub Text6_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   
   If IsEmptyText(Text6) = False Then
      'add by sonia 2024/11/21
      If pa(1) = "FG" Then
         Select Case Text6
            Case "1":
               ReadServicePractice
            Case Else:
               Beep
               Cancel = True
               strTit = "資料檢核"
               strMsg = "只可輸入1"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End Select
      Else
      'end 2024/11/21
         Select Case Text6
            'modify by sonia 2025/4/22 +4部分准駁
            Case "1", "4":
               ReadPatent 1
            Case "2":
               ReadPatent 2
            'Add by Morgan 2006/4/21 裁定駁回
            Case "3"
               ReadPatent 3
            Case Else:
               Beep
               Cancel = True
               strTit = "資料檢核"
               strMsg = "只可輸入1~4"   'modify by sonia 2025/4/22 +4部分准駁
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End Select
      End If  'add by sonia 2024/11/21
   End If
End Sub
