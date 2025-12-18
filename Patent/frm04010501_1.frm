VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010501_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "實審通知日輸入"
   ClientHeight    =   4395
   ClientLeft      =   135
   ClientTop       =   1800
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdOK 
      Caption         =   "內部收文(&E)"
      Height          =   400
      Index           =   3
      Left            =   5136
      TabIndex        =   0
      Top             =   80
      Width           =   1200
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm04010501_1.frx":0000
      Left            =   960
      List            =   "frm04010501_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   1032
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7152
      TabIndex        =   2
      Top             =   80
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6336
      TabIndex        =   1
      Top             =   80
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8376
      TabIndex        =   3
      Top             =   80
      Width           =   800
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   7
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2232
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   9108
      _ExtentX        =   16060
      _ExtentY        =   3942
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
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   5
      Left            =   1200
      TabIndex        =   22
      Top             =   1800
      Width           =   2160
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3810;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   4
      Left            =   960
      TabIndex        =   21
      Top             =   1560
      Width           =   8040
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "14182;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   3
      Left            =   5460
      TabIndex        =   20
      Top             =   1320
      Width           =   2340
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4128;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   2
      Left            =   960
      TabIndex        =   19
      Top             =   1320
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3016;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   1680
      TabIndex        =   18
      Top             =   1050
      Width           =   2790
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4921;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   210
      Index           =   0
      Left            =   5460
      TabIndex        =   17
      Top             =   720
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "4075;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   948
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "代理人:"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   588
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Left            =   4620
      TabIndex        =   14
      Top             =   1320
      Width           =   768
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   768
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   4620
      TabIndex        =   11
      Top             =   720
      Width           =   768
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   768
   End
End
Attribute VB_Name = "frm04010501_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/16 改成Form2.0 (Label2)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim intWhere As Integer
Dim intLastRow As Integer
Dim strReceiveNo As String
Dim m_DefaultPrinter As String
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END

Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer, bolChk As Boolean
Dim strCP09 As String
   Select Case Index
      Case 1
         'frm04010501.Command1_Click
         frm04010501.Show
         Unload frm04010501_2
         Unload Me
      Case 2
         Unload frm04010501
         Unload frm04010501_2
         Unload Me
      Case 3
         mdiMain.mnu1102_Click 1
      Case 0
         bolChk = False
         With MSHFlexGrid1
            .col = 0
            For i = 1 To .Rows - 1
               If .TextMatrix(i, 0) = "v" Then
                  bolChk = True
                  strExc(2) = .TextMatrix(i, 1) 'cp09
                  strExc(3) = .TextMatrix(i, 5) 'cp10
                  strExc(4) = .TextMatrix(i, 6) 'cp12
                  strExc(5) = .TextMatrix(i, 7) 'cp13
                  '93.2.19 ADD BY SONIA 發明申請案檢查實體審查是否已收文已發文
                  'Modified by Morgan 2012/12/26
                  '改和外專一樣判斷
                  'If strExc(3) = 發明申請 And Val(pa(10)) >= 911026 Then
                  If pa(8) = "1" And (strExc(3) = 發明申請 Or strExc(3) = 分割) Then
                     strExc(0) = "SELECT * FROM CASEPROGRESS WHERE " & _
                        ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP27 IS NOT NULL AND CP10='416' AND CP57 IS NULL "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If intI = 0 Then
                        MsgBox "此案未提實體審查, 不可輸入實審通知 !", vbInformation
                        Exit Sub
                     End If
                  End If
                  '93.2.19 END
                  Exit For
               End If
            Next
         End With
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
         Else
            Me.Hide
            'Added by Morgan 2014/1/14
            frm04010501_2.m_AppNo = frm04010501.m_AppNo
            frm04010501_2.m_DocNo = frm04010501.m_DocNo
            frm04010501_2.m_DocWord = frm04010501.m_DocWord
            'end 2014/1/14
            'Add By Sindy 2016/10/5
            frm04010501_2.m_strIR01 = m_strIR01
            frm04010501_2.m_strIR02 = m_strIR02
            frm04010501_2.m_strIR03 = m_strIR03
            frm04010501_2.m_strIR04 = m_strIR04
            '2016/10/5 END
            frm04010501_2.QueryData
            frm04010501_2.Show
         End If
   End Select
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(1) = pa(5)
      Case "英"
         Label2(1) = pa(6)
      Case "日"
         Label2(1) = pa(7)
   End Select
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   
   Text1 = strExc(1)
   Text2 = strExc(2)
   Text3 = strExc(3)
   Text4 = strExc(4)
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010501.m_strIR01
   m_strIR02 = frm04010501.m_strIR02
   m_strIR03 = frm04010501.m_strIR03
   m_strIR04 = frm04010501.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
   ' 90.07.18 modify by louis (暫存預設印表機的名稱)
   m_DefaultPrinter = Printer.DeviceName
   'Add By Cheng 2002/06/20
   SendKeys "{Tab}"
End Sub

Public Function QueryData() As Boolean
   QueryData = False
   
   ReadPatent
   If Combo1.ListCount > 0 Then
      Combo1.ListIndex = 0
   End If
   
   If MSHFlexGrid1.Rows <= 1 Then
      'MsgBox "沒有符合條件的資料", vbOKOnly + vbInformation, "查詢資料"
      QueryData = False
      '91.12.8 CANCEL BY SONIA
      'Unload Me
      '91.12.8 END
   Else
      ' 設定第一筆為預設
      MSHFlexGrid1.row = 1
      GridClick MSHFlexGrid1, intLastRow, 0
      QueryData = True
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010501_1 = Nothing
End Sub

'************************************************
' 取回專利基本資料及收文資料
'
'************************************************
Private Sub ReadPatent()
 Dim Lbl As Object, i As Integer, strTempName As String
   For Each Lbl In Label2
      Lbl = ""
   Next
   Label2(5) = frm04010501.Text5
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   
   If ClsPDReadPatentDatabase(pa(), intWhere) Then
      Label2(1) = pa(5)
      Label2(0) = pa(11)
      If pa(8) <> "" Then ChgType (2) ' Label2(2)
      If pa(9) <> "" Then ChgType (3) ' Label2(3)
      If pa(75) <> "" Then ChgType (4) ' Label2(4)
   End If
   '2005/3/28 MODIFY BY SONIA 加案件性質-402更正
   'Modified by Morgan 2012/12/19 +衍生設計125
   strExc(0) = "SELECT '',CP09,CPM03," & SQLDate("CP27") & ",CP08,CP10,CP12,CP13,CP43,DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40),CP36,CP64 FROM CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER WHERE " & _
      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP27 IS NOT NULL AND( ( CP10 IN ('101','102','103','104','105','125') AND CP09 NOT IN " & _
      "(SELECT CP43 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND " & _
      "CP27 IS NOT NULL AND (CP10='" & 通知實審日 & "' OR CP10='1217' OR SUBSTR(CP10,1,1)='3'))) OR (CP10 IN ('107','203','204','402') AND CP09 NOT IN " & _
      "(SELECT CP43 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND " & _
      "CP27 IS NOT NULL AND (CP10='" & 通知實審日 & "' OR CP10='1217')))) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
      " AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) Union " & _
      "SELECT '',CP09,CPM03," & SQLDate("CP27") & ",CP08,CP10,CP12,CP13,CP43,DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40),CP36,CP64 FROM CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER WHERE " & _
      ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP27 IS NOT NULL AND " & _
      "SUBSTR(CP10,1,1) IN ('3','8') AND CP09 NOT IN " & _
      "(SELECT CP43 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND " & _
      "CP27 IS NOT NULL AND (CP10='" & 通知實審日 & "' OR CP10='1217')) AND CP01=CPM01(+) AND CP10=CPM02(+)" & _
      " AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) ORDER BY CP09 DESC "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then
      Set MSHFlexGrid1.Recordset = RsTemp
   End If
   GridHead

End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 2
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetPatentTrademarkKind(專利, pA(8), strTempName, False, 台灣國家代號) = 1 Then
         If ClsPDGetPatentTrademarkKind(專利, pa(8), strTempName, False, 台灣國家代號) = 1 Then
            Label2(2) = strTempName
         End If
      Case 4
         'Modify By Cheng 2002/07/08
         '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'         If objPublicData.GetAgent(pa(75), strTempName) Then
         If PUB_GetAgentName(pa(1), pa(75), strTempName) Then
            Label2(4) = strTempName
            ChgType = True
         End If
      Case 3
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(pA(9), strTempName) = True Then
         If ClsPDGetNation(pa(9), strTempName) = True Then
            Label2(3) = strTempName
            ChgType = True
         End If
   End Select
End Function

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 2500: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1200: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1500: .Text = "機關文號"
      For i = 5 To 7
         .col = i: .ColWidth(i) = 0
      Next
      'Add By Cheng 2002/06/20
      .col = 8: .ColWidth(8) = 1200: .Text = "相關總收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 9: .ColWidth(9) = 1200: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
        'Add By Cheng 2003/01/27
        '加對造號數
      .col = 10: .ColWidth(10) = 1200: .Text = "對造號數"
      .CellAlignment = flexAlignCenterCenter
      .col = 11: .ColWidth(11) = 1200: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK(0).SetFocus
End Sub
