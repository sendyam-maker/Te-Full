VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060107 
   BorderStyle     =   1  '單線固定
   Caption         =   "翻譯完稿輸入"
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
   Begin VB.CheckBox Check1 
      Caption         =   "已發文補輸字數或改核搞人"
      Height          =   285
      Left            =   6255
      TabIndex        =   18
      Top             =   1500
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7185
      TabIndex        =   6
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8025
      TabIndex        =   7
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6345
      TabIndex        =   4
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   1
      Left            =   1770
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "FCP"
      Top             =   570
      Width           =   495
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   2
      Left            =   2250
      MaxLength       =   6
      TabIndex        =   1
      Top             =   570
      Width           =   855
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   3
      Left            =   3090
      MaxLength       =   1
      TabIndex        =   2
      Top             =   570
      Width           =   255
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   4
      Left            =   3330
      MaxLength       =   2
      TabIndex        =   3
      Top             =   570
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3810
      Left            =   30
      TabIndex        =   5
      Top             =   1830
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   6720
      _Version        =   393216
      Cols            =   8
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   3
      Left            =   1770
      TabIndex        =   17
      Top             =   1530
      Width           =   3195
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   2
      Left            =   1770
      TabIndex        =   16
      Top             =   1200
      Width           =   3195
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Index           =   1
      Left            =   1770
      TabIndex        =   15
      Top             =   870
      Width           =   3195
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblAppDate 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   6120
      TabIndex        =   14
      Top             =   570
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Index           =   1
      Left            =   5400
      TabIndex        =   13
      Top             =   570
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Left            =   450
      TabIndex        =   12
      Top             =   870
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Left            =   1245
      TabIndex        =   11
      Top             =   870
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   1245
      TabIndex        =   10
      Top             =   1200
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(外):"
      Height          =   180
      Index           =   0
      Left            =   1245
      TabIndex        =   9
      Top             =   1530
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   825
      TabIndex        =   8
      Top             =   570
      Width           =   765
   End
End
Attribute VB_Name = "frm060107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/11 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim intLastRow As Integer
'Add By Sindy 2022/5/12
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String
Dim m_Done As Boolean
Dim m_PrevForm As Form
'2022/5/12 END


'Add By Sindy 2022/5/12
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub SetGridHead()
    Dim i As Integer
    FixGrid MSHFlexGrid1
    With MSHFlexGrid1
        .Visible = False
        .Cols = 19
        .row = 0
        .col = 0: .ColWidth(.col) = 200: .Text = "v"
        .CellAlignment = flexAlignCenterCenter
        .col = 1: .ColWidth(.col) = 900: .Text = "收文日"
        .CellAlignment = flexAlignCenterCenter
        .col = 2: .ColWidth(.col) = 1300: .Text = "收文號"
        .CellAlignment = flexAlignCenterCenter
        .col = 3: .ColWidth(.col) = 1400: .Text = "案件性質"
        .CellAlignment = flexAlignCenterCenter
        .col = 4: .ColWidth(.col) = 1200: .Text = "承辦人"
        .CellAlignment = flexAlignCenterCenter
        .col = 5: .ColWidth(.col) = 1200: .Text = "核稿人"
        .CellAlignment = flexAlignCenterCenter
        .col = 6: .ColWidth(.col) = 1200: .Text = "完稿日"
        For i = 7 To .Cols - 1
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
    End With
End Sub

Private Sub ClearGrid()
    Dim rstGrid As New ADODB.Recordset, stSQL As String
    
    stSQL = "SELECT 0, 1,2,3,4,5,6,7,8, 9, 10, 11 FROM DUAL WHERE ROWNUM<1"
    rstGrid.CursorLocation = adUseClient
    rstGrid.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
    Set MSHFlexGrid1.Recordset = rstGrid
    SetGridHead
    Set rstGrid = Nothing
End Sub

Private Sub Check1_Click()
   ClearGrid
End Sub

Private Sub cmdExit_Click()
    blnIsFormBack = False
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim ii As Integer
    
    If MSHFlexGrid1.Rows < 2 Then Exit Sub
    
   'Add By Sindy 2019/5/10
   If m_strIR01 <> "" Then
      If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtCaseNo(1) & txtCaseNo(2) & txtCaseNo(3) & txtCaseNo(4) Then
         MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 & ")一致！"
         Exit Sub
      End If
   End If
   '2019/5/10 END
      
    With MSHFlexGrid1
        .Visible = False
        For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 0) = "v" Then Exit For
        Next ii
        .Visible = True
        If ii = .Rows Then
            MsgBox "請點選欲輸入資料！"
        Else
            '2006/3/10 ADD BY SONIA
            If .TextMatrix(ii, 4) = "" Then
               MsgBox "未輸入承辦人, 不可輸入核稿人 !!"
               Exit Sub
            End If
            '2006/3/10 END
            frm060107_1.bolTfOnly = IIf(Check1.Value = 1, True, False) 'Add by Morgan 2009/6/30
            'Add By Sindy 2022/5/12
            If Not m_PrevForm Is Nothing Then
               Call frm060107_1.SetParent(m_PrevForm)
            End If
            frm060107_1.m_strIR01 = m_strIR01
            frm060107_1.m_strIR02 = m_strIR02
            frm060107_1.m_strIR03 = m_strIR03
            frm060107_1.m_strIR04 = m_strIR04
            '2022/5/12 END
            frm060107_1.Show
            Call frm060107_1.SetData(MSHFlexGrid1.Recordset, ii)
            Me.Hide
        End If
        
    End With
   
End Sub

Public Sub SetGrid(Optional ByVal bolMsg As Boolean = True)

On Error GoTo flgErr

   Dim rstGrid As New ADODB.Recordset
   Dim stSQL As String
   Dim arrCaseNo(1 To 4) As String
   Dim stCon As String
   
   If Check1.Value = 0 Then
      stCon = " AND CP27 IS NULL"
   End If
   
   arrCaseNo(1) = txtCaseNo(1)
   arrCaseNo(2) = Right("000000" & txtCaseNo(2), 6)
   arrCaseNo(3) = Right("0" & txtCaseNo(3), 1)
   arrCaseNo(4) = Right("00" & txtCaseNo(4), 2)
   'Modify by Morgan 2005/9/8 核稿期限改抓ep08(原cp48)
   'Modify by Morgan 2007/6/1 加927
   'Modify by Morgan 2008/10/21 取消209,210 改由分案輸入齊備日控制
   'Modified by Lydia 2016/07/06 案件性質改('201','927')抓常數FCPHaveEP09
   If txtCaseNo(1) = "FCP" Then
      'Modified by Lydia 2018/05/07 +CP113工作時數
      stSQL = "SELECT '' V" & _
         ", DECODE(CP05,NULL,NULL,(SUBSTR(CP05,1,4)-1911)||SUBSTR(CP05,5,2)||SUBSTR(CP05,7,2)) CP05T" & _
         ", CP09,NVL(CPM03,CP10) CP10T" & _
         ", S1.ST02 CP14T, S2.ST02 EP04T" & _
         ", DECODE(EP09,NULL,NULL,(SUBSTR(EP09,1,4)-1911)||SUBSTR(EP09,5,2)||SUBSTR(EP09,7,2)) EP09T" & _
         ", DECODE(PA10,NULL,NULL,(SUBSTR(PA10,1,4)-1911)||SUBSTR(PA10,5,2)||SUBSTR(PA10,7,2)) PA10T" & _
         ", PA08, PTM03 PA08T, CP64, S1.ST15" & _
         ", PA05, PA06, PA07, CP05, CP06, CP10, CP14, EP04, EP09, PA10" & _
         ", DECODE(EP08,NULL,NULL,(SUBSTR(EP08,1,4)-1911)||SUBSTR(EP08,5,2)||SUBSTR(EP08,7,2)) CP48T" & _
         ", CP60, DECODE(CP27,NULL,NULL,CP27-19110000) CP27T,CP113" & _
         " FROM CASEPROGRESS, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1, STAFF S2, PatentTrademarkMap" & _
         " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
         " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
         " AND EP02=CP09" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
         " AND S1.ST01(+)=CP14 AND S2.ST01(+)=EP04" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
         " AND CP10 IN (" & FCPHaveEP09 & ") AND CP57 IS NULL" & stCon & _
         " AND PTM02=PA08 AND PTM01='1' "
         
   'Add by Morgan 2007/8/21
   'Modified by Lydia 2016/07/06 案件性質改('201','927')抓常數FCPHaveEP09
   ElseIf txtCaseNo(1) = "P" Then
      'Modified by Lydia 2018/05/07 +CP113工作時數
      stSQL = "SELECT '' V" & _
         ", DECODE(CP05,NULL,NULL,(SUBSTR(CP05,1,4)-1911)||SUBSTR(CP05,5,2)||SUBSTR(CP05,7,2)) CP05T" & _
         ", CP09,NVL(CPM03,CP10) CP10T" & _
         ", S1.ST02 CP14T, S2.ST02 EP04T" & _
         ", DECODE(EP09,NULL,NULL,(SUBSTR(EP09,1,4)-1911)||SUBSTR(EP09,5,2)||SUBSTR(EP09,7,2)) EP09T" & _
         ", DECODE(PA10,NULL,NULL,(SUBSTR(PA10,1,4)-1911)||SUBSTR(PA10,5,2)||SUBSTR(PA10,7,2)) PA10T" & _
         ", PA08, PTM03 PA08T, CP64, S1.ST15" & _
         ", PA05, PA06, PA07, CP05, CP06, CP10, CP14, EP04, EP09, PA10" & _
         ", DECODE(EP08,NULL,NULL,(SUBSTR(EP08,1,4)-1911)||SUBSTR(EP08,5,2)||SUBSTR(EP08,7,2)) CP48T" & _
         ", CP60, DECODE(CP27,NULL,NULL,CP27-19110000) CP27T,CP113" & _
         " FROM CASEPROGRESS, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1, STAFF S2, PatentTrademarkMap" & _
         " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
         " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
         " AND SUBSTR(CP12,1,1)='F' AND EP02=CP09" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
         " AND S1.ST01(+)=CP14 AND S2.ST01(+)=EP04" & _
         " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
         " AND CP10 IN (" & FCPHaveEP09 & ") AND CP57 IS NULL" & stCon & _
         " AND PTM02=PA08 AND PTM01='1' "
         
   'Add by Morgan 2007/8/21
   'Modified by Lydia 2016/07/06 案件性質改('201','927')抓常數FCPHaveEP09
   ElseIf txtCaseNo(1) = "PS" Then
      stSQL = "SELECT '' V" & _
         ", DECODE(CP05,NULL,NULL,(SUBSTR(CP05,1,4)-1911)||SUBSTR(CP05,5,2)||SUBSTR(CP05,7,2)) CP05T" & _
         ", CP09,NVL(CPM03,CP10) CP10T" & _
         ", S1.ST02 CP14T, S2.ST02 EP04T" & _
         ", DECODE(EP09,NULL,NULL,(SUBSTR(EP09,1,4)-1911)||SUBSTR(EP09,5,2)||SUBSTR(EP09,7,2)) EP09T" & _
         ", DECODE(SP10,NULL,NULL,(SUBSTR(SP10,1,4)-1911)||SUBSTR(SP10,5,2)||SUBSTR(SP10,7,2)) PA10T" & _
         ", '' PA08, '' PA08T, CP64, S1.ST15" & _
         ", SP05 PA05, SP06 PA06, SP07 PA07, CP05, CP06, CP10, CP14, EP04, EP09, SP10" & _
         ", DECODE(EP08,NULL,NULL,(SUBSTR(EP08,1,4)-1911)||SUBSTR(EP08,5,2)||SUBSTR(EP08,7,2)) CP48T" & _
         ", CP60, DECODE(CP27,NULL,NULL,CP27-19110000) CP27T" & _
         " FROM CASEPROGRESS, ENGINEERPROGRESS, SERVICEPRACTICE, CASEPROPERTYMAP, STAFF S1, STAFF S2" & _
         " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
         " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
         " AND substr(CP12,1,1)='F' and EP02=CP09" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
         " AND S1.ST01(+)=CP14 AND S2.ST01(+)=EP04" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
         " AND CP10 IN (" & FCPHaveEP09 & ") AND CP57 IS NULL" & stCon
         
   Else
      'Add by Morgan 2007/7/26 加FG
     'Modified by Lydia 2016/07/06 案件性質改('201','927')抓常數FCPHaveEP09
     'Modified by Lydia 2018/05/07 +CP113工作時數
      stSQL = "SELECT '' V" & _
         ", DECODE(CP05,NULL,NULL,(SUBSTR(CP05,1,4)-1911)||SUBSTR(CP05,5,2)||SUBSTR(CP05,7,2)) CP05T" & _
         ", CP09,NVL(CPM03,CP10) CP10T" & _
         ", S1.ST02 CP14T, S2.ST02 EP04T" & _
         ", DECODE(EP09,NULL,NULL,(SUBSTR(EP09,1,4)-1911)||SUBSTR(EP09,5,2)||SUBSTR(EP09,7,2)) EP09T" & _
         ", DECODE(SP10,NULL,NULL,(SUBSTR(SP10,1,4)-1911)||SUBSTR(SP10,5,2)||SUBSTR(SP10,7,2)) PA10T" & _
         ", '' PA08, '' PA08T, CP64, S1.ST15" & _
         ", SP05 PA05, SP06 PA06, SP07 PA07, CP05, CP06, CP10, CP14, EP04, EP09, SP10" & _
         ", DECODE(EP08,NULL,NULL,(SUBSTR(EP08,1,4)-1911)||SUBSTR(EP08,5,2)||SUBSTR(EP08,7,2)) CP48T" & _
         ", CP60, DECODE(CP27,NULL,NULL,CP27-19110000) CP27T,CP113" & _
         " FROM CASEPROGRESS, ENGINEERPROGRESS, SERVICEPRACTICE, CASEPROPERTYMAP, STAFF S1, STAFF S2" & _
         " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
         " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
         " AND EP02=CP09" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
         " AND S1.ST01(+)=CP14 AND S2.ST01(+)=EP04" & _
         " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
         " AND CP10 IN (" & FCPHaveEP09 & ") AND CP57 IS NULL" & stCon
         
   End If
      
   stSQL = stSQL & " ORDER BY CP05, CP09"

   rstGrid.CursorLocation = adUseClient
   rstGrid.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
    
   If rstGrid.RecordCount > 0 Then
      txtCaseNo(1) = arrCaseNo(1)
      txtCaseNo(2) = arrCaseNo(2)
      txtCaseNo(3) = arrCaseNo(3)
      txtCaseNo(4) = arrCaseNo(4)
      lblAppDate = "" & rstGrid.Fields("PA10T")
      lblCaseName(1) = "" & rstGrid.Fields("PA05")
      lblCaseName(2) = "" & rstGrid.Fields("PA06")
      lblCaseName(3) = "" & rstGrid.Fields("PA07")
   ElseIf bolMsg Then
      ShowNoData
   End If
   
   Set MSHFlexGrid1.Recordset = rstGrid
   SetGridHead
   Set rstGrid = Nothing
    
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If

End Sub

Private Sub cmdSearch_Click()
    SetGrid
    If Me.MSHFlexGrid1.Rows = 2 And Me.Visible = True Then
        MSHFlexGrid1.row = 1
        GridClick MSHFlexGrid1, intLastRow, 0
        cmdOK_Click
   End If
End Sub

Private Sub Form_Activate()
   Static bolActivated As Boolean
   If Not bolActivated Then
      bolActivated = True
      txtCaseNo(2).SetFocus
   End If
   
   'Added by Sindy 2022/5/12
   If m_strIR01 <> "" And m_Done = False Then
      txtCaseNo(1).Text = m_strCP01
      txtCaseNo(2).Text = m_strCP02
      txtCaseNo(3).Text = m_strCP03
      txtCaseNo(4).Text = m_strCP04
      Call cmdSearch_Click
      Call cmdOK_Click
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2022/5/12 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   ClearGrid
   lblAppDate = ""
   lblCaseName(1) = ""
   lblCaseName(2) = ""
   lblCaseName(3) = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Add By Sindy 2022/5/12
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2022/5/12 END
   
    Set frm060107 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK.SetFocus
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK.SetFocus
End Sub

Private Sub txtCaseNo_Change(Index As Integer)
    lblAppDate = ""
    lblCaseName(1) = ""
    lblCaseName(2) = ""
    lblCaseName(3) = ""
    ClearGrid
End Sub

Private Sub txtCaseNo_GotFocus(Index As Integer)
    TextInverse txtCaseNo(Index)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtCaseNo(Index).IMEMode = 2
   CloseIme
End Sub

Private Sub txtCaseNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
