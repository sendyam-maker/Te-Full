VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010601_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "實審通知日輸入"
   ClientHeight    =   5070
   ClientLeft      =   135
   ClientTop       =   930
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   3
      Left            =   6264
      TabIndex        =   0
      Top             =   72
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "內部收文(&E)"
      Height          =   400
      Index           =   0
      Left            =   5040
      TabIndex        =   3
      Top             =   72
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   7
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8328
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7104
      TabIndex        =   1
      Top             =   72
      Width           =   1200
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm06010601_1.frx":0000
      Left            =   960
      List            =   "frm06010601_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   930
      Width           =   615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2715
      Left            =   60
      TabIndex        =   4
      Top             =   2280
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   4789
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   22
      Top             =   930
      Width           =   768
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3360
      TabIndex        =   21
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   120
      TabIndex        =   19
      Top             =   1290
      Width           =   768
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Left            =   3360
      TabIndex        =   18
      Top             =   1290
      Width           =   768
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "代理人:"
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   1620
      Width           =   588
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   120
      TabIndex        =   16
      Top             =   1950
      Width           =   948
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   4170
      TabIndex        =   15
      Top             =   600
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4075;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   1620
      TabIndex        =   14
      Top             =   930
      Width           =   7560
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "13335;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   13
      Top             =   1290
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4075;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   12
      Top             =   1290
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4075;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   11
      Top             =   1620
      Width           =   8220
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "14499;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   5
      Left            =   1140
      TabIndex        =   10
      Top             =   1950
      Width           =   2310
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4075;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm06010601_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/18 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim intWhere As Integer
Dim intLastRow As Integer


Private Sub cmdOK_Click(Index As Integer)
 Dim i As Integer, bolChk As Boolean
   Select Case Index
      Case 0
         mdiMain.mnu1102_Click 1
      Case 1
         frm06010601.Show
         Unload Me
      Case 2
         Unload frm06010601
         Unload Me
      Case 3
         With MSHFlexGrid1
            .col = 0
            For i = 1 To .Rows - 1
               If .TextMatrix(i, 0) = "v" Then
                  bolChk = True
                  strExc(2) = .TextMatrix(i, 1) '收文號
                  strExc(3) = .TextMatrix(i, 5) '案件性質代號
                  strExc(4) = .TextMatrix(i, 6) '業務區別
                  strExc(5) = .TextMatrix(i, 7) '智權人員代號
                  strExc(6) = .TextMatrix(i, 2) '案件性質
                  strExc(7) = .TextMatrix(i, 3) '發文日
                  '93.2.19 ADD BY SONIA 發明申請案檢查實體審查是否已收文已發文
                  'If strExc(3) = 發明申請 Then 'Remove by Morgan 2011/10/20 不必限制因為可能會有分割等其他案件性質
                  'Added by Morgan 2012/12/24
                  '發明案才要檢查--靜芳確認
                  If pa(8) = "1" And (strExc(3) = 發明申請 Or strExc(3) = 分割) Then
                     'Modified by Morgan 2013/10/16 +435續行母案再審
                     strExc(0) = "SELECT * FROM CASEPROGRESS WHERE " & _
                        ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP27 IS NOT NULL AND CP10 in ('416','435') AND CP57 IS NULL "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If intI = 0 Then
                        MsgBox "此案未提實體審查, 不可輸入實審通知 !", vbInformation
                        Exit Sub
                     End If
                     
'Removed by Morgan 2013/1/11 取消,102新法主動修正無期限--毓芳
'
'                     '2007/8/6 ADD BY SONIA 第三人提實審詢問是否管制三個月主動修正
'                     frm06010601_2.m_203check = False
'                     strExc(0) = "select CP50,CP51,CP52 from caseprogress" & _
'                        " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
'                        " and cp10='416' and cp27 is not null and cp50||cp51||cp52 is not null "
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        If MsgBox("本案為第三人" & RsTemp.Fields("cp50") & " " & RsTemp.Fields("cp51") & " " & RsTemp.Fields("cp52") & "提實審，是否要管制三個月主動修正期限？", vbYesNo + vbDefaultButton1) = vbYes Then
'                           frm06010601_2.m_203check = True
'                        End If
'                     End If
'                     '2007/8/6 END
'
'end 2013/1/11

                  End If
                  '93.2.19 END
                  Exit For
               End If
            Next
         End With
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
         Else
            frm06010601_1.Hide
            'Added by Morgan 2017/5/9 電子公文
            frm06010601_2.m_DocWord = frm06010601.m_DocWord
            frm06010601_2.m_DocNo = frm06010601.m_DocNo
            frm06010601_2.m_AppNo = frm06010601.m_AppNo
            'end 2017/5/9
            frm06010601_2.Show
            frm06010601_2.QueryData
         End If
   End Select
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(1) = pa(5)
      Case "英"
         Label2(1) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
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
   'ReadPatent
   'Combo1.ListIndex = 0
   'Label2(5) = Frm06010601.Text5
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010601_1 = Nothing
End Sub

Public Function QueryData() As Boolean
   QueryData = False
   ReadPatent
   If MSHFlexGrid1.Rows > 1 Then
      QueryData = True
      Combo1.ListIndex = 0
      Label2(5) = frm06010601.Text5
   End If
End Function

'************************************************
' 取回專利基本資料及收文資料
'
'************************************************
Private Sub ReadPatent()
 Dim Lbl As Object, i As Integer, strTempName As String
   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label2(1) = pa(5)
      If pa(8) <> "" Then ChgType (2) ' Label2(2)
      If pa(9) <> "" Then ChgType (3) ' Label2(3)
      If pa(75) <> "" Then ChgType (4) ' Label2(4)
      Label2(0) = pa(11)
   End If
   ' 91.09.13 modify by louis
   'strExc(0) = "SELECT '',CP09,CPM03," & SQLDate("CP27") & ",CP08,CP10,CP12,CP13 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE " & _
   '   ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP27 IS NOT NULL AND " & _
   '   "CP10 IN ('101','102','103','104','105','107') AND CP09 NOT IN " & _
   '   "(SELECT CP43 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND " & _
   '   "CP27 IS NOT NULL AND (CP10='" & 通知實審日 & "' OR SUBSTR(CP10,1,1)='3')) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
   '   "Union " & _
   '   "SELECT '',CP09,CPM03," & SQLDate("CP27") & ",CP08,CP10,CP12,CP13 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE " & _
   '   ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP27 IS NOT NULL AND " & _
   '   "SUBSTR(CP10,1,1) IN ('3','8') AND CP09 NOT IN " & _
   '   "(SELECT CP43 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND " & _
   '   "CP27 IS NOT NULL AND CP10='" & 通知實審日 & "') AND CP01=CPM01(+) AND CP10=CPM02(+)"
   
   'Modify by Morgan 2009/8/25 若有3xx性質的收文沒有相關總收文號時會無資料 Ex.FCP-33378
   'strExc(0) = "SELECT '',CP09,CPM03," & SQLDate("CP27") & ",CP08,CP10,CP12,CP13,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
               "FROM CASEPROGRESS,CASEPROPERTYMAP WHERE " & _
                  ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP27 IS NOT NULL AND " & _
                  "CP10 IN ('101','102','103','104','105','107') AND CP09 NOT IN " & _
                  "(SELECT CP43 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND " & _
                  "CP27 IS NOT NULL AND (CP10='" & 通知實審日 & "' OR CP10='1217' OR SUBSTR(CP10,1,1)='3')) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
               "Union SELECT '',CP09,CPM03," & SQLDate("CP27") & ",CP08,CP10,CP12,CP13,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
               "FROM CASEPROGRESS,CASEPROPERTYMAP WHERE " & _
                  ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP27 IS NOT NULL AND " & _
                  "SUBSTR(CP10,1,1) IN ('3','8') AND CP09 NOT IN " & _
                  "(SELECT CP43 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND " & _
                  "CP27 IS NOT NULL AND(CP10='" & 通知實審日 & "' OR CP10='1217')) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
               "ORDER BY SORTFIELD DESC "
   'Modified by Morgan 2012/12/20 +衍生設計125,改請衍生設計308()
   'Modified by Morgan 2013/5/9 +402
   'modify by sonia 2017/6/23 +415專利權延長 FCP-027702
   'Memo by Lydia 2021/04/28 增加欄位,請在GridHead設定是否顯示
   strExc(0) = "SELECT '',CP09,CPM03," & SQLDate("CP27") & ",CP08,CP10,CP12,CP13,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
               " FROM CASEPROGRESS a,CASEPROPERTYMAP" & _
               " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP27>0" & _
               " AND ( CP10 IN ('101','102','103','104','105','107','125','402','415') or SUBSTR(CP10,1,1) IN ('3','8'))" & _
               " AND not exists(SELECT * FROM CASEPROGRESS b WHERE b.cp43=a.cp09 and b.cp27>0 " & _
               " AND (b.CP10 in ('" & 通知實審日 & "','1217') OR SUBSTR(b.CP10,1,1)='3'))" & _
               " AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
               "ORDER BY SORTFIELD DESC "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   ' 90.08.22 modify by louis
   If MSHFlexGrid1.Rows > 1 Then
      MSHFlexGrid1.row = 1
      GridClick MSHFlexGrid1, intLastRow, 0
   End If
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
        'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetAgent(pA(75), strTempName) Then
         If ClsPDGetAgent(pa(75), strTempName) Then
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
      .col = 4: .ColWidth(4) = 2000: .Text = "機關文號"
      'Modified by Lydia 2021/04/28
      'For i = 5 To 7
      For i = 5 To 8
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
End Sub
