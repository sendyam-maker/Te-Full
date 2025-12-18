VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060119_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件進度查詢"
   ClientHeight    =   5736
   ClientLeft      =   120
   ClientTop       =   996
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9336
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7548
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   8376
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060119_1.frx":0000
      Left            =   960
      List            =   "frm060119_1.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   1260
      Width           =   615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3705
      Left            =   60
      TabIndex        =   0
      Top             =   1920
      Width           =   9165
      _ExtentX        =   16171
      _ExtentY        =   6541
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   13
      Top             =   1620
      Width           =   8280
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "14605;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   1710
      TabIndex        =   12
      Top             =   1260
      Width           =   7530
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "13282;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   11
      Top             =   930
      Width           =   3330
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5874;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   6000
      TabIndex        =   10
      Top             =   600
      Width           =   1680
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2963;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   7
      Top             =   600
      Width           =   3330
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5874;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1620
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利號數"
      Height          =   180
      Index           =   2
      Left            =   5160
      TabIndex        =   3
      Top             =   600
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   930
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   720
   End
End
Attribute VB_Name = "frm060119_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/29 Form2.0已修改
'Created by Morgan 2017/5/25 (改frm060101_2)
Option Explicit

Public fmParent As Form
Public strPatent As String

Dim pa(10) As String
Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer


Private Sub cmdok_Click(Index As Integer)
   Dim iRow As Integer
   fmParent.Tag = ""
   Select Case Index
      Case 0
         MSHFlexGrid1.col = 0
         For iRow = 1 To MSHFlexGrid1.Rows - 1
            MSHFlexGrid1.row = iRow
            If MSHFlexGrid1.Text = "v" Then
               If MSHFlexGrid1.TextMatrix(iRow, 1) <> "" Then
                  'Modified by Morgan 2018/10/23 +案件性質
                  'fmParent.Tag = MSHFlexGrid1.TextMatrix(iRow, 1)
                  fmParent.Tag = MSHFlexGrid1.TextMatrix(iRow, 1) & ":" & MSHFlexGrid1.TextMatrix(iRow, 8)
                  'end 2018/10/23
               End If
               Exit For
            End If
         Next
      Case 1
   End Select
   Unload Me
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(3) = pa(1)
      Case "英"
         Label2(3) = pa(2)
      Case "日"
         Label2(3) = pa(3)
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   ReadPatent
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060119_1 = Nothing
End Sub

Private Sub ReadPatent()
   Dim Lbl As Object, i As Integer
   Dim m_PA09 As String
   Dim bolMsg As Boolean 'Added by Morgan 2025/5/21
   
   For Each Lbl In Label2
      Lbl = ""
   Next
   Label2(0) = strPatent
   
   strExc(0) = "SELECT PA11,PA22,PA05,PA06,PA07,CU04,PA09,PA26,PA75,PA01,PA02,PA57 FROM PATENT,CUSTOMER WHERE " & ChgPatent(strPatent) & _
      " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      If Not IsNull(.Fields(0)) Then Label2(2) = .Fields(0)
      If Not IsNull(.Fields(1)) Then Label2(1) = .Fields(1)
      For i = 1 To 3
         If Not IsNull(.Fields(i + 1)) Then
            pa(i) = .Fields(i + 1)
         Else
            pa(i) = ""
         End If
      Next
      If Not IsNull(.Fields(5)) Then Label2(4) = .Fields(5)
      m_PA09 = .Fields("PA09")
      'Added by Morgan 2017/12/4 FCP57047 收到智慧局未出接洽單之來函: 歸卷、一般來函(延期受理、通知補文件)、通知實審日、通知公開 提醒 "收到智慧局來函需退承辦報告客戶" -- Sharon
      'Modified by Morgan 2025/5/21 +073780、073781、073782--劉維芩
      If .Fields("PA01") = "FCP" And (.Fields("PA02") = "057047" Or .Fields("PA02") = "073780" Or .Fields("PA02") = "073781" Or .Fields("PA02") = "073782") Then
            'Modified by Morgan 2025/5/21 訊息統一--Sharon
            'MsgBox "本案收到智慧局來函需退承辦報告客戶！", vbInformation, "提醒"
            bolMsg = True
            
      'Added by Morgan 2017/7/4 若申請人為X48637(大日本印刷股份有限公司),則彈訊息"收到智慧局受理通知,請立即交承辦寄代"--蔡秋舒
      'Modified by Morgan 2017/7/31 +Y34232--吳若芬
      'Modified by Morgan 2019/7/2 +Y52833--洪郁嵐
      'Modified by Lydia 2021/03/11 + Y5133301--葉子寧
      ElseIf .Fields("PA26") = "X48637000" Or .Fields("PA75") = "Y34232000" Or .Fields("PA75") = "Y52833000" Or .Fields("PA75") = "Y51333010" Then
         'Modified by Morgan 2025/5/21 訊息統一--Sharon
         'MsgBox "收到智慧局受理通知,請立即交承辦寄代!!", vbExclamation, "提醒"
         bolMsg = True
      
      'Added by Morgan 2022/4/12 +Y55105 --蘇暐嵐
      '智慧局所有來函彈跳提醒(除C類來函承辦人為工程師及已閉卷的案子)
      ElseIf .Fields("PA75") = "Y55105000" And "" & .Fields("PA57") = "" Then
         'Modified by Morgan 2025/5/21 訊息統一--Sharon
         'MsgBox "收到智慧局來函,請通知承辦寄代！", vbInformation
         bolMsg = True
         
      End If
      'end 2017/7/4
      
      'Added by Morgan 2025/5/21
      If bolMsg = True Then
         MsgBox "收到智慧局來函,請通知承辦寄代！", vbInformation
      End If
      'end 2025/5/21
      
   End If
   End With
   
   'Modified by Morgan 2017/7/3 +通知面詢1401--敏莉
   strExc(0) = "SELECT '',CP09," & SQLDate("CP05") & "," & IIf(m_PA09 = "000", "", " (CPM04)") & " CPM03," & SQLDate("CP27") & ",CP24,CP19," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,CP40),CP10 FROM CASEPROGRESS,CASEPROPERTYMAP" & _
      " WHERE " & ChgCaseprogress(strPatent) & " AND CP09<'C' and CP10<>'701' AND cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " UNION SELECT '',CP09," & SQLDate("CP05") & "," & IIf(m_PA09 = "000", "", " (CPM04)") & " CPM03," & SQLDate("CP27") & ",CP24,CP19," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,CP40),CP10 FROM CASEPROGRESS,CASEPROPERTYMAP" & _
      " WHERE " & ChgCaseprogress(strPatent) & " AND CP10='1401' AND cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " UNION  SELECT '',CP09," & SQLDate("CP05") & "," & IIf(m_PA09 = "000", "", " (CPM04)") & " CPM03," & SQLDate("CP27") & ",CP24,CP19," & _
      " NVL(CU04,NVL(CU05,CU06)),CP10 FROM CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER" & _
      " WHERE " & ChgCaseprogress(strPatent) & " AND CP09<'C' AND CP10='701'AND cpm01(+)=cp01 and cpm02(+)=cp10 " & _
      " AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
End Sub

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
      .col = 2: .ColWidth(2) = 1000: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1000: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1200: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "結果"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1400: .Text = "後金"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 1400: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      For i = 8 To .Cols - 1
         .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub
