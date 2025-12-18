VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050408_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "互惠代理人案件統計表(案件明細)"
   ClientHeight    =   5976
   ClientLeft      =   168
   ClientTop       =   960
   ClientWidth     =   6564
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5976
   ScaleWidth      =   6564
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5130
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   90
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3825
      Left            =   90
      TabIndex        =   1
      Top             =   1470
      Width           =   6345
      _ExtentX        =   11197
      _ExtentY        =   6752
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   4
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label lblContact 
      Height          =   300
      Left            =   1170
      TabIndex        =   15
      Top             =   645
      Width           =   3000
      VariousPropertyBits=   27
      Size            =   "5292;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAgentName 
      Height          =   300
      Left            =   1170
      TabIndex        =   14
      Top             =   360
      Width           =   3000
      VariousPropertyBits=   27
      Size            =   "5292;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "2.案件之盈虧均結算至系統日為止"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   405
      TabIndex        =   13
      Top             =   5670
      Width           =   2865
   End
   Begin VB.Label lblPS 
      AutoSize        =   -1  'True
      Caption         =   "PS: 1.游標所到會變黑色的資料列表示可點選顯示程序盈虧資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   5400
      Width           =   5310
   End
   Begin VB.Label lblNetTot 
      Height          =   180
      Left            =   4950
      TabIndex        =   11
      Top             =   930
      Width           =   1305
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "總盈虧:"
      Height          =   180
      Left            =   4320
      TabIndex        =   10
      Top             =   930
      Width           =   585
   End
   Begin VB.Label lblCondition 
      Height          =   255
      Left            =   1170
      TabIndex        =   9
      Top             =   1170
      Width           =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "資料區間:"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label lblNation 
      Height          =   255
      Left            =   1170
      TabIndex        =   7
      Top             =   930
      Width           =   3000
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "國籍:"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   906
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人:"
      Height          =   255
      Left            =   540
      TabIndex        =   5
      Top             =   644
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "代理人名稱:"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   382
      Width           =   945
   End
   Begin VB.Label lblAgentNo 
      Height          =   255
      Left            =   1170
      TabIndex        =   3
      Top             =   120
      Width           =   3000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "代理人代號:"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frm050408_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/11 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblAgentName、lblContact
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2008/3/4
Option Explicit

Dim m_iRow As Integer
Public m_bolNoDetail As Boolean


'p_bolHeaderOnly:是否只設定表頭 true=是 false=資料一併清除
Public Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim ii As Integer
   With grdDataList
      .Visible = False
      .WordWrap = False
      If p_bolHeaderOnly = False Then
         .Clear
         .Rows = 2: .FixedRows = 1: .FixedCols = 0
      End If
      .row = 0
      .RowHeight(0) = 250
      .RowHeightMin = 250
      ii = 0
      .col = ii: .ColWidth(.col) = 1320: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 3500: .Text = "案件名稱"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      ii = ii + 1
      .col = ii: .ColWidth(.col) = 1140: .Text = "案件盈虧"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      For ii = 3 To .Cols - 1
         .ColWidth(ii) = 0
      Next
      .Refresh
      .Visible = True
   End With
End Sub

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Unload Me
   End Select
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   
   'Added by Lydia 2025/06/06
   If frm050408.Tag <> "" Then
      Me.Caption = "互惠期間統計表(案件明細)"
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050408_2 = Nothing
End Sub

Private Sub GetData(ByVal p_CaseNo As String)
   Dim stSQL As String, stVTB As String
   p_CaseNo = Replace(p_CaseNo, "-", "")
   'modify by sonia 2018/4/10 +CP27
   stVTB = "select cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
      ",nvl(min(decode(nvl(cp16,0)-nvl(cp77,0),0,0,nvl(cp16,0)-nvl(CP18,0)*1000)),0)-nvl(sum(decode(a1512,null,decode(a1901,null,AXF15,AXF04*A1906),AXF04*A1G03)),0) C1,cp27" & _
      " From caseprogress a, acc151, acc150,acc190,acc1g0" & _
      " Where " & ChgCaseprogress(p_CaseNo) & _
      " and (cp16>0 or cp61 is not null)" & _
      " and axf02(+)=cp09 and a1501(+)=axf01 and a1507 is null" & _
      " and A1902(+)=AXF01 AND A1G01(+)=A1512" & _
      " group by cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp27"
      
   'modify by sonia 2018/4/10 +CP27
   strExc(0) = "select sqldatet(cp05),cp09,cpm03,to_char(nvl(C1,0),'9,999,999.00') C02,sqldatet(cp27)" & _
         " from (" & stVTB & ") A,casepropertymap" & _
         " where cpm01(+)=A.cp01 and cpm02(+)=A.cp10 order by cp05"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With frm050408_4
         .lblCaseNo = grdDataList.TextMatrix(m_iRow, 0)
         .lblCaseName = grdDataList.TextMatrix(m_iRow, 1)
         .lblNetTot = grdDataList.TextMatrix(m_iRow, 2)
         'Add by Morgan 2008/7/22
         If Left(.lblCaseNo, 3) = "CFP" Then
            If grdDataList.TextMatrix(m_iRow, 4) <> "" Then
               .lblSecurFound = "安全基金：" & grdDataList.TextMatrix(m_iRow, 4)
            End If
         End If
         'end 2008/7/22
         .grdDataList.Visible = False
         .SetDataListWidth
         Set .grdDataList.Recordset = RsTemp.Clone
         .SetDataListWidth True
         .Show vbModal
      End With
   End If
End Sub

Private Sub GrdDataList_Click()
   If m_iRow <> 0 Then
      GetData grdDataList.TextMatrix(m_iRow, 0)
   End If
End Sub

Private Sub grdDataList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If m_bolNoDetail = True Then Exit Sub
   
   Dim iRow As Integer, iCol As Integer, lBackColor As Long
   Dim ii As Integer
   
   With grdDataList
      iRow = .MouseRow
      iCol = .MouseCol
      
      If iRow = m_iRow Then
         Exit Sub
      End If
      If m_iRow <> 0 Then
         .row = m_iRow
         For iCol = 0 To .Cols - 1
            .col = iCol
            .CellForeColor = .ForeColor
            .CellBackColor = .BackColor
         Next
         m_iRow = 0
      End If
      
      If iRow > 0 Then
         .row = iRow
         For iCol = 0 To .Cols - 1
            .col = iCol
            .CellForeColor = .BackColor
            .CellBackColor = .ForeColor
         Next
         m_iRow = .row
      End If
   End With
End Sub
