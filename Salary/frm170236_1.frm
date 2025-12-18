VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170236_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "翻譯費明細"
   ClientHeight    =   5775
   ClientLeft      =   195
   ClientTop       =   2520
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8115
   Begin VB.CommandButton cmdOK 
      Caption         =   "離開(&X)"
      Height          =   400
      Index           =   0
      Left            =   6840
      TabIndex        =   0
      Top             =   50
      Width           =   1020
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3555
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6271
      _Version        =   393216
      BackColor       =   -2147483628
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblRate 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "9,999"
      Height          =   180
      Index           =   2
      Left            =   2085
      TabIndex        =   17
      Top             =   4400
      Width           =   405
   End
   Begin VB.Label lblRate 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00FFFFFF&
      Caption         =   "9,999"
      Height          =   180
      Index           =   1
      Left            =   2085
      TabIndex        =   16
      Top             =   4200
      Width           =   405
   End
   Begin VB.Label lblAMT 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "999,999"
      Height          =   180
      Index           =   3
      Left            =   7305
      TabIndex        =   14
      Top             =   4680
      Width           =   585
   End
   Begin VB.Label lblAMT 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "999,999"
      Height          =   180
      Index           =   2
      Left            =   7305
      TabIndex        =   13
      Top             =   4440
      Width           =   585
   End
   Begin VB.Label lblAMT 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "999,999"
      Height          =   180
      Index           =   1
      Left            =   7305
      TabIndex        =   12
      Top             =   4200
      Width           =   585
   End
   Begin VB.Label lblAMT 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "999,999"
      Height          =   180
      Index           =   4
      Left            =   7305
      TabIndex        =   11
      Top             =   4920
      Width           =   585
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "實際　金額："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   10
      Top             =   4920
      Width           =   1170
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "代扣補充保費："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5685
      TabIndex        =   9
      Top             =   4680
      Width           =   1365
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "代扣所得稅："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   8
      Top             =   4440
      Width           =   1170
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "總　　　計："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   7
      Top             =   4200
      Width           =   1170
   End
   Begin VB.Label lblRemark 
      AutoSize        =   -1  'True
      Caption         =   "　　　6. ""***""  表示該案已扣除相似折扣部分及瑕疵折扣部分"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   5400
      Width           =   5175
   End
   Begin VB.Label lblRemark 
      AutoSize        =   -1  'True
      Caption         =   "　　　5. ""**""    表示該案已扣除瑕疵折扣部分"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   3810
   End
   Begin VB.Label lblRemark 
      AutoSize        =   -1  'True
      Caption         =   "　　　4. ""*""      表示該案已扣除相似折扣部分"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   4920
      Width           =   3810
   End
   Begin VB.Label lblRemark 
      AutoSize        =   -1  'True
      Caption         =   "　　　3. ""@""    表示該案已計算加成部分"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   3405
   End
   Begin VB.Label lblRemark 
      AutoSize        =   -1  'True
      Caption         =   "　　　2. 中打費 rate：          ／每千個中文字"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   3660
   End
   Begin VB.Label lblRemark 
      AutoSize        =   -1  'True
      Caption         =   "備註：1. 翻譯費 rate：          ／每千個日文字"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   4200
      Width           =   3660
   End
End
Attribute VB_Name = "frm170236_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/14 Form2.0已修改
'create by sonia 2016/8/12
Option Explicit

Public fmParent As Form
Public bNewFormat As Boolean
Public FstYrMn As String
Dim m_iSelRow As Integer

Private Sub cmdok_Click(Index As Integer)
   Unload Me
End Sub

'Private Sub Form_Activate()
'   If Not bNewFormat Then SetDataListWidth
'End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   QueryData
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_ResetSalaryTimer fmParent
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170236_1 = Nothing
End Sub

'Private Sub SetDataListWidth()
'   With grdDataList
'      .FormatString = .FormatString
'      .ColWidth(0) = 800
'      .ColAlignment(0) = flexAlignCenterCenter
'      .ColWidth(1) = 1400
'      .ColAlignment(1) = flexAlignLeftCenter
'      .ColWidth(2) = 1600
'      .ColAlignment(2) = flexAlignLeftCenter
'      .ColWidth(3) = 1200
'      .ColAlignment(3) = flexAlignLeftCenter
'      .ColWidth(4) = 600
'      .ColAlignment(4) = flexAlignRightCenter
'      .ColWidth(5) = 1000
'      .ColAlignment(5) = flexAlignRightCenter
'   End With
'End Sub

Private Sub grdSelected(p_iRow As Integer)
   Dim lColor As Long, ii As Integer
   With GrdDataList
      .row = p_iRow
      .col = 0
      If .CellBackColor = &H80000018 Then
         m_iSelRow = .row
         lColor = &HFFC0C0
      Else
         m_iSelRow = -1
         lColor = &H80000018
      End If
      For ii = 0 To .Cols - 1
         .col = ii
         .CellBackColor = lColor
      Next
   End With
End Sub

Private Sub GrdDataList_Click()
   Dim iRow As Integer
   If bNewFormat Then Exit Sub
   With GrdDataList
      If .MouseRow > 0 And .MouseRow < .Rows Then
         .Visible = False
         iRow = .MouseRow
         If m_iSelRow > 0 Then
            grdSelected m_iSelRow
         End If
         If m_iSelRow <> iRow Then
            grdSelected iRow
         End If
         .Visible = True
      End If
   End With
End Sub

Private Sub grdDataList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_ResetSalaryTimer fmParent
End Sub

'查詢資料
Private Sub QueryData()
   
   strExc(1) = (FstYrMn \ 100 - 1911) & "年" & Val(Right(FstYrMn, 2)) & "月份"
   Me.Caption = strExc(1) & "翻譯費明細"
   
   strExc(0) = "select decode(nvl(tf05,0),0,null,100,null,'*')||decode(nvl(tf06,0),0,null,100,null,'**')||decode(nvl(tf18,0),0,null,100,null,'@')||substr(a1p17, 1, Length(a1p17) - 9)||'-'||substr(a1p17, Length(a1p17) - 8, 6)||decode(substr(a1p17, Length(a1p17) - 2),'000',null,substr(a1p17, Length(a1p17) - 2, 1)||'-'||substr(a1p17, Length(a1p17) - 1, 2)) 本所案號," & _
               " to_char(TF03,'999999') 日文字數,to_char(decode(nvl(tf03,0),0,(nvl(TF02,0)+nvl(TF04,0)) * nvl(SPR02,0) / 1000,(nvl(tf03,0) * nvl(SPR03,0) / 1000)),'99G999G999G999') 翻譯費," & _
               " to_char(nvl(TF02,0)+nvl(TF04,0),'999999') 中文字數,to_char((nvl(TF02,0)+nvl(TF04,0)) * nvl(SPR04,0) / 1000,'999G999') 中打費,to_char(a1p07,'999G999G999G999') 小計" & _
               "  From acc1p0, staff_idmap, staff, TRANSFEE, STAFF_PAYRATE, ACC250" & _
               " where substr(a1p18+19110000,1,6)='" & FstYrMn & "' and a1p15='" & frm170236.cboUser.Tag & "' and a1p05='6130' and sim02(+)=a1p15 and st01(+)=sim01 AND TF07(+)=A1P04 AND TF14(+)=A1P17" & _
               "   AND SPR01(+)=A1P15 and A2502(+)='5' AND A2503(+)=A1P15 AND A2505(+)=A1P04 order by a1p17,a1p15"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With GrdDataList
      .FixedCols = 0
      Set .Recordset = RsTemp.Clone
      End With
      SetGrid
   End If

   strExc(0) = "select sum(a1p07) 小計,nvl(OD06,0) 代扣所得稅,nvl(OD13,0) 代扣補充保費,sum(a1p07)-nvl(od06,0)-nvl(od13,0) 實際金額" & _
               "  from acc1p0,staff_idmap,staff,ACC250,othersalarydata " & _
               " where substr(a1p18+19110000,1,6)='" & FstYrMn & "' and a1p15='" & frm170236.cboUser.Tag & "' and a1p05='6130' and sim02(+)=a1p15 and st01(+)=sim01" & _
               "   and A2502(+)='5' AND A2503(+)=A1P15 AND A2505(+)=A1P04 AND OD03(+)=a1p15 AND OD02(+)=A1P18+19110000 group by od06,od13"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.RecordCount = 1 Then
         lblAMT(1) = "" & RsTemp.Fields(0)
         lblAMT(2) = "" & RsTemp.Fields(1)
         lblAMT(3) = "" & RsTemp.Fields(2)
         lblAMT(4) = "" & RsTemp.Fields(3)
      Else
         lblAMT(1) = 0
         lblAMT(2) = 0
         lblAMT(3) = 0
         lblAMT(4) = 0
      End If
   End If
   
   strExc(0) = "select distinct decode(nvl(tf03,0),0,spr02,spr03),spr04,nvl(tf03,0)" & _
               "  from acc1p0,staff_idmap,staff,TRANSFEE,STAFF_PAYRATE,ACC250 " & _
               " where substr(a1p18+19110000,1,6)='" & FstYrMn & "' and a1p15='" & frm170236.cboUser.Tag & "' and a1p05='6130' and sim02(+)=a1p15 and st01(+)=sim01" & _
               "   AND TF07(+)=A1P04 AND TF14(+)=A1P17 AND SPR01(+)=A1P15 and A2502(+)='5' AND A2503(+)=A1P15 AND A2505(+)=A1P04"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      lblRate(1) = "" & RsTemp.Fields(0)
      lblRate(2) = "" & RsTemp.Fields(1)
      If Val("" & RsTemp.Fields(2)) > 0 Then
         lblRemark(1) = "備註：1. 翻譯費 rate：          ／每千個日文字"
      Else
         lblRemark(1) = "備註：1. 翻譯費 rate：          ／每千個中文字"
      End If
   End If
   lblRate(1) = Format(lblRate(1), "#,##0")
   lblRate(2) = Format(lblRate(2), "#,##0")
   lblAMT(1) = Format(lblAMT(1), "###,##0")
   lblAMT(2) = Format(lblAMT(2), "###,##0")
   lblAMT(3) = Format(lblAMT(3), "###,##0")
   lblAMT(4) = Format(lblAMT(4), "###,##0")
End Sub

Private Sub SetGrid()
Dim iCol As Integer
   
   With GrdDataList
      .Visible = False
      .ForeColor = frm170236.lblTest.ForeColor   '依前畫面之濃淡設定
      .Cols = .Recordset.Fields.Count
      .ColAlignmentFixed = flexAlignCenterCenter
      .ColWidth(0) = 2000
      .ColAlignment(0) = flexAlignCenterCenter
      .ColWidth(1) = 1000
      .ColAlignment(1) = flexAlignRightCenter
      .ColWidth(2) = 1200
      .ColAlignment(2) = flexAlignRightCenter
      .ColWidth(3) = 1000
      .ColAlignment(3) = flexAlignCenterCenter
      .ColWidth(4) = 1200
      .ColAlignment(4) = flexAlignRightCenter
      .ColWidth(5) = 1350
      .ColAlignment(5) = flexAlignRightCenter
      iCol = 5
      For intI = iCol + 1 To .Cols - 1
         .ColWidth(intI) = 0
      Next
      .MergeCol(0) = True
      .MergeCells = flexMergeFree
      .Visible = True
   End With

   lblRate(1).ForeColor = frm170236.lblTest.ForeColor   '依前畫面之濃淡設定
   lblRate(2).ForeColor = lblRate(1).ForeColor
   lblAMT(1).ForeColor = lblRate(1).ForeColor
   lblAMT(2).ForeColor = lblRate(1).ForeColor
   lblAMT(3).ForeColor = lblRate(1).ForeColor
   lblAMT(4).ForeColor = lblRate(1).ForeColor
   
End Sub

