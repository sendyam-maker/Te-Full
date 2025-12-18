VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm880014 
   BorderStyle     =   1  '單線固定
   Caption         =   "參考報價查詢"
   ClientHeight    =   4350
   ClientLeft      =   195
   ClientTop       =   2520
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7230
   Begin VB.CommandButton cmdOK 
      Caption         =   "離開(&X)"
      Height          =   400
      Index           =   0
      Left            =   6075
      TabIndex        =   0
      Top             =   120
      Width           =   1020
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3375
      Left            =   135
      TabIndex        =   1
      Top             =   600
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   6
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "報價日|本所案號|案件名稱|項目|年度|報價"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblRemark 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   2
      Top             =   4050
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "frm880014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/29 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'Create by Morgan 2008/5/16
Option Explicit

Public fmParent As Form
Public bNewFormat As Boolean
Dim m_iSelRow As Integer

Private Sub cmdOK_Click(Index As Integer)
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not bNewFormat Then SetDataListWidth
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub
'Added by Morgan 2016/1/5
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
PUB_ResetSalaryTimer fmParent
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm880014 = Nothing
End Sub

Private Sub SetDataListWidth()
   With grdDataList
   'Added by Lydia 2017/07/03 改變grid的欄位寬度
   If Me.Tag = "1" Then    'CFP案的答辨報價(CFP一般來函frm05010401_3)
      .ColWidth(0) = 800 '報價日
      .ColAlignment(0) = flexAlignCenterCenter
      .ColWidth(1) = 1400  '本所案號
      .ColAlignment(1) = flexAlignLeftCenter
      .ColWidth(2) = 1600  '案件名稱
      .ColAlignment(2) = flexAlignLeftCenter
      .ColWidth(3) = 960  '案件性質
      .ColAlignment(3) = flexAlignLeftCenter
      .ColWidth(4) = 3000   '報價
      .ColAlignment(4) = flexAlignLeftCenter
      '.ColWidth(5) = 0 'Removed by Morgan 2018/7/5 不可設定，因會導致欄位數.cols與實際資料欄數不同而使grdSelected執行錯誤
   Else
   'end 2017/07/03
      .FormatString = .FormatString
      .ColWidth(0) = 800
      .ColAlignment(0) = flexAlignCenterCenter
      .ColWidth(1) = 1400
      .ColAlignment(1) = flexAlignLeftCenter
      .ColWidth(2) = 1600
      .ColAlignment(2) = flexAlignLeftCenter
      .ColWidth(3) = 1200
      .ColAlignment(3) = flexAlignLeftCenter
      .ColWidth(4) = 600
      .ColAlignment(4) = flexAlignRightCenter
      .ColWidth(5) = 1000
      .ColAlignment(5) = flexAlignRightCenter
      'Added by Morgan 2021/7/30 +備註欄
      If .Recordset.Fields.Count > 6 Then
         .WordWrap = True
         .RowHeightMin = 500
         .ColWidth(6) = .Width - 300
         .ColAlignment(6) = flexAlignLeftCenter
      End If
      'end 2021/7/30
   End If 'end 2017/07/03
   End With
End Sub

Private Sub grdSelected(p_iRow As Integer)
   Dim lColor As Long, ii As Integer
   With grdDataList
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
   With grdDataList
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
'Added by Morgan 2016/1/5
Private Sub grdDataList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
PUB_ResetSalaryTimer fmParent
End Sub
