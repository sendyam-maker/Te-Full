VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm072002 
   BorderStyle     =   1  '單線固定
   Caption         =   "庭期資料查詢"
   ClientHeight    =   5820
   ClientLeft      =   90
   ClientTop       =   585
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdQueryFile 
      Caption         =   "卷宗區"
      Height          =   345
      Left            =   6330
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   70
      Width           =   915
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   8424
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   345
      Left            =   7296
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   70
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5172
      Left            =   120
      TabIndex        =   2
      Top             =   528
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   9128
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   16
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
      _Band(0).Cols   =   16
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      Caption         =   "註：開庭日期欄若有**，表示已取消庭期！"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   330
      Width           =   6105
   End
End
Attribute VB_Name = "frm072002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/15 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Private Sub cmdBack_Click()
   frm072001.Show
   Unload Me
End Sub

Private Sub cmdEnd_Click()
   Unload frm072001
   Unload Me
End Sub

'Add By Sindy 2016/6/28 卷宗區
Private Sub cmdQueryFile_Click()
   Call PubShowNextData
End Sub

Public Sub PubShowNextData()
Dim ii As Integer, jj As Integer
   
   For ii = 1 To MSHFlexGrid1.Rows - 1
      If MSHFlexGrid1.TextMatrix(ii, 0) = "V" Then
         'Screen.MousePointer = vbHourglass
         MSHFlexGrid1.TextMatrix(ii, 0) = ""
         MSHFlexGrid1.row = ii
         For jj = 0 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.col = jj
            MSHFlexGrid1.CellBackColor = QBColor(15)
         Next jj
         frm100101_L.m_strKey = MSHFlexGrid1.TextMatrix(ii, 14)
         frm100101_L.SetParent Me
         If frm100101_L.QueryData = True Then
            frm100101_L.Show
            Me.Hide
         Else
            Unload frm100101_L
         End If
         Exit Sub
         'Screen.MousePointer = vbDefault
      End If
   Next ii
End Sub

Private Sub Form_Load()
Dim i As Integer, strText As String
   
   MoveFormToCenter Me
   Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   'Add By Sindy 2011/6/17
   Screen.MousePointer = 11
   If MSHFlexGrid1.Rows > 0 Then
      For i = 1 To MSHFlexGrid1.Rows - 1
         If MSHFlexGrid1.TextMatrix(i, 14) <> "" Then '總收文號
            strText = ""
            strSql = "select distinct st02 from " & _
                     "(select st02 from caselawer,staff " & _
                     "where cl01='" & MSHFlexGrid1.TextMatrix(i, 14) & "' " & _
                     "and cl02=st01(+) " & _
                     "Union " & _
                     "select st02 from caseprogress,staff " & _
                     "where cp09 in (select a0n02 from acc0n0 where a0n02<>a0n01 and a0n01 in (select cp43 from caseprogress where cp09='" & MSHFlexGrid1.TextMatrix(i, 14) & "')) " & _
                     "and cp14=st01(+)) "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  If strText <> "" Then strText = strText & "、"
                  strText = strText & RsTemp(0)
                  RsTemp.MoveNext
               Loop
            End If
            If strText <> "" Then MSHFlexGrid1.TextMatrix(i, 15) = strText '其他開庭律師
         End If
      Next i
   End If
   Screen.MousePointer = 0
   '2011/6/17 End
End Sub

Private Sub GridHead()
 Dim i As Integer
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0
      .MergeCells = flexMergeRestrictRows
      .MergeRow(0) = True
      .col = 0: .ColWidth(0) = 200: .Text = "V" 'Add By Sindy 2016/6/27
      .col = 1: .ColWidth(1) = 1550: .Text = "本所案號"
      .col = 2: .ColWidth(2) = 1200: .Text = "案件名稱"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 800: .Text = "收受日"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 850: .Text = "開庭人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1000: .Text = "法院"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1000: .Text = "法院案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 800: .Text = "開庭日期"
      .col = 8: .ColWidth(8) = 800: .Text = "開庭時間"
      .col = 9: .ColWidth(9) = 800: .Text = "開庭別" 'Add By Sindy 2011/10/20
      .col = 10: .ColWidth(10) = 800: .Text = "開庭種類"
      .col = 11: .ColWidth(11) = 1500: .Text = "法官"
      .col = 12: .ColWidth(12) = 1500: .Text = "檢察官"
      .col = 13: .ColWidth(13) = 1500: .Text = "備註"
      .col = 14: .ColWidth(14) = 0: .Text = ""
      .col = 15: .ColWidth(15) = 3000: .Text = "其他開庭律師或配合開庭承辦人" 'Add By Sindy 2011/6/17
      '判斷是否有資料
      If .Rows > 1 Then
      .Visible = True
         For i = 1 To .Rows - 1
            .MergeRow(i) = True
         Next
         '將第一列反白
         .row = 1
      End If
      .Visible = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm072002 = Nothing
End Sub

'Add By Sindy 2016/6/28
Private Sub MSHFlexGrid1_Click()
Dim ii As Integer
   
   MSHFlexGrid1.Visible = False
   If MSHFlexGrid1.MouseRow <> 0 And MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, 14) <> "" Then
      '資料列清除反白
      MSHFlexGrid1.col = 0
      MSHFlexGrid1.row = MSHFlexGrid1.MouseRow
      If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, 0) = "V" Then
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, 0) = ""
         For ii = 0 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.col = ii
            MSHFlexGrid1.CellBackColor = QBColor(15)
         Next ii
      Else
         '資料列反白
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, 0) = "V"
         For ii = 0 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.col = ii
            MSHFlexGrid1.CellBackColor = &HFFC0C0
         Next ii
      End If
   End If
   MSHFlexGrid1.Visible = True
End Sub

'Add By Sindy 2016/6/28
Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow MSHFlexGrid1, x, y, nCol, nRow
   MSHFlexGrid1.col = nCol
   MSHFlexGrid1.row = nRow
End Sub
