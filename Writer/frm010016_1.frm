VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010016_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外信件清單數量"
   ClientHeight    =   3675
   ClientLeft      =   5580
   ClientTop       =   1860
   ClientWidth     =   5205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5205
   Begin VB.TextBox textLI01 
      Alignment       =   2  '置中對齊
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   2
      Top             =   240
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3810
      TabIndex        =   0
      Top             =   90
      Width           =   1275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2925
      Left            =   210
      TabIndex        =   1
      Top             =   630
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   5159
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   3
      FixedCols       =   0
      HighLight       =   0
      AllowUserResizing=   1
      FormatString    =   "各部門             | 列印次數      | 頁數"
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "信件日期："
      Height          =   180
      Left            =   270
      TabIndex        =   3
      Top             =   300
      Width           =   900
   End
End
Attribute VB_Name = "frm010016_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/11 Form2.0已修改 (無需修改)
'Create By Sindy 2013/3/14
Option Explicit

Private Sub cmdOK_Click(Index As Integer)
   Me.Hide
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Public Function StrMenu() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim intSum As Integer, i As Integer
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   StrMenu = False: intSum = 0
   
   Call GridHead1
   strSql = "SELECT decode(FLC02,'P','專利處','FCP','外專','T','商標處','FCT','外商','FCL','投資法務','A','財務處','X','未分類',FLC02) as 各部門,Count(FLC03) as 列印次數,Count(FLC04) as 總張數 FROM ForeignLetterCount where FLC01=" & DBDATE(textLI01) & " group by FLC02 order by FLC02 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      StrMenu = True
      Set MSHFlexGrid1.Recordset = rsTmp
      rsTmp.MoveFirst
      For i = 1 To rsTmp.RecordCount
         intSum = intSum + Val(rsTmp.Fields(2))
         rsTmp.MoveNext
      Next i
      MSHFlexGrid1.AddItem ""
      MSHFlexGrid1.TextMatrix(rsTmp.RecordCount + 1, 1) = "總計"
      MSHFlexGrid1.TextMatrix(rsTmp.RecordCount + 1, 2) = intSum
   End If
   rsTmp.Close
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm010016_1 = Nothing
End Sub

Private Sub GridHead1()
   MSHFlexGrid1.Clear
   MSHFlexGrid1.Rows = 2
   With MSHFlexGrid1
      .Visible = False
      .Cols = 3
      .row = 0
      .col = 0: .ColWidth(0) = 2000: .Text = "各部門"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(0) = flexAlignCenterCenter
      
      .col = 1: .ColWidth(1) = 1000: .Text = "列印次數"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignCenterCenter
      
      .col = 2: .ColWidth(2) = 1000: .Text = "頁數"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(2) = flexAlignCenterCenter
      .Visible = True
   End With
End Sub
