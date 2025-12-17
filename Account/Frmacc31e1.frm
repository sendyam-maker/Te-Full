VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc31e1 
   AutoRedraw      =   -1  'True
   Caption         =   "智慧局送件明細資料"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   8760
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3825
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   6747
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
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
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1575
      TabIndex        =   2
      Top             =   4170
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "規費合計："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   270
      TabIndex        =   1
      Top             =   4170
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   90
      Top             =   60
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc31e1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/10/20 Form2.0已修改 grdDataList
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8880
   Me.Height = 4875
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   '2012/8/9 modify by sonia
   'Frmacc0000.tool15_enabled
   tool17_enabled
   Frmacc31e0.Enabled = True
   Frmacc31e0.Show
   strFormName = "Frmacc31e0"
   Set Frmacc31e1 = Nothing
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
   SetDataListWidth
   SetGrid
End Sub

Private Sub SetDataListWidth()
   
   With grdDataList
      .Cols = 4: .Rows = 2
      .row = 0:
      .col = 0: .ColWidth(.col) = 1800: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      .col = 1: .ColWidth(.col) = 1200: .Text = "規費"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignRightCenter
      .col = 2: .ColWidth(.col) = 1600: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
      .col = 3: .ColWidth(.col) = 3400: .Text = "申請人"
      .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      .ColAlignment(.col) = flexAlignLeftCenter
   End With
   
End Sub

Private Sub SetGrid()

   Dim i As Integer, lngTotal As Long
   
On Error GoTo ErrHnd

   strSql = Frmacc31e0.GetSql
      
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      adoRecordset.MoveFirst
      Do While Not adoRecordset.EOF
         grdDataList.Rows = grdDataList.Rows + 1: grdDataList.row = grdDataList.Rows - 2
         For i = 0 To grdDataList.Cols - 1
            grdDataList.col = i
            If i = 1 Then
               grdDataList.Text = Format(adoRecordset.Fields(i), DDollar)
               lngTotal = lngTotal + Val("" & adoRecordset.Fields(i)) 'Add by Morgan 2011/9/2
            Else
               grdDataList.Text = "" & adoRecordset.Fields(i)
            End If
            
         Next i
         adoRecordset.MoveNext
      Loop
      If grdDataList.Rows > 2 Then
         grdDataList.Rows = grdDataList.Rows - 1
      End If
   End If
   lblTotal = Format(lngTotal, DDollar) 'Add by Morgan 2011/9/2
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

