VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm082006 
   BorderStyle     =   1  '單線固定
   Caption         =   "開拓客戶資料查詢"
   ClientHeight    =   5745
   ClientLeft      =   105
   ClientTop       =   675
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9345
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5076
      Left            =   132
      TabIndex        =   0
      Top             =   564
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   8943
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
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
      _Band(0).Cols   =   1
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8424
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7296
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   70
      Width           =   1100
   End
End
Attribute VB_Name = "frm082006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Private Sub GridHead()
 Dim i As Integer
   With MSHFlexGrid1
      .row = 0
      .col = 0: .ColWidth(0) = 1000: .Text = "屬性名稱"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 500: .Text = "編號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1700: .Text = "收件人名稱"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1000: .Text = "國籍"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 3000: .Text = "地址"
      .CellAlignment = flexAlignCenterCenter
   End With
End Sub

Private Sub cmdBack_Click()
   frm082005.Show
   Unload frm082006
End Sub

Private Sub cmdEnd_Click()
   Unload frm082006
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm082006 = Nothing
End Sub
