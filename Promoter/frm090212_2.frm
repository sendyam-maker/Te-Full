VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090212_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "公報簡訊查詢列印"
   ClientHeight    =   6105
   ClientLeft      =   -1770
   ClientTop       =   1140
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   9315
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Bindings        =   "frm090212_2.frx":0000
      Height          =   5232
      Left            =   72
      TabIndex        =   1
      Top             =   468
      Width           =   9216
      _ExtentX        =   16245
      _ExtentY        =   9234
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6225
      Top             =   6840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   8208
      TabIndex        =   0
      Top             =   36
      Width           =   1092
   End
   Begin VB.Label lbl 
      Caption         =   "公報簡訊：　　筆"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   90
      TabIndex        =   2
      Top             =   5760
      Width           =   3585
   End
End
Attribute VB_Name = "frm090212_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Private Sub Command1_Click()
frm090212_1.Show
Unload Me
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
GridHead
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090212_2 = Nothing
End Sub

Private Sub GridHead()
   With grd1
      .row = 0
      .col = 0:       .Text = "頁數"
      .ColWidth(0) = 600
      .CellAlignment = flexAlignCenterCenter
      .col = 1:       .Text = "公告號數"
      .ColWidth(1) = 1000
      .CellAlignment = flexAlignCenterCenter
      .col = 2:       .Text = "公告日期"
      .ColWidth(2) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 3:       .Text = "內容"
      .ColWidth(3) = 8500
      .CellAlignment = flexAlignCenterCenter
      .col = 4:       .Text = "國際分類"
      .ColWidth(4) = 2500
      .CellAlignment = flexAlignCenterCenter
      .col = 5:       .Text = "索引"
      .ColWidth(5) = 3000
      .CellAlignment = flexAlignCenterCenter
   End With
End Sub

