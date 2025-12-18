VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090215_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "期刊資料查訊列印"
   ClientHeight    =   6105
   ClientLeft      =   -1905
   ClientTop       =   1410
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9315
   Begin VB.CommandButton cmd 
      Caption         =   "資料維護(&E)"
      Height          =   400
      Left            =   6900
      TabIndex        =   3
      Top             =   30
      Width           =   1200
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   6600
      Top             =   6720
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
   Begin VB.CommandButton Command 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   8100
      TabIndex        =   0
      Top             =   30
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm090215_2.frx":0000
      Height          =   5235
      Left            =   60
      TabIndex        =   1
      Top             =   465
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9234
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      SelectionMode   =   1
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
      _Band(0).Cols   =   7
   End
   Begin VB.Label lbl 
      Caption         =   "期刊資料：　　筆"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   5790
      Width           =   3585
   End
End
Attribute VB_Name = "frm090215_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset
'Add By Cheng 2002/03/04
Public m_strPE01 As String '標題
Public m_strPE03 As String '資料出處
Public m_strPE04 As String '版,頁
Public m_strPE05 As String '出版日期

Private Sub cmd_Click()
If Len(m_strPE01) > 0 Then
   frm090213.Show
End If
End Sub

Private Sub Command_Click()
frm090215_1.Show
Unload Me
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    Set GRD1.Recordset = Adodc1.Recordset
    SetGrd
'Add By Cheng 2002/03/04
m_strPE01 = "": m_strPE03 = "": m_strPE04 = "": m_strPE05 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/03/04
m_strPE01 = "": m_strPE03 = "": m_strPE04 = "": m_strPE05 = ""
    
    Set frm090215_2 = Nothing
End Sub

Sub SetGrd()
With GRD1
   .Cols = 7
   .row = 0
    'Modify By Cheng 2003/11/27
'   .col = 0: .ColWidth(0) = 10000
'    .CellAlignment = flexAlignCenterCenter
'   .Text = "標題"
'   .col = 1: .ColWidth(1) = 1100
'    .CellAlignment = flexAlignCenterCenter
'   .Text = "出版日期"
'   .col = 2: .ColWidth(2) = 1000
'    .CellAlignment = flexAlignCenterCenter
'   .Text = "索引"
'   .col = 3: .ColWidth(3) = 1200
'    .CellAlignment = flexAlignCenterCenter
'   .Text = "資料出處"
'   .col = 4: .ColWidth(4) = 1000
'    .CellAlignment = flexAlignCenterCenter
'   .Text = "作者"
'   .col = 5: .ColWidth(5) = 4000
'    .CellAlignment = flexAlignCenterCenter
'   .Text = "備註"
'   .col = 6: .ColWidth(6) = 700
'    .CellAlignment = flexAlignCenterCenter
'   .Text = "版，頁"
   .col = 0: .ColWidth(0) = 1200
    .CellAlignment = flexAlignCenterCenter
   .Text = "資料出處"
   .col = 1: .ColWidth(1) = 1100
    .CellAlignment = flexAlignCenterCenter
   .Text = "出版日期"
   .col = 2: .ColWidth(2) = 700
    .CellAlignment = flexAlignCenterCenter
   .Text = "版，頁"
   .col = 3: .ColWidth(3) = 1000
    .CellAlignment = flexAlignCenterCenter
   .Text = "索引"
   .col = 4: .ColWidth(4) = 10000
    .CellAlignment = flexAlignCenterCenter
   .Text = "標題"
   .col = 5: .ColWidth(5) = 1000
    .CellAlignment = flexAlignCenterCenter
   .Text = "作者"
   .col = 6: .ColWidth(6) = 4000
    .CellAlignment = flexAlignCenterCenter
   .Text = "備註"
    'End
End With
End Sub

Private Sub Grd1_Click()
'Add By Cheng 2002/03/04
If Me.GRD1.RowSel > 0 Then
    'Modify By Cheng 2003/11/27
'   m_strPE01 = "" & Me.grd1.TextMatrix(Me.grd1.RowSel, 0)
'   m_strPE03 = "" & Me.grd1.TextMatrix(Me.grd1.RowSel, 3)
'   m_strPE04 = "" & Me.grd1.TextMatrix(Me.grd1.RowSel, 6)
'   m_strPE05 = "" & Replace(Me.grd1.TextMatrix(Me.grd1.RowSel, 1), "/", "")
   m_strPE01 = "" & Me.GRD1.TextMatrix(Me.GRD1.RowSel, 4)
   m_strPE03 = "" & Me.GRD1.TextMatrix(Me.GRD1.RowSel, 0)
   m_strPE04 = "" & Me.GRD1.TextMatrix(Me.GRD1.RowSel, 2)
   m_strPE05 = "" & Replace(Me.GRD1.TextMatrix(Me.GRD1.RowSel, 1), "/", "")
    'End
End If
Me.cmd.Default = True
End Sub

Private Sub grd1_RowColChange()
'Add By Cheng 2002/03/04
If Me.GRD1.RowSel > 0 Then
   m_strPE01 = "" & Me.GRD1.TextMatrix(Me.GRD1.RowSel, 0)
   m_strPE03 = "" & Me.GRD1.TextMatrix(Me.GRD1.RowSel, 3)
   m_strPE04 = "" & Me.GRD1.TextMatrix(Me.GRD1.RowSel, 6)
   m_strPE05 = "" & Replace(Me.GRD1.TextMatrix(Me.GRD1.RowSel, 1), "/", "")
End If
Me.cmd.Default = True
End Sub
