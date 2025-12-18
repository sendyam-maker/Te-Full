VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm06010615 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "郵件接收狀況查詢"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8835
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   7890
      TabIndex        =   6
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   330
      Left            =   6990
      TabIndex        =   5
      Top             =   60
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm06010615.frx":0000
      Left            =   840
      List            =   "frm06010615.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   135
      Width           =   1695
   End
   Begin VB.TextBox txtMRL02 
      Height          =   270
      Left            =   3630
      MaxLength       =   7
      TabIndex        =   0
      Top             =   165
      Width           =   885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm06010615.frx":0004
      Height          =   4605
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   8123
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "信箱|接收日期|起始時間|截止時間|新增人員|接收筆數|加密筆數|個案筆數|執行狀況"
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
      _Band(0).Cols   =   9
   End
   Begin VB.Label Label1 
      Caption         =   "信箱："
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   210
      Width           =   555
   End
   Begin VB.Label Label3 
      Caption         =   "接收日期："
      Height          =   195
      Left            =   2670
      TabIndex        =   3
      Top             =   210
      Width           =   915
   End
End
Attribute VB_Name = "frm06010615"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/10 Form2.0已修改
'Create by Sindy 2017/7/6
Option Explicit

Dim dblPrevRow As Double
Public m_QueryType As String


Private Sub cmdExit_Click()
   Unload Me
End Sub

Public Sub cmdQuery_Click()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer
   
   strSql = ""
   If Combo1.Text <> "" Then
      strSql = strSql & " and MRL01='" & Left(Combo1.Text, 2) & "'"
'   Else
'      If m_QueryType = "F" Then '國外部
'         strSql = strSql & " and MRL01 in('01','02')"
'      End If
   End If
   If txtMRL02.Text <> "" Then
      strSql = strSql & " and MRL02='" & DBDATE(txtMRL02.Text) & "'"
   End If
   
   GRD1.Clear
   SetGrd
   
   Screen.MousePointer = vbHourglass
   strSql = "Select " & MRL01CName & " 信箱,sqldatet(MRL02) 接收日期,sqltime6(MRL03) 起始時間,sqltime6(MRL04) 截止時間,st02 新增人員,MRL06 接收筆數,MRL07 加密筆數,MRL08 個案筆數," & MRL09CName & " 執行狀況" & _
            " From MailReceiveLog,Staff" & _
            " Where MRL05=ST01(+)" & strSql & _
            " Order By MRL02||substr('000000'||MRL03,-6) desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
   Else
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   dblPrevRow = GRD1.row
   If rsTmp.RecordCount > 0 Then
      'GRD1.Text = "V"
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = &HFFC0C0
      Next i
   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   '                        0       1           2           3           4           5           6           7           8
   arrGridHeadText = Array("信箱", "接收日期", "起始時間", "截止時間", "新增人員", "接收筆數", "加密筆數", "個案筆數", "執行狀況")
   arrGridHeadWidth = Array(1400, 900, 800, 800, 800, 800, 800, 800, 800)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtMRL02 = strSrvDate(2)
   
   SetCombo1
End Sub

Private Sub SetCombo1()
   If Pub_StrUserSt03 = "M51" Then Combo1.AddItem ""
'   If m_QueryType = "F" Or Pub_StrUserSt03 = "M51" Then '國外部
      Combo1.AddItem "01.IPDept_inbound"
      Combo1.AddItem "02.IPDept_backup"
'   End If
'   If m_QueryType = "P" Or Pub_StrUserSt03 = "M51" Then '專利處
      Combo1.AddItem "03.Patent"
'   End If
'   If m_QueryType = "T" Or Pub_StrUserSt03 = "M51" Then '商標處
      Combo1.AddItem "04.TM"
'   End If
   Combo1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010615 = Nothing
End Sub
