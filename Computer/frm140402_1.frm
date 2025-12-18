VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm140402_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "名稱相同之聯絡人清單"
   ClientHeight    =   5430
   ClientLeft      =   195
   ClientTop       =   2520
   ClientWidth     =   8745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8745
   Begin VB.CommandButton cmdOK 
      Caption         =   "與被點選者為離職關係(&Y)"
      Height          =   400
      Index           =   0
      Left            =   4005
      TabIndex        =   0
      Top             =   60
      Width           =   2550
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "無關係(&N)"
      Height          =   400
      Index           =   1
      Left            =   6585
      TabIndex        =   1
      Top             =   60
      Width           =   1920
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   225
      Top             =   90
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4635
      Left            =   180
      TabIndex        =   2
      Top             =   600
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   8176
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "選擇|聯絡人編號|聯絡人名稱|客戶/代理人名稱|"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frm140402_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/7 改成Form2.0 (無)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'Create by Morgan 2007/12/4
Option Explicit

Dim m_iSelRow As Integer
Public fmParent As Form

Private Sub cmdok_Click(Index As Integer)
   Dim stRefNo2 As String
   Select Case Index
      Case 0
         If CheckCheck(stRefNo2) = False Then
            MsgBox "若與下列聯絡人有為離職關係時請點選該筆！"
            Exit Sub
         End If
      Case 1
         If CheckCheck = True Then
            MsgBox "若與全部聯絡人均非離職關係時請取消點選！"
            Exit Sub
         End If
   End Select
   fmParent.Tag = stRefNo2
   Unload Me
End Sub

Private Function CheckCheck(Optional p_No2 As String) As Boolean
   Dim ii As Integer
   With grdDataList
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) = "V" Then
            p_No2 = .TextMatrix(ii, 4)
            CheckCheck = True
            Exit For
         End If
      Next
   End With
End Function

Private Sub Form_Activate()
   SetDataListWidth
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140402_1 = Nothing
End Sub

Private Sub SetDataListWidth()
   With grdDataList
      .FormatString = .FormatString
      .ColWidth(0) = 465
      .ColAlignment(0) = flexAlignCenterCenter
      .ColWidth(1) = 1110
      .ColAlignment(1) = flexAlignLeftCenter
      .ColWidth(2) = 2685
      .ColAlignment(2) = flexAlignLeftCenter
      .ColWidth(3) = 3720
      .ColAlignment(3) = flexAlignLeftCenter
      .ColWidth(4) = 0
   End With
End Sub

Private Sub grdSelected(p_iRow As Integer)
   Dim stCheck As String, lColor As Long, ii As Integer
   With grdDataList
      .row = p_iRow
      .col = 0
      If .Text = "" Then
         .Text = "V"
         m_iSelRow = .row
         lColor = &HFFC0C0
      Else
         .Text = ""
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
   With grdDataList
      If .MouseRow > 0 And .MouseRow < .Rows Then
         .Visible = False
         iRow = .MouseRow
         If m_iSelRow = iRow Then
            grdSelected m_iSelRow
         Else
            If m_iSelRow > 0 Then
               grdSelected m_iSelRow
            End If
            grdSelected iRow
         End If
         .Visible = True
      End If
   End With
End Sub
