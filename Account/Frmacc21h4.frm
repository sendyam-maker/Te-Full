VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc21h4 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收文號選單"
   ClientHeight    =   3555
   ClientLeft      =   195
   ClientTop       =   2520
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4650
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&Y)"
      Height          =   400
      Index           =   0
      Left            =   2610
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "取消(&N)"
      Height          =   400
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   2835
      Left            =   90
      TabIndex        =   2
      Top             =   600
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   5001
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "選擇|收文號|收文日|案件性質|承辦人"
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "Frmacc21h4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/08 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'Create by Morgan 2010/4/7
Option Explicit
Dim m_iSelRow As Integer
Public fmParent As Form

Private Sub cmdOK_Click(Index As Integer)
   Dim stRefNo2 As String
   Select Case Index
      Case 0
         If CheckCheck(stRefNo2) = False Then
            MsgBox "請點選一筆資料！"
            Exit Sub
         End If
      Case 1
         stRefNo2 = ""
   End Select
   fmParent.Tag = stRefNo2
   Unload Me
End Sub

Private Function CheckCheck(Optional p_No2 As String) As Boolean
   Dim ii As Integer
   With grdDataList
      For ii = 1 To .Rows - 1
         If .TextMatrix(ii, 0) = "V" Then
            p_No2 = .TextMatrix(ii, 1)
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
   PUB_InitForm Me, Me.Width, Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Frmacc21h4 = Nothing
End Sub

Private Sub SetDataListWidth()
   Dim ii As Integer
   With grdDataList
      .FormatString = .FormatString
      .ColWidth(0) = 465
      .ColAlignment(0) = flexAlignCenterCenter
      .ColWidth(1) = 1000
      .ColAlignment(1) = flexAlignLeftCenter
      .ColWidth(2) = 1000
      .ColAlignment(2) = flexAlignLeftCenter
      .ColWidth(3) = 1000
      .ColAlignment(3) = flexAlignLeftCenter
      .ColWidth(4) = 800
      .ColAlignment(4) = flexAlignLeftCenter
      'Added by Morgan 2022/5/2
      For ii = 5 To .Cols - 1
         .ColWidth(ii) = 0
      Next
      'end 2022/5/2
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
   GrdClick
End Sub

Private Sub grdDataList_DblClick()
   GrdClick
   If CheckCheck Then
      cmdOK_Click 0
   End If
End Sub

Private Sub GrdClick()
   Dim iRow As Integer
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
