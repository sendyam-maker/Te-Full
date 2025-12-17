VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc1132 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "拆收據其他收據號選單"
   ClientHeight    =   3555
   ClientLeft      =   195
   ClientTop       =   2520
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6975
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&Y)"
      Height          =   400
      Index           =   0
      Left            =   4860
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "取消(&N)"
      Height          =   400
      Index           =   1
      Left            =   5850
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   2835
      Left            =   90
      TabIndex        =   2
      Top             =   600
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5001
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "V|收據號碼|收據日期|收據抬頭|服務費|規費"
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
         Name            =   "新細明體-ExtB"
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
   Begin VB.Label lblAlert 
      BackStyle       =   0  '透明
      Caption         =   "本收據有拆收據情形，下列收據將會一併作廢，是否確定要繼續！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Visible         =   0   'False
      Width           =   4155
   End
End
Attribute VB_Name = "Frmacc1132"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/14 Form2.0已修改
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Create by Morgan 2011/9/21
Option Explicit

Dim m_iSelRow As Integer
Public fmParent As Form


Private Sub cmdok_Click(Index As Integer)
   Dim stRefNo2 As String
   Select Case Index
      Case 0
         '作廢確認
         If lblAlert.Visible = True Then
            stRefNo2 = "Y"
         Else
            If CheckCheck(stRefNo2) = False Then
               MsgBox "請點選一筆資料！"
               Exit Sub
            End If
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
   Set Frmacc1132 = Nothing
End Sub

Private Sub SetDataListWidth()
   With grdDataList
      .FormatString = .FormatString
      If Me.lblAlert.Visible = True Then
         .ColWidth(0) = 0
      Else
         .ColWidth(0) = 465
      End If
      .ColAlignment(0) = flexAlignCenterCenter
      'Modified by Lydia 2016/09/05  可顯示完整子號
      '.ColWidth(1) = 900
      .ColWidth(1) = 1100
      .ColAlignment(1) = flexAlignLeftCenter
      .ColWidth(2) = 900
      .ColAlignment(2) = flexAlignCenterCenter
      .ColWidth(3) = 2200
      .ColAlignment(3) = flexAlignLeftCenter
      .ColWidth(4) = 850
      .ColAlignment(4) = flexAlignRightCenter
      .ColWidth(5) = 850
      .ColAlignment(5) = flexAlignRightCenter
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
      cmdok_Click 0
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
