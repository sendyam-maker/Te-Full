VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc11c0_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "整張收款單號改扣繳年度"
   ClientHeight    =   5430
   ClientLeft      =   200
   ClientTop       =   2520
   ClientWidth     =   8740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8740
   Begin VB.CommandButton cmdOK 
      Caption         =   "全取消(&C)"
      Height          =   400
      Index           =   1
      Left            =   5580
      TabIndex        =   7
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "更新存檔"
      Height          =   400
      Index           =   3
      Left            =   2460
      TabIndex        =   4
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox TextYear 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   3
      Top             =   150
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "全選(&A)"
      Height          =   400
      Index           =   0
      Left            =   4560
      TabIndex        =   0
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   6600
      TabIndex        =   1
      Top             =   90
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4395
      Left            =   90
      TabIndex        =   2
      Top             =   930
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   7743
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   10
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
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收款日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   570
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   210
      Width           =   1095
   End
End
Attribute VB_Name = "Frmacc11c0_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/10/8 Form2.0已修改
Option Explicit

Dim m_iSelRow As Integer


Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer, lColor As Long, ii As Integer
   
On Error GoTo ErrHand
   
   Select Case Index
      Case 0 '全選
         With grdDataList
            For i = 1 To .Rows - 1
               .row = i
               .col = 0
               .Text = "V"
               lColor = &HFFC0C0
               For ii = 0 To .Cols - 1
                  .col = ii
                  .CellBackColor = lColor
               Next ii
            Next i
         End With
      Case 1 '全取消
         With grdDataList
            For i = 1 To .Rows - 1
               .row = i
               .col = 0
               .Text = ""
               lColor = &H80000018
               For ii = 0 To .Cols - 1
                  .col = ii
                  .CellBackColor = lColor
               Next ii
            Next i
         End With
      Case 2 '結束
         Unload Me
      Case 3 '更新存檔
         If TextYear = "" Then
            MsgBox "請輸入扣繳年度！", , MsgText(5)
            TextYear.SetFocus
            Exit Sub
         ElseIf IsDate(Format(TransDate(CStr(TextYear) & "0101", 2), "####/##/##")) = False Then
            MsgBox "請輸入正確年度！", , MsgText(5)
            TextYear.SetFocus
            Exit Sub
         End If
         With grdDataList
            ii = 0 'add by sonia 2024/1/30
            For i = 1 To .Rows - 1
               .row = i
               .col = 0
               If .Text = "V" Then
                  cnnConnection.BeginTrans
                  strSql = "update acc0k0 set a0k16=" & TextYear & " where a0k01='" & .TextMatrix(i, 3) & "'"
                  cnnConnection.Execute strSql
                  'add by sonia 2023/3/13 改扣繳年度同時更新A0K15
                  If TextYear <> grdDataList.TextMatrix(i, 5) Then
                     strSql = "update acc0k0 set a0k15='" & strSrvDate(2) & "' where a0k01='" & .TextMatrix(i, 3) & "'"
                     cnnConnection.Execute strSql
                  End If
                  'end 2023/3/13
                  strSql = "update acc0m0 set a0m07=" & TextYear & " where a0m02='" & .TextMatrix(i, 3) & "'"
                  cnnConnection.Execute strSql
                  strSql = "update acc1v0 set a1v09=" & TextYear & " where a1v02='" & .TextMatrix(i, 3) & "'"
                  cnnConnection.Execute strSql
                  ii = ii + 1 'add by sonia 2024/1/30
                  cnnConnection.CommitTrans
                  'Frmacc11c0.bolUpdData = True  'cancel by sonia 2024/1/30 下方檢查是否有選取要更新的資料
               End If
            Next i
            'add by sonia 2024/1/30
            If ii > 0 Then
               Frmacc11c0.bolUpdData = True
            Else
               MsgBox "未選取任何資料！", , MsgText(5)
               Exit Sub
            End If
            'end 2024/1/30
         End With
         Unload Me
   End Select
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number > 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
End Sub

Private Sub Form_Activate()
   SetDataListWidth
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Frmacc11c0_1 = Nothing
End Sub

Private Sub SetDataListWidth()
   With grdDataList
      '                0    1    2        3        4        5        6        7
      .FormatString = "V|公司|收據日期|收據編號|收據抬頭|扣繳年度|已扣繳額|未扣繳額"
      .ColWidth(0) = 400
      .ColAlignment(0) = flexAlignCenterCenter
      .ColWidth(1) = 900
      .ColAlignment(1) = flexAlignLeftCenter
      .ColWidth(2) = 850
      .ColAlignment(2) = flexAlignLeftCenter
      .ColWidth(3) = 1000
      .ColAlignment(3) = flexAlignLeftCenter
      .ColWidth(4) = 1800
      .ColAlignment(4) = flexAlignLeftCenter
      .ColWidth(5) = 900
      .ColAlignment(5) = flexAlignLeftCenter
      .ColWidth(6) = 900
      .ColAlignment(6) = flexAlignLeftCenter
      .ColWidth(7) = 900
      .ColAlignment(7) = flexAlignLeftCenter
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
         
'         If m_iSelRow > 0 Then
'            grdSelected m_iSelRow
'         End If
        
         'If m_iSelRow <> iRow Then
            grdSelected iRow
         'End If
         .Visible = True
      End If
   End With
End Sub
