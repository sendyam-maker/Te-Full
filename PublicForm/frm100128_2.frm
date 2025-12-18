VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100128_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "FC客戶直接來所申請比率統計"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   6480
   Begin VB.CommandButton cmdok 
      Caption         =   "Word(&W)"
      Enabled         =   0   'False
      Height          =   405
      Index           =   2
      Left            =   3330
      TabIndex        =   3
      Top             =   60
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4455
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   6345
      _ExtentX        =   11201
      _ExtentY        =   7849
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FormatString    =   "排名|國家|客戶直接來所件數/總件數|比率(%)"
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
      _Band(0).Cols   =   4
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   4275
      TabIndex        =   0
      Top             =   60
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   5445
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.Label lblCondition 
      AutoSize        =   -1  'True
      Caption         =   "lblCondition"
      Height          =   180
      Index           =   3
      Left            =   1035
      TabIndex        =   12
      Top             =   810
      Width           =   885
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Index           =   3
      Left            =   135
      TabIndex        =   11
      Top             =   810
      Width           =   900
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "lblMemo"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   45
      TabIndex        =   10
      Top             =   5520
      Width           =   780
   End
   Begin VB.Label lblCondition 
      AutoSize        =   -1  'True
      Caption         =   "lblCondition"
      Height          =   180
      Index           =   2
      Left            =   1035
      TabIndex        =   9
      Top             =   90
      Width           =   885
   End
   Begin VB.Label lblCondition 
      AutoSize        =   -1  'True
      Caption         =   "lblCondition"
      Height          =   180
      Index           =   1
      Left            =   1035
      TabIndex        =   8
      Top             =   570
      Width           =   885
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "統計部門："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   7
      Top             =   90
      Width           =   900
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   6
      Top             =   570
      Width           =   540
   End
   Begin VB.Label lblCondition 
      AutoSize        =   -1  'True
      Caption         =   "lblCondition"
      Height          =   180
      Index           =   0
      Left            =   1035
      TabIndex        =   5
      Top             =   330
      Width           =   885
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "查詢別："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   330
      Width           =   720
   End
End
Attribute VB_Name = "frm100128_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/22 改成Form2.0(grdDataList改Fonts)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Create by Morgan 2010/8/25
Option Explicit

Public cmdState As Integer

Private Sub cmdOK_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100128_2 = Nothing
End Sub

Sub PubShowNextData()
   Select Case cmdState
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
      Case 2
         Screen.MousePointer = vbHourglass
         runWord
         Screen.MousePointer = vbDefault
      Case Else
   End Select
End Sub

Public Sub SetGrid(p_Rst As ADODB.Recordset, Optional p_Type As Integer = 1)
   Dim iRow As Integer, lngTot1 As Long, lngTot2 As Long
   With grdDataList
      .Visible = False
      Set .Recordset = p_Rst.Clone
      .FormatString = .FormatString
      If p_Type = 2 Then
         .TextMatrix(0, 1) = "洲"
      End If
      .ColWidth(1) = 2000
      .ColAlignment(2) = flexAlignCenterCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColWidth(4) = 0
      .ColWidth(5) = 0
      For iRow = 1 To .Rows - 1
         .TextMatrix(iRow, 0) = iRow
         lngTot1 = lngTot1 + Val(.TextMatrix(iRow, 4))
         lngTot2 = lngTot2 + Val(.TextMatrix(iRow, 5))
      Next
      .AddItem "總計" & vbTab & vbTab & lngTot1 & "/" & lngTot2 & vbTab & Format(100 * lngTot1 / lngTot2, "0") & "%"
      .Visible = True
   End With
End Sub

Private Sub runWord()
   
   Dim stTmp As String
   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   
On Error GoTo ErrHnd
   
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   
   g_WordAp.Visible = True
   g_WordAp.Documents.add
   With g_WordAp.Application
      .WindowState = wdWindowStateMaximize
      .Selection.Font.Name = "標楷體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 14
      '邊界
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1.6)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.4)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      
      stTmp = Me.Caption
      
      .Selection.Font.Size = 18
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      .Selection.TypeText Text:=stTmp
      
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.Font.Size = 14
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.TypeText Text:=lblName(2) & lblCondition(2)
      .Selection.TypeParagraph
      .Selection.TypeText Text:=lblName(0) & lblCondition(0)
      .Selection.TypeParagraph
      .Selection.TypeText Text:=lblName(1) & lblCondition(1)
      .Selection.TypeParagraph
      .Selection.TypeText Text:=lblName(3) & lblCondition(3)
      .Selection.TypeParagraph
      .Selection.ParagraphFormat.Alignment = 2
      .Selection.TypeText Text:=lblMemo
      .Selection.TypeParagraph
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      If grdDataList.Cols - 2 > 7 Then
         .Selection.Font.Size = 12
      End If
      '列數,欄數
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=grdDataList.Rows, NumColumns:=4
      '設定表格高度
      .Selection.Cells.SetHeight RowHeight:=26, HeightRule:=wdRowHeightExactly
      For iRow = 0 To grdDataList.Rows - 1
         .Selection.SelectRow
         If iRow = 0 Then
            .Selection.Cells.SetHeight RowHeight:=52, HeightRule:=wdRowHeightExactly
         End If
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
         .Selection.MoveLeft Unit:=wdCharacter, Count:=1
         For iCol = 0 To 3
            .Selection.TypeText Text:=grdDataList.TextMatrix(iRow, iCol)
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
         Next
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
      Next
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Roman"
      .Selection.MoveRight Unit:=wdCharacter, Count:=1
      .Activate
   End With
   
ErrHnd:

   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91:
            g_WordAp.Documents.add
            Resume Next
         Case 462:
            Set g_WordAp = New Word.Application
            g_WordAp.Documents.add
            Resume Next
         Case Else:
            MsgBox "錯誤 : " & Err.Description, vbCritical
      End Select
   End If
End Sub

