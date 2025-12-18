VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010003 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   5730
   ClientLeft      =   -30
   ClientTop       =   900
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9345
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   5052
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   8916
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
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
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   8052
      TabIndex        =   1
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7224
      TabIndex        =   0
      Top             =   70
      Width           =   800
   End
End
Attribute VB_Name = "frm010003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/16 Form2.0已修改 grdDataList
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean


Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer

If Index = 0 Then
   '將選定的Row之資料填回frm010002
   frm010002.txtSystem = grdDataList.TextMatrix(grdDataList.row, 1)
   If grdDataList.TextMatrix(grdDataList.row, 1) = 馬德里案 Then
      frm010002.fraTF.Visible = True
      frm010002.fraElse.Visible = False
      For i = 0 To 3
         Select Case i
            Case 1
               frm010002.txtTFCode(i) = IIf(grdDataList.TextMatrix(grdDataList.row, i + 2) = " ", "0", grdDataList.TextMatrix(grdDataList.row, i + 2))
            Case 2
               frm010002.txtTFCode(i) = IIf(grdDataList.TextMatrix(grdDataList.row, i + 2) = " ", "0", grdDataList.TextMatrix(grdDataList.row, i + 2))
            Case 3
               frm010002.txtTFCode(i) = IIf(grdDataList.TextMatrix(grdDataList.row, i + 2) = " ", "00", grdDataList.TextMatrix(grdDataList.row, i + 2))
         End Select
      Next
   Else
      frm010002.fraTF.Visible = False
      frm010002.fraElse.Visible = True
      frm010002.txtCode(0) = grdDataList.TextMatrix(grdDataList.row, 2)
      For i = 1 To 2
         Select Case i
            Case 1
               frm010002.txtCode(i) = IIf(grdDataList.TextMatrix(grdDataList.row, i + 3) = " ", "0", grdDataList.TextMatrix(grdDataList.row, i + 3))
            Case 2
               frm010002.txtCode(i) = IIf(grdDataList.TextMatrix(grdDataList.row, i + 3) = " ", "00", grdDataList.TextMatrix(grdDataList.row, i + 3))
         End Select
      Next
   End If
   frm010002.lblPetition.Caption = grdDataList.TextMatrix(grdDataList.row, 6)
   frm010002.CheckCaseCode
End If
Unload Me
End Sub
Private Sub Form_Activate()
If grdDataList.Rows = 1 Then
   ShowMsg MsgText(1030)
   Unload Me
End If
End Sub
Private Sub Form_Load()
Dim strTitle As String, varSaveCursor

MoveFormToCenter Me
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'Modify By Cheng 2003/05/14
'Set grdDataList.Recordset = obj001.ReadCKindRst(frm010002.txtCKind(1), frm010002.txtCKind(2))
'edit by nickc 2007/02/06 不用 dll 了
'Set grdDataList.Recordset = obj001.ReadCKindRst_1(frm010002.txtCKind(1), frm010002.txtCKind(2))
Set grdDataList.Recordset = Cls001ReadCKindRst_1(frm010002.txtCKind(1), frm010002.txtCKind(2))
SetDataListWidth
intLastRow = 0
If grdDataList.Rows > 1 Then
   ShowBar grdDataList, intLastRow, 5
Else
   cmdOK(0).Enabled = False
End If
Select Case frm010002.txtCKind(1)
             Case 專利
                        strTitle = "專利"
             Case 商標
                        strTitle = "商標"
             Case 法務
                        strTitle = "法務"
             Case 顧問
                        strTitle = "顧問"
             Case Else
                        strTitle = "服務業務"
End Select
Me.Caption = "尋找" + strTitle + "資料"
Screen.MousePointer = varSaveCursor
End Sub
Private Sub SetDataListWidth()
Dim varGridWidth() As Variant

varGridWidth = Array(4000, 400, 650, 0, 200, 250, 2000, 2000)
SetGridDataListWidth grdDataList, varGridWidth()
SetDataListVision grdDataList, , True
blnOKtoShow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm010003 = Nothing
End Sub

Private Sub grdDataList_DblClick()
cmdOK_Click (0)
End Sub
Private Sub grdDataList_GotFocus()
GridGotFocus grdDataList
End Sub
Private Sub grdDataList_LostFocus()
GridLostFocus grdDataList
End Sub
Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      ShowBar grdDataList, intLastRow, 5
      blnOKtoShow = True
   End If
End If
End Sub
