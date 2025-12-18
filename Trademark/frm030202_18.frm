VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm030202_18 
   BorderStyle     =   1  '單線固定
   Caption         =   "複製變更事項"
   ClientHeight    =   3990
   ClientLeft      =   450
   ClientTop       =   990
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5310
   Begin VB.TextBox txtCaseNo 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   570
      Width           =   2400
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3075
      TabIndex        =   0
      Top             =   45
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   4050
      TabIndex        =   1
      Top             =   45
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   3030
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   5345
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblFund 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   570
      Width           =   975
   End
End
Attribute VB_Name = "frm030202_18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/2/21 Form2.0已修改
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Public strCP01 As String      '系統別(傳入)
Public strCP02 As String      '流水號(傳入)
Public strCP03 As String      '追加案號(傳入)
Public strCP04 As String      '多國多類碼(傳入)
Public strCP09 As String      '總收文號(傳入)
Public strCE01 As String      '總收文號(回傳)
Public bolOK As Boolean     'True: 確定  False: 取消


Private Sub SetDataListWidth()
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 500
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "總收文號"
grdDataList.ColWidth(1) = 1500
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean, j As Integer
   
   '確定
   If Index = 0 Then
      Cancel = True
      For j = 1 To grdDataList.Rows - 1
         If Trim(grdDataList.TextMatrix(j, 0)) = "V" And Trim(grdDataList.TextMatrix(j, 1)) <> "" Then
            strCE01 = Trim(grdDataList.TextMatrix(j, 1))
            Cancel = False
         End If
      Next j
      If Cancel = True Then
         MsgBox "至少點選一筆資料！", vbExclamation + vbOKOnly, Me.Caption
         Exit Sub
      End If
      bolOK = True
      
   '回前畫面(取消)
   Else
      strCE01 = ""
      bolOK = False
   End If
   Me.Hide
End Sub

Public Function CheckShowList() As Boolean
Dim strSql As String
Dim dblCP27 As Double
Dim intIdx As Integer, i As Integer
   
   CheckShowList = False
   
   strSql = "SELECT ' ' AS V,CE01 FROM CaseProgress,ChangeEvent " & _
                  "WHERE CP01='" & strCP01 & "' AND CP02='" & strCP02 & "' AND CP03='" & strCP03 & "' AND CP04='" & strCP04 & "' " & _
                  "AND CP09=CE01 " & _
                  "AND CE01<>'" & strCP09 & "' "
   Screen.MousePointer = vbHourglass
   grdDataList.Clear
   grdDataList.Rows = 2
   SetDataListWidth
   'GrdDataList.FixedCols = 0
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
       Set grdDataList.Recordset = adoRecordset
       CheckShowList = True
   Else
       'ShowNoData
       grdDataList.Clear
   End If
   SetDataListWidth
   'GrdDataList.FixedCols = 3
   CheckOC
   
'   '若只有一筆資料, 則直接設定為點選此筆資料
'   With Me.GrdDataList
'      If .Rows = 2 Then
'         .row = 1
'         .col = 1
'         If .Text <> "" Then
'           .Visible = False
'           .row = 1
'           .col = 0
'           .Text = "V"
'           For i = 0 To .Cols - 1
'               .col = i
'               .CellBackColor = &HFFC0C0
'               If i <= 2 Then
'                 GrdDataList.CellBackColor = &H8000000F
'               End If
'           Next i
'           .Visible = True
'         End If
'      End If
'   End With
   Screen.MousePointer = vbDefault
   
   bolOK = True
   Exit Function
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grdDataList, x, y, nCol, nRow
grdDataList.col = nCol
grdDataList.row = nRow
End Sub

Private Sub GrdDataList_Click()
Dim tmpMouseRow
Dim j As Integer, i As Integer
   
   grdDataList.Visible = False
   tmpMouseRow = grdDataList.row
   For j = 1 To grdDataList.Rows - 1
      grdDataList.row = j
      grdDataList.col = 0
      grdDataList.Text = ""
      For i = 0 To grdDataList.Cols - 1
           grdDataList.col = i
           grdDataList.CellBackColor = QBColor(15)
'           If i <= 2 Then
'              grdDataList.CellBackColor = &H8000000F
'           End If
      Next i
   Next j
   
   '勾選
   grdDataList.row = tmpMouseRow 'GrdDataList.MouseRow
   grdDataList.col = 0
   If grdDataList.row <> 0 Then
      If grdDataList.Text = "V" Then
           grdDataList.Text = ""
           For i = 0 To grdDataList.Cols - 1
               grdDataList.col = i
               grdDataList.CellBackColor = QBColor(15)
'               If i <= 2 Then
'                  grdDataList.CellBackColor = &H8000000F
'               End If
          Next i
      Else
           grdDataList.Text = "V"
           For i = 0 To grdDataList.Cols - 1
               grdDataList.col = i
               grdDataList.CellBackColor = &HFFC0C0
'               If i <= 2 Then
'                  grdDataList.CellBackColor = &H8000000F
'               End If
           Next i
      End If
   End If
   
   grdDataList.Visible = True
End Sub
