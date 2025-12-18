VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_L_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "副檔名對照表"
   ClientHeight    =   5870
   ClientLeft      =   2800
   ClientTop       =   3720
   ClientWidth     =   7640
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5870
   ScaleWidth      =   7640
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Left            =   5190
      TabIndex        =   8
      Top             =   60
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6180
      TabIndex        =   1
      Top             =   60
      Width           =   1290
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4545
      Left            =   30
      TabIndex        =   2
      Top             =   1290
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   8026
      _Version        =   393216
      Cols            =   5
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
      _Band(0).Cols   =   5
   End
   Begin MSForms.TextBox txtQ 
      Height          =   300
      Left            =   1110
      TabIndex        =   0
      Top             =   450
      Width           =   2205
      VariousPropertyBits=   671105051
      Size            =   "3889;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "模糊比對（副檔名、中文說明）"
      Height          =   180
      Left            =   3360
      TabIndex        =   9
      Top             =   480
      Width           =   2520
   End
   Begin VB.Label Label4 
      Caption         =   "搜尋文字："
      Height          =   225
      Left            =   180
      TabIndex        =   7
      Top             =   480
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "屬來函案件性質時, 檔名若為: 本所案號.案件性質.PDF 示為 官方來函"
      ForeColor       =   &H00C00000&
      Height          =   200
      Left            =   180
      TabIndex        =   6
      Top             =   1050
      Width           =   6140
   End
   Begin VB.Label Label2 
      Caption         =   $"frm100101_L_1.frx":0000
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   1500
      TabIndex        =   5
      Top             =   30
      Width           =   3495
   End
   Begin VB.Label LblSysCode 
      Height          =   225
      Left            =   930
      TabIndex        =   4
      Top             =   750
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統別："
      Height          =   180
      Left            =   180
      TabIndex        =   3
      Top             =   780
      Width           =   720
   End
End
Attribute VB_Name = "frm100101_L_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/27 Form2.0已修改
'Create By Sindy 2014/11/27
Option Explicit

Public m_CP01 As String
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序


Private Sub SetDataListWidth()
Me.grdDataList.Cols = 5
Me.grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "系統別"
grdDataList.ColWidth(0) = 800
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "副檔名"
grdDataList.ColWidth(1) = 1400
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "台灣中文說明"
If m_CP01 = "CFP" And (Pub_StrUserSt03 <> "M51" Or strUserNum = "74001") Then
   grdDataList.ColWidth(2) = 0
Else
   grdDataList.ColWidth(2) = 2000
End If
grdDataList.CellAlignment = flexAlignCenterCenter
If m_CP01 = "CFP" And (Pub_StrUserSt03 <> "M51" Or strUserNum = "74001") Then
   grdDataList.col = 3: grdDataList.Text = "中文說明"
Else
   grdDataList.col = 3: grdDataList.Text = "非台灣中文說明"
End If
grdDataList.ColWidth(3) = 2000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "排序欄"
'If Pub_StrUserSt03 = "M51" And strUserNum <> "74001" Then
   grdDataList.ColWidth(4) = 800
'Else
'   grdDataList.ColWidth(4) = 0
'End If
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub cmdOK_Click()
   Unload Me
   frm100101_L.Show
End Sub

Private Sub cmdQuery_Click()
   Call QueryData
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Call QueryData
End Sub

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strConSql As String 'Add By Sindy 2020/8/28
   
   m_blnColOrderAsc = True
   QueryData = False
   
   If txtQ <> "" Then
      'Modify By Sindy 2023/5/5 + and efc06='Y'
      strConSql = strConSql & " and (efc02 like '%" & ChgSQL(UCase(txtQ)) & "%' or efc03 like '%" & ChgSQL(UCase(txtQ)) & "%' or efc04 like '%" & ChgSQL(UCase(txtQ)) & "%') and efc06='Y'"
   End If
   
   Screen.MousePointer = vbHourglass
   Me.grdDataList.Clear
   SetDataListWidth
   If Pub_StrUserSt03 = "M51" And strUserNum <> "74001" Then
      Label1.Visible = False '系統別
      strSql = "Select efc01 as 系統別,efc02 as 副檔名,efc03 as 台灣中文說明,efc04 as 非台灣中文說明,efc05 as 排序欄" & _
               " From efilecaption Where 1=1" & strConSql & _
               " order by efc01,efc05,efc02 asc"
   Else
      Label1.Visible = True '系統別
      Me.LblSysCode.Caption = m_CP01
      'Modify By Sindy 2018/5/23 取消 efc01 in('ALL','" & m_CP01 & "') and
      strSql = "Select efc01 as 系統別,efc02 as 副檔名,efc03 as 台灣中文說明,efc04 as " & IIf(m_CP01 = "CFP" And (Pub_StrUserSt03 <> "M51" Or strUserNum = "74001"), "中文說明", "非台灣中文說明") & ",efc05 as 排序欄" & _
               " From efilecaption" & _
               " Where efc06='Y'" & strConSql & _
               " order by efc01,efc02 asc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grdDataList.Recordset = rsTmp
      QueryData = True
   Else
      Set grdDataList.Recordset = rsTmp
      grdDataList.AddItem ""
      ShowNoData
   End If
   rsTmp.Close
   Screen.MousePointer = vbDefault
End Function

Private Sub Form_Unload(Cancel As Integer)
Set frm100101_L_1 = Nothing
End Sub

Private Sub GrdDataList_Click()
Dim nCol As Integer, nRow As Integer
   
   With grdDataList
      .Visible = False
      nCol = .MouseCol
      nRow = .MouseRow
      If nRow = 0 Then
         .col = nCol
         If m_blnColOrderAsc = False Then '字串降冪
            If nCol = 4 Then
               .Sort = 3 '數字昇冪
            Else
               .Sort = 5 '字串昇冪
            End If
            m_blnColOrderAsc = True
         Else
            If nCol = 4 Then
               .Sort = 4 '數字降冪
            Else
               .Sort = 6 '字串降冪
            End If
            m_blnColOrderAsc = False
         End If
      End If
      .Visible = True
   End With
End Sub
