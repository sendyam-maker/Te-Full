VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090801_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "舊案申請地址與客戶目前申請地址不同者"
   ClientHeight    =   5865
   ClientLeft      =   2790
   ClientTop       =   3720
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Left            =   7350
      TabIndex        =   0
      Top             =   0
      Width           =   930
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4740
      Left            =   75
      TabIndex        =   1
      Top             =   1050
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   8361
      _Version        =   393216
      Cols            =   16
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
      _Band(0).Cols   =   16
   End
   Begin VB.Label Label3 
      Caption         =   "申請地址：　(勾選一項地址，帶入前一畫面申請地址欄位)"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   825
      Width           =   7920
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   900
      TabIndex        =   4
      Top             =   150
      Width           =   1545
   End
   Begin MSForms.Label lblName 
      Height          =   255
      Left            =   900
      TabIndex        =   3
      Top             =   450
      Width           =   7995
      VariousPropertyBits=   27
      Size            =   "14102;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人1："
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   150
      Width           =   810
   End
End
Attribute VB_Name = "frm090801_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/22 改成Form2.0 (grdDataList,lblName)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim m_blnOneRec As Boolean
Dim i As Integer, ii As Integer
Dim intRow As Integer
Dim intCol As Integer
Public UpForm As Form
Public strAddr As String


Private Sub SetDataListWidth()
Me.grdDataList.Cols = 3
Me.grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "V"
grdDataList.ColWidth(0) = 200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "案件申請地址"
grdDataList.ColWidth(1) = 6000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "備註"
grdDataList.ColWidth(2) = 2500
grdDataList.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub cmdOK_Click()
   For ii = 1 To Me.grdDataList.Rows - 1
      If Me.grdDataList.TextMatrix(ii, 0) = "V" Then
         m_blnOneRec = True
         Exit For
      End If
   Next ii
   If m_blnOneRec = False Then
      MsgBox "請勾選一筆申請地址!!!", vbExclamation + vbOKOnly
      Exit Sub
   Else
      If m_blnOneRec = True Then
         If Trim(Me.grdDataList.TextMatrix(ii, 2)) = "" Then
            UpForm.m_AppAddrChange = True '點選申請地址
         Else
            UpForm.m_AppAddrChange = False '點選客戶目前申請地址
         End If
         UpForm.m_AppAddr = Trim(Me.grdDataList.TextMatrix(ii, 1))
         UpForm.m_Zipcode = ""
         '郵遞區號 5 碼時
         If IsNumeric(Left(UpForm.m_AppAddr, 5)) = True Then
            UpForm.m_Zipcode = Left(UpForm.m_AppAddr, 5)
            UpForm.m_AppAddr = Right(UpForm.m_AppAddr, Len(UpForm.m_AppAddr) - 5)
         '郵遞區號 3 碼時
         ElseIf IsNumeric(Left(UpForm.m_AppAddr, 3)) = True Then
            UpForm.m_Zipcode = Left(UpForm.m_AppAddr, 3)
            UpForm.m_AppAddr = Right(UpForm.m_AppAddr, Len(UpForm.m_AppAddr) - 3)
         End If
      End If
   End If
   Me.Hide
   UpForm.Show
   Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
End Sub

Public Function QueryData() As Boolean
   QueryData = True
   Me.grdDataList.Clear
   If Not UpForm.rsAddrNotAlike.EOF And Not UpForm.rsAddrNotAlike.BOF Then
      Set grdDataList.Recordset = UpForm.rsAddrNotAlike
      Me.grdDataList.AddItem ""
      Me.grdDataList.TextMatrix(grdDataList.Rows - 1, 1) = strAddr
      Me.grdDataList.TextMatrix(grdDataList.Rows - 1, 2) = "客戶目前申請地址"
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
Set frm090801_4 = Nothing
End Sub

Private Sub grdDataList_SelChange()
grdDataList.Visible = False
grdDataList.row = grdDataList.MouseRow
grdDataList.col = 0
intRow = Me.grdDataList.row
intCol = Me.grdDataList.col
If grdDataList.row <> 0 Then
   '先全部清空
   For ii = 1 To Me.grdDataList.Rows - 1
      If Me.grdDataList.TextMatrix(ii, 0) = "V" Then
         Me.grdDataList.row = ii
         Me.grdDataList.col = intCol
         Me.grdDataList.Text = ""
         For i = 0 To grdDataList.Cols - 1
            grdDataList.col = i
            grdDataList.CellBackColor = QBColor(15)
         Next i
      End If
   Next ii
   Me.grdDataList.row = intRow
   Me.grdDataList.col = intCol
      grdDataList.Text = "V"
      For i = 0 To grdDataList.Cols - 1
         grdDataList.col = i
         grdDataList.CellBackColor = &HFFC0C0
      Next i
End If
grdDataList.Visible = True
End Sub
