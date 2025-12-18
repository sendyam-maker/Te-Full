VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210133_INV 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件未付帳款"
   ClientHeight    =   3660
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8604
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   8604
   Begin VB.Frame Frame1 
      Height          =   1400
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8250
      Begin VB.TextBox txtNP07 
         Height          =   270
         Left            =   7920
         MaxLength       =   9
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   192
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   1
         Left            =   1104
         MaxLength       =   1
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   3
         Left            =   960
         MaxLength       =   9
         TabIndex        =   6
         Text            =   "A1K28"
         Top             =   624
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   4
         Left            =   5196
         MaxLength       =   9
         TabIndex        =   5
         Text            =   "Y2776600"
         Top             =   624
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   5
         Left            =   960
         MaxLength       =   9
         TabIndex        =   4
         Text            =   "A1K27"
         Top             =   936
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   6
         Left            =   5196
         MaxLength       =   9
         TabIndex        =   3
         Top             =   936
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   2
         Left            =   5400
         MaxLength       =   1
         TabIndex        =   2
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(1)"
         Height          =   252
         Index           =   1
         Left            =   6096
         TabIndex        =   17
         Top             =   636
         Width           =   2004
      End
      Begin VB.Label Label1 
         Caption         =   "列印申請人：       (要印:Y)"
         Height          =   252
         Index           =   17
         Left            =   50
         TabIndex        =   16
         Top             =   264
         Width           =   2496
      End
      Begin VB.Label Label1 
         Caption         =   "固定請款對象："
         Height          =   252
         Index           =   18
         Left            =   3960
         TabIndex        =   15
         Top             =   660
         Width           =   1296
      End
      Begin VB.Label Label1 
         Caption         =   "列印對象："
         Height          =   252
         Index           =   19
         Left            =   48
         TabIndex        =   14
         Top             =   972
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "請款對象："
         Height          =   252
         Index           =   20
         Left            =   48
         TabIndex        =   13
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "固定列印對象："
         Height          =   252
         Index           =   21
         Left            =   3960
         TabIndex        =   12
         Top             =   972
         Width           =   1296
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(0)"
         Height          =   252
         Index           =   0
         Left            =   1848
         TabIndex        =   11
         Top             =   636
         Width           =   2004
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(2)"
         Height          =   252
         Index           =   2
         Left            =   1848
         TabIndex        =   10
         Top             =   972
         Width           =   2004
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(3)"
         Height          =   252
         Index           =   3
         Left            =   6096
         TabIndex        =   9
         Top             =   972
         Width           =   2004
      End
      Begin VB.Label Label1 
         Caption         =   "合併列印請款單：       (要印:Y)"
         Height          =   252
         Index           =   13
         Left            =   3960
         TabIndex        =   8
         Top             =   264
         Width           =   2556
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHGrid1 
      Height          =   2796
      Left            =   120
      TabIndex        =   0
      Top             =   1812
      Width           =   5800
      _ExtentX        =   10224
      _ExtentY        =   4932
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   1
      FixedCols       =   0
      ScrollBars      =   2
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
      _Band(0).Cols   =   1
   End
End
Attribute VB_Name = "frm210133_INV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create Amy 2025/04/10
Option Explicit

Public intQuery As Integer '1-案件未付帳款 /2-目前請款單設定 資訊
Dim arrCol() As String, arrWidth() As String, i As Integer
Dim stA1K04 As String, stA1K27 As String, stA1K28 As String, stA1K29 As String, stTM56 As String, stTM69 As String

Private Sub Form_Load()
   Frame1.Visible = False
   MSHGrid1.Visible = False
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   intQuery = Empty
   Set frm210133_INV = Nothing
End Sub

Public Function doQuery(ByVal stCaseNo1 As String, ByVal stCaseNo2 As String, ByVal stCaseNo3 As String, ByVal stCaseNo4 As String, Optional ByRef intCntRec As Long) As Boolean
   Dim RsQ As New ADODB.Recordset, intQ As Integer, stQ As String, stTP As String
   
   doQuery = False
   intCntRec = 0
   stTP = stCaseNo1 & "-" & stCaseNo2 & IIf(stCaseNo3 & stCaseNo4 = "000", "", stCaseNo3 & "-" & stCaseNo4)
   If intQuery = 1 Then
'*** 案件未付帳款 資訊 ***
      Me.Caption = "案件未付帳款 資訊 (案號:" & stTP & ")"
      MSHGrid1.Rows = 2
      MSHGrid1.Clear
      
      stQ = PUB_GetNotPayA1K(" And a1k13='" & stCaseNo1 & "' And  a1k14='" & stCaseNo2 & "' And  a1k15='" & stCaseNo3 & "' And a1k16='" & stCaseNo4 & "'", True)
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, stQ)
      If intQ = 1 Then
         doQuery = True
         intCntRec = RsQ.RecordCount
         Set MSHGrid1.Recordset = RsQ
      End If
      Me.Width = 6200
      Me.Height = 3700
      MSHGrid1.Top = 100
      Call SetGridWidth(True)
      Set RsQ = Nothing
      MSHGrid1.Visible = True
   Else
'*** 目前請款單設定 資訊 ***
      Me.Caption = "目前請款單設定 (案號:" & stTP & ")"
      Call SetTxt(0)
      Call Pub_GetCloseA1KData(1, Me.Name, stCaseNo1, stCaseNo2, stCaseNo3, stCaseNo4, stA1K29, txtNP07, stA1K04, stA1K27, stA1K28, stTM56, stTM69)
      txtA1K(1) = stA1K04 '列印申請人
      txtA1K(2) = "" '合併列印請款單
      txtA1K(3) = stA1K28: Call txtA1K_Validate(3, False) '請款對象
      txtA1K(4) = stTM56: Call txtA1K_Validate(4, False) '固定請款對象
      txtA1K(5) = stA1K27: Call txtA1K_Validate(5, False) '列印對象
      txtA1K(6) = stTM69: Call txtA1K_Validate(6, False) '固定列印對象
      Call SetTxt(1, True)
      Me.Width = 8680
      Me.Height = 2030
      Frame1.Visible = True
   End If
End Function

Private Sub SetGridWidth(Optional ByVal IsFirst As Boolean = False)
   Dim stField As String, stWidth As String
   
   MSHGrid1.Visible = False
   If IsFirst = True Then
      stField = "請款單號,請款日,請款幣別金額,請款幣別"
      stWidth = "1000,900,1200,1000"
      arrCol = Split(stField, ",")
      arrWidth = Split(stWidth, ",")
   End If
   MSHGrid1.Cols = UBound(arrCol) + 1
   MSHGrid1.row = 0
   For i = 0 To MSHGrid1.Cols - 1
      MSHGrid1.col = i
      MSHGrid1.Text = arrCol(i)
      MSHGrid1.ColWidth(i) = Val(arrWidth(i))
      MSHGrid1.CellAlignment = flexAlignLeftCenter
   Next i
   MSHGrid1.Visible = True
End Sub

'intChoose:0:清資料 /1-是否Lock
Private Sub SetTxt(intChoose As Integer, Optional ByVal IsLock As Boolean)
   Dim obj As Object
  
   For Each obj In txtA1K
      If intChoose = 0 Then
         obj.Text = ""
      Else
         obj.Locked = IsLock
      End If
   Next
   
   If intChoose = 0 Then
      For Each obj In lblName
         obj.Caption = ""
      Next
   End If
End Sub

Private Sub txtA1K_Validate(Index As Integer, Cancel As Boolean)
   Dim stName As String
   
   If Index >= 3 And Index <= 6 And txtA1K(Index) <> MsgText(601) Then
      If Left(txtA1K(Index), 1) = 代理人編號 Then
         Call ClsPDGetAgent(txtA1K(Index), stName)
      ElseIf Left(txtA1K(Index), 1) = 客戶編號 Then
         stName = GetCustomerName(txtA1K(Index))
      End If
      lblName(Index - 3).Caption = stName
   End If
End Sub
