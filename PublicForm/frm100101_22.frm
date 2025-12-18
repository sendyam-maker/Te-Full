VERSION 5.00
Begin VB.Form frm100101_22 
   BackColor       =   &H80000004&
   BorderStyle     =   1  '單線固定
   Caption         =   "投資法務開拓客戶資料查詢"
   ClientHeight    =   5745
   ClientLeft      =   1440
   ClientTop       =   2310
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   2
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   17
      Top             =   570
      Width           =   615
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   10
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   16
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   13
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   15
      Top             =   3210
      Width           =   5535
   End
   Begin VB.TextBox textECD 
      Height          =   270
      Index           =   1
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   14
      Top             =   900
      Width           =   1215
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   11
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   13
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   9
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   12
      Top             =   2550
      Width           =   2775
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   8
      Left            =   4530
      MaxLength       =   30
      TabIndex        =   11
      Top             =   2220
      Width           =   2775
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   7
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   10
      Top             =   2220
      Width           =   2775
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   6
      Left            =   4530
      MaxLength       =   30
      TabIndex        =   9
      Top             =   1890
      Width           =   2775
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   5
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   8
      Top             =   1890
      Width           =   2775
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   3
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   7
      Top             =   1230
      Width           =   2775
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   4
      Left            =   4530
      MaxLength       =   30
      TabIndex        =   6
      Top             =   1230
      Width           =   2775
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   12
      Left            =   4530
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox textECD 
      Height          =   264
      Index           =   15
      Left            =   1680
      MaxLength       =   12
      TabIndex        =   4
      Top             =   3870
      Width           =   1935
   End
   Begin VB.TextBox textECD 
      Height          =   1215
      Index           =   16
      Left            =   1680
      MaxLength       =   2000
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   4200
      Width           =   5535
   End
   Begin VB.TextBox textECD 
      Height          =   270
      Index           =   14
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   2
      Top             =   3540
      Width           =   405
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "回前畫面"
      Height          =   400
      Index           =   0
      Left            =   6840
      TabIndex        =   0
      Top             =   75
      Width           =   1230
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束"
      Height          =   400
      Index           =   1
      Left            =   8070
      TabIndex        =   1
      Top             =   75
      Width           =   800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "屬性代號："
      Height          =   180
      Left            =   720
      TabIndex        =   31
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "屬性名稱："
      Height          =   180
      Index           =   0
      Left            =   2520
      TabIndex        =   30
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "目前編號："
      Height          =   180
      Index           =   0
      Left            =   720
      TabIndex        =   29
      Top             =   930
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "收  件  人："
      Height          =   180
      Left            =   720
      TabIndex        =   28
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "地        址："
      Height          =   180
      Left            =   720
      TabIndex        =   27
      Top             =   1920
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "國       籍："
      Height          =   180
      Index           =   0
      Left            =   765
      TabIndex        =   26
      Top             =   2910
      Width           =   855
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "公司名稱："
      Height          =   180
      Left            =   720
      TabIndex        =   25
      Top             =   1590
      Width           =   900
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Height          =   180
      Index           =   0
      Left            =   3480
      TabIndex        =   24
      Top             =   600
      Width           =   3525
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      Caption         =   "xxx"
      Height          =   180
      Index           =   1
      Left            =   2550
      TabIndex        =   23
      Top             =   2910
      Width           =   270
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "E-MAIL："
      Height          =   180
      Index           =   1
      Left            =   840
      TabIndex        =   22
      Top             =   3240
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "狀態："
      Height          =   180
      Index           =   2
      Left            =   1080
      TabIndex        =   21
      Top             =   3900
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   180
      Index           =   3
      Left            =   1080
      TabIndex        =   20
      Top             =   4230
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否寄電子報："
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   19
      Top             =   3570
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(N:不寄)"
      Height          =   180
      Index           =   3
      Left            =   2220
      TabIndex        =   18
      Top             =   3570
      Width           =   645
   End
End
Attribute VB_Name = "frm100101_22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/12 改成Form2.0 ;  Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
Option Explicit

Public cmdState As Integer
Dim strTmp As String
Dim rsContact As ADODB.Recordset
Dim m_bReadGrid As Boolean
Dim oText As TextBox
Dim idx As Integer


Private Sub DataGrid1_Click()
   m_bReadGrid = True
End Sub

Private Sub DataGrid1_Validate(Cancel As Boolean)
   m_bReadGrid = False
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100101_22 = Nothing
End Sub

Private Sub cmdok1_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
   End Select
End Sub

Sub StrMenu()
Dim strKey  As String, strKey1 As String
Dim varRef As Variant
Dim ii As Integer
   
   varRef = Split(Me.Tag, "-")
   For ii = LBound(varRef) To UBound(varRef)
      If ii = 0 Then strKey = varRef(ii)
      If ii = 1 Then strKey1 = varRef(ii)
   Next ii
   
   'Add By Sindy 2011/01/03 檢查國內外權限
   If CheckSR12(Me.Tag) = False Then
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
   
   strExc(0) = "select * from expandcusdetail where ecd02='" & strKey & "' and ecd01=" & strKey1 & " "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ShowRecord RsTemp, strKey, strKey1
   Else
      ShowNoData
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub ShowRecord(ByRef p_Rst As ADODB.Recordset, strKey As String, strKey1 As String)
   Dim rsECD As ADODB.Recordset
   Dim CUID(1 To 6) As String
   
   ClearField
   SetCtrlReadOnly True
   Set rsECD = p_Rst.Clone
   With rsECD
      If .RecordCount > 0 Then
         For Each oText In textECD
            idx = oText.Index
            oText.Text = "" & .Fields("ECD" & Format(idx, "0#"))
         Next
         '屬性名稱
         lbl1(0).Caption = ""
         strExc(0) = "select * from expandcusattr where eca01='" & strKey & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            lbl1(0).Caption = RsTemp.Fields("eca02")
         End If
         '國籍
         lbl1(1).Caption = ""
         If ClsPDGetNation(Left(textECD(10), 3), strTmp) = True Then
            lbl1(1).Caption = strTmp
         End If
      End If
   End With
End Sub

Private Sub ClearField()
   Dim oLabel As LABEL
   For Each oText In textECD
      oText.Text = Empty
   Next
   For Each oLabel In lbl1
      oLabel.Caption = Empty
   Next
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In textECD
      oText.Locked = bLocked
   Next
End Sub
