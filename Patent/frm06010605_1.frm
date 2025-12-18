VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm06010605_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "證書號數輸入"
   ClientHeight    =   5745
   ClientLeft      =   -90
   ClientTop       =   1185
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9345
   Begin VB.Frame Frame1 
      Height          =   852
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   9072
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   3672
         TabIndex        =   7
         Top             =   168
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "申請案號"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1500
         MaxLength       =   20
         TabIndex        =   0
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1500
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "FCP"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1980
         MaxLength       =   6
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   2820
         MaxLength       =   1
         TabIndex        =   5
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   3060
         MaxLength       =   2
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7548
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8376
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1620
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1500
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3732
      Left            =   120
      TabIndex        =   9
      Top             =   1860
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   450
      TabIndex        =   13
      Top             =   1530
      Width           =   945
   End
End
Attribute VB_Name = "frm06010605_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/22 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim intLastRow As Integer, intCols As Integer
Dim intWhere As Integer
'Added by Morgan 2023/1/16 電子公文
Public m_RDate As String
Public m_DocWord As String
Public m_DocNo As String
Public m_DocDate As String
Public m_AppNo As String
Public m_DeadLine As String
Public m_NewCP10 As String
Dim m_Done As Boolean
'end 2023/1/16

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Command1_Click()
'Add By Cheng 2002/11/04
Dim strTmp(1 To 2) As String

    'Add By Cheng 2003/12/29
    '若未輸入來函收文日
    If Me.Text5.Text = "" Then
        MsgBox "請輸入來函收文日!!!", vbExclamation + vbOKOnly
        Me.Text5.SetFocus
        Text5_GotFocus
        Exit Sub
    '檢查日期
    ElseIf CheckIsTaiwanDate(Me.Text5.Text) = False Then
        Me.Text5.SetFocus
        Text5_GotFocus
        Exit Sub
    End If
    'End
   intI = 0
   If Option1(0).Value = True Then
      If Text7 = "" Then MsgBox "申請案號不得空白，請重新輸入 !", vbCritical: Exit Sub
      strExc(0) = "select ''," & ChgPatent("", 1) & ",nvl(pa05,nvl(pa06,pa07))," & _
         "pa01,pa02,pa03,pa04 from patent where pa01='FCP' AND pa11='" & Text7 & "'"
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      GridHead
      
      'Added by Morgan 2023/1/16
      If RsTemp.RecordCount = 1 Then
         MSHFlexGrid1.row = 1
         MSHFlexGrid1_Click
         cmdOK(0).Value = True
      End If
      'end 2023/1/16
      
   ElseIf Option1(1).Value = True Then
      If Text3 = "" Then Text3 = "0"
      If Text4 = "" Then Text4 = "00"
      strExc(0) = "select pa01,pa02,pa03,pa04 from patent where pa01='" & Text1 & _
         "' and pa02='" & Text2 & "' and pa03='" & Text3 & "' and pa04='" & Text4 & "'"
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
            'Add By Cheng 2002/11/04
            'edit by nickc 2007/02/05 不用 dll 了
            'If objLawDll.ChkMRec(TransDate(Text5, 2), RsTemp(0).Value & RsTemp(1).Value & RsTemp(2).Value & RsTemp(3).Value, strTmp(1), strTmp(2)) Then
            If ClsLawChkMRec(TransDate(Text5, 2), RsTemp(0).Value & RsTemp(1).Value & RsTemp(2).Value & RsTemp(3).Value, strTmp(1), strTmp(2)) Then
               If strTmp(1) <> "" Then
                  If MsgBox("與櫃台之來函收文記錄 ( " & TransDate(strTmp(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
               End If
            Else
               If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
            End If
         
         If cmdOK(0).Enabled = True Then 'Added by Lydia 2020/01/22 判斷
            '進入畫面二
            strExc(1) = Text1
            strExc(2) = Text2
            strExc(3) = Text3
            strExc(4) = Text4
            frm06010605_2.Show
            Me.Hide
         End If
      End If
   End If
End Sub

Private Sub Form_Activate()
   'Added by Morgan 2023/1/16 電子公文
   If m_AppNo <> "" And m_Done = False Then
      Option1(0).Value = True
      Text7.Text = m_AppNo
      Text5.Text = m_RDate
      Command1.Value = True
      m_Done = True
   End If
   'end 2023/1/16
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   intWhere = 國外_FC
   Option1_Click (0)
   InitGrid 7, MSHFlexGrid1
   GridHead
   Text5 = strSrvDate(2)
   'Add By Cheng 2002/01/31
   Me.Option1(1).Value = True
   SendKeys "{Tab}"
   
   'Added by Lydia 2020/01/22 提前檢查Pat3是否開啟
   'Removed by Morgan 2021/6/25 公報改抓卷宗區，不再往pat3讀取避免當機沒開的情形
   'If Pub_CheckGazetteDir = False Then
   '   cmdOK(0).Enabled = False
   'End If
   'end 2021/6/25
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010605_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK(0).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
 On Error Resume Next
   Select Case Index
      Case 0
         Text7.Enabled = True
         Text2.Enabled = False
         Text3.Enabled = False
         Text4.Enabled = False
         Text7.SetFocus
      Case 1
         Text7.Enabled = False
         Text2.Enabled = True
         Text3.Enabled = True
         Text4.Enabled = True
         Text2.SetFocus
   End Select
End Sub

Private Sub Text1_GotFocus()
  TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_GotFocus()
  TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
  TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
  TextInverse Text4
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> "" Then
      If ChkDate(Text5) Then
         If Val(Text5) > Val(strSrvDate(2)) Then
            MsgBox "來函收文日不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
End Sub

' 確認鈕
Private Sub FormConfirm()
 Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
   If Text5 = "" Then MsgBox "來函收文日不可空白 !", vbCritical: Exit Sub
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" Then
            bolChk = True
            For j = 1 To 4
               strExc(j) = .TextMatrix(i, j + 2)
            Next
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If
   
   'Added by Morgan 2023/1/16 電子公文
   If m_DocNo <> "" Then
      frm06010605_2.m_DocWord = m_DocWord
      frm06010605_2.m_DocNo = m_DocNo
   Else
   'end 2023/1/16
   
      'edit by nickc 2007/02/05 不用 dll 了
      'If objLawDll.ChkMRec(TransDate(Text5, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
      If ClsLawChkMRec(TransDate(Text5, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
         If strTmp(1) <> "" Then
            If MsgBox("與櫃台之來函收文記錄 ( " & TransDate(strTmp(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
         End If
      Else
         If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
      End If
      
   End If 'Added by Morgan 2023/1/16
   frm06010605_2.Show
   Command1.SetFocus
   Me.Hide
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .CellAlignment = flexAlignCenterCenter
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 2000: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 4000: .Text = "專利名稱"
      For i = 3 To 6
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub
