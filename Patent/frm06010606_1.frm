VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm06010606_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "異議/舉發受理函輸入"
   ClientHeight    =   5745
   ClientLeft      =   -1230
   ClientTop       =   1800
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9345
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8388
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7560
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3672
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   6482
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
   Begin VB.Frame Frame1 
      Height          =   852
      Left            =   120
      TabIndex        =   12
      Top             =   540
      Width           =   9072
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   5
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "FCP"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   0
         Top             =   180
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "對照號數"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
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
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   3624
         TabIndex        =   7
         Top             =   168
         Width           =   800
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   9180
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   9180
      Y1              =   1824
      Y2              =   1824
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   450
      TabIndex        =   13
      Top             =   1470
      Width           =   945
   End
End
Attribute VB_Name = "frm06010606_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/23 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim intLastRow As Integer, intCols As Integer
Dim intWhere As Integer
'Added by Morgan 2017/5/10 電子公文
Public m_RDate As String
Public m_DocWord As String
Public m_DocNo As String
Public m_DocDate As String
Public m_AppNo As String
Public m_DeadLine As String
Public m_NewCP10 As String
Dim m_Done As Boolean
Dim m_Retry As Boolean
'end 2017/5/10


Public Sub Clear()
   'Text1 = Empty
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Command1_Click()
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
   If Option1(0).Value = True Then
      If Text7 = "" Then MsgBox "對照號數不得空白，請重新輸入 !", vbCritical: Exit Sub
      'Modify By Cheng 2002/04/12
'      strExc(0) = "SELECT ''," & ChgCaseprogress("", 1) & "||'N',CPM03," & _
'         SQLDate("CP27") & ",NVL(CP37,NVL(CP38,CP39))," & _
'         "NVL(CP40,NVL(CP41,CP42)),CP01,CP02,CP03,CP04,CP09 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP WHERE CP01='FCP' AND CP36='" & Text7 & "' AND " & _
'         "(CP27 IS NOT NULL OR CP27<>'') AND (SUBSTR(CP09,1,1)='A' OR SUBSTR(CP09,1,1)='B') AND " & _
'         "PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA23<>1 AND CP01=CPM01(+) AND CP10=CPM02(+)"
      ' 91.09.13 modify by louis
      'strExc(0) = "SELECT ''," & ChgCaseprogress("", 1) & "||'N',CPM03," & _
      '   SQLDate("CP27") & ",NVL(CP37,NVL(CP38,CP39))," & _
      '   "NVL(CP40,NVL(CP41,CP42)),CP01,CP02,CP03,CP04,CP09 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP WHERE CP01='FCP' AND CP36='" & Text7 & "' AND " & _
      '   "(CP27 IS NOT NULL OR CP27<>'') AND ( CP09<'C' ) AND " & _
      '   "PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA23<>1 AND CP01=CPM01(+) AND CP10=CPM02(+)"
      strExc(0) = "SELECT ''," & ChgCaseprogress("", 1) & "||'N',CPM03," & _
         SQLDate("CP27") & ",NVL(CP37,NVL(CP38,CP39))," & _
         "NVL(CP40,NVL(CP41,CP42)),CP01,CP02,CP03,CP04,CP09,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
         "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP " & _
         "WHERE CP01='FCP' AND CP36='" & Text7 & "' AND " & _
         "(CP27 IS NOT NULL OR CP27<>'') AND ( CP09<'C' ) AND " & _
         "PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA23<>1 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
         "ORDER BY SORTFIELD DESC "
   ElseIf Option1(1).Value = True Then
      If Text3 = "" Then Text3 = "0"
      If Text4 = "" Then Text4 = "00"
      'Modify By Cheng 2002/04/12
'      strExc(0) = "SELECT ''," & ChgCaseprogress("", 1) & "||'N',CPM03," & _
'         SQLDate("CP27") & ",NVL(CP37,NVL(CP38,CP39))," & _
'         "NVL(CP40,NVL(CP41,CP42)),CP01,CP02,CP03,CP04,CP09 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP WHERE CP01='" & Text1 & _
'         "' AND CP02='" & Text2 & "' AND CP03='" & Text3 & "' AND CP04='" & Text4 & "' AND " & _
'         "(CP27 IS NOT NULL OR CP27<>'') AND (SUBSTR(CP09,1,1)='A' OR SUBSTR(CP09,1,1)='B') AND " & _
'         "PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA23<>1 AND CP01=CPM01(+) AND CP10=CPM02(+)"
      ' 91.09.13 modify by louis
      'strExc(0) = "SELECT ''," & ChgCaseprogress("", 1) & "||'N',CPM03," & _
      '   SQLDate("CP27") & ",NVL(CP37,NVL(CP38,CP39))," & _
      '   "NVL(CP40,NVL(CP41,CP42)),CP01,CP02,CP03,CP04,CP09 FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP WHERE CP01='" & Text1 & _
      '   "' AND CP02='" & Text2 & "' AND CP03='" & Text3 & "' AND CP04='" & Text4 & "' AND " & _
      '   "(CP27 IS NOT NULL OR CP27<>'') AND ( CP09<'C' ) AND " & _
      '   "PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA23<>1 AND CP01=CPM01(+) AND CP10=CPM02(+)"
      strExc(0) = "SELECT ''," & ChgCaseprogress("", 1) & "||'N',CPM03," & _
         SQLDate("CP27") & ",NVL(CP37,NVL(CP38,CP39))," & _
         "NVL(CP40,NVL(CP41,CP42)),CP01,CP02,CP03,CP04,CP09,DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
         "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP " & _
         "WHERE CP01='" & Text1 & _
         "' AND CP02='" & Text2 & "' AND CP03='" & Text3 & "' AND CP04='" & Text4 & "' AND " & _
         "(CP27 IS NOT NULL OR CP27<>'') AND ( CP09<'C' ) AND " & _
         "PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA23<>1 AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
         "ORDER BY SORTFIELD DESC "
   End If
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   ' 只有一筆則直接進入到下一畫面
   If MSHFlexGrid1.Rows = 2 Then
      MSHFlexGrid1.row = 1
      GridClick MSHFlexGrid1, intLastRow, 0
      FormConfirm
      
   'Added by Morgan 2017/5/10 電子公文
   ElseIf RsTemp.RecordCount = 0 Then
      If m_AppNo <> "" And m_Retry = False Then m_Retry = True
   'end 2017/5/10
   End If
End Sub

Private Sub Form_Activate()
   'Added by Morgan 2017/5/10 電子公文
   If m_AppNo <> "" And m_Done = False Then
      Option1(0).Value = True
      Text7.Text = Left(m_AppNo, 9)
      Text5.Text = m_RDate
      m_Retry = False
      Command1.Value = True
      If m_Retry = True Then
         Text7.Text = m_AppNo
         Command1.Value = True
      End If
      m_Done = True
   End If
   'end 2017/5/10
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   Option1_Click (0)
   InitGrid 11, MSHFlexGrid1
   GridHead
   Text5 = strSrvDate(2)
   'Add By Cheng 2002/01/31
   Me.Option1(1).Value = True
   SendKeys "{Tab}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010606_1 = Nothing
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
      If ChkDate(Text5) = False Then
         Cancel = True
      ElseIf Val(Text5) > Val(strSrvDate(2)) Then
         MsgBox "來函收文日不可大於系統日 !", vbCritical
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
            For j = 1 To 5
               strExc(j) = .TextMatrix(i, j + 5)
            Next
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If
   'edit by nickc 2007/02/05 不用 dll 了
   'If objLawDll.ChkMRec(TransDate(Text5, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
   If ClsLawChkMRec(TransDate(Text5, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
      If strTmp(1) <> "" Then
         If MsgBox("與櫃台之來函收文記錄 ( " & TransDate(strTmp(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
      End If
   'Modified by Morgan 2017/5/10 電子公文
   'Else
   ElseIf frm06010606_1.m_DocNo = "" Then
   'end 2017/5/10
      If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
   End If
   
   'Added by Morgan 2017/5/10 電子公文
   frm06010606_2.m_DocWord = frm06010606_1.m_DocWord
   frm06010606_2.m_DocNo = frm06010606_1.m_DocNo
   frm06010606_2.m_DocDate = frm06010606_1.m_DocDate
   frm06010606_2.m_AppNo = frm06010606_1.m_AppNo
   frm06010606_2.m_DeadLine = frm06010606_1.m_DeadLine
   'end 2017/5/10
   frm06010606_2.Show
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
      .col = 1: .ColWidth(1) = 1500: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1500: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1200: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 4000: .Text = "專利名稱"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 4000: .Text = "對照名稱"
      For i = 6 To 10
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub
