VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm06010604_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "一般來函輸入"
   ClientHeight    =   5508
   ClientLeft      =   132
   ClientTop       =   948
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6108
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   9336
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1344
      MaxLength       =   7
      TabIndex        =   10
      Top             =   1740
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Index           =   2
      Left            =   8376
      TabIndex        =   13
      Top             =   63
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   360
      Index           =   0
      Left            =   7536
      TabIndex        =   12
      Top             =   63
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   1152
      Left            =   120
      TabIndex        =   14
      Top             =   540
      Width           =   9072
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   8
         Top             =   780
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "專利號數"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   5
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "FCP"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1170
         MaxLength       =   20
         TabIndex        =   0
         Top             =   180
         Width           =   1335
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
         Left            =   3336
         TabIndex        =   9
         Top             =   180
         Width           =   800
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3132
      Left            =   120
      TabIndex        =   11
      Top             =   2220
      Width           =   9072
      _ExtentX        =   16002
      _ExtentY        =   5525
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
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   9180
      Y1              =   2333.194
      Y2              =   2333.194
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   9180
      Y1              =   2359.808
      Y2              =   2359.808
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   240
      TabIndex        =   15
      Top             =   1770
      Width           =   945
   End
End
Attribute VB_Name = "frm06010604_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/4/23 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
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
'end 2017/5/10
'Added by Lydia 2023/09/25
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
'end 2023/09/25

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 2
         Unload Me
   End Select
End Sub

Public Sub Clear()
   Text7 = Empty
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   'Modified by Lydia 2019/11/14 比照frm06010608_1
   'MSHFlexGrid1.Rows = 1
   InitGrid 9, MSHFlexGrid1
   GridHead
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
   intI = 0
   If Option1(0).Value = True Then
      If Text7 = "" Then MsgBox "申請案號不得空白，請重新輸入 !", vbCritical: Exit Sub
      strExc(0) = "select " & ChgService("", 1) & " as No,nvl(SP05,nvl(SP06,SP07)) as Name," & _
         "'' as RName,'',SP01,SP02,SP03,SP04,'' from SERVICEPRACTICE where SP01='FG' AND " & _
         "SP11='" & Text7 & "' and SP09='" & 台灣國家代號 & "' union " & _
         "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,'',pa01,pa02,pa03,pa04,'' from patent where PA01='FCP' AND " & _
         "pa11='" & Text7 & "' and pa09='" & 台灣國家代號 & "' union " & _
         "select distinct(" & ChgCaseprogress("", 1) & "||'N') as No,nvl(cp37,nvl(cp38,cp38)) as Name," & _
         "nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,'' from caseprogress " & _
         "where (CP01='FCP' OR CP01='FG') AND cp36='" & Text7 & "' and (cp01,cp02,cp03,cp04) not in " & _
         "(select pa01,pa02,pa03,pa04 from patent where PA01='FCP' AND " & _
         "pa11='" & Text7 & "' and pa09='" & 台灣國家代號 & "' UNION " & _
         "select SP01,SP02,SP03,SP04 from SERVICEPRACTICE where SP01='FG' AND " & _
         "SP11='" & Text7 & "' and SP09='" & 台灣國家代號 & "')"
         
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      GridHead
      'Added by Morgan 2019/9/24--陳亭妙
      If RsTemp.RecordCount = 1 Then
         '進入畫面二
         strExc(1) = RsTemp("sp01")
         strExc(2) = RsTemp("sp02")
         strExc(3) = RsTemp("sp03")
         strExc(4) = RsTemp("sp04")
         'Added by Lydia 2023/09/25
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> strExc(1) & strExc(2) & strExc(3) & strExc(4) Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Sub
            End If
         End If
         frm06010604_2.m_strIR01 = m_strIR01
         frm06010604_2.m_strIR02 = m_strIR02
         frm06010604_2.m_strIR03 = m_strIR03
         frm06010604_2.m_strIR04 = m_strIR04
         'end 2023/09/25
         frm06010604_2.Show
         Me.Hide
      End If
      'end 2019/9/24
   ElseIf Option1(1).Value = True Then
      If Text3 = "" Then Text3 = "0"
      If Text4 = "" Then Text4 = "00"
      strExc(0) = "select ''," & ChgService("", 1) & " as No,nvl(SP05,nvl(SP06,SP07)) as Name," & _
         "'' as RName,SP11,SP01,SP02,SP03,SP04 from SERVICEPRACTICE where SP01='" & Text1 & _
         "' and SP02='" & Text2 & "' and SP03='" & Text3 & "' and SP04='" & Text4 & _
         "' and SP09='" & 台灣國家代號 & "' union " & _
         "select ''," & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,pa11,pa01,pa02,pa03,pa04 from patent where pa01='" & Text1 & _
         "' and pa02='" & Text2 & "' and " & _
         "pa03='" & Text3 & "' and pa04='" & Text4 & "' and pa09='" & 台灣國家代號 & "'"
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '進入畫面二
         strExc(1) = Text1
         strExc(2) = Text2
         strExc(3) = Text3
         strExc(4) = Text4
         frm06010604_2.Show
         Text2.SetFocus    'add by sonia 2016/3/24 輸完回第一畫面游標停在本所案號欄
         Me.Hide
      End If
   ElseIf Option1(2).Value = True Then
      If Text6 = "" Then MsgBox "專利號數不得空白，請重新輸入 !", vbCritical: Exit Sub
      strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,'',pa01,pa02,pa03,pa04,'' from patent where PA01='FCP' AND " & _
         "pa22='" & Text6 & "' and pa09='" & 台灣國家代號 & "' union " & _
         "select distinct(" & ChgCaseprogress("", 1) & "||'N') as No,nvl(cp37,nvl(cp38,cp38)) as Name," & _
         "nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,'' from caseprogress where " & _
         "(CP01='FCP' OR CP01='FG') AND cp36='" & Text6 & "' and (cp01,cp02,cp03,cp04) not in " & _
         "(select pa01,pa02,pa03,pa04 from patent where PA01='FCP' AND pa22='" & Text6 & "' and " & _
         "pa09='" & 台灣國家代號 & "')"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      GridHead
   End If
   
End Sub

Private Sub Form_Activate()
   'Added by Lydia 2023/09/25
   If m_strIR01 <> "" And m_Done = False Then
      Option1(0).Value = True
      Text5.Text = m_RDate
      Text5.Tag = m_RDate
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   'end 2023/09/25
   'Added by Morgan 2017/5/10 電子公文
   'Modified by Lydia 2023/09/25 + Else
   ElseIf m_AppNo <> "" And m_Done = False Then
      Option1(0).Value = True
      Text7.Text = m_AppNo
      Text5.Text = m_RDate
      Command1.Value = True
      m_Done = True
   End If
   'end 2017/5/10
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   Option1_Click (0)
   InitGrid 9, MSHFlexGrid1
   GridHead
   Text5 = strSrvDate(2)
   'Add By Cheng 2002/01/31
   Me.Option1(1).Value = True
   SendKeys "{Tab}"
   'Modify By Cheng 2002/05/10
'   SendKeys "{Tab}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010604_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 8
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
         Text6.Enabled = False
         Text7.SetFocus
      Case 1
         Text7.Enabled = False
         Text2.Enabled = True
         Text3.Enabled = True
         Text4.Enabled = True
         Text6.Enabled = False
         Text1.SetFocus
         'Add By Cheng 2002/01/31
         SendKeys "{Tab}"
      Case 2
         Text7.Enabled = False
         Text2.Enabled = False
         Text3.Enabled = False
         Text4.Enabled = False
         Text6.Enabled = True
         Text6.SetFocus
   End Select
End Sub

Private Sub Text1_GotFocus()
  TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "FG" And Text1 <> "FCP" Then
      MsgBox "系統別錯誤，請重新輸入 !", vbCritical
      Cancel = True
   End If
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
         If .TextMatrix(i, 8) = "v" Then
            bolChk = True
            For j = 1 To 4
               strExc(j) = .TextMatrix(i, j + 3)
            Next
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If
   frm06010604_2.Show
   Command1.SetFocus
   If Text2.Enabled Then Text2.SetFocus
   Me.Hide
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 1500: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 4000: .Text = "專利名稱"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      For i = 3 To 8
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub
