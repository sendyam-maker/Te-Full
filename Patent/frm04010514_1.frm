VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm04010514_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "初審及公佈通知來函輸入 "
   ClientHeight    =   5745
   ClientLeft      =   240
   ClientTop       =   1335
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9330
   Begin VB.Frame Frame1 
      Height          =   792
      Left            =   192
      TabIndex        =   12
      Top             =   600
      Width           =   9012
      Begin VB.OptionButton Option1 
         Caption         =   "申請案號"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   5940
         MaxLength       =   2
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   5700
         MaxLength       =   1
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   4875
         MaxLength       =   6
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4380
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "P"
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號"
         Height          =   255
         Index           =   1
         Left            =   3300
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   6510
         TabIndex        =   7
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
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
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1500
      MaxLength       =   8
      TabIndex        =   8
      Top             =   1500
      Width           =   1332
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3612
      Left            =   120
      TabIndex        =   9
      Top             =   1980
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   6376
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
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   9180
      Y1              =   1896
      Y2              =   1896
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   120
      X2              =   9180
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Top             =   1500
      Width           =   948
   End
End
Attribute VB_Name = "frm04010514_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/16 改成Form2.0 (MSHFlexGrid1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Create by Morgan 2009/11/24 自內專核准函輸入抽出
Option Explicit

Dim intLastRow As Integer, intCols As Integer
Dim intWhere As Integer
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_AppNo As String
Public m_RDate As String
Dim m_Done As Boolean
'2016/10/5 END


Public Sub Clear()
   Text7 = Empty
   InitGrid 9, MSHFlexGrid1
   GridHead
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
   Dim rsA As New ADODB.Recordset
   Dim StrSQLa As String
   intI = 0
   If Me.Option1(0).Value Then
      If Text7 = "" Then
         MsgBox "申請案號不得空白，請重新輸入 !", vbCritical
         Me.Text7.SetFocus
         Text7_GotFocus
         Exit Sub
      End If
      'Add by Lydia 2014/10/31 設別名f0
      strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,'',pa01,pa02,pa03,pa04,'' from patent f0 where PA01='P' AND PA09<>'000'" & _
         " AND pa11='" & Text7 & "'"
   Else
      If Me.Text1.Text = "" Then
         MsgBox "系統類別不得空白，請重新輸入 !", vbCritical
         Me.Text1.SetFocus
         Text1_GotFocus
         Exit Sub
      Else
         If Me.Text1.Text <> "P" And Me.Text1.Text <> "PS" Then
            MsgBox "系統類別輸入錯誤，請重新輸入 !", vbCritical
            Me.Text1.SetFocus
            Text1_GotFocus
            Exit Sub
         End If
      End If
      If Me.Text2.Text = "" Then
         MsgBox "本所案號不得空白，請重新輸入 !", vbCritical
         Me.Text2.SetFocus
         Text2_GotFocus
         Exit Sub
      End If
      
      If Me.Text3.Text = "" Then Me.Text3.Text = "0"
      If Me.Text4.Text = "" Then Me.Text4.Text = "00"
      '專利基本檔
      'Add by Lydia 2014/10/31 設別名f0
      strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,'',pa01,pa02,pa03,pa04,'' from patent f0 where PA01='" & Me.Text1.Text & "'" & _
         " AND PA02='" & Me.Text2.Text & "' AND PA03='" & Me.Text3.Text & "'" & _
         " AND PA04='" & Me.Text4.Text & "' AND PA09<>'000'"
   
   End If
      
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
     If FMP2open = True And FMP2openSQL <> "" Then
        strExc(0) = strExc(0) & FMP2openSQL
        strExc(0) = Replace(strExc(0), "f0.CP", "f0.PA")
     End If
     
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   If MSHFlexGrid1.Rows = 2 Then
      MSHFlexGrid1.row = 1
      MSHFlexGrid1_Click
      FormConfirm
   End If
End Sub

Private Sub Form_Activate()
   'Add by Sindy 2016/10/5
   If m_strIR01 <> "" And m_Done = False Then
      Option1(0).Value = True
      'Text7.Text = m_AppNo
      Text5.Text = m_RDate
      'Command1.Value = True
      m_Done = True
      'Add By Sindy 2017/12/27
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
      '2017/12/27 END
   End If
   '2016/10/5 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   InitGrid 9, MSHFlexGrid1
   GridHead
   Text5 = strSrvDate(2)
   SendKeys "{Tab}"
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010514_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 8
   cmdOK(0).SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0 '申請案號
   Me.Text7.Enabled = True
   Me.Text2.Enabled = False
   Me.Text3.Enabled = False
   Me.Text4.Enabled = False
   Me.Text7.SetFocus
Case 1 '本所案號
   Me.Text7.Enabled = False
   Me.Text2.Enabled = True
   Me.Text3.Enabled = True
   Me.Text4.Enabled = True
   Me.Text2.SetFocus
End Select
End Sub

Private Sub Text1_GotFocus()
TextInverse Me.Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
If KeyAscii <> 80 And KeyAscii <> 83 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Me.Text1.Text <> "P" Then
   MsgBox "系統類別只能輸入 P !!!", vbExclamation + vbOKOnly
   Cancel = True
   Me.Text1.SetFocus
   Text1_GotFocus
End If
End Sub

Private Sub Text2_GotFocus()
TextInverse Me.Text2
End Sub

Private Sub Text3_GotFocus()
TextInverse Me.Text3
End Sub

Private Sub Text4_GotFocus()
TextInverse Me.Text4
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 <> "" Then
      If ChkDate(Text5) Then
         Text5 = TransDate(Text5, 1) '改可輸西元年但自動轉民國年
         If Val(Text5) > Val(strSrvDate(2)) Then
            MsgBox "來函收文日不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   
   If Text5 = "" Then
      MsgBox "來函收文日不可空白 !", vbCritical
      Text5.SetFocus
      Exit Function
   Else
      Text5_Validate Cancel
      If Cancel = True Then
         Text5.SetFocus
         Text5_GotFocus
         Exit Function
      End If
      
   End If
   TxtValidate = True
   
End Function

' 確認鈕
Private Sub FormConfirm()
 
   Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
   Dim strPA01 As String
   Dim strPA02 As String
   Dim strPA03 As String
   Dim strPA04 As String
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   
   If TxtValidate = False Then Exit Sub
      
   With MSHFlexGrid1
      .col = 8
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 8) = "v" Then
            bolChk = True
            For j = 1 To 4
               strExc(j) = .TextMatrix(i, j + 3)
            Next
            strPA01 = .TextMatrix(i, 4)
            strPA02 = .TextMatrix(i, 5)
            strPA03 = .TextMatrix(i, 6)
            strPA04 = .TextMatrix(i, 7)
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If
   
   'Add By Sindy 2017/12/27
   If m_strIR01 <> "" Then
      If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> strPA01 & strPA02 & strPA03 & strPA04 Then
         MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
         Exit Sub
      End If
   End If
   '2017/12/27 END
   'Add By Sindy 2016/10/5
   frm04010514_2.m_strIR01 = m_strIR01
   frm04010514_2.m_strIR02 = m_strIR02
   frm04010514_2.m_strIR03 = m_strIR03
   frm04010514_2.m_strIR04 = m_strIR04
   '2016/10/5 END
   frm04010514_2.Show
   If Me.Option1(0).Value Then
      Option1_Click 0
   Else
      Option1_Click 1
   End If
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

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
