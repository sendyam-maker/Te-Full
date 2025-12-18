VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm06010601 
   BorderStyle     =   1  '單線固定
   Caption         =   "實審通知日輸入"
   ClientHeight    =   5745
   ClientLeft      =   225
   ClientTop       =   990
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
      Left            =   1344
      MaxLength       =   7
      TabIndex        =   4
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
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   7560
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   9072
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   2
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   5160
         MaxLength       =   1
         TabIndex        =   1
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   3840
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "FCP"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "申請案號"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   6000
         TabIndex        =   3
         Top             =   192
         Width           =   800
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3672
      Left            =   96
      TabIndex        =   11
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "若為實審通知日的假來函輸入，來函收文日請輸入111111"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2640
      TabIndex        =   14
      Top             =   1485
      Width           =   4500
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
      Left            =   240
      TabIndex        =   13
      Top             =   1480
      Width           =   945
   End
End
Attribute VB_Name = "frm06010601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/18 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim intLastRow As Integer, intWhere As Integer
'Added by Morgan 2017/5/9 電子公文
Public m_RDate As String
Public m_DocWord As String
Public m_DocNo As String
Public m_DocDate As String
Public m_AppNo As String
Public m_DeadLine As String
Public m_NewCP10 As String
Dim m_Done As Boolean
'end 2017/5/9


Public Sub Clear()
   Text7 = Empty
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   InitGrid 9, MSHFlexGrid1
   GridHead
   Option1(1).Value = True
   Option1_Click 1
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 2
         Unload Me
   End Select
End Sub

Public Sub Command1_Click()
   intI = 0
   If Option1(0).Value = True Then
      If Text7 = "" Then MsgBox "申請案號不得空白，請重新輸入 !", vbCritical: Exit Sub
      strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,'',pa01,pa02,pa03,pa04,'' from patent where PA01='FCP' AND " & _
         "pa11='" & Text7 & "' and pa09='" & 台灣國家代號 & "' union " & _
         "select distinct(" & ChgCaseprogress("", 1) & "||'N') as No," & _
         "nvl(cp37,nvl(cp38,cp38)) as Name," & _
         "nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,'' from caseprogress where " & _
         "cp36='" & Text7 & "' and cp01='FCP' AND (cp01,cp02,cp03,cp04) not in " & _
         "(select pa01,pa02,pa03,pa04 from patent where PA01='FCP' AND " & _
         "pa11='" & Text7 & "' and pa09='" & 台灣國家代號 & "')"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      Me.Tag = "0"
   Else
      If Trim(Text2) = "" Then MsgBox "請輸入本所案號!!", vbExclamation: Text2.SetFocus: Exit Sub 'Added by Morgan 2015/7/7
      
      If Text3 = "" Then Text3 = "0"
      If Text4 = "" Then Text4 = "00"
      strExc(0) = "select " & ChgPatent("", 1) & ",nvl(pa05,nvl(pa06,pa07)),'',pa11,pa01," & _
         "pa02,pa03,pa04,'' from patent where " & ChgPatent(Text1 & Text2 & Text3 & Text4) & _
         " and pa09='" & 台灣國家代號 & "'"
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      'If intI = 1 Then
      '   If Not IsNull(rsTemp.Fields(3)) Or rsTemp.Fields(3) <> "" Then
      '      GridHead
      '      MsgBox "此案號巳有申請案號，請以申請案號輸入 !", vbCritical
      '      Exit Sub
      '   End If
      'End If
      Me.Tag = "1"
   End If
   
   'Modified by Morgan 2019/10/31 電子公文除外 Ex:FCP-061493
   If intI = 1 And m_DocNo = "" Then
      'Add by Amy 2013/08/27 +讓frm06010601_2可以出現提示"要輸入實審提出日期或系統日"
      strExc(0) = "Select '1' From CaseProgress  Where cp01='" & RsTemp.Fields("pa01") & "' And cp02='" & RsTemp.Fields("pa02") & "' " & _
                       "And cp03='" & RsTemp.Fields("pa03") & "' And cp04='" & RsTemp.Fields("pa04") & "' And InStr('" & NewCasePtyList & "',cp10)>0 " & _
                       "And cp09>'B' "
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI > 0 Then
           Text5 = "111111"
      End If
      'end 2013/08/27
   End If
   GridHead
   If MSHFlexGrid1.Rows = 2 Then
      MSHFlexGrid1.row = 1
      MSHFlexGrid1.col = 8
      MSHFlexGrid1.Text = "v"
      cmdOK_Click 0
   End If
End Sub

Private Sub Form_Activate()
   'Added by Morgan 2017/5/9 電子公文
   If m_AppNo <> "" And m_Done = False Then
      Option1(0).Value = True
      'Modified by Moran2017/8/30
      'Text7.Text = m_AppNo
      'Modified by Morgan 2019/5/10 判斷舉發案才抓9碼,否則衍生設計案會抓不到 Ex:FCP-060687
      If Mid(m_AppNo, 10, 1) = "N" Then
         Text7.Text = Left(m_AppNo, 9)
      Else
         Text7.Text = m_AppNo
      End If
      'end 2019/5/10
      'end 2017/8/30
      Text5.Text = m_RDate
      Command1.Value = True
      m_Done = True
   End If
   'end 2017/5/9
   
On Error Resume Next
   If Option1(1).Value = True Then Text2.SetFocus

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   InitGrid 9, MSHFlexGrid1
   GridHead
   Text5 = strSrvDate(2)
   Option1_Click 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010601 = Nothing
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
         Text7.SetFocus
      Case 1
         Text7.Enabled = False
         Text2.Enabled = True
         Text3.Enabled = True
         Text4.Enabled = True
         Text2.SetFocus
   End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
 Dim strTmp(1 To 2) As String
   If Text5 <> "" Then
      If ChkDate(Text5) Then
         If Val(Text5) > Val(strSrvDate(2)) Then
            MsgBox "來函收文日不可大於系統日 !", vbCritical
            Cancel = True
         
         'Removed by Morgan 2015/7/7 確定時有檢查,User輸案號前會先改日期致都會彈訊息故此處取消
         'Else
         '   'edit by nickc 2007/02/05 不用 dll 了
         '   'If objLawDll.ChkMRec(TransDate(Text5.Text, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
         '   If ClsLawChkMRec(TransDate(Text5.Text, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
         '      If strTmp(1) <> "" Then
         '         If MsgBox("與櫃台之來函收文記錄 ( " & TransDate(strTmp(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
         '      End If
         '   Else
         '      If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
         '   End If
         'end 2015/7/7
         
         End If
      Else
         Cancel = True
      End If
   End If
End Sub

' 確認鈕
Private Sub FormConfirm()
 Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String
   If Text5 = "" Then
      MsgBox "來函收文日不可空白 !", vbCritical
      Text5.SetFocus
      Exit Sub
    ElseIf CheckIsTaiwanDate(Me.Text5.Text) = False Then
      Text5.SetFocus
        Text5_GotFocus
      Exit Sub
   End If
   With MSHFlexGrid1
      .col = 8
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 8) = "v" Then
            bolChk = True
            For j = 1 To 4
               strExc(j) = .TextMatrix(i, j + 3)
            Next
            If Option1(0).Value Then
               strExc(5) = "1"
            Else
               strExc(5) = "2"
            End If
            Exit For
         End If
      Next
   End With
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If
   'edit by nickc 2007/02/05 不用 dll 了
   'If objLawDll.ChkMRec(TransDate(Text5.Text, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
   If ClsLawChkMRec(TransDate(Text5.Text, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
      If strTmp(1) <> "" Then
         If MsgBox("與櫃台之來函收文記錄 ( " & TransDate(strTmp(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
      End If
   'Modified by Morgan 2017/5/10 電子公文
   'Else
   ElseIf Me.m_DocNo = "" Then
   'end 2017/5/10
      If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
   End If
   frm06010601_1.Show
   If frm06010601_1.QueryData() = False Then
      MsgBox "沒有符合條件的資料", vbCritical, "查詢資料"
      Unload frm06010601_1
   Else
      Command1.SetFocus
      Me.Hide
   End If
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 1500: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 5500: .Text = "專利名稱"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      For i = 3 To 8
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
   End With
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text3_GotFocus()
   InverseTextBox Text3
End Sub

Private Sub Text4_GotFocus()
   InverseTextBox Text4
End Sub

Private Sub Text5_GotFocus()
   InverseTextBox Text5
End Sub

Private Sub Text7_GotFocus()
   InverseTextBox Text7
End Sub


