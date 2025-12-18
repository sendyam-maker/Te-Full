VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm04010501 
   BorderStyle     =   1  '單線固定
   Caption         =   "實審通知日輸入"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   945
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
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   9072
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   5460
         MaxLength       =   1
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "尋找(&F)"
         Default         =   -1  'True
         Height          =   375
         Left            =   6300
         TabIndex        =   5
         Top             =   180
         Width           =   800
      End
      Begin VB.OptionButton Option1 
         Caption         =   "本所案號"
         Height          =   255
         Index           =   1
         Left            =   3060
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
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   4140
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "P"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   4632
         MaxLength       =   6
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   5700
         MaxLength       =   2
         TabIndex        =   4
         Top             =   240
         Width           =   375
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
      Left            =   1344
      MaxLength       =   8
      TabIndex        =   6
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
      Left            =   240
      TabIndex        =   13
      Top             =   1500
      Width           =   948
   End
End
Attribute VB_Name = "frm04010501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/16 改成Form2.0 (MSHFlexGrid1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Modify by Morgan 2008/8/18 已改開窗定稿，地址條列印功能取消
Option Explicit

Dim intLastRow As Integer, intWhere As Integer
'Added by Morgan 2014/1/14
Public m_DocNo As String
Public m_AppNo As String
Public m_RDate As String
Dim m_Done As Boolean
'end 2014/1/14
Public m_DocWord As String 'Added by Morgan 2014/4/17
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
'2016/10/5 END


Public Sub Clear()
   Text2.Text = ""
   Text3.Text = ""
   Text4.Text = ""
   Text7.Text = ""
   InitGrid 9, MSHFlexGrid1
   GridHead
   '預設在申請案號欄
   Option1(0).Value = True
   Option1_Click 0
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 2 '結束
         Unload Me
   End Select
End Sub

Public Sub Command1_Click()
   intI = 0
   If Option1(0).Value = True Then
      If Text7 = "" Then MsgBox "申請案號不得空白，請重新輸入 !", vbCritical: Exit Sub
      strExc(0) = "select " & ChgPatent("", 1) & " as No,nvl(pa05,nvl(pa06,pa07)) as Name," & _
         "'' as RName,'',pa01,pa02,pa03,pa04,'' from patent where PA01='P' AND " & _
         "pa11='" & Text7 & "' and pa09='" & 台灣國家代號 & "' union " & _
         "select distinct(" & ChgCaseprogress("", 1) & "||'N') as No," & _
         "nvl(cp37,nvl(cp38,cp38)) as Name," & _
         "nvl(cp37,nvl(cp38,cp39)) as RName,'',cp01,cp02,cp03,cp04,'' from caseprogress where " & _
         "cp36='" & Text7 & "' and cp01='P' AND (cp01,cp02,cp03,cp04) not in " & _
         "(select pa01,pa02,pa03,pa04 from patent where PA01='P' AND " & _
         "pa11='" & Text7 & "' and pa09='" & 台灣國家代號 & "')"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      Me.Tag = "0"
   Else
      If Text3 = "" Then Text3 = "0"
      If Text4 = "" Then Text4 = "00"
      strExc(0) = "select " & ChgPatent("", 1) & ",nvl(pa05,nvl(pa06,pa07)),'',pa11,pa01," & _
         "pa02,pa03,pa04,'' from patent where " & ChgPatent(Text1 & Text2 & Text3 & Text4) & _
         " and pa09='" & 台灣國家代號 & "'"
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
      If intI = 1 Then
         If Not IsNull(RsTemp.Fields(3)) Or RsTemp.Fields(3) <> "" Then
            MsgBox "此案號已有申請案號，請以申請案號輸入 !", vbCritical
            ' 90.07.17 modify by louis (顯示完訊息不可進入下一畫面)
            InitGrid 9, MSHFlexGrid1
            GridHead
            Exit Sub
         End If
      End If
      Me.Tag = "1"
   End If
   GridHead
   If MSHFlexGrid1.Rows = 2 Then
      OnlyOneRec MSHFlexGrid1, 8
      FormConfirm
   End If
   
End Sub

Private Sub Form_Activate()
   'Add By Sindy 2017/12/27
   If m_strIR01 <> "" And m_Done = False Then
      Option1(0).Value = True
      Text5.Text = m_RDate
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   '2017/12/27 END
   'Added by Morgan 2014/1/14
   ElseIf m_AppNo <> "" And m_Done = False Then
      Option1(0).Value = True
      Text7.Text = m_AppNo
      Text5.Text = m_RDate
      Command1.Value = True
      m_Done = True
   End If
   'end 2014/1/14
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   InitGrid 9, MSHFlexGrid1
   GridHead
   Text5 = strSrvDate(2)
   Option1_Click 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010501 = Nothing
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


Private Sub Text5_Validate(Cancel As Boolean)
 Dim strTmp(1 To 2) As String
   If Text5 <> "" Then
      If ChkDate(Text5) Then
         Text5 = TransDate(Text5, 1) 'Add by Morgan 2009/7/31 改可輸西元年但自動轉民國年
         If Val(Text5) > Val(strSrvDate(2)) Then
            MsgBox "來函收文日不可大於系統日 !", vbCritical
            Cancel = True
         Else
            'edit by nickc 2007/02/05 不用 dll 了
            'If objLawDll.ChkMRec(TransDate(Text5.Text, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
            If ClsLawChkMRec(TransDate(Text5.Text, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
               If strTmp(1) <> "" Then
                  If MsgBox("與櫃台之來函收文記錄 ( " & TransDate(strTmp(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
               End If
            'Modified by Morgan 2014/5/5 排除無期限電子公文
            'Else
            ElseIf m_DocNo = "" Then
            'end 2014/5/5
               If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
            End If
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
      
   'Add by Morgan 2009/7/31
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
   
   If TxtValidate = False Then Exit Sub
   
'Remove by Morgan 2009/7/31 日期跳離時已有檢查
'   'edit by nickc 2007/02/05 不用 dll 了
'   'If objLawDll.ChkMRec(TransDate(Text5.Text, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
'   If ClsLawChkMRec(TransDate(Text5.Text, 2), strExc(1) & strExc(2) & strExc(3) & strExc(4), strTmp(1), strTmp(2)) Then
'      If strTmp(1) <> "" Then
'         If MsgBox("與櫃台之來函收文記錄 ( " & TransDate(strTmp(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
'      End If
'   Else
'      If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
'   End If
   
   'Add By Sindy 2017/12/27
   If m_strIR01 <> "" Then
      If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> strExc(1) & strExc(2) & strExc(3) & strExc(4) Then
         MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
         Exit Sub
      End If
   End If
   '2017/12/27 END
   ' 90.07.17 modify by louis (無資料不顯示下一畫面)
   'frm04010501_1.Show
   frm04010501_1.Visible = False
   'Add By Sindy 2016/10/5
   frm04010501_1.m_strIR01 = m_strIR01
   frm04010501_1.m_strIR02 = m_strIR02
   frm04010501_1.m_strIR03 = m_strIR03
   frm04010501_1.m_strIR04 = m_strIR04
   '2016/10/5 END
   If frm04010501_1.QueryData = True Then
      frm04010501_1.Visible = True
      frm04010501_1.Show
'      Text7.SetFocus
      '91.12.8 ADD BY SONIA
      frm04010501_1.cmdOK(2).SetFocus
      '91.12.8 END
      Me.Hide
   Else
      MsgBox "沒有符合條件的資料", vbOKOnly + vbInformation, "查詢資料"
      '91.12.8 ADD BY SONIA
      frm04010501_1.Visible = True
      frm04010501_1.Show
      Me.Hide
      frm04010501_1.cmdOK(3).SetFocus
      '91.12.8 END
   End If
   'Me.Hide
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

