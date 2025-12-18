VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170102 
   BorderStyle     =   1  '單線固定
   Caption         =   "婚喪互助扣款計算"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5730
   Begin VB.TextBox Textwf01 
      Height          =   270
      Left            =   1400
      MaxLength       =   7
      TabIndex        =   0
      Top             =   510
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "扣款計算(&O)"
      Height          =   405
      Index           =   0
      Left            =   3300
      TabIndex        =   2
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   4500
      TabIndex        =   3
      Top             =   60
      Width           =   1065
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1135
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   1485
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   5530
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   2
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Left            =   1395
      TabIndex        =   7
      Top             =   870
      Width           =   2700
      VariousPropertyBits=   27
      Caption         =   "當事人姓名及互助原因"
      Size            =   "4762;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "處理中資料："
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   870
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "婚喪互助日期："
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   1260
   End
End
Attribute VB_Name = "frm170102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/27 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2008/12/23 add by sonia
Option Explicit

Dim i As Integer, j As Integer


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0 '扣款計算
         If Progress Then
            textWF01_Validate (False) '處理完重新顯示資料
            textWF01_GotFocus
         Else
            MsgBox "未點選任何婚喪互助名單！", vbInformation
         End If
      Case 1 '結束
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170102 = Nothing
End Sub

Private Sub grd1_SelChange()
Dim i As Integer

   GRD1.Visible = False
   GRD1.row = GRD1.MouseRow
   '己計算過不可點選
   GRD1.col = 5
   If GRD1.Text <> "" Then
      MsgBox "此筆資料已計算過扣款，不可再點選！", vbInformation
      GRD1.Visible = True
      Exit Sub
   End If
   
   GRD1.col = 0
   If GRD1.row <> 0 Then
      If GRD1.Text = "V" Then
         GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
      Else
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
   GRD1.Visible = True
   cmdok(0).Enabled = True
   cmdok(0).Default = True
End Sub

Private Sub textWF01_GotFocus()
   InverseTextBox Textwf01
   cmdok(0).Enabled = False
   Label2.Visible = False
   Label3.Visible = False
   PBar1.Visible = False
End Sub

Private Sub textWF01_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textWF01_Validate(Cancel As Boolean)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

   If Textwf01 <> "" Then
      'Modified by Morgan 2022/10/31 +wf11
      strSql = "SELECT '',sqldateT(WF01),wf02,st02,wf03||' '||decode(wf03,'1','婚','2','喪','')||decode(wf11,'1','(父親)','2','(母親)','3','(配偶)','4','(兒子)','5','(女兒)',''),sqldateT(WF04) FROM WeddingAndFuneral,staff " & _
               "where WF02=st01(+) and WF01=" & DBDATE(Textwf01) & " order by WF01,WF02 "
      If rsTmp.State = 1 Then rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount = 0 Then
         MsgBox "無此日期的婚喪互助名單！", vbInformation
         Cancel = True
         textWF01_GotFocus
      Else
         cmdok(0).Enabled = True
         cmdok(0).Default = True
      End If
      Set GRD1.Recordset = rsTmp
      SetGrd
      rsTmp.Close
      Set rsTmp = Nothing
   End If
   
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("V", "日期", "互助同仁", "姓名", "原因", "扣款日期")
   arrGridHeadWidth = Array(200, 800, 1000, 1000, 1000, 1000)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Function Progress() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim m_WF01 As String
Dim m_WF02 As String
Dim m_WF03 As String
Dim m_WFA04 As Variant

On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   Progress = False
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If Trim(GRD1.Text) = "V" Then
         GRD1.col = 0
         GRD1.Text = ""
         For j = 0 To GRD1.Cols - 1
             GRD1.col = j
             GRD1.CellBackColor = QBColor(15)
         Next j
         
         GRD1.col = 1
         If Not IsNull(GRD1.Text) Then m_WF01 = GRD1.Text
         GRD1.col = 2
         If Not IsNull(GRD1.Text) Then m_WF02 = GRD1.Text
         '顯示處理中當事人資料及進度BAR
         Label2.Visible = True
         Label3.Visible = True
         PBar1.Visible = True
         GRD1.col = 3
         If Not IsNull(GRD1.Text) Then Label3.Caption = GRD1.Text
         GRD1.col = 4
         If Not IsNull(GRD1.Text) Then
            m_WF03 = Mid(GRD1.Text, 1, 1)
            Label3.Caption = Label3.Caption & " " & GRD1.Text
         End If
         PBar1.Value = 0
         '開始扣款寫檔
         Set rsTmp = New ADODB.Recordset
         'Modified by Morgan 2020/11/4 需判斷 到職日<=互助日 才扣 & 離職日>互助日 也要扣 --辜,劉經理
         'strSql = "select * from staff,salarydata where st04='1' and st01<>" & CNULL(m_WF02) & " and st01=sd01(+) and (nvl(sd09,0)+nvl(sd10,0)>0) order by st01  "
         'Modified by Morgan 2023/5/15 排除配偶
         'strSql = "select * from staff,salarydata where st13<=" & DBDATE(m_WF01) & " and (st04='1' or st51>" & DBDATE(m_WF01) & ") and st01<>" & CNULL(m_WF02) & " and st01=sd01(+) and (nvl(sd09,0)+nvl(sd10,0)>0) order by st01  "
         strSql = "select * from staff,salarydata where st13<=" & DBDATE(m_WF01) & " and (st04='1' or st51>" & DBDATE(m_WF01) & ") and st01<>" & CNULL(m_WF02) & " and st01=sd01(+) and (nvl(sd09,0)+nvl(sd10,0)>0)" & _
            " and not exists(select * from WeddingAndFuneral where wf01=" & CNULL(DBDATE(m_WF01)) & " and wf02=" & CNULL(m_WF02) & " and wf12=st01) order by st01  "
         'end 2020/11/4
         If rsTmp.State = 1 Then rsTmp.Close
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <> 0 Then
            PBar1.Min = 0
            PBar1.max = rsTmp.RecordCount
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
               PBar1.Value = rsTmp.AbsolutePosition
               m_WFA04 = 0
               If m_WF03 = "1" Then
                  If IsNull(rsTmp.Fields("sd09")) = False Then m_WFA04 = rsTmp.Fields("sd09")
               Else
                  If IsNull(rsTmp.Fields("sd10")) = False Then m_WFA04 = rsTmp.Fields("sd10")
               End If
               '2008/12/30 MODIFY BY SONIA 加wfa05
               'If m_WFA04 <> 0 Then cnnConnection.Execute "insert into wfamount (wfa01,wfa02,wfa03,wfa04) values (" & CNULL(DBDATE(m_WF01)) & "," & CNULL(DBDATE(m_WF02)) & "," & CNULL(rsTmp.Fields("st01")) & "," & m_WFA04 & " ) "
               If m_WFA04 <> 0 Then cnnConnection.Execute "insert into wfamount (wfa01,wfa02,wfa03,wfa04,wfa05) values (" & CNULL(DBDATE(m_WF01)) & "," & CNULL(DBDATE(m_WF02)) & "," & CNULL(rsTmp.Fields("st01")) & "," & m_WFA04 & ",'1' ) "
               DoEvents
               rsTmp.MoveNext
            Loop
         End If
         '更新為已計算扣款
         cnnConnection.Execute "update WeddingAndFuneral set wf04=" & strSrvDate(1) & " where wf01=" & CNULL(DBDATE(m_WF01)) & " and wf02=" & CNULL(m_WF02)
         Progress = True
      End If
   Next i
   
   cnnConnection.CommitTrans
   If Progress Then
      MsgBox "婚喪互助扣款計算完畢！", vbInformation
   End If
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
End Function
