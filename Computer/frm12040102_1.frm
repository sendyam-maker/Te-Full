VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040102_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "依案件性質設定各國催審提申期限"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8295
   Begin VB.TextBox txtCF 
      Height          =   315
      Index           =   1
      Left            =   1215
      TabIndex        =   0
      Top             =   660
      Width           =   1185
   End
   Begin VB.TextBox txtInput 
      Alignment       =   1  '靠右對齊
      Appearance      =   0  '平面
      Height          =   300
      Left            =   6345
      TabIndex        =   6
      Text            =   "動態數入框"
      Top             =   240
      Width           =   1050
   End
   Begin VB.TextBox txtCF 
      Height          =   315
      Index           =   3
      Left            =   1215
      TabIndex        =   1
      Top             =   990
      Width           =   1185
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7695
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102_1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102_1.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102_1.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102_1.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102_1.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102_1.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102_1.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102_1.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102_1.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102_1.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040102_1.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1164
      ButtonWidth     =   1138
      ButtonHeight    =   1111
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4875
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1380
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   8599
      _Version        =   393216
      Rows            =   3
      Cols            =   5
      FixedRows       =   2
      FixedCols       =   0
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FormatString    =   "申請國家|提申期限(天)|審查時間(月)|審查時間(天)|是否更新相同案提申期限"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblMemo 
      Caption         =   "審查時間以 12 個月計算年度(每年以365 天計算),餘數再以每個月 30 天計算。(系統存天數)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   4950
      TabIndex        =   8
      Top             =   660
      Width           =   3165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統別："
      Height          =   180
      Index           =   0
      Left            =   270
      TabIndex        =   7
      Top             =   750
      Width           =   720
   End
   Begin MSForms.Label lblCF 
      Height          =   285
      Index           =   3
      Left            =   2415
      TabIndex        =   5
      Top             =   1020
      Width           =   2475
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4366;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   1
      Left            =   270
      TabIndex        =   3
      Top             =   1050
      Width           =   900
   End
End
Attribute VB_Name = "frm12040102_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/29 Form2.0已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
'Create by Morgan 2008/11/19
Option Explicit

Dim m_bUpdate As Boolean
Dim m_bQuery As Boolean
Dim m_EditMode As Integer
Dim m_iRow As Integer
Dim m_iCol As Integer
Dim m_bScroll As Boolean
Dim ii As Integer, jj As Integer


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
         
      Case vbKeyReturn
         If m_EditMode = 4 Then
            OnAction KeyCode
            KeyCode = 0
         End If
         
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

Private Sub Form_Load()
   m_bUpdate = IsUserHasRightOfFunction("frm12040102_1", strEdit, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040102_1", strFind, False)
   m_EditMode = 0
   
   MoveFormToCenter Me
   txtInput.Visible = False
   QueryRecord -2
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040102_1 = Nothing
End Sub

Private Sub FormClear()

End Sub

Private Sub GoNext()
   Dim bFind As Boolean
   Dim iRow As Integer
   Dim iCol As Integer
   
   With grdDataList
      For iRow = m_iRow To .Rows - 1
         .row = iRow
         For iCol = m_iCol + 1 To 4
            .col = iCol
            If .CellBackColor = .BackColor Then
               bFind = True
               Exit For
            End If
         Next
         If bFind = True Then Exit For
         m_iCol = 0
      Next
      If bFind = True Then
         SetBox
      Else
         txtInput.Visible = False
      End If
   End With
End Sub

Private Sub GrdDataList_Click()
   If m_EditMode = 2 Then
      With grdDataList
         .row = .MouseRow
         .col = .MouseCol
         SetBox
      End With
   End If
End Sub

Private Function ModRecord() As Boolean
   
   Dim stUpdate As String
   
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
   
   With grdDataList
   
   For ii = 2 To .Rows - 1
      stUpdate = ""
      '提申期限
      If .TextMatrix(ii, 1) <> .TextMatrix(ii, 5) Then
         stUpdate = stUpdate & ",CF11=" & CNULL(.TextMatrix(ii, 1), True)
      End If
      '審查時間(天)
      If .TextMatrix(ii, 3) <> .TextMatrix(ii, 6) Then
         stUpdate = stUpdate & ",CF05=" & CNULL(.TextMatrix(ii, 3), True)
      End If
      '是否更新相同案提申期限
      If .TextMatrix(ii, 4) <> .TextMatrix(ii, 7) Then
         stUpdate = stUpdate & ",CF29=" & CNULL(.TextMatrix(ii, 4))
      End If
      
      If stUpdate <> "" Then
         stUpdate = Mid(stUpdate, 2)
         strExc(0) = "select 1 from casefee where CF01='" & txtCF(1) & "' and CF03='" & txtCF(3) & "' and CF02='" & .TextMatrix(ii, 8) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '修改
            strSql = "Update CaseFee set " & stUpdate & _
               " where CF01='" & txtCF(1) & "' and CF03='" & txtCF(3) & "' and CF02='" & .TextMatrix(ii, 8) & "'"
            
            strSql = "begin user_data.user_enabled:=1; " & strSql & "; end;" 'Added by Morgan 2019/2/13 改要記錄修改人員時間
            
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
         '新增
         Else
            strSql = "Insert into CaseFee (CF01,CF02,CF03,CF11,CF05,CF29)" & _
               " values ('" & txtCF(1) & "','" & .TextMatrix(ii, 8) & "','" & txtCF(3) & "'" & _
               "," & CNULL(.TextMatrix(ii, 1), True) & "," & CNULL(.TextMatrix(ii, 3), True) & _
               "," & CNULL(.TextMatrix(ii, 4)) & ")"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql, intI
         End If
      End If
   Next
   End With
   
   cnnConnection.CommitTrans
   ModRecord = True
   txtCF(1).Tag = txtCF(1)
   txtCF(3).Tag = txtCF(3)
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Select Case KeyCode
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         grdDataList.row = 2: grdDataList.col = 1
         SetBox
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         FormClear
         txtCF(1) = ""
         txtCF(3) = ""
         txtCF(1).SetFocus
      ' 第一筆
      Case vbKeyHome:
         QueryRecord -2
      ' 前一筆
      Case vbKeyPageUp:
         QueryRecord -1
      ' 後一筆
      Case vbKeyPageDown:
         QueryRecord 1
      ' 最後一筆
      Case vbKeyEnd:
         QueryRecord 2
      ' 確定
      Case vbKeyF9, vbKeyReturn:
         OnWork
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  txtInput.Visible = False
                  m_EditMode = 0
               End If
            Case Else
               m_EditMode = 0
         End Select
         If m_EditMode = 0 Then
            If txtCF(1).Tag <> "" And txtCF(3).Tag <> "" Then
               txtCF(1) = txtCF(1).Tag
               txtCF(3) = txtCF(3).Tag
               QueryRecord
            End If
         End If
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
   
   If KeyCode <> vbKeyEscape Then
      SetCtrlLock
      UpdateToolbarState
   End If
End Sub

' 使用者按下確定的按紐
Private Sub OnWork()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Select Case m_EditMode
      Case 2:
         If txtInput.Visible = True Then
            If UpdateVar = False Then
               Exit Sub
            Else
               txtInput.Visible = False
            End If
         End If
         
         If ModRecord = True Then
            m_EditMode = 0
         End If
      Case 4:
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End If
         m_EditMode = 0
   End Select
End Sub

' 查詢記錄
Private Function QueryRecord(Optional iOpt As Integer) As Boolean
   
   Dim rstQuery As New ADODB.Recordset
   
   strExc(0) = "select cpm01,cpm02 from casepropertymap where cpm01 in ('P','CFP') and length(cpm02)=3"
   
   Select Case iOpt
      Case -2
         strExc(0) = strExc(0) & " order by 1 asc,2 asc"
      Case -1
         strExc(0) = strExc(0) & " and cpm01||cpm02<'" & txtCF(1) & txtCF(3) & "' order by 1 desc,2 desc"
      Case 1
         strExc(0) = strExc(0) & " and cpm01||cpm02>'" & txtCF(1) & txtCF(3) & "' order by 1 asc,2 asc"
      Case 2
         strExc(0) = strExc(0) & " order by 1 desc,2 desc"
      Case Else
         strExc(0) = strExc(0) & " and cpm01='" & txtCF(1) & "' AND cpm02='" & txtCF(3) & "'"
   End Select
   
   With rstQuery
   .MaxRecords = 1
   .CursorLocation = adUseClient
   .Open strExc(0), cnnConnection, adOpenForwardOnly, adLockReadOnly
   If .RecordCount > 0 Then
      txtCF(1) = .Fields(0): txtCF(1).Tag = txtCF(1)
      txtCF(3) = .Fields(1): txtCF(3).Tag = txtCF(3)
      
      strExc(0) = "select NA03 C0,CF11 C1,trunc(Cf05/365)*12+round(MOD(CF05,365)/30,1) C2,CF05 C3,CF29 C4" & _
         ",CF11 C5,CF05 C6,CF29 C7,NA01 C8 from Nation,CaseFee" & _
         " where length(na01)=3 and cf02(+)=na01 and CF01(+)='" & txtCF(1) & "' and CF03(+)='" & txtCF(3) & "'"
      'P 只抓台灣,大陸,香港,澳門,PCT
      If txtCF(1) = "P" Then
         strExc(0) = strExc(0) & " and na01 in ('000','020','013','044','056')"
      Else
         strExc(0) = strExc(0) & " and na01>='010' and na01<'9' and na01 not in ('000','020','013','044')"
      End If
      strExc(0) = strExc(0) & " order by NA01 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Set grdDataList.Recordset = RsTemp.Clone
         SetDataListWidth
         QueryRecord = True
      Else
         MsgBox "查無資料!"
      End If
   ElseIf iOpt = -1 Then
      MsgBox "已經是第一筆!"
   ElseIf iOpt = 1 Then
      MsgBox "已經是最後筆!"
   Else
      MsgBox "查無資料!"
   End If
   End With
   Set rstQuery = Nothing
   UpdateToolbarState
End Function

Private Sub SetBox()
   Dim lngLeft As Long, lngTop As Long
   With grdDataList
      If .CellBackColor = .BackColor Then
         If .row > 0 And .col >= 0 Then
            If .TextMatrix(.row, 0) <> "" Then
               Select Case .col
                  Case 1
                     txtInput.MaxLength = 3
                  Case 2
                     txtInput.MaxLength = 0
                  Case 4
                     txtInput.MaxLength = 1
               End Select
               txtInput.FontName = .CellFontName
               txtInput.FontSize = .CellFontSize
               txtInput.Alignment = .CellAlignment \ 5
               txtInput.Text = .TextMatrix(.row, .col)
               txtInput.Tag = txtInput.Text
               txtInput.Width = .ColWidth(.col)
               txtInput.Height = .RowHeight(.row)
               m_iRow = .row: m_iCol = .col
               
               
               'X 座標
               lngLeft = .Left + 25
               For ii = 0 To .col - 1
                  lngLeft = lngLeft + .ColWidth(ii)
               Next
               'Y 座標
               lngTop = .Top + 25
               '固定列高度
               For ii = 0 To .FixedRows - 1
                  lngTop = lngTop + .RowHeight(ii)
               Next
               '資料列高度
               For ii = .TopRow To .row - 1
                  lngTop = lngTop + .RowHeight(ii)
               Next
               
               '超過畫面時(下)
               If lngTop + txtInput.Height > .Top + .Height Then
                  m_bScroll = True
                  .TopRow = .TopRow + 1
                  m_bScroll = False
                  lngTop = lngTop - txtInput.Height
               '超過畫面時(上)
               ElseIf .row < .TopRow Then
                  m_bScroll = True
                  .TopRow = .TopRow - 1
                  m_bScroll = False
               End If
               
               txtInput.Left = lngLeft: txtInput.Top = lngTop
               If .ColAlignment(.col) < 3 Then
                  txtInput.Alignment = 0
               ElseIf .ColAlignment(.col) < 6 Then
                  txtInput.Alignment = 2
               Else
                  txtInput.Alignment = 1
               End If
               txtInput.Visible = True
               txtInput_GotFocus
               txtInput.SetFocus
            End If
         End If
      End If
   End With
End Sub

Private Sub SetCtrlLock()
   Dim bLock As Boolean
   If m_EditMode = 4 Then
      bLock = False
   Else
      bLock = True
   End If
   txtCF(1).Locked = bLock
   txtCF(3).Locked = bLock
End Sub

Private Sub SetDataListWidth()
   With grdDataList
      .Visible = False
      '.FormatString = "申請國家|提申期限(天)|審查時間(月)|審查時間(天)|是否更新相同案提申期限"
      .RowHeightMin = txtInput.Height
      
      .MergeCol(0) = True
      .MergeCol(4) = True
      .MergeRow(0) = True
      .MergeCells = flexMergeFree
      
      
      .ColWidth(0) = 1800
      .TextMatrix(0, 0) = "申請國家"
      .TextMatrix(1, 0) = .TextMatrix(0, 0)
      .TextMatrix(0, 1) = "提申期限"
      .TextMatrix(1, 1) = "(天)"
      .TextMatrix(0, 2) = "審查時間"
      .TextMatrix(1, 2) = "(月)"
      .TextMatrix(0, 3) = .TextMatrix(0, 2)
      .TextMatrix(1, 3) = "(天)"
      .ColWidth(4) = 2300
      .TextMatrix(0, 4) = "是否更新相同案" & vbCrLf & "提申期限"
      .TextMatrix(1, 4) = .TextMatrix(0, 4)
      
      For jj = 0 To 1
         .row = jj
         For ii = 0 To 4
            .col = ii
            .CellFontBold = True
         Next
      Next
      
      For ii = 5 To .Cols - 1
         .ColWidth(ii) = 0
      Next
      
      .ColAlignmentFixed = flexAlignCenterCenter
      .ColAlignment(1) = flexAlignRightCenter
      .ColAlignment(2) = flexAlignRightCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignCenterCenter
      .Visible = True
   End With
   SetGridColor
End Sub

Private Sub SetGridColor()
   Dim lngColor As Long
   With grdDataList
      .Visible = False
      lngColor = &H8000000F
      For ii = 1 To .Rows - 1
         .row = ii
         .col = 0: .CellBackColor = lngColor
         .col = 3: .CellBackColor = lngColor
      Next
      .Visible = True
   End With
End Sub

Private Sub grdDataList_Scroll()
   If m_bScroll = False Then
      txtInput.Visible = False
   End If
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: OnAction vbKeyF5
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

Private Sub txtCF_Change(Index As Integer)
   If Index = 3 Then
      lblCF(3) = ""
      If Len(txtCF(3)) = 3 Then
         If ClsPDGetCaseProperty(txtCF(1), txtCF(3), strExc(1)) = True Then
            
            If strExc(1) = "（無）" Then
               ClsPDGetCaseProperty txtCF(1), txtCF(3), strExc(1), True
            End If
            lblCF(3) = strExc(1)
         End If
      End If
   End If
End Sub

Private Sub txtCF_GotFocus(Index As Integer)
   TextInverse txtCF(Index)
   CloseIme
End Sub

Private Sub txtCF_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtInput_DblClick()
   If m_iCol = 4 Then
      If txtInput.Text <> "Y" Then
         txtInput.Text = "Y"
      Else
         txtInput.Text = ""
      End If
   End If
End Sub

Private Sub txtInput_GotFocus()
   TextInverse txtInput
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 37, 38, 39, 40
         If UpdateVar = True Then
            With grdDataList
            Select Case KeyCode
               Case 37 '左
                  If m_iCol = 4 Then
                     .row = m_iRow
                     .col = 2
                     SetBox
                  ElseIf m_iCol = 2 Then
                     .row = m_iRow
                     .col = 1
                     SetBox
                  ElseIf m_iCol = 1 Then
                     If m_iRow - 1 > 1 Then
                        .row = m_iRow - 1
                        .col = 4
                        SetBox
                     End If
                  End If
               Case 38 '上
                  If m_iRow - 1 > 1 Then
                     .row = m_iRow - 1
                     .col = m_iCol
                     SetBox
                  End If
               Case 39 '右
                  If m_iCol = 1 Then
                     .row = m_iRow
                     .col = 2
                     SetBox
                  ElseIf m_iCol = 2 Then
                     .row = m_iRow
                     .col = 4
                     SetBox
                  ElseIf m_iCol = 4 Then
                     If m_iRow + 1 < .Rows - 1 Then
                        .row = m_iRow + 1
                        .col = 1
                        SetBox
                     End If
                  End If
               Case 40 '下
                  If m_iRow + 1 < .Rows - 1 Then
                     .row = m_iRow + 1
                     .col = m_iCol
                     SetBox
                  End If
            End Select
            End With
         Else
            TextInverse txtInput
         End If
   End Select
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = vbKeyReturn Then
      If UpdateVar = True Then
         GoNext
      Else
         TextInverse txtInput
      End If
      
   ElseIf KeyAscii = vbKeyEscape Then
      txtInput = txtInput.Tag
      TextInverse txtInput
      
   ElseIf KeyAscii <> 8 Then
      strExc(0) = ""
      If txtInput.SelLength > 0 Then
         If txtInput.SelStart > 0 Then
            strExc(0) = Left(txtInput, txtInput.SelStart)
         End If
         If txtInput.SelLength <> Len(txtInput) Then
            strExc(0) = strExc(0) & Right(txtInput, txtInput.SelStart + txtInput.SelLength)
         End If
      Else
         strExc(0) = txtInput
      End If
      If m_iCol = 4 Then
         KeyAscii = UpperCase(KeyAscii)
         If Chr(KeyAscii) <> "Y" Then
            KeyAscii = 0
            Beep
         End If
      ElseIf KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
         If m_iCol = 2 Then
            '控制到小一位
           If Len(strExc(0)) > 1 And Left(Right(strExc(0), 2), 1) = "." Then
               KeyAscii = 0
               Beep
            End If
         End If
      ElseIf m_iCol = 2 And KeyAscii = Asc(".") Then
         If InStr(strExc(0), ".") > 0 Then
            KeyAscii = 0
            Beep
         End If
      Else
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub txtInput_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 37, 38, 39, 40
         If txtInput.Visible = True Then
            txtInput_GotFocus
         End If
   End Select
End Sub

Private Sub txtInput_LostFocus()
   If txtInput.Visible = True Then
      If UpdateVar = True Then
         txtInput.Visible = False
      Else
         txtInput.SetFocus
      End If
   End If
End Sub
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bUpdate Then
            tlbar.Buttons(2).Enabled = True
         Else
            tlbar.Buttons(2).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(4).Enabled = True
         Else
            tlbar.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(6).Enabled = True
            tlbar.Buttons(7).Enabled = True
            tlbar.Buttons(8).Enabled = True
            tlbar.Buttons(9).Enabled = True
         Else
            tlbar.Buttons(6).Enabled = False
            tlbar.Buttons(7).Enabled = False
            tlbar.Buttons(8).Enabled = False
            tlbar.Buttons(9).Enabled = False
         End If
         tlbar.Buttons(11).Enabled = False
         tlbar.Buttons(12).Enabled = False
         tlbar.Buttons(14).Enabled = True
      
      Case Else
         tlbar.Buttons(2).Enabled = False
         tlbar.Buttons(4).Enabled = False
         tlbar.Buttons(6).Enabled = False
         tlbar.Buttons(7).Enabled = False
         tlbar.Buttons(8).Enabled = False
         tlbar.Buttons(9).Enabled = False
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         tlbar.Buttons(14).Enabled = False
   End Select
   
End Sub
'更新資料表
Private Function UpdateVar() As Boolean
   Dim bolCancel As Boolean, dblMonths As Double
   With grdDataList
   If bolCancel = False Then
      .TextMatrix(m_iRow, m_iCol) = Format(txtInput.Text)
      If m_iCol = 2 Then
         If .TextMatrix(m_iRow, m_iCol) = "" Then
            .TextMatrix(m_iRow, m_iCol + 1) = ""
         '審查時間(月)-->(天)
         Else
            dblMonths = .TextMatrix(m_iRow, m_iCol)
            '12個月以365天計,餘數每1個月以30天計(四捨五入)
            .TextMatrix(m_iRow, m_iCol + 1) = 365 * (Int(dblMonths) \ 12) + Round(30 * (dblMonths - 12 * (Int(dblMonths) \ 12)))
            If Len(.TextMatrix(m_iRow, m_iCol + 1)) > 4 Then
               MsgBox "審查時間超過限制！", vbExclamation
               Exit Function
            End If
         End If
      End If
      UpdateVar = True
   End If
   End With
   
End Function

