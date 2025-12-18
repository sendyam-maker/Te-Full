VERSION 5.00
Begin VB.Form frm060110 
   BorderStyle     =   1  '單線固定
   Caption         =   "補文件整批收/發文"
   ClientHeight    =   4350
   ClientLeft      =   255
   ClientTop       =   990
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   9375
   Begin VB.ListBox lstCaseNo 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Index           =   3
      Left            =   7065
      MultiSelect     =   2  '進階多重選取
      Sorted          =   -1  'True
      TabIndex        =   38
      Top             =   2910
      Width           =   1650
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6390
      Picture         =   "frm060110.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   37
      ToolTipText     =   "取消"
      Top             =   3150
      Width           =   612
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6390
      Picture         =   "frm060110.frx":0442
      Style           =   1  '圖片外觀
      TabIndex        =   36
      ToolTipText     =   "取消"
      Top             =   3600
      Width           =   612
   End
   Begin VB.ListBox lstCaseNo 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Index           =   2
      ItemData        =   "frm060110.frx":0884
      Left            =   4680
      List            =   "frm060110.frx":0886
      MultiSelect     =   2  '進階多重選取
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   2910
      Width           =   1650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "已收文合併、讓與、變更案件"
      CausesValidation=   0   'False
      Height          =   285
      Left            =   315
      TabIndex        =   32
      Top             =   840
      Width           =   2730
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6390
      Picture         =   "frm060110.frx":0888
      Style           =   1  '圖片外觀
      TabIndex        =   29
      ToolTipText     =   "取消"
      Top             =   2220
      Width           =   612
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6390
      Picture         =   "frm060110.frx":0CCA
      Style           =   1  '圖片外觀
      TabIndex        =   28
      ToolTipText     =   "取消"
      Top             =   1770
      Width           =   612
   End
   Begin VB.ListBox lstCaseNo 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Index           =   1
      Left            =   7065
      MultiSelect     =   2  '進階多重選取
      Sorted          =   -1  'True
      TabIndex        =   26
      Top             =   1380
      Width           =   1650
   End
   Begin VB.ListBox lstCaseNo 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Index           =   0
      Left            =   4680
      MultiSelect     =   2  '進階多重選取
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   1380
      Width           =   1650
   End
   Begin VB.TextBox txtAppNo 
      Height          =   270
      Left            =   1980
      TabIndex        =   23
      Top             =   2490
      Width           =   1665
   End
   Begin VB.TextBox txtCP27 
      Height          =   270
      Left            =   1305
      MaxLength       =   6
      TabIndex        =   2
      Top             =   2130
      Width           =   1125
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   285
      Left            =   2385
      TabIndex        =   1
      Top             =   540
      Width           =   660
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請書(&F)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   3
      Left            =   4185
      TabIndex        =   14
      Top             =   60
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發文室送件清單(&L)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   5
      Left            =   315
      TabIndex        =   13
      Top             =   60
      Width           =   1800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "整批發文(&D)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   2
      Left            =   5760
      TabIndex        =   12
      Top             =   60
      Width           =   1425
   End
   Begin VB.TextBox txtLR01 
      Height          =   270
      Left            =   1260
      MaxLength       =   9
      TabIndex        =   0
      Top             =   540
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件清單(P)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   4
      Left            =   2925
      TabIndex        =   5
      Top             =   60
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   8505
      TabIndex        =   4
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "整批收文(&R)"
      Enabled         =   0   'False
      Height          =   400
      Index           =   1
      Left            =   7155
      TabIndex        =   3
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "本次不收文"
      Height          =   975
      Left            =   8775
      TabIndex        =   39
      Top             =   3030
      Width           =   255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "本次要收文"
      Height          =   975
      Left            =   4410
      TabIndex        =   35
      Top             =   3180
      Width           =   255
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "已收文合併、讓與、變更案件:"
      Height          =   180
      Left            =   2205
      TabIndex        =   34
      Top             =   2910
      Width           =   2385
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "本次要收文"
      Height          =   975
      Left            =   4410
      TabIndex        =   31
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "本次不收文"
      Height          =   975
      Left            =   8775
      TabIndex        =   30
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "多申請人案件:"
      Height          =   180
      Left            =   3510
      TabIndex        =   27
      Top             =   1380
      Width           =   1125
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "總委任書申請案號:"
      Height          =   180
      Left            =   315
      TabIndex        =   24
      Top             =   2490
      Width           =   1485
   End
   Begin VB.Label lblCustName 
      Height          =   180
      Index           =   2
      Left            =   3510
      TabIndex        =   22
      Top             =   1080
      Width           =   4275
   End
   Begin VB.Label lblCustName 
      Height          =   180
      Index           =   1
      Left            =   3510
      TabIndex        =   21
      Top             =   840
      Width           =   4275
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(日):"
      Height          =   180
      Index           =   0
      Left            =   3105
      TabIndex        =   20
      Top             =   1080
      Width           =   345
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   3105
      TabIndex        =   19
      Top             =   840
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Left            =   3105
      TabIndex        =   18
      Top             =   585
      Width           =   345
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   315
      TabIndex        =   17
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label lblAppQty3 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      Height          =   180
      Left            =   2115
      TabIndex        =   16
      Top             =   1860
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "已發文案件數:"
      Height          =   180
      Left            =   315
      TabIndex        =   15
      Top             =   1875
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "已收文未發文案件數:"
      Height          =   180
      Index           =   1
      Left            =   315
      TabIndex        =   11
      Top             =   1620
      Width           =   1665
   End
   Begin VB.Label lblAppQty2 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      Height          =   180
      Left            =   2115
      TabIndex        =   10
      Top             =   1605
      Width           =   945
   End
   Begin VB.Label lblCustName 
      Height          =   180
      Index           =   0
      Left            =   3510
      TabIndex        =   9
      Top             =   585
      Width           =   4275
   End
   Begin VB.Label lblAppQty 
      Alignment       =   1  '靠右對齊
      Caption         =   "0"
      Height          =   180
      Left            =   2115
      TabIndex        =   8
      Top             =   1380
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "未收文案件數:"
      Height          =   180
      Left            =   315
      TabIndex        =   7
      Top             =   1365
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號:"
      Height          =   180
      Left            =   315
      TabIndex        =   6
      Top             =   585
      Width           =   765
   End
End
Attribute VB_Name = "frm060110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan2010/8/13 日期欄已修改
'Add by Morgan 2007/8/14
Option Explicit

Dim rsData As New ADODB.Recordset
Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim m_stSubTitle As String
Dim m_strRefDate As String '報表上所謂的當日
Dim m_iChoose As Integer '收文選擇 0:全部 1:已收合併、讓與、變更者

Private Sub cmdok_Click(Index As Integer)
   Dim strAppDate As String
   Screen.MousePointer = vbHourglass
   '結束
   If Index = 0 Then
      Unload Me
   Else
      Select Case Index
         Case 1 '整批收文
            If TxtValidate = True Then
               If m_iChoose = 1 Then
                  Set rsData = GetRst(7, intI)
               Else
                  Set rsData = GetRst(1, intI)
               End If
               If intI = 1 Then
                  If FormSave = False Then
                     MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
                  Else
                     m_strRefDate = strSrvDate(1)
                     PrintCaseList
                     If m_iChoose = 1 Then
                        Command1_Click
                     Else
                        cmdSearch_Click
                     End If
                  End If
               End If
            End If
         
         Case 2 '整批發文
            If TxtValidate(2) = True Then
               Set rsData = GetRst(5, intI)
               If intI = 1 Then
                  If FormSave1 = True Then
                     m_strRefDate = DBDATE(txtCP27)
                     PrintCaseList
                     cmdSearch_Click
                  End If
               End If
            End If
            
         Case 3 '申請書
            If TxtValidate(3) = True Then
               PrintAppForm
            End If
            
         Case 4 '案件清單
            If txtCP27 <> "" Then
               m_strRefDate = DBDATE(txtCP27)
            Else
               m_strRefDate = strSrvDate(1)
            End If
            PrintCaseList
         
         Case 5 '送件清單
            strAppDate = InputBox("發文日：", "請輸入發文日期", strSrvDate(2), Me.Left, Me.Top + Me.Height + 1000)
            If strAppDate <> "" Then
               If ChkDate(strAppDate) = True Then
                  PrintAppList strAppDate
               End If
            End If
            
      End Select
   End If
   Screen.MousePointer = vbDefault
End Sub
'設定客戶名稱
Private Sub setCustName()
   strExc(0) = "select cu04 n1,cu05||' '||cu88||' '||cu89||' '||cu90 n2,cu06 n3 from customer where cu01='" & Left(txtLR01, 8) & "' and cu02='" & Mid(txtLR01, 9, 1) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      lblCustName(0) = "" & RsTemp.Fields(0)
      lblCustName(1) = "" & RsTemp.Fields(1)
      lblCustName(2) = "" & RsTemp.Fields(2)
   Else
      lblCustName(0) = ""
      lblCustName(1) = ""
      lblCustName(2) = ""
   End If
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   If txtLR01 = "" Then
      MsgBox "客戶編號不可空白！"
      txtLR01.SetFocus
   Else
      txtLR01 = Left(txtLR01 & "000", 9)
      setCustName
      txtLR01.Tag = txtLR01
      SetCaseQty
      SetEnable True
   End If
   m_iChoose = 0
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
   Screen.MousePointer = vbHourglass
   If txtLR01 = "" Then
      MsgBox "客戶編號不可空白！"
      txtLR01.SetFocus
   Else
      txtLR01 = Left(txtLR01 & "000", 9)
      setCustName
      txtLR01.Tag = txtLR01
      SetCaseQty 1
      SetEnable True
   End If
   m_iChoose = 1
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
   intI = 0
   Do While intI < lstCaseNo(1).ListCount
      If lstCaseNo(1).Selected(intI) = True Then
         lstCaseNo(0).AddItem lstCaseNo(1).List(intI)
         lstCaseNo(1).RemoveItem intI
      Else
         intI = intI + 1
      End If
   Loop
End Sub

Private Sub Command3_Click()
   intI = 0
   Do While intI < lstCaseNo(0).ListCount
      If lstCaseNo(0).Selected(intI) = True Then
         lstCaseNo(1).AddItem lstCaseNo(0).List(intI)
         lstCaseNo(0).RemoveItem intI
      Else
         intI = intI + 1
      End If
   Loop
End Sub

Private Sub Command4_Click()
   intI = 0
   Do While intI < lstCaseNo(3).ListCount
      If lstCaseNo(3).Selected(intI) = True Then
         lstCaseNo(2).AddItem lstCaseNo(3).List(intI)
         lstCaseNo(3).RemoveItem intI
      Else
         intI = intI + 1
      End If
   Loop
End Sub

Private Sub Command5_Click()
   intI = 0
   Do While intI < lstCaseNo(2).ListCount
      If lstCaseNo(2).Selected(intI) = True Then
         lstCaseNo(3).AddItem lstCaseNo(2).List(intI)
         lstCaseNo(2).RemoveItem intI
      Else
         intI = intI + 1
      End If
   Loop
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   FormClear
   If CheckUse("frm060110", strPrint, False) = True Then
      cmdOK(5).Enabled = True
   Else
      cmdOK(5).Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060110 = Nothing
End Sub

Private Sub txtCP27_GotFocus()
   TextInverse txtCP27
End Sub

Private Sub txtCP27_Validate(Cancel As Boolean)
   If txtCP27 <> "" Then
      If Not ChkDate(txtCP27) Then
        Cancel = True
      End If
   End If
End Sub

Private Sub txtLR01_Change()
   If txtLR01.Tag <> "" Then
      FormClear
   End If
End Sub

Private Sub txtLR01_GotFocus()
   TextInverse txtLR01
End Sub

Private Sub txtLR01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Function TxtValidate(Optional p_iAct As Integer = 1) As Boolean
   Dim bCancel As Boolean
   If txtLR01 = "" Then
      MsgBox "客戶編號不可空白！", vbExclamation
      txtLR01.SetFocus
      Exit Function
   End If
   If Len(txtLR01) <> 6 And Len(txtLR01) <> 9 Then
      MsgBox "客戶編號錯誤！", vbExclamation
      txtLR01_GotFocus
      txtLR01.SetFocus
      Exit Function
   End If
   If txtLR01.Tag <> txtLR01 Then
      MsgBox "客戶編號已變更，請重新按查詢鈕！", vbExclamation
      cmdSearch.SetFocus
      Exit Function
   End If
   '發文
   If p_iAct = 2 Then
      If txtCP27 = "" Then
         MsgBox "請輸入發文日！", vbExclamation
         txtCP27.SetFocus
         Exit Function
      End If
      txtCP27_Validate bCancel
      If bCancel = True Then
         txtCP27_GotFocus
         txtCP27.SetFocus
         Exit Function
      End If

   End If
   '申請書
   If p_iAct = 3 Then
      If txtAppNo = "" Then
         If MsgBox("是否確定沒有總委任書申請案號？", vbYesNo + vbDefaultButton2) = vbNo Then
            txtAppNo.SetFocus
            Exit Function
         End If
      Else
         strExc(0) = "select pa26,pa27,pa28,pa29,pa30 from patent where pa11='" & txtAppNo & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For intI = 0 To 4
               If "" & RsTemp.Fields(intI) = txtLR01 Then
                  Exit For
               End If
            Next
            If intI = 5 Then
               MsgBox "總委任書申請案號並非該客戶案件！", vbExclamation
               txtAppNo_GotFocus
               txtAppNo.SetFocus
               Exit Function
            End If
         Else
            MsgBox "總委任書申請案號不存在！", vbExclamation
            txtAppNo_GotFocus
            txtAppNo.SetFocus
            Exit Function
         End If
      End If
   End If
   
   TxtValidate = True
End Function

Private Function checkList(p_strItem As String) As Boolean
   For intI = 0 To lstCaseNo(1).ListCount - 1
      If lstCaseNo(1).List(intI) = p_strItem Then
         checkList = True
         lstCaseNo(1).RemoveItem intI
         Exit For
      End If
   Next
   
   For intI = 0 To lstCaseNo(3).ListCount - 1
      If lstCaseNo(3).List(intI) = p_strItem Then
         checkList = True
         lstCaseNo(3).RemoveItem intI
         Exit For
      End If
   Next
End Function

Private Function FormSave() As Boolean
   
   Dim cp(1 To 110) As String
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd

   With rsData
      .MoveFirst
      Do While Not .EOF
         cp(1) = .Fields("pa01")
         cp(2) = .Fields("pa02")
         cp(3) = .Fields("pa03")
         cp(4) = .Fields("pa04")
         'Modify by Morgan 2007/10/3 判斷是否為不收文案件
         If checkList(cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)) = False Then
            cp(5) = strSrvDate(1)
            cp(6) = .Fields("np08")
            cp(7) = .Fields("np09")
            cp(9) = AutoNo("B", 6)
            cp(10) = "202"
            cp(13) = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
            cp(12) = GetSalesArea(cp(13))
            cp(14) = strUserNum
            cp(20) = "N"
            cp(26) = "N"
            cp(43) = .Fields("cp09")
            cp(64) = "" & .Fields("np15")
            cp(110) = "" & .Fields("cp110")
            
            strSql = "insert into caseprogress(CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP43,CP64,CP110)" & _
               " VALUES(" & CNULL(cp(1)) & "," & CNULL(cp(2)) & "," & CNULL(cp(3)) & "," & CNULL(cp(4)) & "," & CNULL(cp(5), True) & _
               "," & CNULL(cp(6), True) & "," & CNULL(cp(7), True) & "," & CNULL(cp(9)) & "," & CNULL(cp(10)) & "," & CNULL(cp(12)) & _
               "," & CNULL(cp(13)) & "," & CNULL(cp(14)) & "," & CNULL(cp(20)) & "," & CNULL(cp(26)) & "," & CNULL(cp(43)) & "," & CNULL(ChgSQL(cp(64))) & "," & CNULL(cp(110)) & ")"
               
            cnnConnection.Execute strSql, intI
            
            strSql = "update nextprogress set np06='Y' where np01='" & cp(43) & "' and np06 is null and np07='202'"
            cnnConnection.Execute strSql, intI
            
         End If
         .MoveNext
      Loop
   End With
   cnnConnection.CommitTrans
   FormSave = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
End Function

Private Function FormSave1() As Boolean
   Dim lMax As Long
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd
   
   With rsData
      .MoveFirst
      Do While Not .EOF
         If IsNull(.Fields("cp27")) Then
            strSql = "Update CaseProgress Set CP27=" & DBDATE(txtCP27) & _
               " where cp09='" & .Fields("cp09") & "' and cp10='202' and cp57 is null and cp27 is null"
            cnnConnection.Execute strSql, intI
         Else
            strSql = "Update CaseProgress Set CP27=null" & _
               " where cp09='" & .Fields("cp09") & "' and cp10='202' and cp57 is null and cp27=19221111"
            cnnConnection.Execute strSql, intI
         End If
         .MoveNext
      Loop
   End With
   cnnConnection.CommitTrans
   FormSave1 = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Sub FormClear()
   txtLR01.Tag = ""
   lblCustName(0) = ""
   lblCustName(1) = ""
   lblCustName(2) = ""
   lblAppQty = ""
   lblAppQty2 = ""
   lblAppQty3 = ""
   lstCaseNo(0).Clear
   lstCaseNo(1).Clear
   SetEnable False
End Sub

Private Sub SetEnable(Optional p_bolEnable As Boolean = True)
   Dim ii As Integer
   For ii = 1 To 4
      cmdOK(ii).Enabled = False
   Next
   If p_bolEnable = True Then
      '有未收文
      If Val(lblAppQty) > 0 Then
         cmdOK(1).Enabled = True '整批收文
      End If
      '有未發文
      If Val(lblAppQty2) > 0 Then
         cmdOK(2).Enabled = True '整批發文
         cmdOK(3).Enabled = True '申請書
      End If
      'Modify by Morgan 2007/10/3 只要有案件就可以印
      'If Val(lblAppQty) = 0 And (Val(lblAppQty2) > 0 Or Val(lblAppQty3) > 0) Then
      If Val(lblAppQty) > 0 Or Val(lblAppQty2) > 0 Or Val(lblAppQty3) > 0 Then
         cmdOK(4).Enabled = True '案件清單
      End If
   End If
End Sub

Private Function GetRst(p_iType As Integer, p_iRlt As Integer) As ADODB.Recordset
   Dim stCustNo As String
   stCustNo = Left(txtLR01 & "000", 9)
   Select Case p_iType
      Case 1 '未收文
         strExc(0) = "select pa01,pa02,pa03,pa04,cp09,cp110,np08,np09,NP15 from patent,caseprogress,nextprogress" & _
            " where pa01='FCP' and instr(pa26||pa27||pa28||pa29||pa30,'" & stCustNo & "')>0" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='928' and cp27>20070000" & _
            " and np01(+)=cp09 and np07='202' and np06 is null order by 1,2,3,4"
      
      Case 2 '已收文未發文
         strExc(0) = "select b.cp27,a.cp09,pa11 from patent,caseprogress a,caseprogress b" & _
            " where pa01='FCP' and instr(pa26||pa27||pa28||pa29||pa30,'" & stCustNo & "')>0" & _
            " and a.cp01(+)=pa01 and a.cp02(+)=pa02 and a.cp03(+)=pa03 and a.cp04(+)=pa04 and a.cp10='202' and a.cp27 is null" & _
            " and b.cp09(+)=a.cp43 and b.cp10='928'"
      
      Case 3 '已發文
         strExc(0) = "select pa01,pa02,pa03,pa04 from patent,caseprogress a" & _
            " where pa01='FCP' and instr(pa26||pa27||pa28||pa29||pa30,'" & stCustNo & "')>0" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='202' and cp27>20070000" & _
            " and exists(select * from caseprogress b where b.cp09=a.cp43 and b.cp10='928')"
      
      Case 4 '案件清單
         strExc(0) = "select decode(cp27,null,2,1) C1,substr(cp09,1,1) C2,decode(pa16,'1',1,2) C3,cp05,s1.st02 S1,cp27,pa75,na03" & _
            ",cp01||'-'||cp02||decode(cp03||cp04,'000','',cp03||cp04) C4,decode(pa27,null,'','Y') X1,s2.st02 S2" & _
            " from patent,caseprogress a,staff s1,fagent,nation,staff s2" & _
            " where pa01='FCP' and instr(pa26||pa27||pa28||pa29||pa30,'" & stCustNo & "')>0" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='202'" & _
            " and exists(select * from caseprogress b where b.cp09=a.cp43 and b.cp10='928')" & _
            " and s1.st01(+)=cp14 and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1) and na01(+)=fa10" & _
            " and s2.st01(+)=na16"
            
         'Add by Morgan 2007/10/3 加未收文
         strExc(0) = strExc(0) & " union" & _
            " select 3 C1,Null C2,0 C3,0 cp05,null S1,0 cp27,pa75,na03" & _
            ",cp01||'-'||cp02||decode(cp03||cp04,'000','',cp03||cp04) C4,decode(pa27,null,'','Y') X1,s2.st02 S2" & _
            " from patent,caseprogress,nextprogress,fagent,nation,staff s2" & _
            " where pa01='FCP' and instr(pa26||pa27||pa28||pa29||pa30,'" & stCustNo & "')>0" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='928' and cp27>20070000" & _
            " and np01(+)=cp09 and np07='202' and np06 is null" & _
            " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9,1) and na01(+)=fa10" & _
            " and s2.st01(+)=na16"

         strExc(0) = strExc(0) & " order by C1,C2,C3,C4"
      Case 5 '整批發文
         strExc(0) = "select cp01,cp02,cp03,cp04,cp09,cp13,decode(cp27,null,2,1) C1,substr(cp09,1,1) C2,decode(pa16,'Y',1,2) C3" & _
            ",cp01||'-'||cp02||decode(cp03||cp04,'000','',cp03||cp04) C4,cp27 " & _
            " from patent,caseprogress a" & _
            " where pa01='FCP' and instr(pa26||pa27||pa28||pa29||pa30,'" & stCustNo & "')>0" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='202' and (cp27 is null or cp27=19221111) and cp57 is null" & _
            " and exists(select * from caseprogress b where b.cp09=a.cp43 and b.cp10='928')" & _
            " order by C1,C2,C3,C4"
            
      Case 6 '未收文多申請人案件
         strExc(0) = "select pa01||'-'||pa02||'-'||pa03||'-'||pa04 CNo from patent,caseprogress,nextprogress" & _
            " where pa01='FCP' and instr(pa26||pa27||pa28||pa29||pa30,'" & stCustNo & "')>0 and pa27 is not null" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='928' and cp27>20070000" & _
            " and np01(+)=cp09 and np07='202' and np06 is null order by 1"
            
      Case 7 '未收文已收合併、讓與、變更案件
         strExc(0) = "select pa01||'-'||pa02||'-'||pa03||'-'||pa04 CNo,pa01,pa02,pa03,pa04,cp09,cp110,np08,np09,NP15 from patent,caseprogress a,nextprogress" & _
            " where pa01='FCP' and instr(pa26||pa27||pa28||pa29||pa30,'" & stCustNo & "')>0" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='928' and cp27>20070000" & _
            " and exists(select * from caseprogress b where b.cp01=pa01 and b.cp02=pa02 and b.cp03=pa03 and b.cp04=pa04 and b.cp27 is null and b.cp10 in ('401','701','702'))" & _
            " and np01(+)=cp09 and np07='202' and np06 is null order by 1"
            
   End Select
   p_iRlt = 1
   Set GetRst = ClsLawReadRstMsg(p_iRlt, strExc(0))
   
End Function

'Add by Morgan 2007/10/5
Private Sub SetCombo(p_index As Integer, p_Rst As ADODB.Recordset)
   With p_Rst
   Do While Not .EOF
      lstCaseNo(p_index).AddItem .Fields(0), 0
      .MoveNext
   Loop
   End With
   
End Sub


Private Sub SetCaseQty(Optional p_Choose As Integer)
   
   lblAppQty = ""
   lblAppQty2 = ""
   lblAppQty3 = ""
   lstCaseNo(0).Clear
   lstCaseNo(1).Clear
   lstCaseNo(2).Clear
   lstCaseNo(3).Clear
   'Add by Morgan 2007/10/5 已收合併、讓與、變更案件選擇
   If p_Choose = 1 Then
      Set RsTemp = GetRst(7, intI)
      If RsTemp.RecordCount > 0 Then
         SetCombo 2, RsTemp
      End If
   Else
      Set RsTemp = GetRst(1, intI)
   End If
   If intI = 1 Then
      lblAppQty = RsTemp.RecordCount
   Else
      lblAppQty = 0
   End If
   
   Set RsTemp = GetRst(2, intI)
   If intI = 1 Then
      lblAppQty2 = RsTemp.RecordCount
   Else
      lblAppQty2 = 0
   End If
   
      Set RsTemp = GetRst(3, intI)
   If intI = 1 Then
      lblAppQty3 = RsTemp.RecordCount
   Else
      lblAppQty3 = 0
   End If
   
   'Add by Morgan 2007/10/3 多申請人案件選擇
   intI = 1
   Set RsTemp = GetRst(6, intI)
   If RsTemp.RecordCount > 0 Then
      SetCombo 0, RsTemp
   End If
   'end 2007/10/3

   
   'Add by Morgan 2007/10/5 已收合併、讓與、變更案件選擇
   If p_Choose <> 1 Then
      Set RsTemp = GetRst(7, intI)
      If RsTemp.RecordCount > 0 Then
         SetCombo 2, RsTemp
      End If
   End If
   'end 2007/10/5
   
End Sub
'申請書
Private Sub PrintAppForm()
   Set rsData = GetRst(2, intI)
   If intI = 1 Then
      With rsData
         Do While Not .EOF
            EndLetter "01", .Fields("cp09"), "01", strUserNum
            
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('01','" & .Fields("cp09") & "','01','" & strUserNum & "'" & _
               ",'重新委任發文日','" & .Fields("cp27") & "')"
            cnnConnection.Execute strSql, intI
             
            If txtAppNo <> "" Then
               If txtAppNo = "" & .Fields("pa11") Then
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('01','" & .Fields("cp09") & "','01','" & strUserNum & "'" & _
                     ",'勾選1','■')"
                  cnnConnection.Execute strSql, intI
                  
               Else
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('01','" & .Fields("cp09") & "','01','" & strUserNum & "'" & _
                     ",'總委任書申請案號','" & txtAppNo & "')"
                  cnnConnection.Execute strSql, intI
               
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('01','" & .Fields("cp09") & "','01','" & strUserNum & "'" & _
                     ",'勾選2','■')"
                  cnnConnection.Execute strSql, intI
                  
               End If
               
            End If
            
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('01','" & .Fields("cp09") & "','01','" & strUserNum & "'" & _
               ",'流水號','" & Format(.AbsolutePosition, "000000") & "')"
            cnnConnection.Execute strSql, intI
            NowPrint .Fields("cp09"), "01", "01", False, strUserNum
            .MoveNext
         Loop
         PUB_BatchPrint "5"
      End With
   End If
End Sub

'案件清單
Private Sub PrintCaseList()
   Set rsData = GetRst(4, intI)
   If intI = 1 Then
      DoPrint
   End If
End Sub

'送件清單
Private Sub PrintAppList(Optional p_AppDate As String)
   Dim strTmp As String, iRec As Integer, iRecs As Integer, iCol As Integer
   Dim lngTot As Long
   Dim iCaseNo As Integer '案件筆數
   Dim iPage As Integer '頁次
   Dim nCopys As Integer '份數
   Dim iCopys As Integer
   Dim stCon As String
   Dim strAppDate As String
   
   If p_AppDate <> "" Then
      strAppDate = DBDATE(p_AppDate)
   Else
      strAppDate = strSrvDate(1)
   End If
   
   stCon = " and cp27=" & strAppDate
   
   strExc(0) = "SELECT LPAD(CP01||'-'||CP02||'-'||CP03||'-'||CP04,15,' ') C01, 0 C02, 0 C03" & _
            ",RPAD(NVL(PA11,' '),12,' ') C04,'補文件(委任書)' C05, RPAD(NVL(CU04,' '),20,' ') C06" & _
            ",RPAD(PA05,40,' ') C07" & _
            " from caseprogress a,PATENT,customer" & _
            " where cp01='FCP' and cp10='202'" & stCon & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
            " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9,1)" & _
            " and exists(select * from caseprogress b where b.cp09=a.cp43 and b.cp10='928') order by 1"
            
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '2007/7/6 modify by sonia 張弘郁說改一份
   'nCopys = 2
   nCopys = 1
   If intI = 1 Then
      Printer.Orientation = 2
      Printer.Font = "細明體"
      For iCopys = 1 To nCopys
         If iCopys > 1 Then Printer.NewPage
         iPage = 1: iRec = 0: lngTot = 0: iCaseNo = 0: iRecs = 0
         With RsTemp
         .MoveFirst
         PrintHead strAppDate
         Do While Not .EOF
            iRec = iRec + 1: iRecs = iRecs + 1
            If iRec > 26 Then
               PrintTail iPage
               Printer.NewPage
               iPage = iPage + 1
               PrintHead strAppDate
               iRec = 0
            End If
            strTmp = ""
            For iCol = 0 To 6
               Select Case iCol
                  Case 1
                     strTmp = strTmp & Right(Space(9) & Format(Val("" & .Fields(iCol)), "###,###"), 9) & Space(1)
                  Case 2
                     strTmp = strTmp & Right(Space(8) & Format(Val("" & .Fields(iCol)), "###,###"), 8) & Space(1)
                  Case Else
                     strTmp = strTmp & .Fields(iCol) & Space(1)
               End Select
            Next
            Printer.CurrentY = Printer.CurrentY + 60
            Printer.Print strTmp
            iCaseNo = iCaseNo + 1
            .MoveNext
         Loop
         PrintTail iPage, lngTot, iRecs, iCaseNo
         End With
      Next
      Printer.EndDoc
      MsgBox "列印完成！"
   Else
      MsgBox "無可列印資料！", vbInformation
   End If
  
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = LBound(strData) To UBound(strData)
      Printer.CurrentX = PLeft(iCol)
      Printer.CurrentY = iPrint
      Printer.Print strData(iCol)
    Next
End Sub

Private Sub PrintSubTotal(p_iCount As Integer, Optional p_bSubTotal As Boolean = True, Optional p_ExCount As Integer = -1)
   PrintNewLine
   
   If iPrint - 600 > lngPageHeight Then
      PrintPageHeader
      PrintPageHeader1
   End If
   
   DrawLine
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   If p_bSubTotal = True Then
      strExc(0) = "小計： " & p_iCount & " 筆"
      If p_ExCount > -1 Then
         strExc(0) = strExc(0) & " (當日發文 " & p_ExCount & " 筆)"
      End If
      Printer.Print strExc(0)
   Else
      Printer.Print "總計： " & p_iCount & " 筆"
   End If
   iPrint = iPrint + 300
   
End Sub
Private Sub DoPrint()
   Dim Grp1 As String, Grp2 As String, Grp3 As String, iCount As Integer, iTotal As Integer, iCurCount As Integer
   Dim subTitle As String, iNo As Integer
   Dim iOrientation As Integer
   Dim strTemp() As String
   
   iOrientation = Printer.Orientation
   Printer.Orientation = 1
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With rsData
      GetPleft
      ReDim strTemp(1 To 8)
      iPage = 0
      iCount = 0
      iTotal = 0
      iNo = 0
      iCurCount = -1
      m_stSubTitle = ""
      PrintPageHeader
      Do While Not .EOF
         If Grp1 <> .Fields("C1") Or (.Fields("C1") <> 1 And (Grp2 <> .Fields("C2") Or Grp3 <> .Fields("C3"))) Then
            If iCount > 0 Then
               PrintSubTotal iCount, , iCurCount
               iTotal = iTotal + iCount
               iCount = 0
               iCurCount = -1
            End If
            If .Fields("C1") = 1 Then
               m_stSubTitle = "***已發文***"
               iCurCount = 0
            ElseIf .Fields("C1") = 3 Then
               m_stSubTitle = "***未收文***"
            Else
               If .Fields("C3") = "1" Then
                  m_stSubTitle = "***" & .Fields("C2") & "類未發文(已准)***"
               Else
                  m_stSubTitle = "***" & .Fields("C2") & "類未發文(未准)***"
               End If
            End If
            PrintPageHeader1
            Grp1 = "" & .Fields("C1")
            Grp2 = "" & .Fields("C2")
            Grp3 = "" & .Fields("C3")
         End If
         iCount = iCount + 1
         If "" & .Fields("cp27") = m_strRefDate Then
            iCurCount = iCurCount + 1
         End If
         strTemp(1) = ChangeWStringToTDateString(Format("" & .Fields("cp05"), ""))
         strTemp(2) = "" & .Fields("S1")
         strTemp(3) = ChangeWStringToTDateString(Format("" & .Fields("cp27"), ""))
         strTemp(4) = "" & .Fields("pa75")
         strTemp(5) = "" & .Fields("na03")
         strTemp(6) = "" & .Fields("S2")
         '未發文未核准的加流水號
         If .Fields("C1") = 2 And .Fields("C3") = "2" Then
            iNo = iNo + 1
            strTemp(7) = iNo & "." & .Fields("C4")
         Else
            strTemp(7) = "" & .Fields("C4")
         End If
         strTemp(8) = "" & .Fields("X1")
         PrintDetail strTemp
         .MoveNext
      Loop
      If iCount > 0 Then
         PrintSubTotal iCount, , iCurCount
         iTotal = iTotal + iCount
         PrintSubTotal iTotal, False
      End If
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
End Sub


Sub PrintPageHeader()
   Dim strPTmp As String
   Dim strName(3) As String, idx As Integer
   
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = "補文件案件清單"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   iPrint = iPrint + 100
   strPTmp = "(客戶編號：" & txtLR01 & ")"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   idx = 0
   '中
   If lblCustName(0) <> "" Then
      idx = idx + 1
      strName(idx) = Left(lblCustName(0), 20)
   End If
   '英
   If lblCustName(1) <> "" Then
      idx = idx + 1
      strName(idx) = Left(lblCustName(1), 40)
   End If
   '日
   If lblCustName(2) <> "" Then
      idx = idx + 1
      strName(idx) = Left(lblCustName(2), 20)
   End If
   
   If strName(1) <> "" Then
      PrintNewLine
      strPTmp = strName(1)
      Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPTmp
   End If
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   
   If strName(2) <> "" Then
      strPTmp = strName(2)
      Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPTmp
   End If
   
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
  
   If strName(3) <> "" Then
      strPTmp = strName(3)
      Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPTmp
   End If
   
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   iPage = iPage + 1
   Printer.Print "頁    次：" & str(iPage)
End Sub

Sub PrintPageHeader1()
   
   If iPrint + 1000 > lngPageHeight Then
      Printer.NewPage
      PrintPageHeader
   End If
   
   If m_stSubTitle <> "" Then
      PrintNewLine
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iPrint
      Printer.Print m_stSubTitle
   End If
   
   PrintNewLine
   DrawLine
   
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "收文日"
   
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "發文日"
   
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "代理人編號"
   
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "代理人國籍"
   
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "管制人"
   
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "多申請人(Y)"
   
   PrintNewLine
   DrawLine
   iPrint = iPrint - 300
End Sub

Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   ReDim PLeft(1 To 8)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + 1100
   PLeft(3) = PLeft(2) + 1100
   PLeft(4) = PLeft(3) + 1100
   PLeft(5) = PLeft(4) + 1300
   PLeft(6) = PLeft(5) + 1300
   PLeft(7) = PLeft(6) + 1100
   PLeft(8) = PLeft(7) + 1800
End Sub

Private Sub DrawLine()
   Printer.DrawStyle = vbSolid
   Printer.DrawWidth = 4
   Printer.Line (ciStartX, iPrint)-(lngPageWidth - 500, iPrint)
   iPrint = iPrint + 100
End Sub

Private Sub PrintNewLine(Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      DrawLine
      Printer.NewPage
      PrintPageHeader
      PrintPageHeader1
      iPrint = iPrint + lngLineHeight
   End If
End Sub

Private Sub PrintHead(p_Date As String)

      Dim strTitle As String
      
      strTitle = "外專 " & ChangeTStringToTDateString(TransDate(p_Date, 1)) & " 重新委任補文件送件清單"
      Printer.Print
      Printer.FontSize = 16
      Printer.CurrentX = 5000
      Printer.Print strTitle
      Printer.FontSize = 12
      Printer.CurrentX = 13000
      Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
      Printer.CurrentX = 13000
      Printer.Print "列印時間：" & Format(ServerTime, "##:##:##")
      Printer.CurrentX = 13000
      Printer.Print "列印人員：" & strUserName
      Printer.Print
      Printer.Print "本所案號        規費      規費小計 申請案號     案件性質　　   申請人               案件名稱                                "
      Printer.Print "--------------- --------- -------- ------------ -------------- -------------------- -----------------------------------------------------"
      
End Sub

Private Sub PrintTail(iPage As Integer, Optional p_lngTot As Long, Optional p_iRecs As Integer, Optional p_iCaseCnt As Integer)
   Dim stData As String
   stData = "總計" & Space(4) & Right(Space(8) & "規費 " & Format(p_lngTot, "###,###,##0") & " 元", 16) & Space(4) & Right(Space(2) & "明細 " & Format(p_iRecs) & " 筆", 9) & Space(4) & Right(Space(2) & "案號 " & Format(p_iCaseCnt) & " 筆", 9)
   Printer.FontSize = 12
   Printer.Print String(135, "-")
   Printer.Print stData
   Printer.CurrentY = 10700
   Printer.CurrentX = 7000
   Printer.Print "第 " & Format(iPage) & " 頁"
   
End Sub

Private Sub txtAppNo_GotFocus()
   TextInverse txtAppNo
End Sub

