VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090133 
   BorderStyle     =   1  '單線固定
   Caption         =   "圖形查名路徑-大分類維護"
   ClientHeight    =   4680
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7560
   Begin VB.TextBox txtDB 
      Height          =   300
      Index           =   2
      Left            =   1368
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1248
      Width           =   5880
   End
   Begin VB.TextBox txtDB 
      Height          =   270
      Index           =   1
      Left            =   1368
      MaxLength       =   2
      TabIndex        =   0
      Top             =   936
      Width           =   600
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7695
      Top             =   450
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090133.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090133.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090133.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090133.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090133.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090133.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090133.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090133.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090133.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090133.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090133.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm090133.frx":20F4
      Height          =   2868
      Left            =   456
      TabIndex        =   5
      Top             =   1680
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   5059
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "大分類號|大分類名稱"
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSForms.TextBox textCUID 
      Height          =   228
      Left            =   384
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   648
      Width           =   6264
      VariousPropertyBits=   671105055
      Size            =   "11049;402"
      Value           =   "CREATE :       UPDATE : "
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "名稱："
      Height          =   180
      Index           =   2
      Left            =   744
      TabIndex        =   4
      Top             =   1320
      Width           =   564
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "大分類號："
      Height          =   180
      Index           =   1
      Left            =   408
      TabIndex        =   3
      Top             =   984
      Width           =   900
   End
End
Attribute VB_Name = "frm090133"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2025/11/05 調整欄位:將TMR103(CREATE ID),TMR104(CREATE DATE),TMR105(UPDATE ID),TMR106(UPDATE DATE)=>整合成TMR103,TMR104,TMR105(UPDATE ID+DATE+TIME)
'Memo by Lydia 2024/07/17 改成Form2.0 ; textCUID
'Created by Lydia 2024/07/17
Option Explicit
Dim intLastRow As Integer, intCols As Integer

Dim m_EditMode As Integer '0:瀏覽 1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_StrCon As String 'Form_Load預設抓資料的語法
Dim oText As TextBox

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Screen.MousePointer = vbHourglass
   Select Case KeyCode
      Case vbKeyF2 '新增
         KeyCode = 0: Action 1
      Case vbKeyF3 '修改
         KeyCode = 0: Action 2
      Case vbKeyF4: '查詢
         KeyCode = 0: Action 4
      Case vbKeyF5 '刪除
         KeyCode = 0: Action 3
      Case vbKeyHome '第一筆
         KeyCode = 0: Action 6
      Case vbKeyPageUp '上一筆
         KeyCode = 0: Action 7
      Case vbKeyPageDown '下一筆
         KeyCode = 0: Action 8
      Case vbKeyEnd: '最後筆
         KeyCode = 0: Action 9
      Case vbKeyF9, vbKeyReturn '確定
         KeyCode = 0: Action 11
      Case vbKeyF10 '取消
         KeyCode = 0: Action 12
      Case vbKeyEscape '結束
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            KeyCode = 0: Action 14
         End If
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction("frm090133", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm090133", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm090133", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm090133", strFind, False)
  
   MoveFormToCenter Me
   'Modified by Lydia 2025/11/05 調整欄位
   'm_StrCon = "SELECT TMR101,TMR102,NVL(S1.ST02,TMR103) AS TMR103, SQLDATET(TO_CHAR(TMR104,'YYYYMMDD')) CDATE,SUBSTR(SQLTIME6(TO_CHAR(TMR104,'HH24MISS')),1,5) CTIME," & _
           "NVL(S2.ST02,TMR105) AS TMR105 ,SQLDATET(TO_CHAR(TMR106,'YYYYMMDD')) UDATE,SUBSTR(SQLTIME6(TO_CHAR(TMR106,'HH24MISS')),1,5) UTIME " & _
           "FROM TMQAPPR1,STAFF S1, STAFF S2 WHERE TMR103=S1.ST01(+) AND TMR105=S2.ST01(+) "
   m_StrCon = "SELECT TMR101,TMR102,NVL(S1.ST02,TMR103) AS TMR103, SQLDATET(TMR104) UDATE,SUBSTR(SQLTIME6(TMR105||'00'),1,5) UTIME " & _
           "FROM TMQAPPR1,STAFF S1 WHERE TMR103=S1.ST01(+) "
   textCUID.BackColor = &H8000000F
   Action 6 '預設第一筆
   UpdateToolbarState
   
   ReadData
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090133 = Nothing
End Sub

Private Sub Grd1_Click()
Dim TmpRow As Integer

If m_EditMode = 0 Then
    GridClick GRD1, intLastRow, 7
    
    '帶入textbox
    If GRD1.TextMatrix(intLastRow, 0) <> "" Then
       SetData GRD1.TextMatrix(intLastRow, 0)
    End If
End If

End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
   Action Button.Index
   Screen.MousePointer = vbDefault
End Sub
'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(6).Enabled = True
            TBar1.Buttons(7).Enabled = True
            TBar1.Buttons(8).Enabled = True
            TBar1.Buttons(9).Enabled = True
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
      
      Case 1, 2, 3, 4 '維護
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
End Sub

Private Sub TxtLock()
   Select Case m_EditMode
   Case 0 '瀏覽
      For Each oText In txtDB
         oText.Locked = True
      Next
   Case Else
      For Each oText In txtDB
         oText.Locked = False
         oText.Tag = oText.Text
      Next
      If m_EditMode <> 4 Then
         If m_EditMode = 1 Then
            txtDB(1).SetFocus
            txtDB_GotFocus 1
         Else
            txtDB(1).Locked = True
            txtDB(2).SetFocus
            txtDB_GotFocus 2
         End If
      End If

   End Select
End Sub
Private Sub Action(Index As Integer)
   
   If TBar1.Buttons(Index).Enabled = False Then Exit Sub

On Error GoTo ErrHand

   Select Case Index
      Case 1 '按下新增
        m_EditMode = 1
        FormReset
      Case 2 '按下修改
         If txtDB(1).Text = "" Then
             MsgBox "請先選擇資料!!!", vbExclamation + vbOKOnly
             Exit Sub
         Else
             m_EditMode = 2
         End If
      Case 3 '按下刪除
         If txtDB(1).Text = "" Then
             MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
             Exit Sub
         End If

         If DelMsg() = True Then
            If FormDelete() = False Then
               MsgBox "刪除失敗!", vbCritical
               Exit Sub
            Else
               ReadData '更新GRD1
            End If
         End If

      Case 4 '按下查詢
         FormReset
         m_EditMode = 4
         
      Case 6 '第一筆
         ShowRecord 0
      Case 7 '前一筆
         If txtDB(1) <> "" Then
            ShowRecord 1
         Else
            m_EditMode = -1
         End If
      Case 8 '後一筆
         If txtDB(1) <> "" Then
            ShowRecord 2
         Else
            m_EditMode = -1
         End If
      Case 9 '最後筆
         ShowRecord 3
      Case 11 '按下確定
         Select Case m_EditMode
            '新增,修改
            Case 1, 2
               If TxtValidate = False Then
                  Exit Sub
               Else
                  If m_EditMode = 1 Then
                     If RecIsExist(True, txtDB(1)) = True Then Exit Sub
                  End If
                  If FormSave() = False Then
                     MsgBox "存檔失敗!", vbCritical
                     Exit Sub
                  Else
                     m_EditMode = 0
                     ReadData
                  End If
               End If
            '查詢
            Case 4
               If RecIsExist(False, IIf(Trim(txtDB(1)) = "", "00", Trim(txtDB(1)))) = False Then
                  MsgBox "無資料!", vbExclamation
                  Exit Sub
               Else
                     '清除反白列
                    If intLastRow > 0 Then
                       If GRD1.CellBackColor <> GRD1.BackColor Then
                         GridClick GRD1, intLastRow, 7
                       End If
                    End If
                  m_EditMode = 0
                  SetData txtDB(1)
               End If
         End Select
      Case 12 '按下取消
         m_EditMode = 0
         txtDB(1) = txtDB(1).Tag
         txtDB(2) = txtDB(2).Tag
         If txtDB(1) <> "" Then
            If RecIsExist(False, txtDB(1)) = False Then
               ShowRecord 3
            End If
         End If
      Case 14 '結束
         Unload Me
         Exit Sub
   End Select
   
   If m_EditMode < 0 Then
      m_EditMode = 0
   Else
      UpdateToolbarState
      TxtLock
   End If
   
   Exit Sub
   
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

' 顯示資料
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
 Dim stKey As String
 Dim mDiff As String
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case p_iWay
      Case 0 '第一筆
         strExc(0) = m_StrCon & " order by TMR101 "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            DataErrorMessage 8
         End If
         mDiff = "MIN"
      Case 1 '前一筆
         strExc(0) = m_StrCon & " and TMR101<" & txtDB(1) & " order by TMR101 desc"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            DataErrorMessage 6
         End If
         mDiff = "-1"
      Case 2 '後一筆
         strExc(0) = m_StrCon & " and TMR101>" & txtDB(1) & " order by TMR101 "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            DataErrorMessage 7
         End If
         mDiff = "+1"
      Case 3 '最後筆
         strExc(0) = m_StrCon & " order by TMR101 DESC "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            DataErrorMessage 8
         End If
         mDiff = "MAX"
   End Select
   
         If intI = 1 Then
            txtDB(1) = RsTemp.Fields("TMR101")
            txtDB(2) = RsTemp.Fields("TMR102")
            txtDB(1).Tag = RsTemp.Fields("TMR101")
            txtDB(2).Tag = RsTemp.Fields("TMR102")
            ShowRecord = True
            UpdateCUID RsTemp
         Else
            mDiff = ""
         End If
   
   Screen.MousePointer = vbDefault
   
   '功能鍵可移動反白列
   If intLastRow > 0 And mDiff <> "" Then
       GridClick GRD1, intLastRow, 7
          Select Case mDiff
              Case "MIN"
                 GRD1.row = 1
              Case "-1"
                 GRD1.row = intLastRow - 1
              Case "+1"
                 GRD1.row = intLastRow + 1
              Case "MAX"
                 GRD1.row = GRD1.Rows - 1
          End Select
       GridClick GRD1, intLastRow, 7
   End If
   
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Private Function ReadData() As Boolean
   
   Dim stCon As String
   
   FormReset
   strExc(0) = "SELECT TMR101,TMR102 FROM (" & m_StrCon & ") order by TMR101"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
     Set GRD1.Recordset = RsTemp.Clone
     GRD1.FormatString = "大分類號|大分類名稱"
     GRD1.ColWidth(0) = 960
     GRD1.CellBackColor = &H80000005
     GRD1.ColWidth(1) = 4000
     ReadData = True
   End If
End Function

Private Sub SetData(ByVal p01 As String)
   Dim rsA As New ADODB.Recordset
   Dim strA1 As String, intA As Integer
   
   strA1 = m_StrCon & " and TMR101=" & CNULL(p01)
   If rsA.State <> adStateClosed Then rsA.Close
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, strA1)
   
   With rsA
     For Each oText In txtDB
        oText = "" & .Fields("TMR1" & Format(oText.Index, "00"))
     Next
   End With
   UpdateCUID rsA
   
   txtDB(1).Tag = p01
   txtDB(2).Tag = txtDB(2)
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   strCName = "" & rsSrcTmp.Fields("TMR103")
   'Mark by Lydia 2025/11/05
   'strCDate = "" & rsSrcTmp.Fields("CDATE")
   'strCTime = "" & rsSrcTmp.Fields("CTIME")
   'strUName = "" & rsSrcTmp.Fields("TMR105")
   'end 2025/11/05
   strUDate = "" & rsSrcTmp.Fields("UDATE")
   strUTime = "" & rsSrcTmp.Fields("UTIME")
  
   ' 設定CUID中的文字
   'Modified by Lydia 2025/11/05
   'textCUID.Text = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & " " & String(5, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
   textCUID.Text = "UPDATE : " & strCName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Sub FormReset()
   Dim oText As TextBox
   Dim oLabel As LABEL
   
   For Each oText In txtDB
      oText.Text = ""
   Next

   textCUID = ""
   
     '清除反白列
    If intLastRow > 0 Then
       If GRD1.CellBackColor <> GRD1.BackColor Then
         GridClick GRD1, intLastRow, 7
       End If
    End If
         
End Sub

Private Sub txtDB_GotFocus(Index As Integer)
   TextInverse txtDB(Index)
End Sub

Private Sub txtDB_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDB_Validate(Index As Integer, Cancel As Boolean)
   Dim strCusTemp As String, strTemp As String
   Select Case Index
   Case 1 '大分類號
      If txtDB(Index) <> "" And m_EditMode = 1 Then
         If Trim(txtDB(Index)) <> "" Then
            If Len(txtDB(Index)) = 2 And txtDB(Index) <> "00" Then
               If RecIsExist(True, txtDB(1)) = True Then
                  If m_EditMode <> 0 Then Cancel = True
               End If
            Else
               MsgBox "請輸入2碼大分類號！"
               Cancel = True
            End If
         End If
      End If
   End Select
   
   If Not CheckLengthIsOK(txtDB(Index), txtDB(Index).MaxLength) Then
      If m_EditMode <> 0 Then Cancel = True
   End If
   
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean, idx As Integer
   
   For Each oText In txtDB
      idx = oText.Index
      If Trim(oText.Text) = "" Then
         MsgBox Replace(Label1(oText.Index), "：", "") & "不可空白！", vbExclamation
         txtDB(idx).SetFocus
         txtDB_GotFocus idx
         Exit Function
      End If
      txtDB_Validate idx, bCancel
      If bCancel = True Then
         txtDB(idx).SetFocus
         txtDB_GotFocus idx
         Exit Function
      End If
   Next
   
   TxtValidate = True
End Function

Private Function FormSave() As Boolean
On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
   If m_EditMode = 1 Then
      'Modified by Lydia 2025/11/05
      'strSql = "insert into TMQAppR1(TMR101,TMR102,TMR103,TMR104) values (" & CNULL(txtDB(1)) & "," & CNULL(ChgSQL(txtDB(2))) & "," & CNULL(strUserNum) & ", sysdate) "
      strSql = "insert into TMQAppR1(TMR101,TMR102,TMR103,TMR104,TMR105) values (" & CNULL(txtDB(1)) & "," & CNULL(ChgSQL(txtDB(2))) & "," & CNULL(strUserNum) & ", TO_CHAR(SYSDATE,'YYYYMMDD'),SUBSTR(TO_CHAR(SYSDATE,'HH24MISS'),1,4)) "
   Else
      'Modified by Lydia 2025/11/05
      'strSql = "update TMQAppR1 set TMR102=" & CNULL(ChgSQL(txtDB(2))) & ",TMR105=" & CNULL(strUserNum) & ", TMR106=sysdate where TMR101=" & txtDB(1)
      strSql = "update TMQAppR1 set TMR102=" & CNULL(ChgSQL(txtDB(2))) & ",TMR103=" & CNULL(strUserNum) & ", TMR104=TO_CHAR(SYSDATE,'YYYYMMDD'), TMR105=SUBSTR(TO_CHAR(SYSDATE,'HH24MISS'),1,4) where TMR101=" & CNULL(txtDB(1))
   End If
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

Private Function FormDelete() As Boolean
On Error GoTo ErrHnd

   strExc(0) = "select * from TMQAppR2 Where TMR201='" & txtDB(1) & "'  "
   intI = 1: strExc(1) = ""
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If MsgBox(txtDB(1) & " " & txtDB(2) & vbCrLf & "已有中分類設定，確定要刪除？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
         strExc(1) = "N"
      Else
         strExc(1) = "Y"
      End If
   End If
   If strExc(1) <> "N" Then
      cnnConnection.BeginTrans
      strSql = "delete from TMQAppR1 where TMR101='" & txtDB(1) & "' "
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
      strSql = "delete from TMQAppR2 where TMR201='" & txtDB(1) & "' "
      cnnConnection.Execute strSql, intI
      strSql = "Delete from TMQAppR3 where TMR301='" & txtDB(1) & "' "
      cnnConnection.Execute strSql, intI
      cnnConnection.CommitTrans
      FormDelete = True
   End If
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

Private Function RecIsExist(Optional ByVal bMsg As Boolean = True, Optional ByVal pKey01 As String) As Boolean
   Dim iR As Integer
   Dim rsQa As ADODB.Recordset
   
strExc(0) = ""

If Trim(pKey01) <> "" Then
   strExc(0) = strExc(0) & "and TMR101='" & Trim(pKey01) & "' "
End If

If Left(strExc(0), 3) = "and" Then strExc(0) = Mid(strExc(0), 4, Len(strExc(0)) - 4)

   strExc(1) = " select * from TMQAppR1 where " & strExc(0) & " order by 1"
   iR = 1
   Set rsQa = ClsLawReadRstMsg(iR, strExc(1))
   If iR = 1 Then
      RecIsExist = True
      If bMsg = True Then MsgBox "大分類已存在!!", vbCritical
   Else
      RecIsExist = False
   End If
   Set rsQa = Nothing
   
End Function
