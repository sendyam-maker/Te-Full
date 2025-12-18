VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090120 
   BorderStyle     =   1  '單線固定
   Caption         =   "刪除組群維護"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7350
   Begin VB.TextBox textCD01 
      Height          =   285
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   0
      Top             =   720
      Width           =   930
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   3735
      Left            =   60
      TabIndex        =   2
      Top             =   1410
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      HighLight       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5880
      Top             =   690
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
            Picture         =   "frm090120.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090120.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090120.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090120.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090120.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090120.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090120.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090120.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090120.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090120.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090120.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
   Begin MSForms.TextBox textCD02 
      Height          =   300
      Left            =   1110
      TabIndex        =   5
      Top             =   1080
      Width           =   2190
      VariousPropertyBits=   671105055
      BackColor       =   -2147483633
      Size            =   "3863;529"
      Value           =   "txtFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "維護人員："
      Height          =   180
      Left            =   150
      TabIndex        =   4
      Top             =   1110
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "組群："
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   750
      Width           =   540
   End
End
Attribute VB_Name = "frm090120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/11 改成Form2.0 ; textCD02
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
Option Explicit

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_EditMode As Integer
Dim m_SubMode As Integer
' 第一筆資料的key
Dim m_FirstKEY As String
' 最後一筆資料的key
Dim m_LastKEY As String
' 目前正在顯示的key
Dim m_CurrKEY As String
Dim m_iPreRow As Integer '前次顯示資料列
Dim m_bGridChange As Boolean


Private Sub Form_Initialize()
ReDim m_FieldList(4) As FIELDITEM
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
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
            'PrintData
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
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
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case 13:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
   End Select
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
MoveFormToCenter Me
m_bInsert = IsUserHasRightOfFunction("frm090120", strAdd, False)
m_bUpdate = IsUserHasRightOfFunction("frm090120", strEdit, False)
m_bDelete = IsUserHasRightOfFunction("frm090120", strDel, False)
m_bQuery = IsUserHasRightOfFunction("frm090120", strFind, False)
InitialField
RefreshRange
GetAllData
ShowLastRecord
UpdateToolbarState
SetCtrlReadOnly True
SetGrd
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090120 = Nothing
End Sub

Private Sub SetGrd()
    'grd1.Visible = False
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   arrGridHeadText = Array("組群", "維護人員")
   arrGridHeadWidth = Array(2000, 3000)
   grd1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1.Cols - 1
      grd1.row = 0
      grd1.col = iRow
      grd1.Text = arrGridHeadText(iRow)
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1.CellAlignment = flexAlignCenterCenter
   Next
   grd1.Visible = True
End Sub

Private Sub grd1_SelChange()
   Dim TmpRow As Integer
   'grd1.Visible = False
   TmpRow = grd1.MouseRow
   grd1.col = 0
   If TmpRow <> 0 Then
       m_CurrKEY = grd1.TextMatrix(TmpRow, 0)
       UpdateCtrlData
   End If
   grd1.Visible = True
End Sub

Private Sub ChgGrdData(iRow As Integer)

   Dim i, j, k
   
   'Modify by Morgan 2009/2/19
   'grd1.Visible = False
   'For j = 1 To Grd1.Rows - 1
   '     Grd1.row = j
   '     For k = 0 To Grd1.Cols - 1
   '         Grd1.col = k
   '         Grd1.CellBackColor = QBColor(15)
   '     Next k
   ' Next j
   
   If m_iPreRow > 0 And m_iPreRow < grd1.Rows Then
      grd1.row = m_iPreRow
      For k = 0 To grd1.Cols - 1
          grd1.col = k
          grd1.CellBackColor = QBColor(15)
      Next k
   End If
   'end 2009/2/19

   grd1.row = iRow
   For j = 0 To grd1.Cols - 1
       grd1.col = j
       grd1.CellBackColor = &HFFC0C0
       m_iPreRow = grd1.row 'Add by Morgan 2009/2/19
   Next j
   'grd1.TopRow = iRow Remove by Morgan 2009/2/19
   grd1.Visible = True
End Sub

Private Sub ChgToNowData()
Dim i, j As Integer
 j = 0
For i = 1 To grd1.Rows - 1
    If grd1.TextMatrix(i, 0) = textCD01 Then
        j = i
        Exit For
    End If
Next i
If j <> 0 Then ChgGrdData j
End Sub

Private Sub textCD01_GotFocus()
InverseTextBox textCD01
End Sub

Private Sub textCD01_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textCD01_LostFocus()
   If Trim(textCD01) <> "" And textCD01.Locked = False Then
       m_CurrKEY = ""
       GetAllData
   End If
End Sub

Private Sub textCD02_GotFocus()
InverseTextBox textCD02
OpenIme
End Sub

Private Sub textCD02_Validate(Cancel As Boolean)
If CheckLengthIsOK(textCD02, textCD02.MaxLength) = False Then
    Cancel = True
    Exit Sub
End If
CloseIme
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
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

' 初始化欄位陣列
Private Sub InitialField()
   ' 初始化欄位陣列
    m_FieldList(0).fiName = "CD01"
    m_FieldList(0).fiOldData = Empty
    m_FieldList(0).fiNewData = Empty
    m_FieldList(0).fiType = 0 '文字型態
    m_FieldList(1).fiName = "CD02"
    m_FieldList(1).fiOldData = Empty
    m_FieldList(1).fiNewData = Empty
    m_FieldList(1).fiType = 0  '文字型態
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   m_SubMode = 0
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         textCD01 = ""
         textCD02 = strUserName
         textCD02.Locked = True
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         UpdateCtrlData
         If Pub_StrUserSt03 = "M51" Or InStr(1, textCD02, strUserName) <> 0 Then
                m_EditMode = 2
                SetCtrlReadOnly False
                textCD01.Locked = True
                textCD02.Locked = False
                UpdateToolbarState
                SetInputEntry
        Else
            MsgBox "無此使用權限...", , "警告!!"
        End If
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         SetCtrlReadOnly True
         ClearField
         UpdateToolbarState
         SetInputEntry
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
         CloseIme
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT min(CD01) as CD01 FROM classdelete  "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CD01")) = False Then: m_FirstKEY = rsTmp.Fields("CD01")
   End If
   rsTmp.Close

   strSql = "SELECT max(CD01) as CD01 FROM classdelete "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CD01")) = False Then: m_LastKEY = rsTmp.Fields("CD01")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY = m_FirstKEY

   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset

   If m_CurrKEY = m_FirstKEY Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If

   strSql = "SELECT CD01 FROM classdelete " & _
            "WHERE CD01 in (select max(CD01) from classdelete where  CD01<'" & m_CurrKEY & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CD01")) = False Then: m_CurrKEY = rsTmp.Fields("CD01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   UpdateCtrlData
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset

   If m_CurrKEY = m_LastKEY Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If

   strSql = "SELECT CD01 FROM classdelete " & _
            "WHERE CD01 in (select min(CD01) from classdelete where  CD01>'" & m_CurrKEY & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CD01")) = False Then: m_CurrKEY = rsTmp.Fields("CD01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close

   UpdateCtrlData

EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY = m_LastKEY

   UpdateCtrlData
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            Toolbar1.Buttons(1).Enabled = True
         Else
            Toolbar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            Toolbar1.Buttons(2).Enabled = True
         Else
            Toolbar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            Toolbar1.Buttons(3).Enabled = True
         Else
            Toolbar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            Toolbar1.Buttons(4).Enabled = True
         Else
            Toolbar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            Toolbar1.Buttons(6).Enabled = True
            Toolbar1.Buttons(7).Enabled = True
            Toolbar1.Buttons(8).Enabled = True
            Toolbar1.Buttons(9).Enabled = True
         Else
            Toolbar1.Buttons(6).Enabled = False
            Toolbar1.Buttons(7).Enabled = False
            Toolbar1.Buttons(8).Enabled = False
            Toolbar1.Buttons(9).Enabled = False
         End If
         Toolbar1.Buttons(11).Enabled = False
         Toolbar1.Buttons(12).Enabled = False
         Toolbar1.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         Toolbar1.Buttons(1).Enabled = False
         Toolbar1.Buttons(2).Enabled = False
         Toolbar1.Buttons(3).Enabled = False
         Toolbar1.Buttons(4).Enabled = False
         Toolbar1.Buttons(6).Enabled = False
         Toolbar1.Buttons(7).Enabled = False
         Toolbar1.Buttons(8).Enabled = False
         Toolbar1.Buttons(9).Enabled = False
         Toolbar1.Buttons(11).Enabled = True
         Toolbar1.Buttons(12).Enabled = True
         Toolbar1.Buttons(14).Enabled = False
   End Select
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textCD02.Locked = bEnable
End Sub

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
            If TxtValidate = False Then Exit Function
            ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
            UpdateFieldNewData
            If AddRecord = True Then
                ChgToNowData
            Else
                Exit Function
            End If
      Case 2: '修改
            If TxtValidate = False Then Exit Function
            ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
      Case 3: '刪除
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         If DelRecord = True Then
            RefreshRange
         Else
            Exit Function
         End If
      Case 4: '列印
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
'         UpdateFieldNewData
'         'If CheckDataValid() = True Then
'         If textCD01 <> "" Then
'            If QueryRecord = False Then
'               strMsg = "無此資料"
'               strTit = "查詢資料"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               UpdateCtrlData
'            End If
'         Else
'            GoTo EXITSUB
'         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

Private Sub ClearField()
   Dim nIndex As Integer
   textCD01 = Empty
   textCD02 = Empty
   For nIndex = 0 To 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM classdelete " & _
            "WHERE CD01 = '" & m_CurrKEY & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("CD01")) = False Then: textCD01 = rsTmp.Fields("CD01"): m_CurrKEY = textCD01
      If IsNull(rsTmp.Fields("CD02")) = False Then: textCD02 = GetPrjSalesNM(rsTmp.Fields("CD02"))
      ChgToNowData
   End If
   ' 更新暫存區的資料
   UpdateFieldOldData rsTmp
   rsTmp.Close
EXITSUB:
   Set rsTmp = Nothing
End Sub

'抓當日所有資料
Private Sub GetAllData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
    strSql = "SELECT CD01,ST02 FROM classdelete,staff where cd02=st01(+) order by CD01 "
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    Set grd1.Recordset = rsTmp
    rsTmp.Close
    SetGrd
    
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textCD01.SetFocus: textCD01_GotFocus
      Case 2: textCD02.SetFocus: textCD02_GotFocus
   End Select
End Sub

Private Sub UpdateFieldNewData()
   '若新增資料
   SetFieldNewData "CD01", textCD01
   'SetFieldNewData "CD02", textCD02
End Sub

Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To 3
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False And rsTmp.RecordCount <> 0 Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim rsTmp As New ADODB.Recordset
   
   AddRecord = False
   
   bFirst = True
   strSql = "INSERT INTO classdelete (CD01,CD02,CD03,CD04) values ("
   bFirst = True
   For nIndex = 0 To 0
            strTmp = Empty
            If m_FieldList(nIndex).fiType = 0 Then
               strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
            Else
               strTmp = m_FieldList(nIndex).fiNewData
            End If
            If strTmp <> Empty Then
               If bFirst = True Then
                  strSql = strSql & strTmp
                  bFirst = False
               Else
                  strSql = strSql & "," & strTmp
               End If
            End If
   Next nIndex
   strSql = strSql & ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'HH24Mi')) "
   
On Error GoTo ErrHnd
    cnnConnection.BeginTrans
    Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
    cnnConnection.CommitTrans
    RefreshRange
    GetAllData
    ShowCurrRecord textCD01
    AddRecord = True
   Exit Function
ErrHnd:
    cnnConnection.RollbackTrans
    MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM classdelete " & _
            "WHERE CD01 = '" & strKEY01 & "' "
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To 0
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY = strKEY01
   Else
      strSql = "SELECT CD01 FROM classdelete " & _
               "WHERE CD01 = '" & m_CurrKEY & "'  "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CD01")) = False Then: m_CurrKEY = rsTmp.Fields("CD01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
   strSql = "SELECT CD01 FROM classdelete " & _
            "WHERE CD01 = '" & textCD01 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CD01")) = False Then: m_CurrKEY = rsTmp.Fields("CD01")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

If Trim(textCD01.Text) = "" And textCD01.Locked = False And textCD01.Enabled = True Then
    MsgBox "組群不可空白！", vbInformation, "操作錯誤！"
    textCD01.SetFocus
    Exit Function
End If

If CheckLengthIsOK(textCD01, textCD01.MaxLength) = False Then
   textCD01.SetFocus
   Exit Function
End If
If CheckLengthIsOK(textCD02, textCD02.MaxLength) = False Then
   textCD02.SetFocus
   Exit Function
End If
TxtValidate = True
End Function

' 修改記錄
Private Function ModRecord() As Boolean
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strCD01 As String
   
   ModRecord = False
   
   strCD01 = m_CurrKEY
   strSql = "UPDATE classdelete SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To 0
        strTmp = Empty
        If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
           If m_FieldList(nIndex).fiType = 0 Then
              If m_FieldList(nIndex).fiNewData = Empty Then
                 strTmp = m_FieldList(nIndex).fiName & " = NULL "
              Else
                 strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
              End If
           Else
              If m_FieldList(nIndex).fiNewData = Empty Then
                 strTmp = m_FieldList(nIndex).fiName & " = NULL "
              Else
                 strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
              End If
           End If
        End If
        If strTmp <> Empty Then
           bDifference = True
           If bFirst = True Then
              strSql = strSql & strTmp
              bFirst = False
           Else
              strSql = strSql & "," & strTmp
           End If
        End If
   Next nIndex

   strSql = strSql & " WHERE CD01 = '" & strCD01 & "' "
On Error GoTo ErrHnd
   If bDifference = True Then
      cnnConnection.BeginTrans
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql

      cnnConnection.CommitTrans
      
      GetAllData
      ShowCurrRecord strCD01
   End If
    ModRecord = True
   Exit Function
ErrHnd:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)
    Resume Next
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim strSql As String
   Dim strCD01 As String

   DelRecord = False

On Error GoTo Err

   strCD01 = m_CurrKEY

   strSql = "DELETE FROM classdelete " & _
            "WHERE CD01 = '" & strCD01 & "' "
   
   cnnConnection.Execute strSql
      
   RefreshRange
   GetAllData
   ShowCurrRecord strCD01
   DelRecord = True
   Exit Function
Err:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function
