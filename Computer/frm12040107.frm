VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040107 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工權限檔維護"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   516
   ClientWidth     =   9312
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9312
   Begin VB.TextBox textSR01 
      Height          =   270
      Left            =   1935
      MaxLength       =   5
      TabIndex        =   1
      Top             =   720
      Width           =   945
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   960
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
            Picture         =   "frm12040107.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040107.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040107.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040107.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040107.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040107.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040107.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040107.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040107.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040107.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040107.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9312
      _ExtentX        =   16425
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4680
      Left            =   72
      TabIndex        =   7
      Top             =   1080
      Width           =   9144
      _ExtentX        =   16129
      _ExtentY        =   8255
      _Version        =   393216
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
   Begin MSForms.TextBox textSR01_2 
      Height          =   300
      Left            =   2940
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   705
      Width           =   6255
      VariousPropertyBits=   679495711
      Size            =   "11033;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "注意事項 : "
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   7560
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "點選INSERT,UPDATE,DELETE,QUERY,PRINT,EXECUTE時, 只開放或取消該表單的單項功能"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   7800
      Width           =   8415
   End
   Begin VB.Label Label2 
      Caption         =   "點選表單編號的欄位時, 表示將該筆表單的權限全部開放或取消"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   7560
      Width           =   8415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工編號/等級編號 : "
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   765
      Width           =   1620
   End
End
Attribute VB_Name = "frm12040107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/10/15 改成Form2.0 ; textSR01_2; grdList不用修改UniCode
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer

' 第一筆資料的Key
Dim m_FirstSR As String
' 最後一筆資料的Key
Dim m_LastSR As String
' 目前正在顯示的Key
Dim m_CurrSR As String

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'Add by Morgan 2008/11/14
Dim m_bSortAsc As Boolean '排序方式


Private Sub Form_Load()
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040107", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040107", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040107", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040107", strFind, False)
   
   textSR01_2.BackColor = &H8000000F
   
   m_EditMode = 0
   MoveFormToCenter Me
      
   QueryDB
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
End Sub

' 設定控制項是否可以輸入
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textSR01.Locked = bEnable
End Sub

' 設定Key是否可以輸入
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSR01.Locked = bEnable
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得第一筆及最後一筆的Key
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT MIN(SR01) FROM STAFF_RIGHT "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_FirstSR = rsTmp.Fields(0)
   End If
   rsTmp.Close

   strSql = "SELECT MAX(SR01) FROM STAFF_RIGHT "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_LastSR = rsTmp.Fields(0)
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nRow As Integer
   Dim nCol As Integer
   textSR01 = Empty
   textSR01_2 = Empty
   For nRow = 1 To GrdList.Rows - 1
      For nCol = 2 To 11
         GrdList.TextMatrix(nRow, nCol) = Empty
      Next nCol
   Next nRow
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strSR01 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strSR01) = True Then
      m_CurrSR = strSR01
   Else
      strSql = "SELECT SR01 FROM STAFF_RIGHT " & _
                        "WHERE SR01 IN (SELECT MIN(SR01) FROM STAFF_RIGHT " & _
                                       "WHERE SR01 > '" & m_CurrSR & "' ) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SR01")) = False Then: m_CurrSR = rsTmp.Fields("SR01")
      Else
         RefreshRange
         m_CurrSR = m_LastSR
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrSR = m_FirstSR
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrSR = m_FirstSR Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT SR01 FROM STAFF_RIGHT " & _
                  "WHERE SR01 IN (SELECT MAX(SR01) FROM STAFF_RIGHT " & _
                                 "WHERE SR01 < '" & m_CurrSR & "' ) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SR01")) = False Then: m_CurrSR = rsTmp.Fields("SR01")
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
   
   If m_CurrSR = m_LastSR Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT SR01 FROM STAFF_RIGHT " & _
                  "WHERE SR01 IN (SELECT MIN(SR01) FROM STAFF_RIGHT " & _
                                 "WHERE SR01 > '" & m_CurrSR & "' ) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SR01")) = False Then: m_CurrSR = rsTmp.Fields("SR01")
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrSR = m_LastSR
   UpdateCtrlData
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         ' 90.07.13 modify by louis (依照權限設定其工具列的按紐狀態)
         'tlbar.Buttons(1).Enabled = True
         'tlbar.Buttons(2).Enabled = True
         'tlbar.Buttons(3).Enabled = True
         'tlbar.Buttons(4).Enabled = True
         'tlbar.Buttons(6).Enabled = True
         'tlbar.Buttons(7).Enabled = True
         'tlbar.Buttons(8).Enabled = True
         'tlbar.Buttons(9).Enabled = True
         'tlbar.Buttons(11).Enabled = False
         'tlbar.Buttons(12).Enabled = False
         'tlbar.Buttons(14).Enabled = True
         
         If m_bInsert Then
            tlbar.Buttons(1).Enabled = True
         Else
            tlbar.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            tlbar.Buttons(2).Enabled = True
         Else
            tlbar.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            tlbar.Buttons(3).Enabled = True
         Else
            tlbar.Buttons(3).Enabled = False
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
         ' 新增
      Case 1, 2, 3, 4:
         tlbar.Buttons(1).Enabled = False
         tlbar.Buttons(2).Enabled = False
         tlbar.Buttons(3).Enabled = False
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

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 90.07.13 modify by louis
      ' 新增
      'Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
      '   If m_EditMode = 0 Then
      '      OnAction KeyCode
      '      KeyCode = 0
      '   End If
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
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
         End If
'edit by nickc 2006/11/13
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            OnWork
            UpdateToolbarState
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
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
         OnWork
         UpdateToolbarState
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
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

' 變更INVERSE GridList中的欄位內容
Private Sub OnGrdListField(ByVal nRow As Integer, ByVal nCol As Integer)
   Dim strTemp As String
   Dim nIndex As Integer
   Dim bInverse As Boolean
   If nRow > 0 And nRow < GrdList.Rows Then
      Select Case nCol
         ' 當作用的欄位是第一欄時, 變更該筆記錄的所有權限(全選或全不選)
         'modify by toni 20080925 新增跨部門欄位8其於順延至9,10
         Case 1:
            strTemp = GrdList.TextMatrix(GrdList.row, 2)
            bInverse = True
            For nIndex = 3 To 10
               If nIndex <> 9 Then
                  If GrdList.TextMatrix(GrdList.row, nIndex) <> strTemp Then
                     bInverse = False
                     Exit For
                  End If
               End If
            Next nIndex
            If bInverse = True Then
               'Modify 2008/11/03 Toni V改秀Y
               If GrdList.TextMatrix(GrdList.row, 2) = "Y" Then
                  For nIndex = 2 To 10
                     If nIndex <> 9 Then
                        GrdList.TextMatrix(GrdList.row, nIndex) = Empty
                     End If
                  Next nIndex
               Else
                  For nIndex = 2 To 10
                     If nIndex <> 9 Then
                        GrdList.TextMatrix(GrdList.row, nIndex) = "Y"
                     End If
                  Next nIndex
               End If
            Else
               For nIndex = 2 To 10
                  If nIndex <> 9 Then
                     GrdList.TextMatrix(GrdList.row, nIndex) = "Y"
                  End If
               Next nIndex
            End If
         ' 當作用的欄位是第2,3,4,5,6,7,8欄時, 變更該筆記錄的單項權限(全選或全不選)
         '20080925 add by Toni 8是跨部門權限
         Case 2, 3, 4, 5, 6, 7, 8, 10
            'Modify 2008/11/03 Toni V改秀Y
            If GrdList.TextMatrix(GrdList.row, GrdList.col) = "Y" Then
               GrdList.TextMatrix(GrdList.row, GrdList.col) = Empty
            Else
               'Modify 2008/11/03 Toni  V改秀Y
               GrdList.TextMatrix(GrdList.row, GrdList.col) = "Y"
            End If
         '2008/11/03 ADD BY TONI 9是語文權限
         Case 9
               If nCol = 9 Then
                  If GrdList.TextMatrix(GrdList.row, GrdList.col) = "Y" Then
                     GrdList.TextMatrix(GrdList.row, GrdList.col) = "J"
                  ElseIf GrdList.TextMatrix(GrdList.row, GrdList.col) = "J" Then
                      GrdList.TextMatrix(GrdList.row, GrdList.col) = "E"
                  ElseIf GrdList.TextMatrix(GrdList.row, GrdList.col) = "E" Then
                      GrdList.TextMatrix(GrdList.row, GrdList.col) = Empty
                  Else
                     GrdList.TextMatrix(GrdList.row, GrdList.col) = "Y"
                  End If
               End If
         'Add By Sindy 2010/12/30 國內外
         Case 11
               If nCol = 11 Then
                  If GrdList.TextMatrix(GrdList.row, GrdList.col) = "C" Then
                     GrdList.TextMatrix(GrdList.row, GrdList.col) = "F"
                  ElseIf GrdList.TextMatrix(GrdList.row, GrdList.col) = "F" Then
                      GrdList.TextMatrix(GrdList.row, GrdList.col) = Empty
                  Else
                     GrdList.TextMatrix(GrdList.row, GrdList.col) = "C"
                  End If
               End If
      End Select
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm12040107 = Nothing
End Sub

Private Sub grdList_Click()
   Select Case m_EditMode
   Dim i As Integer
      Case 1, 2:
         'Modified by Morgan 2017/5/15 改用 MouseRow,MouseCol 否則點排序也會觸發( MouseRow=0, row=1)
         'OnGrdListField grdList.row, grdList.col
         OnGrdListField GrdList.MouseRow, GrdList.MouseCol
   End Select
End Sub

Private Sub grdList_KeyPress(KeyAscii As Integer)
'Removed by Morgan 2017/5/15 沒用
'   If KeyAscii = vbKeySpace Then
'      Select Case m_EditMode
'         Case 1, 2:
'            OnGrdListField grdList.row, grdList.col
'      End Select
'   End If
'end 2017/5/15
End Sub

'Add by Morgan 2008/11/14
Private Sub grdList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim iCol As Integer
   iCol = GrdList.MouseCol
   If GrdList.MouseRow < 1 Then
      GrdList.col = iCol
      If m_bSortAsc = True Then
         GrdList.Sort = 1
      Else
         GrdList.Sort = 2
      End If
      m_bSortAsc = Not m_bSortAsc
   End If
End Sub

Private Sub textSR01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 員工等級編號
Private Sub textSR01_Validate(Cancel As Boolean)
   Dim nRecordCount As Integer
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textSR01_2 = Empty
   nRecordCount = 0
   If IsEmptyText(textSR01) = False Then
      
      strSql = "SELECT 2 srt,SL02 FROM STAFF_LEVEL " & _
               "WHERE SL01 = '" & textSR01 & "' "
      'Add by Morgan 2008/9/18
      strSql = strSql & " union SELECT 1 srt,ST02 FROM STAFF " & _
               "WHERE ST01 = '" & textSR01 & "' order by 1,2"
               
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      nRecordCount = rsTmp.RecordCount
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SL02")) = False Then
            textSR01_2 = rsTmp.Fields("SL02")
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
   
   If m_EditMode <> 0 Then
      If nRecordCount <= 0 Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "資料檢核"
               strMsg = "員工權限代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSR01_GotFocus
               GoTo EXITSUB
         End Select
      End If
      
      If m_EditMode = 1 Then
         If IsRecordExist(textSR01) = True Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "該筆記錄已經存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSR01_GotFocus
         End If
      End If
   End If
EXITSUB:
End Sub

' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
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

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strSR01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM STAFF_RIGHT " & _
            "WHERE SR01 = '" & strSR01 & "' "
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

' 清除該筆記錄的原先資料
Private Sub ClearOldRecord(ByVal strSR01 As String)
   Dim strSql As String
   
   strSql = "DELETE FROM STAFF_RIGHT " & _
            "WHERE SR01 = '" & strSR01 & "' "
   cnnConnection.Execute strSql
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' 將該筆資料寫入資料庫中
'20081103 add by Toni 加入strSR10
Private Sub UpdateDB(ByVal strSR01 As String, ByVal strSR02 As String, ByVal strSR03 As String, ByVal strSR04 As String, ByVal strSR05 As String, ByVal strSR06 As String, ByVal strSR07 As String, ByVal strSR08 As String, ByVal strSR09 As String, ByVal strSR10, ByVal strSR11, ByVal strSR12)
   strSql = "INSERT INTO STAFF_RIGHT (SR01, SR02, SR03, SR04, SR05, SR06, SR07, SR08,SR09,SR10,SR11,SR12) " & _
            "VALUES ('" & strSR01 & "','" & strSR02 & "',"
   If strSR03 = Empty Then: strSql = strSql & "NULL,": Else strSql = strSql & "'Y',"
   If strSR04 = Empty Then: strSql = strSql & "NULL,": Else strSql = strSql & "'Y',"
   If strSR05 = Empty Then: strSql = strSql & "NULL,": Else strSql = strSql & "'Y',"
   If strSR06 = Empty Then: strSql = strSql & "NULL,": Else strSql = strSql & "'Y',"
   If strSR07 = Empty Then: strSql = strSql & "NULL,": Else strSql = strSql & "'Y',"
   If strSR08 = Empty Then: strSql = strSql & "NULL,": Else strSql = strSql & "'Y',"
   '20080925 add by Toni 跨部門權限
   If strSR09 = Empty Then: strSql = strSql & "NULL,": Else strSql = strSql & "'Y',"
   'end 20080925
   
   '20081103 add by Toni 語文權限
   If strSR10 = Empty Then
      strSql = strSql & "NULL,"
   Else
      Select Case strSR10
         Case "J"
            strSql = strSql & "'J',"
         Case "E"
             strSql = strSql & "'E',"
         Case "Y"
            strSql = strSql & "'Y',"
      End Select
   End If
   'end 20080925
   
   'Add By Sindy 2010/12/30
   '跨所別
   If strSR11 = Empty Then: strSql = strSql & "NULL,": Else strSql = strSql & "'Y',"
   '國內外
   If strSR12 = Empty Then
      strSql = strSql & "NULL)"
   Else
      Select Case strSR12
         Case "C"
            strSql = strSql & "'C')" '國內
         Case "F"
             strSql = strSql & "'F')" '國外
      End Select
   End If
   '2010/12/30 End
   
   cnnConnection.Execute strSql
End Sub

' 新增記錄
Private Sub AddRecord()
   Dim strSql As String
   Dim strSubSQL As String
   Dim nRow As Integer
   Dim nCol As Integer
   Dim bUpdate As Boolean
   
   '20081103 add by Toni  9是語文權限
   '20080925 add by Toni  8是跨部門權限
   'Add By Sindy 2010/12/30 跨所別
   'Add By Sindy 2010/12/30 國內外
   For nRow = 1 To GrdList.Rows - 1
      bUpdate = False
      For nCol = 2 To 11
         If GrdList.TextMatrix(nRow, nCol) = "Y" Or GrdList.TextMatrix(nRow, nCol) = "J" Or GrdList.TextMatrix(nRow, nCol) = "E" Or GrdList.TextMatrix(nRow, nCol) = "C" Or GrdList.TextMatrix(nRow, nCol) = "F" Then
            bUpdate = True
            Exit For
         End If
         
      Next nCol
      If bUpdate = True Then
         UpdateDB textSR01, GrdList.TextMatrix(nRow, 1), GrdList.TextMatrix(nRow, 2), GrdList.TextMatrix(nRow, 3), GrdList.TextMatrix(nRow, 4), GrdList.TextMatrix(nRow, 5), GrdList.TextMatrix(nRow, 6), GrdList.TextMatrix(nRow, 7), GrdList.TextMatrix(nRow, 8), GrdList.TextMatrix(nRow, 9), GrdList.TextMatrix(nRow, 10), GrdList.TextMatrix(nRow, 11)
      End If
   Next nRow
   ShowCurrRecord textSR01
End Sub

' 修改記錄
Private Sub ModRecord()
   Dim strSql As String
   Dim strSubSQL As String
   Dim nRow As Integer
   Dim bUpdate As Boolean
   Dim nCol As Integer
   
   '20081103 add by Toni  9是語文權限
   '20080925 add by Toni  8是跨部門權限
   'Add By Sindy 2010/12/30 跨所別
   'Add By Sindy 2010/12/30 國內外
   For nRow = 1 To GrdList.Rows - 1
      bUpdate = False
      For nCol = 2 To 11
         If GrdList.TextMatrix(nRow, nCol) = "Y" Or GrdList.TextMatrix(nRow, nCol) = "J" Or GrdList.TextMatrix(nRow, nCol) = "E" Or GrdList.TextMatrix(nRow, nCol) = "C" Or GrdList.TextMatrix(nRow, nCol) = "F" Then
            bUpdate = True
            Exit For
         End If
         
      Next nCol
      If bUpdate = True Then
         UpdateDB m_CurrSR, GrdList.TextMatrix(nRow, 1), GrdList.TextMatrix(nRow, 2), GrdList.TextMatrix(nRow, 3), GrdList.TextMatrix(nRow, 4), GrdList.TextMatrix(nRow, 5), GrdList.TextMatrix(nRow, 6), GrdList.TextMatrix(nRow, 7), GrdList.TextMatrix(nRow, 8), GrdList.TextMatrix(nRow, 9), GrdList.TextMatrix(nRow, 10), GrdList.TextMatrix(nRow, 11)
      End If
   Next nRow
   ShowCurrRecord textSR01
End Sub

' 刪除記錄
Private Sub DelRecord()
   ClearOldRecord m_CurrSR
   ShowCurrRecord m_CurrSR
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   QueryRecord = False
   
   If IsRecordExist(textSR01) = True Then
      m_CurrSR = textSR01
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If
   
   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Select Case m_EditMode
      Case 1:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            AddRecord
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         'Add By Cheng 2002/05/23
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         ClearOldRecord m_CurrSR
         ModRecord
      Case 3:
         DelRecord
         RefreshRange
      Case 4:
         If CheckDataValid() = True Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 將GridList的欄位清成未選取的狀態
Private Sub ClearGrdListField()
   Dim nRow As Integer
   Dim nCol As Integer
   For nRow = 1 To GrdList.Rows - 1
      For nCol = 2 To 11
         GrdList.TextMatrix(nRow, nCol) = Empty
      Next nCol
   Next nRow
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim nRow As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strSrc As String
   Dim strDes As String
   
   textSR01 = m_CurrSR
   'textSR01_Validate False
   If IsEmptyText(textSR01) = False Then
      textSR01_2 = Empty
      strSql = "SELECT 2,SL02 FROM STAFF_LEVEL " & _
            "WHERE SL01 = '" & textSR01 & "' "
      'Add by Morgan 2008/9/18
      strSql = strSql & " union SELECT 1 srt,ST02 FROM STAFF " & _
               "WHERE ST01 = '" & textSR01 & "' order by 1,2"
               
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SL02")) = False Then
            textSR01_2 = rsTmp.Fields("SL02")
         End If
      End If
      rsTmp.Close
   End If
   
   ' 清除GridList中的欄位成未選取的狀態
   ClearGrdListField
   ' 讀取Table中的資料
   strSql = "SELECT * FROM STAFF_RIGHT " & _
            "WHERE SR01 = '" & m_CurrSR & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         strSrc = UCase(rsTmp.Fields("SR02"))
         ' 在GridList找尋該表單編號的列
         For nRow = 1 To GrdList.Rows - 1
            strDes = UCase(GrdList.TextMatrix(nRow, 1))
            ' 若找到該表單邊號則更新其狀態
            'If grdList.TextMatrix(nRow, 1) = rsTmp.Fields("SR02") Then
            If strSrc = strDes Then
               If IsNull(rsTmp.Fields("SR03")) = False Then
                  If rsTmp.Fields("SR03") = "Y" Then
                     GrdList.TextMatrix(nRow, 2) = "Y"
                  End If
               End If
               If IsNull(rsTmp.Fields("SR04")) = False Then
                  If rsTmp.Fields("SR04") = "Y" Then
                     GrdList.TextMatrix(nRow, 3) = "Y"
                  End If
               End If
               If IsNull(rsTmp.Fields("SR05")) = False Then
                  If rsTmp.Fields("SR05") = "Y" Then
                     GrdList.TextMatrix(nRow, 4) = "Y"
                  End If
               End If
               If IsNull(rsTmp.Fields("SR06")) = False Then
                  If rsTmp.Fields("SR06") = "Y" Then
                     GrdList.TextMatrix(nRow, 5) = "Y"
                  End If
               End If
               If IsNull(rsTmp.Fields("SR07")) = False Then
                  If rsTmp.Fields("SR07") = "Y" Then
                     GrdList.TextMatrix(nRow, 6) = "Y"
                  End If
               End If
               If IsNull(rsTmp.Fields("SR08")) = False Then
                  If rsTmp.Fields("SR08") = "Y" Then
                     GrdList.TextMatrix(nRow, 7) = "Y"
                  End If
               End If
               '20080925 add by Toni 跨部門權限
               If IsNull(rsTmp.Fields("SR09")) = False Then
                  If rsTmp.Fields("SR09") = "Y" Then
                     GrdList.TextMatrix(nRow, 8) = "Y"
                  End If
               End If
               'end 20080925
               '20081103 add by Toni 語文權限
               If IsNull(rsTmp.Fields("SR10")) = False Then
                  '2010/8/10 MODIFY BY SONIA 94007加frm140404的A權限
                  'If rsTmp.Fields("SR10") = "J" Then
                  '   grdList.TextMatrix(nRow, 9) = "J"  '日文
                  'ElseIf rsTmp.Fields("SR10") = "E" Then
                  '   grdList.TextMatrix(nRow, 9) = "E"   '英文
                  'ElseIf rsTmp.Fields("SR10") = "Y" Then
                  '   grdList.TextMatrix(nRow, 9) = "Y"  '全部
                  'End If
                  GrdList.TextMatrix(nRow, 9) = rsTmp.Fields("SR10")
                  '2010/8/10 END
               End If
               'end 20081103
               
               'Add By Sindy 2010/12/30
               '跨所別
               If IsNull(rsTmp.Fields("SR11")) = False Then
                  If rsTmp.Fields("SR11") = "Y" Then
                     GrdList.TextMatrix(nRow, 10) = "Y"
                  End If
               End If
               '國內外
               If IsNull(rsTmp.Fields("SR12")) = False Then
                  GrdList.TextMatrix(nRow, 11) = rsTmp.Fields("SR12")
               End If
               '2010/12/30 End
               
               Exit For
            End If
         Next nRow
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'' 更新GridList中的內容
''2008/11/03 add by Toni增加strSR10
''20080925 add by Toni 增加strSR09
'Private Sub UpdateGridListData(ByVal strSR02 As String, ByVal strSR03 As String, ByVal strSR04 As String, ByVal strSR05 As String, ByVal strSR06 As String, ByVal strSR07 As String, ByVal strSR08 As String, ByVal strSR09 As String, ByVal strSR10 As String)
' Dim nRow As Integer
'   For nRow = 1 To grdList.Rows - 1
'      If grdList.TextMatrix(nRow, 1) = strSR02 Then
'         If strSR03 = "Y" Then: grdList.TextMatrix(nRow, 2) = "Y"
'         Exit For
'      End If
'   Next
'End Sub

' 初始化列表
Public Sub InitialGridList()
   GrdList.Clear
   GrdList.Rows = 1
   'grdList.Cols = 10
   'edit by toni 20081103
   'grdList.Cols = 11
   'Modify By Sindy 2010/12/30
   'grdList.Cols = 12
   GrdList.Cols = 14
   
   GrdList.ColWidth(0) = 300
   GrdList.row = 0
   
   GrdList.col = 1
   GrdList.Text = "表單編號"
   GrdList.ColWidth(1) = 1500
   GrdList.ColAlignment(1) = flexAlignLeftCenter
   GrdList.col = 2
   GrdList.Text = "增"
   GrdList.ColWidth(2) = 270
   GrdList.ColAlignment(2) = flexAlignCenterCenter
   GrdList.col = 3
   GrdList.Text = "修"
   GrdList.ColWidth(3) = 270
   GrdList.ColAlignment(3) = flexAlignCenterCenter
   GrdList.col = 4
   GrdList.Text = "刪"
   GrdList.ColWidth(4) = 270
   GrdList.ColAlignment(4) = flexAlignCenterCenter
   GrdList.col = 5
   GrdList.Text = "查"
   GrdList.ColWidth(5) = 270
   GrdList.ColAlignment(5) = flexAlignCenterCenter
   GrdList.col = 6
   GrdList.Text = "印"
   GrdList.ColWidth(6) = 270
   GrdList.ColAlignment(6) = flexAlignCenterCenter
   GrdList.col = 7
   GrdList.Text = "執行"
   GrdList.ColWidth(7) = 420
   GrdList.ColAlignment(7) = flexAlignCenterCenter
   '20080925 add by Toni
   GrdList.col = 8
   GrdList.Text = "跨部門"
   GrdList.ColWidth(8) = 600
   GrdList.ColAlignment(8) = flexAlignCenterCenter
   'end 20080925
   '20081103 add by Toni
   GrdList.col = 9
   GrdList.Text = "語文"
   GrdList.ColWidth(9) = 420
   GrdList.ColAlignment(9) = flexAlignCenterCenter
   'end 20081103
   'Add By Sindy 2010/12/30
   GrdList.col = 10
   GrdList.Text = "跨所"
   GrdList.ColWidth(10) = 420
   GrdList.ColAlignment(10) = flexAlignCenterCenter
   GrdList.col = 11
   GrdList.Text = "國內外"
   GrdList.ColWidth(11) = 600
   GrdList.ColAlignment(11) = flexAlignCenterCenter
   '2010/12/30 End
   GrdList.col = 12
   GrdList.Text = "表單名稱"
   GrdList.ColWidth(12) = 2000
   GrdList.ColAlignment(12) = flexAlignLeftCenter
   GrdList.col = 13
   GrdList.Text = "表單描述"
   GrdList.ColWidth(13) = 3000
   GrdList.ColAlignment(13) = flexAlignLeftCenter
End Sub

' 讀取DATABASE 的 FORM
Private Sub QueryDB()
   Dim nRow As Integer
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset

   InitialGridList

   strSql = "SELECT * FROM FORM ORDER BY FO03, FO01"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         GrdList.Rows = GrdList.Rows + 1
         nRow = GrdList.Rows - 1
         If IsNull(rsTmp.Fields("FO01")) = False Then
            GrdList.TextMatrix(nRow, 1) = rsTmp.Fields("FO01")
         End If
         If IsNull(rsTmp.Fields("FO02")) = False Then
            GrdList.TextMatrix(nRow, 12) = rsTmp.Fields("FO02")
         End If
         If IsNull(rsTmp.Fields("FO03")) = False Then
            GrdList.TextMatrix(nRow, 13) = rsTmp.Fields("FO03")
         End If
         rsTmp.MoveNext
      Loop
      GrdList.FixedRows = 1 'Added by Lydia 2023/10/16
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 設定輸入的位置
Private Sub SetInputEntry()
   textSR01.SetFocus
End Sub

' 檢查輸入的資料是否已經完整
Private Function CheckDataValid() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 4:
         If IsEmptyText(textSR01) = True Then
            strTit = "檢核資料"
            strMsg = "請先輸入員工等級編號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
   End Select
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textSR01_GotFocus()
   InverseTextBox textSR01
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textSR01.Enabled = True Then
   Cancel = False
   textSR01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function
