VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm020508 
   BorderStyle     =   1  '單線固定
   Caption         =   "主張內容分類資料維護"
   ClientHeight    =   5736
   ClientLeft      =   216
   ClientTop       =   996
   ClientWidth     =   9132
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9132
   Begin TabDlg.SSTab tabCtrl 
      Height          =   4932
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8952
      _ExtentX        =   15790
      _ExtentY        =   8700
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "單筆"
      TabPicture(0)   =   "frm020508.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "textCC01"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "textCC02"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm020508.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdList"
      Tab(1).ControlCount=   1
      Begin VB.TextBox textCC02 
         Height          =   1212
         Left            =   1080
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   2
         Top             =   720
         Width           =   7692
      End
      Begin VB.TextBox textCC01 
         Height          =   264
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   1
         Top             =   360
         Width           =   492
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   4392
         Left            =   -74904
         TabIndex        =   6
         Top             =   432
         Width           =   8712
         _ExtentX        =   15367
         _ExtentY        =   7747
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
      Begin VB.Label Label3 
         Caption         =   "主張內容 :"
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "分類代號 :"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   972
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   8580
      Top             =   600
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
            Picture         =   "frm020508.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020508.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020508.frx":0670
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020508.frx":084C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020508.frx":0B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020508.frx":0E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020508.frx":11A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020508.frx":14BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020508.frx":17D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020508.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm020508.frx":1E10
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9132
      _ExtentX        =   16108
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgList"
      DisabledImageList=   "ImgList"
      HotImageList    =   "ImgList"
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
End
Attribute VB_Name = "frm020508"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo By Lydia 2021/11/22 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
Option Explicit

Const MAX_FIELD = 2

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList(MAX_FIELD) As FIELDITEM

' 變數宣告區
Dim m_Recordset As New ADODB.Recordset
Dim m_EditMode As Integer
'
Dim m_CurrSel As Integer

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

' Load Form
Private Sub Form_Load()
   ' 先顯示多筆查詢的畫面
   tabCtrl.Tab = 1
   
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm020508", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm020508", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm020508", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm020508", strFind, False)
   
   m_EditMode = 0
   MoveFormToCenter Me
   
   InitialField
   QueryDB
   ShowFirstRecord
   SetCtrlReadOnly True
   UpdateToolbarState
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CC" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
   Next nIndex
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, ByVal strData As String)
   Dim nIndex As Integer
   For nIndex = 0 To MAX_FIELD - 1
      If strName = m_FieldList(nIndex).fiName Then
         m_FieldList(nIndex).fiNewData = strData
         Exit For
      End If
   Next nIndex
End Sub

' 更新欄位的內容
Private Sub UpdateFieldNewData()
   SetFieldNewData "CC01", textCC01: SetFieldNewData "CC02", textCC02
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData()
   Dim nIndex As Integer
   Dim strTmp As String
   
   If IsRecordsetCorrect = False Then
      GoTo EXITSUB
   End If
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(m_Recordset.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = m_Recordset.Fields(m_FieldList(nIndex).fiName)
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = m_Recordset.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 讀取資料庫所有的資料
Private Sub QueryDB()
   Dim strSql As String
   
   ' 檢查RecordSet的狀態
   If m_Recordset.State <> adStateClosed Then
      m_Recordset.Close
   End If
   ' 設定 Query 的命令
   strSql = "SELECT * FROM ClaimContents " & _
            "ORDER BY CC01"
   ' 讀取資料庫
   m_Recordset.CursorLocation = adUseClient
   m_Recordset.Open strSql, cnnConnection, adOpenDynamic
   
   ' 更新 GridList
   UpdateGridList
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   textCC01 = Empty: textCC02 = Empty
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textCC01.Locked = bEnable: textCC02.Locked = bEnable
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textCC01.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   ' 判斷是否有記錄存在且記錄指標的位置是有資料存在的
   If m_Recordset.RecordCount <= 0 Then: GoTo EXITSUB: 'End If
   If m_Recordset.BOF = True Then: GoTo EXITSUB: 'End If
   If m_Recordset.EOF = True Then: GoTo EXITSUB: 'End If
   
   ClearField
   textCC01 = m_Recordset.Fields("CC01")
   If Not IsNull(m_Recordset.Fields("CC02")) Then: textCC02 = m_Recordset.Fields("CC02"): 'End If
   UpdateFieldOldData
   
EXITSUB:
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strCC01 As String)
   Dim bFind As Boolean
   If m_Recordset.RecordCount > 0 Then
      bFind = False
      m_Recordset.MoveFirst
      While m_Recordset.EOF = False And bFind = False
         If strCC01 = m_Recordset.Fields("CC01") Then
            bFind = True
         Else
            m_Recordset.MoveNext
         End If
      Wend
      
      If bFind Then
         UpdateCtrlData
      Else
         'ShowFirstRecord
         m_Recordset.MoveFirst
         Do While m_Recordset.EOF = False
            If strCC01 < m_Recordset.Fields("CC01") Then
               Exit Do
            End If
            m_Recordset.MoveNext
         Loop
         If m_Recordset.EOF = True Then: m_Recordset.MoveLast
         UpdateCtrlData
      End If
   End If
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   If m_Recordset.RecordCount > 0 Then
      m_Recordset.MoveFirst
      UpdateCtrlData
   End If
End Sub
' 顯示上一筆資料
Private Sub ShowPrevRecord()
   If m_Recordset.RecordCount > 0 Then
      If m_Recordset.BOF = False Then
         m_Recordset.MovePrevious
         ' 若記錄指標在記錄之前則將記錄指標移至第一筆
         If m_Recordset.BOF = True Then
            ShowMsg MsgText(9008)
            m_Recordset.MoveFirst
         End If
         UpdateCtrlData
      End If
   End If
End Sub
' 顯示下一筆資料
Private Sub ShowNextRecord()
   If m_Recordset.RecordCount > 0 Then
      If m_Recordset.EOF = False Then
         m_Recordset.MoveNext
         ' 若記錄指標在記錄之前則將記錄指標移至第一筆
         If m_Recordset.EOF = True Then
            ShowMsg MsgText(9009)
            m_Recordset.MoveLast
         End If
         UpdateCtrlData
      End If
   End If
End Sub
' 顯示最後一筆資料
Private Sub ShowLastRecord()
   If m_Recordset.RecordCount > 0 Then
      m_Recordset.MoveLast
      UpdateCtrlData
   End If
End Sub
' 檢查目前 m_Recordset 的狀態是否正常
Private Function IsRecordsetCorrect() As Boolean
   IsRecordsetCorrect = True
   If m_Recordset.State = adStateClosed Then
      IsRecordsetCorrect = False
      GoTo ExitFun
   End If
   If m_Recordset.RecordCount <= 0 Then
      IsRecordsetCorrect = False
      GoTo ExitFun
   End If
   If m_Recordset.BOF = True Then
      IsRecordsetCorrect = False
      GoTo ExitFun
   End If
   If m_Recordset.EOF = True Then
      IsRecordsetCorrect = False
      GoTo ExitFun
   End If
ExitFun:
End Function
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
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
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
      Case vbKeyF9, vbKeyF10, vbKeyReturn:
         If KeyCode = vbKeyReturn Then: KeyCode = vbKeyF9
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
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
         ClearField
         SetKeyReadOnly False
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
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
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

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm020508 = Nothing
End Sub

Private Sub textCC01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 分類代號
Private Sub textCC01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If m_EditMode = 1 Then
      If IsEmptyText(textCC01) = False Then
         If IsRecordExist(textCC01) = True Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "分類代號已存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCC01_GotFocus
         End If
      End If
   End If
End Sub

' 主張內容
Private Sub textCC02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If Not CheckLengthIsOK(textCC02, 100) Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "主張內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCC02_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCC02.IMEMode = 2
   If Cancel = False Then CloseIme
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
Private Function IsRecordExist(ByVal strCC01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM ClaimContents " & _
            "WHERE CC01 = '" & strCC01 & "'"
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 新增記錄
Private Sub AddRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strCC01 As String
   
   strCC01 = textCC01
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strCC01) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO ClaimContents ("
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
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
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & m_FieldList(nIndex).fiNewData & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
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
   strSql = strSql & ")"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      QueryDB
      ShowCurrRecord strCC01
   End If
   
EXITSUB:
End Sub

' 修改記錄
Private Sub ModRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strCC01 As String
   
   strCC01 = textCC01
   
   strSql = "UPDATE ClaimContents SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_FieldList(nIndex).fiName & " = '" & m_FieldList(nIndex).fiNewData & "'"
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
   
   strSql = strSql & " " & _
                  "WHERE CC01 = '" & strCC01 & "'"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      QueryDB
      ShowCurrRecord strCC01
   End If

End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strCC01 As String
   
   strCC01 = textCC01
   
   strSql = "DELETE FROM ClaimContents " & _
            "WHERE CC01 = '" & strCC01 & "'"
                  
   cnnConnection.Execute strSql
   
   QueryDB
   'ShowFirstRecord
   ShowCurrRecord strCC01
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strSql As String
   Dim nIndex As Integer
   Dim nPos As Integer
   Dim bFind As Boolean
   Dim strCC01 As String
   
   strCC01 = textCC01
   
   nPos = m_Recordset.AbsolutePosition
   QueryRecord = False
   bFind = False
   m_Recordset.MoveFirst
   While (m_Recordset.EOF <> True) And (bFind = False)
      If m_Recordset.Fields("CC01") = strCC01 Then
         bFind = True
      Else
         m_Recordset.MoveNext
      End If
   Wend
   
   If bFind = True Then
      UpdateCtrlData
      UpdateToolbarState
   Else
      m_Recordset.AbsolutePosition = nPos
      UpdateToolbarState
   End If

   QueryRecord = bFind
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
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            ModRecord
         Else
            GoTo EXITSUB
         End If
      Case 3:
         DelRecord
      Case 4:
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 初始化列表
Public Sub InitialGridList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 3
   
   grdList.ColWidth(0) = 300
   grdList.row = 0
      
   grdList.col = 1
   grdList.Text = "代號"
   grdList.ColWidth(1) = 800
   grdList.ColAlignment(1) = flexAlignLeftCenter
   grdList.col = 2
   grdList.Text = "說明"
   grdList.ColWidth(2) = 3000
   grdList.ColAlignment(2) = flexAlignLeftCenter
End Sub

Private Sub UpdateGridList()
   Dim strCC01 As String
   Dim nRow As Integer
   
   grdList.Clear
   InitialGridList
   
   If IsRecordsetCorrect = True Then
      strCC01 = m_Recordset.Fields("CC01")
      
      grdList.Rows = m_Recordset.RecordCount + 1
      m_Recordset.MoveFirst
      nRow = 1
      While m_Recordset.EOF <> True
         grdList.row = nRow
         
         grdList.col = 1
         If IsNull(m_Recordset.Fields("CC01")) = False Then
            grdList.Text = m_Recordset.Fields("CC01")
         End If
         
         grdList.col = 2
         If IsNull(m_Recordset.Fields("CC02")) = False Then
            grdList.Text = m_Recordset.Fields("CC02")
         End If
         
         nRow = nRow + 1
         m_Recordset.MoveNext
      Wend
      grdList.FixedRows = 1  'Added by Lydia 2023/10/16
      ShowCurrRecord strCC01
   End If
End Sub

Private Sub grdList_Click()
   If grdList.row > 0 Then
      grdList_SelChange
   End If
End Sub

Private Sub grdList_SelChange()
   Dim strCC01 As String
   
   If grdList.row > 0 Then
      strCC01 = grdList.TextMatrix(grdList.row, 1)
      ShowCurrRecord strCC01
   End If
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

Private Sub textCC01_GotFocus()
   InverseTextBox textCC01
End Sub

Private Sub textCC02_GotFocus()
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCC02.IMEMode = 1
   OpenIme
   InverseTextBox textCC02
End Sub

Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1, 4:
         textCC01.SetFocus
      Case 2:
         textCC02.SetFocus
   End Select
End Sub

' 檢查欄位內容
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   ' 分類代號
   If IsEmptyText(textCC01) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入分類代號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCC01.SetFocus
      GoTo EXITSUB
   End If
   ' 主張內容
   If IsEmptyText(textCC02) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入主張內容"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCC02.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCC01.Enabled = True Then
   Cancel = False
   textCC01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCC02.Enabled = True Then
   Cancel = False
   textCC02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

