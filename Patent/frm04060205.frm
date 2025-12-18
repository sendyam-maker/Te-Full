VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060205 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸事務所資料維護"
   ClientHeight    =   5760
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   9144
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9144
   Begin VB.TextBox textFNM01 
      Height          =   264
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   0
      Top             =   660
      Width           =   612
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9144
      _ExtentX        =   16129
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
   Begin MSComctlLib.ImageList ImgList 
      Left            =   8400
      Top             =   660
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
            Picture         =   "frm04060205.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060205.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060205.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060205.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060205.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060205.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060205.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060205.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060205.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060205.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060205.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4272
      Left            =   72
      TabIndex        =   5
      Top             =   1392
      Width           =   8892
      _ExtentX        =   15685
      _ExtentY        =   7535
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
   Begin MSForms.TextBox textFNM02 
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   3015
      VariousPropertyBits=   671107099
      MaxLength       =   30
      Size            =   "5318;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   9000
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   9000
      Y1              =   1344
      Y2              =   1344
   End
   Begin VB.Label Label2 
      Caption         =   "事務所編號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "事務所名稱  :"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1212
   End
End
Attribute VB_Name = "frm04060205"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Morgan 2021/12/24 改成Form2.0 (grdList,textFNM02)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
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
' 選取列
Dim m_CurrSel As Integer
'Add By Sindy 2014/4/23 執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'2014/4/23 END


Private Sub Form_Load()
   m_EditMode = 0
   MoveFormToCenter Me
   
   'Add By Sindy 2014/4/23 取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   '2014/4/23 END
   
   InitialField
   QueryDB
   ShowFirstRecord
   SetCtrlReadOnly True
   UpdateToolbarState
End Sub

' 按下按鍵
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'      ' 新增
'      Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd, vbKeyEscape:
'         If m_EditMode = 0 Then
'            OnAction KeyCode
'         End If
'      Case vbKeyF9, vbKeyF10, vbKeyReturn:
'         If m_EditMode <> 0 Then
'            If KeyCode = vbKeyReturn Then: KeyCode = vbKeyF9
'            OnAction KeyCode
'         End If
'   End Select
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
      Case vbKeyF9, vbKeyF10:
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

' 讀取資料庫所有的資料
Private Sub QueryDB()
   Dim strSql As String
   
   ' 檢查RecordSet的狀態
   If m_Recordset.State <> adStateClosed Then
      m_Recordset.Close
   End If
   ' 設定 Query 的命令
   strSql = "SELECT * FROM CAgent ORDER BY FNM01"
            
   ' 讀取資料庫
   m_Recordset.CursorLocation = adUseClient
   m_Recordset.Open strSql, cnnConnection, adOpenDynamic
   
   ' 更新 GridList
   UpdateGridList
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "FNM" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 5:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
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
   SetFieldNewData "FNM01", textFNM01
   SetFieldNewData "FNM02", textFNM02
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

' 清除欄位內的資料內容
Private Sub ClearField()
   textFNM01 = Empty
   textFNM02 = Empty
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textFNM01.Locked = bEnable
   textFNM02.Locked = bEnable
End Sub
' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textFNM01.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   ' 判斷是否有記錄存在且記錄指標的位置是有資料存在的
   If m_Recordset.RecordCount <= 0 Then: GoTo EXITSUB: 'End If
   If m_Recordset.BOF = True Then: GoTo EXITSUB: 'End If
   If m_Recordset.EOF = True Then: GoTo EXITSUB: 'End If
   
   ClearField
   If IsNull(m_Recordset.Fields("FNM01")) = False Then
      textFNM01 = m_Recordset.Fields("FNM01")
   End If
   If IsNull(m_Recordset.Fields("FNM02")) = False Then
      textFNM02 = m_Recordset.Fields("FNM02")
   End If
EXITSUB:
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strFNM01 As String)
   Dim bFind As Boolean
   If m_Recordset.RecordCount > 0 Then
      bFind = False
      m_Recordset.MoveFirst
      While m_Recordset.EOF = False And bFind = False
         If strFNM01 = m_Recordset.Fields("FNM01") Then
            bFind = True
         Else
            m_Recordset.MoveNext
         End If
      Wend
      
      If bFind Then
         UpdateCtrlData
      Else
         ShowFirstRecord
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
' 更新欄位控制項的狀態
Private Sub UpdateFieldState()
   Select Case m_EditMode
      ' 無
      Case 0:
         textFNM01.Locked = True
         textFNM02.Locked = True
      ' 新增
      Case 1:
         textFNM01.Locked = False
         textFNM02.Locked = False
      ' 修改
      Case 2:
         textFNM01.Locked = True
         textFNM02.Locked = False
      ' 查詢
      Case 4:
         textFNM01.Locked = False
         textFNM02.Locked = True
      Case Else:
         textFNM01.Locked = True
         textFNM02.Locked = True
   End Select
End Sub
' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
'         tlbar.Buttons(1).Enabled = True
'         tlbar.Buttons(2).Enabled = True
'         tlbar.Buttons(3).Enabled = True
'         tlbar.Buttons(4).Enabled = True
'         tlbar.Buttons(6).Enabled = True
'         tlbar.Buttons(7).Enabled = True
'         tlbar.Buttons(8).Enabled = True
'         tlbar.Buttons(9).Enabled = True
'         tlbar.Buttons(11).Enabled = False
'         tlbar.Buttons(12).Enabled = False
'         tlbar.Buttons(14).Enabled = True
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
   UpdateFieldState
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
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
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
         ClearField
         m_EditMode = 4
         SetKeyReadOnly False
         UpdateToolbarState
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
         'Added by Morgan 2021/12/24 檢查畫面輸入欄位是否含有Unicode文字
         '放CheckDataValid沒用，要在更新變數值前就檢查
         If PUB_ChkUniText(Me, , True, "TextBox") = False Then
             Exit Sub
         End If
         'end 2021/12/24
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
   SetEntryFocus
End Sub

Private Sub textFNM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmpty(textFNM01) = False Then
      ' 離開欄位時檢查是否該代號有重覆
      If m_EditMode = 1 Then
         ' 檢查記錄是否已存在
         If IsRecordExist(textFNM01) = True Then
            Cancel = True
            strTit = "新增資料"
            strMsg = "該筆記錄已存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textFNM01_GotFocus
         End If
      End If
   End If
End Sub

' 事務所名稱
Private Sub textFNM02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmpty(textFNM02) = False Then
      If StrLength(textFNM02) > 30 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "事務所名稱資料太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFNM02_GotFocus
      End If
   End If
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: textFNM02.IMEMode = 2
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
Private Function IsRecordExist(ByVal strFNM01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM CAgent " & _
            "WHERE FNM01 = '" & strFNM01 & "'"
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
   If (m_Recordset.State <> adStateClosed) Then
      m_Recordset.Close
   End If
   Set m_Recordset = Nothing
   'Add By Cheng 2002/07/18
   Set frm04060205 = Nothing
End Sub

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
   Dim strFNM01 As String
      
   strFNM01 = textFNM01
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strFNM01) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO CAgent ("
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
      ShowCurrRecord strFNM01
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
   Dim strFNM01 As String
      
   strFNM01 = textFNM01
   
   strSql = "UPDATE CAgent SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = m_FieldList(nIndex).fiName & " = '" & m_FieldList(nIndex).fiNewData & "'"
         Else
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = " & 0
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
                  "WHERE FNM01 = '" & strFNM01 & "'"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      QueryDB
      ShowCurrRecord strFNM01
   End If

End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim nIndex As Integer
   Dim nSel As Integer
   Dim strSql As String
   Dim strFNM01 As String
   
   strFNM01 = textFNM01
   
   nSel = 1
   For nIndex = 1 To grdList.Rows - 1
      If strFNM01 = grdList.TextMatrix(nIndex, 1) Then: nSel = nIndex
   Next nIndex
   
   strSql = "DELETE FROM CAgent " & _
            "WHERE FNM01 = '" & strFNM01 & "'"
                  
   cnnConnection.Execute strSql
   
   QueryDB
   grdList_SetSelection nSel
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strSql As String
   Dim nIndex As Index
   Dim nPos
   Dim bFind As Boolean
   Dim strFNM01 As String
   
   strFNM01 = textFNM01
   nPos = m_Recordset.AbsolutePosition
   QueryRecord = False
   bFind = False
   m_Recordset.MoveFirst
   While (m_Recordset.EOF <> True) And (bFind = False)
      If m_Recordset.Fields("FNM01") = strFNM01 Then
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
         If CheckDataValid = True Then
            AddRecord
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid = True Then
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
   QueryDB
   SetCtrlReadOnly True
EXITSUB:
End Sub

Private Sub grdList_SelChange()
   Dim strFNM01 As String
   If grdList.row > 0 Then
      strFNM01 = grdList.TextMatrix(grdList.row, 1)
   End If
   grdList_ShowSelection
   ShowCurrRecord strFNM01
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If grdList.Rows > 0 Then
      If nSel <= 0 Then: nSel = 1
      If nSel > grdList.Rows - 1 Then: nSel = grdList.Rows - 1
      grdList.row = nSel
      grdList_SelChange
      grdList_ShowSelection
   End If
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

Private Sub SetEntryFocus()
   Select Case m_EditMode
      Case 1, 4:
         If textFNM01.Locked = False Then
            textFNM01.SetFocus
         End If
      Case 2:
         If textFNM02.Locked = False Then
            textFNM02.SetFocus
         End If
   End Select
End Sub

' 初始化 Data Grid
Private Sub InitialGridList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 3
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "事務所編號"
   grdList.ColWidth(1) = 1000
   grdList.ColAlignment(1) = flexAlignCenterCenter
   grdList.col = 2
   grdList.Text = "事務所名稱"
   grdList.ColWidth(2) = 2000
   grdList.ColAlignment(2) = flexAlignLeftCenter
End Sub

Private Sub UpdateGridList()
   Dim nRow As Integer
   
   grdList.Clear
   InitialGridList
   
   If m_Recordset.RecordCount > 0 Then
      grdList.Rows = m_Recordset.RecordCount + 1
      m_Recordset.MoveFirst
      nRow = 1
      While m_Recordset.EOF <> True
         grdList.row = nRow
         
         grdList.col = 1
         If IsNull(m_Recordset.Fields("FNM01")) = False Then
            grdList.Text = m_Recordset.Fields("FNM01")
         End If
         
         grdList.col = 2
         If IsNull(m_Recordset.Fields("FNM02")) = False Then
            grdList.Text = m_Recordset.Fields("FNM02")
         End If
         
         nRow = nRow + 1
         m_Recordset.MoveNext
      Wend
      'Added by Lydia 2023/10/17
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/17
   End If
End Sub

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As Object)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Private Sub textFNM01_GotFocus()
   InverseAll textFNM01
End Sub

Private Sub textFNM02_GotFocus()
   InverseAll textFNM02
   'edit by nickc 2007/07/11 切換輸入法改用API
   'textFNM02.IMEMode = 1
   OpenIme
End Sub

Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
      
   If IsEmpty(textFNM01) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入事務所編號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If IsEmpty(textFNM02) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入事務所名稱"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If StrLength(textFNM02) > 30 Then
      strTit = "檢核資料"
      strMsg = "事務所名稱太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Public Function IsEmpty(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   IsEmpty = False
   
   If Len(strData) <= 0 Then
      IsEmpty = True
   Else
      IsEmpty = True
      For nIndex = 1 To Len(strData)
         If Mid(strData, nIndex, 1) <> " " Then
            IsEmpty = False
            Exit For
         End If
      Next nIndex
   End If
End Function

