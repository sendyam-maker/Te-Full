VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030601 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內公報代理人資料維護"
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
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   3
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
   Begin TabDlg.SSTab tabCtrl 
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   660
      Width           =   8955
      _ExtentX        =   15790
      _ExtentY        =   8911
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "單筆"
      TabPicture(0)   =   "frm030601.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "textTA02"
      Tab(0).Control(1)=   "textTA04"
      Tab(0).Control(2)=   "textTA03"
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(6)=   "Label3"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm030601.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grdList"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.TextBox textTA02 
         Height          =   264
         Left            =   -73680
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   492
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   4665
         Left            =   60
         TabIndex        =   9
         Top             =   330
         Width           =   8805
         _ExtentX        =   15536
         _ExtentY        =   8234
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   0
         SelectionMode   =   1
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
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.ComboBox textTA04 
         Height          =   300
         Left            =   -73680
         TabIndex        =   2
         Top             =   1080
         Width           =   2652
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "4678;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTA03 
         Height          =   300
         Left            =   -73680
         TabIndex        =   1
         Top             =   720
         Width           =   2052
         VariousPropertyBits=   679493659
         MaxLength       =   12
         Size            =   "3619;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         Caption         =   "(按<F12>鍵可列出相類似的事務所以供選擇)"
         Height          =   252
         Left            =   -70920
         TabIndex        =   8
         Top             =   1080
         Width           =   3612
      End
      Begin VB.Label Label2 
         Caption         =   "事務所名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   7
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "代理人編號 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   6
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label3 
         Caption         =   "代理人名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   5
         Top             =   720
         Width           =   1212
      End
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   8520
      Top             =   540
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
            Picture         =   "frm030601.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030601.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030601.frx":0670
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030601.frx":084C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030601.frx":0B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030601.frx":0E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030601.frx":11A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030601.frx":14BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030601.frx":17D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030601.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030601.frx":1E10
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm030601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/10 Form2.0已修改 textTA03/textTA04/grdList
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Const MAX_FIELD = 5

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList(MAX_FIELD) As FIELDITEM

' 變數宣告區
Dim m_EditMode As Integer

' 第一筆資料的本所案號
Dim m_FirstTA As String
' 最後一筆資料的本所案號
Dim m_LastTA As String
' 目前正在顯示的本所案號
Dim m_CurrTA As String
'
Dim m_CurrSel As Integer

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Add By Sindy 2014/4/23 執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'2014/4/23 END


Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT TA02 FROM TAGENT " & _
            "WHERE TA02 = (SELECT MIN(To_number(TA02)) FROM TAGENT " & _
                          "WHERE TA01 = 'T') "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TA02")) = False Then: m_FirstTA = rsTmp.Fields("TA02")
   End If
   rsTmp.Close

   strSql = "SELECT TA02 FROM TAGENT " & _
            "WHERE TA02 = (SELECT MAX(To_number(TA02)) FROM TAGENT " & _
                          "WHERE TA01 = 'T') "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TA02")) = False Then: m_LastTA = rsTmp.Fields("TA02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   Select Case KeyCode
'      ' 新增
'      Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd, vbKeyEscape:
'         If m_EditMode = 0 Then
'            OnAction KeyCode
'         End If
'      Case vbKeyF9, vbKeyF10:
'         If m_EditMode <> 0 Then
'            OnAction KeyCode
'         End If
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
'      Case vbKeyEscape:
'         If m_EditMode = 0 Then
'            OnAction KeyCode
'         Else
'            OnAction vbKeyF10
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
            'Mark by Amy 2022/01/10 Form2.0 元件按Enter會觸發存檔
            'OnAction vbKeyF9
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

' Load Form
Private Sub Form_Load()
   tabCtrl.Tab = 1
   
   'Add By Sindy 2014/4/23 取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   '2014/4/23 END
   
   m_EditMode = 0
   MoveFormToCenter Me
   
   InitialField
   RefreshRange
   ShowFirstRecord
   SetCtrlReadOnly True
   UpdateGridList
   UpdateToolbarState
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "TA" & strTmp
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
   SetFieldNewData "TA01", "T"
   SetFieldNewData "TA02", Trim(textTA02)
   SetFieldNewData "TA03", Trim(textTA03)
   SetFieldNewData "TA04", Trim(textTA04)
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
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
   'Dim strSQL As String
   
   ' 檢查RecordSet的狀態
   'If m_Recordset.State <> adStateClosed Then
   '   m_Recordset.Close
   'End If
   ' 設定 Query 的命令
   'strSQL = "SELECT * FROM Tagent " & _
   '         "WHERE TA01 = 'T' " & _
   '         "ORDER BY TA02"
   ' 讀取資料庫
   'm_Recordset.CursorLocation = adUseClient
   'm_Recordset.Open strSQL, cnnConnection, adOpenDynamic
   
   ' 更新 GridList
   'UpdateGridList
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   textTA02 = Empty: textTA03 = Empty: textTA04 = Empty
   
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textTA02.Locked = bEnable
   textTA03.Locked = bEnable
   textTA04.Locked = bEnable
End Sub
' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textTA02.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ClearField
   
   If IsEmptyText(m_CurrTA) = True Then
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM TAGENT " & _
            "WHERE TA01 = 'T' AND " & _
                  "TA02 = '" & m_CurrTA & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("TA02")) Then: textTA02 = Trim(rsTmp.Fields("TA02"))
      If Not IsNull(rsTmp.Fields("TA03")) Then: textTA03 = Trim(rsTmp.Fields("TA03")) 'End If
      If Not IsNull(rsTmp.Fields("TA04")) Then: textTA04 = Trim(rsTmp.Fields("TA04")) 'End If
      UpdateFieldOldData rsTmp
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strTA01 As String, ByVal strTA02 As String)
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   
   If IsRecordExist(strTA01, strTA02) = True Then
      m_CurrTA = strTA02
   Else
      strSql = "SELECT TA02 FROM TAGENT " & _
            "WHERE TA02 = (SELECT MIN(To_number(TA02)) FROM TAGENT " & _
                          "WHERE TA01 = 'T' AND " & _
                                "To_number(TA02) > " & m_CurrTA & ") "
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("TA02")) = False Then
            m_CurrTA = rsTmp.Fields("TA02")
         End If
      Else
         rsTmp.Close
         Set rsTmp = Nothing
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrTA = m_FirstTA
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrTA = m_FirstTA Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT TA02 FROM TAGENT " & _
            "WHERE TA02 = (SELECT MAX(To_number(TA02)) FROM TAGENT " & _
                          "WHERE TA01 = 'T' AND " & _
                                "To_number(TA02) < " & m_CurrTA & ") "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TA02")) = False Then: m_CurrTA = rsTmp.Fields("TA02")
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
   
   If m_CurrTA = m_LastTA Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT TA02 FROM TAGENT " & _
            "WHERE TA02 = (SELECT MIN(To_number(TA02)) FROM TAGENT " & _
                          "WHERE TA01 = 'T' AND " & _
                                "To_number(TA02) > " & m_CurrTA & ") "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TA02")) = False Then: m_CurrTA = rsTmp.Fields("TA02")
   End If
   rsTmp.Close
   
   UpdateCtrlData
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrTA = m_LastTA
   
   UpdateCtrlData
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
         Me.tabCtrl.TabEnabled(1) = False
         tabCtrl.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         Me.tabCtrl.TabEnabled(1) = False
         tabCtrl.Tab = 0
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
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         If CheckDataValid() = True Then
            UpdateFieldNewData
            OnWork
            Me.tabCtrl.TabEnabled(1) = True
            UpdateToolbarState
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
                  Me.tabCtrl.TabEnabled(1) = True
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               Me.tabCtrl.TabEnabled(1) = True
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
   Set frm030601 = Nothing
End Sub

' 代理人編號
Private Sub textTA02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTA02) = False Then
      Select Case m_EditMode
         Case 1:
            If IsRecordExist("T", textTA02) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "代理人編號已存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTA02_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

' 代理人名稱
Private Sub textTA03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTA03) = False Then
      If StrLength(textTA03) > 12 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTA03_GotFocus
      End If
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTA03.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

'Modify by Amy 2022/01/10 原:Integer
Private Sub textTA04_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
   Dim strTemp As String
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   strTemp = textTA04
   If KeyCode = vbKeyF12 And (m_EditMode = 1 Or m_EditMode = 2) Then
      strTemp = textTA04.Text
      textTA04.Clear
      textTA04.Text = strTemp
      strSql = "SELECT DISTINCT TA04 FROM TAGENT " & _
               "WHERE TA01 = 'T' AND " & _
                     "TA04 LIKE '" & strTemp & "%' " & _
               "ORDER BY TA04 "
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            If IsNull(rsTmp.Fields("TA04")) = False Then
               If IsEmptyText(rsTmp.Fields("TA04")) = False Then
                  textTA04.AddItem rsTmp.Fields("TA04")
               End If
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      Set rsTmp = Nothing
      If textTA04.ListCount > 0 Then
         'Modify by Amy 2022/01/10 Form2.0 hWnd無此屬性
         'SendMessage textTA04.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
         textTA04.DropDown
      End If
   End If
End Sub

' 事務所名稱
Private Sub textTA04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTA04) = False Then
      If StrLength(textTA04) > 30 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "事務所名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTA04_GotFocus
      End If
   End If
   ' 事務所名稱空白時, 預設為代理人名稱 91.5.23 MODIFY BY SONIA
   If IsEmptyText(textTA04) = True Then
      textTA04 = textTA03
   End If
   '91.5.23 END
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTA04.IMEMode = 2
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
Private Function IsRecordExist(ByVal strTA01 As String, ByVal strTA02 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM TAGENT " & _
            "WHERE TA01 = '" & strTA01 & "' AND " & _
                  "TA02 = '" & strTA02 & "'"
                  
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
   Dim strTA01, strTA02 As String
   
   strTA01 = "T"
   strTA02 = textTA02
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strTA01, strTA02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO TAGENT ("
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
   
   Pub_SeekTbLog strSql 'Add By Sindy 2019/6/24
   cnnConnection.Execute strSql
   
   If (strTA02 < m_FirstTA) Or (strTA02 > m_LastTA) Then
      RefreshRange
   End If
   
   UpdateGridList
   ShowCurrRecord strTA01, strTA02
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
   Dim strTA01, strTA02 As String
   
   strTA01 = "T"
   strTA02 = m_CurrTA
   
   strSql = "UPDATE TAGENT SET "
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
                  "WHERE TA01 = '" & strTA01 & "' AND " & _
                        "TA02 = '" & strTA02 & "' "
   
   If bDifference = True Then
      Pub_SeekTbLog strSql 'Add By Sindy 2019/6/24
      cnnConnection.Execute strSql
      grdList_ModifiedItem strTA02
      ShowCurrRecord strTA01, strTA02
   End If
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strTA01, strTA02 As String
   
   strTA01 = "T"
   strTA02 = m_CurrTA
   
   strSql = "DELETE FROM TAGENT " & _
            "WHERE TA01 = '" & strTA01 & "' AND " & _
                  "TA02 = '" & strTA02 & "'"
   Pub_SeekTbLog strSql 'Add By Sindy 2019/6/24
   cnnConnection.Execute strSql
   
   If m_CurrTA = m_FirstTA Or m_CurrTA = m_LastTA Then
      RefreshRange
   End If
   
   grdList_DeleteItem strTA02
   ShowCurrRecord strTA01, strTA02
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   QueryRecord = False
   
   If IsRecordExist("T", textTA02) = True Then
      m_CurrTA = textTA02
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
         'Add By Cheng 2002/05/23
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
      
         AddRecord
         RefreshRange
      Case 2:
         'Add By Cheng 2002/05/23
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         ModRecord
      Case 3:
         DelRecord
         RefreshRange
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
   grdList.Cols = 4
   
   grdList.ColWidth(0) = 300
   grdList.row = 0
      
   grdList.col = 1
   grdList.Text = "代理人編號"
   grdList.ColWidth(1) = 1000
   grdList.ColAlignment(1) = flexAlignCenterCenter
   grdList.col = 2
   grdList.Text = "代理人名稱"
   grdList.ColWidth(2) = 2000
   grdList.ColAlignment(2) = flexAlignLeftCenter
   grdList.col = 3
   grdList.Text = "事務所名稱"
   grdList.ColWidth(3) = 3600
   grdList.ColAlignment(3) = flexAlignLeftCenter
End Sub

Private Sub UpdateGridList()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nRow As Integer
   
   grdList.Clear
   InitialGridList
   strSql = "SELECT * FROM TAGENT " & _
            "WHERE TA01 = 'T' " & _
            "ORDER BY To_number(TA02) " 'Modify By Sindy 2023/8/16 改用數字排序
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         grdList.Rows = grdList.Rows + 1
         nRow = grdList.Rows - 1
         If IsNull(rsTmp.Fields("TA02")) = False Then
            grdList.TextMatrix(nRow, 1) = Trim(rsTmp.Fields("TA02"))
         End If
         If IsNull(rsTmp.Fields("TA03")) = False Then
            grdList.TextMatrix(nRow, 2) = Trim(rsTmp.Fields("TA03"))
         End If
         If IsNull(rsTmp.Fields("TA04")) = False Then
            grdList.TextMatrix(nRow, 3) = Trim(rsTmp.Fields("TA04"))
         End If
         rsTmp.MoveNext
      Loop
      grdList.FixedRows = 1 'Add By Sindy 2022/5/2
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub grdList_ModifiedItem(ByVal strTA02 As String)
   Dim nRow As Integer
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   strSql = "SELECT * FROM TAGENT " & _
            "WHERE TA01 = 'T' AND " & _
                  "TA02 = '" & strTA02 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      For nRow = 1 To grdList.Rows - 1
         If grdList.TextMatrix(nRow, 1) = strTA02 Then
            If IsNull(rsTmp.Fields("TA02")) = False Then
               grdList.TextMatrix(nRow, 1) = Trim(rsTmp.Fields("TA02"))
            End If
            If IsNull(rsTmp.Fields("TA03")) = False Then
               grdList.TextMatrix(nRow, 2) = Trim(rsTmp.Fields("TA03"))
            End If
            If IsNull(rsTmp.Fields("TA04")) = False Then
               grdList.TextMatrix(nRow, 3) = Trim(rsTmp.Fields("TA04"))
            Else
               grdList.TextMatrix(nRow, 3) = ""
            End If
            Exit For
         End If
      Next nRow
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub grdList_DeleteItem(ByVal strTA02 As String)
   Dim nRow As Integer
   For nRow = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nRow, 1) = strTA02 Then
         grdList.RemoveItem nRow
         Exit For
      End If
   Next nRow
End Sub

Private Sub grdList_Click()
   Dim strTA01 As String
   Dim strTA02 As String
   Dim nIndex As Integer
   
   If grdList.row > 0 Then
      strTA01 = "T"
      strTA02 = grdList.TextMatrix(grdList.row, 1)
      ShowCurrRecord strTA01, strTA02
   End If
End Sub

Private Sub grdList_SelChange()
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

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2:
         ' 代理人編號不可空白
         If IsEmptyText(textTA02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入代理人編號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTA02.SetFocus
            GoTo EXITSUB
         End If
         ' 代理人名稱不可空白
         If IsEmptyText(textTA03) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入代理人名稱"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTA03.SetFocus
            GoTo EXITSUB
         End If
         ' 事務所名稱不可空白
         'If IsEmptyText(textTA04) = True Then
         '   strTit = "檢核資料"
         '   strMsg = "請輸入事務所名稱"
         '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '   textTA04.SetFocus
         '   GoTo ExitSub
         'End If
         ' 事務所名稱空白時, 預設為代理人名稱 91.5.23 MODIFY BY SONIA
         textTA04_Validate (False)
         '91.5.23 END
         'Add by Amy 2022/01/10檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
        If PUB_ChkUniText(Me, , True) = False Then
            GoTo EXITSUB
        End If
   End Select
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textTA02_GotFocus()
   InverseTextBox textTA02
End Sub

Private Sub textTA03_GotFocus()
   InverseTextBox textTA03
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTA03.IMEMode = 1
   OpenIme
End Sub

Private Sub textTA04_GotFocus()
   InverseTextBox textTA04
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTA04.IMEMode = 1
   OpenIme
End Sub

Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1, 4:
         textTA02.SetFocus
      Case 2:
         textTA03.SetFocus
   End Select
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textTA02.Enabled = True Then
   Cancel = False
   textTA02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTA03.Enabled = True Then
   Cancel = False
   textTA03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTA04.Enabled = True Then
   Cancel = False
   textTA04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function
