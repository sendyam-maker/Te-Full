VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060106 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內公報代理人資料維護"
   ClientHeight    =   5760
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   9150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9150
   Begin TabDlg.SSTab tabCtrl 
      Height          =   4995
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   690
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8811
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   420
      TabCaption(0)   =   "單筆"
      TabPicture(0)   =   "frm04060106.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "textTA05"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "textTA03"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboTA04"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "textTA04"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "textTA02"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm04060106.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdList"
      Tab(1).ControlCount=   1
      Begin VB.TextBox textTA02 
         Height          =   270
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   0
         Top             =   420
         Width           =   1212
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   4635
         Left            =   -74940
         TabIndex        =   11
         Top             =   300
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   8176
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
      Begin MSForms.TextBox textTA04 
         Height          =   300
         Left            =   1500
         TabIndex        =   2
         Top             =   1140
         Width           =   2055
         VariousPropertyBits=   671107099
         Size            =   "3625;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboTA04 
         Height          =   300
         Left            =   1500
         TabIndex        =   3
         Top             =   1140
         Width           =   2325
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "4101;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTA03 
         Height          =   300
         Left            =   1500
         TabIndex        =   1
         Top             =   780
         Width           =   2295
         VariousPropertyBits=   671107099
         MaxLength       =   12
         Size            =   "4048;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "建檔時公告日 :"
         Height          =   252
         Left            =   180
         TabIndex        =   9
         Top             =   1500
         Width           =   1212
      End
      Begin VB.Label Label3 
         Caption         =   "事務所名稱 :"
         Height          =   252
         Left            =   180
         TabIndex        =   8
         Top             =   1140
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "代理人名稱 :"
         Height          =   252
         Left            =   180
         TabIndex        =   7
         Top             =   780
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "代理人編號 :"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   420
         Width           =   1095
      End
      Begin MSForms.TextBox textTA05 
         Height          =   300
         Left            =   1500
         TabIndex        =   4
         Top             =   1500
         Width           =   1215
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
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
      Left            =   8580
      Top             =   600
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
            Picture         =   "frm04060106.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060106.frx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060106.frx":0670
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060106.frx":084C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060106.frx":0B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060106.frx":0E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060106.frx":11A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060106.frx":14BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060106.frx":17D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060106.frx":1AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm04060106.frx":1E10
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm04060106"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/23 改成Form2.0 (textTA03,textTA04,cboTA04,textTA05)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Const MAX_FIELD = 5

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

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

Dim m_CurrSel As Integer
'Add By Sindy 2014/4/23 執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
'2014/4/23 END
Dim m_TA04 As String 'Add By Sindy 2014/9/2

Private Sub Form_Load()
   ' 先顯示多筆查詢的畫面
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
   QueryDB
   ShowFirstRecord
   SetCtrlReadOnly True
   UpdateToolbarState
End Sub

' 按下按鍵
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   
   'Add By Sindy 2014/9/1 當focus在備註欄時按enter鍵維持換行功能而不是存檔功能
   If KeyCode = vbKeyReturn And UCase(Me.ActiveControl.Name) = UCase("textTA04") Then
      Exit Sub
   End If
   '2014/9/1 END
'   Select Case KeyCode
'      ' 新增
'      Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd, vbKeyEscape:
'         If m_EditMode = 0 Then
'            OnAction KeyCode
'         End If
'      Case vbKeyF9, vbKeyF10, vbKeyReturn:
'         If KeyCode = vbKeyReturn Then: KeyCode = vbKeyF9
'         If m_EditMode <> 0 Then
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
         'Modify By Sindy 2023/3/14 Mark:輸入法的選單出現，在選字時按向下箭到3，當按下Enter要確定選此字，程式就直接存檔了(人員反應，覺得麻煩~)
         'If m_EditMode <> 0 Then
         If m_EditMode = 4 Then
         '2023/3/14 END
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
   'Modify by Morgan 2011/5/17 改排序
   'strSql = "SELECT * FROM Tagent " & _
            "WHERE TA01 = 'P' " & _
            "ORDER BY TA02"
   strSql = "SELECT * FROM Tagent " & _
            "WHERE TA01 = 'P' " & _
            "ORDER BY substrb('    '||TA02,-4)"
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
   SetFieldNewData "TA01", "P"
   SetFieldNewData "TA02", textTA02
   SetFieldNewData "TA03", textTA03
   SetFieldNewData "TA04", textTA04
   If IsEmpty(textTA05) = True Then
      SetFieldNewData "TA05", ""
   Else
      SetFieldNewData "TA05", ChangeTStringToWString(textTA05)
   End If
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
   textTA02 = Empty: textTA03 = Empty: textTA04 = Empty: textTA05 = Empty
   m_TA04 = Empty 'Add By Sindy 2014/9/2
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textTA02.Locked = bEnable: textTA03.Locked = bEnable: textTA04.Locked = bEnable: textTA05.Locked = bEnable
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textTA02.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   ' 判斷是否有記錄存在且記錄指標的位置是有資料存在的
   If m_Recordset.RecordCount <= 0 Then: GoTo EXITSUB: 'End If
   If m_Recordset.BOF = True Then: GoTo EXITSUB: 'End If
   If m_Recordset.EOF = True Then: GoTo EXITSUB: 'End If
   
   ClearField
   If IsNull(m_Recordset.Fields("TA02")) = False Then
      textTA02 = m_Recordset.Fields("TA02")
   End If
   If IsNull(m_Recordset.Fields("TA03")) = False Then
      textTA03 = m_Recordset.Fields("TA03")
   End If
   m_TA04 = Empty 'Add By Sindy 2014/9/2
   If IsNull(m_Recordset.Fields("TA04")) = False Then
      textTA04 = m_Recordset.Fields("TA04")
      m_TA04 = m_Recordset.Fields("TA04") 'Add By Sindy 2014/9/2
   End If
   If IsNull(m_Recordset.Fields("TA05")) = False Then
      textTA05 = ChangeWStringToTString(m_Recordset.Fields("TA05"))
   End If
EXITSUB:
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If GrdList.Rows > 0 Then
      If nSel <= 0 Then: nSel = 1
      If nSel > GrdList.Rows - 1 Then: nSel = GrdList.Rows - 1
      GrdList.row = nSel
      grdList_SelChange
      grdList_ShowSelection
   End If
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strTA01 As String, ByVal strTA02 As String)
   Dim bFind As Boolean
   If m_Recordset.RecordCount > 0 Then
      bFind = False
      m_Recordset.MoveFirst
      While m_Recordset.EOF = False And bFind = False
         If strTA02 = m_Recordset.Fields("TA02") And strTA02 = m_Recordset.Fields("TA02") Then
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
         textTA02.Locked = True
         textTA03.Locked = True
         textTA04.Locked = True
         textTA05.Locked = True
      ' 新增
      Case 1:
         textTA02.Locked = False
         textTA03.Locked = False
         textTA04.Locked = False
         textTA05.Locked = False
      ' 修改
      Case 2:
         textTA02.Locked = True
         textTA03.Locked = False
         textTA04.Locked = False
         textTA05.Locked = False
      ' 查詢
      Case 4:
         textTA02.Locked = False
         textTA03.Locked = True
         textTA04.Locked = True
         textTA05.Locked = True
      Case Else:
         textTA02.Locked = True
         textTA03.Locked = True
         textTA04.Locked = True
         textTA05.Locked = True
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
   ' 更新欄位的狀態
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
         m_EditMode = 4
         ClearField
         SetCtrlReadOnly True
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
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         If CheckDataValid() = True Then
            UpdateFieldNewData
            OnWork
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

Private Sub tabCtrl_GotFocus()
   textTA02.SetFocus
End Sub

Private Sub textTA02_GotFocus()
   InverseAll textTA02
End Sub

Private Sub textTA02_KeyPress(KeyAscii As Integer)
   KeyAscii = UCase(KeyAscii)
End Sub

' 代理人編號
Private Sub textTA02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmpty(textTA02) = False Then
      ' 離開欄位時檢查是否該代號有重覆
      If m_EditMode = 1 Then
         ' 檢查記錄是否已存在
         If IsRecordExist("P", textTA02) = True Then
            Cancel = True
            strTit = "新增資料"
            strMsg = "該筆記錄已存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End If
      End If
   End If
End Sub

Private Sub textTA03_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'textTA03.IMEMode = "1"
   OpenIme
   InverseAll textTA03
End Sub

Private Sub textTA03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If StrLength(textTA03) > 12 Then
      Cancel = True
      strMsg = "代理人名稱太長"
      strTit = "檢核資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   
EXITSUB:
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: textTA03.IMEMode = "2"
   If Cancel = False Then CloseIme
End Sub

'Modify By Sindy 2014/9/1
Private Sub textTA04_Change()
Dim strTemp As String
Dim rsTmp As ADODB.Recordset
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If UCase(Me.ActiveControl.Name) = UCase("textTA04") Then
                  cboTA04.Clear
         Set rsTmp = New ADODB.Recordset
         strSql = "SELECT DISTINCT TA04 FROM Tagent " & _
                  "WHERE TA01 = 'P' AND " & _
                        "TA04 LIKE '%" & textTA04 & "%' " & _
                  "ORDER BY TA04"
         ' 讀取資料庫
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenDynamic
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            Do While rsTmp.EOF = False
               If IsNull(rsTmp.Fields("TA04")) = False Then
                  If IsEmpty(rsTmp.Fields("TA04")) = False Then
                     cboTA04.AddItem rsTmp.Fields("TA04")
                  End If
               End If
               rsTmp.MoveNext
            Loop
         End If
         rsTmp.Close
         Set rsTmp = Nothing
         
         If cboTA04.ListCount > 0 Then 'Added by Morgan 2021/12/23
            If cboTA04.ListCount > 1 Or (cboTA04.ListCount = 1 And cboTA04.List(0) <> textTA04) Then
               'Modified by Morgan 2021/12/23 Form2.0無hWnd屬性,可改直接用增加的 DropDown 方法
               'SendMessage cboTA04.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
               cboTA04.DropDown
               'end 2021/12/23
            End If
         End If
      End If
   End If
End Sub

Private Sub textTA04_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'textTA04.IMEMode = "1"
   CloseIme
   textTA04.SelStart = 0
   textTA04.SelLength = Len(textTA04.Text)
End Sub

'Private Sub textTA04_KeyDown(KeyCode As Integer, Shift As Integer)
'   Dim strTemp As String
'   Dim rsTmp As ADODB.Recordset
'
'   'If KeyCode = vbKeyF12 Then
'   If KeyCode = vbkeyenter Then
'      strTemp = textTA04.Text
'      textTA04.Clear
'      textTA04 = strTemp
'
'      Set rsTmp = New ADODB.Recordset
'      strSql = "SELECT DISTINCT TA04 FROM Tagent " & _
'               "WHERE TA01 = 'P' AND " & _
'                     "TA04 LIKE '%" & textTA04.Text & "%' " & _
'               "ORDER BY TA04"
'      ' 讀取資料庫
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSql, cnnConnection, adOpenDynamic
'      If rsTmp.RecordCount > 0 Then
'         rsTmp.MoveFirst
'         Do While rsTmp.EOF = False
'            If IsNull(rsTmp.Fields("TA04")) = False Then
'               If IsEmpty(rsTmp.Fields("TA04")) = False Then
'                  textTA04.AddItem rsTmp.Fields("TA04")
'               End If
'            End If
'            rsTmp.MoveNext
'         Loop
'      End If
'      rsTmp.Close
'      Set rsTmp = Nothing
'      If textTA04.ListCount > 0 Then
'         SendMessage textTA04.hwnd, CB_SHOWDROPDOWN, True, ByVal 0&
'      End If
'   End If
'End Sub

Private Sub textTA04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If StrLength(textTA04) > 30 Then
      Cancel = True
      strMsg = "事務所名稱太長"
      strTit = "檢核資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   ' 事務所名稱空白時, 預設為代理人名稱 91.5.23 modify by sonia
   If IsEmpty(textTA04) = True Then
      textTA04 = textTA03
   End If
   '91.5.23 end
EXITSUB:
   'edit by nickc 2007/07/11 切換輸入法改用API
   'If Cancel = False Then: textTA04.IMEMode = "2"
   If Cancel = False Then CloseIme
End Sub

Private Sub cboTA04_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'cboTA04.IMEMode = "1"
   CloseIme
   cboTA04.SelStart = 0
   cboTA04.SelLength = Len(cboTA04.Text)
End Sub

Private Sub cboTA04_Click()
   textTA04 = cboTA04.Text
   textTA04.SetFocus
End Sub

Private Sub textTA05_GotFocus()
   InverseAll textTA05
End Sub

Private Sub textTA05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmpty(textTA05) = False Then
      If CheckIsTaiwanDate(textTA05, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的日期"
         strTit = "日期不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
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
   strSql = "SELECT * FROM Tagent " & _
            "WHERE TA01 = '" & strTA01 & "' AND " & _
                  "TO_NUMBER(TA02) = '" & Val(strTA02) & "' "
                  
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
Set frm04060106 = Nothing
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
   Dim strTA01, strTA02 As String
   
   strTA01 = "P"
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
   strSql = "INSERT INTO Tagent ("
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
      Pub_SeekTbLog strSql 'Add By Sindy 2019/6/24
      cnnConnection.Execute strSql
      Call UpdTA04 'Add By Sindy 2014/9/2
      QueryDB
      ShowCurrRecord strTA01, strTA02
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
   Dim strTA01, strTA02 As String
   
   strTA01 = "P"
   strTA02 = textTA02
   strSql = "UPDATE Tagent SET "
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
                  "WHERE TA01 = '" & strTA01 & "' AND " & _
                     "TA02 = '" & strTA02 & "'"
   
   If bDifference = True Then
      Pub_SeekTbLog strSql 'Add By Sindy 2019/6/24
      cnnConnection.Execute strSql
      Call UpdTA04 'Add By Sindy 2014/9/2
      QueryDB
      ShowCurrRecord strTA01, strTA02
   End If

End Sub

'Add By Sindy 2014/9/2
Private Sub UpdTA04()
   '事務所名稱欄的值從無到有時,以代理人編號+建檔時公告日更新所有公報資料的事務所名稱欄
   If m_TA04 = "" And textTA04 <> "" Then
      strSql = "update TPBulletin set TPB08='" & textTA04 & "' where TPB07='" & textTA02 & "' and TPB03>=" & DBDATE(Val(textTA05))
      cnnConnection.Execute strSql
      strSql = "update TPBulletin_sonia set TPB08='" & textTA04 & "' where TPB07='" & textTA02 & "' and TPB03>=" & DBDATE(Val(textTA05))
      cnnConnection.Execute strSql
      strSql = "update TPGazette set TPG08='" & textTA04 & "' where TPG07='" & textTA02 & "' and TPG03>=" & DBDATE(Val(textTA05))
      cnnConnection.Execute strSql
   End If
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim nIndex As Integer
   Dim strSql As String
   Dim nSel As Integer
   Dim strTA01, strTA02 As String
   
   nSel = 1
   For nIndex = 1 To GrdList.Rows - 1
      If GrdList.TextMatrix(nIndex, 1) = textTA02 Then: nSel = nIndex
   Next nIndex
   
   strTA01 = "P"
   strTA02 = textTA02
   
   strSql = "DELETE FROM Tagent " & _
            "WHERE TA01 = '" & strTA01 & "' AND " & _
                  "TA02 = '" & strTA02 & "'"
   
   Pub_SeekTbLog strSql 'Add By Sindy 2019/6/24
   cnnConnection.Execute strSql
   
   QueryDB
   'ShowFirstRecord
   grdList_SetSelection nSel
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strSql As String
   Dim nIndex As Index
   Dim nPos
   Dim bFind As Boolean
   Dim strTA01, strTA02 As String
   
   strTA01 = "P"
   strTA02 = textTA02
   nPos = m_Recordset.AbsolutePosition
   QueryRecord = False
   bFind = False
   m_Recordset.MoveFirst
   While (m_Recordset.EOF <> True) And (bFind = False)
      If m_Recordset.Fields("TA01") = strTA01 And m_Recordset.Fields("TA02") = strTA02 Then
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
         AddRecord
      Case 2:
         ModRecord
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
End Sub

Private Sub InitialGridList()
   GrdList.Clear
   GrdList.Rows = 1
   GrdList.Cols = 5
   GrdList.ColWidth(0) = 300
   GrdList.row = 0
   GrdList.col = 1
   GrdList.Text = "代理人編號"
   GrdList.ColWidth(1) = 1200
   GrdList.ColAlignment(1) = flexAlignCenterCenter
   GrdList.col = 2
   GrdList.Text = "代理人名稱"
   GrdList.ColWidth(2) = 1200
   GrdList.ColAlignment(2) = flexAlignLeftCenter
   GrdList.col = 3
   GrdList.Text = "事務所名稱"
   GrdList.ColWidth(3) = 1200
   GrdList.ColAlignment(3) = flexAlignLeftCenter
   GrdList.col = 4
   GrdList.Text = "建檔時公告日"
   GrdList.ColWidth(4) = 1400
   GrdList.ColAlignment(4) = flexAlignCenterCenter
End Sub

Private Sub UpdateGridList()
   Dim strTA01, strTA02, strTA03 As String
   Dim nRow As Integer
   
   GrdList.Clear
   InitialGridList
   
   If IsRecordsetCorrect = True Then
      strTA01 = m_Recordset.Fields("TA01")
      strTA02 = m_Recordset.Fields("TA02")
            
      GrdList.Rows = m_Recordset.RecordCount + 1
      m_Recordset.MoveFirst
      nRow = 1
      While m_Recordset.EOF <> True
         GrdList.row = nRow
         
         GrdList.col = 1
         If IsNull(m_Recordset.Fields("TA02")) = False Then
            GrdList.Text = m_Recordset.Fields("TA02")
         End If
         
         GrdList.col = 2
         If IsNull(m_Recordset.Fields("TA03")) = False Then
            GrdList.Text = m_Recordset.Fields("TA03")
         End If
         
         GrdList.col = 3
         If IsNull(m_Recordset.Fields("TA04")) = False Then
            GrdList.Text = m_Recordset.Fields("TA04")
         End If
         
         GrdList.col = 4
         If IsNull(m_Recordset.Fields("TA05")) = False Then
            GrdList.Text = ChangeWStringToTString(m_Recordset.Fields("TA05"))
         End If
         
         nRow = nRow + 1
         m_Recordset.MoveNext
      Wend
      
      ShowCurrRecord strTA01, strTA02
      GrdList.FixedRows = 1 'Add By Sindy 2022/5/2
   End If
End Sub

Private Sub grdList_Click()
   Dim strTA01, strTA02, strTA03 As String
   Dim nIndex As Integer
   
   If GrdList.row > 0 Then
      GrdList.col = 1
      strTA01 = "P"
      strTA02 = GrdList.Text
      ShowCurrRecord strTA01, strTA02
   End If
End Sub

Private Sub grdList_SelChange()
   Dim strTA01, strTA02, strTA03 As String
   Dim nIndex As Integer
   
   If GrdList.row > 0 Then
      strTA01 = "P"
      strTA02 = GrdList.TextMatrix(GrdList.row, 1)
      ShowCurrRecord strTA01, strTA02
   End If
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = GrdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = GrdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < GrdList.Rows Then
      GrdList.row = m_CurrSel
      GrdList.col = 1
      If GrdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To GrdList.Cols - 1
            GrdList.col = nCol
            If GrdList.CellBackColor <> &H80000005 Then: GrdList.CellBackColor = &H80000005
            If GrdList.CellForeColor <> &H80000008 Then: GrdList.CellForeColor = &H80000008
         Next nCol
      End If
      GrdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < GrdList.Rows Then
      GrdList.row = m_CurrSel
      GrdList.col = 1
      For nCol = 1 To GrdList.Cols - 1
         GrdList.col = nCol
         GrdList.CellBackColor = &H8000000D
         GrdList.CellForeColor = &H80000005
      Next nCol
      GrdList.col = 0
   End If
EXITSUB:
End Sub

Private Sub SetEntryFocus()
   Select Case m_EditMode
      Case 1, 4:
         If textTA02.Locked = False Then
            textTA02.SetFocus
         End If
      Case 2:
         If textTA03.Locked = False Then
            textTA03.SetFocus
         End If
   End Select
End Sub

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

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   'Added by Morgan 2021/12/23 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/23
   
   Select Case m_EditMode
      Case 1, 2:
         ' 代理人編號不可為空白
         If IsEmpty(textTA02) = True Then
            strTit = "資料檢核"
            strMsg = "代理人編號不可為空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         ' 代理人名稱不可為空白
         If IsEmpty(textTA03) = True Then
            strTit = "資料檢核"
            strMsg = "代理人名稱不可為空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         ' 事務所名稱空白時, 預設為代理人名稱91.5.23 modify by sonia
         If IsEmpty(textTA04) = True Then
            textTA04 = textTA03
         End If
         '91.5.23 end
      Case 4:
         ' 代理人編號不可為空白
         If IsEmpty(textTA02) = True Then
            strTit = "資料檢核"
            strMsg = "代理人編號不可為空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
   End Select
   
   ' 代理人名稱不可超過12個字元
   If StrLength(textTA03) > 12 Then
      strTit = "資料檢核"
      strMsg = "代理人名稱太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   ' 事務所名稱不可超過30個字元
   If StrLength(textTA04) > 30 Then
      strTit = "資料檢核"
      strMsg = "事務所名稱太長"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As Object)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub
