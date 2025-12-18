VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030618 
   BorderStyle     =   1  '單線固定
   Caption         =   "公報開拓資料維護"
   ClientHeight    =   5745
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   8955
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   4230
      TabIndex        =   19
      Top             =   720
      Width           =   4155
      Begin VB.TextBox Text2 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "(1.公報 2.撤三)"
         Top             =   60
         Width           =   1365
      End
      Begin VB.TextBox textTBD16 
         Height          =   285
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   20
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label9 
         Caption         =   "種類："
         Height          =   255
         Left            =   510
         TabIndex        =   22
         Top             =   60
         Width           =   615
      End
   End
   Begin VB.TextBox textTBOR02 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2070
      Width           =   2190
   End
   Begin VB.TextBox textTBD02 
      Height          =   285
      Left            =   1470
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1020
      Width           =   1332
   End
   Begin VB.TextBox textTBD01 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   6
      Top             =   750
      Width           =   732
   End
   Begin VB.TextBox textTBD03 
      Height          =   285
      Left            =   1470
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1320
      Width           =   525
   End
   Begin VB.TextBox textTBD04 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   17
      Top             =   1650
      Width           =   1332
   End
   Begin VB.TextBox textTBD03_2 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1350
      Width           =   4332
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   7710
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
            Picture         =   "frm030618.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030618.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030618.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030618.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030618.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030618.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030618.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030618.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030618.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030618.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030618.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   2325
      Left            =   30
      TabIndex        =   5
      Top             =   3360
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4101
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      HighLight       =   0
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
      _Band(0).Cols   =   1
   End
   Begin MSForms.TextBox textTBOR03 
      Height          =   300
      Left            =   1470
      TabIndex        =   2
      Top             =   2370
      Width           =   7395
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "13044;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.TextBox textTBOR04 
      Height          =   285
      Left            =   1470
      MaxLength       =   150
      TabIndex        =   3
      Top             =   2700
      Width           =   7395
   End
   Begin MSForms.TextBox textTBOR05 
      Height          =   300
      Left            =   1470
      TabIndex        =   4
      Top             =   3030
      Width           =   7395
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "13044;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   8910
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "序號："
      Height          =   180
      Left            =   900
      TabIndex        =   16
      Top             =   2100
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "商標權人中文："
      Height          =   180
      Left            =   180
      TabIndex        =   15
      Top             =   2430
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "商標權人英文："
      Height          =   180
      Left            =   180
      TabIndex        =   14
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "商標權人地址："
      Height          =   180
      Left            =   180
      TabIndex        =   13
      Top             =   3090
      Width           =   1260
   End
   Begin VB.Label Label5 
      Caption         =   "申請案號："
      Height          =   255
      Left            =   345
      TabIndex        =   11
      Top             =   1650
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "商標種類："
      Height          =   255
      Left            =   345
      TabIndex        =   10
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "公報卷期："
      Height          =   255
      Left            =   345
      TabIndex        =   9
      Top             =   750
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "審定號："
      Height          =   255
      Left            =   345
      TabIndex        =   8
      Top             =   1050
      Width           =   1095
   End
End
Attribute VB_Name = "frm030618"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/11 Form2.0 TextTBOR03/TextTBOR05/Grd1
'Memo By Sindy 2012/12/5 智權人員欄已修改
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
Dim m_FirstTM(2) As String
' 最後一筆資料的本所案號
Dim m_LastTM(2) As String
' 目前正在顯示的本所案號
Dim m_CurrTM(2) As String


Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT TBD01,TBD02,TBD03 FROM TMBULLETINDATA " & _
            "WHERE TBD16='1' AND TBD02 = (SELECT MIN(TBD02) FROM TMBULLETINDATA WHERE TBD16='1') " & _
              "AND TBD03 = (SELECT MIN(TBD03) FROM TMBULLETINDATA " & _
                            "WHERE TBD02 = (SELECT MIN(TBD02) FROM TMBULLETINDATA WHERE TBD16='1')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TBD02")) = False Then: m_FirstTM(0) = rsTmp.Fields("TBD02")
      If IsNull(rsTmp.Fields("TBD03")) = False Then: m_FirstTM(1) = rsTmp.Fields("TBD03")
   End If
   rsTmp.Close
   
   strSql = "SELECT TBD01,TBD02,TBD03 FROM TMBULLETINDATA " & _
            "WHERE TBD16='1' AND TBD02 = (SELECT MAX(TBD02) FROM TMBULLETINDATA WHERE TBD16='1') " & _
              "AND TBD03 = (SELECT MAX(TBD03) FROM TMBULLETINDATA " & _
                            "WHERE TBD02 = (SELECT MAX(TBD02) FROM TMBULLETINDATA WHERE TBD16='1')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TBD02")) = False Then: m_LastTM(0) = rsTmp.Fields("TBD02")
      If IsNull(rsTmp.Fields("TBD03")) = False Then: m_LastTM(1) = rsTmp.Fields("TBD03")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Load()
   m_EditMode = 0
   MoveFormToCenter Me
   
   InitialField
   RefreshRange
   ShowFirstRecord
   SetCtrlReadOnly True
   UpdateToolbarState
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 3 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "TBOR" & strTmp
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
   SetFieldNewData "TBOR03", textTBOR03
   SetFieldNewData "TBOR04", textTBOR04
   SetFieldNewData "TBOR05", textTBOR05
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
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

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   
   'textTBD01 = Empty
   textTBD02 = Empty
   textTBD03 = Empty
   textTBD03_2 = Empty
   textTBD04 = Empty
   textTBOR02 = Empty
   textTBOR03 = Empty
   textTBOR04 = Empty
   textTBOR05 = Empty
   GRD1.Clear: SetGrd (True)
'   ' 若是新增時卷期不變
'   If m_EditMode <> 1 Then
'      textTBD01 = Empty
'   End If
   
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textTBOR03.Locked = bEnable
   textTBOR04.Locked = bEnable
   textTBOR05.Locked = bEnable
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textTBD02.Locked = bEnable
   textTBD03.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ClearField
   strSql = "SELECT * FROM TMBULLETINDATA " & _
            "WHERE TBD02 = '" & m_CurrTM(0) & "' AND " & _
                  "TBD03 = '" & m_CurrTM(1) & "' AND " & _
                  "TBD16 = '1' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("TBD01")) Then: textTBD01 = rsTmp.Fields("TBD01")
      If Not IsNull(rsTmp.Fields("TBD02")) Then: textTBD02 = rsTmp.Fields("TBD02")
      If Not IsNull(rsTmp.Fields("TBD03")) Then: textTBD03 = rsTmp.Fields("TBD03")
      If Not IsNull(rsTmp.Fields("TBD04")) Then: textTBD04 = rsTmp.Fields("TBD04")
      If Not IsNull(rsTmp.Fields("TBD16")) Then: textTBD16 = rsTmp.Fields("TBD16") 'Add By Sindy 2018/12/17
      
      ' 商標種類
      If IsEmptyText(textTBD03) = False Then
         textTBD03_2 = GetTradeMarkName(textTBD03, 0)
      End If
      
      'Add By Sindy 2018/12/17
      If Pub_StrUserSt03 = "M51" Then
         Frame1.Visible = True
      Else
         Frame1.Visible = False
      End If
      '2018/12/17 END
      
      '商標權人檔
      rsTmp.Close
      'Modify By Sindy 2018/12/17 + TBOR07
      strSql = "SELECT TBOR02,TBOR03,TBOR04,TBOR05,TBOR07 FROM TMBULLETINOWNER WHERE TBOR01='" & textTBD02 & "' and TBOR06='" & textTBD03 & "' and TBOR07='1' order by TBOR02 ASC "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      Set GRD1.Recordset = rsTmp
      SetGrd
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub UpdateCtrlData2()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM TMBULLETINOWNER " & _
            "WHERE TBOR01 = '" & textTBD02 & "' AND " & _
                  "TBOR06 = '" & textTBD03 & "' AND " & _
                  "TBOR02 = '" & textTBOR02 & "' AND " & _
                  "TBOR07 = '1' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      'If Not IsNull(rsTmp.Fields("TBOR02")) Then: textTBOR02 = rsTmp.Fields("TBOR02")
      If Not IsNull(rsTmp.Fields("TBOR03")) Then: textTBOR03 = rsTmp.Fields("TBOR03")
      If Not IsNull(rsTmp.Fields("TBOR04")) Then: textTBOR04 = rsTmp.Fields("TBOR04")
      If Not IsNull(rsTmp.Fields("TBOR05")) Then: textTBOR05 = rsTmp.Fields("TBOR05")
      
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub SetGrd(Optional bolClear As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer, i
   
   'Modify By Sindy 2018/12/17 + 種類
   arrGridHeadText = Array("序號", "商標權人中文", "商標權人英文", "商標權人地址", "種類")
   If Pub_StrUserSt03 = "M51" Then
      arrGridHeadWidth = Array(450, 2500, 1500, 4000, 800)
   Else
      arrGridHeadWidth = Array(450, 2500, 1500, 4000, 0)
   End If
   GRD1.Cols = UBound(arrGridHeadText) + 1
   
   If bolClear = True Then
      GRD1.Rows = 2
   End If
   
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

'Private Sub grd1_SelChange()
'grd1.Visible = False
'If grd1.MouseRow <> 0 Then
'   textTBOR02 = grd1.TextMatrix(grd1.MouseRow, 0)
'   UpdateCtrlData2
'End If
'grd1.Visible = True
'End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   GRD1.col = nCol
   GRD1.row = nRow
   If nRow > 0 Then
      textTBOR02 = GRD1.TextMatrix(GRD1.row, 0)
      UpdateCtrlData2
   End If
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strTBD02 As String, ByVal strTBD03 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strTBD02, strTBD03) = True Then
      m_CurrTM(0) = strTBD02
      m_CurrTM(1) = strTBD03
   Else
      strSql = "SELECT * FROM TMBULLETINDATA " & _
               "WHERE TBD16='1' AND TBD02 = '" & m_CurrTM(0) & "' AND " & _
                     "TBD03 = (SELECT MIN(TBD03) FROM TMBULLETINDATA " & _
                               "WHERE TBD16='1' AND TBD02 = '" & m_CurrTM(0) & "' AND " & _
                                     "TBD03 > '" & m_CurrTM(1) & "') "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("TBD02")) = False Then: m_CurrTM(0) = rsTmp.Fields("TBD02")
         If IsNull(rsTmp.Fields("TBD03")) = False Then: m_CurrTM(1) = rsTmp.Fields("TBD03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT * FROM TMBULLETINDATA " & _
               "WHERE TBD16='1' AND TBD02 = (SELECT MIN(TBD02) FROM TMBULLETINDATA " & _
                               "WHERE TBD16='1' AND TBD02 > '" & m_CurrTM(0) & "') AND " & _
                     "TBD03 = (SELECT MIN(TBD03) FROM TMBULLETINDATA " & _
                               "WHERE TBD02 = (SELECT MIN(TBD02) FROM TMBULLETINDATA " & _
                                               "WHERE TBD16='1' AND TBD02 > '" & m_CurrTM(0) & "')) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("TBD02")) = False Then: m_CurrTM(0) = rsTmp.Fields("TBD02")
         If IsNull(rsTmp.Fields("TBD03")) = False Then: m_CurrTM(1) = rsTmp.Fields("TBD03")
      Else
         rsTmp.Close
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrTM(0) = m_FirstTM(0)
   m_CurrTM(1) = m_FirstTM(1)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrTM(0) = m_FirstTM(0) And m_CurrTM(1) = m_FirstTM(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM TMBULLETINDATA " & _
            "WHERE TBD16='1' AND TBD02 = '" & m_CurrTM(0) & "' AND " & _
                  "TBD03 = (SELECT MAX(TBD03) FROM TMBULLETINDATA " & _
                            "WHERE TBD16='1' AND TBD02 = '" & m_CurrTM(0) & "' AND " & _
                                  "TBD03 < '" & m_CurrTM(1) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TBD02")) = False Then: m_CurrTM(0) = rsTmp.Fields("TBD02")
      If IsNull(rsTmp.Fields("TBD03")) = False Then: m_CurrTM(1) = rsTmp.Fields("TBD03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT * FROM TMBULLETINDATA " & _
            "WHERE TBD16='1' AND TBD02 = (SELECT MAX(TBD02) FROM TMBULLETINDATA " & _
                            "WHERE TBD16='1' AND TBD02 < '" & m_CurrTM(0) & "') AND " & _
                  "TBD03 = (SELECT MAX(TBD03) FROM TMBULLETINDATA " & _
                            "WHERE TBD02 = (SELECT MAX(TBD02) FROM TMBULLETINDATA " & _
                                            "WHERE TBD16='1' AND TBD02 < '" & m_CurrTM(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TBD02")) = False Then: m_CurrTM(0) = rsTmp.Fields("TBD02")
      If IsNull(rsTmp.Fields("TBD03")) = False Then: m_CurrTM(1) = rsTmp.Fields("TBD03")
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
   
   If m_CurrTM(0) = m_LastTM(0) And m_CurrTM(1) = m_LastTM(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM TMBULLETINDATA " & _
            "WHERE TBD16='1' AND TBD02 = '" & m_CurrTM(0) & "' AND " & _
                  "TBD03 = (SELECT MIN(TBD03) FROM TMBULLETINDATA " & _
                            "WHERE TBD16='1' AND TBD02 = '" & m_CurrTM(0) & "' AND " & _
                                  "TBD03 > '" & m_CurrTM(1) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TBD02")) = False Then: m_CurrTM(0) = rsTmp.Fields("TBD02")
      If IsNull(rsTmp.Fields("TBD03")) = False Then: m_CurrTM(1) = rsTmp.Fields("TBD03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT * FROM TMBULLETINDATA " & _
            "WHERE TBD16='1' AND TBD02 = (SELECT MIN(TBD02) FROM TMBULLETINDATA " & _
                            "WHERE TBD16='1' AND TBD02 > '" & m_CurrTM(0) & "') AND " & _
                  "TBD03 = (SELECT MIN(TBD03) FROM TMBULLETINDATA " & _
                            "WHERE TBD02 = (SELECT MIN(TBD02) FROM TMBULLETINDATA " & _
                                            "WHERE TBD16='1' AND TBD02 > '" & m_CurrTM(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TBD02")) = False Then: m_CurrTM(0) = rsTmp.Fields("TBD02")
      If IsNull(rsTmp.Fields("TBD03")) = False Then: m_CurrTM(1) = rsTmp.Fields("TBD03")
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrTM(0) = m_LastTM(0)
   m_CurrTM(1) = m_LastTM(1)
   
   UpdateCtrlData
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   tlbar.Buttons(1).Visible = False
   tlbar.Buttons(3).Visible = False
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         tlbar.Buttons(1).Enabled = True
         tlbar.Buttons(2).Enabled = True
         tlbar.Buttons(3).Enabled = True
         tlbar.Buttons(4).Enabled = True
         tlbar.Buttons(6).Enabled = True
         tlbar.Buttons(7).Enabled = True
         tlbar.Buttons(8).Enabled = True
         tlbar.Buttons(9).Enabled = True
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
      ' 新增
      Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_EditMode = 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyCode = 0
            'OnAction vbKeyF9 'Mrak by Amy 2022/01/11 改Form2.0 不使用Enter 存檔
            KeyCode = 0
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
            KeyCode = 0
         Else
            OnAction vbKeyF10
            KeyCode = 0
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
         SetKeyReadOnly False
         UpdateToolbarState
'         OnPrepareAddRecord
         If IsRecordExist2(textTBD02, textTBD03, textTBOR02) = True Then
            strTit = "檢核資料"
            strMsg = "該筆記錄已經存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTBOR03.SetFocus
         Else
            SetInputEntry
         End If
      ' 修改
      Case vbKeyF3:
         If textTBOR02 = "" Then
            MsgBox "請先點選一筆欲修改的商標權人資料！", vbOKOnly, "檢核資料"
            Exit Sub
         End If
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
         textTBD03 = "1" 'Add By Sindy 2018/8/1 預設為1
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
         ' Moemo 2022/01/11 原 將所有欄位的內容更新到欄位串列中的欄位內容項目 搬至onWork
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
   Set frm030618 = Nothing
End Sub

'Add By Sindy 2014/7/7
Private Sub textTBD02_Validate(Cancel As Boolean)
   If textTBD02 <> "" Then
      If textTBD03 = "" Then
         textTBD03 = "1"
         Call textTBD03_LostFocus
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
Private Function IsRecordExist(ByVal strTBD02 As String, ByVal strTBD03 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM TMBULLETINDATA " & _
            "WHERE TBD02 = '" & strTBD02 & "' AND " & _
                  "TBD03 = '" & strTBD03 & "' AND " & _
                  "TBD16 = '1' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist2(ByVal strTBOR01 As String, ByVal strTBOR06 As String, ByVal strTBOR02 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist2 = False
   strSql = "SELECT * FROM TMBULLETINOWNER " & _
            "WHERE TBOR01 = '" & strTBOR01 & "' AND " & _
                  "TBOR06 = '" & strTBOR06 & "' AND " & _
                  "TBOR02 = '" & strTBOR02 & "' AND " & _
                  "TBOR07 = '1' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsRecordExist2 = True
   Else
      IsRecordExist2 = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 新增記錄
Private Function AddRecord() As Boolean
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strTBD02 As String
   Dim strTBD03 As String
   Dim strTBOR02 As String
   
   strTBD02 = textTBD02
   strTBD03 = textTBD03
   strTBOR02 = textTBOR02
   
   AddRecord = False
   
   ' 檢查記錄是否已存在
   If IsRecordExist2(strTBD02, strTBD03, strTBOR02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If

   bFirst = True
   bDifference = False
   strSql = "INSERT INTO TMBULLETINOWNER ("
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
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
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

   cnnConnection.Execute strSql
   AddRecord = True

   If ((strTBD02 & strTBD03) < (m_FirstTM(0) & m_FirstTM(1))) Or ((strTBD02 & strTBD03) > (m_LastTM(0) & m_LastTM(1))) Then
      RefreshRange
   End If
   
   ShowCurrRecord strTBD02, strTBD03
EXITSUB:
End Function

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
   Dim strTBD02 As String
   Dim strTBD03 As String
   Dim strTBOR02 As String
   Dim strTMBM07 As String
   
   strTBD02 = textTBD02
   strTBD03 = textTBD03
   strTBOR02 = textTBOR02
   
   strSql = "UPDATE TMBULLETINOWNER SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
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
   
   strSql = strSql & " " & _
            "WHERE TBOR01 = '" & strTBD02 & "' AND " & _
                  "TBOR06 = '" & strTBD03 & "' AND " & _
                  "TBOR02 = '" & strTBOR02 & "' "
   If bDifference = True Then
      cnnConnection.Execute strSql
      
      'Add By Sindy 2017/4/26 更新商標公報檔
      strExc(0) = "select distinct tbd01 from tmbulletindata where tbd16='1'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      strTMBM07 = ""
      If intI = 1 Then
         strTMBM07 = RsTemp.Fields(0)
      End If
      strSql = "UPDATE TMBULLETIN SET "
      If strTBOR02 = 1 Then strSql = strSql & "TMBM09='" & ChgSQL(textTBOR03) & "' "
      If strTBOR02 = 2 Then strSql = strSql & "TMBM10='" & ChgSQL(textTBOR03) & "' "
      If strTBOR02 = 3 Then strSql = strSql & "TMBM11='" & ChgSQL(textTBOR03) & "' "
      If strTBOR02 = 4 Then strSql = strSql & "TMBM12='" & ChgSQL(textTBOR03) & "' "
      If strTBOR02 = 5 Then strSql = strSql & "TMBM13='" & ChgSQL(textTBOR03) & "' "
      If strTBOR02 = 6 Then strSql = strSql & "TMBM14='" & ChgSQL(textTBOR03) & "' "
      If strTBOR02 = 7 Then strSql = strSql & "TMBM15='" & ChgSQL(textTBOR03) & "' "
      If strTBOR02 = 8 Then strSql = strSql & "TMBM16='" & ChgSQL(textTBOR03) & "' "
      If strTBOR02 = 9 Then strSql = strSql & "TMBM17='" & ChgSQL(textTBOR03) & "' "
      If strTBOR02 = 10 Then strSql = strSql & "TMBM18='" & ChgSQL(textTBOR03) & "' "
      strSql = strSql & "WHERE TMBM01 = '" & strTBD02 & "' AND TMBM02 = '" & strTBD03 & "' AND TMBM07 = '" & strTMBM07 & "'"
      cnnConnection.Execute strSql
      '2017/4/26 END
      
      ShowCurrRecord strTBD02, strTBD03
   End If
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strTBD02 As String
   Dim strTBD03 As String
   Dim strTBOR02 As String
   
   strTBD02 = textTBD02
   strTBD03 = textTBD03
   strTBOR02 = textTBOR02
   
   strSql = "DELETE FROM TMBULLETINOWNER " & _
            "WHERE TBOR01 = '" & strTBD02 & "' AND " & _
                  "TBOR06 = '" & strTBD03 & "' AND " & _
                  "TBOR02 = '" & strTBOR02 & "' AND " & _
                  "TBOR07 = '1' "
   
   cnnConnection.Execute strSql
      
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strTBD02 = m_LastTM(0) And strTBD03 = m_LastTM(1)) Or (strTBD02 = m_FirstTM(0) And strTBD03 = m_FirstTM(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strTBD02, strTBD03
   
EXITSUB:
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strSql As String
   Dim nIndex As Integer
   Dim nPos As Integer
   Dim bFind As Boolean
   Dim strTBD02 As String
   Dim strTBD03 As String
   
   strTBD02 = textTBD02
   strTBD03 = textTBD03
   
   If IsRecordExist(strTBD02, strTBD03) = True Then
      m_CurrTM(0) = strTBD02
      m_CurrTM(1) = strTBD03
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
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            'Add by Amy 2022/01/11 因From2.0 會將UniCode資料改成?,若於此之前就先執行 UpdateFieldNewData 導致存入的資料應有?,反而沒有了
            UpdateFieldNewData '將所有欄位的內容更新到欄位串列中的欄位內容項目
            If AddRecord = True Then
               RefreshRange
               SetKeyReadOnly True
            Else
               Exit Sub
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            UpdateFieldNewData 'Add by Amy 2022/01/11
            ModRecord
            SetKeyReadOnly True
         Else
            GoTo EXITSUB
         End If
      Case 3:
         DelRecord
         RefreshRange
         SetKeyReadOnly True
      Case 4:
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
         SetKeyReadOnly True
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1:
         If IsEmptyText(textTBD02) Then
            textTBD02.SetFocus
         Else
            textTBD03.SetFocus
         End If
      Case 2: textTBOR03.SetFocus
      Case 4: textTBD02.SetFocus
   End Select
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTemp As String
Dim nPos As Integer
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim strFreeAgentCode As String
   
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2:
'         ' 公報卷期
'         If IsEmptyText(textTBD01) = True Then
'            strTit = "檢核資料"
'            strMsg = "請輸入公報卷期"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTBD01.SetFocus
'            GoTo EXITSUB
'         End If
         ' 審定號不可空白
         If IsEmptyText(textTBD02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入審定號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTBD02.SetFocus
            GoTo EXITSUB
         End If
         ' 商標種類不可空白
         If IsEmptyText(textTBD03) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入商標種類"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTBD03.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
   
   CheckDataValid = True
   
EXITSUB:

   Exit Function
ErrorHandler:
   cnnConnection.RollbackTrans
   MsgBox "(" & Err.Number & ")" & Err.Description
End Function

' 公報卷期
Private Sub textTBD01_GotFocus()
   InverseTextBox textTBD01
End Sub
Private Sub textTBD01_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub
Private Sub textTBD01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rs1 As New ADODB.Recordset
   
   Cancel = False
   If IsEmptyText(textTBD01) = False Then
      If IsNumeric(textTBD01) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "公報卷期只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTBD01_GotFocus
         Exit Sub
      End If
   End If
End Sub

'審定號數
Private Sub textTBD02_GotFocus()
   InverseTextBox textTBD02
End Sub
Private Sub textTBD02_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

' 商標種類
Private Sub textTBD03_GotFocus()
   InverseTextBox textTBD03
End Sub
Private Sub textTBD03_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   textTBD03_2 = Empty
   If IsEmptyText(textTBD03) = False Then
      textTBD03_2 = GetTradeMarkName(textTBD03, 0)
      If IsEmptyText(textTBD03_2) = True Then
         Select Case m_EditMode
            Case 1, 4:
               strTit = "檢核資料"
               strMsg = "商標種類不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTBD03.SetFocus
               GoTo EXITSUB
         End Select
      End If
      If m_EditMode = 1 Then
         If IsEmptyText(textTBD02) = False Then
            If IsRecordExist(textTBD02, textTBD03) = True Then
               strTit = "檢核資料"
               strMsg = "該筆資料已經存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTBD02.SetFocus
               GoTo EXITSUB
            End If
         End If
      End If
   End If
EXITSUB:
End Sub

'商標權人中文
Private Sub textTBOR03_GotFocus()
   OpenIme
   InverseTextBox textTBOR03
End Sub
Private Sub textTBOR03_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTBOR03, textTBOR03.MaxLength) = False Then
      Cancel = True
      textTBOR03_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

'商標權人英文
Private Sub textTBOR04_GotFocus()
   CloseIme
   InverseTextBox textTBOR04
End Sub
Private Sub textTBOR04_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTBOR04, textTBOR04.MaxLength) = False Then
      Cancel = True
      textTBOR04_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

'商標權人地址
Private Sub textTBOR05_GotFocus()
   OpenIme
   InverseTextBox textTBOR05
End Sub
Private Sub textTBOR05_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTBOR05, textTBOR05.MaxLength) = False Then
      Cancel = True
      textTBOR05_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

'Add by Amy 2022/01/11檢查畫面的 TextBox是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

If IsRecordExist(textTBD02, textTBD03) = False Then
   MsgBox "該筆公報審定號資料不存在", vbOKOnly, "檢核資料"
   textTBD02.SetFocus
   Exit Function
End If

If Me.textTBOR03.Enabled = True Then
   Cancel = False
   textTBOR03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTBOR04.Enabled = True Then
   Cancel = False
   textTBOR04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTBOR05.Enabled = True Then
   Cancel = False
   textTBOR05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function
