VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010017 
   BorderStyle     =   1  '單線固定
   Caption         =   "櫃台每日作業統計數輸入"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6960
   Begin VB.TextBox textOI14 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   2490
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1740
      Width           =   465
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢"
      Height          =   285
      Left            =   3030
      TabIndex        =   28
      Top             =   2130
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1890
      MaxLength       =   7
      TabIndex        =   27
      Top             =   2130
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   720
      MaxLength       =   7
      TabIndex        =   26
      Top             =   2130
      Width           =   1065
   End
   Begin VB.TextBox textOI07 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   3420
      MaxLength       =   4
      TabIndex        =   7
      Top             =   1740
      Width           =   465
   End
   Begin VB.TextBox textOI06 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   1500
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1740
      Width           =   465
   End
   Begin VB.TextBox textOI05 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   4410
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1380
      Width           =   465
   End
   Begin VB.TextBox textOI04 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   3420
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1380
      Width           =   465
   End
   Begin VB.TextBox textOI01 
      Alignment       =   2  '置中對齊
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   0
      Top             =   660
      Width           =   1035
   End
   Begin VB.TextBox textOI02 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   1500
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1380
      Width           =   465
   End
   Begin VB.TextBox textOI03 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   2460
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1380
      Width           =   465
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2265
      Left            =   90
      TabIndex        =   10
      Top             =   2430
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   3995
      _Version        =   393216
      Rows            =   3
      Cols            =   1
      FixedRows       =   2
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6990
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":1DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010017.frx":20F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6960
      _ExtentX        =   12277
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
         NumButtons      =   12
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
            Style           =   3
            MixedState      =   -1  'True
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "投資"
      Height          =   180
      Left            =   2100
      TabIndex        =   29
      Top             =   1770
      Width           =   360
   End
   Begin VB.Line Line13 
      X1              =   1620
      X2              =   2220
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   180
      Left            =   150
      TabIndex        =   25
      Top             =   2190
      Width           =   540
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   6900
      X2              =   6900
      Y1              =   1010
      Y2              =   2070
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   5790
      X2              =   5790
      Y1              =   1010
      Y2              =   2070
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   4950
      X2              =   4950
      Y1              =   1010
      Y2              =   2070
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   3960
      X2              =   3960
      Y1              =   1010
      Y2              =   2070
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   3000
      X2              =   3000
      Y1              =   1010
      Y2              =   2070
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   1010
      Y2              =   2070
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   1080
      X2              =   1080
      Y1              =   1010
      Y2              =   2070
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   90
      X2              =   90
      Y1              =   1010
      Y2              =   2070
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   90
      X2              =   6900
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   90
      X2              =   6900
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   90
      X2              =   6900
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   90
      X2              =   6900
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label lblMTotal2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6765
      TabIndex        =   24
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label lblMTotal1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6765
      TabIndex        =   23
      Top             =   1440
      Width           =   105
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "本  月  累  計"
      ForeColor       =   &H00FF00FF&
      Height          =   180
      Left            =   5880
      TabIndex        =   22
      Top             =   1080
      Width           =   990
   End
   Begin VB.Label lblTotal2 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5625
      TabIndex        =   21
      Top             =   1800
      Width           =   105
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "其他"
      Height          =   180
      Left            =   3030
      TabIndex        =   20
      Top             =   1770
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "智權"
      Height          =   180
      Left            =   1110
      TabIndex        =   19
      Top             =   1800
      Width           =   360
   End
   Begin VB.Label lblTotal1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5625
      TabIndex        =   18
      Top             =   1440
      Width           =   105
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "合        計"
      Height          =   180
      Left            =   5010
      TabIndex        =   17
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "其他"
      Height          =   180
      Left            =   4020
      TabIndex        =   16
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "中所"
      Height          =   180
      Left            =   1110
      TabIndex        =   15
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "客戶來訪："
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "長途電話："
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日        期："
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "南所"
      Height          =   180
      Left            =   2070
      TabIndex        =   9
      Top             =   1410
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "高所"
      Height          =   180
      Left            =   3030
      TabIndex        =   8
      Top             =   1440
      Width           =   360
   End
End
Attribute VB_Name = "frm010017"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/23 Form2.0已修改(無需修改)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/23 日期欄已修改
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
Dim m_EditMode As Integer
Dim m_SubMode As Integer
' 第一筆資料的本所案號
Dim m_FirstKEY(1) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(1) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(1) As String
Dim m_TheMCount1 As Long
Dim m_TheMCount2 As Long

Private Sub cmdok_Click()
Dim Cancel As Boolean
Cancel = False
If Text1(0) = "" Or Text1(1) = "" Then MsgBox "日期條件不可以空白!!", vbExclamation, "操作錯誤!!": Exit Sub
Text1_Validate 0, Cancel
If Cancel = True Then Exit Sub
Text1_Validate 1, Cancel
If Cancel = True Then Exit Sub
If Val(Text1(0)) > Val(Text1(1)) Then MsgBox "前面日期應小於後面日期!!", vbExclamation, "操作錯誤!!": Exit Sub
GetAllData
End Sub

Private Sub Form_Initialize()
'edit by nickc 2007/11/01 加入一個欄位
'ReDim m_FieldList(7) As FieldItem
ReDim m_FieldList(8) As FIELDITEM
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
m_bInsert = IsUserHasRightOfFunction("frm010017", strAdd, False)
m_bUpdate = IsUserHasRightOfFunction("frm010017", strEdit, False)
textOI01.Text = strSrvDate(2)
m_CurrKEY(0) = ChangeTStringToWString(textOI01)
InitialField
RefreshRange
UpdateToolbarState
SetCtrlReadOnly True
SetGrd
UpdateCtrlData
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm010017 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth, arrGridHeadText1
   Dim iRow As Integer
'edit by nickc 2007/11/01 加入一個欄位
'   arrGridHeadText = Array("", "長途電話", "長途電話", "長途電話", "長途電話", "客戶來訪", "客戶來訪")
'   arrGridHeadText1 = Array("日期", "中所", "南所", "高所", "其他", "智權", "其他")
'   arrGridHeadWidth = Array(1200, 900, 900, 900, 900, 900, 900)
   arrGridHeadText = Array("", "長途電話", "長途電話", "長途電話", "長途電話", "客戶來訪", "客戶來訪", "客戶來訪")
   arrGridHeadText1 = Array("日期", "中所", "南所", "高所", "其他", "智權", "投資", "其他")
   arrGridHeadWidth = Array(1200, 900, 900, 900, 900, 900, 900, 900)
   grd1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1.Cols - 1
      grd1.row = 0
      grd1.col = iRow
      grd1.Text = arrGridHeadText(iRow)
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1.CellAlignment = flexAlignCenterCenter
      grd1.row = 1
      grd1.col = iRow
      grd1.Text = arrGridHeadText1(iRow)
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1.CellAlignment = flexAlignCenterCenter
   Next
    grd1.MergeCells = flexMergeRestrictRows
    grd1.MergeRow(0) = True
    grd1.MergeCol(0) = True
    If grd1.Rows > 2 Then
        grd1.Visible = False
        For iRow = 2 To grd1.Rows - 1
           grd1.row = iRow
           grd1.col = 0
           grd1.CellAlignment = flexAlignCenterCenter
           grd1.col = 1
           grd1.CellAlignment = flexAlignRightBottom
           grd1.col = 2
           grd1.CellAlignment = flexAlignRightBottom
           grd1.col = 3
           grd1.CellAlignment = flexAlignRightBottom
           grd1.col = 4
           grd1.CellAlignment = flexAlignRightBottom
           grd1.col = 5
           grd1.CellAlignment = flexAlignRightBottom
           grd1.col = 6
           grd1.CellAlignment = flexAlignRightBottom
           'add by nickc 2007/11/01
           grd1.col = 7
           grd1.CellAlignment = flexAlignRightBottom
        Next
        grd1.Visible = True
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
InverseTextBox Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
If Trim(Text1(Index)) <> "" Then
    If CheckIsTaiwanDate(Text1(Index), False) = False Then
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
End Sub

Private Sub textOI01_GotFocus()
InverseTextBox textOI01
End Sub

Private Sub textOI01_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textOI01_Validate(Cancel As Boolean)
If Trim(textOI01) <> "" And m_EditMode = 1 Then
    If CheckIsTaiwanDate(textOI01, False) = False Then
        Cancel = True
        MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
    ElseIf ChkWorkDay(ChangeTStringToWString(textOI01)) = False Then
        Cancel = True
        MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
    End If
End If
End Sub

Private Sub textOI02_GotFocus()
InverseTextBox textOI02
End Sub

Private Sub textOI02_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textOI02_KeyUp(KeyCode As Integer, Shift As Integer)
   lblTotal1.Caption = Trim(Val(textOI02) + Val(textOI03) + Val(textOI04) + Val(textOI05))
   lblMTotal1.Caption = Trim(Val(m_TheMCount1) + Val(lblTotal1.Caption))
End Sub

Private Sub textOI02_Validate(Cancel As Boolean)
If Trim(textOI02) <> "" Then
    If InStr(1, textOI02, ".") = 0 Then
        If IsNumeric(textOI02) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textOI03_GotFocus()
InverseTextBox textOI03
End Sub

Private Sub textOI03_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textOI03_KeyUp(KeyCode As Integer, Shift As Integer)
   lblTotal1.Caption = Trim(Val(textOI02) + Val(textOI03) + Val(textOI04) + Val(textOI05))
   lblMTotal1.Caption = Trim(Val(m_TheMCount1) + Val(lblTotal1.Caption))
End Sub

Private Sub textOI03_Validate(Cancel As Boolean)
If Trim(textOI03) <> "" Then
    If InStr(1, textOI03, ".") = 0 Then
        If IsNumeric(textOI03) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textOI04_GotFocus()
InverseTextBox textOI04
End Sub

Private Sub textOI04_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textOI04_KeyUp(KeyCode As Integer, Shift As Integer)
   lblTotal1.Caption = Trim(Val(textOI02) + Val(textOI03) + Val(textOI04) + Val(textOI05))
   lblMTotal1.Caption = Trim(Val(m_TheMCount1) + Val(lblTotal1.Caption))
End Sub

Private Sub textOI04_Validate(Cancel As Boolean)
If Trim(textOI04) <> "" Then
    If InStr(1, textOI04, ".") = 0 Then
        If IsNumeric(textOI04) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textOI05_GotFocus()
InverseTextBox textOI05
End Sub

Private Sub textOI05_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textOI05_KeyUp(KeyCode As Integer, Shift As Integer)
   lblTotal1.Caption = Trim(Val(textOI02) + Val(textOI03) + Val(textOI04) + Val(textOI05))
   lblMTotal1.Caption = Trim(Val(m_TheMCount1) + Val(lblTotal1.Caption))
End Sub

Private Sub textOI05_Validate(Cancel As Boolean)
If Trim(textOI05) <> "" Then
    If InStr(1, textOI05, ".") = 0 Then
        If IsNumeric(textOI05) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textOI06_GotFocus()
InverseTextBox textOI06
End Sub

Private Sub textOI06_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textOI06_KeyUp(KeyCode As Integer, Shift As Integer)
   'edit by nickc 2007/11/01
   'lblTotal2.Caption = Trim(Val(textOI06) + Val(textOI07))
   lblTotal2.Caption = Trim(Val(textOI06) + Val(textOI07) + Val(textOI14))
   lblMTotal2.Caption = Trim(Val(m_TheMCount2) + Val(lblTotal2.Caption))
End Sub

Private Sub textOI06_Validate(Cancel As Boolean)
If Trim(textOI06) <> "" Then
    If InStr(1, textOI06, ".") = 0 Then
        If IsNumeric(textOI06) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textOI07_GotFocus()
InverseTextBox textOI07
End Sub

Private Sub textOI07_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textOI07_KeyUp(KeyCode As Integer, Shift As Integer)
   'edit by nickc 2007/11/01
   'lblTotal2.Caption = Trim(Val(textOI06) + Val(textOI07))
   lblTotal2.Caption = Trim(Val(textOI06) + Val(textOI07) + Val(textOI14))
   lblMTotal2.Caption = Trim(Val(m_TheMCount2) + Val(lblTotal2.Caption))
End Sub

Private Sub textOI07_Validate(Cancel As Boolean)
If Trim(textOI07) <> "" Then
    If InStr(1, textOI07, ".") = 0 Then
        If IsNumeric(textOI07) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textOI14_GotFocus()
InverseTextBox textOI14
End Sub

Private Sub textOI14_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textOI14_KeyUp(KeyCode As Integer, Shift As Integer)
   lblTotal2.Caption = Trim(Val(textOI06) + Val(textOI07) + Val(textOI14))
   lblMTotal2.Caption = Trim(Val(m_TheMCount2) + Val(lblTotal2.Caption))
End Sub

Private Sub textOI14_Validate(Cancel As Boolean)
If Trim(textOI14) <> "" Then
    If InStr(1, textOI14, ".") = 0 Then
        If IsNumeric(textOI14) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
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
      ' 第一筆
      Case 4: OnAction vbKeyHome
      ' 前一筆
      Case 5: OnAction vbKeyPageUp
      ' 後一筆
      Case 6: OnAction vbKeyPageDown
      ' 最後一筆
      Case 7: OnAction vbKeyEnd
      ' 確定
      Case 9: OnAction vbKeyF9
      ' 取消
      Case 10: OnAction vbKeyF10
      ' 離開
      Case 12: OnAction vbKeyEscape
   End Select
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
         'If IsRecordExist(m_CurrKEY(0)) = True Then MsgBox "已有今日資料，請修改！", vbInformation, "操作錯誤！": Exit Sub
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         'If IsRecordExist(m_CurrKEY(0)) = False Then MsgBox "今日沒有資料，請新增！", vbInformation, "操作錯誤！": Exit Sub
         UpdateCtrlData
         m_EditMode = 2
         SetCtrlReadOnly False
         textOI01.Locked = True
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
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
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
                RefreshRange
            Else
                Exit Function
            End If
      Case 2: '修改
            If TxtValidate = False Then Exit Function
            ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To 7
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "OI" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 1 '數值型態
   Next nIndex
   'add by nickc 2007/11/01
    strTmp = Format(14, "00")
    m_FieldList(7).fiName = "OI" & strTmp
    m_FieldList(7).fiOldData = Empty
    m_FieldList(7).fiNewData = Empty
    m_FieldList(7).fiType = 1 '數值型態
End Sub

'抓當日所有資料
Private Sub GetAllData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
'edit by nickc 2007/11/01
'   strSQL = "SELECT sqldatet(oi01),oi02,oi03,oi04,oi05,oi06,oi07 FROM otherinput " & _
                 "WHERE oi01>=" & Val(ChangeTStringToWString(Text1(0))) & " and oi01<=" & Val(ChangeTStringToWString(Text1(1))) & " order by oi01 asc "
   strSql = "SELECT sqldatet(oi01),oi02,oi03,oi04,oi05,oi06,oi14,oi07 FROM otherinput " & _
                 "WHERE oi01>=" & Val(ChangeTStringToWString(Text1(0))) & " and oi01<=" & Val(ChangeTStringToWString(Text1(1))) & " order by oi01 asc "
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        Set grd1.Recordset = rsTmp
        rsTmp.Close
    SetGrd
EXITSUB:
   Set rsTmp = Nothing
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
         Toolbar1.Buttons(4).Enabled = True
         Toolbar1.Buttons(5).Enabled = True
         Toolbar1.Buttons(6).Enabled = True
         Toolbar1.Buttons(7).Enabled = True
         Toolbar1.Buttons(9).Enabled = False
         Toolbar1.Buttons(10).Enabled = False
         Toolbar1.Buttons(12).Enabled = True
         ' 新增
      Case 1, 2
         Toolbar1.Buttons(1).Enabled = False
         Toolbar1.Buttons(2).Enabled = False
         Toolbar1.Buttons(4).Enabled = False
         Toolbar1.Buttons(5).Enabled = False
         Toolbar1.Buttons(6).Enabled = False
         Toolbar1.Buttons(7).Enabled = False
         Toolbar1.Buttons(9).Enabled = True
         Toolbar1.Buttons(10).Enabled = True
         Toolbar1.Buttons(12).Enabled = False
   End Select
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textOI01.Locked = bEnable
   textOI02.Locked = bEnable
   textOI03.Locked = bEnable
   textOI04.Locked = bEnable
   textOI05.Locked = bEnable
   textOI06.Locked = bEnable
   textOI07.Locked = bEnable
   'add by nickc 2007/11/01
   textOI14.Locked = bEnable
End Sub

Private Sub ClearField()
   Dim nIndex As Integer

   textOI02 = Empty
   textOI03 = Empty
   textOI04 = Empty
   textOI05 = Empty
   textOI06 = Empty
   textOI07 = Empty
   'add by nickc 2007/11/01
   textOI14 = Empty
   
   For nIndex = 0 To 8 'edit by nickc 2007/11/01  7
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textOI01.SetFocus
      Case 2: textOI02.SetFocus
   End Select
End Sub

Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM otherinput " & _
            "WHERE oi01 = " & Val(m_CurrKEY(0)) & " "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("oi01")) = False Then: textOI01 = ChangeWStringToTString("" & rsTmp.Fields("oi01"))
      If IsNull(rsTmp.Fields("oi02")) = False Then: textOI02 = rsTmp.Fields("oi02")
      If IsNull(rsTmp.Fields("oi03")) = False Then: textOI03 = rsTmp.Fields("oi03")
      If IsNull(rsTmp.Fields("oi04")) = False Then: textOI04 = rsTmp.Fields("oi04")
      If IsNull(rsTmp.Fields("oi05")) = False Then: textOI05 = rsTmp.Fields("oi05")
      If IsNull(rsTmp.Fields("oi06")) = False Then: textOI06 = rsTmp.Fields("oi06")
      If IsNull(rsTmp.Fields("oi07")) = False Then: textOI07 = rsTmp.Fields("oi07")
      'add by nickc 2007/11/01
      If IsNull(rsTmp.Fields("oi14")) = False Then: textOI14 = rsTmp.Fields("oi14")
   End If
   ' 更新暫存區的資料
   UpdateFieldOldData rsTmp
   GetAllCounts
   lblTotal1.Caption = Trim(Val(textOI02) + Val(textOI03) + Val(textOI04) + Val(textOI05))
   'edit by nickc 2007/11/01
   'lblTotal2.Caption = Trim(Val(textOI06) + Val(textOI07))
   lblTotal2.Caption = Trim(Val(textOI06) + Val(textOI07) + Val(textOI14))
   lblMTotal1.Caption = Trim(Val(m_TheMCount1) + Val(lblTotal1.Caption))
   lblMTotal2.Caption = Trim(Val(m_TheMCount2) + Val(lblTotal2.Caption))
   rsTmp.Close
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textOI01.Enabled = True Then
   Cancel = False
   textOI01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOI02.Enabled = True Then
   Cancel = False
   textOI02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOI03.Enabled = True Then
   Cancel = False
   textOI03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOI04.Enabled = True Then
   Cancel = False
   textOI04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textOI05.Enabled = True Then
   Cancel = False
   textOI05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textOI06.Enabled = True Then
   Cancel = False
   textOI06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textOI07.Enabled = True Then
   Cancel = False
   textOI07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'add by nickc 2007/11/01
If Me.textOI14.Enabled = True Then
   Cancel = False
   textOI14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
TxtValidate = True
End Function

Private Sub UpdateFieldNewData()
   '若新增資料
   SetFieldNewData "OI01", ChangeTStringToWString(textOI01)
   SetFieldNewData "OI02", textOI02
   SetFieldNewData "OI03", textOI03
   SetFieldNewData "OI04", textOI04
   SetFieldNewData "OI05", textOI05
   SetFieldNewData "OI06", textOI06
   SetFieldNewData "OI07", textOI07
   'add by nickc 2007/11/01
   SetFieldNewData "OI14", textOI14
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
   Dim strOI01 As String
   
   AddRecord = False
   strOI01 = ChangeTStringToWString(textOI01)
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strOI01) = True Then
      strTit = "新增資料"
      strMsg = "該天紀錄已存在，請修改日期或改用修改"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Exit Function
   End If
   
   bFirst = True
   strSql = "INSERT INTO otherinput ("
   For nIndex = 0 To 7 'edit by nickc 2007/11/01 6
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
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
   strSql = strSql & ",oi08,oi09,oi10) "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To 7 'edit by nickc 2007/11/01 6
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
   strSql = strSql & ",'" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI'))) "
On Error GoTo ErrHand
    cnnConnection.BeginTrans
   
   cnnConnection.Execute strSql
   m_CurrKEY(0) = strOI01
   If (Val(strOI01) < Val(m_FirstKEY(0))) Or (Val(strOI01) > Val(m_LastKEY(0))) Then
      RefreshRange
   End If
   
    cnnConnection.CommitTrans
    
   AddRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox " 新增失敗！" & vbCrLf & Err.Description
    Resume Next
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
   Dim strOI01 As String
   
   ModRecord = False
   
   strOI01 = m_CurrKEY(0)
   strSql = "UPDATE otherinput SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To 8 'edit by nickc 2007/11/01 7
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

   strSql = strSql & ",oi11='" & strUserNum & "',oi12=to_number(to_char(sysdate,'YYYYMMDD')),oi13=to_number(to_char(sysdate,'HH24MI')) " & _
                  "WHERE oi01 = " & strOI01 & "  "
On Error GoTo ErrHand
   If bDifference = True Then
      cnnConnection.BeginTrans
      
      cnnConnection.Execute strSql

      cnnConnection.CommitTrans
      
      
   End If
    ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)
    Resume Next
End Function

Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To 7 'edit by nickc 2007/11/01 6
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

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To 7 'edit by nickc 2007/11/01 6
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

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM otherinput " & _
            "WHERE oi01 = " & strKEY01 & " "
                  
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

Sub GetAllCounts()
'取得目前累計
m_TheMCount1 = 0
m_TheMCount2 = 0
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
'edit by nickc 2007/11/01
'strSQL = "SELECT sum(nvl(oi02,0)+nvl(oi03,0)+nvl(oi04,0)+nvl(oi05,0)),sum(nvl(oi06,0)+nvl(oi07,0)) FROM otherinput " & _
              "WHERE substr(oi01,1,6) = " & Val(Mid(m_CurrKEY(0), 1, 6)) & " and oi01<>" & Val(m_CurrKEY(0)) & " "
strSql = "SELECT sum(nvl(oi02,0)+nvl(oi03,0)+nvl(oi04,0)+nvl(oi05,0)),sum(nvl(oi06,0)+nvl(oi07,0)+nvl(oi14,0)) FROM otherinput " & _
              "WHERE substr(oi01,1,6) = " & Val(Mid(m_CurrKEY(0), 1, 6)) & " and oi01<>" & Val(m_CurrKEY(0)) & " "
     If rsTmp.State = 1 Then rsTmp.Close
     rsTmp.CursorLocation = adUseClient
     rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
     If Not rsTmp.EOF And Not rsTmp.BOF Then
           m_TheMCount1 = Val(CheckStr(rsTmp.Fields(0)))
           m_TheMCount2 = Val(CheckStr(rsTmp.Fields(1)))
     End If
     rsTmp.Close
     Set rsTmp = Nothing
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
   Else
      strSql = "SELECT oi01 FROM otherinput " & _
               "WHERE oi01 = '" & m_CurrKEY(0) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("oi01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("oi01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT min(oi01) as oi01 FROM otherinput "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If Not rsTmp.EOF And Not rsTmp.BOF Then
         If IsNull(rsTmp.Fields("oi01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("oi01")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If Val(m_CurrKEY(0)) = Val(m_FirstKEY(0)) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT max(oi01) as oi01 FROM otherinput " & _
            "WHERE oi01 < '" & m_CurrKEY(0) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("oi01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("oi01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT min(oi01) as oi01 FROM otherinput "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("oi01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("oi01")
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
   
   If Val(m_CurrKEY(0)) = Val(m_LastKEY(0)) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT min(oi01) as oi01 FROM otherinput " & _
            "WHERE oi01 > '" & m_CurrKEY(0) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("oi01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("oi01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT max(oi01) as oi01 FROM otherinput  "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("oi01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("oi01")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
  
   UpdateCtrlData
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT min(oi01) as oi01 FROM otherinput  "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("oi01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("oi01")
   End If
   rsTmp.Close

   strSql = "SELECT max(oi01) as oi01 FROM otherinput  "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("oi01")) = False Then: m_LastKEY(0) = rsTmp.Fields("oi01")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub
