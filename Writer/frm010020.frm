VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010020 
   BorderStyle     =   1  '單線固定
   Caption         =   "新案立卷數量輸入"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   6510
   Begin VB.TextBox textCF10 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   1050
      MaxLength       =   8
      TabIndex        =   9
      Top             =   2730
      Width           =   855
   End
   Begin VB.TextBox textCF11 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   3150
      MaxLength       =   8
      TabIndex        =   10
      Top             =   2730
      Width           =   855
   End
   Begin VB.TextBox textCF08 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   1050
      MaxLength       =   8
      TabIndex        =   7
      Top             =   2310
      Width           =   855
   End
   Begin VB.TextBox textCF09 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   3150
      MaxLength       =   8
      TabIndex        =   8
      Top             =   2310
      Width           =   855
   End
   Begin VB.TextBox textCF06 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   1050
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox textCF07 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   3150
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox textCF04 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   1050
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1500
      Width           =   855
   End
   Begin VB.TextBox textCF05 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   3150
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1500
      Width           =   855
   End
   Begin VB.TextBox textCF03 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   3150
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox textCF01 
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
   Begin VB.TextBox textCF02 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   1050
      MaxLength       =   8
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   600
      MaxLength       =   7
      TabIndex        =   11
      Top             =   3270
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1770
      MaxLength       =   7
      TabIndex        =   12
      Top             =   3270
      Width           =   1065
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢"
      Height          =   285
      Left            =   2910
      TabIndex        =   13
      Top             =   3270
      Width           =   1035
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   2955
      Left            =   0
      TabIndex        =   14
      Top             =   3600
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   5212
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6510
      _ExtentX        =   11483
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8220
      Top             =   3990
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
            Picture         =   "frm010020.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010020.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010020.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010020.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010020.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010020.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010020.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010020.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010020.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010020.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010020.frx":1DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010020.frx":20F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   3090
      X2              =   3090
      Y1              =   1020
      Y2              =   3120
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   90
      X2              =   4200
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "著作權："
      Height          =   180
      Left            =   120
      TabIndex        =   27
      Top             =   1530
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "顧　問："
      Height          =   180
      Left            =   2070
      TabIndex        =   26
      Top             =   1530
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "ＣＦＣ："
      Height          =   180
      Left            =   2040
      TabIndex        =   25
      Top             =   1950
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "法　務："
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   1980
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "ＦＣＬ："
      Height          =   180
      Left            =   2070
      TabIndex        =   23
      Top             =   2370
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "ＦＣＴ："
      Height          =   180
      Left            =   150
      TabIndex        =   22
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ＣＦＬ："
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   2820
      Width           =   720
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   90
      X2              =   4200
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   90
      X2              =   4200
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   90
      X2              =   4200
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   1980
      X2              =   1980
      Y1              =   1020
      Y2              =   3120
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   990
      X2              =   990
      Y1              =   1020
      Y2              =   3120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   90
      X2              =   90
      Y1              =   1005
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4200
      X2              =   4200
      Y1              =   1020
      Y2              =   3120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   90
      X2              =   4200
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   90
      X2              =   4200
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專　利："
      Height          =   180
      Left            =   2070
      TabIndex        =   19
      Top             =   1110
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "商　標："
      Height          =   180
      Left            =   150
      TabIndex        =   18
      Top             =   1110
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日        期："
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   180
      Left            =   30
      TabIndex        =   16
      Top             =   3330
      Width           =   540
   End
   Begin VB.Line Line4 
      X1              =   1470
      X2              =   2070
      Y1              =   3390
      Y2              =   3390
   End
   Begin VB.Label Label6 
      Caption         =   "Ｂ、Ｄ、Ｓ、Ｍ、Ｆ："
      Height          =   360
      Left            =   2070
      TabIndex        =   21
      Top             =   2730
      Width           =   990
   End
End
Attribute VB_Name = "frm010020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/23 Form2.0已修改(無需修改)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
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
ReDim m_FieldList(11) As FIELDITEM
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
m_bInsert = IsUserHasRightOfFunction("frm010020", strAdd, False)
m_bUpdate = IsUserHasRightOfFunction("frm010020", strEdit, False)
textCF01.Text = strSrvDate(2)
InitialField
RefreshRange
UpdateToolbarState
SetCtrlReadOnly True
SetGrd
m_CurrKEY(0) = strSrvDate(1)
UpdateCtrlData
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm010020 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   arrGridHeadText = Array("日期", "商標", "專利", "著作權", "顧問", "法務", "CFC", "FCT", "FCL", "CFL", "B、D、S、M、F")
   arrGridHeadWidth = Array(900, 900, 900, 900, 900, 900, 900, 900, 900, 900, 900)
   grd1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To grd1.Cols - 1
      grd1.row = 0
      grd1.col = iRow
      grd1.Text = arrGridHeadText(iRow)
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1.CellAlignment = flexAlignCenterCenter
   Next
    If grd1.Rows >= 2 Then
        grd1.Visible = False
        For iRow = 1 To grd1.Rows - 1
           grd1.row = iRow
           grd1.col = 0
           grd1.CellAlignment = flexAlignCenterCenter
           grd1.col = 1
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 2
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 3
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 4
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 5
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 6
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 7
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 8
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 9
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 10
           grd1.CellAlignment = flexAlignRightCenter
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

Private Sub textCF01_GotFocus()
InverseTextBox textCF01
End Sub

Private Sub textCF01_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textCF01_Validate(Cancel As Boolean)
If Trim(textCF01) <> "" And m_EditMode = 1 Then
    If CheckIsTaiwanDate(textCF01, False) = False Then
        Cancel = True
        MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
    ElseIf ChkWorkDay(ChangeTStringToWString(textCF01)) = False Then
        Cancel = True
        MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
    End If
End If
End Sub

Private Sub textCF02_GotFocus()
InverseTextBox textCF02
End Sub

Private Sub textCF02_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textCF02_Validate(Cancel As Boolean)
If Trim(textCF02) <> "" Then
    If InStr(1, textCF02, ".") = 0 Then
        If IsNumeric(textCF02) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textCF03_GotFocus()
InverseTextBox textCF03
End Sub

Private Sub textCF03_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textCF03_Validate(Cancel As Boolean)
If Trim(textCF03) <> "" Then
    If InStr(1, textCF03, ".") = 0 Then
        If IsNumeric(textCF03) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textCF04_GotFocus()
InverseTextBox textCF04
End Sub

Private Sub textCF04_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textCF04_Validate(Cancel As Boolean)
If Trim(textCF04) <> "" Then
    If InStr(1, textCF04, ".") = 0 Then
        If IsNumeric(textCF04) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textCF05_GotFocus()
InverseTextBox textCF05
End Sub

Private Sub textCF05_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textCF05_Validate(Cancel As Boolean)
If Trim(textCF05) <> "" Then
    If InStr(1, textCF05, ".") = 0 Then
        If IsNumeric(textCF05) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textCF06_GotFocus()
InverseTextBox textCF06
End Sub

Private Sub textCF06_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textCF06_Validate(Cancel As Boolean)
If Trim(textCF06) <> "" Then
    If InStr(1, textCF06, ".") = 0 Then
        If IsNumeric(textCF06) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textCF07_GotFocus()
InverseTextBox textCF07
End Sub

Private Sub textCF07_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textCF07_Validate(Cancel As Boolean)
If Trim(textCF07) <> "" Then
    If InStr(1, textCF07, ".") = 0 Then
        If IsNumeric(textCF07) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textCF08_GotFocus()
InverseTextBox textCF08
End Sub

Private Sub textCF08_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textCF08_Validate(Cancel As Boolean)
If Trim(textCF08) <> "" Then
    If InStr(1, textCF08, ".") = 0 Then
        If IsNumeric(textCF08) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textCF09_GotFocus()
InverseTextBox textCF09
End Sub

Private Sub textCF09_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textCF09_Validate(Cancel As Boolean)
If Trim(textCF09) <> "" Then
    If InStr(1, textCF09, ".") = 0 Then
        If IsNumeric(textCF09) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textCF10_GotFocus()
InverseTextBox textCF10
End Sub

Private Sub textCF10_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textCF10_Validate(Cancel As Boolean)
If Trim(textCF10) <> "" Then
    If InStr(1, textCF10, ".") = 0 Then
        If IsNumeric(textCF10) = False Then
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入數字！", vbInformation, "輸入錯誤"
End Sub

Private Sub textCF11_GotFocus()
InverseTextBox textCF11
End Sub

Private Sub textCF11_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub textCF11_Validate(Cancel As Boolean)
If Trim(textCF11) <> "" Then
    If InStr(1, textCF11, ".") = 0 Then
        If IsNumeric(textCF11) = False Then
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
         textCF01.Locked = True
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
            If IsRecordExist(ChangeTStringToWString(textCF01)) = True Then MsgBox "已有當日資料，請修改！", vbInformation, "操作錯誤！": Exit Function
            UpdateFieldNewData
            If AddRecord = True Then
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
   For nIndex = 1 To 11
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CF" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 1 '數值型態
   Next nIndex
End Sub

'抓當日所有資料
Private Sub GetAllData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT sqldatet(CF01),CF02,CF03,CF04,CF05,CF06,CF07,CF08,CF09,CF10,CF11 FROM CreateFiles " & _
                 "WHERE CF01>=" & Val(ChangeTStringToWString(Text1(0))) & " and CF01<=" & Val(ChangeTStringToWString(Text1(1))) & " order by CF01 asc "
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
   textCF01.Locked = bEnable
   textCF02.Locked = bEnable
   textCF03.Locked = bEnable
   textCF04.Locked = bEnable
   textCF05.Locked = bEnable
   textCF06.Locked = bEnable
   textCF07.Locked = bEnable
   textCF08.Locked = bEnable
   textCF09.Locked = bEnable
   textCF10.Locked = bEnable
   textCF11.Locked = bEnable
End Sub

Private Sub ClearField()
   Dim nIndex As Integer
   
   textCF02 = Empty
   textCF03 = Empty
   textCF04 = Empty
   textCF05 = Empty
   textCF06 = Empty
   textCF07 = Empty
   textCF08 = Empty
   textCF09 = Empty
   textCF10 = Empty
   textCF11 = Empty
   For nIndex = 0 To 11
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textCF01.SetFocus: textCF01_GotFocus
      Case 2: textCF02.SetFocus: textCF02_GotFocus
   End Select
End Sub

Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM CreateFiles " & _
            "WHERE CF01 = " & Val(m_CurrKEY(0)) & " "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("CF01")) = False Then: textCF01 = ChangeWStringToTString("" & rsTmp.Fields("CF01"))
      If IsNull(rsTmp.Fields("CF02")) = False Then: textCF02 = rsTmp.Fields("CF02")
      If IsNull(rsTmp.Fields("CF03")) = False Then: textCF03 = rsTmp.Fields("CF03")
      If IsNull(rsTmp.Fields("CF04")) = False Then: textCF04 = rsTmp.Fields("CF04")
      If IsNull(rsTmp.Fields("CF05")) = False Then: textCF05 = rsTmp.Fields("CF05")
      If IsNull(rsTmp.Fields("CF06")) = False Then: textCF06 = rsTmp.Fields("CF06")
      If IsNull(rsTmp.Fields("CF07")) = False Then: textCF07 = rsTmp.Fields("CF07")
      If IsNull(rsTmp.Fields("CF08")) = False Then: textCF08 = rsTmp.Fields("CF08")
      If IsNull(rsTmp.Fields("CF09")) = False Then: textCF09 = rsTmp.Fields("CF09")
      If IsNull(rsTmp.Fields("CF10")) = False Then: textCF10 = rsTmp.Fields("CF10")
      If IsNull(rsTmp.Fields("CF11")) = False Then: textCF11 = rsTmp.Fields("CF11")
      
   End If
   ' 更新暫存區的資料
   UpdateFieldOldData rsTmp
   rsTmp.Close
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

If Me.textCF02.Enabled = True Then
   Cancel = False
   textCF02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF03.Enabled = True Then
   Cancel = False
   textCF03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF04.Enabled = True Then
   Cancel = False
   textCF04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF05.Enabled = True Then
   Cancel = False
   textCF05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF06.Enabled = True Then
   Cancel = False
   textCF06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF07.Enabled = True Then
   Cancel = False
   textCF07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF08.Enabled = True Then
   Cancel = False
   textCF08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF09.Enabled = True Then
   Cancel = False
   textCF09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF10.Enabled = True Then
   Cancel = False
   textCF10_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCF11.Enabled = True Then
   Cancel = False
   textCF11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
TxtValidate = True
End Function

Private Sub UpdateFieldNewData()
   '若新增資料
   SetFieldNewData "CF01", ChangeTStringToWString(textCF01)
   SetFieldNewData "CF02", textCF02
   SetFieldNewData "CF03", textCF03
   SetFieldNewData "CF04", textCF04
   SetFieldNewData "CF05", textCF05
   SetFieldNewData "CF06", textCF06
   SetFieldNewData "CF07", textCF07
   SetFieldNewData "CF08", textCF08
   SetFieldNewData "CF09", textCF09
   SetFieldNewData "CF10", textCF10
   SetFieldNewData "CF11", textCF11
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
   Dim strCF01 As String
   
   AddRecord = False
   strCF01 = textCF01

   bFirst = True
   strSql = "INSERT INTO CreateFiles ("
   For nIndex = 0 To 11
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
   strSql = strSql & ",cf12,cf13,cf14) "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To 11
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
   strSql = strSql & ",'" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI')))"
On Error GoTo ErrHand
    cnnConnection.BeginTrans
   
   cnnConnection.Execute strSql
   
    cnnConnection.CommitTrans
    m_CurrKEY(0) = strCF01
   If (Val(strCF01) < Val(m_FirstKEY(0))) Or (Val(strCF01) > Val(m_LastKEY(0))) Then
      RefreshRange
   End If
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
   Dim strCF01 As String
   
   ModRecord = False
   
   strCF01 = m_CurrKEY(0)
   strSql = "UPDATE CreateFiles SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To 11
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

   strSql = strSql & ",cf15='" & strUserNum & "',cf16=to_number(to_char(sysdate,'YYYYMMDD')),cf17=to_number(to_char(sysdate,'HH24MI')) " & _
                  "WHERE CF01 = " & strCF01 & "  "
On Error GoTo ErrHand
   If bDifference = True Then
      cnnConnection.BeginTrans
      
      cnnConnection.Execute strSql

      cnnConnection.CommitTrans
      
      RefreshRange
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
   
   For nIndex = 0 To 11
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
   For nIndex = 0 To 11
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
   strSql = "SELECT * FROM CreateFiles " & _
            "WHERE CF01 = " & strKEY01 & " "
                  
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
   
   strSql = "SELECT max(CF01) as CF01 FROM CreateFiles " & _
            "WHERE CF01 < '" & m_CurrKEY(0) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CF01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT min(CF01) as CF01 FROM CreateFiles "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CF01")
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
   
   strSql = "SELECT min(CF01) as  CF01 FROM CreateFiles " & _
            "WHERE CF01 > '" & m_CurrKEY(0) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CF01")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT max(CF01) as CF01 FROM CreateFiles  "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("CF01")
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
   
   strSql = "SELECT min(CF01) as CF01 FROM CreateFiles  "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("CF01")
   End If
   rsTmp.Close

   strSql = "SELECT max(CF01) as CF01 FROM CreateFiles  "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsTmp.EOF And Not rsTmp.BOF Then
      If IsNull(rsTmp.Fields("CF01")) = False Then: m_LastKEY(0) = rsTmp.Fields("CF01")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub
