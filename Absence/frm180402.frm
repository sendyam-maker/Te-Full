VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180402 
   BorderStyle     =   1  '單線固定
   Caption         =   "簽核主管特殊對象的簽核職代"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Begin VB.TextBox txtB0209 
      Height          =   270
      Left            =   5760
      MaxLength       =   1
      TabIndex        =   1
      Top             =   630
      Width           =   260
   End
   Begin VB.TextBox txtB0208 
      Height          =   270
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1020
      Width           =   735
   End
   Begin VB.TextBox txtB0201 
      Height          =   270
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   0
      Top             =   630
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
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
            Picture         =   "frm180402.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180402.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180402.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180402.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180402.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180402.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180402.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180402.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180402.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180402.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm180402.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8950
      _ExtentX        =   15787
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
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
      BorderStyle     =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Bindings        =   "frm180402.frx":20F4
      Height          =   3165
      Left            =   45
      TabIndex        =   9
      Top             =   2310
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   5592
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
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
      _Band(0).Cols   =   13
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "簽核種類："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   7
      Left            =   4830
      TabIndex        =   25
      Top             =   690
      Width           =   900
   End
   Begin MSForms.Label Label4 
      Height          =   290
      Left            =   6060
      TabIndex        =   24
      Top             =   660
      Width           =   2210
      VariousPropertyBits=   27
      Caption         =   "(1.人事 2.案件)"
      Size            =   "3898;512"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   6
      Left            =   3840
      TabIndex        =   8
      Top             =   1890
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   4
      Left            =   3840
      TabIndex        =   6
      Top             =   1590
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   2
      Left            =   3840
      TabIndex        =   4
      Top             =   1290
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   5
      Left            =   2040
      TabIndex        =   7
      Top             =   1890
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   3
      Left            =   2040
      TabIndex        =   5
      Top             =   1590
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   300
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   1290
      Width           =   1520
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2672;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註：若有異動資料需通知人事室！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   21
      Left            =   210
      TabIndex        =   23
      Top             =   5520
      Width           =   3120
   End
   Begin VB.Label Label3 
      Caption         =   "(部門別或員工代號)"
      Height          =   200
      Left            =   4440
      TabIndex        =   22
      Top             =   1050
      Width           =   1820
   End
   Begin VB.Label Label2 
      Caption         =   "備註：簽核對象若為空白，代表沒有指定固定的簽核對象。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   500
      Left            =   5910
      TabIndex        =   21
      Top             =   1520
      Width           =   2900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "被簽核的對象："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   750
      TabIndex        =   20
      Top             =   1050
      Width           =   1260
   End
   Begin MSForms.Label txtB0208_2 
      Height          =   290
      Left            =   2820
      TabIndex        =   19
      Top             =   1020
      Width           =   1580
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2778;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label txtB0201_2 
      Height          =   290
      Left            =   2820
      TabIndex        =   18
      Top             =   660
      Width           =   1490
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(2)"
      Height          =   180
      Index           =   11
      Left            =   3600
      TabIndex        =   17
      Top             =   1950
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(2)"
      Height          =   180
      Index           =   3
      Left            =   3600
      TabIndex        =   16
      Top             =   1650
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(2)"
      Height          =   180
      Index           =   2
      Left            =   3600
      TabIndex        =   15
      Top             =   1350
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人1：(1)"
      Height          =   180
      Index           =   6
      Left            =   630
      TabIndex        =   14
      Top             =   1350
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人2：(1)"
      Height          =   180
      Index           =   5
      Left            =   630
      TabIndex        =   13
      Top             =   1650
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "職務代理人3：(1)"
      Height          =   180
      Index           =   4
      Left            =   630
      TabIndex        =   12
      Top             =   1950
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "簽核主管："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   1110
      TabIndex        =   11
      Top             =   690
      Width           =   900
   End
End
Attribute VB_Name = "frm180402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2023/12/20 修改抓新部門程式
'Memo By Sindy 2021/11/17 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Create by Sindy 2011/8/8
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 第一筆資料的Key
Dim m_FirstKEY(3) As String
' 最後一筆資料的Key
Dim m_LastKEY(3) As String
' 目前正在顯示的Key
Dim m_CurrKEY(3) As String
Dim i As Integer, j As Integer
Dim dblPrevRow As Double


Private Sub Combo2_GotFocus(Index As Integer)
   InverseTextBox Combo2(Index)
End Sub

Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
   If m_EditMode <> 0 And Combo2(Index).Text > "" And Len(Trim(Combo2(Index).Text)) = 5 Then
      '抓取員工姓名
      Combo2(Index).Text = SetCboStaffName(Combo2(Index).Text)
   End If
End Sub

Private Sub Combo2_Validate(Index As Integer, Cancel As Boolean)
   If m_EditMode <> 0 And Combo2(Index) <> "" Then
      If Index <> 0 Then
         If Left(Combo2(Index), 5) = txtB0201 Then
            MsgBox "不可為本人！", vbExclamation
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(Combo2(Index), 5)) = True Then
         Call Combo2_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查 員工不可為”不寄信”
      If ChkStaffST14(Left(Combo2(Index), 5)) = True Then
         Call Combo2_GotFocus(Index)
         Cancel = True
         Exit Sub
      End If
      '檢查職代輸入順序
      If Index >= 1 And Index <= 6 Then
         If (Trim(Combo2(2)) <> "" And Trim(Combo2(1)) = "") Or _
            (Trim(Combo2(4)) <> "" And Trim(Combo2(3)) = "") Or _
            (Trim(Combo2(6)) <> "" And Trim(Combo2(5)) = "") Then
            MsgBox "請依序輸入職務代理人！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         If (Trim(Combo2(3)) <> "" And Trim(Combo2(1)) = "") Or _
            (Trim(Combo2(5)) <> "" And Trim(Combo2(3)) = "") Then
            MsgBox "請依序輸入職務代理人！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         If (Combo2(2) <> "" And Left(Combo2(2), 5) = Left(Combo2(1), 5)) Or _
            (Combo2(4) <> "" And Left(Combo2(4), 5) = Left(Combo2(3), 5)) Or _
            (Combo2(6) <> "" And Left(Combo2(6), 5) = Left(Combo2(5), 5)) Then
            MsgBox "資料重覆！", vbExclamation
            Combo2(Index).SetFocus
            Call Combo2_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

' 按下按鍵
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
'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub

Private Sub Form_Load()
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   
   SetDataListWidth
   
   ClearField
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   ReadAllData
   'OnAction vbKeyF4
   OnAction vbKeyF10
   
   'Modify By Sindy 2025/1/24 Mark
'   'Add By Sindy 2023/5/4
'   If Pub_StrUserSt03 = "M51" Then
'      txtB0209.Enabled = True
'   Else
'      txtB0209.Enabled = False
'   End If
'   '2023/5/4 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180402 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grd1, x, y, nCol, nRow
grd1.col = nCol
grd1.row = nRow
End Sub

Private Sub grd1_SelChange()
grd1.Visible = False
If grd1.MouseRow <> 0 Then
'   '上一筆資料列清除反白
'   If dblPrevRow > 0 Then
'      grd1.col = 2
'      grd1.row = dblPrevRow
'      For i = 0 To 1
'         grd1.col = i
'         grd1.CellBackColor = &H8000000F
'      Next i
'      For i = 2 To grd1.Cols - 1
'         grd1.col = i
'         grd1.CellBackColor = QBColor(15)
'      Next i
'   End If
'   '目前資料列反白
'   grd1.col = 0
'   grd1.row = grd1.MouseRow
'   dblPrevRow = grd1.row
'   For i = 0 To grd1.Cols - 1
'      grd1.col = i
'      grd1.CellBackColor = &HFFC0C0
'   Next i
   '查詢目前資料列
   'ShowCurrRecord grd1.TextMatrix(grd1.row, 8), grd1.TextMatrix(grd1.row, 9)
   m_CurrKEY(0) = grd1.TextMatrix(grd1.row, 8)
   m_CurrKEY(1) = grd1.TextMatrix(grd1.row, 9)
   m_CurrKEY(2) = grd1.TextMatrix(grd1.row, 12) 'Add By Sindy 2023/5/5
   UpdateCtrlData
End If
grd1.Visible = True
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean

TxtValidate = False

If txtB0201.Text = "" Then
    MsgBox "簽核主管不可以空白！", vbExclamation
    txtB0201.SetFocus
    Exit Function
End If

'Modify By Sindy 2012/2/14
'If txtB0208.Text = "" Then
'    MsgBox "被簽核的對象不可以空白！", vbExclamation
'    txtB0208.SetFocus
'    Exit Function
'End If

'Add By Sindy 2023/5/4
Cancel = False
txtB0209_Validate Cancel
If Cancel = True Then
   Exit Function
End If

If m_EditMode = 1 Then
   ' 檢查記錄是否已存在
   If IsRecordExist(txtB0201, txtB0208, txtB0209) = True Then
      MsgBox "該筆記錄已存在", vbOKOnly, "更新資料"
'      If txtB0201_2 = "" Then Call txtB0201_LostFocus
'      If txtB0208 = "" Then Call txtB0208_LostFocus
      txtB0201.SetFocus
      Exit Function
   End If
End If

'Modify By Sindy 2025/1/24 Mark
''Add By Sindy 2023/5/5
'If (m_EditMode = 1 Or m_EditMode = 2 Or m_EditMode = 3) _
'   And Pub_StrUserSt03 <> "M51" _
'   And txtB0209 <> "1" Then
'   MsgBox "無權限異動簽核種類不是〔人事〕的資料！請洽電腦中心。", vbOKOnly, "更新資料"
'   txtB0201.SetFocus
'   Exit Function
'End If
''2023/5/5 END

Cancel = False
txtB0201_Validate Cancel
If Cancel = True Then
   Exit Function
End If

Cancel = False
txtB0208_Validate Cancel
If Cancel = True Then
   Exit Function
End If

'Modify By Sindy 2012/2/14
'If txtB0201.Text = txtB0208.Text Then
'    MsgBox "(簽核主管)和(被簽核的對象)不可重覆！", vbExclamation
'    txtB0208.SetFocus
'    Exit Function
'End If

If Combo2(1).Text = "" _
   And Combo2(2).Text = "" _
   And Combo2(3).Text = "" _
   And Combo2(4).Text = "" _
   And Combo2(5).Text = "" _
   And Combo2(6).Text = "" Then
    MsgBox "職務代理人不可以空白！", vbExclamation
    Combo2(1).SetFocus
    Exit Function
End If

For i = 1 To Combo2.UBound
   Cancel = False
   Combo2_Validate i, Cancel
   If Cancel = True Then
      Exit Function
   End If
Next i

TxtValidate = True
End Function

' 更新資料
Private Function SaveData(strEditMode As Integer) As Boolean
Dim strKEY01 As String, strKEY02 As String, strKEY03 As String
   
On Error GoTo ErrHand
   
   SaveData = False
   
   If txtB0208.Text = "" Then txtB0208.Text = txtB0201.Text 'Add By Sindy 2012/2/15
   strKEY01 = txtB0201
   strKEY02 = txtB0208
   strKEY03 = txtB0209 'Add By Sindy 2023/5/4
   
   cnnConnection.BeginTrans
   '新增
   If strEditMode = 1 Then
      'Modify By Sindy 2023/5/4 + txtB0209
      strSql = "INSERT INTO ABS002 VALUES(" & CNULL(strKEY01) & _
                  "," & CNULL(Left(Trim(Combo2(1)), 5)) & "," & CNULL(Left(Trim(Combo2(2)), 5)) & _
                  "," & CNULL(Left(Trim(Combo2(3)), 5)) & "," & CNULL(Left(Trim(Combo2(4)), 5)) & _
                  "," & CNULL(Left(Trim(Combo2(5)), 5)) & "," & CNULL(Left(Trim(Combo2(6)), 5)) & _
                  "," & CNULL(strKEY02) & "," & CNULL(txtB0209) & ")"
   '修改
   ElseIf strEditMode = 2 Then
      'Modify By Sindy 2023/5/4 + txtB0209
      strSql = "UPDATE ABS002 SET " & _
                  "B0202=" & CNULL(Left(Trim(Combo2(1)), 5)) & ",B0203=" & CNULL(Left(Trim(Combo2(2)), 5)) & _
                  ",B0204=" & CNULL(Left(Trim(Combo2(3)), 5)) & ",B0205=" & CNULL(Left(Trim(Combo2(4)), 5)) & _
                  ",B0206=" & CNULL(Left(Trim(Combo2(5)), 5)) & ",B0207=" & CNULL(Left(Trim(Combo2(6)), 5)) & _
                  ",B0209=" & CNULL(txtB0209) & _
               " WHERE B0201=" & CNULL(strKEY01) & " and B0208=" & CNULL(strKEY02) & " and B0209=" & CNULL(strKEY03)
   End If
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   
   If (strKEY01 < m_FirstKEY(0)) Or (strKEY01 > m_LastKEY(0)) Then
      RefreshRange
   End If
   ShowCurrRecord strKEY01, strKEY02, strKEY03
   
   SaveData = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 更新失敗！" & vbCrLf & Err.Description
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strKEY01 As String, strKEY02 As String, strKEY03 As String
   
On Error GoTo ErrHand
   
   DelRecord = False
   
   If txtB0208.Text = "" Then txtB0208.Text = txtB0201.Text 'Add By Sindy 2012/2/15
   strKEY01 = txtB0201
   strKEY02 = txtB0208
   strKEY03 = txtB0209
   
   cnnConnection.BeginTrans
   
   strSql = "DELETE FROM ABS002 WHERE b0201 = " & CNULL(strKEY01) & " and b0208 = " & CNULL(strKEY02) & " and b0209 = " & CNULL(strKEY03)
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
'Dim strKEY01 As String
'Dim strKEY02 As String
'
'   QueryRecord = False
'
'   strKEY01 = txtB0201
'   strKEY02 = txtB0208
'
'   If IsRecordExist(strKEY01, strKEY02) = True Then
'      m_CurrKEY(0) = strKEY01
'      m_CurrKEY(1) = strKEY02
'      QueryRecord = True
'      UpdateCtrlData
''      ReadAllData
'   Else
'      QueryRecord = False
'   End If
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
   
   'Modify By Sindy 2012/2/14
   QueryRecord = False
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   strCon = ""
   If Trim(txtB0201) <> "" Then
      strCon = strCon & "and B0201='" & txtB0201 & "' "
   End If
   If Trim(txtB0208) <> "" Then
      strCon = strCon & "and B0208='" & txtB0208 & "' "
   End If
   
   grd1.Rows = 2
   grd1.Clear
   grd1.FixedCols = 0
   dblPrevRow = 0
   'Modify By Sindy 2012/2/15 nvl(A0902,nvl(s7.ST02,B0208)) 簽核主管與被簽核的對象相同時,被簽核的對象顯示空白
   strSql = "select s0.ST02,decode(B0208,B0201,' ',nvl(A0922,nvl(s7.ST02,B0208)))" & _
            ",s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02,B0201,B0208,s0.ST04,decode(B0209,'1','人事','2','案件',B0209),B0209 " & _
            "from ABS002,ACC090NEW,STAFF s0 " & _
            ",STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF s7 " & _
            "where B0201=s0.ST01(+) " & _
            "and B0208=s7.ST01(+) " & _
            "and B0208=A0921(+) " & _
            "and B0202=s1.ST01(+) " & _
            "and B0203=s2.ST01(+) " & _
            "and B0204=s3.ST01(+) " & _
            "and B0205=s4.ST01(+) " & _
            "and B0206=s5.ST01(+) " & _
            "and B0207=s6.ST01(+) " & strCon & _
            "order by B0201,B0208"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd1.Recordset = rsTmp
      grd1.FixedCols = 2
      QueryRecord = True
      rsTmp.MoveFirst
      m_CurrKEY(0) = rsTmp.Fields("B0201")
      m_CurrKEY(1) = rsTmp.Fields("B0208")
      m_CurrKEY(2) = rsTmp.Fields("B0209")
      UpdateCtrlData
      dblPrevRow = 1
   End If
   rsTmp.Close
   SetDataListWidth
   GetSelChage
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   UpdateToolbarState
   
EXITSUB:
   Set rsTmp = Nothing
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse
   
   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         If SaveData(m_EditMode) = True Then
             RefreshRange
             ReadAllData
             SetKeyReadOnly True
         Else
             Exit Function
         End If
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Function
         If SaveData(m_EditMode) = False Then Exit Function
         ReadAllData
         SetKeyReadOnly True
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1), m_CurrKEY(2)
            ReadAllData
            SetKeyReadOnly True
         Else
            Exit Function
         End If
      Case 4: '查詢
         If txtB0201.Text = "" Then
             MsgBox "簽核主管不可以空白！", vbExclamation
             txtB0201.SetFocus
         End If
         'Modify By Sindy 2012/2/14
'         If txtB0208.Text = "" Then
'             MsgBox "被簽核的對象不可以空白！", vbExclamation
'             txtB0208.SetFocus
'         End If
'         If txtB0201 <> "" And txtB0208 <> "" Then
         If txtB0201 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
               ReadAllData 'Add By Sindy 2012/2/14
            End If
            SetKeyReadOnly True
         Else
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
   
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 0, 1, 4: If Me.txtB0201.Visible = True Then txtB0201.SetFocus
      Case 2: Combo2(1).SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM ABS002 WHERE b0201=" & CNULL(strKEY01) & " and b0208=" & CNULL(strKEY02) & " and b0209=" & CNULL(strKEY03)
   
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

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02, strKEY03) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
      m_CurrKEY(2) = strKEY03
   Else
      strSql = "SELECT B0201,B0208,B0209 FROM ABS002 WHERE B0201='" & m_CurrKEY(0) & "' and B0208='" & m_CurrKEY(1) & "' and B0209='" & m_CurrKEY(2) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
         If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT B0201,B0208,B0209 FROM ABS002 order by B0201 asc,B0208 asc,B0209 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
         If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
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
   m_CurrKEY(1) = m_FirstKEY(1)
   m_CurrKEY(2) = m_FirstKEY(2)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) And m_CurrKEY(2) = m_FirstKEY(2) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT B0201,B0208,B0209 FROM ABS002 WHERE B0201||B0208||B0209<'" & m_CurrKEY(0) & m_CurrKEY(1) & m_CurrKEY(2) & "' order by B0201 desc,B0208 desc,B0209 desc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT B0201,B0208,B0209 FROM ABS002 order by B0201 asc,B0208 asc,B0209 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) And m_CurrKEY(2) = m_LastKEY(2) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT B0201,B0208,B0209 FROM ABS002 WHERE B0201||B0208||B0209>'" & m_CurrKEY(0) & m_CurrKEY(1) & m_CurrKEY(2) & "' order by B0201 asc,B0208 asc,B0209 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT B0201,B0208,B0209 FROM ABS002 order by B0201 asc,B0208 asc,B0209 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      If IsNull(rsTmp.Fields(2)) = False Then: m_CurrKEY(2) = rsTmp.Fields(2)
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   m_CurrKEY(2) = m_LastKEY(2)
   UpdateCtrlData
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
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         Call txtB0201_LostFocus
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
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
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
                  SetKeyReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               ReadAllData 'Add By Sindy 2012/2/14
               UpdateCtrlData
               SetCtrlReadOnly True
               SetKeyReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   txtB0201.Locked = bEnable
   If bEnable Then txtB0201.BackColor = &H8000000F Else txtB0201.BackColor = &H80000005
   txtB0208.Locked = bEnable
   If bEnable Then txtB0208.BackColor = &H8000000F Else txtB0208.BackColor = &H80000005
   'Add By Sindy 2023/5/4
   txtB0209.Locked = bEnable
   If bEnable Then txtB0209.BackColor = &H8000000F Else txtB0209.BackColor = &H80000005
   '2023/5/4 END
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   For i = 1 To Combo2.UBound
      Combo2(i).Locked = bEnable
      If bEnable Then Combo2(i).BackColor = &H8000000F Else Combo2(i).BackColor = &H80000005
   Next i
End Sub

Private Sub ClearField()
   txtB0201 = Empty
   txtB0201_2 = Empty
   txtB0208 = Empty
   txtB0208_2 = Empty
   For i = 1 To Combo2.UBound
      'Combo2(i).Clear
      Combo2(i).Text = Empty
   Next i
   txtB0209 = "1" 'Add By Sindy 2023/5/4
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub ReadAllData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   grd1.Rows = 2
   grd1.Clear
   grd1.FixedCols = 0
   'Modify By Sindy 2012/2/15 nvl(A0902,nvl(s7.ST02,B0208)) 簽核主管與被簽核的對象相同時,被簽核的對象顯示空白
   strSql = "select s0.ST02,decode(B0208,B0201,' ',nvl(A0922,nvl(s7.ST02,B0208)))" & _
            ",s1.ST02,s2.ST02,s3.ST02,s4.ST02,s5.ST02,s6.ST02,B0201,B0208,s0.ST04,decode(B0209,'1','人事','2','案件',B0209),B0209 " & _
            "from ABS002,ACC090NEW,STAFF s0 " & _
            ",STAFF s1,STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF s7 " & _
            "where B0201=s0.ST01(+) " & _
            "and B0208=s7.ST01(+) " & _
            "and B0208=A0921(+) " & _
            "and B0202=s1.ST01(+) " & _
            "and B0203=s2.ST01(+) " & _
            "and B0204=s3.ST01(+) " & _
            "and B0205=s4.ST01(+) " & _
            "and B0206=s5.ST01(+) " & _
            "and B0207=s6.ST01(+) " & _
            "order by B0201,B0208"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd1.Recordset = rsTmp
      grd1.FixedCols = 2
   End If
   rsTmp.Close
   SetDataListWidth
   GetSelChage
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub GetSelChage()
grd1.Visible = False
If grd1.Rows - 1 > 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      grd1.col = 2
      grd1.row = dblPrevRow
      For i = 0 To 1
         grd1.col = i
         grd1.CellBackColor = &H8000000F
      Next i
      For i = 2 To grd1.Cols - 1
         grd1.col = i
         grd1.CellBackColor = QBColor(15)
      Next i
   End If
   '尋找目前資料列
   For j = 1 To grd1.Rows - 1
      If grd1.TextMatrix(j, 8) = m_CurrKEY(0) And grd1.TextMatrix(j, 9) = m_CurrKEY(1) And grd1.TextMatrix(j, 12) = m_CurrKEY(2) Then
         grd1.col = 0
         grd1.row = j
         dblPrevRow = grd1.row
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = &HFFC0C0
         Next i
'            grd1.TopRow = j
         Exit For
      End If
   Next j
End If
grd1.Visible = True
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   ClearField
   
   strSql = "SELECT ABS002.*,A0921,A0922,A0925,s1.ST02 s1_ST02 " & _
            ",s2.ST02 s2_ST02,s3.ST02 s3_ST02,s4.ST02 s4_ST02,s5.ST02 s5_ST02,s6.ST02 s6_ST02,s7.ST02 s7_ST02 " & _
            "FROM ABS002,STAFF s1,ACC090NEW " & _
            ",STAFF s2,STAFF s3,STAFF s4,STAFF s5,STAFF s6,STAFF s7 " & _
            "WHERE B0201=s1.ST01(+) and s1.ST93=A0921(+) and B0201='" & m_CurrKEY(0) & "' and B0208='" & m_CurrKEY(1) & "' " & _
            "and B0202=s2.ST01(+) and B0203=s3.ST01(+) and B0204=s4.ST01(+) " & _
            "and B0205=s5.ST01(+) and B0206=s6.ST01(+) and B0207=s7.ST01(+) " & _
            "and B0209='" & m_CurrKEY(2) & "' "
   'strSql = "SELECT ABS002.*,A0901,A0902,A0911 FROM ABS002,STAFF,ACC090 WHERE B0201=ST01(+) and ST03=A0901(+) and B0201='" & m_CurrKEY(0) & "' and B0208='" & m_CurrKEY(1) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If m_EditMode = 1 And txtB0201.Enabled = True Then
         '員工代號欄位為空白,使用者自行輸入欲新增的員工代號
      Else
         If IsNull(rsTmp.Fields("B0201")) = False Then txtB0201 = rsTmp.Fields("B0201"): txtB0201_2 = rsTmp.Fields("s1_ST02")
      End If
      If IsNull(rsTmp.Fields("B0202")) = False Then Combo2(1).Text = Left(Trim(rsTmp.Fields("B0202")) & Space(5), 7) & rsTmp.Fields("s2_ST02")
      If IsNull(rsTmp.Fields("B0203")) = False Then Combo2(2).Text = Left(Trim(rsTmp.Fields("B0203")) & Space(5), 7) & rsTmp.Fields("s3_ST02")
      If IsNull(rsTmp.Fields("B0204")) = False Then Combo2(3).Text = Left(Trim(rsTmp.Fields("B0204")) & Space(5), 7) & rsTmp.Fields("s4_ST02")
      If IsNull(rsTmp.Fields("B0205")) = False Then Combo2(4).Text = Left(Trim(rsTmp.Fields("B0205")) & Space(5), 7) & rsTmp.Fields("s5_ST02")
      If IsNull(rsTmp.Fields("B0206")) = False Then Combo2(5).Text = Left(Trim(rsTmp.Fields("B0206")) & Space(5), 7) & rsTmp.Fields("s6_ST02")
      If IsNull(rsTmp.Fields("B0207")) = False Then Combo2(6).Text = Left(Trim(rsTmp.Fields("B0207")) & Space(5), 7) & rsTmp.Fields("s7_ST02")
      'Modify By Sindy 2012/2/15 簽核主管與被簽核的對象相同時,被簽核的對象顯示空白
      txtB0208 = "": txtB0208_2 = ""
      If IsNull(rsTmp.Fields("B0208")) = False Then
         If rsTmp.Fields("B0201") <> rsTmp.Fields("B0208") Then
            txtB0208 = rsTmp.Fields("B0208")
            'Modify By Sindy 2025/3/5
            If Len(txtB0208) = 5 Then
               txtB0208_2 = GetPrjSalesNM(Trim(txtB0208))
            Else
               txtB0208_2 = GetPrjSalesBlack(Trim(txtB0208), True)
            End If
            '2025/3/5 END
            If txtB0208_2 = "" Then
               txtB0208_2 = GetDeptNameA0922(Trim(txtB0201))
            End If
         End If
      End If
      '2012/2/15 End
      If IsNull(rsTmp.Fields("B0209")) = False Then txtB0209 = rsTmp.Fields("B0209") 'Add By Sindy 2023/5/4
   End If
   rsTmp.Close
   GetSelChage
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "select B0201,B0208,B0209 from ABS002 order by B0201 asc,B0208 asc,B0209 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then m_FirstKEY(0) = Trim(rsTmp.Fields(0))
      If IsNull(rsTmp.Fields(1)) = False Then m_FirstKEY(1) = Trim(rsTmp.Fields(1))
      If IsNull(rsTmp.Fields(2)) = False Then m_FirstKEY(2) = Trim(rsTmp.Fields(2))
   End If
   rsTmp.Close
   
   strSql = "select B0201,B0208,B0209 from ABS002 order by B0201 desc,B0208 desc,B0209 desc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then m_LastKEY(0) = Trim(rsTmp.Fields(0))
      If IsNull(rsTmp.Fields(1)) = False Then m_LastKEY(1) = Trim(rsTmp.Fields(1))
      If IsNull(rsTmp.Fields(2)) = False Then m_LastKEY(2) = Trim(rsTmp.Fields(2))
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
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
         Else
            TBar1.Buttons(6).Enabled = False
            TBar1.Buttons(7).Enabled = False
            TBar1.Buttons(8).Enabled = False
            TBar1.Buttons(9).Enabled = False
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
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

Private Sub SetDataListWidth()
grd1.row = 0
grd1.col = 0: grd1.Text = "簽核主管"
grd1.ColWidth(0) = 1200
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 1: grd1.Text = "被簽核的對象"
grd1.ColWidth(1) = 1200
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 2: grd1.Text = "職代一(1)"
grd1.ColWidth(2) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 3: grd1.Text = "職代一(2)"
grd1.ColWidth(3) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 4: grd1.Text = "職代二(1)"
grd1.ColWidth(4) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 5: grd1.Text = "職代二(2)"
grd1.ColWidth(5) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 6: grd1.Text = "職代三(1)"
grd1.ColWidth(6) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 7: grd1.Text = "職代三(2)"
grd1.ColWidth(7) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 8: grd1.Text = "B0201"
grd1.ColWidth(8) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 9: grd1.Text = "B0208"
grd1.ColWidth(9) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 10: grd1.Text = "ST04"
grd1.ColWidth(10) = 0
grd1.CellAlignment = flexAlignLeftCenter
'Add By Sindy 2023/5/4
grd1.col = 11: grd1.Text = "簽核種類"
grd1.ColWidth(11) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 12: grd1.Text = "B0209"
grd1.ColWidth(12) = 0
grd1.CellAlignment = flexAlignLeftCenter
'2023/5/4 END
End Sub

Private Sub txtB0201_GotFocus()
   InverseTextBox txtB0201
End Sub

Private Sub txtB0201_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtB0201_LostFocus()
Dim Rs As New ADODB.Recordset
Dim strText(14) As String
Dim strA0911 As String, strA0925 As String
   
   If m_EditMode <> 0 And txtB0201 <> "" Then
      txtB0201_2 = GetStaffName(txtB0201, True)
      strA0911 = GetStaffA0911(txtB0201, strA0925) 'Modify By Sindy 2023/12/20
      
      For i = 1 To Combo2.UBound
         strText(i) = Combo2(i).Text
         'Modify By Sindy 2023/12/20
         'Combo2(i).Clear
         'Combo2(i).AddItem ""
         Call SetB1003Combo(Combo2(i), strA0911, strA0925)
         Combo2(i).Text = strText(i)
         '2023/12/20 END
      Next i
   End If
End Sub

Private Sub txtB0201_Validate(Cancel As Boolean)
   If txtB0201.Text = "" Then txtB0201_2 = ""
   
   If m_EditMode <> 0 And txtB0201 <> "" Then
      ' 檢查員工編號規則
      If ChkStaffID(txtB0201) Then
         Call txtB0201_GotFocus
         Cancel = True
         Exit Sub
      End If
      txtB0201_2 = GetStaffName(txtB0201, True)
      If txtB0201_2 = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         Call txtB0201_GotFocus
         Cancel = True
         Exit Sub
      End If
      If m_EditMode = 1 And txtB0208 <> "" Then
         ' 檢查記錄是否已存在
         If IsRecordExist(txtB0201, txtB0208, txtB0209) = True Then
            MsgBox "該筆記錄已存在", vbInformation
            Call txtB0201_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtB0208_GotFocus()
   InverseTextBox txtB0208
End Sub

Private Sub txtB0208_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtB0208_LostFocus()
   If m_EditMode <> 0 And Trim(txtB0208) <> "" Then
      txtB0208_2 = GetPrjSalesNM(Trim(txtB0208))
      If txtB0208_2 = "" Then
         txtB0208_2 = GetDeptNameA0922(Trim(txtB0208))
      End If
   End If
End Sub

Private Sub txtB0208_Validate(Cancel As Boolean)
   If txtB0208.Text = "" Then txtB0208_2 = ""
   
   If m_EditMode <> 0 And txtB0208 <> "" Then
      txtB0208_2 = GetPrjSalesBlack(Trim(txtB0208))
      If txtB0208_2 = "" Then
         txtB0208_2 = GetPrjSalesNM(Trim(txtB0208))
      End If
      If txtB0208_2 = "" Then
         MsgBox "查無此部門或員工！", vbInformation
         Call txtB0208_GotFocus
         Cancel = True
         Exit Sub
      End If
      If m_EditMode = 1 And txtB0201 <> "" Then
         ' 檢查記錄是否已存在
         If IsRecordExist(txtB0201, txtB0208, txtB0209) = True Then
            MsgBox "該筆記錄已存在", vbInformation
            Call txtB0208_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtB0209_GotFocus()
   InverseTextBox txtB0209
End Sub

Private Sub txtB0209_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 49 And KeyAscii <> 50 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtB0209_Validate(Cancel As Boolean)
If m_EditMode <> 0 Then
   If txtB0209 <> "" Then
      If CheckLengthIsOK(txtB0209, txtB0209.MaxLength) = False Then
         Call txtB0209_GotFocus
         Cancel = True
         Exit Sub
      End If
   Else
      MsgBox "簽核種類不可以空白！", vbExclamation
      Call txtB0209_GotFocus
      Cancel = True
      Exit Sub
   End If
End If
End Sub
