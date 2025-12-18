VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010614 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "郵件分信關鍵字對照表維護"
   ClientHeight    =   6380
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6380
   ScaleWidth      =   8950
   Begin VB.ComboBox Combo5 
      Height          =   300
      ItemData        =   "frm06010614.frx":0000
      Left            =   60
      List            =   "frm06010614.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   26
      Top             =   1980
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      ItemData        =   "frm06010614.frx":0004
      Left            =   1080
      List            =   "frm06010614.frx":000E
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   960
      Width           =   1905
   End
   Begin VB.CheckBox ChkLK14 
      Caption         =   "用單字索引"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7230
      TabIndex        =   6
      Top             =   1290
      Width           =   1275
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "frm06010614.frx":002A
      Left            =   7290
      List            =   "frm06010614.frx":0037
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   960
      Width           =   1485
   End
   Begin VB.TextBox txtLK04 
      Height          =   270
      Left            =   3810
      MaxLength       =   150
      TabIndex        =   11
      Top             =   1890
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.TextBox txtLK11 
      Height          =   270
      Left            =   4110
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1260
      Width           =   345
   End
   Begin VB.TextBox txtLK13 
      Height          =   270
      Left            =   4110
      MaxLength       =   4
      TabIndex        =   2
      Top             =   960
      Width           =   765
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "→"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3420
      TabIndex        =   9
      Top             =   1890
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "←"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3420
      TabIndex        =   8
      Top             =   1620
      Width           =   375
   End
   Begin VB.ListBox lstUsers 
      Height          =   760
      ItemData        =   "frm06010614.frx":0055
      Left            =   1080
      List            =   "frm06010614.frx":005C
      TabIndex        =   7
      Top             =   1590
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm06010614.frx":006A
      Left            =   3810
      List            =   "frm06010614.frx":006C
      TabIndex        =   10
      Top             =   1590
      Width           =   3195
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm06010614.frx":006E
      Left            =   1080
      List            =   "frm06010614.frx":0070
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1260
      Width           =   1485
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010614.frx":0072
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010614.frx":038E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010614.frx":06AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010614.frx":0886
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010614.frx":0BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010614.frx":0EBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010614.frx":11DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010614.frx":14F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010614.frx":1812
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010614.frx":1B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm06010614.frx":1E4A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   13
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
      Bindings        =   "frm06010614.frx":2166
      Height          =   2685
      Left            =   45
      TabIndex        =   12
      Top             =   2640
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   4745
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|關鍵字|種類|分類|收受者|優先排序|比 Initial 優先檢查|新增人員日期|更新人員日期|使用單位|單字索引"
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
   Begin MSForms.TextBox txtLK01 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   630
      Width           =   6195
      VariousPropertyBits=   679495707
      MaxLength       =   100
      Size            =   "10927;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox ListSysSet 
      Height          =   975
      Left            =   630
      TabIndex        =   29
      Top             =   5370
      Width           =   8295
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "14631;1720"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label23 
      Height          =   195
      Index           =   1
      Left            =   3870
      TabIndex        =   28
      Top             =   2430
      Width           =   4875
      VariousPropertyBits=   27
      Caption         =   "UPDATE：ID  Date  Time"
      Size            =   "8599;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label23 
      Height          =   195
      Index           =   0
      Left            =   3870
      TabIndex        =   27
      Top             =   2190
      Width           =   4875
      VariousPropertyBits=   27
      Caption         =   "CREATE：ID  Date  Time"
      Size            =   "8599;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "單字索引:是指在關鍵字的前後,會加上""空白""字元一同索引"
      ForeColor       =   &H000000C0&
      Height          =   555
      Left            =   7170
      TabIndex        =   25
      Top             =   1590
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(模糊比對)"
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   8
      Left            =   7320
      TabIndex        =   24
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   180
      Index           =   7
      Left            =   60
      TabIndex        =   23
      Top             =   5400
      Width           =   540
   End
   Begin VB.Label LblCnt 
      AutoSize        =   -1  'True
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   60
      TabIndex        =   22
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "使用信箱："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   4
      Left            =   6360
      TabIndex        =   21
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label LblLK11 
      Caption         =   "Y.比個案、 Initial 優先檢查"
      Height          =   195
      Left            =   4500
      TabIndex        =   20
      Top             =   1290
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "規則類別："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   3
      Left            =   3210
      TabIndex        =   19
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "優先排序："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   2
      Left            =   3210
      TabIndex        =   18
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "種類："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   510
      TabIndex        =   17
      Top             =   1020
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "分類："
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   6
      Left            =   510
      TabIndex        =   16
      Top             =   1290
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收受者："
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   5
      Left            =   330
      TabIndex        =   15
      Top             =   1620
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "關鍵字："
      ForeColor       =   &H00000080&
      Height          =   180
      Index           =   1
      Left            =   330
      TabIndex        =   14
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "frm06010614"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/10 Form2.0已修改
'Create by Sindy 2017/1/19
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 第一筆資料的Key
Dim m_FirstKEY(2) As String
' 最後一筆資料的Key
Dim m_LastKEY(2) As String
' 目前正在顯示的Key
Dim m_CurrKEY(2) As String
Dim i As Integer, j As Integer
Dim dblPrevRow As Double
Public m_strLK12 As String
Dim ii As Integer
Dim m_blnColOrderAsc As Boolean
Dim varTemp As Variant


'Add By Sindy 2017/3/14
Private Sub Combo1_Click()
   If Combo2.Enabled = True Then
      If m_EditMode = 1 Or m_EditMode = 2 Then '新增,修改時
         If Combo1.Text <> Combo1.Tag Then '分類
            lstUsers.Clear
            txtLK04 = "" 'Add By Sindy 2017/8/28
         End If
         If Left(Combo3, 1) = "F" Then 'IPDept
            If Left(Combo1.Text, 1) = "2" Then '2 外商
               Combo2.Text = "國外部轉信外商群組"
            ElseIf Left(Combo1.Text, 1) = "3" Then '3 外專
            ElseIf Left(Combo1.Text, 1) = "4" Then '4 專利處
               Combo2.Text = "patent"
            ElseIf Left(Combo1.Text, 1) = "5" Then '5 外法
            ElseIf Left(Combo1.Text, 1) = "6" Then '6 新知
               Combo2.Text = "國外部轉信新知群組"
            ElseIf Left(Combo1.Text, 1) = "7" Then '7 財務
               Combo2.Text = "account"
            ElseIf Left(Combo1.Text, 1) = "8" Then '8 開拓
               Combo2.Text = "國外部轉信開拓群組"
            End If
         ElseIf Left(Combo3, 1) = "P" Then 'Patent
            If Left(Combo1.Text, 1) = "1" Then '1 P程序1
               Combo2.Text = "專利處轉信非台灣程序1"
            ElseIf Left(Combo1.Text, 1) = "2" Then '2 P程序2
               Combo2.Text = "專利處轉信非台灣程序2"
            ElseIf Left(Combo1.Text, 1) = "3" Then '3 亞洲
               Combo2.Text = "專利處轉信美日單號程序"
            ElseIf Left(Combo1.Text, 1) = "4" Then '4 歐洲
               Combo2.Text = "專利處轉信美日雙號程序"
            ElseIf Left(Combo1.Text, 1) = "5" Then '5 美洋非(單)
               Combo2.Text = "專利處轉信美日以外單號程序"
            ElseIf Left(Combo1.Text, 1) = "6" Then '6 美洋非(雙)
               Combo2.Text = "專利處轉信美日以外雙號程序"
            'Add By Sindy 2020/3/18
            Else
               varTemp = Split(Combo1.Text, " ")
               If UBound(varTemp) > 0 Then
                  For ii = 0 To Combo5.ListCount - 1
                     If InStr(Combo5.List(ii), varTemp(0)) > 0 Then
                        varTemp = Split(Combo5.List(ii), " ")
                        Combo2.Text = varTemp(1)
                        Exit For
                     End If
                  Next ii
               End If
            '2020/3/18 END
            End If
         End If
         If Combo2.Text <> "" Then
            For i = 0 To lstUsers.ListCount - 1
               If lstUsers.List(i) = Combo2.Text Then
                  Combo2.Text = ""
               End If
            Next i
            If Combo2.Text <> "" Then Combo2_LostFocus: Call cmdAdd_Click
         End If
         Combo1.Tag = Combo1.Text
      End If
   End If
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
   If Combo3.Text = "" Then Exit Sub
   'Add By Sindy 2017/9/28
   If Combo3.Tag <> Combo3.Text Then '使用信箱: 使用單位
      Call SetCombo
   End If
   Combo3.Tag = Combo3.Text
   '2017/9/28 END
   If m_EditMode = 1 And txtLK01.Text <> "" Then
      ' 檢查記錄是否已存在
      If IsRecordExist(txtLK01, Left(Combo3, 1)) = True Then
         MsgBox "該筆記錄已存在", vbOKOnly, "更新資料"
         Cancel = True
         txtLK01.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub SetCombo3()
   For ii = 0 To Combo3.ListCount - 1
      If Left(Combo3.List(ii), 1) = m_strLK12 Then
         Combo3.ListIndex = ii
         Exit For
      End If
   Next ii
   'Add By Sindy 2017/9/28
   If Combo3.Tag <> Combo3.Text Then
      Call SetCombo
   End If
   Combo3.Tag = Combo3.Text
   '2017/9/28 END
   If Pub_StrUserSt03 = "M51" Then
      Combo3.Locked = False
   Else
      Combo3.Locked = True
   End If
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

Private Sub Form_Load()
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   ClearField
   
   SetCombo3
   SetDataListWidth
   
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   'ReadAllData
   'OnAction vbKeyF4
   OnAction vbKeyF10
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010614 = Nothing
End Sub

Private Sub SetCombo()
   LblLK11.Caption = "Y.比個案、 Initial 優先檢查"
   If Combo3.Text <> Combo3.Tag Then
      If Left(Combo3.Text, 1) = "F" Then 'IPDept.國外部
         Combo1.Clear
         Combo1.AddItem "2 外商"
         Combo1.AddItem "3 外專"
         Combo1.AddItem "4 專利處"
         Combo1.AddItem "5 外法"
         Combo1.AddItem "6 新知"
         Combo1.AddItem "7 財務"
         Combo1.AddItem "8 開拓"
         Combo1.AddItem "Z 其他"
         '收受者
         Combo2.Clear
         Combo2.AddItem ""
         Combo2.AddItem "patent"
         Combo2.AddItem "account"
         Combo2.AddItem "國內信件管理人員" 'Add By Sindy 2023/7/27
         Combo2.AddItem "國外部轉信外專群組"
         Combo2.AddItem "國外部轉信外專承辦英文組長"
         Combo2.AddItem "國外部轉信外專承辦日文組長"
         Combo2.AddItem "國外部轉信外專非日文承辦主管"
         Combo2.AddItem "國外部轉信外商群組"
         Combo2.AddItem "國外部轉信外法群組"
         Combo2.AddItem "國外部轉信外法英文組群組"
         Combo2.AddItem "國外部轉信外法日文組群組"
         Combo2.AddItem "國外部轉信新知群組"
         Combo2.AddItem "國外部轉信開拓群組"
         Combo2.AddItem "寰華案外專程序窗口"
         'Add By Sindy 2017/3/28 + 外專承辦組人員
         strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' and st03='F23' order by st03,st01 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               RsTemp.MoveFirst
               Do While RsTemp.EOF = False
                  Combo2.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
                  RsTemp.MoveNext
               Loop
            End With
         End If
         '2017/3/28 END
      ElseIf Left(Combo3.Text, 1) = "P" Then 'Patent.專利處
         Combo1.Clear
         'Add By Sindy 2025/1/9
         If strSrvDate(1) >= P業務區劃分啟用日 Then
            Call PUB_AddItemCFPHandler(Combo1, Combo5, , "P")
         Else
         '2025/1/9 END
            Combo1.AddItem "1 P程序1"
            Combo1.AddItem "2 P程序2"
         End If
         'Modify by Sindy 2020/3/18
         '109/4/1以後改業務區劃分
         If strSrvDate(1) >= CFP業務區劃分啟用日 Then
            Call PUB_AddItemCFPHandler(Combo1, Combo5)
         Else
         '2020/3/18 END
            Combo1.AddItem "3 美日(單)"
            Combo1.AddItem "4 美日(雙)"
            Combo1.AddItem "5 美日外(單)"
            Combo1.AddItem "6 美日外(雙)"
         End If
         Combo1.AddItem "7 其他"
         Combo1.AddItem "8 垃圾信箱"
         '收受者
         Combo2.Clear
         Combo2.AddItem ""
         strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' and st03>='P10' and st03<='P19' order by st03,st01 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               RsTemp.MoveFirst
               Do While RsTemp.EOF = False
                  Combo2.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
                  RsTemp.MoveNext
               Loop
            End With
         End If
      'Add By Sindy 2019/4/3
      ElseIf Left(Combo3.Text, 1) = "T" Then 'TM.商標處
         LblLK11.Caption = "甲類 乙類 丙類 或 空白"
         Combo1.Clear
         Combo1.AddItem "1 MCTF"
         Combo1.AddItem "2 大陸案"
         Combo1.AddItem "3 個人"
         Combo1.AddItem "4 非大陸案"
         Combo1.AddItem "5 其他"
         '收受者
         Combo2.Clear
         Combo2.AddItem ""
         Combo2.AddItem "MCTF01 MCTF01"
         Combo2.AddItem "MCTF02 MCTF02"
         Combo2.AddItem "MCTF03 MCTF03"
         Combo2.AddItem "MCTF04 MCTF04"
         Combo2.AddItem "MCTF05 MCTF05"
         strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' and st03>='P20' and st03<='P29' and st01 not in('96029','96030') order by st03,st01 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               RsTemp.MoveFirst
               Do While RsTemp.EOF = False
                  Combo2.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
                  RsTemp.MoveNext
               Loop
            End With
         End If
         'Combo2.AddItem "76012 " & GetPrjSalesNM("76012") 'Modify By Sindy 2023/7/11 Mark
         Combo2.AddItem "98020 " & GetPrjSalesNM("98020")  'add by sonia 2021/11/9
      '2019/4/3 END
      Else
         Combo1.Clear
         Combo2.Clear
      End If
   End If
End Sub

Private Sub Grd1_Click()
Dim nRow As Long, nCol As Long
   
   grd1.Visible = False
   grd1.row = grd1.MouseRow
   grd1.col = grd1.MouseCol
   nRow = grd1.row
   nCol = grd1.col
   If nRow = 0 Then
      If Me.grd1.row < 1 And Me.grd1.Text <> "V" Then
         If Me.grd1.Text = "無" Then
            If m_blnColOrderAsc = True Then
               Me.grd1.Sort = 3  '數值昇冪
               m_blnColOrderAsc = False
            Else
               Me.grd1.Sort = 4 '數值降冪
               m_blnColOrderAsc = True
            End If
         Else
            If m_blnColOrderAsc = True Then
               Me.grd1.Sort = 5 '字串昇冪
               m_blnColOrderAsc = False
            Else
               Me.grd1.Sort = 6 '字串降冪
               m_blnColOrderAsc = True
            End If
         End If
      End If
   End If
   grd1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grd1, x, y, nCol, nRow
If nCol < 0 Then Exit Sub
If nRow < 0 Then Exit Sub
grd1.col = nCol
grd1.row = nRow
End Sub

Private Sub grd1_SelChange()
grd1.Visible = False
'If grd1.MouseRow <> 0 Then
If grd1.row <> 0 Then
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
   m_CurrKEY(0) = grd1.TextMatrix(grd1.row, 1) '關鍵字
   m_CurrKEY(1) = grd1.TextMatrix(grd1.row, 12) '使用信箱: 使用單位
   UpdateCtrlData
   If m_CurrKEY(1) <> Left(Trim(Combo3), 1) Then
      RefreshRange
   End If
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
   
   If txtLK01.Text = "" Then
       MsgBox "關鍵字不可以空白！", vbExclamation
       txtLK01.SetFocus
       Exit Function
   End If
   
   If Combo3.Text = "" Then
       MsgBox "使用信箱不可以空白！", vbExclamation
       Combo3.SetFocus
       Exit Function
   End If
   
   If m_EditMode = 1 Then
      ' 檢查記錄是否已存在
      If IsRecordExist(txtLK01, Left(Combo3, 1)) = True Then
         MsgBox "該筆記錄已存在", vbOKOnly, "更新資料"
         txtLK01.SetFocus
         Exit Function
      End If
   End If
   
   If Combo4.Text = "" Then
       MsgBox "種類不可以空白！", vbExclamation
       Combo4.SetFocus
       Exit Function
   End If
   
   If Combo1.Text = "" Then
       MsgBox "分類不可以空白！", vbExclamation
       Combo1.SetFocus
       Exit Function
   End If
   
   'Add By Sindy 2020/7/6
   Cancel = False
   txtLK11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   
   'Add By Sindy 2017/3/14
   'Modify By Sindy 2017/7/6 其他可以不輸入收受者
   'If txtLK04.Text = "" And Left(Trim(Combo1), 1) <> "Z" Then
   If txtLK04.Text = "" And InStr(Trim(Combo1), "其他") = 0 And InStr(Trim(Combo1), "垃圾信箱") = 0 Then
       MsgBox "收受者不可以空白！", vbExclamation
       Combo2.SetFocus
       Exit Function
   End If
   'Add By Sindy 2025/9/22
   If Me.Tag <> "" Then
      MsgBox Me.Tag, vbExclamation
      Exit Function
   End If
   '2025/9/22 END
   
   Cancel = False
   txtLK01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   
   Cancel = False
   txtLK13_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   
   Cancel = False
   Combo2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If

   'Add by Sindy 2021/4/27 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/4/27 END

   TxtValidate = True
End Function

' 更新資料
Private Function SaveData(strEditMode As Integer) As Boolean
Dim strKEY01 As String, strKEY02 As String
   
On Error GoTo ErrHand
   
   SaveData = False
   
   'Modify By Sindy 2018/7/6
'   'Add By Sindy 2018/1/10
'   If ChkLK14.Value = 1 And strEditMode = 1 Then
'      txtLK01 = Trim(txtLK01)
'   End If
'   '2018/1/10 END
   strKEY01 = txtLK01
   strKEY02 = Left(Combo3, 1)
   
   cnnConnection.BeginTrans
   '新增
   If strEditMode = 1 Then
      strSql = "INSERT INTO IPDeptKeyWord(LK01,LK02,LK03,LK04,LK11,LK12" & IIf(Trim(txtLK13) <> "", ",LK13", "") & ",LK14) VALUES(" & _
                  CNULL(ChgSQL(strKEY01)) & "," & CNULL(Trim(Left(Combo1, 2))) & "," & CNULL(Left(Combo4.Text, 1)) & _
                  "," & CNULL(txtLK04) & "," & CNULL(txtLK11) & _
                  "," & CNULL(strKEY02) & IIf(Trim(txtLK13) <> "", "," & CNULL(txtLK13), "") & "," & IIf(ChkLK14.Value = 1, "'Y'", "null") & ")"
   '修改
   ElseIf strEditMode = 2 Then
      strSql = "UPDATE IPDeptKeyWord SET " & _
                  "LK02=" & CNULL(Trim(Left(Combo1, 2))) & ",LK03=" & CNULL(Left(Combo4.Text, 1)) & _
                  ",LK04=" & CNULL(txtLK04) & ",LK11=" & CNULL(txtLK11) & _
                  ",LK14=" & IIf(ChkLK14.Value = 1, "'Y'", "null")
      If Trim(txtLK13) <> "" Then
         strSql = strSql & ",LK13=" & CNULL(txtLK13)
      End If
      strSql = strSql & " WHERE LK01=" & CNULL(ChgSQL(txtLK01)) & " and LK12=" & CNULL(strKEY02)
   End If
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   
   If (strKEY01 & strKEY02 < m_FirstKEY(0) & m_FirstKEY(1)) Or (strKEY01 & strKEY02 > m_LastKEY(0) & m_LastKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strKEY01, strKEY02
   
   SaveData = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 更新失敗！" & vbCrLf & Err.Description
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   
On Error GoTo ErrHand
   
   DelRecord = False
   
   cnnConnection.BeginTrans
   
   strSql = "DELETE FROM IPDeptKeyWord WHERE LK01 = " & CNULL(ChgSQL(txtLK01)) & " and LK12 = " & CNULL(Left(Combo3, 1))
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
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String, strTemp As String
   
   QueryRecord = False
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   m_blnColOrderAsc = False
   
   'Add By Sindy 2023/6/21
   txtLK01.Tag = ""
   Combo1.Tag = ""
   txtLK04.Tag = ""
   '2023/6/21 END
   strCon = ""
   '關鍵字
   If Trim(txtLK01) <> "" Then
      'Modify By Sindy 2018/7/5 查詢時輸入的條件要Trim前後空白, 這樣才能抓到全部的資料
      '                         ex:"FICPI",可以查到"FICPI","FICPI "
      'strCon = strCon & " and instr(upper(LK01),upper('" & txtLK01 & "'))>0 "
      'Modify By Sindy 2024/8/12 +ChgSQL()
      strCon = strCon & " and instr(upper(ltrim(rtrim(LK01))),upper(ltrim(rtrim('" & ChgSQL(txtLK01) & "'))))>0 "
      txtLK01.Tag = txtLK01.Text
   End If
   '使用信箱: 使用單位
   If Trim(Combo3.Text) <> "" Then
      strCon = strCon & " and LK12='" & Left(Trim(Combo3.Text), 1) & "' "
   End If
   'Add By Sindy 2017/2/21 加入分類做查詢條件
   If Trim(Combo1.Text) <> "" Then
      strCon = strCon & " and LK02='" & Trim(Left(Combo1.Text, 2)) & "' "
      Combo1.Tag = Trim(Combo1.Text) 'Add By Sindy 2022/7/25
   End If
   '2017/2/21 END
   'Add By Sindy 2018/5/15 加入收受者做查詢條件
   If lstUsers.ListCount > 0 Then
      varTemp = Split(txtLK04, ";")
      strTemp = ""
      For ii = 0 To UBound(varTemp)
         strTemp = strTemp & "and instr(LK04,'" & varTemp(ii) & "')>0"
      Next ii
      If strTemp <> "" Then
         strTemp = Trim(Mid(strTemp, 4))
         strCon = strCon & " and (" & strTemp & ")"
      End If
      txtLK04.Tag = txtLK04 'Add By Sindy 2022/7/25
   End If
   '2017/2/21 END
   
   grd1.Rows = 2
   grd1.Clear
   dblPrevRow = 0
   strSql = "select '' V,LK01 關鍵字,decode(LK03,'1','主旨','2','寄件者或網域','3','收件者',LK03) 種類" & _
            ",decode(LK12,'F',decode(LK02," & Show國外部信件分類 & ",LK02),'P',decode(LK02," & Show專利處信件分類 & ",LK02),'T',decode(LK02," & Show商標處信件分類 & ",LK02)) 分類" & _
            ",GETSTAFFNAMELIST(replace(LK04,';',',')) 收受者,LK13 優先排序" & _
            ",LK11 規則類別,s1.st02||' '||sqldatet(LK06)||' '||LK07 新增人員日期" & _
            ",s2.st02||' '||sqldatet(LK09)||' '||LK10 更新人員日期" & _
            ",decode(LK12,'F','IPDept','P','Patent','T','TM',LK12) 使用單位" & _
            ",LK14 單字索引,LK03,LK12" & _
            " from IPDeptKeyWord,staff s1,staff s2" & _
            " where LK05=s1.st01(+) and LK02<>'S'" & _
            " and LK08=s2.st01(+)" & strCon & _
            " order by LK01 asc,LK12 asc"
            'LK13 asc,LK01 asc,LK12 asc
            'LK12 asc,LK02 asc,LK03 asc,LK13 asc,LK01||LK12 asc
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   LblCnt.Caption = ""
   If rsTmp.RecordCount > 0 Then
      Set grd1.Recordset = rsTmp
      LblCnt.Caption = rsTmp.RecordCount & " 筆"
'      '解析收受者
'      For i = 1 To grd1.Rows - 1
'         grd1.TextMatrix(i, 4) = PUB_ReadUserData(grd1.TextMatrix(i, 4))
'      Next i
      QueryRecord = True
      rsTmp.MoveFirst
      m_CurrKEY(0) = rsTmp.Fields("關鍵字")
      m_CurrKEY(1) = rsTmp.Fields("LK12")
      UpdateCtrlData
      dblPrevRow = 1
   End If
   rsTmp.Close
   SetDataListWidth
   GetSelChage
   
   Call QuerySysSet 'Add By Sindy 2017/10/20
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   UpdateToolbarState
   RefreshRange
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'Add By Sindy 2017/10/20
Private Sub QuerySysSet()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   ListSysSet.Clear
   If Left(Trim(Combo3.Text), 1) = "F" Then 'IPDept: 國外部
      'Modify By Sindy 2022/8/22 取消;關鍵字設定即可
      ListSysSet.AddItem "[111/8/22取消] 主旨＝APAA 並且(有??? 或 寄信者網域為 jp)   -- 轉寄給 王文安;國外部轉信開拓群組"
      ListSysSet.AddItem "[111/8/22取消] 主旨＝APAA 並且(無??? 並且 寄信者網域非 jp) -- 轉寄給 洪琬姿;國外部轉信開拓群組"
   End If
   'Add By Sindy 2018/3/6
   ListSysSet.AddItem "寄件者＝tipo@tiponet.tipo.gov.tw並且主旨＝(收件成功通知 或 電子公文通知) -- 不分信直接刪除"
   strSql = "select decode(LK03,'1','主旨','2','寄件者或網域','3','收件者',LK03),LK04,LK01" & _
            " from ipdeptkeyword" & _
            " where LK12='" & Left(Trim(Combo3.Text), 1) & "' and lk02='S'" & _
            " order by LK13 asc,LK01 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      For j = 1 To rsTmp.RecordCount
         strSql = rsTmp.Fields(0) & "＝"
         '解析收受者
         varTemp = Split(rsTmp.Fields(1), ";")
         strSql = strSql & PUB_ReadUserData(rsTmp.Fields(1), True)
         strSql = strSql & " -- " & rsTmp.Fields(2)
         ListSysSet.AddItem strSql
      Next j
   End If
   
   'Add By Sindy 2019/4/18
   ListSysSet.AddItem "檢索簡體字的文字檔,原始檔存於\\Linux\polycom\TaieNew\TaRevOutLook"
   ListSysSet.AddItem "　　　　　　　　　 執行檔存於\\M51-WIN7\Program Files\AutoBatch"
   ListSysSet.AddItem "　專利處檢索簡體字的文字檔名:executePatent.txt"
   ListSysSet.AddItem "　商標處檢索簡體字的文字檔名:executeTM.txt"
   ListSysSet.AddItem "優先順序：小至大，數值小為優先" 'Add By Sindy 2020/3/6
   
   'If ListSysSet.ListCount > 0 Then Call PUB_SetListScroll(Me, ListSysSet)
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

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
         If txtLK01.Tag <> "" Or Combo1.Tag <> "" Or txtLK04.Tag <> "" Then
            txtLK01.Text = txtLK01.Tag
            Combo1.Text = Combo1.Tag
            txtLK04.Text = txtLK04.Tag
            QueryRecord
         Else
            ReadAllData
         End If
         SetKeyReadOnly True
      Case 3: '刪除
         If txtLK01.Text = "" Then
             MsgBox "關鍵字不可以空白！", vbExclamation
             txtLK01.SetFocus
             GoTo EXITSUB
         End If
         If Combo3.Text = "" Then
             MsgBox "使用信箱不可以空白！", vbExclamation
             GoTo EXITSUB
         End If
         If DelRecord = True Then
            RefreshRange
            ClearField
            If txtLK01.Tag <> "" Or Combo1.Tag <> "" Or txtLK04.Tag <> "" Then
               txtLK01.Text = txtLK01.Tag
               Combo1.Text = Combo1.Tag
               txtLK04.Text = txtLK04.Tag
               If QueryRecord = False Then
                  ReadAllData
               End If
            Else
               ReadAllData
            End If
            SetKeyReadOnly True
         Else
            Exit Function
         End If
      Case 4: '查詢
'         If txtLK01.Text = "" Then
'             MsgBox "關鍵字不可以空白！", vbExclamation
'             txtLK01.SetFocus
'         End If
'         If txtLK01 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
               ReadAllData
            End If
            SetKeyReadOnly True
'         Else
'            GoTo EXITSUB
'         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
   
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 0, 1, 4: If Me.txtLK01.Visible = True Then txtLK01.SetFocus
      Case 2: Combo4.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   'strSql = "SELECT * FROM IPDeptKeyWord WHERE instr(upper(LK01)," & CNULL(Trim(UCase(ChgSQL(strKEY01)))) & ")>0 and LK12=" & CNULL(strKEY02)
   'Modify By Sindy 2018/7/5 不要Trim前後空白比對 ex:關鍵字 FICPI 已設為網域寄件者; 無法在主旨再設關鍵字 FICPI
   'strSql = "SELECT * FROM IPDeptKeyWord WHERE upper(ltrim(rtrim(LK01)))=" & CNULL(Trim(UCase(ChgSQL(strKEY01)))) & " and LK02<>'S' and LK12=" & CNULL(strKEY02)
   strSql = "SELECT * FROM IPDeptKeyWord WHERE upper(LK01)=" & CNULL(UCase(ChgSQL(strKEY01))) & " and LK02<>'S' and LK12=" & CNULL(strKEY02)
   
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
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "SELECT LK01,LK12,LK13 FROM IPDeptKeyWord WHERE LK01='" & ChgSQL(m_CurrKEY(0)) & "' and LK02<>'S' and LK12='" & m_CurrKEY(1) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT LK01,LK12,LK13 FROM IPDeptKeyWord"
      '使用信箱: 使用單位
      If Trim(Combo3.Text) <> "" Then
         strSql = strSql & " WHERE LK12='" & Left(Trim(Combo3.Text), 1) & "' and LK02<>'S'"
      Else
         strSql = strSql & " WHERE LK02<>'S'"
      End If
      'strSql = strSql & " order by LK13 asc,LK01 asc,LK12 asc"
      'strSql = strSql & " order by LK12 asc,LK02 asc,LK03 asc,LK13 asc,LK01||LK12 asc"
      strSql = strSql & " order by LK01 asc,LK12 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
         If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      If Combo3.Locked = False Then
         strSql = "SELECT LK01,LK12,LK13 FROM IPDeptKeyWord where LK02<>'S'"
         'strSql = strSql & " order by LK12 asc,LK02 asc,LK03 asc,LK13 asc,LK01||LK12 asc"
         strSql = strSql & " order by LK01 asc,LK12 asc"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
            If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
         Else
            ShowLastRecord
            GoTo EXITSUB
         End If
         rsTmp.Close
      End If
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   ReadAllData 'Add By Sindy 2017/2/6
   strSql = "SELECT LK01,LK12,LK13 FROM IPDeptKeyWord WHERE LK01||LK12<'" & ChgSQL(m_CurrKEY(0)) & m_CurrKEY(1) & "' and LK02<>'S'"
   If Trim(Combo3.Text) <> "" Then
      strSql = strSql & " and LK12='" & Left(Trim(Combo3.Text), 1) & "'"
   End If
   'strSql = strSql & " order by LK13 desc,LK01 desc,LK12 desc"
   'strSql = strSql & " order by LK12 desc,LK02 desc,LK03 desc,LK13 desc,LK01||LK12 desc"
   strSql = strSql & " order by LK01 desc,LK12 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT LK01,LK12,LK13 FROM IPDeptKeyWord"
   If Trim(Combo3.Text) <> "" Then
      strSql = strSql & " where LK12='" & Left(Trim(Combo3.Text), 1) & "' and LK02<>'S'"
   Else
      strSql = strSql & " where LK02<>'S'"
   End If
   'strSql = strSql & " order by LK13 desc,LK01 desc,LK12 desc"
   'strSql = strSql & " order by LK12 desc,LK02 desc,LK03 desc,LK13 desc,LK01||LK12 desc"
   strSql = strSql & " order by LK01 desc,LK12 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   ReadAllData 'Add By Sindy 2017/2/6
   strSql = "SELECT LK01,LK12,LK13 FROM IPDeptKeyWord WHERE LK01||LK12>'" & ChgSQL(m_CurrKEY(0)) & m_CurrKEY(1) & "' and LK02<>'S'"
   If Trim(Combo3.Text) <> "" Then
      strSql = strSql & " and LK12='" & Left(Trim(Combo3.Text), 1) & "'"
   End If
   'strSql = strSql & " order by LK13 asc,LK01 asc,LK12 asc"
   'strSql = strSql & " order by LK12 asc,LK02 asc,LK03 asc,LK13 asc,LK01||LK12 asc"
   strSql = strSql & " order by LK01 asc,LK12 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT LK01,LK12,LK13 FROM IPDeptKeyWord"
   If Trim(Combo3.Text) <> "" Then
      strSql = strSql & " where LK12='" & Left(Trim(Combo3.Text), 1) & "' and LK02<>'S'"
   Else
      strSql = strSql & " where LK02<>'S'"
   End If
   'strSql = strSql & " order by LK13 asc,LK01 asc,LK12 asc"
   'strSql = strSql & " order by LK12 asc,LK02 asc,LK03 asc,LK13 asc,LK01||LK12 asc"
   strSql = strSql & " order by LK01 asc,LK12 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then: m_CurrKEY(1) = rsTmp.Fields(1)
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
'         If Pub_StrUserSt03 <> "M51" Then
'            For ii = 0 To Combo3.ListCount - 1
'               If Left(Combo3.List(ii), 1) = m_strLK12 Then
'                  Combo3.ListIndex = ii
'                  Exit For
'               End If
'            Next ii
'            Combo3.Locked = True
'         End If
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         'Call txtB0201_LostFocus
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
         dblPrevRow = 0
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
               ReadAllData
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
   txtLK01.Locked = bEnable
   If bEnable Then txtLK01.BackColor = &H8000000F Else txtLK01.BackColor = &H80000005
   If Pub_StrUserSt03 = "M51" Then '電腦中心才開放此欄位
      Combo3.Locked = bEnable
      If bEnable Then Combo3.BackColor = &H8000000F Else Combo3.BackColor = &H80000005
   Else
      Combo3.Locked = True
      Combo3.BackColor = &H8000000F
   End If
   If m_EditMode <> 2 Then
      Combo1.Locked = bEnable
      If bEnable Then Combo1.BackColor = &H8000000F Else Combo1.BackColor = &H80000005
      'Add By Sindy 2018/5/15
      Combo2.Locked = bEnable
      If bEnable Then Combo2.BackColor = &H8000000F Else Combo2.BackColor = &H80000005
      cmdAdd.Enabled = Not bEnable
      If Not bEnable Then cmdAdd.BackColor = &H8000000F Else cmdAdd.BackColor = &H80000005
      cmdRemove.Enabled = Not bEnable
      If Not bEnable Then cmdRemove.BackColor = &H8000000F Else cmdRemove.BackColor = &H80000005
      '2018/5/15 END
   End If
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   txtLK13.Locked = bEnable
   If bEnable Then txtLK13.BackColor = &H8000000F Else txtLK13.BackColor = &H80000005
   txtLK11.Locked = bEnable
   If bEnable Then txtLK11.BackColor = &H8000000F Else txtLK11.BackColor = &H80000005
   Combo1.Locked = bEnable
   If bEnable Then Combo1.BackColor = &H8000000F Else Combo1.BackColor = &H80000005
   Combo2.Locked = bEnable
   If bEnable Then Combo2.BackColor = &H8000000F Else Combo2.BackColor = &H80000005
   cmdAdd.Enabled = Not bEnable
   If Not bEnable Then cmdAdd.BackColor = &H8000000F Else cmdAdd.BackColor = &H80000005
   cmdRemove.Enabled = Not bEnable
   If Not bEnable Then cmdRemove.BackColor = &H8000000F Else cmdRemove.BackColor = &H80000005
   'Add By Sindy 2018/1/10
   ChkLK14.Enabled = Not bEnable
   If Not bEnable Then ChkLK14.BackColor = &H8000000F Else ChkLK14.BackColor = &H80000005
   '2018/1/10 END
   'Add By Sindy 2018/5/15
   Combo4.Locked = bEnable
   If bEnable Then Combo4.BackColor = &H8000000F Else Combo4.BackColor = &H80000005
   '2018/5/15 END
End Sub

Private Sub ClearField()
   txtLK01 = Empty
   Combo4.ListIndex = -1
   txtLK13 = Empty
   Combo1.ListIndex = -1
   txtLK11 = Empty
   lstUsers.Clear
   txtLK04 = Empty
   Combo2.Text = ""
'   Combo3.ListIndex = 0
'   'Add By Sindy 2017/9/28
'   If Combo3.Tag <> Combo3.Text Then
'      Call SetCombo
'   End If
'   Combo3.Tag = Combo3.Text
'   '2017/9/28 END
   Label23(0).Caption = ""
   Label23(1).Caption = ""
   ChkLK14.Value = 0 'Add By Sindy 2018/9/6
   Me.Tag = "" 'Add By Sindy 2025/9/22
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub ReadAllData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   txtLK01.Tag = ""
   Combo1.Tag = ""
   txtLK04.Tag = ""
   grd1.Rows = 2
   grd1.Clear
   strSql = "select '' V,LK01 關鍵字,decode(LK03,'1','主旨','2','寄件者或網域',LK03) 種類" & _
            ",decode(LK12,'F',decode(LK02," & Show國外部信件分類 & ",LK02),'P',decode(LK02," & Show專利處信件分類 & ",LK02),'T',decode(LK02," & Show商標處信件分類 & ",LK02)) 分類" & _
            ",GETSTAFFNAMELIST(replace(LK04,';',',')) 收受者,LK13 優先排序" & _
            ",LK11 規則類別,s1.st02||' '||sqldatet(LK06)||' '||LK07 新增人員日期" & _
            ",s2.st02||' '||sqldatet(LK09)||' '||LK10 更新人員日期" & _
            ",decode(LK12,'F','國外部','P','專利處','T','商標處','L','法務部','S','智權部',LK12) 使用信箱" & _
            ",LK14 單字索引,LK03,LK12" & _
            " from IPDeptKeyWord,staff s1,staff s2" & _
            " where LK05=s1.st01(+) and LK02<>'S'" & _
            " and LK08=s2.st01(+)"
   If Trim(Combo3.Text) <> "" Then
      strSql = strSql & " and LK12='" & Left(Trim(Combo3.Text), 1) & "'"
   End If
   'strSql = strSql & " order by LK13 asc,LK01 asc,LK12 asc"
   'strSql = strSql & " order by LK12 asc,LK02 asc,LK03 asc,LK13 asc,LK01||LK12 asc"
   'Modify By Sindy 2017/10/18 改為顯示程式抓關鍵字的順序
   'strSql = strSql & " order by LK01 asc,LK12 asc"
   strSql = strSql & " order by LK13 asc,LK01 asc"
   '2017/10/18 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   LblCnt.Caption = ""
   If rsTmp.RecordCount > 0 Then
      Set grd1.Recordset = rsTmp
      LblCnt.Caption = rsTmp.RecordCount & " 筆"
'      '解析收受者
'      For i = 1 To GRD1.Rows - 1
'         GRD1.TextMatrix(i, 4) = PUB_ReadUserData(GRD1.TextMatrix(i, 13))
'      Next i
   End If
   rsTmp.Close
   SetDataListWidth
   GetSelChage
   
   Call QuerySysSet 'Add By Sindy 2017/10/20
   
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
      For i = 0 To grd1.Cols - 1
         grd1.col = i
         grd1.CellBackColor = QBColor(15)
      Next i
   End If
   '尋找目前資料列
   For j = 1 To grd1.Rows - 1
      If grd1.TextMatrix(j, 1) = m_CurrKEY(0) And grd1.TextMatrix(j, 12) = m_CurrKEY(1) Then
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
Dim tmpArr As Variant, strTempName As String
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   ClearField
   strSql = "select *" & _
            " from IPDeptKeyWord" & _
            " where LK01='" & ChgSQL(m_CurrKEY(0)) & "' and LK12='" & m_CurrKEY(1) & "' and LK02<>'S'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If m_EditMode = 1 And txtLK01.Enabled = True Then
         '關鍵字欄位為空白,使用者自行輸入欲新增的關鍵字
      Else
         If IsNull(rsTmp.Fields("LK01")) = False Then txtLK01 = rsTmp.Fields("LK01")
      End If
      If IsNull(rsTmp.Fields("LK12")) = False Then
         For ii = 0 To Combo3.ListCount - 1
            If Left(Combo3.List(ii), 1) = rsTmp.Fields("LK12") Then
               Combo3.ListIndex = ii
               Exit For
            End If
         Next ii
         'Add By Sindy 2017/9/28
         If Combo3.Tag <> Combo3.Text Then
            Call SetCombo
         End If
         Combo3.Tag = Combo3.Text
         '2017/9/28 END
      End If
      If IsNull(rsTmp.Fields("LK02")) = False Then '分類
         For ii = 0 To Combo1.ListCount - 1
            If Trim(Left(Combo1.List(ii), 2)) = rsTmp.Fields("LK02") Then
               Combo1.ListIndex = ii
               Exit For
            End If
         Next ii
      End If
      If IsNull(rsTmp.Fields("LK03")) = False Then
         Combo4.ListIndex = Val(rsTmp.Fields("LK03")) - 1
      End If
      '收受者
      txtLK04 = "" & rsTmp.Fields("LK04")
      SetlstUsers "" & rsTmp.Fields("LK04")
      
      If IsNull(rsTmp.Fields("LK11")) = False Then txtLK11 = rsTmp.Fields("LK11")
      If IsNull(rsTmp.Fields("LK13")) = False Then txtLK13 = rsTmp.Fields("LK13")
      'Add By Sindy 2018/1/10
      If "" & rsTmp.Fields("LK14") = "Y" Then
         ChkLK14.Value = 1
      Else
         ChkLK14.Value = 0
      End If
      '2018/1/10 END
      
      Combo1.Tag = Combo1.Text 'Add By Sindy 2017/3/14
      ' 更新CUID
      UpdateCUID rsTmp
   End If
   rsTmp.Close
   GetSelChage
   
   Me.Enabled = True
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("LK05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LK05")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("LK05"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("LK06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LK06")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("LK06"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("LK07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LK07")) = False Then
         strTemp = rsSrcTmp.Fields("LK07")
         strCTime = Format(strTemp, "00:00:00")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("LK08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LK08")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("LK08"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("LK09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LK09")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("LK09"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("LK10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("LK10")) = False Then
         strTemp = rsSrcTmp.Fields("LK10")
         strUTime = Format(strTemp, "00:00:00")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23(0).Caption = "CREATE：" & strCName & _
                        " " & strCDate & _
                        " " & strCTime
   Label23(1).Caption = "UPDATE：" & strUName & _
                        " " & strUDate & _
                        " " & strUTime
End Sub

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "select LK01,LK12,LK13 from IPDeptKeyWord"
   If Trim(Combo3.Text) <> "" Then
      strSql = strSql & " where LK12='" & Left(Trim(Combo3.Text), 1) & "' and LK02<>'S'"
   Else
      strSql = strSql & " where LK02<>'S'"
   End If
   'strSql = strSql & " order by LK13 asc,LK01 asc,LK12 asc"
   'strSql = strSql & " order by LK12 asc,LK02 asc,LK03 asc,LK13 asc,LK01||LK12 asc"
   strSql = strSql & " and rownum<=10 order by LK01 asc,LK12 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then m_FirstKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then m_FirstKEY(1) = rsTmp.Fields(1)
   End If
   rsTmp.Close
   
   strSql = "select LK01,LK12,LK13 from IPDeptKeyWord"
   If Trim(Combo3.Text) <> "" Then
      strSql = strSql & " where LK12='" & Left(Trim(Combo3.Text), 1) & "' and LK02<>'S'"
   Else
      strSql = strSql & " where LK02<>'S'"
   End If
   'strSql = strSql & " order by LK13 desc,LK01 desc,LK12 desc"
   'strSql = strSql & " order by LK12 desc,LK02 desc,LK03 desc,LK13 desc,LK01||LK12 desc"
   strSql = strSql & " and rownum<=10 order by LK01 desc,LK12 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then m_LastKEY(0) = rsTmp.Fields(0)
      If IsNull(rsTmp.Fields(1)) = False Then m_LastKEY(1) = rsTmp.Fields(1)
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
grd1.col = 0: grd1.Text = "V"
grd1.ColWidth(0) = 200
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 1: grd1.Text = "關鍵字"
grd1.ColWidth(1) = 1200
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 2: grd1.Text = "種類"
grd1.ColWidth(2) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 3: grd1.Text = "分類"
grd1.ColWidth(3) = 850
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 4: grd1.Text = "收受者"
grd1.ColWidth(4) = 2000
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 5: grd1.Text = "優先排序"
grd1.ColWidth(5) = 500
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 6: grd1.Text = "規則類別"
grd1.ColWidth(6) = 500
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 7: grd1.Text = "新增人員日期"
grd1.ColWidth(7) = 900
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 8: grd1.Text = "更新人員日期"
grd1.ColWidth(8) = 900
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 9: grd1.Text = "使用信箱"
If Pub_StrUserSt03 = "M51" Then
   grd1.ColWidth(9) = 800
Else
   grd1.ColWidth(9) = 0
End If
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 10: grd1.Text = "單字索引"
grd1.ColWidth(10) = 800
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 11: grd1.Text = "LK03"
grd1.ColWidth(11) = 0
grd1.CellAlignment = flexAlignLeftCenter
grd1.col = 12: grd1.Text = "LK12"
grd1.ColWidth(12) = 0
grd1.CellAlignment = flexAlignLeftCenter
End Sub

Private Sub txtLK01_GotFocus()
   InverseTextBox txtLK01
End Sub

Private Sub txtLK01_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtLK01
End Sub

Private Sub txtLK01_Validate(Cancel As Boolean)
   If txtLK01.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtLK01, txtLK01.MaxLength) Then
      Cancel = True
   End If
   If m_EditMode = 1 And Combo3.Text <> "" Then
      ' 檢查記錄是否已存在
      If IsRecordExist(txtLK01, Left(Combo3, 1)) = True Then
         MsgBox "該筆記錄已存在", vbOKOnly, "更新資料"
         Cancel = True
         txtLK01.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub txtLK11_GotFocus()
   InverseTextBox txtLK11
End Sub

Private Sub txtLK11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Left(Combo3, 1) <> "T" Then
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         MsgBox "「規則類別」只能輸入 Y 或空白 !!!", vbExclamation + vbOKOnly
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtLK11_Validate(Cancel As Boolean)
   If txtLK11 <> "" Then
      If Left(Combo3, 1) = "T" Then
         'If KeyAscii <> 65 And KeyAscii <> 66 And KeyAscii <> 67 And KeyAscii <> 8 Then
         If txtLK11 <> "甲" And txtLK11 <> "乙" And txtLK11 <> "丙" Then
            MsgBox "「規則類別」只能輸入 甲,乙,丙 或空白 !!!", vbExclamation + vbOKOnly
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtLK13_GotFocus()
   InverseTextBox txtLK13
End Sub

Private Sub txtLK13_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtLK13_Validate(Cancel As Boolean)
   If Not CheckLengthIsOK(txtLK13, txtLK13.MaxLength) Then
      Cancel = True
   End If
End Sub

'加項
Private Sub cmdAdd_Click()
Dim strText As String 'Add By Sindy 2025/9/22
   
   Me.Tag = Trim(Combo2.Text) 'Add By Sindy 2025/9/22
   AddlstUsers
   strText = ComposeListX
   txtLK04 = strText
   'Add By Sindy 2025/9/22
   If txtLK04 <> strText Then
      Me.Tag = "欲增加收受者( " & Me.Tag & " )，目前欄位長度(" & Len(txtLK04) & ")不足，必須至少為(" & Len(strText) & ")，請通知電腦中心！"
      MsgBox Me.Tag, vbExclamation
   Else
      Me.Tag = ""
   End If
   '2025/9/22 END
   Combo2.SetFocus
End Sub

'減項
Private Sub cmdRemove_Click()
   RemovelstUsers
   txtLK04 = ComposeListX
   Combo2.SetFocus
End Sub

Private Sub Combo2_GotFocus()
   InverseTextBox Combo2
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_LostFocus()
Dim strText As String
   
   If m_EditMode <> 0 And Combo2.Text > "" And Len(Trim(Combo2.Text)) = 5 Then
      '抓取員工姓名
      Combo2.Text = SetCboStaffName(Combo2.Text)
   'Add By Sindy 2021/6/22 檢查是否輸入員工姓名
   Else
      '依員工姓名抓取員工編號
      strText = GetPrjSalesNM_2(CStr(Combo2.Text))
      If strText <> "" Then
         Combo2.Text = Left(Left(strText, 5) & Space(5), 7) & Combo2.Text
      End If
      '2021/6/22 END
   End If
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If Combo2.Text > "" And Len(Trim(Combo2.Text)) = 5 Then
         '檢查人員是否存在或離職
         If ChkStaffST04(Left(Combo2, 5)) = True Then
            Call Combo2_GotFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Function ComposeListX() As String
   strExc(1) = ""
   If lstUsers.ListCount > 0 Then
      For intI = 0 To lstUsers.ListCount - 1
         If Len(lstUsers.List(intI)) > 6 And InStr(lstUsers.List(intI), "@") = 0 Then
            If Mid(lstUsers.List(intI), 6, 1) = " " Then
               If GetPrjSalesNM(Mid(lstUsers.List(intI), 1, 5)) <> "" Then
                  strExc(1) = strExc(1) & ";" & Mid(lstUsers.List(intI), 1, 5) '員工編號
                  GoTo RunNext
               End If
            End If
         End If
         strExc(1) = strExc(1) & ";" & lstUsers.List(intI) '員工編號
RunNext:
      Next
      If strExc(1) <> "" Then strExc(1) = Mid(strExc(1), 2)
   End If
   ComposeListX = strExc(1)
End Function

Private Sub SetlstUsers(p_stNums As String)
   Dim arrID, strTempName As String
   lstUsers.Clear
   If p_stNums <> "" Then
      arrID = Split(p_stNums, ";")
      '照原順序排
      For ii = 0 To UBound(arrID)
         strTempName = ""
         If Trim(arrID(ii)) <> "" Then
            If Len(arrID(ii)) = 5 And InStr(arrID(ii), "@") = 0 Then '員工編號
               strTempName = GetPrjSalesNM(CStr(arrID(ii)))
            End If
            If strTempName <> "" Then
               lstUsers.AddItem arrID(ii) & " " & strTempName, 0
               lstUsers.ItemData(0) = PUB_Id2Num(CStr(arrID(ii))) '員工編號
            Else
               lstUsers.AddItem arrID(ii), 0
            End If
         End If
      Next ii
   End If
End Sub

Private Sub AddlstUsers()
   Dim idx As Integer, bFound As Boolean
   Dim strUserId As String
   
   Combo2.Text = Trim(Combo2.Text) 'Add By Sindy 2017/3/14
   If Combo2.Text <> "" Then
      If Len(Combo2.Text) > 6 And InStr(Combo2.Text, "@") = 0 Then '員工編號
         If Mid(Combo2.Text, 6, 1) = " " Then
            If GetPrjSalesNM(Mid(Combo2.Text, 1, 5)) <> "" Then
               strUserId = Mid(Combo2.Text, 1, 5)
            End If
         End If
      'Add By Sindy 2025/9/22 ex:潘子微(外專.主任.Anny) <A4011@taie.com.tw>
      ElseIf InStr(Combo2.Text, "@") > 0 _
         And (InStr(Combo2.Text, "<") > 0 Or InStr(Combo2.Text, ">") > 0 _
              Or InStr(Combo2.Text, "(") > 0 Or InStr(Combo2.Text, ")") > 0) Then
         MsgBox "請輸入員工編號 或 員工姓名 或 E-Mail！"
         Combo2.SetFocus
         Combo2_GotFocus
         Exit Sub
      End If
      If strUserId <> "" Then
         For idx = 0 To lstUsers.ListCount - 1
            If lstUsers.ItemData(idx) = PUB_Id2Num(strUserId) Then
               MsgBox "員工已存在於收受者清單中！"
               Combo2.SetFocus
               Combo2_GotFocus
               bFound = True
               Exit For
            End If
         Next
      Else
         For idx = 0 To lstUsers.ListCount - 1
            If lstUsers.List(idx) = Combo2.Text Then
               MsgBox "員工已存在於收受者清單中！"
               Combo2.SetFocus
               Combo2_GotFocus
               bFound = True
               Exit For
            End If
         Next
      End If
      If bFound = False Then
         lstUsers.AddItem Combo2.Text, 0
         If strUserId <> "" Then
            lstUsers.ItemData(0) = PUB_Id2Num(strUserId)
         End If
         Combo2.Text = ""
      End If
   End If
End Sub

Private Sub RemovelstUsers()
   Dim idx As Integer, ii As Integer
   If lstUsers.ListCount > 0 Then
      ii = 0
      For idx = 0 To lstUsers.ListCount - 1
         If lstUsers.Selected(ii) = True Then
            lstUsers.RemoveItem ii
            ii = ii - 1
         End If
         ii = ii + 1
      Next
   End If
End Sub
