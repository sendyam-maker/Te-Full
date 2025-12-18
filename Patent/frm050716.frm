VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050716 
   BorderStyle     =   1  '單線固定
   Caption         =   "系統特殊設定"
   ClientHeight    =   5700
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   7500
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7500
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C0FFC0&
      Caption         =   "關鍵字查詢"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4860
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   660
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5970
      MaxLength       =   30
      TabIndex        =   14
      Top             =   690
      Width           =   1395
   End
   Begin VB.TextBox TxtUpdMsg 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1620
      TabIndex        =   11
      Text            =   "資料更新中，請稍後..."
      Top             =   2130
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox textAuthGrp 
      Height          =   270
      Left            =   1080
      MaxLength       =   14
      TabIndex        =   3
      Top             =   2565
      Width           =   6285
   End
   Begin VB.TextBox textCode 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   0
      Top             =   690
      Width           =   3630
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   2730
      Left            =   30
      TabIndex        =   5
      Top             =   2865
      Width           =   7425
      _ExtentX        =   13102
      _ExtentY        =   4784
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   1200
      _ExtentX        =   974
      _ExtentY        =   974
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050716.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050716.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050716.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050716.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050716.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050716.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050716.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050716.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050716.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050716.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050716.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7500
      _ExtentX        =   13229
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
   Begin MSForms.TextBox textMan 
      Height          =   465
      Left            =   690
      TabIndex        =   2
      Top             =   1560
      Width           =   4500
      VariousPropertyBits=   -1463795685
      ScrollBars      =   2
      Size            =   "7937;820"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textExplain 
      Height          =   495
      Left            =   750
      TabIndex        =   1
      Top             =   1020
      Width           =   6615
      VariousPropertyBits=   -1463795685
      MaxLength       =   300
      ScrollBars      =   2
      Size            =   "11668;873"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   420
      Left            =   1100
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2100
      Width           =   6270
      VariousPropertyBits=   -1467989985
      Size            =   "11060;741"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "顯示姓名："
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   2100
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "可維護群組："
      Height          =   180
      Left            =   0
      TabIndex        =   10
      Top             =   2595
      Width           =   1080
   End
   Begin VB.Label Label4 
      Caption         =   "人員請輸員工編號，並以 , 或 ; 分隔"
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   5460
      TabIndex        =   9
      Top             =   1620
      Width           =   1704
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "內容："
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   1590
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "說明："
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   1020
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "設定代號："
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   930
   End
End
Attribute VB_Name = "frm050716"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/13 Form2.0已修改(textExplain,textMan,Text1,Grd1)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
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
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_EditMode As Integer
Dim m_SubMode As Integer
' 第一筆資料的key
Dim m_FirstKEY As String
' 最後一筆資料的key
Dim m_LastKEY As String
' 目前正在顯示的key
Dim m_CurrKEY As String
Dim m_iPreRow As Integer '前次顯示資料列
Dim m_bGridChange As Boolean
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序 Add By Sindy 2019/12/27
Dim stCon As String              'add by sonia 2021/12/23 改為公用

'add by sonia 2022/3/9
Private Sub cmdSearch_Click()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   If Trim(Text2) <> "" Then
      'Modify By Sindy 2022/4/28 + or instr(oExplain,'" & Text2 & "')>0 or instr(oMan,'" & Text2 & "')>0
      'Modified by Lydia 2022/07/15
      'modify by sonia 2024/4/30 +upper
      strSql = "SELECT * FROM SetSpecMan WHERE instr(upper(oCode),upper('" & UCase(ChgSQL(Text2)) & "'))>0 or instr(upper(oExplain),upper('" & UCase(ChgSQL(Text2)) & "'))>0 or instr(upper(oMan),upper('" & UCase(ChgSQL(Text2)) & "'))>0 order by oCode"
   Else
      strSql = "SELECT oCode,oExplain,oMan,oAuthGrp FROM SetSpecMan where 1=1 " & stCon & " order by oCode "
   End If
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set GRD1.Recordset = rsTmp
   rsTmp.Close
   SetGrd
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub
'end 2022/3/9

Private Sub Form_Initialize()
   ReDim m_FieldList(4) As FIELDITEM
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
'            PrintData
      ' 刪除
      Case vbKeyF5:
'         If m_bDelete Then
'            If m_EditMode = 0 Then
'               OnAction KeyCode
'               KeyCode = 0
'            End If
'         End If
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
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

'modify by sonia 2021/12/13
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

   'add by sonia 2021/12/23
   If Pub_StrUserSt03 <> "M51" Then
      stCon = " and instr(OAUTHGRP,'" & Pub_strUserST05 & "')>0"
      'Added by Lydia 2024/01/12 非電腦中心人員隱藏按鈕; (內商程序人員增加維護自己的設定權限)
      cmdSearch.Visible = False
      Text2.Visible = False
      'end 2024/01/12
   End If
   'end 2021/12/23
   
   MoveFormToCenter Me
   m_bInsert = IsUserHasRightOfFunction("frm050716", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm050716", strEdit, False)
   m_bQuery = IsUserHasRightOfFunction("frm050716", strFind, False)
   InitialField
   RefreshRange
   GetAllData
   ShowLastRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   SetGrd
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm050716 = Nothing
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
    
    'grd1.Visible = False
   arrGridHeadText = Array("設定代號", "說明", "內容", "可維護群組")
   arrGridHeadWidth = Array(2000, 3000, 2500, 2500)
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

'Add By Sindy 2019/12/27
Private Sub Grd1_Click()
Dim nCol As Long, nRow As Long

   GRD1.Visible = False
   GRD1.row = GRD1.MouseRow
   GRD1.col = GRD1.MouseCol
   nRow = GRD1.row
   nCol = GRD1.col
   If nRow = 0 Then
      If GRD1.Text <> "V" Then
         If GRD1.Text = "無" Then
            If m_blnColOrderAsc = True Then
               GRD1.Sort = 3  '數值昇冪
               m_blnColOrderAsc = False
            Else
               GRD1.Sort = 4 '數值降冪
               m_blnColOrderAsc = True
            End If
         Else
            If m_blnColOrderAsc = True Then
               GRD1.Sort = 5 '字串昇冪
               m_blnColOrderAsc = False
            Else
               GRD1.Sort = 6 '字串降冪
               m_blnColOrderAsc = True
            End If
         End If
      End If
   End If
   GRD1.Visible = True
End Sub

Private Sub grd1_SelChange()
Dim TmpRow As Integer
   
   'grd1.Visible = False
   TmpRow = GRD1.MouseRow
   GRD1.col = 0
   If TmpRow <> 0 Then
       m_CurrKEY = GRD1.TextMatrix(TmpRow, 0)
       UpdateCtrlData
   End If
   GRD1.Visible = True
End Sub

Private Sub ChgGrdData(iRow As Integer)
Dim i, j, k
   
   'Modify by Morgan 2009/2/19
   'grd1.Visible = False
   'For j = 1 To Grd1.Rows - 1
   '     Grd1.row = j
   '     For k = 0 To Grd1.Cols - 1
   '         Grd1.col = k
   '         Grd1.CellBackColor = QBColor(15)
   '     Next k
   ' Next j
   
   If m_iPreRow > 0 And m_iPreRow < GRD1.Rows Then
      GRD1.row = m_iPreRow
      For k = 0 To GRD1.Cols - 1
          GRD1.col = k
          GRD1.CellBackColor = QBColor(15)
      Next k
   End If
   'end 2009/2/19

   GRD1.row = iRow
   For j = 0 To GRD1.Cols - 1
       GRD1.col = j
       GRD1.CellBackColor = &HFFC0C0
       m_iPreRow = GRD1.row 'Add by Morgan 2009/2/19
   Next j
   'grd1.TopRow = iRow Remove by Morgan 2009/2/19
   GRD1.Visible = True
End Sub

Private Sub ChgToNowData()
Dim i, j As Integer
   j = 0
   For i = 1 To GRD1.Rows - 1
       If GRD1.TextMatrix(i, 0) = textCode Then
           j = i
           Exit For
       End If
   Next i
   If j <> 0 Then ChgGrdData j
End Sub

Private Sub textAuthGrp_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCode_GotFocus()
   InverseTextBox textCode
End Sub

Private Sub textcode_LostFocus()
   If Trim(textCode) <> "" And textCode.Locked = False Then
       'm_CurrKEY = ""   'CANCEL BY SONIA 2021/12/23 否則跳離此欄後再按前後筆無效
       GetAllData
   End If
End Sub

Private Sub textExplain_GotFocus()
   InverseTextBox textExplain
   OpenIme
End Sub

Private Sub textexplain_Validate(Cancel As Boolean)
   If CheckLengthIsOK(textExplain, textExplain.MaxLength) = False Then
       Cancel = True
       Exit Sub
   End If
   CloseIme
End Sub

Private Sub textman_GotFocus()
   InverseTextBox textMan
   CloseIme
End Sub

'modify by sonia 2021/12/13
'Private Sub textMan_KeyPress(KeyAscii As Integer)
Private Sub textMan_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textman_Validate(Cancel As Boolean)
   'Modified by Lydia 2025/07/28 將OMAN的TextBox.maxlength改成0，避免放大欄位長度未能修改TextBox
   'If CheckLengthIsOK(textMan, textMan.MaxLength) = False Then
   If CheckLengthIsOK(textMan, 300) = False Then
       Cancel = True
   End If
End Sub

Private Sub textauthgrp_GotFocus()
   InverseTextBox textAuthGrp
   CloseIme
End Sub

Private Sub textAuthGrp_Validate(Cancel As Boolean)
   If CheckLengthIsOK(textAuthGrp, textAuthGrp.MaxLength) = False Then
       Cancel = True
   End If
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
      ' 刪除
      Case 3: 'OnAction vbKeyF5
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

' 初始化欄位陣列
Private Sub InitialField()
   ' 初始化欄位陣列
    m_FieldList(0).fiName = "oCode"
    m_FieldList(0).fiOldData = Empty
    m_FieldList(0).fiNewData = Empty
    m_FieldList(0).fiType = 0 '文字型態
    m_FieldList(1).fiName = "oExplain"
    m_FieldList(1).fiOldData = Empty
    m_FieldList(1).fiNewData = Empty
    m_FieldList(1).fiType = 0  '文字型態
    m_FieldList(2).fiName = "oMan"
    m_FieldList(2).fiOldData = Empty
    m_FieldList(2).fiNewData = Empty
    m_FieldList(2).fiType = 0  '文字型態
    m_FieldList(3).fiName = "oAuthGrp"
    m_FieldList(3).fiOldData = Empty
    m_FieldList(3).fiNewData = Empty
    m_FieldList(3).fiType = 0  '文字型態
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
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         textCode = ""
         textExplain = ""
         textAuthGrp = ""
         textCode.Locked = False
         textExplain.Locked = False
         textAuthGrp.Locked = False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         UpdateCtrlData
         If Pub_StrUserSt03 = "M51" Or InStr(1, textAuthGrp, PUB_GetST05(strUserNum)) <> 0 Then
                m_EditMode = 2
                SetCtrlReadOnly False
                textCode.Locked = True
                If Pub_StrUserSt03 = "M51" Then
                   textExplain.Locked = False
                   textAuthGrp.Locked = False
                Else
                   textExplain.Locked = True
                   textAuthGrp.Locked = True
                End If
                UpdateToolbarState
                SetInputEntry
                If textExplain.Locked = True Then textMan.SetFocus   'add by sonia 2021/12/15
        Else
            MsgBox "無此使用權限...", , "警告!!"
        End If
      ' 刪除
      Case vbKeyF5:
'         strTit = "詢問"
'         strMsg = "是否要刪除此筆資料?"
'         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
'         If nResponse = vbYes Then
'            m_EditMode = 3
'            If OnWork = True Then
'                UpdateToolbarState
'            Else
'                Exit Sub
'            End If
'         End If
      ' 查詢
      Case vbKeyF4:
'         SetCtrlReadOnly True
'         ClearField
'         UpdateToolbarState
'         SetInputEntry
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
         CloseIme
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select

End Sub

'Modifie by Morgan 2015/3/19 +限制讀取的資料
Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
'Dim stCon As String   'cancel by sonia 移到公用區
   
   'cancel by sonia 移到Form_Load
   'If Pub_StrUserSt03 <> "M51" Then
   '   stCon = " and instr(OAUTHGRP,'" & Pub_strUserST05 & "')>0"
   'End If
   
   strSql = "SELECT min(ocode) as oCode FROM SetSpecMan where 1=1 " & stCon

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("oCode")) = False Then: m_FirstKEY = rsTmp.Fields("oCode")
   End If
   rsTmp.Close

   strSql = "SELECT max(ocode) as oCode FROM SetSpecMan where 1=1 " & stCon
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("oCode")) = False Then: m_LastKEY = rsTmp.Fields("oCode")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY = m_FirstKEY

   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If m_CurrKEY = m_FirstKEY Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If

   'modify by sonia 2021/12/23 內層加stCon條件
   'strSql = "SELECT oCode FROM SetSpecMan " & _
            "WHERE oCode in (select max(oCode) from SetSpecMan where  oCode<'" & m_CurrKEY & "') " & stCon
   strSql = "SELECT oCode FROM SetSpecMan " & _
            "WHERE oCode in (select max(oCode) from SetSpecMan where  oCode<'" & m_CurrKEY & "'" & stCon & ")"
   'end 2021/12/23
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("oCode")) = False Then: m_CurrKEY = rsTmp.Fields("oCode")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
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

   If m_CurrKEY = m_LastKEY Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If

   'modify by sonia 2021/12/23 內層加stCon條件
   'strSql = "SELECT oCode FROM SetSpecMan " & _
            "WHERE oCode in (select min(oCode) from SetSpecMan where  oCode>'" & m_CurrKEY & "') "
   strSql = "SELECT oCode FROM SetSpecMan " & _
            "WHERE oCode in (select min(oCode) from SetSpecMan where  oCode>'" & m_CurrKEY & "'" & stCon & ")"
   'end 2021/12/23
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("oCode")) = False Then: m_CurrKEY = rsTmp.Fields("oCode")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close

   UpdateCtrlData

EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY = m_LastKEY

   UpdateCtrlData
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
         If m_bDelete Then
            Toolbar1.Buttons(3).Enabled = True
         Else
            Toolbar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            Toolbar1.Buttons(4).Enabled = True
         Else
            Toolbar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            Toolbar1.Buttons(6).Enabled = True
            Toolbar1.Buttons(7).Enabled = True
            Toolbar1.Buttons(8).Enabled = True
            Toolbar1.Buttons(9).Enabled = True
         Else
            Toolbar1.Buttons(6).Enabled = False
            Toolbar1.Buttons(7).Enabled = False
            Toolbar1.Buttons(8).Enabled = False
            Toolbar1.Buttons(9).Enabled = False
         End If
         Toolbar1.Buttons(11).Enabled = False
         Toolbar1.Buttons(12).Enabled = False
         Toolbar1.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         Toolbar1.Buttons(1).Enabled = False
         Toolbar1.Buttons(2).Enabled = False
         Toolbar1.Buttons(3).Enabled = False
         Toolbar1.Buttons(4).Enabled = False
         Toolbar1.Buttons(6).Enabled = False
         Toolbar1.Buttons(7).Enabled = False
         Toolbar1.Buttons(8).Enabled = False
         Toolbar1.Buttons(9).Enabled = False
         Toolbar1.Buttons(11).Enabled = True
         Toolbar1.Buttons(12).Enabled = True
         Toolbar1.Buttons(14).Enabled = False
   End Select
   
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textMan.Locked = bEnable
   textExplain.Locked = bEnable
   textAuthGrp.Locked = bEnable
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
                ChgToNowData
            Else
                Exit Function
            End If
      Case 2: '修改
            If TxtValidate = False Then Exit Function
            ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
      Case 3: '刪除
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
'         UpdateFieldNewData
'         If DelRecord = True Then
'            RefreshRange
'         Else
'            Exit Function
'         End If
      Case 4: '列印
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
'         UpdateFieldNewData
'         'If CheckDataValid() = True Then
'         If textCU01 <> "" Then
'            If QueryRecord = False Then
'               strMsg = "無此資料"
'               strTit = "查詢資料"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               UpdateCtrlData
'            End If
'         Else
'            GoTo EXITSUB
'         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

Private Sub ClearField()
   Dim nIndex As Integer
   textCode = Empty
   textExplain = Empty
   textMan = Empty
   textAuthGrp = Empty
   For nIndex = 0 To 3
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   strSql = "SELECT * FROM SetSpecMan " & _
            "WHERE oCode = '" & m_CurrKEY & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("oCode")) = False Then: textCode = rsTmp.Fields("oCode"): m_CurrKEY = textCode
      If IsNull(rsTmp.Fields("oExplain")) = False Then: textExplain = rsTmp.Fields("oExplain")
      textMan.Tag = "" 'Add By Sindy 2014/9/11
      Text1 = "" 'Add By Sindy 2019/12/27
      If IsNull(rsTmp.Fields("oMan")) = False Then
         textMan = rsTmp.Fields("oMan")
         textMan.Tag = rsTmp.Fields("oMan") 'Modify By Sindy 2014/9/11
         'Add By Sindy 2019/12/27
         Text1.Tag = PUB_ReadUserData(textMan)
         If Text1.Tag <> textMan And Replace(Text1.Tag, ",", ";") <> textMan Then
            Text1 = Text1.Tag
         End If
         '2019/12/27 END
      End If
      If IsNull(rsTmp.Fields("oAuthGrp")) = False Then: textAuthGrp = rsTmp.Fields("oAuthGrp")
      'cancel by sonia 2021/12/15 移到OnAction的vbKeyF3
      'If textMan.Visible = True Then textMan.SetFocus 'Add By Sindy 2019/12/27
      'end 2021/12/15
      ChgToNowData
   End If
   ' 更新暫存區的資料
   UpdateFieldOldData rsTmp
   rsTmp.Close
EXITSUB:
   Set rsTmp = Nothing
End Sub

'抓當日所有資料
'Modifie by Morgan 2015/3/19 +限制讀取的資料
Private Sub GetAllData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
'Dim stCon As String    'cancel by sonia 移到公用區
   
   'cancel by sonia 移到Form_Load
   'If Pub_StrUserSt03 <> "M51" Then
   '   stCon = " and instr(OAUTHGRP,'" & Pub_strUserST05 & "')>0"
   'End If
   
    strSql = "SELECT oCode,oExplain,oMan,oAuthGrp FROM SetSpecMan where 1=1 " & stCon & " order by oCode "
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    Set GRD1.Recordset = rsTmp
    rsTmp.Close
    SetGrd
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textCode.SetFocus: textCode_GotFocus
      Case 2: textExplain.SetFocus: textExplain_GotFocus
   End Select
End Sub

Private Sub UpdateFieldNewData()
   '若新增資料
   SetFieldNewData "oCode", textCode
   SetFieldNewData "oExplain", textExplain
   SetFieldNewData "oMan", textMan
   SetFieldNewData "oAuthGrp", textAuthGrp
   'add by sonia 2022/3/7
   If Left(textCode, 4) = "MCTF" Then
      MsgBox "修改MCTF人員時，請同時修改該編號員工檔的內部郵件收件員工編號！", , "提醒!!"
      'add by sonia 2024/4/24
      If InStr(textCode, "收信人員") = 0 Then
         MsgBox "修改MCTF人員時，請考慮該編號是否應加入特殊設定「MCTMember」內，會影響收文記錄CP161欄！", , "提醒!!"
      End If
      'end 2024/4/24
   End If
   'end 2022/3/7
End Sub

Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To 3
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
   
   AddRecord = False
   
   bFirst = True
   strSql = "INSERT INTO SetSpecMan (oCode,oExplain,oMan,oAuthGrp) values ("
   bFirst = True
   For nIndex = 0 To 3
            strTmp = Empty
            If m_FieldList(nIndex).fiType = 0 Then
               strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
            Else
               strTmp = m_FieldList(nIndex).fiNewData
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
   strSql = strSql & ") "
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   RefreshRange
   GetAllData
   ShowCurrRecord textCode
   AddRecord = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM SetSpecMan " & _
            "WHERE oCode = '" & strKEY01 & "' "
                  
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

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To 3
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

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY = strKEY01
   Else
      strSql = "SELECT oCode FROM SetSpecMan " & _
               "WHERE oCode = '" & m_CurrKEY & "'  "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("oCode")) = False Then: m_CurrKEY = rsTmp.Fields("oCode")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
   strSql = "SELECT oCode FROM SetSpecMan " & _
            "WHERE oCode = '" & textCode & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("oCode")) = False Then: m_CurrKEY = rsTmp.Fields("oCode")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   If Trim(textCode.Text) = "" And textCode.Locked = False And textCode.Enabled = True Then
      MsgBox "識別代號不可空白！", vbInformation, "操作錯誤！"
      textCode.SetFocus
      Exit Function
   End If

'   'Add By Sindy 2016/8/24 會更新CFP程序人員相關資料,因此要做一些資料檢查
'   If UCase(m_CurrKEY) = UCase("專利處轉信亞洲程序") Or UCase(m_CurrKEY) = UCase("專利處轉信歐洲程序") _
'      Or UCase(m_CurrKEY) = UCase("專利處轉信美洋非洲單號程序") Or UCase(m_CurrKEY) = UCase("專利處轉信美洋非洲雙號程序") Then
'      If Len(Trim(textMan)) <> 5 Then
'         MsgBox "請輸入員工編號！", vbInformation, "操作錯誤！"
'         textMan.SetFocus
'         Exit Function
'      End If
'      '檢查人員是否存在或離職
'      If ChkStaffST04(textMan) = True Then
'         textMan.SetFocus
'         Exit Function
'      End If
'      If PUB_GetST05(textMan) <> "83" And PUB_GetST05(textMan) <> "85" Then
'         MsgBox "請輸入CFP程序人員！", vbInformation, "操作錯誤！"
'         textMan.SetFocus
'         Exit Function
'      End If
'   End If
'   '2016/8/24 END

   'Add by Morgan 2009/2/19
   If CheckLengthIsOK(textCode, textCode.MaxLength) = False Then
      textCode.SetFocus
      Exit Function
   End If
   If CheckLengthIsOK(textExplain, textExplain.MaxLength) = False Then
      textExplain.SetFocus
      Exit Function
   End If
   If CheckLengthIsOK(textMan, textMan.MaxLength) = False Then
      textMan.SetFocus
      Exit Function
   End If
   If CheckLengthIsOK(textAuthGrp, textAuthGrp.MaxLength) = False Then
      textAuthGrp.SetFocus
      Exit Function
   End If
   'end 2009/2/19
   
   'add by sonia 2021/12/24 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/24
   
   TxtValidate = True
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
Dim strCode As String
'Add By Sindy 2016/8/25
Dim rsTmp As New ADODB.Recordset
Dim strMinCP05 As String
'2016/8/25 END
'Dim strSQLCon As String, strCFPEmp As String 'Add By Sindy 2018/6/21
   
   ModRecord = False
   
   strCode = m_CurrKEY
   strSql = "UPDATE SetSpecMan SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To 3
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

   strSql = strSql & " WHERE oCode = '" & strCode & "' "
   
On Error GoTo ErrHnd
   
   If bDifference = True Then
      Screen.MousePointer = vbHourglass 'Add By Sindy 2016/8/24
      cnnConnection.BeginTrans
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      
      'Add By Sindy 2014/9/11
      'CFT承辦人有異動時,更新CFT該國案件之下一程序未處理催審期限之智權人員
      'Modified by Lydia 2016/11/16 + Or UCase(strCode) = UCase("CFT_101239A") Or UCase(strCode) = UCase("CFT_101239B")
      'Modified by Lydia 2018/10/03 分成CFT_101239B(北所)和CFT_101239C(中所)兩個設定
      'If UCase(strCode) = UCase("CFT_011A") Or UCase(strCode) = UCase("CFT_011B") Or UCase(strCode) = UCase("CFT_101239A") Or UCase(strCode) = UCase("CFT_101239B") Then
      'Mark by Lydia 2022/08/05 改寫成Trigger: NATION_BEFORE 和 SETSPECMAN_BEFORE
'      If UCase(strCode) = UCase("CFT_011A") Or UCase(strCode) = UCase("CFT_011B") Or UCase(strCode) = UCase("CFT_101239A") Or UCase(strCode) = UCase("CFT_101239B") Or UCase(strCode) = UCase("CFT_101239C") Then
'         If Len(Trim(textMan)) = 5 And Trim(textMan.Text) <> Trim(textMan.Tag) Then
'            TxtUpdMsg.Visible = True 'Add By Sindy 2016/8/24
'            'Modifeid by Lydia 2016/11/16 加判斷
'            'If PUB_UpdNpCFT305Np10(Trim(textMan.Text), "011", IIf(UCase(strCode) = UCase("CFT_011A"), 2, 1)) = False Then GoTo ErrHnd
'            If Mid(UCase(strCode), 1, 7) = "CFT_011" Then
'                If PUB_UpdNpCFT305Np10(Trim(textMan.Text), "011", IIf(UCase(strCode) = UCase("CFT_011A"), 2, 1)) = False Then GoTo ErrHnd
'            ElseIf Mid(UCase(strCode), 1, 10) = "CFT_101239" Then
'                'Modified by Lydia 2018/10/03 分成CFT_101239B(北所)和CFT_101239C(中所)兩個設定
'                'If PUB_UpdNpCFT305Np10(Trim(textMan.Text), "101", IIf(UCase(strCode) = UCase("CFT_101239A"), 4, 5)) = False Then GoTo ErrHnd
'                'If PUB_UpdNpCFT305Np10(Trim(textMan.Text), "239", IIf(UCase(strCode) = UCase("CFT_101239A"), 4, 5)) = False Then GoTo ErrHnd
'                strExc(1) = ""
'                Select Case Right(strCode, 1)
'                      '南高所
'                      Case "A": strExc(1) = "4"
'                      '北所
'                      Case "B": strExc(1) = "5"
'                      '中所
'                      Case "C": strExc(1) = "6"
'                End Select
'                If strExc(1) <> "" Then
'                      If PUB_UpdNpCFT305Np10(Trim(textMan.Text), "101", Val(strExc(1))) = False Then GoTo ErrHnd
'                      If PUB_UpdNpCFT305Np10(Trim(textMan.Text), "239", Val(strExc(1))) = False Then GoTo ErrHnd
'                End If
'                'end 2018/10/03
'            End If
'            TxtUpdMsg.Visible = False 'Add By Sindy 2016/8/24
'         End If
'      End If
'      '2014/9/11 END
      'end 2022/08/05
      
'      'Add By Sindy 2016/8/24 更新CFP程序人員相關資料
'      textMan.Text = Trim(textMan.Text)
'      textMan.Tag = Trim(textMan.Tag)
'      If Len(textMan.Text) = 5 _
'         And textMan.Text <> textMan.Tag _
'         And (UCase(strCode) = UCase("專利處轉信亞洲程序") Or UCase(strCode) = UCase("專利處轉信歐洲程序") _
'              Or UCase(strCode) = UCase("專利處轉信美洋非洲單號程序") Or UCase(strCode) = UCase("專利處轉信美洋非洲雙號程序")) Then
'         TxtUpdMsg.Visible = True
'         '先找出最小收文日,速度才會快一點
'         strSql = "SELECT MIN(CP05) FROM CASEPROGRESS,PATENT,staff" & _
'                  " WHERE CP01 in ('CFP','CPS') AND CP27 IS NULL AND CP57 IS NULL AND CP05>=20030201 and cp14=st01(+) and st03='P12'" & _
'                  " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL"
'         rsTmp.CursorLocation = adUseClient
'         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         strMinCP05 = ""
'         If rsTmp.RecordCount > 0 Then
'            strMinCP05 = rsTmp.Fields(0)
'         End If
'         rsTmp.Close
'
'         If UCase(strCode) = UCase("專利處轉信亞洲程序") Then
'            strSql = "UPDATE NATION SET NA73='" & textMan.Text & "',NA74='" & textMan.Text & "' WHERE NA02<'C1'"
'            cnnConnection.Execute strSql
'            '更新亞洲
'            strSql = "UPDATE CASEPROGRESS SET CP14='" & textMan.Text & "'" & _
'                     " WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,PATENT,NATION" & _
'                     " WHERE CP14='" & textMan.Tag & "' AND CP04='00' AND CP27 IS NULL AND CP57 IS NULL" & IIf(strMinCP05 <> "", " AND CP05>=" & strMinCP05, "") & _
'                     " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL AND PA09=NA01(+) AND SUBSTR(NA02,1,2) IN ('A0','B0','C0'))"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'            '更新待送件區
'            strSql = "UPDATE EmpElectronProcess" & _
'                     " set eep05='" & textMan.Text & "'" & _
'                     " where (eep01,eep02,eep04) in" & _
'                     " (select eep01,eep02,eep04 from caseprogress,EmpElectronProcess,patent,nation WHERE cp01 in ('CFP','CPS') and nvl(cp27,0)=0 and nvl(cp57,0)=0" & _
'                     " and cp09=eep01(+) and eep05='" & textMan.Tag & "'" & _
'                     " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+) AND SUBSTR(NA02,1,2) IN ('A0','B0','C0')" & _
'                     " )"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'         ElseIf UCase(strCode) = UCase("專利處轉信歐洲程序") Then
'            strSql = "UPDATE NATION SET NA73='" & textMan.Text & "',NA74='" & textMan.Text & "' WHERE NA02='C20'"
'            cnnConnection.Execute strSql
'            '更新歐洲
'            strSql = "UPDATE CASEPROGRESS SET CP14='" & textMan.Text & "'" & _
'                     " WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,PATENT,NATION" & _
'                     " WHERE CP14='" & textMan.Tag & "' AND CP04='00' AND CP27 IS NULL AND CP57 IS NULL" & IIf(strMinCP05 <> "", " AND CP05>=" & strMinCP05, "") & _
'                     " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL AND PA09=NA01(+) AND SUBSTR(NA02,1,2)='C2')"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'            '子案各國也要更新(會有子案的只有EPC所以才沒有下國家別)
'            strSql = "UPDATE CASEPROGRESS SET CP14='" & textMan.Text & "'" & _
'                     " WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,PATENT" & _
'                     " WHERE CP14='" & textMan.Tag & "' AND CP04<>'00' AND CP27 IS NULL AND CP57 IS NULL" & IIf(strMinCP05 <> "", " AND CP05>=" & strMinCP05, "") & _
'                     " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL)"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'            '更新待送件區
'            strSql = "UPDATE EmpElectronProcess" & _
'                     " set eep05='" & textMan.Text & "'" & _
'                     " where (eep01,eep02,eep04) in" & _
'                     " (select eep01,eep02,eep04 from caseprogress,EmpElectronProcess,patent,nation WHERE cp01 in ('CFP','CPS') and nvl(cp27,0)=0 and nvl(cp57,0)=0" & _
'                     " and cp09=eep01(+) and eep05='" & textMan.Tag & "'" & _
'                     " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+) AND SUBSTR(NA02,1,2)='C2'" & _
'                     " )"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'         ElseIf UCase(strCode) = UCase("專利處轉信美洋非洲單號程序") Then
'            '美洲
'            strSql = "UPDATE NATION SET NA73='" & textMan.Text & "' WHERE NA02='C10'"
'            cnnConnection.Execute strSql
'            '非洲
'            strSql = "UPDATE NATION SET NA73='" & textMan.Text & "' WHERE NA02='C30'"
'            cnnConnection.Execute strSql
'            '大洋洲
'            strSql = "UPDATE NATION SET NA73='" & textMan.Text & "' WHERE NA02='C40'"
'            cnnConnection.Execute strSql
'            '更新美洲,大洋洲,非洲(單號)
'            strSql = "UPDATE CASEPROGRESS SET CP14='" & textMan.Text & "'" & _
'                     " WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,PATENT,NATION" & _
'                     " WHERE CP14='" & textMan.Tag & "' AND CP04='00' AND CP27 IS NULL AND CP57 IS NULL" & IIf(strMinCP05 <> "", " AND CP05>=" & strMinCP05, "") & _
'                     " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL AND PA09=NA01(+) AND SUBSTR(NA02,1,2) IN ('C1','C3','C4')" & _
'                     " AND mod(CP02,2)=1)"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'            '更新待送件區
'            strSql = "UPDATE EmpElectronProcess" & _
'                     " set eep05='" & textMan.Text & "'" & _
'                     " where (eep01,eep02,eep04) in" & _
'                     " (select eep01,eep02,eep04 from caseprogress,EmpElectronProcess,patent,nation WHERE cp01 in ('CFP','CPS') and nvl(cp27,0)=0 and nvl(cp57,0)=0" & _
'                     " and cp09=eep01(+) and eep05='" & textMan.Tag & "'" & _
'                     " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+) AND SUBSTR(NA02,1,2) IN ('C1','C3','C4')" & _
'                     " and mod(CP02,2)=1)"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'         ElseIf UCase(strCode) = UCase("專利處轉信美洋非洲雙號程序") Then
'            '美洲
'            strSql = "UPDATE NATION SET NA74='" & textMan.Text & "' WHERE NA02='C10'"
'            cnnConnection.Execute strSql
'            '非洲
'            strSql = "UPDATE NATION SET NA74='" & textMan.Text & "' WHERE NA02='C30'"
'            cnnConnection.Execute strSql
'            '大洋洲
'            strSql = "UPDATE NATION SET NA74='" & textMan.Text & "' WHERE NA02='C40'"
'            cnnConnection.Execute strSql
'            '更新美洲,大洋洲,非洲(雙號)
'            strSql = "UPDATE CASEPROGRESS SET CP14='" & textMan.Text & "'" & _
'                     " WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,PATENT,NATION" & _
'                     " WHERE CP14='" & textMan.Tag & "' AND CP04='00' AND CP27 IS NULL AND CP57 IS NULL" & IIf(strMinCP05 <> "", " AND CP05>=" & strMinCP05, "") & _
'                     " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL AND PA09=NA01(+) AND SUBSTR(NA02,1,2) IN ('C1','C3','C4')" & _
'                     " AND mod(CP02,2)=0)"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'            '更新待送件區
'            strSql = "UPDATE EmpElectronProcess" & _
'                     " set eep05='" & textMan.Text & "'" & _
'                     " where (eep01,eep02,eep04) in" & _
'                     " (select eep01,eep02,eep04 from caseprogress,EmpElectronProcess,patent,nation WHERE cp01 in ('CFP','CPS') and nvl(cp27,0)=0 and nvl(cp57,0)=0" & _
'                     " and cp09=eep01(+) and eep05='" & textMan.Tag & "'" & _
'                     " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+) AND SUBSTR(NA02,1,2) IN ('C1','C3','C4')" & _
'                     " and mod(CP02,2)=0)"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'         End If
'         TxtUpdMsg.Visible = False
'      End If
'      '2016/8/24 END
      
      'Modify By Sindy 2020/3/12 Mark
'      'Modify By Sindy 2018/6/21 更新CFP程序人員相關資料
'      textMan.Text = Trim(textMan.Text)
'      textMan.Tag = Trim(textMan.Tag)
'      If Len(textMan.Text) = 5 _
'         And textMan.Text <> textMan.Tag _
'         And (UCase(strCode) = UCase("專利處轉信美日單號程序") Or UCase(strCode) = UCase("專利處轉信美日雙號程序") _
'              Or UCase(strCode) = UCase("專利處轉信美日以外單號程序") Or UCase(strCode) = UCase("專利處轉信美日以外雙號程序")) Then
'         TxtUpdMsg.Visible = True
'         '抓出CFP程序人員名單
'         'strSql = "SELECT DISTINCT NA73 FROM NATION UNION SELECT DISTINCT NA74 FROM NATION"
'         strSql = "select st01,st02 from staff where st05 in('85','83') and st04='1'"
'         rsTmp.CursorLocation = adUseClient
'         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         strCFPEmp = ""
'         If rsTmp.RecordCount > 0 Then
'            rsTmp.MoveFirst
'            Do While rsTmp.EOF = False
'               strCFPEmp = strCFPEmp & ",'" & rsTmp.Fields(0) & "'"
'               rsTmp.MoveNext
'            Loop
'            If strCFPEmp <> "" Then
'               strCFPEmp = Mid(strCFPEmp, 2)
'            Else
'               MsgBox "讀取CFP程序人員名單有誤,無法更新資料!!"
'               rsTmp.Close
'               GoTo ErrHnd
'            End If
'         End If
'         rsTmp.Close
'         '先找出最小收文日,速度才會快一點
'         strSql = "SELECT MIN(CP05) FROM CASEPROGRESS,PATENT,staff" & _
'                  " WHERE CP01 in ('CFP','CPS') AND CP158=0 AND CP159=0 AND CP05>=20030201 and cp14=st01(+) and st03='P12'" & _
'                  " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL"
'         rsTmp.CursorLocation = adUseClient
'         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         strMinCP05 = ""
'         If rsTmp.RecordCount > 0 Then
'            strMinCP05 = rsTmp.Fields(0)
'         End If
'         rsTmp.Close
'         If UCase(strCode) = UCase("專利處轉信美日單號程序") Or UCase(strCode) = UCase("專利處轉信美日雙號程序") Then
'            '美日單號
'            If UCase(strCode) = UCase("專利處轉信美日單號程序") Then
'               strSql = "UPDATE NATION SET NA73='" & textMan.Text & "' WHERE substr(na01,1,3) in('011','101')"
'               Pub_SeekTbLog strSql
'               cnnConnection.Execute strSql
'               strSQLCon = " and mod(CP02,2)=1" '單號
'            '美日雙號
'            Else
'               strSql = "UPDATE NATION SET NA74='" & textMan.Text & "' WHERE substr(na01,1,3) in('011','101')"
'               Pub_SeekTbLog strSql
'               cnnConnection.Execute strSql
'               strSQLCon = " and mod(CP02,2)=0" '雙號
'            End If
''            '更新進度檔CP14(承辦人)
''            ' AND CP04='00'
''            strSql = "UPDATE CASEPROGRESS SET CP14='" & textMan.Text & "'" & _
''                     " WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,PATENT,NATION" & _
''                     " WHERE CP14='" & textMan.Tag & "' AND CP158=0 AND CP159=0" & IIf(strMinCP05 <> "", " AND CP05>=" & strMinCP05, "") & _
''                     " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL AND PA09=NA01(+) AND substr(na01,1,3) in('011','101')" & _
''                     strSQLCon & " )"
''            Pub_SeekTbLog strSql
''            cnnConnection.Execute strSql
''            '更新待送件區EEP05(收受者)
''            strSql = "UPDATE EmpElectronProcess" & _
''                     " set eep05='" & textMan.Text & "'" & _
''                     " where (eep01,eep02,eep04) in" & _
''                     " (select eep01,eep02,eep04 from caseprogress,EmpElectronProcess,patent,nation WHERE cp01 in ('CFP','CPS') and CP158=0 and CP159=0" & _
''                     " and cp09=eep01(+) and eep05='" & textMan.Tag & "'" & _
''                     " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+) AND substr(na01,1,3) in('011','101')" & _
''                     strSQLCon & " )"
''            Pub_SeekTbLog strSql
''            cnnConnection.Execute strSql
'         Else
'            '美日以外單號
'            If UCase(strCode) = UCase("專利處轉信美日以外單號程序") Then
'               strSql = "UPDATE NATION SET NA73='" & textMan.Text & "' WHERE substr(na01,1,3) not in('011','101')"
'               Pub_SeekTbLog strSql
'               cnnConnection.Execute strSql
'               strSQLCon = " and mod(CP02,2)=1" '單號
'            '美日以外雙號
'            Else
'               strSql = "UPDATE NATION SET NA74='" & textMan.Text & "' WHERE substr(na01,1,3) not in('011','101')"
'               Pub_SeekTbLog strSql
'               cnnConnection.Execute strSql
'               strSQLCon = " and mod(CP02,2)=0" '雙號
'            End If
'''            '歐洲:子案各國也要更新(會有子案的只有EPC所以才沒有下國家別)
'''            strSql = "UPDATE CASEPROGRESS SET CP14='" & textMan.Text & "'" & _
'''                     " WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,PATENT" & _
'''                     " WHERE CP14='" & textMan.Tag & "' AND CP04<>'00' AND CP158=0 AND CP159=0" & IIf(strMinCP05 <> "", " AND CP05>=" & strMinCP05, "") & _
'''                     " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL" & _
'''                     strSQLCon & " )"
'''            Pub_SeekTbLog strSql
'''            cnnConnection.Execute strSql
''            '更新進度檔CP14(承辦人)
''            ' AND CP04='00'
''            strSql = "UPDATE CASEPROGRESS SET CP14='" & textMan.Text & "'" & _
''                     " WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,PATENT,NATION" & _
''                     " WHERE CP14='" & textMan.Tag & "' AND CP158=0 AND CP159=0" & IIf(strMinCP05 <> "", " AND CP05>=" & strMinCP05, "") & _
''                     " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL AND PA09=NA01(+) AND substr(na01,1,3) not in('011','101')" & _
''                     strSQLCon & " )"
''            Pub_SeekTbLog strSql
''            cnnConnection.Execute strSql
''            '更新待送件區EEP05(收受者)
''            strSql = "UPDATE EmpElectronProcess" & _
''                     " set eep05='" & textMan.Text & "'" & _
''                     " where (eep01,eep02,eep04) in" & _
''                     " (select eep01,eep02,eep04 from caseprogress,EmpElectronProcess,patent,nation WHERE cp01 in ('CFP','CPS') and CP158=0 and CP159=0" & _
''                     " and cp09=eep01(+) and eep05='" & textMan.Tag & "'" & _
''                     " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+) AND substr(na01,1,3) not in('011','101')" & _
''                     strSQLCon & " )"
''            Pub_SeekTbLog strSql
''            cnnConnection.Execute strSql
'         End If
'         '整批語法時(107/6/22):
''         UPDATE NATION SET NA73='85037',NA74='86032' WHERE SUBSTR(NA01,1,3) NOT IN ('011','101')
''         UPDATE NATION SET NA73='99043',NA74='79017' WHERE SUBSTR(NA01,1,3) IN ('011','101')
''         UPDATE CASEPROGRESS SET CP14=(select decode(mod(CP02,2),1,na73,na74) from patent,nation where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and pa09=na01) WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,PATENT,NATION WHERE CP14 IN (SELECT DISTINCT NA73 FROM NATION UNION SELECT DISTINCT NA74 FROM NATION) AND CP158=0 AND CP159=0 AND CP05>=20151130 AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL AND PA09=NA01(+))
''         UPDATE EmpElectronProcess set eep05=(select decode(mod(CP02,2),1,na73,na74) from caseprogress,patent,nation where eep01=cp09 and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and pa09=na01) where (eep01,eep02,eep04) in (select eep01,eep02,eep04 from caseprogress,EmpElectronProcess,patent,nation WHERE cp01 in ('CFP','CPS') and CP158=0 AND CP159=0 and cp09=eep01(+) and eep05 IN (SELECT DISTINCT NA73 FROM NATION UNION SELECT DISTINCT NA74 FROM NATION) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+))
'         'Modify By Sindy 2018/6/22 不區分國家,一起更新
'         If strCFPEmp <> "" Then
'            '更新進度檔CP14 (承辦人)
'            'modify by sonia 2018/6/28 加AND CP10<>'1913'條件
'            strSql = "UPDATE CASEPROGRESS SET CP14=(select decode(mod(CP02,2),1,na73,na74) from patent,nation where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and pa09=na01)" & _
'                     " WHERE CP09 IN (SELECT CP09 FROM CASEPROGRESS,PATENT,NATION" & _
'                     " WHERE CP14 IN (" & strCFPEmp & ")" & _
'                     " AND CP158=0 AND CP159=0 AND CP10<>'1913'" & IIf(strMinCP05 <> "", " AND CP05>=" & strMinCP05, "") & _
'                     " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL AND PA09=NA01(+)" & _
'                     strSQLCon & ")"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'            '更新待送件區EEP05(收受者)
'            strSql = "UPDATE EmpElectronProcess set eep05=(select decode(mod(CP02,2),1,na73,na74) from caseprogress,patent,nation where eep01=cp09 and cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and pa09=na01)" & _
'                     " where (eep01,eep02,eep04) in (select eep01,eep02,eep04" & _
'                     " from caseprogress,EmpElectronProcess,patent,nation" & _
'                     " WHERE cp01 in ('CFP','CPS') and CP158=0 AND CP159=0 and cp09=eep01(+)" & _
'                     " and eep05 IN (" & strCFPEmp & ")" & _
'                     " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+)" & _
'                     strSQLCon & ")"
'            Pub_SeekTbLog strSql
'            cnnConnection.Execute strSql
'         End If
'         '2018/6/22 END
'         TxtUpdMsg.Visible = False
'      End If
'      '2016/8/24 END
      
      textMan.Tag = textMan.Text 'Add By Sindy 2014/9/11
      cnnConnection.CommitTrans
      Screen.MousePointer = vbDefault 'Add By Sindy 2016/8/24
      
      GetAllData
      ShowCurrRecord strCode
   End If
   ModRecord = True
   
   Set rsTmp = Nothing 'Add By Sindy 2016/8/25
   Exit Function
   
ErrHnd:
   TxtUpdMsg.Visible = False
   Screen.MousePointer = vbDefault 'Add By Sindy 2016/8/24
   cnnConnection.RollbackTrans
   Set rsTmp = Nothing 'Add By Sindy 2016/8/25
   If Err.Description <> "" Then 'Add By Sindy 2014/9/11 +if
      MsgBox (Err.Description)
      Resume Next
   End If
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
'   Dim strSQL As String
'   Dim strLI01 As String
'   Dim strLI02 As String
'   Dim strLI08 As String
'
'   DelRecord = False
'
'On Error GoTo Err
'
'   strLI01 = m_CurrKEY(0)
'   strLI02 = m_CurrKEY(1)
'   strLI08 = m_CurrKEY(2)
'
'   strSQL = "DELETE FROM letterinput " & _
'            "WHERE li01 = " & strLI01 & " AND " & _
'                  "li02 = " & strLI02 & " and li08='" & strLI08 & "'"
'
'   cnnConnection.Execute strSQL
'
'   DelRecord = True
'
'   Exit Function
'Err:
'    cnnConnection.RollbackTrans
'    MsgBox "修改失敗！" & vbCrLf & Err.Description
End Function

