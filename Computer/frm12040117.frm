VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040117 
   BorderStyle     =   1  '單線固定
   Caption         =   "CFP核駁報價資料維護"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7605
   Begin VB.TextBox textYF02_2 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1380
      Width           =   2355
   End
   Begin VB.TextBox textYF01_2 
      BorderStyle     =   0  '沒有框線
      Height          =   270
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2115
   End
   Begin VB.TextBox txtScore 
      Height          =   270
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   3
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtFee 
      Height          =   270
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   2
      Top             =   1710
      Width           =   855
   End
   Begin VB.TextBox textYF02 
      Height          =   270
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1380
      Width           =   372
   End
   Begin VB.TextBox textYF01 
      Height          =   270
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1020
      Width           =   612
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6345
      Top             =   1050
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
            Picture         =   "frm12040117.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040117.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040117.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040117.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040117.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040117.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040117.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040117.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040117.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040117.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040117.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7605
      _ExtentX        =   13414
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
   End
   Begin MSForms.TextBox textCUID 
      Height          =   270
      Left            =   495
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   690
      Width           =   6630
      VariousPropertyBits=   671105055
      Size            =   "11695;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   630
      Left            =   1440
      TabIndex        =   4
      Top             =   2370
      Width           =   5715
      VariousPropertyBits=   -1467989989
      MaxLength       =   150
      ScrollBars      =   2
      Size            =   "10081;1111"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "說明 :"
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   12
      Top             =   2370
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點數 :"
      Height          =   180
      Index           =   6
      Left            =   480
      TabIndex        =   8
      Top             =   2085
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "費用 :"
      Height          =   180
      Index           =   5
      Left            =   480
      TabIndex        =   7
      Top             =   1755
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利種類 :"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   1425
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家 :"
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   1065
      Width           =   930
   End
End
Attribute VB_Name = "frm12040117"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/22 改成Form2.0 (txtDesc,textCUID)
'Create by Lydia 2017/06/06 CFP核駁報價資料維護
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Private Const mYF04 As String = "107" '案件性質
Private Const mYF03 As String = "Y00000000" '代理人代號

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type

Dim m_FieldList() As FIELDITEM

Dim oText
Dim TF_YF As Integer
Dim idx As Integer

Private Sub Form_Initialize()
   strExc(0) = "select * from patentyearfee where rownum<1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   TF_YF = RsTemp.Fields.Count
   ReDim m_FieldList(TF_YF) As FIELDITEM
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         KeyCode = 0
         If tlbar.Buttons(1).Enabled = True Then
            If m_EditMode = 0 Then
               OnAction vbKeyF2
            End If
         End If
      
      Case vbKeyF3 ' 修改
         KeyCode = 0
         If tlbar.Buttons(2).Enabled = True Then
            If m_EditMode = 0 Then
               OnAction vbKeyF3
            End If
         End If
         
      Case vbKeyF5 ' 刪除
         KeyCode = 0
         If tlbar.Buttons(3).Enabled = True Then
            If m_EditMode = 0 Then
               OnAction vbKeyF5
            End If
         End If
      
      Case vbKeyF4 ' 查詢
         KeyCode = 0
         If tlbar.Buttons(4).Enabled = True Then
            If m_EditMode = 0 Then
               OnAction vbKeyF4
            End If
         End If
      
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd
         If tlbar.Buttons(6).Enabled = True Then
            If m_EditMode = 0 Then
               OnAction KeyCode
            End If
         End If
         KeyCode = 0
         
      Case vbKeyF9, vbKeyF10
         If tlbar.Buttons(11).Enabled = True Then
            If m_EditMode <> 0 Then
               OnAction KeyCode
            End If
         End If
         KeyCode = 0
         
      Case vbKeyEscape
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            KeyCode = 0
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction vbKeyEscape
            End If
         End If
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 KeyCode 事件的話，也會做到
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
         
   End Select
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   textYF01_2.BackColor = &H8000000F
   textYF02_2.BackColor = &H8000000F
   InitialField
   m_EditMode = 0
   ShowRecord -2
   UpdateToolbarState
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040117 = Nothing
End Sub

Private Sub textYF01_Change()
   If Len(textYF01) = 3 Then
      textYF01_2 = GetNationName(textYF01, 0)
   Else
      textYF01_2 = ""
   End If
End Sub

Private Sub textYF02_Change()
   If Len(textYF02) = 1 Then
      If textYF01 < "010" Then
         textYF02_2 = GetPatentName(textYF02, 0)
      Else
         textYF02_2 = GetPatentName(textYF02, 1)
      End If
   Else
      textYF02_2 = ""
   End If
End Sub

Private Sub textYF02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textYF02) = False Then
      If m_EditMode <> 0 Then
         If IsEmptyText(textYF02_2) = True Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "專利種類不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textYF02_GotFocus
         End If
      End If
   End If
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
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
' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF2 ' 新增
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         textYF01.SetFocus
         
      Case vbKeyF3 ' 修改
         m_EditMode = 2
         SetCtrlReadOnly False
         UpdateToolbarState
         
      Case vbKeyF5 ' 刪除
         If MsgBox("是否要刪除此筆資料?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
         
      Case vbKeyF4 ' 查詢
         m_EditMode = 4
         ClearField
         SetCtrlReadOnly True
         textYF01.Locked = False
         textYF02.Locked = False
         UpdateToolbarState
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
         
      Case vbKeyF9 ' 確定
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         
      Case vbKeyF10 ' 取消
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  m_EditMode = 0
                  If textYF01.Tag <> "" Then
                     textYF01 = textYF01.Tag
                     textYF02 = textYF02.Tag
                     ShowRecord
                  Else
                     ClearField
                  End If
                  UpdateToolbarState
               End If
               
            Case Else
               m_EditMode = 0
               If textYF01.Tag <> "" Then
                  textYF01 = textYF01.Tag
                  textYF02 = textYF02.Tag
                  ShowRecord
               Else
                  ClearField
               End If
               UpdateToolbarState
         End Select
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In Me.Controls
      If TypeName(oText) = "TextBox" Then
         oText.Locked = bLocked
      End If
   Next
End Sub
'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bInsert Then
            tlbar.Buttons(1).Enabled = True
         Else
            tlbar.Buttons(1).Enabled = False
         End If
         If m_bUpdate And textYF01 <> "" Then
            tlbar.Buttons(2).Enabled = True
         Else
            tlbar.Buttons(2).Enabled = False
         End If
         If m_bDelete And textYF01 <> "" Then
            tlbar.Buttons(3).Enabled = True
         Else
            tlbar.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(4).Enabled = True
         Else
            tlbar.Buttons(4).Enabled = False
         End If
         If m_bQuery And textYF01 <> "" Then
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
      
      Case 1, 2, 3, 4 '維護
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

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0, Optional ByVal p_bGridOnly As Boolean) As Boolean
   Dim stCon As String, stKey As String
   
   stCon = " AND YF04='" & mYF04 & "' AND YF05='1' AND INSTR('000,013,020,044,056',YF01)=0 AND YF03 = '" & mYF03 & "'"
   
   stKey = textYF01 & textYF02 & mYF03 & mYF04

   Select Case p_iWay
      Case 0, 3 '當筆
         strExc(0) = "SELECT *" & _
            " FROM patentyearfee WHERE YF01||YF02||YF03||YF04='" & stKey & "'" & stCon
         
      Case -2 '首筆
         strExc(0) = "SELECT *" & _
            " FROM patentyearfee A WHERE YF01||YF02||YF03||YF04||YF05=(SELECT MIN(B.YF01||B.YF02||B.YF03||B.YF04||B.YF05)" & _
            " FROM patentyearfee B WHERE 1=1" & stCon & ")"

      Case -1 '前筆
         strExc(0) = "SELECT *" & _
            " FROM patentyearfee A WHERE YF01||YF02||YF03||YF04||YF05=(SELECT MAX(B.YF01||B.YF02||B.YF03||B.YF04||B.YF05)" & _
            " FROM patentyearfee B WHERE B.YF01||B.YF02||B.YF03||B.YF04<'" & stKey & "'" & stCon & ")"

      Case 1 '後筆
         strExc(0) = "SELECT *" & _
            " FROM patentyearfee A WHERE YF01||YF02||YF03||YF04||YF05=(SELECT MIN(B.YF01||B.YF02||B.YF03||B.YF04||B.YF05)" & _
            " FROM patentyearfee B WHERE B.YF01||B.YF02||B.YF03||B.YF04>'" & stKey & "'" & stCon & ")"

      Case 2 '末筆
         strExc(0) = "SELECT *" & _
            " FROM patentyearfee A WHERE YF01||YF02||YF03||YF04||YF05=(SELECT MAX(B.YF01||B.YF02||B.YF03||B.YF04||B.YF05)" & _
            " FROM patentyearfee B WHERE 1=1" & stCon & ")"
   End Select
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      UpdateCtrlData RsTemp
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         MsgBox "無資料！", vbInformation
         ClearField
      End If
   End If
      
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   
   If Me.Visible = True Then
      textYF01.SetFocus
      textYF01_GotFocus
   End If
   
End Function

Private Sub ClearField()
   For Each oText In Me.Controls
      If TypeName(oText) = "TextBox" Then
         oText.Text = Empty
      End If
   Next
   For intI = 1 To TF_YF
      m_FieldList(intI).fiOldData = Empty
      m_FieldList(intI).fiNewData = Empty
   Next
   textCUID = Empty
   textYF01_2 = Empty
   textYF02_2 = Empty
End Sub

Private Function OnWork() As Boolean

   Select Case m_EditMode
      Case 1: '新增
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            '檢查是否重複
            If PUB_PYFIsExists(textYF01, textYF02, mYF03, mYF04, "1") = True Then
               strExc(0) = "新增資料"
               strExc(1) = "該筆記錄已存在"
               MsgBox strExc(1), vbOKOnly, strExc(0)
               Exit Function
            End If
            UpdateFieldNewData
            If AddRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord 3
            End If
         End If
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            UpdateFieldNewData
            If ModRecord = True Then
               OnWork = True
               m_EditMode = 0
               ShowRecord 3
            End If
         End If
         
      Case 3: '刪除
         If DelRecord = True Then
            OnWork = True
            m_EditMode = 0
            ShowRecord 2
         End If
      
       Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               textYF01.SetFocus
               textYF01_GotFocus
            End If
         End If
   End Select
   
End Function

Private Sub UpdateFieldNewData()
   m_FieldList(1).fiNewData = textYF01.Text
   m_FieldList(2).fiNewData = textYF02.Text
   m_FieldList(3).fiNewData = mYF03
   m_FieldList(4).fiNewData = mYF04
   m_FieldList(5).fiNewData = "1"
   m_FieldList(6).fiNewData = Val(Format(txtScore)) * 1000
   m_FieldList(7).fiNewData = Val(Format(txtFee)) - m_FieldList(6).fiNewData
   m_FieldList(8).fiNewData = txtDesc.Text
End Sub

Private Sub textYF01_GotFocus()
   TextInverse textYF01
End Sub

Private Sub textYF01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textYF01) = False Then
      If m_EditMode <> 0 Then
         If IsEmptyText(textYF01_2) = True Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "申請國家代號不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         ElseIf InStr("000,013,020,044,056", textYF01) > 0 Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "不可輸入該申請國家！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         End If
         If Cancel = True Then
            textYF01_GotFocus
         End If
      End If
   End If
End Sub

Private Sub textYF02_GotFocus()
   TextInverse textYF02
End Sub

Private Sub textYF02_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not InStr("1,2,3", Chr(KeyAscii)) > 0 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDesc_GotFocus()
   TextInverse txtDesc
End Sub


Private Sub txtFee_GotFocus()
   TextInverse txtFee
End Sub

Private Sub txtFee_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtScore_GotFocus()
   TextInverse txtScore
End Sub

Private Sub txtScore_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

' 新增記錄
Private Function AddRecord() As Boolean
   Dim stSQL As String, stCols As String, stValues As String
     
   cnnConnection.BeginTrans
   
On Error GoTo ErrHand
   
   '畫面有的欄位才更新
   For idx = 1 To TF_YF
      If m_FieldList(idx).fiNewData <> "" Then
         stCols = stCols & "," & m_FieldList(idx).fiName
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stValues = stValues & "," & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stValues = stValues & "," & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   If stCols <> "" Then
      stCols = Mid(stCols, 2)
      stValues = Mid(stValues, 2)
   End If
   stSQL = "INSERT INTO patentyearfee (" & stCols & ") Values (" & stValues & ")"
   
   Pub_SeekTbLog stSQL
   
   cnnConnection.Execute stSQL, intI

   cnnConnection.CommitTrans
   
   textYF01.Tag = textYF01
   textYF02.Tag = textYF02
   
   AddRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.NUMBER = -2147217873 Then
      MsgBox "相同資料已存在！"
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   Dim CUID(1 To 6) As String
   
   With p_Rst
      If .RecordCount > 0 Then
         For idx = 1 To TF_YF
            m_FieldList(idx).fiOldData = "" & .Fields(m_FieldList(idx).fiName)
            m_FieldList(idx).fiNewData = m_FieldList(idx).fiOldData
            'Modified by Lydia 2017/06/29 O12和O8的Type不同,統一做文字處理
            'If .Fields(m_FieldList(idx).fiName).Type = 200 Then
               m_FieldList(idx).fiType = 0
            'Else
            '   m_FieldList(idx).fiType = 1
            'End If
            'end 2017/06/29
         Next
         
         textYF01.Text = m_FieldList(1).fiNewData
         textYF01.Tag = textYF01.Text
         textYF02.Text = m_FieldList(2).fiNewData
         textYF02.Tag = textYF02.Text
         txtScore.Text = Val(m_FieldList(6).fiNewData) / 1000
         txtFee.Text = Val(m_FieldList(6).fiNewData) + Val(m_FieldList(7).fiNewData)
         txtDesc.Text = m_FieldList(8).fiNewData
   
         CUID(1) = "" & .Fields("YF09")
         CUID(2) = "" & .Fields("YF10")
         CUID(3) = "" & .Fields("YF11")
         CUID(4) = "" & .Fields("YF12")
         CUID(5) = "" & .Fields("YF13")
         CUID(6) = "" & .Fields("YF14")
      End If
   End With
   UpdateCUID CUID, textCUID
   
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   For idx = 1 To TF_YF
      m_FieldList(idx).fiName = "YF" & Format(idx, "00")
   Next
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub
' 刪除記錄
Private Function DelRecord() As Boolean
   Dim stSQL As String
   
   cnnConnection.BeginTrans
On Error GoTo ErrHand
   
   '刪除資料
   stSQL = "delete from PATENTYEARFEE where YF01='" & m_FieldList(1).fiOldData & "' AND YF02='" & m_FieldList(2).fiOldData & "'" & _
      " AND YF03='" & m_FieldList(3).fiOldData & "' AND YF04='" & m_FieldList(4).fiOldData & "'"
   
   Pub_SeekTbLog stSQL
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   
   DelRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

End Function

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stSQL = "UPDATE PATENTYEARFEE SET "
   stSet = ""
   For idx = 1 To TF_YF
      If m_FieldList(idx).fiNewData <> m_FieldList(idx).fiOldData Then
         bDifference = True
         '文字
         If m_FieldList(idx).fiType = 0 Then
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(ChgSQL(m_FieldList(idx).fiNewData))
         '數字
         Else
            stSet = stSet & "," & m_FieldList(idx).fiName & "=" & CNULL(m_FieldList(idx).fiNewData, True)
         End If
      End If
   Next
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where YF01='" & m_FieldList(1).fiOldData & "'" & _
         " AND YF02='" & m_FieldList(2).fiOldData & "' AND YF03='" & m_FieldList(3).fiOldData & "'" & _
         " AND YF04='" & m_FieldList(4).fiOldData & "'"
      Pub_SeekTbLog stSQL
      
      cnnConnection.Execute stSQL, intI
   End If
   
   cnnConnection.CommitTrans
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

Private Function TxtValidate() As Boolean
   
   Dim Cancel As Boolean
   
   'Added by Morgan 2021/12/22 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/22
   
   If textYF01 = "" Then
      ShowMsg "申請國家不可空白 !"
      textYF01.SetFocus
      Exit Function
   End If
   
   If textYF02 = "" Then
      ShowMsg "專利種類不可空白 !"
      textYF02.SetFocus
      Exit Function
   End If
   
   textYF01_Validate Cancel
   If Cancel = True Then
      textYF01.SetFocus
      Exit Function
   End If
            
   '維護
   If m_EditMode <> 4 Then
      If txtFee = "" Then
         ShowMsg "費用不可空白 !"
         txtFee.SetFocus
         txtFee_GotFocus
         Exit Function
      End If
      If txtScore = "" Then
         ShowMsg "點數不可空白 !"
         txtScore.SetFocus
         txtScore_GotFocus
         Exit Function
      End If
      If Val(txtScore) * 1000 > Val(txtFee) Then
         ShowMsg "點數不可大於費用 !"
         txtScore.SetFocus
         txtScore_GotFocus
         Exit Function
      End If
   End If
   TxtValidate = True
End Function
