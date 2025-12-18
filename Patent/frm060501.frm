VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060501 
   BorderStyle     =   1  '單線固定
   Caption         =   "外翻人員資料維護"
   ClientHeight    =   3045
   ClientLeft      =   1755
   ClientTop       =   1860
   ClientWidth     =   7605
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   7605
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印地址條(&P)"
      Height          =   375
      Left            =   6060
      TabIndex        =   9
      Top             =   750
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   780
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
            Picture         =   "frm060501.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060501.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060501.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060501.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060501.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060501.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060501.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060501.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060501.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060501.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm060501.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   4
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
   Begin MSForms.TextBox txtA0I 
      Height          =   285
      Index           =   16
      Left            =   1020
      TabIndex        =   11
      Top             =   1860
      Width           =   6510
      VariousPropertyBits=   671105051
      MaxLength       =   50
      Size            =   "11483;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA0I 
      Height          =   690
      Index           =   6
      Left            =   1020
      TabIndex        =   3
      Top             =   2250
      Width           =   6510
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "11483;1217"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA0I 
      Height          =   285
      Index           =   1
      Left            =   1410
      TabIndex        =   0
      Top             =   750
      Width           =   885
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1561;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA0I 
      Height          =   285
      Index           =   4
      Left            =   1020
      TabIndex        =   1
      Top             =   1110
      Width           =   1275
      VariousPropertyBits=   671105051
      MaxLength       =   5
      Size            =   "2249;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtA0I 
      Height          =   285
      Index           =   3
      Left            =   1020
      TabIndex        =   2
      Top             =   1470
      Width           =   6510
      VariousPropertyBits=   671105051
      MaxLength       =   35
      Size            =   "11483;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "EMAIL:"
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   12
      Top             =   1920
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註:"
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   10
      Top             =   2250
      Width           =   405
   End
   Begin MSForms.Label lblName 
      Height          =   285
      Left            =   2370
      TabIndex        =   8
      Top             =   750
      Width           =   1275
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2249;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "郵遞區號:"
      Height          =   180
      Index           =   155
      Left            =   210
      TabIndex        =   7
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "聯絡地址:"
      Height          =   180
      Index           =   154
      Left            =   210
      TabIndex        =   6
      Top             =   1530
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "外翻員工編號:"
      Height          =   180
      Index           =   55
      Left            =   210
      TabIndex        =   5
      Top             =   810
      Width           =   1125
   End
End
Attribute VB_Name = "frm060501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/07 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
'Modify by Morgan 2009/1/21 消新增及刪除功能(改回寫到員工基本檔，原Table TransAddr 已不再使用)
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢

'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim TF_A0I As Integer '欄位數
Dim oText As Object, idx As Integer
Dim m_bConfirmCheck As Boolean


'Modify By Sindy 2022/3/2 淑華說:現在沒有在用了,都用寄電子檔
'Private Sub cmdOK_Click()
'   If txtA0I(1).Text <> "" Then
'       frm060501_1.Show
'       Me.Hide
'   End If
'End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      Form_KeyDown vbKeyF2, 0
   End If
   SetInputEntry
   UpdateToolbarState
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060501 = Nothing
End Sub

' 顯示資料
'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
   
   Dim stKey01 As String
   Dim adoRst As New ADODB.Recordset
   Dim stCols As String
   
   '先抓廠商檔沒有則抓員工基本檔
   'Modified by Morgan 2012/11/19 改都抓員工基本檔
   'stCols = " st01 a0i01,decode(a0i01,null,st08,a0i03) a0i03,decode(a0i01,null,st33,a0i04) a0i04,a0i06,a0i09,a0i10,a0i11,a0i16"
   stCols = " a0i01,st08 a0i03,st33 a0i04,a0i06,a0i09,a0i10,a0i11,st18 a0i16"
   
   stKey01 = txtA0I(1)
   
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT " & stCols & " FROM staff,acc0i0" & _
            " WHERE st03='F51' and st01 = '" & stKey01 & "' and a0i01(+)=st01 and a0i01 is not null"
      Case -2
         strExc(0) = "SELECT " & stCols & " FROM staff,acc0i0" & _
            " where st03='F51' and a0i01(+)=st01 and a0i01 is not null order by 1 ASC"
      Case -1
         strExc(0) = "SELECT " & stCols & " FROM staff,acc0i0" & _
            " WHERE st03='F51' and st01 <'" & stKey01 & "' and a0i01(+)=st01 and a0i01 is not null order by 1 DESC"
      Case 1
         strExc(0) = "SELECT " & stCols & " FROM staff,acc0i0" & _
            " WHERE st03='F51' and st01 >'" & stKey01 & "' and a0i01(+)=st01 and a0i01 is not null order by 1 ASC"
      Case 2
         strExc(0) = "SELECT " & stCols & " FROM staff,acc0i0" & _
            " where st03='F51' and a0i01(+)=st01 and a0i01 is not null order by 1 DESC"
   End Select
   intI = 1
   adoRst.MaxRecords = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      UpdateCtrlData adoRst
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         MsgBox "查無資料！", vbInformation
         ClearField
      End If
   End If
   
   If m_EditMode = 0 Then
      SetCtrlReadOnly True
   End If
   Set adoRst = Nothing
   If Me.Visible = True Then
      txtA0I(1).SetFocus
      txtA0I_GotFocus 1
   End If
End Function

Private Sub txtA0I_Change(Index As Integer)
   If Index = 1 Then
      If txtA0I(Index) = "" Then
         lblName = ""
      End If
   End If
End Sub

Private Sub txtA0I_GotFocus(Index As Integer)
   TextInverse txtA0I(Index)
   Select Case Index
   Case 1, 4, 16
      CloseIme
   Case Else
      OpenIme
   End Select
End Sub

Private Sub ClearField()
   For Each oText In txtA0I
      oText.Text = Empty
   Next
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(ByRef p_Rst As ADODB.Recordset)
   With p_Rst
   If .RecordCount > 0 Then
      txtA0I(1) = "" & .Fields("a0i01")
      txtA0I(4) = "" & .Fields("a0i04"): txtA0I(4).Tag = txtA0I(4)
      txtA0I(3) = "" & .Fields("a0i03"): txtA0I(3).Tag = txtA0I(3)
      txtA0I(16) = "" & .Fields("a0i16"): txtA0I(16).Tag = txtA0I(16)
      txtA0I(6) = "" & .Fields("a0i06"): txtA0I(6).Tag = txtA0I(6)
      If ClsPDGetStaffN(txtA0I(1), strExc(1), , True) = True Then
         lblName = strExc(1)
      End If
   End If
   End With
   txtA0I(1).Tag = txtA0I(1)
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtA0I
      oText.Locked = bLocked
   Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
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
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction KeyCode
            End If
         End If
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If

   End Select
End Sub
' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim bCancel As Boolean
   
   Select Case KeyCode
      Case vbKeyF3 ' 修改
         m_EditMode = 2
         UpdateToolbarState
         SetInputEntry
         
      Case vbKeyF4 ' 查詢
         m_EditMode = 4
         ClearField
         SetInputEntry
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
         If OnWork = False Then
            Exit Sub
         End If
         SetInputEntry
         UpdateToolbarState
         
      Case vbKeyF10 ' 取消
         bCancel = False
         Select Case m_EditMode
            Case 1, 2:
               If MsgBox("你並未存檔, 確定離開嗎?", vbYesNo + vbQuestion + vbDefaultButton2, "詢問") = vbYes Then
                  bCancel = True
               End If
            Case Else
               bCancel = True
         End Select
         If bCancel = True Then
            txtA0I(1) = txtA0I(1).Tag
            m_EditMode = 0
            SetInputEntry
            ShowRecord
            UpdateToolbarState
         End If
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub


'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bUpdate And txtA0I(1) <> "" Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtA0I(1) <> "" Then
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
      
      Case 1, 2, 3, 4 '維護
         TBar1.Buttons(2).Enabled = False
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

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 2
         SetCtrlReadOnly False
         txtA0I(1).Locked = True
         If Me.Visible = True Then
            txtA0I(4).SetFocus
         End If
      Case 4
         SetCtrlReadOnly True
         txtA0I(1).Locked = False
         If Me.Visible = True Then
            txtA0I(1).SetFocus
         End If
      Case Else
         SetCtrlReadOnly True
         If Me.Visible = True Then
            txtA0I(1).SetFocus
         End If
   End Select
   PUB_ChangeCaption Me, m_EditMode
End Sub

Private Function OnWork() As Boolean
   Select Case m_EditMode
         
      Case 2: '修改
         '重新檢查欄位有效性
         If TxtValidate() = True Then
            'Add by Sindy 2021/12/07 檢查畫面上的物件是否含有Unicode文字
            If PUB_ChkUniText(Me, True, True) = False Then
               Exit Function
            End If

            If ModRecord = True Then
               'Modified by Morgan 2012/11/19 改直接回寫員工檔Staff(利用Trigger回寫廠商檔acc0i0)
               'MsgBox "若扣單地址也需修改時請通知人事做相同修改！"
               OnWork = True
               m_EditMode = 0
               ShowRecord
            End If
         End If
      
      Case 4: '查詢
         If TxtValidate() = True Then
            If ShowRecord = True Then
               OnWork = True
               m_EditMode = 0
            Else
               txtA0I(1).SetFocus
               txtA0I_GotFocus 1
            End If
         End If
         
   End Select
End Function

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   m_bConfirmCheck = True
      
   For Each oText In txtA0I
      If oText.Locked = False And oText.Visible = True And oText.Enabled = True Then
         idx = oText.Index
         bCancel = False
         txtA0I_Validate idx, bCancel
         If bCancel = True Then
            txtA0I(idx).SetFocus
            txtA0I_GotFocus idx
            GoTo EscPoint
         End If
      End If
   Next
   
   TxtValidate = True
   
EscPoint:
   m_bConfirmCheck = False
End Function

Private Sub txtA0I_KeyPress(Index As Integer, KeyAscii As ReturnInteger)
   Select Case Index
   Case 1
      KeyAscii = UpperCase(KeyAscii)
   Case 3, 4
      KeyAscii = ChangeZIP(KeyAscii)
   Case 16
      PUB_EMailFilter Val(KeyAscii)
   End Select
End Sub

Private Sub txtA0I_Validate(Index As Integer, Cancel As Boolean)
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If Index = 1 Then
         If ClsPDGetStaffN(txtA0I(Index), strExc(1), , True) = False Then
            Cancel = True
         Else
            lblName = strExc(1)
         End If
      End If
   End If
End Sub

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   'Modified by Morgan 2012/11/19 改直接回寫員工檔Staff(利用Trigger回寫廠商檔acc0i0)
   If txtA0I(6) <> txtA0I(6).Tag Then
      stSQL = "UPDATE ACC0I0 SET A0I06=" & CNULL(ChgSQL(txtA0I(6))) & " WHERE a0i01='" & txtA0I(1) & "'"
      Pub_SeekTbLog stSQL
      cnnConnection.Execute stSQL, intI
   End If
   
   stSQL = "UPDATE staff SET "
   stSet = ""
   If txtA0I(3) <> txtA0I(3).Tag Then
      stSet = stSet & ",ST08=" & CNULL(ChgSQL(txtA0I(3)))
      bDifference = True
   End If
   
   If txtA0I(16) <> txtA0I(16).Tag Then
      stSet = stSet & ",ST18=" & CNULL(ChgSQL(txtA0I(16)))
      bDifference = True
   End If
   
   If txtA0I(4) <> txtA0I(4).Tag Then
      stSet = stSet & ",ST33=" & CNULL(ChgSQL(txtA0I(4)))
      bDifference = True
   End If
   
   If bDifference = True Then
      stSet = Mid(stSet, 2)
      stSQL = stSQL & stSet & " where ST01='" & txtA0I(1) & "'"
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

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      ' 修改
      Case 2: OnAction vbKeyF3
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

