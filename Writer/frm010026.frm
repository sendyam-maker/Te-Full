VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010026 
   BorderStyle     =   1  '單線固定
   Caption         =   "分所案號維護"
   ClientHeight    =   4584
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7560
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4584
   ScaleWidth      =   7560
   Begin VB.TextBox txtEdit 
      Height          =   270
      Index           =   0
      Left            =   5010
      MaxLength       =   50
      TabIndex        =   4
      Top             =   712
      Width           =   2025
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Index           =   4
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   3
      Top             =   712
      Width           =   330
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Index           =   3
      Left            =   2970
      MaxLength       =   1
      TabIndex        =   2
      Top             =   712
      Width           =   240
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Index           =   2
      Left            =   1980
      MaxLength       =   6
      TabIndex        =   1
      Top             =   712
      Width           =   960
   End
   Begin VB.TextBox txtKey 
      Height          =   270
      Index           =   1
      Left            =   1425
      MaxLength       =   3
      TabIndex        =   0
      Top             =   712
      Width           =   510
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   4050
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
            Picture         =   "frm010026.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010026.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010026.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010026.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010026.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010026.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010026.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010026.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010026.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010026.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm010026.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
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
            Object.Tag             =   "F2"
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
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   14
      Left            =   3780
      TabIndex        =   32
      Top             =   3990
      Width           =   1125
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   13
      Left            =   2610
      TabIndex        =   31
      Top             =   3990
      Width           =   1125
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   12
      Left            =   1440
      TabIndex        =   30
      Top             =   3990
      Width           =   1125
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   11
      Left            =   1440
      TabIndex        =   29
      Top             =   3690
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   10
      Left            =   1440
      TabIndex        =   28
      Top             =   3390
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   9
      Left            =   1440
      TabIndex        =   27
      Top             =   3090
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   8
      Left            =   1440
      TabIndex        =   26
      Top             =   2790
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   7
      Left            =   1440
      TabIndex        =   25
      Top             =   2490
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   6
      Left            =   5010
      TabIndex        =   24
      Top             =   2190
      Width           =   2115
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請案號："
      Height          =   255
      Index           =   7
      Left            =   3750
      TabIndex        =   23
      Top             =   2190
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　(英)："
      Height          =   255
      Index           =   8
      Left            =   180
      TabIndex        =   22
      Top             =   1290
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "分所案號："
      Height          =   255
      Index           =   17
      Left            =   3780
      TabIndex        =   21
      Top             =   720
      Width           =   1200
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   5
      Left            =   1440
      TabIndex        =   20
      Top             =   2190
      Width           =   2115
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   4
      Left            =   1440
      TabIndex        =   19
      Top             =   1890
      Width           =   2115
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   3
      Left            =   1425
      TabIndex        =   18
      Top             =   1590
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Update："
      Height          =   255
      Index           =   15
      Left            =   180
      TabIndex        =   17
      Top             =   3990
      Width           =   1200
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   2
      Left            =   1440
      TabIndex        =   16
      Top             =   1290
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "申請人２："
      Height          =   255
      Index           =   12
      Left            =   180
      TabIndex        =   15
      Top             =   2790
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "申請人１："
      Height          =   255
      Index           =   11
      Left            =   180
      TabIndex        =   14
      Top             =   2490
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "申請日："
      Height          =   255
      Index           =   10
      Left            =   180
      TabIndex        =   13
      Top             =   2190
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   255
      Index           =   9
      Left            =   180
      TabIndex        =   12
      Top             =   1890
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　(日)："
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   11
      Top             =   1590
      Width           =   1200
   End
   Begin MSForms.Label lblDisp 
      Height          =   270
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   990
      Width           =   6030
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "申請人５："
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   9
      Top             =   3690
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "申請人４："
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   8
      Top             =   3390
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "申請人３："
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   7
      Top             =   3090
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱(中)："
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   6
      Top             =   1005
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   720
      Width           =   1200
   End
End
Attribute VB_Name = "frm010026"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/07/30 Form2.0已修改 lblDisp
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

'目前狀態
Dim iCurState As Integer
'前一案號
Dim lst_KEY(1 To 4) As String
'目前案號
Dim cur_KEY(1 To 4) As String
'最後收文智權人員
Dim stKeySales As String
'最後收文承辦人或協辦人員 2011/8/15多判斷協辦人員
Dim stRcvEmp As String
Dim stLos04 As String 'Added by Lydia 2023/09/08 案源案件之第一介紹人

Private Sub Form_Load()
   MoveFormToCenter Me
   Call SetToolBar(0)
   Call FormReset
   '預設為查詢
   Call SetToolBar(4)
   Call SetInputs(4)
   iCurState = 4
End Sub

'工具列控制
Private Sub SetToolBar(iStatus As Integer)
Dim i As Integer
   
   For i = 1 To 13
      tlbar.Buttons(i).Enabled = False
   Next
   tlbar.Buttons(14).Enabled = True
   
   Select Case iStatus
      Case 0
      '瀏覽
         tlbar.Buttons(2).Enabled = True
         tlbar.Buttons(4).Enabled = True
      Case 1
      '新增
      Case 2
      '修改
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
      Case 3
      '刪除
      Case 4
      '查詢
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
      Case Else
   End Select
End Sub

Private Sub FormReset()
Dim oText As TextBox, oLabel 'Modify by Amy 2021/07/30 原:oLabel As LABEL
   
   For Each oText In txtKey
      oText.Text = ""
      oText.Locked = True
   Next
   
   For Each oText In txtEdit
      oText.Text = ""
      oText.Locked = True
   Next

   For Each oLabel In lblDisp
      oLabel.Caption = ""
   Next
End Sub

Private Sub SetInputs(Optional ByVal iStatus As Integer = 0)
Dim oText As TextBox, oLabel 'Modify by Amy 2021/07/30 原:oLabel As LABEL
   
   Select Case iStatus
      Case 2
      '修改
         For Each oText In txtKey
            oText.Locked = True
            oText.Enabled = False
         Next
         For Each oText In txtEdit
            oText.Locked = False
         Next
      Case 4
      '查詢
         For Each oText In txtKey
            oText.Text = ""
            oText.Enabled = True
            oText.Locked = False
         Next
         For Each oText In txtEdit
            oText.Text = ""
            oText.Locked = True
         Next
         For Each oLabel In lblDisp
            oLabel.Caption = ""
         Next
      Case Else
      '其他
         For Each oText In txtKey
            oText.Enabled = True
            oText.Locked = True
         Next
         For Each oText In txtEdit
            oText.Locked = True
         Next
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010026 = Nothing
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)

   Select Case Button.Index
      Case 1
      '新增
      Case 2
      '修改
         '最後收文智權人員或承辦人或協辦人員為該所人員才可修改
         'Modified by Lydia 2015/10/05 '承辦法務'改為'協辦人員'
         'Modified by Lydia 2023/09/08 +案源案件之第一介紹人stLos04
         'If (PUB_GetST06(stKeySales) <> pub_strUserOffice And _
            PUB_GetST06(stRcvEmp) <> pub_strUserOffice) And _
            Pub_StrUserSt03 <> "M51" Then
         If Pub_StrUserSt03 <> "M51" And InStr(PUB_GetST06(stKeySales) & "," & PUB_GetST06(stRcvEmp) & "," & PUB_GetST06(stLos04), pub_strUserOffice) = 0 Then
            MsgBox "此案與你的所別不同！", vbCritical
            Exit Sub
         Else
            Call SetToolBar(2)
            Call SetInputs(2)
            txtEdit(0).SetFocus
            iCurState = 2
         End If
      Case 3
      '刪除
      Case 4
      '查詢
         lst_KEY(1) = cur_KEY(1)
         lst_KEY(2) = cur_KEY(2)
         lst_KEY(3) = cur_KEY(3)
         lst_KEY(4) = cur_KEY(4)
         Call SetToolBar(4)
         Call SetInputs(4)
         txtKey(1).SetFocus
         Call txtkey_GotFocus(1)
         iCurState = 4
      Case 11
      '確定
         '查詢
         If iCurState = 4 Then
            If txtKey(1) = "" Then
               MsgBox "系統別不可空白！", vbCritical
               Exit Sub
            End If
            txtKey(1) = Trim(txtKey(1).Text)
            If txtKey(2) = "" Then
               MsgBox "案號不可空白！", vbCritical
               Exit Sub
            End If
            txtKey(2) = Left(Me.txtKey(2).Text & "000000000", 6)
            txtKey(3) = Left(Me.txtKey(3).Text & "0", 1)
            txtKey(4) = Left(Me.txtKey(4).Text & "00", 2)
            cur_KEY(1) = txtKey(1)
            cur_KEY(2) = txtKey(2)
            cur_KEY(3) = txtKey(3)
            cur_KEY(4) = txtKey(4)
         '修改
         ElseIf iCurState = 2 Then
            If UpdateData() = True Then
               'MsgBox "修改成功", vbInformation
            Else
               Exit Sub
            End If
         End If
         
         If doQuery(4) = True Then
            Call SetToolBar(0)
            Call SetInputs
            iCurState = 0
         End If
         txtKey(1).SetFocus
         Call txtkey_GotFocus(1)
         
      Case 12
      '取消
         If iCurState = 4 Then
            cur_KEY(1) = lst_KEY(1)
            cur_KEY(2) = lst_KEY(2)
            cur_KEY(3) = lst_KEY(3)
            cur_KEY(4) = lst_KEY(4)
            If cur_KEY(1) = "" Then
               MsgBox "無前次查詢紀錄，不可取消！", vbCritical
               Exit Sub
            Else
               Call doQuery(4)
            End If
         ElseIf iCurState = 2 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
               Exit Sub
            Else
               Call doQuery(4)
            End If
         End If
         Call SetToolBar(0)
         Call SetInputs
         iCurState = 0
      Case 14
      '結束
         Unload Me
   End Select
End Sub

Private Function UpdateData() As Boolean
Dim strSql As String, lngEffRec As Long
Dim strCustContNo As String
   
On Error GoTo ErrHand
   
   Select Case cur_KEY(1)
      '專利
      Case "P", "CFP", "FCP"
         strSql = "UPDATE PATENT SET PA47='" & txtEdit(0) & "',PA95='" & strUserNum & "',PA96=to_number(to_char(sysdate,'YYYYMMDD')),PA97=to_number(to_char(sysdate,'HH24MI'))" & _
            " WHERE PA01='" & cur_KEY(1) & "' AND PA02='" & cur_KEY(2) & "' AND PA03='" & cur_KEY(3) & "' AND PA04='" & cur_KEY(4) & "'"
      '商標
      Case "T", "CFT", "FCT", "TF"
         strSql = "UPDATE TRADEMARK SET TM34='" & txtEdit(0) & "',TM62='" & strUserNum & "',TM63=to_number(to_char(sysdate,'YYYYMMDD')),TM64=to_number(to_char(sysdate,'HH24MI'))" & _
            " WHERE TM01='" & cur_KEY(1) & "' AND TM02='" & cur_KEY(2) & "' AND TM03='" & cur_KEY(3) & "' AND TM04='" & cur_KEY(4) & "'"
      '法務
      Case "CFL", "FCL", "L", "LIN"
         strSql = "UPDATE LAWCASE SET LC16='" & txtEdit(0) & "',LC31='" & strUserNum & "',LC32=to_number(to_char(sysdate,'YYYYMMDD')),LC33=to_number(to_char(sysdate,'HH24MI'))" & _
            " WHERE LC01='" & cur_KEY(1) & "' AND LC02='" & cur_KEY(2) & "' AND LC03='" & cur_KEY(3) & "' AND LC04='" & cur_KEY(4) & "'"
      '顧問
      Case "LA"
         strSql = "UPDATE HIRECASE SET HC07='" & txtEdit(0) & "',HC16='" & strUserNum & "',HC17=to_number(to_char(sysdate,'YYYYMMDD')),HC18=to_number(to_char(sysdate,'HH24MI'))" & _
            " WHERE HC01='" & cur_KEY(1) & "' AND HC02='" & cur_KEY(2) & "' AND HC03='" & cur_KEY(3) & "' AND HC04='" & cur_KEY(4) & "'"
      '服務
      '"CFC","CPS","FG","PS","S","TB","TC","TD","TM","TR","TS","TT"
      Case Else
         strSql = "UPDATE SERVICEPRACTICE SET SP28='" & txtEdit(0) & "',SP55='" & strUserNum & "',SP56=to_number(to_char(sysdate,'YYYYMMDD')),SP57=to_number(to_char(sysdate,'HH24MI'))" & _
            " WHERE SP01='" & cur_KEY(1) & "' AND SP02='" & cur_KEY(2) & "' AND SP03='" & cur_KEY(3) & "' AND SP04='" & cur_KEY(4) & "'"
   End Select
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, lngEffRec
   cnnConnection.CommitTrans
   UpdateData = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF3
      '修改
         If tlbar.Buttons(2).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(2))
         End If
      Case vbKeyF4
      '查詢
         If tlbar.Buttons(4).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(4))
         End If
      Case vbKeyF9, vbKeyReturn
      '確定
         If tlbar.Buttons(11).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(11))
         End If
      Case vbKeyF10
      '取消
         If tlbar.Buttons(12).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(12))
         End If
      Case vbKeyEscape
      '結束
        If tlbar.Buttons(14).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(14))
         End If
    End Select
End Sub

Private Function doQuery(ByVal iAct As Integer) As Boolean
Dim strSql As String, rsQuery As New ADODB.Recordset
   
   rsQuery.MaxRecords = 2
   rsQuery.CursorLocation = adUseClient
   doQuery = False
   
   Select Case iAct
      Case 4
      '查詢
         Select Case cur_KEY(1)
            '專利
            Case "P", "CFP", "FCP"
               strSql = "Select CU12,CU13 From PATENT, Customer WHERE CU01=SUBSTR(PA26,1,8) AND CU02=SUBSTR(PA26,9,1)" & _
                  " AND PA01='" & cur_KEY(1) & "' AND PA02='" & cur_KEY(2) & "' AND PA03='" & cur_KEY(3) & "' AND PA04='" & cur_KEY(4) & "'"
            '商標
            Case "T", "CFT", "FCT", "TF"
               strSql = "Select CU12,CU13 From TRADEMARK, Customer WHERE CU01=SUBSTR(TM23,1,8) AND CU02=SUBSTR(TM23,9,1)" & _
                  " AND TM01='" & cur_KEY(1) & "' AND TM02='" & cur_KEY(2) & "' AND TM03='" & cur_KEY(3) & "' AND TM04='" & cur_KEY(4) & "'"
            '法務
            Case "CFL", "FCL", "L", "LIN"
               strSql = "Select CU12,CU13 From LAWCASE, Customer WHERE CU01=SUBSTR(LC11,1,8) AND CU02=SUBSTR(LC11,9,1)" & _
                  " AND LC01='" & cur_KEY(1) & "' AND LC02='" & cur_KEY(2) & "' AND LC03='" & cur_KEY(3) & "' AND LC04='" & cur_KEY(4) & "'"
            '顧問
            Case "LA"
               strSql = "Select CU12,CU13 From HIRECASE, Customer WHERE CU01=SUBSTR(HC05,1,8) AND CU02=SUBSTR(HC05,9,1)" & _
                  " AND HC01='" & cur_KEY(1) & "' AND HC02='" & cur_KEY(2) & "' AND HC03='" & cur_KEY(3) & "' AND HC04='" & cur_KEY(4) & "'"
            '服務
            '"CFC","CPS","FG","PS","S","TB","TC","TD","TM","TR","TS","TT"
            Case Else
               strSql = "Select CU12,CU13,SP09 From SERVICEPRACTICE, Customer WHERE CU01=SUBSTR(SP08,1,8) AND CU02=SUBSTR(SP08,9,1)" & _
                  " AND SP01='" & cur_KEY(1) & "' AND SP02='" & cur_KEY(2) & "' AND SP03='" & cur_KEY(3) & "' AND SP04='" & cur_KEY(4) & "'"
         End Select
   End Select
   
On Error GoTo ErrHand
   
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      '讀取最後收文智權人員
      If cur_KEY(1) = "FCP" Or cur_KEY(1) = "FG" Then
        stKeySales = PUB_GetFCPSalesNo(cur_KEY(1), cur_KEY(2), cur_KEY(3), cur_KEY(4))
      ElseIf cur_KEY(1) = "FCL" Or cur_KEY(1) = "LIN" Then
        stKeySales = PUB_GetFCLSalesNo(cur_KEY(1), cur_KEY(2), cur_KEY(3), cur_KEY(4))
      ElseIf cur_KEY(1) = "FCT" Then
        stKeySales = PUB_GetFCTSalesNo(cur_KEY(1), cur_KEY(2), cur_KEY(3), cur_KEY(4))
      ElseIf cur_KEY(1) = "S" Then
        If rsQuery.Fields("SP09") = "000" Then
           stKeySales = PUB_GetFCTSalesNo(cur_KEY(1), cur_KEY(2), cur_KEY(3), cur_KEY(4))
        Else
           stKeySales = PUB_GetAKindSalesNo(cur_KEY(1), cur_KEY(2), cur_KEY(3), cur_KEY(4))
        End If
      Else
        stKeySales = PUB_GetAKindSalesNo(cur_KEY(1), cur_KEY(2), cur_KEY(3), cur_KEY(4))
      End If
      '讀取最後收文承辦人,2011/8/15同時抓協辦人員
      stRcvEmp = ""
      strSql = "SELECT * FROM CaseProgress WHERE CP01='" & cur_KEY(1) & "' AND CP02='" & cur_KEY(2) & "' AND CP03='" & cur_KEY(3) & "' AND CP04='" & cur_KEY(4) & "' AND CP14||CP29 is not null ORDER BY cp05 desc,cp09 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         stRcvEmp = "" & RsTemp.Fields("cp14")
         '2011/8/15 ADD BY SONIA 法務顧問案無承辦人則抓協辦人員
         If cur_KEY(1) = "L" Or cur_KEY(1) = "LA" Then
            If stRcvEmp = "" Then stRcvEmp = "" & RsTemp.Fields("cp29")
         End If
         '2011/8/15 END
      End If
      'Added by Lydia 2023/09/08 案源案件以介紹人所別為該所也可操作
      stLos04 = ""
      If InStr(cur_KEY(1), "L") > 0 Then
         strSql = "select los04 from lawofficesource where los15 in (select max(cp162) from caseprogress where cp01='" & cur_KEY(1) & "' and cp02='" & cur_KEY(2) & "' and cp03='" & cur_KEY(3) & "' and cp04='" & cur_KEY(4) & "' and cp162 is not null)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            stLos04 = "" & RsTemp.Fields("los04")
            If InStr(stLos04, ",") > 0 Then
               stLos04 = Mid(stLos04, 1, InStr(stLos04, ",") - 1)
            End If
         End If
      End If
      'end 2023/09/08
      
      '最後收文智權人員或承辦人為該所人員才可查詢
      'Modified by Lydia 2023/09/08 +案源案件之第一介紹人stLos04
      'If (PUB_GetST06(stKeySales) <> pub_strUserOffice And _
         PUB_GetST06(stRcvEmp) <> pub_strUserOffice) And _
         Pub_StrUserSt03 <> "M51" Then
      If Pub_StrUserSt03 <> "M51" And InStr(PUB_GetST06(stKeySales) & "," & PUB_GetST06(stRcvEmp) & "," & PUB_GetST06(stLos04), pub_strUserOffice) = 0 Then
         MsgBox "此案與你的所別不同！", vbCritical
      Else
         lst_KEY(1) = cur_KEY(1)
         lst_KEY(2) = cur_KEY(2)
         lst_KEY(3) = cur_KEY(3)
         lst_KEY(4) = cur_KEY(4)
         If ReQuery() = True Then doQuery = True
      End If
   Else
      MsgBox "無此記錄之資料！", vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   Exit Function
   
ErrHand:
   MsgBox Err.Description, vbCritical
End Function

'完整資料查詢
Private Function ReQuery() As Boolean
Dim strSql As String, rsQuery As New ADODB.Recordset, intI As Integer
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   
   ReQuery = False
   Select Case cur_KEY(1)
      '專利
      Case "P", "CFP", "FCP"
         strSql = "select PA01 K01,PA02 K02,PA03 K03,PA04 K04" & _
            ", PA47 E00, PA05 D01, PA06 D02, PA07 D03, NA03 D04" & _
            ", DECODE(PA10,NULL,NULL,(SUBSTRB(PA10,1,4)-1911)||'/'||SUBSTR(PA10,5,2)||'/'||SUBSTR(PA10,7,2)) D05" & _
            ", PA11 D06" & _
            ", NVL(NVL(C1.CU04,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90),C1.CU06) D07" & _
            ", NVL(NVL(C2.CU04,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90),C2.CU06) D08" & _
            ", NVL(NVL(C3.CU04,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90),C3.CU06) D09" & _
            ", NVL(NVL(C4.CU04,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90),C4.CU06) D10" & _
            ", NVL(NVL(C5.CU04,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90),C5.CU06) D11" & _
            ", S1.ST02 D12" & _
            ", DECODE(PA96,NULL,NULL,(SUBSTRB(PA96,1,4)-1911)||'/'||SUBSTR(PA96,5,2)||'/'||SUBSTR(PA96,7,2)) D13" & _
            ", FLOOR(PA97/100)||':'||MOD(PA97,100) D14" & _
            ", PA48 D00, C1.CU13 R01,S2.ST04 R02,PA149 R03,C1.CU01 R04" & _
            " from PATENT, NATION, CUSTOMER C1, CUSTOMER C2, CUSTOMER C3, CUSTOMER C4, CUSTOMER C5, STAFF S1, STAFF S2" & _
            " WHERE PA01='" & cur_KEY(1) & "' AND PA02='" & cur_KEY(2) & "' AND PA03='" & cur_KEY(3) & "' AND PA04='" & cur_KEY(4) & "'" & _
            " AND C1.CU01(+)=SUBSTR(PA26,1,8) AND C1.CU02(+)=SUBSTR(PA26,9,1)" & _
            " AND C2.CU01(+)=SUBSTR(PA27,1,8) AND C2.CU02(+)=SUBSTR(PA27,9,1)" & _
            " AND C3.CU01(+)=SUBSTR(PA28,1,8) AND C3.CU02(+)=SUBSTR(PA28,9,1)" & _
            " AND C4.CU01(+)=SUBSTR(PA29,1,8) AND C4.CU02(+)=SUBSTR(PA29,9,1)" & _
            " AND C5.CU01(+)=SUBSTR(PA30,1,8) AND C5.CU02(+)=SUBSTR(PA30,9,1)" & _
            " AND NA01(+) = PA09 AND S1.ST01(+)=PA95 AND S2.ST01(+)=C1.CU13"
      '商標
      Case "T", "CFT", "FCT", "TF"
         strSql = "select TM01 K01,TM02 K02,TM03 K03,TM04 K04" & _
            ",TM34 E00, TM05 D01, TM06 D02, TM07 D03, NA03 D04" & _
            ", DECODE(TM11,NULL,NULL,(SUBSTRB(TM11,1,4)-1911)||'/'||SUBSTR(TM11,5,2)||'/'||SUBSTR(TM11,7,2)) D05" & _
            ", TM12 D06" & _
            ", NVL(NVL(C1.CU04,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90),C1.CU06) D07" & _
            ", NVL(NVL(C2.CU04,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90),C2.CU06) D08" & _
            ", NVL(NVL(C3.CU04,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90),C3.CU06) D09" & _
            ", NVL(NVL(C4.CU04,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90),C4.CU06) D10" & _
            ", NVL(NVL(C5.CU04,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90),C5.CU06) D11" & _
            ", S1.ST02 D12" & _
            ", DECODE(TM63,NULL,NULL,(SUBSTRB(TM63,1,4)-1911)||'/'||SUBSTR(TM63,5,2)||'/'||SUBSTR(TM63,7,2)) D13" & _
            ", FLOOR(TM64/100)||':'||MOD(TM64,100) D14" & _
            ", TM35 D00, C1.CU13 R01,S2.ST04 R02,TM123 R03,C1.CU01 R04" & _
            " from TRADEMARK, NATION, CUSTOMER C1, CUSTOMER C2, CUSTOMER C3, CUSTOMER C4, CUSTOMER C5, STAFF S1, STAFF S2" & _
            " WHERE TM01='" & cur_KEY(1) & "' AND TM02='" & cur_KEY(2) & "' AND TM03='" & cur_KEY(3) & "' AND TM04='" & cur_KEY(4) & "'" & _
            " AND C1.CU01(+)=SUBSTR(TM23,1,8) AND C1.CU02(+)=SUBSTR(TM23,9,1)" & _
            " AND C2.CU01(+)=SUBSTR(TM78,1,8) AND C2.CU02(+)=SUBSTR(TM78,9,1)" & _
            " AND C3.CU01(+)=SUBSTR(TM79,1,8) AND C3.CU02(+)=SUBSTR(TM79,9,1)" & _
            " AND C4.CU01(+)=SUBSTR(TM80,1,8) AND C4.CU02(+)=SUBSTR(TM80,9,1)" & _
            " AND C5.CU01(+)=SUBSTR(TM81,1,8) AND C5.CU02(+)=SUBSTR(TM81,9,1)" & _
            " AND NA01(+) = TM10 AND S1.ST01(+)=TM62 AND S2.ST01(+)=C1.CU13"
      '法務
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      Case "CFL", "FCL", "L", "LIN"
         'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46
         strSql = "select LC01 K01,LC02 K02,LC03 K03,LC04 K04" & _
            ", LC16 E00, LC05 D01, LC06 D02, LC07 D03, NA03 D04" & _
            ", '' D05, '' D06" & _
            ", NVL(NVL(C1.CU04,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90),C1.CU06) D07" & _
            ", NVL(NVL(C2.CU04,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90),C2.CU06) D08" & _
            ", NVL(NVL(C3.CU04,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90),C3.CU06) D09" & _
            ", NVL(NVL(C4.CU04,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90),C4.CU06) D10" & _
            ", NVL(NVL(C5.CU04,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90),C5.CU06) D11" & _
            ", S1.ST02 D12" & _
            ", DECODE(LC32,NULL,NULL,(SUBSTRB(LC32,1,4)-1911)||'/'||SUBSTR(LC32,5,2)||'/'||SUBSTR(LC32,7,2)) D13" & _
            ", FLOOR(LC33/100)||':'||MOD(LC33,100) D14" & _
            ", LC17 D00, C1.CU13 R01,S2.ST04 R02,LC42 R03,C1.CU01 R04" & _
            " from LAWCASE, NATION, CUSTOMER C1, CUSTOMER C2, CUSTOMER C3, CUSTOMER C4, CUSTOMER C5, STAFF S1, STAFF S2" & _
            " WHERE LC01='" & cur_KEY(1) & "' AND LC02='" & cur_KEY(2) & "' AND LC03='" & cur_KEY(3) & "' AND LC04='" & cur_KEY(4) & "'" & _
            " AND C1.CU01(+)=SUBSTR(LC11,1,8) AND C1.CU02(+)=SUBSTR(LC11,9,1)" & _
            " AND C2.CU01(+)=SUBSTR(LC43,1,8) AND C2.CU02(+)=SUBSTR(LC43,9,1)" & _
            " AND C3.CU01(+)=SUBSTR(LC44,1,8) AND C3.CU02(+)=SUBSTR(LC44,9,1)" & _
            " AND C4.CU01(+)=SUBSTR(LC45,1,8) AND C4.CU02(+)=SUBSTR(LC45,9,1)" & _
            " AND C5.CU01(+)=SUBSTR(LC46,1,8) AND C5.CU02(+)=SUBSTR(LC46,9,1)" & _
            " AND NA01(+) = LC15 AND S1.ST01(+)=LC31 AND S2.ST01(+)=C1.CU13"
      '顧問
      Case "LA"
         'Modify By Sindy 2011/2/18 增加HC24,HC25,HC26,HC27
         strSql = "select HC01 K01,HC02 K02,HC03 K03,HC04 K04" & _
            ", HC07 E00, HC06 D01, '' D02, '' D03, '' D04" & _
            ", '' D05, '' D06" & _
            ", NVL(NVL(C1.CU04,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90),C1.CU06) D07" & _
            ", NVL(NVL(C2.CU04,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90),C2.CU06) D08" & _
            ", NVL(NVL(C3.CU04,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90),C3.CU06) D09" & _
            ", NVL(NVL(C4.CU04,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90),C4.CU06) D10" & _
            ", NVL(NVL(C5.CU04,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90),C5.CU06) D11" & _
            ", S1.ST02 D12" & _
            ", DECODE(HC17,NULL,NULL,(SUBSTRB(HC17,1,4)-1911)||'/'||SUBSTR(HC17,5,2)||'/'||SUBSTR(HC17,7,2)) D13" & _
            ", FLOOR(HC18/100)||':'||MOD(HC18,100) D14" & _
            ", '' D00, C1.CU13 R01,S2.ST04 R02,HC23 R03,C1.CU01 R04" & _
            " from HIRECASE, CUSTOMER C1, CUSTOMER C2, CUSTOMER C3, CUSTOMER C4, CUSTOMER C5, STAFF S1, STAFF S2" & _
            " WHERE HC01='" & cur_KEY(1) & "' AND HC02='" & cur_KEY(2) & "' AND HC03='" & cur_KEY(3) & "' AND HC04='" & cur_KEY(4) & "'" & _
            " AND C1.CU01(+)=SUBSTR(HC05,1,8) AND C1.CU02(+)=SUBSTR(HC05,9,1)" & _
            " AND C2.CU01(+)=SUBSTR(HC24,1,8) AND C2.CU02(+)=SUBSTR(HC24,9,1)" & _
            " AND C3.CU01(+)=SUBSTR(HC25,1,8) AND C3.CU02(+)=SUBSTR(HC25,9,1)" & _
            " AND C4.CU01(+)=SUBSTR(HC26,1,8) AND C4.CU02(+)=SUBSTR(HC26,9,1)" & _
            " AND C5.CU01(+)=SUBSTR(HC27,1,8) AND C5.CU02(+)=SUBSTR(HC27,9,1)" & _
            " AND S1.ST01(+)=HC16 AND S2.ST01(+)=C1.CU13"
      '服務
      '"CFC","CPS","FG","PS","S","TB","TC","TD","TM","TR","TS","TT"
      Case Else
         strSql = "select SP01 K01,SP02 K02,SP03 K03,SP04 K04" & _
            ", SP28 E00, SP05 D01, SP06 D02, SP07 D03, NA03 D04" & _
            ", DECODE(SP10,NULL,NULL,(SUBSTRB(SP10,1,4)-1911)||'/'||SUBSTR(SP10,5,2)||'/'||SUBSTR(SP10,7,2)) D05" & _
            ", SP11 D06" & _
            ", NVL(NVL(C1.CU04,C1.cu05||' '||C1.cu88||' '||C1.cu89||' '||C1.cu90),C1.CU06) D07" & _
            ", NVL(NVL(C2.CU04,C2.cu05||' '||C2.cu88||' '||C2.cu89||' '||C2.cu90),C2.CU06) D08" & _
            ", NVL(NVL(C3.CU04,C3.cu05||' '||C3.cu88||' '||C3.cu89||' '||C3.cu90),C3.CU06) D09" & _
            ", NVL(NVL(C4.CU04,C4.cu05||' '||C4.cu88||' '||C4.cu89||' '||C4.cu90),C4.CU06) D10" & _
            ", NVL(NVL(C5.CU04,C5.cu05||' '||C5.cu88||' '||C5.cu89||' '||C5.cu90),C5.CU06) D11" & _
            ", S1.ST02 D12" & _
            ", DECODE(SP56,NULL,NULL,(SUBSTRB(SP56,1,4)-1911)||'/'||SUBSTR(SP56,5,2)||'/'||SUBSTR(SP56,7,2)) D13" & _
            ", FLOOR(SP57/100)||':'||MOD(SP57,100) D14" & _
            ", SP29 D00, C1.CU13 R01,S2.ST04 R02,SP78 R03,C1.CU01 R04" & _
            " from SERVICEPRACTICE, NATION, CUSTOMER C1, CUSTOMER C2, CUSTOMER C3, CUSTOMER C4, CUSTOMER C5, STAFF S1, STAFF S2" & _
            " WHERE SP01='" & cur_KEY(1) & "' AND SP02='" & cur_KEY(2) & "' AND SP03='" & cur_KEY(3) & "' AND SP04='" & cur_KEY(4) & "'" & _
            " AND C1.CU01(+)=SUBSTR(SP08,1,8) AND C1.CU02(+)=SUBSTR(SP08,9,1)" & _
            " AND C2.CU01(+)=SUBSTR(SP58,1,8) AND C2.CU02(+)=SUBSTR(SP58,9,1)" & _
            " AND C3.CU01(+)=SUBSTR(SP59,1,8) AND C3.CU02(+)=SUBSTR(SP59,9,1)" & _
            " AND C4.CU01(+)=SUBSTR(SP65,1,8) AND C4.CU02(+)=SUBSTR(SP65,9,1)" & _
            " AND C5.CU01(+)=SUBSTR(SP66,1,8) AND C5.CU02(+)=SUBSTR(SP66,9,1)" & _
            " AND NA01(+) = SP09 AND S1.ST01(+)=SP55 AND S2.ST01(+)=C1.CU13"
   End Select

   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      txtKey(1) = "" & rsQuery.Fields("K01")
      txtKey(2) = "" & rsQuery.Fields("K02")
      txtKey(3) = "" & rsQuery.Fields("K03")
      txtKey(4) = "" & rsQuery.Fields("K04")
      For intI = 0 To 0
         txtEdit(intI) = "" & rsQuery.Fields("E" & Format(intI, "00"))
      Next intI
      For intI = 1 To 14
         lblDisp(intI) = "" & rsQuery.Fields("D" & Format(intI, "00"))
      Next intI
      ReQuery = True
   Else
      MsgBox "案件〔" & cur_KEY(1) & cur_KEY(2) & cur_KEY(3) & cur_KEY(4) & "〕已被刪除！", vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   MsgBox Err.Description, vbCritical
   Screen.MousePointer = vbDefault
End Function

Private Sub txtkey_GotFocus(Index As Integer)
   TextInverse txtKey(Index)
   If txtKey(Index).Enabled = True Then
      CloseIme
   End If
End Sub

Private Sub txtKey_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
