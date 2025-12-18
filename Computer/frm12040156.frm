VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040156 
   BorderStyle     =   1  '單線固定
   Caption         =   "LEDES基本資料維護"
   ClientHeight    =   5484
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5484
   ScaleWidth      =   9120
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8505
      Top             =   300
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
            Picture         =   "frm12040156.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040156.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040156.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040156.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040156.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040156.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040156.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040156.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040156.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040156.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040156.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
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
   Begin MSForms.TextBox Text1 
      Height          =   675
      Index           =   18
      Left            =   810
      TabIndex        =   37
      Top             =   4710
      Width           =   7995
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "14102;1191"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   6435
      TabIndex        =   34
      Top             =   4380
      Width           =   2355
      VariousPropertyBits=   671105051
      MaxLength       =   100
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   2370
      TabIndex        =   16
      Top             =   4380
      Width           =   945
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "1667;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   285
      Left            =   2310
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   750
      Width           =   6540
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1170
      TabIndex        =   0
      Top             =   750
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   2070
      TabIndex        =   15
      Top             =   4080
      Width           =   6750
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "11906;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   7470
      TabIndex        =   13
      Top             =   3480
      Width           =   1170
      VariousPropertyBits=   671105051
      MaxLength       =   10
      Size            =   "2064;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   2070
      TabIndex        =   14
      Top             =   3780
      Width           =   4365
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "7699;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   2070
      TabIndex        =   12
      Top             =   3480
      Width           =   2385
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4207;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   6465
      TabIndex        =   11
      Top             =   3180
      Width           =   810
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1429;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   2070
      TabIndex        =   10
      Top             =   3180
      Width           =   2385
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4207;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   2475
      TabIndex        =   9
      Top             =   2880
      Width           =   6345
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "11192;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   2070
      TabIndex        =   8
      Top             =   2580
      Width           =   6750
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "11906;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   2070
      TabIndex        =   7
      Top             =   2280
      Width           =   6750
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "11906;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   2070
      TabIndex        =   6
      Top             =   1980
      Width           =   6750
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "11906;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   2070
      TabIndex        =   5
      Top             =   1680
      Width           =   6750
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "11906;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   6435
      TabIndex        =   4
      Top             =   1380
      Width           =   2385
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4207;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2070
      TabIndex        =   3
      Top             =   1380
      Width           =   2385
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4207;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "備註："
      Height          =   180
      Index           =   17
      Left            =   240
      TabIndex        =   36
      Top             =   4740
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PO_NUMBER："
      Height          =   180
      Index           =   16
      Left            =   4725
      TabIndex        =   35
      Top             =   4432
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LINE_ITEM_UNIT_COST："
      Height          =   180
      Index           =   15
      Left            =   225
      TabIndex        =   33
      Top             =   4432
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "帳單格式：            ( 1. 1998B    2. 1998BI )"
      Height          =   180
      Index           =   14
      Left            =   225
      TabIndex        =   31
      Top             =   1125
      Width           =   3210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LAW_FIRM_ID："
      Height          =   180
      Index           =   13
      Left            =   225
      TabIndex        =   30
      Top             =   4125
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TIMEKEEPER_CLASSIFICATION："
      Height          =   180
      Index           =   12
      Left            =   4725
      TabIndex        =   29
      Top             =   3532
      Width           =   2715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TIMEKEEPER_NAME："
      Height          =   180
      Index           =   11
      Left            =   225
      TabIndex        =   28
      Top             =   3825
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TIMEKEEPER_ID："
      Height          =   180
      Index           =   10
      Left            =   225
      TabIndex        =   27
      Top             =   3532
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CLIENT_COUNTRY："
      Height          =   180
      Index           =   9
      Left            =   4725
      TabIndex        =   26
      Top             =   3232
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CLIENT_POSTCODE："
      Height          =   180
      Index           =   8
      Left            =   225
      TabIndex        =   25
      Top             =   3232
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CLIENT_STATEorREGION："
      Height          =   180
      Index           =   7
      Left            =   225
      TabIndex        =   24
      Top             =   2925
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CLIENT_CITY："
      Height          =   180
      Index           =   6
      Left            =   225
      TabIndex        =   23
      Top             =   2625
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CLIENT_ADDRESS_2："
      Height          =   180
      Index           =   5
      Left            =   225
      TabIndex        =   22
      Top             =   2325
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CLIENT_ADDRESS_1："
      Height          =   180
      Index           =   4
      Left            =   225
      TabIndex        =   21
      Top             =   2025
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CLIENT_NAME："
      Height          =   180
      Index           =   3
      Left            =   225
      TabIndex        =   20
      Top             =   1725
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CLIENT_TAX_ID："
      Height          =   180
      Index           =   2
      Left            =   4725
      TabIndex        =   19
      Top             =   1425
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CLIENT_ID："
      Height          =   180
      Index           =   1
      Left            =   225
      TabIndex        =   18
      Top             =   1425
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "請款對象："
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   17
      Top             =   810
      Width           =   900
   End
End
Attribute VB_Name = "frm12040156"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/29 Form2.0已修改
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Create by Morgan 2012/4/20
Option Explicit

Dim ActionEdit As Integer '0:新增 1:修改 3:瀏覽
'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim oText As Object
   
   
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Screen.MousePointer = vbHourglass
   Select Case KeyCode
      Case vbKeyF2: Action 1 '新增
      Case vbKeyF3: Action 2 '修改
      Case vbKeyF5: Action 3 '刪除
      Case vbKeyF4: Action 4 '查詢
      Case vbKeyHome: Action 6 '第一筆
      Case vbKeyPageUp: Action 7 '前一筆
      Case vbKeyPageDown: Action 8 '後一筆
      Case vbKeyEnd: Action 9 '最後筆
      Case vbKeyF9, vbKeyReturn: Action 11 '確定
      Case vbKeyF10: Action 12 '取消
      Case vbKeyEscape: Action 14 '結束
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction("frm12040156", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040156", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040156", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040156", strFind, False)
   
   MoveFormToCenter Me
   
   ActionEdit = 3
   Action 6 '預設第一筆
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040156 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Screen.MousePointer = vbHourglass
   Action Button.Index
   Screen.MousePointer = vbDefault
End Sub

Private Sub Action(Index As Integer)
   
   If TBar1.Buttons(Index).Enabled = False Then Exit Sub

On Error GoTo ErrHand
   Select Case Index
      Case 1 '按下新增
        ActionEdit = 0
        FormReset
        
      Case 2 '按下修改
         ActionEdit = 1
         
      Case 3 '按下刪除
         If Text1(1).Text = "" Then
             MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
             Exit Sub
         End If
         
         If DelMsg Then
            If FormDelete() = False Then
               MsgBox "刪除失敗!", vbCritical
               Exit Sub
            '刪除後移到最末筆
            Else
               RsAction 3
            End If
         End If
         
      Case 4 '按下查詢
         FormReset
         ActionEdit = 2
         
      Case 6 '第一筆
         RsAction 0
      Case 7 '前一筆
         RsAction 1
      Case 8 '後一筆
         RsAction 2
      Case 9 '最後筆
         RsAction 3
      Case 11 '按下確定
         Select Case ActionEdit
            Case 0, 1 '新增,修改
               If TxtValidate = False Then
                  Exit Sub
               Else
                  'Add by Sindy 2021/11/29 檢查畫面上的物件是否含有Unicode文字
                  If PUB_ChkUniText(Me, True, True) = False Then
                     Exit Sub
                  End If
                  
                  If FormSave() = False Then
                     MsgBox "存檔失敗!", vbCritical
                     Exit Sub
                  End If
               End If
            Case 2
               If ReadData(Text1(1)) = False Then
                  MsgBox "無資料!", vbExclamation
                  Exit Sub
               End If
         End Select
         ActionEdit = 3
         
      Case 12 '按下取消
         Text1(1) = Text1(1).Tag
         If Text1(1) <> "" Then
            If ReadData(Text1(1)) = True Then
               ActionEdit = 3
            End If
         End If
         
      Case 14 '結束
         Unload Me
         Exit Sub
   End Select
   
   CmdSitu
   TxtLock
   Exit Sub
   
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

Private Function FormSave() As Boolean
   Dim stSQL As String, bInTrans As Boolean
   Dim stCols As String, stValues As String
   
On Error GoTo ErrHandle
   
   cnnConnection.BeginTrans
   bInTrans = True
   
   If ActionEdit = 0 Then
      stCols = "": stValues = ""
      For Each oText In Text1
         stCols = stCols & IIf(stCols <> "", ",", "") & "LD" & Format(oText.Index, "00")
         stValues = stValues & IIf(stValues <> "", ",", "") & "'" & ChgSQL(oText.Text) & "'"
      Next
      stSQL = "insert into LEDES(" & stCols & ") values (" & stValues & ")"
   Else
      stValues = ""
      For Each oText In Text1
         stValues = stValues & IIf(stValues <> "", ",", "") & "LD" & Format(oText.Index, "00") & "='" & ChgSQL(oText.Text) & "'"
      Next
      stSQL = "update LEDES set " & stValues & " where LD01='" & Text1(1) & "'"
   End If
   cnnConnection.Execute stSQL, intI
   cnnConnection.CommitTrans
   FormSave = True
   
ErrHandle:
   If Err.Number <> 0 Then
      If bInTrans Then cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function


Private Sub CmdSitu()
   Dim bolTF As Boolean
   Dim ii As Integer
   
   If ActionEdit = 3 Then
      bolTF = True
   Else
      bolTF = False
   End If
   
   For ii = 1 To 4
      TBar1.Buttons(ii).Enabled = bolTF
      If Text1(1).Tag <> "" Then
         TBar1.Buttons(ii + 5).Enabled = bolTF
      Else
         TBar1.Buttons(ii + 5).Enabled = False
      End If
   Next
   TBar1.Buttons(11).Enabled = Not bolTF
   TBar1.Buttons(12).Enabled = Not bolTF
   TBar1.Buttons(14).Enabled = bolTF
   
End Sub

Private Sub TxtLock()
   Select Case ActionEdit
   Case 0 '新增
      For Each oText In Text1
         oText.Locked = False
      Next
      Text1(1).SetFocus
      Text1_GotFocus 1
      
   Case 1 '修改
      For Each oText In Text1
         oText.Locked = False
      Next
      Text1(1).Locked = True
      Text1(2).SetFocus
      Text1_GotFocus 2
      
   Case 2 '查詢
      Text1(1).Locked = False
      Text1(1).SetFocus
      Text1_GotFocus 1
      
   Case 3 '按下取消後的狀態
      For Each oText In Text1
         oText.Locked = True
      Next
   End Select
End Sub


Private Sub RsAction(ByVal Sty As Integer)
 Dim stKEY As String
    
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case Sty
      Case 0 '第一筆
         strExc(0) = "SELECT nvl(min(ld01),0) FROM LEDES"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 1 '前一筆
         strExc(0) = "SELECT nvl(max(LD01),'" & Text1(1) & "') FROM LEDES where LD01<'" & Text1(1) & "'"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = Text1(1) Then
               DataErrorMessage 6
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 2 '後一筆
         strExc(0) = "SELECT nvl(min(ld01),'" & Text1(1) & "') FROM LEDES where ld01>'" & Text1(1) & "'"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = Text1(1) Then
               DataErrorMessage 7
            Else
               stKEY = RsTemp.Fields(0)
            End If
         End If
         
      Case 3 '最後筆
         strExc(0) = "SELECT nvl(max(ld01),0) FROM LEDES"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
               stKEY = RsTemp.Fields(0)
            End If
         End If
   End Select
   
   
   If stKEY <> "" Then ReadData stKEY
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

Private Function ReadData(ByVal pKey As String) As Boolean
   FormReset
   pKey = Left(pKey & "000", 9)
   strExc(0) = "select * from LEDES where ld01='" & pKey & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      For Each oText In Text1
         oText = "" & .Fields("LD" & Format(oText.Index, "00"))
      Next
      End With
      Text1(1).Tag = Text1(1)
      Text2 = GetClientName(Text1(1))
      ReadData = True
   End If
End Function

Private Function FormDelete() As Boolean
   Dim stSQL As String, bInTrans As Boolean
   
On Error GoTo ErrHandle

   cnnConnection.BeginTrans
   bInTrans = True

   stSQL = "delete LEDES where ld01='" & Text1(1) & "'"
   cnnConnection.Execute stSQL, intI
   
   cnnConnection.CommitTrans
   FormDelete = True
   
ErrHandle:
   If Err.Number <> 0 Then
      If bInTrans Then cnnConnection.RollbackTrans
      MsgBox Err.Description
   End If
End Function

Private Sub FormReset()
   For Each oText In Text1
      If Not (ActionEdit = 2 And oText.Index = 1) Then
         oText.Text = ""
      End If
   Next
   Text2 = ""
End Sub

Private Function TxtValidate() As Boolean
   If Text1(1) = "" Then
      MsgBox "請款對象不可空白!!"
      Text1(1).SetFocus
      Exit Function
   End If
   
   If Text1(15) = "" Then
      MsgBox "帳單格式不可空白!!"
      Text1(1).SetFocus
      Exit Function
   End If
   TxtValidate = True
End Function

Private Sub Text1_Change(Index As Integer)
   If Index = 1 Then
      Text2 = ""
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index = 1 Then
      KeyAscii = UpperCase(KeyAscii)
   End If
   If KeyAscii <> 8 Then
      If Index = 15 Then
         If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
            KeyAscii = 0
            Beep
         End If
      ElseIf Index = 16 Then
         If KeyAscii <> Asc(".") And IsNumeric(Chr(KeyAscii)) = False Then
            KeyAscii = 0
            Beep
         End If
      End If
   End If
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Index = 1 Then
      If Text1(1) <> "" Then
         Text2 = GetClientName(Text1(1))
      End If
   End If
End Sub

Private Function GetClientName(pNo) As String
   If Left(pNo, 1) = "X" Then
      strExc(0) = "select nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),nvl(cu04,cu06)) from customer WHERE cu01='" & Left(pNo, 8) & "' and cu02='" & Mid(pNo, 9) & "'"
   Else
      strExc(0) = "select nvl(rtrim(fa05||' '||fa63||' '||fa64||' '||fa65),nvl(fa04,fa06)) from fagent WHERE fa01='" & Left(pNo, 8) & "' and fa02='" & Mid(pNo, 9) & "'"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetClientName = "" & RsTemp(0)
   End If
End Function
