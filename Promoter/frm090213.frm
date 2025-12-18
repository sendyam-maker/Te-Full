VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090213 
   BorderStyle     =   1  '單線固定
   Caption         =   "期刊資料維護"
   ClientHeight    =   3630
   ClientLeft      =   405
   ClientTop       =   1770
   ClientWidth     =   7845
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7845
   Tag             =   "期刊資料維護"
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7230
      Top             =   570
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
            Picture         =   "frm090213.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090213.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090213.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090213.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090213.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090213.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090213.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090213.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090213.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090213.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090213.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7845
      _ExtentX        =   13838
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
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1050
      TabIndex        =   2
      Top             =   1590
      Width           =   2415
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4260;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   6
      Left            =   1050
      TabIndex        =   6
      Top             =   3225
      Width           =   855
      VariousPropertyBits=   671107099
      MaxLength       =   2
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   1050
      TabIndex        =   5
      Top             =   2730
      Width           =   6735
      VariousPropertyBits=   671107099
      MaxLength       =   40
      Size            =   "11880;661"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   4
      Left            =   1050
      TabIndex        =   4
      Top             =   2325
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   1050
      TabIndex        =   3
      Top             =   1950
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   6
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   1050
      TabIndex        =   1
      Top             =   1215
      Width           =   2415
      VariousPropertyBits=   671107099
      MaxLength       =   20
      Size            =   "4260;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1050
      TabIndex        =   0
      Top             =   870
      Width           =   5655
      VariousPropertyBits=   671107099
      MaxLength       =   60
      Size            =   "9975;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label8 
      Height          =   255
      Left            =   1950
      TabIndex        =   15
      Top             =   3240
      Width           =   5025
      VariousPropertyBits=   27
      Size            =   "8864;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "索　　引："
      Height          =   180
      Left            =   90
      TabIndex        =   14
      Top             =   3270
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "備　　註："
      Height          =   180
      Left            =   90
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "出版日期：                               (中國專利報用西元年,其他用民國年)"
      Height          =   180
      Left            =   90
      TabIndex        =   12
      Top             =   2370
      Width           =   5415
   End
   Begin VB.Label Label4 
      Caption         =   "版／頁　："
      Height          =   180
      Left            =   90
      TabIndex        =   11
      Top             =   2010
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "資料出處："
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   1605
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "作　　者："
      Height          =   180
      Left            =   105
      TabIndex        =   9
      Top             =   1245
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "標　　題："
      Height          =   180
      Left            =   105
      TabIndex        =   8
      Top             =   930
      Width           =   975
   End
End
Attribute VB_Name = "frm090213"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (Text1,Combo1,Label8)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset, p As New ADODB.Recordset
Dim EDITSELECT As Integer, i As Integer
Dim a As String, str As String, s As Integer, CheckData As Boolean

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

'Add By Cheng 2002/03/01
Dim m_blnAddNewOK As Boolean '新增成功或失敗

Private Sub Combo1_GotFocus()
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
End Sub

Private Sub Combo1_LostFocus()
'If Combo1.Text <> "中國專利報" And Combo1.Text <> "智慧財產權" Then
'    MsgBox ("只能選擇 中國專利報 OR 智慧財產權 !!")
'    Combo1.SetFocus
'    Exit Sub
'End If
End Sub

Private Sub Form_Activate()
If pemain.State = adStateOpen Then pemain.Close
strExc(0) = "SELECT PE01,PE02,PE03,PE04,decode(pe03,'中國專利報',PE05,DECODE(PE05,'','',(SUBSTR(PE05,1,4)-1911)||(SUBSTR(PE05,5,2))||(SUBSTR(PE05,7,2)))),PE06,PE07,pe01||PE03||PE04||decode(pe03,'中國專利報',PE05,DECODE(PE05,'','',(SUBSTR(PE05,1,4)-1911)||(SUBSTR(PE05,5,2))||(SUBSTR(PE05,7,2)))) AS AD FROM PERIODICAL ORDER BY PE03,PE04,decode(pe03,'中國專利報',PE05,DECODE(PE05,'','',(SUBSTR(PE05,1,4)-1911)||(SUBSTR(PE05,5,2))||(SUBSTR(PE05,7,2)))),PE06 "
pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
If pemain.EOF And pemain.BOF Then Exit Sub
'Add By Cheng 2002/03/04
If Len("" & frm090215_2.m_strPE01) > 0 Then
   Do While Not pemain.EOF
      If pemain.Fields(0).Value = frm090215_2.m_strPE01 And _
         pemain.Fields(2).Value = frm090215_2.m_strPE03 And _
         pemain.Fields(3).Value = frm090215_2.m_strPE04 And _
         pemain.Fields(4).Value = frm090215_2.m_strPE05 Then Exit Do
      pemain.MoveNext
   Loop
End If
For i = 0 To 6
If i = 2 Then
    If IsNull(pemain.Fields(i).Value) Then
        Combo1.Text = ""
    Else
        Combo1.Text = pemain.Fields(i).Value
    End If
Else
    If IsNull(pemain.Fields(i).Value) Then
        Text1(i) = ""
    Else
        Text1(i) = pemain.Fields(i).Value
    End If
End If
Next i
For i = 1 To 4
    TBar1.Buttons(i).Enabled = True
Next i
For i = 6 To 9
    TBar1.Buttons(i).Enabled = True
Next i
TBar1.Buttons(11).Enabled = False
TBar1.Buttons(12).Enabled = False
TBar1.Buttons(14).Enabled = True
locktext (1)
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyF2
         If TBar1.Buttons(1).Enabled = True Then
            EDITTOOL (1)
            Text1(0).SetFocus
         End If
       Case vbKeyF3
         If TBar1.Buttons(2).Enabled = True Then
            EDITTOOL (2)
         End If
       Case vbKeyF5
         If TBar1.Buttons(3).Enabled = True Then
            EDITTOOL (3)
         End If
       Case vbKeyF4
         If TBar1.Buttons(4).Enabled = True Then
            EDITTOOL (4)
            Text1(0).SetFocus
         End If
       Case vbKeyF9, vbKeyReturn
         If TBar1.Buttons(11).Enabled = True Then
            EDITTOOL (9)
         End If
       Case vbKeyHome
         If TBar1.Buttons(6).Enabled = True Then
            EDITTOOL (5)
         End If
       Case vbKeyEnd
         If TBar1.Buttons(9).Enabled = True Then
            EDITTOOL (8)
         End If
       Case vbKeyPageUp
         If TBar1.Buttons(7).Enabled = True Then
            EDITTOOL (6)
         End If
       Case vbKeyPageDown
          If TBar1.Buttons(8).Enabled = True Then
            EDITTOOL (7)
          End If
       Case vbKeyF10
          If TBar1.Buttons(12).Enabled = True Then
            EDITTOOL (10)
          End If
       Case vbKeyEscape
         If TBar1.Buttons(14).Enabled = True Then
            EDITTOOL (11)
         End If
End Select
   If KeyCode <> vbKeyF2 And KeyCode <> vbKeyF3 And KeyCode <> vbKeyF4 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyEscape Then
      If EDITSELECT > 4 Then
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
      End If
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'edit by nickc 2007/08/27
'Select Case KeyCode
'       Case vbKeyF2
'         If TBar1.Buttons(1).Enabled = True Then
'            EDITTOOL (1)
'            Text1(0).SetFocus
'         End If
'       Case vbKeyF3
'         If TBar1.Buttons(2).Enabled = True Then
'            EDITTOOL (2)
'         End If
'       Case vbKeyF5
'         If TBar1.Buttons(3).Enabled = True Then
'            EDITTOOL (3)
'         End If
'       Case vbKeyF4
'         If TBar1.Buttons(4).Enabled = True Then
'            EDITTOOL (4)
'            Text1(0).SetFocus
'         End If
'       Case vbKeyF9, vbKeyReturn
'         If TBar1.Buttons(11).Enabled = True Then
'            EDITTOOL (9)
'         End If
'       Case vbKeyHome
'         If TBar1.Buttons(6).Enabled = True Then
'            EDITTOOL (5)
'         End If
'       Case vbKeyEnd
'         If TBar1.Buttons(9).Enabled = True Then
'            EDITTOOL (8)
'         End If
'       Case vbKeyPageUp
'         If TBar1.Buttons(7).Enabled = True Then
'            EDITTOOL (6)
'         End If
'       Case vbKeyPageDown
'          If TBar1.Buttons(8).Enabled = True Then
'            EDITTOOL (7)
'          End If
'       Case vbKeyF10
'          If TBar1.Buttons(12).Enabled = True Then
'            EDITTOOL (10)
'          End If
'       Case vbKeyEscape
'         If TBar1.Buttons(14).Enabled = True Then
'            EDITTOOL (11)
'         End If
'End Select
'   If KeyCode <> vbKeyF2 And KeyCode <> vbKeyF3 And KeyCode <> vbKeyF4 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyEscape Then
'      If EDITSELECT > 4 Then
'         If m_bInsert Then
'             TBar1.Buttons(1).Enabled = True
'         Else
'             TBar1.Buttons(1).Enabled = False
'         End If
'         If m_bUpdate Then
'             TBar1.Buttons(2).Enabled = True
'         Else
'             TBar1.Buttons(2).Enabled = False
'         End If
'         If m_bDelete Then
'             TBar1.Buttons(3).Enabled = True
'         Else
'             TBar1.Buttons(3).Enabled = False
'         End If
'      End If
'   End If


End Sub

Private Sub Form_Load()
   m_bInsert = IsUserHasRightOfFunction("frm090213", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm090213", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm090213", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm090213", strFind, False)
MoveFormToCenter Me

'Added by Morgan 2022/1/14
Combo1.Clear
Combo1.AddItem "中國專利報"
Combo1.AddItem "智慧財產權"
'end 2022/1/14

pemain.CursorLocation = adUseClient
p.CursorLocation = adUseClient
EDITSELECT = 0
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
'Add By Cheng 2002/03/01
Me.Caption = Me.Tag
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090213 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Add By Cheng 2002/03/01
Me.Caption = Me.Tag
Select Case Button.Index
       Case 1
          EDITTOOL (1)
          Text1(0).SetFocus
          'Add By Cheng 2002/03/01
          Me.Caption = Me.Tag + " － 新增"
       Case 2
          EDITTOOL (2)
          'Add By Cheng 2002/03/01
          Me.Caption = Me.Tag + " － 修改"
       Case 3
          EDITTOOL (3)
          'Add By Cheng 2002/03/01
          Me.Caption = Me.Tag + " － 刪除"
       Case 4
          EDITTOOL (4)
          Text1(0).SetFocus
          'Add By Cheng 2002/03/01
          Me.Caption = Me.Tag + " － 查詢"
       Case 6
          EDITTOOL (5)
       Case 7
          EDITTOOL (6)
       Case 8
          EDITTOOL (7)
       Case 9
          EDITTOOL (8)
       Case 11 '確定
          EDITTOOL (9)
       Case 12
          EDITTOOL (10)
       Case 14
          EDITTOOL (11)
End Select
   If Button.Index <> 14 And Button.Index <> 1 And Button.Index <> 2 And Button.Index <> 3 And Button.Index <> 4 Then
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
   End If

'Add By Cheng 2002/03/01
If Button.Index = 11 And EDITSELECT = 1 And m_blnAddNewOK = True Then
   Tbar1_ButtonClick Me.TBar1.Buttons(1)
   m_blnAddNewOK = False
End If

End Sub

Private Function EDITTOOL(Index As Integer)
Select Case Index
       Case 1 'NEW
          EDITSELECT = 1
          locktext (2)
          For i = 0 To 6
            If i = 2 Then
                Combo1.Text = ""
            Else
                Text1(i).Text = ""
            End If
          Next i
          For i = 1 To 4
            TBar1.Buttons(i).Enabled = False
          Next i
          For i = 6 To 9
            TBar1.Buttons(i).Enabled = False
          Next i
          TBar1.Buttons(11).Enabled = True
          TBar1.Buttons(12).Enabled = True
          TBar1.Buttons(14).Enabled = False
          Label8.Caption = ""
       Case 2 'UPDATA
          EDITSELECT = 2
          locktext (3)
          For i = 1 To 4
             TBar1.Buttons(i).Enabled = False
          Next i
          For i = 6 To 9
            TBar1.Buttons(i).Enabled = False
          Next i
          TBar1.Buttons(11).Enabled = True
          TBar1.Buttons(12).Enabled = True
          TBar1.Buttons(14).Enabled = False
       Case 3 'DELETE
          EDITSELECT = 3
          locktext (1)
          For i = 1 To 4
            TBar1.Buttons(i).Enabled = False
          Next i
          For i = 6 To 9
            TBar1.Buttons(i).Enabled = False
          Next i
          TBar1.Buttons(11).Enabled = True
          TBar1.Buttons(12).Enabled = True
          TBar1.Buttons(14).Enabled = False
          If MsgBox("是否要刪除此筆資料", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
          pemain.MoveNext
          If pemain.EOF Then
            pemain.MoveFirst
            a = CheckStr(pemain.Fields("AD").Value)
          Else
            a = CheckStr(pemain.Fields("AD").Value)
          End If
            If Combo1.Text = "中國專利報" Then
               cnnConnection.Execute "delete periodical where pe01='" & Text1(0) & "' and pE03='" & Combo1.Text & "' AND PE04='" & Text1(3).Text & "' AND PE05=" & Text1(4).Text & " "
            Else
               cnnConnection.Execute "delete periodical where pe01='" & Text1(0) & "' and pE03='" & Combo1.Text & "' AND PE04='" & Text1(3).Text & "' AND PE05=" & ChangeTStringToWString(Text1(4).Text) & " "
            End If
            pemain.ReQuery
            pemain.Find "AD='" & a & "'"
           For i = 0 To 6
            If i = 2 Then
                If IsNull(pemain.Fields(i).Value) Then
                    Combo1.Text = ""
                Else
                    Combo1.Text = pemain.Fields(i).Value
                End If
            Else
                If IsNull(pemain.Fields(i).Value) Then
                    Text1(i) = ""
                Else
                    Text1(i) = pemain.Fields(i).Value
                End If
            End If
            Next i
              End If
                 For i = 1 To 4
                      TBar1.Buttons(i).Enabled = True
                  Next i
                  For i = 6 To 9
                      TBar1.Buttons(i).Enabled = True
                  Next i
                  TBar1.Buttons(11).Enabled = False
                  TBar1.Buttons(12).Enabled = False
                  TBar1.Buttons(14).Enabled = True
       Case 4 'QUTION
            EDITSELECT = 4
            locktext (4)
            For i = 0 To 6
                If i = 2 Then
                   Combo1.Text = ""
                Else
                   Text1(i).Text = ""
                   DoEvents
                End If
            Next i
            For i = 1 To 4
                TBar1.Buttons(i).Enabled = False
            Next i
            For i = 6 To 9
                TBar1.Buttons(i).Enabled = False
            Next i
            TBar1.Buttons(11).Enabled = True
            TBar1.Buttons(12).Enabled = True
            TBar1.Buttons(14).Enabled = False
            Label8.Caption = ""
            Combo1.SetFocus
       Case 5 'FIRST
            EDITSELECT = 5
          If Not (pemain.BOF And pemain.EOF) Then
            If TBar1.Buttons(5).Enabled = True Then
                pemain.MoveFirst
            End If
            For i = 0 To 6
              If i = 2 Then
                  If IsNull(pemain.Fields(i).Value) Then
                      Combo1.Text = ""
                  Else
                      Combo1.Text = pemain.Fields(i).Value
                  End If
              Else
                  If IsNull(pemain.Fields(i).Value) Then
                      Text1(i) = ""
                  Else
                      Text1(i) = pemain.Fields(i).Value
                  End If
              End If
            Next i
            ShowRs
          End If
          EDITSELECT = 0
       Case 6 'PRIVATE
         EDITSELECT = 6
          If TBar1.Buttons(6).Enabled = True Then
               If Not pemain.RecordCount = 0 Then
                   pemain.MovePrevious
                   If pemain.BOF Then
                      DataErrorMessage (6)
                      pemain.MoveFirst
                   End If
                   For i = 0 To 6
                         If i = 2 Then
                            If IsNull(pemain.Fields(i).Value) Then
                                Combo1.Text = ""
                            Else
                                Combo1.Text = pemain.Fields(i).Value
                            End If
                        Else
                            If IsNull(pemain.Fields(i).Value) Then
                                Text1(i) = ""
                            Else
                                Text1(i) = pemain.Fields(i).Value
                            End If
                        End If
                   Next i
                   ShowRs
               End If
          End If
            EDITSELECT = 0
       Case 7 'NEXT
               EDITSELECT = 7
             If TBar1.Buttons(7).Enabled = True Then
                If Not pemain.RecordCount = 0 Then
                    If pemain.EOF Then
                        pemain.MoveLast
                    End If
                    pemain.MoveNext
                    If pemain.EOF Then
                        DataErrorMessage (7)
                        pemain.MoveLast
                    End If
                    For i = 0 To 6
                            If i = 2 Then
                                If IsNull(pemain.Fields(i).Value) Then
                                    Combo1.Text = ""
                                Else
                                    Combo1.Text = pemain.Fields(i).Value
                                End If
                            Else
                                If IsNull(pemain.Fields(i).Value) Then
                                    Text1(i) = ""
                                Else
                                    Text1(i) = CheckStr(pemain.Fields(i).Value)
                                End If
                            End If
                    Next i
                    ShowRs
                End If
             End If
             EDITSELECT = 0
       Case 8 'LAST
            EDITSELECT = 8
          If TBar1.Buttons(8).Enabled = True Then
            If Not pemain.RecordCount = 0 Then
                pemain.MoveLast
                For i = 0 To 6
                  If i = 2 Then
                     If IsNull(pemain.Fields(i).Value) Then
                        Combo1.Text = ""
                     Else
                        Combo1.Text = pemain.Fields(i).Value
                     End If
                  Else
                     If IsNull(pemain.Fields(i).Value) Then
                        Text1(i) = ""
                     Else
                        Text1(i) = pemain.Fields(i).Value
                     End If
                  End If
                Next i
                ShowRs
            End If
          End If
          EDITSELECT = 9
       Case 9 'ENTER
          If EDITSELECT = 1 Then
            If Len(Trim(Text1(0).Text)) = 0 Or Len(Trim(Combo1.Text)) = 0 Or Len(Trim(Text1(3).Text)) = 0 Or Len(Trim(Text1(4))) = 0 Then
               s = MsgBox("標題、資料出處、版頁、出版日期，不可空白！！", , "User 輸入錯誤")
               Exit Function
            End If
          If p.State = adStateOpen Then p.Close
            'Modify By Cheng 2003/03/25
            '加標題條件
'          strExc(1) = "select count(pE01) from periodical where PE03='" & Combo1.Text & "' AND PE04='" & Text1(3) & "' AND PE05=" & Text1(4).Text & " "
          strExc(1) = "select count(pE01) from periodical where PE01='" & ChgSQL(Me.Text1(0).Text) & "' And PE03='" & Combo1.Text & "' AND PE04='" & Text1(3) & "' AND PE05=" & Text1(4).Text & " "
          p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
          If p.Fields(0).Value <> "0" Then
            MsgBox "此資料已存在"
            Text1(0).SetFocus
            Exit Function
            'Add By Cheng 2002/03/01
            m_blnAddNewOK = False
          End If
          End If
          CheckData = False
          If Len(Text1(6)) <> 0 Then
            Text1_LostFocus (6)
            If CheckData = False Then
               'S = MsgBox("索引錯誤", , "User 輸入錯誤")
               Text1(6) = Trim(Text1(6))
               Text1(6).SetFocus
               Text1_GotFocus (6)
               Exit Function
            End If
         End If
          Select Case EDITSELECT
                 Case 1
                     str = Text1(0).Text & Combo1.Text & Text1(3).Text & Text1(4).Text
                     If Combo1.Text = "中國專利報" Then
                        cnnConnection.Execute "INSERT INTO PERIODICAL(pe01,pe02,pe03,pe04,pe05,pe06,pe07) VALUES('" & Text1(0).Text & "','" & Text1(1).Text & "','" & Combo1.Text & "','" & Text1(3).Text & "'," & Text1(4).Text & ",'" & Text1(5).Text & "','" & Text1(6).Text & "')"
                     Else
                        cnnConnection.Execute "INSERT INTO PERIODICAL(pe01,pe02,pe03,pe04,pe05,pe06,pe07) VALUES('" & Text1(0).Text & "','" & Text1(1).Text & "','" & Combo1.Text & "','" & Text1(3).Text & "'," & ChangeTStringToWString(Text1(4).Text) & ",'" & Text1(5).Text & "','" & Text1(6).Text & "')"
                     End If
                     pemain.ReQuery
                     pemain.Find "AD='" & str & "'"
                 Case 2
                     str = Text1(0).Text & Combo1.Text & Text1(3).Text & Text1(4).Text
                     If Combo1.Text = "中國專利報" Then
                        cnnConnection.Execute "UPDATE PERIODICAL SET PE02='" & Text1(1).Text & "',PE06='" & Text1(5).Text & "',PE07='" & Text1(6).Text & "' WHERE PE03='" & Combo1.Text & "' AND PE04='" & Text1(3).Text & "' AND PE05=" & Text1(4).Text & " and PE01='" & Text1(0).Text & "' "
                     Else
                        cnnConnection.Execute "UPDATE PERIODICAL SET PE02='" & Text1(1).Text & "',PE06='" & Text1(5).Text & "',PE07='" & Text1(6).Text & "' WHERE PE03='" & Combo1.Text & "' AND PE04='" & Text1(3).Text & "' AND PE05=" & ChangeTStringToWString(Text1(4).Text) & " and PE01='" & Text1(0).Text & "' "
                     End If
                     pemain.ReQuery
                     pemain.Find "AD='" & str & "'"
                 Case 4
                     pemain.ReQuery
                     pemain.Find "AD='" & Text1(0).Text & Combo1.Text & Text1(3).Text & Text1(4).Text & "'"
                     'If P.State = adStateOpen Then P.Close
                     'strExc(1) = "SELECT PI01,PI02 FROM PERIODICALINDEX WHERE PI01='" & Text1(0) & "'"
                     'P.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
                     If pemain.EOF Then
                        MsgBox "查無資料"
                        If Not pemain.RecordCount = 0 Then
                           pemain.MoveFirst
                           For i = 0 To 6
                               If i = 2 Then
                                   If IsNull(pemain.Fields(i).Value) Then
                                       Combo1.Text = ""
                                   Else
                                       Combo1.Text = pemain.Fields(i).Value
                                   End If
                               Else
                                   If IsNull(pemain.Fields(i).Value) Then
                                       Text1(i) = ""
                                   Else
                                       Text1(i) = pemain.Fields(i).Value
                                   End If
                               End If
                           Next i
                        End If
                     Else
                        If Not pemain.RecordCount = 0 Then
                            For i = 0 To 6
                                If i = 2 Then
                                    If IsNull(pemain.Fields(i).Value) Then
                                        Combo1.Text = ""
                                    Else
                                        Combo1.Text = pemain.Fields(i).Value
                                    End If
                                Else
                                    If IsNull(pemain.Fields(i).Value) Then
                                        Text1(i) = ""
                                    Else
                                        Text1(i) = pemain.Fields(i).Value
                                    End If
                                End If
                            Next i
                            'pemain.Find "AD='" & Combo1.Text + Text1(3) + Text1(4) & "'"
                        End If
                     End If
          End Select
              For i = 1 To 4
                  TBar1.Buttons(i).Enabled = True
              Next i
              For i = 6 To 9
                  TBar1.Buttons(i).Enabled = True
              Next i
              TBar1.Buttons(11).Enabled = False
              TBar1.Buttons(12).Enabled = False
              TBar1.Buttons(14).Enabled = True
              locktext (1)
              EDITSELECT = 0
              ShowRs
              'Add By Cheng 2002/03/01
              If EDITSELECT = 1 Then
                  m_blnAddNewOK = True
              End If
       Case 10 'CANCEL
         If EDITSELECT <> 4 Then
          If MsgBox("你尚未存檔,確定離開?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
             If EDITSELECT = 1 And Not pemain.RecordCount = 0 Then pemain.MoveFirst
             EDITSELECT = 0
             If pemain.RecordCount <> 0 Then
                  For i = 0 To 6
             If i = 2 Then
                 If IsNull(pemain.Fields(i).Value) Then
                     Combo1.Text = ""
                 Else
                     Combo1.Text = pemain.Fields(i).Value
                 End If
             Else
                 If IsNull(pemain.Fields(i).Value) Then
                     Text1(i) = ""
                 Else
                     Text1(i) = pemain.Fields(i).Value
                 End If
             End If
             Next i
             End If
             For i = 1 To 4
                 TBar1.Buttons(i).Enabled = True
             Next i
             For i = 6 To 9
                 TBar1.Buttons(i).Enabled = True
             Next i
             TBar1.Buttons(11).Enabled = False
             TBar1.Buttons(12).Enabled = False
             TBar1.Buttons(14).Enabled = True
             locktext (1)
             End If
           Else
            If pemain.RecordCount <> 0 Then
                For i = 0 To 6
                    If i = 2 Then
                        If IsNull(pemain.Fields(i).Value) Then
                            Combo1.Text = ""
                        Else
                            Combo1.Text = pemain.Fields(i).Value
                        End If
                    Else
                        If IsNull(pemain.Fields(i).Value) Then
                            Text1(i) = ""
                        Else
                            Text1(i) = pemain.Fields(i).Value
                        End If
                    End If
                 Next i
            End If
            For i = 1 To 4
                TBar1.Buttons(i).Enabled = True
            Next i
            For i = 6 To 9
                TBar1.Buttons(i).Enabled = True
            Next i
            TBar1.Buttons(11).Enabled = False
            TBar1.Buttons(12).Enabled = False
            TBar1.Buttons(14).Enabled = True
            locktext (1)
           End If
           EDITSELECT = 0
           ShowRs
       Case 11 'END
           Unload Me
End Select
End Function


Private Sub Text1_GotFocus(Index As Integer)
Select Case Index
Case 0, 1, 2, 5
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text1(Index).IMEMode = 1
   OpenIme
Case Else
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text1(Index).IMEMode = 2
   CloseIme
End Select
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
If Index = 6 Then
   KeyAscii = UpperCase(KeyAscii)
End If
End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
Select Case Index
Case 3
         'If Len(Text(Index)) <> 0 Then
            If IsNumeric(Text1(Index)) = False Then
               MsgBox "請輸入數字", vbInformation
               Text1(Index).SetFocus
               Text1_GotFocus (Index)
               Exit Sub
            End If
         'End If
Case Else
End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Select Case Index
       Case 0
        If CheckLengthIsOK(Text1(0), 60) = False Then
            Text1(0).SetFocus
            Text1(0).SelStart = 0
            Text1(0).SelLength = Len(Text1(0).Text)
            Exit Sub
        End If
       Case 1
          If CheckLengthIsOK(Text1(1), 20) = False Then
            Text1(1).SetFocus
            Text1(1).SelStart = 0
            Text1(1).SelLength = Len(Text1(1).Text)
            Exit Sub
          End If
       Case 2
          If CheckLengthIsOK(Text1(2), 20) = False Then
            Text1(2).SetFocus
            Text1(2).SelStart = 0
            Text1(2).SelLength = Len(Text1(2).Text)
            Exit Sub
        End If
       Case 4
       If Combo1.Text = "中國專利報" Then
        If CheckIsDate(Text1(4).Text) = False Then
            Text1(4).SetFocus
            Text1(4).SelStart = 0
            Text1(4).SelLength = Len(Text1(4).Text)
            Exit Sub
        End If
       Else
           If Len(Combo1.Text) <> 0 Then
                If CheckIsTaiwanDate(Text1(4).Text) = False Then
                    Text1(4).SetFocus
                    Text1(4).SelStart = 0
                    Text1(4).SelLength = Len(Text1(4).Text)
                    Exit Sub
                End If
            End If
       End If
       Case 5
          If CheckLengthIsOK(Text1(5), 40) = False Then
            Text1(5).SetFocus
            Text1(5).SelStart = 0
            Text1(5).SelLength = Len(Text1(5).Text)
            Exit Sub
        End If
       Case 6
         If Len(Text1(Index)) <> 0 Then
              If p.State = adStateOpen Then p.Close
                  strExc(1) = "SELECT PI02 FROM PERIODICALINDEX WHERE PI01='" & Trim(Text1(6).Text) & "'"
                  p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
                  If p.EOF And p.BOF Then
                     s = MsgBox("無此索引代號,請重新輸入", , "索引錯誤")
                     'ShowNoData
                     Text1(6).SetFocus
                     Text1(6).SelStart = 0
                     Text1(6).SelLength = Len(Text1(6))
                     Exit Sub
                     CheckData = False
                  Else
                     If IsNull(p.Fields(0).Value) Then
                        Label8.Caption = ""
                     Else
                        Label8.Caption = p.Fields(0).Value
                     End If
                     CheckData = True
                  End If
         End If
Case Else
End Select
End Sub

Private Sub locktext(Index As Integer) '鎖住輸入項
Dim j As Integer
Select Case Index
       Case 1 '初值
          For j = 0 To 6
             If j = 2 Then
                Combo1.Enabled = False
             Else
                Text1(j).Locked = True
             End If
          Next j
       Case 2 '新增
          For j = 0 To 6
             If j = 2 Then
                Combo1.Enabled = True
             Else
                Text1(j).Locked = False
             End If
          Next j
       Case 3 '修改
          For j = 0 To 6
             If j = 2 Then
                Combo1.Enabled = False
             ElseIf j = 3 Or j = 0 Then
                Text1(j).Locked = True
             ElseIf j = 4 Then
                Text1(j).Locked = True
             Else
                Text1(j).Locked = False
             End If
          Next j
       Case 4 '查詢
          For j = 0 To 6
             If j = 2 Then
                Combo1.Enabled = True
                ElseIf j = 3 Or j = 0 Then
                Text1(j).Locked = False
             ElseIf j = 4 Then
                Text1(j).Locked = False
             Else
                Text1(j).Locked = True
             End If
          Next j
End Select
End Sub

Sub ShowRs()
   If Len(Trim(Text1(6))) <> 0 Then
      If p.State = adStateOpen Then p.Close
       strExc(1) = "SELECT PI02 FROM PERIODICALINDEX WHERE PI01='" & Text1(6).Text & "'"
       p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
       If p.EOF And p.BOF Then Label8.Caption = "": Exit Sub
      If IsNull(p.Fields(0).Value) Then
          Label8.Caption = ""
      Else
          Label8.Caption = p.Fields(0).Value
      End If
   Else
      Label8.Caption = ""
   End If
End Sub
