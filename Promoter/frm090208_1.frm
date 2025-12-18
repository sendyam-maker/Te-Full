VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090208_1 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "公報簡訊個人輸入作業"
   ClientHeight    =   4605
   ClientLeft      =   1155
   ClientTop       =   1530
   ClientWidth     =   8070
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7485
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
            Picture         =   "frm090208_1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090208_1.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090208_1.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090208_1.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090208_1.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090208_1.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090208_1.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090208_1.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090208_1.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090208_1.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090208_1.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8070
      _ExtentX        =   14235
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
   Begin MSForms.TextBox Text 
      Height          =   1575
      Index           =   7
      Left            =   1155
      TabIndex        =   7
      Top             =   2550
      Width           =   6855
      VariousPropertyBits=   -1467989989
      MaxLength       =   300
      ScrollBars      =   2
      Size            =   "12091;2778"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   6
      Left            =   1155
      TabIndex        =   6
      Top             =   2205
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   7
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   5
      Left            =   1155
      TabIndex        =   5
      Top             =   1845
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   2
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   4
      Left            =   3930
      TabIndex        =   4
      Top             =   1500
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   11
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   3
      Left            =   2550
      TabIndex        =   3
      Top             =   1500
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   11
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   2
      Left            =   1155
      TabIndex        =   2
      Top             =   1515
      Width           =   1185
      VariousPropertyBits=   671107099
      MaxLength       =   11
      Size            =   "2090;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   1
      Left            =   1155
      TabIndex        =   1
      Top             =   1170
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   10
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   0
      Left            =   1155
      TabIndex        =   0
      Top             =   825
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   5
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "索引只可輸入P或U開頭"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   6060
      TabIndex        =   19
      Top             =   1920
      Width           =   1920
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      Caption         =   "內容摘要字數限制：300中英文字,一個中文算2個"
      Height          =   180
      Left            =   3960
      TabIndex        =   18
      Top             =   2340
      Width           =   3975
   End
   Begin MSForms.Label Label9 
      Height          =   255
      Left            =   1155
      TabIndex        =   17
      Top             =   4320
      Width           =   6660
      VariousPropertyBits=   27
      Size            =   "11747;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      Height          =   210
      Left            =   1635
      TabIndex        =   16
      Top             =   1875
      Width           =   4335
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "Create  ID："
      Height          =   180
      Left            =   150
      TabIndex        =   15
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "內容摘要："
      Height          =   180
      Left            =   150
      TabIndex        =   14
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      Caption         =   "公告日期："
      Height          =   180
      Left            =   150
      TabIndex        =   13
      Top             =   2250
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "索引："
      Height          =   180
      Left            =   510
      TabIndex        =   12
      Top             =   1890
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "國際分類："
      Height          =   180
      Left            =   150
      TabIndex        =   11
      Top             =   1575
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Caption         =   "公告號數："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   10
      Top             =   1215
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "公告頁數："
      Height          =   180
      Left            =   150
      TabIndex        =   9
      Top             =   855
      Width           =   975
   End
End
Attribute VB_Name = "frm090208_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (Text,Label9)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset, PESUB As New ADODB.Recordset
Dim UserStaff As String, s As Integer
Dim EDITSELECT As Integer, i As Integer
Dim NOWSTR As String, NEXTSTR As String

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

'Add By Cheng 2002/03/01
Dim m_blnAddNewOK As Boolean '新增成功或失敗

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyF2
         If TBar1.Buttons(1).Enabled = True Then
             EDITTOOL (1)
            Text(0).SetFocus
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
        'Modify By Cheng 2004/02/20
        '取消按Esc鍵退出畫面的功能
'       Case vbKeyEscape
'         If TBar1.Buttons(14).Enabled = True Then
'            EDITTOOL (11)
'         End If
        'End
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

Private Sub Form_Load()
   m_bInsert = IsUserHasRightOfFunction("frm090208_1", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm090208_1", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm090208_1", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm090208_1", strFind, False)

MoveFormToCenter Me
If pemain.State = adStateOpen Then pemain.Close
pemain.CursorLocation = adUseClient
PESUB.CursorLocation = adUseClient
strExc(0) = "SELECT ST01 FROM STAFF WHERE ST01='" & strUserNum & "'"
pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
If pemain.BOF And pemain.EOF Then MsgBox "無此LOGIN人員之資料", vbInformation: Unload Me
pemain.Close
'Label9.Caption = strUserNum
'Modify By Cheng 2002/03/01
'strExc(0) = "SELECT BB01,BB02,nvl(BB03,' '),nvl(BB04,' '),nvl(BB05,' '),nvl(BB06,' '),nvl(BB07,0),nvl(BB08,' '),nvl(BB09,' '),st02,bb11,bb12,bb01||bb02 as seek FROM BULLETINBRIEF,staff WHERE BB10='" & strUserNum & "' AND (BB09 <>'1' or bb09 is null) and bb10=st01(+) ORDER BY 1,2"
strExc(0) = "SELECT BB01,BB02,nvl(BB03,' '),nvl(BB04,' '),nvl(BB05,' '),nvl(BB06,' '),nvl(BB07,0),nvl(BB08,' '),nvl(BB09,' '),st02,bb11,bb12,bb01||bb02 as seek,BB10 FROM BULLETINBRIEF,staff WHERE BB10='" & strUserNum & "' AND (BB09 <>'1' or bb09 is null) and bb10=st01(+) ORDER BY 1,2"
pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'Modify By Cheng 2002/03/01
If pemain.EOF And pemain.BOF Then
   MsgBox "資料庫內無資料", vbInformation
'   Exit Sub
Else
   For i = 0 To 7
       If i = 6 Then
           Text(i).Text = ChangeWStringToTString(CheckStr(pemain.Fields(i).Value))
       Else
           Text(i).Text = CheckStr(pemain.Fields(i).Value)
       End If
   Next i
   Label9.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
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
Me.Tag = Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm090208_1 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Add By Cheng 2002/03/01
Me.Caption = Me.Tag
Select Case Button.Index
       Case 1 '新增
         EDITTOOL (1)
         'Add By Cheng 2002/03/01
         Me.Caption = Me.Tag + " － 新增"
       Case 2 '修改
         'Modify By Cheng 2002/03/01
         If pemain.RecordCount > 0 Then
            EDITTOOL (2)
            'Add By Cheng 2002/03/01
            Me.Caption = Me.Tag + " － 修改"
         Else
            Exit Sub
         End If
       Case 3 '刪除
         'Modify By Cheng 2002/03/01
         If pemain.RecordCount > 0 Then
            EDITTOOL (3)
            'Add By Cheng 2002/03/01
            Me.Caption = Me.Tag + " － 刪除"
         Else
            Exit Sub
         End If
       Case 4 '查詢
         EDITTOOL (4)
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
'Add By Cheng 2002/03/01
Dim strBB10 As String
Dim strBB11 As String
Dim strBB12 As String
Select Case Index
       Case 1 'NEW
          EDITSELECT = 1
          locktext (2)
          For i = 0 To 7
          Text(i).Text = ""
          Next i
          Label8.Caption = ""
          For i = 1 To 4
          TBar1.Buttons(i).Enabled = False
          Next i
          For i = 6 To 9
          TBar1.Buttons(i).Enabled = False
          Next i
          TBar1.Buttons(11).Enabled = True
          TBar1.Buttons(12).Enabled = True
          TBar1.Buttons(14).Enabled = False
          Text(0).SetFocus
          Text_GotFocus (0)
       Case 2 'UPDATA
          'Add By Cheng 2002/03/01
          If pemain.RecordCount <= 0 Then Exit Function
          If pemain.Fields(8).Value = "1" Then MsgBox "此筆資料彙整旗標為1不可做修改動作", vbInformation: Exit Function
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
          Text(0).SetFocus
          Text_GotFocus (0)
       Case 3 'DELETE
         'Modify By Cheng 2002/03/01
'          If pemain.Fields(8).Value = "1" Then MsgBox "此筆資料彙整旗標為1不可做刪除動作", vbInformation: Exit Function
          If pemain.RecordCount > 0 Then
            If pemain.Fields(8).Value = "1" Then MsgBox "此筆資料彙整旗標為1不可做刪除動作", vbInformation: Exit Function
          End If
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
          NOWSTR = CheckStr(pemain.Fields(0).Value) & CheckStr(pemain.Fields(1).Value)
          pemain.MoveNext
          If pemain.EOF Then
            pemain.MoveFirst
            NEXTSTR = CheckStr(pemain.Fields(0).Value) & CheckStr(pemain.Fields(1).Value)
          Else
            NEXTSTR = CheckStr(pemain.Fields(0).Value) & CheckStr(pemain.Fields(1).Value)
          End If
              cnnConnection.Execute "DELETE BULLETINBRIEF WHERE bb01||BB02='" & NOWSTR & "' AND BB10='" & strUserNum & "'"
            pemain.ReQuery
            pemain.Find "seek='" & NEXTSTR & "'"
            If pemain.RecordCount > 0 Then
               For i = 0 To 7
               If i = 6 Then
                   Text(i).Text = ChangeWStringToTString(CheckStr(pemain.Fields(i)))
               Else
                   Text(i).Text = CheckStr(pemain.Fields(i).Value)
               End If
               Next i
               Label9.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
            Else
               For i = 0 To Me.Text.Count - 1
                  Me.Text(i).Text = Empty
               Next i
            End If
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
          For i = 0 To 7
          Text(i).Text = ""
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
          Text(0).SetFocus
          Text_GotFocus (0)
       Case 5 'FIRST
          If Not (pemain.BOF And pemain.EOF) Then
            If TBar1.Buttons(5).Enabled = True Then
            pemain.MoveFirst
            End If
            For i = 0 To 7
            If IsNull(pemain.Fields(i).Value) Then
            Text(i).Text = ""
            Else
            If i = 6 Then
                Text(i).Text = ChangeWStringToTString(CheckStr(pemain.Fields(i).Value))
            Else
                Text(i).Text = CheckStr(pemain.Fields(i).Value)
            End If
            End If
            Next i
            Label9.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
            Text_LostFocus (5)
          End If
       Case 6 'PRIVATE
          If Not (pemain.BOF And pemain.EOF) Then
            If TBar1.Buttons(6).Enabled = True Then
                pemain.MovePrevious
                If pemain.BOF Then
                    DataErrorMessage (6)
                    pemain.MoveFirst
                End If
            End If
            For i = 0 To 7
                If IsNull(pemain.Fields(i).Value) Then
                Text(i).Text = ""
            Else
                If i = 6 Then
                    Text(i).Text = ChangeWStringToTString(pemain.Fields(i).Value)
                Else
                    Text(i).Text = pemain.Fields(i).Value
                End If
            End If
            Next i
            Label9.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
            Text_LostFocus (5)
          End If
       Case 7 'NEXT
          If Not (pemain.BOF And pemain.EOF) Then
                If TBar1.Buttons(7).Enabled = True Then
                    pemain.MoveNext
                    If pemain.EOF Then
                        DataErrorMessage (7)
                        pemain.MoveLast
                    End If
                End If
                For i = 0 To 7
                    If IsNull(pemain.Fields(i).Value) Then
                       Text(i).Text = ""
                    Else
                        If i = 6 Then
                            Text(i).Text = ChangeWStringToTString(pemain.Fields(i).Value)
                       Else
                       Text(i).Text = pemain.Fields(i).Value
                       End If
                End If
                Next i
                Label9.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
                Text_LostFocus (5)
          End If
       Case 8 'LAST
          If Not (pemain.BOF And pemain.EOF) Then
            If TBar1.Buttons(8).Enabled = True Then
                pemain.MoveLast
            End If
            For i = 0 To 7
                If IsNull(pemain.Fields(i).Value) Then
                    Text(i).Text = ""
                Else
                    If i = 6 Then
                        Text(i).Text = ChangeWStringToTString(pemain.Fields(i).Value)
                    Else
                        Text(i).Text = pemain.Fields(i).Value
                    End If
                End If
            Next i
            Label9.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
            Text_LostFocus (5)
          End If
       Case 9 'ENTER
          Dim str As String
          If EDITSELECT = 1 Then
          If PESUB.State = adStateOpen Then PESUB.Close
          strExc(1) = "select count(bb02) from bulletinbrief where bb01='" & Text(0) & "' and bb02='" & Text(1) & "'"
          PESUB.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
          If PESUB.Fields(0).Value <> "0" Then
            MsgBox "此資料已存在"
            Text(0).SetFocus
            Exit Function
            m_blnAddNewOK = False
          End If
          End If
          If EDITSELECT < 4 Then
                'Add By Cheng 2002/12/25
                '若為新增或修改時, 檢查輸入資料完整性
                If EDITSELECT = 1 Or EDITSELECT = 2 Then
                    If TxtValidate = False Then m_blnAddNewOK = False: Exit Function
                End If
                
               If CheckIsTaiwanDate(Text(6)) = False Then
                   Text(6).SetFocus
                   Text(6).SelStart = 0
                   Text(6).SelLength = Len(Text(6))
                   Exit Function
               Else
                    'Modify By Cheng 2003/04/25
                    '取消檢查工作天
'                   If Not ChkWork(ChangeTStringToWString(Text(6))) Then
'                      Text(6).SetFocus
'                      Text(6).SelStart = 0
'                      Text(6).SelLength = Len(Text(6))
'                      Exit Function
'                   End If
               End If
               If CheckLengthIsOK(Text(7), 300) = False Then
                   Text(7).SetFocus
                   Text(7).SelStart = 0
                   Text(7).SelLength = Len(Text(7))
                   Exit Function
               End If
           End If
          Select Case EDITSELECT
                 Case 1
                     str = Trim(Text(1).Text)
                     cnnConnection.Execute "INSERT INTO BULLETINBRIEF (BB01,BB02,BB03,BB04,BB05,BB06,BB07,BB08,BB09,BB10) VALUES('" & Text(0).Text & "','" & Text(1).Text & "','" & Text(2).Text & "','" & Text(3).Text & "','" & Text(4).Text & "','" & Text(5).Text & "'," & ChangeTStringToWString(Text(6).Text) & ",'" & Text(7).Text & "','0','" & strUserNum & "')"
                     pemain.ReQuery
                     pemain.Find "seek='" & Trim(Text(0)) & str & "'"
                 Case 2
                     str = Trim(Text(1).Text)
                     'Modify By Cheng 2002/03/01
                     '更新方式採用先刪除後新增(因為PK可改)
'                     cnnConnection.Execute "UPDATE BULLETINBRIEF SET BB03='" & Text(2) & "',BB04='" & Text(3) & "',BB05='" & Text(4).Text & "',BB06='" & Text(5).Text & "',BB07=" & ChangeTStringToWString(Text(6).Text) & ",BB08='" & Text(7) & "' WHERE BB10='" & strUserNum & "' AND BB02='" & Text(1).Text & "' AND BB01='" & Text(0).Text & "'  "
                     strBB10 = pemain.Fields("BB10")
                     strBB11 = pemain.Fields("BB11")
                     strBB12 = pemain.Fields("BB12")
                     
                     '911107 nickchen
                     On Error GoTo CheckingErr
                     cnnConnection.BeginTrans

                     cnnConnection.Execute "Delete From BULLETINBRIEF Where BB01='" & pemain.Fields("BB01") & "' And BB02='" & pemain.Fields("BB02") & "'"
                     DoEvents
                     cnnConnection.Execute "INSERT INTO BULLETINBRIEF (BB01,BB02,BB03,BB04,BB05,BB06,BB07,BB08,BB09,BB10,BB11,BB12) VALUES('" & Text(0).Text & "','" & Text(1).Text & "','" & Text(2).Text & "','" & Text(3).Text & "','" & Text(4).Text & "','" & Text(5).Text & "'," & ChangeTStringToWString(Text(6).Text) & ",'" & Text(7).Text & "','0','" & strUserNum & "'," & strBB11 & "," & strBB12 & ")"
                     DoEvents
                     cnnConnection.Execute "UPDATE BULLETINBRIEF SET BB11=" & strBB11 & ", BB12=" & strBB12 & " WHERE BB02='" & Text(1).Text & "' AND BB01='" & Text(0).Text & "'  "
                     
                     '911107 nickchen
                     cnnConnection.CommitTrans
                     
                     pemain.ReQuery
                     pemain.Find "seek='" & Trim(Text(0)) & str & "'"
                 Case 4
                     pemain.ReQuery
                     pemain.Find "seek='" & Trim(Text(0)) & Trim(Text(1)) & "'"
                     'Modify By Cheng 2002/03/01
                     If pemain.RecordCount > 0 Then
                        If pemain.EOF Then
                           MsgBox "查無資料"
                           pemain.MoveFirst
                           For i = 0 To 7
                           If i = 6 Then
                               Text(i).Text = ChangeWStringToTString(CheckStr(pemain.Fields(i).Value))
                           Else
                               Text(i).Text = CheckStr(pemain.Fields(i).Value)
                           End If
                           Next i
                           Label9.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
                        Else
                           For i = 0 To 7
                           If i = 6 Then
                               Text(i).Text = ChangeWStringToTString(CheckStr(pemain.Fields(i).Value))
                           Else
                               Text(i).Text = CheckStr(pemain.Fields(i).Value)
                           End If
                           Next i
                           Label9.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
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
              Text(0).SetFocus
              'Add By Cheng 2002/03/01
              If EDITSELECT = 1 Then
                  m_blnAddNewOK = True
              End If
       Case 10 'CANCEL
            'Modify by Morgan 2005/1/17 查詢不用問
            'If MsgBox("你尚未存檔,確定離開?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
            If EDITSELECT <> 4 Then
               If MsgBox("你尚未存檔,確定離開?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
                 Exit Function
               End If
            End If
          
             If EDITSELECT = 1 Then
               If pemain.State = 1 Then
                  'Modify By Cheng 2002/03/01
                  If pemain.RecordCount > 0 Then
                     pemain.MoveFirst
                  End If
               End If
            End If
             EDITSELECT = 0
               'Modify By Cheng 2002/03/01
               If pemain.RecordCount > 0 Then
                  For i = 0 To 7
                  If i = 6 Then
                        Text(i).Text = ChangeWStringToTString(pemain.Fields(i).Value)
                  Else
                        Text(i).Text = pemain.Fields(i).Value
                  End If
                  Next i
                  Label9.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
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
          'End If
       Case 11 'END
           Unload Me
End Select
 '911107 nick transation
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
     Resume Next
End Function

Private Sub Text_GotFocus(Index As Integer)
Select Case Index
Case 7
      'edit by nickc 2007/07/11 切換輸入法改用API
      'Text(Index).IMEMode = 1
      OpenIme
Case Else
      'edit by nickc 2007/07/11 切換輸入法改用API
      'Text(Index).IMEMode = 2
      CloseIme
End Select
    Text(Index).SelStart = 0
    Text(Index).SelLength = Len(Text(Index))
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
'Add by Morgan 2005/1/17 公告頁數只可書數字
Case 0
   If KeyAscii <> 8 And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
'Modify by Morgan 2005/1/17 加 1 也要控制
Case 1, 2, 3, 4, 5
   KeyAscii = UpperCase(KeyAscii)
Case Else
End Select
End Sub

Private Sub Text_KeyUp(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
'Modify By Cheng 2002/12/25
''Modify By Cheng 2002/03/01
'If Me.Text(Index).Locked = False Then
'   Select Case Index
'   Case 0, 1
''            If Len(Text(Index)) <> 0 Then
'               If IsNumeric(Text(Index)) = False Then
'                  MsgBox "請輸入數字", vbInformation
'                  Text(Index).SetFocus
'                  Text_GotFocus (Index)
'                  Exit Sub
'               End If
''            End If
'   Case Else
'   End Select
'End If
End Sub

Private Sub Text_LostFocus(Index As Integer)
''Add By Cheng 2002/03/01
'If Me.Text(Index).Locked = False Then
'   If CheckKeyIn(Index) <> 0 Then
'      Me.Text(Index).SetFocus
'      Exit Sub
'   End If
'End If
'
''Modify By Cheng 2002/03/01
'If Me.Text(Index).Locked = False Then
'   Select Case Index
'         Case 0, 1
'            If Len(Text(Index)) <> 0 Then
'               If IsNumeric(Text(Index)) = False Then
'                  MsgBox "請輸入數字", vbInformation
'                  Text(Index).SetFocus
'                  Text_GotFocus (Index)
'                  Exit Sub
'               End If
'            End If
'          Case 5
'           'If TBar1.Buttons(6).Enabled = False Then
'           If PESUB.State = adStateOpen Then PESUB.Close
'           strExc(0) = "SELECT BBI02 FROM BULLETINBRIEFINDEX WHERE BBI01='" & Text(5).Text & "'"
'           PESUB.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'           If PESUB.BOF And PESUB.EOF Then
'               MsgBox "無此索引代號", vbInformation
'               Text(5).SetFocus
'               Text(5).SelStart = 0
'               Text(5).SelLength = Len(Text(5))
'               Exit Sub
'           End If
'           If Not PESUB.BOF Then PESUB.MoveFirst
'           Label8.Caption = PESUB.Fields(0).Value
'           PESUB.Close
'           'End If
'          Case 6
'            If Len(Trim(Text(6))) <> 0 Then
'              If CheckIsTaiwanDate(Text(6)) = False Then
'                  Text(6).SetFocus
'                  Text(6).SelStart = 0
'                  Text(6).SelLength = Len(Text(6))
'                  Exit Sub
'              Else
'                  If Val(GetTodayDate) < Val(Text(6)) + 19110000 Then
'                     s = MsgBox("公告日期不可大於系統日", , "USER 輸入錯誤！！")
'                     Text(6).SetFocus
'                     Text_GotFocus (6)
'                     Exit Sub
'                  End If
'              End If
'           End If
'          Case 7
'            If Len(Trim(Text(7))) <> 0 Then
'           If CheckLengthIsOK(Text(7), 200) = False Then
'               Text(7).SetFocus
'               Text(7).SelStart = 0
'               Text(7).SelLength = Len(Text(7))
'               Exit Sub
'           End If
'           End If
'          Case Else
'   End Select
'End If
End Sub

Private Sub locktext(Index As Integer) '鎖住輸入項
Dim j As Integer
Select Case Index
       Case 1 '初值
          For j = 0 To 7
             Text(j).Locked = True
          Next j
       Case 2 '新增
          For j = 0 To 7
             Text(j).Locked = False
          Next j
       Case 3 '修改
          For j = 0 To 7
             'Modify By Cheng 2002/03/01
             '公告頁數及公告號數可修改
'             If j = 1 Or j = 0 Then
'                Text(j).Locked = True
'             Else
                Text(j).Locked = False
'             End If
          Next j
       Case 4 '查詢
          For j = 0 To 7
             If j = 1 Or j = 0 Then
                Text(j).Locked = False
             Else
                Text(j).Locked = True
             End If
          Next j
          
End Select
End Sub

Private Function CheckKeyIn(Index As Integer)
CheckKeyIn = -1
Select Case Index
Case 2, 3, 4 '國際分類
   If Len(Me.Text(Index).Text) > 0 Then
        'Modify By Cheng 2002/12/25
        '不檢查第一碼一定介於A －H
'      '檢查第一碼
'      If UCase(Mid(Me.Text(Index).Text, 1, 1)) < "A" Or UCase(Mid(Me.Text(Index).Text, 1, 1)) > "H" Then
'         MsgBox "第一碼必須介於 A － H 之間!!!", vbExclamation
''         Me.Text(Index).SelStart = 0
''         Me.Text(Index).SelLength = 1
'         Exit Function
'      End If
   End If
   If Len(Me.Text(Index).Text) > 1 Then
      '檢查第二碼
      If IsNumeric(Mid(Me.Text(Index), 2, 1)) = False Then
         MsgBox "第二碼必須為數字!!!", vbExclamation
         Me.Text(Index).SelStart = 1
         Me.Text(Index).SelLength = 1
         Exit Function
      End If
   End If
   If Len(Me.Text(Index).Text) > 2 Then
      '檢查第三碼
      If IsNumeric(Mid(Me.Text(Index), 3, 1)) = False Then
         MsgBox "第三碼必須為數字!!!", vbExclamation
         Me.Text(Index).SelStart = 2
         Me.Text(Index).SelLength = 1
         Exit Function
      End If
   End If
   If Len(Me.Text(Index).Text) > 3 Then
      '檢查第四碼
      If UCase(Mid(Me.Text(Index).Text, 4, 1)) < "A" Or UCase(Mid(Me.Text(Index).Text, 4, 1)) > "Z" Then
         MsgBox "第四碼必須為英文字!!!", vbExclamation
         Me.Text(Index).SelStart = 3
         Me.Text(Index).SelLength = 1
         Exit Function
      End If
   End If
End Select
CheckKeyIn = 0
End Function

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
'Add By Cheng 2002/03/01
If Me.Text(Index).Locked = False Then
   If CheckKeyIn(Index) <> 0 Then
      Me.Text(Index).SetFocus
        'Add By Cheng 2002/12/25
        Select Case Index
        Case 2, 3, 4
            '無動作
        Case Else
            Text_GotFocus Index
        End Select
      Cancel = True
      Exit Sub
   End If
End If

'Modify By Cheng 2002/03/01
If Me.Text(Index).Locked = False Then
   Select Case Index
        'edit by nick 2004/11/05
         'Case 0, 1
         Case 0
            If Len(Text(Index)) <> 0 Then
               If IsNumeric(Text(Index)) = False Then
                  MsgBox "請輸入數字", vbInformation
                  Text(Index).SetFocus
                  Text_GotFocus (Index)
                  Cancel = True
               End If
            End If
          Case 5
           'If TBar1.Buttons(6).Enabled = False Then
           If PESUB.State = adStateOpen Then PESUB.Close
           strExc(0) = "SELECT BBI02 FROM BULLETINBRIEFINDEX WHERE BBI01='" & Text(5).Text & "'"
           PESUB.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
           If PESUB.BOF And PESUB.EOF Then
               MsgBox "無此索引代號", vbInformation
               Text(5).SetFocus
               Text(5).SelStart = 0
               Text(5).SelLength = Len(Text(5))
                Cancel = True
           End If
           If Not PESUB.BOF Then PESUB.MoveFirst
            'Modify By Cheng 2002/12/25
'           Label8.Caption = PESUB.Fields(0).Value
            If PESUB.EOF = False Then
               Label8.Caption = "" & PESUB.Fields(0).Value
            Else
               Label8.Caption = ""
            End If
           PESUB.Close
           'End If
          Case 6
            If Len(Trim(Text(6))) <> 0 Then
              If CheckIsTaiwanDate(Text(6)) = False Then
                  Text(6).SetFocus
                  Text(6).SelStart = 0
                  Text(6).SelLength = Len(Text(6))
                  Cancel = True
              Else
                  If Val(strSrvDate(1)) < Val(Text(6)) + 19110000 Then
                     s = MsgBox("公告日期不可大於系統日", , "USER 輸入錯誤！！")
                     Text(6).SetFocus
                     Text_GotFocus (6)
                      Cancel = True
                  End If
              End If
           End If
          Case 7
            If Len(Trim(Text(7))) <> 0 Then
           If CheckLengthIsOK(Text(7), 300) = False Then
               Text(7).SetFocus
               Text(7).SelStart = 0
               Text(7).SelLength = Len(Text(7))
                Cancel = True
           End If
           End If
          Case Else
   End Select
End If

End Sub
   
'Add By Cheng 2002/12/25
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.Text(0).Text = "" Then
    MsgBox "請輸入公告頁數!!!", vbExclamation + vbOKOnly
    Me.Text(0).SetFocus
    Text_GotFocus 0
    Exit Function
End If
If Me.Text(1).Text = "" Then
    MsgBox "請輸入公告號數!!!", vbExclamation + vbOKOnly
    Me.Text(1).SetFocus
    Text_GotFocus 1
    Exit Function
End If

For Each objTxt In Me.Text
   If objTxt.Enabled = True And objTxt.Visible = True Then
      Cancel = False
      Text_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

TxtValidate = True
End Function


