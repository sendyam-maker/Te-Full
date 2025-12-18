VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090210_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "公報簡訊資料維護"
   ClientHeight    =   4560
   ClientLeft      =   480
   ClientTop       =   1935
   ClientWidth     =   7755
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7755
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6810
      Top             =   615
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
            Picture         =   "frm090210_1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090210_1.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090210_1.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090210_1.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090210_1.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090210_1.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090210_1.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090210_1.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090210_1.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090210_1.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090210_1.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   7755
      _ExtentX        =   13679
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
      Height          =   300
      Index           =   8
      Left            =   1275
      TabIndex        =   16
      Top             =   3885
      Width           =   495
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "873;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   1335
      Index           =   7
      Left            =   1275
      TabIndex        =   15
      Top             =   2370
      Width           =   6015
      VariousPropertyBits=   -1467989989
      MaxLength       =   300
      ScrollBars      =   2
      Size            =   "10610;2355"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   6
      Left            =   1275
      TabIndex        =   14
      Top             =   2025
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   5
      Left            =   1275
      TabIndex        =   13
      Top             =   1695
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
      Left            =   4155
      TabIndex        =   12
      Top             =   1350
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
      Left            =   2715
      TabIndex        =   11
      Top             =   1350
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
      Left            =   1275
      TabIndex        =   10
      Top             =   1350
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
      Index           =   1
      Left            =   1275
      TabIndex        =   9
      Top             =   1020
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
      Left            =   1275
      TabIndex        =   8
      Top             =   690
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   5
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "索引只可輸入P或U開頭"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   5580
      TabIndex        =   21
      Top             =   1740
      Width           =   1920
   End
   Begin VB.Label Label11 
      Alignment       =   1  '靠右對齊
      Caption         =   "內容摘要字數限制：300中英文字,一個中文算2個"
      Height          =   180
      Left            =   3480
      TabIndex        =   20
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label10 
      Height          =   180
      Left            =   1695
      TabIndex        =   18
      Top             =   1740
      Width           =   3675
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      Caption         =   "彙整旗標："
      Height          =   180
      Left            =   225
      TabIndex        =   17
      Top             =   3930
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      Caption         =   "內容摘要："
      Height          =   180
      Left            =   -15
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin MSForms.Label Label7 
      Height          =   255
      Left            =   1275
      TabIndex        =   6
      Top             =   4230
      Width           =   6045
      VariousPropertyBits=   27
      Size            =   "10663;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "Create  ID："
      Height          =   180
      Left            =   -15
      TabIndex        =   5
      Top             =   4230
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      Caption         =   "公告日期："
      Height          =   180
      Left            =   -15
      TabIndex        =   4
      Top             =   2070
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "索引："
      Height          =   180
      Left            =   585
      TabIndex        =   3
      Top             =   1725
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "國際分類："
      Height          =   180
      Left            =   -15
      TabIndex        =   2
      Top             =   1395
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Caption         =   "公告號數："
      Height          =   180
      Left            =   0
      TabIndex        =   1
      Top             =   1050
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "公告頁數："
      Height          =   180
      Left            =   -15
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frm090210_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (Text,Label7)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan 2010/8/16 日期欄已修改
Option Explicit

Dim pemain As New ADODB.Recordset, p As New ADODB.Recordset, PESUB As New ADODB.Recordset
Dim UserStaff As String, i As Integer, EDITSELECT As Integer, str As String
Dim NEXTSTR As String, NOWSTR As String, s As Integer
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

'Add By Cheng 2002/03/01
Dim m_blnAddNewOK As Boolean '新增成功或失敗

'Add by Morgan 2004/4/27
'是否已觸發 Form Active 事件
Dim bolActive As Boolean


Private Sub Form_Activate()
   'Add by Morgan 2004/4/27
   If bolActive = True Then Exit Sub
   bolActive = True
   
    If pemain.State = adStateOpen Then pemain.Close
    locktext (1)
    strExc(0) = "SELECT ST01 FROM STAFF WHERE ST02='" & strUserName & "'"
    pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
    pemain.Close
    'Label7.Caption = strUserNum
    strExc(0) = "SELECT BB01,BB02,BB03,BB04,BB05,BB06,BB07,BB08,BB09,st02,bb11,bb12,bb01||bb02 as seek FROM BULLETINBRIEF,staff where bb10=st01(+) ORDER BY bb01,BB02"
    pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   For i = 0 To 8
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
   Label7.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
   'Modify By Sindy 2014/4/25
'   For i = 1 To 4
'       TBar1.Buttons(i).Enabled = True
'   Next i
'   For i = 6 To 9
'       TBar1.Buttons(i).Enabled = True
'   Next i
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
   '2014/4/25 END
   TBar1.Buttons(11).Enabled = False
   TBar1.Buttons(12).Enabled = False
   TBar1.Buttons(14).Enabled = True
End Sub

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
            Text(1).SetFocus
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
   m_bInsert = IsUserHasRightOfFunction("frm090210_1", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm090210_1", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm090210_1", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm090210_1", strFind, False)
   MoveFormToCenter Me
   pemain.CursorLocation = adUseClient
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
   'Add by Morgan 2004/4/27
   bolActive = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm090210_1 = Nothing
End Sub


Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Add By Cheng 2002/03/01
Me.Caption = Me.Tag
    Select Case Button.Index
       Case 1 '新增
         If TBar1.Buttons(1).Enabled = True Then
            EDITTOOL (1)
            Text(0).SetFocus
            'Add By Cheng 2002/03/01
            Me.Caption = Me.Tag + " － 新增"
         End If
       Case 2 '修改
         If TBar1.Buttons(2).Enabled = True Then
            EDITTOOL (2)
            'Add By Cheng 2002/03/01
            Me.Caption = Me.Tag + " － 修改"
         End If
         'End If
       Case 3 '刪除
          If TBar1.Buttons(3).Enabled = True Then
            EDITTOOL (3)
            'Add By Cheng 2002/03/01
            Me.Caption = Me.Tag + " － 刪除"
          End If
       Case 4 '查詢
         If TBar1.Buttons(4).Enabled = True Then
            EDITTOOL (4)
            Text(0).SetFocus
            'Add By Cheng 2002/03/01
            Me.Caption = Me.Tag + " － 查詢"
         End If
       Case 6
         If TBar1.Buttons(6).Enabled = True Then
            EDITTOOL (5)
         End If
       Case 7
         If TBar1.Buttons(7).Enabled = True Then
            EDITTOOL (6)
         End If
       Case 8
         If TBar1.Buttons(8).Enabled = True Then
            EDITTOOL (7)
         End If
       Case 9
         If TBar1.Buttons(9).Enabled = True Then
            EDITTOOL (8)
         End If
       Case 11
         If TBar1.Buttons(11).Enabled = True Then
            EDITTOOL (9)
         End If
       Case 12
         If TBar1.Buttons(12).Enabled = True Then
            EDITTOOL (10)
         End If
       Case 14
         If TBar1.Buttons(14).Enabled = True Then
            EDITTOOL (11)
         End If
End Select
   If Button.Index <> 14 And Button.Index <> 1 And Button.Index <> 2 And Button.Index <> 3 And Button.Index <> 4 Then
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
          For i = 0 To 8
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
          locktext (1)
          EDITSELECT = 3
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
            cnnConnection.Execute "DELETE BULLETINBRIEF WHERE bb01||BB02='" & NOWSTR & "' "
            pemain.ReQuery
            pemain.Find "seek='" & NEXTSTR & "'"
            For i = 0 To 8
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
          End If
          Label7.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
'             For i = 1 To 4
'                  TBar1.Buttons(i).Enabled = True
'              Next i
'              For i = 6 To 9
'                  TBar1.Buttons(i).Enabled = True
'              Next i
            'Modify By Sindy 2014/4/25
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
            '2014/4/25 END
            TBar1.Buttons(11).Enabled = False
            TBar1.Buttons(12).Enabled = False
            TBar1.Buttons(14).Enabled = True
       Case 4 'QUTION
          locktext (4)
          EDITSELECT = 4
          For i = 0 To 8
          Text(i).Text = ""
          Next i
          Label7.Caption = "" 'Add by Morgan 2005/1/17
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
       Case 5 'FIRST
       If pemain.BOF And pemain.EOF Then Exit Function
          If TBar1.Buttons(5).Enabled = True Then
          pemain.MoveFirst
          End If
           For i = 0 To 8
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
             Label7.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
       Case 6 'PRIVATE
       If pemain.BOF And pemain.EOF Then Exit Function
          If TBar1.Buttons(6).Enabled = True Then
            pemain.MovePrevious
          If pemain.BOF Then
            DataErrorMessage (6)
            pemain.MoveFirst
          End If
       End If
        For i = 0 To 8
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
             Label7.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
       Case 7 'NEXT
          If pemain.BOF And pemain.EOF Then Exit Function
          If TBar1.Buttons(7).Enabled = True Then
          pemain.MoveNext
          If pemain.EOF Then
          DataErrorMessage (7)
          pemain.MoveLast
          End If
          End If
           For i = 0 To 8
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
             Label7.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
       Case 8 'LAST
       If pemain.BOF And pemain.EOF Then Exit Function
          If TBar1.Buttons(8).Enabled = True Then
          pemain.MoveLast
          End If
           For i = 0 To 8
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
             Label7.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
       Case 9 'ENTER
            '若為新增資料
            If EDITSELECT = 1 Then
                '預設新增失敗
                m_blnAddNewOK = False
            End If
            If EDITSELECT = 1 Then
                If Len(Trim(Text(0))) = 0 Or Len(Trim(Text(1))) = 0 Then
                    s = MsgBox("公告頁數與公告號數不可空白！！", , "User 輸入錯誤")
                    Exit Function
                End If
                If p.State = adStateOpen Then p.Close
                strExc(1) = "select count(*) from BULLETINBRIEF where BB02='" & Text(1) & "' and bb01='" & Text(0) & "' "
                p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
                If p.Fields(0).Value <> "0" Then
                    MsgBox "此資料已存在"
                    Text(0).SetFocus
                    Exit Function
                End If
            End If
            'Add By Cheng 2003/04/10
            '若為新增或修改資料時
            If EDITSELECT = 1 Or EDITSELECT = 2 Then
                'Add By Cheng 2003/04/14
                '若未輸入索引
                If Me.Text(5).Text = "" Then
                    MsgBox "索引不可空白！！", , "User 輸入錯誤"
                    Me.Text(5).SetFocus
                    Text_GotFocus 5
                    Exit Function
                End If
                '若未輸入公告日期
                If Me.Text(6).Text = "" Then
                    MsgBox "公告日期不可空白！！", , "User 輸入錯誤"
                    Me.Text(6).SetFocus
                    Text_GotFocus 6
                    Exit Function
                '若有輸入公告日期
                Else
                    If CheckIsTaiwanDate(Text(6)) = False Then
                        Text(6).SetFocus
                        Text_GotFocus (6)
                        Exit Function
                    End If
                    If Val(strSrvDate(1)) < Val(Text(6)) + 19110000 Then
                       s = MsgBox("公告日期不可大於系統日", , "USER 輸入錯誤！！")
                       Text(6).SetFocus
                       Text_GotFocus (6)
                       Exit Function
                    End If
                End If
            End If
          Select Case EDITSELECT
                 Case 1
                     str = Trim(Text(1).Text)
                     'Modified by Morgan 2022/1/14 +Trim否則寫入後就無法更新
                     cnnConnection.Execute "INSERT INTO BULLETINBRIEF(bb01,bb02,bb03,bb04,bb05,bb06,bb07,bb08,bb09,bb10) VALUES('" & Trim(Text(0)) & "','" & Trim(Text(1)) & "','" & Text(2) & "','" & Text(3) & "','" & Text(4) & "','" & Text(5) & "','" & ChangeTStringToWString(Text(6)) & "','" & Text(7) & "','" & Text(8) & "','" & strUserNum & "')"
                     pemain.ReQuery
                     pemain.Find "seek='" & Trim(Text(0)) & str & "'"
                 Case 2
                     str = Trim(Text(1).Text)
                     cnnConnection.Execute "UPDATE BULLETINBRIEF SET BB03='" & Text(2) & "',BB04='" & Text(3) & "',BB05='" & Text(4) & "',BB06='" & Text(5) & "',BB07='" & ChangeTStringToWString(Text(6)) & "',BB08='" & Text(7) & "',BB09='" & Text(8) & "' WHERE BB02='" & Text(1) & "' and BB01='" & Text(0) & "' "
                     pemain.ReQuery
                     pemain.Find "seek='" & Trim(Text(0)) & str & "'"
                 Case 4
                     pemain.ReQuery
                     pemain.Find "seek='" & Trim(Text(0)) & Trim(Text(1)) & "'"
                     'If P.State = adStateOpen Then P.Close
                     'strExc(1) = "SELECT BB01,BB02 FROM BULLETINBRIEF WHERE BB02='" & Text(1) & "' AND BB10 ='" & UserStaff & "'"
                     'P.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
                     If pemain.EOF Then
                     MsgBox "查無資料"
                     pemain.MoveFirst
                     End If
                  End Select
               
            If Not pemain.EOF Then 'Added by Morgan 2022/1/14 程式有錯，先加判斷
              For i = 0 To 8
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
              Label7.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
            End If
'              For i = 1 To 4
'                  TBar1.Buttons(i).Enabled = True
'              Next i
'              For i = 6 To 9
'                  TBar1.Buttons(i).Enabled = True
'              Next i
              'Modify By Sindy 2014/4/25
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
               '2014/4/25 END
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
             If EDITSELECT = 1 Then pemain.MoveFirst
             EDITSELECT = 0
             For i = 0 To 8
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
             Label7.Caption = CheckStr(pemain.Fields(9).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(10).Value))) & "      " & Format(CheckStr(pemain.Fields(11).Value), "@@:@@")
'              For i = 1 To 4
'                  TBar1.Buttons(i).Enabled = True
'              Next i
'              For i = 6 To 9
'                  TBar1.Buttons(i).Enabled = True
'              Next i
              'Modify By Sindy 2014/4/25
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
               '2014/4/25 END
              TBar1.Buttons(11).Enabled = False
              TBar1.Buttons(12).Enabled = False
              TBar1.Buttons(14).Enabled = True
              locktext (1)
              'End If
       Case 11 'END
           Unload Me
End Select
End Function

Private Sub Text_Change(Index As Integer)
Select Case Index
       Case 5
        If PESUB.State = adStateOpen Then PESUB.Close
        strExc(0) = "SELECT BBI02 FROM BULLETINBRIEFINDEX WHERE BBI01='" & Text(5).Text & "'"
        PESUB.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
        If PESUB.BOF And PESUB.EOF Then Label10 = "": Exit Sub
        If Not PESUB.BOF Then PESUB.MoveFirst
                If IsNull(PESUB.Fields(0).Value) Then
            Label10.Caption = ""
        Else
            Label10.Caption = PESUB.Fields(0).Value
        End If
        PESUB.Close
End Select
End Sub

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
If Index = 7 Then
    If LenB(StrConv(Text(Index), vbFromUnicode)) > 301 Then
        MsgBox ("超出長度")
        KeyAscii = 0
        Text(7).Text = StrConv(MidB(StrConv(Text(7), vbFromUnicode), 1, 300), vbUnicode)
    End If
End If
End Sub

Private Sub Text_KeyUp(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
Select Case Index
Case 0, 1
         'If Len(Text(Index)) <> 0 Then
            '91.12.26 cancel by sonia
            'If IsNumeric(Text(Index)) = False Then
            '   MsgBox "請輸入數字", vbInformation
            '   Text(Index).SetFocus
            '   Text_GotFocus (Index)
            '   Exit Sub
            'End If
             '91.12.26 end
         'End If
Case Else
End Select
If Index = 7 Then
    If LenB(StrConv(Text(Index), vbFromUnicode)) >= 301 Then
         MsgBox ("超出長度")
        
        Text(7).Text = StrConv(MidB(StrConv(Text(7), vbFromUnicode), 1, 300), vbUnicode)
    End If
End If
End Sub

Private Sub Text_LostFocus(Index As Integer)
Select Case Index
Case 6 '公告日期
     If Len(Text(6)) <> 0 Then
         If Val(strSrvDate(1)) < Val(Text(6)) + 19110000 Then
            s = MsgBox("公告日期不可大於系統日", , "USER 輸入錯誤！！")
            Text(6).SetFocus
            Text_GotFocus (6)
            Exit Sub
         End If
     End If
Case Else
End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
'    'add by nick 2004/11/05
'    Case 1
'        If UCase(Mid(Text(1).Text, 1, 1)) <> "I" And UCase(Mid(Text(1).Text, 1, 1)) <> "M" And UCase(Mid(Text(1).Text, 1, 1)) <> "U" Then
'
'        End If
    Case 5 '索引
        'Modify By Cheng 2003/04/14
        '若有輸入索引
        If Me.Text(5).Text <> "" Then
            If TBar1.Buttons(6).Enabled = False Then
                If PESUB.State = adStateOpen Then PESUB.Close
                strExc(0) = "SELECT BBI02 FROM BULLETINBRIEFINDEX WHERE BBI01='" & Text(5).Text & "'"
                PESUB.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
                If PESUB.BOF And PESUB.EOF Then
                    MsgBox "無此索引代號", vbInformation
                    Text(5).SetFocus
                    Text(5).SelStart = 0
                    Text(5).SelLength = Len(Text(5))
                    Cancel = True
                    Exit Sub
                Else
                    Cancel = False
                End If
                If Not PESUB.BOF Then PESUB.MoveFirst
                If IsNull(PESUB.Fields(0).Value) Then
                    Label10.Caption = ""
                Else
                    Label10.Caption = PESUB.Fields(0).Value
                End If
                PESUB.Close
            End If
        End If
    Case 6 '公告日期
        'Modify By Cheng 2003/04/10
        '若有輸入公告日期
        If Me.Text(6).Text <> "" Then
            If CheckIsTaiwanDate(Text(6)) = False Then
                Text(6).SetFocus
                Text(6).SelStart = 0
                Text(6).SelStart = 0
                Cancel = True
            Else
                Cancel = False
            End If
        '若未輸入公告日期
        Else
            Cancel = False
        End If
    Case 7
        If CheckLengthIsOK(Text(7), 300) = False Then
            Text(7).SetFocus
            Text(7).SelStart = 0
            Text(7).SelLength = Len(Text(7))
            Cancel = True
            Exit Sub
        Else
            Cancel = False
        End If
    Case 8
        If Text(8).Text <> "" And Text(8).Text <> "1" Then
            MsgBox "只可為空白或 1", vbInformation
            Text(8).SetFocus
            Text(8).SelStart = 0
            Text(8).SelLength = Len(Text(8))
            Cancel = True
            Exit Sub
        Else
            Cancel = False
        End If
    End Select
End Sub

Private Sub locktext(Index As Integer) '鎖住輸入項
Dim j As Integer
Select Case Index
       Case 1 '初值
          For j = 0 To 8
             Text(j).Locked = True
          Next j
       Case 2 '新增
          For j = 0 To 8
             Text(j).Locked = False
          Next j
       Case 3 '修改
          For j = 0 To 8
             If j = 1 Or j = 0 Then
                Text(j).Locked = True
             Else
                Text(j).Locked = False
             End If
          Next j
       Case 4 '查詢
          For j = 0 To 8
             If j = 1 Or j = 0 Then
                Text(j).Locked = False
             Else
                Text(j).Locked = True
             End If
          Next j
End Select
End Sub
