VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090206_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利案例個人輸入作業"
   ClientHeight    =   5115
   ClientLeft      =   1095
   ClientTop       =   2340
   ClientWidth     =   7905
   ControlBox      =   0   'False
   DrawWidth       =   2
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7905
   Begin VB.TextBox txtPC20 
      Height          =   270
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   19
      Top             =   2190
      Width           =   1215
   End
   Begin VB.ComboBox cboPC19 
      Height          =   300
      ItemData        =   "frm090206_1.frx":0000
      Left            =   1125
      List            =   "frm090206_1.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   18
      Top             =   1890
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Height          =   300
      Left            =   1125
      TabIndex        =   17
      Top             =   1590
      Width           =   4725
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   1125
      TabIndex        =   16
      Top             =   1290
      Width           =   4725
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1125
      TabIndex        =   15
      Top             =   1005
      Width           =   4725
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1125
      TabIndex        =   14
      Top             =   720
      Width           =   4725
   End
   Begin VB.CommandButton Command2 
      Caption         =   "印表(&P)"
      Height          =   405
      Left            =   7092
      TabIndex        =   13
      Top             =   660
      Width           =   756
   End
   Begin VB.CommandButton Command1 
      Caption         =   "索引(&I)"
      Default         =   -1  'True
      Height          =   405
      Left            =   6300
      TabIndex        =   12
      Top             =   660
      Width           =   756
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6930
      Top             =   1860
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
            Picture         =   "frm090206_1.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090206_1.frx":0320
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090206_1.frx":063C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090206_1.frx":0818
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090206_1.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090206_1.frx":0E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090206_1.frx":116C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090206_1.frx":1488
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090206_1.frx":17A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090206_1.frx":1AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090206_1.frx":1DDC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   7905
      _ExtentX        =   13944
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
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   11
      Left            =   1140
      TabIndex        =   27
      Top             =   4170
      Visible         =   0   'False
      Width           =   495
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "873;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   375
      Index           =   9
      Left            =   1125
      TabIndex        =   25
      Top             =   3150
      Width           =   6735
      VariousPropertyBits=   671107099
      MaxLength       =   60
      Size            =   "11880;661"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   1125
      TabIndex        =   24
      Top             =   2760
      Width           =   6735
      VariousPropertyBits=   671107099
      MaxLength       =   60
      Size            =   "11880;661"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   7
      Left            =   3045
      TabIndex        =   23
      Top             =   2445
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   2
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   6
      Left            =   2685
      TabIndex        =   22
      Top             =   2475
      Width           =   255
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "450;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   5
      Left            =   1725
      TabIndex        =   21
      Top             =   2475
      Width           =   855
      VariousPropertyBits=   671107099
      MaxLength       =   6
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   4
      Left            =   1125
      TabIndex        =   20
      Top             =   2475
      Width           =   495
      VariousPropertyBits=   671107099
      MaxLength       =   3
      Size            =   "873;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   615
      Index           =   10
      Left            =   1125
      TabIndex        =   26
      Top             =   3540
      Width           =   6735
      VariousPropertyBits=   -1467989989
      MaxLength       =   400
      ScrollBars      =   2
      Size            =   "11880;1085"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "文書日期："
      Height          =   180
      Left            =   150
      TabIndex        =   31
      Top             =   2205
      Width           =   900
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "文書類型："
      Height          =   180
      Left            =   150
      TabIndex        =   30
      Top             =   1950
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "（ 0:未彙整, 1:已彙整 ）"
      Height          =   180
      Left            =   1800
      TabIndex        =   29
      Top             =   4230
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Caption         =   "彙整旗標："
      Height          =   180
      Left            =   90
      TabIndex        =   28
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   1290
      X2              =   3270
      Y1              =   2565
      Y2              =   2565
   End
   Begin MSForms.Label Label16 
      Height          =   255
      Left            =   1830
      TabIndex        =   11
      Top             =   4770
      Width           =   4635
      VariousPropertyBits=   27
      Size            =   "8176;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label15 
      Height          =   255
      Left            =   1830
      TabIndex        =   10
      Top             =   4470
      Width           =   4650
      VariousPropertyBits=   27
      Size            =   "8202;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label14 
      Caption         =   "Update   ID  DATE ：  "
      Height          =   180
      Left            =   150
      TabIndex        =   9
      Top             =   4815
      Width           =   1620
   End
   Begin VB.Label Label13 
      Caption         =   "Create    ID  DATE ："
      Height          =   180
      Left            =   150
      TabIndex        =   8
      Top             =   4500
      Width           =   1620
   End
   Begin VB.Label Label12 
      Alignment       =   1  '靠右對齊
      Caption         =   "案情摘要："
      Height          =   180
      Left            =   75
      TabIndex        =   7
      Top             =   3570
      Width           =   975
   End
   Begin VB.Label Label11 
      Alignment       =   1  '靠右對齊
      Caption         =   "案例字號："
      Height          =   180
      Left            =   75
      TabIndex        =   6
      Top             =   3165
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   1  '靠右對齊
      Caption         =   "主旨："
      Height          =   180
      Left            =   435
      TabIndex        =   5
      Top             =   2805
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所案號："
      Height          =   180
      Left            =   75
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "備用類："
      Height          =   180
      Left            =   315
      TabIndex        =   3
      Top             =   1635
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      Caption         =   "次次類："
      Height          =   180
      Left            =   195
      TabIndex        =   2
      Top             =   1350
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "次類："
      Height          =   180
      Left            =   435
      TabIndex        =   1
      Top             =   1050
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "主類："
      Height          =   180
      Left            =   435
      TabIndex        =   0
      Top             =   765
      Width           =   615
   End
End
Attribute VB_Name = "frm090206_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/14 改成Form2.0 (Text1,Label15,Label16,Printer列印未改)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset, p As New ADODB.Recordset
Dim EDITSELECT As Integer
Dim NEXTSTR As String, str As String
Dim NOWSTR As String, TXT090206_1 As Object
Dim SYSERR As Integer

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_bPrint As Boolean

Private Sub Combo1_Click()
   'Modify by Morgan 2004/6/1
   '不檢查，存檔時固定抓3碼
   'Combo1.Text = Trim(Left(Combo1.Text, 3))
   If Combo1.Tag <> Combo1.Text Then
      Combo2.Clear
      Combo2 = ""
      Combo1.Tag = Combo1.Text
   End If
End Sub

Private Sub Combo1_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Combo1.IMEMode = 2
   CloseIme
End Sub

Private Sub Combo1_LostFocus()
   SYSERR = 0
   'Modify by Morgan 2004/6/1
   '不檢查，存檔時固定抓3碼
   If Combo1.Text <> "" Then
      If LenB(StrConv(Left(Combo1, 3), vbFromUnicode)) > 3 Then
         MsgBox "前3碼只可輸入半形文數字！", vbInformation
          SYSERR = 1
      End If
   End If
End Sub

Private Sub Combo2_Click()
   'Modify by Morgan 2004/6/1
   '不檢查，存檔時固定抓2碼
   'Combo2.Text = Trim(Left(Combo2.Text, 2))
   If Combo2.Tag <> Combo2.Text Then
      Combo3.Clear
      Combo3 = ""
      Combo2.Tag = Combo2.Text
   End If
End Sub

Private Sub Combo2_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Combo2.IMEMode = 2
   CloseIme
End Sub

Private Sub Combo2_LostFocus()
   SYSERR = 0
   'Modify by Morgan 2004/6/1
   '不檢查，存檔時固定抓2碼
   If Combo2.Text <> "" Then
      If LenB(StrConv(Left(Combo2, 2), vbFromUnicode)) > 2 Then
         MsgBox "前2碼只可輸入半形文數字！", vbInformation
          SYSERR = 1
      End If
   End If
End Sub

Private Sub Combo3_Click()
   Combo3.Text = Trim(Left(Combo3.Text, 2))
   If Combo3.Tag <> Combo3.Text Then
      Combo4.Clear
      Combo4 = ""
      Combo3.Tag = Combo3.Text
   End If
End Sub

Private Sub Combo3_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Combo3.IMEMode = 2
   CloseIme
End Sub

Private Sub Combo3_LostFocus()
   SYSERR = 0
   If Combo3.Text <> "" Then
      If CheckLengthIsOK(Combo3.Text, 2) = False Then
'          Combo3.SetFocus
          SYSERR = 1
          Exit Sub
      End If
   End If
End Sub

Private Sub Combo4_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Combo4.IMEMode = 2
   CloseIme
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo4_LostFocus()
   SYSERR = 0
   If Combo4.Text <> "" Then
      If CheckLengthIsOK(Combo4.Text, 2) = False Then
'          Combo4.SetFocus
          SYSERR = 1
          Exit Sub
      End If
   End If
End Sub

Private Sub Command1_Click()
   If Combo1.Text = "" Then
      If p.State = adStateOpen Then p.Close
      'Modify by Morgan 2004/5/13
      '加主旨
      'strExc(1) = "SELECT DISTINCT PC01 FROM PATENTCASE "
      strExc(1) = "SELECT RPAD(PC01,5,' ')||MAX(DECODE(PC02,'*',PC09,'')) FROM PATENTCASE GROUP BY RPAD(PC01,5,' ')"
      p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
      If p.BOF And p.EOF Then Exit Sub
      If Not p.BOF Then p.MoveFirst
      Combo1.Clear
      Do While Not p.EOF
          Combo1.AddItem p.Fields(0).Value
          p.MoveNext
      Loop
   ElseIf Combo2.Text = "" Then
      If p.State = adStateOpen Then p.Close
      'Modify by Morgan 2004/5/13
      '加主旨
      'strExc(1) = "SELECT DISTINCT PC02 FROM PATENTCASE WHERE PC01='" & Combo1.Text & "' "
      strExc(1) = "SELECT RPAD(PC02,6,' ')||MAX(DECODE(PC03,'*',PC09,'')) FROM PATENTCASE WHERE PC01='" & Left(Combo1.Text, 3) & "' GROUP BY RPAD(PC02,6,' ') "
      p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
      If p.BOF And p.EOF Then Exit Sub
      If Not p.BOF Then p.MoveFirst
      Combo2.Clear
      Do While Not p.EOF
          Combo2.AddItem p.Fields(0).Value
          p.MoveNext
      Loop
   ElseIf Combo3.Text = "" Then
      If p.State = adStateOpen Then p.Close
      'Modify by Morgan 2004/5/13
      '加主旨
      'strExc(1) = "SELECT DISTINCT PC03 FROM PATENTCASE WHERE PC01='" & Combo1.Text & "' AND PC02='" & Combo2.Text & "'"
      strExc(1) = "SELECT RPAD(PC03,6,' ')||MAX(DECODE(PC04,'*',PC09,'')) FROM PATENTCASE WHERE PC01='" & Left(Combo1.Text, 3) & "' AND PC02='" & Left(Combo2.Text, 2) & "' GROUP BY RPAD(PC03,6,' ') "
      p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
      If p.BOF And p.EOF Then Exit Sub
      If Not p.BOF Then p.MoveFirst
      Combo3.Clear
      Do While Not p.EOF
          Combo3.AddItem p.Fields(0).Value
          p.MoveNext
      Loop
      If Combo3.ListCount > 0 Then
         Combo3.Text = Format(Val(Combo3.List(Combo3.ListCount - 1)) + 1, "00")
      End If
   ElseIf Combo4.Text = "" Then
      If p.State = adStateOpen Then p.Close
      strExc(1) = "SELECT DISTINCT PC04 FROM PATENTCASE WHERE PC01='" & Left(Combo1.Text, 3) & "' AND PC02='" & Left(Combo2.Text, 2) & "' AND PC03='" & Combo3.Text & "'"
      p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
      If p.BOF And p.EOF Then Exit Sub
      If Not p.BOF Then p.MoveFirst
      Combo4.Clear
      Do While Not p.EOF
          Combo4.AddItem p.Fields(0).Value
          p.MoveNext
      Loop
   End If
End Sub

Private Sub Command2_Click()
    PrintOption
End Sub

Private Sub PrintOption()
   Dim i As Integer, j As Integer, k As Integer, stTmp As String
   i = 500
   Printer.Orientation = 2
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = GetPrPosX(12000, "專利案例資料"):        Printer.CurrentY = i
   Printer.Print "專利案例資料"
   Printer.Font.Underline = False
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   i = i + 800
   Printer.CurrentX = 500: Printer.CurrentY = i
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 7000: Printer.CurrentY = i
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   i = i + 400
   Printer.Line (500, i)-(10000, i)
   i = i + 500
   Printer.CurrentX = 500: Printer.CurrentY = i
   Printer.Print "類    別：" & CheckStr(pemain.Fields(0)) & " - " & CheckStr(pemain.Fields(1)) & " - " & CheckStr(pemain.Fields(2)) & " - " & CheckStr(pemain.Fields(3))
   i = i + 300
   Printer.CurrentX = 500: Printer.CurrentY = i
   Printer.Print "文書種類：" & cboPC19.Text
   i = i + 300
   Printer.CurrentX = 500: Printer.CurrentY = i
   Printer.Print "文書日期：" & ChangeTStringToTDateString(txtPC20.Text)
   i = i + 300
   Printer.CurrentX = 500: Printer.CurrentY = i
   Printer.Print "本所案號：" & CheckStr(pemain.Fields(4).Value) & "-" & CheckStr(pemain.Fields(5).Value) & "-" & CheckStr(pemain.Fields(6).Value) & "-" & CheckStr(pemain.Fields(7).Value)
   i = i + 300
   Printer.CurrentX = 500: Printer.CurrentY = i
   Printer.Print "案例字號：" & CheckStr(pemain.Fields(9).Value)
   i = i + 300
   Printer.CurrentX = 500: Printer.CurrentY = i
   Printer.Print "主    旨：" & CheckStr(pemain.Fields(8).Value)
   i = i + 300
   Printer.CurrentX = 500: Printer.CurrentY = i
   Printer.Print "案情摘要："
   stTmp = pemain.Fields(10).Value
   For j = 10 To Len(stTmp)
      k = k + 1
      If Printer.TextWidth(Left(stTmp, k)) > 14000 Then
         Printer.CurrentX = 500 + Printer.TextWidth("案情摘要："): Printer.CurrentY = i
         Printer.Print Left(stTmp, k - 1)
         stTmp = Mid(stTmp, k)
         i = i + 300
         k = 0
      End If
   Next
   Printer.CurrentX = 500 + Printer.TextWidth("案情摘要："): Printer.CurrentY = i
   Printer.Print stTmp
   
   i = i + 600
   Printer.CurrentX = 500: Printer.CurrentY = i
   Printer.Print "CREATE  ID：" & CheckStr(pemain.Fields(11).Value)
   Printer.CurrentX = 6000: Printer.CurrentY = i
   Printer.Print "CREATE  DATE：" & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(12).Value)))
   
   i = i + 300
   Printer.CurrentX = 500: Printer.CurrentY = i
   Printer.Print "UPDATE  ID：" & CheckStr(pemain.Fields(14).Value)
   Printer.CurrentX = 6000: Printer.CurrentY = i
   Printer.Print "UPDATE  DATE：" & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(15).Value)))
   Printer.EndDoc
   ShowPrintOk
End Sub

'快速鍵
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

   Select Case KeyCode
       Case vbKeyF2
         If TBar1.Buttons(1).Enabled Then
            Tbar1_ButtonClick Me.TBar1.Buttons(1)
         End If
       Case vbKeyF3
         If TBar1.Buttons(2).Enabled Then
            Tbar1_ButtonClick Me.TBar1.Buttons(2)
         End If
       Case vbKeyF5
         If TBar1.Buttons(3).Enabled Then
            Tbar1_ButtonClick Me.TBar1.Buttons(3)
         End If
       Case vbKeyF4
         If TBar1.Buttons(4).Enabled Then
            Tbar1_ButtonClick Me.TBar1.Buttons(4)
         End If
       Case vbKeyF9, vbKeyReturn
         If TBar1.Buttons(11).Enabled Then
            Tbar1_ButtonClick Me.TBar1.Buttons(11)
         End If
       Case vbKeyHome
         If TBar1.Buttons(6).Enabled Then
            Tbar1_ButtonClick Me.TBar1.Buttons(6)
         End If
       Case vbKeyEnd
         If TBar1.Buttons(9).Enabled Then
            Tbar1_ButtonClick Me.TBar1.Buttons(9)
         End If
       Case vbKeyPageUp
         If TBar1.Buttons(7).Enabled Then
            Tbar1_ButtonClick Me.TBar1.Buttons(7)
         End If
       Case vbKeyPageDown
         If TBar1.Buttons(8).Enabled Then
            Tbar1_ButtonClick Me.TBar1.Buttons(8)
         End If
       Case vbKeyF10
         If TBar1.Buttons(12).Enabled Then
            Tbar1_ButtonClick Me.TBar1.Buttons(12)
         End If
        'Modify By Cheng 2004/02/20
        '取消按Esc鍵離開此畫面的功能
'       Case vbKeyEscape
'         EDITTOOL (11)
        'End
   End Select

End Sub

Private Sub Form_Load()
   m_bInsert = IsUserHasRightOfFunction("frm090206_1", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm090206_1", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm090206_1", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm090206_1", strFind, False)
   m_bPrint = IsUserHasRightOfFunction("frm090206_1", strPrint, False)

   MoveFormToCenter Me
   If pemain.State = adStateOpen Then pemain.Close
   pemain.CursorLocation = adUseClient
   strExc(0) = "SELECT ST01 FROM STAFF WHERE ST02='" & strUserName & "' and st04='1' "
   pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If pemain.EOF And pemain.BOF Then MsgBox "無此LOGIN人員之資料": Unload Me: Exit Sub
   pemain.Close
   'strExc(0) = "SELECT SR03,SR04,SR05,SR06,SR07,SR08 FROM STAFF,STAFF_RIGHT WHERE SR01=ST05(+) AND ST01='" & strUserNum & "' AND SR02='frm090206' "
   'pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   'If pemain.BOF And pemain.EOF Then MsgBox "無此人員之權限資料", vbInformation: Exit Sub
   'NEWDATE = pemain.Fields(0).Value
   'UPDATE = pemain.Fields(1).Value
   'DELETE = pemain.Fields(2).Value
   'QUITE = pemain.Fields(3).Value
   
   'Add by Morgan 2004/5/3
   cboPC19.AddItem "判決", 0
   cboPC19.AddItem "決定書", 1
   cboPC19.AddItem "其他", 2
   
   If pemain.State = adStateOpen Then pemain.Close
   strExc(0) = "select pc01,pc02,pc03,pc04,pc05,pc06,pc07,pc08,pc09,pc10,pc11,s1.st02,pc13,pc14,s2.st02,pc16,pc17,PC01||PC02||PC03||PC04 AS A, PC18, PC19, PC20 from patentcase,staff s1,staff s2 where  pc12='" & strUserNum & "'  AND pc12=s1.st01(+) and pc15=s2.st01(+) and (PC18<>'1' OR PC18 IS NULL)  ORDER BY PC01,PC02,PC03,PC04 "
   pemain.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If pemain.EOF And pemain.BOF Then
      MsgBox "資料庫內無資料", vbInformation
   Else
      Call FormRead
   End If
   
   locktext (1)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text1(9).IMEMode = 1
   
   ToolControl 1

   'Add By Cheng 2002/03/01
   Me.Tag = Me.Caption
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090206_1 = Nothing
End Sub

Private Function EDITTOOL(Index As Integer)
   Select Case Index
       Case 1 'NEW
          EDITSELECT = 1
          Call FormClear
          cboPC19.ListIndex = 0
          Text1(11).Text = "0"
          locktext (2)
          ToolControl 0

       Case 2 'UPDATE
          EDITSELECT = 2
          locktext (3)
          ToolControl 0
          
       Case 3 'DELETE
         If MsgBox("是否要刪除此筆資料", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
            NOWSTR = CheckStr(pemain.Fields("A").Value)
            pemain.MoveNext
            If pemain.EOF Then
               pemain.MoveFirst
               NEXTSTR = CheckStr(pemain.Fields("A").Value)
            Else
              NEXTSTR = CheckStr(pemain.Fields("A").Value)
            End If
            
            cnnConnection.Execute "delete PATENTCASE where PC01||PC02||PC03||PC04='" & NOWSTR & "'"
            pemain.ReQuery
            pemain.Find "A='" & NEXTSTR & "'"
            
            '資料清空時
            If pemain.EOF Then
               Call FormClear
            Else
               Call FormRead
            End If
            ToolControl 1
         End If
            
       Case 4 'QUERY
       
          EDITSELECT = 4
          Call FormClear
          locktext (4)
          ToolControl 0
          
      Case 5 'FIRST
         If pemain.BOF And pemain.EOF Then
            MsgBox "資料庫內無資料", vbInformation
            Exit Function
         End If
         
         If TBar1.Buttons(5).Enabled = True Then
           pemain.MoveFirst
           Call FormRead
         End If
          
      Case 6 'PREVIOUS
         If pemain.BOF And pemain.EOF Then
            MsgBox "資料庫內無資料", vbInformation
         Else
            pemain.MovePrevious
            If pemain.BOF Then
               DataErrorMessage (6)
               pemain.MoveFirst
            End If
            Call FormRead
         End If
         
      Case 7 'NEXT
         If pemain.BOF Or pemain.EOF Then
             MsgBox "資料庫內無資料", vbInformation
         Else
           pemain.MoveNext
           If pemain.EOF Then
              DataErrorMessage (7)
              pemain.MoveLast
           End If
           Call FormRead
         End If
          
       Case 8 'LAST
         If pemain.BOF Or pemain.EOF Then
            MsgBox "資料庫內無資料", vbInformation
         Else
            pemain.MoveLast
            Call FormRead
         End If
         
      Case 9 'ENTER
         
         If EDITSELECT = 1 Or EDITSELECT = 2 Then  '新增, 修改
         
            '資料檢查
            If EDITSELECT = 1 Then
               Combo1_LostFocus
               If SYSERR = 1 Then
                  Combo1.SetFocus
                  Combo1_GotFocus
                  Exit Function
               End If
               Combo2_LostFocus
               If SYSERR = 1 Then
                  Combo2.SetFocus
                  Combo2_GotFocus
                  Exit Function
               End If
               Combo3_LostFocus
               If SYSERR = 1 Then
                  Combo3.SetFocus
                  Combo3_GotFocus
                  Exit Function
               End If
               Combo4_LostFocus
               If SYSERR = 1 Then
                  Combo4.SetFocus
                  Combo4_GotFocus
                  Exit Function
               End If
            End If
            
            For Each TXT090206_1 In Text1
               Text1_LostFocus (TXT090206_1.Index)
               If SYSERR = 1 Then Exit Function
            Next
            
            Dim stPC19 As String, stPC20 As String, stPC07 As String, stPC08 As String
               
            If cboPC19.ListIndex >= 0 Then
               stPC19 = "'" & cboPC19.ListIndex & "'"
            Else
               stPC19 = "Null"
            End If
            
            If txtPC20 <> "" Then
               stPC20 = Val(txtPC20) + 19110000
            Else
               stPC20 = "NULL"
            End If
            
            stPC07 = IIf(Len(Trim(Text1(4))) <> 0, IIf(Len(Trim(Text1(5))) <> 0, IIf(Len(Trim(Text1(6).Text)) = 0, "0", Text1(6).Text), ""), "")
            stPC08 = IIf(Len(Trim(Text1(4))) <> 0, IIf(Len(Trim(Text1(5))) <> 0, IIf(Len(Trim(Text1(7).Text)) = 0, "00", Text1(7).Text), ""), "")
         End If
         
         If EDITSELECT = 1 Then  '新增
            If Combo1.Text = "" Then Combo1.Text = "000"
            If Combo2.Text = "" Then Combo2.Text = "00"
            If Combo3.Text = "" Then Combo3.Text = "00"
            If Combo4.Text = "" Then Combo4.Text = "00"
            
            If p.State = adStateOpen Then p.Close
            strExc(1) = "select count(PC01) from PATENTCASE where PC01='" & Left(Combo1.Text, 3) & "' AND PC02='" & Left(Combo2.Text, 2) & "' AND PC03='" & Combo3.Text & "' AND PC04='" & Combo4.Text & "'"
            p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
            If p.Fields(0).Value <> "0" Then
               MsgBox "此資料已存在"
               Combo1.SetFocus
            ElseIf txtPC20.Text = "" Then  '文書日期不可空白
               MsgBox "文書日期不可空白！", vbExclamation
               txtPC20.SetFocus
            Else
               str = Left(Combo1.Text, 3) & Left(Combo2.Text, 2) & Combo3.Text & Combo4.Text
               cnnConnection.Execute "INSERT INTO PATENTCASE(pc01,pc02,pc03,pc04,pc05,pc06,pc07,pc08,pc09,pc10,pc11,pc18,pc19,pc20) VALUES('" & Left(Combo1.Text, 3) & "','" & Left(Combo2.Text, 2) & "','" & Combo3.Text & "', '" & Combo4.Text & "','" & ChgSQL(Text1(4)) & "','" & ChgSQL(Text1(5)) & "','" & ChgSQL(stPC07) & "','" & ChgSQL(stPC08) & "','" & ChgSQL(Text1(8)) & "','" & ChgSQL(Text1(9)) & "','" & ChgSQL(Text1(10)) & "','" & ChgSQL(Text1(11)) & "'," & stPC19 & "," & stPC20 & ")"
               pemain.ReQuery
               pemain.Find "A='" & str & "'"
               Tbar1_ButtonClick Me.TBar1.Buttons(1)
            End If
            
         Else
            If EDITSELECT = 2 Then '修改
               str = CheckStr(pemain.Fields("A").Value)
               cnnConnection.Execute "begin user_data.user_enabled:=1; UPDATE PATENTCASE  SET PC05='" & ChgSQL(Text1(4)) & "',PC06='" & ChgSQL(Text1(5)) & "',PC07='" & ChgSQL(stPC07) & "',PC08='" & ChgSQL(stPC08) & "',PC09='" & ChgSQL(Text1(8)) & "',PC10='" & ChgSQL(Text1(9)) & "',PC11='" & ChgSQL(Text1(10)) & "',PC18='" & ChgSQL(Text1(11)) & "',PC19=" & stPC19 & ",PC20=" & stPC20 & " WHERE PC01='" & Left(Combo1.Text, 3) & "' AND PC02='" & Left(Combo2.Text, 2) & "' AND PC03='" & Combo3.Text & "' AND PC04='" & Combo4.Text & "'; end;"
               pemain.ReQuery
               pemain.Find "A='" & str & "'"
         
            ElseIf EDITSELECT = 4 Then '查詢
               pemain.ReQuery
               pemain.Find "A='" & Left(Combo1.Text, 3) + Left(Combo2.Text, 2) + Combo3.Text + Combo4.Text & "'"
               If pemain.EOF Then
                  MsgBox "查無資料"
                  pemain.ReQuery
                  If pemain.RecordCount > 0 Then pemain.MoveFirst
               End If
            End If
            If pemain.EOF Then
               Call FormClear
            Else
               Call FormRead
            End If
            EDITSELECT = 0
            locktext (1)
            ToolControl 1
         End If
            
       Case 10 'CANCEL
         
         If EDITSELECT = 1 Or EDITSELECT = 2 Then  '新增, 修改
            If MsgBox("你尚未存檔,確定離開?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
               Exit Function
            End If
         End If
         
         If pemain.EOF Then
            FormClear
         Else
            FormRead
         End If
         EDITSELECT = 0
         locktext (1)
         ToolControl 1
         
      Case 11 'END
         Unload Me
           
   End Select
   
End Function

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   
   Me.Caption = Me.Tag
   Select Case Button.Index
      Case 1 '新增
         EDITTOOL (1)
         Combo1.SetFocus
         Me.Caption = Me.Tag + " － 新增"
      Case 2 '修改
         EDITTOOL (2)
         Me.Caption = Me.Tag + " － 修改"
      Case 3 '刪除
         EDITTOOL (3)
         Me.Caption = Me.Tag + " － 刪除"
      Case 4 '查詢
         EDITTOOL (4)
         Combo1.SetFocus
         Me.Caption = Me.Tag + " － 查詢"
      Case 6
         EDITTOOL (5)
      Case 7
         EDITTOOL (6)
      Case 8
         EDITTOOL (7)
      Case 9
         EDITTOOL (8)
      Case 11  '確定
         EDITTOOL (9)
      Case 12  '取消
         EDITTOOL (10)
      Case 14  '結束
         EDITTOOL (11)
   End Select

End Sub

Private Sub Text1_GotFocus(Index As Integer)
   Select Case Index
      Case 8, 9, 10
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
   If Index >= 4 And Index <= 7 Then
      KeyAscii = UpperCase(KeyAscii)
   ElseIf Index = 11 Then
      If KeyAscii <> 49 And KeyAscii <> 48 And KeyAscii <> 8 Then KeyAscii = 0
   End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
   Case 7
      If p.State = adStateOpen Then p.Close
      SYSERR = 0
      If Text1(4) <> "" And Text1(5) <> "" Then
         Select Case Text1(4)
            Case "P", "CFP", "FCP"
               strExc(1) = "SELECT PA01 FROM PATENT WHERE PA01='" & Text1(4).Text & "' AND PA02='" & Text1(5).Text & "' AND PA03='" & IIf(Len(Trim(Text1(6).Text)) = 0, "0", Text1(6).Text) & "' AND PA04='" & IIf(Len(Trim(Text1(7).Text)) = 0, "00", Text1(7).Text) & "'"
            Case "T", "CFT", "FCT"
               strExc(1) = "SELECT TM01 FROM TRADEMARK WHERE TM01='" & Text1(4).Text & "' AND TM02='" & Text1(5).Text & "' AND TM03='" & IIf(Len(Trim(Text1(6).Text)) = 0, "0", Text1(6).Text) & "' AND TM04='" & IIf(Len(Trim(Text1(7).Text)) = 0, "00", Text1(7).Text) & "'"
            Case "L", "FLC", "CFL"
               strExc(1) = "SELECT LC FROM LAWCASE WHERE LC01='" & Text1(4).Text & "' AND LC02='" & Text1(5).Text & "' AND LC03='" & IIf(Len(Trim(Text1(6).Text)) = 0, "0", Text1(6).Text) & "' AND LC04='" & IIf(Len(Trim(Text1(7).Text)) = 0, "00", Text1(7).Text) & "'"
            Case "LA"
               strExc(1) = "SELECT HC01 FROM HIRECASE WHERE HC01='" & Text1(4).Text & "' AND HC02='" & Text1(5).Text & "' AND HC03='" & IIf(Len(Trim(Text1(6).Text)) = 0, "0", Text1(6).Text) & "' AND HC04='" & IIf(Len(Trim(Text1(7).Text)) = 0, "00", Text1(7).Text) & "'"
            Case Else
               strExc(1) = "SELECT SP01 FROM SERVICEPRACTICE WHERE SP01='" & Text1(4).Text & "' AND SP02='" & Text1(5).Text & "' AND SP03='" & IIf(Len(Trim(Text1(6).Text)) = 0, "0", Text1(6).Text) & "' AND SP04='" & IIf(Len(Trim(Text1(7).Text)) = 0, "00", Text1(7).Text) & "'"
         End Select
         p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
         If p.BOF And p.EOF Then
            MsgBox "輸入之本所案號不存在於基本檔中"
            Text1(5).SetFocus
            Text1(5).SelStart = 0
            Text1(5).SelLength = Len(Text1(5))
            SYSERR = 1
         Else
            SYSERR = 0
         End If
      End If
   'Add By Cheng 2004/02/20
   Case 8 '主旨
      SYSERR = 0
      If Me.Text1(Index).Text <> "" Then
         If CheckLengthIsOK(Me.Text1(Index).Text, 60) = False Then
             Me.Text1(Index).SetFocus
             SYSERR = 1
             Exit Sub
         End If
      End If
   Case 9 '案例字號
      SYSERR = 0
      If Me.Text1(Index).Text <> "" Then
         If CheckLengthIsOK(Me.Text1(Index).Text, 60) = False Then
             Me.Text1(Index).SetFocus
             SYSERR = 1
             Exit Sub
         End If
      End If
   Case 10 '案情摘要
      SYSERR = 0
      If Me.Text1(Index).Text <> "" Then
         If CheckLengthIsOK(Me.Text1(Index).Text, 400) = False Then
             Me.Text1(Index).SetFocus
             SYSERR = 1
             Exit Sub
         End If
      End If
   End Select
End Sub
'index1: 1=預設(唯讀) , 2=新增 , 3=修改 ,4=查詢
Private Sub locktext(index1 As Integer)   '鎖住輸入項
   Dim j As Integer
   Select Case index1
      Case 1 '初值
         For j = 0 To 11
            If j = 0 Then
               'Modify by Morgan 2004/5/13
               'Combo1.Locked = True
               Combo1.Enabled = False
               
            ElseIf j = 1 Then
               'Modify by Morgan 2004/5/13
               'Combo2.Locked = True
               Combo2.Enabled = False
               
            ElseIf j = 2 Then
               'Modify by Morgan 2004/5/13
               'Combo3.Locked = True
               Combo3.Enabled = False
               
            ElseIf j = 3 Then
               'Modify by Morgan 2004/5/13
               'Combo4.Locked = True
               Combo4.Enabled = False
            Else
               Text1(j).Locked = True
            End If
         Next j
         cboPC19.Locked = True
         txtPC20.Locked = True
         
      Case 2 '新增
         For j = 0 To 11
            If j = 0 Then
               'Modify by Morgan 2004/5/13
               'Combo1.Locked = False
               Combo1.Enabled = True
               
            ElseIf j = 1 Then
               'Modify by Morgan 2004/5/13
               'Combo2.Locked = False
               Combo2.Enabled = True
               
            ElseIf j = 2 Then
               'Modify by Morgan 2004/5/13
               'Combo3.Locked = False
               Combo3.Enabled = True
               
            ElseIf j = 3 Then
               'Modify by Morgan 2004/5/13
               'Combo4.Locked = False
               Combo4.Enabled = True
            Else
               Text1(j).Locked = False
            End If
         Next j
         cboPC19.Locked = False
         txtPC20.Locked = False
      
      Case 3 '修改
         For j = 0 To 11
            If j = 0 Then
               'Modify by Morgan 2004/5/13
               'Combo1.Locked = True
               Combo1.Enabled = False
               
            ElseIf j = 1 Then
               'Modify by Morgan 2004/5/13
               'Combo2.Locked = True
               Combo2.Enabled = False
               
            ElseIf j = 2 Then
               'Modify by Morgan 2004/5/13
               'Combo3.Locked = True
               Combo3.Enabled = False
               
            ElseIf j = 3 Then
               'Modify by Morgan 2004/5/13
               'Combo4.Locked = True
               Combo4.Enabled = False
            Else
               Text1(j).Locked = False
            End If
         Next j
         cboPC19.Locked = False
         txtPC20.Locked = False
          
      Case 4 '查詢
         For j = 0 To 11
            If j = 0 Then
               'Modify by Morgan 2004/5/13
               'Combo1.Locked = False
               Combo1.Enabled = True
               
            ElseIf j = 1 Then
               'Modify by Morgan 2004/5/13
               'Combo2.Locked = False
               Combo2.Enabled = True
               
            ElseIf j = 2 Then
               'Modify by Morgan 2004/5/13
               'Combo3.Locked = False
               Combo3.Enabled = True
               
            ElseIf j = 3 Then
               'Modify by Morgan 2004/5/13
               'Combo4.Locked = False
               Combo4.Enabled = True
            Else
               Text1(j).Locked = True
            End If
         Next j
         cboPC19.Locked = True
         txtPC20.Locked = True
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Text1(Index).Locked = True Then Exit Sub
   Select Case Index
      Case 11
         If Text1(Index) = "" Then Text1(Index) = "0"
   End Select
End Sub
'Add by Morgan 2004/5/3
Private Sub txtPC20_GotFocus()
   TextInverse txtPC20
End Sub
'Add by Morgan 2004/5/3
Private Sub txtPC20_Validate(Cancel As Boolean)
   If txtPC20.Locked = False Then
      If Len(Trim(txtPC20)) <> 0 Then
        If CheckIsTaiwanDate(txtPC20) = False Then
            txtPC20.SetFocus
            txtPC20_GotFocus
            Cancel = True
        ElseIf Val(strSrvDate(1)) < Val(txtPC20) + 19110000 Then
            MsgBox "文書日期不可大於系統日", vbExclamation, "USER 輸入錯誤！！"
            txtPC20.SetFocus
            txtPC20_GotFocus
            Cancel = True
        End If
      End If
   End If
End Sub

'Add by Morgan 2004/5/3
'清除畫面
Private Sub FormClear()
          
   Dim i As Integer
   
   Combo1.Text = "": Combo1.Clear
   Combo2.Text = "": Combo2.Clear
   Combo3.Text = "": Combo3.Clear
   Combo4.Text = "": Combo4.Clear
   For i = 4 To 11
      Text1(i).Text = ""
   Next i
   cboPC19.ListIndex = -1
   txtPC20.Text = ""
   Label15.Caption = ""
   Label16.Caption = ""
   
End Sub

'Add by Morgan 2004/5/3
'讀取資料
Private Sub FormRead()
   
   Dim i As Integer
   
   If IsNull(pemain.Fields(0).Value) Then
       Combo1.Text = ""
   Else
       Combo1.Text = pemain.Fields(0).Value
   End If
   If IsNull(pemain.Fields(1).Value) Then
       Combo2.Text = ""
   Else
       Combo2.Text = pemain.Fields(1).Value
   End If
   If IsNull(pemain.Fields(2).Value) Then
       Combo3.Text = ""
   Else
       Combo3.Text = pemain.Fields(2).Value
   End If
   If IsNull(pemain.Fields(3).Value) Then
       Combo4.Text = ""
   Else
       Combo4.Text = pemain.Fields(3).Value
   End If
   For i = 4 To 10
      Text1(i) = "" & pemain.Fields(i).Value
   Next i
   Label15.Caption = CheckStr(pemain.Fields(11).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(12).Value))) & "      " & Format(CheckStr(pemain.Fields(13).Value), "@@:@@")
   Label16.Caption = CheckStr(pemain.Fields(14).Value) & "      " & ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(pemain.Fields(15).Value))) & "      " & Format(CheckStr(pemain.Fields(16).Value), "@@:@@")
   Text1(11) = "" & pemain.Fields("PC18").Value
   cboPC19.ListIndex = -1
   If Not IsNull(pemain.Fields("PC19")) Then
      If Val(pemain.Fields("PC19")) >= 0 And Val(pemain.Fields("PC19")) <= 2 Then
         cboPC19.ListIndex = Val(pemain.Fields("PC19"))
      End If
   End If
   txtPC20.Text = ChangeWStringToTString(CheckStr("" & pemain.Fields("PC20").Value))
            
End Sub
'Add by Morgan 2004/5/3
'iCtrl:0=輸入狀態，1=瀏覽狀態
Private Sub ToolControl(iCtrl As Integer)

   Dim i As Integer
   
   Select Case iCtrl
   
      Case 0
         For i = 1 To 4
           TBar1.Buttons(i).Enabled = False
         Next i
         For i = 6 To 9
           TBar1.Buttons(i).Enabled = False
         Next i
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
         Command2.Enabled = False
      Case 1
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate And pemain.RecordCount > 0 Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete And pemain.RecordCount > 0 Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery And pemain.RecordCount > 0 Then
            TBar1.Buttons(4).Enabled = True
            If m_bPrint Then
               Command2.Enabled = True
            Else
               Command2.Enabled = False
            End If
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         For i = 6 To 9
            TBar1.Buttons(i).Enabled = True
         Next i
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
   End Select
   
End Sub
