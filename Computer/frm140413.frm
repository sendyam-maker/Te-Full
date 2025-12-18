VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140413 
   BorderStyle     =   1  '單線固定
   Caption         =   "程式修改公告維護"
   ClientHeight    =   5676
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8052
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5676
   ScaleWidth      =   8052
   Begin VB.CommandButton CmdPaper 
      Caption         =   "附件"
      Height          =   540
      Left            =   60
      TabIndex        =   41
      Top             =   4080
      Width           =   700
   End
   Begin VB.CommandButton CmdSelect 
      Caption         =   "全取消"
      Height          =   300
      Index           =   0
      Left            =   7120
      TabIndex        =   43
      Top             =   800
      Width           =   700
   End
   Begin VB.CommandButton CmdSelect 
      Caption         =   "全選"
      Height          =   300
      Index           =   1
      Left            =   6360
      TabIndex        =   42
      Top             =   800
      Width           =   700
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   1
      Top             =   930
      Width           =   615
   End
   Begin VB.Frame FrmSysKind 
      Height          =   4455
      Left            =   5040
      TabIndex        =   37
      Top             =   840
      Width           =   2895
      Begin VB.CheckBox Check1 
         Caption         =   "Trademark1"
         Height          =   375
         Index           =   15
         Left            =   1560
         TabIndex        =   18
         Top             =   2400
         Width           =   1200
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Finance"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Patpro"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Trademark"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Promoter"
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   19
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Writer"
         Height          =   375
         Index           =   10
         Left            =   1560
         TabIndex        =   20
         Top             =   2880
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Person"
         Height          =   375
         Index           =   12
         Left            =   1560
         TabIndex        =   22
         Top             =   3360
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Account"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Casher"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Salary"
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Query"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Patpro1"
         Height          =   375
         Index           =   6
         Left            =   1560
         TabIndex        =   16
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Law"
         Height          =   375
         Index           =   8
         Left            =   1560
         TabIndex        =   14
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "File"
         Height          =   375
         Index           =   11
         Left            =   240
         TabIndex        =   21
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Computer"
         Height          =   375
         Index           =   13
         Left            =   240
         TabIndex        =   23
         Top             =   3975
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "AutoBatch"
         Height          =   375
         Index           =   14
         Left            =   1560
         TabIndex        =   24
         Top             =   3975
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  '置中對齊
         Caption         =   "公佈系統別"
         Height          =   255
         Left            =   165
         TabIndex        =   38
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.Frame FrmYN 
      Height          =   615
      Left            =   1200
      TabIndex        =   36
      Top             =   2340
      Width           =   1455
      Begin VB.OptionButton Option1 
         Caption         =   "是"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "否"
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   600
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
            Picture         =   "frm140413.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140413.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140413.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140413.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140413.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140413.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140413.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140413.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140413.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140413.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140413.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   8052
      _ExtentX        =   14203
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
      Index           =   5
      Left            =   3795
      TabIndex        =   6
      Top             =   2490
      Width           =   750
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1323;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   1785
      Index           =   4
      Left            =   840
      TabIndex        =   8
      Top             =   3540
      Width           =   4095
      VariousPropertyBits=   -1466941413
      MaxLength       =   1000
      ScrollBars      =   2
      Size            =   "7223;3149"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   840
      TabIndex        =   7
      Top             =   3000
      Width           =   4095
      VariousPropertyBits=   -1466941413
      MaxLength       =   200
      ScrollBars      =   2
      Size            =   "7223;873"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   2010
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   1290
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   5
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   930
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "時數："
      Height          =   255
      Index           =   8
      Left            =   3225
      TabIndex        =   45
      Top             =   2520
      Width           =   600
   End
   Begin MSForms.Label Label23 
      Height          =   300
      Left            =   240
      TabIndex        =   44
      Top             =   5370
      Width           =   6615
      VariousPropertyBits=   27
      Caption         =   "Create ID:           Date         Time             Update ID:                Date                  Time"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "請將上線日期及流水號抄至請作單紙本。"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3360
      TabIndex        =   40
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "內容："
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   39
      Top             =   3540
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "(民國年月日)"
      Height          =   255
      Left            =   2085
      TabIndex        =   35
      Top             =   960
      Width           =   1020
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   2
      Left            =   1200
      TabIndex        =   34
      Top             =   1695
      Width           =   1935
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3413;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "流水號："
      Height          =   255
      Index           =   1
      Left            =   3225
      TabIndex        =   32
      Top             =   960
      Width           =   720
   End
   Begin MSForms.Label Label2 
      Height          =   300
      Index           =   1
      Left            =   2100
      TabIndex        =   31
      Top             =   1320
      Width           =   960
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "摘要："
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   30
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "是否公布："
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   29
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "上線日期："
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   28
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "請作單日："
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   27
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "需求部門："
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   26
      Top             =   1695
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "需求人員："
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "frm140413"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/15 Form2.0已修改
'2013/04/30 Modify By Amy 增加開啟請作單檔
'2013/03/21 Create by Amy
Option Explicit

Dim RbMain As New ADODB.Recordset, bp As New ADODB.Recordset
Dim ActionEdit As Integer '0:add 1:update 2:query 3:cancel
Dim m_AttachPath As String 'Add By Amy 2013/04/30

Dim i As Integer
'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean

' 第一筆資料
Dim m_FirstKEY(2) As String
' 最後一筆資料
Dim m_LastKEY(2) As String
' 目前正在顯示
Dim m_CurrKEY(2) As String

Dim oText As Object, oCheck As CheckBox, idx As Integer
Dim strCDate As String


'Add By Amy 2013/04/30 開啟請作單檔
Private Sub CmdPaper_Click()
   Dim hLocalFile As Long
   Dim stFileName As String

   Screen.MousePointer = vbHourglass
  
  stFileName = Text1(0) & Text2
   If GetAttachFile(m_AttachPath, stFileName) = False Then
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   stFileName = m_AttachPath & "\" & stFileName & ".pdf"
   ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Select_All (Index)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
        Text1(0).SetFocus
        RbEdit 0
        Case vbKeyF3
        Text1(1).SetFocus
        Text1(0).TabStop = False
        RbEdit 1
        Case vbKeyF5
        RbEdit 2
        Case vbKeyF4
        RbEdit 5
        Case vbKeyHome
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                ActionRb 0
             End If
        Case vbKeyPageUp
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                 ActionRb 1
             End If
        Case vbKeyPageDown
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                 ActionRb 2
             End If
        Case vbKeyEnd
             If Not (ActionEdit = 0 Or ActionEdit = 1) Then
                  ActionRb 3
             End If
        Case vbKeyF9
        '欄位驗證
        RbEdit 3
        Text1(0).TabStop = True
        '抓資料
        Case vbKeyF10
         RbEdit 4
        Case vbKeyEscape
        Unload Me
        Set frm140413 = Nothing
End Select
End Sub

'Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
   'Add By Amy 2014/09/10 當focus在備註欄時按enter鍵維持換行功能而不是存檔功能
   If KeyAscii = 13 And UCase(Me.ActiveControl.Name) = UCase("Text1") Then
      If Me.ActiveControl.Index = 4 Then
         Exit Sub
      End If
   End If
   'end 2014/09/10
    Select Case KeyAscii
      Case vbKeyReturn:
         If ActionEdit <> 3 Then
            KeyAscii = 0
            Form_KeyDown vbKeyF9, 0
         End If
    End Select
End Sub

Private Sub Form_Load()
 '取得使用者執行各項功能的權限
   'm_bInsert = IsUserHasRightOfFunction("frm140413", strAdd, False)
   'm_bUpdate = IsUserHasRightOfFunction("frm140413", strEdit, False)
   'm_bDelete = IsUserHasRightOfFunction("frm140413", strDel, False)
   
   MoveFormToCenter Me
   
   ActionEdit = 3 'Modify By Sindy 2016/9/5 此句程式由後往前移,先預設值
   m_AttachPath = App.path & "\PGMBulletinAttach"
   RefreshRange
   
   'Modify By Sindy 2021/12/15
   'GetFirstRecordVal '設定第一筆key值
   GetLastRecordVal '設定最後一筆資料
   '2021/12/15 END
   
   ToolBarSet 1        '設定ToolBar按鈕顯示
   'ActionEdit = 3           'cancel/第一次進入
End Sub

Private Sub ActionRb(ByVal sta As Integer) ', rsTmp As ADODB.Recordset
   TxtLock 2
      Select Case sta
         Case 0 'MoveFirst
           GetFirstRecordVal
         Case 1 'MovePrv
            GetPreRecordVal
         Case 2 'MoveNext
            GetNextRecordVal
         Case 3 'MoveLast
            GetLastRecordVal
      End Select
   
End Sub

Private Sub SetTxtValue()
Dim strTmp As String, m_ibf01 As String, m_ibf02 As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

   'Modify by Amy 2014/07/16 +BU15
   strSql = "SELECT BU01, BU03,BU04,BU05,BU14,Decode(NVL(BU15,0),0,null,BU15) As BU15,BU06,BU07,BU02,NVL(ST02,'NULL') As ST02,NVL(ST03,'NULL') As ST03," & _
               "BU08,BU09,BU10,BU11,BU12,BU13 FROM PGMBulletin,STAFF " & _
                "WHERE BU03=ST01(+) And BU01='" & m_CurrKEY(0) & "' And BU02='" & m_CurrKEY(1) & "' ORDER BY BU01,BU03"
         
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
    For Each oText In Text1
      idx = oText.Index
      If IsNull(rsTmp(idx)) Then
         oText = ""
      Else
        Select Case idx
           Case 0, 2
            oText = ChangeWStringToTString(rsTmp(idx))
           Case 1, 3, 4
            oText = rsTmp(idx)
           'Modify by Amy 2014/07/16 +BU15
           Case 5
            oText = Val(rsTmp(idx))
        End Select
      End If
    Next
    
    Text1(1).Tag = Text1(1).Text  'Added by Lydia 2020/04/27 記錄員工編號
    
    '是否公佈
    Select Case rsTmp.Fields("BU06")
        Case 0
          Option1(1).Value = True
        Case 1
          Option1(0).Value = True
    End Select
   
   Text2.Text = IIf(rsTmp.Fields("BU02") <= 9, Format(Val(rsTmp.Fields("BU02")), "00"), rsTmp.Fields("BU02"))    '流水號
    
    If rsTmp.Fields("ST02") = "NULL" Then  '需求人員名稱
        MsgBox ("需求人員有誤")
        Label2(1).Caption = ""
        Label2(2).Caption = ""
    Else
        Label2(1).Caption = rsTmp.Fields("ST02")
        'Added by Lydia 2023/12/27
        If DBDATE(Text1(0)) >= 新部門啟用日 Then
           Label2(2).Caption = GetDeptNameA0922("" & rsTmp.Fields("BU03"))
        Else
        'end 2023/12/27
           Label2(2).Caption = IIf(ClsPDGetStaffDeptName(rsTmp.Fields("ST03"), strTmp), strTmp, "")  '需求部門名稱
        End If
    End If
        
   RedSystemKind IIf(IsNull(rsTmp.Fields("BU07")), "", rsTmp.Fields("BU07"))   '系統別
   End If
   
   'Add By Amy 2013/05/03 Start
   '判斷ImgByteFile是否有資料,沒有請作單鈕設Disabled
   'Modify by Amy 2023/06/09 +if 放99年以前的附件會抓不到
   If Len(Text1(0)) = 7 Then
      m_ibf01 = Left(Text1(0), 3)
   Else
      m_ibf01 = "0" & Left(Text1(0), 2)
   End If
   m_ibf02 = Right(Text1(0), 4) & Text2
   
   strExc(0) = "Select  * From ImgByteFile Where IBF01='" & m_ibf01 & "' And IBF02='" & m_ibf02 & "' And IBF03='0' And IBF04='00'"
   If RsTemp.State <> adStateClosed Then RsTemp.Close
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If RsTemp.RecordCount > 0 Then
      CmdPaper.Enabled = True
   Else
      CmdPaper.Enabled = False
   End If
   
   '更新CUID
   UpdateCUID rsTmp
   '2013/05/03 End
End Sub

'Add By Amy 2013/04/30
Private Sub Form_Unload(Cancel As Integer)
    KillAttach
   Set frm140413 = Nothing
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
      Case 1 'Add
         Text1(0).SetFocus
         RbEdit 0
      Case 2 'Update
         Text1(1).SetFocus
         Text1(0).TabStop = False
         RbEdit 1
      Case 3 'Del
         RbEdit 2
      Case 4 'query
        TxtClear
        RbEdit 5
      Case 6 'MoveFirst
         ActionRb 0
      Case 7 'MovePrv
         ActionRb 1
      Case 8 'MoveNext
         ActionRb 2
      Case 9 'MoveLast
         ActionRb 3
      Case 11 'OK
        RbEdit 3
        Text1(0).TabStop = True
      Case 12 'Cancel
        RbEdit 4
      Case 14 'Exit
        Unload Me
        Set frm140413 = Nothing
       
   End Select
End Sub
Private Sub RbEdit(sta As Integer)
Dim StrSQLa, strBU01 As String
Dim SeqNo As Integer
Dim BulletinYN As Integer
Dim BuSystemKind As String
Dim NowTime As Long 'Modify by Amy 2013/07/31 改型態(原integer)
Dim nTime As String

    Select Case sta
      Case 0 'add
         TxtClear
         ToolBarSet 0
         CmdPaper.Enabled = False
         ActionEdit = 0
         Text1(0).SetFocus
         TextInverse Text1(0)
         
      Case 1 'update
         ToolBarSet 0
         ActionEdit = 1
         Text1(0).Locked = True
         Text1(1).SetFocus
         TextInverse Text1(1)
         
      Case 2 'delete
         If MsgBox("是否要刪除此筆資料?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            If DelRecord = True Then
                RefreshRange
            Else
                Exit Sub
            End If
         End If
      Case 3 'ok
        If ActionEdit = 0 Then  '在新增狀態按Enter鍵
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            If Option1(0).Value Then
               BulletinYN = 1
            Else
               BulletinYN = 0
            End If
             strBU01 = ChangeTStringToWString(Me.Text1(0).Text)
             SeqNo = GetSerialNo(ChangeTStringToWString(Me.Text1(0).Text))
             BuSystemKind = ChkSystemKind()
             nTime = ServerTime
             NowTime = IIf(Len(nTime) = 6, Left(nTime, 4), Left(nTime, 3))
            
             'Modify by Amy 2014/07/16 +BU15
             StrSQLa = "Insert Into PGMBulletin (BU01, BU02, BU03, BU04, BU05, BU06, BU07, BU08, BU09, BU10,BU14,BU15) " & _
                              "Values('" & strBU01 & "','" & SeqNo & "','" & Me.Text1(1).Text & "','" & ChangeTStringToWString(Me.Text1(2).Text) & "','" & ChgSQL(Me.Text1(3)) & "'," & _
                              "'" & BulletinYN & "','" & BuSystemKind & "','" & strUserNum & "','" & strSrvDate(1) & "','" & NowTime & "'," & _
                              "'" & ChgSQL(Me.Text1(4)) & "'," & Me.Text1(5) & ")"
             cnnConnection.Execute StrSQLa
             Text2.Text = IIf(SeqNo <= 9, Format(Val(SeqNo), "00"), SeqNo)
             GetCurrRecordVal strBU01, SeqNo
                       
         ElseIf ActionEdit = 1 Then '在修改狀態按Enter鍵
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            If Option1(0).Value Then
               BulletinYN = 1
            Else
               BulletinYN = 0
            End If
             BuSystemKind = ChkSystemKind()
             nTime = ServerTime
             NowTime = IIf(Len(nTime) = 6, Left(nTime, 4), Left(nTime, 3))
             'Modify by Amy 2014/07/16 +BU15
             StrSQLa = "Update PGMBulletin set BU03='" & Me.Text1(1).Text & "',BU04= '" & ChangeTStringToWString(Me.Text1(2).Text) & "',BU05='" & ChgSQL(Me.Text1(3)) & "', " & _
                             "BU06='" & BulletinYN & "', BU07='" & BuSystemKind & "',BU11='" & strUserNum & "', BU12='" & strSrvDate(1) & "', BU13='" & NowTime & "', " & _
                             "BU14='" & ChgSQL(Me.Text1(4)) & "', BU15=" & Val(Me.Text1(5)) & " " & _
                             "Where BU01='" & ChangeTStringToWString(Me.Text1(0).Text) & "' And BU02='" & Text2.Text & "'"
             cnnConnection.Execute StrSQLa
             RefreshRange
             SetTxtValue
             
         ElseIf ActionEdit = 2 Then '在查詢狀態按Enter鍵
            If Len(Trim(Text1(0))) > 0 And Len(Trim(Text2)) > 0 Then
                If CheckIsTaiwanDate(Me.Text1(0).Text) = False Then Text1(0).SetFocus: Exit Sub
                    QueryRecord Text1(0).Text, Text2.Text
            ElseIf Len(Trim(Text1(0))) = 0 Then
                MsgBox ("請輸入上線日期")
                Text1(0).SetFocus
                Exit Sub
            ElseIf Len(Trim(Text2)) = 0 Then
                MsgBox ("請輸入流水號")
                Text2.SetFocus
                Exit Sub
            End If
            
         End If
         ToolBarSet 1
         ActionEdit = 3
      Case 4 'cancel
       If ActionEdit <> 2 Then
         If MsgBox("妳並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
            If ActionEdit = 0 Then
               ActionRb 3
            ElseIf ActionEdit = 1 Then
               SetTxtValue
            End If
            ToolBarSet 1
            ActionEdit = 3 'cancel
            SetTxtValue
         Else
            Exit Sub
         End If
         Text1(0).SetFocus
        Else
        ToolBarSet 1
         ActionEdit = 3 'cancel
         SetTxtValue
        End If
       Case 5 'query
         ToolBarSet 0
         TxtLock 3
         ActionEdit = 2
         Text1(0).Locked = False
         Text1(0).SetFocus
         Text2.Locked = False
         Text2.Appearance = 1
         Text2.BorderStyle = 1
         
   End Select
End Sub

Private Function TxtValidate() As Boolean
   TxtValidate = False
   Dim TotalChk As Integer
  
   'Add by Sindy 2021/12/15 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Function
   End If
   
   If Text1(0) = "" Then MsgBox "上線日期不可為空值", vbInformation: Text1(0).SetFocus: Exit Function
   If CheckIsTaiwanDate(Me.Text1(0)) = False Then Text1(0).SetFocus: Exit Function
   If Not ChkWorkDay(DBDATE(Me.Text1(0))) Then MsgBox "上線日期必須是工作天 !", vbInformation: Text1(0).SetFocus: Exit Function
               
   If ActionEdit = 0 Or ActionEdit = 1 Then
      If Text1(1) = "" Then MsgBox "需求人員不可為空值", vbInformation: Text1(1).SetFocus: Exit Function
      If Label2(1).Caption = "" Then MsgBox "無此員工", vbInformation: Text1(1).SetFocus: Exit Function
      If Text1(2) = "" Then MsgBox "請作單日不可為空值", vbInformation: Text1(2).SetFocus: Exit Function
      If CheckIsTaiwanDate(Me.Text1(2).Text) = False Then Text1(2).SetFocus: Exit Function
      If ChkEndDate(Text1(0), Text1(2), "日期") = False Then Text1(2).SetFocus: Exit Function
      
      'Add by Amy 2014/07/16 +時數BU15必填
      If ActionEdit = 0 Or (ActionEdit = 1 And Val(ChangeTDateStringToTString(strCDate)) > 1030716) Then
         If Len(Trim(Me.Text1(5))) = 0 Then MsgBox "時數" & MsgText(52), vbInformation: Text1(5).SetFocus: Exit Function
      End If
      'end 2014/07/16
      
      If Text1(3) = "" Then MsgBox "摘要不可為空值", vbInformation: Text1(3).SetFocus: Exit Function
      'Modify By Amy 2013/05/10 判斷主旨及說明是否超過最大字數
      If CheckLengthIsOK(Text1(3), 200) = False Then Text1(3).SetFocus: Exit Function
      If CheckLengthIsOK(Text1(4), 1000) = False Then Text1(4).SetFocus: Exit Function
      
      'Modify By Amy 2013/04/23 說明不需必填
      'If Text1(4) = "" Then MsgBox "說明不可為空值", vbInformation: Text1(4).SetFocus: Exit Function
      If Me.Option1(0).Value = True Then
         '選擇公佈則系統別需至少勾選一項
         For Each oCheck In Check1
           If oCheck.Value = 1 Then
               Exit For
           ElseIf TotalChk = 15 Then 'Modify by Amy 2018/11/14 原：１４
               MsgBox "請輸入系統別", vbInformation: Check1(0).SetFocus: Exit Function
           Else
               TotalChk = TotalChk + 1
           End If
         Next
      End If
   End If
      
   TxtValidate = True
End Function

Private Sub TxtClear()
   Dim txt As Object, Lbl As Object, Chk As Object
   For Each txt In Text1
      txt.Text = ""
   Next
   Text2.Text = "" '流水號
   For Each Lbl In Label2
      Lbl = ""
   Next
   Option1(0).Value = 1
   For Each Chk In Check1
      Chk.Value = 0
   Next
   Label23 = Empty
End Sub
Private Sub ToolBarSet(ByVal sta As Integer)
 Dim i As Integer, txt As Object
 Select Case sta
    Case 0
        TxtLock 1
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
    Case 1
        TxtLock 0
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = True
         TBar1.Buttons(i + 5).Enabled = True
      Next
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
    End Select
   
End Sub
Private Sub TxtLock(ByVal Lt As Integer)
 Dim txt As Object, Chk As Object, i As Integer
   Select Case Lt
      Case 0 'cancel/OK
         For Each txt In frm140413.Text1
            txt.Locked = True
            txt.Enabled = True
         Next
         Text2.Locked = True
         Text2.Appearance = 0
         Text2.BorderStyle = 0
         FrmYN.Enabled = False
         FrmSysKind.Enabled = False
         CmdSelect(0).Visible = False
         CmdSelect(1).Visible = False
         
      Case 1 'add/upd
         For Each txt In frm140413.Text1
            txt.Locked = False
         Next
         Text2.Locked = True
         Text2.Appearance = 0
         Text2.BorderStyle = 0
         FrmYN.Enabled = True
         FrmSysKind.Enabled = True
         CmdSelect(0).Visible = True
         CmdSelect(1).Visible = True
         
      Case 2 '第一次進入
         For Each txt In frm140413.Text1
            txt.Locked = True
         Next
         Text2.Locked = True
         Text2.Appearance = 0
         Text2.BorderStyle = 0
         FrmYN.Enabled = False
         FrmSysKind.Enabled = False
         CmdSelect(0).Visible = False
         CmdSelect(1).Visible = False
         'TxtClear
         
      Case 3 'query
        Text2.Text = "01" '查詢時預設流水號01
        Text1(1).Enabled = False
        Text1(2).Enabled = False
        Text1(3).Enabled = False
        FrmYN.Enabled = False
        FrmSysKind.Enabled = False
        CmdSelect(0).Visible = False
        CmdSelect(1).Visible = False
      End Select
End Sub

'取得序號
Private Function GetSerialNo(strBU01 As String) As String
  Dim StrSQLa As String
  Dim rsA As New ADODB.Recordset

  '抓PGMBulletin(程式修改公告)流水號 每日重編
  StrSQLa = "SELECT NVL(MAX(BU02),0) as BU02 From PGMBulletin Where BU01='" & strBU01 & "'  Order By BU02 Desc "
  rsA.CursorLocation = adUseClient
  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly

  If Not rsA.EOF And Not rsA.BOF Then
      GetSerialNo = Format(Val(rsA("BU02").Value) + 1, "00")
  Else
      GetSerialNo = "01"
  End If
  If rsA.State <> adStateClosed Then rsA.Close
  Set rsA = Nothing
End Function

Private Sub RefreshRange()

Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   strSql = "Select BU01,BU02 From PGMBulletin " & _
               "Where BU01 = (Select MIN(BU01) From PGMBulletin) AND " & _
                          "BU02 = (Select MIN(BU02) From PGMBulletin " & _
                                         "Where BU01 = (Select MIN(BU01) FROM PGMBulletin)) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("BU01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("BU01")
      If IsNull(rsTmp.Fields("BU02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("BU02")
   End If
   rsTmp.Close

    strSql = "Select BU01,BU02 From PGMBulletin " & _
               "Where BU01 = (Select MAX(BU01) From PGMBulletin) AND " & _
                          "BU02 = (Select MAX(BU02) From PGMBulletin " & _
                                         "Where BU01 = (Select MAX(BU01) FROM PGMBulletin)) "
 
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("BU01")) = False Then: m_LastKEY(0) = rsTmp.Fields("BU01")
      If IsNull(rsTmp.Fields("BU02")) = False Then: m_LastKEY(1) = rsTmp.Fields("BU02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "Select * From PGMBulletin " & _
                "Where BU01 = '" & strKEY01 & "' AND " & _
                          "BU02 = '" & strKEY02 & "' "
                  
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
Private Sub GetCurrRecordVal(ByVal strKEY01 As String, ByVal strKEY02 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "Select BU01,BU02 From PGMBulletin " & _
                  "Where BU01 = '" & m_CurrKEY(0) & "' AND " & _
                             "BU02 = (Select MIN(BU02) From PGMBulletin " & _
                                          "Where BU01 = '" & m_CurrKEY(0) & "' AND " & _
                                                    "BU02 > '" & m_CurrKEY(1) & "' )"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("BU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("BU01")
         If IsNull(rsTmp.Fields("BU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("BU02")
         rsTmp.Close
         RefreshRange
         SetTxtValue
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "Select BU01,BU02 From PGMBulletin " & _
                  "Where BU01 = (Select MIN(BU01) From PGMBulletin " & _
                                         "Where BU01 > '" & m_CurrKEY(0) & "') AND " & _
                                                    "BU02 = (Select MIN(BU02) From PGMBulletin " & _
                                                                 "Where BU01 = (Select MIN(BU01) From PGMBulletin " & _
                                                                                        "Where BU01 > '" & m_CurrKEY(0) & "')) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("BU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("BU01")
         If IsNull(rsTmp.Fields("BU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("BU02")
      Else
         GetLastRecordVal
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   RefreshRange
   SetTxtValue
EXITSUB:
End Sub
' 第一筆資料
Private Sub GetFirstRecordVal()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   
   SetTxtValue
End Sub
'上一筆資料
Private Sub GetPreRecordVal()
    Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "Select BU01,BU02 From PGMBulletin " & _
               "Where BU01 = '" & m_CurrKEY(0) & "' And " & _
                         "BU02 = (Select MAX(BU02) From PGMBulletin " & _
                                        "Where BU01 = '" & m_CurrKEY(0) & "' AND " & _
                                                   "BU02 < '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("BU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("BU01")
      If IsNull(rsTmp.Fields("BU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("BU02")
      rsTmp.Close
      SetTxtValue
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "Select BU01,BU02 From PGMBulletin " & _
               "Where BU01 = (Select MAX(BU01) From PGMBulletin " & _
                                       "Where BU01 < '" & m_CurrKEY(0) & "') AND " & _
                                       "BU02 = (Select MAX(BU02) From PGMBulletin " & _
                                       "Where BU01 = (Select MAX(BU01) From PGMBulletin " & _
                                       "Where BU01 < '" & m_CurrKEY(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("BU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("BU01")
      If IsNull(rsTmp.Fields("BU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("BU02")
   End If
   rsTmp.Close
   SetTxtValue
   
EXITSUB:
   Set rsTmp = Nothing
End Sub
'下一筆資料
Private Sub GetNextRecordVal()
    Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "Select BU01,BU02 From PGMBulletin " & _
               "Where BU01 = '" & m_CurrKEY(0) & "' AND " & _
                         "BU02 = (Select MIN(BU02) From PGMBulletin " & _
                                       "Where BU01 = '" & m_CurrKEY(0) & "' AND " & _
                                                  "BU02 > '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("BU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("BU01")
      If IsNull(rsTmp.Fields("BU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("BU02")
      rsTmp.Close
      SetTxtValue
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "Select BU01,BU02 From PGMBulletin " & _
            "Where BU01 = (Select MIN(BU01) From PGMBulletin " & _
                                    "Where BU01 > '" & m_CurrKEY(0) & "') AND " & _
                                              "BU02 = (Select MIN(BU02) FROM customer " & _
                                                           "Where BU01 = (Select MIN(BU01) From PGMBulletin " & _
                                                                                     "Where BU01 > '" & m_CurrKEY(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("BU01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("BU01")
      If IsNull(rsTmp.Fields("BU02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("BU02")
   End If
   rsTmp.Close
   
   SetTxtValue
   
EXITSUB:
   Set rsTmp = Nothing
End Sub
' 最後一筆資料
Private Sub GetLastRecordVal()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   SetTxtValue
End Sub

' 查詢記錄
Private Function QueryRecord(ByVal strBU01 As String, ByVal strBU02 As String) As Boolean

   QueryRecord = False
   strBU01 = ChangeTStringToWString(strBU01)
   strBU02 = Val(strBU02)
  
   If IsRecordExist(strBU01, strBU02) = True Then
      m_CurrKEY(0) = strBU01
      m_CurrKEY(1) = strBU02
      QueryRecord = True
      
   Else
      QueryRecord = False
      MsgBox ("無此資料")
   End If
    SetTxtValue
   ToolBarSet 1
End Function

Private Function ChkSystemKind() As String
   Dim returnVal As String

    For Each oCheck In Check1
      idx = oCheck.Index
           
      If oCheck.Value Then
      Select Case idx
        Case 0
          returnVal = "Account"
        Case 1
          returnVal = "Finance"
        Case 2
          returnVal = "Casher"
        Case 3
          returnVal = "Salary"
        Case 4
          returnVal = "Query"
        Case 5
          returnVal = "Patpro"
        Case 6
          returnVal = "Patpro1"
        Case 7
          returnVal = "Trademark"
        Case 8
          returnVal = "Law"
        Case 9
          returnVal = "Promoter"
        Case 10
          returnVal = "Writer"
        Case 11
          returnVal = "File"
        Case 12
          returnVal = "Person"
        Case 13
          returnVal = "Computer"
        Case 14
          returnVal = "AutoBatch"
        'Add by Amy 2018/11/14
        Case 15
          returnVal = "Trademark1"
      End Select
           ChkSystemKind = ChkSystemKind & returnVal & ","
      End If
    Next
    'Modify By Amy 2013/04/22 改存系統別,
    'ChkSystemKind = Mid(ChkSystemKind, 2) '去掉第一個,
End Function
Private Sub RedSystemKind(ByVal SysKind As String)
  If Trim(SysKind) <> "" Then
     Dim Chk As CheckBox
     Dim strTmp() As String
     
    For Each Chk In Check1
      Chk.Value = 0
    Next
    strTmp = Split(SysKind, ",")
     
     For i = 0 To UBound(strTmp)
        Select Case strTmp(i)
          Case "Account"
            Check1(0).Value = 1
          Case "Finance"
            Check1(1).Value = 1
          Case "Casher"
            Check1(2).Value = 1
          Case "Salary"
            Check1(3).Value = 1
          Case "Query"
            Check1(4).Value = 1
          Case "Patpro"
            Check1(5).Value = 1
          Case "Patpro1"
            Check1(6).Value = 1
          Case "Trademark"
           Check1(7).Value = 1
          Case "Law"
            Check1(8).Value = 1
          Case "Promoter"
            Check1(9).Value = 1
          Case "Writer"
            Check1(10).Value = 1
          Case "File"
            Check1(11).Value = 1
          Case "Person"
            Check1(12).Value = 1
          Case "Computer"
            Check1(13).Value = 1
          Case "AutoBatch"
            Check1(14).Value = 1
          'Add by Amy 2018/11/14
          Case "Trademark1"
           Check1(15).Value = 1
      End Select
   Next
  End If
End Sub

Private Sub Text1_Change(Index As Integer)
    Select Case Index
     Case 1   '輸入員編帶姓名及部門
       If ActionEdit = 0 Or ActionEdit = 1 Then
        
         If Len(Me.Text1(1)) = 5 Then
            If Me.Text1(1).Tag <> Me.Text1(1).Text Then  'Added by Lydia 2020/04/27 判斷有修改才抓資料
                'Modified by Lydia 2016/08/17 改用模組
                'Modify By Sindy 2016/9/5 解開Mark,恢復原程式
                If bp.State = adStateOpen Then bp.Close
                  'Added by Lydia 2023/12/27
                  If strSrvDate(1) >= 新部門啟用日 Then
                     strExc(1) = "SELECT ST02,NVL(A0922,A0902) AS A0902 FROM STAFF,ACC090,ACC090NEW WHERE ST01='" & Text1(1) & "' And ST03=A0901 AND ST93=A0921(+) And ST04<>2"
                  Else
                  'end 2023/12/27
                     strExc(1) = "SELECT ST02,A0902 FROM STAFF,ACC090 WHERE ST01='" & Text1(1) & "' And ST03=A0901 And ST04<>2"
                  End If
                  bp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
    
                 If bp.BOF And bp.EOF Then
                    Label2(1).Caption = ""
                    Label2(2).Caption = ""
                    MsgBox ("無此員工")
                    Text1_GotFocus 1
                 Else
                   If IsNull(bp.Fields(0).Value) Then
                        Label2(1).Caption = ""
                    Else
                        Label2(1).Caption = bp.Fields(0).Value
                    End If
                    If IsNull(bp.Fields(1).Value) Then
                        Label2(2).Caption = ""
                    Else
                        Label2(2).Caption = bp.Fields(1).Value
                    End If
                 End If
               bp.Close
    '            Label2(1).Caption = GetStaffName(Me.Text1(1), True, strExc(1))
    '            Label2(2).Caption = strExc(1)
           End If 'Added by Lydia 2020/04/27
         Else
            Label2(1).Caption = ""
            Label2(2).Caption = ""
         End If
       End If
       Text1(1).Tag = Text1(1).Text  'Added by Lydia 2020/04/27 記錄員工編號
    End Select
    
   'Added by Morgan 2022/3/18
   If Index = 3 Or Index = 4 Then
      PUB_RefreshText Text1(Index)
   End If
   'end 2022/3/18
End Sub

'反白
Public Sub TextInverse(ByRef txtTemp As Object)
txtTemp.SelStart = 0
txtTemp.SelLength = Len(txtTemp.Text)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   If ActionEdit <> 3 Then
        TextInverse Text1(Index)
   End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Select Case Index
    Case 1
        KeyAscii = UpperCase(KeyAscii)
    End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Select Case Index
     Case 0
         If Len(Trim(Me.Text1(0).Text)) > 0 Then
            If CheckIsTaiwanDate(Me.Text1(0)) = False Then
               Text1(0).SetFocus
               TextInverse Text1(0)
                Exit Sub
            End If
            If Not ChkWorkDay(DBDATE(Me.Text1(0))) Then
                MsgBox "上線日期必須是工作天 !"
                Text1(0).SetFocus
                TextInverse Text1(0)
                Exit Sub
            End If
          End If
      Case 1   '輸入員編帶姓名及部門
       If (ActionEdit = 0 Or ActionEdit = 1) And Len(Me.Text1(1)) > 0 Then
         If Len(Me.Text1(1)) = 5 Then
            If Me.Text1(1).Tag <> Me.Text1(1).Text Then  'Added by Lydia 2020/04/27 判斷有修改才抓資料
                'Modified by Lydia 2016/08/17 改用模組
                'Modify By Sindy 2016/9/5 解開Mark,恢復原程式
                If bp.State = adStateOpen Then bp.Close
                  'Added by Lydia 2023/12/27
                  If strSrvDate(1) >= 新部門啟用日 Then
                     strExc(1) = "SELECT ST02,NVL(A0922,A0902) AS A0902 FROM STAFF,ACC090,ACC090NEW WHERE ST01='" & Text1(1) & "' And ST03=A0901 AND ST93=A0921(+) And ST04<>2"
                  Else
                  'end 2023/12/27
                     strExc(1) = "SELECT ST02,A0902 FROM STAFF,ACC090 WHERE ST01='" & Text1(1) & "' And ST03=A0901 And ST04<>2"
                  End If
                  bp.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
    
                 If bp.BOF And bp.EOF Then
                    Label2(1).Caption = ""
                    Label2(2).Caption = ""
                    MsgBox ("無此員工")
                    Text1(1).SetFocus
                    TextInverse Text1(1)
                    Exit Sub
                 Else
                   If IsNull(bp.Fields(0).Value) Then
                        Label2(1).Caption = ""
                    Else
                        Label2(1).Caption = bp.Fields(0).Value
                    End If
                    If IsNull(bp.Fields(1).Value) Then
                        Label2(2).Caption = ""
                    Else
                        Label2(2).Caption = bp.Fields(1).Value
                    End If
                 End If
                bp.Close
    '            Label2(1).Caption = GetStaffName(Me.Text1(1), True, strExc(1))
    '            Label2(2).Caption = strExc(1)
            End If 'Added by Lydia 2020/04/27
         Else
            Label2(1).Caption = ""
            Label2(2).Caption = ""
            MsgBox ("無此員工")
            Text1(1).SetFocus
            TextInverse Text1(1)
           Exit Sub
         End If
        End If
        Text1(1).Tag = Text1(1).Text  'Added by Lydia 2020/04/27 記錄員工編號
      Case 2
         If Len(Trim(Me.Text1(2).Text)) > 0 Then
            If CheckIsTaiwanDate(Me.Text1(2).Text) = False Then
               Text1(2).SetFocus
               TextInverse Text1(2)
                Exit Sub
            ElseIf ChkEndDate(Text1(0), Text1(2), "日期") = False Then
                Text1(2).SetFocus
                TextInverse Text1(2)
                Exit Sub
            End If
          End If
         'Add by Amy +時數BU15
         If Len(Trim(Me.Text1(5))) > 0 Then
            If Not IsNumeric(Me.Text1(5)) Then
                MsgBox "時數" & MsgText(63), , MsgText(5)
                Text1(5).SetFocus
                TextInverse Text1(5)
                Exit Sub
            End If
         End If
            
   End Select
End Sub

Public Function ChkEndDate(txt1 As Object, txt2 As Object, St As String) As Boolean
Dim sss
If Val(txt2.Text) > Val(txt1.Text) And txt2.Text <> "" Then
   sss = MsgBox(St & "區間錯誤", , "錯誤！")
   ChkEndDate = False
Else
   ChkEndDate = True
End If
End Function

Private Sub Select_All(inD As Integer)
    Dim Chk As CheckBox
    For Each Chk In Check1
       Chk.Value = inD
    Next
End Sub

'Add By Amy 2013/05/03 刪除記錄
Private Function DelRecord() As Boolean
    Dim strSQLD As String, m_bu01 As String, m_bu02 As String, m_ibf01 As String, m_ibf02 As String
    DelRecord = False
On Error GoTo ErrHand
    cnnConnection.BeginTrans
    m_bu01 = ChangeTStringToWString(Me.Text1(0))
    m_bu02 = Val(Text2)
    strSQLD = "Delete PGMBulletin Where BU01='" & m_bu01 & "' And BU02='" & m_bu02 & "'"
    cnnConnection.Execute strSQLD
    
    'Add by Amy 2013/05/03 刪除ImgByteFile
    m_ibf01 = Left(Me.Text1(0), 3)
    m_ibf02 = Right(Me.Text1(0), 4) & Text2
    
    PUB_DelFtpFile2 m_ibf01 & "-" & m_ibf02 & "-0-00-5", , UCase("ImgByteFile") 'Add By Sindy 2017/8/10 檔案改放 FTP,必須在DB資料刪除前執行
    strSQLD = "Delete ImgByteFile Where IBF01='" & m_ibf01 & "' And IBF02='" & m_ibf02 & "' And IBF03='0' And IBF04='00'"
    cnnConnection.Execute strSQLD
        
    ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (m_bu01 = m_LastKEY(0) And m_bu02 = m_LastKEY(1)) Or (m_bu01 = m_FirstKEY(0) And m_bu02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
    GetCurrRecordVal m_bu01, m_bu02
    DelRecord = True
    cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

'Add By Amy 2013/04/30 pFileName為上線日(民國年月日+序號)
Private Function GetAttachFile(ByVal m_AttachPath As String, ByVal pFileName As String) As Boolean
   
   Dim stAttPath As String, m_ibf01 As String, m_ibf02 As String
   m_ibf01 = "": m_ibf02 = ""
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   
On Error GoTo ErrHnd
   'Modify by Amy 2023/06/09 +if 放99年以前的附件會抓不到
   'Modify By Sindy 2023/6/12 pFileName=年度+6碼流水號
   'If Len(pFileName) = 7 Then
   If Len(pFileName) = 9 Then
   '2023/6/12 END
      m_ibf01 = Left(pFileName, 3)
   Else
      m_ibf01 = "0" & Left(pFileName, 2)
   End If
   m_ibf02 = Right(pFileName, 6)
    If Dir(m_AttachPath, vbDirectory) = "" Then
        MkDir m_AttachPath
    End If
    stAttPath = m_AttachPath & "\" & pFileName & ".pdf"
    '檔案已存在時不必重新下載
    If Dir(stAttPath) <> "" Then
        pFileName = stAttPath
        GetAttachFile = True
        Exit Function
    End If
      
   strExc(0) = "Select * From ImgByteFile Where IBF01='" & m_ibf01 & "' And IBF02='" & m_ibf02 & "' And IBF03='0' And IBF04='00' And IBF05='5' "
   If RsTemp.State <> adStateClosed Then RsTemp.Close
   RsTemp.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
   If RsTemp.RecordCount > 0 Then
      If Dir(stAttPath) <> "" Then Kill stAttPath
      'Add By Sindy 2017/8/10
'      If "" & RsTemp.Fields("IBF15") <> "" Then
         GetAttachFile = PUB_GetFtpFile(RsTemp.Fields("IBF15"), stAttPath, UCase("ImgByteFile"))
'      Else
'      '2017/8/10 END
'         With RsTemp
'            lngSize = Val(.Fields("IBF13").Value)
'            ReDim bytes(lngSize)
'            If lngSize > 0 Then bytes() = .Fields("IBF14").GetChunk(lngSize)
'         End With
'         iFileNo = FreeFile
'         Open stAttPath For Binary Access Write As #iFileNo
'         If lngSize > 0 Then Put #iFileNo, , bytes()
'         Close #iFileNo
'         GetAttachFile = True
'      End If
      pFileName = stAttPath
   Else
      Close #iFileNo
      MsgBox ("無此請作單資料")
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

'Add By Amy 2013/04/30
Private Sub KillAttach()
    Dim strPath As String '防刪到c:\
On Error Resume Next
    strPath = App.path & "\PGMBulletinAttach"
   If Dir(strPath & "\.") <> "" Then
      Kill strPath & "\*.*"
   End If
End Sub

'Add By Amy 2013/05/03 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   strCDate = ""
   If IsNull(rsSrcTmp.Fields("BU08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU08")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("BU08"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("BU09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU09")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("BU09"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("BU10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU10")) = False Then
         strTemp = rsSrcTmp.Fields("BU10")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("BU11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU11")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("BU11"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("BU12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU12")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("BU12"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("BU13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("BU13")) = False Then
         strTemp = rsSrcTmp.Fields("BU13")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

