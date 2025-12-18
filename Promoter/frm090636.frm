VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090636 
   BorderStyle     =   1  '單線固定
   Caption         =   "價目表資料維護"
   ClientHeight    =   5460
   ClientLeft      =   6090
   ClientTop       =   1550
   ClientWidth     =   9140
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9140
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   30
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
            Picture         =   "frm090636.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090636.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090636.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090636.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090636.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090636.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090636.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090636.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090636.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090636.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090636.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   60
      TabIndex        =   12
      Top             =   720
      Width           =   9015
      _ExtentX        =   15893
      _ExtentY        =   8273
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm090636.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label16(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label23"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label16(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text1(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textDate(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdRemAtt(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdAddAtt(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdOpenAtt(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lstAtt(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textDate(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm090636.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textDate(1)"
      Tab(1).Control(1)=   "textDate(0)"
      Tab(1).Control(2)=   "Text1(0)"
      Tab(1).Control(3)=   "cmdOK"
      Tab(1).Control(4)=   "GRD1"
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(6)=   "Label16(0)"
      Tab(1).Control(7)=   "Line5(0)"
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(9)=   "Label6"
      Tab(1).Control(10)=   "Label5"
      Tab(1).Control(11)=   "Label1(4)"
      Tab(1).ControlCount=   12
      Begin VB.TextBox textDate 
         Height          =   270
         Index           =   3
         Left            =   3480
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1140
         Width           =   915
      End
      Begin VB.ListBox lstAtt 
         Height          =   580
         Index           =   0
         ItemData        =   "frm090636.frx":212C
         Left            =   1620
         List            =   "frm090636.frx":2133
         Sorted          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1620
         Width           =   7140
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟附件"
         Height          =   345
         Index           =   0
         Left            =   3930
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2250
         Width           =   1065
      End
      Begin VB.CommandButton cmdAddAtt 
         Caption         =   "新增附件"
         Height          =   345
         Index           =   0
         Left            =   1650
         TabIndex        =   4
         Top             =   2250
         Width           =   1065
      End
      Begin VB.CommandButton cmdRemAtt 
         Caption         =   "刪除附件"
         Height          =   345
         Index           =   0
         Left            =   2790
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2250
         Width           =   1065
      End
      Begin VB.TextBox textDate 
         Height          =   270
         Index           =   2
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   1
         Top             =   1140
         Width           =   915
      End
      Begin VB.TextBox textDate 
         Height          =   270
         Index           =   1
         Left            =   -72780
         MaxLength       =   7
         TabIndex        =   9
         Top             =   1110
         Width           =   915
      End
      Begin VB.TextBox textDate 
         Height          =   270
         Index           =   0
         Left            =   -73830
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1110
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   0
         Top             =   570
         Width           =   465
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   -73830
         MaxLength       =   2
         TabIndex        =   7
         Top             =   450
         Width           =   465
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "查詢"
         Height          =   400
         Left            =   -68310
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm090636.frx":213F
         Height          =   2990
         Left            =   -74910
         TabIndex        =   18
         Top             =   1410
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   5274
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "系統類別|啟用日期|新增人員|日期|時間|修改人員|日期|時間"
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
         _Band(0).Cols   =   8
      End
      Begin VB.Label Label9 
         Caption         =   "註：資料列點二下才會查詢明細資料。"
         ForeColor       =   &H000000C0&
         Height          =   250
         Left            =   -74880
         TabIndex        =   28
         Top             =   4410
         Width           =   4690
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "下架日期："
         Height          =   180
         Index           =   3
         Left            =   2580
         TabIndex        =   27
         Top             =   1170
         Width           =   900
      End
      Begin MSForms.Label Label23 
         Height          =   195
         Left            =   570
         TabIndex        =   26
         Top             =   4410
         Width           =   7905
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13944;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         Caption         =   "註：電子檔命名規則為啟用日期（民國）＋系統類別中文＋價目表.PDF"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3300
         TabIndex        =   25
         Top             =   2670
         Width           =   5565
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "價目表附件檔案："
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   24
         Top             =   1650
         Width           =   1440
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "啟用日期："
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   1170
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "啟用日期："
         Height          =   180
         Index           =   0
         Left            =   -74730
         TabIndex        =   22
         Top             =   1140
         Width           =   900
      End
      Begin VB.Line Line5 
         Index           =   0
         X1              =   -73110
         X2              =   -72510
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Label Label4 
         Caption         =   "（01.國內專利   02.大陸專利       03.香港澳門專利   04.CFP"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   1590
         TabIndex        =   21
         Top             =   390
         Width           =   5805
      End
      Begin VB.Label Label3 
         Caption         =   "　05.國內商標   06.大陸商標       07.馬德里商標       08.國內著作權   09.大陸著作權"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   1590
         TabIndex        =   20
         Top             =   630
         Width           =   6855
      End
      Begin VB.Label Label2 
         Caption         =   "　10.CFT            11.美國著作權   12.顧問及法務）"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   1590
         TabIndex        =   19
         Top             =   870
         Width           =   5805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "系統類別："
         Height          =   180
         Index           =   5
         Left            =   180
         TabIndex        =   17
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label7 
         Caption         =   "　10.CFT            11.美國著作權   12.顧問及法務）"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -73320
         TabIndex        =   16
         Top             =   870
         Width           =   5805
      End
      Begin VB.Label Label6 
         Caption         =   "　05.國內商標   06.大陸商標       07.馬德里商標       08.國內著作權   09.大陸著作權"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -73320
         TabIndex        =   15
         Top             =   630
         Width           =   6855
      End
      Begin VB.Label Label5 
         Caption         =   "（01.國內專利   02.大陸專利       03.香港澳門專利   04.CFP"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -73320
         TabIndex        =   14
         Top             =   390
         Width           =   5805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "系統類別："
         Height          =   180
         Index           =   4
         Left            =   -74730
         TabIndex        =   13
         Top             =   510
         Width           =   900
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9140
      _ExtentX        =   16122
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
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
End
Attribute VB_Name = "frm090636"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/1 Form2.0已修改
'Create By Sindy 2014/3/3
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
' 第一筆資料的系統類別、啟用日期
Dim m_FirstKEY(2) As String
' 最後一筆資料的系統類別、啟用日期
Dim m_LastKEY(2) As String
' 目前正在顯示的系統類別、啟用日期
Dim m_CurrKEY(2) As String
Dim rsA As New ADODB.Recordset
Dim tf_PLF As Integer
Dim m_AttachPath As String
Private Declare Function SendMessageByNum Lib "user32" _
   Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
   wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
Dim m_strCName As String, m_strCDate As String, m_strCTime As String


Private Sub cmdOK_Click()
Dim rsTmp As New ADODB.Recordset
Dim strSQLCon As String
   
   If Text1(0) <> "" Then
      strSQLCon = strSQLCon & " and plf01='" & Text1(0) & "'"
   'Modify By Sindy 2024/8/5
   Else
      strSQLCon = strSQLCon & " and plf01<='12'"
   '2024/8/5 END
   End If
   If textDate(0) <> "" Then
      strSQLCon = strSQLCon & " and plf02>=" & DBDATE(textDate(0))
   End If
   If textDate(1) <> "" Then
      strSQLCon = strSQLCon & " and plf02<=" & DBDATE(textDate(1))
   End If
   'Modify By Sindy 2024/8/5 +plf04
   strSql = "SELECT decode(plf01,'01','國內專利','02','大陸專利','03','香港澳門專利','04','CFP','05','國內商標','06','大陸商標','07','馬德里商標','08','國內著作權','09','大陸著作權','10','CFT','11','美國著作權','12','顧問及法務',plf01)" & _
            ",sqldatet(plf02),plf01,s1.st02,sqldatet(plf06),sqltime(plf07),s2.st02,sqldatet(plf09),sqltime(plf10),sqldatet(plf04)" & _
            " FROM pricelistfile,staff s1,staff s2 where plf05=s1.st01(+) and plf08=s2.st01(+)" & strSQLCon & " order by plf02 desc,plf01 asc"
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set GRD1.Recordset = rsTmp
   SetGrd
   If rsTmp.RecordCount = 0 Then
      GRD1.Rows = 2
      GRD1.row = 1
      GRD1.col = 0
      MsgBox "無此資料", vbOKOnly, "查詢資料"
   End If
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from pricelistfile where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_PLF = rsA.Fields.Count
   SetGrd
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
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
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
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

Private Sub Form_Load()
   ReDim m_FieldList(tf_PLF) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   
   ClearField
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum
   
'   InitialField
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   KillAttach
   Set frm090636 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub GRD1_DblClick()
Dim tmpMouseRow
   
   tmpMouseRow = GRD1.row
   If tmpMouseRow <> 0 Then
      GRD1.row = tmpMouseRow
      GRD1.col = 0
      If GRD1.CellBackColor = &HFFC0C0 Then
         Text1(1).Text = GRD1.TextMatrix(tmpMouseRow, 2)
         textDate(2).Text = Replace(GRD1.TextMatrix(tmpMouseRow, 1), "/", "")
         QueryRecord
      End If
   End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow GRD1, x, y, nCol, nRow
   If nRow < 0 Then nRow = 0
   GRD1.col = nCol
   GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j

   tmpMouseRow = GRD1.row
   If tmpMouseRow <> 0 Then
      GRD1.row = tmpMouseRow
      GRD1.col = 0
      If GRD1.CellBackColor <> &HFFC0C0 Then
         GRD1.Visible = False
         For j = 1 To GRD1.Rows - 1
            GRD1.row = j
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               GRD1.CellBackColor = QBColor(15)
            Next i
         Next j
         GRD1.row = tmpMouseRow
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
         GRD1.Visible = True
      End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 1 Then
      If GRD1.Rows - 1 >= 1 Then
         If GRD1.TextMatrix(1, 0) <> "" Then '有查出資料時
            Call cmdOK_Click
         End If
      End If
   End If
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If Me.Text1(1).Enabled = True Then
      Cancel = False
      Text1_Validate 1, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textDate(2).Enabled = True Then
      Cancel = False
      textDate_Validate 2, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Add By Sindy 2024/8/5
   If Me.textDate(3).Enabled = True Then
      Cancel = False
      textDate_Validate 3, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2024/8/5 END
   
   TxtValidate = True
End Function

'' 設定欄位的內容
'Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
'Dim nIndex As Integer
'
'   For nIndex = 0 To tf_PLF - 1
'      If strName = m_FieldList(nIndex).fiName Then
'         If strData = "#==#" Then
'            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
'         Else
'            m_FieldList(nIndex).fiNewData = strData
'         End If
'         Exit For
'      End If
'   Next nIndex
'End Sub
'
'' 從記錄中更新欄位內容
'Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
'Dim nIndex As Integer
'Dim strTmp As String
'
'   For nIndex = 0 To tf_PLF - 1
'      If m_FieldList(nIndex).fiName <> Empty Then
'         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
'            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
'            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
'         Else
'            m_FieldList(nIndex).fiOldData = Empty
'            m_FieldList(nIndex).fiNewData = Empty
'         End If
'      End If
'   Next nIndex
'EXITSUB:
'End Sub

' 新增記錄
Private Function AddRecord() As Boolean
Dim strPLF01 As String, strPLF02 As String
   
   AddRecord = False
   
   strPLF01 = Text1(1)
   strPLF02 = DBDATE(textDate(2))
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   If SaveAttFile(strPLF01, strPLF02, 0) = False Then GoTo ErrHand
   
   cnnConnection.CommitTrans
   
   If (strPLF01 & strPLF02 < m_FirstKEY(0) & m_FirstKEY(1)) Or (strPLF01 & strPLF02 > m_LastKEY(0) & m_LastKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strPLF01, strPLF02
   
   AddRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

Private Function SaveAttFile(strKEY01 As String, strKEY02 As String, Index As Integer) As Boolean
Dim ii As Integer, jj As Integer
Dim stFilePath As String
Dim iFileNo As Integer
Dim bytes() As Byte
Dim lngSize As Long '檔案大小
Dim adoRst As New ADODB.Recordset
Const BlockSize = 500000
Dim Numblocks As Integer
Dim LeftOver As Long
Dim stReName As String, strFtpPath As String 'Add By Sindy 2017/5/31
   
   SaveAttFile = True
   
   For ii = 0 To 0 'lstAtt(Index).ListCount - 1
      If lstAtt(Index).ItemData(ii) = 0 Then
         stFilePath = lstAtt(Index).List(ii)
         stFilePath = Left(stFilePath, InStrRev(stFilePath, " (") - 1)
         If iFileNo > 0 Then Close #iFileNo
         iFileNo = FreeFile
         Open stFilePath For Binary Access Read As #iFileNo
         lngSize = LOF(iFileNo)
         Close #iFileNo
         If lngSize = 0 Then
            SaveAttFile = False
            ShowMsg stFilePath & MsgText(9221)
            Exit Function
         End If
         
         'Add By Sindy 2017/5/31
         '上傳FTP File Server
         stReName = strKEY02 & "." & lngSize & ".pdf"
         PUB_PutFtpFile stFilePath, strKEY01, stReName, strFtpPath, "PRICELISTFILE"
         If strFtpPath <> "" Then
            strSql = "insert into PRICELISTFILE(PLF01,PLF02,PLF03,PLF11) " & _
                     "values(" & CNULL(strKEY01) & "," & strKEY02 & _
                     "," & lngSize & "," & CNULL(strFtpPath) & ")"
            cnnConnection.Execute strSql
         End If
         'Call PUB_DelPCOrgFile(stFilePath) '一併將PC上的實體檔案刪除
'         '2017/5/31 END
'         With adoRst
'            If adoRst.State = adStateClosed Then
'               strExc(0) = "select * from pricelistfile where rownum<1"
'               .CursorLocation = adUseClient
'               .Open strExc(0), cnnConnection, adOpenStatic, adLockOptimistic
'            End If
'            .AddNew
'            .Fields("plf01").Value = strKEY01
'            .Fields("plf02").Value = strKEY02
'            .Fields("plf03").Value = lngSize
'            Numblocks = lngSize / BlockSize
'            LeftOver = lngSize Mod BlockSize
'
'            ReDim bytes(LeftOver)
'            Get #iFileNo, , bytes()
'            .Fields("plf04").AppendChunk bytes()
'            ReDim bytes(BlockSize)
'            For jj = 1 To Numblocks
'                Get #iFileNo, , bytes()
'                .Fields("plf04").AppendChunk bytes()
'            Next jj
'
'            Close #iFileNo
'            .UPDATE
'         End With
      End If
   Next ii
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim strPLF01 As String, strPLF02 As String
Dim stFileName As String
   
   ModRecord = False
   
   If Text1(1) <> m_CurrKEY(0) And DBDATE(textDate(2)) <> m_CurrKEY(1) Then
      MsgBox "系統記錄的目前系統類別（" & m_CurrKEY(0) & "）及啟用日期（" & m_CurrKEY(1) & "）與畫面上的系統類別及啟用日期不一致，請重新確認！"
      Exit Function
   End If
   
   strPLF01 = m_CurrKEY(0)
   strPLF02 = m_CurrKEY(1)
   
   '產生附件
   stFileName = lstAtt(0).List(0)
   If InStr(stFileName, "\") = 0 Then
      If GetAttachFile(stFileName, , 1) = False Then Exit Function
   End If
   lstAtt(0).Clear
   lstAtt(0).AddItem stFileName, 0
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   PUB_DelFtpFile2 strPLF01, " and PLF02=" & strPLF02, UCase("PRICELISTFILE") 'Add By Sindy 2017/5/31 檔案改放 FTP,必須在DB資料刪除前執行
   strSql = "DELETE FROM pricelistfile WHERE plf01='" & strPLF01 & "' and plf02=" & strPLF02
   cnnConnection.Execute strSql
   
   If SaveAttFile(strPLF01, strPLF02, 0) = False Then GoTo ErrHand
   
   strSql = "update pricelistfile Set" & _
                   " plf05='" & m_strCName & "',plf06=" & m_strCDate & ",plf07=" & Right("000000" & m_strCTime, 6) & _
                   ",plf08='" & strUserNum & "',plf09=" & strSrvDate(1) & ",plf10=" & Right("000000" & ServerTime, 6) & _
            " WHERE plf01='" & strPLF01 & "' and plf02=" & strPLF02
   cnnConnection.Execute strSql
   
   'Add By Sindy 2024/8/5 by系統類別更新下架日期
   strSql = "update pricelistfile Set plf04=" & IIf(Val(textDate(3)) = 0, "null", DBDATE(textDate(3))) & " where plf01='" & strPLF01 & "'"
   cnnConnection.Execute strSql
   '2024/8/5 END
   
   cnnConnection.CommitTrans
   
   ShowCurrRecord strPLF01, strPLF02
   
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strPLF01 As String, strPLF02 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   If Text1(1) <> m_CurrKEY(0) And DBDATE(textDate(2)) <> m_CurrKEY(1) Then
      MsgBox "系統記錄的目前系統類別（" & m_CurrKEY(0) & "）及啟用日期（" & m_CurrKEY(1) & "）與畫面上的系統類別及啟用日期不一致，請重新確認！"
      Exit Function
   End If
   
   cnnConnection.BeginTrans
   
   strPLF01 = m_CurrKEY(0)
   strPLF02 = m_CurrKEY(1)
   
   PUB_DelFtpFile2 strPLF01, " and PLF02=" & strPLF02, UCase("PRICELISTFILE") 'Add By Sindy 2017/5/31 檔案改放 FTP,必須在DB資料刪除前執行
   strSql = "DELETE FROM pricelistfile WHERE plf01='" & strPLF01 & "' and plf02=" & strPLF02
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   
   If (strPLF01 = m_LastKEY(0) And strPLF02 = m_LastKEY(1)) Or (strPLF01 = m_FirstKEY(0) And strPLF02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strPLF01, strPLF02
   
   DelRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strPLF01 As String, strPLF02 As String
   
   QueryRecord = False
   strPLF01 = Text1(1)
   strPLF02 = DBDATE(textDate(2))
   If IsRecordExist(strPLF01, strPLF02) = True Then
      m_CurrKEY(0) = strPLF01
      m_CurrKEY(1) = strPLF02
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
      ClearField
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse

   OnWork = False
   Select Case m_EditMode
      Case 1: '新增
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
'            UpdateFieldNewData
            If AddRecord = True Then
                RefreshRange
            Else
                Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '修改
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
'            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If Text1(1) <> "" And textDate(2) <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            If Text1(1) = "" Then
               MsgBox "請輸入系統類別才可進行查詢動作！", vbInformation
               Text1(1).SetFocus
               GoTo EXITSUB
            End If
            If textDate(2) = "" Then
               MsgBox "請輸入啟用日期才可進行查詢動作！", vbInformation
               textDate(2).SetFocus
               GoTo EXITSUB
            End If
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: If Me.Visible = True Then Text1(1).SetFocus
      'Case 2: If Me.Visible = True Then textDate(2).SetFocus
      Case 4: If Me.Visible = True Then Text1(1).SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM pricelistfile WHERE plf01='" & strKEY01 & "'" & _
            " and plf02=" & Val(DBDATE(strKEY02))
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
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "SELECT plf01,plf02 FROM pricelistfile WHERE plf01='" & m_CurrKEY(0) & "' and plf02=" & Val(DBDATE(m_CurrKEY(1)))
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("plf01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("plf01")
         If IsNull(rsTmp.Fields("plf02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("plf02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT plf01,plf02 FROM pricelistfile" & _
               " WHERE plf02=(SELECT MIN(plf02) FROM pricelistfile" & _
                             " where plf01=(select min(plf01) from pricelistfile))" & _
                 " and plf01=(select min(plf01) from pricelistfile)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("plf01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("plf01")
         If IsNull(rsTmp.Fields("plf02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("plf02")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   m_CurrKEY(1) = m_FirstKEY(1)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT plf01,plf02 FROM pricelistfile " & _
            "WHERE plf01 = '" & m_CurrKEY(0) & "' AND " & _
                  "plf02 = (SELECT MAX(plf02) FROM pricelistfile " & _
                            "WHERE plf01 = '" & m_CurrKEY(0) & "' AND " & _
                                  "plf02 < '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("plf01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("plf01")
      If IsNull(rsTmp.Fields("plf02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("plf02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT plf01,plf02 FROM pricelistfile " & _
            "WHERE plf01 = (SELECT MAX(plf01) FROM pricelistfile " & _
                            "WHERE plf01 < '" & m_CurrKEY(0) & "') AND " & _
                  "plf02 = (SELECT MAX(plf02) FROM pricelistfile " & _
                            "WHERE plf01 = (SELECT MAX(plf01) FROM pricelistfile " & _
                                            "WHERE plf01 < '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("plf01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("plf01")
      If IsNull(rsTmp.Fields("plf02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("plf02")
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT plf01,plf02 FROM pricelistfile " & _
            "WHERE plf01 = '" & m_CurrKEY(0) & "' AND " & _
                  "plf02 = (SELECT MIN(plf02) FROM pricelistfile " & _
                            "WHERE plf01 = '" & m_CurrKEY(0) & "' AND " & _
                                  "plf02 > '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("plf01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("plf01")
      If IsNull(rsTmp.Fields("plf02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("plf02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT plf01,plf02 FROM pricelistfile " & _
            "WHERE plf01 = (SELECT MIN(plf01) FROM pricelistfile " & _
                            "WHERE plf01 > '" & m_CurrKEY(0) & "') AND " & _
                  "plf02 = (SELECT MIN(plf02) FROM pricelistfile " & _
                            "WHERE plf01 = (SELECT MIN(plf01) FROM pricelistfile " & _
                                            "WHERE plf01 > '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("plf01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("plf01")
      If IsNull(rsTmp.Fields("plf02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("plf02")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         UpdateToolbarState
         SetInputEntry
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
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
'         UpdateFieldNewData
         If OnWork = True Then
            Me.SSTab1.TabEnabled(1) = True
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
                  Me.SSTab1.TabEnabled(1) = True
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               Me.SSTab1.TabEnabled(1) = True
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT plf01,plf02 FROM pricelistfile " & _
            "WHERE plf01 = (SELECT MIN(plf01) FROM pricelistfile) AND " & _
                  "plf02 = (SELECT MIN(plf02) FROM pricelistfile " & _
                           "WHERE plf01 = (SELECT MIN(plf01) FROM pricelistfile)) "
                           
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("plf01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("plf01")
      If IsNull(rsTmp.Fields("plf02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("plf02")
   End If
   rsTmp.Close

   strSql = "SELECT plf01,plf02 FROM pricelistfile " & _
            "WHERE plf01 = (SELECT MAX(plf01) FROM pricelistfile) AND " & _
                  "plf02 = (SELECT MAX(plf02) FROM pricelistfile " & _
                           "WHERE plf01 = (SELECT MAX(plf01) FROM pricelistfile)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("plf01")) = False Then: m_LastKEY(0) = rsTmp.Fields("plf01")
      If IsNull(rsTmp.Fields("plf02")) = False Then: m_LastKEY(1) = rsTmp.Fields("plf02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   strSql = "SELECT plf01,plf02,plf03,plf05,plf06,plf07,plf08,plf09,plf10" & _
             " FROM pricelistfile WHERE plf01='" & m_CurrKEY(0) & "' and plf02=" & Val(m_CurrKEY(1))
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      
      If IsNull(rsTmp.Fields("plf01")) = False Then: Text1(1) = rsTmp.Fields("plf01")
      If IsNull(rsTmp.Fields("plf02")) = False Then: textDate(2) = TransDate(rsTmp.Fields("plf02"), 1)
      
      'Add By Sindy 2024/8/5 下架日期
      textDate(3) = ""
      strExc(0) = "select plf04 from pricelistfile where plf01='" & Text1(1) & "' and plf02=" & Val(DBDATE(textDate(2)))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         textDate(3) = TransDate("" & RsTemp.Fields("plf04"), 1)
      End If
      '2024/8/5 END
      
      If IsNull(rsTmp.Fields("plf05")) = False Then: m_strCName = rsTmp.Fields("plf05")
      If IsNull(rsTmp.Fields("plf06")) = False Then: m_strCDate = rsTmp.Fields("plf06")
      If IsNull(rsTmp.Fields("plf07")) = False Then: m_strCTime = rsTmp.Fields("plf07")
      
      lstAtt(0).AddItem SetFileName(Text1(1), textDate(2)), 0
      'lstAtt(0).ItemData(0) = 1
      SetListScroll lstAtt(0)
      
      ' 更新CUID
      Call UpdateCUID(rsTmp)
'      ' 更新暫存區的資料
'      UpdateFieldOldData rsTmp
      SSTab1.Tab = 0
   End If
   
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("plf05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("plf05")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("plf05"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("plf06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("plf06")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("plf06"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("plf07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("plf07")) = False Then
         strTemp = rsSrcTmp.Fields("plf07")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("plf08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("plf08")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("plf08"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("plf09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("plf09")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("plf09"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("plf10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("plf10")) = False Then
         strTemp = rsSrcTmp.Fields("plf10")
         strUTime = Format(strTemp, "##:##:##")
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

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
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
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
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

Private Function CheckDataValid() As Boolean
   CheckDataValid = False
   
   If Text1(1).Text = "" Then
      MsgBox "系統類別不可以空白！", vbExclamation
      Text1(1).SetFocus
      Exit Function
   End If
   
   If textDate(2) = "" Then
      MsgBox "啟用日期不可以空白！", vbExclamation
      textDate(2).SetFocus
      Exit Function
   Else
      If m_EditMode = 1 Then
         If Val(DBDATE(textDate(2))) < Val(strSrvDate(1)) Then
            If MsgBox("確定啟用日期小於系統日嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               textDate(2).SetFocus
               Exit Function
            End If
         End If
         'Modify By Sindy 2024/8/5 新增時,此系統類別不可以有下架日期存在
         strExc(0) = "select plf01 from pricelistfile where plf01='" & Text1(1) & "' and nvl(plf04,0)>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "此系統類別已有下架日期！" & vbCrLf & _
                   "若要新增資料，請先取消下架日期。", vbExclamation
            Exit Function
         End If
         '2024/8/5 END
      End If
   End If
   
   If lstAtt(0).ListCount = 0 Then
      MsgBox "請加入附件！", vbExclamation
      cmdAddAtt(0).SetFocus
      Exit Function
   End If
   
   If UCase(Right(GetFileName(Trim(lstAtt(0).List(0))), 4)) <> UCase(".pdf") Then
      MsgBox "附件必須是PDF檔！", vbExclamation
      cmdRemAtt(0).SetFocus
      Exit Function
   End If
   
   CheckDataValid = True
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   Text1(1).Enabled = Not bEnable
   textDate(2).Enabled = Not bEnable
   If bEnable Then Text1(1).BackColor = &H8000000F Else Text1(1).BackColor = &H80000005
   If bEnable Then textDate(2).BackColor = &H8000000F Else textDate(2).BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   Text1(1).Enabled = Not bEnable
   textDate(2).Enabled = Not bEnable
   textDate(3).Enabled = Not bEnable 'Add By Sindy 2024/8/5
   If bEnable Then Text1(1).BackColor = &H8000000F Else Text1(1).BackColor = &H80000005
   If bEnable Then textDate(2).BackColor = &H8000000F Else textDate(2).BackColor = &H80000005
   If bEnable Then textDate(3).BackColor = &H8000000F Else textDate(3).BackColor = &H80000005 'Add By Sindy 2024/8/5
   cmdAddAtt(0).Enabled = Not bEnable
   cmdRemAtt(0).Enabled = Not bEnable
   'cmdOpenAtt(0).Enabled = Not bEnable
   cmdOpenAtt(0).Enabled = True
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   Text1(1) = Empty
   textDate(2) = Empty
   textDate(3) = Empty 'Add By Sindy 2024/8/5
   lstAtt(0).Clear
   Label23 = Empty
   
   m_strCName = ""
   m_strCDate = ""
   m_strCTime = ""
   
   SetGrd
   For nIndex = 0 To tf_PLF - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

'Private Sub UpdateFieldNewData()
'Dim MyArr As Variant
'   '若新增資料
'   If m_EditMode = 1 Then
'      SetFieldNewData "plf01", Text1(1)
'      SetFieldNewData "plf02", dbdate(Textdate(2))
'   End If
'   SetFieldNewData "plf03", ""
'End Sub
'
'' 初始化欄位陣列
'Private Sub InitialField()
'Dim nIndex As Integer
'Dim strTmp As String
'   ' 初始化欄位陣列
'   For nIndex = 1 To tf_PLF
'      strTmp = Format(nIndex, "00")
'      m_FieldList(nIndex - 1).fiName = "PLF" & strTmp
'      m_FieldList(nIndex - 1).fiOldData = Empty
'      m_FieldList(nIndex - 1).fiNewData = Empty
'      m_FieldList(nIndex - 1).fiType = 0 '文字型態
'      Select Case nIndex
'         Case 2, 3, 4, 5, 6, 7:
'            m_FieldList(nIndex - 1).fiType = 1 '數值型態
'      End Select
'   Next nIndex
'End Sub

'帶預設資料
Private Sub InitialData()
   SetGrd
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   'Modify By Sindy 2024/8/5 +下架日期
   arrGridHeadText = Array("系統類別", "啟用日期", "plf01", "新增人員", "日期", "時間", "修改人員", "日期", "時間", "下架日期")
   arrGridHeadWidth = Array(1500, 800, 0, 900, 800, 800, 900, 800, 800, 800)
   GRD1.Visible = False
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

'系統類別
Private Sub Text1_GotFocus(Index As Integer)
   InverseTextBox Text1(Index)
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub
Private Sub Text1_LostFocus(Index As Integer)
   If Trim(Text1(Index)) <> "" And Len(Trim(Text1(Index))) <> 2 Then
      Text1(Index) = Right("00" & Trim(Text1(Index)), 2)
   End If
End Sub
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Text1(Index) <> "" Then
      If Val(Text1(Index)) > 12 Or Val(Text1(Index)) < 1 Then
         MsgBox "系統類別只可以輸入01 ~ 12！", vbInformation
         Call Text1_GotFocus(Index)
         Cancel = True
         Exit Sub
      Else
         Call Text1_LostFocus(Index)
      End If
      If m_EditMode = 1 Then
         If IsRecordExist(Text1(1), DBDATE(textDate(2))) = True And Text1(1).Enabled = True And Text1(1).Locked = False Then
            MsgBox "此筆資料已存在，請修改！", vbInformation
            Call Text1_GotFocus(1)
            Cancel = True
            Exit Sub
         End If
         'Modify By Sindy 2024/8/5 新增時,此系統類別不可以有下架日期存在
         strExc(0) = "select plf01 from pricelistfile where plf01='" & Text1(1) & "' and nvl(plf04,0)>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "此系統類別已有下架日期！" & vbCrLf & _
                   "若要新增資料，請先取消下架日期。", vbExclamation
            Call Text1_GotFocus(Index)
            Cancel = True
            Exit Sub
         End If
         '2024/8/5 END
      End If
   End If
End Sub

'啟用日期
Private Sub textDate_GotFocus(Index As Integer)
   InverseTextBox textDate(Index)
   CloseIme
End Sub
Private Sub textDate_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub
Private Sub textDate_Validate(Index As Integer, Cancel As Boolean)
   If textDate(Index) <> "" Then
      If ChkDate(textDate(Index)) = False Then
          Call textDate_GotFocus(Index)
          Cancel = True
          Exit Sub
      End If
      If Index = 2 And (m_EditMode = 1 Or m_EditMode = 2) Then
         If ChkWorkDay(DBDATE(textDate(2))) = False Then
            MsgBox "啟用日期必須是工作日！", vbExclamation
            Call textDate_GotFocus(2)
            Cancel = True
            Exit Sub
         End If
      End If
      Select Case Index
         Case 0
            If textDate(Index) <> "" And textDate(Index + 1) = "" Then
               textDate(Index + 1) = textDate(Index)
            End If
         Case 1
            If RunNick2(textDate(Index - 1), textDate(Index)) Then
               Call textDate_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
      End Select
   End If
End Sub

Private Function GetAttachFile(ByRef pFileName As String, Optional pSavePath As String, Optional pFileSize As Integer = 0) As Boolean
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte
   
On Error GoTo ErrHnd
   
   strExc(0) = "select * from pricelistfile where plf01='" & Text1(1) & "' and plf02=" & Val(DBDATE(textDate(2)))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      pFileName = SetFileName(RsTemp.Fields("plf01"), RsTemp.Fields("plf02"))
      
      If pSavePath = "" Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
         stAttPath = m_AttachPath & "\" & pFileName
         '檔案已存在時
         If Dir(stAttPath) <> "" Then
            '檢查檔案是否正在使用中
            If PUB_ChkFileOpening(stAttPath) = True Then
               MsgBox stAttPath & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
               Exit Function
            End If
            Kill stAttPath
         End If
      Else
         stAttPath = pSavePath
      End If
      
      If Dir(stAttPath) <> "" Then Kill stAttPath
      'Add By Sindy 2017/5/31
      If "" & RsTemp.Fields("plf11") <> "" Then
         GetAttachFile = PUB_GetFtpFile(RsTemp.Fields("plf11"), stAttPath, UCase("PRICELISTFILE"))
      Else
      '2017/5/31 END
         With RsTemp
            lngSize = Val(.Fields("plf03").Value)
            ReDim bytes(lngSize)
            If lngSize > 0 Then bytes() = .Fields("plf04").GetChunk(lngSize)
         End With
         iFileNo = FreeFile
         Open stAttPath For Binary Access Write As #iFileNo
         If lngSize > 0 Then Put #iFileNo, , bytes()
         Close #iFileNo
      End If
      pFileName = stAttPath
      If pFileSize = 1 Then
         pFileName = pFileName & " (" & Round(RsTemp.Fields("plf03") / 1024, 2) & " KB)"
      End If
      GetAttachFile = True
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   If iFileNo > 0 Then Close #iFileNo
End Function

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String
   Dim bolIsSelect As Boolean
   Dim ii As Integer
   
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   
   strAtt = lstAtt(Index).List(0)
   
   If strAtt = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
'      For ii = 0 To lstAtt(Index).ListCount - 1
'         If lstAtt(Index).Selected(ii) Then
'            bolIsSelect = True
            stFileName = lstAtt(Index).List(ii)
            If InStrRev(stFileName, " (") > 0 Then
               stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            End If
            
            If InStr(stFileName, "\") = 0 Then
               If GetAttachFile(stFileName) = False Then Exit Sub
            End If
            
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
'         End If
'      Next ii
'      If bolIsSelect = False Then
'         MsgBox "請選擇欲開啟的附件！"
'      End If
   End If
   
   Screen.MousePointer = vbDefault
End Sub

'新增
Private Sub cmdAddAtt_Click(Index As Integer)
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s
   Dim strFile As String
   
On Error GoTo ErrHnd
   
   If lstAtt(Index).List(0) <> "" Then
      MsgBox "已有附件，不可新增！"
      Exit Sub
   End If
   
   stFileName = "*.*"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.*)|*.*"
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
            For ii = 1 To UBound(sFile)
               If InStr(CStr(sFile(ii)), "#") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
                  Exit Sub
               End If
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               '檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  ShowMsg sFile(ii) & MsgText(9221)
                  Exit Sub
               End If
               AddListX lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(Index)
            Next
         Else
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(strFile, "#") > 0 Then
               MsgBox strFile & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
               Exit Sub
            End If
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     Exit For
                  End If
               Next ii
            End If
            
            stFileName = .FileName
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            '檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg strFile & MsgText(9221)
               Exit Sub
            End If
            AddListX lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#", lstAtt(Index)
         End If
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Function AddListX(oList As Object, stNewItem As String, oList1 As Object) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String
      
   If stNewItem <> "" Then
      For idx = 0 To 0 'oList.ListCount - 1
         stFileName = GetFileName(oList.List(idx))
         If GetFileName(stNewItem) = stFileName Then
            MsgBox "附件 " & stFileName & " 已存在！"
            AddListX = False
            bFound = True
            Exit For
         End If
      Next
      
'      If bFound = False Then
'         For idx = 0 To 0 'oList1.ListCount - 1
'            stFileName = GetFileName(oList1.List(idx))
'            If GetFileName(stNewItem) = stFileName Then
'               MsgBox "附件 " & stFileName & " 已存在！"
'               AddListX = False
'               bFound = True
'               Exit For
'            End If
'         Next
'      End If
      
      oList.Clear
      If bFound = False Then
         oList.AddItem stNewItem, 0
         SetListScroll oList
         AddListX = True
      End If
   End If
End Function

Private Sub SetListScroll(oList As Object)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

'刪除
Private Sub cmdRemAtt_Click(Index As Integer)
   Call RemoveList(lstAtt(Index), Index)
End Sub

Private Function RemoveList(oList As Object, Index As Integer) As Boolean
   Dim ii As Integer
   If oList.ListCount > 0 Then
'      ii = 0
'      Do While ii < oList.ListCount
'         If oList.Selected(ii) = True Then
            
'            'Add By Sindy 2017/5/31
'            If MsgBox("確定要永久刪除" & oList.List(0) & "電子檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
'               Screen.MousePointer = vbDefault
'               Exit Function
'            End If
'            '直接從資料庫刪除檔案
'            If PUB_DelFtpFile2(m_CurrKEY(0), " and plf02=" & m_CurrKEY(1), "PRICELISTFILE") = True Then  '檔案改放FTP,必須在DB資料刪除前執行
'               strSql = "DELETE FROM pricelistfile WHERE plf01='" & m_CurrKEY(0) & "' and plf02=" & m_CurrKEY(1)
'               cnnConnection.Execute strSql
'               'Call ReadAttachFile
'            End If
'            '2017/5/31 END
            
            oList.RemoveItem 0 'ii
            SetListScroll oList
            RemoveList = True
'            ii = ii - 1
'         End If
'         ii = ii + 1
'      Loop
   End If
End Function

'檔案名稱：啟用日期(民國日期)＋系統類別中文名稱＋價目表.pdf
Private Function SetFileName(strSysKind As String, strDate As String) As String
   Select Case strSysKind
      Case "01"
         strSysKind = "國內專利"
      Case "02"
         strSysKind = "大陸專利"
      Case "03"
         strSysKind = "香港澳門專利"
      Case "04"
         strSysKind = "CFP"
      Case "05"
         strSysKind = "國內商標"
      Case "06"
         strSysKind = "大陸商標"
      Case "07"
         strSysKind = "馬德里商標"
      Case "08"
         strSysKind = "國內著作權"
      Case "09"
         strSysKind = "大陸著作權"
      Case "10"
         strSysKind = "CFT"
      Case "11"
         strSysKind = "美國著作權"
      Case "12"
         strSysKind = "顧問及法務"
   End Select
   SetFileName = TransDate(strDate, 1) & strSysKind & "價目表.pdf"
End Function
