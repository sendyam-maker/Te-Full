VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210101_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "接洽人資料維護"
   ClientHeight    =   5784
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   8376
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5784
   ScaleWidth      =   8376
   Begin VB.Frame fraContact 
      Height          =   3456
      Left            =   30
      TabIndex        =   28
      Top             =   2300
      Width           =   8300
      Begin VB.CheckBox Check2 
         Caption         =   "是臺灣聯絡地址"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   5
         Top             =   830
         Width           =   2000
      End
      Begin VB.TextBox textCUID1 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   210
         Width           =   6270
      End
      Begin VB.CheckBox Check1 
         Caption         =   "聯絡地址同客戶"
         Height          =   195
         Left            =   3300
         TabIndex        =   3
         Top             =   570
         Width           =   1680
      End
      Begin VB.CommandButton cmdTW 
         Caption         =   "臺灣地址格式"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   70
         TabIndex        =   7
         Top             =   1245
         Width           =   1160
      End
      Begin VB.CommandButton cmdSearchZip 
         Caption         =   "郵遞區號查詢"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   1620
         Width           =   1160
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "傳真："
         Height          =   180
         Index           =   9
         Left            =   4680
         TabIndex        =   42
         Top             =   2304
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "手機："
         Height          =   180
         Index           =   11
         Left            =   708
         TabIndex        =   41
         Top             =   2592
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "電話："
         Height          =   180
         Index           =   12
         Left            =   708
         TabIndex        =   40
         Top             =   2256
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LINE ID："
         Height          =   180
         Index           =   28
         Left            =   4416
         TabIndex        =   39
         Top             =   2568
         Width           =   804
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   288
         Index           =   30
         Left            =   1275
         TabIndex        =   13
         Top             =   2232
         Width           =   2976
         VariousPropertyBits=   679493659
         MaxLength       =   20
         Size            =   "5249;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   288
         Index           =   32
         Left            =   1275
         TabIndex        =   15
         Top             =   2544
         Width           =   2976
         VariousPropertyBits=   679493659
         MaxLength       =   20
         Size            =   "5249;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   288
         Index           =   31
         Left            =   5208
         TabIndex        =   14
         Top             =   2208
         Width           =   2976
         VariousPropertyBits=   679493659
         MaxLength       =   20
         Size            =   "5249;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   288
         Index           =   33
         Left            =   5208
         TabIndex        =   16
         Top             =   2520
         Width           =   2976
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5239;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   285
         Index           =   5
         Left            =   1275
         TabIndex        =   2
         Top             =   510
         Width           =   1875
         VariousPropertyBits=   679493659
         MaxLength       =   10
         Size            =   "3307;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   285
         Index           =   8
         Left            =   1275
         TabIndex        =   10
         Top             =   1920
         Width           =   3180
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5609;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   552
         Index           =   13
         Left            =   1272
         TabIndex        =   17
         Top             =   2844
         Width           =   6972
         VariousPropertyBits=   679493659
         MaxLength       =   250
         ScrollBars      =   2
         Size            =   "12303;979"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   285
         Index           =   2
         Left            =   1275
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   600
         VariousPropertyBits=   679493659
         Size            =   "1058;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   285
         Index           =   21
         Left            =   1275
         TabIndex        =   8
         Top             =   1620
         Width           =   1200
         VariousPropertyBits=   679493659
         MaxLength       =   10
         Size            =   "2117;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   540
         Index           =   22
         Left            =   1275
         TabIndex        =   6
         Top             =   1050
         Width           =   6960
         VariousPropertyBits=   -1467989989
         MaxLength       =   70
         ScrollBars      =   2
         Size            =   "12277;952"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   285
         Index           =   10
         Left            =   6660
         TabIndex        =   4
         Top             =   510
         Width           =   330
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "582;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   285
         Index           =   23
         Left            =   6690
         TabIndex        =   12
         Top             =   1920
         Width           =   330
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "582;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   285
         Index           =   24
         Left            =   6690
         TabIndex        =   11
         Top             =   1620
         Width           =   330
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "582;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label63 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "名稱："
         Height          =   180
         Index           =   0
         Left            =   735
         TabIndex        =   38
         Top             =   510
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   14
         Left            =   708
         TabIndex        =   37
         Top             =   2868
         Width           =   540
      End
      Begin VB.Label Label63 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "編號："
         Height          =   180
         Index           =   7
         Left            =   735
         TabIndex        =   36
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label63 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "E-MAIL："
         Height          =   180
         Index           =   5
         Left            =   468
         TabIndex        =   35
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label Label63 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "郵遞區號："
         Height          =   180
         Index           =   1
         Left            =   348
         TabIndex        =   34
         Top             =   1620
         Width           =   900
      End
      Begin VB.Label Label63 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "聯絡地址："
         Height          =   180
         Index           =   2
         Left            =   375
         TabIndex        =   33
         Top             =   1050
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄國內電子報：       （N:不寄）"
         Height          =   180
         Index           =   17
         Left            =   5055
         TabIndex        =   32
         Top             =   585
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄顧問電子報：        （Y:寄/N:不寄）"
         Height          =   180
         Index           =   1
         Left            =   5055
         TabIndex        =   31
         Top             =   1965
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄專利雙週報：       （N:不寄）"
         Height          =   180
         Index           =   21
         Left            =   5055
         TabIndex        =   30
         Top             =   1665
         Width           =   2820
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "存檔(&S)"
      Height          =   400
      Index           =   1
      Left            =   6300
      TabIndex        =   21
      Top             =   30
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.TextBox txtCuNo 
      Height          =   276
      Left            =   1035
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   7110
      TabIndex        =   22
      Top             =   30
      Width           =   1125
   End
   Begin VB.CommandButton cmdContact 
      Caption         =   "新增"
      Height          =   285
      Index           =   1
      Left            =   5895
      TabIndex        =   18
      Top             =   2010
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdContact 
      Caption         =   "刪除"
      Height          =   285
      Index           =   3
      Left            =   7425
      TabIndex        =   20
      Top             =   2010
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdContact 
      Caption         =   "加入"
      Height          =   285
      Index           =   2
      Left            =   6660
      TabIndex        =   19
      Top             =   2010
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   90
      Top             =   1950
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   572
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm210101_2.frx":0000
      Height          =   1095
      Left            =   90
      TabIndex        =   23
      Top             =   840
      Width           =   8175
      _ExtentX        =   14415
      _ExtentY        =   1926
      _Version        =   393216
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   14
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "PCC02"
         Caption         =   "編號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "PCC05"
         Caption         =   "名稱"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "PCC21"
         Caption         =   "郵遞區號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "PCC22"
         Caption         =   "聯絡地址"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "PCC08"
         Caption         =   "EMail"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PCC10"
         Caption         =   "是否寄電子報"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   315
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1272.189
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   875.906
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   2844.284
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1632.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   432
         EndProperty
      EndProperty
   End
   Begin MSForms.Label lblCustAddress 
      Height          =   350
      Left            =   1080
      TabIndex        =   27
      Top             =   420
      Width           =   7000
      VariousPropertyBits=   27
      Size            =   "12347;617"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "聯絡地址："
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   26
      Top             =   420
      Width           =   900
   End
   Begin MSForms.Label lblCustName 
      Height          =   195
      Left            =   2160
      TabIndex        =   25
      Top             =   150
      Width           =   3780
      VariousPropertyBits=   27
      Size            =   "6667;353"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   24
      Top             =   150
      Width           =   900
   End
End
Attribute VB_Name = "frm210101_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/14 Form2.0已修改 lblCustName/lblCustAddress/txtPcc()/DataGrid1
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
'Create by Morgan 2008/7/31
Option Explicit

Dim m_iConEditMode As Integer '聯絡人狀態 1:新增 2:修改
Dim m_bReadGrid As Boolean '是否要讀取被點選聯絡人資料
Dim rsContact As ADODB.Recordset
Dim rsContactOld As ADODB.Recordset
Public strCU127 As String 'Add by Morgan 2009/1/9 預設聯絡人
Dim strCheck As String '檢查欄位是否變動字串
Dim oText 'Modify by Amy 2021/12/14 原:As TextBox
Public bolCU87IsTW As Boolean  '客戶國籍是否為台灣
Dim strECustMsg As String 'Added by Morgan 2024/6/3 全E化客戶提醒

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      txtPCC(21).Text = Empty
      txtPCC(21).Enabled = False
      txtPCC(22).Text = Empty
      txtPCC(22).Enabled = False
   'Added by Morgan 2022/9/2 預設接洽人不可設聯絡地址
   ElseIf Val(strCU127) = Val(txtPCC(2)) Then
      MsgBox "【" & txtPCC(5) & "】為預設接洽人，不可設【聯絡地址】！", vbExclamation
      Check1.Value = vbChecked
   'end 2022/9/2
   Else
      txtPCC(21).Enabled = True
      txtPCC(22).Enabled = True
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim strData As String, iRtn As Integer
   If Index = 1 Then
      'Add by Morgan 2009/1/9 檢查是否資料有變動但未按 [ 加入 ] 鈕
      For Each oText In txtPCC
         strData = strData & oText
      Next
      If strCheck <> strData Then
         iRtn = MsgBox("目前編輯的連絡人資料尚未更新是否要加入？", vbYesNoCancel + vbDefaultButton3)
         '取消
         If iRtn = 2 Then
            Exit Sub
         '加入
         ElseIf iRtn = 6 Then
            If UpdContact = False Then
               Exit Sub
            End If
         End If
      End If
      If ModRecord = False Then
         Exit Sub
         
      'Added by Morgan 2024/6/3
      ElseIf strECustMsg <> "" Then
         MsgBox strECustMsg, vbInformation
      'end 2024/6/3
      End If
   End If
   Unload Me
End Sub

'Add by Amy 2016/05/16
Private Sub cmdSearchZip_Click()
    Call frm100134.SetParent(Me)
    Me.Hide
    frm100134.BFormZip = "txtPCC(21)"
    frm100134.strPrevFormMon = "frm210101_1"
    frm100134.GetStreet txtPCC(22).Text, 2
    frm100134.Show
End Sub

Private Sub cmdTW_Click()
    frm100135.Show vbModal
End Sub
'end 2016/05/16

'Modified by Morgan 2022/9/23 表格內
Private Sub DataGrid1_Click()
   '點選同一列可能不會觸發RowColChange
   If DataGrid1.col = -1 Then
      ReadContact
   End If
   m_bReadGrid = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If m_bReadGrid = True Then
      ReadContact
   End If
End Sub

Private Sub DataGrid1_Validate(Cancel As Boolean)
   m_bReadGrid = False
End Sub

Private Sub Form_Load()
   'Add by Amy 2016/05/16
   If 案件預設收據公司別啟用日 >= Val(strSrvDate(1)) Then
        cmdSearchZip.Visible = False
        cmdTW.Visible = False
   End If
   MoveFormToCenter Me
   textCUID1.BackColor = Me.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
   bolCU87IsTW = False 'Add by Amy 2016/05/16
   Set frm210101_2 = Nothing
End Sub

Public Sub OpenContactTable()
   
On Error GoTo Checking
   
   strExc(0) = "select * from potcustcont where pcc01='" & txtCuNo.Text & "' order by pcc02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/06/16 +FormName 改暫存TB
   Set rsContact = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   Set rsContactOld = PUB_CreateRecordset(RsTemp, , , , Me.Name)
   'end 2014/06/16
   Set RsTemp = Nothing
   Set Adodc1.Recordset = rsContact
   Set DataGrid1.DataSource = Adodc1
   DataGrid1.Refresh
   DataGrid1.col = 0
   DataGrid1.CurrentCellVisible = True
   If rsContact.RecordCount > 0 Then
      ReadContact
   End If
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
End Sub

'讀取該筆聯絡人資料
Private Sub ReadContact()
   Dim CUID(1 To 6) As String
   
   strCheck = ""
   ClearField1
   With Adodc1.Recordset
      If Not (.EOF Or .BOF) Then
         For Each oText In txtPCC
            'Modify by Amy 2021/12/14 改Form2.0 +.Text 否則資料無法顯示
            oText.Text = "" & .Fields("PCC" & Format(oText.Index, "00"))
            strCheck = strCheck & oText
         Next
         
         If txtPCC(22) <> "" Then
            Check1.Value = 0
         End If
         
         CUID(1) = "" & .Fields("PCC14")
         CUID(2) = "" & .Fields("PCC15")
         CUID(3) = "" & .Fields("PCC16")
         CUID(4) = "" & .Fields("PCC17")
         CUID(5) = "" & .Fields("PCC18")
         CUID(6) = "" & .Fields("PCC19")
         
         UpdateCUID CUID, textCUID1
      End If
   End With
End Sub

'清除舊資料
Private Sub ClearField1()
   For Each oText In txtPCC
      oText.Text = Empty
   Next
   textCUID1.Text = Empty
   Check1.Value = 1
   strCheck = ""
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef pText As TextBox)
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
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   pText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Sub txtPCC_GotFocus(Index As Integer)
   Select Case Index
      Case 5, 22, 13
         OpenIme
      Case Else
         CloseIme
   End Select
   TextInverse txtPCC(Index)
End Sub

'Modify by Amy 2021/12/14 原:Integer
Private Sub txtPCC_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Dim strAddr As String, strNewArea As String, strZipCode As String, strCountry As String, strROC As String, strIndArea As String
    Dim intArea As Integer, intFocus As Integer
    Dim bolMany As Boolean
    
   'Modify by Amy 2021/12/14 +txtPCC(Index)
   Select Case Index
      'Modify by Amy 2016/05/16 將地址拆出22
      Case 5, 21
         KeyAscii = ChangeZIP(KeyAscii, txtPCC(Index))
      'Add by Morgan 2008/12/22
      Case 10, 24
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
     'Add by Amy 2016/05/16
     Case 22
         KeyAscii = ChangeZIP(KeyAscii, txtPCC(Index))
         If LTrim(Me.txtPCC(Index)) <> MsgText(601) Then
            strROC = ""
            strAddr = Me.txtPCC(Index)
            If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
            If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
            If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
            '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
            strIndArea = "True"
            strAddr = ReplaceIndArea(strAddr, strIndArea)
            If strIndArea = "True" Then strIndArea = MsgText(601)
            If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
                strIndArea = "新竹" & strIndArea
                strAddr = Mid(strAddr, 3)
            End If
            If Len(LTrim(strAddr)) > 4 And (Mid(strAddr, 3, 1) = "市" Or Mid(strAddr, 3, 1) = "縣") Then
                '輸到路/街/段.
                If Asc("路") = KeyAscii Or Asc("街") = KeyAscii Or Asc("段") = KeyAscii Then
                    intFocus = Val(Me.txtPCC(Index).SelStart) - Len(strROC) - Len(strIndArea)
                    strAddr = Mid(strAddr, 1, intFocus) & Chr(KeyAscii) & Mid(strAddr, intFocus + 1) 'KeyPress未完成時地址欄位尚未顯示目前字,故先加入當下的字查
                    '有鄉/鎮/市/區
                    'Modify by Amy 2018/12/19 +判斷第七個字 ex:嘉義縣阿里山鄉 X80024
                    If Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" _
                      Or Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" _
                      Or Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
                        strZipCode = GetZipCode_Tai(1, strAddr, , bolMany, , strCountry)
                        If strZipCode <> MsgText(601) Then
                            If bolMany = False Then
                                Call ChkZipData(2, Me.txtPCC(Index), strZipCode, , strCountry)
                                Me.txtPCC(Index).SelStart = intFocus + Len(strROC) + Len(strIndArea)
                                Me.txtPCC(Index).SelLength = 0
                            Else
                                '多筆以縣/市+鄉/鎮/市/區及路名查
                                bolMany = False
                                strZipCode = GetZipCode_Tai(3, Mid(strAddr, 1, intFocus + 1), intArea, bolMany, , strCountry)
                                If strZipCode <> MsgText(601) And bolMany = False Then
                                    Call ChkZipData(2, Me.txtPCC(Index), strZipCode, , strCountry)
                                    Me.txtPCC(Index).SelStart = intFocus + Len(strROC) + Len(strIndArea)
                                    Me.txtPCC(Index).SelLength = 0
                                End If
                            End If
                        End If
                    '沒鄉/鎮/市/區
                    Else
                        '取 段/路/街 查
                        strZipCode = GetZipCode_Tai(2, Mid(strAddr, 1, intFocus + 1), intArea, bolMany, strNewArea, strCountry)
                        If strZipCode <> MsgText(601) And bolMany = False Then
                            '補上查到的區,避免輸入兩個同樣的字(路/街/段)被取代,故不用Replace
                            Me.txtPCC(Index) = strROC & Left(strAddr, 3) & strNewArea & strIndArea & Mid(strAddr, 4, intArea - 4) & Mid(strAddr, intFocus + 2)
                            Call ChkZipData(2, Me.txtPCC(Index), strZipCode, , strCountry)
                            Me.txtPCC(Index).SelStart = intFocus + Len(strROC) + Len(strNewArea) + Len(strIndArea)
                            Me.txtPCC(Index).SelLength = 0
                        End If
                    End If
                End If
            End If
         End If
      'Add by Sindy 2011/3/18
      Case 23
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
            Beep
         End If
   End Select
   'end 2021/12/14
End Sub

'聯絡人
Private Sub cmdContact_Click(Index As Integer)
   Dim bDupeCheck As Boolean, sPCC(1 To 5) As String
   Select Case Index
      Case 1 '新增
         ClearField1
         txtPCC(2).Text = getNewNo
         txtPCC(5).SetFocus
         m_iConEditMode = 1
         
      Case 2 '加入
         'Add by Amy 2016/05/16
         'modify by sonia 2016/5/31
         'If Check2.Value = 1 Then
         If Check2.Value = 1 And Check1.Value = 0 Then
            If FormCheck() = False Then Exit Sub
         End If
         'end f2016/05/16
         UpdContact
         
      Case 3 '刪除
         If txtPCC(2) <> "" Then
            If Not (rsContact.EOF Or rsContact.BOF) Then
               If PUB_PCCDelCheck(txtCuNo, , txtPCC(2)) = True Then
                  If fnCaseCheck(txtCuNo, txtPCC(2), txtPCC(5)) = True Then  'Added by Morgan 2025/5/27
                     rsContact.Delete
                     rsContact.UpdateBatch
                     ClearField1
                  End If
               End If
            End If
         End If
   End Select
End Sub

Private Function UpdContact() As Boolean
   If TxtValidate1 = True Then
      PUB_FilterFormText Me 'Add by Morgan 2008/12/22 修正畫面所有含跳行符號的文字框
      UpdateContact
      DataGrid1.Refresh
      ClearField1
      UpdContact = True
   End If
End Function

Private Function getNewNo() As String
   Dim myTemp As ADODB.Recordset
   Dim iUsableNo As Integer
   
   Set myTemp = rsContact.Clone
   With myTemp
      .Sort = "PCC02 asc"
      iUsableNo = 1
      If .RecordCount > 0 Then
         'Added by Morgan 2021/8/27 新增若未超過99時改抓最大+1(順序比較明確,新的在後面),超過99才抓中間空的
         .MoveLast
         iUsableNo = Val("" & .Fields(1)) + 1
         If iUsableNo > 99 Then
         'end 2021/8/23
         
            .MoveFirst
            Do While Not .EOF
               If iUsableNo = Val("" & .Fields(1)) Then
                  iUsableNo = iUsableNo + 1
               Else
                  Exit Do
               End If
               .MoveNext
            Loop
            
         End If 'Added by Morgan 2021/8/27
      End If
      getNewNo = Format(iUsableNo, "00")
      
   End With
   Set myTemp = Nothing
End Function

Private Function TxtValidate1() As Boolean
   Dim idx As Integer
   Dim Cancel As Boolean
   
   For Each oText In txtPCC
      If oText.Locked = False Then
         idx = oText.Index
         Cancel = False
         txtPCC_Validate idx, Cancel
         If Cancel = True Then
            txtPCC_GotFocus idx
            Exit Function
         End If
      End If
   Next
   'Add by Amy 2021/11/26 +國內同業控制
   strExc(0) = "Select cu80 From Customer Where cu01 = '" & txtCuNo & "' And cu02= '0'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
        If "" & RsTemp(0) = "國內同業" Then
            If txtPCC(8) <> "" Then
                ShowMsg "此客戶為國內同業,不可輸入E-MAIL以免誤發電子郵件, 如有需要請加註於備註欄 !"
                txtPCC(8).SetFocus
                txtPCC_GotFocus 8
                Exit Function
            End If
            If txtPCC(10) <> "N" Then
                ShowMsg "此客戶為國內同業, 不可寄國內電子報 !"
                txtPCC(10).SetFocus
                txtPCC_GotFocus 10
                Exit Function
            End If
            If txtPCC(23) <> "N" Then
                ShowMsg "此客戶為國內同業, 不可寄顧問電子報 !"
                txtPCC(23).SetFocus
                txtPCC_GotFocus 23
                Exit Function
            End If
            If txtPCC(24) <> "N" Then
                ShowMsg "此客戶為國內同業, 不可寄專利雙週報 !"
                txtPCC(24).SetFocus
                txtPCC_GotFocus 24
                Exit Function
            End If
        End If
   End If
   'end 2021/11/26
   If Trim(txtPCC(5)) = "" Then
      ShowMsg "接洽人名稱不可為空白 !"
      txtPCC(5).SetFocus
      Exit Function
   ElseIf Trim(txtPCC(5)) = Trim(lblCustName) Then
      ShowMsg "接洽人名稱不可與申請人相同 !"
      txtPCC(5).SetFocus
      Exit Function
   End If
   
   If Check1.Value <> 1 Then

      If txtPCC(22) = "" Then
         ShowMsg "聯絡地址不可空白!"
         txtPCC(22).SetFocus
         Exit Function
      'Added by Morgan 2017/11/1
      'Modified by Morgan 2022/3/28 +改以字數計算
      ElseIf Not CheckLengthIsOK(txtPCC(22), txtPCC(22).MaxLength, , , True) Then
         ShowMsg "聯絡地址過長!"
         txtPCC(22).SetFocus
         Exit Function
      '2011/10/18 add by sonia
      'Modify by Amy 2016/05/16 勾選「是臺灣聯絡地址」才檢查
      ElseIf Check2.Value = 1 Then
         
         'Added by Morgan 2022/3/28
         If txtPCC(21) = "" Then
            ShowMsg "郵遞區號不可空白!"
            txtPCC(21).SetFocus
            Exit Function
         End If
         'end 2022/3/28
         
        If CheckTaiwanAddr(txtPCC(22), "000", "聯絡地址") = False Then
            txtPCC(22).SetFocus
            Exit Function
        End If
      '2011/10/18 end
      End If
   End If
   
   'Add by Morgan 2009/1/9
   If CheckContactAddr = False Then Exit Function
   
   'Add by Amy 2025/05/19 名稱前後有空白要Trim再查
   If Left(txtPCC(5), 1) = " " Or Left(txtPCC(5), 1) = "　" Then
      txtPCC(5) = Mid(txtPCC(5), 2)
   End If
   If Right(txtPCC(5), 1) = " " Or Right(txtPCC(5), 1) = "　" Then
      txtPCC(5) = Mid(txtPCC(5), 1, Len(txtPCC(5)) - 1)
   End If
   'end 2025/05/19
   
   '檢查聯絡人是否重複
   Set RsTemp = rsContact.Clone
   With RsTemp
   If .RecordCount > 0 Then
      .MoveFirst
      Do While Not .EOF
         If .Fields("PCC02") <> txtPCC(2) And UCase("" & .Fields("PCC05")) = UCase(txtPCC(5)) Then
            If MsgBox("名稱與接洽人[" & .Fields("PCC02") & "]重複！是否仍然要加入？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
               Exit Function
            End If
         End If
         .MoveNext
      Loop
   End If
   End With
   Set RsTemp = Nothing
   TxtValidate1 = True
End Function

Private Sub UpdateContact()
   With rsContact
   If txtPCC(2) = "" Then
      m_iConEditMode = 1
      txtPCC(2) = getNewNo
      .AddNew
   Else
      If .RecordCount > 0 Then
         .MoveFirst
         .Find "PCC02='" & txtPCC(2) & "'"
         If .EOF Then
            .AddNew
         End If
      Else
         .AddNew
      End If
   End If
   For Each oText In txtPCC
      'Modify by Morgan 2008/12/5 去除跳行符號
      '.Fields("PCC" & Format(oText.Index, "00")) = oText.Text
      'Add by Amy 2025/05/19 名稱前後有空白要Trim
      If oText.Index = 5 Then
         If Left(oText.Text, 1) = " " Or Left(oText.Text, 1) = "　" Then
            oText.Text = Mid(oText.Text, 2)
         End If
         If Right(oText.Text, 1) = " " Or Right(oText.Text, 1) = "　" Then
            oText.Text = Mid(oText.Text, 1, Len(oText.Text) - 1)
         End If
      End If
      'end 2025/05/19
      .Fields("PCC" & Format(oText.Index, "00")) = PUB_StringFilter(oText.Text)
   Next
   .UPDATE
   End With
End Sub

Private Function ModRecord() As Boolean
   Dim stSQL As String, stSet As String, stCols As String, stValues As String
   Dim bDifference As Boolean, bAddNew As Boolean
   Dim idx As Integer
   Dim bMailChanged As Boolean 'Added by Morgan 2024/6/3
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   '更新聯絡人資料
   If rsContact.RecordCount = 0 Then
      If rsContactOld.RecordCount > 0 Then
         '刪除聯絡人資料
         'Modified by Morgan 2017/8/30 批次刪除未紀錄所有接洽人改逐筆刪除
         'stSQL = "delete from potcustcont where pcc01='" & txtCuNo & "'"
         'Pub_SeekTbLog stSQL
         'cnnConnection.Execute stSQL, intI
         rsContactOld.MoveFirst
         Do While Not rsContactOld.EOF
            If "" & rsContactOld.Fields("pcc08") <> "" Then bMailChanged = True 'Added by Morgan 2024/6/3
            stSQL = "delete from potcustcont where pcc01='" & txtCuNo & "' and pcc02='" & rsContactOld.Fields("pcc02") & "'"
            Pub_SeekTbLog stSQL
            cnnConnection.Execute stSQL, intI
            rsContactOld.MoveNext
         Loop
         'end 2017/8/30
      End If
   Else
      '刪除聯絡人(原來的編號在新的聯絡人資料中找不到的)
      With rsContactOld
      If .RecordCount > 0 Then
         .MoveFirst
         Do While Not .EOF
            rsContact.MoveFirst
            rsContact.Find "PCC02='" & .Fields("PCC02") & "'"
            If rsContact.EOF Then
               If "" & rsContactOld.Fields("pcc08") <> "" Then bMailChanged = True 'Added by Morgan 2024/6/3
               '刪除聯絡人資料
               stSQL = "delete from potcustcont where pcc01='" & txtCuNo & "' and pcc02='" & .Fields("PCC02") & "'"
               Pub_SeekTbLog stSQL
               cnnConnection.Execute stSQL, intI
            End If
            .MoveNext
         Loop
      End If
      End With
      '更新(新增)連絡人
      With rsContact
      .MoveFirst
      Do While Not .EOF
         If rsContactOld.RecordCount = 0 Then
            bAddNew = True
         Else
            rsContactOld.MoveFirst
            rsContactOld.Find "PCC02='" & .Fields("PCC02") & "'"
            If rsContactOld.EOF Then
               bAddNew = True
            Else
               bAddNew = False
            End If
         End If
         
         '新增
         If bAddNew = True Then
            stCols = "PCC01"
            stValues = "'" & txtCuNo & "'"
            For idx = 1 To .Fields.Count - 1
               If .Fields(idx) <> "" Then
                  stCols = stCols & "," & .Fields(idx).Name
                  If .Fields(idx).Name = "PCC11" Then
                     stValues = stValues & "," & .Fields(idx)
                  Else
                     stValues = stValues & ",'" & ChgSQL(.Fields(idx)) & "'"
                  End If
               End If
            Next
            
            If "" & .Fields("pcc08") <> "" Then bMailChanged = True 'Added by Morgan 2024/6/3
            
            stSQL = "INSERT INTO PotCustCont (" & stCols & ") Values (" & stValues & ")"
            Pub_SeekTbLog stSQL
            cnnConnection.Execute stSQL, intI
         '修改
         Else
            bDifference = False
            stSet = ""
            For idx = 2 To .Fields.Count - 1
               If "" & .Fields(idx) <> "" & rsContactOld.Fields(idx) Then
                  bDifference = True
                  If .Fields(idx).Name = "PCC11" Then
                     stSet = stSet & "," & .Fields(idx).Name & "=" & CNULL(.Fields(idx), True)
                  Else
                     stSet = stSet & "," & .Fields(idx).Name & "=" & CNULL(ChgSQL(.Fields(idx)))
                  End If
               End If
            Next
            If bDifference = True Then
               If "" & rsContactOld.Fields("pcc08") <> "" & .Fields("pcc08") Then bMailChanged = True 'Added by Morgan 2024/6/3
               
               stSet = Mid(stSet, 2)
               stSQL = "begin user_data.user_enabled:=1; Update PotCustCont set " & stSet & " where PCC01='" & txtCuNo & "' and PCC02='" & .Fields("PCC02") & "'; end;"
               Pub_SeekTbLog stSQL
               cnnConnection.Execute stSQL
            End If
         End If
         .MoveNext
      Loop
      End With
   End If
   
   If bMailChanged Then PUB_ECustEmailChangeInform txtCuNo & "0", strECustMsg, "2" 'Added by Morgan 2024/6/3 全E化客戶任一信箱異動時(含聯絡人)，發信通知智權人員(一併彈提醒)--杜協理
   
   cnnConnection.CommitTrans
   ModRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical
End Function

Public Sub SetState(p_iMode As Integer)
   If p_iMode = 2 Then
      cmdOK(1).Visible = True
      cmdContact(1).Visible = True
      cmdContact(2).Visible = True
      If Pub_StrUserSt03 = "M51" Then
         cmdContact(3).Visible = True
      End If
   Else
      Me.fraContact.Enabled = False
   End If
End Sub

'Add by Amy 2016/05/16
Private Sub txtPCC_LostFocus(Index As Integer)
    Dim strZipCode As String, strAddr As String, strCountry As String, strCityN As String, strIndArea As String, strNewArea As String, strROC As String
    Dim bolMany As Boolean, intArea As Integer
    
    If Index <> 22 Or (Index = 22 And cmdOK(1).Visible = False) Then Exit Sub
    If Check2.Value = 0 Or Trim(txtPCC(Index)) = MsgText(601) Then Exit Sub
    
    If Check2.Value = 1 Then
        Me.txtPCC(Index) = ReplaceAddrTW(Me.txtPCC(Index))
AgainCheck1:
        strROC = ""
        strAddr = Me.txtPCC(Index)
        If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
        If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
        If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
        '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
        strIndArea = "True"
        strAddr = ReplaceIndArea(strAddr, strIndArea)
        If strIndArea = "True" Then strIndArea = MsgText(601)
        If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
            strIndArea = "新竹" & strIndArea
            strAddr = Mid(strAddr, 3)
        End If
        '第3個字是 縣 / 市
        If Mid(strAddr, 3, 1) = "市" Or Mid(strAddr, 3, 1) = "縣" Or Mid(strAddr, 1, 3) = "釣魚臺" Or Mid(strAddr, 1, 3) = "海南島" Then
            'Modify by Amy 2018/12/19 +判斷第七個字 ex:嘉義縣阿里山鄉 X80024
            If Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" _
              Or Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" _
              Or Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
                '傳入地址前6個字抓到郵遞區號
                intArea = 6
                strZipCode = GetPostZip(Left(strAddr, 6), 6, , , bolMany)
                '傳入地址前5個字取郵遞區號
                If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddr, 5), 5, , , bolMany): intArea = 5
                '抓到郵遞區號
                If strZipCode <> MsgText(601) Then
                    If bolMany = True Then
                        '多筆以縣/市+鄉/鎮/市/區及路名查
                        bolMany = False
                        strZipCode = GetZipCode_Tai(3, strAddr, , bolMany)
                        If strZipCode <> MsgText(601) Then
                            '限制縣/市+鄉/鎮/市/區及路名查:一筆-直接帶/多筆-進查詢畫面
                            If bolMany = False Then
                                Call ChkZipData(2, Me.txtPCC(Index), strZipCode)
                            Else
                                Call ChkZipData(1, Me.txtPCC(Index), strZipCode, intArea)
                            End If
                        End If
                    Else
                        '非多筆
                        Call ChkZipData(2, Me.txtPCC(Index), strZipCode, intArea)
                    End If
                Else
                    '判斷是否有此區/鄉/鎮
                    strZipCode = GetPostZip(Mid(strAddr, 4, intArea - 3), intArea - 3, , , bolMany, "Pzd03")
                    If strZipCode <> MsgText(601) Then
                        '區別錯,進入查詢畫面
                        Call ChkZipData(3, Me.txtPCC(Index), strZipCode, intArea)
                    Else
                        '當作沒區只有路 ex:新竹縣or市園區二路
                        bolMany = False
                        strZipCode = GetZipCode_Tai(2, strAddr, intArea, bolMany, strNewArea)
                        If strZipCode <> MsgText(601) Then
                            '以縣/市及路名查:一筆-直接帶/多筆-進查詢畫面
                            If bolMany = False Then
                                Me.txtPCC(Index) = strROC & Left(strAddr, 3) & strNewArea & strIndArea & Mid(strAddr, 4)
                                Call ChkZipData(2, Me.txtPCC(Index), strZipCode, intArea)
                            Else
                                intArea = 0
                                Call ChkZipData(4, Me.txtPCC(Index), strZipCode, intArea)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
            '無鄉/鎮/市/區
            Else
                '以路/街 抓是否有zip
                strZipCode = GetZipCode_Tai(2, strAddr, intArea, bolMany, strNewArea)
                If strZipCode <> MsgText(601) Then
                    If bolMany = True Then
                        '多筆
                        intArea = 0
                        Call ChkZipData(4, Me.txtPCC(Index), strZipCode, intArea)
                    Else
                        '非多筆
                        Me.txtPCC(Index) = strROC & Left(strAddr, 3) & strNewArea & strIndArea & Mid(strAddr, 4)
                        Call ChkZipData(2, Me.txtPCC(Index), strZipCode, intArea)
                    End If
                End If
                '都抓不到ZipCode
                If strZipCode = MsgText(601) Then
                    If CheckTaiwanAddr_Tai(Me.txtPCC(Index), Me.txtPCC(Index - 1), "000", "", strZipCode, , False, Me.Name) = False Then
                        If MsgBox("無法依地址聯絡地址" & "的郵遞區號，請問是臺灣地址嗎？", vbYesNo + vbCritical) = vbYes Then
                            If strZipCode = "格式錯誤" Then
                                frm100135.Show vbModal
                                Call ChkZipData(9, Me.txtPCC(Index), strZipCode)
                            Else
                                Call ChkZipData(3, Me.txtPCC(Index), strZipCode)
                            End If
                            Exit Sub
                        Else
                            '非臺灣
                            Check2.Value = 0
                            Exit Sub
                        End If
                    End If
                End If
            End If
        
        '第3個字無 縣 / 市
        Else
            '傳入地址前2個字判斷是否有其縣/市
            strCityN = "Pzd02"
            strZipCode = GetPostZip(Left(strAddr, 2), 3, 1, , bolMany, "Pzd02", strCityN)
            If strZipCode <> MsgText(601) Then
                If bolMany = False Then
                    '只有一筆
                    Me.txtPCC(Index) = strROC & strCityN & strIndArea & Mid(strAddr, 3)
                    GoTo AgainCheck1
                Else
                    '新竹、嘉義會有2筆
                    intArea = 0
                    Call ChkZipData(5, Me.txtPCC(Index), strZipCode, intArea)
                End If
            '查無郵遞區號
            Else
                If MsgBox("無法依地址帶出聯絡地址的郵遞區號，請問是臺灣地址嗎？", vbYesNo + vbCritical) = vbYes Then
                    frm100135.Show vbModal
                    Me.txtPCC(Index).SetFocus
                    Exit Sub
                Else
                    '非臺灣
                    Check2.Value = 0
                    Exit Sub
                End If
            End If
                        
        End If
        
    End If
End Sub

Private Sub txtPCC_Validate(Index As Integer, Cancel As Boolean)
   Dim stMsg As String 'Add by Amy 2024/04/10
   'Modify by Amy 2024/04/10 加電子報相關欄位檢查
   'Modify by Amy 2024/05/14 顧問電子報可輸Y 拆開
   If Index = 23 And txtPCC(Index) <> "N" And txtPCC(Index) <> "Y" And txtPCC(Index) <> MsgText(601) Then
      MsgBox "顧問電子報只允許輸入N或Y,不可輸小寫或全型..."
      Cancel = True
      Call txtPCC_GotFocus(Index)
      Exit Sub
   ElseIf Index = 10 Or Index = 24 Then
      If Index = 10 Then
         stMsg = "國內電子報"
      Else
         stMsg = "專利雙週報"
      End If
   'end 2024/05/14
      If txtPCC(Index) <> "N" And txtPCC(Index) <> MsgText(601) Then
         MsgBox stMsg & "只允許輸入N,不可輸小寫或全型..."
         Cancel = True
         Call txtPCC_GotFocus(Index)
         Exit Sub
      End If
   ElseIf Index = 8 Then
      If txtPCC(Index) <> "" Then
         If PUB_CheckMail(txtPCC(Index)) = False Then
            Cancel = True
         End If
      End If
   End If
End Sub

'Add by Morgan 2009/1/9
'檢查預設接洽人聯絡地址是否與客戶聯絡地址一致
Private Function CheckContactAddr() As Boolean
   '客戶無預設接洽人或新增接洽人時不必檢查
   If Val(strCU127) = 0 Or Val(txtPCC(2)) = 0 Then
      CheckContactAddr = True
   '若為預設接洽人且有輸入聯絡地址
   ElseIf Val(strCU127) = Val(txtPCC(2)) And Trim(txtPCC(21)) & Trim(txtPCC(22)) <> "" Then
      'Modified by Morgan 2022/9/2
      'strExc(0) = "select cu30,cu31 from customer where cu01='" & txtCuNo & "' and cu02='0'"
      'intI = 1
      'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      'If intI = 1 Then
      '   With RsTemp
      '   If "" & .Fields("cu30") = Trim(txtPCC(21)) And "" & .Fields("cu31") = Trim(txtPCC(22)) Then
      '      CheckContactAddr = True
      '   Else
      '      MsgBox "預設接洽人的聯絡地址必須與客戶聯絡地址【" & .Fields("cu30") & .Fields("cu31") & "】" & vbCrLf & "相同！"
      '   End If
      '   End With
      'End If
       MsgBox "預設接洽人不可有【聯絡地址】！" & vbCrLf & "系統將自動勾選【聯絡地址同客戶】！", vbExclamation
       Check1.Value = vbChecked
       'end 2022/9/2
   Else
      CheckContactAddr = True
   End If
End Function

'Add by Amy 2016/05/16
'Modify by Amy 2021/12/14 原:As TextBox
Private Sub ChkZipData(ByVal intChoose As Integer, ByRef objTxt As Control, Optional ByRef stZipCode As String = "", Optional ByRef intArea As Integer = 0, Optional ByRef stCountryCode As String = "")
    Dim intZipIdx As Integer '地址相對應Zip欄位
    Dim strMsg As String, strAddr As String
    Dim intCount As Integer
    
     Select Case objTxt.Index
        Case 22
            strMsg = "聯絡地址"
            intZipIdx = 21
        Case Else
            intZipIdx = objTxt.Index
    End Select
    
    Select Case intChoose
        Case 1 'ZipCode多筆(同區/鄉 ZipCode不同)
            '且與畫面上欄位資料前3碼不同或空值,彈郵遞區號查詢畫面
            If InStr(stZipCode, Left(Trim(txtPCC(intZipIdx)), 3)) = 0 Or Trim(txtPCC(intZipIdx)) = MsgText(601) Then
                If Trim(txtPCC(intZipIdx)) <> MsgText(601) Then MsgBox strMsg & "郵遞區號有誤,請選擇正確郵遞區號！"
                Call frm100134.SetParent(Me)
                Me.Hide
                frm100134.BFormZip = "txtPCC(" & intZipIdx & ")"
                frm100134.strPrevFormMon = "frm210101_1"
                frm100134.GetStreet objTxt.Text, 1, intArea, stZipCode
                Call frm100134.QueryData
                frm100134.Show
                Exit Sub
            End If
        Case 2 'ZipCode非多筆
            '判斷抓到的郵遞區號是否與畫面上欄位資料前3碼相同
            If Left(txtPCC(intZipIdx), 3) <> stZipCode Then
                If txtPCC(intZipIdx) <> MsgText(601) Then MsgBox strMsg & "郵遞區號有誤,系統將自動更正！", , MsgText(5)
                txtPCC(intZipIdx) = stZipCode
                txtPCC_GotFocus (objTxt.Index)
            End If
            Exit Sub
        Case 3, 4, 5 '抓不到ZipCode-3.區錯/4.只有路且郵遞區號為多筆/5.抓到2個字縣市,但多筆
            MsgBox strMsg & "無法解析郵遞區號，請由下一畫面選取！"
            Call frm100134.SetParent(Me)
            Me.Hide
            frm100134.BFormZip = "txtPCC(" & intZipIdx & ")"
            frm100134.strPrevFormMon = "frm210101_1"
            frm100134.GetStreet objTxt.Text, IIf(intChoose = 6, 2, intChoose), intArea, stZipCode
            Call frm100134.QueryData
            frm100134.Show
            Exit Sub
        Case 9
            If txtPCC(objTxt.Index).Enabled = True Then
                txtPCC(objTxt.Index).SetFocus
                txtPCC_GotFocus (objTxt.Index)
            End If
    End Select
End Sub

Private Function FormCheck() As Boolean
    Dim strZipCode As String, strAddr As String, strCountry As String, strIndArea As String, strROC As String
    Dim bolMany As Boolean, intArea As Integer
    Dim bCancel As Boolean 'Add by Amy 2024/04/10
    
    FormCheck = False
    
    If Check2.Value = 1 And Trim(txtPCC(22)) <> MsgText(601) And txtPCC(21) = MsgText(601) Then
        MsgBox "郵遞區號不可為空 ！"
        txtPCC(21).SetFocus
        Exit Function
    End If
    
    '聯絡地址判斷
    txtPCC(22) = ReplaceAddrTW(txtPCC(22))
    strROC = ""
    strAddr = txtPCC(22)
    If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
    If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
    If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
    '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
    strIndArea = "True"
    strAddr = ReplaceIndArea(strAddr, strIndArea)
    If strIndArea = "True" Then strIndArea = MsgText(601)
    If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
        strIndArea = "新竹" & strIndArea
        strAddr = Mid(strAddr, 3)
    End If
    intArea = 6
    strZipCode = GetPostZip(Left(strAddr, 6), 6, , strCountry, bolMany)
    '傳入地址前5個字取郵遞區號
    If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddr, 5), 5, , strCountry, bolMany): intArea = 5
    If InStr(strZipCode, Left(txtPCC(21), 3)) = 0 Then MsgBox "地址對應之郵遞區號有誤請確認！": txtPCC(22).SetFocus: Exit Function
    
    'Add by Amy 2024/04/10 電子報相關欄位檢查
    '國內電子報
    If txtPCC(10).Text <> MsgText(601) Then
      Call txtPCC_Validate(10, bCancel)
      If bCancel = True Then Exit Function
    End If
    '專利雙週報
    If txtPCC(24).Text <> MsgText(601) Then
      Call txtPCC_Validate(24, bCancel)
      If bCancel = True Then Exit Function
    End If
    '顧問電子報
    If txtPCC(23).Text <> MsgText(601) Then
      Call txtPCC_Validate(23, bCancel)
      If bCancel = True Then Exit Function
    End If
    'end 2024/04/10
    
    FormCheck = True
End Function
'end 2016/05/16

'Added by Morgan 2025/5/27
'檢查是否有被設為個案接洽人
'Modified by Morgan 2025/10/13 +pName
Private Function fnCaseCheck(pCuNo As String, pPCC02 As String, Optional pName As String) As Boolean
   strExc(0) = "Select pa01||'-'||pa02||'-'||pa03||'-'||pa04 From patent Where pa26='" & pCuNo & "0' and pa149='" & pPCC02 & "'" & _
      " union Select tm01||'-'||tm02||'-'||tm03||'-'||tm04 From TRADEMARK Where tm23='" & pCuNo & "0' and tm123='" & pPCC02 & "'" & _
      " union Select sp01||'-'||sp02||'-'||sp03||'-'||sp04 From SERVICEPRACTICE Where sp08='" & pCuNo & "0' and sp78='" & pPCC02 & "'" & _
      " union Select lc01||'-'||lc02||'-'||lc03||'-'||lc04 From LAWCASE Where lc11='" & pCuNo & "0' and lc42='" & pPCC02 & "'" & _
      " union Select hc01||'-'||hc02||'-'||hc03||'-'||hc04 From HIRECASE Where hc05='" & pCuNo & "0' and hc23='" & pPCC02 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      MsgBox pPCC02 & "(" & pName & ") 有被設為下列案件的接洽人，不可刪除！" & vbCrLf & vbCrLf & RsTemp.GetString, vbExclamation
   Else
      fnCaseCheck = True
   End If
End Function
