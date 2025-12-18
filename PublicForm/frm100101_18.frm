VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_18 
   BorderStyle     =   1  '單線固定
   Caption         =   "接洽人資料查詢"
   ClientHeight    =   5784
   ClientLeft      =   420
   ClientTop       =   4416
   ClientWidth     =   8376
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5784
   ScaleWidth      =   8376
   Begin VB.CommandButton CmdOk1 
      Caption         =   "地址條"
      Enabled         =   0   'False
      Height          =   400
      Index           =   2
      Left            =   6210
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   510
      Width           =   1230
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "回前畫面"
      Height          =   400
      Index           =   0
      Left            =   6210
      TabIndex        =   26
      Top             =   60
      Width           =   1230
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束"
      Height          =   400
      Index           =   1
      Left            =   7440
      TabIndex        =   25
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox txtCuNo 
      Height          =   300
      Left            =   1035
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1092
   End
   Begin VB.Frame fraContact 
      Height          =   3228
      Left            =   72
      TabIndex        =   16
      Top             =   2490
      Width           =   8160
      Begin VB.CheckBox Check1 
         Caption         =   "聯絡地址同客戶"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3285
         TabIndex        =   3
         Top             =   552
         Width           =   1680
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   288
         Index           =   33
         Left            =   4980
         TabIndex        =   13
         Top             =   2184
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
         Height          =   288
         Index           =   31
         Left            =   4980
         TabIndex        =   11
         Top             =   1896
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
         Left            =   1044
         TabIndex        =   12
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
         Index           =   30
         Left            =   1044
         TabIndex        =   10
         Top             =   1920
         Width           =   2976
         VariousPropertyBits=   679493659
         MaxLength       =   20
         Size            =   "5249;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LINE ID："
         Height          =   180
         Index           =   28
         Left            =   4188
         TabIndex        =   37
         Top             =   2232
         Width           =   804
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "電話："
         Height          =   180
         Index           =   12
         Left            =   504
         TabIndex        =   36
         Top             =   1944
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "手機："
         Height          =   180
         Index           =   11
         Left            =   480
         TabIndex        =   35
         Top             =   2256
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "傳真："
         Height          =   180
         Index           =   9
         Left            =   4452
         TabIndex        =   34
         Top             =   1992
         Width           =   540
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   24
         Left            =   6768
         TabIndex        =   5
         Top             =   480
         Width           =   336
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   23
         Left            =   6456
         TabIndex        =   6
         Top             =   780
         Width           =   336
         VariousPropertyBits=   671105052
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   10
         Left            =   6708
         TabIndex        =   9
         Top             =   1620
         Width           =   336
         VariousPropertyBits=   671105055
         MaxLength       =   1
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   540
         Index           =   22
         Left            =   1032
         TabIndex        =   7
         Top             =   1080
         Width           =   6960
         VariousPropertyBits=   -1466941409
         MaxLength       =   70
         ScrollBars      =   2
         Size            =   "12277;952"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   21
         Left            =   1032
         TabIndex        =   4
         Top             =   780
         Width           =   1200
         VariousPropertyBits=   671105055
         MaxLength       =   10
         Size            =   "2117;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   2
         Left            =   1035
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   600
         VariousPropertyBits=   671105055
         Size            =   "1058;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   636
         Index           =   13
         Left            =   1032
         TabIndex        =   14
         Top             =   2496
         Width           =   6972
         VariousPropertyBits=   -1466941409
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "12303;1111"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   8
         Left            =   1032
         TabIndex        =   8
         Top             =   1620
         Width           =   3180
         VariousPropertyBits=   671105055
         MaxLength       =   50
         Size            =   "5609;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPCC 
         Height          =   300
         Index           =   5
         Left            =   1032
         TabIndex        =   2
         Top             =   480
         Width           =   1872
         VariousPropertyBits=   671105055
         MaxLength       =   10
         Size            =   "3307;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCUID1 
         Height          =   300
         Left            =   1680
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   180
         Width           =   6345
         VariousPropertyBits=   -2147467233
         BackColor       =   16777215
         Size            =   "11192;529"
         Caption         =   "LblFM2"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄發專利雙週報：      （N:不寄）"
         Height          =   180
         Index           =   21
         Left            =   4956
         TabIndex        =   30
         Top             =   540
         Width           =   2952
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄顧問電子報：        （Y:寄/N:不寄）"
         Height          =   180
         Index           =   1
         Left            =   4800
         TabIndex        =   29
         Top             =   840
         Width           =   3252
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄電子報：         （N:不寄）"
         Height          =   180
         Index           =   17
         Left            =   5400
         TabIndex        =   28
         Top             =   1680
         Width           =   2556
      End
      Begin VB.Label Label63 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "聯絡地址："
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   1128
         Width           =   900
      End
      Begin VB.Label Label63 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "郵遞區號："
         Height          =   180
         Index           =   1
         Left            =   132
         TabIndex        =   21
         Top             =   828
         Width           =   900
      End
      Begin VB.Label Label63 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "E-MAIL："
         Height          =   180
         Index           =   5
         Left            =   252
         TabIndex        =   20
         Top             =   1644
         Width           =   780
      End
      Begin VB.Label Label63 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "編號："
         Height          =   180
         Index           =   7
         Left            =   495
         TabIndex        =   19
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   14
         Left            =   492
         TabIndex        =   18
         Top             =   2496
         Width           =   540
      End
      Begin VB.Label Label63 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "名稱："
         Height          =   180
         Index           =   0
         Left            =   492
         TabIndex        =   17
         Top             =   552
         Width           =   540
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   90
      Top             =   2130
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
      Bindings        =   "frm100101_18.frx":0000
      Height          =   1530
      Left            =   90
      TabIndex        =   15
      Top             =   960
      Width           =   8175
      _ExtentX        =   14415
      _ExtentY        =   2709
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
      ColumnCount     =   5
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
            ColumnWidth     =   972.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3060.284
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1764.284
         EndProperty
      EndProperty
   End
   Begin MSForms.Label lblCustAddress 
      Height          =   510
      Left            =   1050
      TabIndex        =   33
      Top             =   420
      Width           =   4875
      Size            =   "8599;900"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCustName 
      Height          =   300
      Left            =   2160
      TabIndex        =   32
      Top             =   120
      Width           =   3735
      VariousPropertyBits=   27
      Size            =   "6588;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "聯絡地址："
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   24
      Top             =   450
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   23
      Top             =   150
      Width           =   900
   End
End
Attribute VB_Name = "frm100101_18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2024/03/13 拿掉A4名條印表機Combo1的物件和程式
'Memo by Lydia 2022/01/10 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、lblCustAddress、lblCustName、textCUID1、txtPCC(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
'Create by Morgan 2008/8/5
Option Explicit
Public cmdState As Integer
Dim m_bReadGrid As Boolean '是否要讀取被點選聯絡人資料
Dim rsContact As ADODB.Recordset
Dim SeekPrintL As Integer
Dim SeekPrint As Integer
Dim mPrevForm As Form 'Added by Lydia 2016/10/28

'Added by Lydia 2016/10/28
Public Sub SetParent(ByRef fm As Form)
   Set mPrevForm = fm
End Sub

Sub StrMenu()
   'Add By Sindy 2011/01/03 檢查國內外權限
   If CheckSR12(Me.Tag) = False Then
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
   pub_QL05 = pub_QL05 & IIf(PUB_CheckQL05("編號：" & Me.Tag & "(接洽人資料)") = "", "", ";編號：" & Me.Tag & "(接洽人資料)") 'Add By Sindy 2025/8/13
   
   txtCuNo = Left(Me.Tag, 8)
   OpenContactTable
End Sub

Private Sub cmdok1_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
         
      'Add by Morgan 2008/8/26
      Case 2 '地址條
         If txtPCC(2) = "" Then
            MsgBox "請勾選欲列印地址條的資料!!!", vbExclamation + vbOKOnly
         Else
            'Added by Lydia 2016/10/28 從申請人查詢來,地址條採16格列印
            If TypeName(mPrevForm) <> "Nothing" Then
                If PUB_AddAddressA4List(Mid(txtCuNo & "000", 1, 9) & "-" & txtPCC(2), strExc(0)) Then
                End If
                If Val(strExc(0)) > 0 Then
                   'Modified by Lydia 2017/11/22 +國內
                   CmdOk1(2).Caption = "國內A4名條 (" & Val(strExc(0)) & ")"
                   mPrevForm.cmdOK(4).Caption = "國內A4名條 (" & Val(strExc(0)) & ")"
                   'end 2017/11/22
                End If
            Else
            'end 2016/10/28
                'Modified by Lydia 2017/11/03 改成操作介面
'                Screen.MousePointer = vbHourglass
'                Set Printer = Printers(Combo1.ListIndex)
'                Load frm083014
'                frm083014.Hide
'                frm083014.Opt1(0).Value = True
'                frm083014.m_ContactNo = txtPCC(2)
'                frm083014.Text1(0).Text = txtCuNo
'                frm083014.Text1(4).Text = "1"
'                frm083014.SetPrinter Printer.DeviceName
'                frm083014.cmdPrint_Click
'                Unload frm083014
'                Screen.MousePointer = vbDefault
                frm083014.iStiu = 1
                frm083014.Show
                Me.Hide
                'end 2017/11/03
            End If
         End If
   End Select
End Sub

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
   MoveFormToCenter Me
   cmdState = -1
   textCUID1.BackColor = Me.BackColor
   SeekPrintL = Printer.Orientation
   'PUB_SetPrinter Me.Name, Me.Combo1, , , SeekPrint 'Mark by Lydia 2024/03/13
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '若印表機或偏移值有變動, 則更新列印設定
   'Mark by Lydia 2024/03/13
   'If Me.Combo1.Text <> Me.Combo1.Tag Then
   '    PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, 0, 0, Me.Combo1.Text
   'End If
   'Set Printer = Printers(SeekPrint)
   'If SeekPrintL <> 0 Then
   '    Printer.Orientation = SeekPrintL
   'End If
   'end 2024/03/13
   Set frm100101_18 = Nothing
End Sub

Private Sub OpenContactTable()
   
On Error GoTo Checking
   
   strExc(0) = "select cu04 CName,RPAD( NVL(CU30,' '),11,' ')||CU31 CAddr,pcc.*" & _
      " from customer,potcustcont pcc where cu01='" & txtCuNo.Text & "' and cu02='0' and pcc01(+)=cu01 order by pcc02"
   intI = 1
   Set rsContact = ClsLawReadRstMsg(intI, strExc(0))
   Set Adodc1.Recordset = rsContact
   DataGrid1.col = 0
   DataGrid1.CurrentCellVisible = True
   If rsContact.RecordCount > 0 Then
      If pub_QL04 <> "" Then InsertQueryLog (rsContact.RecordCount) 'Add By Sindy 2025/8/13
      lblCustName = "" & rsContact.Fields("CName")
      lblCustAddress = "" & rsContact.Fields("CAddr")
      ReadContact
   Else
      If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
   End If
   
Checking:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
   End If
   
End Sub

'讀取該筆聯絡人資料
Private Sub ReadContact()
   Dim CUID(1 To 6) As String
   Dim oText As Object
   
   ClearField1
   With Adodc1.Recordset
      If Not (.EOF Or .BOF) Then
         For Each oText In txtPCC
            oText = "" & .Fields("PCC" & Format(oText.Index, "00"))
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
   Dim oText As Object
   For Each oText In txtPCC
      oText.Text = Empty
   Next
   textCUID1.Text = Empty
   Check1.Value = 1
End Sub
' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
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
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub
