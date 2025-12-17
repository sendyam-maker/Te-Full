VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc3280 
   AutoRedraw      =   -1  'True
   Caption         =   "託收票據資料查詢"
   ClientHeight    =   5232
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5232
   ScaleWidth      =   8760
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   150
      Width           =   3500
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3240
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   870
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   870
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5868
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1632
      Width           =   480
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   17
      Top             =   4712
      Width           =   1260
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   528
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc3280.frx":0000
      Height          =   2500
      Left            =   240
      TabIndex        =   9
      Top             =   2120
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   4382
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "託收票據資料"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "a0g02"
         Caption         =   "託收銀行"
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
      BeginProperty Column01 
         DataField       =   "a0e20"
         Caption         =   "託收帳號"
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
      BeginProperty Column02 
         DataField       =   "a0e02"
         Caption         =   "票據號碼"
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
         DataField       =   "a0e14"
         Caption         =   "託收日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "a0e10"
         Caption         =   "到期日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "a0e21"
         Caption         =   "兌現日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "a0e11"
         Caption         =   "票據金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "a0e05"
         Caption         =   "往來類別"
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
      BeginProperty Column08 
         DataField       =   "contect"
         Caption         =   "往來對象"
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
         Size            =   655
         BeginProperty Column00 
            ColumnWidth     =   1967.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1344.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1116.284
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1068.094
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   972.284
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   5328
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   2000
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   550
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
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   1260
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   1260
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   6
      Top             =   1620
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3240
      TabIndex        =   7
      Top             =   1620
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   348
      TabIndex        =   21
      Top             =   190
      Width           =   972
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "Y:是 N:否 空白:全部"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6420
      TabIndex        =   20
      Top             =   1656
      Width           =   2064
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "是否兌現"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4920
      TabIndex        =   19
      Top             =   1644
      Width           =   972
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "合計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6288
      TabIndex        =   18
      Top             =   4712
      Width           =   492
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   16
      Top             =   900
      Width           =   252
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "託收帳號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   15
      Top             =   900
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "票據號碼"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   348
      TabIndex        =   14
      Top             =   552
      Width           =   972
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   13
      Top             =   1620
      Width           =   252
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "託收日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   12
      Top             =   1620
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3000
      TabIndex        =   11
      Top             =   1260
      Width           =   252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "託收銀行"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   10
      Top             =   1260
      Width           =   972
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1892
      Left            =   240
      Top             =   108
      Width           =   8292
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc3280"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoadodc1 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strSql As String
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/04/17


'Add by Sindy 2020/04/17
Private Sub SetCompN()
    strCmpN = "": strCmp = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, True, True)
End Sub

'Add by Sindy 2020/04/17
Private Sub CboCmp_GotFocus()
    TextInverse CboCmp
End Sub

Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label11 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/17

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8850
   Me.Height = 5670 'Modify by Amy 2023/08/18 原:5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   'Add by Sindy 2020/04/17 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/04/17
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   'Add by Morgan 2006/7/25
   PUB_SetAccount Combo1
   PUB_SetAccount Combo2
   'end 2006/7/25
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc3280 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   '20140120START Modify By eric
   'Modify By Sindy 2020/4/20
   'adoadodc1.Open "select * from acc0e0 where a0e23 = '" & IIf(Text5 = "2", "J", "1") & "' and a0e20 >= '" & Combo1 & "' and a0e20 <= '" & Combo2 & "' and a0e19 >= '" & Text1 & "' and a0e19 <= '" & Text2 & "' and a0e14 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e14 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e21 = 0 and a0e25 = 0 and a0e04 = '" & MsgText(18) & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "select * from acc0e0 where a0e23 = '" & Left(CboCmp.Text, 1) & "' and a0e20 >= '" & Combo1 & "' and a0e20 <= '" & Combo2 & "' and a0e19 >= '" & Text1 & "' and a0e19 <= '" & Text2 & "' and a0e14 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e14 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e21 = 0 and a0e25 = 0 and a0e04 = '" & MsgText(18) & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenStatic, adLockReadOnly
   'adoadodc1.Open "select * from acc0e0 where a0e20 >= '" & Combo1 & "' and a0e20 <= '" & Combo2 & "' and a0e19 >= '" & Text1 & "' and a0e19 <= '" & Text2 & "' and a0e14 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e14 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e21 = 0 and a0e25 = 0 and a0e04 = '" & MsgText(18) & "' order by a0e01 asc, a0e02 asc", adoTaie, adOpenStatic, adLockReadOnly
   '20140120END
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
Dim strUnion As String

On Error GoTo Checking

   strSql = ""
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   
   Call SetCompN 'Add by Sindy 2020/04/17
   
'   If Text3 = MsgText(601) Then
      
      '20140120START Modify By eric
      'Modify By Sindy 2020/4/17 公司別改變數
'      If Text5 <> "" Then
'         strSql = " and a0e23 = '" & IIf(Text5 = "2", "J", "1") & "' "
'      Else
'         strSql = ""
'      End If
      If strCmp <> MsgText(601) Then
          If InStr(strCmp, "+") > 0 Then
             strSql = " And (a0e23 is null or a0e23 In ('" & Replace(strCmp, "+", "','") & "')) "
          Else
             strSql = " And (a0e23 is null or a0e23='" & strCmp & "') "
          End If
      Else
         strSql = ""
      End If
      '2020/4/17 END
      
      If Text3 <> MsgText(601) Then
         strSql = strSql & " and a0e02 = '" & Text3 & "'"
      End If
      'If Text3 <> MsgText(601) Then
      '   strSql = " and a0e02 = '" & Text3 & "'"
      'End If
      '20140120END
       
      If Combo1 <> MsgText(601) Then
         strSql = strSql & " and a0e20 >= '" & Combo1 & "'"
      End If
      If Combo2 <> MsgText(601) Then
         strSql = strSql & " and a0e20 <= '" & Combo2 & "'"
      End If
      If Text1 <> MsgText(601) Then
         strSql = strSql & " and a0e19 >= '" & Text1 & "'"
      End If
      If Text2 <> MsgText(601) Then
         strSql = strSql & " and a0e19 <= '" & Text2 & "'"
      End If
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         strSql = strSql & " and a0e14 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      End If
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         strSql = strSql & " and a0e14 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      End If
      Select Case Text7
         Case MsgText(602)
            strSql = strSql & " and (a0e21 is not null and a0e21 <> 0)"
         Case MsgText(603)
            strSql = strSql & " and (a0e21 is null or a0e21 = 0)"
      End Select
      strUnion = "select a0e01, a0e02, a0g02, a0e20, a0e14, a0e10, a0e11, a0e05, cu04 as contect, a0e21 from acc0e0, acc0g0, customer where a0e19 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e25 = 0 and (a0e14 <> 0 and a0e14 is not null) and a0e04 = '" & MsgText(18) & "'" & strSql
      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e20, a0e14, a0e10, a0e11, a0e05, a0i02 as contect, a0e21 from acc0e0, acc0g0, acc0i0 where a0e19 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e25 = 0 and (a0e14 <> 0 and a0e14 is not null) and a0e04 = '" & MsgText(18) & "'" & strSql
      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e20, a0e14, a0e10, a0e11, a0e05, st02 as contect, a0e21 from acc0e0, acc0g0, staff where a0e19 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e25 = 0 and (a0e14 <> 0 and a0e14 is not null) and a0e04 = '" & MsgText(18) & "'" & strSql
      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e20, a0e14, a0e10, a0e11, a0e05, '' as contect, a0e21 from acc0e0, acc0g0 where a0e19 = a0g01 and a0e05 = '4' and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e25 = 0 and (a0e14 <> 0 and a0e14 is not null) and a0e04 = '" & MsgText(18) & "'" & strSql & " order by a0e01 asc, a0e02 asc"
      adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      strUnion = "select a0e01, a0e02, a0g02, a0e20, a0e14, a0e10, a0e11, a0e05, cu04 as contect from acc0e0, acc0g0, customer where a0e19 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and a0e02 = '" & Text3 & "' and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e25 = 0 and a0e14 <> 0 and a0e04 = '" & MsgText(18) & "'"
'      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e20, a0e14, a0e10, a0e11, a0e05, a0i02 as contect from acc0e0, acc0g0, acc0i0 where a0e19 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and a0e02 = '" & Text3 & "' and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e25 = 0 and a0e14 <> 0 and a0e04 = '" & MsgText(18) & "'"
'      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e20, a0e14, a0e10, a0e11, a0e05, st02 as contect from acc0e0, acc0g0, staff where a0e19 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and a0e02 = '" & Text3 & "' and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e25 = 0 and a0e14 <> 0 and a0e04 = '" & MsgText(18) & "'"
'      strUnion = strUnion & " union select a0e01, a0e02, a0g02, a0e20, a0e14, a0e10, a0e11, a0e05, '' as contect from acc0e0, acc0g0 where a0e19 = a0g01 and a0e05 = '4' and a0e02 = '" & Text3 & "' and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e25 = 0 and a0e14 <> 0 and a0e04 = '" & MsgText(18) & "' order by a0e01 asc, a0e02 asc"
'      adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
'   End If
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            AdodcRefresh
            SumShow
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   If Text3 = MsgText(601) Then
      adoaccsum.Open "select SUM(A0E11) from ACC0E0 WHERE a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e25 = 0 and a0e14 <> 0 and a0e04 = '" & MsgText(18) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoaccsum.Open "select SUM(A0E11) from ACC0E0 WHERE a0e02 = '" & Text3 & "' and a0e15 = 0 and a0e16 = 0 and a0e17 = 0 and a0e25 = 0 and a0e14 <> 0 and a0e04 = '" & MsgText(18) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
   End If
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text4 = MsgText(601)
      Else
         Text4 = Format(adoaccsum.Fields(0).Value, DDollar)
      End If
   Else
      Text4 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Combo1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Combo2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modify by Sindy 2020/04/17 公司別改下拉
''20140120START By eric
'Private Sub Text5_LostFocus()
'   If Text5.Text <> "1" And Text5.Text <> "2" And Text5.Text <> "" Then
'      MsgBox "公司別僅可為 1 / 2 或不輸入  ! "
'      Text5.Text = ""
'      Text5.SetFocus
'      Exit Sub
'   End If
'End Sub
'
''20140120START By eric
'Private Sub Text5_GotFocus()
'   TextInverse Text5
'   CloseIme
'End Sub
'
''20140120START By eric
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
