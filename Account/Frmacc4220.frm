VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc4220 
   AutoRedraw      =   -1  'True
   Caption         =   "科目分類帳查詢"
   ClientHeight    =   5016
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9408
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5016
   ScaleWidth      =   9408
   Begin VB.ComboBox Combo1 
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
      Left            =   1308
      TabIndex        =   0
      Top             =   210
      Width           =   3500
   End
   Begin VB.TextBox Text10 
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
      Height          =   300
      Left            =   7800
      TabIndex        =   18
      Top             =   4560
      Width           =   1500
   End
   Begin VB.TextBox Text9 
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
      Height          =   324
      Left            =   1308
      TabIndex        =   16
      Top             =   912
      Width           =   1800
   End
   Begin VB.TextBox Text3 
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
      Height          =   300
      Left            =   5232
      TabIndex        =   14
      Top             =   4560
      Width           =   1500
   End
   Begin VB.TextBox Text1 
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
      Height          =   300
      Left            =   3720
      TabIndex        =   13
      Top             =   4560
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4220.frx":0000
      Height          =   3108
      Left            =   240
      TabIndex        =   12
      Top             =   1344
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   5482
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
      Caption         =   "科目分類帳資料"
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "a0205"
         Caption         =   "傳票日期"
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
      BeginProperty Column01 
         DataField       =   "ax202"
         Caption         =   "傳票號碼"
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
         DataField       =   "a0902"
         Caption         =   "部門別"
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
         DataField       =   "ax206"
         Caption         =   "借方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "ax207"
         Caption         =   "貸方金額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "ax212"
         Caption         =   "摘要"
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
      BeginProperty Column06 
         DataField       =   "ax208"
         Caption         =   "對沖代號(客)"
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
      BeginProperty Column07 
         DataField       =   "ax209"
         Caption         =   "對沖代號(業)"
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
         DataField       =   "ax214"
         Caption         =   "對沖代號(本所案號)"
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
      BeginProperty Column09 
         DataField       =   "ax213"
         Caption         =   "對沖代號(其他)"
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
      BeginProperty Column10 
         DataField       =   "a0201"
         Caption         =   "公司別"
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
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1391.811
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   708.095
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1488.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   4356.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1391.811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1391.811
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1620.284
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   780.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1260
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   6225
      TabIndex        =   10
      Top             =   570
      Width           =   1215
      _ExtentX        =   2138
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
   Begin VB.TextBox Text8 
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
      Height          =   300
      Left            =   6840
      TabIndex        =   9
      Top             =   210
      Width           =   1572
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
      Height          =   300
      Left            =   6240
      TabIndex        =   1
      Top             =   210
      Width           =   612
   End
   Begin VB.TextBox Text7 
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
      Left            =   2745
      TabIndex        =   7
      Top             =   576
      Width           =   2050
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1308
      TabIndex        =   2
      Top             =   576
      Width           =   1400
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   7785
      TabIndex        =   11
      Top             =   570
      Width           =   1215
      _ExtentX        =   2138
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
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "餘額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   19
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "上月餘額"
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
      TabIndex        =   17
      Top             =   936
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
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "部門別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   210
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1170
      Left            =   270
      Top             =   120
      Width           =   9000
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   7545
      TabIndex        =   6
      Top             =   570
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
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
      Left            =   5280
      TabIndex        =   5
      Top             =   576
      Width           =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "會計科目"
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
      TabIndex        =   4
      Top             =   576
      Width           =   972
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   210
      Width           =   735
   End
End
Attribute VB_Name = "Frmacc4220"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/13 Form2.0已修改 DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoadodc1 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strSql As String
Dim strSQL1 As String

'Add by Amy 2020/04/09
Private Sub Combo1_GotFocus()
    TextInverse Combo1
    CloseIme
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(Combo1) = MsgText(601) Then Exit Sub
    
    strCmp = Combo1
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label2 & MsgText(63), , MsgText(5)
        Cancel = True
        Combo1.SetFocus
        Exit Sub
    ElseIf Len(Trim(Combo1)) = 1 Then
        Combo1 = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/09

Private Sub DataGrid1_DblClick()
   If DataGrid1.row >= 0 Then
      If DataGrid1.Columns(1).Text <> "" Then
         Load Frmacc4221
         Frmacc4221.p_stA0202 = DataGrid1.Columns(1).Text
         Frmacc4221.p_stA0201 = DataGrid1.Columns(10).Text 'Added by Morgan 2014/2/24
         Frmacc4221.QueryTable
         Set Frmacc4221.p_oForm = Me
         Me.Enabled = False
      End If
   End If
End Sub

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
   Me.Width = 9700 'Modify by Amy 2023/07/19 原:9500
   Me.Height = 5600 'Modify by Amy 2023/0719 原:5400
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   '92.10.22 modify by sonia 原為年月範圍改為日期範圍
   'MaskEdBox1.Mask = Mid(DFormat, 1, 6)
   'MaskEdBox2.Mask = Mid(DFormat, 1, 6)
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   OpenTable
   '20140122REMARK By eric (公司別欄位可選擇 1本所 或 2智權)
   'Text4 = "1"
   'Add by Amy 2020/04/09
   Combo1.AddItem "", 0
   Call Pub_SetCboCmp(Combo1, False, False, False, , 1)
   'end 2020/04/09
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc4220 = Nothing
End Sub

Private Sub Text2_Change()
   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   Text8 = A0902Query(Text2)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Mark by Amy 2020/04/09 改下拉
'Private Sub Text4_Change()
'   If Text4 = MsgText(601) Then
'      Exit Sub
'   End If
'   '20140122START Add By eric
'   If Text4 <> "1" And Text4 <> "J" Then
'      MsgBox "公司別僅能為 1 或 J ! (1:台一/J:智權)"
'      Text4.Text = ""
'      Text4.SetFocus
'   End If
'   '20140122END
'
'   Text5 = A0802Query(Text4)
'End Sub
'
'Private Sub Text4_GotFocus()
'   TextInverse Text4
'   '20140122START Add By eric
'   CloseIme
'   '20140122END
'End Sub

''20140122START Add By eric
'Private Sub Text4_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
''20140120START By eric
'Private Sub Text4_LostFocus()
'   If Text4.Text = "" Then
'      MsgBox "公司別僅可為 1 / 2 !"
'      Text4.Text = ""
'      Text4.SetFocus
'      Exit Sub
'   End If
'End Sub
'end 2020/04/09

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
    Dim strCmp As String 'Add by Amy 2020/04/09
On Error GoTo Checking

   adoadodc1.CursorLocation = adUseClient
   'Modify by Amy 2020/04/09 公司別改下拉 原:Text4
   If Trim(Combo1) <> MsgText(601) Then
      strCmp = Combo1
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
   End If
   Select Case strAccount
      Case "2"
         adoadodc1.Open "select a0305 as a0205, ax302 as ax202, ax304 as ax204, ax308 as ax208, ax309 as ax209, ax306 as ax206, ax307 as ax207, ax312 as ax212 from acc031, acc030 where acc031.ax301 = acc030.a0301 and acc031.ax302 = acc030.a0302 and ax301 = '" & strCmp & "' and ax304 = '" & Text2 & "' and ax305 = '" & Text6 & "' order by a0301 asc, a0305 asc, a0302 asc", adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         adoadodc1.Open "select a0205, ax202, ax204, ax208, ax209, ax206, ax207, ax212 from acc021, acc020 where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and ax201 = '" & strCmp & "' and ax204 = '" & Text2 & "' and ax205 = '" & Text6 & "' order by a0201 asc, a0205 asc, a0202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
   End Select
   'end 2020/04/09
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text6_Change()
   If Text6 = MsgText(601) Then
      Exit Sub
   End If
   Text7 = A0102Query(Text6)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

'*************************************************
'  查詢資料表(傳票資料)
'
'*************************************************
Public Sub QueryTable()
Dim lngStartDate As Long
Dim lngEndDate As Long
Dim strCmp As String 'Add by Amy 2020/04/09

On Error GoTo Checking
   strSql = ""
   '92.10.22 modify by sonia
   'lngStartDate = Val(FCDate(MaskEdBox1.Text & "/" & MsgText(12)))
   'lngEndDate = Val(FCDate(MaskEdBox2.Text & "/" & MsgText(13)))
   lngStartDate = Val(FCDate(MaskEdBox1.Text))
   lngEndDate = Val(FCDate(MaskEdBox2.Text))
   '92.10.22 end
   'Modify by Amy 2020/04/09 公司別改下拉 原:Text4
   If Trim(Combo1) <> MsgText(601) Then
        strCmp = Combo1
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
   End If
   'end 2020/04/09
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         'Modify by Amy 2020/04/09 公司別改下拉 原:Text4
         If strCmp <> MsgText(601) Then
            strSql = " and ax301 = '" & strCmp & "'"
         End If
         If Text2 <> MsgText(601) Then
            strSql = strSql & " and ax304 = '" & Text2 & "'"
'         Else
'            strSQL = strSQL & " and ax304 = '" & MsgText(55) & "'"
         End If
         If Text6 <> MsgText(601) Then
            strSql = strSql & " and ax305 = '" & Text6 & "'"
         End If
         '92.10.22 modify by sonia
         'If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         '92.10.22 end
            strSql = strSql & " and a0305 >= " & lngStartDate & ""
         End If
         '92.10.22 modify by sonia
         'If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         '92.10.22 end
            strSql = strSql & " and a0305 <= " & lngEndDate & ""
         End If
         
         'Added by Morgan 2025/5/26
         If pub_strUserOffice <> "1" Then
            If pub_strUserOffice = "2" Then
               strSql = strSql & " and ax305 like '1911%'"
            ElseIf pub_strUserOffice = "3" Then
               strSql = strSql & " and ax305 like '1912%'"
            ElseIf pub_strUserOffice = "4" Then
               strSql = strSql & " and ax305 like '1913%'"
            End If
         End If
         'end 2025/5/26
         
         If strSql <> MsgText(601) Then
            strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
         End If
         'Modified by Morgan 2014/2/24 +a0201
         adoadodc1.Open "select /*+FIRST_ROWS */ a0305 as a0205, ax302 as ax202, ax304 as ax204, ax308 as ax208, ax309 as ax209, ax306 as ax206, ax307 as ax207, ax312 as ax212, a0902, ax314 as ax214, ax313 as ax213, a0301 as a0201 from (select * from acc031, acc030, acc090 where ax301 = a0301 and ax302 = a0302 and ax304 = a0901) new " & strSql & " order by a0301 asc, a0305 asc, a0302 asc", adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         'Modify by Amy 2020/04/09 公司別改下拉 原:Text4
         If strCmp <> MsgText(601) Then
            'Modify by Morgan 2004/11/16
            'strSQL = " and ax201 = '" & Text4 & "'"
            strSql = " and a0201 = '" & strCmp & "'"
         End If
         If Text2 <> MsgText(601) Then
            strSql = strSql & " and ax204 = '" & Text2 & "'"
'         Else
'            strSQL = strSQL & " and ax204 = '" & MsgText(55) & "'"
         End If
         If Text6 <> MsgText(601) Then
            strSql = strSql & " and ax205 = '" & Text6 & "'"
         End If
         
         'Added by Morgan 2025/5/26
         If pub_strUserOffice <> "1" Then
            If pub_strUserOffice = "2" Then
               strSql = strSql & " and ax205 like '1911%'"
            ElseIf pub_strUserOffice = "3" Then
               strSql = strSql & " and ax205 like '1912%'"
            ElseIf pub_strUserOffice = "4" Then
               strSql = strSql & " and ax205 like '1913%'"
            End If
         End If
         'end 2025/5/26
         
         '92.10.22 modify by sonia
         'If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> Mid(MsgText(29), 1, 6) Then
         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         '92.10.22 end
            strSql = strSql & " and a0205 >= " & lngStartDate & ""
         End If
         '92.10.22 modify by sonia
         'If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> Mid(MsgText(29), 1, 6) Then
         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         '92.10.22 end
            strSql = strSql & " and a0205 <= " & lngEndDate & ""
         End If
         'If strSQL <> MsgText(601) Then
         '   strSQL = " where " & Mid(strSQL, 5, Len(strSQL) - 4)
         'End If
         'Modify by  Morgan 2004/11/16 太慢，改語法
         'adoadodc1.Open "select /*+FIRST_ROWS */ a0205, ax202, ax204, ax208, ax209, ax206, ax207, ax212, ax204 as a0902, ax214, ax213 from acc020, acc021 where ax201 = a0201 and ax202 = a0202" & strSQL & " order by a0201 asc, a0205 asc, a0202 asc", adoTaie, adOpenStatic, adLockReadOnly
         '2013/9/18 MODIF BY SONIA 加入排序, ax203 asc
         'Modified by Morgan 2014/2/24 +a0201
         strSql = "select a0205, ax202, ax204, ax208, ax209, ax206, ax207, ax212, ax204 as a0902, ax214, ax213, a0201 " & _
            " From acc020, acc021" & _
            " where ax201(+) = a0201 and ax202(+) = a0202" & strSql & _
            " order by a0201 asc, a0205 asc, a0202 asc, ax203 asc"
         adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   End Select
   strSQL1 = ""
   'Modify by Amy 2020/04/09 改下拉
   If strCmp <> MsgText(601) Then
      strSQL1 = strSQL1 & " and a0403 = '" & strCmp & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSQL1 = strSQL1 & " and a0404 = '" & Text2 & "'"
   Else
      strSQL1 = strSQL1 & " and a0404 = '" & MsgText(55) & "'"
   End If
   If Text6 <> MsgText(601) Then
      strSQL1 = strSQL1 & " and a0405 = '" & Text6 & "'"
   End If
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoaccsum.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         adoaccsum.Open "select decode(substr(a0505, 1, 1), '1', a0508, '2', a0508, '3', a0508, 0) from acc050 where a0501 = " & IIf((Val(Mid(MaskEdBox1.Text, 5, 2)) - 1) < 1, Val(Mid(MaskEdBox1.Text, 1, 3)) - 1, Val(Mid(MaskEdBox1.Text, 1, 3))) & " and a0502 = " & IIf((Val(Mid(MaskEdBox1.Text, 5, 2)) - 1) < 1, 12, Val(Mid(MaskEdBox1.Text, 5, 2)) - 1) & " and substr(a0505, 1, 1) in ('1', '2', '3')" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
      Case Else
         adoaccsum.Open "select decode(substr(a0405, 1, 1), '1', a0408, '2', a0408, '3', a0408, 0) from acc040 where a0401 = " & IIf((Val(Mid(MaskEdBox1.Text, 5, 2)) - 1) < 1, Val(Mid(MaskEdBox1.Text, 1, 3)) - 1, Val(Mid(MaskEdBox1.Text, 1, 3))) & " and a0402 = " & IIf((Val(Mid(MaskEdBox1.Text, 5, 2)) - 1) < 1, 12, Val(Mid(MaskEdBox1.Text, 5, 2)) - 1) & " and substr(a0405, 1, 1) in ('1', '2', '3')" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
   End Select
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text9 = MsgText(601)
      Else
         Text9 = Format(adoaccsum.Fields(0).Value, FDollar)
      End If
   Else
      Text9 = MsgText(601)
   End If
   adoaccsum.Close
   Adodc1.Recordset.ReQuery
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

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
Dim bolShowMsg As Boolean 'Add by Amy 2020/04/09

   Select Case KeyCode
      Case vbKeyF12
         'Modify by Amy 2020/04/09 +bolShowMsg
         If FormCheck(bolShowMsg) Then
            Screen.MousePointer = vbHourglass
            QueryTable
            SumShow
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            If bolShowMsg = False Then MsgBox MsgText(181), , MsgText(5)
         End If
         'end 2020/04/09
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
Dim douDebit As Double
Dim douCredit As Double
Dim douAmt As Double  '2005/11/18 ADD BY SONIA

   douDebit = 0
   douCredit = 0
   If adoadodc1.State = adStateOpen Then
      Set adoaccsum = adoadodc1.Clone
      Do While adoaccsum.EOF = False
         If IsNull(adoaccsum.Fields("ax206").Value) = False Then
            douDebit = douDebit + Val(adoaccsum.Fields("ax206").Value)
         End If
         If IsNull(adoadodc1.Fields("ax207").Value) = False Then
            douCredit = douCredit + Val(adoaccsum.Fields("ax207").Value)
         End If
         adoaccsum.MoveNext
      Loop
      If adoaccsum.RecordCount <> 0 Then
         adoaccsum.MoveFirst
      End If
      adoaccsum.Close
   End If
   Text1 = Format(douDebit, FDollar)
   Text3 = Format(douCredit, FDollar)
   '2005/11/18 ADD BY SONIA
   If adoaccsum.State = adStateOpen Then
      adoaccsum.Close
   End If
   adoaccsum.CursorLocation = adUseClient
   adoaccsum.Open "select A0103 from acc010 where a0101 = '" & Text6 & "' ", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccsum.RecordCount <> 0 Then
      '2012/9/18 modify by sonia 改同帳務,資產1XXX,費用6XXX及9XXX(分攤費用)科目為合計餘額=上期餘額+合計借方-合計貸方,其餘科目為合計餘額=上期餘額-合計借方+合計貸方
      'If adoaccsum.Fields("a0103").Value = "1" Then
      '   douAmt = Val(Format(Text9, FAmount)) + douDebit - douCredit
      'Else
      '   douAmt = Val(Format(Text9, FAmount)) + douCredit - douDebit
      'End If
      If Left(Text6, 1) = "1" Or Left(Text6, 1) = "6" Or Left(Text6, 1) = "9" Then
         douAmt = Val(Format(Text9, FAmount)) + douDebit - douCredit
      Else
         douAmt = Val(Format(Text9, FAmount)) + douCredit - douDebit
      End If
      '2012/9/18 END
   End If
   adoaccsum.Close
   Text10 = Format(douAmt, FDollar)
   '2005/11/18 END
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck(ByRef bolMsg As Boolean) As Boolean
    Dim bolCancel As Boolean 'Add by Amy 2020/04/09
    
   'Modify by Amy 2020/04/09 改下拉 原:Text4
   If Trim(Combo1) <> MsgText(601) Then
      Call Combo1_Validate(bolCancel)
      If bolCancel = False Then
        FormCheck = True
        Exit Function
      Else
        bolMsg = True
        FormCheck = False
        Exit Function
      End If
   'add by sonia 2021/11/17 公司別不可空白
   Else
      MsgBox "公司別不可空白，請選擇！"
      Combo1.SetFocus
      bolMsg = True
      FormCheck = False
      Exit Function
   'end 2021/11/17
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text6 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function



