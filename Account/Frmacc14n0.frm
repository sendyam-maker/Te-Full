VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc14n0 
   AutoRedraw      =   -1  'True
   Caption         =   "已開發票未收款明細查詢"
   ClientHeight    =   5400
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9048
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   9048
   Begin VB.CommandButton Command2 
      Caption         =   "E-mail 通知智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   690
      Width           =   2355
   End
   Begin VB.TextBox Text4 
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
      Left            =   3840
      MaxLength       =   6
      TabIndex        =   3
      Top             =   450
      Width           =   915
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
      Left            =   2010
      MaxLength       =   3
      TabIndex        =   2
      Top             =   480
      Width           =   735
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
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   1
      Top             =   480
      Width           =   735
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
      Left            =   1140
      MaxLength       =   1
      TabIndex        =   0
      Top             =   90
      Width           =   465
   End
   Begin VB.CommandButton Command1 
      Caption         =   "產生Excel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   210
      Width           =   1605
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   1140
      TabIndex        =   4
      Top             =   870
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   572
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   330
      Left            =   2550
      TabIndex        =   5
      Top             =   870
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   572
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7650
      Top             =   4920
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
      Bindings        =   "Frmacc14n0.frx":0000
      Height          =   3915
      Left            =   60
      TabIndex        =   14
      Top             =   1260
      Width           =   8805
      _ExtentX        =   15536
      _ExtentY        =   6900
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   16
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "st06"
         Caption         =   "所"
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
         DataField       =   "a0902"
         Caption         =   "業務區"
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
         DataField       =   "st02"
         Caption         =   "智權人員"
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
         DataField       =   "a0k01"
         Caption         =   "請款單號"
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
      BeginProperty Column04 
         DataField       =   "a4302"
         Caption         =   "發票日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "yyyy/M/d"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "a4301"
         Caption         =   "發票號碼"
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
         DataField       =   "a0j02"
         Caption         =   "本所案號"
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
         DataField       =   "a0k04"
         Caption         =   "收據抬頭"
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
         DataField       =   "amt"
         Caption         =   "發票總額"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "a4305"
         Caption         =   "營業稅"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "a4307"
         Caption         =   "列印日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "ee/mm/dd"
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
            Locked          =   -1  'True
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   671.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   815.811
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   912.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   912.189
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   1031.811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column07 
            Locked          =   -1  'True
            ColumnWidth     =   1068.094
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   684.284
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   912.189
         EndProperty
      EndProperty
   End
   Begin MSForms.Label Label5 
      Height          =   300
      Left            =   4800
      TabIndex        =   13
      Top             =   480
      Width           =   1485
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2619;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      Left            =   2850
      TabIndex        =   12
      Top             =   510
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   1860
      X2              =   2070
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "業  務 區"
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
      Left            =   180
      TabIndex        =   11
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "發票日期"
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
      Left            =   180
      TabIndex        =   10
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "所　　別"
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
      Left            =   180
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "（1.北 2.中 3.南 4.高 5.其他 空白.全部）"
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
      Left            =   1830
      TabIndex        =   8
      Top             =   120
      Width           =   4065
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2610
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   6090
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc14n0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改
'Created by Sindy 2014/1/23
Option Explicit

Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt424 As New Worksheet
Dim m_lngRow As Long
Dim m_intPage As Integer
Dim strCaption As String


'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            doQuery
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
Dim Cancel As Boolean
   
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
      MsgBox Label1 & MsgText(52), , MsgText(5)
      MaskEdBox1.SetFocus
      FormCheck = False
      Exit Function
   End If
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label1 & MsgText(63), , MsgText(5)
      MaskEdBox1.SetFocus
      FormCheck = False
      Exit Function
   End If
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then
      MsgBox Label1 & MsgText(52), , MsgText(5)
      MaskEdBox2.SetFocus
      FormCheck = False
      Exit Function
   End If
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label1 & MsgText(63), , MsgText(5)
      MaskEdBox2.SetFocus
      FormCheck = False
      Exit Function
   End If
   Cancel = False
   Call Text4_Validate(Cancel)
   If Cancel = True Then
      FormCheck = False
      Exit Function
   End If
   
   FormCheck = True
End Function

'產生Excel
Private Sub Command1_Click()
Dim ii As Integer
Dim strTemp(1 To 11) As String
   
   strCaption = "已開發票未收款明細查詢"
   If Dir(strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43)
   End If
   
   Screen.MousePointer = vbHourglass
   m_intPage = 0: m_lngRow = 0
   xlsSalesPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   Set wksaccrpt424 = xlsSalesPoint.Worksheets(1)
   'xlsSalesPoint.Visible = True
   wksaccrpt424.PageSetup.Orientation = xlLandscape '橫印
   'wksaccrpt424.PageSetup.Orientation = wdOrientLandscape '直印
   wksaccrpt424.PageSetup.LeftMargin = 28.34
   wksaccrpt424.PageSetup.RightMargin = 28.34
   wksaccrpt424.PageSetup.TopMargin = 42.51
   wksaccrpt424.PageSetup.BottomMargin = 42.51
   wksaccrpt424.PageSetup.HeaderMargin = 28.34
   wksaccrpt424.PageSetup.FooterMargin = 28.34
   wksaccrpt424.Columns("a:a").ColumnWidth = 3
   wksaccrpt424.Columns("b:b").ColumnWidth = 10
   wksaccrpt424.Columns("c:c").ColumnWidth = 10
   wksaccrpt424.Columns("d:d").ColumnWidth = 10
   wksaccrpt424.Columns("e:e").ColumnWidth = 10
   wksaccrpt424.Columns("f:f").ColumnWidth = 12
   wksaccrpt424.Columns("g:g").ColumnWidth = 12 'Add By Sindy 2015/3/18
   wksaccrpt424.Columns("h:h").ColumnWidth = 35
   wksaccrpt424.Columns("i:i").ColumnWidth = 10
   'wksaccrpt424.Range("g:j").Select
   'xlsSalesPoint.Selection.NumberFormatLocal = "@"
   wksaccrpt424.Columns("j:j").ColumnWidth = 10
   wksaccrpt424.Columns("k:k").ColumnWidth = 10
   Call ExcelHead '頁首
   
   With Adodc1.Recordset
      .MoveFirst
      Do While Not .EOF
         If m_lngRow Mod 32 = 0 Then
            '換頁
            wksaccrpt424.Range("A" & (m_lngRow + 1)).Select
            wksaccrpt424.HPageBreaks.add Before:=wksaccrpt424.Application.ActiveCell
            Call ExcelHead '頁首
         End If
         '清空變數值
         For ii = 1 To 11
            strTemp(ii) = ""
         Next ii
         '讀取欄位值
         For ii = 0 To 10
            Select Case ii
               Case 8, 9 '金額
                  strTemp(ii + 1) = Format("" & .Fields(ii), "#,##0")
               Case Else
                  strTemp(ii + 1) = "" & .Fields(ii)
            End Select
         Next ii
         '存放欄位
         m_lngRow = m_lngRow + 1
         wksaccrpt424.Range("a" & m_lngRow).Value = strTemp(1)
         wksaccrpt424.Range("b" & m_lngRow).Value = strTemp(2)
         wksaccrpt424.Range("c" & m_lngRow).Value = strTemp(3)
         wksaccrpt424.Range("d" & m_lngRow).Value = strTemp(4)
         wksaccrpt424.Range("e" & m_lngRow).Value = strTemp(5)
         wksaccrpt424.Range("f" & m_lngRow).Value = strTemp(6)
         wksaccrpt424.Range("g" & m_lngRow).Value = strTemp(7) 'Add By Sindy 2015/3/18
         wksaccrpt424.Range("h" & m_lngRow).Value = strTemp(8)
         wksaccrpt424.Range("i" & m_lngRow).Value = strTemp(9)
         wksaccrpt424.Range("j" & m_lngRow).Value = strTemp(10)
         wksaccrpt424.Range("k" & m_lngRow).Value = strTemp(11)
         .MoveNext
      Loop
   End With
   'Modify by Amy 2016/06/23 +判斷版本
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=56
   End If
   'end 2016/06/23
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set xlsSalesPoint = Nothing
   Set wksaccrpt424 = Nothing
   StatusClear
   Screen.MousePointer = vbDefault
   MsgBox "Excel檔產生完畢！（" & strExcelPath & strCaption & ACDate(ServerDate) & ServerTime & MsgText(43) & "）"
End Sub

Private Sub ExcelHead()
   m_intPage = m_intPage + 1
   m_lngRow = m_lngRow + 1
   wksaccrpt424.Range("a" & m_lngRow).Value = strCaption
   wksaccrpt424.Range("a" & m_lngRow & ":k" & m_lngRow).Select
   With xlsSalesPoint.Selection
       .HorizontalAlignment = xlCenter
       .VerticalAlignment = xlBottom
       .WrapText = False
       .Orientation = 0
       .AddIndent = False
       .ShrinkToFit = False
       .MergeCells = True
   End With
   wksaccrpt424.Application.Selection.Font.Bold = True
   wksaccrpt424.Application.Selection.Font.Size = 16
   
   m_lngRow = m_lngRow + 1
   wksaccrpt424.Range("a" & m_lngRow).Value = "列印人：" & GetStaffName(strUserNum)
   wksaccrpt424.Range("f" & m_lngRow).Value = "發票日期：" & MaskEdBox1.Text & "~" & MaskEdBox2.Text
   wksaccrpt424.Range("i" & m_lngRow).Value = "列印日期：" & CFDate(strSrvDate(2))
   m_lngRow = m_lngRow + 1
   wksaccrpt424.Range("i" & m_lngRow).Value = "頁　　次：" & m_intPage
   m_lngRow = m_lngRow + 1
   wksaccrpt424.Range("a" & m_lngRow).Value = "所"
   wksaccrpt424.Range("b" & m_lngRow).Value = "業務區"
   wksaccrpt424.Range("c" & m_lngRow).Value = "智權人員"
   wksaccrpt424.Range("d" & m_lngRow).Value = "請款單號"
   wksaccrpt424.Range("e" & m_lngRow).Value = "發票日期"
   wksaccrpt424.Range("f" & m_lngRow).Value = "發票號碼"
   wksaccrpt424.Range("g" & m_lngRow).Value = "本所案號" 'Add By Sindy 2015/3/18
   wksaccrpt424.Range("h" & m_lngRow).Value = "收據抬頭"
   wksaccrpt424.Range("i" & m_lngRow).Value = "發票總額"
   wksaccrpt424.Range("j" & m_lngRow).Value = "營業稅"
   wksaccrpt424.Range("k" & m_lngRow).Value = "列印日期"
   wksaccrpt424.Range("A" & m_lngRow & ":K" & m_lngRow).Select
   wksaccrpt424.Application.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlEdgeLeft).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlEdgeTop).LineStyle = xlNone
   With wksaccrpt424.Application.Selection.Borders(xlEdgeBottom)
      .LineStyle = xlContinuous
      .Weight = xlThin
      .ColorIndex = xlAutomatic
   End With
   wksaccrpt424.Application.Selection.Borders(xlEdgeRight).LineStyle = xlNone
   wksaccrpt424.Application.Selection.Borders(xlInsideVertical).LineStyle = xlNone
End Sub

'E-mail 通知智權人員
Private Sub Command2_Click()
Dim TempFileName As String, strFileName As String
Dim strSales As String, strSales_Nm As String
Dim ff As Integer
Dim A01 As String, A02 As String, A03 As String, A04 As String
Dim A05 As String, A06 As String, A07 As String, A08 As String
   
   Screen.MousePointer = vbHourglass
   With Adodc1.Recordset
      If .RecordCount > 0 Then
         .MoveFirst
         TempFileName = FCDate(MaskEdBox1) & "-" & FCDate(MaskEdBox2) & "已開發票未收款明細"
         strSales = "": strSales_Nm = ""
         Do While Not .EOF
            If strSales <> .Fields("a0k20") Then
               If strSales <> "" Then
                  Close ff
                  PUB_SendMail strUserNum, strSales, "", "您尚有已開發票未收款！", "Dear Sirs," & vbCrLf & "          " & "請參考附件資料。" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心", , App.path & strFileName & ".txt"
               End If
               strSales = .Fields("a0k20")
               strSales_Nm = .Fields("st02")
               ff = FreeFile
               strFileName = "\" & strSales_Nm & TempFileName
               If ff > 0 Then Close #ff
               ff = FreeFile
               Open App.path & strFileName & ".txt" For Output As ff
               Print #ff, "發票日期  發票號碼   請款單號   收據抬頭             本所案號        申請國家       案件性質       金　額"
               Print #ff, "========= ========== ========== ==================== =============== ============== ========== =========="
            End If
            strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','',cp03||decode(cp04,'00','',cp04)),na03,decode(a0j04,'000',cpm03,cpm04),nvl(a0j09,0)+nvl(a0j10,0)" & _
                        " From acc0j0,caseprogress,casepropertymap,Nation" & _
                        " where a0j13='" & .Fields("a0k01") & "'" & _
                        " and a0j01=cp09(+)" & _
                        " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                        " and a0j04=na01(+)" & _
                        " order by cp01,cp02,cp03,cp04 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If RsTemp.RecordCount > 0 Then
               RsTemp.MoveFirst
               intI = 0
               Do While Not RsTemp.EOF
                  intI = intI + 1
                  If intI = 1 Then
                     A01 = convForm(.Fields("a4302"), 9)
                     A02 = convForm(.Fields("a4301"), 10)
                     A03 = convForm(.Fields("a0k01"), 10)
                     A04 = convForm(.Fields("a0k04"), 20)
                  Else
                     A01 = convForm("", 9)
                     A02 = convForm("", 10)
                     A03 = convForm("", 10)
                     A04 = convForm("", 20)
                  End If
                  A05 = convForm(RsTemp.Fields(0), 15)
                  A06 = convForm(RsTemp.Fields(1), 14)
                  A07 = convForm(RsTemp.Fields(2), 10)
                  A08 = Right("          " & Format(RsTemp.Fields(3), "#,##0"), 10)
                  Print #ff, A01 & " " & A02 & " " & A03 & " " & A04 & " " & A05 & " " & A06 & " " & A07 & " " & A08
                  RsTemp.MoveNext
               Loop
            End If
            .MoveNext
         Loop
         Close ff
         PUB_SendMail strUserNum, strSales, "", "您尚有已開發票未收款！", "Dear Sirs," & vbCrLf & "          " & "請參考附件資料。" & vbCrLf & vbCrLf & vbCrLf & "                                                        電腦中心", , App.path & strFileName & ".txt"
         MsgBox "E-mail已發文完畢!!" 'Add By Sindy 2014/2/18
      End If
   End With
   Screen.MousePointer = vbDefault
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
   'Modify by Amy 2023/08/18 W9045 H5700
   Me.Width = 9140
   Me.Height = 5840
   'end 2023/08/18
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Label5.Caption = ""
   Command1.Enabled = False
   Command2.Enabled = False
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_KillTempFile "*已開發票未收款明細*.txt" '清除暫存檔
   StatusClear
   strFormName = MsgText(601)
   MenuEnabled
   Set Frmacc14n0 = Nothing
End Sub

Private Sub doQuery()
   Dim strCon As String
   
   strCon = ""
   '所別
   If Text1 <> "" Then
      strCon = strCon & " and st06='" & Text1 & "'"
   End If
   '業務區
   If Text2 <> "" Then
      strCon = strCon & " and st15>='" & Text2 & "'"
   End If
   If Text3 <> "" Then
      strCon = strCon & " and st15<='" & Text3 & "'"
   End If
   '智權人員
   If Text4 <> "" Then
      strCon = strCon & " and a0k20='" & Text4 & "'"
   End If
   
   Screen.MousePointer = vbHourglass
   '未作廢無轉開已結清
   'Modify By Sindy 2015/3/18 +acc0j0抓取本所案號(a0j02)
   'Modify By Sindy 2015/5/6 +group by st06,a0902,st02,a0k01,a4302,a4301,a0j02,a0k04,a4304,a4305,a4307,a0k20,st01 : 過濾E10409939有二個收文號同本所案號不要重覆出現
   strExc(0) = "select decode(st06,'1','北','2','中','3','南','4','高','其他') st06,a0902,st02,a0k01,sqldatet(a4302) a4302,a4301,a0j02,a0k04,nvl(a4304,0)+nvl(a4305,0) amt,nvl(a4305,0) a4305,sqldatet(to_char(a4307,'yyyymmdd')) a4307,a0k20,st01" & _
               " From acc430, acc431, acc0k0, staff, acc090,acc0j0" & _
               " where a4301=axc01(+)" & _
               " and axc02=a0k01(+)" & _
               " and axc02=a0j13(+)" & _
               " and a0k20=st01(+) and st15=a0901(+)" & _
               " and a4302 between " & Val(FCDate(MaskEdBox1.Text)) & " and " & Val(FCDate(MaskEdBox2.Text)) & _
               " and nvl(a4308,0)=0 and nvl(a4310,0)=0 and a0k37 is null" & strCon & _
               " group by st06,a0902,st02,a0k01,a4302,a4301,a0j02,a0k04,a4304,a4305,a4307,a0k20,st01" & _
               " order by st06,st01,a0k01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2014/06/26 改不用離線資料集，避免資料多時新增至暫存檔慢
   'Set Adodc1.Recordset = PUB_CreateRecordset(RsTemp)
   Set Adodc1.Recordset = RsTemp
   If RsTemp.RecordCount > 0 Then
      Command1.Enabled = True
      Command2.Enabled = True
   Else
      Screen.MousePointer = vbDefault
      Command1.Enabled = False
      Command2.Enabled = False
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") _
      And KeyAscii <> Asc("3") And KeyAscii <> Asc("4") And KeyAscii <> Asc("5") Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
Dim strTemp As String, strTemp1 As String
   
   Label5.Caption = ""
   If Text4.Text <> "" Then
      If Not ClsPDGetStaff(Text4.Text, strTemp, strTemp1) Then
         Cancel = True
         Exit Sub
      End If
      Label5.Caption = strTemp
   End If
End Sub
