VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc4160 
   AutoRedraw      =   -1  'True
   Caption         =   "預算資料"
   ClientHeight    =   6036
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8808
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6036
   ScaleWidth      =   8808
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   5760
      Picture         =   "Frmacc4160.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   768
      Width           =   350
   End
   Begin VB.TextBox Text9 
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
      Left            =   6480
      TabIndex        =   21
      Top             =   408
      Width           =   1095
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
      Left            =   5880
      TabIndex        =   19
      Top             =   5580
      Width           =   2532
   End
   Begin VB.TextBox Text8 
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
      Left            =   1560
      TabIndex        =   17
      Top             =   5580
      Width           =   2532
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4160.frx":0102
      Height          =   4350
      Left            =   240
      TabIndex        =   14
      Top             =   1110
      Width           =   4095
      _ExtentX        =   7218
      _ExtentY        =   7684
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "去年預算資料"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "a0402"
         Caption         =   "月份"
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
         DataField       =   "a0409"
         Caption         =   "預算金額"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   135
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1031.811
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   2352.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   312
      Left            =   228
      Top             =   1044
      Visible         =   0   'False
      Width           =   960
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
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton Command2 
      Caption         =   "不複製前一年預算"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   228
      TabIndex        =   13
      Top             =   768
      Width           =   2052
   End
   Begin VB.CommandButton Command1 
      Caption         =   "複製前一年預算"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2400
      TabIndex        =   12
      Top             =   768
      Width           =   1932
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00C0FFFF&
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
      Left            =   5160
      MaxLength       =   3
      TabIndex        =   3
      Top             =   768
      Width           =   612
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   2
      Top             =   396
      Width           =   1092
   End
   Begin VB.TextBox Text4 
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
      Left            =   7080
      TabIndex        =   10
      Top             =   48
      Width           =   1452
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
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
      Left            =   6495
      MaxLength       =   3
      TabIndex        =   1
      Top             =   48
      Width           =   612
   End
   Begin VB.TextBox Text2 
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
      Left            =   1800
      TabIndex        =   8
      Top             =   48
      Width           =   3492
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   0
      Top             =   48
      Width           =   612
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Frmacc4160.frx":0117
      Height          =   4350
      Left            =   4560
      TabIndex        =   15
      Top             =   1110
      Width           =   4095
      _ExtentX        =   7218
      _ExtentY        =   7684
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "預算資料"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "a0402"
         Caption         =   "月份"
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
         DataField       =   "a0409"
         Caption         =   "預算金額"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   2352.189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   4560
      Top             =   1008
      Visible         =   0   'False
      Width           =   960
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
   Begin MSForms.TextBox Text6 
      Height          =   300
      Left            =   2310
      TabIndex        =   11
      Top             =   375
      Width           =   3000
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "5292;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "借/貸方"
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
      Left            =   5520
      TabIndex        =   20
      Top             =   408
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
      Left            =   4920
      TabIndex        =   18
      Top             =   5580
      Width           =   495
   End
   Begin VB.Label Label3 
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
      Left            =   480
      TabIndex        =   16
      Top             =   5580
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   252
      Left            =   5520
      TabIndex        =   9
      Top             =   48
      Width           =   732
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4728
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "年度"
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
      Left            =   4560
      TabIndex        =   7
      Top             =   768
      Width           =   492
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   5205
      Left            =   120
      Top             =   750
      Width           =   4335
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   6
      Top             =   48
      Width           =   732
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
      Left            =   240
      TabIndex        =   5
      Top             =   408
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc4160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/23 Form2.0已修改 Text6
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit
Public adoacc040T As New ADODB.Recordset
Public adoacc040 As New ADODB.Recordset
Public adoacc010 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public adoadodc2 As New ADODB.Recordset

Private Sub Command1_Click_Old()
'Dim intCounter As Integer
'
'   If Text1 = MsgText(601) Or Text5 = MsgText(601) Or Text7 = MsgText(601) Then
'        Exit Sub
'   'Added by Lydia 2019/11/28
'   Else
'        Acc040DeleteUpd
'   'end 2019/11/28
'   End If
'
'   adoadodc1.Close
'   adoadodc1.CursorLocation = adUseClient
'   If Text3 = MsgText(601) Then
'      adoadodc1.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   Else
'      adoadodc1.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   End If
'   Adodc1.Recordset.Requery
'   Acc040Check
'   If strControlButton = MsgText(602) Then
'      Exit Sub
'   End If
'   strControlButton = MsgText(601)
'   'Modify by Amy 2014/01/02
'   Text1.Locked = True
'   If adoacc040.State <> adStateClosed Then
'        adoacc040.Close
'   End If
'   'end 2014/01/02
'   adoacc040.CursorLocation = adUseClient
'   If Text3 = MsgText(601) Then
'      adoacc040.Open "select * from acc040 where a0401 = " & Val(Text7) - 1 & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc040.Open "select * from acc040 where a0401 = " & Val(Text7) - 1 & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   If adoacc040.RecordCount = 0 Then
'      If Adodc1.Recordset.RecordCount <> 0 Then
'         Exit Sub
'      End If
'      For intCounter = 1 To 12
'         If adoaccsum.State = adStateOpen Then
'            adoaccsum.Close
'         End If
'         adoaccsum.CursorLocation = adUseClient
'         If Text3 = MsgText(601) Then
'            adoaccsum.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0402 = " & intCounter & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
'         Else
'            adoaccsum.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0402 = " & intCounter & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
'         End If
'         If adoaccsum.RecordCount = 0 Then
'            Adodc1.Recordset.AddNew
'         End If
'         adoaccsum.Close
'         Acc040Save
'         Adodc1.Recordset.Fields("a0402").Value = intCounter
'         Adodc1.Recordset.Fields("a0406").Value = 0
'         Adodc1.Recordset.Fields("a0407").Value = 0
'         Adodc1.Recordset.Fields("a0408").Value = 0
'         Adodc1.Recordset.Fields("a0409").Value = 0
'         Adodc1.Recordset.UpdateBatch
'      Next intCounter
'   Else
'      Do While adoacc040.EOF = False
'         If adoaccsum.State = adStateOpen Then
'            adoaccsum.Close
'         End If
'         adoaccsum.CursorLocation = adUseClient
'         If Text3 = MsgText(601) Then
'            adoaccsum.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0402 = " & adoacc040.Fields("a0402").Value & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
'         Else
'            adoaccsum.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0402 = " & adoacc040.Fields("a0402").Value & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
'         End If
'         If adoaccsum.RecordCount = 0 Then
'            Adodc1.Recordset.AddNew
'         End If
'         adoaccsum.Close
'         Acc040Save
'         Adodc1.Recordset.Fields("a0402").Value = adoacc040.Fields("a0402").Value
'         Adodc1.Recordset.Fields("a0406").Value = 0
'         Adodc1.Recordset.Fields("a0407").Value = 0
'         Adodc1.Recordset.Fields("a0408").Value = 0
'         Adodc1.Recordset.Fields("a0409").Value = adoacc040.Fields("a0409").Value
'         Adodc1.Recordset.UpdateBatch
'         adoacc040.MoveNext
'      Loop
'   End If
'   adoacc040.Close
'   adoadodc2.Close
'   adoadodc2.CursorLocation = adUseClient
'   If Text3 = MsgText(601) Then
'      adoadodc2.Open "select * from acc040 where a0401 = " & (Val(Text7) - 1) & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   Else
'      adoadodc2.Open "select * from acc040 where a0401 = " & (Val(Text7) - 1) & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   End If
'   Adodc2.Recordset.Requery
'   SumShow
End Sub

'不複製前一年預算
Private Sub Command2_Click_Old()
''add by nickc 2007/02/08
'Dim intCounter As Integer
'
'   If Text1 = MsgText(601) Or Text5 = MsgText(601) Or Text7 = MsgText(601) Then
'        Exit Sub
'   'Added by Lydia 2019/11/28
'   Else
'        Acc040DeleteUpd
'   'end 2019/11/28
'   End If
'
'   adoadodc1.Close
'   adoadodc1.CursorLocation = adUseClient
'   If Text3 = MsgText(601) Then
'      adoadodc1.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   Else
'      adoadodc1.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   End If
'   Adodc1.Recordset.Requery
'   If Adodc1.Recordset.RecordCount <> 0 Then
'      Exit Sub
'   End If
'   Acc040Check
'   If strControlButton = MsgText(602) Then
'      Exit Sub
'   End If
'   strControlButton = MsgText(601)
'   Text1.Locked = True 'Modify by Amy 2014/01/02
'   For intCounter = 1 To 12
'      If adoaccsum.State = adStateOpen Then
'         adoaccsum.Close
'      End If
'      adoaccsum.CursorLocation = adUseClient
'      If Text3 = MsgText(601) Then
'         adoaccsum.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0402 = " & intCounter & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Else
'         adoaccsum.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0402 = " & intCounter & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenStatic, adLockReadOnly
'      End If
'      If adoaccsum.RecordCount = 0 Then
'         Adodc1.Recordset.AddNew
'      End If
'      adoaccsum.Close
'      Acc040Save
'      Adodc1.Recordset.Fields("a0402").Value = intCounter
'      Adodc1.Recordset.Fields("a0406").Value = 0
'      Adodc1.Recordset.Fields("a0407").Value = 0
'      Adodc1.Recordset.Fields("a0408").Value = 0
'      Adodc1.Recordset.Fields("a0409").Value = 0
'      Adodc1.Recordset.UpdateBatch
'   Next intCounter
'   adoadodc2.Close
'   adoadodc2.CursorLocation = adUseClient
'   If Text3 = MsgText(601) Then
'      adoadodc2.Open "select * from acc040 where a0401 = " & (Val(Text7) - 1) & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   Else
'      adoadodc2.Open "select * from acc040 where a0401 = " & (Val(Text7) - 1) & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   End If
'   Adodc2.Recordset.Requery
'   SumShow
End Sub

'Add by Amy 2024/08/28 避免產生 違反唯一限制條件及 調整預算時,導致調整後的餘額變0(因先刪資料再寫入,ex:只調5月,其他月餘額也變0)
'複製前一年預算
Private Sub Command1_Click()
   Dim intCounter As Integer, stQ As String
   
   If Acc040Check("Cmd1") = False Then
      strControlButton = MsgText(602)
      Exit Sub
   End If
   
   Text1.Locked = True
   strControlButton = MsgText(601)
   
   Call ReadAdodc2 '去年預算資料
   Call ReadAdodc1("Cmd1") '複製前一年預算
   
   Command1.Enabled = False
   Command2.Enabled = False
End Sub

'不複製前一年預算
Private Sub Command2_Click()
   Dim intCounter As Integer

   If Acc040Check("Cmd2") = False Then
      strControlButton = MsgText(602)
      Exit Sub
   End If
   
   Text1.Locked = True
   strControlButton = MsgText(601)
   
   Call ReadAdodc2 '去年預算資料
   Call ReadAdodc1("Cmd2") '不複製前一年預算
   
   Command1.Enabled = False
   Command2.Enabled = False
End Sub

Private Sub ReadAdodc1(stState As String)
   Dim strIns As String, stra0404 As String, i As Integer
  
   '部門
   If Text3 = MsgText(601) Then
      stra0404 = MsgText(55)
   Else
      stra0404 = Text3
   End If
   '不複製前一年預算 或 (複製前一年預算 且 前一年無資料)
   If stState = "Cmd2" Or (stState = "Cmd1" And adoadodc2.RecordCount = 0) Then
      For i = 1 To 12
         strIns = "Insert Into Acc040 (a0401,a0402,a0403,a0404,a0405,a0406,a0407,a0408,a0409,a0413) " & _
                        "Values(" & Text7 & "," & i & ",'" & Text1 & "','" & stra0404 & "','" & Text5 & "',0,0,0,0,'" & strUserNum & "')"
         adoTaie.Execute strIns
      Next i
   Else
      strIns = "Insert Into Acc040 (a0401,a0402,a0403,a0404,a0405,a0406,a0407,a0408,a0409,a0413) " & _
                     "Select " & Text7 & ",a0402,a0403,a0404,a0405,0,0,0,a0409,'" & strUserNum & "' " & _
                     "From Acc040 Where a0401 = " & Val(Text7) - 1 & " And a0403 = '" & Text1 & "' " & _
                     "And a0405 = '" & Text5 & "' And a0404='" & stra0404 & "' "
      adoTaie.Execute strIns
   End If
   adoacc040T.Requery
   AdodcRefresh
   SumShow
   RecordShow
End Sub

Private Sub ReadAdodc2()
   Dim strQ As String, stra0404 As String
  
   '部門
   If Text3 = MsgText(601) Then
      stra0404 = MsgText(55)
   Else
      stra0404 = Text3
   End If
   
   strQ = "Select * From Acc040 Where a0401 = " & Val(Text7) - 1 & " And a0403 = '" & Text1 & "' " & _
               "And a0404 = '" & stra0404 & "' And a0405 = '" & Text5 & "' Order by a0402 asc"
   If adoadodc2.State = adStateOpen Then adoadodc2.Close
   adoadodc2.CursorLocation = adUseClient
   adoadodc2.Open strQ, adoTaie, adOpenDynamic, adLockBatchOptimistic
   Adodc2.Recordset.Requery
   SumShow
End Sub

Private Sub ShowBt(Optional ByRef bCancel As Boolean)
   Dim stWhere As String
   
   If Text7 = MsgText(601) Or Text1 = MsgText(601) Or Text5 = MsgText(601) Then Exit Sub
   
   bCancel = False
   Command1.Enabled = False
   Command2.Enabled = False
   '新增
   If strSaveConfirm = MsgText(3) Then
      Call ReadAdodc2 '去年預算資料
      '部門
      If Text3 = MsgText(601) Then
         stWhere = "a0404='TOT' "
      Else
         stWhere = "a0404='" & Text3 & "' "
      End If
      stWhere = stWhere & "And a0401=" & Text7 & " And a0403='" & Text1 & "' And a0405='" & Text5 & "'"
      If Pub_GetField("Acc040", stWhere, "a0405", True) = "NoData" Then
         Command1.Enabled = True
         Command2.Enabled = True
      '新增時已按[不/複製前一年預算] 鈕,又在年度欄跳離開不需彈訊息
      ElseIf adoadodc1.RecordCount = 0 Then
         MsgBox Text7 & "年 " & Text1 & "公司" & vbCrLf & _
                           "會計科目[" & Text5 & "]資料已存在" & vbCrLf & _
                           "請按查詢再按修改", , MsgText(5)
         bCancel = True
      End If
   End If
End Sub

Public Sub SetData(State As String)
   Select Case State
      Case "F10"
         '修改->某月預算資料值(Gird已更新)->取消 鈕,畫面[不應該]維持已修改後的值
         If Text1 <> MsgText(601) And Text5 <> MsgText(601) And Text7 <> MsgText(601) Then
            Command3_Click
         End If
   End Select
End Sub

Public Sub Frmacc4160_Delete() '從aacc_del搬回修改
   Dim stWhere As String
On Error GoTo Checking
   
   With Frmacc4160
      Select Case strAccount
         Case "2"
            If DeleteCheck("select a0501 from acc050 where a0501 = " & Val(.Text7) & " and a0503 = '" & .Text1 & "' and a0504 = '" & .Text3 & "' and a0505 = '" & .Text5 & "'") = MsgText(603) Then
               Exit Sub
            End If
            adoTaie.Execute "delete from acc050 where a0501 = " & Val(.Text7) & " and a0503 = '" & .Text1 & "' and a0504 = '" & .Text3 & "' and a0505 = '" & .Text5 & "'"
         Case Else
            'Modify by Amy 2024/08/28 無資料會彈訊息,故不使用DeleteCheck
            '部門
            If Text3 = MsgText(601) Then
               stWhere = "a0404='TOT' "
            Else
               stWhere = "a0404='" & Text3 & "' "
            End If
            stWhere = stWhere & "And a0401 = " & Val(.Text7) & " and a0403 = '" & .Text1 & "' and a0405 = '" & .Text5 & "'"
            If Pub_GetField("Acc040", stWhere, "a0405", True) = "NoData" Then
               Exit Sub
            End If
            adoTaie.Execute "delete from acc040 where " & stWhere
            'end 2024/08/28
      End Select
      Frmacc4160_Clear 'Add by Amy 2024/08/28 避免刪了畫面仍存在
      .AdodcRefresh
      .SumShow
      .adoacc040T.Requery
      If .adoacc040T.RecordCount <> 0 Then
         .adoacc040T.MoveFirst
         .RecordShow
      Else
         StatusClear
      End If
   End With
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Public Sub MoveData(stMove As String)  '從aacc_fst/aacc_pre/aacc_nxt/aacc_lst搬回修改
   Select Case UCase(stMove)
      Case "FST"
         If adoacc040T.RecordCount <> 0 Then
            adoacc040T.MoveFirst
         End If
      Case "PRE"
         If adoacc040T.BOF = False Then
            adoacc040T.MovePrevious
            If adoacc040T.BOF Then
               adoacc040T.MoveFirst
               MsgBox MsgText(7), , MsgText(5)
            End If
         End If
      Case "NXT"
         If adoacc040T.EOF = False Then
            adoacc040T.MoveNext
            If adoacc040T.EOF Then
               adoacc040T.MoveLast
               MsgBox MsgText(8), , MsgText(5)
            End If
         End If
      Case "LST"
         If adoacc040T.RecordCount <> 0 Then
            adoacc040T.MoveLast
         End If
   End Select
   
   If adoacc040T.BOF = False Then
      FormShow
      AdodcRefresh
      SumShow
      RecordShow
   End If
End Sub
'end 2024/08/28

Private Sub Command3_Click()
   Dim strDepart As String
   Dim bolCancel As Boolean 'Add by Amy 2020/04/08
On Error GoTo Checking
 
   'Modify by Amy 2024/08/28 原只檢查text1,加新增查詢
   If Acc040Check("F5") = False Then
      Exit Sub
   End If
   
   '新增 按查詢鈕
   If strSaveConfirm = MsgText(3) Then
      Call Text7_Validate(bolCancel)
   '一般查詢
   Else
   
      If Text3 = "" Then
         strDepart = "TOT"
      Else
         strDepart = Text3
      End If
     
      adoacc040T.MoveFirst
      adoacc040T.Find "A0401 = " & Val(Text7) & "", 0, adSearchForward, 1
      If adoacc040T.EOF = False Then
         adoacc040T.Find "A0403 = '" & Text1 & "'", 0, adSearchForward, adoacc040T.Bookmark
         If adoacc040T.EOF = False Then
            adoacc040T.Find "A0404 = '" & strDepart & "'", 0, adSearchForward, adoacc040T.Bookmark
            If adoacc040T.EOF = False Then
               adoacc040T.Find "A0405 = '" & Text5 & "'", 0, adSearchForward, adoacc040T.Bookmark
               'Mark by Amy 2024/08/28 查第一筆有資料,查二筆無資料時,畫面不會更新,避免又按修改鈕
   '            If adoacc040T.EOF Then
   '               MsgBox MsgText(33), , MsgText(5)
   '               Exit Sub
   '            End If
            End If
         End If
      End If
      RecordShow
      AdodcRefresh
      If Adodc1.Recordset.RecordCount = 0 Then
         If Adodc2.Recordset.RecordCount = 0 Then
            MsgBox MsgText(33), , MsgText(5)
         Else
            MsgBox Text7 & "年資料不存在！", , MsgText(5)
         End If
      End If
   End If
   'end 2024/08/28
   Exit Sub
Checking:
   MsgBox MsgText(33), , MsgText(5)
   Exit Sub
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
   Adodc2.Recordset.UpdateBatch
   SumShow
End Sub

Private Sub DataGrid2_AfterColUpdate(ByVal ColIndex As Integer)
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   Select Case ColIndex
      Case 1
         If DataGrid2.Columns(1).Text = MsgText(601) Then
            DataGrid2.Columns(1).Value = 0
         End If
         'Add by Amy 2024/08/28 Grid修改值,只需更新當筆資料
         If strSaveConfirm = MsgText(4) Then  '修改
            Call Acc040Save(DataGrid2.Columns(0).Text)
         End If
         'end 2024/08/28
         Adodc1.Recordset.UpdateBatch
         SumShow
   End Select
End Sub

Private Sub DataGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn
         SendKeys "{DOWN}"
   End Select
End Sub

Private Sub Form_Activate()
   strFormName = Name
   If strCon1 = MsgText(601) Then
      Exit Sub
   End If
   
   adoacc040T.Find "a0401 = " & Val(strCon1) & "", 0, adSearchForward, 1
   If adoacc040T.EOF = False Then
      adoacc040T.Find "a0403 = '" & strCon2 & "'", 0, adSearchForward, adoacc040T.Bookmark
      If adoacc040T.EOF = False Then
         adoacc040T.Find "a0404 = '" & strCon3 & "'", 0, adSearchForward, adoacc040T.Bookmark
         If adoacc040T.EOF = False Then
            adoacc040T.Find "a0405 = '" & strCon4 & "'", 0, adSearchForward, adoacc040T.Bookmark
            If adoacc040T.EOF = False Then
               FormShow
               AdodcRefresh
               SumShow
               RecordShow
            End If
         End If
      End If
   End If
   strCon1 = MsgText(601)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 8900
   
'Modified by Morgan 2017/1/19 輸入預算若有Grid捲動時游標會消失,改加高能完整顯示12個月份
'   Me.Height = 5500
   
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   Me.Height = 6450
   PUB_InitForm Me, Me.Width, Me.Height
   'end 2017/1/19
   
   OpenTable
   'Text1 = "1" 'Modify by Amy 2014/01/02 不預帶
   FormDisabled
   'Modify by Amy 2024/08/28
   'If adoacc040T.RecordCount <> 0 Then
   If adoadodc1.RecordCount <> 0 Then
      RecordShow
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Cancel = 1
      Exit Sub
   End If
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc4160 = Nothing
End Sub

Private Sub Text1_Change()
   Text2 = ""
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   'Add by Amy 2020/04/08
   If InStr(GetBookKeepCmp, Text1) = 0 Then
     Text2 = ""
     Exit Sub
   End If
   Text2 = A0802Query(Text1)
   'Add by Amy 2024/08/28 新增時公司別/會科/部門/年度 可能都是空,需重抓資料
   If strSaveConfirm = MsgText(3) Then
      Command1.Enabled = False
      Command2.Enabled = False
      If adoadodc2.RecordCount = 0 Then
      Else
         ReadAdodc2 '去年預算資料
      End If
   End If
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
   adoacc040T.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         adoacc040T.Open "select a0501 as a0401, a0503 as a0403, a0504 as a0404, a0505 as a0405 from acc050 group by a0501, a0503, a0504, a0505", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Case Else
         adoacc040T.Open "select a0401, a0403, a0404, a0405 from acc040 group by a0401, a0403, a0404, a0405", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End Select
   
   adoadodc1.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         adoadodc1.Open "select a0501 as a0401, a0502 as a0402, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0506 as a0406, a0507 as a0407, a0508 as a0408, a0509 as a0409, a0511 as a0411, a0512 as a0412, a0513 as a0413, a0514 as a0414, a0515 as a0415, a0516 as a0416 from acc050 where a0501 = " & Val(Text7) & " and a0503 = '" & Text1 & "' and a0504 = '" & Text3 & "' and a0505 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Case Else
         adoadodc1.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End Select
   Set Adodc1.Recordset = adoadodc1
   
   adoadodc2.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         adoadodc2.Open "select a0501 as a0401, a0502 as a0402, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0506 as a0406, a0507 as a0407, a0508 as a0408, a0509 as a0409, a0511 as a0411, a0512 as a0412, a0513 as a0413, a0514 as a0414, a0515 as a0415, a0516 as a0416 from acc050 where a0501 = " & (Val(Text7) - 1) & " and a0503 = '" & Text1 & "' and a0504 = '" & Text3 & "' and a0505 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
      Case Else
         adoadodc2.Open "select * from acc040 where a0401 = " & (Val(Text7) - 1) & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   End Select
   Set Adodc2.Recordset = adoadodc2
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Amy 2013/12/24
Private Sub Text1_KeyPress(KeyAscii As Integer)
     KeyAscii = UpperCase(KeyAscii)
End Sub
'end 2013/12/24

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 = MsgText(601) Then
      Exit Sub
   End If
   'Add by Amy 2014/01/02
   'Modify by Amy 2020/04/08
   'If Text1 <> "1" And Text1 <> "J" Then
   If InStr(GetBookKeepCmp, Text1) = 0 Then
         MsgBox Label4 & MsgText(63), , MsgText(5) '原:"公司別只可輸入 1 或 J"
   'end 2020/04/08
         Cancel = True
         Text1.SetFocus
         Exit Sub
   End If
   'end 2014/01/02
   If ExistCheck("acc080", "a0801", Text1, Label4) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text3_Change()
   Text4 = A0902Query(Text3)
   'Add by Amy 2024/08/28 新增時公司別/會科/部門/年度 可能都是空,需重抓資料
   If strSaveConfirm = MsgText(3) Then
      Command1.Enabled = False
      Command2.Enabled = False
      If adoadodc2.RecordCount = 0 Then
      Else
         ReadAdodc2 '去年預算資料
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   If Text3 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc090", "a0901", Text3, Label1) = False Then
      Cancel = True
      Exit Sub
   End If
End Sub

Private Sub Text5_Change()
   If Text5 = MsgText(601) Then
      Exit Sub
   End If
   Text6 = A0102Query(Text5)
   adoacc010.CursorLocation = adUseClient
   adoacc010.Open "select * from acc010 where a0101 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc010.RecordCount <> 0 Then
      If IsNull(adoacc010.Fields("a0103").Value) Then
         Text9 = MsgText(601)
      Else
         Select Case adoacc010.Fields("a0103").Value
            Case Mid(ComboItem(1), 1, 1)
               Text9 = ComboItem(1)
            Case Mid(ComboItem(2), 1, 1)
               Text9 = ComboItem(2)
         End Select
      End If
   Else
      Text9 = MsgText(601)
   End If
   adoacc010.Close
   'Add by Amy 2024/08/28 新增時公司別/會科/部門/年度 可能都是空,需重抓資料
   If strSaveConfirm = MsgText(3) Then
      Command1.Enabled = False
      Command2.Enabled = False
      If adoadodc2.RecordCount = 0 Then
      Else
         ReadAdodc2 '去年預算資料
      End If
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   'Modify by Amy 2014/01/07 +公司別確認改用PUB_CheckCompany
'   If ExistCheck("acc010", "a0101", Text5, Label2) = False Then
'      Cancel = True
'      Exit Sub
'   End If
   If PUB_CheckCompany(Text5, Text1) = False Then
         Cancel = True
         Exit Sub
   End If
   'end 2014/01/07
   '2012/1/12 add by sonia 科目名稱內有 '不用' 二字者不可新增
   If InStr(Text6, "不用") > 0 Then
      MsgBox "此科目已不再使用, 不可輸入預算資料 ! ", , MsgText(5)
      Cancel = True
      Exit Sub
   End If
   '2012/1/12 end
End Sub

'Add by Amy 2024/08/28
Private Sub Text7_Change()
   'Add by Amy 2024/08/28 新增時公司別/會科/部門/年度 可能都是空,需重抓資料
   If strSaveConfirm = MsgText(3) Then
      Command1.Enabled = False
      Command2.Enabled = False
   End If
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

'*************************************************
'  儲存欄位資料(會計科目餘額資料)
'
'*************************************************
'Modify by Amy 2024/08/28  原:Private
Public Sub Acc040Save(Optional ByVal stMon As String)
   Dim strUpd As String, strField As String, stra0404 As String
  
   '部門
   If Text3 = MsgText(601) Then
      stra0404 = MsgText(55)
   Else
      stra0404 = Text3
   End If
   
   '新增為前一年全部月份資料,記錄人/日/時,設為同樣的值
   If strSaveConfirm = MsgText(3) Then
      strField = "a0411=" & Val(strSrvDate(2)) & ",a0412=" & ServerTime & " "
      strUpd = "Update Acc040 Set " & strField & _
                     "Where a0401 = " & Val(Text7) & " And a0403 = '" & Text1 & "' " & _
                     "And a0405 = '" & Text5 & "' And a0404='" & stra0404 & "' "
   'Grid修改值,只需更新當筆資料
   Else
      strField = "a0416='" & strUserNum & "',a0414=" & Val(strSrvDate(2)) & ",a0415=" & ServerTime & " "
      strUpd = "Update Acc040 Set " & strField & _
                     "Where a0401 = " & Val(Text7) & " And a0403 = '" & Text1 & "' " & _
                     "And a0405 = '" & Text5 & "' And a0404='" & stra0404 & "' And a0402=" & Val(stMon)
   End If
   adoTaie.Execute strUpd
   'end 2024/08/28
   
   'Mark by Amy 2024/08/28  改寫法下面程式不使用
'   Adodc1.Recordset.Fields("a0401").Value = Val(Text7)
'   Adodc1.Recordset.Fields("a0403").Value = Text1
'   If Text3 = MsgText(601) Then
'      Adodc1.Recordset.Fields("a0404").Value = MsgText(55)
'   Else
'      Adodc1.Recordset.Fields("a0404").Value = Text3
'   End If
'   Adodc1.Recordset.Fields("a0405").Value = Text5
'   Adodc1.Recordset.Fields("a0413").Value = strUserNum
End Sub

'Mark by Amy 2024/08/28 不使用,財務有調整預算(修改資料,會先刪資料再寫入),導致調整後的餘額變0
'Added by Lydia 2019/11/28 刪除先前輸入資料,避免重複主鍵的程式錯誤
'Private Sub Acc040DeleteUpd()
'    If strSaveConfirm = MsgText(3) Then  '新增
'         cnnConnection.Execute "delete from acc040 where a0401=" & Val(Text7) & " and a0403='" & Text1 & "' and a0404='" & IIf(Text3 = MsgText(601), MsgText(55), Text3) & "' and a0405='" & Text5 & "' "
'    End If
'End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         If Text3 = MsgText(601) Then
            adoadodc1.Open "select a0501 as a0401, a0502 as a0402, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0506 as a0406, a0507 as a0407, a0508 as a0408, a0509 as a0409, a0511 as a0411, a0512 as a0412, a0513 as a0413, a0514 as a0414, a0515 as a0415, a0516 as a0416 from acc050 where a0501 = " & Val(Text7) & " and a0503 = '" & Text1 & "' and a0504 = '" & MsgText(55) & "' and a0505 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         Else
            adoadodc1.Open "select a0501 as a0401, a0502 as a0402, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0506 as a0406, a0507 as a0407, a0508 as a0408, a0509 as a0409, a0511 as a0411, a0512 as a0412, a0513 as a0413, a0514 as a0414, a0515 as a0415, a0516 as a0416 from acc050 where a0501 = " & Val(Text7) & " and a0503 = '" & Text1 & "' and a0504 = '" & Text3 & "' and a0505 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         End If
      Case Else
         If Text3 = MsgText(601) Then
            adoadodc1.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         Else
            adoadodc1.Open "select * from acc040 where a0401 = " & Val(Text7) & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         End If
   End Select
   Adodc1.Recordset.Requery
   
   adoadodc2.Close
   adoadodc2.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         If Text3 = MsgText(601) Then
            adoadodc2.Open "select a0501 as a0401, a0502 as a0402, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0506 as a0406, a0507 as a0407, a0508 as a0408, a0509 as a0409, a0511 as a0411, a0512 as a0412, a0513 as a0413, a0514 as a0414, a0515 as a0415, a0516 as a0416 from acc050 where a0501 = " & (Val(Text7) - 1) & " and a0503 = '" & Text1 & "' and a0504 = '" & MsgText(55) & "' and a0505 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         Else
            adoadodc2.Open "select a0501 as a0401, a0502 as a0402, a0503 as a0403, a0504 as a0404, a0505 as a0405, a0506 as a0406, a0507 as a0407, a0508 as a0408, a0509 as a0409, a0511 as a0411, a0512 as a0412, a0513 as a0413, a0514 as a0414, a0515 as a0415, a0516 as a0416 from acc050 where a0501 = " & (Val(Text7) - 1) & " and a0503 = '" & Text1 & "' and a0504 = '" & Text3 & "' and a0505 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         End If
      Case Else
         If Text3 = MsgText(601) Then
            adoadodc2.Open "select * from acc040 where a0401 = " & (Val(Text7) - 1) & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         Else
            adoadodc2.Open "select * from acc040 where a0401 = " & (Val(Text7) - 1) & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
         End If
   End Select
   Adodc2.Recordset.Requery
   SumShow
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  計算並顯示合計
'
'*************************************************
Public Sub SumShow()
   adoaccsum.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         If Text3 = MsgText(601) Then
            adoaccsum.Open "select sum(a0509) from acc050 where a0501 = " & Val(Text7) & " and a0503 = '" & Text1 & "' and a0504 = '" & MsgText(55) & "' and a0505 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoaccsum.Open "select sum(a0509) from acc050 where a0501 = " & Val(Text7) & " and a0503 = '" & Text1 & "' and a0504 = '" & Text3 & "' and a0505 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
         End If
      Case Else
         If Text3 = MsgText(601) Then
            adoaccsum.Open "select sum(a0409) from acc040 where a0401 = " & Val(Text7) & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoaccsum.Open "select sum(a0409) from acc040 where a0401 = " & Val(Text7) & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
         End If
   End Select
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text10 = MsgText(601)
      Else
         Text10 = Format(adoaccsum.Fields(0).Value, DDollar)
      End If
   Else
      Text10 = MsgText(601)
   End If
   adoaccsum.Close
   adoaccsum.CursorLocation = adUseClient
   Select Case strAccount
      Case "2"
         If Text3 = MsgText(601) Then
            adoaccsum.Open "select sum(a0509) from acc050 where a0501 = " & Val(Text7) & " and a0503 = '" & Text1 & "' and a0504 = '" & MsgText(55) & "' and a0505 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoaccsum.Open "select sum(a0509) from acc050 where a0501 = " & Val(Text7) & " and a0503 = '" & Text1 & "' and a0504 = '" & Text3 & "' and a0505 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
         End If
      Case Else
         If Text3 = MsgText(601) Then
            adoaccsum.Open "select sum(a0409) from acc040 where a0401 = " & (Val(Text7) - 1) & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adoaccsum.Open "select sum(a0409) from acc040 where a0401 = " & (Val(Text7) - 1) & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "'", adoTaie, adOpenStatic, adLockReadOnly
         End If
   End Select
   If adoaccsum.RecordCount <> 0 Then
      If IsNull(adoaccsum.Fields(0).Value) Then
         Text8 = MsgText(601)
      Else
         Text8 = Format(adoaccsum.Fields(0).Value, DDollar)
      End If
   Else
      Text8 = MsgText(601)
   End If
   adoaccsum.Close
End Sub

'*************************************************
'  功能鍵設定-非使用狀態
'
'*************************************************
Public Sub FormDisabled()
   Text1.Locked = False
   Command2.Enabled = False
   Command1.Enabled = False
   DataGrid2.Enabled = False
End Sub

'*************************************************
'  功能鍵設定-可使用狀態
'
'*************************************************
Public Sub FormEnabled()
   'Add byAmy 2013/12/26 修改鎖公司別
   If strSaveConfirm = MsgText(4) Then
      Text1.Locked = True
   Else
       Text1.Locked = False
   'Modify by Amy 2024/08/28 只有新增可使用 [不/複製前一年預算]鈕
        Call ShowBt
   End If
   'end 2013/12/26
'   Command2.Enabled = True
'   Command1.Enabled = True
    'end 2024/08/28
   DataGrid2.Enabled = True
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
   Dim bCancel As Boolean 'Add by Amy 2024/08/28
   
   If strSaveConfirm <> MsgText(3) And strSaveConfirm <> MsgText(4) Then
      Exit Sub
   End If
   'Mark by Amy 2024/08/28 新增->輸完條件->年度跳離開彈已有資料,一定要輸沒資料的年度才能跳離
   '年度
'   If Text7 = MsgText(601) Then
'      MsgBox Label5 & MsgText(52), , MsgText(5)
'      Cancel = True
'      Exit Sub
'   End If
   'Add by Amy 2024/08/28 判斷已有資料彈訊息
   '                        查 1公司113年110201->修改->改Grid2某個月資料->按[複製前一年預算] 會出現 違反唯一...錯誤
   Call Text5_Validate(bCancel) '會計科目
   If bCancel = True Then
      Cancel = True
      Exit Sub
   End If
   Call ShowBt(bCancel)
   If bCancel = True Then
      If Me.ActiveControl.Name <> "Text7" Then Cancel = True '游標跳離,避免無法改其他欄
      Exit Sub
   End If
   'end 2024/08/28
End Sub

'*************************************************
'  儲存前檢查欄位是否正確
'
'*************************************************
'Modify by Amy 2024/08/28 讓F9 也可使用,加 stState 拿掉strControlButton = MsgText(602) ,原Sub
Public Function Acc040Check(stState As String) As Boolean
   Dim Cancel As Boolean 'Add by Amy 2014/01/02
   
   Acc040Check = False
   '公司別
   If Text1 = MsgText(601) Then
      MsgBox MsgText(10) & Label4, , MsgText(5)
      Text1.SetFocus
      Exit Function
   Else
      'Add by Amy 2014/01/02
      Call Text1_Validate(Cancel)
      If Cancel = True Then
         Text1.SetFocus
         Exit Function
      End If
      '會計科目
      If Text5 = MsgText(601) Then
         MsgBox MsgText(10) & Label2, , MsgText(5)
         Text5.SetFocus
         Exit Function
      End If
      Call Text5_Validate(Cancel)
      If Cancel = True Then
         Text5.SetFocus
         Exit Function
      End If
      'end 2014/01/02
      '年度
      If Text7 = MsgText(601) Then
         MsgBox Label5 & MsgText(52), , MsgText(5)
         Text7.SetFocus
         Exit Function
      End If
      'Add by Amy 2024/08/28
      Call Text7_Validate(Cancel)
      If Cancel = True Then
         Text7.SetFocus
         Exit Function
      End If
      'end 2024/08/28
      If ExistCheck("acc080", "a0801", Text1, Label4) = False Then
         Text1.SetFocus
         Exit Function
      End If
      '部門
      If Text3 <> MsgText(601) Then
         Call Text3_Validate(Cancel)
         If Cancel = True Then
            Text3.SetFocus
            Exit Function
         End If
         'Add by Amy 2024/08/28 110年起L公司才有L部門,其他公司不可有L部門
         If stState <> "F5" And stState <> "F3" And Val(Text7) >= 110 Then
            If (Text1 <> "L" And Text3 = "L") Or (Text1 = "L" And InStr(Text3, "L") = 0 And Text3 <> "" And Text3 <> "TOT") Then
               MsgBox "110年開始 [" & Text1 & "]公司 不可有 [" & Text3 & "]部門", , MsgText(5)
               Text3.SetFocus
               Exit Function
            End If
         End If
      End If
   End If
   
   If stState = "Cmd1" Or stState = "Cmd2" Then
      If Left(strSrvDate(1), 4) > Val(Text7) + 1911 Then
         If MsgBox("確定要複製 [" & Val(Text7) - 1 & "]年 資料 到 [" & Text7 & "]年" & vbCrLf & _
            "是:繼續 否:不複製", vbYesNo, "請確認") = vbNo Then
            Command1.Enabled = True
            Command2.Enabled = True
            Exit Function
         End If
      End If
   '修改
   ElseIf stState = "F3" Then
      If adoadodc1.RecordCount = 0 Then
         MsgBox "無資料可修改", , MsgText(5)
         Command3.SetFocus
         Exit Function
      End If
   '存檔
   ElseIf stState = "F9" Then
      If adoadodc1.RecordCount = 0 Then
         MsgBox "無資料需存檔", , MsgText(5)
         Exit Function
      End If
   End If
   Acc040Check = True
End Function

'*************************************************
'  顯示筆數
'
'*************************************************
Public Sub RecordShow()
   'Modify by Amy 2024/08/28 +if
   If adoacc040T.EOF = False Then
      Frmacc0000.StatusBar1.Panels(2).Text = adoacc040T.Bookmark & MsgText(35) & adoacc040T.RecordCount
   Else
      Frmacc0000.StatusBar1.Panels(2).Text = ""
   End If
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   'Modify by Amy 2024/08/28 +if
   If adoacc040T.EOF = False Then
      Text1 = adoacc040T.Fields("a0403").Value
      If adoacc040T.Fields("a0404").Value = MsgText(55) Then
         Text3 = MsgText(601)
      Else
         Text3 = adoacc040T.Fields("a0404").Value
      End If
      Text5 = adoacc040T.Fields("a0405").Value
      Text7 = adoacc040T.Fields("a0401").Value
   End If
End Sub

'Add by Amy 2013/12/24 從aacc_cls.bas搬回
'Modified by Lydia 2017/01/26 新增時預設上一筆輸入的下一會計科目
'Public Sub Frmacc4160_Clear()
Public Sub Frmacc4160_Clear(Optional bolAddNew As Boolean = False)
'Added by Lydia 2017/01/26
Dim rsB As New ADODB.Recordset
Dim bCancel As Boolean 'Add by Amy 2024/08/28

    If bolAddNew = False Then
       Text1.Tag = "": Text3.Tag = "": Text5.Tag = "": Text7.Tag = ""
    Else
       Text1.Tag = Text1.Text: Text3.Tag = Text3.Text: Text5.Tag = Text5.Text: Text7.Tag = Text7.Text
    End If
'end 2017/01/26

      'Modify by Amy 2014/01/02
      Text1.Locked = False
      Text1 = "" '原:"1"
      Text2 = ""
      'end 2014/01/02
      Text3 = ""
      Text4 = ""
      Text5 = ""
      Text6 = ""
      Text9 = ""
      Text7 = ""
      AdodcRefresh
      Text1.SetFocus
      
   'Added by Lydia 2017/01/26 新增時預設上一筆輸入的下一會計科目
   If bolAddNew = True And Text5.Tag <> "" Then
      'Ｍodify by Amy 2024/08/28  4 字頭只抓4碼,6字頭照舊-秀玲與財務討論(4 字頭及 6 字頭會輸預算)
      If Left(Text5.Tag, 1) = "4" Then
         strSql = "select min(a0101) from acc010 where a0101>'" & Text5.Tag & "' and length(a0101)=4 and instr(a0102,'不用')=0  "
      Else
         strSql = "select min(a0101) from acc010 where a0101>'" & Text5.Tag & "' and length(a0101)>=4 and instr(a0102,'不用')=0  "
      End If
      'end 2024/08/28
      intI = 1
      Set rsB = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If "" & rsB(0) <> "" Then
            Text1.Text = Text1.Tag: Text3.Text = Text3.Tag: Text5.Text = "" & rsB(0): Text7.Text = Text7.Tag
            Text1.Tag = "": Text3.Tag = "": Text5.Tag = "": Text7.Tag = ""
            'Add by Amy 2024/08/28 已有資料不可按新增
            Call ShowBt(bCancel)
            If bCancel = True Then
               strSaveConfirm = MsgText(601)
               Set rsB = Nothing
               Exit Sub
            End If
            'end 2024/08/28
            'Added by Lydia 2019/11/28 帶出該科目的去年預算(DataGrid2) ,今年預算由操作者決定複製or不複製
            'Modify by Amy 2024/08/28 改抓共用函數
'            adoadodc2.Close
'            adoadodc2.CursorLocation = adUseClient
'            If Text3 = MsgText(601) Then
'               adoadodc2.Open "select * from acc040 where a0401 = " & (Val(Text7) - 1) & " and a0403 = '" & Text1 & "' and a0404 = '" & MsgText(55) & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'            Else
'               adoadodc2.Open "select * from acc040 where a0401 = " & (Val(Text7) - 1) & " and a0403 = '" & Text1 & "' and a0404 = '" & Text3 & "' and a0405 = '" & Text5 & "' order by a0402 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
'            End If
'            Adodc2.Recordset.Requery
'            SumShow
'            'end 2019/11/28
            Call ReadAdodc2 '去年預算資料
            'end 2024/08/28
         End If
      End If
      Set rsB = Nothing
      Text5.SetFocus
      Text5_GotFocus
   End If
End Sub
'end 2013/12/24

