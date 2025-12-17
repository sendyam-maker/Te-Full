VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2212 
   AutoRedraw      =   -1  'True
   Caption         =   "收款資料查詢"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   8730
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4350
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   15
      Top             =   210
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4350
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   14
      Top             =   570
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4350
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   13
      Top             =   930
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   6
      Top             =   5002
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   1572
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6960
      TabIndex        =   3
      Top             =   5002
      Width           =   1572
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   248
      Value           =   -1  'True
      Width           =   240
   End
   Begin VB.OptionButton Option2 
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   608
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Height          =   255
      Left            =   4080
      TabIndex        =   0
      Top             =   968
      Width           =   255
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2212.frx":0000
      Height          =   3015
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "a0x02"
         Caption         =   "請款編號"
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
         DataField       =   "a0x03"
         Caption         =   "規費"
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
      BeginProperty Column02 
         DataField       =   "a0x04"
         Caption         =   "案件性質"
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
         DataField       =   "a0x05"
         Caption         =   "請款外幣"
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
         DataField       =   "a0x06"
         Caption         =   "折讓外幣"
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
         DataField       =   "a0x08"
         Caption         =   "幣別"
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
         DataField       =   "a0x11"
         Caption         =   "本次外幣"
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
      BeginProperty Column07 
         DataField       =   "a0x10"
         Caption         =   "結清"
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
         DataField       =   "a0x07"
         Caption         =   "已收外幣"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   510.236
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1065.26
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   360
      Top             =   1800
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSForms.TextBox Text5 
      Height          =   330
      Left            =   5820
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   210
      Width           =   2775
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "4895;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   330
      Left            =   5820
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   570
      Width           =   2775
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "4895;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text9 
      Height          =   330
      Left            =   5820
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   930
      Width           =   2775
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "4895;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "溢收金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "收款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "代理人1"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   248
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "代理人2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   608
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "代理人3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   968
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1245
      Left            =   360
      Top             =   120
      Width           =   8295
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   120
      Top             =   5400
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "轉入暫收款單號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   5040
      Width           =   1695
   End
End
Attribute VB_Name = "Frmacc2212"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/07 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Text4~Text9
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit
Public adoacc0y0 As New ADODB.Recordset
Public adoacc1k0 As New ADODB.Recordset
Public adoacc0x0 As New ADODB.Recordset
Public adocaseprogress As New ADODB.Recordset
Public adoacc0z0 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Dim douAmount As Double

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/09 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 8850
'   Me.Height = 6000
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath1)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   PUB_InitForm Me, 8850, 5900, strBackPicPath1
   'end 2021/12/09
   Text3 = strItemNo
   Acc0x0Show
   OpenTable
End Sub

Private Sub Form_Unload(Cancel As Integer)
   adoTaie.Execute "delete from acc0x0"
   strItemNo = ""
   tool3_enabled
   Select Case strFormLink
      Case "Frmacc2210"
         Frmacc2210.Enabled = True
      Case "Frmacc2220"
         Frmacc2220.Enabled = True
   End Select
   Set Frmacc2212 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoacc0y0.CursorLocation = adUseClient
   adoacc0y0.Open "select * from acc0y0 where a0y01 = '" & strItemNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   FormShow
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc0x0 where a0x01 = '" & strItemNo & "' order by a0x02 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  顯示資料表
'
'*************************************************
Public Sub FormShow()
   If IsNull(adoacc0y0.Fields("a0y07").Value) Then
      Text4 = MsgText(601)
   Else
      Text4 = adoacc0y0.Fields("a0y07").Value
      Text5 = FagentQuery(Text4, 1)
      If Text5 = "" Then
         Text5 = FagentQuery(Text4, 2)
      End If
      If Text5 = "" Then
         Text5 = CustomerQuery(Text4, 1)
      End If
      If Text5 = "" Then
         Text5 = CustomerQuery(Text4, 2)
      End If
   End If
   If IsNull(adoacc0y0.Fields("a0y08").Value) Then
      Text6 = MsgText(601)
   Else
      Text6 = adoacc0y0.Fields("a0y08").Value
   End If
   If IsNull(adoacc0y0.Fields("a0y09").Value) Then
      Text8 = MsgText(601)
   Else
      Text8 = adoacc0y0.Fields("a0y09").Value
      If Len(Text8) = 6 Then
         Text9 = FagentQuery(AfterZero(Text8), 2)
      Else
         Text9 = FagentQuery(Text8, 2)
      End If
   End If
   If IsNull(adoacc0y0.Fields("a0y10").Value) Then
      Text10 = MsgText(601)
   Else
      Text10 = adoacc0y0.Fields("a0y10").Value
   End If
   Select Case adoacc0y0.Fields("a0y18").Value
      Case 1
         Option1.Value = True
      Case 2
         Option2.Value = True
      Case 3
         Option3.Value = True
   End Select
   If IsNull(adoacc0y0.Fields("a0y06").Value) Then
      Text2 = MsgText(601)
   Else
      Text2 = adoacc0y0.Fields("a0y06").Value
   End If
End Sub

'*************************************************
'  儲存資料表(請款單收款記錄資料)
'
'*************************************************
Private Sub Acc0x0Show()
   adoacc0x0.CursorLocation = adUseClient
   adoacc0x0.Open "select * from acc0x0 where a0x01 = '" & Text3 & "' order by a0x02 asc", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0z0.CursorLocation = adUseClient
   adoacc0z0.Open "select * from acc0z0, acc0y0 where a0z01 = a0y01 and a0z01 = '" & Text3 & "'", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc0z0.EOF = False
      adoacc0x0.AddNew
      adoacc0x0.Fields("a0x01").Value = adoacc0z0.Fields("a0z01").Value
      adoacc0x0.Fields("a0x02").Value = adoacc0z0.Fields("a0z02").Value
      adoacc1k0.CursorLocation = adUseClient
      adoacc1k0.Open "select * from acc1k0 where a1k01 = '" & adoacc0z0.Fields("a0z02").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc1k0.RecordCount <> 0 Then
         If IsNull(adoacc1k0.Fields("a1k09").Value) Then
            adoacc0x0.Fields("a0x03").Value = 0
         Else
            adoacc0x0.Fields("a0x03").Value = adoacc1k0.Fields("a1k09").Value
         End If
         adocaseprogress.CursorLocation = adUseClient
         adocaseprogress.Open "select nvl(cpm03, cpm04) from caseprogress, casepropertymap where cp01 = cpm01 and cp10 = cpm02 and cp60 = '" & adoacc1k0.Fields("a1k01").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adocaseprogress.RecordCount <> 0 Then
            If IsNull(adocaseprogress.Fields(0).Value) Then
               adoacc0x0.Fields("a0x04").Value = Null
            Else
               adoacc0x0.Fields("a0x04").Value = adocaseprogress.Fields(0).Value
            End If
         Else
            adoacc0x0.Fields("a0x04").Value = Null
         End If
         adocaseprogress.Close
         If IsNull(adoacc1k0.Fields("a1k10").Value) Then
            adoacc0x0.Fields("a0x05").Value = 0
            adoacc0x0.Fields("a0x06").Value = 0
            adoacc0x0.Fields("a0x07").Value = 0
         Else
            If IsNull(adoacc1k0.Fields("a1k08").Value) Then
               adoacc0x0.Fields("a0x05").Value = 0
            Else
               adoacc0x0.Fields("a0x05").Value = adoacc1k0.Fields("a1k08").Value
            End If
            If IsNull(adoacc1k0.Fields("a1k06").Value) Then
               adoacc0x0.Fields("a0x06").Value = 0
            Else
               adoacc0x0.Fields("a0x06").Value = adoacc1k0.Fields("a1k06").Value
            End If
            If IsNull(adoacc1k0.Fields("a1k30").Value) Then
               adoacc0x0.Fields("a0x07").Value = 0
            Else
               adoacc0x0.Fields("a0x07").Value = Val(Format(adoacc1k0.Fields("a1k30").Value / adoacc0z0.Fields("a0y04").Value, FAmount))
            End If
         End If
         If IsNull(adoacc1k0.Fields("a1k29").Value) Then
            adoacc0x0.Fields("a0x10").Value = Null
         Else
            adoacc0x0.Fields("a0x10").Value = adoacc1k0.Fields("a1k29").Value
         End If
         If IsNull(adoacc1k0.Fields("a1k29").Value) Then
            adoacc0x0.Fields("a0x10").Value = ""
         Else
            adoacc0x0.Fields("A0X10").Value = adoacc1k0.Fields("a1k29").Value
         End If
      End If
      adoacc1k0.Close
      '2005/6/21 MODIFY BY SONIA
      'adoacc0x0.Fields("a0x08").Value = "US$"
      'Modify By Sindy 2012/10/23 原程式抓A0Z03改抓A0Y03
      'adoacc0x0.Fields("a0x08").Value = adoacc0z0.Fields("A0Z03").Value
      adoacc0x0.Fields("a0x08").Value = adoacc0z0.Fields("A0Y03").Value
      '2012/10/23 End
      '2005/6/21 END
      adoacc0x0.Fields("A0X11").Value = adoacc0z0.Fields("A0Z04").Value
      adoacc0x0.Fields("A0X09").Value = Val(Format(Val(adoacc0x0.Fields("A0X11").Value) * Val(strCon3), FAmount))
      adoacc0x0.UpdateBatch
      adoacc0z0.MoveNext
   Loop
   adoacc0z0.Close
   adoacc0x0.Close
End Sub
