VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc12c0 
   AutoRedraw      =   -1  'True
   Caption         =   "發票資料查詢"
   ClientHeight    =   5112
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9528
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5112
   ScaleWidth      =   9528
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1650
      TabIndex        =   2
      Top             =   600
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
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
      Left            =   3570
      TabIndex        =   1
      Top             =   240
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
      Height          =   300
      Left            =   1650
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc12c0.frx":0000
      Height          =   3495
      Left            =   105
      TabIndex        =   10
      Top             =   1410
      Width           =   9210
      _ExtentX        =   16235
      _ExtentY        =   6160
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.4
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
      Caption         =   "發票資料查詢"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "a4302"
         Caption         =   "發票日期"
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
         DataField       =   "a4301"
         Caption         =   "發票號碼"
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
      BeginProperty Column02 
         DataField       =   "a4303"
         Caption         =   "統一編號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "a4201"
         Caption         =   "收據抬頭/客戶名稱"
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
         DataField       =   "a4304"
         Caption         =   "銷售額"
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
      BeginProperty Column05 
         DataField       =   "a4305"
         Caption         =   "稅額"
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
      BeginProperty Column06 
         DataField       =   "axc02"
         Caption         =   "請款單編號"
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
      BeginProperty Column07 
         DataField       =   "a4306"
         Caption         =   "列印次數"
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
         DataField       =   "a4319"
         Caption         =   "發票開立上傳"
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
         DataField       =   "a4308"
         Caption         =   "作廢日期"
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
      BeginProperty Column10 
         DataField       =   "a4321"
         Caption         =   "發票作廢上傳"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "###/##/##"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "a4324"
         Caption         =   "折讓/銷退上傳日"
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
            ColumnWidth     =   1031.811
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1044.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2027.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1008
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnWidth     =   1031.811
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   -1  'True
            ColumnWidth     =   1476.284
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   120
      Top             =   1290
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3570
      TabIndex        =   3
      Top             =   600
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   1650
      TabIndex        =   4
      Top             =   960
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   300
      Left            =   3570
      TabIndex        =   5
      Top             =   960
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "作廢日期"
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
      Left            =   690
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label6 
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
      Left            =   3330
      TabIndex        =   11
      Top             =   960
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   135
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
      Height          =   255
      Left            =   3330
      TabIndex        =   9
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "發票/銷退日期"
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
      Left            =   150
      TabIndex        =   8
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label2 
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
      Left            =   3330
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "發票號碼"
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
      Left            =   690
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc12c0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改
'20140121Create By eric  發票作廢查詢
'2016/3/11 瑞婷說發票作廢查詢改發票資料查詢
Option Explicit

Public adoadodc1 As New ADODB.Recordset


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
Dim stDef As String 'Add by Amy 2020/02/18
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/08/18 W9500 H5400
   Me.Width = 9620
   Me.Height = 5560
   'end 2023/08/18
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Modify by Amy 2020/02/24 預設前一工作日~當日
   stDef = PUB_GetWorkDay1(strSrvDate(1) - 1, 1)
   MaskEdBox1.Text = CFDate(TransDate(stDef, 1))
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Text = CFDate(strSrvDate(2))
   MaskEdBox2.Mask = DFormat
   'end 2020/02/24
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
   
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)

   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc1260 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   
   'Modify by Amy 2019/07/01 +上傳日期時間
   'Modifyb by Amy 2019/11/25 改顯示欄位順序及加a1k01
   If adoadodc1.State <> adStateClosed Then
      adoadodc1.Close
      adoadodc1.CursorLocation = adUseClient
      adoadodc1.Open "select a4302 ,a4301,a4303 ,'' a4201, a4304 , a4305,'' as axc02, a4306,a4319||a4320,a4308,a4321||a4322 From acc430 WHERE ROWNUM<1", adoTaie, adOpenStatic, adLockReadOnly
   Else
      adoadodc1.CursorLocation = adUseClient
      adoadodc1.Open "select a4302 ,a4301,a4303 ,'' a4201, a4304 , a4305,'' as axc02, a4306,a4319||a4320,a4308,a4321||a4322 From acc430 WHERE ROWNUM<1", adoTaie, adOpenStatic, adLockReadOnly
   End If
   'end 2019/07/01
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'*************************************************
'  重新整理 Adodc 之資料
'*************************************************
Public Sub AdodcRefresh()
Dim txtNum3 As Long
Dim txtNum4 As Long
Dim txtNum5 As Long
Dim txtNum6 As Long
Dim strQ As String 'Add by Amy 2019/11/25
 
On Error GoTo Checking
  
   strSql = ""
   '畫面發票號碼
   If Text1 <> MsgText(601) Then
     strSql = " and a4301 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a4301 <= '" & Text2 & "'"
   End If
   '畫面發票日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      txtNum3 = Val(FCDate(MaskEdBox1.Text))
      strSql = strSql & " and a4302 >= " & txtNum3 & " "
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      txtNum4 = Val(FCDate(MaskEdBox2.Text))
      strSql = strSql & " and a4302 <= " & txtNum4 & " "
   End If
   '畫面作廢日期
   If MaskEdBox3.Text <> MsgText(601) And MaskEdBox3.Text <> MsgText(29) Then
      txtNum5 = Val(FCDate(MaskEdBox3.Text))
      strSql = strSql & " and a4308 >= " & txtNum5 & " "
   End If
   If MaskEdBox4.Text <> MsgText(601) And MaskEdBox4.Text <> MsgText(29) Then
      txtNum6 = Val(FCDate(MaskEdBox4.Text))
      strSql = strSql & " and a4308 <= " & txtNum6 & " "
   End If
   'GET 作廢日期欄位非0及NULL的資料
'2016/3/11 cancel by sonia 瑞婷說發票作廢查詢改發票資料查詢
'   strSql = strSql & " and ( a4308<>0 or a4308 is not null ) "
   
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
     
   adoadodc1.CursorLocation = adUseClient
   '2016/3/11 modify by sonia 瑞婷說發票作廢查詢改發票資料查詢
   'adoadodc1.Open "select A4302 ,A4301 ,A4308 ,A4303 ,NVL(A4201,'') A4201, A4304 , A4305, A4306 From acc430 ,acc420  where A4303 = A4202(+) " & strSql & _
                  "union select A4302 ,A4301 ,A4308 ,A4303 ,NVL(CU04,'') A4201, A4304 , A4305, A4306 From acc430 ,customer where A4303 = CU11(+) " & strSql & _
                  " and NOT EXISTS (select a4202 from acc420,customer where a4202=CU11) " & _
                  strSql, adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2019/5/9 MX14798002會重覆
   'adoadodc1.Open "select A4302 ,A4301 ,A4308 ,A4303 ,a0k04 A4201, A4304 , A4305, A4306 From acc430,acc431,acc0k0 where nvl(a4308,0)=0 and a4301=axc01(+) and axc02=a0k01(+) and a0k01 is not null " & strSql & _
                  "union all select A4302 ,A4301 ,A4308 ,A4303 ,nvl(cu04,a4201) A4201, A4304 , A4305, A4306 From acc430,acc420,customer where nvl(a4308,0)>0 and A4303 = A4202(+) and A4303 = CU11(+) " & strSql & _
                  " order by a4302,a4301", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2019/07/01 + 上傳日期時間
   'Modifyb by Amy 2019/11/25 改顯示欄位順序及收據號碼
'   adoadodc1.Open "select A4302 ,A4301 ,A4308 ,A4303 ,a0k04 A4201, A4304 , A4305, A4306,Decode(a4319||a4320,'111111240000','',sqldatet(a4319)||' '||sqltime(a4320)) a4319,Decode(a4321||a4322,'11111124000000','',sqldatet(a4321)||' '||sqltime(a4322)) a4321 From acc430,acc431,acc0k0 where nvl(a4308,0)=0 and a4301=axc01(+) and axc02=a0k01(+) and a0k01 is not null " & strSql & _
'                  "union all select A4302 ,A4301 ,A4308 ,A4303 ,nvl(cu04,a4201) A4201, A4304 , A4305, A4306,Decode(a4319||a4320,'111111240000','',sqldatet(a4319)||' '||sqltime(a4320)) a4319,Decode(a4321||a4322,'11111124000000','',sqldatet(a4321)||' '||sqltime(a4322)) a4321 From acc430,acc420,customer where nvl(a4308,0)>0 and decode(A4303,'00000000',null,A4303) = A4202(+) and decode(A4303,'00000000',null,A4303) = CU11(+) and '0'=cu02(+) " & strSql & _
'                  " order by a4302,a4301", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2020/07/10 原:Union All ,1090710 ED96429020會出現兩筆
       strQ = "select A4302 ,A4301,A4303 ,a0k04 A4201, A4304,Axc02, A4305, A4306,Decode(a4319||a4320,'111111240000','',sqldatet(a4319)||' '||sqltime(a4320)) a4319,A4308,Decode(a4321||a4322,'11111124000000','',sqldatet(a4321)||' '||sqltime(a4322)) a4321,'' a4324 From acc430,acc431,acc0k0 where nvl(a4308,0)=0 and a4301=axc01(+) and axc02=a0k01(+) and a0k01 is not null " & strSql & _
        "union select A4302 ,A4301,A4303,nvl(cu04,a4201) A4201, A4304,Axc02, A4305, A4306,Decode(a4319||a4320,'111111240000','',sqldatet(a4319)||' '||sqltime(a4320)) a4319,A4308,Decode(a4321||a4322,'11111124000000','',sqldatet(a4321)||' '||sqltime(a4322)) a4321,'' a4324 From acc430,acc420,customer,acc431 where nvl(a4308,0)>0 and a4301=axc01(+) and decode(A4303,'00000000',null,A4303) = A4202(+) and decode(A4303,'00000000',null,A4303) = CU11(+) and '0'=cu02(+) " & strSql
   'Add by Amy 2020/09/17 +銷退查詢(含先給客戶發票但未付款,已申報後轉開),增加折讓/銷退上傳日
   strQ = strQ & _
         "union select a0s03 A4302 ,a0s01 A4301,A4303 ,A0K04 A4201, A4304,Axc02, A4305, A4306,'' a4319,A4308,'' a4321,Decode(A0S28||A0S29,'111111240000','',sqldatet(A0S28)||' '||sqltime(A0S29)) A4324 From Acc0S0,Acc430,Acc431,Acc0K0 Where A0S26=A4301(+) and A0S26 is not null and A4301=Axc01(+) and Axc02=A0K01(+)  and A0S03 >= 1080701 " & Replace(strSql, "a4302", "a0s03") & _
         "union select A4302 ,A4301,A4303 ,a0k04 A4201, A4304,Axc02, A4305, A4306,'' a4319,A4308,'' a4321,Decode(a4324||a4325,'11111124000000','',sqldatet(a4324)||' '||sqltime(a4325)) a4324 From Acc430,Acc431,Acc0k0 Where a4301=axc01(+) And SubStr(axc02,1,9)=a0k01(+) And Nvl(a4310,0)>=1080701" & Replace(strSql, "a4302", "a4310") & _
                  " order by a4302,a4301"
   adoadodc1.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   'end 2019/11/25
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

Private Sub Text1_Validate(Cancel As Boolean)
   Text2 = Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  功能鍵定義
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         If FormCheck Then
            Screen.MousePointer = vbHourglass
            AdodcRefresh
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
'  畫面輸入檢查
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox3.Text <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox4.Text <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

