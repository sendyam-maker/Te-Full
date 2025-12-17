VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2162 
   AutoRedraw      =   -1  'True
   Caption         =   "抵帳單資料選取"
   ClientHeight    =   5110
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5110
   ScaleWidth      =   8760
   Begin VB.TextBox Text12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3168
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   396
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2820
      TabIndex        =   3
      Top             =   240
      Width           =   348
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2580
      TabIndex        =   2
      Top             =   240
      Width           =   240
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   1
      Top             =   240
      Width           =   780
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4644
      MaxLength       =   14
      TabIndex        =   5
      Top             =   240
      Width           =   1572
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1332
      MaxLength       =   3
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1332
      TabIndex        =   11
      Top             =   996
      Width           =   2196
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6816
      TabIndex        =   10
      Top             =   996
      Visible         =   0   'False
      Width           =   1596
   End
   Begin VB.CommandButton Command2 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6636
      TabIndex        =   8
      Top             =   264
      Width           =   816
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7572
      TabIndex        =   9
      Top             =   264
      Width           =   816
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2162.frx":0000
      Height          =   3396
      Left            =   396
      TabIndex        =   7
      Top             =   1440
      Width           =   8052
      _ExtentX        =   14199
      _ExtentY        =   5997
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "cp09"
         Caption         =   "收文號"
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
         DataField       =   "PropertyName"
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
      BeginProperty Column02 
         DataField       =   "cp05"
         Caption         =   "收文日"
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
      BeginProperty Column03 
         DataField       =   "cp27"
         Caption         =   "發文日"
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
         DataField       =   "FagentName"
         Caption         =   "代理人名稱"
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
      BeginProperty Column05 
         DataField       =   "cp44"
         Caption         =   "代理人編號"
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
         Size            =   284
         BeginProperty Column00 
            ColumnWidth     =   1179.78
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1399.748
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   980.221
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   980.221
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   4360.252
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1280.126
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   396
      Top             =   1308
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   547
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
   Begin MSForms.ComboBox Combo1 
      Height          =   345
      Left            =   1332
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   590
      Width           =   7065
      VariousPropertyBits=   679495707
      BackColor       =   16777215
      DisplayStyle    =   3
      Size            =   "12462;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "帳單金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3684
      TabIndex        =   16
      Top             =   252
      Width           =   972
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   15
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "案件名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   372
      TabIndex        =   14
      Top             =   636
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "申請國家"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   372
      TabIndex        =   13
      Top             =   1044
      Width           =   972
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "目前盈虧"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5856
      TabIndex        =   12
      Top             =   1032
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   12
      Top             =   4716
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc2162"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/07 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB、Combo1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
'2006/4/28整理
Option Explicit
Public adocase As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset

Private Sub Combo1_GotFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
OpenIme
End Sub

Private Sub Combo1_LostFocus()
'edit by nickc 2007/06/11  切換輸入法改用API
CloseIme
End Sub

Private Sub Command1_Click()
   strCon9 = ""
   KeyEnter vbKeyEscape
End Sub

Private Sub Command2_Click()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select cp01||cp02||cp03||cp04 as CaseNo, pa26 as CustomerNo, a0k04 from caseprogress, patent, acc0k0 where cp01 = pa01 and cp02 = pa02 and cp03 = pa03 and cp04 = pa04 and cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' union " & _
                 "select cp01||cp02||cp03||cp04 as CaseNo, tm23 as CustomerNo, a0k04 from caseprogress, trademark, acc0k0 where cp01 = tm01 and cp02 = tm02 and cp03 = tm03 and cp04 = tm04 and cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' union " & _
                 "select cp01||cp02||cp03||cp04 as CaseNo, lc11 as CustomerNo, a0k04 from caseprogress, lawcase, acc0k0 where cp01 = lc01 and cp02 = lc02 and cp03 = lc03 and cp04 = lc04 and cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' union " & _
                 "select cp01||cp02||cp03||cp04 as CaseNo, hc05 as CustomerNo, a0k04 from caseprogress, hirecase, acc0k0 where cp01 = hc01 and cp02 = hc02 and cp03 = hc03 and cp04 = hc04 and cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "' union " & _
                 "select cp01||cp02||cp03||cp04 as CaseNo, sp08 as CustomerNo, a0k04 from caseprogress, servicepractice, acc0k0 where cp01 = sp01 and cp02 = sp02 and cp03 = sp03 and cp04 = sp04 and cp60 = a0k01 (+) and cp09 = '" & Adodc1.Recordset.Fields("cp09").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      adoTaie.Execute "delete from acc161 where axg01 = '" & strCon8 & "' and axg02 = '" & Adodc1.Recordset.Fields("cp09").Value & "'"
      
      If IsNull(adoquery.Fields("CustomerNo").Value) = False Then
         '2014/3/18 modify by sonia axg12限100bytes
         'strCon9 = "insert into acc161 (axg01, axg02, axg03, axg04, axg05, axg06, axg07, axg08, axg12, axg13) values ('" & strCon8 & "', '" & Adodc1.Recordset.Fields("cp09").Value & "', '" & adoquery.Fields("CaseNo").Value & "', " & Val(Text11) & ", '" & adoquery.Fields("CustomerNo").Value & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & Mid(Combo1, 4, Len(Combo1)) & "', '" & adoquery.Fields("a0k04").Value & "')"
         'Modified by Morgan 2024/10/7 修正資料有單引號會發生錯誤問題
         'strCon9 = "insert into acc161 (axg01, axg02, axg03, axg04, axg05, axg06, axg07, axg08, axg12, axg13) values ('" & strCon8 & "', '" & Adodc1.Recordset.Fields("cp09").Value & "', '" & adoquery.Fields("CaseNo").Value & "', " & Val(Text11) & ", '" & adoquery.Fields("CustomerNo").Value & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & PUB_StrToStr_byVal(Mid(Combo1, 4, Len(Combo1)), 100) & "', '" & adoquery.Fields("a0k04").Value & "')"
         'Modified by Morgan 2024/10/15 修正收據抬頭Null的錯誤(FMP案抵帳單)
         strCon9 = "insert into acc161 (axg01, axg02, axg03, axg04, axg05, axg06, axg07, axg08, axg12, axg13) values ('" & strCon8 & "', '" & Adodc1.Recordset.Fields("cp09").Value & "', '" & adoquery.Fields("CaseNo").Value & "', " & Val(Text11) & ", '" & adoquery.Fields("CustomerNo").Value & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & ChgSQL(PUB_StrToStr_byVal(Mid(Combo1, 4, Len(Combo1)), 100)) & "', '" & ChgSQL("" & adoquery.Fields("a0k04").Value) & "')"
      Else
         '2014/3/18 modify by sonia axg12限100bytes
         'strCon9 = "insert into acc161 (axg01, axg02, axg03, axg04, axg05, axg06, axg07, axg08, axg12, axg13) values ('" & strCon8 & "', '" & Adodc1.Recordset.Fields("cp09").Value & "', '" & adoquery.Fields("CaseNo").Value & "', " & Val(Text11) & ", '" & adoquery.Fields("CustomerNo").Value & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & Mid(Combo1, 4, Len(Combo1)) & "', '" & adoquery.Fields("a0k04").Value & "')"
         'Modified by Morgan 2024/10/7 修正資料有單引號會發生錯誤問題
         'strCon9 = "insert into acc161 (axg01, axg02, axg03, axg04, axg05, axg06, axg07, axg08, axg12, axg13) values ('" & strCon8 & "', '" & Adodc1.Recordset.Fields("cp09").Value & "', '" & adoquery.Fields("CaseNo").Value & "', " & Val(Text11) & ", '" & adoquery.Fields("CustomerNo").Value & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & PUB_StrToStr_byVal(Mid(Combo1, 4, Len(Combo1)), 100) & "', '" & adoquery.Fields("a0k04").Value & "')"
         'Modified by Morgan 2024/10/15 修正收據抬頭Null的錯誤(FMP案抵帳單)
         strCon9 = "insert into acc161 (axg01, axg02, axg03, axg04, axg05, axg06, axg07, axg08, axg12, axg13) values ('" & strCon8 & "', '" & Adodc1.Recordset.Fields("cp09").Value & "', '" & adoquery.Fields("CaseNo").Value & "', " & Val(Text11) & ", '" & adoquery.Fields("CustomerNo").Value & "', " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "', '" & ChgSQL(PUB_StrToStr_byVal(Mid(Combo1, 4, Len(Combo1)), 100)) & "', '" & ChgSQL("" & adoquery.Fields("a0k04").Value) & "')"
      End If
   Else
      strCon9 = ""
   End If
   adoquery.Close
   If IsNull(Adodc1.Recordset.Fields("cp44").Value) = False Then
      strCustNo = Adodc1.Recordset.Fields("cp44").Value
   End If
   Unload Me
   tool2_enabled
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2021/12/07 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 8850
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
   strFormName = Name
   PUB_InitForm Me, 8850, 5500, strBackPicPath1
   'end 2021/12/07
   
   Text10 = strCon2
   If Text10 = "TF" Then
      Text12.Visible = True
   Else
      Text12.Visible = False
   End If
   Text5 = strCon3
   Text7 = strCon4
   Text9 = strCon5
   Text12 = strCon6
   Text11 = strCon7
   FormShow
   AdodcRefresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strItemNo = ""
   tool3_enabled
   Frmacc2160.Enabled = True
   Frmacc2160.Show
   Set Frmacc2162 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
On Error GoTo Checking
   If adocase.State = adStateOpen Then
      adocase.Close
   End If
   adocase.CursorLocation = adUseClient
   Select Case Text10
      Case "TF"
         '2006/4/28 MODIFY BY SONIA 加判斷是否有輸入代理人編號
         'adocase.Open "select cp09, nvl(cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp44 = '" & strCustNo & "'", adoTaie, adOpenStatic, adLockReadOnly
         If strCustNo <> "" Then
            adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "' and cp44 = '" & strCustNo & "'", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & Text7 & "' and cp03 = '" & Text9 & "' and cp04 = '" & Text12 & "'", adoTaie, adOpenStatic, adLockReadOnly
         End If
         '2006/4/28 END
      Case Else
         '2006/4/28 MODIFY BY SONIA 加判斷是否有輸入代理人編號
         'adocase.Open "select cp09, nvl(cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp44 = '" & strCustNo & "'", adoTaie, adOpenStatic, adLockReadOnly
         If strCustNo <> "" Then
            adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "' and cp44 = '" & strCustNo & "'", adoTaie, adOpenStatic, adLockReadOnly
         Else
            adocase.Open "select cp09, decode(cpm01, 'FCP', cpm03, 'FCT', cpm03, 'FCL', cpm03, 'LIN', cpm03, 'ACS', cpm03, cpm04) as PropertyName, nvl(cp05 - 19110000, 0) as cp05, nvl(cp27 - 19110000, 0) as cp27, cp44, nvl(fa05||fa63||fa64||fa65, nvl(fa04, fa06)) as FagentName, cp18 from caseprogress, casepropertymap, fagent where cp01 = cpm01 and cp10 = cpm02 and substr(cp44, 1, 8) = fa01 (+) and substr(cp44, 9, 1) = fa02 (+) and cp01 = '" & Text10 & "' and cp02 = '" & Text5 & "' and cp03 = '" & Text7 & "' and cp04 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
         End If
         '2006/4/28 END
   End Select
   Set Adodc1.Recordset = adocase
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   KeyEnter KeyCode
End Sub

'*************************************************
'  顯示畫面
'
'*************************************************
Public Sub FormShow()
   Combo1.Clear
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select pa05 as Name1, pa06 as Name2, pa07 as Name3, nvl(na03, na04) as NationName from patent, nation where pa09 = na01 (+) and pa01 = '" & Text10 & "' and pa02 = '" & Text5 & "' and pa03 = '" & Text7 & "' and pa04 = '" & Text9 & "' union " & _
                 "select tm05 as Name1, tm06 as Name2, tm07 as Name3, nvl(na03, na04) as NationName from trademark, nation where tm10 = na01 (+) and tm01 = '" & Text10 & "' and tm02 = '" & Text5 & "' and tm03 = '" & Text7 & "' and tm04 = '" & Text9 & "' and tm01 <> 'TF' union " & _
                 "select tm05 as Name1, tm06 as Name2, tm07 as Name3, nvl(na03, na04) as NationName from trademark, nation where tm10 = na01 (+) and tm01 = '" & Text10 & "' and tm02 = '" & Text5 & Text7 & "' and tm03 = '" & Text9 & "' and tm04 = '" & Text12 & "' and tm01 = 'TF' union " & _
                 "select lc05 as Name1, lc06 as Name2, lc07 as Name3, nvl(na03, na04) as NationName from lawcase, nation where lc15 = na01 (+) and lc01 = '" & Text10 & "' and lc02 = '" & Text5 & "' and lc03 = '" & Text7 & "' and lc04 = '" & Text9 & "' union " & _
                 "select hc06 as Name1, '' as Name2, '' as Name3, '' as NationName from hirecase where hc01 = '" & Text10 & "' and hc02 = '" & Text5 & "' and hc03 = '" & Text7 & "' and hc04 = '" & Text9 & "' union " & _
                 "select sp05 as Name1, sp06 as Name2, sp07 as Name3, nvl(na03, na04) as NationName from servicepractice, nation where sp09 = na01 (+) and sp01 = '" & Text10 & "' and sp02 = '" & Text5 & "' and sp03 = '" & Text7 & "' and sp04 = '" & Text9 & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields("Name1").Value) = False Then
         Combo1 = "中--" & adoquery.Fields("Name1").Value
         Combo1.AddItem "中--" & adoquery.Fields("Name1").Value
      End If
      If IsNull(adoquery.Fields("Name2").Value) = False Then
         Combo1.AddItem "英--" & adoquery.Fields("Name2").Value
      End If
      If IsNull(adoquery.Fields("Name3").Value) = False Then
         Combo1.AddItem "日--" & adoquery.Fields("Name3").Value
      End If
      If IsNull(adoquery.Fields("NationName").Value) = False Then
         Text1 = adoquery.Fields("NationName").Value
      Else
         Text1 = MsgText(601)
      End If
   Else
      Text1 = MsgText(601)
   End If
   adoquery.Close
End Sub
