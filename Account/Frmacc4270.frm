VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc4270 
   AutoRedraw      =   -1  'True
   Caption         =   "智權人員結餘點數查詢"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   9405
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1290
      TabIndex        =   0
      Top             =   180
      Width           =   3050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "產生Excel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      TabIndex        =   6
      Top             =   210
      Width           =   1350
   End
   Begin VB.TextBox txtSalesNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1305
      MaxLength       =   5
      TabIndex        =   4
      Top             =   960
      Width           =   1305
   End
   Begin VB.TextBox Text1 
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
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "報表列印"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   5
      Top             =   210
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4270.frx":0000
      Height          =   3480
      Left            =   240
      TabIndex        =   7
      Top             =   1470
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   6138
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "智權人員結餘點數資料"
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Name"
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
      BeginProperty Column01 
         DataField       =   "T1"
         Caption         =   "大陸商標"
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
      BeginProperty Column02 
         DataField       =   "P1"
         Caption         =   "大陸專利"
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
      BeginProperty Column03 
         DataField       =   "T2"
         Caption         =   "國外商標"
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
         DataField       =   "P2"
         Caption         =   "國外專利"
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
         DataField       =   "CFL"
         Caption         =   "ＣＦＬ"
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
         DataField       =   "TOTAL"
         Caption         =   "合計"
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
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1454.74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   225
      Top             =   1365
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
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
      Left            =   1296
      TabIndex        =   1
      Top             =   588
      Width           =   1332
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
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
      Left            =   3000
      TabIndex        =   2
      Top             =   588
      Width           =   1332
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "查詢當日結餘放出須等批次後才能查詢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   5280
      TabIndex        =   14
      Top             =   1080
      Width           =   3825
   End
   Begin MSForms.Label lblSalesName 
      Height          =   285
      Left            =   2790
      TabIndex        =   13
      Top             =   960
      Width           =   2000
      VariousPropertyBits=   19
      Size            =   "3528;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "員工代號"
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
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "資料別           (1放出 2全部 3餘額 4空白)    "
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
      Left            =   4560
      TabIndex        =   11
      Top             =   600
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4620
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   10
      Top             =   228
      Width           =   732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "傳票日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   9
      Top             =   588
      Width           =   972
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2760
      TabIndex        =   8
      Top             =   588
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1245
      Left            =   240
      Top             =   105
      Width           =   9000
   End
End
Attribute VB_Name = "Frmacc4270"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/05 Form2.0已修改 lblSalesName/DataGrid1
'Memo by Amy 2013/11/05 與frmacc41f0共用accrpt423 Table, 隱藏報表列鈕
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoadodc1 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Dim dllaccrpt423 As Object
Dim strSql As String

'Add by Amy 2020/04/14
Private Sub Combo1_GotFocus()
    TextInverse Combo1
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
    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label3 & MsgText(63), , MsgText(5)
        Cancel = True
        Combo1.SetFocus
        Exit Sub
    ElseIf Len(Trim(Combo1)) = 1 Then
        Combo1 = Trim(strCmp) & "　" & A0802Query(strCmp, True)
    End If
End Sub
'end 2020/04/14

'Added by Morgan 2011/11/14
Private Sub Command1_Click()
   Export2Excel adoadodc1
End Sub

'Private Sub Command2_Click()
''92.11.20 ADD BY SONIA
'Dim strTmp As String
'
'   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
'      Exit Sub
'   End If
'   Screen.MousePointer = vbHourglass
'   'ProduceData
'   '92.11.20 add by sonia
'   strTmp = ""
'   Select Case Text1
'      Case "1"
'         strTmp = "放出"
'      Case "2"
'         strTmp = "全部"
'      Case "4"
'         strTmp = "空白"
'      Case Else
'         strTmp = "餘額"
'   End Select
'   '92.11.20 end
'   'Modify by Amy 2013/11/05 因與frmacc41f0共用accrpt423 Table 所以需傳入strUserNum
'   dllaccrpt423.Acc4270 ReportTitle(423), Text4, Text5, MaskEdBox1.Text, MaskEdBox2.Text, StaffQuery(strUserNum), strTmp, CFDate(strSrvDate(2))
'   Screen.MousePointer = vbDefault
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
'End Sub

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
   Me.Width = 9500
   Me.Height = 5500
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   'Modify by Amy 2020/04/14 公司別改下拉
   'Text4 = "1" 'Add by Amy 2017/01/11 預帶1公司-瑞婷
   Combo1.AddItem "", 0
   Call Pub_SetCboCmp(Combo1, True, True, False, "1", 1)
   'end 2020/04/14
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   MaskEdBox1.Text = "088/10/01"
   
   OpenTable
'   Text4 = "1"
'   Text4.Enabled = False 'Added by Morgan 2011/11/14 目前只有1公司,且程式也沒有過濾
   Text1 = "3"
'   Command2.Enabled = False
'   Command1.Enabled = Command2.Enabled 'Added by Morgan 2011/11/14
   Command1.Enabled = False
   Set dllaccrpt423 = CreateObject("AccReport.ReportSelect")
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set dllaccrpt423 = Nothing
   Set Frmacc4270 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
Dim strCmp As String 'Add by Amy 2020/04/14

On Error GoTo Checking
   'Modif by Amy 2020/04/14 公司別改下拉 原: IIf(Text4 = "2", "J", IIf(Text4 = "1", "1", ""))
   If Trim(Combo1) <> MsgText(601) Then
      strCmp = Combo1
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
   End If
   adoadodc1.CursorLocation = adUseClient
'   Select Case strAccount
'      Case "2"
'         adoadodc1.Open "select ax301 as ax201, ax302 as ax202, a0305 as a0205, ax303 as ax203, ax306 as ax206, ax307 as ax207, ax304 as ax204, ax312 as ax212, a0102 from acc031, acc010, acc030 where ax305 = a0101 and ax301 = a0301 and ax302 = a0302 and ax301 = '" & IIf(Text4 = "2", "J", IIf(Text4 = "1", "1", "")) & "' and a0305 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0305 <= " & Val(FCDate(MaskEdBox2.Text)) & " order by a0305 desc, ax302 asc, ax303 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Case Else
         adoadodc1.Open "select * from acc021, acc010, acc020 where acc021.ax205 = acc010.a0101 and acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and ax201 = '" & strCmp & "' and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & " order by a0205 desc, ax202 asc, ax203 asc", adoTaie, adOpenStatic, adLockReadOnly
'   End Select
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  查詢資料表(傳票資料)
'  Memo 2017/11/01 原未讀暫存資料,因增加目標欄位(R42311)故改抓暫存資料
'*************************************************
Public Sub QueryTable()
    Dim bolCancel As Boolean
    Dim strQ As String
    Dim StrSQLa As String, StrSqlB As String
    Dim strCmp As String 'Add by Amy 2020/04/14
    
    Call txtSalesNo_Validate(bolCancel)
    If bolCancel = True Then Exit Sub
    
On Error GoTo Checking
    strSql = ""
    'Add by Amy 2020/04/14 公司別改下拉
    If Trim(Combo1) <> MsgText(601) Then
        strCmp = Combo1
        If InStr(strCmp, "　") > 0 Then
              strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    If adoadodc1.State = adStateOpen Then
        adoadodc1.Close
    End If
    adoadodc1.CursorLocation = adUseClient
    
    StrSqlB = ""
    Select Case Text1
        Case "1"
            StrSqlB = "ax206"
        Case "2"
            StrSqlB = "ax207"
        Case Else
            StrSqlB = "ax207-ax206"
    End Select
 
    
    '查詢'1放出'時,剔除整張傳票都沒有收入科目的傳票
    If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
        strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
        If Text1 = "1" Then
            strQ = strQ & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text))
        End If
    End If
    If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
        strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
        If Text1 = "1" Then
            strQ = strQ & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text))
        End If
    End If
         
    '智權人員條件
    If txtSalesNo <> "" Then
        strSql = strSql & " and ax209='" & txtSalesNo.Text & "'"
        If Text1 = "1" Then
            strQ = strQ & " and ax209='" & txtSalesNo.Text & "'"
        End If
    End If
    
    'Modify by Amy 2020/04/14 公司別改下拉 原每句 IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", ""))
    If strCmp <> MsgText(601) Then
        If InStr(strCmp, "+") = 0 Then
            strSql = strSql & " And a0201='" & strCmp & "' "
        Else
            strSql = strSql & " And a0201 In ('" & Replace(strCmp, "+", "','") & "') "
        End If
    End If
    StrSQLa = " Having (Nvl(sum(decode(ax205, '249101', " & StrSqlB & " , 0)),0)<>0 Or Nvl(sum(decode(ax205, '249102', " & StrSqlB & " , 0)),0)<>0 Or Nvl(sum(decode(ax205, '249103', " & StrSqlB & " , 0)),0)<>0 Or Nvl(sum(decode(ax205, '249104', " & StrSqlB & " , 0)),0)<>0 Or Nvl(sum(decode(ax205, '249105', " & StrSqlB & " , 0)),0)<>0) "
    If Text1 = "4" Then
        strSql = "select st15 as Dept, st01 as ID, st02 as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, staff where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st15, st01, st02" & StrSQLa & _
                " union select a0901||'T' as Dept, a0901||'T' as ID, a0902 as Name, '' as T1, '' as P1, '' as T2, '' as CFL,'' as P2, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by a0901, a0902" & StrSQLa & _
                " union select 'S1T' as Dept, 'S1T' as ID, '北所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S1' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'S2T' as Dept, 'S2T' as ID, '中所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S2' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'S3T' as Dept, 'S3T' as ID, '南所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S3' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'S4T' as Dept, 'S4T' as ID, '高所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S4' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'S9T' as Dept, 'S9T' as ID, '廣東' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S9' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'X'||st15 as Dept, st01 as ID, st02 as Name, '' as T1, '' as P1, '' as T2, '' as P2,'' as CFL, '' as TOTAL from acc021, acc020, staff where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st15,st01, st02" & StrSQLa & _
                " union select 'XZT' as Dept, 'X0T' as ID, '其他合計' as Name, '' as T1, '' as P1, '' as T2, '' as P2,'' as CFL, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'ZZZ' as Dept, 'ZZZ' as ID, '全所合計' as Name, '' as T1, '' as P1, '' as T2, '' as P2,'' as CFL, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & " ORDER BY DEPT,ID"
    '查詢'1放出'時,剔除整張傳票都沒有收入科目的傳票ex:D106092547
    ElseIf Text1 = "1" Then
        cnnConnection.Execute "Delete From Accrpt4270 Where ID='" & strUserNum & "'"
        cnnConnection.Execute "Delete From Accrpt4270_1 Where ID='" & strUserNum & "'"
        
        '收入科目的傳票
        If strCmp <> MsgText(601) Then
            If InStr(strCmp, "+") = 0 Then
                strQ = " And a0201='" & strCmp & "' "
            Else
                strQ = " And a0201 In ('" & Replace(strCmp, "+", "','") & "') "
            End If
        End If
        strQ = "Insert into Accrpt4270_1 (ID,R001,R002) " & _
                "Select Distinct '" & strUserNum & "',ax201,ax202 From acc020,acc021 Where a0201=ax201(+) And a0202=ax202(+) " & _
                " And substr(ax205,1,1)='4' And ax206+ax207>0 " & strQ
        cnnConnection.Execute strQ
                    
        '原始資料
        strQ = "Insert into Accrpt4270 (ID,R001,R002,R003,R004,R005) " & _
                "Select '" & strUserNum & "',ax201,ax202,ax205,Sum(" & StrSqlB & "),ax209 From Acc020,Acc021 " & _
                "Where ax201(+) = a0201 and ax202(+) = a0202 " & _
                " and ax205 in ('249101', '249102', '249103', '249104', '249105') " & strSql & " " & _
                "Group by ax201,ax202,ax205,ax209 " & StrSQLa
        cnnConnection.Execute strQ
            
        '刪除無收入科目的傳票
        strQ = "Delete From Accrpt4270 Where ID='" & strUserNum & "' And (r001,r002) Not In (Select R001,R002 From Accrpt4270_1 Where ID='" & strUserNum & "')"
        cnnConnection.Execute strQ
        
        strSql = "select st15 as Dept, st01 as ID, st02 as Name, sum(decode(R003, '249101', R004 , 0)) as T1, sum(decode(R003, '249102', R004 , 0)) as P1, sum(decode(R003, '249103', R004 , 0)) as T2, sum(decode(R003, '249104', R004 , 0)) as P2, sum(decode(R003, '249105', R004 , 0)) as CFL, sum(R004 ) as TOTAL From Accrpt4270,staff Where ID='" & strUserNum & "' and R005=st01(+) and substr(st15,1,1)='S' Group by st15, st01, st02" & _
                 " Union select a0901||'T' as Dept, a0901||'T' as ID, a0902 as Name, sum(decode(R003, '249101', R004 , 0)) as T1, sum(decode(R003, '249102', R004 , 0)) as P1, sum(decode(R003, '249103', R004 , 0)) as T2, sum(decode(R003, '249104', R004 , 0)) as P2, sum(decode(R003, '249105', R004 , 0)) as CFL, sum(R004 ) as TOTAL From Accrpt4270,(select * from staff, acc090 where st15 = a0901) new Where ID='" & strUserNum & "' and R005=st01(+) and substr(st15,1,1)='S' Group by a0901, a0902" & _
                 " Union select 'S1T' as Dept, 'S1T' as ID, '北所' as Name, sum(decode(R003, '249101', R004 , 0)) as T1, sum(decode(R003, '249102', R004 , 0)) as P1, sum(decode(R003, '249103', R004 , 0)) as T2, sum(decode(R003, '249104', R004 , 0)) as P2, sum(decode(R003, '249105', R004 , 0)) as CFL, sum(R004 ) as TOTAL From Accrpt4270,(select * from staff, acc090 where st15 = a0901) new Where ID='" & strUserNum & "' and R005=st01 (+) and substr(st15,1,2)='S1' " & _
                 " Union select 'S2T' as Dept, 'S2T' as ID, '中所' as Name, sum(decode(R003, '249101', R004 , 0)) as T1, sum(decode(R003, '249102', R004 , 0)) as P1, sum(decode(R003, '249103', R004 , 0)) as T2, sum(decode(R003, '249104', R004 , 0)) as P2, sum(decode(R003, '249105', R004 , 0)) as CFL, sum(R004 ) as TOTAL From Accrpt4270,(select * from staff, acc090 where st15 = a0901) new Where ID='" & strUserNum & "' and R005=st01 (+) and substr(st15,1,2)='S2' " & _
                 " Union select 'S3T' as Dept, 'S3T' as ID, '南所' as Name, sum(decode(R003, '249101', R004 , 0)) as T1, sum(decode(R003, '249102', R004 , 0)) as P1, sum(decode(R003, '249103', R004 , 0)) as T2, sum(decode(R003, '249104', R004 , 0)) as P2, sum(decode(R003, '249105', R004 , 0)) as CFL, sum(R004 ) as TOTAL From Accrpt4270,(select * from staff, acc090 where st15 = a0901) new Where ID='" & strUserNum & "' and R005=st01 (+) and substr(st15,1,2)='S3' " & _
                 " Union select 'S4T' as Dept, 'S4T' as ID, '高所' as Name, sum(decode(R003, '249101', R004 , 0)) as T1, sum(decode(R003, '249102', R004 , 0)) as P1, sum(decode(R003, '249103', R004 , 0)) as T2, sum(decode(R003, '249104', R004 , 0)) as P2, sum(decode(R003, '249105', R004 , 0)) as CFL, sum(R004 ) as TOTAL From Accrpt4270,(select * from staff, acc090 where st15 = a0901) new Where ID='" & strUserNum & "' and R005=st01 (+) and substr(st15,1,2)='S4' " & _
                 " Union select 'S9T' as Dept, 'S9T' as ID, '廣東' as Name, sum(decode(R003, '249101', R004 , 0)) as T1, sum(decode(R003, '249102', R004 , 0)) as P1, sum(decode(R003, '249103', R004 , 0)) as T2, sum(decode(R003, '249104', R004 , 0)) as P2, sum(decode(R003, '249105', R004 , 0)) as CFL, sum(R004 ) as TOTAL From Accrpt4270,(select * from staff, acc090 where st15 = a0901) new Where ID='" & strUserNum & "' and R005=st01 (+) and substr(st15,1,2)='S9' " & _
                 " Union select 'X'||st15 as Dept, st01 as ID, st02 as Name, sum(decode(R003, '249101', R004 , 0)) as T1, sum(decode(R003, '249102', R004 , 0)) as P1, sum(decode(R003, '249103', R004 , 0)) as T2, sum(decode(R003, '249104', R004 , 0)) as P2, sum(decode(R003, '249105', R004 , 0)) as CFL, sum(R004 ) as TOTAL From Accrpt4270,staff Where ID='" & strUserNum & "' and R005=st01(+) and substr(st15,1,1)<>'S' Group by st15,st01, st02" & _
                 " Union select 'XZT' as Dept, 'X0T' as ID, '其他合計' as Name, sum(decode(R003, '249101', R004 , 0)) as T1, sum(decode(R003, '249102', R004 , 0)) as P1, sum(decode(R003, '249103', R004 , 0)) as T2, sum(decode(R003, '249104', R004 , 0)) as P2, sum(decode(R003, '249105', R004 , 0)) as CFL, sum(R004 ) as TOTAL From Accrpt4270,(select * from staff, acc090 where st15 = a0901) new Where ID='" & strUserNum & "' and R005=st01(+) and substr(st15,1,1)<>'S' " & _
                 " Union select 'ZZZ' as Dept, 'ZZZ' as ID, '全所合計' as Name, sum(decode(R003, '249101', R004 , 0)) as T1, sum(decode(R003, '249102', R004 , 0)) as P1, sum(decode(R003, '249103', R004 , 0)) as T2, sum(decode(R003, '249104', R004 , 0)) as P2, sum(decode(R003, '249105', R004 , 0)) as CFL, sum(R004 ) as TOTAL From Accrpt4270,(select * from staff, acc090 where st15 = a0901) new Where ID='" & strUserNum & "' and R005=st01(+) "

      
    Else
        strSql = "select st15 as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, staff where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st15, st01, st02" & StrSQLa & _
                " union select a0901||'T' as Dept, a0901||'T' as ID, a0902 as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by a0901, a0902" & StrSQLa & _
                " union select 'S1T' as Dept, 'S1T' as ID, '北所' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S1' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'S2T' as Dept, 'S2T' as ID, '中所' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S2' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'S3T' as Dept, 'S3T' as ID, '南所' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S3' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'S4T' as Dept, 'S4T' as ID, '高所' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S4' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'S9T' as Dept, 'S9T' as ID, '廣東' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S9' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'X'||st15 as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, staff where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st15,st01, st02" & StrSQLa & _
                " union select 'XZT' as Dept, 'X0T' as ID, '其他合計' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
                " union select 'ZZZ' as Dept, 'ZZZ' as ID, '全所合計' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201 and ax202(+) = a0202 and ax209 = st01 (+) and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & " ORDER BY DEPT,ID"
    End If
    'end 2020/04/14
    adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
    adoTaie.Execute "Delete from accrpt423 Where R42301='" & strUserNum & "' "
    Adodc1.Recordset.Requery
    If Adodc1.Recordset.RecordCount = 0 Then
       Adodc1.Recordset.Close
       Command1.Enabled = False
       MsgBox MsgText(28), , MsgText(5)
       Exit Sub
    Else
        Command1.Enabled = True
        adoadodc1.MoveFirst
        Do While adoadodc1.EOF = False
            adoTaie.Execute "insert into accrpt423 (R42301,R42302,R42303,R42304,R42305,R42306,R42307,R42308,R42309,R42310) values ('" & strUserNum & "', '" & adoadodc1.Fields("Dept").Value & "', '" & adoadodc1.Fields("Name").Value & "', " & _
            IIf(IsNull(adoadodc1.Fields("T1").Value), "Null", adoadodc1.Fields("T1").Value) & ", " & IIf(IsNull(adoadodc1.Fields("P1").Value), "Null", adoadodc1.Fields("P1").Value) & ", " & IIf(IsNull(adoadodc1.Fields("T2").Value), "Null", adoadodc1.Fields("T2").Value) & ", " & IIf(IsNull(adoadodc1.Fields("P2").Value), "Null", adoadodc1.Fields("P2").Value) & ", " & IIf(IsNull(adoadodc1.Fields("CFL").Value), "Null", adoadodc1.Fields("CFL").Value) & ", " & IIf(IsNull(adoadodc1.Fields("TOTAL").Value), "Null", adoadodc1.Fields("TOTAL").Value) & _
            ", '" & adoadodc1.Fields("ID").Value & "' )"
            adoadodc1.MoveNext
        Loop
        '目標欄位(抓系統當月)
        strSql = "Update Accrpt423 Set R42311=(" & _
                   "Select sum(nvl(PE04,0))*1000 PE04 From PerFormance " & _
                   "Where PE01=R42310 And PE02='TOT' And PE03= " & Val(Left(strSrvDate(1), 6)) & " Group by PE01) " & _
                 "Where R42301='" & strUserNum & "' And R42310 is not null And InStr(R42302,'T')=0 And R42302<>'ZZZ'"
        adoTaie.Execute strSql
        
        'Add by Amy 2018/01/11 刪除目票及合計是0 ex:廣東不出現
        strSql = "Delete From Accrpt423 Where R42301='" & strUserNum & "' And Nvl(R42309,0)=0 And Nvl(R42311,0)=0 "
        adoTaie.Execute strSql
        'end 2018/01/11
        
        '抓取暫存資料
        strSql = "Select R42302 as Dept,R42310 as ID,R42303 as Name,R42304 as T1,R42305 as P1,R42306 as T2,R42307 as P2,R42308 as CFL,R42309 as TOTAL,R42311 as PE04 " & _
                 "From Accrpt423 Where R42301='" & strUserNum & "' Order by R42302,R42310"
        If adoadodc1.State = adStateOpen Then adoadodc1.Close
        adoadodc1.CursorLocation = adUseClient
        adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
        Adodc1.Recordset.Requery
    End If
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

'Private Sub Text4_Change()
'   If Text4 = MsgText(601) Then
'      Exit Sub
'   End If
'   Text5 = A0802Query(Text4)
'End Sub

'Mark by Amy 2020/04/14 公司別改下拉
'Private Sub Text4_GotFocus()
'   TextInverse Text4
'End Sub
'
''Add By Sindy 2014/1/23
'Private Sub Text4_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
   Dim HasShowMsg As Boolean 'Add by Amy 2020/04/14
   
   Select Case KeyCode
      Case vbKeyF12
         'Modify by Amy 2020/04/14
         If FormCheck(HasShowMsg) Then
            Screen.MousePointer = vbHourglass
            QueryTable
'            Command2.Enabled = True
'            Command1.Enabled = Command2.Enabled 'Added by Morgan 2011/11/14
            Screen.MousePointer = vbDefault
            Exit Sub
         ElseIf HasShowMsg = False Then
         'end 2020/04/14
            MsgBox MsgText(181), , MsgText(5)
         End If
         
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
'Modify by Amy 2020/04/14 +HasShowMsg
Public Function FormCheck(HasShowMsg As Boolean) As Boolean
   Dim bCancel As Boolean 'Modify by Amy 2020/04/14
   
   'Modify by Amy 2020/04/14 公司別改下拉  原:Text4
   If Trim(Combo1) <> MsgText(601) Then
      Call Combo1_Validate(bCancel)
      If bCancel = True Then
        HasShowMsg = True
      Else
        FormCheck = True
      End If
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

''*************************************************
''  將查詢結果儲存於暫存檔中(ACCRPT423)
''
''*************************************************
'Public Sub ProduceData()
'   adoTaie.Execute "delete from accrpt423"
'   If adoquery.State = adStateOpen Then
'      adoquery.Close
'   End If
'   adoquery.CursorLocation = adUseClient
'   adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   Do While adoquery.EOF = False
'      adoTaie.Execute "insert into accrpt423 values ('" & strUserNum & "', '" & adoquery.Fields("Dept").Value & "', '" & adoquery.Fields("Name").Value & "', " & _
'                      IIf(IsNull(adoquery.Fields("T1").Value), "Null", adoquery.Fields("T1").Value) & ", " & IIf(IsNull(adoquery.Fields("P1").Value), "Null", adoquery.Fields("P1").Value) & ", " & IIf(IsNull(adoquery.Fields("T2").Value), "Null", adoquery.Fields("T2").Value) & ", " & IIf(IsNull(adoquery.Fields("P2").Value), "Null", adoquery.Fields("P2").Value) & ", " & IIf(IsNull(adoquery.Fields("TOTAL").Value), "Null", adoquery.Fields("TOTAL").Value) & ")"
'      adoquery.MoveNext
'   Loop
'   adoquery.Close
'End Sub

Private Sub txtSalesNo_GotFocus()
   TextInverse txtSalesNo
End Sub

Private Sub txtSalesNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesNo_Validate(Cancel As Boolean)
   If txtSalesNo = "" Then
      lblSalesName = ""
   Else
      lblSalesName = GetStaffName(txtSalesNo, True)
      If lblSalesName = "" Then
         MsgBox "員工代碼輸入錯誤！", vbCritical
         Cancel = True
      End If
   End If
End Sub

'Add by Morgan 2011/11/14
'*************************************************
'  轉成Excel檔案
'
'*************************************************
Private Sub Export2Excel(adoRst As ADODB.Recordset)
   Dim xlsSalesPoint As New Excel.Application
   Dim wksaccrpt418 As New Worksheet
   Dim xlsSelect As Selection
   Dim strFileName As String
   Dim iRow As Integer, iRowCount As Integer
   Dim stCode1 As String, stCode2 As String
   Dim stRptName As String, stCellID As String, stCellFormat As String
   Dim bolShowZezo As Boolean, stCode1Name As String
   Dim Rc As String '欄位座標
   Dim MaxCol As String '最右的欄位代碼
   Dim strTmp As String
   Dim strAreaFrom As String, strBranchFormula As String, strTotalFormula As String, strLastDept As String
   'Add by Amy 2014/12/16 '+隱藏列
   Dim strHiddenRow As String, ii As Integer
   Dim arrRow
   'Add by Amy 2020/04/14
   Dim strCmp As String
   Dim arrCmp
   
On Error GoTo ErrHnd

   MaxCol = Chr(Asc("a") + 10)
   
   stCellFormat = "#,##0.00 ;[紅色]-#,##0.00 "
   
   stRptName = ReportTitle(423)
   
   strFileName = strExcelPath & Trim(Replace(stRptName, "*", "")) & Format(Now, "yyyymmddhhmmss") & MsgText(43)
   If Dir(strFileName) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFileName
   End If

   xlsSalesPoint.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   Set wksaccrpt418 = xlsSalesPoint.Worksheets(1)
   With wksaccrpt418
      iRow = 1
      .Range("a" & iRow).Value = stRptName
      Rc = MaxCol & iRow
      With .Range("a" & iRow & ":" & Rc)
         .Font.Size = 18
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
         .MergeCells = True
      End With
      
      iRow = iRow + 2
      .Range("c" & iRow).Value = "公司別："
      With .Range("c" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlRight
      End With

      '.Range("d" & iRow).Value = Text4 & "  " & Text5
      'Modify by Amy 2020/04/14 公司名稱改抓function
      '.Range("d" & iRow).Value = IIf(Text4 = "2", "智權", IIf(Text4 = "1", "台一", "台一　專利商標/智權"))
      strCmp = Combo1
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
      strCmp = GetAccReportCmpN(strCmp, True, True) 'Add by Amy 2020804/16
      .Range("d" & iRow).Value = strCmp
      'end 2020/04/14
      With .Range("d" & iRow & ":f" & iRow)
         .Font.Size = 12
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
         .MergeCells = True
      End With
      
      iRow = iRow + 1
      .Range("c" & iRow).Value = "傳票日期："
      With .Range("c" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlRight
      End With

      .Range("d" & iRow).Value = MaskEdBox1 & " ~ " & MaskEdBox2
      With .Range("d" & iRow & ":f" & iRow)
         .Font.Size = 12
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
         .MergeCells = True
      End With
      
      iRow = iRow + 1
      
      .Range("a" & iRow).Value = "列印人員："
      With .Range("a" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlRight
      End With

       strTmp = StaffQuery(strUserNum)
      .Range("b" & iRow).Value = strTmp
      With .Range("b" & iRow)
         .Font.Size = 12
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
      End With
      
      .Range("c" & iRow).Value = "資料別："
      With .Range("c" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlRight
      End With
      
      Select Case Text1
         Case "1": strTmp = "放出"
         Case "2": strTmp = "全部"
         Case "4": strTmp = "空白"
         Case Else: strTmp = "餘額"
      End Select
   
      .Range("d" & iRow).Value = strTmp
      With .Range("d" & iRow & ":f" & iRow)
         .Font.Size = 12
         .Font.Bold = False
         .HorizontalAlignment = xlLeft
         .MergeCells = True
      End With
      
      'Modify by Amy 2017/05/23 隱藏 CFL 調整列印日期顯示
'      .Range("g" & iRow).Value = "列印日期："
'      With .Range("g" & iRow & ":h" & iRow)
'         .Font.Size = 12
'         .Font.Bold = True
'         .HorizontalAlignment =xlRight
'         .MergeCells = True
'      End With
      
'      strTmp = CFDate(strSrvDate(2))
'      .Range("i" & iRow).Value = strTmp
'      With .Range("i" & iRow)
'         .Font.Size = 12
'         .Font.Bold = False
'         .HorizontalAlignment = xlLeft
'      End With
      .Range("j" & iRow).Value = "列印日期：" & CFDate(strSrvDate(2))
      .Range("j" & iRow).Font.Size = 12
      .Range("j" & iRow).Font.Bold = True
      .Range("j" & iRow).HorizontalAlignment = xlLeft
      'end 2017/05/23
      
      iRow = iRow + 2
      
      .Range("a" & iRow).Value = "智權人員"
      .Columns("a").ColumnWidth = 11.5
      With .Range("a" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
      End With
      
      .Range("b" & iRow).Value = "大陸商標"
      .Columns("b").ColumnWidth = 14
      With .Range("b" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
      End With
      
      
      .Range("c" & iRow).Value = "大陸專利"
      .Columns("c").ColumnWidth = 11.5
      .Columns("d").ColumnWidth = 2
      With .Range("c" & iRow & ":d" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
         .MergeCells = True
      End With
      
      .Range("e" & iRow).Value = "國外商標"
      .Columns("e").ColumnWidth = 14
      With .Range("e" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
      End With
      
      .Range("f" & iRow).Value = "國外專利"
      .Columns("f").ColumnWidth = 12 '10.5
      .Columns("g").ColumnWidth = 3
      With .Range("f" & iRow & ":g" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
         .MergeCells = True
      End With
      
      'Add By Sindy 2013/1/15
      .Range("h" & iRow).Value = "ＣＦＬ"
      .Columns("h").ColumnWidth = 11.5
      .Columns("i").ColumnWidth = 2
      With .Range("h" & iRow & ":i" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
         .MergeCells = True
      End With
      '2013/1/15 End
      
      .Range("j" & iRow).Value = "合計"
      .Columns("j").ColumnWidth = 9
      .Columns("k").ColumnWidth = 10
      With .Range("j" & iRow & ":k" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
         .MergeCells = True
      End With
      
      'Add by Amy 2017/11/01 +目標及撥出結餘
      .Range("l" & iRow).Value = "目標"
      .Columns("l").ColumnWidth = 14
       With .Range("l" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
      End With
      
      .Range("m" & iRow).Value = "撥出結餘"
      .Columns("m").ColumnWidth = 16.5
       With .Range("m" & iRow)
         .Font.Size = 12
         .Font.Bold = True
         .HorizontalAlignment = xlCenter
      End With
      strAreaFrom = iRow + 1
      'end 2017/11/01
      
      adoRst.MoveFirst
      Do While Not adoRst.EOF
         iRow = iRow + 1
         
         With .Range("a" & iRow)
            .Value = "" & adoRst.Fields("Name")
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
         End With
         'Modify by Amy 2017/11/01
         With .Range("b" & iRow & ":m" & iRow)
            .Font.Size = 12
            .HorizontalAlignment = xlRight
            .NumberFormatLocal = stCellFormat
         End With
            
         '合計欄位(放公式)
         '全所
         If "" & adoRst.Fields("dept") = "ZZZ" Then
            .Range("a" & iRow & ":m" & iRow).Font.Bold = True 'Modify by Amy 2017/11/01
            .Range("b" & iRow).Formula = "=SUM(" & strTotalFormula & ")"
            .Range("c" & iRow).Formula = "=SUM(" & Replace(strTotalFormula, "b", "c") & ")"
            .Range("e" & iRow).Formula = "=SUM(" & Replace(strTotalFormula, "b", "e") & ")"
            .Range("f" & iRow).Formula = "=SUM(" & Replace(strTotalFormula, "b", "f") & ")"
            .Range("h" & iRow).Formula = "=SUM(" & Replace(strTotalFormula, "b", "h") & ")"
            'Add by Amy 2017/11/01 +目標及撥出結餘
            .Range("l" & iRow).Formula = "=SUM(" & Replace(strTotalFormula, "b", "l") & ")"
            .Range("m" & iRow).Formula = "=SUM(" & Replace(strTotalFormula, "b", "m") & ")"
         '部門/所合計
         'Modify by Amy 2018/01/11 總所為M0100
         ElseIf Right("" & adoRst.Fields("dept"), 1) = "T" And "" & adoRst.Fields("dept") <> "XTOT" Then
            'Add by Amy 2014/12/16 +隱藏台南所和高雄所合計-瑞婷
            If "" & adoRst.Fields("dept") = "S31T" Or "" & adoRst.Fields("Dept") = "S41T" Then
                strHiddenRow = strHiddenRow & iRow & ";"
            End If
            'end 2014/12/16
            .Range("a" & iRow & ":m" & iRow).Font.Bold = True 'Modify by Amy 2017/11/01
            '區小計
            If Left("" & adoRst.Fields("dept"), 3) = Left(strLastDept, 3) Then
               .Range("b" & iRow).Formula = "=SUM(b" & strAreaFrom & ":b" & (iRow - 1) & ")"
               .Range("c" & iRow).Formula = "=SUM(c" & strAreaFrom & ":c" & (iRow - 1) & ")"
               .Range("e" & iRow).Formula = "=SUM(e" & strAreaFrom & ":e" & (iRow - 1) & ")"
               .Range("f" & iRow).Formula = "=SUM(f" & strAreaFrom & ":f" & (iRow - 1) & ")"
               .Range("h" & iRow).Formula = "=SUM(h" & strAreaFrom & ":h" & (iRow - 1) & ")"
               'Add by Amy 2017/11/01 +目標及撥出結餘
               .Range("l" & iRow).Formula = "=SUM(l" & strAreaFrom & ":l" & (iRow - 1) & ")"
               .Range("m" & iRow).Formula = "=SUM(m" & strAreaFrom & ":m" & (iRow - 1) & ")"
               strBranchFormula = strBranchFormula & IIf(strBranchFormula <> "", ",", "") & "b" & iRow
            '所小計
            Else
               If strBranchFormula <> "" Then
                  .Range("b" & iRow).Formula = "=SUM(" & strBranchFormula & ")"
                  .Range("c" & iRow).Formula = "=SUM(" & Replace(strBranchFormula, "b", "c") & ")"
                  .Range("e" & iRow).Formula = "=SUM(" & Replace(strBranchFormula, "b", "e") & ")"
                  .Range("f" & iRow).Formula = "=SUM(" & Replace(strBranchFormula, "b", "f") & ")"
                  .Range("h" & iRow).Formula = "=SUM(" & Replace(strBranchFormula, "b", "h") & ")"
                  'Add by Amy 2017/11/01 +目標及撥出結餘
                  .Range("l" & iRow).Formula = "=SUM(" & Replace(strBranchFormula, "b", "l") & ")"
                  .Range("m" & iRow).Formula = "=SUM(" & Replace(strBranchFormula, "b", "m") & ")"
                'Mark by Amy 2018/01/11 語法有剔除目標和合計是0
'               'Add by Amy 2017/11/01 ex:廣東下無資料
'               ElseIf Val(strAreaFrom) = iRow - 1 Then
'                .Range("b" & iRow).Value = Val("" & adoRst.Fields("T1"))
'                .Range("c" & iRow).Value = Val("" & adoRst.Fields("P1"))
'                .Range("e" & iRow).Value = Val("" & adoRst.Fields("T2"))
'                .Range("f" & iRow).Value = Val("" & adoRst.Fields("P2"))
'                .Range("h" & iRow).Value = Val("" & adoRst.Fields("CFL"))
'                '目標(撥出結餘不需帶值)
'                .Range("l" & iRow).Value = Val("" & adoRst.Fields("PE04"))
               '沒有分區的
               Else
                  .Range("b" & iRow).Formula = "=SUM(b" & strAreaFrom & ":b" & (iRow - 1) & ")"
                  .Range("c" & iRow).Formula = "=SUM(c" & strAreaFrom & ":c" & (iRow - 1) & ")"
                  .Range("e" & iRow).Formula = "=SUM(e" & strAreaFrom & ":e" & (iRow - 1) & ")"
                  .Range("f" & iRow).Formula = "=SUM(f" & strAreaFrom & ":f" & (iRow - 1) & ")"
                  .Range("h" & iRow).Formula = "=SUM(h" & strAreaFrom & ":h" & (iRow - 1) & ")"
                  'Add by Amy 2017/11/01 +目標及撥出結餘
                  .Range("l" & iRow).Formula = "=SUM(l" & strAreaFrom & ":l" & (iRow - 1) & ")"
                  .Range("m" & iRow).Formula = "=SUM(m" & strAreaFrom & ":m" & (iRow - 1) & ")"
               End If
               strBranchFormula = ""
               strTotalFormula = strTotalFormula & IIf(strTotalFormula <> "", ",", "") & "b" & iRow
            End If
         Else
            .Range("a" & iRow & ":k" & iRow).Font.Bold = False
            .Range("b" & iRow).Value = Val("" & adoRst.Fields("T1"))
            .Range("c" & iRow).Value = Val("" & adoRst.Fields("P1"))
            .Range("e" & iRow).Value = Val("" & adoRst.Fields("T2"))
            .Range("f" & iRow).Value = Val("" & adoRst.Fields("P2"))
            .Range("h" & iRow).Value = Val("" & adoRst.Fields("CFL"))
            'Add by Amy 2017/11/01 +目標(撥出結餘不需帶值)
            .Range("l" & iRow).Value = Val("" & adoRst.Fields("PE04"))
            
         End If
         'Modify by Amy 2018/01/11 +其他合計加總起始位置判斷
         If (Left(strLastDept, 1) = "S" And Len(strLastDept) = "3" And Left("" & adoRst.Fields("dept"), 1) = "X") Or _
               ("" & adoRst.Fields("dept") <> strLastDept And Left("" & adoRst.Fields("dept"), 1) <> "X") Then
            strAreaFrom = iRow
         End If
         .Range("j" & iRow).Formula = "=SUM(a" & iRow & ":h" & iRow & ")"
         .Range("c" & iRow & ":d" & iRow).MergeCells = True
         .Range("f" & iRow & ":g" & iRow).MergeCells = True
         .Range("h" & iRow & ":i" & iRow).MergeCells = True
         .Range("j" & iRow & ":k" & iRow).MergeCells = True
            
         strLastDept = "" & adoRst.Fields("dept")
         adoRst.MoveNext
      Loop
      
      'Add by Amy 2014/12/16 +隱藏台南所和高雄所合計
      arrRow = Split(strHiddenRow, ";")
      For ii = 0 To UBound(arrRow) - 1
        .Rows(arrRow(ii) & ":" & arrRow(ii)).Hidden = True
      Next ii
      'end 2014/12/16
      
      'Add by Amy 2017/05/23 隱藏CFL
      .Range("h:i").EntireColumn.Hidden = True
      
      '.Columns("a:" & MaxCol).EntireColumn.AutoFit
      'Modify by Amy 2017/11/01 +目標及撥出結餘
      With .Range("A8:m" & iRow).Borders(xlEdgeLeft)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Range("A8:m" & iRow).Borders(xlEdgeTop)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Range("A8:m" & iRow).Borders(xlEdgeBottom)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Range("A8:m" & iRow).Borders(xlEdgeRight)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Range("A8:m" & iRow).Borders(xlInsideVertical)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
      With .Range("A8:m" & iRow).Borders(xlInsideHorizontal)
          .LineStyle = xlContinuous
          .Weight = xlThin
          .ColorIndex = xlAutomatic
      End With
    
      'Modify by Amy 2015/05/21 原使用函數以為是抓A4紙張
      .PageSetup.PaperSize = 9 '設定紙張 A4
      .PageSetup.Orientation = xlPortrait '直印
      .PageSetup.PrintTitleRows = "$1:$7" '表頭保留7列
      .PageSetup.PrintArea = "$A$1:$M$" & iRow '設定列印範圍
      'end 2017/11/01
      .PageSetup.LeftMargin = xlsSalesPoint.InchesToPoints(0.5) '左邊界
      .PageSetup.RightMargin = xlsSalesPoint.InchesToPoints(0.5) '右邊界
      .PageSetup.Zoom = False '100 '縮放比例
   End With
   'Modify by Amy 2016/06/23 +判斷版本
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   'end 2016/06/23
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   
   MsgBox "Excel檔已產生!" & vbCrLf & vbCrLf & strFileName, vbInformation, Me.Caption
   Exit Sub

ErrHnd:
   MsgBox Err.Description
   
End Sub

'*************************************************
'  查詢資料表(傳票資料)
'
'*************************************************
'Mark by Amy 2017/11/01 發現資料未讀暫存資料(可能之前使用AccRepot),因增加目標欄位故改抓暫存資料
Public Sub QueryTable_Old()
'
''Add by Morgan 2004/11/15
'   Dim bolCancel As Boolean
'   Call txtSalesNo_Validate(bolCancel)
'   If bolCancel = True Then Exit Sub
''2004/11/15 end
'
'Dim StrSQLa As String, StrSqlB As String
'
'On Error GoTo Checking
'   strSql = ""
'   If adoadodc1.State = adStateOpen Then
'      adoadodc1.Close
'   End If
'   adoadodc1.CursorLocation = adUseClient
'   '92.11.19 ADD BY SONIA
'   StrSqlB = ""
'   Select Case Text1
'      Case "1"
'         StrSqlB = "ax206"
'      Case "2"
'         StrSqlB = "ax207"
'      Case Else
'         StrSqlB = "ax207-ax206"
'   End Select
'   '92.11.19 END
''   Select Case strAccount
''      Case "2"
''         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
''            strSql = strSql & " and a0305 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
''         End If
''         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
''            strSql = strSql & " and a0305 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
''         End If
''         'Add by Morgan 2004/11/15 加智權人員條件
''         If txtSalesNo <> "" Then
''            strSql = strSql & " and ax209='" & txtSalesNo.Text & "'"
''         End If
''
''         '92.11.19 MODIFY BY SONIA
''         'strSQL = "select st15 as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, staff where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & " group by st15, st01, st02" & _
''         '         " union select a0901||'T' as Dept, a0901||'T' as ID, a0902 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & " group by a0901, a0902" & _
''         '         " union select 'S1T' as Dept, 'S1T' as ID, '北所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S1' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''         '         " union select 'S2T' as Dept, 'S2T' as ID, '中所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S2' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''         '         " union select 'S3T' as Dept, 'S3T' as ID, '南所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S3' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''         '         " union select 'S4T' as Dept, 'S4T' as ID, '高所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S4' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''         '         " union select 'S9T' as Dept, 'S9T' as ID, '廣東' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S9' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''         '         " union select 'X01' as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, staff where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & " group by st01, st02" & _
''         '         " union select 'X0T' as Dept, 'X0T' as ID, '其他合計' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''         '         " union select 'ZZZ' as Dept, 'ZZZ' as ID, '全所合計' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and ax205 in ('249101', '249102', '249103', '249104')" & strSQL
''         'Modify By Sindy 2013/1/16 +249105
''         strSql = "select st15 as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(decode(ax205, '249105', ax207-ax206, 0)) as CFL, sum(ax207-ax206) as TOTAL from acc021, acc020, staff where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st15, st01, st02" & _
''                  " union select a0901||'T' as Dept, a0901||'T' as ID, a0902 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(decode(ax205, '249105', ax207-ax206, 0)) as CFL, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by a0901, a0902" & _
''                  " union select 'S1T' as Dept, 'S1T' as ID, '北所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(decode(ax205, '249105', ax207-ax206, 0)) as CFL, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S1' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & _
''                  " union select 'S2T' as Dept, 'S2T' as ID, '中所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(decode(ax205, '249105', ax207-ax206, 0)) as CFL, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S2' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & _
''                  " union select 'S3T' as Dept, 'S3T' as ID, '南所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(decode(ax205, '249105', ax207-ax206, 0)) as CFL, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S3' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & _
''                  " union select 'S4T' as Dept, 'S4T' as ID, '高所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(decode(ax205, '249105', ax207-ax206, 0)) as CFL, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S4' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & _
''                  " union select 'S9T' as Dept, 'S9T' as ID, '廣東' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(decode(ax205, '249105', ax207-ax206, 0)) as CFL, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S9' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & _
''                  " union select 'X01' as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(decode(ax205, '249105', ax207-ax206, 0)) as CFL, sum(ax207-ax206) as TOTAL from acc021, acc020, staff where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st01, st02" & _
''                  " union select 'X0T' as Dept, 'X0T' as ID, '其他合計' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(decode(ax205, '249105', ax207-ax206, 0)) as CFL, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & _
''                  " union select 'ZZZ' as Dept, 'ZZZ' as ID, '全所合計' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(decode(ax205, '249105', ax207-ax206, 0)) as CFL, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql
''         '92.11.19 END
''         If Trim(Text4) <> MsgText(601) Then
''            strSql = "select * from (" & strSql & ") New where ax201 = '" & IIf(Text4 = "2", "J", "1") & "'"
''         End If
''         adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
''      Case Else
'         'Mark by Amy 2017/07/10 有加下列條件人員可能抓不完整
'         'Modify by Amy 2013/08/13 改帶系統年1月1日避免run太久
''         If Text1 = "4" Then
''            MaskEdBox1.Text = Left(strSrvDate(2), 3) & "/01/01"
''         End If
'
'         If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'            strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'         End If
'         If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'            strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'         End If
'
'         'Add by Morgan 2004/11/15 加智權人員條件
'         If txtSalesNo <> "" Then
'            strSql = strSql & " and ax209='" & txtSalesNo.Text & "'"
'         End If
'
'            'Modify By Cheng 2003/07/21
'            '大陸商標, 大陸專利, 國外商標, 國外專利的金額不可同時為0
''         strSQL = "select st15 as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, staff where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & " group by st15, st01, st02" & _
''                  " union select a0901||'T' as Dept, a0901||'T' as ID, a0902 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & " group by a0901, a0902" & _
''                  " union select 'S1T' as Dept, 'S1T' as ID, '北所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S1' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''                  " union select 'S2T' as Dept, 'S2T' as ID, '中所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S2' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''                  " union select 'S3T' as Dept, 'S3T' as ID, '南所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S3' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''                  " union select 'S4T' as Dept, 'S4T' as ID, '高所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S4' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''                  " union select 'S9T' as Dept, 'S9T' as ID, '廣東' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S9' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''                  " union select 'T01' as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, staff where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & " group by st01, st02" & _
''                  " union select 'X0T' as Dept, 'X0T' as ID, '其他合計' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & _
''                  " union select 'ZZZ' as Dept, 'ZZZ' as ID, '全所合計' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and ax205 in ('249101', '249102', '249103', '249104')" & strSQL
'         '92.11.19 MODIFY BY SONIA
'         '   strSQLA = " Having (Nvl(sum(decode(ax205, '249101', ax207-ax206, 0)),0)<>0 Or Nvl(sum(decode(ax205, '249102', ax207-ax206, 0)),0)<>0 Or Nvl(sum(decode(ax205, '249103', ax207-ax206, 0)),0)<>0 Or Nvl(sum(decode(ax205, '249104', ax207-ax206, 0)),0)<>0) "
'         'strSQL = "select st15 as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, staff where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & " group by st15, st01, st02" & strSQLA & _
'         '         " union select a0901||'T' as Dept, a0901||'T' as ID, a0902 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & " group by a0901, a0902" & strSQLA & _
'         '         " union select 'S1T' as Dept, 'S1T' as ID, '北所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S1' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & strSQLA & _
'         '         " union select 'S2T' as Dept, 'S2T' as ID, '中所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S2' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & strSQLA & _
'         '         " union select 'S3T' as Dept, 'S3T' as ID, '南所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S3' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & strSQLA & _
'         '         " union select 'S4T' as Dept, 'S4T' as ID, '高所' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S4' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & strSQLA & _
'         '         " union select 'S9T' as Dept, 'S9T' as ID, '廣東' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S9' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & strSQLA & _
'         '         " union select 'X01' as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, staff where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & " group by st01, st02" & strSQLA & _
'         '         " union select 'X0T' as Dept, 'X0T' as ID, '其他合計' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & strSQLA & _
'         '         " union select 'ZZZ' as Dept, 'ZZZ' as ID, '全所合計' as Name, sum(decode(ax205, '249101', ax207-ax206, 0)) as T1, sum(decode(ax205, '249102', ax207-ax206, 0)) as P1, sum(decode(ax205, '249103', ax207-ax206, 0)) as T2, sum(decode(ax205, '249104', ax207-ax206, 0)) as P2, sum(ax207-ax206) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and ax205 in ('249101', '249102', '249103', '249104')" & strSQL & strSQLA
'         'Modify By Sindy 2013/1/16 +249105
'         StrSQLa = " Having (Nvl(sum(decode(ax205, '249101', " & StrSqlB & " , 0)),0)<>0 Or Nvl(sum(decode(ax205, '249102', " & StrSqlB & " , 0)),0)<>0 Or Nvl(sum(decode(ax205, '249103', " & StrSqlB & " , 0)),0)<>0 Or Nvl(sum(decode(ax205, '249104', " & StrSqlB & " , 0)),0)<>0 Or Nvl(sum(decode(ax205, '249105', " & StrSqlB & " , 0)),0)<>0) "
'         If Text1 = "4" Then
'            'Modify by Amy 2013/07/30 select + CFL
'            'strSql = "select st15 as Dept, st01 as ID, st02 as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as TOTAL from acc021, acc020, staff where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st15, st01, st02" & StrSQLa & _
'                     " union select a0901||'T' as Dept, a0901||'T' as ID, a0902 as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by a0901, a0902" & StrSQLa & _
'                     " union select 'S1T' as Dept, 'S1T' as ID, '北所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S1' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S2T' as Dept, 'S2T' as ID, '中所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S2' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S3T' as Dept, 'S3T' as ID, '南所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S3' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S4T' as Dept, 'S4T' as ID, '高所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S4' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S9T' as Dept, 'S9T' as ID, '廣東' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S9' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'X01' as Dept, st01 as ID, st02 as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as TOTAL from acc021, acc020, staff where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st01, st02" & StrSQLa & _
'                     " union select 'X0T' as Dept, 'X0T' as ID, '其他合計' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'ZZZ' as Dept, 'ZZZ' as ID, '全所合計' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201 = a0201 and ax202 = a0202 and ax209 = st01 (+) and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & " ORDER BY DEPT,ID"
'           strSql = "select st15 as Dept, st01 as ID, st02 as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, staff where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st15, st01, st02" & StrSQLa & _
'                     " union select a0901||'T' as Dept, a0901||'T' as ID, a0902 as Name, '' as T1, '' as P1, '' as T2, '' as CFL,'' as P2, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by a0901, a0902" & StrSQLa & _
'                     " union select 'S1T' as Dept, 'S1T' as ID, '北所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S1' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S2T' as Dept, 'S2T' as ID, '中所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S2' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S3T' as Dept, 'S3T' as ID, '南所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S3' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S4T' as Dept, 'S4T' as ID, '高所' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S4' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S9T' as Dept, 'S9T' as ID, '廣東' as Name, '' as T1, '' as P1, '' as T2, '' as P2, '' as CFL,'' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S9' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'X01' as Dept, st01 as ID, st02 as Name, '' as T1, '' as P1, '' as T2, '' as P2,'' as CFL, '' as TOTAL from acc021, acc020, staff where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st01, st02" & StrSQLa & _
'                     " union select 'X0T' as Dept, 'X0T' as ID, '其他合計' as Name, '' as T1, '' as P1, '' as T2, '' as P2,'' as CFL, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'ZZZ' as Dept, 'ZZZ' as ID, '全所合計' as Name, '' as T1, '' as P1, '' as T2, '' as P2,'' as CFL, '' as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & " ORDER BY DEPT,ID"
'         Else
'            strSql = "select st15 as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, staff where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st15, st01, st02" & StrSQLa & _
'                     " union select a0901||'T' as Dept, a0901||'T' as ID, a0902 as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) = 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by a0901, a0902" & StrSQLa & _
'                     " union select 'S1T' as Dept, 'S1T' as ID, '北所' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S1' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S2T' as Dept, 'S2T' as ID, '中所' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S2' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S3T' as Dept, 'S3T' as ID, '南所' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S3' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S4T' as Dept, 'S4T' as ID, '高所' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S4' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'S9T' as Dept, 'S9T' as ID, '廣東' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 2) = 'S9' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'X01' as Dept, st01 as ID, st02 as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, staff where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & " group by st01, st02" & StrSQLa & _
'                     " union select 'X0T' as Dept, 'X0T' as ID, '其他合計' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and substr(st15, 1, 1) <> 'S' and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & _
'                     " union select 'ZZZ' as Dept, 'ZZZ' as ID, '全所合計' as Name, sum(decode(ax205, '249101', " & StrSqlB & " , 0)) as T1, sum(decode(ax205, '249102', " & StrSqlB & " , 0)) as P1, sum(decode(ax205, '249103', " & StrSqlB & " , 0)) as T2, sum(decode(ax205, '249104', " & StrSqlB & " , 0)) as P2, sum(decode(ax205, '249105', " & StrSqlB & " , 0)) as CFL, sum(" & StrSqlB & " ) as TOTAL from acc021, acc020, (select * from staff, acc090 where st15 = a0901) new where ax201(+) = a0201" & IIf(Text4 = "2", " and a0201='J'", IIf(Text4 = "1", " and a0201='1'", "")) & " and ax202(+) = a0202 and ax209 = st01 (+) and ax205 in ('249101', '249102', '249103', '249104', '249105')" & strSql & StrSQLa & " ORDER BY DEPT,ID"
'         End If
'         '92.11.19 END
''         If Trim(Text4) <> MsgText(601) Then
''            strSql = "select * from (" & strSql & ") New where ax201 = '" & IIf(Text4 = "2", "J", "1") & "'"
''         End If
'         adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
''   End Select
'   adoTaie.Execute "delete from accrpt423 Where R42301='" & strUserNum & "' " 'Modify By Sindy 2011/01/07 'Modify by Amy 2013/07/26 +where
'   Adodc1.Recordset.Requery
'   If Adodc1.Recordset.RecordCount = 0 Then
'      Adodc1.Recordset.Close
'      Command1.Enabled = False
'      MsgBox MsgText(28), , MsgText(5)
'      Exit Sub
'   Else
'      Command1.Enabled = True
'      'Modify By Sindy 2011/01/07
'      adoadodc1.MoveFirst
'      Do While adoadodc1.EOF = False
'         'Modify by AMy 2013/11/05 +存ID欄位(R42310)
'         'Modify by Amy2013/07/26 +欄位名稱
'         'Modify By Sindy 2013/1/16 +249105
'         adoTaie.Execute "insert into accrpt423 (R42301,R42302,R42303,R42304,R42305,R42306,R42307,R42308,R42309,R42310) values ('" & strUserNum & "', '" & adoadodc1.Fields("Dept").Value & "', '" & adoadodc1.Fields("Name").Value & "', " & _
'         IIf(IsNull(adoadodc1.Fields("T1").Value), "Null", adoadodc1.Fields("T1").Value) & ", " & IIf(IsNull(adoadodc1.Fields("P1").Value), "Null", adoadodc1.Fields("P1").Value) & ", " & IIf(IsNull(adoadodc1.Fields("T2").Value), "Null", adoadodc1.Fields("T2").Value) & ", " & IIf(IsNull(adoadodc1.Fields("P2").Value), "Null", adoadodc1.Fields("P2").Value) & ", " & IIf(IsNull(adoadodc1.Fields("CFL").Value), "Null", adoadodc1.Fields("CFL").Value) & ", " & IIf(IsNull(adoadodc1.Fields("TOTAL").Value), "Null", adoadodc1.Fields("TOTAL").Value) & _
'         ", '" & adoadodc1.Fields("ID").Value & "' )"
'         adoadodc1.MoveNext
'      Loop
'      '2011/01/07 End
'   End If
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
End Sub
