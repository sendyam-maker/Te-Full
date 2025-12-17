VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc4250 
   AutoRedraw      =   -1  'True
   Caption         =   "智權人員點數查詢"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   8175
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
      Left            =   1800
      TabIndex        =   11
      Top             =   240
      Width           =   3500
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6192
      TabIndex        =   9
      Top             =   4620
      Width           =   1500
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc4250.frx":0000
      Height          =   3000
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5292
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
      Caption         =   "智權人員點數資料"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "ax201"
         Caption         =   "公司"
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
         DataField       =   "a0205"
         Caption         =   "傳票日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "@@@/@@/@@"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "ax212"
         Caption         =   "摘要內容"
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
         DataField       =   "Amount"
         Caption         =   "點數"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2984.882
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1289.764
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   1440
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
   Begin VB.TextBox Text2 
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
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   1572
      _ExtentX        =   2778
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
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   1572
      _ExtentX        =   2778
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
   Begin MSForms.TextBox Text1 
      Height          =   300
      Left            =   3390
      TabIndex        =   6
      Top             =   960
      Width           =   2295
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      Size            =   "4048;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "合計"
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
      Left            =   5544
      TabIndex        =   10
      Top             =   4632
      Width           =   612
   End
   Begin VB.Label Label2 
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
      TabIndex        =   8
      Top             =   240
      Width           =   972
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   24
      Top             =   4788
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1335
      Left            =   225
      Top             =   135
      Width           =   7695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "對沖代號(業)"
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
      TabIndex        =   5
      Top             =   960
      Width           =   1452
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
      Left            =   3480
      TabIndex        =   4
      Top             =   600
      Width           =   132
   End
   Begin VB.Label Label3 
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
      TabIndex        =   3
      Top             =   600
      Width           =   972
   End
End
Attribute VB_Name = "Frmacc4250"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/05 Form2.0已修改 text1/DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit

Public adoadodc1 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strSql As String

'Add by Amy 2020/03/31
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
    If InStr(GetBookKeepCmp, strCmp) = 0 Then
        MsgBox Label2 & MsgText(63), , MsgText(5)
        Cancel = True
        Combo1.SetFocus
        Exit Sub
    ElseIf Len(Trim(Combo1)) = 1 Then
        Combo1 = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/03/31

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call PUB_SaveTrackMode(0, KeyCode)  ' Add by Amy 2022/01/05 Form2.0 記錄鍵盤傳入順序
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
   Me.Width = 8300
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
   'Text3 = "1"
   'Add by Amy 2020/03/31
   Combo1.AddItem "", 0
   Call Pub_SetCboCmp(Combo1, False, False, False, , 1)
   'end 2020/03/31
   'add by sonia 2013/8/6 預設當月1日至月底
   MaskEdBox1.Text = Mid(CFDate(ACDate(ServerDate)), 1, 7) & "01"
   MaskEdBox2.Text = CFDate(ACDate(GetLastDay(DBDATE(FCDate(MaskEdBox1)))))
   'end 2013/8/6
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   strTrackMode = "" 'Add by Amy 2022/01/05 Form2.0 記錄鍵盤傳入順序
   MenuEnabled
   Set Frmacc4250 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc021, acc020 where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & " and ax209 = '" & Text2 & "' order by ax202 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  查詢並顯示資料
'
'*************************************************
Public Sub AdodcQuery()
Dim strCmp As String, strSql As String 'Add by Amy 2020/03/31

On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   strSql = ""
   'Modify by Amy 2020/03/31 改下拉 原:Text3
   If Trim(Combo1) <> MsgText(601) Then
      strCmp = Combo1
      If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
      End If
'      'Add By Sindy 2014/1/21
'      If Text3 = "2" Then
'         strSql = " and ax201='J'"
'      Else
'      '2014/1/21 END
         strSql = " and ax201='" & strCmp & "'"
'      End If
   End If
   'end 2020/03/31
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0205 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0205 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and ax209 = '" & Text2 & "'"
   End If
   'Add By Cheng 2004/01/14
   '若非北所員工, 只能列印該所資料
   If pub_strUserOffice <> "1" Then
      strSql = strSql & " And ''||ST06='" & pub_strUserOffice & "' "
   End If
   'End
   '         adoadodc1.Open "select distinct ax201, ax202, a0k04, decode(ax207, 0, ax206 * -1, ax207) as Amount, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Caseno, ax212 from acc021, acc020, (select distinct a1p22, a0k04 from acc0m0, acc1p0, acc0k0, acc021, acc020 where a0m01 = a1p04 and a0m02 = a0k01 and a1p22 = ax202 and a1p05 = ax205 and ax201 = a0201 and ax202 = a0202 and substr(ax205, 1, 1) = '4'" & strSQL & ") new where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and ax202 = a1p22 (+) and substr(ax205, 1, 1) = '4'" & strSQL & _
   '                        " union select distinct ax201, ax202, a0k04, decode(ax207, 0, ax206 * -1, ax207) as Amount, ax209, substr(ax214, 1, length(ax214) - 9)||'-'||substr(ax214, length(ax214) - 8, 6)||'-'||substr(ax214, length(ax214) - 2, 1)||'-'||substr(ax214, length(ax214) - 1, 2) as Caseno, ax212 from acc021, acc020, (select distinct a1p22, a0k04 from acc0s0, acc1p0, acc0k0, acc021, acc020 where a0s01 = a1p04 and a0s02 = a0k01 and a1p22 = ax202 and a1p05 = ax205 and ax201 = a0201 and ax202 = a0202 and substr(ax205, 1, 1) = '4'" & strSQL & ") new where acc021.ax201 = acc020.a0201 and acc021.ax202 = acc020.a0202 and ax202 = a1p22 and substr(ax205, 1, 1) = '4'" & strSQL & " order by ax202 asc", adoTaie, adOpenStatic, adLockReadOnly
      'Modify By Cheng 2003/06/03
   '         adoadodc1.Open "select ax201, ax202, ax212, decode(ax207, 0, ax206 * -1, ax207) as Amount from acc021, acc020 where ax201 = a0201 (+) and ax202 = a0202 (+) and substr(ax205, 1, 2) = '41'" & strSQL & " order by ax202 asc", adoTaie, adOpenStatic, adLockReadOnly
      'Modify By Cheng 2004/01/14
   '         adoadodc1.Open "select ax201, ax202, ax212, decode(ax207, 0, ax206 * -1, ax207) as Amount from acc021, acc020 where ax201 = a0201 (+) and ax202 = a0202 (+) and (substr(ax205, 1, 2) = '41' Or (ax205 = '7121' and ax209 Is Not Null)) " & strSQL & " order by ax202 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Morgan 2005/6/16 加傳票日
   'Modify By Sindy 2014/1/21 order by ax202 asc ==> order by a0205,ax201,ax202,ax203 asc
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   'Modify by Amy 2020/03/31 +7129
   strSql = "select ax201, ax202, ax212, decode(ax207, 0, ax206 * -1, ax207) as Amount,a0205 from acc021, acc020, Staff " & _
                "where ax201(+)=a0201 and ax202(+)=a0202 and (substr(ax205, 1, 1) = '4' Or ((ax205 = '7121' or ax205='7129' ) and ax209 Is Not Null)) And AX209=ST01(+) " & strSql & _
                " order by a0205,ax201,ax202,ax203 asc"
   adoadodc1.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   'End
   Adodc1.Recordset.Requery
   SumShow
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

'ADD BY SONIA 2013/8/6 預設止日為起日的該月月底
Private Sub MaskEdBox1_LostFocus()
   'Add by Amy 2017/11/20 年月日輸錯會錯誤
   If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then Exit Sub
     
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      MaskEdBox1.SetFocus
      Exit Sub
   End If
   'end 2017/11/20
   MaskEdBox2.Text = CFDate(ACDate(GetLastDay(DBDATE(FCDate(MaskEdBox1)))))
End Sub
'END 2013/8/6

'Add by Amy 2017/11/20 年月日輸錯會錯誤
Private Sub MaskEdBox2_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = MsgText(601) Or MaskEdBox2.Text = MsgText(29) Then Exit Sub
     
   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox Label3 & MsgText(63), , MsgText(5)
      Cancel = True
      MaskEdBox2.SetFocus
      Exit Sub
   End If
  
End Sub

Private Sub Text2_Change()
   If Text2 = MsgText(601) Then
      Exit Sub
   End If
   Text1 = StaffQuery(Text2)
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)   '2009/6/5 add by sonia
End Sub

'Mark by Amy 2020/03/31 改下拉
'Private Sub Text3_Change()
'   If Text3 = MsgText(601) Then
'      Exit Sub
'   End If
''   Text4 = A0802Query(Text3)
'End Sub
'
'Private Sub Text3_GotFocus()
'   TextInverse Text3
'End Sub
'
'Private Sub Text3_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub
'end 2020/03/31

'*************************************************
'  功能鍵定義
'
'*************************************************
Public Sub KeyDefine(KeyCode As Integer)
   Call PUB_SaveTrackMode(1, KeyCode) 'Add by Amy 2022/01/05 Form2.0
   Select Case KeyCode
      Case vbKeyF12
         'Add by Amy 2022/01/05 Form2.0控制Function鍵：記錄鍵盤傳入順序，判斷是否可執行
         If PUB_ChkTrackMode = False Then
                Exit Sub
         End If
         'end 2022/01/05

         If FormCheck Then
            Screen.MousePointer = vbHourglass
            AdodcQuery
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
   Text6 = ""
   If Adodc1.Recordset.State <> adStateOpen Then
      Exit Sub
   End If
   Set adoaccsum = Adodc1.Recordset.Clone
   Do While adoaccsum.EOF = False
      Text6 = Val(Text6) + adoaccsum.Fields("Amount").Value
      adoaccsum.MoveNext
   Loop
   adoaccsum.Close
   Text6 = Format(Text6, FDollar)
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   'Add by Amy 2020/03/31
   Dim bolCancel As Boolean
   
   If Trim(Combo1) <> MsgText(601) Then
      Combo1_Validate (bolCancel)
      If bolCancel = False Then
        FormCheck = True
        Exit Function
      End If
   End If
   'end 2020/03/31
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
