VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc7112 
   AutoRedraw      =   -1  'True
   Caption         =   "分所收款智權人員統計資料"
   ClientHeight    =   5025
   ClientLeft      =   1515
   ClientTop       =   2505
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   8760
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1230
      MaxLength       =   100
      TabIndex        =   8
      Top             =   960
      Width           =   6045
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc7112.frx":0000
      Height          =   3675
      Left            =   150
      TabIndex        =   6
      Top             =   1320
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   6482
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   20
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
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "收款人"
         Caption         =   "收款人"
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
         DataField       =   "現金"
         Caption         =   "現金"
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
      BeginProperty Column02 
         DataField       =   "支票"
         Caption         =   "支票"
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
      BeginProperty Column03 
         DataField       =   "合計"
         Caption         =   "合計"
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
      BeginProperty Column04 
         DataField       =   "點數"
         Caption         =   "點數"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.000"
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
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1590.236
         EndProperty
      EndProperty
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
      Height          =   315
      Left            =   1230
      TabIndex        =   2
      Top             =   570
      Width           =   975
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   1230
      TabIndex        =   0
      Top             =   210
      Width           =   1575
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   300
      Left            =   3150
      TabIndex        =   1
      Top             =   210
      Width           =   1575
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   240
      Top             =   1200
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
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "收款抬頭"
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
      Left            =   240
      TabIndex        =   9
      Top             =   990
      Width           =   975
   End
   Begin VB.Label lblSalesName 
      BackStyle       =   0  '透明
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
      Left            =   2400
      TabIndex        =   7
      Top             =   570
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   570
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "收款日期"
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
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   2910
      TabIndex        =   3
      Top             =   210
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc7112"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit
Public adoadodc1 As New ADODB.Recordset

Private Sub Command1_Click()
   If Adodc1.Recordset.RecordCount = 0 Then
      Exit Sub
   End If
    If Adodc1.Recordset.Fields("R43008").Value & "," & Adodc1.Recordset.Fields("R43007").Value = "," Then
        Exit Sub
    End If
   strItemNo = Adodc1.Recordset.Fields("R43008").Value & "," & Adodc1.Recordset.Fields("R43007").Value
   strCon4 = MaskEdBox2.Text
   strCon5 = MaskEdBox3.Text
   strExitControl = MsgText(601)
   tool3_enabled
   Screen.MousePointer = vbHourglass
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   Frmacc7111.Show
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   Me.Hide
End Sub

Private Sub Form_Activate()
'   MaskEdBox2.Mask = ""
'   MaskEdBox2.Text = strCon4
'   MaskEdBox2.Mask = DFormat
'   MaskEdBox3.Mask = ""
'   MaskEdBox3.Text = strCon5
'   MaskEdBox3.Mask = DFormat
'   If MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> MsgText(601) Then
'      AdodcRefresh
'   End If
'   strExitControl = MsgText(602)
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
   Me.Width = 8850
   Me.Height = 5400
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox2.Mask = MsgText(601)
   MaskEdBox2.Text = MsgText(601)
   MaskEdBox2.Mask = DFormat
   Me.MaskEdBox2.Text = strCon4
   MaskEdBox3.Mask = MsgText(601)
   MaskEdBox3.Text = MsgText(601)
   MaskEdBox3.Mask = DFormat
   Me.MaskEdBox3.Text = strCon5
   Me.Text1.Text = strCon6
   Me.lblSalesName.Caption = strCon7
   Me.Text2.Text = strCon8   'add by sonia 2018/4/19
   OpenTable
   AdodcRefresh
   strExitControl = MsgText(602)
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strExitControl = MsgText(602) Then
      StatusClear
      strFormName = MsgText(601)
      KeyEnter vbKeyEscape
      MenuEnabled
      Set Frmacc7112 = Nothing
        tool3_enabled
        MenuDisabled
        Frmacc7110.Show
        strFormName = "Frmacc7110"
      Exit Sub
   End If
   strExitControl = MsgText(602)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "Select * From ACC310 Where A3101='" & pub_strUserOffice & "' And A3103='" & Text1 & "' And RowNum < 1 Order By A3103 ", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  重新整理 Adodc 之資料
'
'*************************************************
Public Sub AdodcRefresh()
Dim strSql As String
Dim StrSQLa As String

On Error GoTo Checking
    strSql = ""
    If MaskEdBox2.Text <> "" And MaskEdBox2.Text <> MsgText(29) Then
        strSql = " And A3102 >= " & Val(FCDate(MaskEdBox2.Text)) & ""
    End If
    If MaskEdBox3.Text <> "" And MaskEdBox3.Text <> MsgText(29) Then
        strSql = strSql & " And A3102 <= " & Val(FCDate(MaskEdBox3.Text)) & ""
    End If
    If Me.Text1.Text <> "" Then
        'edit by nick 2004/12/30
        'StrSql = StrSql & " And A0K20='" & Me.Text1.Text & "' "
        strSql = strSql & " And A3121='" & Me.Text1.Text & "' "
    End If
    'add by nick 2005/01/05 增加收據抬頭用模糊
    If Me.Text2.Text <> "" Then
        strSql = strSql & " and instr(A3122, '" & Me.Text2.Text & "') > 0 "
    End If
   If adoadodc1.State = adStateOpen Then
        adoadodc1.Close
    End If
    adoadodc1.CursorLocation = adUseClient
    'edit by nick 2004/08/20 可以查他所 cancel
    'strSQLA = "Select ST02 As 收款人, Sum(A3105) As 現金, Sum(A3106) As 支票, Sum(Nvl(A3106,0) + Nvl(A3105,0)) As 合計, A0K20, ST06 From ACC310, ACC0k0, Staff Where A3103=A0K01(+) And A0K20=ST01(+) And A3101='" & pub_strUserOffice & "' " & strSQL & " Group By ST02, A0K20, ST06 "
    'strSQLA = strSQLA & " Union Select '總　計' As 收款人, Sum(A3105) As 現金, Sum(A3106) As 支票, Sum(Nvl(A3106,0) + Nvl(A3105,0)) As 合計, 'zzzzz' As A0K20, 'z' As ST06 From ACC310, ACC0K0 Where A3103=A0K01(+) And A3101='" & pub_strUserOffice & "' " & strSQL
    'strSQLA = strSQLA & " Order By 6, 5 "
    StrSQLa = "Select ST02 As 收款人, Sum(A3105) As 現金, Sum(A3106) As 支票, Sum(Nvl(A3106,0) + Nvl(A3105,0)) As 合計,sum(Round(A3123,3)) as 點數, A3121, nvl(ST06,0) From ACC310,  Staff Where  A3121=ST01(+)  And A3101='" & pub_strUserOffice & "' " & strSql & " Group By ST02, A3121, ST06 "
    StrSQLa = StrSQLa & " Union Select '總　計' As 收款人, Sum(A3105) As 現金, Sum(A3106) As 支票, Sum(Nvl(A3106,0) + Nvl(A3105,0)) As 合計,sum(Round(A3123,3)) as 點數,  'zzzzz' As A3121, 'z' As ST06 From ACC310 Where A3101='" & pub_strUserOffice & "' " & strSql
    StrSQLa = StrSQLa & " Order By 7, 6 "
    adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
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
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub MaskEdBox2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub MaskEdBox3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    KeyAscii = 0
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) And Text1 <> MsgText(802) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox3.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

Private Function ReConBNOurCaseNO(strCaseNo As String) As String
If strCaseNo <> "" Then
    ReConBNOurCaseNO = Replace(Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "-" & Right(Left(strCaseNo, Len(strCaseNo) - 3), 6) & "-" & Right(Left(strCaseNo, Len(strCaseNo) - 2), 1) & "-" & Right(strCaseNo, 2), "-0-00", "")
Else
    ReConBNOurCaseNO = ""
End If
End Function

Private Sub Text1_Validate(Cancel As Boolean)
    If Me.Text1.Text = "" Then Me.lblSalesName.Caption = "": Exit Sub
    Me.lblSalesName.Caption = GetStaffName(Me.Text1.Text)
    If Me.lblSalesName.Caption = "" Then
        MsgBox "智權人員輸入錯誤!!!", vbExclamation + vbOKOnly
        Cancel = True
    End If
    If Cancel = True Then Text1_GotFocus
End Sub
