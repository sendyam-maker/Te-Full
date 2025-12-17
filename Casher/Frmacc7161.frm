VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc7161 
   AutoRedraw      =   -1  'True
   Caption         =   "分所收款-智權人員繳款統計資料"
   ClientHeight    =   5020
   ClientLeft      =   1520
   ClientTop       =   2510
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5020
   ScaleWidth      =   8760
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc7161.frx":0000
      Height          =   4035
      Left            =   150
      TabIndex        =   6
      Top             =   960
      Width           =   8445
      _ExtentX        =   14887
      _ExtentY        =   7108
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   26
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
         Size            =   11
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
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
      BeginProperty Column02 
         DataField       =   "票據金額"
         Caption         =   "票據金額"
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
         DataField       =   "北所電匯"
         Caption         =   "北所電匯"
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
         DataField       =   "分所電匯"
         Caption         =   "分所電匯"
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
      BeginProperty Column06 
         DataField       =   "抵暫收款"
         Caption         =   "抵暫收款"
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
      BeginProperty Column07 
         DataField       =   "其他"
         Caption         =   "其他"
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
      BeginProperty Column08 
         DataField       =   "溢收款"
         Caption         =   "溢收款"
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
         DataField       =   "手續費"
         Caption         =   "手續費"
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
         DataField       =   "補扣繳"
         Caption         =   "補扣繳"
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
      BeginProperty Column11 
         DataField       =   "外幣"
         Caption         =   "外幣"
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
      BeginProperty Column12 
         DataField       =   "total00"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1069.795
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column12 
            Alignment       =   1
            Object.Visible         =   0   'False
            ColumnWidth     =   1399.748
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
      Left            =   1710
      TabIndex        =   2
      Top             =   570
      Width           =   975
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   1710
      TabIndex        =   0
      Top             =   210
      Width           =   1575
      _ExtentX        =   2787
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
      Left            =   3630
      TabIndex        =   1
      Top             =   210
      Width           =   1575
      _ExtentX        =   2787
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
      Top             =   810
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
   Begin MSForms.Label lblSalesName 
      Height          =   290
      Left            =   2880
      TabIndex        =   7
      Top             =   590
      Width           =   1725
      VariousPropertyBits=   19
      Size            =   "3043;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
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
      Caption         =   "出納確認日期"
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
      Width           =   1695
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
      Left            =   3390
      TabIndex        =   3
      Top             =   210
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc7161"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/1/2 Form2.0已修改(lblSalesName,DataGrid1改Fonts)
'Create by Lydia 2014/10/3 分所收款資料查詢-智權人員繳款(統計)
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
      Set Frmacc7161 = Nothing
        tool3_enabled
        MenuDisabled
        Frmacc7160.Show
        strFormName = "Frmacc7160"
      Exit Sub
   End If
   strExitControl = MsgText(602)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
'on error GoTo Checking
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

'on error GoTo Checking
    strSql = ""
    If MaskEdBox2.Text <> "" And MaskEdBox2.Text <> MsgText(29) Then
        strSql = " And A4413 >= " & Val(FCDate(MaskEdBox2.Text)) + 19110000 & ""
    End If
    If MaskEdBox3.Text <> "" And MaskEdBox3.Text <> MsgText(29) Then
        strSql = strSql & " And A4413 <= " & Val(FCDate(MaskEdBox3.Text)) + 19110000 & ""
    End If
    If Me.Text1.Text <> "" Then

        strSql = strSql & " And A4401='" & Me.Text1.Text & "' "
    End If
    If adoadodc1.State = adStateOpen Then
        adoadodc1.Close
    End If
    adoadodc1.CursorLocation = adUseClient
   
'StrSQLa = " Select nvl(ST06,0) as ST06,ST02 As 收款人,ST01, Sum(A4405) As 票據金額, Sum(A4406) As 北所電匯,"
'StrSQLa = StrSQLa & " sum(A4407) as 分所電匯, sum(A4408) as 現金, sum(A4409) as 抵暫收款, sum(A4410) as 溢收款, sum(A4411) as 手續費,"
'StrSQLa = StrSQLa & " sum(A4422) as 補扣繳, sum(A4426) as 外幣,Sum(A4405+A4406+A4407+A4408+A4409-A4410+A4411+A4422+A4426) total00  From ACC440,  Staff Where  A4401=ST01(+) and ST06 ='" & pub_strUserOffice & "' "
'StrSQLa = StrSQLa & strSql & " Group By ST06,ST02,ST01 Union Select 'z' as ST06,'總　計' As 收款人,'z' as st01, Sum(A4405) As 票據金額, Sum(A4406) As 北所電匯, "
'StrSQLa = StrSQLa & " sum(A4407) as 分所電匯, sum(A4408) as 現金, sum(A4409) as 抵暫收款, sum(A4410) as 溢收款, sum(A4411) as 手續費, "
'StrSQLa = StrSQLa & " sum(A4422) as 補扣繳, sum(A4426) as 外幣,Sum(A4405+A4406+A4407+A4408+A4409-A4410+A4411+A4422+A4426) total00  From ACC440,  Staff Where  A4401=ST01(+) and ST06 ='" & pub_strUserOffice & "' "
'StrSQLa = StrSQLa & strSql & " Order By 1,3 "
'---有些金額（如：外幣）可能是null，怕計算錯誤，改加總前一畫面產生的記錄
'Modified by Lydia 2015/02/04 + 點數(AXD06/1000)
'Added by Lydia 2015/07/15 +其他(A4430=R43120)
StrSQLa = " select nvl(ST06,0) as ST06,ST02 As 收款人,ST01,Sum(R43119) As 點數,Sum(R43109) As 票據金額, Sum(R43110) As 北所電匯, "
StrSQLa = StrSQLa & " sum(R43111) as 分所電匯, sum(R43112) as 現金, sum(R43113) as 抵暫收款, sum(R43120) as 其他, sum(R43114) as 溢收款, "
StrSQLa = StrSQLa & " sum(R43115) as 手續費, sum(R43116) as 補扣繳, sum(R43117) as 外幣,Sum(R43118) total00 "
StrSQLa = StrSQLa & " from ACCRPT431, STAFF where R43106=st01 and R43100='" & strUserNum & "' "
StrSQLa = StrSQLa & " group by st06,st02,st01 Union Select 'z' as ST06,'總　計' As 收款人,'z' as st01,Sum(R43119) As 點數, Sum(R43109) As 票據金額, Sum(R43110) As 北所電匯, "
StrSQLa = StrSQLa & " sum(R43111) as 分所電匯, sum(R43112) as 現金, sum(R43113) as 抵暫收款, sum(R43120) as 其他, sum(R43114) as 溢收款, "
StrSQLa = StrSQLa & " sum(R43115) as 手續費, sum(R43116) as 補扣繳, sum(R43117) as 外幣,Sum(R43118) total00 "
StrSQLa = StrSQLa & " From ACCRPT431, STAFF Where R43106=st01 and R43100='" & strUserNum & "' "
StrSQLa = StrSQLa & " Order By 1,3 "

    adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
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


Private Sub Text1_Validate(Cancel As Boolean)
    If Me.Text1.Text = "" Then Me.lblSalesName.Caption = "": Exit Sub
    Me.lblSalesName.Caption = GetStaffName(Me.Text1.Text)
    If Me.lblSalesName.Caption = "" Then
        MsgBox "智權人員輸入錯誤!!!", vbExclamation + vbOKOnly
        Cancel = True
    End If
    If Cancel = True Then Text1_GotFocus
End Sub




