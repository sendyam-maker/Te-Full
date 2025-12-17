VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Frmacc7101 
   AutoRedraw      =   -1  'True
   Caption         =   "分所收款資料查詢"
   ClientHeight    =   5025
   ClientLeft      =   5295
   ClientTop       =   2250
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   8760
   Begin VB.ComboBox Combo3 
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
      Left            =   5760
      TabIndex        =   2
      Top             =   240
      Width           =   2772
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2772
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc7101.frx":0000
      Height          =   4092
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   7223
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
      ColumnCount     =   16
      BeginProperty Column00 
         DataField       =   "收款日"
         Caption         =   "收款日"
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
         DataField       =   "電腦收據"
         Caption         =   "電腦收據"
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
         DataField       =   "人工收據"
         Caption         =   "人工收據"
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
      BeginProperty Column04 
         DataField       =   "收據抬頭"
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
      BeginProperty Column05 
         DataField       =   "案件性質"
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
      BeginProperty Column06 
         DataField       =   "點數"
         Caption         =   "點數"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "現金"
         Caption         =   "現金"
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
         DataField       =   "支票"
         Caption         =   "支票"
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
         DataField       =   "到期日"
         Caption         =   "到期日"
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
         DataField       =   "帳號"
         Caption         =   "帳號"
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
      BeginProperty Column11 
         DataField       =   "票號"
         Caption         =   "票號"
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
      BeginProperty Column12 
         DataField       =   "付款地"
         Caption         =   "付款地"
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
      BeginProperty Column13 
         DataField       =   "扣繳日"
         Caption         =   "扣繳日"
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
      BeginProperty Column14 
         DataField       =   "扣繳金額"
         Caption         =   "扣繳金額"
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
      BeginProperty Column15 
         DataField       =   "留分所金額"
         Caption         =   "留分所金額"
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
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1335.118
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
            Alignment       =   2
         EndProperty
         BeginProperty Column14 
            Alignment       =   1
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   312
      Left            =   240
      Top             =   600
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
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5520
      TabIndex        =   4
      Top             =   240
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc7101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit
Public adoacc310 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset

Private Sub Combo1_Change()
'   SelectScope
End Sub

Private Sub Combo1_Click()
'   SelectScope
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Combo3 = Combo2
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
   strCon1 = "收款日"
   strCon2 = "電腦收據"
   strCon3 = "人工收據"
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1 = MsgText(31)
   SelectScope
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandler
       
    If Adodc1.Recordset.RecordCount <> 0 Then
        strItemNo = Adodc1.Recordset.Fields("電腦收據").Value & "," & Adodc1.Recordset.Fields("人工收據").Value
    Else
       strItemNo = MsgText(601)
    End If
    StatusClear
    tool1_enabled
    Frmacc7100.Enabled = True
    Frmacc7100.Show
    strFormName = "Frmacc7100"
    Set Frmacc7101 = Nothing
Exit Sub
ErrorHandler:
    strItemNo = MsgText(601)
    Resume Next
End Sub

'*************************************************
'  搜尋條件範圍值，並代入 Combo2、Combo3 之中
'
'*************************************************
Private Sub SelectScope()
   strCondition = MsgText(601)
   If Combo1 = MsgText(31) Then
      Exit Sub
   End If
   Select Case Combo1
      Case strCon1
         strCondition = "A0302"
      Case strCon2
         strCondition = "A0303"
      Case strCon3
         strCondition = "A0304"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc310.CursorLocation = adUseClient
   'edit by nick 2004/08/20  可查分所
   'adoacc310.Open "Select Distinct " & strCondition & " From ACC310 Where A3101='" & pub_strUserOffice & "' Order By " & strCondition & " ", adoTaie, adOpenStatic, adLockReadOnly
   adoacc310.Open "Select Distinct " & strCondition & " From ACC310 Order By " & strCondition & " ", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc310.EOF = False
      If IsNull(adoacc310.Fields(0).Value) = False Then
         Combo2.AddItem adoacc310.Fields(0).Value
         Combo3.AddItem adoacc310.Fields(0).Value
      End If
      adoacc310.MoveNext
   Loop
   adoacc310.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Screen.MousePointer = vbHourglass
         Acc310Query
         Screen.MousePointer = vbDefault
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  分所收款資料查詢
'
'*************************************************
Private Sub Acc310Query()
On Error GoTo Checking
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1
         strCondition = "A3102"
      Case strCon2
         strCondition = "A3103"
      Case strCon3
         strCondition = "A3104"
      Case MsgText(31)
         'edit by nick 2004/08/20  可查分所
         'adoadodc1.Open "Select A3102 As 收款日, A3103 As 電腦收據, A3104 As 人工收據, ST02 As 收款人, A0K04 As 收據抬頭, A0J20 As 案件性質, Round(Nvl(A0J09,0)/1000,1) As 點數, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳金額, A3113 As 留分所金額 From ACC310, ACC0K0, ACC0J0, Staff Where A3103=A0K01(+) And A0K01=A0J13(+) And A0K20=ST01(+) And A3101='" & pub_strUserOffice & "' ", adoTaie, adOpenStatic, adLockReadOnly
         'Modified by Morgan 2011/12/27 取消 a0j20
         adoadodc1.Open "Select A3102 As 收款日, A3103 As 電腦收據, A3104 As 人工收據, ST02 As 收款人, A3122 As 收據抬頭, getcp10desc(cp01,cp10,a0j04) As 案件性質, Round(Nvl(A0J09,0)/1000,3) As 點數, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳金額, A3113 As 留分所金額 From ACC310,  ACC0J0, Staff,caseprogress Where A3103=A0J13(+) And A3121=ST01(+) and cp09(+)=a0j01", adoTaie, adOpenStatic, adLockReadOnly
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   If Combo3 = MsgText(601) Then
      'edit by nick 2004/08/20  可查分所
      'adoadodc1.Open "Select A3102 As 收款日, A3103 As 電腦收據, A3104 As 人工收據, ST02 As 收款人, A0K04 As 收據抬頭, A0J20 As 案件性質, Round(Nvl(A0J09,0)/1000,1) As 點數, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳金額, A3113 As 留分所金額 From ACC310, ACC0K0, ACC0J0, Staff Where A3103=A0K01(+) And A0K01=A0J13(+) And A0K20=ST01(+) And A3101='" & pub_strUserOffice & "' And " & strCondition & " = '" & Combo2 & "' Order By A3102, A3103, A3104 ", adoTaie, adOpenStatic, adLockReadOnly
      'Modified by Morgan 2011/12/27 取消 a0j20
      adoadodc1.Open "Select A3102 As 收款日, A3103 As 電腦收據, A3104 As 人工收據, ST02 As 收款人, A3122 As 收據抬頭, getcp10desc(cp01,cp10,a0j04) As 案件性質, Round(Nvl(A0J09,0)/1000,3) As 點數, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳金額, A3113 As 留分所金額 From ACC310, ACC0J0, Staff,caseprogress Where A3103=A0J13(+) And A3121=ST01(+) And " & strCondition & " = '" & Combo2 & "' and cp09(+)=a0j01 Order By A3102, A3103, A3104 ", adoTaie, adOpenStatic, adLockReadOnly
   Else
      If Combo2 = MsgText(601) Then
         'edit by nick 2004/08/20  可查分所
         'adoadodc1.Open "Select A3102 As 收款日, A3103 As 電腦收據, A3104 As 人工收據, ST02 As 收款人, A0K04 As 收據抬頭, A0J20 As 案件性質, Round(Nvl(A0J09,0)/1000,1) As 點數, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳金額, A3113 As 留分所金額 From ACC310, ACC0K0, ACC0J0, Staff Where A3103=A0K01(+) And A0K01=A0J13(+) And A0K20=ST01(+) And A3101='" & pub_strUserOffice & "' And " & strCondition & " <= '" & Combo3 & "' Order By A3102, A3103, A3104 ", adoTaie, adOpenStatic, adLockReadOnly
         'Modified by Morgan 2011/12/27 取消 a0j20
         adoadodc1.Open "Select A3102 As 收款日, A3103 As 電腦收據, A3104 As 人工收據, ST02 As 收款人, A3122 As 收據抬頭, getcp10desc(cp01,cp10,a0j04) As 案件性質, Round(Nvl(A0J09,0)/1000,3) As 點數, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳金額, A3113 As 留分所金額 From ACC310,  ACC0J0, Staff,caseprogress Where A3103=A0J13(+) And A3121=ST01(+) And " & strCondition & " <= '" & Combo3 & "' and cp09(+)=a0j01 Order By A3102, A3103, A3104 ", adoTaie, adOpenStatic, adLockReadOnly
      Else
         'edit by nick 2004/08/20  可查分所
         'adoadodc1.Open "Select A3102 As 收款日, A3103 As 電腦收據, A3104 As 人工收據, ST02 As 收款人, A0K04 As 收據抬頭, A0J20 As 案件性質, Round(Nvl(A0J09,0)/1000,1) As 點數, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳金額, A3113 As 留分所金額 From ACC310, ACC0K0, ACC0J0, Staff Where A3103=A0K01(+) And A0K01=A0J13(+) And A0K20=ST01(+) And A3101='" & pub_strUserOffice & "' And " & strCondition & " >= '" & Combo2 & "' And " & strCondition & " <= '" & Combo3 & "' Order By A3102, A3103, A3104 ", adoTaie, adOpenStatic, adLockReadOnly
         'Modified by Morgan 2011/12/27 取消 a0j20
         adoadodc1.Open "Select A3102 As 收款日, A3103 As 電腦收據, A3104 As 人工收據, ST02 As 收款人, A3122 As 收據抬頭, getcp10desc(cp01,cp10,a0j04) As 案件性質, Round(Nvl(A0J09,0)/1000,3) As 點數, A3105 As 現金, A3106 As 支票, A3107 As 到期日, A3108 As 帳號, A3109 As 票號, A3110 As 付款地, A3111 As 扣繳日, A3112 As 扣繳金額, A3113 As 留分所金額 From ACC310,  ACC0J0, Staff,caseprogress Where A3103=A0J13(+) And A3121=ST01(+) And " & strCondition & " >= '" & Combo2 & "' And " & strCondition & " <= '" & Combo3 & "' and cp09(+)=a0j01 Order By A3102, A3103, A3104 ", adoTaie, adOpenStatic, adLockReadOnly
      End If
   End If
   Adodc1.Recordset.Requery
   If Adodc1.Recordset.RecordCount = 0 Then
      MsgBox MsgText(33), , MsgText(5)
   End If
   Exit Sub
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   'edit by nick 2004/08/20  可查分所
   'adoadodc1.Open "Select * From ACC310 Where A3101='" & pub_strUserOffice & "' And A3102 = '" & Combo1 & "' Order By A3102, A3103, A3104 ", adoTaie, adOpenStatic, adLockReadOnly
   adoadodc1.Open "Select * From ACC310 Where A3102 = '" & Combo1 & "' Order By A3102, A3103, A3104 ", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub



