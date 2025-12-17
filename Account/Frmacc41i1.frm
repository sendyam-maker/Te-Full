VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc41i1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "財產目錄資料查詢"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8730
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
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2052
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc41i1.frx":0000
      Height          =   4095
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   11.25
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "a2b01"
         Caption         =   "編號"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "a2b16"
         Caption         =   "公司別"
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
         DataField       =   "a2b05"
         Caption         =   "取得日期"
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
         DataField       =   "a2b02t"
         Caption         =   "類別"
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
         DataField       =   "a2b03t"
         Caption         =   "所在地"
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
      BeginProperty Column05 
         DataField       =   "a2b04"
         Caption         =   "財產名稱"
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
         DataField       =   "a2b19"
         Caption         =   "報廢日期"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         Size            =   275
         BeginProperty Column00 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2190.047
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1035.213
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
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
   Begin MSForms.ComboBox Combo3 
      Height          =   330
      Left            =   5700
      TabIndex        =   2
      Top             =   240
      Width           =   2775
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4895;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   330
      Left            =   2670
      TabIndex        =   1
      Top             =   240
      Width           =   2775
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4895;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
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
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   240
      Width           =   135
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
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc41i1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/1 改成Form2.0 ; Combo2、Combo3、DataGrid1改字型=新細明體-ExtB
'Create by Lydia 2017/05/19 財產目錄/報廢資料查詢
Option Explicit
Public adoacc2b0 As New ADODB.Recordset
Public adoadodc1 As New ADODB.Recordset
Public iType As String  '1 = 財產目錄 ,2 =財產報廢

'Modified by Lydia 2021/12/01 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Combo2_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   Combo3 = Combo2
End Sub

'Modified by Lydia 2021/12/01 改成Form 2.0; KeyAscii As Integer=>MSForms.ReturnInteger
Private Sub Combo3_KeyPress(KeyAscii As MSForms.ReturnInteger)
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
   
   If iType = "2" Then
      Me.Caption = "財產報廢資料查詢"
   End If
   
   '表單初始化
   PUB_InitForm Me, 8850, 5400, strBackPicPath2
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name

   strCon1 = "公司別"
   If iType = "1" Then
      strCon2 = "取得日期"
   Else
      strCon2 = "報廢日期"
   End If
   strCon3 = "類別"
   strCon4 = "所在地"
   Combo1.AddItem MsgText(31)
   Combo1.AddItem strCon1
   Combo1.AddItem strCon2
   Combo1.AddItem strCon3
   Combo1.AddItem strCon4
   Combo1 = MsgText(31)
   SelectScope
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Adodc1.Recordset.RecordCount <> 0 Then
      strCompanyNo = Adodc1.Recordset.Fields("a2b16").Value
      strItemNo = Adodc1.Recordset.Fields("a2b01").Value
   Else
      strCompanyNo = MsgText(601)
      strItemNo = MsgText(601)
   End If
   StatusClear
   tool1_enabled
   
   If iType = "1" Then
      Frmacc41i0.Enabled = True
      Frmacc41i0.Show
   Else
      Frmacc41i0_1.Enabled = True
      Frmacc41i0_1.Show
   End If
   Set Frmacc41i1 = Nothing
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
      Case strCon1 '公司別
         strCondition = "a2b16"
      Case strCon2 '取得/報廢日期
         If iType = "1" Then
            strCondition = "a2b05"
         Else
            strCondition = "a2b19"
         End If
      Case strCon3 '類別
         strCondition = "a2b02"
      Case strCon4 '所在地
         strCondition = "a2b03"
   End Select
   If strCondition = MsgText(601) Then
      Exit Sub
   End If
   Combo2.Clear
   Combo3.Clear
   adoacc2b0.CursorLocation = adUseClient
   adoacc2b0.Open "select distinct " & strCondition & " from acc2b0 order by " & strCondition & " asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoacc2b0.EOF = False
      If IsNull(adoacc2b0.Fields(0).Value) = False Then
         Combo2.AddItem adoacc2b0.Fields(0).Value
         Combo3.AddItem adoacc2b0.Fields(0).Value
      End If
      adoacc2b0.MoveNext
   Loop
   adoacc2b0.Close
End Sub

'*************************************************
'  功能鍵定義
'
'*************************************************
Private Sub KeyDefine(KeyCode As Integer)
   Select Case KeyCode
      Case vbKeyF12
         Screen.MousePointer = vbHourglass
         Acc2b0Query
         Screen.MousePointer = vbDefault
   End Select
   KeyEnter KeyCode
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

'*************************************************
'  傳票資料查詢
'
'*************************************************
Private Sub Acc2b0Query()
Dim stSQL As String
On Error GoTo Checking

   stSQL = "select a.*,decode(a2b02,'1','交通運輸設備','2','生財器具','3','電腦硬體','4','電腦軟體',a2b02) a2b02t," & _
           "decode(a2b03,'1','北所','2','中所','3','南所','4','高所','5','其他',a2b03) a2b03t " & _
           "from acc2b0 a where 1=1 " & IIf(iType = "1", "", " and nvl(a2b19,0) > 0 ")
           
   adoadodc1.Close
   adoadodc1.CursorLocation = adUseClient
   Select Case Combo1
      Case strCon1 '公司別
         strCondition = "a2b16"
      Case strCon2 '取得/報廢日期
         If iType = "1" Then
            strCondition = "a2b05"
         Else
            strCondition = "a2b19"
         End If
      Case strCon3 '類別
         strCondition = "a2b02"
      Case strCon4 '所在地
         strCondition = "a2b03"
      Case MsgText(31) '全部
         If iType = "1" Then
            strExc(1) = stSQL & " order by a2b01 asc "
         Else
            strExc(1) = stSQL & " and nvl(a2b19,0) > 0 order by a2b01 asc"
         End If
         adoadodc1.Open strExc(1), adoTaie, adOpenStatic, adLockReadOnly
         Adodc1.Recordset.Requery
         Exit Sub
      Case Else
         Exit Sub
   End Select
   If Combo3 = MsgText(601) Then
      If Combo1 = strCon1 Then
         strExc(1) = stSQL & " and " & strCondition & " = '" & Combo2 & "' order by " & strCondition & " asc, a2b01 asc"
      Else
         strExc(1) = stSQL & " and " & strCondition & " = " & Val(Combo2) & " order by " & strCondition & " asc, a2b01 asc"
      End If
   Else
      If Combo2 = MsgText(601) Then
         If Combo1 = strCon1 Then
            strExc(1) = stSQL & " and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc, a2b01 asc"
         Else
            strExc(1) = stSQL & " and " & strCondition & " <= " & Val(Combo3) & " order by " & strCondition & " asc, a2b01 asc"
         End If
      Else
         If Combo1 = strCon1 Then
            strExc(1) = stSQL & " and " & strCondition & " >= '" & Combo2 & "' and " & strCondition & " <= '" & Combo3 & "' order by " & strCondition & " asc, a2b01 asc"
         Else
            strExc(1) = stSQL & " and " & strCondition & " >= '" & Val(Combo2) & "' and " & strCondition & " <= '" & Val(Combo3) & "' order by " & strCondition & " asc, a2b01 asc"
         End If
      End If
   End If
   adoadodc1.Open strExc(1), adoTaie, adOpenStatic, adLockReadOnly
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

   strExc(1) = "select a.*,decode(a2b02,'1','交通運輸設備','2','生財器具','3','電腦硬體','4','電腦軟體',a2b02) a2b02t," & _
               "decode(a2b03,'1','北所','2','中所','3','南所','4','高所','5','其他',a2b03) a2b03t " & _
               "from acc2b0 a where a2b01= '" & Combo2 & "' " & IIf(iType = "1", "", " and nvl(a2b19,0) > 0 ") & " order by a2b01 asc"
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open strExc(1), adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub


