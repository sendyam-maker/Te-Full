VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc2230 
   AutoRedraw      =   -1  'True
   Caption         =   "國外請款金額查詢"
   ClientHeight    =   5244
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8772
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5244
   ScaleWidth      =   8772
   Begin VB.CheckBox Check1 
      Caption         =   "已扣除折讓金額"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   630
      TabIndex        =   3
      Top             =   450
      Width           =   1755
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      TabIndex        =   2
      Top             =   90
      Width           =   852
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc2230.frx":0000
      Height          =   3870
      Left            =   240
      TabIndex        =   4
      Top             =   1095
      Width           =   8295
      _ExtentX        =   14626
      _ExtentY        =   6816
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   17
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
         Name            =   "新細明體-ExtB"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "國外請款金額查詢"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "a1k01"
         Caption         =   "請款單號"
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
         DataField       =   "a1k03"
         Caption         =   "代理人"
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
         DataField       =   "代理人名稱"
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
      BeginProperty Column03 
         DataField       =   "a1k1316"
         Caption         =   "本所案號"
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
         DataField       =   "a1k02"
         Caption         =   "請款日期"
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
      BeginProperty Column05 
         DataField       =   "a1k18"
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
         DataField       =   "a1k11"
         Caption         =   "請款台幣金額"
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
      BeginProperty Column07 
         DataField       =   "a1k08"
         Caption         =   "請款外幣金額"
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
         DataField       =   "Damount"
         Caption         =   "折讓金額"
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
      BeginProperty Column09 
         DataField       =   "a1k09"
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
      BeginProperty Column10 
         DataField       =   "a1k30"
         Caption         =   "已收金額"
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
      BeginProperty Column11 
         DataField       =   "a1k29"
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
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1152
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1008
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   515.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1512
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1488.189
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   1008
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            ColumnWidth     =   1307.906
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnWidth     =   659.906
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Top             =   90
      Width           =   1572
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   90
      Width           =   1572
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   225
      Top             =   945
      Visible         =   0   'False
      Width           =   1200
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
   Begin VB.Label Label18 
      BackStyle       =   0  '透明
      Caption         =   "(*表示作廢、@表示有折讓、$表示銷帳)"
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
      Left            =   330
      TabIndex        =   8
      Top             =   765
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "代理人國籍："
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
      Left            =   5640
      TabIndex        =   7
      Top             =   90
      Width           =   1215
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
      Left            =   3480
      TabIndex        =   6
      Top             =   90
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款外幣金額："
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
      Left            =   360
      TabIndex        =   5
      Top             =   90
      Width           =   1455
   End
End
Attribute VB_Name = "Frmacc2230"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; DataGrid1改字型=新細明體-ExtB
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/30 日期欄已修改
Option Explicit

Public adoadodc1 As New ADODB.Recordset
Dim strSql As String


Private Sub Form_Activate()
'edit by nickc 2007/02/08
'   '93.3.16 ADD BY SONIA
'   If IsObject(mdiMain) Then
'      mdiMain.toolshow
'   End If
'   '93.3.16 END
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyDefine KeyCode
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
'   Me.Height = 5400
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath2)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   strFormName = Name
   'Modify by Amy 2023/10/11 原8850, 5500
   PUB_InitForm Me, 8895, 5715, strBackPicPath2
   'end 2021/12/09
   
   OpenTable
   StatusView MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc2230 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and a1k01 = 'X' order by a1k01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Set Adodc1.Recordset = adoadodc1
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  查詢資料表(傳票資料)
'
'*************************************************
Public Sub QueryTable()
Dim strSql As String
'Add By Cheng 2003/08/18
Dim StrSQLa As String
Dim strSQL1 As String
   
On Error GoTo Checking
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/22 清除查詢印表記錄檔欄位
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   strSql = ""
   strSQL1 = ""
   'Add By Sindy 2013/1/14
   If Check1.Value = 1 Then '已扣除折讓金額
      pub_QL05 = pub_QL05 & ";" & Check1.Caption
      If Text1 <> MsgText(601) Then
         strSql = " and (a1k08-nvl(a1k31,0)) >= " & Val(Text1) & ""
         strSQL1 = " and (a1k08-nvl(a1k31,0)) >= " & Val(Text1) & ""
      End If
      If Text2 <> MsgText(601) Then
         strSql = strSql & " and (a1k08-nvl(a1k31,0)) <= " & Val(Text2) & ""
         strSQL1 = strSQL1 & " and (a1k08-nvl(a1k31,0)) <= " & Val(Text2) & ""
      End If
   Else
   '2013/1/14 End
      If Text1 <> MsgText(601) Then
         strSql = " and a1k08 >= " & Val(Text1) & ""
         strSQL1 = " and a1k08 >= " & Val(Text1) & ""
      End If
      If Text2 <> MsgText(601) Then
         strSql = strSql & " and a1k08 <= " & Val(Text2) & ""
         strSQL1 = strSQL1 & " and a1k08 <= " & Val(Text2) & ""
      End If
   End If
   If Text1 <> MsgText(601) Or Text2 <> MsgText(601) Then
      pub_QL05 = pub_QL05 & ";" & Label1 & Text1 & "-" & Text2 'Add By Sindy 2010/12/22
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and instr(fa10, '" & Text3 & "') = 1"
      strSQL1 = strSQL1 & " and instr(cu10, '" & Text3 & "') = 1"
      pub_QL05 = pub_QL05 & ";" & Label3 & Text3 'Add By Sindy 2010/12/22
   End If
   'Modify By Cheng 2002/09/20
   '增加顯示本所案號
'   adoadodc1.Open "select a1k01, a1k03, a1k02, a1k11, (nvl(a1k06, 0) * nvl(a1k10, 1)) as Damount, nvl(a1k09, 0) as a1k09, a1k30, a1k29, a1k08 from acc1k0, fagent where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+)" & strSQL & " order by a1k01 asc", adoTaie, adOpenStatic, adLockReadOnly
    'Modify By Cheng 2003/08/18
'   adoadodc1.Open "select decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as a1k01, a1k03, a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k02, a1k11, (nvl(a1k06, 0) * nvl(a1k10, 1)) as Damount, nvl(a1k09, 0) as a1k09, a1k30, a1k29, a1k08 from acc1k0, fagent, acc140 where substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+) and a1k01 = a1403 (+)" & strSQL & " order by a1k01 asc", adoTaie, adOpenStatic, adLockReadOnly
    'Modify By Sindy 2012/12/7 (nvl(a1k06, 0) * nvl(a1k10, 1)) as Damount ==> nvl(a1k06, 0) as Damount
'    StrSQLa = "select decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as a1k01, a1k03, Decode(FA04, Null, Decode(FA05, Null, FA06, FA05||' '||FA63||' '||FA64||' '||FA65), FA04) As 代理人名稱 , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k02, a1k11, (nvl(a1k06, 0) * nvl(a1k10, 1)) as Damount, nvl(a1k09, 0) as a1k09, a1k30, a1k29, a1k08 from acc1k0, fagent, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and a1k01 = a1403 (+)" & strSql
'    StrSQLa = StrSQLa & " Union select decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as a1k01, a1k03, Decode(CU04, Null, Decode(CU05, Null, CU06, CU05||' '||CU88||' '||CU89||' '||CU90), CU04) As 代理人名稱 , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k02, a1k11, (nvl(a1k06, 0) * nvl(a1k10, 1)) as Damount, nvl(a1k09, 0) as a1k09, a1k30, a1k29, a1k08 from acc1k0, Customer, acc140 where substr(a1k03, 1, 8) = CU01 and substr(a1k03, 9, 1) = CU02 and a1k01 = a1403 (+)" & strSQL1
    'Modify By Sindy 2013/1/14 +, a1k18
    'Modified by Lydia 2025/06/24 +a1k08
    StrSQLa = "select decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as a1k01, a1k03, Decode(FA04, Null, Decode(FA05, Null, FA06, FA05||' '||FA63||' '||FA64||' '||FA65), FA04) As 代理人名稱 , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k02, a1k18, a1k11,a1k08, nvl(a1k06, 0) as Damount, nvl(a1k09, 0) as a1k09, a1k30, a1k29, a1k08 from acc1k0, fagent, acc140 where substr(a1k03, 1, 8) = fa01 and substr(a1k03, 9, 1) = fa02 and a1k01 = a1403 (+)" & strSql
    StrSQLa = StrSQLa & " Union select decode(a1k12, null, decode(a1k07, null, decode(a1401, null, a1k01, a1k01||'$'), a1k01||'@'), a1k01||'*') as a1k01, a1k03, Decode(CU04, Null, Decode(CU05, Null, CU06, CU05||' '||CU88||' '||CU89||' '||CU90), CU04) As 代理人名稱 , a1k13||'-'||a1k14||'-'||a1k15||'-'||a1k16 as a1k1316, a1k02, a1k18, a1k11,a1k08, nvl(a1k06, 0) as Damount, nvl(a1k09, 0) as a1k09, a1k30, a1k29, a1k08 from acc1k0, Customer, acc140 where substr(a1k03, 1, 8) = CU01 and substr(a1k03, 9, 1) = CU02 and a1k01 = a1403 (+)" & strSQL1
    '2012/12/7 End
    StrSQLa = StrSQLa & " Order By 1 "
   adoadodc1.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   Adodc1.Recordset.ReQuery
   If Adodc1.Recordset.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/22
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   Else
      InsertQueryLog (Adodc1.Recordset.RecordCount) 'Add By Sindy 2010/12/22
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
            QueryTable
            Screen.MousePointer = vbDefault
            Exit Sub
         Else
            MsgBox MsgText(181), , MsgText(5)
         End If
   End Select
   KeyEnter KeyCode
   StatusView MsgText(98)
End Sub

'*************************************************
'  畫面輸入檢查
'
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
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function
