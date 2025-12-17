VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmacc32f0 
   AutoRedraw      =   -1  'True
   Caption         =   "兌現日別票據明細查詢"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5520
   ScaleWidth      =   9255
   Begin VB.ComboBox CboCmp 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   180
      Width           =   3500
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  '靠右對齊
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
      Height          =   315
      Left            =   5640
      TabIndex        =   14
      Top             =   4910
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  '靠右對齊
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
      Height          =   315
      Left            =   3285
      TabIndex        =   12
      Top             =   4910
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   5
      Top             =   950
      Width           =   1575
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
      Height          =   315
      Left            =   6720
      TabIndex        =   2
      Top             =   590
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Frmacc32f0.frx":0000
      Height          =   3435
      Left            =   120
      TabIndex        =   6
      Top             =   1395
      Width           =   8750
      _ExtentX        =   15425
      _ExtentY        =   6059
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
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
      Caption         =   "兌現日別票據明細資料"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "a0e02"
         Caption         =   "票據號碼"
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
         DataField       =   "a0e08"
         Caption         =   "票別"
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
         DataField       =   "a0e04"
         Caption         =   "應收/付別"
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
      BeginProperty Column03 
         DataField       =   "a0e11"
         Caption         =   "票據金額"
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
      BeginProperty Column04 
         DataField       =   "a0e13"
         Caption         =   "開票日期"
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
         DataField       =   "a0e10"
         Caption         =   "到期日期"
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
      BeginProperty Column06 
         DataField       =   "a0e21"
         Caption         =   "兌現日期"
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
      BeginProperty Column07 
         DataField       =   "contect"
         Caption         =   "往來對象"
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
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   4724.788
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
      Left            =   1200
      TabIndex        =   1
      Top             =   590
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   945
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3120
      TabIndex        =   4
      Top             =   945
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
      Left            =   120
      Top             =   1275
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
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   225
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "共               筆"
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
      Left            =   5280
      TabIndex        =   15
      Top             =   4905
      Width           =   1575
   End
   Begin VB.Label Label10 
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
      Height          =   255
      Left            =   2565
      TabIndex        =   13
      Top             =   4905
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "銀行帳號"
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
      Left            =   5640
      TabIndex        =   11
      Top             =   585
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "應收/付"
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
      Left            =   5640
      TabIndex        =   10
      Top             =   945
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   -120
      Top             =   4800
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "銀行別"
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
      Top             =   585
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   1215
      Left            =   100
      Top             =   120
      Width           =   8750
   End
   Begin VB.Label Label4 
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
      Left            =   2880
      TabIndex        =   8
      Top             =   945
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "兌現日期"
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
      TabIndex        =   7
      Top             =   945
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc32f0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/07 Form2.0已修改 DataGrid1
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/3 日期欄已修改
Option Explicit

Public adoadodc1 As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Dim strSql As String
'Add By Cheng 2002/01/30
Dim m_dbl_RsCnts As Double '筆數
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/04/17


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
   'Modify by Amy 2023/10/11 避免切畫面仍要調整,故調大小及位置,原W9500 H5860,(lngWidth - Me.Width) / 2-瑞婷
   Me.Width = 9375
   '20140122START Modify By eric
   Me.Height = 5985 'Modify by Amy 2023/08/18 原: 5700
   'Me.Height = 5500
   '20140122END
   Me.Move 0, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath2)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   'Add by Sindy 2020/04/17 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/04/17
   
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Combo1.AddItem ComboItem(181)
   Combo1.AddItem ComboItem(182)
   OpenTable
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(98)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StatusClear
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   Set Frmacc32f0 = Nothing
End Sub

'*************************************************
'  開啟資料表
'
'*************************************************
Private Sub OpenTable()
On Error GoTo Checking
   adoadodc1.CursorLocation = adUseClient
   adoadodc1.Open "select * from acc0e0 where a0e21 >= " & Val(FCDate(MaskEdBox1.Text)) & " and a0e21 <= " & Val(FCDate(MaskEdBox2.Text)) & " and a0e19 = '" & Text1 & "' order by a0e02 asc", adoTaie, adOpenStatic, adLockReadOnly
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
Dim strUnion As String
'Add By Cheng 2002/01/30
Dim strSQLR As String
Dim strSQLP As String

On Error GoTo Checking
   strSql = ""
   If adoadodc1.State = adStateOpen Then
      adoadodc1.Close
   End If
   adoadodc1.CursorLocation = adUseClient
   
   Call SetCompN 'Add by Sindy 2020/04/17
   
   '選擇應收或應付票據
   If Combo1 <> "" Then
      strSql = strSql & " and a0e04 = '" & Mid(Combo1, 1, 1) & "'"
      Select Case Mid(Combo1, 1, 1)
         '應收票據
         Case "R"
            '20140122START Add By eric
            'Modify By Sindy 2020/4/17 公司別改變數
'            If Text5 <> "" Then
'               strSql = strSql & " and a0e23 = '" & IIf(Text5 = "2", "J", "1") & "'"
'            End If
            If strCmp <> MsgText(601) Then
                If InStr(strCmp, "+") > 0 Then
                   strSql = strSql & " And (a0e23 is null or a0e23 In ('" & Replace(strCmp, "+", "','") & "')) "
                Else
                   strSql = strSql & " And (a0e23 is null or a0e23='" & strCmp & "') "
                End If
            End If
            '2020/4/17 END
            If Text1 <> MsgText(601) Then
               strSql = strSql & " and a0e19 = '" & Text1 & "'"
            End If
            'If Text1 <> MsgText(601) Then
            '   strSql = " and a0e19 = '" & Text1 & "'"
            'End If
            '20140122END
                        
            If Text2 <> MsgText(601) Then
               strSql = strSql & " and a0e20 = '" & Text2 & "'"
            End If
            
            'Modify By Cheng 2002/01/30
'            strSQL = strSQL & " and a0e04 = 'R'"
            
            If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
               strSql = strSql & " and (a0e21 <> 0 AND A0E21 IS NOT NULL) and a0e21 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
            End If
            If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
               strSql = strSql & " and (a0e21 <> 0 AND A0E21 IS NOT NULL) and a0e21 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
            End If
         '應付票據
         Case "P"
            '20140122START Add By eric
            'Modify By Sindy 2020/4/17 公司別改變數
'            If Text5 <> "" Then
'               strSql = " and a0e23 = '" & IIf(Text5 = "2", "J", "1") & "'"
'            End If
            If strCmp <> MsgText(601) Then
                If InStr(strCmp, "+") > 0 Then
                   strSql = " And (a0e23 is null or a0e23 In ('" & Replace(strCmp, "+", "','") & "')) "
                Else
                   strSql = " And (a0e23 is null or a0e23='" & strCmp & "') "
                End If
            End If
            '2020/4/17 END
            If Text1 <> MsgText(601) Then
               strSql = strSql & " and a0e01 = '" & Text1 & "'"
            End If
            'If Text1 <> MsgText(601) Then
            '   strSql = " and a0e01 = '" & Text1 & "'"
            'End If
            '20140122END
            
            If Text2 <> MsgText(601) Then
               strSql = strSql & " and a0e07 = '" & Text2 & "'"
            End If
            
            'Modify By Cheng 2002/01/30
'            strSQL = strSQL & " and a0e04 = 'P'"

            If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
               'Modify By Cheng 2002/01/30
'               strSQL = strSQL & " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
               strSql = strSql & " and a0e37 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
            End If
            If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
               'Modify By Cheng 2002/01/30
'               strSQL = strSQL & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
               strSql = strSql & " and a0e37 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
            End If
      End Select
   '應收與應付一起查
   Else
      strSQLR = "": strSQLP = ""
      '20140122START Add By eric
      'Modify By Sindy 2020/4/17 公司別改變數
'      If Text5 <> "" Then
'         strSQLR = strSQLR & " and a0e23 = '" & IIf(Text5 = "2", "J", "1") & "'"
'      End If
      If strCmp <> MsgText(601) Then
          If InStr(strCmp, "+") > 0 Then
             strSQLR = strSQLR & " And (a0e23 is null or a0e23 In ('" & Replace(strCmp, "+", "','") & "')) "
          Else
             strSQLR = strSQLR & " And (a0e23 is null or a0e23='" & strCmp & "') "
          End If
      End If
      '2020/4/17 END
      '20140122END
      If Text1 <> MsgText(601) Then
         strSQLR = strSQLR & " and a0e19 = '" & Text1 & "'"
      End If
      If Text2 <> MsgText(601) Then
         strSQLR = strSQLR & " and a0e20 = '" & Text2 & "'"
      End If
      
      '20140122START Add By eric
      'Modify By Sindy 2020/4/17 公司別改變數
'      If Text5 <> "" Then
'         strSQLP = strSQLP & " and a0e23 = '" & IIf(Text5 = "2", "J", "1") & "'"
'      End If
      If strCmp <> MsgText(601) Then
          If InStr(strCmp, "+") > 0 Then
             strSQLP = strSQLP & " And (a0e23 is null or a0e23 In ('" & Replace(strCmp, "+", "','") & "')) "
          Else
             strSQLP = strSQLP & " And (a0e23 is null or a0e23='" & strCmp & "') "
          End If
      End If
      '2020/4/17 END
      '20140122END
      If Text1 <> MsgText(601) Then
         strSQLP = strSQLP & " and a0e01 = '" & Text1 & "'"
      End If
      If Text2 <> MsgText(601) Then
         strSQLP = strSQLP & " and a0e07 = '" & Text2 & "'"
      End If
      
      If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
         'Modify By Cheng 2002/01/30
'         strSQL = strSQL & " and a0e10 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
         strSQLR = strSQLR & " And (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null and a0e21 >= " & Val(FCDate(MaskEdBox1.Text)) & ") "
         strSQLP = strSQLP & " And (a0e04 = 'P' and a0e37 <> 0 and a0e37 >= " & Val(FCDate(MaskEdBox1.Text)) & ") "
      End If
      If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
         'Modify By Cheng 2002/01/30
'         strSQL = strSQL & " and a0e10 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
         strSQLR = strSQLR & " And (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null and a0e21 <= " & Val(FCDate(MaskEdBox2.Text)) & ") "
         strSQLP = strSQLP & " And (a0e04 = 'P' and a0e37 <> 0 and a0e37 <= " & Val(FCDate(MaskEdBox2.Text)) & ") "
      End If
   End If
   strUnion = ""
   If strSql <> MsgText(601) Or strSQLR <> MsgText(601) Or strSQLP <> MsgText(601) Then
      'Modify By Cheng 2002/01/30
'      strUnion = "select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, cu04 as contect from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and ((a0e04 = 'P' and a0e10 <> 0) or (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null))" & strSQL
'      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, a0i02 as contect from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and ((a0e04 = 'P' and a0e10 <> 0) or (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null))" & strSQL
'      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, st02 as contect from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and ((a0e04 = 'P' and a0e10 <> 0) or (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null))" & strSQL
'      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, '' as contect from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and ((a0e04 = 'P' and a0e10 <> 0) or (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null))" & strSQL & " order by a0e10 asc, a0e02 asc"
      Select Case Left(Me.Combo1.Text, 1)
      Case "R"
         '處理應收
         strUnion = "select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, cu04 as contect from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) " & strSql
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, a0i02 as contect from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) " & strSql
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, st02 as contect from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) " & strSql
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, '' as contect from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) " & strSql
      Case "P"
         '處理應付
         strUnion = IIf(strUnion = "", "", " union ") & "select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, cu04 as contect from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and (a0e04 = 'P' and a0e37 <> 0) " & strSql
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, a0i02 as contect from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and (a0e04 = 'P' and a0e37 <> 0) " & strSql
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, st02 as contect from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and (a0e04 = 'P' and a0e37 <> 0) " & strSql
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, '' as contect from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and (a0e04 = 'P' and a0e37 <> 0) " & strSql
      Case Else
         '處理應收
         strUnion = "select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, cu04 as contect from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) " & strSql & strSQLR
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, a0i02 as contect from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) " & strSql & strSQLR
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, st02 as contect from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) " & strSql & strSQLR
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, '' as contect from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) " & strSql & strSQLR
         '處理應付
         strUnion = IIf(strUnion = "", "", strUnion & " union ") & "select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, cu04 as contect from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and (a0e04 = 'P' and a0e37 <> 0) " & strSql & strSQLP
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, a0i02 as contect from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and (a0e04 = 'P' and a0e37 <> 0) " & strSql & strSQLP
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, st02 as contect from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and (a0e04 = 'P' and a0e37 <> 0) " & strSql & strSQLP
         strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, '' as contect from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and (a0e04 = 'P' and a0e37 <> 0) " & strSql & strSQLP
      End Select
      strUnion = strUnion & " order by a0e10 asc, a0e02 asc"
      adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
   '未設定任何條件時
   Else
      'Modify By Cheng 2002/01/30
'      strUnion = "select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, cu04 as contect from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and ((a0e04 = 'P' and a0e10 <> 0) or (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null)) and a0e01 like '%'"
'      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, a0i02 as contect from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and ((a0e04 = 'P' and a0e10 <> 0) or (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null)) and a0e01 like '%'"
'      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, st02 as contect from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and ((a0e04 = 'P' and a0e10 <> 0) or (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null)) and a0e01 like '%'"
'      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, '' as contect from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and ((a0e04 = 'P' and a0e10 <> 0) or (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null)) and a0e01 like '%'" & strSQL & " order by a0e10 asc, a0e02 asc"
      '處理應收
      strUnion = "select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, cu04 as contect from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) "
      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, a0i02 as contect from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) "
      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, st02 as contect from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) "
      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e21, '' as contect from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null) "
      
      '處理應付
      strUnion = IIf(strUnion = "", "", strUnion & " union ") & "select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, cu04 as contect from acc0e0, acc0g0, customer where a0e01 = a0g01 and a0e05 = '1' and substr(a0e06, 1, 8) = cu01 (+) and substr(a0e06, 9, 1) = cu02 (+) and (a0e04 = 'P' and a0e37 <> 0) "
      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, a0i02 as contect from acc0e0, acc0g0, acc0i0 where a0e01 = a0g01 and a0e05 = '2' and a0e06 = a0i01 (+) and (a0e04 = 'P' and a0e37 <> 0) "
      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, st02 as contect from acc0e0, acc0g0, staff where a0e01 = a0g01 and a0e05 = '3' and a0e06 = st01 (+) and (a0e04 = 'P' and a0e37 <> 0) "
      strUnion = strUnion & " union select a0e01, a0e02, a0e20, a0e08, a0e04, a0e11, a0e13, a0e10, a0e37 as a0e21, '' as contect from acc0e0, acc0g0 where a0e01 = a0g01 and a0e05 = '4' and (a0e04 = 'P' and a0e37 <> 0) "
      
      strUnion = strUnion & " order by a0e10 asc, a0e02 asc"
      adoadodc1.Open strUnion, adoTaie, adOpenStatic, adLockReadOnly
   End If
   Adodc1.Recordset.Requery
   'Add / Modify By Cheng 2002/01/30
   m_dbl_RsCnts = 0
   If Adodc1.Recordset.RecordCount = 0 Then
      Adodc1.Recordset.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   Else
      m_dbl_RsCnts = Me.adoadodc1.RecordCount
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

Private Sub Text2_GotFocus()
   TextInverse Text2
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
            SumShow
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
'   If strSQL <> "" Then
'      strSQL = " where " & Mid(strSQL, 5, Len(strSQL) - 4)
'   End If
   
   'Modify By Cheng 2002/01/30
'   adoaccsum.CursorLocation = adUseClient
'   adoaccsum.Open "select SUM(A0E11) from ACC0E0 where ((a0e04 = 'P' and a0e10 <> 0) or (a0e04 = 'R' and a0e21 <> 0 and a0e21 is not null))" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccsum.RecordCount <> 0 Then
'      If IsNull(adoaccsum.Fields(0).Value) Then
'         Text4 = MsgText(601)
'      Else
'         Text4 = Format(adoaccsum.Fields(0).Value, DDollar)
'      End If
'   Else
'      Text4 = MsgText(601)
'   End If
'   adoaccsum.Close
   
   'Add By Cheng 2002/01/30
   Text4 = MsgText(601)
   Text3 = MsgText(601)   'ADD BY SONIA 2013/6/18
   With Me.Adodc1
      If .Recordset.State = adStateOpen Then
         If .Recordset.RecordCount > 0 Then .Recordset.MoveFirst
         While Not Me.Adodc1.Recordset.EOF
            Text4.Text = Val(Text4.Text) + .Recordset.Fields(5).Value
            Text3.Text = Val(Text3.Text) + 1   'ADD BY SONIA 2013/6/18
            .Recordset.MoveNext
         Wend
         Text4.Text = Format(Text4.Text, DDollar)
         If .Recordset.RecordCount > 0 Then Me.Adodc1.Recordset.MoveFirst
      End If
   End With
   
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

'Modify by Sindy 2020/04/17 公司別改下拉
''20140122START By eric
'Private Sub Text5_LostFocus()
'   If Text5.Text <> "1" And Text5.Text <> "2" And Text5.Text <> "" Then
'      MsgBox "公司別僅可為 1 / 2 或不輸入!"
'      Text5.Text = ""
'      Text5.SetFocus
'      Exit Sub
'   End If
'End Sub
'
''20140122START By eric
'Private Sub Text5_GotFocus()
'   TextInverse Text5
'   CloseIme
'End Sub
'
''20140122START By eric
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub

'Add by Sindy 2020/04/17
Private Sub SetCompN()
    strCmpN = "": strCmp = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, True, True)
End Sub

'Add by Sindy 2020/04/17
Private Sub CboCmp_GotFocus()
    TextInverse CboCmp
End Sub

Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label6 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/04/17
